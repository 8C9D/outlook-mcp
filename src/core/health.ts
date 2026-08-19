// The daily self-monitoring check for the hosted server.
//
// Once a day (cron "37 13 * * *", see wrangler.jsonc) the Worker verifies the
// four things that break silently: the Graph change-notification subscription,
// the refresh-token rotation the whole connector hangs on, the error counters
// the two LLM features increment when they swallow failures, and KV itself.
//
// A healthy run writes a heartbeat record to KV and nothing else. Any failing
// check produces a DRAFT in the owner's own inbox — subject
// "outlook-mcp health: <failing checks>" — naming what failed, since when, and
// how to fix it. The draft is NEVER sent: it is created directly in the inbox
// as an unsent message, so send_draft remains the only send path in this
// codebase and a dead server cannot mail anyone.
//
// Everything is injectable (store, Graph transport, the rotation, the drafting)
// so every failure mode is unit-testable offline; the Worker wiring in
// src/worker/health.ts supplies the real dependencies.
import {
  readFeatureErrors,
  torontoDateOf,
  type FeatureErrorRecord,
} from "./auto-filing.js";
import { GraphError, callGraphServer } from "./graph.js";
import { HEALTH_PROBE_KEY, STATE_HEALTH, STATE_SUBSCRIPTION } from "./kv-keys.js";
import { readJson, writeJson, type StateStore } from "./state.js";
import type { SubscriptionRecord } from "./subscriptions.js";

/** A feature whose swallowed errors reach this many in one Toronto day fails. */
export const HEALTH_ERROR_THRESHOLD = 5;

/** Every check the health run performs, in the order it performs them. */
export const HEALTH_CHECKS = [
  "kv",
  "token_refresh",
  "subscription",
  "filing_errors",
  "digest_errors",
] as const;

export type HealthCheckName = (typeof HEALTH_CHECKS)[number];

export type HealthCheckResult = {
  name: HealthCheckName;
  ok: boolean;
  detail: string;
  /** When this check first started failing; carried forward across runs. */
  failingSince?: string;
};

export type HealthReport = {
  at: string;
  healthy: boolean;
  checks: HealthCheckResult[];
  /** Id of the alert draft placed in the inbox, when one was created. */
  alertDraftId?: string;
  /** Why the alert draft could not be created, when it could not. */
  alertError?: string;
};

type GraphCall = (path: string, init?: RequestInit) => Promise<any>;

export type HealthDeps = {
  store: StateStore;
  /** Force one refresh-token rotation; must throw when Microsoft refuses. */
  refreshToken: () => Promise<void>;
  /** Graph transport for the subscription check and the alert draft. */
  graph?: GraphCall;
  /** Create the alert DRAFT in the inbox and return its id. Never sends. */
  draftAlert?: (subject: string, body: string) => Promise<string>;
  now?: () => Date;
  errorThreshold?: number;
};

/** The stored heartbeat, or null when the health cron has never completed. */
export async function readHealthReport(store: StateStore): Promise<HealthReport | null> {
  return readJson<HealthReport>(store, STATE_HEALTH);
}

/** "outlook-mcp health: subscription, token_refresh" — the draft's subject. */
export function healthAlertSubject(failing: HealthCheckResult[]): string {
  return `outlook-mcp health: ${failing.map((check) => check.name).join(", ")}`;
}

/** Per-check fix pointers, written for the person who finds the draft. */
const FIX_POINTERS: Record<HealthCheckName, string> = {
  kv:
    "KV is this server's only storage; nothing works without it. Check the Cloudflare " +
    "dashboard and status page, and `npx wrangler tail outlook-mcp` for the live errors.",
  token_refresh:
    "Re-seed the mailbox credential: locally run `npm run login` (interactive sign-in), " +
    "then `npm run seed:kv` to put the fresh refresh token back into KV.",
  subscription:
    "The inbox subscription normally heals itself — the 6-hourly upkeep cron and every " +
    "authenticated MCP request re-create it. If it stays down, watch " +
    "`npx wrangler tail outlook-mcp` while sending yourself a message, and check that " +
    "PUBLIC_BASE_URL still matches the deployed hostname.",
  filing_errors:
    "Auto-filing swallowed repeated failures today. `get_auto_filing_log` shows each " +
    "one's reason; `npx wrangler tail outlook-mcp` shows them live. `manage_auto_filing " +
    "disable_filing` stops the feature while you look.",
  digest_errors:
    "The morning digest swallowed repeated failures today. `get_auto_filing_log` shows " +
    "the reasons; `manage_auto_filing disable_digest` stops it while you look.",
};

/** The alert draft's body: what failed, since when, and what to do about it. */
export function healthAlertBody(report: HealthReport, failing: HealthCheckResult[]): string {
  const lines: string[] = [
    `The daily outlook-mcp health check at ${report.at} found ` +
      `${failing.length} failing check(s).`,
    "",
  ];
  for (const check of failing) {
    lines.push(
      `FAILING: ${check.name}`,
      `  Since: ${check.failingSince ?? report.at}`,
      `  Detail: ${check.detail}`,
      `  Fix: ${FIX_POINTERS[check.name]}`,
      ""
    );
  }
  const passing = report.checks.filter((check) => check.ok);
  if (passing.length > 0) {
    lines.push(`Passing: ${passing.map((check) => check.name).join(", ")}`, "");
  }
  lines.push(
    "—",
    "This draft was created (never sent) by the health cron on the hosted outlook-mcp " +
      "server. get_health shows the latest results; delete this draft once handled."
  );
  return lines.join("\n");
}

/** Default drafter: a plain unsent message created directly in the inbox. */
async function graphDraftAlert(graph: GraphCall, subject: string, body: string): Promise<string> {
  const me = await graph("/me?$select=mail,userPrincipalName");
  const address = me?.mail ?? me?.userPrincipalName;
  if (!address) throw new Error("Graph did not report an address for this mailbox.");
  const draft = await graph("/me/mailFolders/inbox/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      subject,
      body: { contentType: "Text", content: body },
      toRecipients: [{ emailAddress: { address: String(address) } }],
    }),
  });
  if (!draft?.id) throw new Error("Graph did not return an id for the health alert draft.");
  return String(draft.id);
}

/**
 * Run every check, write the heartbeat, and deliver an alert draft when
 * anything failed. Individual checks never throw out of this function: each
 * failure becomes a failing check result, because the whole point is a report.
 */
export async function runHealthCheck(deps: HealthDeps): Promise<HealthReport> {
  const now = deps.now ?? (() => new Date());
  const graph = deps.graph ?? callGraphServer;
  const threshold = deps.errorThreshold ?? HEALTH_ERROR_THRESHOLD;
  const at = now().toISOString();
  const checks: HealthCheckResult[] = [];

  // 1. KV round-trip. Runs first because every other record read below is
  // meaningless if storage itself is down.
  checks.push(
    await runCheck("kv", async () => {
      const probe = `${at} ${Math.random().toString(16).slice(2)}`;
      await deps.store.put(HEALTH_PROBE_KEY, probe, { ttlSeconds: 300 });
      const back = await deps.store.get(HEALTH_PROBE_KEY);
      if (back !== probe) {
        throw new Error(`the probe value did not round-trip (read back: ${back ?? "null"})`);
      }
      return "a probe value round-tripped through the store";
    })
  );

  // 2. One forced refresh-token rotation, through the same machinery every
  // Graph call uses. If this breaks, the connector locks out within the hour.
  checks.push(
    await runCheck("token_refresh", async () => {
      await deps.refreshToken();
      return "a forced refresh-token rotation succeeded and the new token is stored";
    })
  );

  // 3. The subscription named by the KV record must be alive in Graph with a
  // future expiry — Graph deletes lapsed subscriptions silently.
  checks.push(
    await runCheck("subscription", async () => {
      const record = await readJson<SubscriptionRecord>(deps.store, STATE_SUBSCRIPTION);
      if (!record?.id) {
        throw new Error(
          "no subscription record in KV — the upkeep cron has never completed, or KV was cleared"
        );
      }
      let live: any;
      try {
        live = await graph(`/subscriptions/${encodeURIComponent(record.id)}`);
      } catch (err) {
        if (err instanceof GraphError) {
          throw new Error(
            `Graph does not have subscription ${record.id} (HTTP ${err.status}) — ` +
              "deliveries stopped; the upkeep cron should recreate it"
          );
        }
        throw err;
      }
      const expiresAt = Date.parse(live?.expirationDateTime ?? "");
      if (!Number.isFinite(expiresAt) || expiresAt <= now().getTime()) {
        throw new Error(
          `subscription ${record.id} expired at ${live?.expirationDateTime ?? "(unknown)"}`
        );
      }
      return `subscription ${record.id} is live, expires ${live.expirationDateTime}`;
    })
  );

  // 4 + 5. The error counters the two LLM features increment when their
  // background paths swallow a failure.
  const today = torontoDateOf(now());
  for (const feature of ["filing", "digest"] as const) {
    checks.push(
      await runCheck(`${feature}_errors`, async () => {
        const record = await readFeatureErrors(deps.store, feature, today).catch(
          () => null as FeatureErrorRecord | null
        );
        const count = record?.count ?? 0;
        if (count >= threshold) {
          throw new Error(
            `${count} swallowed error(s) today (threshold ${threshold}); ` +
              `last at ${record!.lastAt}: ${record!.lastReason}`
          );
        }
        return `${count} swallowed error(s) today (threshold ${threshold})`;
      })
    );
  }

  // Carry "since when" across runs: a check failing in the previous report
  // keeps that report's failingSince, so the alert can say how long.
  const previous = await readHealthReport(deps.store).catch(() => null);
  for (const check of checks) {
    if (check.ok) continue;
    const before = previous?.checks.find((c) => c.name === check.name && !c.ok);
    check.failingSince = before ? (before.failingSince ?? previous!.at) : at;
  }

  const failing = checks.filter((check) => !check.ok);
  const report: HealthReport = { at, healthy: failing.length === 0, checks };

  if (failing.length > 0) {
    const draft = deps.draftAlert ?? ((s: string, b: string) => graphDraftAlert(graph, s, b));
    try {
      report.alertDraftId = await draft(
        healthAlertSubject(failing),
        healthAlertBody(report, failing)
      );
    } catch (err) {
      // The report still stands; the failure to alert is itself reported.
      report.alertError = `could not create the alert draft: ${
        err instanceof Error ? err.message : String(err)
      }`;
    }
  }

  // The heartbeat is written on every run, healthy or not. If KV itself is
  // down this write fails too — the draft above is then the only signal, which
  // is exactly why the alert is drafted before the heartbeat is written.
  try {
    await writeJson(deps.store, STATE_HEALTH, report);
  } catch {
    // Reported through the kv check already.
  }

  return report;
}

/** One check: any throw becomes {ok: false} with the error as the detail. */
async function runCheck(
  name: HealthCheckName,
  fn: () => Promise<string>
): Promise<HealthCheckResult> {
  try {
    return { name, ok: true, detail: await fn() };
  } catch (err) {
    return { name, ok: false, detail: err instanceof Error ? err.message : String(err) };
  }
}
