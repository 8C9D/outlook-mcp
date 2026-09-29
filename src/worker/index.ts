// Cloudflare Worker entry point: the same tools, prompts and resources as the
// stdio server (both build from core/registry.js, so neither can drift), served
// over MCP Streamable HTTP and gated by OAuth, plus the three things only a
// hosted server can do — receive Graph change notifications, keep their
// subscription alive on a schedule, and hand out short-lived authenticated
// links to attachment bytes it cannot save to disk. The opt-in self-alert
// routes (core/self-alert.js) are answered before OAuthProvider sees the
// request: they carry their own shared secret, and are 404 unless it is set.
//
// OAuthProvider owns the whole authorization-server surface — discovery
// metadata, dynamic client registration, PKCE, the token endpoint, bearer
// validation and the 401 + WWW-Authenticate challenge that tells an MCP client
// where to start. It routes an authenticated request to `apiHandler` and
// everything else to `defaultHandler`; nothing anonymous can reach /mcp.
//
// It exposes only a fetch handler, so the cron trigger is wired by wrapping it
// rather than exporting it directly.
import { OAuthProvider } from "@cloudflare/workers-oauth-provider";
import { DOWNLOAD_ROUTE_PREFIX } from "../core/downloads.js";
import { defaultHandler } from "./authorize.js";
import { downloadHandler } from "./download.js";
import { mcpHandler } from "./mcp-handler.js";
import { keepSubscriptionAlive, publicBaseUrl } from "./notifications.js";
import { draftMorningBrief, reconcileFilingCorrections } from "./llm.js";
import { runWorkerHealthCheck } from "./health.js";
import { handleSelfAlert, isSelfAlertPath, runWorkerSelfAlertWatchdog } from "./self-alert.js";
import { scheduledJobFor } from "../core/schedule.js";
import type { Env } from "./env.js";

// The schedule constants (which must match wrangler.jsonc) and the tick →
// job dispatch live in core/schedule.js, where the offline tier can test them.
export { DIGEST_HOUR_TORONTO, HEALTH_CRON, SELF_ALERT_WATCHDOG_CRON } from "../core/schedule.js";

/** The only scope this server issues; the mailbox permissions are fixed at consent time. */
const SCOPES_SUPPORTED = ["outlook"];

/**
 * Both protected surfaces behind one handler: the MCP endpoint itself, and the
 * attachment downloads get_attachment hands out. The downloads live under
 * /mcp/, so the single apiRoute below covers them (it is prefix-matched) and
 * they carry the same bearer check — a link with no token gets OAuthProvider's
 * 401, not the file. See DOWNLOAD_ROUTE_PREFIX for why they cannot sit at the
 * root instead.
 */
const apiHandler = {
  fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    return new URL(request.url).pathname.startsWith(DOWNLOAD_ROUTE_PREFIX)
      ? downloadHandler.fetch(request, env, ctx)
      : mcpHandler.fetch(request, env, ctx);
  },
};

/**
 * RFC 9728 requires the advertised resource to match the URL pasted into the
 * client exactly, so it has to be this deployment's own origin — which lives in
 * the PUBLIC_BASE_URL binding (injected at deploy time), not in this file. Bindings are not readable
 * at module scope, so the provider is built on the first request and memoized
 * per origin (a Worker isolate only ever sees one).
 */
let cached: { origin: string; provider: OAuthProvider<Env> } | undefined;

function providerFor(request: Request, env: Env): OAuthProvider<Env> {
  // Falling back to the request's own origin keeps `wrangler dev` and any fork
  // working with no configuration; a deployment sets the var so the value does
  // not depend on which hostname a caller happened to use.
  const origin = publicBaseUrl(env) ?? new URL(request.url).origin;
  if (cached?.origin === origin) return cached.provider;

  const provider = new OAuthProvider<Env>({
    apiRoute: "/mcp",
    apiHandler,
    defaultHandler,

    authorizeEndpoint: "/authorize",
    tokenEndpoint: "/oauth/token",
    // claude.ai registers itself as a public client via RFC 7591; without this
    // endpoint the connector has no way to obtain a client_id.
    clientRegistrationEndpoint: "/oauth/register",

    scopesSupported: SCOPES_SUPPORTED,

    resourceMetadata: {
      resource: `${origin}/mcp`,
      scopes_supported: SCOPES_SUPPORTED,
      bearer_methods_supported: ["header"],
      resource_name: "Outlook MCP",
    },
  });
  cached = { origin, provider };
  return provider;
}

export default {
  fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    // Shared-secret routes, not OAuth ones: answered before the provider, which
    // would otherwise hand them to defaultHandler. 404 unless SELF_ALERT_SECRET
    // is set.
    if (isSelfAlertPath(new URL(request.url).pathname)) return handleSelfAlert(request, env);
    return providerFor(request, env).fetch(request, env, ctx);
  },

  /**
   * Cron triggers (see `triggers.crons` in wrangler.jsonc). Four jobs share
   * the handler, dispatched by scheduledJobFor (core/schedule.js): the health
   * check and the self-alert watchdog on their exact cron expressions, first,
   * and the remaining ticks by the hour America/Toronto is actually on:
   *
   *  - The health check, daily at 13:37 UTC (HEALTH_CRON). Verifies
   *    KV, a forced token rotation, the Graph subscription and the two LLM
   *    error counters; healthy runs write a heartbeat, failing ones also leave
   *    an unsent alert draft in the inbox.
   *  - The self-alert watchdog, hourly at minute 47 (SELF_ALERT_WATCHDOG_CRON).
   *    Emails the owner once when a job that registered a heartbeat has gone
   *    quiet for longer than it asked. Does nothing unless SELF_ALERT_SECRET is
   *    set. One of its 11:47/12:47 UTC ticks falls in Toronto's 07:00 hour,
   *    which is why it is dispatched before the hour check below.
   *  - Subscription upkeep, every 6 hours. Mail subscriptions expire after ~2.9
   *    days and Graph drops them silently, so this runs far more often than
   *    that and creates, renews or leaves the subscription alone as needed.
   *  - The morning digest, at 07:00 America/Toronto. Cloudflare crons are UTC
   *    only, and 07:00 Toronto is 11:00 UTC in EDT and 12:00 UTC in EST, so
   *    BOTH are scheduled and this guard drops the one that is not 07:00 right
   *    now. That is what keeps the brief at 07:00 local across a DST change
   *    with no redeploy; core/digest.js additionally refuses to draft a second
   *    brief for a date it has already covered, so a double fire cannot double
   *    up either.
   *
   * A failure must not throw out of the scheduled handler — it would be retried
   * on the next tick anyway — so everything is logged instead.
   */
  async scheduled(event: ScheduledController, env: Env, ctx: ExecutionContext): Promise<void> {
    const job = scheduledJobFor(event.cron, new Date(event.scheduledTime));

    if (job === "health") {
      ctx.waitUntil(
        runWorkerHealthCheck(env).then(
          (report) =>
            console.log(
              `Cron ${event.cron}: health check ${report.healthy ? "healthy" : "UNHEALTHY"}` +
                (report.alertDraftId ? ` — alert draft ${report.alertDraftId}` : "") +
                (report.alertError ? ` — ${report.alertError}` : "")
            ),
          (err) => console.error(`Cron ${event.cron}: health check failed: ${String(err)}`)
        )
      );
      return;
    }

    if (job === "self-alert-watchdog") {
      ctx.waitUntil(
        runWorkerSelfAlertWatchdog(env).then(
          (result) => {
            if (!result.enabled) return; // off: stay quiet every hour
            console.log(
              `Cron ${event.cron}: self-alert watchdog — ${result.watched} watched, ` +
                `${result.stale} stale, ${result.alerted.length} alerted` +
                (result.alerted.length ? ` (${result.alerted.join(", ")})` : "") +
                "."
            );
            for (const problem of result.problems) {
              console.error(`Cron ${event.cron}: self-alert watchdog: ${problem}`);
            }
          },
          (err) => console.error(`Cron ${event.cron}: self-alert watchdog failed: ${String(err)}`)
        )
      );
      return;
    }

    if (job === "digest") {
      ctx.waitUntil(
        draftMorningBrief(env).then(
          (outcome) => console.log(`Cron ${event.cron}: morning brief — ${outcome.reason}.`),
          (err) => console.error(`Cron ${event.cron}: morning brief failed: ${String(err)}`)
        )
      );
      return;
    }

    ctx.waitUntil(
      keepSubscriptionAlive(env).then(
        (result) => console.log(`Cron ${event.cron}: mail subscription ${result.action}.`),
        (err) => console.error(`Cron ${event.cron}: subscription upkeep failed: ${String(err)}`)
      )
    );
    // The auto-filer's feedback loop also reconciles on every accepted
    // notification delivery; this tick covers quiet stretches. No-op while
    // filing is disabled.
    ctx.waitUntil(
      reconcileFilingCorrections(env).catch((err) =>
        console.error(`Cron ${event.cron}: correction reconcile failed: ${String(err)}`)
      )
    );
  },
} satisfies ExportedHandler<Env>;
