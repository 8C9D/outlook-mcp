// The latest self-monitoring results. On the hosted server this reads the
// heartbeat the daily health cron wrote to KV (core/health.js); on the local
// stdio server there is no cron, no KV and no subscription, so the tool runs
// the checks that mean something locally — silent sign-in and mailbox access —
// and says plainly which checks only exist on the hosted server.
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { HEALTH_CHECKS, readHealthReport } from "../core/health.js";
import { requireStateStore } from "../core/state.js";
import { ToolResult, formatLocal, runTool, textResult } from "./common.js";

export const getHealthSchema = {
  include_details: z
    .boolean()
    .default(true)
    .describe("Include each check's detail line, not just the per-check verdicts."),
};

const getHealthArgs = z.object(getHealthSchema);

export const getHealthDescription =
  "Report the server's own health. On the hosted (remote) server: the latest results of the " +
  "daily self-monitoring cron, which verifies KV, a forced Microsoft token rotation, the Graph " +
  "change-notification subscription, and the auto-filing/digest error counters — healthy runs " +
  "write only a heartbeat, and any failing check also leaves an UNSENT alert draft in the inbox " +
  'titled "outlook-mcp health: …". On the local stdio server: live checks of what matters ' +
  "locally (silent sign-in from the token cache, mailbox access), with the remote-only checks " +
  "named as such rather than faked. Changes nothing either way.";

/** The checks that exist only on the hosted server, for the local report. */
const REMOTE_ONLY_NOTE =
  "Remote-only checks (run daily at 13:37 UTC on the hosted Worker; ask its get_health for " +
  `the results): ${HEALTH_CHECKS.join(", ")} — the Graph change-notification subscription, ` +
  "the KV-held refresh token's rotation, the LLM feature error counters, and KV itself all " +
  "live on that server, not here.";

export async function getHealthHandler(input: z.input<typeof getHealthArgs>): Promise<ToolResult> {
  return runTool(async () => {
    const { include_details } = getHealthArgs.parse(input);
    const store = requireStateStore();

    if (store.mode === "remote") {
      const report = await readHealthReport(store);
      if (!report) {
        return textResult(
          "No health report yet: the daily health cron (13:37 UTC) has not completed since " +
            "this feature was deployed. Check again after it has run, or see `wrangler tail` " +
            "if it never appears."
        );
      }
      const lines = [
        `Server health: ${report.healthy ? "HEALTHY" : "UNHEALTHY"} (checked ${formatLocal(report.at)})`,
        "",
        ...report.checks.map((check) => {
          const head = `${check.ok ? "OK  " : "FAIL"}  ${check.name}`;
          const since = !check.ok && check.failingSince ? ` (failing since ${formatLocal(check.failingSince)})` : "";
          return include_details ? `${head}${since} — ${check.detail}` : `${head}${since}`;
        }),
      ];
      if (report.alertDraftId) {
        lines.push(
          "",
          `An unsent alert draft was left in the inbox (message id ${report.alertDraftId}).`
        );
      }
      if (report.alertError) {
        lines.push("", `The alert draft could not be created: ${report.alertError}`);
      }
      return textResult(lines.join("\n"));
    }

    // Local stdio server: run what can honestly be checked from here.
    const lines: string[] = ["Local server health (checked live, just now):", ""];
    let healthy = true;
    const local = async (name: string, fn: () => Promise<string>) => {
      try {
        const detail = await fn();
        lines.push(include_details ? `OK    ${name} — ${detail}` : `OK    ${name}`);
      } catch (err) {
        healthy = false;
        const detail = err instanceof Error ? err.message : String(err);
        lines.push(`FAIL  ${name} — ${detail.split("\n")[0]}`);
      }
    };

    await local("sign-in", async () => {
      const me = await callGraphServer("/me?$select=mail,userPrincipalName");
      const address = me?.mail ?? me?.userPrincipalName ?? "(unknown address)";
      return `silent token acquisition works; signed in as ${address}`;
    });
    await local("mailbox", async () => {
      const inbox = await callGraphServer(
        "/me/mailFolders/inbox?$select=displayName,totalItemCount"
      );
      return `inbox reachable (${inbox?.totalItemCount ?? "?"} message(s)) — mail scopes consented`;
    });

    lines.push("", healthy ? "Both local checks pass." : "Run `npm run doctor` for a diagnosis.");
    lines.push("", REMOTE_ONLY_NOTE);
    return textResult(lines.join("\n"));
  });
}
