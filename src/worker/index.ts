// Cloudflare Worker entry point: the same tools, prompts and resources as the
// stdio server (both build from core/registry.js, so neither can drift), served
// over MCP Streamable HTTP and gated by OAuth, plus the three things only a
// hosted server can do — receive Graph change notifications, keep their
// subscription alive on a schedule, and hand out short-lived authenticated
// links to attachment bytes it cannot save to disk.
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
import { keepSubscriptionAlive } from "./notifications.js";
import { draftMorningBrief } from "./llm.js";
import { runWorkerHealthCheck } from "./health.js";
import { torontoHourOf } from "../core/auto-filing.js";
import type { Env } from "./env.js";

/** The only scope this server issues; the mailbox permissions are fixed at consent time. */
const SCOPES_SUPPORTED = ["outlook"];

/** The hour, America/Toronto, at which the morning brief is drafted. */
export const DIGEST_HOUR_TORONTO = 7;

/**
 * The daily health-check schedule (must match wrangler.jsonc). 13:37 UTC is
 * 09:37 Toronto in EDT and 08:37 in EST — always morning for the owner, after
 * the digest, and colliding with neither the 6-hourly upkeep ticks (minute 17
 * of hours 5/11/17/23 UTC) nor the digest ticks (11:00 and 12:00 UTC). This
 * one is dispatched on the cron expression rather than the local hour, since
 * unlike the digest it has no wall-clock meaning to preserve across DST.
 */
export const HEALTH_CRON = "37 13 * * *";

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

const oauthProvider = new OAuthProvider<Env>({
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
    // Filled from the wrangler var so the advertised resource matches the URL
    // pasted into the client exactly, which RFC 9728 requires.
    resource: "https://outlook-mcp.arthur-yuhao-zhang.workers.dev/mcp",
    scopes_supported: SCOPES_SUPPORTED,
    bearer_methods_supported: ["header"],
    resource_name: "Outlook MCP",
  },
});

export default {
  fetch: (request: Request, env: Env, ctx: ExecutionContext) =>
    oauthProvider.fetch(request, env, ctx),

  /**
   * Cron triggers (see `triggers.crons` in wrangler.jsonc). Three jobs share
   * the handler: the daily health check is dispatched on its exact cron
   * expression, and the remaining ticks are told apart by the hour
   * America/Toronto is actually on:
   *
   *  - The health check, daily at 13:37 UTC (see HEALTH_CRON above). Verifies
   *    KV, a forced token rotation, the Graph subscription and the two LLM
   *    error counters; healthy runs write a heartbeat, failing ones also leave
   *    an unsent alert draft in the inbox.
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
    const hour = torontoHourOf(new Date(event.scheduledTime));

    if (event.cron === HEALTH_CRON) {
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

    if (hour === DIGEST_HOUR_TORONTO) {
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
  },
} satisfies ExportedHandler<Env>;
