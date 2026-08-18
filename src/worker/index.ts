// Cloudflare Worker entry point: the same 24 tools, 2 prompts and 2 resources
// as the stdio server, served over MCP Streamable HTTP and gated by OAuth,
// plus the three things only a hosted server can do — receive Graph change
// notifications, keep their subscription alive on a schedule, and hand out
// short-lived authenticated links to attachment bytes it cannot save to disk.
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
import type { Env } from "./env.js";

/** The only scope this server issues; the mailbox permissions are fixed at consent time. */
const SCOPES_SUPPORTED = ["outlook"];

/**
 * Both protected routes behind one handler: the MCP endpoint, and the
 * attachment downloads get_attachment hands out. Listing /download/ as an
 * apiRoute is what puts it behind the same bearer check as /mcp — a link with
 * no token gets OAuthProvider's 401, not the file.
 */
const apiHandler = {
  fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    return new URL(request.url).pathname.startsWith(DOWNLOAD_ROUTE_PREFIX)
      ? downloadHandler.fetch(request, env, ctx)
      : mcpHandler.fetch(request, env, ctx);
  },
};

const oauthProvider = new OAuthProvider<Env>({
  apiRoute: ["/mcp", DOWNLOAD_ROUTE_PREFIX],
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
   * Cron trigger (see `triggers.crons` in wrangler.jsonc). Mail subscriptions
   * expire after ~2.9 days and Graph drops them silently, so this runs far more
   * often than that and creates, renews or leaves the subscription alone as
   * needed. A failure must not throw out of the scheduled handler — it would be
   * retried on the next tick anyway — so it is logged instead.
   */
  async scheduled(event: ScheduledController, env: Env, ctx: ExecutionContext): Promise<void> {
    ctx.waitUntil(
      keepSubscriptionAlive(env).then(
        (result) => console.log(`Cron ${event.cron}: mail subscription ${result.action}.`),
        (err) => console.error(`Cron ${event.cron}: subscription upkeep failed: ${String(err)}`)
      )
    );
  },
} satisfies ExportedHandler<Env>;
