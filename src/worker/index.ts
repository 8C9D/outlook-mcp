// Cloudflare Worker entry point: the same 21 tools and 2 prompts as the stdio
// server, served over MCP Streamable HTTP and gated by OAuth.
//
// OAuthProvider owns the whole authorization-server surface — discovery
// metadata, dynamic client registration, PKCE, the token endpoint, bearer
// validation and the 401 + WWW-Authenticate challenge that tells an MCP client
// where to start. It routes an authenticated request to `apiHandler` and
// everything else to `defaultHandler`; nothing anonymous can reach /mcp.
import { OAuthProvider } from "@cloudflare/workers-oauth-provider";
import { defaultHandler } from "./authorize.js";
import { mcpHandler } from "./mcp-handler.js";
import type { Env } from "./env.js";

/** The only scope this server issues; the mailbox permissions are fixed at consent time. */
const SCOPES_SUPPORTED = ["outlook"];

export default new OAuthProvider<Env>({
  apiRoute: "/mcp",
  apiHandler: mcpHandler,
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
