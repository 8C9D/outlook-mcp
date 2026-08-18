// Bindings the Worker expects: the two KV namespaces from wrangler.jsonc, the
// secrets set with `wrangler secret put`, and the OAuth helpers OAuthProvider
// injects into every handler it dispatches to.
import type { OAuthHelpers } from "@cloudflare/workers-oauth-provider";

export interface Env {
  /** Injected by OAuthProvider; not declared in wrangler.jsonc. */
  OAUTH_PROVIDER: OAuthHelpers;
  /** Microsoft token store: the rotating refresh token and the cached access token. */
  OUTLOOK_KV: KVNamespace;
  /** Grants, authorization codes and bearer tokens owned by workers-oauth-provider. */
  OAUTH_KV: KVNamespace;
  /** Application (client) ID of the outlook-mcp Entra app registration (public client). */
  MS_CLIENT_ID: string;
  /** Graph /me `id` of the single account allowed to authorize this connector. */
  ALLOWED_MS_USER_ID: string;
  /** Graph /me `userPrincipalName`/`mail` of that same account (second accepted match). */
  ALLOWED_MS_UPN: string;
}
