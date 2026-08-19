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
  /**
   * Anthropic API key for the two LLM features (auto-filing on the change
   * notification, and the morning digest cron). Absent means both are simply
   * unavailable — they ship disabled anyway, and every call path logs the
   * missing key to the audit trail rather than failing loudly. Never logged,
   * never returned by a tool, never written to KV.
   */
  ANTHROPIC_API_KEY?: string;
  /**
   * Public origin of this Worker (wrangler var, not a secret). Graph is told to
   * deliver change notifications to `${PUBLIC_BASE_URL}/notifications`, so it
   * has to match the deployed hostname exactly.
   */
  PUBLIC_BASE_URL?: string;
  /**
   * Exactly "true" enables the non-interactive POST-with-ms_access_token path
   * on /authorize. Deliberately unset on the deployed Worker (neither a var in
   * wrangler.jsonc nor a secret), so production only ever authorizes through
   * the interactive device-code flow; local `wrangler dev` runs turn it on via
   * the gitignored `.dev.vars`. Remote test r5 asserts production refuses it.
   */
  ALLOW_DIRECT_AUTHORIZE?: string;
}
