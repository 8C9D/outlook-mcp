// Microsoft Graph tokens for the Worker, held in KV.
//
// MSAL Node does not run on workerd (it reaches for node crypto/network APIs
// that nodejs_compat does not provide), so the refresh-token grant is issued
// here with plain fetch. The mailbox refresh token is seeded once from the
// local MSAL cache by scripts/seed-kv.ts; from then on this module owns it.
//
// Microsoft rotates the refresh token on every exchange and invalidates the
// previous one, so the new value MUST be written back or the connector locks
// itself out on the following call.
import { KV_ACCESS_TOKEN, KV_REFRESH_TOKEN } from "../core/kv-keys.js";
import { AuthRequiredError } from "../core/token.js";
import type { Env } from "./env.js";

export const TOKEN_ENDPOINT = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";
export const DEVICE_CODE_ENDPOINT =
  "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode";
export const GRAPH_ME = "https://graph.microsoft.com/v1.0/me";


/**
 * Exactly the scopes the local MSAL cache was consented for. Re-requesting the
 * same set means the refresh grant never triggers a new consent prompt.
 */
export const MAILBOX_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
  "Contacts.ReadWrite",
  "MailboxSettings.ReadWrite",
  "Tasks.ReadWrite",
].join(" ");

/** Identity-only scope for proving who is authorizing the connector. */
export const IDENTITY_SCOPES = "offline_access User.Read";

type CachedAccess = { token: string; expiresAt: number };

/** Renew this many seconds before the access token actually expires. */
const EXPIRY_SKEW_SECONDS = 300;

type TokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

/**
 * Exchange the stored refresh token for a fresh access token and persist the
 * rotated refresh token. Returns the new access token.
 */
export async function refreshMailboxToken(env: Env): Promise<string> {
  const refreshToken = await env.OUTLOOK_KV.get(KV_REFRESH_TOKEN);
  if (!refreshToken) {
    throw new AuthRequiredError(
      `no refresh token in KV under ${KV_REFRESH_TOKEN}; run \`npm run seed:kv\` locally`
    );
  }

  const response = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: MAILBOX_SCOPES,
    }),
  });

  const payload = (await response.json().catch(() => ({}))) as TokenResponse;
  if (!response.ok || !payload.access_token) {
    throw new AuthRequiredError(
      `Microsoft refused the refresh grant (HTTP ${response.status} ${payload.error ?? ""}: ` +
        `${payload.error_description ?? "no detail"}). Re-run \`npm run login\` locally, ` +
        "then `npm run seed:kv` to reseed KV."
    );
  }

  // Order matters: persist the rotated refresh token before anything can fail,
  // otherwise the token just spent is lost and the connector is locked out.
  if (payload.refresh_token && payload.refresh_token !== refreshToken) {
    await env.OUTLOOK_KV.put(KV_REFRESH_TOKEN, payload.refresh_token);
  }

  const lifetime = payload.expires_in ?? 3600;
  const cached: CachedAccess = {
    token: payload.access_token,
    expiresAt: Date.now() + Math.max(lifetime - EXPIRY_SKEW_SECONDS, 60) * 1000,
  };
  await env.OUTLOOK_KV.put(KV_ACCESS_TOKEN, JSON.stringify(cached), {
    // KV's floor is 60s; never outlive the token itself.
    expirationTtl: Math.max(lifetime, 60),
  });

  return payload.access_token;
}

/** Cached mailbox access token, refreshing through KV only when it is stale. */
export async function getMailboxAccessToken(env: Env): Promise<string> {
  const raw = await env.OUTLOOK_KV.get(KV_ACCESS_TOKEN);
  if (raw) {
    try {
      const cached = JSON.parse(raw) as CachedAccess;
      if (cached.token && cached.expiresAt > Date.now()) return cached.token;
    } catch {
      // Unparseable cache entry: fall through and mint a new token.
    }
  }
  return refreshMailboxToken(env);
}

/** Build the TokenProvider the tool layer consumes for one Worker request. */
export function mailboxTokenProvider(env: Env): () => Promise<string> {
  return () => getMailboxAccessToken(env);
}

export type GraphIdentity = {
  id: string;
  userPrincipalName?: string;
  mail?: string;
  displayName?: string;
};

/** Resolve a Microsoft access token to a Graph identity, or null if it is not valid. */
export async function resolveIdentity(accessToken: string): Promise<GraphIdentity | null> {
  const response = await fetch(`${GRAPH_ME}?$select=id,userPrincipalName,mail,displayName`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok) return null;
  return (await response.json()) as GraphIdentity;
}

/**
 * The single-user allowlist. A Microsoft identity may authorize this connector
 * only if its immutable Graph id matches, or — as a fallback for the same
 * account seen through a different id shape — its UPN/mail matches.
 */
export function isAllowedIdentity(identity: GraphIdentity, env: Env): boolean {
  if (env.ALLOWED_MS_USER_ID && identity.id === env.ALLOWED_MS_USER_ID) return true;
  const upn = env.ALLOWED_MS_UPN?.toLowerCase();
  if (!upn) return false;
  return (
    identity.userPrincipalName?.toLowerCase() === upn || identity.mail?.toLowerCase() === upn
  );
}
