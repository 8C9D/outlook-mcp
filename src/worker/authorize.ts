// The connector's authorization UI — everything OAuthProvider does not own.
//
// Proving who is authorizing is done with Microsoft's *device code* flow rather
// than a redirect-based auth code flow. The Entra app registration behind this
// server is a public native client registered for device code and has no web
// redirect URI; adding one would mean changing the registration. Device code
// needs no redirect URI at all, so the Worker can verify the user's Microsoft
// identity with the app exactly as it is registered today.
//
// The device-code token is used for one thing only: calling Graph /me to learn
// who signed in, so it can be checked against the single-user allowlist. It is
// never stored. The mailbox refresh token in KV is a wholly separate credential.
import { AuthorizationError, type AuthRequest } from "@cloudflare/workers-oauth-provider";
import type { Env } from "./env.js";
import {
  DEVICE_CODE_ENDPOINT,
  IDENTITY_SCOPES,
  TOKEN_ENDPOINT,
  isAllowedIdentity,
  resolveIdentity,
} from "./ms-token.js";
import type { GrantProps } from "./mcp-handler.js";
import { NOTIFICATIONS_PATH, handleNotifications } from "./notifications.js";
import { VERSION } from "../core/version.js";

/** Pending device-code authorizations, keyed by an unguessable flow id. */
const FLOW_PREFIX = "flow:";
const FLOW_TTL_SECONDS = 900;

/** Scope granted to the connector when the client asks for nothing specific. */
const DEFAULT_SCOPE = ["outlook"];

type PendingFlow = {
  authRequest: AuthRequest;
  deviceCode: string;
};

function html(body: string, status = 200): Response {
  return new Response(body, {
    status,
    headers: { "content-type": "text/html; charset=utf-8" },
  });
}

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json" },
  });
}

function escapeHtml(value: string): string {
  return value.replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!
  );
}

function page(title: string, body: string): string {
  return `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
  :root { color-scheme: light dark; }
  body { font: 16px/1.5 system-ui, sans-serif; max-width: 34rem; margin: 4rem auto; padding: 0 1.25rem; }
  h1 { font-size: 1.35rem; }
  code { font-size: 1.6rem; letter-spacing: .12em; font-weight: 600; }
  .muted { opacity: .7; font-size: .9rem; }
  .err { color: #b00020; }
</style></head><body>${body}</body></html>`;
}

function errorPage(title: string, detail: string, status: number): Response {
  return html(
    page(title, `<h1>${escapeHtml(title)}</h1><p class="err">${escapeHtml(detail)}</p>`),
    status
  );
}

/** Redirect an OAuth error back to the client, or render it when we cannot. */
function authorizationErrorResponse(err: AuthorizationError): Response {
  if (!err.redirectUri) {
    return errorPage("Authorization request rejected", `${err.code}: ${err.description}`, 400);
  }
  const target = new URL(err.redirectUri);
  target.searchParams.set("error", err.code);
  target.searchParams.set("error_description", err.description);
  if (err.state) target.searchParams.set("state", err.state);
  if (err.issuer) target.searchParams.set("iss", err.issuer);
  return Response.redirect(target.toString(), 302);
}

async function startDeviceCode(env: Env): Promise<{
  device_code: string;
  user_code: string;
  verification_uri: string;
  message?: string;
}> {
  const response = await fetch(DEVICE_CODE_ENDPOINT, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: env.MS_CLIENT_ID, scope: IDENTITY_SCOPES }),
  });
  if (!response.ok) {
    throw new Error(`Microsoft refused to start a device code flow: HTTP ${response.status}`);
  }
  return response.json() as Promise<{
    device_code: string;
    user_code: string;
    verification_uri: string;
    message?: string;
  }>;
}

type DevicePoll =
  | { state: "pending" }
  | { state: "ready"; accessToken: string }
  | { state: "failed"; detail: string };

async function pollDeviceCode(env: Env, deviceCode: string): Promise<DevicePoll> {
  const response = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: deviceCode,
    }),
  });
  const payload = (await response.json().catch(() => ({}))) as {
    access_token?: string;
    error?: string;
    error_description?: string;
  };
  if (response.ok && payload.access_token) {
    return { state: "ready", accessToken: payload.access_token };
  }
  if (payload.error === "authorization_pending" || payload.error === "slow_down") {
    return { state: "pending" };
  }
  return { state: "failed", detail: payload.error_description ?? payload.error ?? "unknown error" };
}

/**
 * The single gate: turn a Microsoft access token into a completed authorization,
 * refusing any identity that is not the one this server was set up for.
 */
async function completeForIdentity(
  env: Env,
  authRequest: AuthRequest,
  msAccessToken: string
): Promise<{ ok: true; redirectTo: string } | { ok: false; status: number; detail: string }> {
  const identity = await resolveIdentity(msAccessToken);
  if (!identity) {
    return { ok: false, status: 401, detail: "That Microsoft token is not valid for Graph." };
  }
  if (!isAllowedIdentity(identity, env)) {
    return {
      ok: false,
      status: 403,
      detail:
        `${identity.userPrincipalName ?? identity.mail ?? identity.id} is not the account this ` +
        "connector belongs to. Only its owner can authorize it.",
    };
  }

  const props: GrantProps = {
    userId: identity.id,
    displayName: identity.displayName,
    upn: identity.userPrincipalName ?? identity.mail,
  };
  const { redirectTo } = await env.OAUTH_PROVIDER.completeAuthorization({
    request: authRequest,
    userId: identity.id,
    metadata: { authorizedAt: new Date().toISOString() },
    scope: authRequest.scope.length > 0 ? authRequest.scope : DEFAULT_SCOPE,
    props,
  });
  return { ok: true, redirectTo };
}

async function handleAuthorizeStart(request: Request, env: Env): Promise<Response> {
  let authRequest: AuthRequest;
  try {
    authRequest = await env.OAUTH_PROVIDER.parseAuthRequest(request);
  } catch (err) {
    if (err instanceof AuthorizationError) return authorizationErrorResponse(err);
    throw err;
  }

  const client = await env.OAUTH_PROVIDER.lookupClient(authRequest.clientId);
  const clientName = client?.clientName ?? authRequest.clientId;

  const device = await startDeviceCode(env);
  const flowId = crypto.randomUUID() + crypto.randomUUID().replace(/-/g, "");
  const pending: PendingFlow = { authRequest, deviceCode: device.device_code };
  await env.OUTLOOK_KV.put(FLOW_PREFIX + flowId, JSON.stringify(pending), {
    expirationTtl: FLOW_TTL_SECONDS,
  });

  const body = `
<h1>Connect ${escapeHtml(clientName)} to your Outlook</h1>
<p>Sign in as the owner of this mailbox to approve the connection. In another tab, open
<a href="${escapeHtml(device.verification_uri)}" target="_blank" rel="noopener">${escapeHtml(device.verification_uri)}</a>
and enter this code:</p>
<p><code id="code">${escapeHtml(device.user_code)}</code></p>
<p class="muted" id="status">Waiting for you to finish signing in…</p>
<p class="muted">This grants ${escapeHtml(clientName)} access to your mail, calendar, contacts and
tasks. Only this mailbox's own account can complete this step. The code expires in 15 minutes.</p>
<script>
const flow = ${JSON.stringify(flowId)};
const status = document.getElementById("status");
async function poll() {
  try {
    const res = await fetch("/authorize/poll?flow=" + encodeURIComponent(flow));
    const data = await res.json();
    if (data.status === "ok") { status.textContent = "Approved. Returning to the app…"; location.href = data.redirectTo; return; }
    if (data.status === "pending") { setTimeout(poll, 5000); return; }
    status.className = "err";
    status.textContent = data.detail || "Authorization failed.";
  } catch (e) {
    status.className = "err";
    status.textContent = "Lost contact with the server: " + e;
  }
}
setTimeout(poll, 5000);
</script>`;
  return html(page("Connect to Outlook", body));
}

async function handleAuthorizePoll(request: Request, env: Env): Promise<Response> {
  const flowId = new URL(request.url).searchParams.get("flow");
  if (!flowId) return json({ status: "failed", detail: "Missing flow id." }, 400);

  const raw = await env.OUTLOOK_KV.get(FLOW_PREFIX + flowId);
  if (!raw) {
    return json({ status: "failed", detail: "This sign-in expired. Reconnect to try again." }, 404);
  }
  const pending = JSON.parse(raw) as PendingFlow;

  const poll = await pollDeviceCode(env, pending.deviceCode);
  if (poll.state === "pending") return json({ status: "pending" });
  if (poll.state === "failed") {
    await env.OUTLOOK_KV.delete(FLOW_PREFIX + flowId);
    return json({ status: "failed", detail: poll.detail }, 400);
  }

  const result = await completeForIdentity(env, pending.authRequest, poll.accessToken);
  // One-shot: the device code has been redeemed either way.
  await env.OUTLOOK_KV.delete(FLOW_PREFIX + flowId);
  if (!result.ok) return json({ status: "failed", detail: result.detail }, result.status);
  return json({ status: "ok", redirectTo: result.redirectTo });
}

/**
 * Non-interactive authorization for local and test Workers only. The caller
 * supplies a Microsoft access token it already holds instead of completing the
 * device-code flow in a browser; the allowlist check is the identical one.
 *
 * Reachable only when ALLOW_DIRECT_AUTHORIZE=true (see env.ts), which
 * production deliberately leaves unset: even allowlist-gated, this path lets
 * anonymous callers feed arbitrary tokens to Graph through the Worker, and
 * would turn a leaked (short-lived) Microsoft access token into a persistent
 * grant without any interactive sign-in.
 */
async function handleAuthorizeDirect(request: Request, env: Env): Promise<Response> {
  let authRequest: AuthRequest;
  try {
    authRequest = await env.OAUTH_PROVIDER.parseAuthRequest(request);
  } catch (err) {
    if (err instanceof AuthorizationError) {
      return json({ status: "failed", detail: `${err.code}: ${err.description}` }, 400);
    }
    throw err;
  }

  const form = await request.formData().catch(() => null);
  const msAccessToken = form?.get("ms_access_token");
  if (typeof msAccessToken !== "string" || msAccessToken.length === 0) {
    return json({ status: "failed", detail: "Missing ms_access_token." }, 400);
  }

  const result = await completeForIdentity(env, authRequest, msAccessToken);
  if (!result.ok) return json({ status: "failed", detail: result.detail }, result.status);
  return json({ status: "ok", redirectTo: result.redirectTo });
}

export const defaultHandler = {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/authorize") {
      if (request.method === "GET") return handleAuthorizeStart(request, env);
      if (request.method === "POST") {
        // Refused before the request is even parsed unless this Worker is
        // explicitly a local/test one; the deployed Worker never sets the flag.
        if (env.ALLOW_DIRECT_AUTHORIZE !== "true") {
          return json(
            {
              status: "failed",
              detail:
                "Direct authorization is disabled on this server. " +
                "Use the interactive sign-in at GET /authorize.",
            },
            403
          );
        }
        return handleAuthorizeDirect(request, env);
      }
      return new Response("Method not allowed", { status: 405 });
    }
    if (url.pathname === "/authorize/poll" && request.method === "GET") {
      return handleAuthorizePoll(request, env);
    }
    if (url.pathname === NOTIFICATIONS_PATH) {
      // Microsoft Graph posts here with no credential of its own; the endpoint
      // authenticates deliveries with the subscription's clientState instead.
      // ctx carries the waitUntil the follow-up work is scheduled on.
      return handleNotifications(request, env, ctx);
    }
    if (url.pathname === "/health") {
      // Deliberately says nothing about the mailbox or whether anyone is
      // authorized. The version is the one thing worth serving anonymously: it
      // is how a headless check tells whether the deployed build is the one in
      // the working tree, and so whether its tools/list carries this registry's
      // annotations.
      return json({ status: "ok", service: "outlook-mcp", version: VERSION });
    }
    if (url.pathname === "/") {
      return html(
        page(
          "outlook-mcp",
          `<h1>outlook-mcp</h1><p>A private Model Context Protocol server for a single Outlook
mailbox. The MCP endpoint is <code style="font-size:1rem">/mcp</code> and requires OAuth; only
this mailbox's owner can authorize a client.</p>`
        )
      );
    }
    return new Response("Not found", { status: 404 });
  },
};
