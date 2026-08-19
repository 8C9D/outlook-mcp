// Live tests against the deployed Cloudflare Worker.
//
// Everything here runs from the local machine against the real workers.dev
// endpoint and the real Microsoft tenant: discovery metadata, refusal of
// anonymous callers, a full OAuth authorization-code exchange, an MCP
// round-trip over Streamable HTTP, proof that the mailbox refresh token in KV
// actually rotates, and the things only the hosted server can do — KV-backed
// delta positions, MCP resources, a Graph change notification travelling from a
// real send all the way into get_mailbox_activity, and attachments and MIME
// exports moving in and out of a mailbox on a server with no filesystem.
//
// Like the local harness it cleans up after itself — the OAuth client it
// registers and every grant and token derived from it are deleted from KV, the
// probe mail is purged, the notifications it caused are removed from the ring
// buffer and the delta position is restored — and the final sweep asserts
// nothing is left. The mail subscription itself is deliberately NOT torn down:
// it is the production subscription for the real inbox, not a test artifact.
//
// The deployed Worker refuses the non-interactive POST-with-ms_access_token
// authorize path (r5 asserts exactly that), so the bearer that the
// authenticated tests need can only come from a real device-code sign-in.
// When run in a terminal (or with MCP_REMOTE_INTERACTIVE=1) r6 drives the
// production /authorize page and waits for the owner to enter the code at
// microsoft.com/devicelogin; in a headless run (no TTY, or
// MCP_REMOTE_HEADLESS=1) the bearer-dependent tests are reported as SKIP and
// everything unauthenticated still runs.
//
// Prerequisites: `npm run deploy`, the three wrangler secrets, and `npm run
// seed:kv`. Run with `npm run test:remote`.
import { spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import os from "node:os";
import { createHash, randomBytes } from "node:crypto";
import path from "node:path";
import { PROJECT_ROOT } from "./project-root.js";
import { installMsalTokenProvider } from "./auth.js";
import { callGraphServer, callGraphServerBytes } from "./core/graph.js";
import {
  KV_ACCESS_TOKEN,
  KV_REFRESH_TOKEN,
  STATE_ACTIVITY,
  STATE_LLM_AUDIT,
  STATE_LLM_CONFIG,
  STATE_SUBSCRIPTION,
  deltaKey,
  downloadKey,
} from "./core/kv-keys.js";
import {
  NEVER_FILE_INTO,
  readLlmConfig,
  type AuditEntry,
  type LlmConfig,
} from "./core/auto-filing.js";
import { createMemoryStateStore } from "./core/state.js";
import { DOWNLOAD_ROUTE_PREFIX } from "./core/downloads.js";
import type { ActivityEntry } from "./core/notifications.js";
import { TOOLS } from "./core/registry.js";
import { FOLDERS_URI, RECENT_INBOX_URI } from "./core/resources.js";
import { SUBSCRIPTION_RESOURCE, type SubscriptionRecord } from "./core/subscriptions.js";
import { VERSION } from "./core/version.js";

const BASE_URL = process.env.MCP_REMOTE_URL ?? "https://outlook-mcp.arthur-yuhao-zhang.workers.dev";
const MCP_URL = `${BASE_URL}/mcp`;
const TEST_CLIENT_NAME = "[MCP TEST] remote harness";

type Outcome = { name: string; passed: boolean; skipped?: boolean; detail?: string };
const outcomes: Outcome[] = [];

async function test(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    outcomes.push({ name, passed: true });
    console.log(`PASS  ${name}`);
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    outcomes.push({ name, passed: false, detail });
    console.log(`FAIL  ${name}\n      ${detail.split("\n").join("\n      ")}`);
  }
}

function skip(name: string, why: string): void {
  outcomes.push({ name, passed: true, skipped: true, detail: why });
  console.log(`SKIP  ${name}\n      ${why}`);
}

// Production only authorizes through the interactive device-code flow, so the
// authenticated tests need a human to enter a code at microsoft.com/devicelogin.
// A TTY implies one is present; MCP_REMOTE_INTERACTIVE / MCP_REMOTE_HEADLESS
// override the guess in either direction.
const interactive =
  process.env.MCP_REMOTE_INTERACTIVE === "1" ||
  (process.stdout.isTTY === true && process.env.MCP_REMOTE_HEADLESS !== "1");
const HEADLESS_SKIP =
  "needs a bearer, and only an interactive device-code sign-in can mint one against " +
  "production (the direct authorize path is disabled there) — rerun in a terminal or " +
  "with MCP_REMOTE_INTERACTIVE=1";

/** A test that needs the bearer from r6: reported as SKIP in headless runs. */
async function testAuthed(name: string, fn: () => Promise<void>): Promise<void> {
  if (!interactive) return skip(name, HEADLESS_SKIP);
  return test(name, fn);
}

function assert(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

// ---------------------------------------------------------------- wrangler KV

function wrangler(args: string[]): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn("npx", ["wrangler", ...args], { cwd: PROJECT_ROOT });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (d) => (stdout += d));
    child.stderr.on("data", (d) => (stderr += d));
    child.on("error", reject);
    child.on("exit", (code) => resolve({ code: code ?? -1, stdout, stderr }));
  });
}

async function namespaceId(binding: string): Promise<string> {
  const raw = await fs.readFile(path.join(PROJECT_ROOT, "wrangler.jsonc"), "utf8");
  const config = JSON.parse(raw.replace(/^\s*\/\/.*$/gm, "")) as {
    kv_namespaces?: { binding: string; id: string }[];
  };
  const entry = config.kv_namespaces?.find((ns) => ns.binding === binding);
  assert(entry, `wrangler.jsonc has no ${binding} namespace`);
  return entry.id;
}

async function kvGet(nsId: string, key: string): Promise<string | null> {
  const { code, stdout } = await wrangler([
    "kv",
    "key",
    "get",
    key,
    "--namespace-id",
    nsId,
    "--remote",
    "--text",
  ]);
  if (code !== 0) return null;
  return stdout;
}

async function kvDelete(nsId: string, key: string): Promise<void> {
  await wrangler(["kv", "key", "delete", key, "--namespace-id", nsId, "--remote"]);
}

/** Write a KV value through a 0600 temp file rather than argv, as seed:kv does. */
async function kvPut(nsId: string, key: string, value: string): Promise<void> {
  const tmpFile = path.join(os.tmpdir(), `outlook-mcp-remote-${randomBytes(8).toString("hex")}`);
  await fs.writeFile(tmpFile, value, { mode: 0o600 });
  try {
    const { code, stderr } = await wrangler([
      "kv",
      "key",
      "put",
      key,
      "--path",
      tmpFile,
      "--namespace-id",
      nsId,
      "--remote",
    ]);
    assert(code === 0, `wrangler kv key put ${key} failed: ${stderr}`);
  } finally {
    await fs.rm(tmpFile, { force: true });
  }
}

async function kvListKeys(nsId: string): Promise<string[]> {
  const { code, stdout, stderr } = await wrangler([
    "kv",
    "key",
    "list",
    "--namespace-id",
    nsId,
    "--remote",
  ]);
  assert(code === 0, `wrangler kv key list failed: ${stderr}`);
  const start = stdout.indexOf("[");
  assert(start >= 0, `unexpected kv key list output: ${stdout.slice(0, 200)}`);
  return (JSON.parse(stdout.slice(start)) as { name: string }[]).map((k) => k.name);
}

/** A stable, non-revealing identifier for a secret value. */
function fingerprint(secret: string): string {
  return createHash("sha256").update(secret).digest("hex").slice(0, 16);
}

// -------------------------------------------------------------------- helpers

function base64url(bytes: Buffer): string {
  return bytes.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

type McpResponse = { result?: any; error?: { code: number; message: string } };

async function mcpCall(
  bearer: string | null,
  method: string,
  params?: unknown,
  id = 1
): Promise<{ status: number; headers: Headers; body: McpResponse | null }> {
  const headers: Record<string, string> = {
    "content-type": "application/json",
    accept: "application/json, text/event-stream",
  };
  if (bearer) headers.authorization = `Bearer ${bearer}`;
  const response = await fetch(MCP_URL, {
    method: "POST",
    headers,
    body: JSON.stringify({ jsonrpc: "2.0", id, method, ...(params ? { params } : {}) }),
  });
  const text = await response.text();
  let body: McpResponse | null = null;
  try {
    body = JSON.parse(text);
  } catch {
    body = null;
  }
  return { status: response.status, headers: response.headers, body };
}

function resultOf(
  response: { status: number; body: McpResponse | null },
  context: string
): any {
  assert(response.status === 200, `${context}: HTTP ${response.status}`);
  assert(response.body, `${context}: response was not JSON`);
  assert(!response.body.error, `${context}: ${JSON.stringify(response.body.error)}`);
  return response.body.result;
}

/** Call a tool over the remote endpoint, keeping whatever it answered. */
async function callToolRaw(
  name: string,
  args: Record<string, unknown> = {}
): Promise<{ text: string; isError: boolean }> {
  assert(bearer, `no bearer token (r6 must pass first) — cannot call ${name}`);
  const result = resultOf(
    await mcpCall(bearer!, "tools/call", { name, arguments: args }),
    `tools/call ${name}`
  );
  return {
    text: (result.content as { text: string }[]).map((part) => part.text).join("\n"),
    isError: result.isError === true,
  };
}

/** Call a tool over the remote endpoint and return its text, refusing isError. */
async function callTool(name: string, args: Record<string, unknown> = {}): Promise<string> {
  const { text, isError } = await callToolRaw(name, args);
  assert(!isError, `${name} returned isError: ${text}`);
  return text;
}

/** Retry `fn` until it returns a value or the deadline passes. */
async function poll<T>(
  what: string,
  timeoutMs: number,
  intervalMs: number,
  fn: () => Promise<T | undefined>
): Promise<T> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const value = await fn();
    if (value !== undefined) return value;
    if (Date.now() > deadline) {
      throw new Error(`Timed out after ${Math.round(timeoutMs / 1000)}s waiting for ${what}`);
    }
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}

/** MCP requires an initialize before anything useful; stateless mode needs it per POST. */
const INIT_PARAMS = {
  protocolVersion: "2025-06-18",
  capabilities: {},
  clientInfo: { name: "outlook-mcp-remote-test", version: "1" },
};

// ------------------------------------------------------------------ the tests

installMsalTokenProvider();

console.log(`Remote endpoint: ${MCP_URL}\n`);

const oauthNs = await namespaceId("OAUTH_KV");
const outlookNs = await namespaceId("OUTLOOK_KV");

// Used to send the change-notification probe to this mailbox, and to clean up
// afterwards with the local Microsoft token rather than through the connector.
const me = await callGraphServer("/me?$select=mail,userPrincipalName");
const ownAddress: string = me.mail ?? me.userPrincipalName;

let registeredClientId: string | undefined;
let bearer: string | undefined;

// Everything this run adds to the OAuth store must be gone by the end, so the
// teardown compares against what was there before rather than guessing at the
// provider's key naming.
const baselineOauthKeys = new Set(await kvListKeys(oauthNs));

await test("r1. protected resource metadata (RFC 9728) advertises this exact endpoint", async () => {
  const response = await fetch(`${BASE_URL}/.well-known/oauth-protected-resource/mcp`);
  assert(response.ok, `HTTP ${response.status}`);
  const metadata = (await response.json()) as any;
  assert(
    metadata.resource === MCP_URL,
    `resource is ${metadata.resource}, expected ${MCP_URL} — claude.ai requires an exact match`
  );
  assert(
    metadata.authorization_servers?.[0] === BASE_URL,
    `authorization_servers[0] is ${metadata.authorization_servers?.[0]}`
  );
});

await test("r2. authorization server metadata offers DCR and S256 PKCE", async () => {
  const response = await fetch(`${BASE_URL}/.well-known/oauth-authorization-server`);
  assert(response.ok, `HTTP ${response.status}`);
  const metadata = (await response.json()) as any;
  assert(metadata.issuer === BASE_URL, `issuer is ${metadata.issuer}`);
  assert(
    typeof metadata.registration_endpoint === "string",
    "no registration_endpoint: claude.ai could not register itself"
  );
  assert(
    metadata.code_challenge_methods_supported?.includes("S256"),
    "S256 PKCE is not advertised"
  );
  assert(
    metadata.grant_types_supported?.includes("authorization_code"),
    "authorization_code grant is not advertised"
  );
});

await test("r3. anonymous and bogus-bearer requests to the protected routes are refused", async () => {
  const anonymous = await mcpCall(null, "tools/list");
  assert(anonymous.status === 401, `anonymous POST returned ${anonymous.status}, expected 401`);
  const challenge = anonymous.headers.get("www-authenticate") ?? "";
  assert(
    challenge.includes("resource_metadata="),
    `401 lacks a resource_metadata challenge: ${challenge}`
  );

  const bogus = await mcpCall("not-a-real-token", "tools/list");
  assert(bogus.status === 401, `bogus bearer returned ${bogus.status}, expected 401`);

  // The GET (SSE) and DELETE verbs must be gated too, not just POST.
  for (const method of ["GET", "DELETE"]) {
    const response = await fetch(MCP_URL, { method });
    assert(
      response.status === 401,
      `anonymous ${method} /mcp returned ${response.status}, expected 401`
    );
  }

  // The attachment download route is gated by the same bearer check. A
  // well-formed but unissued id must be refused for want of a token, not
  // answered with 404 — that ordering is what keeps the links private.
  const madeUpId = randomBytes(32).toString("hex");
  for (const bearerValue of [null, "not-a-real-token"]) {
    const response = await fetch(`${BASE_URL}${DOWNLOAD_ROUTE_PREFIX}${madeUpId}`, {
      headers: bearerValue ? { authorization: `Bearer ${bearerValue}` } : {},
    });
    assert(
      response.status === 401,
      `${bearerValue ? "bogus-bearer" : "anonymous"} GET of a download link returned ` +
        `${response.status}, expected 401`
    );
  }
});

await test("r4. dynamic client registration (RFC 7591) issues a client_id", async () => {
  const response = await fetch(`${BASE_URL}/oauth/register`, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      client_name: TEST_CLIENT_NAME,
      redirect_uris: ["http://localhost:8976/callback"],
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      token_endpoint_auth_method: "none",
    }),
  });
  const registrationBody = await response.text();
  assert(response.ok, `HTTP ${response.status}: ${registrationBody}`);
  const client = JSON.parse(registrationBody) as any;
  assert(typeof client.client_id === "string", "registration returned no client_id");
  registeredClientId = client.client_id;
});

/** Build the /authorize URL for the registered test client. */
function authorizeUrl(codeChallenge: string, state: string): string {
  const url = new URL(`${BASE_URL}/authorize`);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("client_id", registeredClientId!);
  url.searchParams.set("redirect_uri", "http://localhost:8976/callback");
  url.searchParams.set("scope", "outlook");
  url.searchParams.set("state", state);
  url.searchParams.set("code_challenge", codeChallenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("resource", MCP_URL);
  return url.toString();
}

await test("r5. the direct POST-with-token authorize path is disabled in production", async () => {
  assert(registeredClientId, "no registered client (r4 must pass first)");
  const verifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(verifier).digest());
  // The value of the token must never matter: the deployed Worker has
  // ALLOW_DIRECT_AUTHORIZE unset, so the request is refused before the token
  // is looked at (a live token would be refused identically).
  const response = await fetch(authorizeUrl(challenge, "state-direct"), {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ ms_access_token: "any-value-must-not-matter" }),
  });
  assert(
    response.status === 403,
    `POST /authorize with an ms_access_token returned ${response.status}, expected 403 — ` +
      "the direct authorization path must be disabled on the deployed Worker"
  );
  const body = (await response.json()) as any;
  assert(body.status === "failed", `expected status "failed", got ${JSON.stringify(body)}`);
  assert(
    /disabled/i.test(body.detail ?? ""),
    `the refusal must say the path is disabled, not that the token is bad: ${JSON.stringify(body)}`
  );
});

await testAuthed("r6. an interactive device-code sign-in completes authorization and yields a bearer", async () => {
  assert(registeredClientId, "no registered client (r4 must pass first)");
  const codeVerifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(codeVerifier).digest());
  const state = "state-" + randomBytes(6).toString("hex");

  // GET /authorize starts a real device-code flow and renders the user code;
  // pull the code, the verification URL and the flow id out of the page and
  // ask the human running this suite to complete the sign-in.
  const pageResponse = await fetch(authorizeUrl(challenge, state));
  const pageBody = await pageResponse.text();
  assert(pageResponse.ok, `GET /authorize failed: HTTP ${pageResponse.status} ${pageBody.slice(0, 300)}`);
  const userCode = pageBody.match(/<code id="code">([^<]+)<\/code>/)?.[1];
  const verificationUri = pageBody.match(/<a href="([^"]+)" target="_blank"/)?.[1];
  const flowId = pageBody.match(/const flow = "([^"]+)"/)?.[1];
  assert(userCode && verificationUri && flowId, `could not parse the /authorize page: ${pageBody.slice(0, 500)}`);

  // How long to wait for the human sign-in. 10 minutes by default; a run that
  // has to route around throttled verification emails can extend it via
  // MCP_REMOTE_SIGNIN_TIMEOUT_MS without touching the flow being tested.
  const signinTimeoutMs = Number(process.env.MCP_REMOTE_SIGNIN_TIMEOUT_MS) || 600_000;

  console.log(`\n      ACTION REQUIRED: open ${verificationUri} and enter the code ${userCode},`);
  console.log(`      signing in as the mailbox owner. Waiting up to ${Math.round(signinTimeoutMs / 60_000)} minutes...\n`);

  // Poll exactly as the page's own script does, until the sign-in lands.
  const deadline = Date.now() + signinTimeoutMs;
  let redirectTo: string | undefined;
  for (;;) {
    const pollResponse = await fetch(
      `${BASE_URL}/authorize/poll?flow=${encodeURIComponent(flowId)}`
    );
    const poll = (await pollResponse.json()) as any;
    if (poll.status === "ok") {
      redirectTo = poll.redirectTo;
      break;
    }
    assert(poll.status === "pending", `authorization failed: ${JSON.stringify(poll)}`);
    assert(Date.now() < deadline, "timed out waiting for the device-code sign-in");
    await new Promise((resolve) => setTimeout(resolve, 5_000));
  }

  const redirect = new URL(redirectTo!);
  assert(redirect.searchParams.get("state") === state, "authorization returned the wrong state");
  const code = redirect.searchParams.get("code");
  assert(code, `no authorization code in ${redirectTo}`);

  const tokenResponse = await fetch(`${BASE_URL}/oauth/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: "http://localhost:8976/callback",
      client_id: registeredClientId!,
      code_verifier: codeVerifier,
    }),
  });
  const tokenBody = await tokenResponse.text();
  assert(
    tokenResponse.ok,
    `token exchange failed: HTTP ${tokenResponse.status} ${tokenBody}`
  );
  const token = JSON.parse(tokenBody) as any;
  assert(typeof token.access_token === "string", "token endpoint returned no access_token");
  bearer = token.access_token;
});

await test("r7. token exchange rejects a wrong PKCE verifier", async () => {
  assert(registeredClientId, "no registered client");
  const response = await fetch(`${BASE_URL}/oauth/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      code: "made-up-code",
      redirect_uri: "http://localhost:8976/callback",
      client_id: registeredClientId!,
      code_verifier: base64url(randomBytes(32)),
    }),
  });
  assert(response.status >= 400, `a bogus code was accepted with HTTP ${response.status}`);
});

await testAuthed("r8. initialize over Streamable HTTP reports this server", async () => {
  assert(bearer, "no bearer token (r6 must pass first)");
  const response = await mcpCall(bearer!, "initialize", INIT_PARAMS);
  const result = resultOf(response, "initialize");
  assert(result.serverInfo?.name === "outlook", `serverInfo.name is ${result.serverInfo?.name}`);
  assert(result.protocolVersion, "no protocolVersion in the initialize result");
  // Stateless: the server must not hand out a session it cannot remember.
  assert(
    !response.headers.get("mcp-session-id"),
    "the stateless endpoint returned an Mcp-Session-Id"
  );
});

await testAuthed(
  `r9. tools/list serves all ${TOOLS.length} tools, with the same annotations as the stdio surface`,
  async () => {
    assert(bearer, "no bearer token");
    const result = resultOf(await mcpCall(bearer!, "tools/list"), "tools/list");
    const remote = result.tools as { name: string; annotations?: Record<string, boolean> }[];
    const remoteNames = remote.map((t) => t.name).sort();
    const localNames = TOOLS.map((t) => t.name).sort();
    assert(
      remoteNames.length === localNames.length,
      `remote has ${remoteNames.length} tools, the registry has ${localNames.length}`
    );
    assert(
      remoteNames.join(",") === localNames.join(","),
      `tool surfaces differ:\n  remote: ${remoteNames.join(", ")}\n  local:  ${localNames.join(", ")}`
    );

    // The annotation hints have to survive the second transport too: a client
    // that trusts them on stdio must see the same ones over HTTP.
    const hints = (a: Record<string, boolean> | undefined) =>
      [a?.readOnlyHint, a?.destructiveHint, a?.idempotentHint, a?.openWorldHint].join(",");
    for (const tool of TOOLS) {
      const served = remote.find((t) => t.name === tool.name)!;
      assert(
        hints(served.annotations) === hints(tool.annotations),
        `${tool.name} is annotated [${hints(served.annotations)}] over HTTP, ` +
          `[${hints(tool.annotations)}] in the registry ` +
          "(readOnly, destructive, idempotent, openWorld)"
      );
    }
  }
);

await testAuthed("r10. prompts/list serves both prompts", async () => {
  assert(bearer, "no bearer token");
  const result = resultOf(await mcpCall(bearer!, "prompts/list"), "prompts/list");
  const names = (result.prompts as { name: string }[]).map((p) => p.name).sort();
  assert(
    names.join(",") === "morning_brief,triage_inbox",
    `prompts are ${names.join(", ")}`
  );
});

await testAuthed("r11. tools/call list_events reaches Graph through the KV token", async () => {
  assert(bearer, "no bearer token");
  const result = resultOf(
    await mcpCall(bearer!, "tools/call", { name: "list_events", arguments: { days: 3 } }),
    "tools/call list_events"
  );
  const text = (result.content as { text: string }[]).map((c) => c.text).join("\n");
  assert(result.isError !== true, `list_events returned isError: ${text}`);
  assert(
    /America\/Toronto|No events in this window/.test(text),
    `unexpected list_events output: ${text.slice(0, 300)}`
  );
});

await testAuthed("r12. the mailbox refresh token in KV rotates on every exchange", async () => {
  assert(bearer, "no bearer token");
  const before = await kvGet(outlookNs, KV_REFRESH_TOKEN);
  assert(before, `${KV_REFRESH_TOKEN} is missing from KV — run \`npm run seed:kv\``);

  // Drop the cached access token so the next Graph call must spend the refresh
  // token, which is what makes Microsoft rotate it.
  await kvDelete(outlookNs, KV_ACCESS_TOKEN);
  const result = resultOf(
    await mcpCall(bearer!, "tools/call", { name: "list_events", arguments: { days: 1 } }),
    "tools/call forcing a refresh"
  );
  assert(result.isError !== true, "the refreshing call failed");

  // KV is eventually consistent; give the rotated value a moment to land.
  let after: string | null = null;
  for (let attempt = 0; attempt < 10; attempt++) {
    after = await kvGet(outlookNs, KV_REFRESH_TOKEN);
    if (after && after !== before) break;
    await new Promise((resolve) => setTimeout(resolve, 3000));
  }
  assert(
    after && after !== before,
    `the refresh token did not rotate (still sha256:${fingerprint(before)}) — ` +
      "the Worker would lock itself out once Microsoft expires it"
  );
  console.log(
    `      rotated sha256:${fingerprint(before)} -> sha256:${fingerprint(after!)}`
  );

  // And the rotated token must itself work, or the chain is broken.
  await kvDelete(outlookNs, KV_ACCESS_TOKEN);
  const again = resultOf(
    await mcpCall(bearer!, "tools/call", { name: "list_events", arguments: { days: 1 } }),
    "tools/call after rotation"
  );
  assert(again.isError !== true, "the rotated refresh token could not mint an access token");
});

// ------------------------------------------- v6: notifications, delta, resources

const TEST_PREFIX = "[MCP TEST]";
const notificationsUrl = `${BASE_URL}/notifications`;

/** What the run has to put back or clean out afterwards. */
const v6: {
  deltaBefore?: string | null;
  activityBefore?: ActivityEntry[];
  probeSubject?: string;
} = {};

await test("r13. the notification endpoint is public, echoes the handshake, and enforces clientState", async () => {
  // Graph presents no credential, so this route must answer unauthenticated —
  // and must still refuse to record anything without the shared secret.
  const token = `validation-${randomBytes(8).toString("hex")}`;
  const handshake = await fetch(`${notificationsUrl}?validationToken=${encodeURIComponent(token)}`, {
    method: "POST",
  });
  assert(handshake.status === 200, `handshake returned HTTP ${handshake.status}, expected 200`);
  assert(
    (handshake.headers.get("content-type") ?? "").startsWith("text/plain"),
    `handshake content-type is ${handshake.headers.get("content-type")}`
  );
  assert((await handshake.text()) === token, "the deployed endpoint did not echo the token");

  const forged = await fetch(notificationsUrl, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      value: [
        {
          subscriptionId: "forged",
          clientState: "not-the-secret",
          changeType: "created",
          resourceData: { id: "forged-message" },
        },
      ],
    }),
  });
  assert(forged.status === 202, `forged delivery returned HTTP ${forged.status}, expected 202`);
  const outcome = (await forged.json()) as { accepted: number; discarded: number };
  assert(
    outcome.accepted === 0 && outcome.discarded === 1,
    `a forged delivery was not discarded: ${JSON.stringify(outcome)}`
  );
});

await testAuthed("r14. resources/list and resources/read serve the same two resources as stdio", async () => {
  assert(bearer, "no bearer token");
  const result = resultOf(await mcpCall(bearer!, "resources/list"), "resources/list");
  const uris = (result.resources as { uri: string }[]).map((r) => r.uri).sort();
  assert(
    uris.join(",") === [FOLDERS_URI, RECENT_INBOX_URI].sort().join(","),
    `remote resources are ${uris.join(", ") || "(none)"}`
  );

  const read = resultOf(
    await mcpCall(bearer!, "resources/read", { uri: RECENT_INBOX_URI }),
    "resources/read"
  );
  const contents = read.contents as { uri: string; text?: string }[];
  assert(contents?.length === 1, `resources/read returned ${contents?.length} contents`);
  assert(contents[0]!.uri === RECENT_INBOX_URI, `read back ${contents[0]!.uri}`);
  assert(
    /latest message\(s\) in inbox|No messages in inbox/.test(contents[0]!.text ?? ""),
    `unexpected resource body: ${(contents[0]!.text ?? "").slice(0, 200)}`
  );
});

await testAuthed("r15. check_new_mail keeps its delta position in KV between requests", async () => {
  // The stateless Worker builds a new server per POST, so the only thing that
  // can carry a delta position between calls is KV.
  v6.deltaBefore = await kvGet(outlookNs, deltaKey("inbox"));

  const baseline = await callTool("check_new_mail", { folder: "inbox", reset: true });
  assert(
    /Starting position recorded for inbox/.test(baseline),
    `unexpected baseline output: ${baseline}`
  );

  const stored = await poll("the delta position to appear in KV", 60_000, 3_000, async () => {
    const raw = await kvGet(outlookNs, deltaKey("inbox"));
    return raw && raw !== v6.deltaBefore ? raw : undefined;
  });
  assert(
    /\$deltatoken=/.test(stored),
    `the stored position does not look like a Graph delta link: ${stored.slice(0, 120)}`
  );

  const second = await callTool("check_new_mail", {});
  assert(
    /No changes in inbox since|new or changed message/.test(second),
    `a second call did not resume from the stored position: ${second}`
  );
  assert(
    !/Starting position recorded/.test(second),
    `the position was lost between requests: ${second}`
  );
});

await test("r16. a mail subscription exists, points here, and expires in the future", async () => {
  const record = await poll(
    "the Worker to create the inbox subscription",
    120_000,
    5_000,
    async () => {
      const raw = await kvGet(outlookNs, STATE_SUBSCRIPTION);
      if (!raw) return undefined;
      try {
        return JSON.parse(raw) as SubscriptionRecord;
      } catch {
        return undefined;
      }
    }
  );
  assert(record.id, `subscription record has no id: ${JSON.stringify(record)}`);
  assert(
    record.notificationUrl === notificationsUrl,
    `subscription notifies ${record.notificationUrl}, expected ${notificationsUrl}`
  );
  assert(
    record.resource === SUBSCRIPTION_RESOURCE,
    `subscription watches ${record.resource}, expected ${SUBSCRIPTION_RESOURCE}`
  );
  assert(
    record.clientState?.length >= 32,
    "the stored clientState is not a long random secret"
  );

  const remaining = (Date.parse(record.expirationDateTime) - Date.now()) / 60000;
  assert(remaining > 0, `the stored subscription already expired at ${record.expirationDateTime}`);
  assert(remaining <= 4230, `expiry is ${Math.round(remaining)} min away, beyond Graph's maximum`);

  // Graph's own view has to agree; a record for a subscription Graph forgot is
  // exactly the failure the cron exists to repair.
  const live = await callGraphServer(`/subscriptions/${encodeURIComponent(record.id)}`);
  assert(live.id === record.id, `Graph returned subscription ${live.id}`);
  assert(
    live.notificationUrl === notificationsUrl,
    `Graph has notificationUrl ${live.notificationUrl}`
  );
  assert(
    Date.parse(live.expirationDateTime) > Date.now(),
    `Graph says the subscription expired at ${live.expirationDateTime}`
  );
  console.log(
    `      subscription ${record.id} expires ${live.expirationDateTime} ` +
      `(${Math.round(remaining / 60)}h away); cron: see wrangler.jsonc triggers`
  );

  // The cron that keeps it alive must actually be configured.
  const config = JSON.parse(
    (await fs.readFile(path.join(PROJECT_ROOT, "wrangler.jsonc"), "utf8")).replace(
      /^\s*\/\/.*$/gm,
      ""
    )
  ) as { triggers?: { crons?: string[] } };
  assert(
    (config.triggers?.crons ?? []).length > 0,
    "wrangler.jsonc declares no cron trigger, so nothing would renew this subscription"
  );
});

await testAuthed("r17. end-to-end: a message sent to this mailbox shows up in get_mailbox_activity", async () => {
  v6.activityBefore = JSON.parse((await kvGet(outlookNs, STATE_ACTIVITY)) || "[]") as ActivityEntry[];

  const subject = `${TEST_PREFIX} v6 webhook ${randomBytes(4).toString("hex")}`;
  v6.probeSubject = subject;
  const createText = await callTool("create_draft", {
    to: [ownAddress],
    subject,
    body: "Change-notification probe from the remote harness. Safe to delete.",
  });
  const draftId = createText.match(/Draft id: (\S+)/)?.[1];
  assert(draftId, `could not extract a draft id from: ${createText}`);
  await callTool("send_draft", { draft_id: draftId });

  // Graph delivers notifications asynchronously and sometimes slowly; this is a
  // bounded wait with a diagnostic, not an indefinite one.
  let lastSeen = "";
  try {
    await poll("Graph to deliver the change notification", 240_000, 10_000, async () => {
      lastSeen = await callTool("get_mailbox_activity", { since_hours: 1, limit: 25 });
      return lastSeen.includes(subject) ? true : undefined;
    });
  } catch (err) {
    const arrived = await callGraphServer(
      `/me/mailFolders/inbox/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$select=id,receivedDateTime`
    );
    throw new Error(
      `${err instanceof Error ? err.message : String(err)}\n` +
        `The message itself ${arrived?.value?.length ? "did" : "did NOT"} reach the inbox, so the ` +
        `gap is ${arrived?.value?.length ? "in Graph's delivery to /notifications" : "in sending"}.\n` +
        `Last get_mailbox_activity output:\n${lastSeen}`
    );
  }

  const line = lastSeen.split("\n").find((entry) => entry.includes(subject));
  assert(line, "the activity output changed between the poll and the assertion");
  assert(
    /Notified: \d{4}-\d{2}-\d{2} \d{2}:\d{2}/.test(lastSeen),
    `activity entries carry no notification timestamp:\n${lastSeen}`
  );
  assert(
    /Message id: \S+/.test(lastSeen),
    `activity entries carry no message id:\n${lastSeen}`
  );
});

// ------------------------------------------- v7: attachments without a filesystem

// A stable, public https resource for the url source. The Worker fetches it
// itself, which is the whole point: the bytes never touch this machine.
const ATTACH_URL = "https://example.com/";

await testAuthed("r18. add_attachment fetches an https URL from the Worker; file_path is refused", async () => {
  const subject = `${TEST_PREFIX} v7 url attach ${randomBytes(4).toString("hex")}`;
  const createText = await callTool("create_draft", {
    to: [ownAddress],
    subject,
    body: "URL-attachment probe from the remote harness. Safe to delete.",
  });
  const draftId = createText.match(/Draft id: (\S+)/)?.[1];
  assert(draftId, `could not extract a draft id from: ${createText}`);

  const attachText = await callTool("add_attachment", {
    draft_id: draftId,
    url: ATTACH_URL,
    attachment_name: "probe.html",
  });
  assert(/Name: probe\.html/.test(attachText), `unexpected add_attachment output: ${attachText}`);
  assert(
    /Type: text\/html/.test(attachText),
    `the content type did not come from the response: ${attachText}`
  );

  // Graph's own view, read with this machine's token rather than the connector's.
  const listed = await callGraphServer(
    `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name,size,contentType`
  );
  const attachment = (listed?.value ?? []).find((a: any) => a.name === "probe.html");
  assert(attachment, `the attachment is not in the inventory: ${JSON.stringify(listed?.value)}`);
  assert(attachment.size > 500, `the attachment is suspiciously small: ${attachment.size} bytes`);
  const full = await callGraphServer(
    `/me/messages/${encodeURIComponent(draftId)}/attachments/${encodeURIComponent(attachment.id)}`
  );
  assert(
    Buffer.from(full.contentBytes ?? "", "base64").toString("utf8").includes("Example Domain"),
    "the attached bytes are not what the URL serves"
  );

  // The local-only source has to fail honestly rather than half-work.
  const refused = await callToolRaw("add_attachment", {
    draft_id: draftId,
    file_path: "/etc/hosts",
  });
  assert(refused.isError, `file_path was accepted by the hosted server: ${refused.text}`);
  assert(
    /no access to your filesystem|only on the local/i.test(refused.text),
    `unexpected file_path refusal: ${refused.text}`
  );

  await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}/permanentDelete`, {
    method: "POST",
  });
});

await testAuthed("r19. get_attachment hands out a bearer-gated download link that expires", async () => {
  const payload = Buffer.alloc(2048);
  for (let i = 0; i < payload.length; i++) payload[i] = (i * 31) % 256;

  const subject = `${TEST_PREFIX} v7 download ${randomBytes(4).toString("hex")}`;
  const createText = await callTool("create_draft", {
    to: [ownAddress],
    subject,
    body: "Download-link probe from the remote harness. Safe to delete.",
  });
  const draftId = createText.match(/Draft id: (\S+)/)?.[1];
  assert(draftId, `could not extract a draft id from: ${createText}`);

  try {
    await callTool("add_attachment", {
      draft_id: draftId,
      content_base64: payload.toString("base64"),
      attachment_name: "blob.bin",
    });
    const listed = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name`
    );
    const attachment = (listed?.value ?? []).find((a: any) => a.name === "blob.bin");
    assert(attachment, `the attachment is not in the inventory: ${JSON.stringify(listed?.value)}`);

    // One minute is the shortest link this server will issue, and the shortest
    // the expiry assertion below can wait for.
    const text = await callTool("get_attachment", {
      message_id: draftId,
      attachment_id: attachment.id,
      link_ttl_minutes: 1,
    });
    const link = text.match(/Download: (\S+)/)?.[1];
    assert(link, `no download link in the output: ${text}`);
    assert(
      link.startsWith(`${BASE_URL}${DOWNLOAD_ROUTE_PREFIX}`) &&
        /\/[0-9a-f]{64}$/.test(link),
      `the link is not an unguessable id on this Worker: ${link}`
    );
    assert(!/Saved to:/.test(text), `the hosted server claimed to save a file: ${text}`);

    // Anonymous first: the bytes must never be reachable without the bearer.
    const anonymous = await fetch(link);
    assert(anonymous.status === 401, `anonymous download returned ${anonymous.status}, expected 401`);

    const authed = await fetch(link, { headers: { authorization: `Bearer ${bearer}` } });
    assert(authed.status === 200, `authenticated download returned ${authed.status}`);
    assert(
      (authed.headers.get("content-disposition") ?? "").includes("blob.bin"),
      `content-disposition does not name the file: ${authed.headers.get("content-disposition")}`
    );
    const downloaded = Buffer.from(await authed.arrayBuffer());
    assert(
      downloaded.equals(payload),
      `the download is ${downloaded.length} bytes, expected the ${payload.length} that went in`
    );

    // A well-formed id that was never issued is a 404, not somebody's file.
    const unissued = await fetch(
      `${BASE_URL}${DOWNLOAD_ROUTE_PREFIX}${randomBytes(32).toString("hex")}`,
      { headers: { authorization: `Bearer ${bearer}` } }
    );
    assert(unissued.status === 404, `an unissued id returned ${unissued.status}, expected 404`);

    // And the link really dies: the record carries its own deadline, so this is
    // exact rather than dependent on KV getting round to the expiry.
    console.log("      waiting for the 1-minute download link to expire...");
    const expiredStatus = await poll("the download link to expire", 180_000, 10_000, async () => {
      const response = await fetch(link, { headers: { authorization: `Bearer ${bearer}` } });
      return response.status === 200 ? undefined : response.status;
    });
    assert(expiredStatus === 404, `the expired link returned ${expiredStatus}, expected 404`);
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}/permanentDelete`, {
      method: "POST",
    }).catch(() => {});
  }
});

// ------------------------------------------- v8: MIME export without a filesystem

await testAuthed("r23. export_message serves a message's MIME through a bearer-gated link", async () => {
  const inbox = await callGraphServer(
    "/me/mailFolders/inbox/messages?$top=1&$select=id,subject,internetMessageId"
  );
  const message = inbox?.value?.[0];
  assert(message, "the inbox is empty; no message to export");

  const text = await callTool("export_message", {
    message_id: message.id,
    link_ttl_minutes: 1,
  });
  assert(!/Saved to:/.test(text), `the hosted server claimed to save a file: ${text}`);
  const link = text.match(/Download: (\S+)/)?.[1];
  assert(link, `no download link in the output: ${text}`);
  assert(
    link.startsWith(`${BASE_URL}${DOWNLOAD_ROUTE_PREFIX}`) && /\/[0-9a-f]{64}$/.test(link),
    `the link is not an unguessable id on this Worker: ${link}`
  );
  const downloadId = link.slice(link.lastIndexOf("/") + 1);

  try {
    const anonymous = await fetch(link);
    assert(anonymous.status === 401, `anonymous download returned ${anonymous.status}, expected 401`);

    const authed = await fetch(link, { headers: { authorization: `Bearer ${bearer}` } });
    assert(authed.status === 200, `authenticated download returned ${authed.status}`);
    assert(
      (authed.headers.get("content-type") ?? "").startsWith("message/rfc822"),
      `the link does not serve MIME: ${authed.headers.get("content-type")}`
    );
    assert(
      (authed.headers.get("content-disposition") ?? "").includes(".eml"),
      `content-disposition does not name a .eml: ${authed.headers.get("content-disposition")}`
    );

    const served = Buffer.from(await authed.arrayBuffer());
    // Byte-for-byte what Graph hands this machine for the same message.
    const direct = await callGraphServerBytes(
      `/me/messages/${encodeURIComponent(message.id)}/$value`
    );
    assert(
      served.equals(Buffer.from(direct.bytes)),
      `the served MIME is ${served.length} bytes, Graph's own copy is ${direct.bytes.length}`
    );
    const separator = served.indexOf("\r\n\r\n");
    assert(separator > 0, "the served bytes have no RFC 822 header/body separator");
    if (message.internetMessageId) {
      assert(
        served.toString("utf8", 0, separator).includes(message.internetMessageId),
        "the served MIME is not this message (Message-ID does not match)"
      );
    }
  } finally {
    // The record would expire on its own within the minute; drop it now so the
    // run leaves nothing at all in KV.
    await kvDelete(outlookNs, downloadKey(downloadId));
  }
});

// -------------------------------------------- v9: the LLM features, off by default

/** What the LLM tests have to put back. */
const v9: { configBefore?: string | null; auditBefore?: string | null } = {};

await test("r24. the deployed server ships with both LLM features DISABLED", async () => {
  // The gate criterion: whatever this run does later, the state the Worker was
  // deployed in must have auto-filing and the digest off. An absent record is
  // the shipped default and reads back as both off; a present one must say so.
  v9.configBefore = await kvGet(outlookNs, STATE_LLM_CONFIG);
  v9.auditBefore = await kvGet(outlookNs, STATE_LLM_AUDIT);

  const config = v9.configBefore ? (JSON.parse(v9.configBefore) as Partial<LlmConfig>) : null;
  assert(
    config?.filingEnabled !== true,
    `auto-filing is ENABLED on the deployed server: ${v9.configBefore}`
  );
  assert(
    config?.digestEnabled !== true,
    `the morning digest is ENABLED on the deployed server: ${v9.configBefore}`
  );
  console.log(
    `      deployed LLM config: ${v9.configBefore ?? "(absent — the shipped default, both off)"}`
  );

  // And the code really does default to off, so an absent record cannot mean
  // anything else.
  const defaults = await readLlmConfig(createMemoryStateStore("remote"));
  assert(
    !defaults.filingEnabled && !defaults.digestEnabled,
    "readLlmConfig does not default both features off"
  );

  // Both cron schedules the digest needs are declared, alongside the upkeep one.
  const wranglerConfig = JSON.parse(
    (await fs.readFile(path.join(PROJECT_ROOT, "wrangler.jsonc"), "utf8")).replace(/^\s*\/\/.*$/gm, "")
  ) as { triggers?: { crons?: string[] } };
  const crons = wranglerConfig.triggers?.crons ?? [];
  for (const cron of ["0 11 * * *", "0 12 * * *"]) {
    assert(crons.includes(cron), `wrangler.jsonc is missing the "${cron}" digest schedule`);
  }
});

await testAuthed("r25. end-to-end: auto-filing classifies a notified message, then goes back off", async () => {
  const before = JSON.parse((await kvGet(outlookNs, STATE_LLM_AUDIT)) || "[]") as AuditEntry[];

  // Turn it on through the tool, exactly as the owner would. The threshold is
  // lowered for the run because the model is (correctly) more cautious about a
  // subject carrying an obvious test marker than about real mail.
  const enabled = await callTool("manage_auto_filing", { action: "enable_filing" });
  assert(/Auto-filing:\s+ON/.test(enabled), `enable_filing did not report ON:\n${enabled}`);
  await callTool("manage_auto_filing", { action: "set_threshold", threshold: 0.5 });

  let subject = "";
  try {
    subject = `${TEST_PREFIX} Your receipt from Acme Hardware ${randomBytes(4).toString("hex")}`;
    const createText = await callTool("create_draft", {
      to: [ownAddress],
      subject,
      body:
        "Thank you for your order #A-8827 at Acme Hardware. Total charged: $41.20 CAD to your " +
        "Visa ending 4242. This is your receipt; no action is needed.",
    });
    const draftId = createText.match(/Draft id: (\S+)/)?.[1];
    assert(draftId, `could not extract a draft id from: ${createText}`);
    await callTool("send_draft", { draft_id: draftId });

    // Delivery to self can lag, and a Graph notification lags behind that, so
    // this waits generously rather than calling a slow delivery a failure.
    const decision = await poll(
      "the Worker to classify the notified message",
      300_000,
      10_000,
      async () => {
        const log = JSON.parse((await kvGet(outlookNs, STATE_LLM_AUDIT)) || "[]") as AuditEntry[];
        return log.find((entry) => entry.subject?.includes(subject.slice(-8)));
      }
    );

    console.log(
      `      classified as ${decision.action}` +
        (decision.folder ? ` → ${decision.folder}` : "") +
        ` (confidence ${decision.confidence ?? "n/a"}, model ${decision.model}): ${decision.reason}`
    );
    assert(decision.model?.startsWith("claude-haiku-4-5"), `unexpected model: ${decision.model}`);
    assert(
      (decision.usage?.output ?? 0) > 0,
      `the audit entry records no token usage: ${JSON.stringify(decision)}`
    );
    assert(
      decision.action === "moved" || decision.action === "moved+categorized",
      `a plain shop receipt was not filed: ${decision.action} — ${decision.reason}`
    );
    assert(decision.folder, "a move was audited without naming a folder");

    // The move really happened, and not into a folder the allowlist excludes.
    // A probe sent to self exists twice: the Sent Items copy (which never
    // moves) and the delivered copy the classifier filed — only a copy outside
    // both well-known folders can be the filed one.
    const filed = await callGraphServer(
      `/me/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$select=id,parentFolderId`
    );
    const inbox = await callGraphServer("/me/mailFolders/inbox?$select=id");
    const sent = await callGraphServer("/me/mailFolders/sentitems?$select=id");
    const received = (filed?.value ?? []).filter(
      (m: any) => m.parentFolderId !== inbox.id && m.parentFolderId !== sent.id
    );
    assert(received.length > 0, "the probe is still in the inbox — no move took place");
    // The audit names nested folders as "Parent/Child" (the filing allowlist
    // convention); Graph's displayName is the leaf alone, so compare per
    // segment — leaf against the folder, parent (if any) against its parent.
    const folder = await callGraphServer(
      `/me/mailFolders/${encodeURIComponent(received[0].parentFolderId)}?$select=displayName,parentFolderId`
    );
    const segments = String(decision.folder).split("/");
    const leaf = segments[segments.length - 1];
    assert(
      folder.displayName === leaf,
      `the audit log says ${decision.folder} but the message is in ${folder.displayName}`
    );
    if (segments.length > 1) {
      const parent = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(folder.parentFolderId)}?$select=displayName`
      );
      assert(
        parent.displayName === segments[segments.length - 2],
        `the audit log says ${decision.folder} but the message is in ` +
          `${parent.displayName}/${folder.displayName}`
      );
    }
    assert(
      !NEVER_FILE_INTO.includes(String(folder.displayName).toLowerCase()),
      `the classifier filed into ${folder.displayName}, which is on the never-file list`
    );

    // The tool surfaces the same decision to a reader.
    const logText = await callTool("get_auto_filing_log", { limit: 5 });
    assert(logText.includes(decision.folder!), `get_auto_filing_log does not show the move:\n${logText}`);
  } finally {
    // Off again, whatever happened above.
    const disabled = await callTool("manage_auto_filing", { action: "disable_filing" });
    assert(/Auto-filing:\s+OFF/.test(disabled), `disable_filing did not report OFF:\n${disabled}`);
  }
});

// ------------------------------------------------- v10: is the deployment current

// r9 compares the annotations the deployed server actually serves, but tools/list
// needs a bearer and so cannot run headless. This can: /health names the build's
// version, so a headless run still fails loudly when the deployed Worker predates
// the working tree — which is the only way r9 could be passing about stale hints.
await test("r26. the deployed Worker is running this checkout's version", async () => {
  const response = await fetch(`${BASE_URL}/health`);
  assert(response.ok, `GET /health returned HTTP ${response.status}`);
  const body = (await response.json()) as { status?: string; version?: string };
  assert(body.status === "ok", `/health reports status ${body.status}`);
  assert(
    body.version === VERSION,
    `the deployed Worker is v${body.version ?? "(unreported)"}, this checkout is v${VERSION} — ` +
      "run `npm run deploy`"
  );
});

await test("r20. cleanup: the probe message, its notifications and the delta position are gone", async () => {
  // The mail itself: both copies, permanently (the tools only soft-delete).
  const messages = await callGraphServer(
    `/me/messages?$filter=${encodeURIComponent(`startswith(subject,'${TEST_PREFIX}')`)}&$select=id,subject`
  );
  for (const message of messages?.value ?? []) {
    await callGraphServer(`/me/messages/${encodeURIComponent(message.id)}/permanentDelete`, {
      method: "POST",
    });
  }
  const leftover = await callGraphServer(
    `/me/messages?$filter=${encodeURIComponent(`startswith(subject,'${TEST_PREFIX}')`)}&$select=id,subject`
  );
  assert(
    (leftover?.value ?? []).length === 0,
    `leftover test messages: ${JSON.stringify(leftover.value)}`
  );

  // The ring buffer: drop what this run put there, keep everything else. The
  // subscription itself stays — it is the point of the feature, not an artifact.
  const current = JSON.parse((await kvGet(outlookNs, STATE_ACTIVITY)) || "[]") as ActivityEntry[];
  const kept = current.filter(
    (entry) => !String(entry.subject ?? "").startsWith(TEST_PREFIX)
  );
  await kvPut(outlookNs, STATE_ACTIVITY, JSON.stringify(kept));

  const after = await poll("the ring buffer to lose the test entries", 60_000, 5_000, async () => {
    const raw = (await kvGet(outlookNs, STATE_ACTIVITY)) || "[]";
    return raw.includes(TEST_PREFIX) ? undefined : (JSON.parse(raw) as ActivityEntry[]);
  });
  assert(
    !after.some((entry) => String(entry.subject ?? "").startsWith(TEST_PREFIX)),
    "the ring buffer still holds test notifications"
  );
  assert(
    after.length === kept.length,
    `ring buffer holds ${after.length} entries, expected the ${kept.length} that predate this run`
  );

  // The delta position: put back exactly what was there, or remove ours.
  if (v6.deltaBefore) {
    await kvPut(outlookNs, deltaKey("inbox"), v6.deltaBefore);
  } else {
    await kvDelete(outlookNs, deltaKey("inbox"));
  }

  // The LLM state: the config back to exactly what was deployed (usually
  // absent), the audit log back to what predates this run, and the day's
  // budget counter reduced by what this run spent. The deployed server must
  // end this run in the state the gate asserts: both features disabled.
  if (v9.configBefore) {
    await kvPut(outlookNs, STATE_LLM_CONFIG, v9.configBefore);
  } else {
    await kvDelete(outlookNs, STATE_LLM_CONFIG);
  }
  if (v9.auditBefore) {
    await kvPut(outlookNs, STATE_LLM_AUDIT, v9.auditBefore);
  } else {
    await kvDelete(outlookNs, STATE_LLM_AUDIT);
  }

  const restored = await poll("the LLM config to go back to its deployed state", 60_000, 5_000, async () => {
    const raw = await kvGet(outlookNs, STATE_LLM_CONFIG);
    if (!v9.configBefore) return raw === null ? "(absent)" : undefined;
    return raw === v9.configBefore ? raw : undefined;
  });
  const finalConfig = restored === "(absent)" ? null : (JSON.parse(restored) as Partial<LlmConfig>);
  assert(
    finalConfig?.filingEnabled !== true && finalConfig?.digestEnabled !== true,
    `the run left an LLM feature enabled on the deployed server: ${restored}`
  );

  const leftoverAudit = JSON.parse((await kvGet(outlookNs, STATE_LLM_AUDIT)) || "[]") as AuditEntry[];
  assert(
    !leftoverAudit.some((entry) => String(entry.subject ?? "").includes(TEST_PREFIX)),
    "the audit log still holds entries this run created"
  );
});

// ------------------------------------------------------------------- teardown

await test("r21. cleanup: every OAuth record this run created is deleted", async () => {
  const keys = await kvListKeys(oauthNs);
  const created = keys.filter((key) => !baselineOauthKeys.has(key));
  for (const key of created) await kvDelete(oauthNs, key);

  const remaining = (await kvListKeys(oauthNs)).filter((key) => !baselineOauthKeys.has(key));
  assert(
    remaining.length === 0,
    `OAUTH_KV still holds ${remaining.length} key(s) from this run: ${remaining.join(", ")}`
  );
});

await testAuthed("r22. final sweep: the revoked bearer no longer opens /mcp", async () => {
  assert(bearer, "no bearer token");
  // KV caches reads at the edge for up to a minute, so a just-deleted token can
  // still validate briefly. Poll rather than accept that as a pass.
  let status = 0;
  for (let attempt = 0; attempt < 15; attempt++) {
    status = (await mcpCall(bearer!, "tools/list")).status;
    if (status === 401) break;
    await new Promise((resolve) => setTimeout(resolve, 10_000));
  }
  assert(status === 401, `the revoked bearer still returned HTTP ${status} after ~2 minutes`);
});

const passed = outcomes.filter((o) => o.passed).length;
const skipped = outcomes.filter((o) => o.skipped).length;
console.log(
  `\n${passed}/${outcomes.length} remote tests passed` +
    (skipped
      ? ` (${skipped} skipped — the authenticated surface needs an interactive device-code sign-in)`
      : "") +
    "."
);
if (passed !== outcomes.length) {
  for (const outcome of outcomes.filter((o) => !o.passed)) {
    console.log(`  FAILED: ${outcome.name}`);
  }
  process.exitCode = 1;
}