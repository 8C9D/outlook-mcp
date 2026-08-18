// Live tests against the deployed Cloudflare Worker.
//
// Everything here runs from the local machine against the real workers.dev
// endpoint and the real Microsoft tenant: discovery metadata, refusal of
// anonymous callers, a full OAuth authorization-code exchange, an MCP
// round-trip over Streamable HTTP, and proof that the mailbox refresh token in
// KV actually rotates. Like the local harness it cleans up after itself — the
// OAuth client it registers and every grant and token derived from it are
// deleted from KV at the end, and the final sweep asserts nothing is left.
//
// Prerequisites: `npm run deploy`, the three wrangler secrets, and `npm run
// seed:kv`. Run with `npm run test:remote`.
import { spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import { createHash, randomBytes } from "node:crypto";
import path from "node:path";
import { PROJECT_ROOT } from "./project-root.js";
import { installMsalTokenProvider, getAccessTokenSilent } from "./auth.js";
import { KV_ACCESS_TOKEN, KV_REFRESH_TOKEN } from "./core/kv-keys.js";
import { TOOLS } from "./core/registry.js";

const BASE_URL = process.env.MCP_REMOTE_URL ?? "https://outlook-mcp.arthur-yuhao-zhang.workers.dev";
const MCP_URL = `${BASE_URL}/mcp`;
const TEST_CLIENT_NAME = "[MCP TEST] remote harness";

type Outcome = { name: string; passed: boolean; detail?: string };
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

await test("r3. anonymous and bogus-bearer requests to /mcp are refused", async () => {
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

await test("r5. authorization refuses an identity it cannot verify (allowlist gate)", async () => {
  assert(registeredClientId, "no registered client (r4 must pass first)");
  const verifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(verifier).digest());
  const response = await fetch(authorizeUrl(challenge, "state-reject"), {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ ms_access_token: "definitely-not-a-microsoft-token" }),
  });
  assert(
    response.status === 401,
    `an unverifiable Microsoft token was answered with ${response.status}, expected 401`
  );
  const body = (await response.json()) as any;
  assert(body.status === "failed", `expected status "failed", got ${JSON.stringify(body)}`);
});

let codeVerifier = "";

await test("r6. the owner's Microsoft identity completes authorization and yields a bearer", async () => {
  assert(registeredClientId, "no registered client (r4 must pass first)");
  codeVerifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(codeVerifier).digest());
  const state = "state-" + randomBytes(6).toString("hex");

  // Stands in for the browser device-code sign-in: the same Graph /me allowlist
  // check runs, using a Microsoft token this machine already holds.
  const msAccessToken = await getAccessTokenSilent();
  const authorizeResponse = await fetch(authorizeUrl(challenge, state), {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ ms_access_token: msAccessToken }),
  });
  const authorizeBody = await authorizeResponse.text();
  assert(
    authorizeResponse.ok,
    `authorize failed: HTTP ${authorizeResponse.status} ${authorizeBody}`
  );
  const authorized = JSON.parse(authorizeBody) as any;
  assert(authorized.status === "ok", `authorize said ${JSON.stringify(authorized)}`);

  const redirect = new URL(authorized.redirectTo);
  assert(redirect.searchParams.get("state") === state, "authorization returned the wrong state");
  const code = redirect.searchParams.get("code");
  assert(code, `no authorization code in ${authorized.redirectTo}`);

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

await test("r8. initialize over Streamable HTTP reports this server", async () => {
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

await test(`r9. tools/list serves all ${TOOLS.length} tools, identical to the stdio surface`, async () => {
  assert(bearer, "no bearer token");
  const result = resultOf(await mcpCall(bearer!, "tools/list"), "tools/list");
  const remoteNames = (result.tools as { name: string }[]).map((t) => t.name).sort();
  const localNames = TOOLS.map((t) => t.name).sort();
  assert(
    remoteNames.length === localNames.length,
    `remote has ${remoteNames.length} tools, the registry has ${localNames.length}`
  );
  assert(
    remoteNames.join(",") === localNames.join(","),
    `tool surfaces differ:\n  remote: ${remoteNames.join(", ")}\n  local:  ${localNames.join(", ")}`
  );
});

await test("r10. prompts/list serves both prompts", async () => {
  assert(bearer, "no bearer token");
  const result = resultOf(await mcpCall(bearer!, "prompts/list"), "prompts/list");
  const names = (result.prompts as { name: string }[]).map((p) => p.name).sort();
  assert(
    names.join(",") === "morning_brief,triage_inbox",
    `prompts are ${names.join(", ")}`
  );
});

await test("r11. tools/call list_events reaches Graph through the KV token", async () => {
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

await test("r12. the mailbox refresh token in KV rotates on every exchange", async () => {
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

// ------------------------------------------------------------------- teardown

await test("r13. cleanup: every OAuth record this run created is deleted", async () => {
  const keys = await kvListKeys(oauthNs);
  const created = keys.filter((key) => !baselineOauthKeys.has(key));
  for (const key of created) await kvDelete(oauthNs, key);

  const remaining = (await kvListKeys(oauthNs)).filter((key) => !baselineOauthKeys.has(key));
  assert(
    remaining.length === 0,
    `OAUTH_KV still holds ${remaining.length} key(s) from this run: ${remaining.join(", ")}`
  );
});

await test("r14. final sweep: the revoked bearer no longer opens /mcp", async () => {
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
console.log(`\n${passed}/${outcomes.length} remote tests passed.`);
if (passed !== outcomes.length) {
  for (const outcome of outcomes.filter((o) => !o.passed)) {
    console.log(`  FAILED: ${outcome.name}`);
  }
  process.exitCode = 1;
}