// One-time (and re-runnable) seeding of the Worker's Microsoft token store.
//
// Runs locally, where the MSAL disk cache lives. It lifts the current refresh
// token out of .token-cache.json and writes it to Workers KV via wrangler, so
// the Worker can mint mailbox access tokens without MSAL and without asking the
// user to consent again. Nothing is printed that would reveal the token, and it
// is handed to wrangler through a 0600 temp file rather than argv, which would
// otherwise be visible to every process on the machine.
import { spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { createHash, randomBytes } from "node:crypto";
import { PROJECT_ROOT } from "../project-root.js";
import { installMsalTokenProvider } from "../auth.js";
import { callGraphServer } from "../core/graph.js";
import { KV_ACCESS_TOKEN, KV_REFRESH_TOKEN } from "../core/kv-keys.js";

const CACHE_PATH = path.join(PROJECT_ROOT, ".token-cache.json");
const WRANGLER_CONFIG = path.join(PROJECT_ROOT, "wrangler.jsonc");

function fingerprint(secret: string): string {
  return createHash("sha256").update(secret).digest("hex").slice(0, 12);
}

function run(args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn("npx", ["wrangler", ...args], {
      cwd: PROJECT_ROOT,
      stdio: ["ignore", "inherit", "inherit"],
    });
    child.on("error", reject);
    child.on("exit", (code) =>
      code === 0 ? resolve() : reject(new Error(`wrangler ${args[0]} ${args[1]} exited ${code}`))
    );
  });
}

/** The OUTLOOK_KV namespace id, read from the committed wrangler config. */
async function outlookNamespaceId(): Promise<string> {
  const raw = await fs.readFile(WRANGLER_CONFIG, "utf8");
  // Strip // comments so the jsonc config can be parsed as JSON.
  const config = JSON.parse(raw.replace(/^\s*\/\/.*$/gm, "")) as {
    kv_namespaces?: { binding: string; id: string }[];
  };
  const entry = config.kv_namespaces?.find((ns) => ns.binding === "OUTLOOK_KV");
  if (!entry) throw new Error("wrangler.jsonc has no OUTLOOK_KV kv_namespaces entry");
  return entry.id;
}

async function readRefreshToken(): Promise<string> {
  let raw: string;
  try {
    raw = await fs.readFile(CACHE_PATH, "utf8");
  } catch {
    throw new Error(`No token cache at ${CACHE_PATH}. Run \`npm run login\` first.`);
  }
  const cache = JSON.parse(raw) as { RefreshToken?: Record<string, { secret?: string }> };
  const entries = Object.values(cache.RefreshToken ?? {});
  const secret = entries[0]?.secret;
  if (!secret) {
    throw new Error(
      `${CACHE_PATH} contains no refresh token. Run \`npm run login\` to re-authenticate.`
    );
  }
  if (entries.length > 1) {
    throw new Error(
      `${CACHE_PATH} holds ${entries.length} refresh tokens (more than one account). ` +
        "Delete the cache and run `npm run login` with the single intended account."
    );
  }
  return secret;
}

installMsalTokenProvider();

// Confirm the cached credentials still work, and report the identity the
// Worker's allowlist has to be set to.
const me = await callGraphServer("/me?$select=id,userPrincipalName,mail,displayName");
const upn = me.userPrincipalName ?? me.mail;
console.log(`Cached sign-in is live: ${me.displayName} <${upn}>`);
console.log(`  Graph /me id : ${me.id}`);
console.log(`  UPN / mail   : ${upn}`);
console.log("  -> these are the values for the ALLOWED_MS_USER_ID / ALLOWED_MS_UPN secrets.\n");

const refreshToken = await readRefreshToken();
const namespaceId = await outlookNamespaceId();
console.log(`Seeding refresh token (sha256:${fingerprint(refreshToken)}…) into OUTLOOK_KV.`);

const tmpFile = path.join(os.tmpdir(), `outlook-mcp-seed-${randomBytes(8).toString("hex")}`);
await fs.writeFile(tmpFile, refreshToken, { mode: 0o600 });
try {
  await run([
    "kv",
    "key",
    "put",
    KV_REFRESH_TOKEN,
    "--path",
    tmpFile,
    "--namespace-id",
    namespaceId,
    "--remote",
  ]);
  // Any access token cached from the previous refresh token is now misleading.
  await run(["kv", "key", "delete", KV_ACCESS_TOKEN, "--namespace-id", namespaceId, "--remote"]);
} finally {
  await fs.rm(tmpFile, { force: true });
}

console.log(
  `\nSeeded. The Worker will rotate ${KV_REFRESH_TOKEN} on every exchange from here on;\n` +
    "re-run this script only after `npm run login` (i.e. after the chain has been broken)."
);
