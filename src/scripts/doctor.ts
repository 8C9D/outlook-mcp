// `npm run doctor` — what to run first when something does not work.
//
// It answers, in order, the three questions a broken install can be stuck on:
// is this checkout configured at all, is there a usable sign-in on disk, and
// does that sign-in still open the mailbox with the permissions the tools need.
// Each check prints PASS / WARN / FAIL and, when it fails, the fix — including
// translations of the three Microsoft errors this project itself hit while the
// Entra app registration was wrong (see SETUP.md).
//
// The environment stage is deliberately free of network and credentials, so it
// is also the check a fresh clone can run: `npm run doctor -- --env-only`.
//
// Everything here is exported and side-effect-free; the CLI entry at the bottom
// only runs when this file is the process's entry point, so the test harness can
// import the checks and assert on them.
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SCOPES, TOKEN_CACHE_PATH, getAccessTokenSilent, getGrantedScopes } from "../auth.js";
import { GraphError, callGraphWithToken } from "../core/graph.js";
import { PROJECT_ROOT } from "../project-root.js";
import { VERSION } from "../core/version.js";

/** WARN is a fact worth knowing that does not stop the server working. */
export type CheckStatus = "pass" | "warn" | "fail";

export type Check = { name: string; status: CheckStatus; detail: string };

/** Node 20 is the oldest runtime this is tested on (and Wrangler's floor). */
const MIN_NODE_MAJOR = 20;

/**
 * The failures worth translating: each one is a real message Microsoft returns
 * for a misconfigured app registration, and each has a single fix. Matched
 * against the whole error text, so an MSAL message or a Graph body both work.
 */
const KNOWN_FAILURES: { match: RegExp; help: string }[] = [
  {
    match: /AADSTS70002\b/,
    help:
      'AADSTS70002 — Entra is treating this app as a confidential client and wants a secret. ' +
      'The fix is on the app registration: Authentication → "Allow public client flows" = Yes ' +
      '(manifest: "allowPublicClient": true). The device-code flow this server signs in with is ' +
      "a public-client flow and cannot present a secret.",
  },
  {
    match: /AADSTS50020\b/,
    help:
      "AADSTS50020 — the signed-in personal Microsoft account does not exist in the directory " +
      "this app registration is scoped to. The fix is the app's audience: supported account " +
      'types must be "Personal Microsoft accounts only" (or "Accounts in any organizational ' +
      'directory and personal Microsoft accounts"). This server signs in against the ' +
      "/consumers authority, which only a personal-account audience accepts.",
  },
  {
    match: /AADSTS700016\b/,
    help:
      "AADSTS700016 — the client id was not found in the consumers directory. Either " +
      "AZURE_CLIENT_ID is not this app's Application (client) ID, or the registration's " +
      "audience excludes personal Microsoft accounts.",
  },
];

/**
 * The fix for a known failure, or undefined. Exported so the harness can assert
 * that each error this project hit still has its translation.
 */
export function translateFailure(message: string): string | undefined {
  return KNOWN_FAILURES.find((known) => known.match.test(message))?.help;
}

/** A Graph 403 here always means the same thing: consent, not code. */
export const FORBIDDEN_HELP =
  "HTTP 403 from Microsoft Graph — the token is valid but carries no permission for that " +
  "resource. Add the delegated permission to the app registration, then run `npm run login` " +
  "again: consent is granted per scope at sign-in, so an existing cached token never gains one.";

/** The message to print for a failed check, with the translation appended. */
export function explain(err: unknown): string {
  const message = err instanceof Error ? err.message : String(err);
  if (err instanceof GraphError && err.status === 403) return `${message}\n${FORBIDDEN_HELP}`;
  const help = translateFailure(message);
  return help ? `${message}\n${help}` : message;
}

/** The scopes this server asks for that a token does not carry. */
export function missingScopes(granted: string[]): string[] {
  const held = new Set(granted.map((scope) => scope.toLowerCase()));
  return SCOPES.filter((scope) => !held.has(scope.toLowerCase()));
}

async function exists(target: string): Promise<boolean> {
  return fs
    .access(target)
    .then(() => true)
    .catch(() => false);
}

/**
 * Stage 1: configuration only — no network, no token, nothing account-specific.
 * A fresh clone with `.env` in place passes this.
 */
export async function environmentChecks(): Promise<Check[]> {
  const checks: Check[] = [];

  const major = Number(process.versions.node.split(".")[0]);
  checks.push({
    name: "node runtime",
    status: major >= MIN_NODE_MAJOR ? "pass" : "fail",
    detail:
      major >= MIN_NODE_MAJOR
        ? `node ${process.versions.node}`
        : `node ${process.versions.node} is older than the supported v${MIN_NODE_MAJOR}.`,
  });

  const sdkInstalled = await exists(
    path.join(PROJECT_ROOT, "node_modules", "@modelcontextprotocol", "sdk", "package.json")
  );
  checks.push({
    name: "dependencies installed",
    status: sdkInstalled ? "pass" : "fail",
    detail: sdkInstalled
      ? `node_modules present under ${PROJECT_ROOT}`
      : "node_modules is missing or incomplete. Run `npm install` in the project root.",
  });

  const envPath = path.join(PROJECT_ROOT, ".env");
  const envPresent = await exists(envPath);
  const clientId = process.env.AZURE_CLIENT_ID ?? "";
  const looksLikeGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    clientId
  );
  checks.push({
    name: "AZURE_CLIENT_ID",
    status: looksLikeGuid ? "pass" : clientId ? "warn" : "fail",
    detail: looksLikeGuid
      ? `set (${clientId})${envPresent ? " from .env or the environment" : " from the environment"}`
      : clientId
        ? `set, but "${clientId}" is not shaped like an Application (client) ID (a GUID).`
        : `not set. Create ${envPath} containing AZURE_CLIENT_ID=<Application (client) ID of the ` +
          "Entra app registration> — see SETUP.md for how to create one.",
  });

  // Only a `node dist/server.js` client config needs the build, so its absence
  // is a fact rather than a fault: `npm run serve` and the tests use tsx.
  const built = await exists(path.join(PROJECT_ROOT, "dist", "server.js"));
  checks.push({
    name: "built server (dist/server.js)",
    status: built ? "pass" : "warn",
    detail: built
      ? "present — a client can be pointed at `node dist/server.js`"
      : "not built. Run `npm run build` before configuring a client to run `node dist/server.js`.",
  });

  return checks;
}

/** Stage 2: the cached sign-in on disk — still no network. */
export async function tokenCacheChecks(): Promise<Check[]> {
  const present = await exists(TOKEN_CACHE_PATH);
  if (!present) {
    return [
      {
        name: "token cache",
        status: "fail",
        detail:
          `${TOKEN_CACHE_PATH} does not exist — nobody has signed in on this machine. ` +
          "Run `npm run login` and complete the device-code sign-in.",
      },
    ];
  }

  const checks: Check[] = [];
  const stat = await fs.stat(TOKEN_CACHE_PATH);
  const mode = stat.mode & 0o777;
  checks.push({
    name: "token cache permissions",
    status: mode === 0o600 ? "pass" : "warn",
    detail:
      mode === 0o600
        ? `${TOKEN_CACHE_PATH} is 0600`
        : `${TOKEN_CACHE_PATH} is ${mode.toString(8).padStart(3, "0")}; it holds a refresh token ` +
          "for the whole mailbox. Run `chmod 600` on it.",
  });

  let accounts = 0;
  let readable = true;
  try {
    const cache = JSON.parse(await fs.readFile(TOKEN_CACHE_PATH, "utf8")) as {
      Account?: Record<string, unknown>;
    };
    accounts = Object.keys(cache.Account ?? {}).length;
  } catch {
    readable = false;
  }
  checks.push({
    name: "token cache contents",
    status: readable && accounts > 0 ? "pass" : "fail",
    detail:
      readable && accounts > 0
        ? `${accounts} cached account${accounts === 1 ? "" : "s"}`
        : readable
          ? "the cache names no account. Run `npm run login` again."
          : "the cache is not readable JSON. Delete it and run `npm run login` again.",
  });

  return checks;
}

/** Stage 3: the live probe — a silent token, its scopes, and two Graph calls. */
export async function liveChecks(): Promise<Check[]> {
  const checks: Check[] = [];

  let token: string;
  try {
    token = await getAccessTokenSilent();
    checks.push({
      name: "silent token acquisition",
      status: "pass",
      detail: "the cached refresh token still works — no sign-in needed",
    });
  } catch (err) {
    checks.push({
      name: "silent token acquisition",
      status: "fail",
      detail:
        `${explain(err)}\nIf the refresh token has simply expired or been revoked, ` +
        "`npm run login` fixes it.",
    });
    return checks;
  }

  const granted = await getGrantedScopes().catch(() => [] as string[]);
  const missing = missingScopes(granted);
  checks.push({
    name: "granted scopes",
    status: missing.length ? "fail" : "pass",
    detail: missing.length
      ? `the sign-in is missing ${missing.join(", ")}. Add the delegated permission(s) to the ` +
        "app registration and run `npm run login` again — consent is granted per scope at " +
        "sign-in, so an existing cached token never gains one."
      : `all ${SCOPES.length} requested scopes granted: ${granted.join(", ")}`,
  });

  const call = (graphPath: string) => callGraphWithToken(async () => token, graphPath);

  try {
    const me = await call("/me?$select=displayName,userPrincipalName,id");
    checks.push({
      name: "Graph /me",
      status: "pass",
      detail: `${me.displayName} <${me.userPrincipalName}> (id ${me.id})`,
    });
  } catch (err) {
    checks.push({ name: "Graph /me", status: "fail", detail: explain(err) });
  }

  // /me needs only User.Read, so a second probe is what proves the mail scopes
  // were really consented to rather than merely requested.
  try {
    const inbox = await call("/me/mailFolders/inbox?$select=displayName,totalItemCount");
    checks.push({
      name: "Graph inbox",
      status: "pass",
      detail: `${inbox.displayName}: ${inbox.totalItemCount} messages`,
    });
  } catch (err) {
    checks.push({ name: "Graph inbox", status: "fail", detail: explain(err) });
  }

  return checks;
}

/**
 * Stage 4: the deployed Worker, when this checkout has one. A local-only install
 * has no Worker, and a Worker that is merely out of date still works, so nothing
 * here can fail the run.
 */
export async function deploymentChecks(): Promise<Check[]> {
  const raw = await fs
    .readFile(path.join(PROJECT_ROOT, "wrangler.jsonc"), "utf8")
    .catch(() => undefined);
  if (!raw) return [];
  const config = JSON.parse(raw.replace(/^\s*\/\/.*$/gm, "")) as {
    vars?: { PUBLIC_BASE_URL?: string };
  };
  const baseUrl = config.vars?.PUBLIC_BASE_URL;
  if (!baseUrl) return [];

  try {
    const response = await fetch(`${baseUrl}/health`);
    const body = (await response.json()) as { status?: string; version?: string };
    const current = body.version === VERSION;
    return [
      {
        name: "deployed Worker",
        status: response.ok ? (current ? "pass" : "warn") : "warn",
        detail: !response.ok
          ? `${baseUrl}/health answered HTTP ${response.status}`
          : current
            ? `${baseUrl} is serving v${body.version}`
            : `${baseUrl} is serving ${body.version ? `v${body.version}` : "a build too old to " +
                "report its version"}, this checkout is v${VERSION}. Run \`npm run deploy\` to ` +
              "bring them level.",
      },
    ];
  } catch (err) {
    return [
      {
        name: "deployed Worker",
        status: "warn",
        detail: `${baseUrl}/health could not be reached: ${explain(err)}`,
      },
    ];
  }
}

function report(checks: Check[]): void {
  for (const check of checks) {
    console.log(`${check.status.toUpperCase().padEnd(4)}  ${check.name}`);
    for (const line of check.detail.split("\n")) console.log(`      ${line}`);
  }
}

async function main(): Promise<void> {
  const envOnly = process.argv.includes("--env-only");
  console.log(`outlook-mcp doctor (v${VERSION})\n`);

  console.log("-- environment --");
  const checks = await environmentChecks();
  report(checks);

  if (!envOnly) {
    console.log("\n-- sign-in on disk --");
    const cache = await tokenCacheChecks();
    report(cache);
    checks.push(...cache);

    // A live probe with no usable cache would only repeat what stage 2 said.
    if (!cache.some((check) => check.status === "fail")) {
      console.log("\n-- live mailbox probe --");
      const live = await liveChecks();
      report(live);
      checks.push(...live);
    }

    console.log("\n-- deployment --");
    const deployment = await deploymentChecks();
    if (deployment.length) report(deployment);
    else console.log("      no PUBLIC_BASE_URL in wrangler.jsonc — nothing deployed to check.");
    checks.push(...deployment);
  }

  const failed = checks.filter((check) => check.status === "fail");
  const warned = checks.filter((check) => check.status === "warn");
  console.log(
    `\n${checks.length - failed.length}/${checks.length} checks passed` +
      (warned.length ? ` (${warned.length} warning${warned.length === 1 ? "" : "s"})` : "") +
      (envOnly ? " — environment stage only." : ".")
  );
  if (failed.length) {
    console.log("\nFix these first:");
    for (const check of failed) console.log(`  - ${check.name}`);
    process.exitCode = 1;
  }
}

// Only when run as a script: importing this module must not start a probe.
const entry = process.argv[1];
if (entry && path.resolve(entry) === fileURLToPath(import.meta.url)) {
  await main();
}
