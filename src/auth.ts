import dotenv from "dotenv";
import { promises as fs } from "node:fs";
import path from "node:path";
import {
  PublicClientApplication,
  type AuthenticationResult,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { PROJECT_ROOT } from "./project-root.js";
import { AuthRequiredError, setDefaultTokenProvider } from "./core/token.js";

export { AuthRequiredError };

// Explicit path: bare dotenv.config() reads .env from process.cwd(), which is
// arbitrary under Claude Desktop. quiet keeps dotenv's tip off stdout, which
// must carry only JSON-RPC in server mode.
dotenv.config({ path: path.join(PROJECT_ROOT, ".env"), quiet: true });

/** Where MSAL's serialized cache lives. `npm run doctor` reports on this file. */
export const TOKEN_CACHE_PATH = path.join(PROJECT_ROOT, ".token-cache.json");

// offline_access is added by MSAL automatically; openid/profile must not be listed.
export const SCOPES = [
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.ReadWrite",
  "Contacts.ReadWrite",
  "MailboxSettings.ReadWrite",
  "Tasks.ReadWrite",
];

function requireClientId(): string {
  const clientId = process.env.AZURE_CLIENT_ID;
  if (!clientId) {
    throw new Error(
      "AZURE_CLIENT_ID is not set. Create .env in the project root containing " +
        "AZURE_CLIENT_ID=<Application (client) ID of the outlook-mcp Entra app registration>."
    );
  }
  return clientId;
}

const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    try {
      context.tokenCache.deserialize(await fs.readFile(TOKEN_CACHE_PATH, "utf8"));
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
    }
  },
  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (context.cacheHasChanged) {
      await fs.writeFile(TOKEN_CACHE_PATH, context.tokenCache.serialize(), { mode: 0o600 });
      await fs.chmod(TOKEN_CACHE_PATH, 0o600);
    }
  },
};

let pcaInstance: PublicClientApplication | undefined;

function getPca(): PublicClientApplication {
  pcaInstance ??= new PublicClientApplication({
    auth: {
      clientId: requireClientId(),
      authority: "https://login.microsoftonline.com/consumers",
    },
    cache: { cachePlugin },
  });
  return pcaInstance;
}

async function acquireByDeviceCode(pca: PublicClientApplication): Promise<string> {
  const result = await pca.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (deviceCode) => {
      // stderr: stdout must stay clean in case this ever runs under an MCP client.
      console.error(`\n=== Sign-in required ===\n${deviceCode.message}\n`);
    },
  });
  if (!result?.accessToken) {
    throw new Error("Device-code flow completed without returning an access token.");
  }
  return result.accessToken;
}

async function acquireSilentResult(
  pca: PublicClientApplication
): Promise<AuthenticationResult | undefined> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0];
  if (!account) return undefined;
  return (await pca.acquireTokenSilent({ account, scopes: SCOPES })) ?? undefined;
}

async function acquireSilent(pca: PublicClientApplication): Promise<string | undefined> {
  return (await acquireSilentResult(pca))?.accessToken ?? undefined;
}

/**
 * The scopes the cached sign-in actually carries, as MSAL reports them. Graph
 * access tokens for a personal Microsoft account are not JWTs, so this — not
 * decoding the token — is the only way to see what was consented to.
 * Throws AuthRequiredError on the same conditions as getAccessTokenSilent.
 */
export async function getGrantedScopes(): Promise<string[]> {
  const pca = getPca();
  try {
    const result = await acquireSilentResult(pca);
    if (result) return result.scopes ?? [];
    throw new AuthRequiredError("no cached account");
  } catch (err) {
    if (err instanceof AuthRequiredError) throw err;
    throw new AuthRequiredError(err instanceof Error ? err.message : String(err));
  }
}

/**
 * Server-mode token getter: silent acquisition only. Never starts a device-code
 * flow; throws AuthRequiredError when the cache is missing or unusable.
 */
export async function getAccessTokenSilent(): Promise<string> {
  const pca = getPca();
  try {
    const token = await acquireSilent(pca);
    if (token) return token;
    throw new AuthRequiredError("no cached account");
  } catch (err) {
    if (err instanceof AuthRequiredError) throw err;
    const reason = err instanceof Error ? err.message : String(err);
    throw new AuthRequiredError(reason);
  }
}

/**
 * Make MSAL + the on-disk cache the token source for core/graph.js. Called by
 * every Node entry point (stdio server, CLI scripts, test harness); the Worker
 * installs a KV-backed provider instead and never loads this module.
 */
export function installMsalTokenProvider(): void {
  setDefaultTokenProvider(getAccessTokenSilent);
}

export async function getAccessToken(): Promise<string> {
  const pca = getPca();
  try {
    const token = await acquireSilent(pca);
    if (token) return token;
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    console.error(
      `Cached sign-in could not be refreshed (refresh token expired or revoked): ${reason}\n` +
        "Re-authentication is required; starting a fresh device-code sign-in."
    );
  }
  return acquireByDeviceCode(pca);
}
