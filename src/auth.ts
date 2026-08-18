import "dotenv/config";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  PublicClientApplication,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";

const PROJECT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const CACHE_PATH = path.join(PROJECT_ROOT, ".token-cache.json");

// offline_access is added by MSAL automatically; openid/profile must not be listed.
const SCOPES = ["User.Read", "Mail.Read", "Mail.ReadWrite", "Calendars.ReadWrite"];

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
      context.tokenCache.deserialize(await fs.readFile(CACHE_PATH, "utf8"));
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
    }
  },
  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (context.cacheHasChanged) {
      await fs.writeFile(CACHE_PATH, context.tokenCache.serialize(), { mode: 0o600 });
      await fs.chmod(CACHE_PATH, 0o600);
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

/** Thrown when no token can be acquired without user interaction. */
export class AuthRequiredError extends Error {
  constructor(reason?: string) {
    super(
      "Authentication expired. Run `npm run login` in a terminal in ~/dev/outlook-mcp, then retry." +
        (reason ? ` (${reason})` : "")
    );
    this.name = "AuthRequiredError";
  }
}

async function acquireSilent(pca: PublicClientApplication): Promise<string | undefined> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts[0];
  if (!account) return undefined;
  const result = await pca.acquireTokenSilent({ account, scopes: SCOPES });
  return result?.accessToken ?? undefined;
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
