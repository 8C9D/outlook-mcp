// Transport-agnostic Microsoft Graph token indirection.
//
// The tool layer must run unchanged in two very different hosts: the local
// stdio server (MSAL + disk cache, Node) and the Cloudflare Worker (refresh
// token in KV, no MSAL). Neither implementation may be imported by the tools,
// so callers install a provider and the tools just ask for a token.
//
// AsyncLocalStorage scopes a provider to one request: on Workers the KV binding
// arrives per-request via `env`, and concurrent fetches must not observe each
// other's provider. The default provider is the fallback for single-tenant
// hosts (stdio, CLI scripts) that install one provider for the process.
import { AsyncLocalStorage } from "node:async_hooks";

/** Returns a Graph access token without any user interaction. */
export type TokenProvider = () => Promise<string>;

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

const storage = new AsyncLocalStorage<TokenProvider>();
let defaultProvider: TokenProvider | undefined;

/** Install the process-wide provider (stdio server, CLI scripts, test harness). */
export function setDefaultTokenProvider(provider: TokenProvider): void {
  defaultProvider = provider;
}

/** Run `fn` with `provider` scoped to it, overriding the default (Worker requests). */
export function runWithTokenProvider<T>(provider: TokenProvider, fn: () => T): T {
  return storage.run(provider, fn);
}

/** Token for a server-mode Graph call. Never prompts. */
export function getAccessTokenSilent(): Promise<string> {
  const provider = storage.getStore() ?? defaultProvider;
  if (!provider) {
    throw new AuthRequiredError("no token provider installed");
  }
  return provider();
}
