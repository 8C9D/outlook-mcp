import { getAccessToken, getAccessTokenSilent } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * In-process log of Graph requests: one entry per Graph call (method + path;
 * the transport-level 429 retry does not add a second entry). Used by the test
 * harness to assert request shapes (e.g. that manage_message issued a single
 * /$batch call); not read in server mode.
 */
export const graphRequestLog: { method: string; path: string }[] = [];

/** A non-2xx response from Microsoft Graph, carrying the status and error body. */
export class GraphError extends Error {
  constructor(
    public readonly status: number,
    public readonly statusText: string,
    public readonly requestPath: string,
    public readonly body: string
  ) {
    super(`Graph request failed: ${status} ${statusText} for ${requestPath}\n${body}`);
    this.name = "GraphError";
  }
}

async function callGraphWithToken(
  getToken: () => Promise<string>,
  path: string,
  init?: RequestInit
): Promise<any> {
  const token = await getToken();
  graphRequestLog.push({ method: (init?.method ?? "GET").toUpperCase(), path });
  const doFetch = () =>
    fetch(path.startsWith("https://") ? path : GRAPH_BASE + path, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(init?.headers as Record<string, string> | undefined),
      },
    });

  let response = await doFetch();
  if (response.status === 429) {
    // Honor Retry-After once, then fail readably if still throttled.
    const retryAfter = Number(response.headers.get("Retry-After") ?? "2");
    const waitSeconds = Number.isFinite(retryAfter) && retryAfter > 0 ? Math.min(retryAfter, 60) : 2;
    console.error(`Graph throttled (429) on ${path}; retrying once after ${waitSeconds}s.`);
    await new Promise((resolve) => setTimeout(resolve, waitSeconds * 1000));
    response = await doFetch();
    if (response.status === 429) {
      throw new GraphError(
        429,
        "Too Many Requests",
        path,
        "Microsoft Graph is throttling requests. Wait a minute and retry."
      );
    }
  }
  if (!response.ok) {
    throw new GraphError(response.status, response.statusText, path, await response.text());
  }
  if (response.status === 204) return null;
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

/** Interactive-capable Graph call (may trigger device-code sign-in). For CLI scripts only. */
export async function callGraph(path: string, init?: RequestInit): Promise<any> {
  return callGraphWithToken(getAccessToken, path, init);
}

/**
 * Server-mode Graph call: silent token acquisition only. Throws AuthRequiredError
 * (never prompts) when re-authentication is needed. `path` may also be a full
 * @odata.nextLink URL.
 */
export async function callGraphServer(path: string, init?: RequestInit): Promise<any> {
  return callGraphWithToken(getAccessTokenSilent, path, init);
}
