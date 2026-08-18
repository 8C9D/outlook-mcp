// Node-only Graph entry point that may start a device-code sign-in. Kept out of
// core/graph.ts so the Worker bundle never pulls in MSAL or the disk cache.
import { getAccessToken } from "./auth.js";
import { callGraphWithToken } from "./core/graph.js";

/** Interactive-capable Graph call (may trigger device-code sign-in). For CLI scripts only. */
export async function callGraph(path: string, init?: RequestInit): Promise<any> {
  return callGraphWithToken(getAccessToken, path, init);
}
