import { getAccessToken } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export async function callGraph(path: string, init?: RequestInit): Promise<any> {
  const token = await getAccessToken();
  const response = await fetch(GRAPH_BASE + path, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init?.headers as Record<string, string> | undefined),
    },
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(
      `Graph request failed: ${response.status} ${response.statusText} for ${path}\n${body}`
    );
  }
  if (response.status === 204) return null;
  return response.json();
}
