// Short-lived attachment downloads for the hosted server.
//
// The local stdio server can save an attachment to ~/Downloads and hand back a
// path. The Worker has no filesystem and its answer travels as text through the
// MCP client, so binary attachments are parked in the state store under an
// unguessable id and served by an authenticated route instead.
//
// Two separate expiries guard a download, deliberately:
//   * `expiresAt` inside the record, enforced on every read here. This is the
//     real deadline — it is exact, and it holds even when the store still has
//     the bytes (KV expiry is eventual, and its minimum TTL is 60 s).
//   * the store's own TTL, which is only garbage collection.
//
// Nothing in here is a substitute for authentication: the route that serves
// these records is OAuth-gated, and the id is unguessable so a leaked link is
// still useless to anyone without a bearer.
import { downloadKey } from "./kv-keys.js";
import type { StateStore } from "./state.js";

/** Default lifetime of a download link. */
export const DOWNLOAD_TTL_DEFAULT_MINUTES = 15;

/** Hard ceiling on a download link's lifetime. */
export const DOWNLOAD_TTL_MAX_MINUTES = 15;

/**
 * Largest attachment served this way. A KV value may not exceed 25 MB and the
 * bytes are stored base64 (4/3 expansion), so 18 MB of attachment is the most
 * that fits with room to spare for the record's own fields.
 */
export const DOWNLOAD_MAX_BYTES = 18 * 1024 * 1024;

/**
 * The path a download is served from. Deliberately *under* the MCP endpoint:
 * an OAuth client binds its token's audience to the resource it was told about
 * (this server advertises ".../mcp"), and audience matching is path-prefixed on
 * path boundaries. A link at /download/… would be refused as out-of-audience
 * even for the client that asked for it; /mcp/download/… is inside it.
 */
export const DOWNLOAD_ROUTE_PREFIX = "/mcp/download/";

export type DownloadRecord = {
  /** Filename to serve the bytes under. */
  name: string;
  contentType: string;
  /** Size of the decoded bytes, in bytes. */
  size: number;
  /** ISO instant after which this download must be refused. */
  expiresAt: string;
  /** The attachment itself, base64 (exactly as Graph returned it). */
  base64: string;
};

/** 256 bits of randomness as hex — not enumerable, not derived from the message id. */
function newDownloadId(): string {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

/** Park an attachment for `ttlMinutes` and return the id and deadline. */
export async function storeDownload(
  store: StateStore,
  attachment: { name: string; contentType: string; base64: string; size: number },
  ttlMinutes: number
): Promise<{ id: string; expiresAt: string; url?: string }> {
  const minutes = Math.min(Math.max(ttlMinutes, 1), DOWNLOAD_TTL_MAX_MINUTES);
  const expiresAt = new Date(Date.now() + minutes * 60_000).toISOString();
  const id = newDownloadId();
  const record: DownloadRecord = { ...attachment, expiresAt };
  await store.put(downloadKey(id), JSON.stringify(record), { ttlSeconds: minutes * 60 });
  return {
    id,
    expiresAt,
    ...(store.publicBaseUrl ? { url: store.publicBaseUrl + DOWNLOAD_ROUTE_PREFIX + id } : {}),
  };
}

/**
 * Read a parked download, or null when there is none, the id is malformed, or
 * it has expired (in which case the leftover bytes are dropped).
 */
export async function readDownload(
  store: StateStore,
  id: string
): Promise<DownloadRecord | null> {
  if (!/^[0-9a-f]{64}$/.test(id)) return null;
  const raw = await store.get(downloadKey(id));
  if (!raw) return null;
  let record: DownloadRecord;
  try {
    record = JSON.parse(raw) as DownloadRecord;
  } catch {
    return null;
  }
  if (!record?.base64 || Date.parse(record.expiresAt) <= Date.now()) {
    await store.delete(downloadKey(id));
    return null;
  }
  return record;
}
