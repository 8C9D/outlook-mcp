// OneDrive (Microsoft Graph /me/drive) helpers shared by the file tools on
// both transports. Free of Node-only imports, like the rest of core/.
//
// Live behaviour this module is built on, verified against the real personal
// drive before anything was designed in (details in ASSUMPTIONS.md "v13"):
//   * PUT …:/content DEFAULTS TO REPLACE on a name collision — silently. Every
//     upload therefore states @microsoft.graph.conflictBehavior explicitly:
//     "rename" unless the caller asked to overwrite.
//   * Upload by path auto-creates missing parent folders (mkdir -p semantics).
//   * DELETE is soft (recycle bin). The bin cannot be LISTED on a personal
//     drive (every documented shape 400s), but a deleted item can be restored
//     by id (POST /restore) and then permanently deleted — which is how tests
//     verify bin presence and clean up after themselves.
//   * createLink with no scope defaults to "anonymous" on a personal drive, and
//     is idempotent per link type (the same URL and permission id come back).
//   * Search indexing lags: a fresh file appears in /search after ~15 s on a
//     good day and minutes on a bad one.
import { callGraphServer, callGraphServerBytes, type GraphBytes } from "./graph.js";

/** Cap on children returned when listing one folder. */
export const LIST_FOLDER_CAP = 500;

/** Simple-upload limit: Graph accepts single PUTs up to 4 MB; above, a session. */
export const DRIVE_SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;

const UPLOAD_CHUNK_SIZE = 4 * 1024 * 1024; // multiple of 320 KiB, as Graph requires

/** Fields every drive read selects; enough to describe an item fully. */
export const DRIVE_ITEM_SELECT =
  "id,name,size,folder,file,root,parentReference,createdDateTime,lastModifiedDateTime,webUrl";

// ---------------------------------------------------------------- pure logic

export type DrivePathResult = { ok: true; path: string } | { ok: false; message: string };

/**
 * Normalize a caller-supplied OneDrive path: backslashes become slashes, empty
 * segments collapse, and "" (or "/") means the drive root. "." and ".." are
 * refused — a drive path names a location, it does not navigate.
 */
export function normalizeDrivePath(raw: string): DrivePathResult {
  const segments = raw.replace(/\\/g, "/").split("/").filter((s) => s.trim() !== "");
  for (const segment of segments) {
    if (segment === "." || segment === "..") {
      return { ok: false, message: `Path segments "." and ".." are not allowed (got: ${raw}).` };
    }
  }
  return { ok: true, path: segments.map((s) => s.trim()).join("/") };
}

/** The folder part of a normalized path ("" for a root-level name). */
export function parentPathOf(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

/** The final segment of a normalized path. */
export function baseNameOf(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? path : path.slice(idx + 1);
}

/**
 * The Graph URL for a normalized path, with an optional endpoint suffix
 * ("/children", "/content", …). Handles the two addressing grammars: the root
 * is `/me/drive/root/children`, a path is `/me/drive/root:/a/b:/children`.
 */
export function drivePathUrl(path: string, suffix = ""): string {
  if (path === "") return `/me/drive/root${suffix}`;
  const encoded = path.split("/").map(encodeURIComponent).join("/");
  return `/me/drive/root:/${encoded}${suffix ? `:${suffix}` : ""}`;
}

/**
 * Map the tool-level overwrite flag to Graph's conflict vocabulary. The default
 * is rename-not-overwrite because Graph's own default (replace) destroys the
 * existing file silently.
 */
export function conflictBehaviorFor(overwrite: boolean): "replace" | "rename" {
  return overwrite ? "replace" : "rename";
}

/**
 * Whether an item passes the caller's type filter: "file", "folder", or a
 * filename extension (with or without the dot, any case).
 */
export function matchesTypeFilter(
  item: { name?: string; folder?: unknown; file?: unknown },
  type: string | undefined
): boolean {
  if (!type) return true;
  const wanted = type.trim().toLowerCase();
  if (wanted === "folder") return item.folder !== undefined;
  if (wanted === "file") return item.folder === undefined;
  const ext = wanted.startsWith(".") ? wanted : `.${wanted}`;
  return item.folder === undefined && (item.name ?? "").toLowerCase().endsWith(ext);
}

/** Sort children the way a file browser does: folders first, then names. */
export function compareDriveChildren(
  a: { name?: string; folder?: unknown },
  b: { name?: string; folder?: unknown }
): number {
  const aFolder = a.folder !== undefined;
  const bFolder = b.folder !== undefined;
  if (aFolder !== bFolder) return aFolder ? -1 : 1;
  return (a.name ?? "").localeCompare(b.name ?? "", undefined, { sensitivity: "base" });
}

/** Human path of an item, from its parentReference ("/" for the root itself). */
export function itemDisplayPath(item: {
  name?: string;
  parentReference?: { path?: string };
  root?: unknown;
}): string {
  if (item.root !== undefined) return "/";
  let parent = item.parentReference?.path ?? "";
  parent = parent.replace(/^\/drive\/root:?/, "");
  try {
    parent = decodeURIComponent(parent);
  } catch {
    // keep as is — some special characters arrive already decoded
  }
  return `${parent}/${item.name ?? ""}`;
}

/** The machine-readable shape of one drive item (used in structuredContent too). */
export function describeDriveItem(item: any): Record<string, unknown> {
  const folder = item?.folder !== undefined;
  return {
    id: String(item?.id ?? ""),
    name: String(item?.name ?? ""),
    path: itemDisplayPath(item ?? {}),
    kind: folder ? "folder" : "file",
    size: Number(item?.size ?? 0),
    ...(folder ? { childCount: Number(item?.folder?.childCount ?? 0) } : {}),
    ...(item?.file?.mimeType ? { mimeType: String(item.file.mimeType) } : {}),
    ...(item?.lastModifiedDateTime ? { modified: String(item.lastModifiedDateTime) } : {}),
    ...(item?.webUrl ? { webUrl: String(item.webUrl) } : {}),
  };
}

// ------------------------------------------------------------- Graph traffic

/** GET one item by id, selecting the standard fields. */
export async function getDriveItemById(itemId: string): Promise<any> {
  return callGraphServer(
    `/me/drive/items/${encodeURIComponent(itemId)}?$select=${DRIVE_ITEM_SELECT}`
  );
}

/** GET one item by normalized path ("" = the root folder itself). */
export async function getDriveItemByPath(path: string): Promise<any> {
  return callGraphServer(`${drivePathUrl(path)}?$select=${DRIVE_ITEM_SELECT}`);
}

/** The bytes of a file item (Graph answers /content with a 302 fetch follows). */
export async function downloadDriveItem(itemId: string): Promise<GraphBytes> {
  return callGraphServerBytes(`/me/drive/items/${encodeURIComponent(itemId)}/content`);
}

/**
 * Upload bytes to a normalized destination path, honoring the conflict
 * behavior. Small files go up in one PUT; larger ones through an upload
 * session in 4 MB chunks (the session's uploadUrl carries its own auth token
 * and lives outside graph.microsoft.com, so chunk PUTs are plain fetches).
 * Returns the created item — whose name may differ from the destination's
 * basename when "rename" resolved a collision.
 */
export async function uploadDriveFile(
  destinationPath: string,
  buffer: Buffer | Uint8Array,
  contentType: string,
  behavior: "replace" | "rename"
): Promise<any> {
  if (buffer.length <= DRIVE_SIMPLE_UPLOAD_LIMIT) {
    return callGraphServer(
      `${drivePathUrl(destinationPath, "/content")}?@microsoft.graph.conflictBehavior=${behavior}`,
      {
        method: "PUT",
        headers: { "Content-Type": contentType },
        body: new Uint8Array(buffer),
      }
    );
  }

  const session = await callGraphServer(drivePathUrl(destinationPath, "/createUploadSession"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": behavior,
        name: baseNameOf(destinationPath),
      },
    }),
  });
  const uploadUrl: string | undefined = session?.uploadUrl;
  if (!uploadUrl) throw new Error("Graph did not return an uploadUrl for the file upload session.");

  let item: any = null;
  for (let offset = 0; offset < buffer.length; offset += UPLOAD_CHUNK_SIZE) {
    const chunk = buffer.subarray(offset, Math.min(offset + UPLOAD_CHUNK_SIZE, buffer.length));
    const range = `bytes ${offset}-${offset + chunk.length - 1}/${buffer.length}`;
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": range,
      },
      body: new Uint8Array(chunk),
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(
        `File upload failed at ${range}: HTTP ${response.status} ${response.statusText}\n${body.slice(0, 500)}`
      );
    }
    const text = await response.text();
    if (text) {
      try {
        item = JSON.parse(text);
      } catch {
        // intermediate chunks answer with ranges, not the item
      }
    }
  }
  if (!item?.id) {
    // The final chunk should have carried the driveItem; fall back to a read.
    return getDriveItemByPath(destinationPath);
  }
  return item;
}
