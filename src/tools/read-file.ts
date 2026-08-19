import { z } from "zod";
import {
  downloadDriveItem,
  getDriveItemById,
  getDriveItemByPath,
  itemDisplayPath,
  normalizeDrivePath,
} from "../core/drive.js";
import {
  DOWNLOAD_MAX_BYTES,
  DOWNLOAD_TTL_DEFAULT_MINUTES,
  DOWNLOAD_TTL_MAX_MINUTES,
  storeDownload,
} from "../core/downloads.js";
import { getStateStore } from "../core/state.js";
import { ToolResult, errorResult, formatLocal, isNotFound, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";
import { saveToDownloads } from "./save-local.js";

export const readFileSchema = {
  path: z
    .string()
    .min(1)
    .optional()
    .describe('OneDrive path of the file to read, e.g. "Documents/notes.txt".'),
  item_id: z
    .string()
    .min(1)
    .optional()
    .describe("The file's item id (from search_files or list_folder), instead of path."),
  link_ttl_minutes: z
    .number()
    .int()
    .min(1)
    .max(DOWNLOAD_TTL_MAX_MINUTES)
    .default(DOWNLOAD_TTL_DEFAULT_MINUTES)
    .describe(
      `Hosted server only: how long a binary file's download link stays valid, in minutes (default ${DOWNLOAD_TTL_DEFAULT_MINUTES}, max ${DOWNLOAD_TTL_MAX_MINUTES}). Ignored by the local server, which saves binaries to disk instead.`
    ),
};

const readFileArgs = z.object(readFileSchema);

export const readFileDescription =
  "Read one OneDrive file. Small text files (text/* or JSON under 50 KB) are returned inline on both servers. Anything else depends on where this server runs: the local (stdio) server saves the file to ~/Downloads/outlook-mcp-attachments/ and returns the path, while the hosted server returns a sign-in-required download link that expires within 15 minutes. Same thresholds as get_attachment. Find files with search_files or list_folder.";

const INLINE_LIMIT = 50 * 1024; // same threshold as get_attachment

export async function readFileHandler(input: z.input<typeof readFileArgs>): Promise<ToolResult> {
  return runTool(async () => {
    const { path, item_id, link_ttl_minutes } = readFileArgs.parse(input);
    if ((path === undefined) === (item_id === undefined)) {
      return errorResult("Give exactly one of path or item_id.");
    }

    let normalized = "";
    if (path !== undefined) {
      const result = normalizeDrivePath(path);
      if (!result.ok) return errorResult(result.message);
      normalized = result.path;
      if (normalized === "") return errorResult("path names the drive root, not a file.");
    }

    let item: any;
    try {
      item = item_id !== undefined ? await getDriveItemById(item_id) : await getDriveItemByPath(normalized);
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          item_id !== undefined
            ? `No OneDrive item with id ${item_id}.`
            : `No OneDrive file at path ${JSON.stringify(normalized)}. Find it with search_files or list_folder.`
        );
      }
      throw err;
    }
    if (item.folder !== undefined) {
      return errorResult(
        `${itemDisplayPath(item)} is a folder — use list_folder to see what is inside it.`
      );
    }

    const displayPath = itemDisplayPath(item);
    const size = Number(item.size ?? 0);
    const contentType: string = item.file?.mimeType ?? "application/octet-stream";
    const textLike = /^text\//i.test(contentType) || /^application\/json\b/i.test(contentType);
    const store = getStateStore();
    const remote = store?.mode === "remote";

    // Refuse an oversize remote binary BEFORE downloading it — the size is known.
    if (remote && !(textLike && size < INLINE_LIMIT) && size > DOWNLOAD_MAX_BYTES) {
      return errorResult(
        `${displayPath} is ${formatSize(size)}, too large for this server to hand over ` +
          `(limit ${formatSize(DOWNLOAD_MAX_BYTES)}). Open it in OneDrive (${item.webUrl ?? "onedrive.live.com"}), ` +
          "or read it from the local stdio server, which saves files straight to disk."
      );
    }

    const { bytes } = await downloadDriveItem(item.id);
    const buffer = Buffer.from(bytes);
    const head =
      `File: ${displayPath}\n` +
      `Type: ${contentType}\n` +
      `Size: ${formatSize(buffer.length)}\n` +
      `Modified: ${formatLocal(item.lastModifiedDateTime)}\n` +
      `id: ${item.id}`;

    if (textLike && buffer.length < INLINE_LIMIT) {
      return textResult(`${head}\n\nContent:\n${buffer.toString("utf8")}`);
    }

    if (remote) {
      if (buffer.length > DOWNLOAD_MAX_BYTES) {
        return errorResult(
          `${displayPath} is ${formatSize(buffer.length)}, too large for this server to hand over ` +
            `(limit ${formatSize(DOWNLOAD_MAX_BYTES)}).`
        );
      }
      const parked = await storeDownload(
        store!,
        {
          name: item.name ?? "file",
          contentType,
          base64: buffer.toString("base64"),
          size: buffer.length,
        },
        link_ttl_minutes
      );
      if (!parked.url) {
        return errorResult(
          "This server does not know its own public URL, so it cannot hand out a download link. " +
            "Set PUBLIC_BASE_URL and redeploy."
        );
      }
      return textResult(
        `${head}\n` +
          `Download: ${parked.url}\n` +
          `Link expires: ${formatLocal(parked.expiresAt)} (in ${link_ttl_minutes} min)\n` +
          "The link needs the same sign-in as this connector — it is useless to anyone else, and " +
          "it stops working when it expires. Fetch it again for a fresh link."
      );
    }

    const savePath = await saveToDownloads(item.name ?? "file", buffer);
    return textResult(`${head}\nSaved to: ${savePath}`);
  });
}
