import { z } from "zod";
import {
  baseNameOf,
  conflictBehaviorFor,
  itemDisplayPath,
  normalizeDrivePath,
  uploadDriveFile,
} from "../core/drive.js";
import { getStateStore } from "../core/state.js";
import {
  SourceReadError,
  contentTypeForFile,
  prepareBase64Source,
  prepareFileSource,
  prepareUrlSource,
} from "./file-sources.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";

export const uploadFileSchema = {
  destination_path: z
    .string()
    .min(1)
    .describe(
      'Where in OneDrive to put the file, INCLUDING the filename — e.g. "Documents/Reports/summary.pdf". Missing folders are created automatically.'
    ),
  file_path: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 1 of 3: absolute local path of a file on the machine running this server. Works only on the local (stdio) server — the hosted server has no filesystem and rejects it."
    ),
  url: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 2 of 3: an https:// URL the server downloads and uploads (max 25 MB). The URL must be reachable without credentials."
    ),
  content_base64: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 3 of 3: the file's bytes, base64-encoded, for content you already hold (max 3 MB decoded)."
    ),
  overwrite: z
    .boolean()
    .default(false)
    .describe(
      "If a file already exists at destination_path: false (default) uploads under an auto-renamed name like 'summary 1.pdf' and reports it; true REPLACES the existing file's content."
    ),
};

const uploadFileArgs = z.object(uploadFileSchema);

export const uploadFileDescription =
  "Upload a file to OneDrive (max 25 MB) from exactly one of three sources: file_path (a local file — local stdio server only), url (an https link this server fetches), or content_base64 (bytes you supply, max 3 MB). destination_path names the target folder AND filename; missing folders are created. On a name collision the default is rename-not-overwrite (the upload gets a numbered name); pass overwrite: true to deliberately replace the existing file.";

export async function uploadFileHandler(
  input: z.input<typeof uploadFileArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { destination_path, file_path, url, content_base64, overwrite } =
      uploadFileArgs.parse(input);

    const given = [
      ["file_path", file_path],
      ["url", url],
      ["content_base64", content_base64],
    ].filter(([, value]) => value !== undefined);
    if (given.length !== 1) {
      return errorResult(
        given.length === 0
          ? "Give exactly one source: file_path (local server only), url, or content_base64."
          : `Give exactly one source, not ${given.length} (${given.map(([n]) => n).join(", ")}).`
      );
    }
    if (file_path !== undefined && getStateStore()?.mode === "remote") {
      return errorResult(
        "file_path works only on the local (stdio) server: this hosted server has no access to " +
          "your filesystem. Upload with url (an https link, up to 25 MB) or content_base64 " +
          "(bytes inline, up to 3 MB) instead."
      );
    }

    const normalized = normalizeDrivePath(destination_path);
    if (!normalized.ok) return errorResult(normalized.message);
    const destination = normalized.path;
    if (destination === "") {
      return errorResult("destination_path must include a filename, e.g. \"Documents/report.pdf\".");
    }
    const filename = baseNameOf(destination);

    const prepared =
      file_path !== undefined
        ? await prepareFileSource(file_path, filename)
        : url !== undefined
          ? await prepareUrlSource(url, filename)
          : prepareBase64Source(content_base64!, filename);
    if (!prepared.ok) return errorResult(prepared.message);

    let buffer: Buffer;
    let contentType: string;
    try {
      const loaded = await prepared.source.read();
      buffer = loaded.buffer;
      contentType = loaded.contentType ?? contentTypeForFile(filename);
    } catch (err) {
      if (err instanceof SourceReadError) return errorResult(err.message);
      throw err;
    }

    const item = await uploadDriveFile(destination, buffer, contentType, conflictBehaviorFor(overwrite));

    const finalPath = itemDisplayPath(item);
    const renamed = item?.name && item.name !== filename;
    return textResult(
      `File uploaded to OneDrive.\n` +
        `Path: ${finalPath}\n` +
        `Type: ${contentType}\n` +
        `Size: ${formatSize(buffer.length)}\n` +
        `id: ${item?.id ?? "(unknown)"}` +
        (renamed
          ? `\nNote: ${JSON.stringify(filename)} already existed, so the upload was auto-renamed to ` +
            `${JSON.stringify(item.name)} (pass overwrite: true to replace instead).`
          : "")
    );
  });
}
