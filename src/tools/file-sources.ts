// The three caller-supplied byte sources shared by add_attachment and
// upload_file: a local file path (stdio server only), an https URL, or inline
// base64. Each source is validated as cheaply as possible up front (existence,
// scheme, size where it is knowable) and only read after the caller has
// verified its own preconditions (the draft exists, the destination parses), so
// a 25 MB download is never spent on a call that was going to fail anyway.
import { promises as fs } from "node:fs";
import path from "node:path";
import { formatSize } from "./read-message.js";

/** Hard cap for any file this server moves around, matching Outlook's own limits. */
export const MAX_FILE_SIZE = 25 * 1024 * 1024;

/** Cap for inline content_base64 (decoded) — as much base64 as is sane in context. */
export const BASE64_MAX = 3 * 1024 * 1024;

const CONTENT_TYPES: Record<string, string> = {
  ".txt": "text/plain",
  ".md": "text/markdown",
  ".csv": "text/csv",
  ".html": "text/html",
  ".htm": "text/html",
  ".ics": "text/calendar",
  ".json": "application/json",
  ".xml": "application/xml",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".heic": "image/heic",
  ".svg": "image/svg+xml",
  ".zip": "application/zip",
  ".gz": "application/gzip",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".mp3": "audio/mpeg",
  ".mp4": "video/mp4",
  ".mov": "video/quicktime",
};

export function contentTypeForFile(filename: string): string {
  return CONTENT_TYPES[path.extname(filename).toLowerCase()] ?? "application/octet-stream";
}

/**
 * A source that has passed every check cheap enough to make before the
 * expensive part. `read` produces the bytes, and may sharpen the content type
 * with what the transfer revealed.
 */
export type PreparedSource = {
  name: string;
  read: () => Promise<{ buffer: Buffer; contentType?: string }>;
};

export type SourcePreparation =
  | { ok: true; source: PreparedSource }
  | { ok: false; message: string };

/** A transfer that failed for a reason the caller can act on. */
export class SourceReadError extends Error {}

/** Last path segment of a URL, as a filename. */
function nameFromUrl(url: URL): string {
  const segment = url.pathname.split("/").filter(Boolean).pop() ?? "";
  let decoded = segment;
  try {
    decoded = decodeURIComponent(segment);
  } catch {
    // keep the raw segment
  }
  return decoded.replace(/[/\\:\0-\x1f]/g, "_").trim() || "download";
}

export async function prepareFileSource(
  filePath: string,
  overrideName?: string
): Promise<SourcePreparation> {
  if (!path.isAbsolute(filePath)) {
    return { ok: false, message: `file_path must be an absolute path, got: ${filePath}` };
  }
  let stat;
  try {
    stat = await fs.stat(filePath);
  } catch {
    return { ok: false, message: `File not found or unreadable: ${filePath}` };
  }
  if (!stat.isFile()) return { ok: false, message: `Not a regular file: ${filePath}` };
  try {
    await fs.access(filePath, fs.constants.R_OK);
  } catch {
    return { ok: false, message: `File exists but is not readable: ${filePath}` };
  }
  // Size guard BEFORE reading any bytes — a stat is enough to refuse oversize files.
  if (stat.size > MAX_FILE_SIZE) {
    return {
      ok: false,
      message:
        `File is ${formatSize(stat.size)}, over this tool's 25 MB cap. ` +
        "Share large files another way (e.g. a cloud link pasted into the draft body).",
    };
  }
  return {
    ok: true,
    source: {
      name: overrideName ?? path.basename(filePath),
      read: async () => ({ buffer: await fs.readFile(filePath) }),
    },
  };
}

/**
 * Download the URL, refusing anything over the cap. The body is read chunk by
 * chunk and abandoned the moment it grows too large, so a server that lies
 * about (or omits) Content-Length cannot make this buffer 2 GB.
 */
export async function prepareUrlSource(
  rawUrl: string,
  overrideName?: string
): Promise<SourcePreparation> {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    return { ok: false, message: `Not a valid URL: ${rawUrl}` };
  }
  if (url.protocol !== "https:") {
    return {
      ok: false,
      message: `url must be https:// (got ${url.protocol}//). Plaintext and local-file URLs are refused.`,
    };
  }

  return {
    ok: true,
    source: {
      name: overrideName ?? nameFromUrl(url),
      read: async () => {
        const response = await fetch(url, { redirect: "follow" });
        if (!response.ok) {
          throw new SourceReadError(
            `Could not download ${url.href}: HTTP ${response.status} ${response.statusText}`
          );
        }
        const declared = Number(response.headers.get("content-length") ?? "");
        if (Number.isFinite(declared) && declared > MAX_FILE_SIZE) {
          throw new SourceReadError(
            `That URL is ${formatSize(declared)}, over this tool's 25 MB cap.`
          );
        }
        const chunks: Uint8Array[] = [];
        let total = 0;
        const reader = response.body?.getReader();
        if (reader) {
          for (;;) {
            const { done, value } = await reader.read();
            if (done) break;
            total += value.length;
            if (total > MAX_FILE_SIZE) {
              await reader.cancel();
              throw new SourceReadError(
                "That URL is over this tool's 25 MB cap (the download was stopped)."
              );
            }
            chunks.push(value);
          }
        } else {
          const body = new Uint8Array(await response.arrayBuffer());
          if (body.length > MAX_FILE_SIZE) {
            throw new SourceReadError(
              `That URL is ${formatSize(body.length)}, over this tool's 25 MB cap.`
            );
          }
          chunks.push(body);
        }
        const header = (response.headers.get("content-type") ?? "").split(";")[0]?.trim();
        return {
          buffer: Buffer.concat(chunks),
          // The server's own type wins when it says anything specific; a generic
          // octet-stream tells us nothing the filename does not.
          ...(header && header !== "application/octet-stream" ? { contentType: header } : {}),
        };
      },
    },
  };
}

export function prepareBase64Source(content: string, overrideName?: string): SourcePreparation {
  const cleaned = content.replace(/^data:[^,]*,/, "").replace(/\s+/g, "");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(cleaned) || cleaned.length % 4 !== 0) {
    return { ok: false, message: "content_base64 is not valid base64." };
  }
  const buffer = Buffer.from(cleaned, "base64");
  if (buffer.length === 0) return { ok: false, message: "content_base64 decoded to zero bytes." };
  if (buffer.length > BASE64_MAX) {
    return {
      ok: false,
      message:
        `content_base64 decodes to ${formatSize(buffer.length)}, over the 3 MB limit for inline ` +
        "content. Use url for anything larger (up to 25 MB).",
    };
  }
  return {
    ok: true,
    source: { name: overrideName ?? "attachment", read: async () => ({ buffer }) },
  };
}
