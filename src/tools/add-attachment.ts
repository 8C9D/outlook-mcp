import { promises as fs } from "node:fs";
import path from "node:path";
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { getStateStore } from "../core/state.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatSize } from "./read-message.js";

export const addAttachmentSchema = {
  draft_id: z
    .string()
    .min(1)
    .describe("The id of the draft to attach the file to (from create_draft)."),
  file_path: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 1 of 3: absolute local path of a file on the machine running this server (e.g. /Users/me/report.pdf). Works only on the local (stdio) server — the hosted server has no filesystem and rejects it."
    ),
  url: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 2 of 3: an https:// URL the server downloads and attaches (max 25 MB). The URL must be reachable without credentials."
    ),
  content_base64: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 3 of 3: the file's bytes, base64-encoded, for content you already hold (max 3 MB decoded). Give attachment_name too, so the file arrives with a sensible name and type."
    ),
  attachment_name: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Name the attachment should carry in the email. Defaults to the file's basename or the last segment of the URL; required in practice for content_base64."
    ),
};

const addAttachmentArgs = z.object(addAttachmentSchema);

export const addAttachmentDescription =
  "Attach a file to an existing email draft (max 25 MB) from exactly one of three sources: file_path (a local file — local stdio server only), url (an https link this server fetches), or content_base64 (bytes you supply, max 3 MB). Natural flow: create_draft → add_attachment (once per file) → send_draft. Fails if the id is not a draft, no source or more than one is given, or the file is missing/oversized. This tool never sends — the draft stays in Drafts until send_draft is called.";

const SMALL_LIMIT = 3 * 1024 * 1024; // single-POST fileAttachment below this
const MAX_SIZE = 25 * 1024 * 1024; // hard cap for this tool
const BASE64_MAX = 3 * 1024 * 1024; // cap for inline content_base64 (decoded)
const CHUNK_SIZE = 4 * 1024 * 1024; // upload-session chunk size

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
 * Upload a 3–25 MB file via an Outlook attachment upload session: 4 MB PUT
 * chunks against the session's uploadUrl. The uploadUrl carries its own auth
 * token, so chunk PUTs are plain fetches — NOT callGraphServer (whose bearer
 * token is for graph.microsoft.com, a different audience).
 */
async function uploadViaSession(
  draftId: string,
  name: string,
  contentType: string,
  buffer: Buffer
): Promise<void> {
  const session = await callGraphServer(
    `/me/messages/${encodeURIComponent(draftId)}/attachments/createUploadSession`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        AttachmentItem: { attachmentType: "file", name, contentType, size: buffer.length },
      }),
    }
  );
  const uploadUrl: string | undefined = session?.uploadUrl;
  if (!uploadUrl) throw new Error("Graph did not return an uploadUrl for the attachment session.");

  for (let offset = 0; offset < buffer.length; offset += CHUNK_SIZE) {
    const chunk = buffer.subarray(offset, Math.min(offset + CHUNK_SIZE, buffer.length));
    const range = `bytes ${offset}-${offset + chunk.length - 1}/${buffer.length}`;
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Length": String(chunk.length),
        "Content-Range": range,
      },
      body: new Uint8Array(chunk),
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(
        `Attachment upload failed at ${range}: HTTP ${response.status} ${response.statusText}\n${body.slice(0, 500)}`
      );
    }
  }
}

/**
 * A source that has passed every check cheap enough to make before the draft is
 * verified (existence, scheme, size where it is knowable). `read` produces the
 * bytes, and may sharpen the content type with what the transfer revealed.
 */
type Prepared = {
  name: string;
  read: () => Promise<{ buffer: Buffer; contentType?: string }>;
};

type Preparation = { ok: true; source: Prepared } | { ok: false; message: string };

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

async function prepareFile(filePath: string, attachmentName?: string): Promise<Preparation> {
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
  if (stat.size > MAX_SIZE) {
    return {
      ok: false,
      message:
        `File is ${formatSize(stat.size)}, over this tool's 25 MB attachment cap. ` +
        "Share large files another way (e.g. a cloud link pasted into the draft body).",
    };
  }
  return {
    ok: true,
    source: {
      name: attachmentName ?? path.basename(filePath),
      read: async () => ({ buffer: await fs.readFile(filePath) }),
    },
  };
}

/** A download that failed for a reason the caller can act on. */
class ToolFetchError extends Error {}

/**
 * Download the URL, refusing anything over the cap. The body is read chunk by
 * chunk and abandoned the moment it grows too large, so a server that lies
 * about (or omits) Content-Length cannot make this buffer 2 GB.
 */
async function prepareUrl(rawUrl: string, attachmentName?: string): Promise<Preparation> {
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
      name: attachmentName ?? nameFromUrl(url),
      read: async () => {
        const response = await fetch(url, { redirect: "follow" });
        if (!response.ok) {
          throw new ToolFetchError(
            `Could not download ${url.href}: HTTP ${response.status} ${response.statusText}`
          );
        }
        const declared = Number(response.headers.get("content-length") ?? "");
        if (Number.isFinite(declared) && declared > MAX_SIZE) {
          throw new ToolFetchError(
            `That URL is ${formatSize(declared)}, over this tool's 25 MB attachment cap.`
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
            if (total > MAX_SIZE) {
              await reader.cancel();
              throw new ToolFetchError(
                "That URL is over this tool's 25 MB attachment cap (the download was stopped)."
              );
            }
            chunks.push(value);
          }
        } else {
          const body = new Uint8Array(await response.arrayBuffer());
          if (body.length > MAX_SIZE) {
            throw new ToolFetchError(
              `That URL is ${formatSize(body.length)}, over this tool's 25 MB attachment cap.`
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

function prepareBase64(content: string, attachmentName?: string): Preparation {
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
    source: { name: attachmentName ?? "attachment", read: async () => ({ buffer }) },
  };
}

export async function addAttachmentHandler(
  input: z.input<typeof addAttachmentArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id, file_path, url, content_base64, attachment_name } =
      addAttachmentArgs.parse(input);

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
          "your filesystem. Attach the file with url (an https link, up to 25 MB) or " +
          "content_base64 (bytes inline, up to 3 MB) instead."
      );
    }

    const prepared =
      file_path !== undefined
        ? await prepareFile(file_path, attachment_name)
        : url !== undefined
          ? await prepareUrl(url, attachment_name)
          : prepareBase64(content_base64!, attachment_name);
    if (!prepared.ok) return errorResult(prepared.message);

    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(draft_id)}?$select=isDraft,subject`
    );
    if (!msg.isDraft) {
      return errorResult(
        `Message ${draft_id} is not a draft (subject: ${JSON.stringify(msg.subject ?? "")}) — attachments can only be added to drafts.`
      );
    }

    const name = prepared.source.name;
    let buffer: Buffer;
    let contentType: string;
    try {
      const loaded = await prepared.source.read();
      buffer = loaded.buffer;
      contentType = loaded.contentType ?? contentTypeForFile(name);
    } catch (err) {
      if (err instanceof ToolFetchError) return errorResult(err.message);
      throw err;
    }

    if (buffer.length < SMALL_LIMIT) {
      await callGraphServer(`/me/messages/${encodeURIComponent(draft_id)}/attachments`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name,
          contentType,
          contentBytes: buffer.toString("base64"),
        }),
      });
    } else {
      await uploadViaSession(draft_id, name, contentType, buffer);
    }

    return textResult(
      `Attachment added.\n` +
        `Name: ${name}\n` +
        `Type: ${contentType}\n` +
        `Size: ${formatSize(buffer.length)}\n` +
        `Draft subject: ${msg.subject || "(no subject)"}\n` +
        `Draft id: ${draft_id}\n` +
        "Still in Drafts — use send_draft to send it."
    );
  });
}
