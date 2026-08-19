import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  downloadDriveItem,
  getDriveItemByPath,
  itemDisplayPath,
  normalizeDrivePath,
} from "../core/drive.js";
import { getStateStore } from "../core/state.js";
import {
  MAX_FILE_SIZE,
  SourcePreparation,
  SourceReadError,
  contentTypeForFile,
  prepareBase64Source,
  prepareFileSource,
  prepareUrlSource,
} from "./file-sources.js";
import { ToolResult, errorResult, isNotFound, runTool, textResult } from "./common.js";
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
      "Source 1 of 4: absolute local path of a file on the machine running this server (e.g. /Users/me/report.pdf). Works only on the local (stdio) server — the hosted server has no filesystem and rejects it."
    ),
  url: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 2 of 4: an https:// URL the server downloads and attaches (max 25 MB). The URL must be reachable without credentials."
    ),
  content_base64: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Source 3 of 4: the file's bytes, base64-encoded, for content you already hold (max 3 MB decoded). Give attachment_name too, so the file arrives with a sensible name and type."
    ),
  onedrive_path: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Source 4 of 4: the OneDrive path of a file to attach (e.g. "Documents/report.pdf") — the file\'s bytes are copied into the draft; the OneDrive file is not modified. Works on both servers.'
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
  "Attach a file to an existing email draft (max 25 MB) from exactly one of four sources: file_path (a local file — local stdio server only), url (an https link this server fetches), content_base64 (bytes you supply, max 3 MB), or onedrive_path (a file in OneDrive, attached as a copy of its bytes). Natural flow: create_draft → add_attachment (once per file) → send_draft. Fails if the id is not a draft, no source or more than one is given, or the file is missing/oversized. This tool never sends — the draft stays in Drafts until send_draft is called.";

const SMALL_LIMIT = 3 * 1024 * 1024; // single-POST fileAttachment below this
const CHUNK_SIZE = 4 * 1024 * 1024; // upload-session chunk size

export { contentTypeForFile };

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
 * Pre-flight a OneDrive source: resolve the item and check its size BEFORE the
 * draft check, exactly as the other sources are pre-flighted; the bytes are
 * only downloaded after the draft is verified.
 */
async function prepareOneDriveSource(
  rawPath: string,
  overrideName?: string
): Promise<SourcePreparation> {
  const normalized = normalizeDrivePath(rawPath);
  if (!normalized.ok) return { ok: false, message: normalized.message };
  if (normalized.path === "") {
    return { ok: false, message: "onedrive_path names the drive root, not a file." };
  }
  let item: any;
  try {
    item = await getDriveItemByPath(normalized.path);
  } catch (err) {
    if (isNotFound(err)) {
      return {
        ok: false,
        message: `No OneDrive file at ${JSON.stringify(normalized.path)}. Find it with search_files or list_folder.`,
      };
    }
    throw err;
  }
  if (item.folder !== undefined) {
    return {
      ok: false,
      message: `${itemDisplayPath(item)} is a folder — attachments must be single files.`,
    };
  }
  const size = Number(item.size ?? 0);
  if (size > MAX_FILE_SIZE) {
    return {
      ok: false,
      message:
        `${itemDisplayPath(item)} is ${formatSize(size)}, over this tool's 25 MB attachment cap. ` +
        "Share it as a link instead (share_link) and paste the URL into the draft body.",
    };
  }
  return {
    ok: true,
    source: {
      name: overrideName ?? item.name ?? "attachment",
      read: async () => {
        const { bytes } = await downloadDriveItem(item.id);
        return {
          buffer: Buffer.from(bytes),
          ...(item.file?.mimeType ? { contentType: item.file.mimeType } : {}),
        };
      },
    },
  };
}

export async function addAttachmentHandler(
  input: z.input<typeof addAttachmentArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { draft_id, file_path, url, content_base64, onedrive_path, attachment_name } =
      addAttachmentArgs.parse(input);

    const given = [
      ["file_path", file_path],
      ["url", url],
      ["content_base64", content_base64],
      ["onedrive_path", onedrive_path],
    ].filter(([, value]) => value !== undefined);
    if (given.length !== 1) {
      return errorResult(
        given.length === 0
          ? "Give exactly one source: file_path (local server only), url, content_base64, or onedrive_path."
          : `Give exactly one source, not ${given.length} (${given.map(([n]) => n).join(", ")}).`
      );
    }
    if (file_path !== undefined && getStateStore()?.mode === "remote") {
      return errorResult(
        "file_path works only on the local (stdio) server: this hosted server has no access to " +
          "your filesystem. Attach the file with url (an https link, up to 25 MB), " +
          "content_base64 (bytes inline, up to 3 MB), or onedrive_path instead."
      );
    }

    const prepared =
      file_path !== undefined
        ? await prepareFileSource(file_path, attachment_name)
        : url !== undefined
          ? await prepareUrlSource(url, attachment_name)
          : content_base64 !== undefined
            ? prepareBase64Source(content_base64, attachment_name)
            : await prepareOneDriveSource(onedrive_path!, attachment_name);
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
      if (err instanceof SourceReadError) return errorResult(err.message);
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
