// Where the local stdio server puts files it hands back: attachments from
// get_attachment and .eml exports from export_message. The hosted Worker has no
// filesystem and uses core/downloads.ts instead, so nothing here runs there —
// but the module is still imported into the Worker bundle, which is why it
// touches node:fs only inside the functions.
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

/** Local download directory, shared by every tool that saves a file. */
export const SAVE_DIR = path.join(os.homedir(), "Downloads", "outlook-mcp-attachments");

/** Strip path separators/control chars so a Graph-supplied name cannot escape SAVE_DIR. */
export function sanitizeFilename(name: string): string {
  const cleaned = name.replace(/[/\\:\0-\x1f]/g, "_").replace(/^\.+/, "_").trim();
  return cleaned || "attachment";
}

/** Return a path in SAVE_DIR that does not collide with an existing file. */
export async function collisionFreePath(filename: string): Promise<string> {
  const ext = path.extname(filename);
  const stem = filename.slice(0, filename.length - ext.length);
  for (let i = 0; ; i++) {
    const candidate = path.join(SAVE_DIR, i === 0 ? filename : `${stem} (${i})${ext}`);
    try {
      await fs.access(candidate);
    } catch {
      return candidate;
    }
  }
}

/** Write bytes into SAVE_DIR under a safe, non-colliding name; return the path. */
export async function saveToDownloads(
  filename: string,
  bytes: Uint8Array | Buffer
): Promise<string> {
  await fs.mkdir(SAVE_DIR, { recursive: true });
  const savePath = await collisionFreePath(sanitizeFilename(filename));
  await fs.writeFile(savePath, bytes);
  return savePath;
}
