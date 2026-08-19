import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  drivePathUrl,
  getDriveItemById,
  getDriveItemByPath,
  itemDisplayPath,
  normalizeDrivePath,
} from "../core/drive.js";
import { GraphError } from "../core/graph.js";
import { ToolResult, errorResult, isNotFound, runTool, textResult } from "./common.js";

export const manageFileSchema = {
  action: z
    .enum(["move", "rename", "delete"])
    .describe(
      "move (to another folder, keeping the name), rename (in place), or delete (into the OneDrive recycle bin, recoverable from onedrive.live.com)."
    ),
  path: z
    .string()
    .min(1)
    .optional()
    .describe('OneDrive path of the file or folder to act on, e.g. "Documents/old.txt".'),
  item_id: z
    .string()
    .min(1)
    .optional()
    .describe("The item's id (from search_files or list_folder), instead of path."),
  destination_folder: z
    .string()
    .optional()
    .describe('move only: the folder to move into, e.g. "Documents/Archive" ("" or "/" is the drive root). Must already exist.'),
  new_name: z
    .string()
    .min(1)
    .optional()
    .describe('rename only: the new name, e.g. "report-final.pdf".'),
};

const manageFileArgs = z.object(manageFileSchema);

export const manageFileDescription =
  "Move, rename, or delete one OneDrive file or folder (by path or item id). delete is soft: the item goes to the OneDrive recycle bin, where it can be restored for ~30 days. Moving or renaming keeps the item id stable; a name collision at the destination is refused rather than overwriting anything.";

export async function manageFileHandler(
  input: z.input<typeof manageFileArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, path, item_id, destination_folder, new_name } = manageFileArgs.parse(input);
    if ((path === undefined) === (item_id === undefined)) {
      return errorResult("Give exactly one of path or item_id.");
    }

    let normalized = "";
    if (path !== undefined) {
      const result = normalizeDrivePath(path);
      if (!result.ok) return errorResult(result.message);
      normalized = result.path;
      if (normalized === "") return errorResult("The drive root itself cannot be moved, renamed or deleted.");
    }

    let item: any;
    try {
      item = item_id !== undefined ? await getDriveItemById(item_id) : await getDriveItemByPath(normalized);
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          item_id !== undefined
            ? `No OneDrive item with id ${item_id}.`
            : `No OneDrive item at path ${JSON.stringify(normalized)}. Find it with search_files or list_folder.`
        );
      }
      throw err;
    }
    if (item.root !== undefined) {
      return errorResult("The drive root itself cannot be moved, renamed or deleted.");
    }
    const displayPath = itemDisplayPath(item);
    const kind = item.folder !== undefined ? "folder" : "file";

    if (action === "delete") {
      await callGraphServer(`/me/drive/items/${encodeURIComponent(item.id)}`, { method: "DELETE" });
      return textResult(
        `Deleted ${kind} ${displayPath} into the OneDrive recycle bin.\n` +
          "It can be restored from onedrive.live.com → Recycle bin (kept ~30 days)."
      );
    }

    if (action === "rename") {
      if (!new_name) return errorResult("rename needs new_name.");
      if (/[/\\]/.test(new_name)) {
        return errorResult("new_name is a single name, not a path — use move to change folders.");
      }
      let renamed: any;
      try {
        renamed = await callGraphServer(`/me/drive/items/${encodeURIComponent(item.id)}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ name: new_name }),
        });
      } catch (err) {
        if (err instanceof GraphError && /nameAlreadyExists/i.test(err.body)) {
          return errorResult(
            `A sibling named ${JSON.stringify(new_name)} already exists — nothing was renamed or overwritten.`
          );
        }
        throw err;
      }
      return textResult(
        `Renamed ${kind} ${displayPath} → ${renamed?.name ?? new_name}.\nid (unchanged): ${item.id}`
      );
    }

    // move
    if (destination_folder === undefined) {
      return errorResult('move needs destination_folder ("" or "/" for the drive root).');
    }
    const dest = normalizeDrivePath(destination_folder);
    if (!dest.ok) return errorResult(dest.message);
    let destFolder: any;
    try {
      destFolder = await getDriveItemByPath(dest.path);
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No OneDrive folder at ${JSON.stringify(dest.path || "/")} — create it first (upload_file creates folders, or create one in OneDrive) or check the path with list_folder.`
        );
      }
      throw err;
    }
    if (destFolder.folder === undefined) {
      return errorResult(`${itemDisplayPath(destFolder)} is a file — destination_folder must be a folder.`);
    }
    let moved: any;
    try {
      moved = await callGraphServer(`/me/drive/items/${encodeURIComponent(item.id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ parentReference: { id: destFolder.id } }),
      });
    } catch (err) {
      if (err instanceof GraphError && /nameAlreadyExists/i.test(err.body)) {
        return errorResult(
          `${JSON.stringify(item.name)} already exists in ${itemDisplayPath(destFolder)} — nothing was moved or overwritten.`
        );
      }
      throw err;
    }
    return textResult(
      `Moved ${kind} ${displayPath} → ${itemDisplayPath(moved ?? { name: item.name, parentReference: { path: `/drive/root:${dest.path ? `/${dest.path}` : ""}` } })}.\n` +
        `id (unchanged): ${item.id}`
    );
  });
}
