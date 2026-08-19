import { z } from "zod";
import {
  DRIVE_ITEM_SELECT,
  LIST_FOLDER_CAP,
  compareDriveChildren,
  describeDriveItem,
  drivePathUrl,
  getDriveItemById,
  getDriveItemByPath,
  itemDisplayPath,
  normalizeDrivePath,
} from "../core/drive.js";
import {
  ToolResult,
  errorResult,
  fetchPaged,
  formatLocal,
  isNotFound,
  runTool,
  structuredResult,
} from "./common.js";
import { formatSize } from "./read-message.js";

export const listFolderSchema = {
  path: z
    .string()
    .optional()
    .describe(
      'OneDrive folder path to list, e.g. "Documents" or "Documents/Receipts". Omit both path and item_id for the drive root.'
    ),
  item_id: z
    .string()
    .min(1)
    .optional()
    .describe("A folder's item id (from search_files or an earlier listing), instead of path."),
};

const listFolderArgs = z.object(listFolderSchema);

/** Permissive machine-readable folder listing; every field optional. */
export const listFolderOutputSchema = {
  path: z.string().optional().describe("The listed folder's path ('/' is the root)."),
  id: z.string().optional(),
  folderCount: z.number().optional(),
  fileCount: z.number().optional(),
  items: z
    .array(
      z.looseObject({
        id: z.string().optional(),
        name: z.string().optional(),
        path: z.string().optional(),
        kind: z.string().optional(),
        size: z.number().optional(),
        childCount: z.number().optional(),
        mimeType: z.string().optional(),
        modified: z.string().optional(),
      })
    )
    .optional()
    .describe("Folders first, then files, each group alphabetical."),
};

export const listFolderDescription =
  "List one OneDrive folder's contents (folders first, then files, with sizes, modified dates and item ids) — by path, by item id, or the drive root when neither is given. Unlike search_files this reads the folder directly, so it always reflects the current state. For MAIL folders use list_folders instead.";

export async function listFolderHandler(
  input: z.input<typeof listFolderArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { path, item_id } = listFolderArgs.parse(input);
    if (path !== undefined && item_id !== undefined) {
      return errorResult("Give path OR item_id, not both (or neither, for the drive root).");
    }

    let normalized = "";
    if (path !== undefined) {
      const result = normalizeDrivePath(path);
      if (!result.ok) return errorResult(result.message);
      normalized = result.path;
    }

    let folder: any;
    try {
      folder = item_id !== undefined ? await getDriveItemById(item_id) : await getDriveItemByPath(normalized);
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          item_id !== undefined
            ? `No OneDrive item with id ${item_id}.`
            : `No OneDrive folder at path ${JSON.stringify(normalized || "/")}. Check it with list_folder on the parent, or search_files.`
        );
      }
      throw err;
    }
    if (folder.folder === undefined) {
      return errorResult(
        `${itemDisplayPath(folder)} is a file, not a folder — use read_file to read it.`
      );
    }

    const children = await fetchPaged(
      `/me/drive/items/${encodeURIComponent(folder.id)}/children?$select=${DRIVE_ITEM_SELECT}&$top=200`,
      LIST_FOLDER_CAP
    );
    children.sort(compareDriveChildren);

    const displayPath = itemDisplayPath(folder);
    const folderCount = children.filter((c) => c.folder !== undefined).length;
    const fileCount = children.length - folderCount;
    const structured = {
      path: displayPath,
      id: String(folder.id ?? ""),
      folderCount,
      fileCount,
      items: children.map(describeDriveItem),
    };

    if (children.length === 0) {
      return structuredResult(`OneDrive folder ${displayPath} is empty.\nid: ${folder.id}`, structured);
    }

    const truncated =
      children.length >= LIST_FOLDER_CAP ? `\n(only the first ${LIST_FOLDER_CAP} items are shown)` : "";
    const lines = children.map((item) => {
      if (item.folder !== undefined) {
        return `[folder] ${item.name} — ${item.folder?.childCount ?? 0} item(s)\n  id: ${item.id}`;
      }
      return (
        `${item.name} — ${formatSize(item.size)}, modified ${formatLocal(item.lastModifiedDateTime)}\n` +
        `  id: ${item.id}`
      );
    });
    return structuredResult(
      `OneDrive folder ${displayPath} — ${folderCount} folder(s), ${fileCount} file(s)\n` +
        `id: ${folder.id}\n\n${lines.join("\n")}${truncated}`,
      structured
    );
  });
}
