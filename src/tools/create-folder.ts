import { z } from "zod";
import { GraphError, callGraphServer } from "../core/graph.js";
import {
  ToolResult,
  errorResult,
  fetchPaged,
  isNotFound,
  runTool,
  textResult,
} from "./common.js";

export const createFolderSchema = {
  name: z
    .string()
    .min(1)
    .max(255)
    .describe("Display name for the new folder."),
  parent_folder: z
    .string()
    .optional()
    .describe(
      'Where to create it: a well-known name ("inbox", "archive", "deleteditems", …) or a folder id from list_folders. Omit to create a top-level folder at the mailbox root.'
    ),
};

const createFolderArgs = z.object(createFolderSchema);

export const createFolderDescription =
  "Create a mail folder, either at the mailbox root (default) or as a subfolder of parent_folder. Returns the new folder's id, usable straight away with manage_message's move action, search_mail's folder input, and manage_rules' move_to_folder. Fails with the existing folder's id if a folder of that name already exists at the same level — folder names must be unique among siblings.";

const FOLDER_CAP = 500;

export async function createFolderHandler(
  input: z.input<typeof createFolderArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { name, parent_folder } = createFolderArgs.parse(input);
    const trimmed = name.trim();
    if (!trimmed) return errorResult("name cannot be blank.");

    // Resolve the parent first so a bad parent is reported as such, rather than
    // surfacing as a confusing 404 on the create call.
    let parentLabel = "the mailbox root";
    let childrenPath = "/me/mailFolders";
    let createPath = "/me/mailFolders";
    if (parent_folder) {
      let parent: any;
      try {
        parent = await callGraphServer(
          `/me/mailFolders/${encodeURIComponent(parent_folder)}?$select=id,displayName`
        );
      } catch (err) {
        if (isNotFound(err)) {
          return errorResult(
            `parent_folder ${JSON.stringify(parent_folder)} does not exist — use a well-known name or a folder id from list_folders.`
          );
        }
        throw err;
      }
      parentLabel = `"${parent.displayName ?? parent_folder}"`;
      const encoded = encodeURIComponent(parent.id);
      childrenPath = `/me/mailFolders/${encoded}/childFolders`;
      createPath = childrenPath;
    }

    // Outlook treats sibling folder names as case-insensitively unique; check up
    // front so the failure names the folder that is already there.
    const siblings = await fetchPaged(
      `${childrenPath}?$select=id,displayName&$top=100`,
      FOLDER_CAP
    );
    const clash = siblings.find(
      (f: any) => String(f.displayName ?? "").toLowerCase() === trimmed.toLowerCase()
    );
    if (clash) {
      return errorResult(
        `A folder named "${clash.displayName}" already exists in ${parentLabel} (id: ${clash.id}). ` +
          "Use that folder, or pick a different name."
      );
    }

    let created: any;
    try {
      created = await callGraphServer(`${createPath}?$select=id,displayName`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName: trimmed }),
      });
    } catch (err) {
      // Lost a race, or Outlook considers the name a clash for a reason the
      // sibling scan missed (e.g. a hidden folder).
      if (err instanceof GraphError && /ErrorFolderExists/i.test(err.body)) {
        return errorResult(
          `A folder named "${trimmed}" already exists in ${parentLabel}. Use list_folders to find its id.`
        );
      }
      throw err;
    }

    return textResult(
      `Folder created.\n` +
        `Name: ${created.displayName}\n` +
        `Parent: ${parentLabel}\n` +
        `Folder id: ${created.id}`
    );
  });
}
