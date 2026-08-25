// Mail-folder rename and move — the organizing acts between create_folder and
// delete_folder. Verified live: PATCH displayName renames in place, and
// POST /move keeps the folder's id (unlike a moved message, whose id changes).
import { z } from "zod";
import { GraphError, callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, isNotFound, runTool, textResult } from "./common.js";
import { WELL_KNOWN_FOLDERS } from "./delete-folder.js";

export const manageFolderSchema = {
  action: z
    .enum(["rename", "move"])
    .describe(
      "rename: change the folder's display name; move: make it a subfolder of destination_folder (its messages and subfolders travel with it)."
    ),
  folder: z
    .string()
    .min(1)
    .describe("The folder to act on: a folder id from list_folders (preferred), or a name Graph resolves."),
  new_name: z
    .string()
    .min(1)
    .max(255)
    .optional()
    .describe('rename: the new display name. Required for action "rename".'),
  destination_folder: z
    .string()
    .min(1)
    .optional()
    .describe(
      'move: where the folder should go — a well-known name ("inbox", "archive", …), a folder id from list_folders, or "root" for the mailbox root. Required for action "move".'
    ),
};

const manageFolderArgs = z.object(manageFolderSchema);

export const manageFolderDescription =
  "Rename a mail folder, or move it (with everything in it) under another folder or to the mailbox root. The folder's id stays the same through both, so saved rules and references keep working. Well-known system folders (Inbox, Drafts, Sent Items, …) are refused; so is a rename or move that would collide with a sibling folder of the same name. Creating is create_folder; deleting is delete_folder.";

/** The ids of the mailbox's well-known folders, resolved in one $batch. */
export async function fetchWellKnownFolderIds(): Promise<Set<string>> {
  const batch = await callGraphServer("/$batch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      requests: WELL_KNOWN_FOLDERS.map((name) => ({
        id: name,
        method: "GET",
        url: `/me/mailFolders/${name}?$select=id`,
      })),
    }),
  });
  return new Set(
    (batch?.responses ?? [])
      .filter((r: any) => r.status >= 200 && r.status < 300)
      .map((r: any) => r.body?.id)
      .filter((id: unknown): id is string => typeof id === "string")
  );
}

export async function manageFolderHandler(
  input: z.input<typeof manageFolderArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, folder: folderInput, new_name, destination_folder } =
      manageFolderArgs.parse(input);

    let folder: any;
    try {
      folder = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(folderInput)}?$select=id,displayName,totalItemCount,childFolderCount`
      );
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(`No folder ${JSON.stringify(folderInput)} — use a folder id from list_folders.`);
      }
      throw err;
    }
    const wellKnownIds = await fetchWellKnownFolderIds();
    if (wellKnownIds.has(folder.id)) {
      return errorResult(
        `"${folder.displayName}" is a well-known system folder and cannot be renamed or moved. ` +
          "Only user-created folders can be."
      );
    }

    if (action === "rename") {
      if (!new_name) return errorResult('Action "rename" requires new_name.');
      let renamed: any;
      try {
        renamed = await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ displayName: new_name.trim() }),
        });
      } catch (err) {
        // A rename collision comes back as ErrorFolderExists or (verified
        // live) as ErrorFolderSavePropertyError — Graph's phrasing for "the
        // displayName you PATCHed is not acceptable", which for a valid string
        // means a sibling already holds it.
        if (
          err instanceof GraphError &&
          /ErrorFolderExists|ErrorFolderSavePropertyError/i.test(err.body)
        ) {
          return errorResult(
            `Could not rename to "${new_name.trim()}" — folder names must be unique among siblings, and a sibling folder most likely already carries that name.`
          );
        }
        throw err;
      }
      return textResult(
        `Folder renamed from "${folder.displayName}" to "${renamed.displayName}". ` +
          "Its messages, subfolders, and id are unchanged.\n" +
          `Folder id: ${renamed.id}`
      );
    }

    if (!destination_folder) return errorResult('Action "move" requires destination_folder.');
    let destinationLabel = "the mailbox root";
    let destinationId = "msgfolderroot";
    if (destination_folder.toLowerCase() !== "root") {
      let destination: any;
      try {
        destination = await callGraphServer(
          `/me/mailFolders/${encodeURIComponent(destination_folder)}?$select=id,displayName`
        );
      } catch (err) {
        if (isNotFound(err)) {
          return errorResult(
            `destination_folder ${JSON.stringify(destination_folder)} does not exist — use a well-known name, a folder id from list_folders, or "root".`
          );
        }
        throw err;
      }
      if (destination.id === folder.id) {
        return errorResult("A folder cannot be moved into itself.");
      }
      destinationLabel = `"${destination.displayName ?? destination_folder}"`;
      destinationId = destination.id;
    }

    let moved: any;
    try {
      moved = await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/move`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ destinationId }),
      });
    } catch (err) {
      if (err instanceof GraphError && /ErrorFolderExists/i.test(err.body)) {
        return errorResult(
          `${destinationLabel} already contains a folder named "${folder.displayName}" — rename one of them first.`
        );
      }
      throw err;
    }
    return textResult(
      `Folder "${folder.displayName}" moved under ${destinationLabel}` +
        ` with its ${folder.totalItemCount ?? 0} message(s) and ${folder.childFolderCount ?? 0} subfolder(s).\n` +
        `Folder id: ${moved?.id ?? folder.id} (unchanged — folder ids survive moves)`
    );
  });
}
