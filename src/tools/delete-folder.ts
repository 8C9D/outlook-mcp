import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, fetchPaged, isNotFound, runTool, textResult } from "./common.js";

export const deleteFolderSchema = {
  folder: z
    .string()
    .min(1)
    .describe(
      "The folder to delete: a folder id from list_folders (preferred), or a name Graph resolves. Well-known system folders are always refused."
    ),
  force: z
    .boolean()
    .default(false)
    .describe(
      "Delete a folder that still contains messages: they are moved into Deleted Items first (individually, so they are visible there), then the folder itself is moved to Deleted Items. Without force a non-empty folder is refused. A folder with subfolders is refused even with force — delete the subfolders first."
    ),
};

const deleteFolderArgs = z.object(deleteFolderSchema);

export const deleteFolderDescription =
  "Soft-delete a mail folder by MOVING it into Deleted Items, where it stays recoverable. (Verified live on this personal account: Graph's actual DELETE on a folder permanently destroys the folder AND every message in it with no Deleted Items copy, so this tool never issues it — the move is the only honest soft delete.) Guards: well-known folders (Inbox, Drafts, Sent Items, Deleted Items, Junk, Archive, Outbox, Conversation History, …) are always refused; a folder that still contains messages is refused unless force is set, in which case its messages are moved to Deleted Items first and the result says so; a folder with subfolders is always refused. To undo, move the folder back out of Deleted Items with Outlook; permanent removal must be done from Outlook deliberately.";

/**
 * Graph well-known folder names that may never be deleted. Resolved to ids at
 * call time and compared by id, because consumer mailFolder objects expose no
 * wellKnownName property (verified live: $select=wellKnownName → 400).
 */
export const WELL_KNOWN_FOLDERS = [
  "inbox",
  "drafts",
  "sentitems",
  "deleteditems",
  "junkemail",
  "archive",
  "outbox",
  "conversationhistory",
  "clutter",
  "syncissues",
  "scheduled",
] as const;

/** With force, at most this many messages are moved out before the delete. */
export const FORCE_MOVE_CAP = 500;

export type FolderFacts = {
  displayName: string;
  totalItemCount: number;
  childFolderCount: number;
  wellKnown: boolean;
};

/**
 * The guard logic, pure so the offline tier can cover the whole matrix.
 * Returns the refusal text, or null when the delete may proceed.
 */
export function deleteFolderRefusal(folder: FolderFacts, force: boolean): string | null {
  if (folder.wellKnown) {
    return (
      `"${folder.displayName}" is a well-known system folder and cannot be deleted. ` +
      "Only user-created folders can be."
    );
  }
  if (folder.childFolderCount > 0) {
    return (
      `"${folder.displayName}" contains ${folder.childFolderCount} subfolder(s). Delete the ` +
      "subfolders first (list_folders shows them) — this tool never takes a folder tree in one call."
    );
  }
  if (folder.totalItemCount > 0 && !force) {
    return (
      `"${folder.displayName}" still contains ${folder.totalItemCount} message(s). Move them out ` +
      "first, or pass force: true to have them moved into Deleted Items before the folder is deleted."
    );
  }
  if (folder.totalItemCount > FORCE_MOVE_CAP) {
    return (
      `"${folder.displayName}" contains ${folder.totalItemCount} message(s), more than the ` +
      `${FORCE_MOVE_CAP} this tool will move in one call. Empty it from Outlook instead.`
    );
  }
  return null;
}

/** Move up to 20 messages per Graph $batch; returns how many moves succeeded. */
async function moveMessagesToDeletedItems(messageIds: string[]): Promise<number> {
  let moved = 0;
  for (let start = 0; start < messageIds.length; start += 20) {
    const chunk = messageIds.slice(start, start + 20);
    const result = await callGraphServer("/$batch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: chunk.map((id, i) => ({
          id: String(i),
          method: "POST",
          url: `/me/messages/${encodeURIComponent(id)}/move`,
          headers: { "Content-Type": "application/json" },
          body: { destinationId: "deleteditems" },
        })),
      }),
    });
    for (const response of result?.responses ?? []) {
      if (response.status >= 200 && response.status < 300) moved++;
    }
  }
  return moved;
}

export async function deleteFolderHandler(
  input: z.input<typeof deleteFolderArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { folder: folderInput, force } = deleteFolderArgs.parse(input);

    let folder: any;
    try {
      folder = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(folderInput)}?$select=id,displayName,totalItemCount,childFolderCount`
      );
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No folder ${JSON.stringify(folderInput)} — use a folder id from list_folders.`
        );
      }
      throw err;
    }

    // Well-known folders are matched by ID: the same folder is reachable under
    // its well-known name, its id, or a localized display name, and only the
    // id is unambiguous. One $batch resolves all eleven names in one request
    // (eleven parallel GETs draw per-mailbox 429s); a name a consumer account
    // does not have simply 404s inside the batch.
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
    const wellKnownIds = new Set(
      (batch?.responses ?? [])
        .filter((r: any) => r.status >= 200 && r.status < 300)
        .map((r: any) => r.body?.id)
        .filter((id: unknown): id is string => typeof id === "string")
    );

    const facts: FolderFacts = {
      displayName: String(folder.displayName ?? folderInput),
      totalItemCount: Number(folder.totalItemCount ?? 0),
      childFolderCount: Number(folder.childFolderCount ?? 0),
      wellKnown: wellKnownIds.has(folder.id),
    };
    const refusal = deleteFolderRefusal(facts, force);
    if (refusal) return errorResult(refusal);

    // Force path: the folder's messages go to Deleted Items FIRST — visible
    // there individually, not buried inside a deleted subfolder.
    let movedOut = 0;
    if (force && facts.totalItemCount > 0) {
      const messages = await fetchPaged(
        `/me/mailFolders/${encodeURIComponent(folder.id)}/messages?$select=id&$top=100`,
        FORCE_MOVE_CAP
      );
      movedOut = await moveMessagesToDeletedItems(messages.map((m: any) => String(m.id)));
      const remaining = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(folder.id)}?$select=totalItemCount`
      );
      if (Number(remaining?.totalItemCount ?? 0) > 0) {
        return errorResult(
          `Moved ${movedOut} message(s) from "${facts.displayName}" to Deleted Items, but ` +
            `${remaining.totalItemCount} message(s) could not be moved. The folder was NOT deleted; ` +
            "retry, or empty it from Outlook."
        );
      }
    }

    // The soft delete itself: move the folder under Deleted Items.
    await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/move`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: "deleteditems" }),
    });

    return textResult(
      [
        `Folder "${facts.displayName}" moved to Deleted Items (soft delete — it can be moved back out).`,
        ...(movedOut > 0
          ? [
              `Its ${movedOut} message(s) were moved into Deleted Items first, so they sit there ` +
                "individually rather than inside the deleted folder.",
            ]
          : []),
        "Permanent removal, if wanted, must be done from Outlook (empty Deleted Items).",
      ].join("\n")
    );
  });
}
