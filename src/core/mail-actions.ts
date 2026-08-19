// The ONLY mailbox surface the LLM classifier can reach.
//
// This module is the structural half of the injection defence. core/classifier.js
// imports no Graph transport at all — it is handed a ClassifierMailbox and can
// therefore do exactly five things: list folders, list categories, read one
// message, move a message, set a message's categories. Send, delete, reply,
// forward, rule creation and settings changes are not expressible through this
// interface, so no prompt the model ever sees can produce them.
//
// Two tests hold that line: an import-graph assertion over core/classifier.js,
// and a source scan of this file for mutating Graph verbs that do not belong.
// Keep this module small enough that both stay easy to read.
import { callGraphServer } from "./graph.js";
import { isNeverFileFolder } from "./auto-filing.js";
import type { ClassifierMailbox, FilingFolder } from "./classifier.js";

/** How many folders may be offered to the model (and so how long the prompt is). */
const FOLDER_CAP = 60;

const MESSAGE_SELECT = "id,subject,from,receivedDateTime,bodyPreview,categories,parentFolderId";

/** The live implementation, against whatever Graph token is in scope. */
export function graphClassifierMailbox(): ClassifierMailbox {
  return {
    async listFilingFolders() {
      // One level deep, same shape list_folders reports, minus the folders a
      // classifier must never target (see NEVER_FILE_INTO).
      const top = await collect("/me/mailFolders?$select=id,displayName,childFolderCount&$top=100");
      const folders: FilingFolder[] = [];
      for (const folder of top) {
        if (!folder?.id || !folder?.displayName) continue;
        if (!isNeverFileFolder(folder.displayName)) {
          folders.push({ id: folder.id, displayName: folder.displayName });
        }
        if (folders.length >= FOLDER_CAP) break;
        if (folder.childFolderCount > 0) {
          const children = await collect(
            `/me/mailFolders/${encodeURIComponent(folder.id)}/childFolders?$select=id,displayName&$top=100`
          );
          for (const child of children) {
            if (!child?.id || !child?.displayName) continue;
            if (isNeverFileFolder(child.displayName)) continue;
            folders.push({ id: child.id, displayName: `${folder.displayName}/${child.displayName}` });
            if (folders.length >= FOLDER_CAP) break;
          }
        }
        if (folders.length >= FOLDER_CAP) break;
      }
      return folders.slice(0, FOLDER_CAP);
    },

    async listCategories() {
      const data = await callGraphServer("/me/outlook/masterCategories?$top=100");
      return (data?.value ?? [])
        .map((category: any) => String(category?.displayName ?? ""))
        .filter((name: string) => name !== "");
    },

    async readMessage(messageId) {
      const message = await callGraphServer(
        `/me/messages/${encodeURIComponent(messageId)}?$select=${MESSAGE_SELECT}`
      ).catch(() => null);
      if (!message?.id) return null;
      const address = message?.from?.emailAddress;
      const from = address?.address
        ? address.name && address.name !== address.address
          ? `${address.name} <${address.address}>`
          : String(address.address)
        : "(unknown sender)";
      return {
        id: String(message.id),
        subject: String(message.subject ?? ""),
        from,
        receivedDateTime: message.receivedDateTime ?? undefined,
        bodyPreview: String(message.bodyPreview ?? ""),
        categories: Array.isArray(message.categories) ? message.categories.map(String) : [],
        parentFolderId: message.parentFolderId ?? undefined,
      };
    },

    async move(messageId, folderId) {
      const moved = await callGraphServer(`/me/messages/${encodeURIComponent(messageId)}/move`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ destinationId: folderId }),
      });
      return String(moved?.id ?? messageId);
    },

    async categorize(messageId, categories) {
      await callGraphServer(`/me/messages/${encodeURIComponent(messageId)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ categories }),
      });
    },
  };
}

/** GET a Graph collection, following @odata.nextLink. */
async function collect(firstPath: string): Promise<any[]> {
  const items: any[] = [];
  let next: string | undefined = firstPath;
  while (next && items.length < 500) {
    const page: any = await callGraphServer(next);
    items.push(...(page?.value ?? []));
    next = page?.["@odata.nextLink"];
  }
  return items;
}
