import { z } from "zod";
import { ToolResult, fetchPaged, runTool, structuredResult } from "./common.js";

export const listFoldersSchema = {
  // MCP clients reject empty schemas less gracefully than a harmless optional flag.
  include_hidden: z
    .boolean()
    .default(false)
    .describe("Also list folders Outlook normally hides (default false)."),
};

const listFoldersArgs = z.object(listFoldersSchema);

/** Permissive machine-readable folder tree; every field optional. */
export const listFoldersOutputSchema = {
  count: z.number().optional().describe("Top-level folder count."),
  folders: z
    .array(
      z.looseObject({
        id: z.string().optional(),
        name: z.string().optional(),
        unread: z.number().optional(),
        total: z.number().optional(),
        childFolderCount: z.number().optional(),
        children: z.array(z.looseObject({})).optional(),
      })
    )
    .optional(),
};

export const listFoldersDescription =
  "List the mailbox folder tree (top-level folders and one level of subfolders) with unread/total message counts and folder ids, as text and as structuredContent. Folder ids can be used with manage_message's move action, search_mail's folder input, and delete_folder.";

const FOLDER_CAP = 200;
const SELECT = "id,displayName,unreadItemCount,totalItemCount,childFolderCount";

export async function listFoldersHandler(
  input: z.input<typeof listFoldersArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { include_hidden } = listFoldersArgs.parse(input);
    const hiddenParam = include_hidden ? "&includeHiddenFolders=true" : "";
    const top = await fetchPaged(
      `/me/mailFolders?$select=${SELECT}&$top=100${hiddenParam}`,
      FOLDER_CAP
    );

    const lines: string[] = [];
    const structuredFolders: Record<string, unknown>[] = [];
    for (const folder of top) {
      lines.push(formatFolder(folder, 0));
      const entry = describeFolder(folder);
      structuredFolders.push(entry);
      if (folder.childFolderCount > 0) {
        const children = await fetchPaged(
          `/me/mailFolders/${encodeURIComponent(folder.id)}/childFolders?$select=${SELECT}&$top=100${hiddenParam}`,
          FOLDER_CAP
        );
        entry.children = children.map(describeFolder);
        for (const child of children) {
          lines.push(formatFolder(child, 1));
          if (child.childFolderCount > 0) {
            lines.push(`    … ${child.childFolderCount} deeper subfolder(s) not shown`);
          }
        }
      }
    }
    return structuredResult(`Mail folders (${top.length} top-level):\n\n${lines.join("\n")}`, {
      count: top.length,
      folders: structuredFolders,
    });
  });
}

function describeFolder(folder: any): Record<string, unknown> & { children?: unknown[] } {
  return {
    id: String(folder.id ?? ""),
    name: String(folder.displayName ?? ""),
    unread: Number(folder.unreadItemCount ?? 0),
    total: Number(folder.totalItemCount ?? 0),
    childFolderCount: Number(folder.childFolderCount ?? 0),
  };
}

function formatFolder(folder: any, depth: number): string {
  const indent = "  ".repeat(depth);
  return (
    `${indent}${folder.displayName} — ${folder.unreadItemCount ?? 0} unread / ${folder.totalItemCount ?? 0} total\n` +
    `${indent}  id: ${folder.id}`
  );
}
