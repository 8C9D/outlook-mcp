import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  getDriveItemById,
  getDriveItemByPath,
  itemDisplayPath,
  normalizeDrivePath,
} from "../core/drive.js";
import { ToolResult, errorResult, isNotFound, runTool, textResult } from "./common.js";

export const shareLinkSchema = {
  action: z
    .enum(["create", "list", "revoke"])
    .describe(
      "create a sharing link, list the item's existing sharing links (with their permission ids), or revoke one link by permission_id."
    ),
  path: z
    .string()
    .min(1)
    .optional()
    .describe('OneDrive path of the file or folder, e.g. "Documents/report.pdf".'),
  item_id: z
    .string()
    .min(1)
    .optional()
    .describe("The item's id (from search_files or list_folder), instead of path."),
  link_type: z
    .enum(["view", "edit"])
    .default("view")
    .describe("create only: view (read-only) or edit (anyone with the link can modify). Default view."),
  permission_id: z
    .string()
    .min(1)
    .optional()
    .describe("revoke only: the permission id of the link to revoke (from create or list)."),
};

const shareLinkArgs = z.object(shareLinkSchema);

export const shareLinkDescription =
  "Create, list, or revoke sharing links for a OneDrive file or folder. BE AWARE: a created link is ANONYMOUS — anyone who obtains it can open (view) or modify (edit) the item without signing in, until the link is revoked. Creating the same link type twice returns the same link. revoke deletes the link's permission so the URL stops working (a later create issues a different URL).";

/** The lines describing one sharing-link permission, for create and list output. */
function describeLinkPermission(perm: any): string {
  const roles = (perm?.roles ?? []).join(", ") || "(no roles)";
  return (
    `${perm?.link?.type ?? "?"} link (${roles}; scope: ${perm?.link?.scope ?? "?"})\n` +
    `  URL: ${perm?.link?.webUrl ?? "(none)"}\n` +
    `  permission id: ${perm?.id ?? "(none)"}`
  );
}

export async function shareLinkHandler(input: z.input<typeof shareLinkArgs>): Promise<ToolResult> {
  return runTool(async () => {
    const { action, path, item_id, link_type, permission_id } = shareLinkArgs.parse(input);
    if ((path === undefined) === (item_id === undefined)) {
      return errorResult("Give exactly one of path or item_id.");
    }

    let normalized = "";
    if (path !== undefined) {
      const result = normalizeDrivePath(path);
      if (!result.ok) return errorResult(result.message);
      normalized = result.path;
      if (normalized === "") return errorResult("Refusing to share the entire drive root.");
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
    if (item.root !== undefined) return errorResult("Refusing to share the entire drive root.");
    const displayPath = itemDisplayPath(item);

    if (action === "create") {
      const perm = await callGraphServer(`/me/drive/items/${encodeURIComponent(item.id)}/createLink`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: link_type, scope: "anonymous" }),
      });
      return textResult(
        `Sharing link created for ${displayPath}.\n` +
          describeLinkPermission(perm) +
          `\n\nANYONE who obtains this URL can ${link_type === "edit" ? "open AND MODIFY" : "open"} ` +
          "the item without signing in. Revoke it with share_link action: revoke and the permission id above."
      );
    }

    if (action === "list") {
      const perms = await callGraphServer(`/me/drive/items/${encodeURIComponent(item.id)}/permissions`);
      // A sharing link has link.type (view/edit). The OWNER permission on a
      // personal drive also carries a bare link object (no type) — verified
      // live — so filtering on link alone would miscount the owner as a link.
      const links = (perms?.value ?? []).filter((p: any) => p.link?.type);
      const others = (perms?.value ?? []).length - links.length;
      if (links.length === 0) {
        return textResult(
          `${displayPath} has no sharing links.` +
            (others > 0 ? `\n(${others} non-link permission(s), e.g. the owner, not shown.)` : "")
        );
      }
      return textResult(
        `Sharing links on ${displayPath} (${links.length}):\n\n` +
          links.map(describeLinkPermission).join("\n") +
          (others > 0 ? `\n\n(${others} non-link permission(s), e.g. the owner, not shown.)` : "")
      );
    }

    // revoke
    if (!permission_id) return errorResult("revoke needs permission_id (from create or list).");
    try {
      await callGraphServer(
        `/me/drive/items/${encodeURIComponent(item.id)}/permissions/${encodeURIComponent(permission_id)}`,
        { method: "DELETE" }
      );
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No permission ${permission_id} on ${displayPath} — list the item's links with action: list.`
        );
      }
      throw err;
    }
    return textResult(
      `Sharing link revoked on ${displayPath}.\n` +
        "The URL stops working for everyone who has it. Creating a new link later issues a different URL."
    );
  });
}
