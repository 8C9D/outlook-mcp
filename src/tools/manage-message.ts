import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { fetchMasterCategories } from "./manage-categories.js";

export const manageMessageSchema = {
  message_ids: z
    .array(z.string().min(1))
    .min(1)
    .max(20)
    .describe("Ids of the messages to act on (1–20)."),
  action: z
    .enum([
      "move",
      "archive",
      "delete",
      "mark_read",
      "mark_unread",
      "flag",
      "unflag",
      "categorize",
    ])
    .describe(
      "move: to destination_folder; archive: to the Archive folder; delete: to Deleted Items (soft delete, recoverable); mark_read/mark_unread; flag/unflag: follow-up flag; categorize: set the messages' categories to `categories` (replaces any it already had)."
    ),
  destination_folder: z
    .string()
    .optional()
    .describe(
      'Target folder for action "move": a well-known name ("inbox", "archive", "deleteditems", "junkemail", "drafts") or a folder id from list_folders or create_folder.'
    ),
  categories: z
    .array(z.string().min(1))
    .optional()
    .describe(
      'Category names for action "categorize". These REPLACE each message\'s existing categories rather than adding to them, so include every category the message should end up with. Pass an empty array to strip all categories. Every name must already exist in the mailbox (see manage_categories).'
    ),
};

const manageMessageArgs = z.object(manageMessageSchema);

export const manageMessageDescription =
  "Act on up to 20 email messages: move, archive, delete (soft — moves to Deleted Items), mark read/unread, flag/unflag, or categorize. categorize REPLACES each message's categories with the list you pass (it does not append), so read the current categories first if you mean to add one; an empty list clears them. Returns a per-message success/failure list; moved messages get a NEW message id (reported in the output). Before calling with delete or move, state to the user exactly which messages (subjects/senders) will be affected. Deletion is never permanent from this tool — messages remain recoverable in Deleted Items.";

/** One Graph JSON-batch request item for a single message id. */
function batchRequest(
  itemId: string,
  messageId: string,
  action: string,
  destinationFolder: string | undefined,
  categories: string[] | undefined
): object {
  const base = `/me/messages/${encodeURIComponent(messageId)}`;
  const json = { "Content-Type": "application/json" };
  switch (action) {
    case "move":
      return {
        id: itemId,
        method: "POST",
        url: `${base}/move`,
        headers: json,
        body: { destinationId: destinationFolder },
      };
    case "archive":
      return {
        id: itemId,
        method: "POST",
        url: `${base}/move`,
        headers: json,
        body: { destinationId: "archive" },
      };
    case "delete":
      return { id: itemId, method: "DELETE", url: base };
    case "mark_read":
      return { id: itemId, method: "PATCH", url: base, headers: json, body: { isRead: true } };
    case "mark_unread":
      return { id: itemId, method: "PATCH", url: base, headers: json, body: { isRead: false } };
    case "flag":
      return {
        id: itemId,
        method: "PATCH",
        url: base,
        headers: json,
        body: { flag: { flagStatus: "flagged" } },
      };
    case "unflag":
      return {
        id: itemId,
        method: "PATCH",
        url: base,
        headers: json,
        body: { flag: { flagStatus: "notFlagged" } },
      };
    case "categorize":
      return { id: itemId, method: "PATCH", url: base, headers: json, body: { categories } };
    default:
      throw new Error(`Unknown action ${action}`);
  }
}

function describeSuccess(
  action: string,
  destinationFolder: string | undefined,
  categories: string[] | undefined,
  body: any
): string {
  const newId = body?.id as string | undefined;
  switch (action) {
    case "move":
      return `moved to ${destinationFolder}${newId ? ` (new id: ${newId})` : ""}`;
    case "archive":
      return `archived${newId ? ` (new id: ${newId})` : ""}`;
    case "delete":
      return "deleted (moved to Deleted Items)";
    case "mark_read":
      return "marked read";
    case "mark_unread":
      return "marked unread";
    case "flag":
      return "flagged";
    case "unflag":
      return "unflagged";
    case "categorize":
      return categories?.length
        ? `categories set to ${categories.map((c) => JSON.stringify(c)).join(", ")}`
        : "categories cleared";
    default:
      return "done";
  }
}

function describeFailure(status: number, body: any): string {
  const code = body?.error?.code;
  const message = body?.error?.message;
  return `HTTP ${status} ${code ?? ""}: ${message ?? "(no error detail from Graph)"}`;
}

/** POST one JSON batch and return its per-item responses keyed by item id. */
async function postBatch(requests: object[]): Promise<Map<string, any>> {
  const result = await callGraphServer("/$batch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ requests }),
  });
  const byId = new Map<string, any>();
  for (const response of result?.responses ?? []) byId.set(String(response.id), response);
  return byId;
}

export async function manageMessageHandler(
  input: z.input<typeof manageMessageArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_ids, action, destination_folder, categories } = manageMessageArgs.parse(input);
    if (action === "move" && !destination_folder) {
      return errorResult('Action "move" requires destination_folder.');
    }

    // Graph silently accepts unknown category names on a message, which would
    // leave colourless labels behind, so validate against the master list first
    // and write back the master list's exact spelling.
    let resolvedCategories = categories;
    if (action === "categorize") {
      if (!categories) {
        return errorResult(
          'Action "categorize" requires categories (an empty array clears all categories).'
        );
      }
      const master = await fetchMasterCategories();
      const byLower = new Map<string, string>(
        master.map((c: any) => [String(c.displayName ?? "").toLowerCase(), c.displayName])
      );
      const unknown = categories.filter((name) => !byLower.has(name.toLowerCase()));
      if (unknown.length > 0) {
        const available = master.map((c: any) => JSON.stringify(c.displayName)).join(", ");
        return errorResult(
          `Unknown categor(y/ies): ${unknown.map((n) => JSON.stringify(n)).join(", ")}. ` +
            `This mailbox has: ${available || "(none)"}. Create it with manage_categories first.`
        );
      }
      resolvedCategories = categories.map((name) => byLower.get(name.toLowerCase())!);
    }

    // All ids go out in ONE Graph JSON batch ($batch caps at 20 requests,
    // matching the schema's max). Each item succeeds or fails independently.
    const requests = message_ids.map((id, i) =>
      batchRequest(String(i), id, action, destination_folder, resolvedCategories)
    );
    let responses = await postBatch(requests);

    // Individual items can be throttled (429) even when the batch itself
    // succeeds. Retry just those items once, after the longest Retry-After.
    const throttled = message_ids
      .map((_, i) => String(i))
      .filter((itemId) => responses.get(itemId)?.status === 429);
    if (throttled.length > 0) {
      const waitSeconds = Math.min(
        Math.max(
          ...throttled.map((itemId) => {
            const headers = responses.get(itemId)?.headers ?? {};
            const key = Object.keys(headers).find((k) => k.toLowerCase() === "retry-after");
            const retryAfter = Number(key ? headers[key] : "2");
            return Number.isFinite(retryAfter) && retryAfter > 0 ? retryAfter : 2;
          })
        ),
        60
      );
      console.error(
        `Graph throttled ${throttled.length} batch item(s) (429); retrying once after ${waitSeconds}s.`
      );
      await new Promise((resolve) => setTimeout(resolve, waitSeconds * 1000));
      const retryResponses = await postBatch(
        throttled.map((itemId) =>
          batchRequest(
            itemId,
            message_ids[Number(itemId)]!,
            action,
            destination_folder,
            resolvedCategories
          )
        )
      );
      for (const [itemId, response] of retryResponses) responses.set(itemId, response);
    }

    const lines: string[] = [];
    let failures = 0;
    message_ids.forEach((id, i) => {
      const response = responses.get(String(i));
      if (!response) {
        failures++;
        lines.push(`FAILED  ${id}: no response for this item in the Graph batch reply`);
        return;
      }
      if (response.status >= 200 && response.status < 300) {
        lines.push(
          `OK      ${id}: ${describeSuccess(action, destination_folder, resolvedCategories, response.body)}`
        );
      } else if (response.status === 429) {
        failures++;
        lines.push(
          `FAILED  ${id}: still throttled (HTTP 429) after one retry — wait a minute and retry this id.`
        );
      } else {
        failures++;
        lines.push(`FAILED  ${id}: ${describeFailure(response.status, response.body)}`);
      }
    });

    const summary = `${message_ids.length - failures}/${message_ids.length} message(s) ${action} succeeded${failures ? `, ${failures} failed` : ""}.`;
    const result = `${summary}\n\n${lines.join("\n")}`;
    return failures === message_ids.length ? errorResult(result) : textResult(result);
  });
}
