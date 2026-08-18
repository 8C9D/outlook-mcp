import { z } from "zod";
import { GraphError, callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";

export const manageMessageSchema = {
  message_ids: z
    .array(z.string().min(1))
    .min(1)
    .max(20)
    .describe("Ids of the messages to act on (1–20)."),
  action: z
    .enum(["move", "archive", "delete", "mark_read", "mark_unread", "flag", "unflag"])
    .describe(
      "move: to destination_folder; archive: to the Archive folder; delete: to Deleted Items (soft delete, recoverable); mark_read/mark_unread; flag/unflag: follow-up flag."
    ),
  destination_folder: z
    .string()
    .optional()
    .describe(
      'Target folder for action "move": a well-known name ("inbox", "archive", "deleteditems", "junkemail", "drafts") or a folder id from list_folders.'
    ),
};

const manageMessageArgs = z.object(manageMessageSchema);

export const manageMessageDescription =
  "Act on up to 20 email messages: move, archive, delete (soft — moves to Deleted Items), mark read/unread, or flag/unflag. Returns a per-message success/failure list; moved messages get a NEW message id (reported in the output). Before calling with delete or move, state to the user exactly which messages (subjects/senders) will be affected. Deletion is never permanent from this tool — messages remain recoverable in Deleted Items.";

async function applyAction(
  id: string,
  action: string,
  destinationFolder: string | undefined
): Promise<string> {
  const base = `/me/messages/${encodeURIComponent(id)}`;
  const patch = (body: object) =>
    callGraphServer(base, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  const move = async (destinationId: string) => {
    const moved = await callGraphServer(`${base}/move`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId }),
    });
    return moved?.id as string | undefined;
  };

  switch (action) {
    case "move": {
      const newId = await move(destinationFolder!);
      return `moved to ${destinationFolder}${newId ? ` (new id: ${newId})` : ""}`;
    }
    case "archive": {
      const newId = await move("archive");
      return `archived${newId ? ` (new id: ${newId})` : ""}`;
    }
    case "delete":
      await callGraphServer(base, { method: "DELETE" });
      return "deleted (moved to Deleted Items)";
    case "mark_read":
      await patch({ isRead: true });
      return "marked read";
    case "mark_unread":
      await patch({ isRead: false });
      return "marked unread";
    case "flag":
      await patch({ flag: { flagStatus: "flagged" } });
      return "flagged";
    case "unflag":
      await patch({ flag: { flagStatus: "notFlagged" } });
      return "unflagged";
    default:
      throw new Error(`Unknown action ${action}`);
  }
}

export async function manageMessageHandler(
  input: z.input<typeof manageMessageArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { message_ids, action, destination_folder } = manageMessageArgs.parse(input);
    if (action === "move" && !destination_folder) {
      return errorResult('Action "move" requires destination_folder.');
    }

    const lines: string[] = [];
    let failures = 0;
    for (const id of message_ids) {
      try {
        const outcome = await applyAction(id, action, destination_folder);
        lines.push(`OK      ${id}: ${outcome}`);
      } catch (err) {
        failures++;
        let reason = err instanceof Error ? err.message : String(err);
        if (err instanceof GraphError) {
          try {
            const parsed = JSON.parse(err.body);
            reason = `HTTP ${err.status} ${parsed?.error?.code ?? ""}: ${parsed?.error?.message ?? err.statusText}`;
          } catch {
            reason = `HTTP ${err.status} ${err.statusText}`;
          }
        }
        lines.push(`FAILED  ${id}: ${reason}`);
      }
    }

    const summary = `${message_ids.length - failures}/${message_ids.length} message(s) ${action} succeeded${failures ? `, ${failures} failed` : ""}.`;
    const result = `${summary}\n\n${lines.join("\n")}`;
    return failures === message_ids.length ? errorResult(result) : textResult(result);
  });
}
