import { z } from "zod";
import { callGraphServer } from "../graph.js";
import {
  TIMEZONE,
  TZ_PREFER,
  ToolInputError,
  ToolResult,
  errorResult,
  formatLocal,
  formatSender,
  isNotFound,
  runTool,
  textResult,
} from "./common.js";
import { resolveTaskList } from "./list-tasks.js";

export const manageTaskSchema = {
  action: z
    .enum(["create", "complete", "reopen", "update", "delete"])
    .describe(
      "create: a new task (title required); complete: mark task_id done; reopen: mark task_id not started again; update: change fields on task_id; delete: PERMANENTLY remove task_id."
    ),
  task_id: z
    .string()
    .min(1)
    .optional()
    .describe(
      "The task to act on (from list_tasks). Required for complete, reopen, update, and delete."
    ),
  task_list: z
    .string()
    .optional()
    .describe(
      "Microsoft To Do list to work in: its name or id. Omit for the default list. For complete/reopen/update/delete this must be the list the task actually lives in."
    ),
  title: z
    .string()
    .min(1)
    .optional()
    .describe('Task title. Required for action "create"; optional for update.'),
  due_date: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
    .optional()
    .describe("Due date as an ISO date (YYYY-MM-DD), interpreted in America/Toronto."),
  body: z
    .string()
    .optional()
    .describe("Plain-text notes for the task. On update this replaces the existing notes."),
  reminder: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/, "ISO local datetime (YYYY-MM-DDTHH:MM)")
    .optional()
    .describe(
      "Reminder time as YYYY-MM-DDTHH:MM in America/Toronto (no timezone suffix). Setting it turns the task's reminder on."
    ),
  linked_message_id: z
    .string()
    .min(1)
    .optional()
    .describe(
      'For action "create": turn an email into a task. The message\'s subject, sender, date, and a link that opens it in Outlook are appended to the task notes.'
    ),
};

const manageTaskArgs = z.object(manageTaskSchema);

export const manageTaskDescription =
  "Create, complete, reopen, update, or delete a Microsoft To Do task. Dates are America/Toronto. Pass linked_message_id on create to turn an email into a task — the mail's subject, sender, and an Outlook link go into the task notes. WARNING: delete here is PERMANENT — unlike mail, a deleted To Do task does not go to a recoverable folder and cannot be restored, so state the task's title to the user and get agreement before deleting. To finish a task while keeping it, use complete instead of delete.";

/** A Graph dateTimeTimeZone for a naive local wall-clock value. */
function localStamp(dateTime: string): { dateTime: string; timeZone: string } {
  return { dateTime, timeZone: TIMEZONE };
}

/** The "from this email" block appended to a task's notes. */
async function linkedMessageNote(messageId: string): Promise<string> {
  let message: any;
  try {
    message = await callGraphServer(
      `/me/messages/${encodeURIComponent(messageId)}?$select=subject,from,receivedDateTime,webLink`,
      { headers: { Prefer: TZ_PREFER } }
    );
  } catch (err) {
    if (isNotFound(err)) {
      throw new ToolInputError(
        `linked_message_id ${messageId} does not match any message — use a message id from search_mail.`
      );
    }
    throw err;
  }
  return [
    "From email:",
    `Subject: ${message.subject || "(no subject)"}`,
    `From: ${formatSender(message.from)}`,
    `Received: ${formatLocal(message.receivedDateTime)}`,
    ...(message.webLink ? [`Open in Outlook: ${message.webLink}`] : []),
  ].join("\n");
}

export async function manageTaskHandler(
  input: z.input<typeof manageTaskArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, task_id, task_list, title, due_date, body, reminder, linked_message_id } =
      manageTaskArgs.parse(input);

    if (action !== "create" && !task_id) {
      return errorResult(`Action "${action}" requires task_id (from list_tasks).`);
    }
    if (action !== "create" && linked_message_id) {
      return errorResult('linked_message_id is only supported on action "create".');
    }

    const list = await resolveTaskList(task_list);
    const tasksPath = `/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;
    const taskPath = task_id ? `${tasksPath}/${encodeURIComponent(task_id)}` : "";

    if (action === "create") {
      if (!title) return errorResult('Action "create" requires title.');
      const notes = [
        ...(body ? [body] : []),
        ...(linked_message_id ? [await linkedMessageNote(linked_message_id)] : []),
      ].join("\n\n");
      const created = await callGraphServer(tasksPath, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          title,
          ...(notes ? { body: { content: notes, contentType: "text" } } : {}),
          ...(due_date ? { dueDateTime: localStamp(`${due_date}T00:00:00`) } : {}),
          ...(reminder
            ? { reminderDateTime: localStamp(`${reminder}:00`), isReminderOn: true }
            : {}),
        }),
      });
      return textResult(
        `Task created in "${list.displayName}".\n` +
          `Title: ${created.title}\n` +
          (due_date ? `Due: ${due_date} (${TIMEZONE})\n` : "") +
          (reminder ? `Reminder: ${reminder.replace("T", " ")} (${TIMEZONE})\n` : "") +
          (linked_message_id ? "Linked email details added to the task notes.\n" : "") +
          `Task id: ${created.id}`
      );
    }

    if (action === "delete") {
      // Read the title first so the confirmation names what was destroyed; To Do
      // deletes are permanent, so there is no recovering it afterwards.
      let doomedTitle = task_id;
      try {
        const existing = await callGraphServer(taskPath);
        if (existing?.title) doomedTitle = existing.title;
      } catch (err) {
        if (isNotFound(err)) {
          return errorResult(
            `No task ${task_id} in "${list.displayName}" — check the id with list_tasks, and pass task_list if it lives in another list.`
          );
        }
        throw err;
      }
      await callGraphServer(taskPath, { method: "DELETE" });
      return textResult(
        `Task "${doomedTitle}" deleted permanently from "${list.displayName}". ` +
          "To Do has no recoverable deleted-items folder, so this cannot be undone."
      );
    }

    // complete / reopen / update all PATCH the task.
    const patch: any = {};
    let summary: string;
    if (action === "complete") {
      patch.status = "completed";
      summary = "marked complete";
    } else if (action === "reopen") {
      patch.status = "notStarted";
      summary = "reopened (status: not started)";
    } else {
      if (title !== undefined) patch.title = title;
      if (body !== undefined) patch.body = { content: body, contentType: "text" };
      if (due_date !== undefined) patch.dueDateTime = localStamp(`${due_date}T00:00:00`);
      if (reminder !== undefined) {
        patch.reminderDateTime = localStamp(`${reminder}:00`);
        patch.isReminderOn = true;
      }
      if (Object.keys(patch).length === 0) {
        return errorResult(
          'Action "update" needs at least one of title, due_date, body, or reminder.'
        );
      }
      summary = `updated (${Object.keys(patch).join(", ")})`;
    }

    let updated: any;
    try {
      updated = await callGraphServer(taskPath, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(patch),
      });
    } catch (err) {
      if (isNotFound(err)) {
        return errorResult(
          `No task ${task_id} in "${list.displayName}" — check the id with list_tasks, and pass task_list if it lives in another list.`
        );
      }
      throw err;
    }

    return textResult(
      `Task ${summary}.\n` +
        `Title: ${updated.title}\n` +
        `List: "${list.displayName}"\n` +
        `Status: ${updated.status}\n` +
        `Task id: ${updated.id}`
    );
  });
}
