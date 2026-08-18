import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
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
import { recurrenceSchema, toGraphRecurrence, describeRecurrence } from "./create-event.js";

export const manageTaskSchema = {
  action: z
    .enum([
      "create",
      "complete",
      "reopen",
      "update",
      "delete",
      "add_subtask",
      "complete_subtask",
      "remove_subtask",
      "create_list",
      "rename_list",
    ])
    .describe(
      "create: a new task (title required); complete: mark task_id done; reopen: mark task_id not started again; update: change fields on task_id; delete: PERMANENTLY remove task_id; add_subtask/complete_subtask/remove_subtask: work on task_id's checklist (subtasks); create_list: a new To Do list named list_name; rename_list: rename task_list to list_name."
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
  recurrence: recurrenceSchema
    .optional()
    .describe(
      'For action "create" only: make this a repeating task. due_date is required with it (Microsoft To Do anchors the repeat on the due date). Microsoft To Do refuses recurrence changes on an existing task, so to change a repeat, delete the task and create it again; "update" with clear_recurrence stops the repeat.'
    ),
  clear_recurrence: z
    .boolean()
    .optional()
    .describe(
      'For action "update": stop a repeating task from repeating (the task itself stays). This is the only recurrence change Microsoft To Do accepts after creation.'
    ),
  subtask: z
    .string()
    .min(1)
    .optional()
    .describe(
      "The checklist item (subtask) to act on: its text for add_subtask, or the exact text of an existing item for complete_subtask/remove_subtask when subtask_id is not known."
    ),
  subtask_id: z
    .string()
    .min(1)
    .optional()
    .describe(
      "The id of the checklist item for complete_subtask/remove_subtask (shown by add_subtask and by every subtask listing). Takes precedence over subtask."
    ),
  list_name: z
    .string()
    .min(1)
    .optional()
    .describe(
      "For create_list: the name of the new To Do list. For rename_list: the new name for the list named by task_list."
    ),
};

const manageTaskArgs = z.object(manageTaskSchema);

export const manageTaskDescription =
  "Create, complete, reopen, update, or delete a Microsoft To Do task; manage its subtasks (checklist items); create or rename a To Do list. Dates are America/Toronto. Pass linked_message_id on create to turn an email into a task — the mail's subject, sender, and an Outlook link go into the task notes. Pass recurrence on create for a repeating task (due_date required). WARNING: delete here is PERMANENT — unlike mail, a deleted To Do task does not go to a recoverable folder and cannot be restored, so state the task's title to the user and get agreement before deleting. To finish a task while keeping it, use complete instead of delete. This tool deliberately CANNOT delete a To Do list: deleting a list destroys every task in it with no recoverable copy, which no soft-delete convention can undo — delete a list in the Microsoft To Do app if that is really what the user wants.";

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

/** The checklist of a task, newest state, rendered compactly with ids. */
async function checklistOf(tasksPath: string, taskId: string): Promise<any[]> {
  const data = await callGraphServer(
    `${tasksPath}/${encodeURIComponent(taskId)}/checklistItems?$top=100`
  );
  return data?.value ?? [];
}

function describeChecklist(items: any[]): string {
  if (items.length === 0) return "Subtasks: none";
  const done = items.filter((i) => i.isChecked).length;
  const lines = items.map(
    (i) => `  [${i.isChecked ? "x" : " "}] ${i.displayName || "(untitled)"}\n    subtask id: ${i.id}`
  );
  return `Subtasks (${done}/${items.length} done):\n${lines.join("\n")}`;
}

/** Find one checklist item by id or by exact text, or say why it could not be found. */
function resolveChecklistItem(
  items: any[],
  subtaskId: string | undefined,
  subtaskText: string | undefined
): any {
  if (subtaskId) {
    const byId = items.find((i) => i.id === subtaskId);
    if (!byId) {
      throw new ToolInputError(
        `No subtask ${subtaskId} on this task. Current subtasks:\n${describeChecklist(items)}`
      );
    }
    return byId;
  }
  const matches = items.filter(
    (i) => String(i.displayName ?? "").toLowerCase() === subtaskText!.toLowerCase()
  );
  if (matches.length === 0) {
    throw new ToolInputError(
      `No subtask reading ${JSON.stringify(subtaskText)} on this task. Current subtasks:\n${describeChecklist(items)}`
    );
  }
  if (matches.length > 1) {
    throw new ToolInputError(
      `${matches.length} subtasks read ${JSON.stringify(subtaskText)} — pass subtask_id instead:\n${describeChecklist(items)}`
    );
  }
  return matches[0];
}

export async function manageTaskHandler(
  input: z.input<typeof manageTaskArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const {
      action,
      task_id,
      task_list,
      title,
      due_date,
      body,
      reminder,
      linked_message_id,
      recurrence,
      clear_recurrence,
      subtask,
      subtask_id,
      list_name,
    } = manageTaskArgs.parse(input);

    const subtaskActions = ["add_subtask", "complete_subtask", "remove_subtask"];
    const listActions = ["create_list", "rename_list"];

    if (!listActions.includes(action) && action !== "create" && !task_id) {
      return errorResult(`Action "${action}" requires task_id (from list_tasks).`);
    }
    if (action !== "create" && linked_message_id) {
      return errorResult('linked_message_id is only supported on action "create".');
    }
    if (action !== "create" && recurrence) {
      return errorResult(
        'recurrence can only be set when the task is created ("create"). Microsoft To Do rejects ' +
          "recurrence changes on an existing task, so to change how a task repeats, delete it and " +
          'create it again. To stop it repeating, use action "update" with clear_recurrence.'
      );
    }
    if (clear_recurrence && action !== "update") {
      return errorResult('clear_recurrence is only supported on action "update".');
    }

    // ---- list management (no task involved) -------------------------------
    if (action === "create_list") {
      if (!list_name) return errorResult('Action "create_list" requires list_name.');
      const existing = await callGraphServer("/me/todo/lists?$top=100");
      const clash = (existing?.value ?? []).find(
        (l: any) => String(l.displayName ?? "").toLowerCase() === list_name.toLowerCase()
      );
      if (clash) {
        return errorResult(
          `A To Do list named "${clash.displayName}" already exists (id ${clash.id}) — use it, or pick another name.`
        );
      }
      const created = await callGraphServer("/me/todo/lists", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName: list_name }),
      });
      return textResult(
        `To Do list "${created.displayName}" created.\nList id: ${created.id}\n` +
          "Pass its name as task_list to put tasks in it."
      );
    }

    if (action === "rename_list") {
      if (!list_name) return errorResult('Action "rename_list" requires list_name (the new name).');
      const target = await resolveTaskList(task_list);
      const renamed = await callGraphServer(`/me/todo/lists/${encodeURIComponent(target.id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ displayName: list_name }),
      });
      return textResult(
        `To Do list renamed from "${target.displayName}" to "${renamed.displayName}". ` +
          "Its tasks are untouched.\n" +
          `List id: ${renamed.id}`
      );
    }

    const list = await resolveTaskList(task_list);
    const tasksPath = `/me/todo/lists/${encodeURIComponent(list.id)}/tasks`;
    const taskPath = task_id ? `${tasksPath}/${encodeURIComponent(task_id)}` : "";

    if (action === "create") {
      if (!title) return errorResult('Action "create" requires title.');
      if (recurrence && !due_date) {
        return errorResult(
          "A repeating task needs due_date: Microsoft To Do anchors the repeat on the due date and " +
            "rejects recurrence without one."
        );
      }
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
          ...(recurrence ? { recurrence: toGraphRecurrence(recurrence, due_date!) } : {}),
        }),
      });
      return textResult(
        `Task created in "${list.displayName}".\n` +
          `Title: ${created.title}\n` +
          (due_date ? `Due: ${due_date} (${TIMEZONE})\n` : "") +
          (reminder ? `Reminder: ${reminder.replace("T", " ")} (${TIMEZONE})\n` : "") +
          (created.recurrence ? `${describeRecurrence(created.recurrence)}\n` : "") +
          (linked_message_id ? "Linked email details added to the task notes.\n" : "") +
          `Task id: ${created.id}`
      );
    }

    // ---- subtasks (checklist items) ---------------------------------------
    if (subtaskActions.includes(action)) {
      if (action === "add_subtask") {
        if (!subtask) return errorResult('Action "add_subtask" requires subtask (its text).');
        let added: any;
        try {
          added = await callGraphServer(`${taskPath}/checklistItems`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ displayName: subtask }),
          });
        } catch (err) {
          if (isNotFound(err)) {
            return errorResult(
              `No task ${task_id} in "${list.displayName}" — check the id with list_tasks, and pass task_list if it lives in another list.`
            );
          }
          throw err;
        }
        const items = await checklistOf(tasksPath, task_id!);
        return textResult(
          `Subtask added to the task in "${list.displayName}".\n` +
            `Added: ${added.displayName}\nSubtask id: ${added.id}\n\n${describeChecklist(items)}`
        );
      }

      let items: any[];
      try {
        items = await checklistOf(tasksPath, task_id!);
      } catch (err) {
        if (isNotFound(err)) {
          return errorResult(
            `No task ${task_id} in "${list.displayName}" — check the id with list_tasks, and pass task_list if it lives in another list.`
          );
        }
        throw err;
      }
      if (!subtask_id && !subtask) {
        return errorResult(
          `Action "${action}" requires subtask_id or subtask (the item's exact text). Current subtasks:\n${describeChecklist(items)}`
        );
      }
      const item = resolveChecklistItem(items, subtask_id, subtask);
      const itemPath = `${taskPath}/checklistItems/${encodeURIComponent(item.id)}`;

      if (action === "complete_subtask") {
        await callGraphServer(itemPath, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ isChecked: true }),
        });
        return textResult(
          `Subtask "${item.displayName}" marked done.\n\n${describeChecklist(await checklistOf(tasksPath, task_id!))}`
        );
      }

      await callGraphServer(itemPath, { method: "DELETE" });
      return textResult(
        `Subtask "${item.displayName}" removed from the task. Removing a subtask is permanent — ` +
          "To Do keeps no recoverable copy of it.\n\n" +
          describeChecklist(await checklistOf(tasksPath, task_id!))
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
      // Graph accepts recurrence: null on an existing task (it refuses every
      // other recurrence change), which is what makes "stop repeating" possible.
      if (clear_recurrence) patch.recurrence = null;
      if (Object.keys(patch).length === 0) {
        return errorResult(
          'Action "update" needs at least one of title, due_date, body, reminder, or clear_recurrence.'
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
        (updated.recurrence
          ? `${describeRecurrence(updated.recurrence)}\n`
          : clear_recurrence
            ? "Repeat: off (this task no longer repeats)\n"
            : "") +
        `Task id: ${updated.id}`
    );
  });
}
