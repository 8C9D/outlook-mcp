import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import {
  TIMEZONE,
  TZ_PREFER,
  ToolInputError,
  ToolResult,
  fetchPaged,
  runTool,
  textResult,
} from "./common.js";
import { addDays, torontoToday } from "./list-events.js";

export const listTasksSchema = {
  task_list: z
    .string()
    .optional()
    .describe(
      "Which Microsoft To Do list to read: its name (e.g. \"Tasks\", \"Groceries\") or its id. Omit for the account's default list."
    ),
  include_completed: z
    .boolean()
    .default(false)
    .describe("Also list tasks already marked complete (default false — open tasks only)."),
  due_within_days: z
    .number()
    .int()
    .min(0)
    .max(365)
    .optional()
    .describe(
      "Only list tasks due on or before this many days from today (0 = due today or earlier). Overdue tasks are always kept; tasks with no due date are dropped when this filter is set."
    ),
};

const listTasksArgs = z.object(listTasksSchema);

export const listTasksDescription =
  "List Microsoft To Do tasks, grouped as overdue / today / upcoming / no due date in America/Toronto local time. Shows each task's title, due date, reminder, and task id (use the id with manage_task). Defaults to open tasks in the default list; set include_completed to also see finished ones, or due_within_days to narrow to what is due soon.";

const TASK_CAP = 300;

export type TaskList = { id: string; displayName: string };

/**
 * Resolve a task_list input (name or id, or nothing) to one of the account's To Do
 * lists. One GET covers every case: personal accounts have few lists, and matching
 * locally means a bad name reports the available lists instead of a bare 404.
 */
export async function resolveTaskList(taskList: string | undefined): Promise<TaskList> {
  const data = await callGraphServer("/me/todo/lists?$top=100");
  const lists: any[] = data?.value ?? [];
  if (lists.length === 0) {
    throw new ToolInputError("This account has no Microsoft To Do lists.");
  }
  if (!taskList) {
    const preferred =
      lists.find((l) => l.wellknownListName === "defaultList") ?? lists[0];
    return { id: preferred.id, displayName: preferred.displayName ?? "(unnamed)" };
  }
  const match =
    lists.find((l) => l.id === taskList) ??
    lists.find(
      (l) => String(l.displayName ?? "").toLowerCase() === taskList.toLowerCase()
    );
  if (!match) {
    const names = lists.map((l) => JSON.stringify(l.displayName ?? "")).join(", ");
    throw new ToolInputError(
      `No To Do list named or identified by ${JSON.stringify(taskList)}. Available lists: ${names}.`
    );
  }
  return { id: match.id, displayName: match.displayName ?? "(unnamed)" };
}

/** The naive local wall-clock string Graph returns for a task date, given Prefer: outlook.timezone. */
function localDateTime(value: any): string | undefined {
  const dateTime = value?.dateTime as string | undefined;
  return dateTime ? dateTime.slice(0, 16) : undefined;
}

/** Render a task's due/reminder stamp: date only at midnight, date + time otherwise. */
function stamp(local: string): string {
  return local.endsWith("T00:00") ? local.slice(0, 10) : `${local.slice(0, 10)} ${local.slice(11, 16)}`;
}

export async function listTasksHandler(input: z.input<typeof listTasksArgs>): Promise<ToolResult> {
  return runTool(async () => {
    const { task_list, include_completed, due_within_days } = listTasksArgs.parse(input);
    const list = await resolveTaskList(task_list);

    const filter = include_completed ? "" : `&$filter=${encodeURIComponent("status ne 'completed'")}`;
    const tasks = await fetchPaged(
      `/me/todo/lists/${encodeURIComponent(list.id)}/tasks?$top=100${filter}`,
      TASK_CAP,
      { Prefer: TZ_PREFER }
    );

    const today = torontoToday();
    const cutoff = due_within_days === undefined ? undefined : addDays(today, due_within_days);

    const groups: Record<"overdue" | "today" | "upcoming" | "none", string[]> = {
      overdue: [],
      today: [],
      upcoming: [],
      none: [],
    };
    let shown = 0;
    for (const task of tasks) {
      const due = localDateTime(task.dueDateTime);
      const dueDate = due?.slice(0, 10);
      if (cutoff !== undefined && (!dueDate || dueDate > cutoff)) continue;
      shown++;
      const reminder = task.isReminderOn ? localDateTime(task.reminderDateTime) : undefined;
      const details = [
        due ? `due ${stamp(due)}` : undefined,
        reminder ? `reminder ${stamp(reminder)}` : undefined,
        task.status === "completed" ? "completed" : undefined,
      ].filter(Boolean);
      const line =
        `  ${task.title || "(untitled)"}${details.length ? ` — ${details.join(" · ")}` : ""}\n` +
        `    id: ${task.id}`;
      if (!dueDate) groups.none.push(line);
      else if (dueDate < today) groups.overdue.push(line);
      else if (dueDate === today) groups.today.push(line);
      else groups.upcoming.push(line);
    }

    const scope = [
      include_completed ? "open + completed" : "open",
      cutoff !== undefined ? `due on or before ${cutoff}` : undefined,
    ]
      .filter(Boolean)
      .join(", ");
    if (shown === 0) {
      return textResult(`No tasks in "${list.displayName}" (${scope}).`);
    }

    const sections: string[] = [];
    const labels: [keyof typeof groups, string][] = [
      ["overdue", "Overdue"],
      ["today", "Today"],
      ["upcoming", "Upcoming"],
      ["none", "No due date"],
    ];
    for (const [key, label] of labels) {
      if (groups[key].length === 0) continue;
      sections.push(`${label} (${groups[key].length}):\n${groups[key].join("\n")}`);
    }
    return textResult(
      `${shown} task(s) in "${list.displayName}" — ${scope} (${TIMEZONE}):\n\n${sections.join("\n\n")}`
    );
  });
}
