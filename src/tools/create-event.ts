import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { TIMEZONE, ToolInputError, ToolResult, errorResult, runTool, textResult } from "./common.js";
import { addDays } from "./list-events.js";
import { calendarBasePath, resolveCalendar } from "./list-calendars.js";

/** Graph's weekday names, indexed by JavaScript's day-of-week. */
const WEEKDAYS = [
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
] as const;

/**
 * A repeat rule in the terms a caller thinks in. Mapped onto Graph's
 * pattern/range pair by toGraphRecurrence; shared by create_event and
 * manage_event so a series is described the same way whichever tool builds it.
 */
export const recurrenceSchema = z
  .object({
    frequency: z
      .enum(["daily", "weekly", "monthly", "yearly"])
      .describe("How often the event repeats."),
    interval: z
      .number()
      .int()
      .min(1)
      .max(99)
      .default(1)
      .describe("Repeat every N days/weeks/months/years (default 1, e.g. 2 with weekly = fortnightly)."),
    weekdays: z
      .array(z.enum(WEEKDAYS))
      .min(1)
      .optional()
      .describe(
        'weekly only: which days it falls on, e.g. ["monday","thursday"]. Defaults to the start date\'s own weekday.'
      ),
    day_of_month: z
      .number()
      .int()
      .min(1)
      .max(31)
      .optional()
      .describe("monthly/yearly only: day of the month. Defaults to the start date's day."),
    month: z
      .number()
      .int()
      .min(1)
      .max(12)
      .optional()
      .describe("yearly only: month number (1-12). Defaults to the start date's month."),
    until: z
      .string()
      .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
      .optional()
      .describe("Repeat until this date inclusive (YYYY-MM-DD). Mutually exclusive with count."),
    count: z
      .number()
      .int()
      .min(1)
      .max(999)
      .optional()
      .describe("Stop after this many occurrences. Mutually exclusive with until."),
  })
  .describe(
    "Make this a repeating series. Omit for a one-off event. With neither until nor count the series never ends."
  );

export type RecurrenceInput = z.output<typeof recurrenceSchema>;

/** Map a repeat rule onto Graph's recurrence object, anchored at the start date. */
export function toGraphRecurrence(rec: RecurrenceInput, startDate: string): any {
  if (rec.until && rec.count !== undefined) {
    throw new ToolInputError("recurrence takes either until or count, not both.");
  }
  if (rec.until && rec.until < startDate) {
    throw new ToolInputError(
      `recurrence.until (${rec.until}) is before the event's start date (${startDate}).`
    );
  }
  const anchor = new Date(`${startDate}T12:00:00Z`);
  const pattern: any = { interval: rec.interval };
  switch (rec.frequency) {
    case "daily":
      pattern.type = "daily";
      break;
    case "weekly":
      pattern.type = "weekly";
      pattern.daysOfWeek = rec.weekdays ?? [WEEKDAYS[anchor.getUTCDay()]];
      break;
    case "monthly":
      pattern.type = "absoluteMonthly";
      pattern.dayOfMonth = rec.day_of_month ?? anchor.getUTCDate();
      break;
    case "yearly":
      pattern.type = "absoluteYearly";
      pattern.dayOfMonth = rec.day_of_month ?? anchor.getUTCDate();
      pattern.month = rec.month ?? anchor.getUTCMonth() + 1;
      break;
  }
  const range: any = { startDate, recurrenceTimeZone: TIMEZONE };
  if (rec.count !== undefined) {
    range.type = "numbered";
    range.numberOfOccurrences = rec.count;
  } else if (rec.until) {
    range.type = "endDate";
    range.endDate = rec.until;
  } else {
    range.type = "noEnd";
  }
  return { pattern, range };
}

/** One compact line describing a Graph recurrence, for tool output. */
export function describeRecurrence(recurrence: any): string {
  const pattern = recurrence?.pattern ?? {};
  const range = recurrence?.range ?? {};
  const every = pattern.interval > 1 ? `every ${pattern.interval} ` : "";
  const unit =
    pattern.type === "daily"
      ? "day(s)"
      : pattern.type === "weekly"
        ? `week(s) on ${(pattern.daysOfWeek ?? []).join(", ")}`
        : /Monthly$/i.test(pattern.type ?? "")
          ? `month(s) on day ${pattern.dayOfMonth}`
          : `year(s) on ${pattern.month}-${String(pattern.dayOfMonth).padStart(2, "0")}`;
  const bound =
    range.type === "numbered"
      ? `, ${range.numberOfOccurrences} occurrence(s)`
      : range.type === "endDate"
        ? `, until ${range.endDate}`
        : ", no end date";
  return `Repeats ${every}${unit}${bound}`;
}

export const createEventSchema = {
  subject: z.string().min(1).describe("Event title."),
  start: z
    .string()
    .min(1)
    .describe(
      'Start as ISO datetime, e.g. "2026-08-20T09:00" — interpreted in America/Toronto when no UTC offset is given. For all-day events a date (YYYY-MM-DD) is enough.'
    ),
  end: z
    .string()
    .optional()
    .describe("End as ISO datetime, same interpretation as start. Default: start + 60 minutes."),
  all_day: z.boolean().default(false).describe("Create an all-day event (default false)."),
  location: z.string().optional().describe("Optional location text."),
  body: z.string().optional().describe("Optional event description (plain text)."),
  attendees: z
    .array(z.string().email())
    .optional()
    .describe(
      "Optional attendee email addresses. CAUTION: when attendees are included, Outlook sends them meeting invitations as soon as the event is created."
    ),
  reminder_minutes: z
    .number()
    .int()
    .min(0)
    .max(40320)
    .optional()
    .describe(
      "Remind this many minutes before the event starts (0 = at start time, max 40320 = 4 weeks). Omit to leave the calendar's own default reminder in place."
    ),
  calendar: z
    .string()
    .optional()
    .describe(
      "Which calendar to create the event on: its name or id (see list_calendars). Omit for the account's default calendar."
    ),
  recurrence: recurrenceSchema.optional(),
};

const createEventArgs = z.object(createEventSchema);

export const createEventDescription =
  "Create an event on the user's Outlook calendar (times in America/Toronto unless an explicit UTC offset is given; default duration 60 minutes). Optionally set a reminder, target a non-default calendar by name, and make it a repeating series with recurrence. CAUTION: if attendees are provided, Outlook emails them invitations immediately when the event is created — and for a series they are invited to every occurrence. Omit attendees to create a private event with no notifications.";

const NAIVE_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/;
const OFFSET_RE = /(Z|[+-]\d{2}:?\d{2})$/;

export type GraphDateTime = { dateTime: string; timeZone: string };

/** Normalize an input datetime: naive strings stay Toronto wall-clock; offset strings become UTC. */
export function toGraphDateTime(value: string): GraphDateTime | undefined {
  if (NAIVE_RE.test(value)) {
    const dateTime = value.includes("T") ? value : `${value}T00:00:00`;
    return { dateTime: dateTime.length === 16 ? `${dateTime}:00` : dateTime, timeZone: TIMEZONE };
  }
  if (OFFSET_RE.test(value)) {
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return undefined;
    return { dateTime: d.toISOString().replace(/\.\d{3}Z$/, ""), timeZone: "UTC" };
  }
  return undefined;
}

function plusMinutes(g: GraphDateTime, minutes: number): GraphDateTime {
  const d = new Date(`${g.dateTime}Z`);
  d.setUTCMinutes(d.getUTCMinutes() + minutes);
  return { dateTime: d.toISOString().replace(/\.\d{3}Z$/, ""), timeZone: g.timeZone };
}

export async function createEventHandler(
  input: z.input<typeof createEventArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const {
      subject,
      start,
      end,
      all_day,
      location,
      body,
      attendees,
      reminder_minutes,
      calendar,
      recurrence,
    } = createEventArgs.parse(input);

    let startObj = toGraphDateTime(start);
    if (!startObj) return errorResult(`Could not parse start datetime: ${JSON.stringify(start)}.`);
    let endObj = end !== undefined ? toGraphDateTime(end) : undefined;
    if (end !== undefined && !endObj) {
      return errorResult(`Could not parse end datetime: ${JSON.stringify(end)}.`);
    }

    if (all_day) {
      // Graph requires all-day events to run midnight-to-midnight, at least one full day.
      const startDate = startObj.dateTime.slice(0, 10);
      let endDate = endObj ? endObj.dateTime.slice(0, 10) : startDate;
      if (endDate <= startDate) endDate = addDays(startDate, 1);
      startObj = { dateTime: `${startDate}T00:00:00`, timeZone: TIMEZONE };
      endObj = { dateTime: `${endDate}T00:00:00`, timeZone: TIMEZONE };
    } else if (!endObj) {
      endObj = plusMinutes(startObj, 60);
    }

    const target = await resolveCalendar(calendar);
    const graphRecurrence = recurrence
      ? toGraphRecurrence(recurrence, startObj.dateTime.slice(0, 10))
      : undefined;

    const event = await callGraphServer(`${calendarBasePath(target)}/events`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject,
        start: startObj,
        end: endObj,
        isAllDay: all_day,
        ...(location ? { location: { displayName: location } } : {}),
        ...(body ? { body: { contentType: "Text", content: body } } : {}),
        ...(attendees?.length
          ? { attendees: attendees.map((a) => ({ emailAddress: { address: a }, type: "required" })) }
          : {}),
        ...(reminder_minutes !== undefined
          ? { isReminderOn: true, reminderMinutesBeforeStart: reminder_minutes }
          : {}),
        ...(graphRecurrence ? { recurrence: graphRecurrence } : {}),
      }),
    });

    const localSpan = all_day
      ? `all day ${startObj.dateTime.slice(0, 10)}` +
        (endObj!.dateTime.slice(0, 10) !== addDays(startObj.dateTime.slice(0, 10), 1)
          ? ` to ${addDays(endObj!.dateTime.slice(0, 10), -1)}`
          : "")
      : `${startObj.dateTime.slice(0, 16).replace("T", " ")}–${endObj!.dateTime.slice(11, 16)} (${startObj.timeZone})`;
    return textResult(
      `Event created.\n` +
        `Subject: ${event.subject}\n` +
        `When: ${localSpan}\n` +
        (target ? `Calendar: ${target.name}\n` : "") +
        (graphRecurrence ? `${describeRecurrence(event.recurrence ?? graphRecurrence)}\n` : "") +
        (reminder_minutes !== undefined
          ? `Reminder: ${reminder_minutes === 0 ? "at start time" : `${reminder_minutes} min before`}\n`
          : "") +
        `Event id: ${event.id}${graphRecurrence ? " (the series master — use list_events with include_ids for a single occurrence)" : ""}\n` +
        (attendees?.length
          ? `Note: ${attendees.length} attendee(s) were included — Outlook is sending them invitations now` +
            (graphRecurrence ? ", for the whole series." : ".")
          : "No attendees — no invitations sent.")
    );
  });
}
