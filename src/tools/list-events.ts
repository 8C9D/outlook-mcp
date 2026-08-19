import { z } from "zod";
import { TIMEZONE, TZ_PREFER, ToolResult, fetchPaged, runTool, structuredResult } from "./common.js";
import { calendarBasePath, resolveCalendar } from "./list-calendars.js";

export const listEventsSchema = {
  start_date: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
    .optional()
    .describe("First day of the window as an ISO date (YYYY-MM-DD) in America/Toronto. Default: today."),
  days: z
    .number()
    .int()
    .min(1)
    .max(31)
    .default(7)
    .describe("Number of days to list, starting at start_date (default 7, max 31)."),
  calendar: z
    .string()
    .optional()
    .describe(
      "Which calendar to read: its name or id (see list_calendars). Omit for the account's default calendar."
    ),
  include_ids: z
    .boolean()
    .default(false)
    .describe(
      "Also print each event's id and, for repeating events, whether the line is one occurrence of a series (default false). Needed before manage_event can act on a specific event or on a single occurrence."
    ),
};

const listEventsArgs = z.object(listEventsSchema);

/** Permissive machine-readable day grouping; every field optional. */
export const listEventsOutputSchema = {
  startDate: z.string().optional(),
  endDate: z.string().optional().describe("Last day of the window, inclusive."),
  timezone: z.string().optional(),
  calendar: z.string().optional(),
  count: z.number().optional(),
  days: z
    .array(
      z.looseObject({
        date: z.string().optional(),
        events: z
          .array(
            z.looseObject({
              subject: z.string().optional(),
              start: z.string().optional().describe("Local wall-clock HH:MM; absent when all-day."),
              end: z.string().optional(),
              allDay: z.boolean().optional(),
              location: z.string().optional(),
              id: z.string().optional().describe("Present when include_ids is set."),
              occurrence: z.boolean().optional().describe("One date of a repeating series."),
            })
          )
          .optional(),
      })
    )
    .optional(),
};

export const listEventsDescription =
  "List the user's Outlook calendar events for a date window, grouped by day in America/Toronto local time. Shows each event as start–end time, subject, and location; all-day events are listed first within each day, and repeating events appear once per occurrence. Defaults to the next 7 days of the default calendar starting today; set calendar to read another one and include_ids to get the ids manage_event needs.";

/** Today's date in America/Toronto as YYYY-MM-DD. */
export function torontoToday(): string {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

export function addDays(isoDate: string, days: number): string {
  const d = new Date(`${isoDate}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

const EVENT_CAP = 500;

export async function listEventsHandler(
  input: z.input<typeof listEventsArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const parsed = listEventsArgs.parse(input);
    const startDate = parsed.start_date ?? torontoToday();
    const endDate = addDays(startDate, parsed.days);
    const target = await resolveCalendar(parsed.calendar);
    const where = target ? ` in ${target.name}` : "";
    // Naive datetime bounds are interpreted in the Prefer: outlook.timezone zone.
    const path =
      `${calendarBasePath(target)}/calendarView?startDateTime=${encodeURIComponent(`${startDate}T00:00:00`)}` +
      `&endDateTime=${encodeURIComponent(`${endDate}T00:00:00`)}` +
      `&$select=id,subject,start,end,location,isAllDay,organizer,type&$orderby=start/dateTime&$top=50`;
    const events = await fetchPaged(path, EVENT_CAP, { Prefer: TZ_PREFER });

    const lastDay = addDays(startDate, parsed.days - 1);
    const structuredBase = {
      startDate,
      endDate: lastDay,
      timezone: TIMEZONE,
      ...(target ? { calendar: target.name } : {}),
    };
    if (events.length === 0) {
      return structuredResult(`No events in this window (${startDate} to ${lastDay})${where}.`, {
        ...structuredBase,
        count: 0,
        days: [],
      });
    }

    const byDay = new Map<string, { allDay: any[]; timed: any[] }>();
    for (const ev of events) {
      const day = String(ev.start?.dateTime ?? "").slice(0, 10);
      if (!byDay.has(day)) byDay.set(day, { allDay: [], timed: [] });
      (ev.isAllDay ? byDay.get(day)!.allDay : byDay.get(day)!.timed).push(ev);
    }

    const sections: string[] = [];
    const structuredDays: Record<string, unknown>[] = [];
    for (const day of [...byDay.keys()].sort()) {
      const { allDay, timed } = byDay.get(day)!;
      const lines = [
        ...allDay.map((ev) => `  all day     ${describe(ev)}${identify(ev, parsed.include_ids)}`),
        ...timed.map(
          (ev) => `  ${hhmm(ev.start)}–${hhmm(ev.end)} ${describe(ev)}${identify(ev, parsed.include_ids)}`
        ),
      ];
      sections.push(`${day}\n${lines.join("\n")}`);
      structuredDays.push({
        date: day,
        events: [...allDay, ...timed].map((ev) => describeStructured(ev, parsed.include_ids)),
      });
    }
    return structuredResult(
      `Events ${startDate} to ${lastDay}${where} (${TIMEZONE}):\n\n${sections.join("\n\n")}`,
      { ...structuredBase, count: events.length, days: structuredDays }
    );
  });
}

function describeStructured(ev: any, includeIds: boolean): Record<string, unknown> {
  const repeating = ev.type === "occurrence" || ev.type === "exception";
  return {
    subject: ev.subject || "(no subject)",
    allDay: Boolean(ev.isAllDay),
    ...(ev.isAllDay ? {} : { start: hhmm(ev.start), end: hhmm(ev.end) }),
    ...(ev.location?.displayName ? { location: ev.location.displayName } : {}),
    ...(includeIds ? { id: String(ev.id ?? ""), occurrence: repeating } : {}),
  };
}

function hhmm(when: any): string {
  return String(when?.dateTime ?? "").slice(11, 16) || "??:??";
}

function describe(ev: any): string {
  const location = ev.location?.displayName;
  return `${ev.subject || "(no subject)"}${location ? ` (${location})` : ""}`;
}

/** The id line, when asked for. Occurrences say so, since editing one differs from editing the series. */
function identify(ev: any, includeIds: boolean): string {
  if (!includeIds) return "";
  const repeating = ev.type === "occurrence" || ev.type === "exception";
  return `\n      id: ${ev.id}${repeating ? " (one occurrence of a repeating event)" : ""}`;
}
