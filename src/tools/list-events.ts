import { z } from "zod";
import { TIMEZONE, TZ_PREFER, ToolResult, fetchPaged, runTool, textResult } from "./common.js";

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
};

const listEventsArgs = z.object(listEventsSchema);

export const listEventsDescription =
  "List the user's Outlook calendar events for a date window, grouped by day in America/Toronto local time. Shows each event as start–end time, subject, and location; all-day events are listed first within each day. Defaults to the next 7 days starting today.";

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
    // Naive datetime bounds are interpreted in the Prefer: outlook.timezone zone.
    const path =
      `/me/calendarView?startDateTime=${encodeURIComponent(`${startDate}T00:00:00`)}` +
      `&endDateTime=${encodeURIComponent(`${endDate}T00:00:00`)}` +
      `&$select=subject,start,end,location,isAllDay,organizer&$orderby=start/dateTime&$top=50`;
    const events = await fetchPaged(path, EVENT_CAP, { Prefer: TZ_PREFER });

    const lastDay = addDays(startDate, parsed.days - 1);
    if (events.length === 0) {
      return textResult(`No events in this window (${startDate} to ${lastDay}).`);
    }

    const byDay = new Map<string, { allDay: any[]; timed: any[] }>();
    for (const ev of events) {
      const day = String(ev.start?.dateTime ?? "").slice(0, 10);
      if (!byDay.has(day)) byDay.set(day, { allDay: [], timed: [] });
      (ev.isAllDay ? byDay.get(day)!.allDay : byDay.get(day)!.timed).push(ev);
    }

    const sections: string[] = [];
    for (const day of [...byDay.keys()].sort()) {
      const { allDay, timed } = byDay.get(day)!;
      const lines = [
        ...allDay.map((ev) => `  all day     ${describe(ev)}`),
        ...timed.map((ev) => `  ${hhmm(ev.start)}–${hhmm(ev.end)} ${describe(ev)}`),
      ];
      sections.push(`${day}\n${lines.join("\n")}`);
    }
    return textResult(
      `Events ${startDate} to ${lastDay} (${TIMEZONE}):\n\n${sections.join("\n\n")}`
    );
  });
}

function hhmm(when: any): string {
  return String(when?.dateTime ?? "").slice(11, 16) || "??:??";
}

function describe(ev: any): string {
  const location = ev.location?.displayName;
  return `${ev.subject || "(no subject)"}${location ? ` (${location})` : ""}`;
}
