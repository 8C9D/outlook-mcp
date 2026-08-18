import { z } from "zod";
import { callGraphServer } from "../graph.js";
import { TIMEZONE, ToolResult, errorResult, runTool, textResult } from "./common.js";
import { addDays } from "./list-events.js";

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
};

const createEventArgs = z.object(createEventSchema);

export const createEventDescription =
  "Create an event on the user's Outlook calendar (times in America/Toronto unless an explicit UTC offset is given; default duration 60 minutes). CAUTION: if attendees are provided, Outlook emails them invitations immediately when the event is created — omit attendees to create a private event with no notifications.";

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
    const { subject, start, end, all_day, location, body, attendees } =
      createEventArgs.parse(input);

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

    const event = await callGraphServer("/me/events", {
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
        `Event id: ${event.id}\n` +
        (attendees?.length
          ? `Note: ${attendees.length} attendee(s) were included — Outlook is sending them invitations now.`
          : "No attendees — no invitations sent.")
    );
  });
}
