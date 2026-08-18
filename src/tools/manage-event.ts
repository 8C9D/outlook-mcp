import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { TIMEZONE, TZ_PREFER, ToolResult, errorResult, runTool, textResult } from "./common.js";
import { toGraphDateTime } from "./create-event.js";
import { addDays } from "./list-events.js";

export const manageEventSchema = {
  event_id: z.string().min(1).describe("The id of the event to act on."),
  action: z
    .enum(["update", "cancel", "respond"])
    .describe(
      "update: change event fields; cancel: cancel/remove the event (as organizer) or decline it (as attendee); respond: accept/decline/tentative an invitation."
    ),
  subject: z.string().optional().describe("update: new event title."),
  start: z
    .string()
    .optional()
    .describe(
      'update: new start as ISO datetime, e.g. "2026-08-20T09:00" — interpreted in America/Toronto when no UTC offset is given.'
    ),
  end: z.string().optional().describe("update: new end, same interpretation as start."),
  location: z.string().optional().describe("update: new location text."),
  body: z.string().optional().describe("update: new event description (plain text)."),
  all_day: z.boolean().optional().describe("update: make the event all-day (or not)."),
  response: z
    .enum(["accept", "decline", "tentative"])
    .optional()
    .describe('respond: your response to the invitation (required for action "respond").'),
  comment: z
    .string()
    .optional()
    .describe("respond/cancel: optional message included in the notification to the organizer/attendees."),
  send_response: z
    .boolean()
    .default(true)
    .describe("respond: whether to notify the organizer of your response (default true)."),
};

const manageEventArgs = z.object(manageEventSchema);

export const manageEventDescription =
  "Update, cancel, or respond to a calendar event. CAUTION — visible to other people: on events with attendees, updates and cancellations send notification emails to every attendee, and responses notify the organizer. Before calling, state the event's subject, date, and what will change. cancel picks the right operation automatically: organizer with attendees → cancellation notices; organizer without attendees → the event is removed (to Deleted Items); attendee → decline.";

export async function manageEventHandler(
  input: z.input<typeof manageEventArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const args = manageEventArgs.parse(input);
    const base = `/me/events/${encodeURIComponent(args.event_id)}`;

    const event = await callGraphServer(
      `${base}?$select=subject,start,end,isAllDay,isOrganizer,attendees,organizer`,
      { headers: { Prefer: TZ_PREFER } }
    );
    const attendeeCount = (event.attendees ?? []).length;
    const label = `"${event.subject || "(no subject)"}" (${String(event.start?.dateTime ?? "").slice(0, 16).replace("T", " ")})`;

    switch (args.action) {
      case "update": {
        const patch: any = {};
        if (args.subject !== undefined) patch.subject = args.subject;
        if (args.location !== undefined) patch.location = { displayName: args.location };
        if (args.body !== undefined) patch.body = { contentType: "Text", content: args.body };
        if (args.all_day !== undefined) patch.isAllDay = args.all_day;
        if (args.start !== undefined) {
          const start = toGraphDateTime(args.start);
          if (!start) return errorResult(`Could not parse start datetime: ${JSON.stringify(args.start)}.`);
          patch.start = start;
        }
        if (args.end !== undefined) {
          const end = toGraphDateTime(args.end);
          if (!end) return errorResult(`Could not parse end datetime: ${JSON.stringify(args.end)}.`);
          patch.end = end;
        }
        if (patch.isAllDay === true) {
          // Graph requires all-day events to span midnight-to-midnight full days.
          const startDate = (patch.start?.dateTime ?? event.start?.dateTime ?? "").slice(0, 10);
          let endDate = (patch.end?.dateTime ?? event.end?.dateTime ?? "").slice(0, 10);
          if (!startDate) return errorResult("Cannot determine the event's start date.");
          if (endDate <= startDate) endDate = addDays(startDate, 1);
          patch.start = { dateTime: `${startDate}T00:00:00`, timeZone: TIMEZONE };
          patch.end = { dateTime: `${endDate}T00:00:00`, timeZone: TIMEZONE };
        }
        if (Object.keys(patch).length === 0) {
          return errorResult(
            "Nothing to update — provide at least one of subject, start, end, location, body, all_day."
          );
        }
        const updated = await callGraphServer(base, {
          method: "PATCH",
          headers: { "Content-Type": "application/json", Prefer: TZ_PREFER },
          body: JSON.stringify(patch),
        });
        return textResult(
          `Event updated (${Object.keys(patch).join(", ")}).\n` +
            `Subject: ${updated.subject || "(no subject)"}\n` +
            `When: ${String(updated.start?.dateTime ?? "").slice(0, 16).replace("T", " ")}–${String(updated.end?.dateTime ?? "").slice(11, 16)} (${TIMEZONE})\n` +
            `Event id: ${updated.id}\n` +
            (attendeeCount
              ? `Note: ${attendeeCount} attendee(s) are being notified of this change.`
              : "No attendees — no notifications sent.")
        );
      }

      case "cancel": {
        if (event.isOrganizer && attendeeCount > 0) {
          await callGraphServer(`${base}/cancel`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(args.comment ? { comment: args.comment } : {}),
          });
          return textResult(
            `Event ${label} cancelled. ${attendeeCount} attendee(s) are being sent cancellation notices.`
          );
        }
        if (event.isOrganizer) {
          await callGraphServer(base, { method: "DELETE" });
          return textResult(
            `Event ${label} removed (moved to Deleted Items). No attendees — no notifications sent.`
          );
        }
        // Attendee "cancelling" = declining the invitation.
        await callGraphServer(`${base}/decline`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sendResponse: true, ...(args.comment ? { comment: args.comment } : {}) }),
        });
        return textResult(
          `You are an attendee of ${label}, not the organizer — the invitation was declined instead (organizer notified).`
        );
      }

      case "respond": {
        if (!args.response) {
          return errorResult('Action "respond" requires response (accept | decline | tentative).');
        }
        const graphAction =
          args.response === "accept"
            ? "accept"
            : args.response === "decline"
              ? "decline"
              : "tentativelyAccept";
        await callGraphServer(`${base}/${graphAction}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            sendResponse: args.send_response,
            ...(args.comment ? { comment: args.comment } : {}),
          }),
        });
        return textResult(
          `Responded "${args.response}" to ${label}.` +
            (args.send_response
              ? " The organizer is being notified."
              : " No response sent to the organizer.")
        );
      }
    }
  });
}
