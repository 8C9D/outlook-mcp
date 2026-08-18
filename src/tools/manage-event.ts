import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { TIMEZONE, TZ_PREFER, ToolResult, errorResult, runTool, textResult } from "./common.js";
import { describeRecurrence, recurrenceSchema, toGraphDateTime, toGraphRecurrence } from "./create-event.js";
import { addDays } from "./list-events.js";

export const manageEventSchema = {
  event_id: z.string().min(1).describe("The id of the event to act on."),
  action: z
    .enum(["update", "cancel", "respond"])
    .describe(
      "update: change event fields; cancel: cancel/remove the event (as organizer) or decline it (as attendee); respond: accept/decline/tentative an invitation."
    ),
  scope: z
    .enum(["this_event_only", "entire_series"])
    .optional()
    .describe(
      "Repeating events only: whether update/cancel applies to the one occurrence named by event_id or to the whole series. Defaults to whichever event_id names — an occurrence id changes just that occurrence, a series id changes the series. Pass an occurrence id (list_events with include_ids) to use this_event_only."
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
  reminder_minutes: z
    .number()
    .int()
    .min(-1)
    .max(40320)
    .optional()
    .describe(
      "update: remind this many minutes before the start (0 = at start time, max 40320 = 4 weeks). Use -1 to turn the reminder off."
    ),
  recurrence: recurrenceSchema
    .optional()
    .describe(
      "update: replace the repeat rule, or turn a one-off event into a series. Applies to the whole series, so it cannot be combined with scope this_event_only."
    ),
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
  "Update, cancel, or respond to a calendar event, including one occurrence of a repeating event or a whole series. CAUTION — visible to other people: on events with attendees, updates and cancellations send notification emails to every attendee, and responses notify the organizer. Editing or cancelling an entire series notifies every attendee about all of its occurrences (and a changed repeat rule re-issues the whole series); editing one occurrence notifies them about that date only. Before calling, state the event's subject, date, whether you are touching one occurrence or the series, and what will change. cancel picks the right operation automatically: organizer with attendees → cancellation notices; organizer without attendees → the event is removed (to Deleted Items); attendee → decline.";

const EVENT_SELECT =
  "subject,start,end,isAllDay,isOrganizer,attendees,organizer,type,seriesMasterId,recurrence,isReminderOn,reminderMinutesBeforeStart";

export async function manageEventHandler(
  input: z.input<typeof manageEventArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const args = manageEventArgs.parse(input);

    let event = await callGraphServer(
      `/me/events/${encodeURIComponent(args.event_id)}?$select=${EVENT_SELECT}`,
      { headers: { Prefer: TZ_PREFER } }
    );

    // Which event the action actually lands on. Graph models a repeating event
    // as a seriesMaster plus per-date occurrences, each with its own id: acting
    // on the occurrence id creates (or removes) an exception, acting on the
    // master changes every date.
    const isOccurrence = event.type === "occurrence" || event.type === "exception";
    const isSeriesMaster = event.type === "seriesMaster";
    let targetId: string = args.event_id;
    let seriesWide = isSeriesMaster;

    if (isOccurrence && args.scope === "entire_series") {
      if (!event.seriesMasterId) {
        return errorResult(
          "This occurrence does not report a series master, so the series cannot be edited from it."
        );
      }
      targetId = event.seriesMasterId;
      seriesWide = true;
      event = await callGraphServer(
        `/me/events/${encodeURIComponent(targetId)}?$select=${EVENT_SELECT}`,
        { headers: { Prefer: TZ_PREFER } }
      );
    } else if (isSeriesMaster && args.scope === "this_event_only") {
      return errorResult(
        "That id is the whole repeating series, not one occurrence. Call list_events with " +
          "include_ids for the window you mean and use the id of the occurrence you want to change."
      );
    }

    const base = `/me/events/${encodeURIComponent(targetId)}`;
    const attendeeCount = (event.attendees ?? []).length;
    const what = seriesWide
      ? "the entire series"
      : isOccurrence
        ? "this occurrence only"
        : "this event";
    const label = `"${event.subject || "(no subject)"}" (${String(event.start?.dateTime ?? "").slice(0, 16).replace("T", " ")})`;

    switch (args.action) {
      case "update": {
        const patch: any = {};
        if (args.subject !== undefined) patch.subject = args.subject;
        if (args.location !== undefined) patch.location = { displayName: args.location };
        if (args.body !== undefined) patch.body = { contentType: "Text", content: args.body };
        if (args.all_day !== undefined) patch.isAllDay = args.all_day;
        if (args.reminder_minutes !== undefined) {
          if (args.reminder_minutes < 0) {
            patch.isReminderOn = false;
          } else {
            patch.isReminderOn = true;
            patch.reminderMinutesBeforeStart = args.reminder_minutes;
          }
        }
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
        if (args.recurrence) {
          if (!seriesWide && isOccurrence) {
            return errorResult(
              "A repeat rule belongs to the whole series, not one date. Repeat the call with " +
                "scope entire_series (or the series id) to change how often it recurs."
            );
          }
          const anchor = (patch.start?.dateTime ?? event.start?.dateTime ?? "").slice(0, 10);
          if (!anchor) return errorResult("Cannot determine the event's start date.");
          patch.recurrence = toGraphRecurrence(args.recurrence, anchor);
        }
        if (Object.keys(patch).length === 0) {
          return errorResult(
            "Nothing to update — provide at least one of subject, start, end, location, body, " +
              "all_day, reminder_minutes, recurrence."
          );
        }
        const updated = await callGraphServer(base, {
          method: "PATCH",
          headers: { "Content-Type": "application/json", Prefer: TZ_PREFER },
          body: JSON.stringify(patch),
        });
        return textResult(
          `Event updated — ${what} (${Object.keys(patch).join(", ")}).\n` +
            `Subject: ${updated.subject || "(no subject)"}\n` +
            `When: ${String(updated.start?.dateTime ?? "").slice(0, 16).replace("T", " ")}–${String(updated.end?.dateTime ?? "").slice(11, 16)} (${TIMEZONE})\n` +
            (updated.recurrence ? `${describeRecurrence(updated.recurrence)}\n` : "") +
            (args.reminder_minutes !== undefined
              ? `Reminder: ${
                  args.reminder_minutes < 0
                    ? "off"
                    : args.reminder_minutes === 0
                      ? "at start time"
                      : `${args.reminder_minutes} min before`
                }\n`
              : "") +
            `Event id: ${updated.id}\n` +
            (attendeeCount
              ? `Note: ${attendeeCount} attendee(s) are being notified of this change` +
                (seriesWide ? " to every occurrence of the series." : ".")
              : "No attendees — no notifications sent.")
        );
      }

      case "cancel": {
        const scopeSuffix = seriesWide
          ? " (the entire series)"
          : isOccurrence
            ? " (this occurrence only; the rest of the series is untouched)"
            : "";
        if (event.isOrganizer && attendeeCount > 0) {
          await callGraphServer(`${base}/cancel`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(args.comment ? { comment: args.comment } : {}),
          });
          return textResult(
            `Event ${label} cancelled${scopeSuffix}. ${attendeeCount} attendee(s) are being sent cancellation notices.`
          );
        }
        if (event.isOrganizer) {
          await callGraphServer(base, { method: "DELETE" });
          return textResult(
            `Event ${label} removed${scopeSuffix} (moved to Deleted Items). No attendees — no notifications sent.`
          );
        }
        // Attendee "cancelling" = declining the invitation.
        await callGraphServer(`${base}/decline`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sendResponse: true, ...(args.comment ? { comment: args.comment } : {}) }),
        });
        return textResult(
          `You are an attendee of ${label}, not the organizer — the invitation was declined instead${scopeSuffix} (organizer notified).`
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
