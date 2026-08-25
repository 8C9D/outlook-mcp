import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { TIMEZONE, TZ_PREFER, ToolResult, errorResult, runTool, textResult } from "./common.js";
import { describeRecurrence, recurrenceSchema, toGraphDateTime, toGraphRecurrence } from "./create-event.js";
import { addDays } from "./list-events.js";
import { resolveCategoryNames } from "./manage-categories.js";

export const manageEventSchema = {
  event_id: z.string().min(1).describe("The id of the event to act on."),
  action: z
    .enum(["update", "cancel", "respond", "forward"])
    .describe(
      "update: change event fields; cancel: cancel/remove the event (as organizer) or decline it (as attendee); respond: accept/decline/tentative an invitation, optionally proposing a new time; forward: email the event to forward_to so they can join or copy it."
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
  show_as: z
    .enum(["free", "tentative", "busy", "oof", "workingElsewhere"])
    .optional()
    .describe('update: how the event blocks the free/busy view ("oof" = out of office).'),
  private: z
    .boolean()
    .optional()
    .describe("update: mark the event Private (true) or Normal (false)."),
  categories: z
    .array(z.string().min(1))
    .optional()
    .describe(
      "update: REPLACE the event's categories with these names (empty array clears them). Every name must exist in the mailbox's category list (manage_categories)."
    ),
  attendees: z
    .array(z.string().email())
    .optional()
    .describe(
      "update: REPLACE the event's whole attendee list — these become the Required attendees, optional_attendees the Optional ones (pass both together; an empty attendees array with no optional_attendees removes everyone). Added attendees are emailed an invitation; removed ones a cancellation."
    ),
  optional_attendees: z
    .array(z.string().email())
    .optional()
    .describe("update: the Optional attendees of the replaced list (see attendees)."),
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
  proposed_start: z
    .string()
    .optional()
    .describe(
      'respond with "tentative" or "decline": propose a different start to the organizer, as an ISO datetime (America/Toronto when no UTC offset is given). Requires proposed_end, and the organizer must allow new-time proposals.'
    ),
  proposed_end: z
    .string()
    .optional()
    .describe("respond: the proposed new end, paired with proposed_start."),
  forward_to: z
    .array(z.string().email())
    .min(1)
    .optional()
    .describe('forward: who to send the event to (required for action "forward").'),
};

const manageEventArgs = z.object(manageEventSchema);

export const manageEventDescription =
  "Update, cancel, respond to, or forward a calendar event, including one occurrence of a repeating event or a whole series. update covers times, subject, location, body, reminder, show_as (free/busy status), private, categories, the repeat rule — and the attendee list (replaced wholesale; changes email the people affected). respond accepts/declines/tentatives an invitation, optionally proposing a new time (proposed_start/proposed_end with tentative or decline). forward emails the event to forward_to. CAUTION — visible to other people: on events with attendees, updates and cancellations send notification emails to every attendee, responses notify the organizer, and forward emails the event to new people. Editing or cancelling an entire series notifies every attendee about all of its occurrences (and a changed repeat rule re-issues the whole series); editing one occurrence notifies them about that date only. Before calling, state the event's subject, date, whether you are touching one occurrence or the series, and what will change. cancel picks the right operation automatically: organizer with attendees → cancellation notices; organizer without attendees → the event is removed (to Deleted Items); attendee → decline.";

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
        if (args.show_as !== undefined) patch.showAs = args.show_as;
        if (args.private !== undefined) patch.sensitivity = args.private ? "private" : "normal";
        if (args.categories !== undefined) {
          patch.categories = args.categories.length
            ? await resolveCategoryNames(args.categories)
            : [];
        }
        if (args.optional_attendees !== undefined && args.attendees === undefined) {
          return errorResult(
            "optional_attendees replaces the attendee list together with attendees — pass both " +
              "(attendees may be an empty array), so the required attendees are not wiped by accident."
          );
        }
        if (args.attendees !== undefined) {
          patch.attendees = [
            ...args.attendees.map((a) => ({ emailAddress: { address: a }, type: "required" })),
            ...(args.optional_attendees ?? []).map((a) => ({
              emailAddress: { address: a },
              type: "optional",
            })),
          ];
        }
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
              "all_day, reminder_minutes, show_as, private, categories, attendees, recurrence."
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
            (patch.attendees
              ? `Attendees replaced: now ${(updated.attendees ?? []).length} (was ${attendeeCount}). ` +
                "Added attendees are being invited; removed ones are being sent cancellations."
              : attendeeCount
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
        let proposedNewTime: { start: any; end: any } | undefined;
        if (args.proposed_start !== undefined || args.proposed_end !== undefined) {
          if (!args.proposed_start || !args.proposed_end) {
            return errorResult("A time proposal needs both proposed_start and proposed_end.");
          }
          if (args.response === "accept") {
            return errorResult(
              'A new time can only be proposed with response "tentative" or "decline" — accepting ' +
                "means the current time works."
            );
          }
          const start = toGraphDateTime(args.proposed_start);
          const end = toGraphDateTime(args.proposed_end);
          if (!start) return errorResult(`Could not parse proposed_start: ${JSON.stringify(args.proposed_start)}.`);
          if (!end) return errorResult(`Could not parse proposed_end: ${JSON.stringify(args.proposed_end)}.`);
          proposedNewTime = { start, end };
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
            // Graph requires sendResponse: true when a new time is proposed —
            // a proposal the organizer never sees would be no proposal at all.
            sendResponse: proposedNewTime ? true : args.send_response,
            ...(args.comment ? { comment: args.comment } : {}),
            ...(proposedNewTime ? { proposedNewTime } : {}),
          }),
        });
        return textResult(
          `Responded "${args.response}" to ${label}.` +
            (proposedNewTime
              ? ` A new time was proposed to the organizer: ${args.proposed_start} to ${args.proposed_end} — they decide whether to take it.`
              : args.send_response
                ? " The organizer is being notified."
                : " No response sent to the organizer.")
        );
      }

      case "forward": {
        if (!args.forward_to?.length) {
          return errorResult('Action "forward" requires forward_to (email addresses).');
        }
        await callGraphServer(`${base}/forward`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            toRecipients: args.forward_to.map((a) => ({ emailAddress: { address: a } })),
            ...(args.comment ? { comment: args.comment } : {}),
          }),
        });
        return textResult(
          `Event ${label} forwarded to ${args.forward_to.join(", ")} — they receive it by email ` +
            "and can add it to their calendar. (On a personal account the organizer is not " +
            "otherwise notified of the forward.)"
        );
      }
    }
  });
}
