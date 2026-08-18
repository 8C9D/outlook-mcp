// Mailbox settings other than the automatic reply: working hours and the
// Focused-Inbox per-sender overrides.
//
// auto_reply deliberately stays its own tool. Its get/set/clear vocabulary is
// about one thing (the out-of-office message) and carries an outward-facing
// caution that belongs to that tool alone; folding it in here would have made
// "set" mean four different things and would have broken every existing caller.
// This tool's "get" therefore reports the auto-reply state read-only and points
// at auto_reply for changes. See ASSUMPTIONS.md (Batch B) for the decision.
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { WEEKDAYS } from "./create-event.js";

export const mailboxSettingsSchema = {
  action: z
    .enum(["get", "set_working_hours", "set_focus_override", "clear_focus_override"])
    .describe(
      "get: read the mailbox's time zone, working hours, Focused-Inbox overrides and auto-reply status; set_working_hours: change the working days and hours; set_focus_override: always file a sender's mail in Focused or Other; clear_focus_override: drop that sender's override and let Outlook decide again."
    ),
  days: z
    .array(z.enum(WEEKDAYS))
    .min(1)
    .optional()
    .describe(
      'set_working_hours: the working days, e.g. ["monday","tuesday","wednesday","thursday","friday"]. Omit to keep the current days.'
    ),
  start_time: z
    .string()
    .regex(/^\d{2}:\d{2}$/, "24-hour time (HH:MM)")
    .optional()
    .describe("set_working_hours: start of the working day as HH:MM (24-hour), e.g. \"09:00\"."),
  end_time: z
    .string()
    .regex(/^\d{2}:\d{2}$/, "24-hour time (HH:MM)")
    .optional()
    .describe("set_working_hours: end of the working day as HH:MM (24-hour), e.g. \"17:30\"."),
  sender: z
    .string()
    .email()
    .optional()
    .describe(
      "set_focus_override / clear_focus_override: the sender's email address. Required for both."
    ),
  classify_as: z
    .enum(["focused", "other"])
    .optional()
    .describe(
      'set_focus_override: where this sender\'s mail always lands — "focused" (the Focused tab) or "other". Required for set_focus_override.'
    ),
};

const mailboxSettingsArgs = z.object(mailboxSettingsSchema);

export const mailboxSettingsDescription =
  "Read and change mailbox settings: working hours (which days and hours the user works) and Focused-Inbox overrides (senders pinned to the Focused or Other tab). CAUTION — working hours are not private: they drive the free/busy view and the meeting times Outlook suggests, so anyone who can schedule with this account sees the effect of a change. Focused-Inbox overrides are local to this mailbox and invisible to senders, but they change where mail lands, so a sender pinned to \"other\" will be easy to miss. Use get before changing anything, and state the exact before/after to the user. The automatic reply (out-of-office) is reported here read-only — change it with the auto_reply tool.";

/** "08:00:00.0000000" → "08:00"; anything unexpected is passed through. */
function shortTime(value: string | undefined): string {
  return /^\d{2}:\d{2}/.test(value ?? "") ? value!.slice(0, 5) : (value ?? "(unset)");
}

/** "09:00" → the "09:00:00.0000000" Graph stores. */
function graphTime(value: string): string {
  return `${value}:00.0000000`;
}

function describeWorkingHours(wh: any): string {
  if (!wh) return "Working hours: not set for this mailbox.";
  const days: string[] = wh.daysOfWeek ?? [];
  return (
    `Working hours: ${days.length ? days.join(", ") : "(no days set — effectively off)"} ` +
    `${shortTime(wh.startTime)}–${shortTime(wh.endTime)} (${wh.timeZone?.name ?? "unknown time zone"})`
  );
}

function describeOverrides(overrides: any[]): string {
  if (overrides.length === 0) {
    return "Focused-Inbox overrides: none (Outlook classifies every sender on its own).";
  }
  const lines = overrides.map(
    (o) =>
      `  ${o.senderEmailAddress?.address ?? "(unknown)"} → ${o.classifyAs}` +
      (o.senderEmailAddress?.name && o.senderEmailAddress.name !== o.senderEmailAddress.address
        ? ` (${o.senderEmailAddress.name})`
        : "")
  );
  return `Focused-Inbox overrides (${overrides.length}):\n${lines.join("\n")}`;
}

/** Every override, paged; the collection is small on a personal mailbox. */
async function readOverrides(): Promise<any[]> {
  const data = await callGraphServer("/me/inferenceClassification/overrides?$top=100");
  return data?.value ?? [];
}

export async function mailboxSettingsHandler(
  input: z.input<typeof mailboxSettingsArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, days, start_time, end_time, sender, classify_as } =
      mailboxSettingsArgs.parse(input);

    if (action === "get") {
      const settings = await callGraphServer(
        "/me/mailboxSettings?$select=timeZone,workingHours,automaticRepliesSetting"
      );
      const overrides = await readOverrides();
      const autoReply = settings?.automaticRepliesSetting?.status ?? "disabled";
      return textResult(
        `Mailbox time zone: ${settings?.timeZone ?? "(unknown)"}\n` +
          `${describeWorkingHours(settings?.workingHours)}\n` +
          `${describeOverrides(overrides)}\n` +
          `Auto-reply: ${autoReply === "disabled" ? "OFF" : `ON (${autoReply})`} — read/change it with the auto_reply tool.`
      );
    }

    if (action === "set_working_hours") {
      if (!days && !start_time && !end_time) {
        return errorResult(
          'Action "set_working_hours" needs at least one of days, start_time, or end_time.'
        );
      }
      if (start_time && end_time && end_time <= start_time) {
        return errorResult(
          `end_time (${end_time}) must be later than start_time (${start_time}) — Outlook has no overnight working day.`
        );
      }
      const current = (
        await callGraphServer("/me/mailboxSettings?$select=workingHours")
      )?.workingHours;
      if (start_time && !end_time && shortTime(current?.endTime) <= start_time) {
        return errorResult(
          `start_time (${start_time}) is not earlier than the current end of day (${shortTime(current?.endTime)}) — pass end_time too.`
        );
      }
      if (end_time && !start_time && end_time <= shortTime(current?.startTime)) {
        return errorResult(
          `end_time (${end_time}) is not later than the current start of day (${shortTime(current?.startTime)}) — pass start_time too.`
        );
      }
      // Graph replaces workingHours wholesale, so unspecified fields are carried
      // over from what is there rather than silently reset. The time zone is
      // never changed here: Graph normalizes it to the mailbox's own zone anyway.
      const workingHours = {
        daysOfWeek: days ?? current?.daysOfWeek ?? [],
        startTime: start_time ? graphTime(start_time) : (current?.startTime ?? "08:00:00.0000000"),
        endTime: end_time ? graphTime(end_time) : (current?.endTime ?? "17:00:00.0000000"),
        ...(current?.timeZone ? { timeZone: current.timeZone } : {}),
      };
      const updated = await callGraphServer("/me/mailboxSettings", {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ workingHours }),
      });
      return textResult(
        "Working hours updated.\n" +
          `Before: ${describeWorkingHours(current).replace(/^Working hours: /, "")}\n` +
          `After:  ${describeWorkingHours(updated?.workingHours ?? workingHours).replace(/^Working hours: /, "")}\n` +
          "This is visible to anyone who schedules with this account."
      );
    }

    if (!sender) {
      return errorResult(`Action "${action}" requires sender (an email address).`);
    }
    const overrides = await readOverrides();
    const existing = overrides.find(
      (o) => String(o.senderEmailAddress?.address ?? "").toLowerCase() === sender.toLowerCase()
    );

    if (action === "set_focus_override") {
      if (!classify_as) {
        return errorResult('Action "set_focus_override" requires classify_as ("focused" or "other").');
      }
      if (existing) {
        // Graph refuses a second override for the same address, so change the
        // one that is already there rather than reporting a duplicate.
        const patched = await callGraphServer(
          `/me/inferenceClassification/overrides/${encodeURIComponent(existing.id)}`,
          {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ classifyAs: classify_as }),
          }
        );
        return textResult(
          `Focused-Inbox override updated: ${sender} → ${patched?.classifyAs ?? classify_as} ` +
            `(was ${existing.classifyAs}).\nOverride id: ${existing.id}`
        );
      }
      const created = await callGraphServer("/me/inferenceClassification/overrides", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          classifyAs: classify_as,
          senderEmailAddress: { name: sender, address: sender },
        }),
      });
      return textResult(
        `Focused-Inbox override added: ${sender} → ${created.classifyAs}.\n` +
          `Every future message from that address lands in ${created.classifyAs === "focused" ? "Focused" : "Other"}.\n` +
          `Override id: ${created.id}`
      );
    }

    if (!existing) {
      return errorResult(
        `There is no Focused-Inbox override for ${sender}.\n${describeOverrides(overrides)}`
      );
    }
    await callGraphServer(
      `/me/inferenceClassification/overrides/${encodeURIComponent(existing.id)}`,
      { method: "DELETE" }
    );
    return textResult(
      `Focused-Inbox override for ${sender} removed (it was "${existing.classifyAs}"). ` +
        "Outlook classifies that sender on its own again."
    );
  });
}
