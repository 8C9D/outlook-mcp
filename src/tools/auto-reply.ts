import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { TIMEZONE, ToolResult, errorResult, formatLocal, runTool, textResult } from "./common.js";
import { toGraphDateTime } from "./create-event.js";

export const autoReplySchema = {
  action: z
    .enum(["get", "set", "clear"])
    .describe("get: current auto-reply state; set: enable with a message; clear: disable."),
  message: z
    .string()
    .optional()
    .describe("set: the auto-reply text shown to people who email you. Required for set."),
  external_message: z
    .string()
    .optional()
    .describe(
      "set: a different text for senders outside the user's organization/contacts. Defaults to the same as message."
    ),
  start: z
    .string()
    .optional()
    .describe(
      'set: ISO datetime when auto-reply starts, e.g. "2026-08-20T17:00" (America/Toronto unless an offset is given). Omit start/end to enable immediately until cleared.'
    ),
  end: z.string().optional().describe("set: ISO datetime when auto-reply stops."),
};

const autoReplyArgs = z.object(autoReplySchema);

export const autoReplyDescription =
  "Get, set, or clear the mailbox's automatic reply (out-of-office). CAUTION — outward-facing: set and clear change what EVERY person who emails this account receives, immediately. Before calling set or clear, state the exact auto-reply text (or that it will be turned off). set without start/end enables the reply until it is cleared.";

function describeSetting(s: any): string {
  const status = s?.status ?? "disabled";
  if (status === "disabled") return "Auto-reply is OFF.";
  const lines = [`Auto-reply is ON (${status}).`];
  if (status === "scheduled") {
    lines.push(
      `Window: ${formatLocal(s.scheduledStartDateTime?.dateTime)} to ${formatLocal(s.scheduledEndDateTime?.dateTime)} (${s.scheduledStartDateTime?.timeZone ?? TIMEZONE})`
    );
  }
  const internal = (s.internalReplyMessage ?? "").replace(/<[^>]+>/g, "").trim();
  const external = (s.externalReplyMessage ?? "").replace(/<[^>]+>/g, "").trim();
  lines.push(`Internal message: ${internal || "(empty)"}`);
  if (external && external !== internal) lines.push(`External message: ${external}`);
  lines.push(`External audience: ${s.externalAudience ?? "none"}`);
  return lines.join("\n");
}

export async function autoReplyHandler(
  input: z.input<typeof autoReplyArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, message, external_message, start, end } = autoReplyArgs.parse(input);

    if (action === "get") {
      const settings = await callGraphServer(
        "/me/mailboxSettings?$select=automaticRepliesSetting"
      );
      return textResult(describeSetting(settings?.automaticRepliesSetting));
    }

    let setting: any;
    if (action === "set") {
      if (!message) return errorResult('Action "set" requires message.');
      if ((start === undefined) !== (end === undefined)) {
        return errorResult("Provide both start and end for a scheduled window, or neither.");
      }
      setting = {
        status: start !== undefined ? "scheduled" : "alwaysEnabled",
        internalReplyMessage: message,
        externalReplyMessage: external_message ?? message,
        externalAudience: "all",
      };
      if (start !== undefined) {
        const startObj = toGraphDateTime(start);
        const endObj = toGraphDateTime(end!);
        if (!startObj) return errorResult(`Could not parse start datetime: ${JSON.stringify(start)}.`);
        if (!endObj) return errorResult(`Could not parse end datetime: ${JSON.stringify(end)}.`);
        setting.scheduledStartDateTime = startObj;
        setting.scheduledEndDateTime = endObj;
      }
    } else {
      setting = { status: "disabled" };
    }

    const result = await callGraphServer("/me/mailboxSettings", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ automaticRepliesSetting: setting }),
    });
    return textResult(
      `Auto-reply ${action === "set" ? "enabled" : "cleared"}.\n` +
        describeSetting(result?.automaticRepliesSetting ?? setting)
    );
  });
}
