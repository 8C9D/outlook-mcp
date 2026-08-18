import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolInputError, ToolResult, runTool, textResult } from "./common.js";

export const listCalendarsSchema = {
  // MCP clients reject empty schemas less gracefully than a harmless optional flag.
  writable_only: z
    .boolean()
    .default(false)
    .describe(
      "Only list calendars events can be added to, hiding read-only ones such as subscribed holiday calendars (default false)."
    ),
};

const listCalendarsArgs = z.object(listCalendarsSchema);

export const listCalendarsDescription =
  "List the calendars in this Outlook account with their names and ids, marking the default one and whether each can be written to. Use a name (or id) from here as the `calendar` input of create_event and list_events; omit that input to work with the default calendar.";

export type CalendarRef = { id: string; name: string; isDefault: boolean };

const SELECT = "id,name,isDefaultCalendar,canEdit,owner";

/**
 * Resolve a `calendar` input (name or id, or nothing) to one calendar. One GET
 * covers every case: an account has few calendars, and matching locally means a
 * bad name reports the available calendars instead of a bare 404.
 */
export async function resolveCalendar(calendar: string | undefined): Promise<CalendarRef | undefined> {
  if (!calendar) return undefined; // the default calendar: /me/events, no lookup needed
  const data = await callGraphServer(`/me/calendars?$select=${SELECT}&$top=100`);
  const calendars: any[] = data?.value ?? [];
  const wanted = calendar.toLowerCase();
  const match =
    calendars.find((c) => c.id === calendar) ??
    calendars.find((c) => String(c.name ?? "").toLowerCase() === wanted) ??
    calendars.find((c) => String(c.name ?? "").toLowerCase().startsWith(wanted));
  if (!match) {
    const names = calendars.map((c) => JSON.stringify(c.name ?? "")).join(", ");
    throw new ToolInputError(
      `No calendar named or identified by ${JSON.stringify(calendar)}. Available calendars: ${names}.`
    );
  }
  return { id: match.id, name: match.name ?? "(unnamed)", isDefault: !!match.isDefaultCalendar };
}

/** The Graph path prefix for a calendar's events/calendarView, default calendar included. */
export function calendarBasePath(calendar: CalendarRef | undefined): string {
  return calendar ? `/me/calendars/${encodeURIComponent(calendar.id)}` : "/me";
}

export async function listCalendarsHandler(
  input: z.input<typeof listCalendarsArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { writable_only } = listCalendarsArgs.parse(input);
    const data = await callGraphServer(`/me/calendars?$select=${SELECT}&$top=100`);
    const calendars: any[] = (data?.value ?? []).filter((c: any) => !writable_only || c.canEdit);

    if (calendars.length === 0) {
      return textResult(
        writable_only ? "No writable calendars in this account." : "No calendars in this account."
      );
    }

    const lines = calendars.map((c) => {
      const flags = [
        c.isDefaultCalendar ? "default" : null,
        c.canEdit ? null : "read-only",
      ].filter(Boolean);
      return (
        `${c.name || "(unnamed)"}${flags.length ? ` — ${flags.join(", ")}` : ""}\n` +
        `  id: ${c.id}`
      );
    });
    return textResult(
      `Calendars (${calendars.length}${writable_only ? ", writable only" : ""}):\n\n${lines.join("\n")}`
    );
  });
}
