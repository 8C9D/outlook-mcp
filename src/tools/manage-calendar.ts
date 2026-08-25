// Calendar management beyond listing: create, rename, recolor.
//
// There is deliberately NO delete action, for the same reason manage_task
// cannot delete a To Do list: deleting a calendar destroys every event in it
// with no recoverable copy anywhere (Graph's calendar DELETE is permanent on a
// personal account — no Deleted Items, no recycle bin), which is exactly the
// outcome the soft-delete policy exists to prevent. Someone who really wants
// that can do it in Outlook. Calendar SHARING is also not offered: granting a
// third party standing access could not be live-verified without actually
// granting it, and standing outbound access grants sit outside this server's
// design (the same instinct that keeps forwarding out of inbox rules).
import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { resolveCalendar } from "./list-calendars.js";

/** Graph's calendar palette ("auto" = let Outlook pick). Verified live: PATCH color works. */
export const CALENDAR_COLORS = [
  "auto",
  "lightBlue",
  "lightGreen",
  "lightOrange",
  "lightGray",
  "lightYellow",
  "lightTeal",
  "lightPink",
  "lightBrown",
  "lightRed",
] as const;

export const manageCalendarSchema = {
  action: z
    .enum(["create", "rename", "set_color"])
    .describe(
      "create: a new calendar named name; rename: change calendar's name to name; set_color: change calendar's colour. (Deleting a calendar is deliberately not offered — it would permanently destroy every event in it; do that in Outlook if it is really wanted.)"
    ),
  calendar: z
    .string()
    .optional()
    .describe(
      "rename / set_color: which calendar to change — its name or id (see list_calendars)."
    ),
  name: z
    .string()
    .min(1)
    .max(255)
    .optional()
    .describe("create: the new calendar's name. rename: the name to change it to."),
  color: z
    .enum(CALENDAR_COLORS)
    .optional()
    .describe('set_color (or create): the calendar\'s colour in Outlook, e.g. "lightGreen".'),
};

const manageCalendarArgs = z.object(manageCalendarSchema);

export const manageCalendarDescription =
  "Create, rename, or recolor a calendar in this Outlook account. create makes an additional calendar (events go to it via create_event's calendar input); rename and set_color change an existing one, found by name or id. Renaming or recoloring never touches the events. This tool deliberately CANNOT delete a calendar — deleting one permanently destroys every event in it with no recoverable copy, so that is left to Outlook itself. The default calendar cannot be renamed.";

export async function manageCalendarHandler(
  input: z.input<typeof manageCalendarArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, calendar, name, color } = manageCalendarArgs.parse(input);

    if (action === "create") {
      if (!name) return errorResult('Action "create" requires name.');
      const existing = await callGraphServer("/me/calendars?$select=id,name&$top=100");
      const clash = (existing?.value ?? []).find(
        (c: any) => String(c.name ?? "").toLowerCase() === name.toLowerCase()
      );
      if (clash) {
        return errorResult(
          `A calendar named "${clash.name}" already exists (id ${clash.id}) — use it, or pick another name.`
        );
      }
      const created = await callGraphServer("/me/calendars", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, ...(color ? { color } : {}) }),
      });
      return textResult(
        `Calendar "${created.name}" created${color ? ` (${created.color ?? color})` : ""}.\n` +
          `Calendar id: ${created.id}\n` +
          "Pass its name as the calendar input of create_event and list_events to use it."
      );
    }

    if (!calendar) {
      return errorResult(`Action "${action}" requires calendar (a name or id from list_calendars).`);
    }
    const target = await resolveCalendar(calendar);
    if (!target) return errorResult("Could not resolve the calendar."); // unreachable: calendar was given

    if (action === "rename") {
      if (!name) return errorResult('Action "rename" requires name (the new name).');
      if (target.isDefault) {
        return errorResult(
          `"${target.name}" is the account's default calendar — Outlook does not allow renaming it.`
        );
      }
      const renamed = await callGraphServer(`/me/calendars/${encodeURIComponent(target.id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name }),
      });
      return textResult(
        `Calendar renamed from "${target.name}" to "${renamed.name}". Its events are untouched.\n` +
          `Calendar id: ${renamed.id}`
      );
    }

    if (!color) return errorResult('Action "set_color" requires color.');
    const recolored = await callGraphServer(`/me/calendars/${encodeURIComponent(target.id)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ color }),
    });
    return textResult(
      `Calendar "${recolored.name ?? target.name}" is now ${recolored.color ?? color}.\n` +
        `Calendar id: ${target.id}`
    );
  });
}
