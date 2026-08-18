import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";

const CATEGORIES_PATH = "/me/outlook/masterCategories";

/** Graph's fixed palette for master categories: preset id → the colour Outlook shows. */
export const PRESET_COLORS = {
  preset0: "red",
  preset1: "orange",
  preset2: "brown",
  preset3: "yellow",
  preset4: "green",
  preset5: "teal",
  preset6: "olive",
  preset7: "blue",
  preset8: "purple",
  preset9: "cranberry",
  preset10: "steel",
  preset11: "dark steel",
  preset12: "gray",
  preset13: "dark gray",
  preset14: "black",
  preset15: "dark red",
  preset16: "dark orange",
  preset17: "dark brown",
  preset18: "dark yellow",
  preset19: "dark green",
  preset20: "dark teal",
  preset21: "dark olive",
  preset22: "dark blue",
  preset23: "dark purple",
  preset24: "dark cranberry",
  none: "no colour",
} as const;

type PresetId = keyof typeof PRESET_COLORS;
const PRESET_IDS = Object.keys(PRESET_COLORS) as [PresetId, ...PresetId[]];

export const manageCategoriesSchema = {
  action: z
    .enum(["list", "create", "delete"])
    .describe(
      "list: the mailbox's category names, colours, and ids; create: a new category (display_name + color); delete: remove a category by category_id."
    ),
  display_name: z
    .string()
    .min(1)
    .optional()
    .describe('Category name (required for action "create"). Names are unique per mailbox.'),
  color: z
    .enum(PRESET_IDS)
    .optional()
    .describe(
      'Colour for action "create" — one of Graph\'s preset ids: ' +
        Object.entries(PRESET_COLORS)
          .map(([id, name]) => `${id}=${name}`)
          .join(", ") +
        "."
    ),
  category_id: z
    .string()
    .min(1)
    .optional()
    .describe('The category to remove (required for action "delete"; from action "list").'),
};

const manageCategoriesArgs = z.object(manageCategoriesSchema);

export const manageCategoriesDescription =
  "List, create, or delete the mailbox's Outlook categories (the coloured labels applied to mail). Use manage_message's categorize action to put categories on messages; this tool only manages the master list. Deleting a category removes it from the master list but does NOT strip it from messages that already carry it — those keep the name, without a colour. Colours come from Graph's fixed preset palette.";

/** Fetch the mailbox's master categories. Exported for manage_message's validation. */
export async function fetchMasterCategories(): Promise<any[]> {
  const data = await callGraphServer(`${CATEGORIES_PATH}?$top=100`);
  return data?.value ?? [];
}

export async function manageCategoriesHandler(
  input: z.input<typeof manageCategoriesArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, display_name, color, category_id } = manageCategoriesArgs.parse(input);

    if (action === "list") {
      const categories = await fetchMasterCategories();
      if (categories.length === 0) return textResult("No categories defined in this mailbox.");
      const lines = categories.map(
        (c, i) =>
          `${i + 1}. ${c.displayName || "(unnamed)"} — ${describeColor(c.color)}\n` +
          `   Category id: ${c.id}`
      );
      return textResult(`${categories.length} categor(y/ies):\n\n${lines.join("\n")}`);
    }

    if (action === "delete") {
      if (!category_id) {
        return errorResult('Action "delete" requires category_id (from action "list").');
      }
      const existing = await fetchMasterCategories();
      const target = existing.find((c) => c.id === category_id);
      if (!target) {
        return errorResult(
          `No category with id ${category_id} in this mailbox. Use action "list" to see the current categories.`
        );
      }
      await callGraphServer(`${CATEGORIES_PATH}/${encodeURIComponent(category_id)}`, {
        method: "DELETE",
      });
      return textResult(
        `Category "${target.displayName}" deleted from the master list. ` +
          "Messages already carrying that category keep the name (now without a colour)."
      );
    }

    // action === "create"
    if (!display_name) return errorResult('Action "create" requires display_name.');
    if (!color) {
      return errorResult(
        'Action "create" requires color — one of ' + PRESET_IDS.join(", ") + "."
      );
    }
    const existing = await fetchMasterCategories();
    const clash = existing.find(
      (c) => String(c.displayName ?? "").toLowerCase() === display_name.toLowerCase()
    );
    if (clash) {
      return errorResult(
        `A category named "${clash.displayName}" already exists (id: ${clash.id}, ${describeColor(clash.color)}).`
      );
    }

    const created = await callGraphServer(CATEGORIES_PATH, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ displayName: display_name, color }),
    });
    return textResult(
      `Category created.\n` +
        `Name: ${created.displayName}\n` +
        `Colour: ${describeColor(created.color)}\n` +
        `Category id: ${created.id}`
    );
  });
}

function describeColor(color: string | undefined): string {
  if (!color) return "no colour";
  const name = PRESET_COLORS[color as keyof typeof PRESET_COLORS];
  return name ? `${name} (${color})` : color;
}
