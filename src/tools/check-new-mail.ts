import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { deltaKey } from "../core/kv-keys.js";
import { readJson, requireStateStore, writeJson } from "../core/state.js";
import {
  TZ_PREFER,
  ToolInputError,
  ToolResult,
  formatLocal,
  formatSender,
  runTool,
  textResult,
} from "./common.js";

export const checkNewMailSchema = {
  folder: z
    .string()
    .default("inbox")
    .describe(
      'Mail folder to watch. Well-known names: "inbox" (default), "archive", "junkemail"; also accepts a folder id. Each folder keeps its own independent position.'
    ),
  reset: z
    .boolean()
    .default(false)
    .describe(
      "Throw away the stored position and start again from now, reporting nothing this call (default false). Use when the stored position is stale or you want a clean baseline."
    ),
  max_results: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(25)
    .describe("Maximum number of changes to describe (default 25, max 50)."),
};

const checkNewMailArgs = z.object(checkNewMailSchema);

export const checkNewMailDescription =
  "Report what has changed in a mail folder since the last time this tool was called, using a Microsoft Graph delta query. The first call (or one with reset) only establishes the starting position and lists nothing; every later call returns just the messages added, changed or removed since — no re-reading of the whole folder. Use this for \"anything new?\" follow-ups instead of re-listing the inbox with search_mail. The position advances on every successful call, so each change is reported exactly once.";

/** Delta pages are small by default; the Prefer header must ride every request. */
const DELTA_PAGE_PREFER = "odata.maxpagesize=500";
const DELTA_SELECT = "id,subject,from,receivedDateTime,isRead";

/** Hard stop on a first-time enumeration of a very large folder. */
const MAX_PAGES = 200;

type StoredDelta = { deltaLink: string; updatedAt: string; folder: string };

type DeltaDrain = { items: any[]; deltaLink: string; pages: number; truncated: boolean };

/** Follow @odata.nextLink until Graph hands back the @odata.deltaLink. */
async function drainDelta(firstPath: string): Promise<DeltaDrain> {
  const headers = { Prefer: `${DELTA_PAGE_PREFER}, ${TZ_PREFER}` };
  const items: any[] = [];
  let next: string | undefined = firstPath;
  let deltaLink: string | undefined;
  let pages = 0;
  while (next && pages < MAX_PAGES) {
    const page: any = await callGraphServer(next, { headers });
    pages++;
    items.push(...(page?.value ?? []));
    deltaLink = page?.["@odata.deltaLink"];
    next = page?.["@odata.nextLink"];
  }
  if (!deltaLink) {
    throw new ToolInputError(
      `Microsoft Graph did not finish the delta enumeration of "${firstPath}" within ${MAX_PAGES} pages. ` +
        "The folder is too large to baseline in one call; watch a smaller folder."
    );
  }
  return { items, deltaLink, pages, truncated: !!next };
}

/** True for the tombstone Graph emits when an item leaves the folder. */
function isRemoved(item: any): boolean {
  return !!item["@removed"];
}

export async function checkNewMailHandler(
  input: z.input<typeof checkNewMailArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { folder, reset, max_results } = checkNewMailArgs.parse(input);
    const store = requireStateStore();
    const key = deltaKey(folder);
    const previous = await readJson<StoredDelta>(store, key);
    const stored = reset ? null : previous;

    if (!stored?.deltaLink) {
      const baseline = await drainDelta(
        `/me/mailFolders/${encodeURIComponent(folder)}/messages/delta?$select=${DELTA_SELECT}`
      );
      await writeJson(store, key, {
        deltaLink: baseline.deltaLink,
        updatedAt: new Date().toISOString(),
        folder,
      } satisfies StoredDelta);
      return textResult(
        `Starting position recorded for ${folder} (${baseline.items.length} message(s) currently there).` +
          `${reset && previous?.deltaLink ? " Previous position discarded." : ""}\n` +
          "Nothing is reported on this call. Call check_new_mail again to see what has changed since now."
      );
    }

    const changes = await drainDelta(stored.deltaLink);
    await writeJson(store, key, {
      deltaLink: changes.deltaLink,
      updatedAt: new Date().toISOString(),
      folder,
    } satisfies StoredDelta);

    const since = formatLocal(stored.updatedAt);
    if (changes.items.length === 0) {
      return textResult(`No changes in ${folder} since ${since}.`);
    }

    const removed = changes.items.filter(isRemoved);
    const present = changes.items.filter((item) => !isRemoved(item));

    // A delta entry for an edited message carries only the properties that
    // changed, so anything without a subject is looked up to stay readable.
    const shown = present.slice(0, max_results);
    for (const item of shown) {
      if (item.subject !== undefined || !item.id) continue;
      try {
        const full = await callGraphServer(
          `/me/messages/${encodeURIComponent(item.id)}?$select=${DELTA_SELECT}`,
          { headers: { Prefer: TZ_PREFER } }
        );
        Object.assign(item, { ...full, id: item.id, partial: true });
      } catch {
        // Message already gone: report what the delta entry itself carried.
      }
    }

    const lines = shown.map((item, index) => {
      const state = item.isRead === false ? "  [unread]" : "";
      return [
        `${index + 1}. ${item.subject ?? "(subject not reported)"}${item.partial ? "  [changed]" : ""}`,
        `   From: ${formatSender(item.from)}  At: ${formatLocal(item.receivedDateTime)}${state}`,
        `   Message id: ${item.id}`,
      ].join("\n");
    });

    const header =
      `${present.length} new or changed message(s) in ${folder} since ${since}` +
      (removed.length ? `, ${removed.length} removed` : "") +
      (present.length > shown.length ? ` (showing ${shown.length})` : "") +
      ":";
    return textResult(`${header}\n\n${lines.join("\n\n")}`);
  });
}
