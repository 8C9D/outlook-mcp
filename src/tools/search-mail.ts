import { z } from "zod";
import {
  TIMEZONE,
  TZ_PREFER,
  ToolResult,
  errorResult,
  fetchPaged,
  formatLocal,
  formatSender,
  runTool,
  structuredResult,
} from "./common.js";

export const searchMailSchema = {
  query: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Search text matched against message subjects, senders, and bodies (Microsoft Graph $search). Plain words work best; results are relevance-ranked, not newest-first. OMIT this entirely for \"latest/most recent mail\" requests — with no query the folder's newest messages are returned newest-first."
    ),
  folder: z
    .string()
    .default("inbox")
    .describe(
      'Mail folder to search. Well-known names: "inbox" (default), "sentitems", "drafts", "archive", "deleteditems", "junkemail". Also accepts a folder id. Ignored when all_folders is true.'
    ),
  all_folders: z
    .boolean()
    .default(false)
    .describe(
      "Search the whole mailbox — every folder, including Sent Items and Deleted Items — instead of one folder (default false). The folder input is ignored when set."
    ),
  date_from: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
    .optional()
    .describe(
      "Only messages received ON or AFTER this ISO date (YYYY-MM-DD), interpreted as an America/Toronto calendar date."
    ),
  date_to: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
    .optional()
    .describe(
      "Only messages received ON or BEFORE this ISO date (YYYY-MM-DD), interpreted as an America/Toronto calendar date."
    ),
  has_attachments: z
    .boolean()
    .optional()
    .describe(
      "true: only messages with attachments; false: only messages without. Omit for both."
    ),
  max_results: z
    .number()
    .int()
    .min(1)
    .max(25)
    .default(10)
    .describe("Maximum number of results to return (default 10, max 25)."),
  include_body_preview: z
    .boolean()
    .default(true)
    .describe("Include a one-line preview of each message body (default true)."),
};

const searchMailArgs = z.object(searchMailSchema);

/** Permissive machine-readable shape of the answer; every field optional. */
export const searchMailOutputSchema = {
  mode: z.enum(["search", "latest"]).optional().describe("How the results were produced."),
  scope: z.string().optional().describe('The folder searched, or "all folders".'),
  query: z.string().optional(),
  filters: z
    .looseObject({
      date_from: z.string().optional(),
      date_to: z.string().optional(),
      has_attachments: z.boolean().optional(),
    })
    .optional(),
  count: z.number().optional(),
  messages: z
    .array(
      z.looseObject({
        subject: z.string().optional(),
        from: z.string().optional(),
        receivedDateTime: z.string().optional().describe("As Graph returned it."),
        receivedLocal: z.string().optional().describe(`Local ${TIMEZONE} rendering.`),
        messageId: z.string().optional(),
        conversationId: z.string().optional(),
        hasAttachments: z.boolean().optional(),
        preview: z.string().optional(),
      })
    )
    .optional(),
};

export const searchMailDescription =
  "Search the user's Outlook mail, or list a folder's newest messages. With query: full-text search, relevance-ranked (use for topical requests). WITHOUT query: the latest messages, genuinely newest-first (use for \"latest/most recent email\" requests — do not invent a query for those). Both modes accept date_from/date_to (America/Toronto calendar dates), has_attachments, and all_folders (whole-mailbox scope, Sent and Deleted Items included). Returns for each hit: subject, sender, received datetime (America/Toronto), message id, conversation id, attachment flag, and optionally a body preview — as text and as structuredContent. Use the returned conversation id with read_thread to read the full conversation, or the message id with create_draft to draft a reply.";

// ---------------------------------------------------------------------------
// Query building. Graph's rules for messages, verified live on this mailbox:
//  - $search cannot be combined with $filter (400 SearchWithFilter) or with
//    $orderby, so query-mode filters must ride INSIDE the KQL: "received>=" /
//    "received<=" and "hasattachments:true|false" all work in $search.
//  - KQL dates are day-granular in an unspecified zone, so the KQL range is
//    widened by one day on each side and the exact America/Toronto boundary is
//    enforced client-side on receivedDateTime.
//  - $filter + $orderby=receivedDateTime works ONLY when receivedDateTime
//    appears in the filter first; a hasAttachments-only filter is refused with
//    InefficientFilter, so latest mode always leads with a receivedDateTime
//    clause (a sentinel "ge 1900-01-01" when no date bound was asked for).
// ---------------------------------------------------------------------------

export type SearchFilters = {
  dateFrom?: string;
  dateTo?: string;
  hasAttachments?: boolean;
};

/** Shift an ISO date by whole days (pure calendar arithmetic, UTC-safe). */
function shiftDate(isoDate: string, days: number): string {
  const d = new Date(`${isoDate}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

/** The wall-clock rendering of an instant in America/Toronto, to the minute. */
function torontoWallClock(when: Date): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(when);
  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? "";
  return `${get("year")}-${get("month")}-${get("day")}T${get("hour")}:${get("minute")}`;
}

/**
 * The UTC instant of midnight America/Toronto on `isoDate`. Toronto is UTC-4
 * or UTC-5; midnight always exists (DST shifts at 02:00), so trying both
 * offsets and checking which one Toronto agrees with is exact year-round.
 */
export function torontoMidnightUtc(isoDate: string): string {
  for (const offset of ["-04:00", "-05:00"]) {
    const candidate = new Date(`${isoDate}T00:00:00${offset}`);
    if (torontoWallClock(candidate) === `${isoDate}T00:00`) return candidate.toISOString();
  }
  // Unreachable for real dates; fall back to the EST reading.
  return new Date(`${isoDate}T00:00:00-05:00`).toISOString();
}

/** The exact UTC window [start, end) a Toronto date range means. */
export function utcWindow(filters: SearchFilters): { fromUtc?: string; toUtcExclusive?: string } {
  return {
    ...(filters.dateFrom ? { fromUtc: torontoMidnightUtc(filters.dateFrom) } : {}),
    ...(filters.dateTo ? { toUtcExclusive: torontoMidnightUtc(shiftDate(filters.dateTo, 1)) } : {}),
  };
}

/**
 * The KQL string for query mode: the caller's query plus server-side narrowing
 * terms. Date bounds are widened ±1 day (exactness comes from the client-side
 * window); the string is quoted and escaped exactly as before.
 */
export function buildSearchKql(query: string, filters: SearchFilters): string {
  const terms = [query];
  if (filters.dateFrom) terms.push(`received>=${shiftDate(filters.dateFrom, -1)}`);
  if (filters.dateTo) terms.push(`received<=${shiftDate(filters.dateTo, 1)}`);
  if (filters.hasAttachments !== undefined) {
    terms.push(`hasattachments:${filters.hasAttachments}`);
  }
  return terms.join(" AND ");
}

/**
 * The $filter expression for latest mode, or undefined when no filter is
 * needed. receivedDateTime clauses come first — Graph refuses
 * $orderby=receivedDateTime unless the ordered property leads the filter.
 */
export function buildLatestFilter(filters: SearchFilters): string | undefined {
  const { fromUtc, toUtcExclusive } = utcWindow(filters);
  const clauses: string[] = [];
  if (fromUtc) clauses.push(`receivedDateTime ge ${fromUtc}`);
  if (toUtcExclusive) clauses.push(`receivedDateTime lt ${toUtcExclusive}`);
  if (filters.hasAttachments !== undefined) {
    if (clauses.length === 0) clauses.push("receivedDateTime ge 1900-01-01T00:00:00Z");
    clauses.push(`hasAttachments eq ${filters.hasAttachments}`);
  }
  return clauses.length > 0 ? clauses.join(" and ") : undefined;
}

/** The exact client-side check both modes apply to every returned message. */
export function matchesFilters(
  message: { receivedDateTime?: string; hasAttachments?: boolean },
  filters: SearchFilters
): boolean {
  if (filters.hasAttachments !== undefined && Boolean(message.hasAttachments) !== filters.hasAttachments) {
    return false;
  }
  const { fromUtc, toUtcExclusive } = utcWindow(filters);
  if (fromUtc || toUtcExclusive) {
    const received = Date.parse(message.receivedDateTime ?? "");
    if (!Number.isFinite(received)) return false;
    if (fromUtc && received < Date.parse(fromUtc)) return false;
    if (toUtcExclusive && received >= Date.parse(toUtcExclusive)) return false;
  }
  return true;
}

export async function searchMailHandler(
  input: z.input<typeof searchMailArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const parsed = searchMailArgs.parse(input);
    const { query, folder, all_folders, date_from, date_to, has_attachments, max_results, include_body_preview } =
      parsed;
    if (date_from && date_to && date_from > date_to) {
      return errorResult(`date_from ${date_from} is after date_to ${date_to}.`);
    }
    const filters: SearchFilters = {
      ...(date_from ? { dateFrom: date_from } : {}),
      ...(date_to ? { dateTo: date_to } : {}),
      ...(has_attachments !== undefined ? { hasAttachments: has_attachments } : {}),
    };
    const filtersActive = date_from !== undefined || date_to !== undefined || has_attachments !== undefined;

    const select = "id,conversationId,subject,from,receivedDateTime,bodyPreview,hasAttachments";
    const base = all_folders
      ? "/me/messages"
      : `/me/mailFolders/${encodeURIComponent(folder)}/messages`;
    const scope = all_folders ? "all folders" : folder;

    let path: string;
    let fetchCap = max_results;
    if (query === undefined) {
      // Latest-mail mode: no $search, so $filter + $orderby are allowed and the
      // filtering is exact server-side (UTC instants of the Toronto dates).
      const filter = buildLatestFilter(filters);
      path =
        `${base}?` +
        (filter ? `$filter=${encodeURIComponent(filter)}&` : "") +
        `$orderby=receivedDateTime%20desc&$select=${select}&$top=${max_results}`;
    } else {
      // KQL string literal: escape embedded backslashes and double quotes. The
      // filters ride inside the KQL ($search combines with neither $filter nor
      // $orderby), day-widened; the exact window is re-applied below, so fetch
      // a few extra in case the widening let boundary hits in.
      const kqlSource = buildSearchKql(query, filters);
      const kql = `"${kqlSource.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
      fetchCap = filtersActive ? Math.min(100, max_results * 3) : max_results;
      path = `${base}?$search=${encodeURIComponent(kql)}&$select=${select}&$top=${Math.min(fetchCap, 25)}`;
    }
    const fetched = await fetchPaged(path, fetchCap, { Prefer: TZ_PREFER });
    // Belt and braces in latest mode; the real boundary enforcement in search
    // mode (KQL dates are coarse and their timezone unspecified).
    const messages = fetched.filter((m) => matchesFilters(m, filters)).slice(0, max_results);

    const filterNote = [
      date_from ? `from ${date_from}` : undefined,
      date_to ? `to ${date_to}` : undefined,
      has_attachments === undefined ? undefined : has_attachments ? "with attachments" : "without attachments",
    ]
      .filter(Boolean)
      .join(", ");
    const suffix = filterNote ? ` (${filterNote})` : "";

    const structuredBase = {
      mode: query === undefined ? "latest" : "search",
      scope,
      ...(query !== undefined ? { query } : {}),
      ...(filtersActive
        ? {
            filters: {
              ...(date_from ? { date_from } : {}),
              ...(date_to ? { date_to } : {}),
              ...(has_attachments !== undefined ? { has_attachments } : {}),
            },
          }
        : {}),
    };

    if (messages.length === 0) {
      return structuredResult(
        query === undefined
          ? `No messages in ${scope}${suffix}.`
          : `No messages matching ${JSON.stringify(query)} in ${scope}${suffix}.`,
        { ...structuredBase, count: 0, messages: [] }
      );
    }

    const lines = messages.map((m, i) => {
      const preview = (m.bodyPreview ?? "").replace(/\s+/g, " ").trim().slice(0, 150);
      return [
        `${i + 1}. ${m.subject || "(no subject)"}`,
        `   From: ${formatSender(m.from)}  At: ${formatLocal(m.receivedDateTime)}${m.hasAttachments ? "  [has attachments]" : ""}`,
        `   Message id: ${m.id}`,
        `   Conversation id: ${m.conversationId}`,
        ...(include_body_preview && preview ? [`   Preview: ${preview}`] : []),
      ].join("\n");
    });
    const header =
      query === undefined
        ? `${messages.length} latest message(s) in ${scope}${suffix} (newest first):`
        : `${messages.length} result(s) for ${JSON.stringify(query)} in ${scope}${suffix} (relevance-ranked):`;

    return structuredResult(`${header}\n\n` + lines.join("\n\n"), {
      ...structuredBase,
      count: messages.length,
      messages: messages.map((m) => ({
        subject: m.subject || "(no subject)",
        from: formatSender(m.from),
        ...(m.receivedDateTime ? { receivedDateTime: m.receivedDateTime } : {}),
        receivedLocal: formatLocal(m.receivedDateTime),
        messageId: m.id,
        conversationId: m.conversationId,
        hasAttachments: Boolean(m.hasAttachments),
        ...(include_body_preview
          ? { preview: (m.bodyPreview ?? "").replace(/\s+/g, " ").trim().slice(0, 150) }
          : {}),
      })),
    });
  });
}
