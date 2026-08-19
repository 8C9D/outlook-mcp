import { z } from "zod";
import {
  DRIVE_ITEM_SELECT,
  describeDriveItem,
  itemDisplayPath,
  matchesTypeFilter,
} from "../core/drive.js";
import { ToolResult, fetchPaged, formatLocal, runTool, structuredResult } from "./common.js";
import { formatSize } from "./read-message.js";

export const searchFilesSchema = {
  query: z
    .string()
    .min(1)
    .describe(
      "What to search for. Matches file and folder names and, for many file types, file content. A distinctive single word works best."
    ),
  type: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Optional filter: "file" (only files), "folder" (only folders), or a filename extension like "pdf" or ".docx".'
    ),
  max_results: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(20)
    .describe("Most results to return (default 20, max 50)."),
};

const searchFilesArgs = z.object(searchFilesSchema);

/** Permissive machine-readable search results; every field optional. */
export const searchFilesOutputSchema = {
  query: z.string().optional(),
  count: z.number().optional().describe("Results returned (after the type filter)."),
  files: z
    .array(
      z.looseObject({
        id: z.string().optional(),
        name: z.string().optional(),
        path: z.string().optional(),
        kind: z.string().optional(),
        size: z.number().optional(),
        mimeType: z.string().optional(),
        modified: z.string().optional(),
        webUrl: z.string().optional(),
      })
    )
    .optional(),
};

export const searchFilesDescription =
  "Search OneDrive for files and folders by name or content, optionally filtered to files, folders, or one extension. Results come from OneDrive's search index, which lags reality: a file created or renamed in the last few minutes may not appear yet (use list_folder for an exact, current listing of one folder). Read a result with read_file; ids work with every file tool.";

/** How many raw hits to page through before the type filter is applied. */
const SEARCH_FETCH_CAP = 200;

export async function searchFilesHandler(
  input: z.input<typeof searchFilesArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { query, type, max_results } = searchFilesArgs.parse(input);

    // The q value is embedded inside quotes in the URL path; encode it fully
    // and escape single quotes the OData way so a quote cannot break the call.
    const q = encodeURIComponent(query.replace(/'/g, "''"));
    const hits = await fetchPaged(
      `/me/drive/root/search(q='${q}')?$select=${DRIVE_ITEM_SELECT}&$top=50`,
      SEARCH_FETCH_CAP
    );

    const matched = hits.filter((item) => matchesTypeFilter(item, type)).slice(0, max_results);
    const files = matched.map(describeDriveItem);

    const filterNote = type ? ` (type: ${type})` : "";
    if (matched.length === 0) {
      return structuredResult(
        `No OneDrive items match ${JSON.stringify(query)}${filterNote}. ` +
          "Note the search index can lag a few minutes behind new or renamed files — " +
          "list_folder sees them immediately.",
        { query, count: 0, files: [] }
      );
    }

    const lines = matched.map((item) => {
      const isFolder = item.folder !== undefined;
      const detail = isFolder
        ? `folder, ${item.folder?.childCount ?? 0} item(s)`
        : `${formatSize(item.size)}${item.file?.mimeType ? `, ${item.file.mimeType}` : ""}`;
      return (
        `${itemDisplayPath(item)} — ${detail}\n` +
        `  modified: ${formatLocal(item.lastModifiedDateTime)}\n` +
        `  id: ${item.id}`
      );
    });
    return structuredResult(
      `OneDrive search ${JSON.stringify(query)}${filterNote}: ${matched.length} result(s)\n\n` +
        lines.join("\n"),
      { query, count: files.length, files }
    );
  });
}
