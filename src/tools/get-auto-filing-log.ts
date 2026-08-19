import { z } from "zod";
import { AUDIT_CAP, readAuditLog, type AuditEntry } from "../core/auto-filing.js";
import { requireStateStore } from "../core/state.js";
import { ToolResult, errorResult, formatLocal, runTool, textResult } from "./common.js";

export const getAutoFilingLogSchema = {
  limit: z
    .number()
    .int()
    .min(1)
    .max(AUDIT_CAP)
    .default(25)
    .describe(`Maximum number of entries to return (default 25, max ${AUDIT_CAP}).`),
  actions_only: z
    .boolean()
    .default(false)
    .describe(
      "Only show entries where something actually happened (a move, a categorize, a drafted brief). Default false, which also shows every decision NOT to act and why."
    ),
};

const getAutoFilingLogArgs = z.object(getAutoFilingLogSchema);

export const getAutoFilingLogDescription =
  "Read the audit trail of the LLM mail features: every message the auto-filer classified and where it moved it (each entry's `source` says whether a model decided or a learned preference filed it with no model call), every message it deliberately left alone and why (low confidence, a protected subject, a discarded model answer, the daily budget cap), every correction it learned from the user re-filing something, and every morning brief that was drafted. This is the record of what ran on the user's behalf — check it before trusting or enabling auto-filing. The last 100 decisions are kept, newest first. Available only on the hosted (remote) server, where the features actually run; the local stdio server returns an error.";

export async function getAutoFilingLogHandler(
  input: z.input<typeof getAutoFilingLogArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { limit, actions_only } = getAutoFilingLogArgs.parse(input);
    const store = requireStateStore();
    if (store.mode !== "remote") {
      return errorResult(
        "get_auto_filing_log works only on the hosted (remote) server. The auto-filer runs " +
          "there, off the change notifications Microsoft Graph pushes to it, and its audit log " +
          "lives in that server's KV store — the local stdio server has neither."
      );
    }

    const all = await readAuditLog(store);
    const entries = (actions_only ? all.filter((e) => e.action !== "none") : all).slice(0, limit);

    if (entries.length === 0) {
      return textResult(
        all.length === 0
          ? "The auto-filing log is empty: nothing has been classified yet. If you have just " +
              "enabled auto-filing, entries appear as new mail arrives."
          : `No entries with an action among the last ${all.length} decision(s). Call again with ` +
              "actions_only false to see what was left alone and why."
      );
    }

    const lines = entries.map((entry, index) => `${index + 1}. ${formatEntry(entry)}`);
    return textResult(
      `${entries.length} of ${all.length} logged decision(s), newest first:\n\n${lines.join("\n\n")}`
    );
  });
}

function formatEntry(entry: AuditEntry): string {
  const parts = [
    `${formatLocal(entry.at)}  [${entry.feature}]  ${entry.action.toUpperCase()}`,
    `   ${entry.subject ?? "(no subject recorded)"}`,
  ];
  const detail: string[] = [];
  if (entry.folder) detail.push(`folder: ${entry.folder}`);
  if (entry.categories?.length) detail.push(`categories: ${entry.categories.join(", ")}`);
  if (entry.confidence !== undefined) detail.push(`confidence: ${entry.confidence.toFixed(2)}`);
  // "preference" marks the no-model fast path; "llm" a genuine model decision.
  if (entry.source) detail.push(`source: ${entry.source}`);
  if (entry.sender) detail.push(`sender: ${entry.sender}`);
  if (detail.length) parts.push(`   ${detail.join("  ")}`);
  parts.push(`   Reason: ${entry.reason}`);
  if (entry.model) {
    const usage = entry.usage ? ` (${entry.usage.input} in / ${entry.usage.output} out tokens)` : "";
    parts.push(`   Model: ${entry.model}${usage}`);
  }
  if (entry.messageId) parts.push(`   Message id: ${entry.messageId}`);
  return parts.join("\n");
}
