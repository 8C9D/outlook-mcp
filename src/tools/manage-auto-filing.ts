import { z } from "zod";
import { LLM_MODEL } from "../core/anthropic.js";
import {
  DAILY_CAP_MAX,
  PROTECTED_SUBJECT_PATTERNS,
  THRESHOLD_MAX,
  THRESHOLD_MIN,
  readApiCallsToday,
  readAuditLog,
  readLastDigestDate,
  readLlmConfig,
  torontoDateOf,
  writeLlmConfig,
  type LlmConfig,
} from "../core/auto-filing.js";
import { requireStateStore, type StateStore } from "../core/state.js";
import { ToolInputError, ToolResult, errorResult, runTool, textResult } from "./common.js";

export const manageAutoFilingSchema = {
  action: z
    .enum([
      "status",
      "enable_filing",
      "disable_filing",
      "enable_digest",
      "disable_digest",
      "set_threshold",
      "set_daily_cap",
      "add_skip_pattern",
      "remove_skip_pattern",
    ])
    .default("status")
    .describe(
      "status: report what is on, the tunables and today's API usage (default). enable_filing/disable_filing: LLM classification of newly-arrived mail into existing folders. enable_digest/disable_digest: the 07:00 America/Toronto morning brief, left as a draft. set_threshold, set_daily_cap, add_skip_pattern, remove_skip_pattern: the tunables."
    ),
  threshold: z
    .number()
    .min(THRESHOLD_MIN)
    .max(THRESHOLD_MAX)
    .optional()
    .describe(
      `Minimum confidence (${THRESHOLD_MIN}–${THRESHOLD_MAX}) for the classifier to move or categorize a message; below it nothing happens. Required for action "set_threshold". Default 0.8.`
    ),
  daily_cap: z
    .number()
    .int()
    .min(0)
    .max(DAILY_CAP_MAX)
    .optional()
    .describe(
      `Maximum Anthropic API calls per America/Toronto day, across both features; once reached everything is skipped and logged until midnight. Required for action "set_daily_cap". Default 200, max ${DAILY_CAP_MAX}. Set 0 to stop all API calls without changing the enable flags.`
    ),
  pattern: z
    .string()
    .min(2)
    .max(120)
    .optional()
    .describe(
      'Subject substring (matched case-insensitively) that must never be classified — mail matching it is never sent to the model and never moved. Required for actions "add_skip_pattern" and "remove_skip_pattern". The built-in list (one-time passcodes, verification codes, password resets and so on) always applies and cannot be removed.'
    ),
};

const manageAutoFilingArgs = z.object(manageAutoFilingSchema);

export const manageAutoFilingDescription =
  "Turn the two LLM mail features on or off and tune them: auto-filing (when new mail arrives, a model classifies it against the mailbox's EXISTING folders and categories and files it, never sending, deleting or replying) and the morning digest (a brief of overnight mail, the day's calendar and tasks due soon, left as an unsent DRAFT at 07:00 America/Toronto). BOTH SHIP DISABLED and cost money per message classified — tell the user what the README's cost section says before enabling either. Use get_auto_filing_log to see what the classifier actually did. Available only on the hosted (remote) server, where both features run; the local stdio server returns an error.";

export async function manageAutoFilingHandler(
  input: z.input<typeof manageAutoFilingArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, threshold, daily_cap, pattern } = manageAutoFilingArgs.parse(input);
    const store = requireStateStore();
    if (store.mode !== "remote") {
      return errorResult(
        "manage_auto_filing works only on the hosted (remote) server. Auto-filing runs off the " +
          "change notifications Microsoft Graph pushes there and the digest off its cron " +
          "trigger; its settings live in that server's KV store, so changing them here would " +
          "have no effect on anything."
      );
    }

    switch (action) {
      case "status":
        return textResult(await describeStatus(store, await readLlmConfig(store)));

      case "enable_filing":
        return applied(store, await writeLlmConfig(store, { filingEnabled: true }),
          "Auto-filing is now ON. New mail arriving in the inbox will be classified against your " +
            "existing folders and moved when the model is confident. Check get_auto_filing_log " +
            "after a few messages; disable_filing turns it off again immediately.");

      case "disable_filing":
        return applied(store, await writeLlmConfig(store, { filingEnabled: false }),
          "Auto-filing is now OFF. Nothing will be classified or moved. Mail already filed stays " +
            "where it is; get_auto_filing_log still shows what was done.");

      case "enable_digest":
        return applied(store, await writeLlmConfig(store, { digestEnabled: true }),
          "The morning digest is now ON. At 07:00 America/Toronto a brief of overnight unread " +
            'mail, today\'s calendar and tasks due within three days is left as a DRAFT titled ' +
            '"Morning brief — <date>". It is never sent; you send it, or just read it and delete it.');

      case "disable_digest":
        return applied(store, await writeLlmConfig(store, { digestEnabled: false }),
          "The morning digest is now OFF. No brief will be drafted. Any drafts already made stay " +
            "in Drafts.");

      case "set_threshold": {
        if (threshold === undefined) {
          throw new ToolInputError('Action "set_threshold" needs a threshold between 0.5 and 1.');
        }
        const config = await writeLlmConfig(store, { threshold });
        return applied(store, config,
          `Confidence threshold is now ${config.threshold.toFixed(2)}. Below it the classifier ` +
            "logs its reasoning and does nothing.");
      }

      case "set_daily_cap": {
        if (daily_cap === undefined) {
          throw new ToolInputError(`Action "set_daily_cap" needs a daily_cap between 0 and ${DAILY_CAP_MAX}.`);
        }
        const config = await writeLlmConfig(store, { dailyCallCap: daily_cap });
        return applied(store, config,
          `Daily Anthropic call cap is now ${config.dailyCallCap}. Once reached, both features ` +
            "log a skip and do nothing until midnight America/Toronto.");
      }

      case "add_skip_pattern": {
        if (!pattern) throw new ToolInputError('Action "add_skip_pattern" needs a pattern.');
        const existing = await readLlmConfig(store);
        const config = await writeLlmConfig(store, {
          skipPatterns: [...existing.skipPatterns, pattern],
        });
        return applied(store, config,
          `Mail whose subject contains "${pattern.trim().toLowerCase()}" will never be sent to the ` +
            "model or moved.");
      }

      case "remove_skip_pattern": {
        if (!pattern) throw new ToolInputError('Action "remove_skip_pattern" needs a pattern.');
        const needle = pattern.trim().toLowerCase();
        const existing = await readLlmConfig(store);
        if (!existing.skipPatterns.includes(needle)) {
          throw new ToolInputError(
            `"${needle}" is not one of your added skip patterns. The built-in patterns ` +
              `(${PROTECTED_SUBJECT_PATTERNS.join(", ")}) always apply and cannot be removed.`
          );
        }
        const config = await writeLlmConfig(store, {
          skipPatterns: existing.skipPatterns.filter((p) => p !== needle),
        });
        return applied(store, config, `Removed the skip pattern "${needle}".`);
      }
    }
  });
}

/** Report the change and the resulting state in one breath. */
async function applied(store: StateStore, config: LlmConfig, note: string): Promise<ToolResult> {
  return textResult(`${note}\n\n${await describeStatus(store, config)}`);
}

async function describeStatus(store: StateStore, config: LlmConfig): Promise<string> {
  const today = torontoDateOf(new Date());
  const used = await readApiCallsToday(store, today);
  const lastDigest = await readLastDigestDate(store);
  const logged = (await readAuditLog(store)).length;

  return [
    "LLM mail features:",
    `  Auto-filing:     ${config.filingEnabled ? "ON" : "OFF"}`,
    `  Morning digest:  ${config.digestEnabled ? "ON" : "OFF"}  (07:00 America/Toronto, drafted, never sent)`,
    "",
    "Settings:",
    `  Confidence threshold: ${config.threshold.toFixed(2)}`,
    `  Daily API call cap:   ${config.dailyCallCap}`,
    `  Model:                ${LLM_MODEL}`,
    `  Extra skip patterns:  ${config.skipPatterns.length ? config.skipPatterns.join(", ") : "(none)"}`,
    `  Built-in skip list:   ${PROTECTED_SUBJECT_PATTERNS.join(", ")}`,
    "",
    "Usage:",
    `  API calls today (${today}): ${used} of ${config.dailyCallCap}`,
    `  Last morning brief drafted: ${lastDigest ?? "(never)"}`,
    `  Decisions in the audit log: ${logged} (see get_auto_filing_log)`,
  ].join("\n");
}
