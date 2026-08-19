import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { torontoDateOf } from "../core/auto-filing.js";
import {
  RULES_BACKUP_FORMAT,
  buildRulesBackup,
  diffRules,
  forwardingRuleNames,
  hasAnyCondition,
  normalizeForDiff,
  parseRulesBackup,
  type PortableRule,
  type RulesDiff,
} from "../core/rules-backup.js";
import { getStateStore } from "../core/state.js";
import { saveToDownloads } from "./save-local.js";
import {
  ToolResult,
  errorResult,
  isNotFound,
  runTool,
  textResult,
  toRecipients,
} from "./common.js";

// Inbox rules run server-side on ALL future incoming mail with no per-message
// approval, so creation is deliberately conservative. Forwarding/redirect rule
// actions are EXCLUDED by design: a standing silent forward is an exfiltration
// primitive and stays out of this server.

const RULES_PATH = "/me/mailFolders/inbox/messageRules";

/** The predicate shape shared by conditions and exceptions (Graph messageRulePredicates). */
function predicateShape(purpose: string) {
  return z
    .object({
      from_addresses: z
        .array(z.string().email())
        .optional()
        .describe(`${purpose} messages whose sender is one of these exact email addresses.`),
      sender_contains: z
        .array(z.string().min(1))
        .optional()
        .describe(`${purpose} messages whose sender name/address contains any of these substrings.`),
      subject_contains: z
        .array(z.string().min(1))
        .optional()
        .describe(`${purpose} messages whose subject contains any of these substrings.`),
      body_contains: z
        .array(z.string().min(1))
        .optional()
        .describe(`${purpose} messages whose body contains any of these substrings.`),
    })
    .describe(purpose);
}

type Predicates = z.infer<ReturnType<typeof predicateShape>>;

export const manageRulesSchema = {
  action: z
    .enum(["list", "create", "update", "delete", "export", "import"])
    .describe(
      "list: all inbox rules; create: a new rule; update: change fields on an existing rule in place (rule_id); delete: remove a rule by rule_id; export: the full rule set (conditions, exceptions, actions) as portable JSON — returned inline, and on the local server also saved to a dated file; import: restore rules from an exported backup (dry-run by default, see apply)."
    ),
  display_name: z
    .string()
    .min(1)
    .optional()
    .describe('Rule name (required for action "create"; optional rename on update).'),
  conditions: predicateShape("Match")
    .optional()
    .describe(
      'When the rule fires. Required for "create"; on "update" the whole condition set is REPLACED by what you pass (at least one condition is always required).'
    ),
  exceptions: predicateShape("Exempt")
    .optional()
    .describe(
      'Optional carve-outs: matching messages the rule must NOT act on, same fields as conditions. On "update" the whole exception set is REPLACED by what you pass; pass an empty object to clear all exceptions.'
    ),
  actions: z
    .object({
      move_to_folder: z
        .string()
        .min(1)
        .optional()
        .describe(
          'Move matching messages to this folder: a well-known name ("archive", "junkemail", …) or a folder id from list_folders or create_folder. The folder must exist.'
        ),
      mark_as_read: z.boolean().optional().describe("Mark matching messages as read."),
      delete: z
        .boolean()
        .optional()
        .describe("Move matching messages to Deleted Items (soft delete, recoverable)."),
    })
    .optional()
    .describe(
      'What the rule does. Required for "create"; on "update" the whole action set is REPLACED by what you pass (at least one action is always required). Forwarding actions are deliberately not supported.'
    ),
  enabled: z
    .boolean()
    .optional()
    .describe('For action "update": turn the rule on (true) or off (false) without deleting it.'),
  rule_id: z
    .string()
    .min(1)
    .optional()
    .describe(
      'The rule to change or remove (required for actions "update" and "delete"; from action "list").'
    ),
  backup_json: z
    .string()
    .min(2)
    .optional()
    .describe(
      `The backup document to import (required for action "import"): the ${RULES_BACKUP_FORMAT} JSON that action "export" produced, verbatim.`
    ),
  apply: z
    .boolean()
    .default(false)
    .describe(
      'For action "import": false (the default) is a DRY RUN that only shows the diff against the live rules; true applies the listed creates and updates. Import NEVER deletes — live rules absent from the backup are listed and left alone either way.'
    ),
};

const manageRulesArgs = z.object(manageRulesSchema);

export const manageRulesDescription =
  "List, create, update, delete, export, or import Outlook inbox rules. WARNING: a rule acts automatically on ALL FUTURE incoming mail with no per-message approval — before creating or updating one, state the complete resulting rule (every condition, exception, and action) to the user, and keep rules conservative (prefer narrow conditions, and use exceptions to carve out mail the rule must not touch). Rules can move, mark read, or soft-delete matching mail; forwarding actions are deliberately not supported, and importing a backup that carries them is refused. update patches the rule in place, keeping its id and position; conditions, exceptions, and actions are each replaced wholesale by what you pass. export returns the whole rule set as portable JSON (and saves a dated file on the local server); import is a dry-run diff against the live rules unless apply is true, and NEVER deletes — live rules missing from the backup are only listed. Existing messages are never affected, only future arrivals.";

/** Graph messageRulePredicates for a predicate input, or undefined when nothing was set. */
function buildPredicates(input: Predicates | undefined): any {
  if (!input) return undefined;
  const out: any = {};
  if (input.from_addresses?.length) out.fromAddresses = toRecipients(input.from_addresses);
  if (input.sender_contains?.length) out.senderContains = input.sender_contains;
  if (input.subject_contains?.length) out.subjectContains = input.subject_contains;
  if (input.body_contains?.length) out.bodyContains = input.body_contains;
  return out;
}

function hasPredicate(input: Predicates | undefined): boolean {
  const built = buildPredicates(input);
  return Boolean(built && Object.keys(built).length > 0);
}

/** Human-readable AND-joined summary of a Graph messageRulePredicates object. */
function summarizePredicates(p: any): string[] {
  const parts: string[] = [];
  if (p?.fromAddresses?.length) {
    parts.push(
      `from ${p.fromAddresses.map((r: any) => r.emailAddress?.address).filter(Boolean).join(" or ")}`
    );
  }
  if (p?.senderContains?.length) parts.push(`sender contains ${p.senderContains.join(" or ")}`);
  if (p?.subjectContains?.length) parts.push(`subject contains ${p.subjectContains.join(" or ")}`);
  if (p?.bodyContains?.length) parts.push(`body contains ${p.bodyContains.join(" or ")}`);
  const known = new Set(["fromAddresses", "senderContains", "subjectContains", "bodyContains"]);
  for (const key of Object.keys(p ?? {})) {
    if (!known.has(key)) parts.push(`${key}: ${JSON.stringify(p[key])}`);
  }
  return parts;
}

/** Human-readable "conditions → actions EXCEPT exceptions" summary of a Graph messageRule. */
function summarizeRule(rule: any, folderNames: Map<string, string>): string {
  const conditions = summarizePredicates(rule.conditions);
  const exceptions = summarizePredicates(rule.exceptions);

  const actions: string[] = [];
  const a = rule.actions ?? {};
  if (a.moveToFolder) actions.push(`move to ${folderNames.get(a.moveToFolder) ?? a.moveToFolder}`);
  if (a.markAsRead) actions.push("mark as read");
  if (a.delete) actions.push("delete (to Deleted Items)");
  if (a.forwardTo?.length || a.forwardAsAttachmentTo?.length || a.redirectTo?.length) {
    actions.push("FORWARD/REDIRECT (created outside this server)");
  }
  const knownActions = new Set([
    "moveToFolder",
    "markAsRead",
    "delete",
    "forwardTo",
    "forwardAsAttachmentTo",
    "redirectTo",
    "stopProcessingRules",
  ]);
  for (const key of Object.keys(a)) {
    if (!knownActions.has(key) && a[key]) actions.push(`${key}: ${JSON.stringify(a[key])}`);
  }

  const head = `${conditions.join(" AND ") || "(no conditions — matches everything)"} → ${actions.join(", ") || "(no actions)"}`;
  return exceptions.length ? `${head}\n   EXCEPT ${exceptions.join(" AND ")}` : head;
}

/** Resolve the display names of any moveToFolder targets, for readable summaries. */
async function resolveFolderNames(rules: any[]): Promise<Map<string, string>> {
  const names = new Map<string, string>();
  const ids = new Set<string>(
    rules.map((r) => r.actions?.moveToFolder).filter((id): id is string => Boolean(id))
  );
  for (const id of ids) {
    try {
      const folder = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(id)}?$select=displayName`
      );
      if (folder?.displayName) names.set(id, `"${folder.displayName}"`);
    } catch {
      // keep the raw id in the summary
    }
  }
  return names;
}

/**
 * Translate an actions input into Graph rule actions, resolving the move target.
 * Returns an error result instead when the target folder does not exist.
 */
async function buildActions(
  actions: NonNullable<z.infer<typeof manageRulesArgs>["actions"]>
): Promise<{ graphActions: any; moveTargetLabel: string } | { error: ToolResult }> {
  const graphActions: any = {};
  let moveTargetLabel = "";
  if (actions.move_to_folder) {
    let folder: any;
    try {
      folder = await callGraphServer(
        `/me/mailFolders/${encodeURIComponent(actions.move_to_folder)}?$select=id,displayName`
      );
    } catch (err) {
      if (isNotFound(err)) {
        return {
          error: errorResult(
            `move_to_folder target ${JSON.stringify(actions.move_to_folder)} does not exist — use a well-known name, a folder id from list_folders, or create_folder first.`
          ),
        };
      }
      throw err;
    }
    graphActions.moveToFolder = folder.id;
    moveTargetLabel = folder.displayName ?? actions.move_to_folder;
  }
  if (actions.mark_as_read) graphActions.markAsRead = true;
  if (actions.delete) graphActions.delete = true;
  return { graphActions, moveTargetLabel };
}

/**
 * Refuse an import whose creates/updates would leave a rule with no conditions
 * (it would act on ALL mail) or a create with no actions — the same guards the
 * create and update actions enforce. Returns null when the plan is acceptable.
 */
function guardImport(diff: RulesDiff): ToolResult | null {
  const conditionless = [
    ...diff.creates.filter((rule) => !hasAnyCondition(rule.conditions)),
    ...diff.updates
      .filter(
        (update) =>
          update.fields.some((f) => f.field === "conditions") &&
          !hasAnyCondition(update.backup.conditions)
      )
      .map((update) => update.backup),
  ];
  if (conditionless.length > 0) {
    return errorResult(
      `Import refused: rule(s) ${conditionless.map((r) => JSON.stringify(r.displayName)).join(", ")} ` +
        "would end up with no conditions and act on ALL incoming mail. Fix the backup JSON first."
    );
  }
  const actionless = diff.creates.filter(
    (rule) => !rule.actions || Object.keys(rule.actions).length === 0
  );
  if (actionless.length > 0) {
    return errorResult(
      `Import refused: rule(s) ${actionless.map((r) => JSON.stringify(r.displayName)).join(", ")} ` +
        "have no actions. Fix the backup JSON first."
    );
  }
  return null;
}

/** Human-readable diff: creates with their full rule, field-level updates, and the never-deleted rest. */
function renderRulesDiff(diff: RulesDiff): string {
  const none = new Map<string, string>();
  const sections: string[] = [];

  sections.push(
    diff.creates.length > 0
      ? `Would CREATE ${diff.creates.length} rule(s):\n` +
          diff.creates
            .map((rule) => `  + "${rule.displayName}": ${summarizeRule(rule, none)}`)
            .join("\n")
      : "Nothing to create."
  );

  sections.push(
    diff.updates.length > 0
      ? `Would UPDATE ${diff.updates.length} rule(s) in place:\n` +
          diff.updates
            .map(
              (update) =>
                `  ~ "${update.backup.displayName}" (rule id ${update.live.id}):\n` +
                update.fields
                  .map(
                    (f) =>
                      `      ${f.field}: live ${JSON.stringify(normalizeForDiff(f.live) ?? null)} ` +
                      `→ backup ${JSON.stringify(normalizeForDiff(f.backup) ?? null)}`
                  )
                  .join("\n")
            )
            .join("\n")
      : "Nothing to update."
  );

  if (diff.unchanged.length > 0) {
    sections.push(
      `Already identical: ${diff.unchanged.map((rule) => `"${rule.displayName}"`).join(", ")}`
    );
  }

  sections.push(
    diff.liveOnly.length > 0
      ? `Live but NOT in the backup (import never deletes — these stay as they are):\n` +
          diff.liveOnly
            .map((rule: any) => `  = "${rule.displayName}" (rule id ${rule.id})`)
            .join("\n")
      : "No live rules are missing from the backup."
  );

  return sections.join("\n\n");
}

export async function manageRulesHandler(
  input: z.input<typeof manageRulesArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, display_name, conditions, exceptions, actions, enabled, rule_id, backup_json, apply } =
      manageRulesArgs.parse(input);

    if (action === "list") {
      const data = await callGraphServer(RULES_PATH);
      const rules: any[] = data?.value ?? [];
      if (rules.length === 0) return textResult("No inbox rules.");
      const folderNames = await resolveFolderNames(rules);
      const lines = rules.map(
        (r, i) =>
          `${i + 1}. ${r.displayName || "(unnamed)"}${r.isEnabled ? "" : "  [DISABLED]"}\n` +
          `   Rule id: ${r.id}\n` +
          `   ${summarizeRule(r, folderNames)}`
      );
      return textResult(`${rules.length} inbox rule(s):\n\n${lines.join("\n\n")}`);
    }

    if (action === "export") {
      const data = await callGraphServer(RULES_PATH);
      const rules: any[] = data?.value ?? [];
      const backup = buildRulesBackup(rules, new Date());
      const json = JSON.stringify(backup, null, 2);

      const lines = [
        `Exported ${rules.length} inbox rule(s) as ${RULES_BACKUP_FORMAT} ` +
          "(complete: conditions, exceptions, and actions).",
      ];
      const forwarding = forwardingRuleNames(backup.rules);
      if (forwarding.length > 0) {
        lines.push(
          `WARNING: ${forwarding.length} rule(s) carry forward/redirect actions created outside ` +
            `this server (${forwarding.join(", ")}). The backup records them faithfully, but ` +
            "importing it will be refused until those entries are removed from the JSON."
        );
      }
      // The local stdio server also writes a dated file, like export_message
      // does; the hosted Worker has no filesystem, so there the inline JSON is
      // the artifact.
      if (getStateStore()?.mode !== "remote") {
        const savePath = await saveToDownloads(
          `inbox-rules-${torontoDateOf(new Date())}.json`,
          Buffer.from(json, "utf8")
        );
        lines.push(`Saved to: ${savePath}`);
      } else {
        lines.push("(Hosted server: no file was written — keep the JSON below.)");
      }
      lines.push("", "Backup JSON:", json);
      return textResult(lines.join("\n"));
    }

    if (action === "import") {
      if (!backup_json) {
        return errorResult(
          'Action "import" requires backup_json — the JSON that action "export" produced.'
        );
      }
      const parsed = parseRulesBackup(backup_json);
      if (!parsed.ok) return errorResult(`That backup cannot be imported: ${parsed.error}`);

      // The no-forwarding discipline holds on the way back in too: this server
      // never creates forwarding rules, so it will not restore them either.
      const forwarding = forwardingRuleNames(parsed.backup.rules);
      if (forwarding.length > 0) {
        return errorResult(
          `Import refused: rule(s) ${forwarding.map((n) => JSON.stringify(n)).join(", ")} carry ` +
            "forward/redirect actions, which this server deliberately never creates (a standing " +
            "silent forward is an exfiltration primitive). Remove those entries from the backup " +
            "JSON to import the rest."
        );
      }

      const data = await callGraphServer(RULES_PATH);
      const liveRules: any[] = data?.value ?? [];
      const diff = diffRules(liveRules, parsed.backup.rules);

      // The create/update guards hold for imported rules as well: nothing may
      // end up as a rule with no conditions (it would act on ALL mail) or with
      // no actions.
      const guarded = guardImport(diff);
      if (guarded) return guarded;

      const summary = renderRulesDiff(diff);
      const toApply = diff.creates.length + diff.updates.length;

      if (!apply) {
        return textResult(
          `DRY RUN — nothing was changed. Backup of ${parsed.backup.exportedAt} ` +
            `(${parsed.backup.rules.length} rule(s)) against ${liveRules.length} live rule(s):\n\n` +
            summary +
            "\n\n" +
            (toApply > 0
              ? `Pass apply: true to apply the ${diff.creates.length} create(s) and ${diff.updates.length} update(s) above.`
              : "Nothing to apply — the backup creates or changes no rule.")
        );
      }

      if (toApply === 0) {
        return textResult(
          `Nothing to apply — the backup creates or changes no rule.\n\n${summary}`
        );
      }

      const maxSequence = Math.max(
        0,
        ...liveRules.map((rule: any) => Number(rule.sequence) || 0)
      );
      const results: string[] = [];
      let failures = 0;

      for (const [index, rule] of diff.creates.entries()) {
        try {
          const created = await callGraphServer(RULES_PATH, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              displayName: rule.displayName,
              sequence: rule.sequence ?? maxSequence + index + 1,
              isEnabled: rule.isEnabled ?? true,
              conditions: rule.conditions,
              ...(rule.exceptions !== undefined ? { exceptions: rule.exceptions } : {}),
              actions: rule.actions,
            }),
          });
          results.push(`OK      created "${rule.displayName}" (rule id ${created?.id})`);
        } catch (err) {
          failures++;
          results.push(
            `FAILED  create "${rule.displayName}": ${err instanceof Error ? err.message.split("\n")[0] : String(err)}`
          );
        }
      }

      for (const update of diff.updates) {
        const patch: Record<string, unknown> = {};
        for (const field of update.fields) {
          // A field the backup does not carry is only ever cleared for
          // exceptions (an empty exception set is safe; see the update action).
          patch[field.field] = field.backup === undefined ? {} : field.backup;
        }
        try {
          await callGraphServer(`${RULES_PATH}/${encodeURIComponent(String(update.live.id))}`, {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(patch),
          });
          results.push(
            `OK      updated "${update.backup.displayName}" in place ` +
              `(${update.fields.map((f) => f.field).join(", ")})`
          );
        } catch (err) {
          failures++;
          results.push(
            `FAILED  update "${update.backup.displayName}": ${err instanceof Error ? err.message.split("\n")[0] : String(err)}`
          );
        }
      }

      const applied =
        `Import applied: ${toApply - failures} of ${toApply} change(s) succeeded.\n\n` +
        results.join("\n") +
        (diff.liveOnly.length > 0
          ? `\n\nLeft untouched (live but not in the backup — import never deletes): ` +
            diff.liveOnly.map((rule: any) => `"${rule.displayName}"`).join(", ")
          : "");
      return failures === toApply && toApply > 0 ? errorResult(applied) : textResult(applied);
    }

    if (action === "delete") {
      if (!rule_id) return errorResult('Action "delete" requires rule_id (from action "list").');
      await callGraphServer(`${RULES_PATH}/${encodeURIComponent(rule_id)}`, { method: "DELETE" });
      return textResult(`Rule ${rule_id} deleted. It no longer acts on incoming mail.`);
    }

    // create and update share the condition/action validation and folder resolution.
    if (conditions && !hasPredicate(conditions)) {
      return errorResult(
        "A rule needs at least one condition (from_addresses, sender_contains, subject_contains, or body_contains) — a rule with no conditions would act on ALL incoming mail."
      );
    }
    const hasAnyAction =
      actions && Boolean(actions.move_to_folder || actions.mark_as_read || actions.delete);
    if (actions && !hasAnyAction) {
      return errorResult(
        "A rule needs at least one action (move_to_folder, mark_as_read, or delete)."
      );
    }

    let graphActions: any | undefined;
    let moveTargetLabel = "";
    if (actions) {
      const built = await buildActions(actions);
      if ("error" in built) return built.error;
      graphActions = built.graphActions;
      moveTargetLabel = built.moveTargetLabel;
    }

    if (action === "update") {
      if (!rule_id) return errorResult('Action "update" requires rule_id (from action "list").');
      const patch: any = {};
      if (display_name !== undefined) patch.displayName = display_name;
      if (conditions !== undefined) patch.conditions = buildPredicates(conditions);
      if (exceptions !== undefined) patch.exceptions = buildPredicates(exceptions);
      if (graphActions !== undefined) patch.actions = graphActions;
      if (enabled !== undefined) patch.isEnabled = enabled;
      if (Object.keys(patch).length === 0) {
        return errorResult(
          'Action "update" needs at least one of display_name, conditions, exceptions, actions, or enabled.'
        );
      }

      let updated: any;
      try {
        updated = await callGraphServer(`${RULES_PATH}/${encodeURIComponent(rule_id)}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(patch),
        });
      } catch (err) {
        if (isNotFound(err)) {
          return errorResult(
            `No inbox rule with id ${rule_id} — use action "list" to see the current rules.`
          );
        }
        throw err;
      }

      const folderNames = await resolveFolderNames([updated]);
      if (graphActions?.moveToFolder && moveTargetLabel) {
        folderNames.set(graphActions.moveToFolder, `"${moveTargetLabel}"`);
      }
      return textResult(
        `Inbox rule updated in place (same rule id, same position).\n` +
          `Name: ${updated.displayName}${updated.isEnabled ? "" : "  [DISABLED]"}\n` +
          `Rule id: ${updated.id}\n` +
          `Changed: ${Object.keys(patch).join(", ")}\n` +
          `Rule now: ${summarizeRule(updated, folderNames)}\n` +
          "It acts automatically on all future incoming mail that matches (existing mail is unaffected)."
      );
    }

    // action === "create"
    if (!display_name) return errorResult('Action "create" requires display_name.');
    if (!conditions) {
      return errorResult(
        "A rule needs at least one condition (from_addresses, sender_contains, subject_contains, or body_contains) — a rule with no conditions would act on ALL incoming mail."
      );
    }
    if (!graphActions) {
      return errorResult(
        "A rule needs at least one action (move_to_folder, mark_as_read, or delete)."
      );
    }

    // Existing rules serve two purposes: an early loud failure if rules are not
    // accessible, and the auto-assigned sequence number.
    const existing = await callGraphServer(RULES_PATH);
    const sequence =
      Math.max(0, ...(existing?.value ?? []).map((r: any) => Number(r.sequence) || 0)) + 1;

    const created = await callGraphServer(RULES_PATH, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        displayName: display_name,
        sequence,
        isEnabled: enabled ?? true,
        conditions: buildPredicates(conditions),
        ...(exceptions !== undefined ? { exceptions: buildPredicates(exceptions) } : {}),
        actions: graphActions,
      }),
    });

    const folderNames = new Map<string, string>();
    if (graphActions.moveToFolder) {
      folderNames.set(graphActions.moveToFolder, `"${moveTargetLabel}"`);
    }
    return textResult(
      `Inbox rule created.\n` +
        `Name: ${created.displayName}\n` +
        `Rule id: ${created.id}\n` +
        `Rule: ${summarizeRule(created, folderNames)}\n` +
        "It now acts automatically on all future incoming mail that matches (existing mail is unaffected)."
    );
  });
}
