import { z } from "zod";
import { callGraphServer } from "../graph.js";
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
    .enum(["list", "create", "update", "delete"])
    .describe(
      "list: all inbox rules; create: a new rule; update: change fields on an existing rule in place (rule_id); delete: remove a rule by rule_id."
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
};

const manageRulesArgs = z.object(manageRulesSchema);

export const manageRulesDescription =
  "List, create, update, or delete Outlook inbox rules. WARNING: a rule acts automatically on ALL FUTURE incoming mail with no per-message approval — before creating or updating one, state the complete resulting rule (every condition, exception, and action) to the user, and keep rules conservative (prefer narrow conditions, and use exceptions to carve out mail the rule must not touch). Rules can move, mark read, or soft-delete matching mail; forwarding actions are deliberately not supported. update patches the rule in place, keeping its id and position; conditions, exceptions, and actions are each replaced wholesale by what you pass. Existing messages are never affected, only future arrivals.";

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

export async function manageRulesHandler(
  input: z.input<typeof manageRulesArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, display_name, conditions, exceptions, actions, enabled, rule_id } =
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
