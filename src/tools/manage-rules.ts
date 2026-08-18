import { z } from "zod";
import { GraphError, callGraphServer } from "../graph.js";
import { ToolResult, errorResult, runTool, textResult, toRecipients } from "./common.js";

// Inbox rules run server-side on ALL future incoming mail with no per-message
// approval, so creation is deliberately conservative. Forwarding/redirect rule
// actions are EXCLUDED by design: a standing silent forward is an exfiltration
// primitive and stays out of this server.

const RULES_PATH = "/me/mailFolders/inbox/messageRules";

export const manageRulesSchema = {
  action: z
    .enum(["list", "create", "delete"])
    .describe("list: all inbox rules; create: a new rule; delete: remove a rule by rule_id."),
  display_name: z
    .string()
    .min(1)
    .optional()
    .describe('Rule name (required for action "create").'),
  conditions: z
    .object({
      from_addresses: z
        .array(z.string().email())
        .optional()
        .describe("Match messages whose sender is one of these exact email addresses."),
      sender_contains: z
        .array(z.string().min(1))
        .optional()
        .describe("Match messages whose sender name/address contains any of these substrings."),
      subject_contains: z
        .array(z.string().min(1))
        .optional()
        .describe("Match messages whose subject contains any of these substrings."),
      body_contains: z
        .array(z.string().min(1))
        .optional()
        .describe("Match messages whose body contains any of these substrings."),
    })
    .optional()
    .describe('For action "create": when the rule fires. At least one condition is required.'),
  actions: z
    .object({
      move_to_folder: z
        .string()
        .min(1)
        .optional()
        .describe(
          'Move matching messages to this folder: a well-known name ("archive", "junkemail", …) or a folder id from list_folders. The folder must exist.'
        ),
      mark_as_read: z.boolean().optional().describe("Mark matching messages as read."),
      delete: z
        .boolean()
        .optional()
        .describe("Move matching messages to Deleted Items (soft delete, recoverable)."),
    })
    .optional()
    .describe(
      'For action "create": what the rule does. At least one action is required. Forwarding actions are deliberately not supported.'
    ),
  rule_id: z
    .string()
    .min(1)
    .optional()
    .describe('The id of the rule to remove (required for action "delete"; from action "list").'),
};

const manageRulesArgs = z.object(manageRulesSchema);

export const manageRulesDescription =
  "List, create, or delete Outlook inbox rules. WARNING: a rule acts automatically on ALL FUTURE incoming mail with no per-message approval — before creating one, state the complete rule (every condition and every action) to the user, and keep rules conservative (prefer narrow conditions). Rules can move, mark read, or soft-delete matching mail; forwarding actions are deliberately not supported. Existing messages are not affected, only future arrivals.";

/** Human-readable "conditions → actions" summary of a Graph messageRule. */
function summarizeRule(rule: any, folderNames: Map<string, string>): string {
  const conditions: string[] = [];
  const c = rule.conditions ?? {};
  if (c.fromAddresses?.length) {
    conditions.push(
      `from ${c.fromAddresses.map((r: any) => r.emailAddress?.address).filter(Boolean).join(" or ")}`
    );
  }
  if (c.senderContains?.length) conditions.push(`sender contains ${c.senderContains.join(" or ")}`);
  if (c.subjectContains?.length)
    conditions.push(`subject contains ${c.subjectContains.join(" or ")}`);
  if (c.bodyContains?.length) conditions.push(`body contains ${c.bodyContains.join(" or ")}`);
  const knownConditions = new Set([
    "fromAddresses",
    "senderContains",
    "subjectContains",
    "bodyContains",
  ]);
  for (const key of Object.keys(c)) {
    if (!knownConditions.has(key)) conditions.push(`${key}: ${JSON.stringify(c[key])}`);
  }

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

  return `${conditions.join(" AND ") || "(no conditions — matches everything)"} → ${actions.join(", ") || "(no actions)"}`;
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

export async function manageRulesHandler(
  input: z.input<typeof manageRulesArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, display_name, conditions, actions, rule_id } = manageRulesArgs.parse(input);

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

    // action === "create"
    if (!display_name) return errorResult('Action "create" requires display_name.');
    const hasCondition =
      conditions &&
      Boolean(
        conditions.from_addresses?.length ||
          conditions.sender_contains?.length ||
          conditions.subject_contains?.length ||
          conditions.body_contains?.length
      );
    if (!hasCondition) {
      return errorResult(
        "A rule needs at least one condition (from_addresses, sender_contains, subject_contains, or body_contains) — a rule with no conditions would act on ALL incoming mail."
      );
    }
    const hasAction =
      actions && Boolean(actions.move_to_folder || actions.mark_as_read || actions.delete);
    if (!hasAction) {
      return errorResult(
        "A rule needs at least one action (move_to_folder, mark_as_read, or delete)."
      );
    }

    // Existing rules serve two purposes: an early loud failure if rules are not
    // accessible, and the auto-assigned sequence number.
    const existing = await callGraphServer(RULES_PATH);
    const sequence =
      Math.max(0, ...(existing?.value ?? []).map((r: any) => Number(r.sequence) || 0)) + 1;

    const graphActions: any = {};
    let moveTargetLabel = "";
    if (actions!.move_to_folder) {
      let folder: any;
      try {
        folder = await callGraphServer(
          `/me/mailFolders/${encodeURIComponent(actions!.move_to_folder)}?$select=id,displayName`
        );
      } catch (err) {
        if (err instanceof GraphError && err.status === 404) {
          return errorResult(
            `move_to_folder target ${JSON.stringify(actions!.move_to_folder)} does not exist — use a well-known name or a folder id from list_folders.`
          );
        }
        throw err;
      }
      graphActions.moveToFolder = folder.id;
      moveTargetLabel = folder.displayName ?? actions!.move_to_folder;
    }
    if (actions!.mark_as_read) graphActions.markAsRead = true;
    if (actions!.delete) graphActions.delete = true;

    const graphConditions: any = {};
    if (conditions!.from_addresses?.length)
      graphConditions.fromAddresses = toRecipients(conditions!.from_addresses);
    if (conditions!.sender_contains?.length)
      graphConditions.senderContains = conditions!.sender_contains;
    if (conditions!.subject_contains?.length)
      graphConditions.subjectContains = conditions!.subject_contains;
    if (conditions!.body_contains?.length) graphConditions.bodyContains = conditions!.body_contains;

    const created = await callGraphServer(RULES_PATH, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        displayName: display_name,
        sequence,
        isEnabled: true,
        conditions: graphConditions,
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
