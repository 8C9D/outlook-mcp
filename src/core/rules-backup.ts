// Portable backup format for inbox rules, and the pure diff logic behind
// `manage_rules export` / `manage_rules import`. Nothing here touches Graph —
// the tool feeds Graph's rule objects in and applies the plan out — so every
// branch is unit-testable offline.
//
// Two invariants the tool relies on:
//  - Import NEVER deletes: live rules absent from a backup are only ever
//    listed (diffRules puts them in `liveOnly`).
//  - Forwarding stays out: a backup carrying forward/redirect actions is
//    detected here (forwardingRuleNames) and refused by the tool, consistent
//    with manage_rules never offering those actions in the first place.

/** The format marker written into every export and required on import. */
export const RULES_BACKUP_FORMAT = "outlook-mcp-rules/1";

/** One rule as stored in a backup: Graph's shape, minus the volatile bits. */
export type PortableRule = {
  /** The Graph rule id at export time; used to match live rules on import. */
  id?: string;
  displayName: string;
  sequence?: number;
  isEnabled?: boolean;
  conditions?: Record<string, unknown>;
  exceptions?: Record<string, unknown>;
  actions?: Record<string, unknown>;
};

export type RulesBackup = {
  format: string;
  exportedAt: string;
  rules: PortableRule[];
};

/** The fields a backup carries and the diff compares, in a stable order. */
export const RULE_FIELDS = [
  "displayName",
  "sequence",
  "isEnabled",
  "conditions",
  "exceptions",
  "actions",
] as const;

export type RuleField = (typeof RULE_FIELDS)[number];

/** Build the portable backup document from Graph's own rule objects. */
export function buildRulesBackup(graphRules: any[], now: Date): RulesBackup {
  return {
    format: RULES_BACKUP_FORMAT,
    exportedAt: now.toISOString(),
    rules: graphRules.map((rule) => {
      const portable: PortableRule = { displayName: String(rule?.displayName ?? "(unnamed)") };
      if (rule?.id) portable.id = String(rule.id);
      if (rule?.sequence !== undefined && rule?.sequence !== null) {
        portable.sequence = Number(rule.sequence);
      }
      if (typeof rule?.isEnabled === "boolean") portable.isEnabled = rule.isEnabled;
      for (const field of ["conditions", "exceptions", "actions"] as const) {
        if (rule?.[field] && typeof rule[field] === "object") portable[field] = rule[field];
      }
      return portable;
    }),
  };
}

/** Parse and validate a backup document. Never throws. */
export function parseRulesBackup(
  json: string
): { ok: true; backup: RulesBackup } | { ok: false; error: string } {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch (err) {
    return { ok: false, error: `not valid JSON: ${err instanceof Error ? err.message : String(err)}` };
  }
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    return { ok: false, error: "the backup must be a JSON object" };
  }
  const doc = parsed as Record<string, unknown>;
  if (doc.format !== RULES_BACKUP_FORMAT) {
    return {
      ok: false,
      error: `unrecognized format ${JSON.stringify(doc.format)} — expected "${RULES_BACKUP_FORMAT}" (from manage_rules export)`,
    };
  }
  if (!Array.isArray(doc.rules)) {
    return { ok: false, error: 'the backup has no "rules" array' };
  }
  for (const [index, rule] of (doc.rules as unknown[]).entries()) {
    if (typeof rule !== "object" || rule === null || Array.isArray(rule)) {
      return { ok: false, error: `rules[${index}] is not an object` };
    }
    const r = rule as Record<string, unknown>;
    if (typeof r.displayName !== "string" || r.displayName.trim() === "") {
      return { ok: false, error: `rules[${index}] has no displayName` };
    }
    for (const field of ["conditions", "exceptions", "actions"] as const) {
      if (r[field] !== undefined && (typeof r[field] !== "object" || Array.isArray(r[field]))) {
        return { ok: false, error: `rules[${index}].${field} is not an object` };
      }
    }
  }
  return { ok: true, backup: doc as unknown as RulesBackup };
}

/** Names of backup rules carrying forward/redirect actions. Must be empty to import. */
export function forwardingRuleNames(rules: PortableRule[]): string[] {
  return rules
    .filter((rule) => {
      const actions = (rule.actions ?? {}) as Record<string, unknown>;
      return ["forwardTo", "forwardAsAttachmentTo", "redirectTo"].some(
        (key) => Array.isArray(actions[key]) && (actions[key] as unknown[]).length > 0
      );
    })
    .map((rule) => rule.displayName);
}

/** True when a rule's conditions match at least one thing (Graph predicate shape). */
export function hasAnyCondition(conditions: Record<string, unknown> | undefined): boolean {
  if (!conditions) return false;
  return Object.values(conditions).some((value) =>
    Array.isArray(value) ? value.length > 0 : value !== undefined && value !== null && value !== ""
  );
}

export type RuleFieldDiff = { field: RuleField; live: unknown; backup: unknown };

export type RuleUpdate = {
  /** The live Graph rule being changed. */
  live: any;
  backup: PortableRule;
  fields: RuleFieldDiff[];
};

export type RulesDiff = {
  /** Backup rules with no live counterpart: would be created. */
  creates: PortableRule[];
  /** Matched rules whose fields differ: would be patched in place. */
  updates: RuleUpdate[];
  /** Matched rules that are already identical. */
  unchanged: PortableRule[];
  /** Live rules absent from the backup. NEVER deleted — listed only. */
  liveOnly: any[];
};

/**
 * Canonical JSON for comparing rule fields: keys sorted, null/empty values
 * dropped, and — one Graph quirk — senderContains values uppercased, because
 * Graph stores them uppercased and a backup would otherwise diff forever
 * against what it itself created.
 */
export function normalizeForDiff(value: unknown, key?: string): unknown {
  if (Array.isArray(value)) {
    const items = value
      .map((item) =>
        key === "senderContains" && typeof item === "string"
          ? item.toUpperCase()
          : normalizeForDiff(item)
      )
      .filter((item) => item !== undefined);
    return items.length === 0 ? undefined : items;
  }
  if (typeof value === "object" && value !== null) {
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(value as Record<string, unknown>).sort()) {
      const normalized = normalizeForDiff((value as Record<string, unknown>)[k], k);
      if (normalized !== undefined) out[k] = normalized;
    }
    return Object.keys(out).length === 0 ? undefined : out;
  }
  if (value === null || value === undefined) return undefined;
  return value;
}

function fieldEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(normalizeForDiff(a) ?? null) === JSON.stringify(normalizeForDiff(b) ?? null);
}

/**
 * Compare live rules against a backup. Matching is by rule id first (the ids
 * survive a same-mailbox round trip), then by display name case-insensitively
 * (for rules that were deleted and recreated since the export).
 */
export function diffRules(liveRules: any[], backupRules: PortableRule[]): RulesDiff {
  const diff: RulesDiff = { creates: [], updates: [], unchanged: [], liveOnly: [] };
  const matchedLiveIds = new Set<string>();

  for (const backup of backupRules) {
    let live = backup.id ? liveRules.find((rule) => rule?.id === backup.id) : undefined;
    if (!live) {
      live = liveRules.find(
        (rule) =>
          !matchedLiveIds.has(String(rule?.id)) &&
          String(rule?.displayName ?? "").toLowerCase() === backup.displayName.toLowerCase()
      );
    }
    if (!live) {
      diff.creates.push(backup);
      continue;
    }
    matchedLiveIds.add(String(live.id));

    const fields: RuleFieldDiff[] = [];
    for (const field of RULE_FIELDS) {
      if (backup[field] === undefined && field !== "exceptions") continue; // absent in backup: not asserted
      if (!fieldEqual(live?.[field], backup[field])) {
        fields.push({ field, live: live?.[field], backup: backup[field] });
      }
    }
    if (fields.length === 0) diff.unchanged.push(backup);
    else diff.updates.push({ live, backup, fields });
  }

  for (const rule of liveRules) {
    if (!matchedLiveIds.has(String(rule?.id))) diff.liveOnly.push(rule);
  }
  return diff;
}
