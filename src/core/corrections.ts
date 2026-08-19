// The auto-filer's feedback loop: notice when the user re-files a message the
// filer moved, and turn that correction into a sender→folder preference that
// future arrivals honor BEFORE any model call.
//
// Detection is reconciliation, not events: recent "moved" audit entries are
// compared against each message's CURRENT parentFolderId. (The alternative —
// change-notification subscriptions on every destination folder — would need
// one Graph subscription per folder against a hard platform cap, and still
// could not see moves between two folders that carry no subscription. One read
// per unreconciled recent move is bounded, transport-agnostic and cannot miss
// a move, only see it late.) It runs where the filer itself runs: on every
// accepted notification delivery — so a correction is learned before the next
// message from that sender is classified — and on the 6-hourly upkeep cron.
//
// This module lives INSIDE the classifier-side boundary: it imports no Graph
// transport and acts through the same ClassifierMailbox port the classifier
// uses, whose only mutations are move and categorize — and the reconciler
// itself calls neither; it only reads messages/folders and writes KV state.
// The offline boundary test walks this file's imports too.
import {
  AUDIT_CAP,
  auditSubject,
  extractAddress,
  isNeverFileFolder,
  isProtectedSubject,
  readAuditLog,
  upsertPreferenceFromCorrection,
  type AuditEntry,
} from "./auto-filing.js";
import { writeJson, type StateStore } from "./state.js";
import { STATE_LLM_AUDIT } from "./kv-keys.js";
import type { ClassifierMailbox, MailFacts } from "./classifier.js";

/** Moves older than this are no longer checked — the user has accepted them. */
export const RECONCILE_WINDOW_MS = 72 * 3600 * 1000;

/**
 * A move still in place after this long is marked confirmed and never re-read.
 * Younger ones stay unmarked so a correction made hours later is still seen.
 */
export const CONFIRM_AFTER_MS = 24 * 3600 * 1000;

/** At most this many message reads per reconcile pass. */
export const MAX_CHECKS_PER_RUN = 20;

export type ReconcileOutcome = {
  /** Entries whose message was actually read this pass. */
  checked: number;
  /** Corrections detected (preferences learned or reinforced). */
  corrections: number;
};

export type ReconcileContext = {
  store: StateStore;
  mailbox: ClassifierMailbox;
  /** The config's extra skip patterns; OTP-protected subjects never learn. */
  skipPatterns?: string[];
  now?: () => Date;
};

/**
 * Find the message an audit entry filed, wherever it is NOW. The direct read
 * by newMessageId only works while the message has not moved again: a Graph
 * move mints a NEW id and the old one 404s (verified live) — and the user's
 * correction IS a move. So when the direct read misses, the message is
 * re-found through its conversation id, which is stable across moves; among
 * the conversation's messages, only those with the same (audit-truncated)
 * subject count, and copies sitting in never-file folders (the Sent Items
 * copy of self-sent mail, a junked copy) are skipped rather than mistaken
 * for the filed one.
 */
async function locateFiledMessage(
  mailbox: ClassifierMailbox,
  entry: AuditEntry
): Promise<MailFacts | null> {
  const direct = await mailbox.readMessage(entry.newMessageId!);
  if (direct) return direct;
  if (!entry.conversationId) return null;

  const candidates = (await mailbox.findByConversation(entry.conversationId)).filter(
    (message) => entry.subject === undefined || auditSubject(message.subject) === entry.subject
  );
  // Still where the filer put it (under yet another id)? Then it is that copy.
  const inPlace = candidates.find((message) => message.parentFolderId === entry.folderId);
  if (inPlace) return inPlace;
  // Otherwise the first copy in a folder worth learning from.
  for (const message of candidates) {
    if (!message.parentFolderId) continue;
    const folder = await mailbox.getFolder(message.parentFolderId);
    if (folder && !isNeverFileFolder(folder.displayName)) return message;
  }
  return null;
}

/** True when this audit entry is a filing move the reconciler still watches. */
function isOpenMove(entry: AuditEntry, now: Date): boolean {
  return (
    entry.feature === "filing" &&
    (entry.action === "moved" || entry.action === "moved+categorized") &&
    typeof entry.folderId === "string" &&
    typeof entry.newMessageId === "string" &&
    entry.reconciled === undefined &&
    now.getTime() - Date.parse(entry.at) < RECONCILE_WINDOW_MS
  );
}

/**
 * One reconcile pass. Never throws — it runs in the background next to the
 * classifier, and a failed pass simply runs again on the next trigger.
 *
 * The audit ring is read once, mutated in memory (reconciled marks plus any
 * new "correction" entries) and written back once; like every ring-buffer
 * write in this server that is read-modify-write without a lock, an exactly
 * concurrent append can be lost — accepted at single-mailbox scale.
 */
export async function reconcileCorrections(context: ReconcileContext): Promise<ReconcileOutcome> {
  const now = context.now?.() ?? new Date();
  const outcome: ReconcileOutcome = { checked: 0, corrections: 0 };
  try {
    const audit = await readAuditLog(context.store);
    const candidates = audit.filter((entry) => isOpenMove(entry, now)).slice(0, MAX_CHECKS_PER_RUN);
    if (candidates.length === 0) return outcome;

    let mutated = false;
    const corrections: AuditEntry[] = [];

    for (const entry of candidates) {
      const message = await locateFiledMessage(context.mailbox, entry);
      outcome.checked++;

      if (!message) {
        // Deleted (or its folder was) — nothing to learn from a delete.
        entry.reconciled = "gone";
        mutated = true;
        continue;
      }
      if (message.parentFolderId === entry.folderId) {
        // Still where the filer put it. Old enough → settled; else keep watching.
        if (now.getTime() - Date.parse(entry.at) >= CONFIRM_AFTER_MS) {
          entry.reconciled = "confirmed";
          mutated = true;
        }
        continue;
      }

      // The message moved. Where to, and is that a correction worth learning?
      const target = message.parentFolderId
        ? await context.mailbox.getFolder(message.parentFolderId)
        : null;
      const sender = entry.sender ?? extractAddress(message.from) ?? undefined;
      if (!target || !sender) {
        entry.reconciled = "gone";
        mutated = true;
        continue;
      }
      if (
        isNeverFileFolder(target.displayName) ||
        isProtectedSubject(entry.subject ?? message.subject, context.skipPatterns ?? [])
      ) {
        // Deleted Items / Junk / Sent and friends are not filing choices, and a
        // protected (OTP-and-friends) subject may never teach a preference —
        // the compiled-in + configured skip list outranks the feedback loop.
        entry.reconciled = "ignored";
        mutated = true;
        continue;
      }

      const preference = await upsertPreferenceFromCorrection(
        context.store,
        sender,
        { id: target.id, name: target.displayName },
        now
      );
      entry.reconciled = "corrected";
      mutated = true;
      outcome.corrections++;

      const leaveAlone = target.displayName.trim().toLowerCase() === "inbox";
      corrections.push({
        at: now.toISOString(),
        feature: "filing",
        action: "correction",
        messageId: message.id,
        subject: entry.subject,
        sender,
        folder: target.displayName,
        folderId: target.id,
        reason:
          `you moved this out of ${entry.folder ?? "the filed folder"} to ${target.displayName}; ` +
          (leaveAlone
            ? `mail from ${sender} will now be left in the Inbox`
            : `mail from ${sender} will now be filed to ${target.displayName} without a model call`) +
          ` (correction ${preference.corrections})`,
      });
    }

    if (mutated || corrections.length > 0) {
      // One write: the corrections newest-first on top of the mutated ring.
      await writeJson(
        context.store,
        STATE_LLM_AUDIT,
        [...corrections.reverse(), ...audit].slice(0, AUDIT_CAP)
      );
    }
  } catch {
    // Background work; the next pass will try again.
  }
  return outcome;
}
