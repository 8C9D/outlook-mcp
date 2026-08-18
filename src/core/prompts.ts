// MCP prompts: reusable workflows the user can invoke by name. Each returns a
// single user message that instructs the calling model how to drive this server's
// tools. Both are deliberately zero-argument — every client renders those, and
// there is nothing here a follow-up turn cannot supply.
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

const TRIAGE_INBOX = `Triage my Outlook inbox. Work read-only until I approve, in this order:

1. Call list_folders to learn the folder tree (names, ids, unread counts) — you need real folder ids for any proposal.
2. Call manage_rules with action "list" to learn the rules that already exist. Every proposal you make must be consistent with these: do not propose a rule that duplicates, contradicts, or shadows one already there, and prefer extending an existing rule (manage_rules action "update", or an exception on it) over adding a near-duplicate.
3. Call search_mail with NO query (get_latest mode) on the inbox to sample the most recent arrivals — up to 25. Group them by sender and by evident type (newsletters, receipts, notifications, real correspondence, anything that looks like it needs a reply).
4. Report what you found as a compact summary: the groups, how many messages in each, and which ones look like they need my attention.
5. Then PROPOSE, do not perform:
   - batched moves: for each group worth filing, the exact message ids and the exact destination folder (naming a folder that exists, or a new folder to create with create_folder);
   - rules: for each recurring pattern, the complete rule in plain language — every condition, every exception, every action — and say plainly that a rule will act automatically on all future matching mail with no further approval;
   - anything you think should simply be deleted (soft delete, recoverable from Deleted Items).

Do not call create_folder, create_draft, manage_message, manage_categories, manage_rules with create/update/delete, or any other writing tool during this triage. Present the proposals and stop. Wait for me to say explicitly which proposals to apply; apply only those, and nothing you inferred alongside them.`;

const MORNING_BRIEF = `Give me my morning brief. Gather all three of these first, then write the brief:

1. Recent mail: call search_mail with NO query (get_latest mode) on the inbox, max_results 25. Keep only messages received within the last 24 hours, judging by the "At:" timestamps (America/Toronto). Note that this listing does not report read/unread state — if whether I have already read something changes what you would say about it, call read_message on that message rather than guessing.
2. Today's calendar: call list_events with days 1 for today's schedule.
3. Open tasks: call list_tasks with due_within_days 3 (open tasks only) — this returns overdue, today, and upcoming groups.

Then produce a compact brief, in this shape:

- Calendar — today's events in time order, one line each; call out anything starting within the next two hours and any gaps big enough to work in.
- Mail — the last 24 hours grouped by what they are, newest first. One line per message that matters (sender, subject, why it matters). Collapse bulk mail into a count instead of listing it.
- Tasks — overdue first (these are the ones I have already missed), then due today, then the next few days.
- ACTIONS NEEDED — a short numbered list of the things only I can decide or do today: mail that needs a reply, tasks colliding with meetings, deadlines about to pass. Reference the specific message subject or task title for each, and keep it to the items that genuinely need me.

Be brief and concrete: no restating the raw tool output, no filler. If a section is empty, say so in one line. Do not send mail, create tasks, or change anything — this is a read-only briefing.`;

export function registerPrompts(server: McpServer): void {
  server.registerPrompt(
    "triage_inbox",
    {
      title: "Triage inbox",
      description:
        "Survey the inbox's recent arrivals against the existing folder and rule set, then propose batched moves and new/updated inbox rules for explicit approval. Proposes only — makes no changes.",
    },
    () => ({
      messages: [{ role: "user", content: { type: "text", text: TRIAGE_INBOX } }],
    })
  );

  server.registerPrompt(
    "morning_brief",
    {
      title: "Morning brief",
      description:
        "Compact start-of-day briefing: mail from the last 24 hours, today's calendar, tasks due within 3 days, and an actions-needed list. Read-only.",
    },
    () => ({
      messages: [{ role: "user", content: { type: "text", text: MORNING_BRIEF } }],
    })
  );
}
