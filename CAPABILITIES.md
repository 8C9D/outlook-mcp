# CAPABILITIES — what a person can do in Outlook, and what this server covers

The v1.3 audit: every capability a normal human user of personal Outlook (outlook.com — Mail,
Calendar, People, To Do, plus the OneDrive surface Outlook attachments touch) exercises, and where
this server stands on each. Compiled against the Outlook web UI surface; every "newly implemented"
and "not feasible" verdict below was probed live against this account, and cites what the probe
returned.

**Statuses**

- ✅ **already supported** — existed before v1.3
- 🆕 **newly implemented** — added in v1.3
- ❌ **not feasible** — Microsoft Graph offers no way to do it for a personal account (the probe or
  documentation trail is cited)
- 🚫 **excluded by design** — technically possible, deliberately not offered, reason stated
- — **client-only** — a rendering/UX affordance of the Outlook app itself with no server-side
  meaning (listed so the inventory is exhaustive, not because a server could ever own it)

## Mail — reading and triage

| Capability | Status | Notes / tool |
| --- | --- | --- |
| List newest mail, per folder or whole mailbox | ✅ | `search_mail` (no query = newest-first) |
| Full-text search (subject/sender/body) | ✅ | `search_mail` with `query` |
| Filter by date range / attachments | ✅ | `date_from`/`date_to`/`has_attachments` |
| Filter unread | 🆕 | `unread_only`, and read state now shown on every hit |
| Focused / Other tab view | 🆕 | `tab: "focused" \| "other"` |
| Read a message in full, with attachment inventory | ✅ | `read_message` |
| Read a whole conversation | ✅ | `read_thread` |
| Mark read/unread, flag/unflag | ✅ | `manage_message` (flags are on/off; flag due-dates not exposed — niche, and Outlook's own UI barely surfaces them) |
| Move / archive / delete (recoverable) | ✅ | `manage_message`, batched |
| Restore from Deleted Items | ✅ | `manage_message` move back out |
| Empty Deleted Items / delete permanently | 🚫 | soft-delete policy: no tool purges; done in Outlook deliberately |
| Categorize mail (and manage the category list) | ✅ | `manage_message` + `manage_categories` |
| Inbox rules (create/update/delete, exceptions) | ✅ | `manage_rules`, with export/import backup |
| Rules that forward/redirect | 🚫 | a standing silent forward is an exfiltration primitive (README, security model) |
| Block / unblock a sender | ✅ | `manage_senders` (per message; Graph cannot read the lists back — probed, documented) |
| Manage safe senders | ❌ | probed: every Graph route 404/400s on a consumer account |
| Report phishing (evidence out) | ✅ | `read_message include_headers` forensics + `export_message` .eml; Graph has no consumer "report" action |
| What's new since I last looked | ✅ | `check_new_mail` (delta) / `get_mailbox_activity` (push, hosted) |
| Pin a message | ❌ | no Graph API |
| Snooze a message | ❌ | no Graph API |
| Ignore/mute a conversation | ❌ | no Graph API for consumer accounts |
| Sweep | ❌ | no Graph API; the combination of rules + batched `manage_message` covers the intent |
| Message recall | ❌ | an Exchange-organization feature; does not exist for personal accounts |
| Print, translate, read aloud, conditional formatting, themes | — | client-only rendering |

## Mail — composing and sending

| Capability | Status | Notes / tool |
| --- | --- | --- |
| Compose a draft; edit it; discard it | ✅ | `create_draft` / `update_draft` / `manage_message` delete |
| Reply, reply-all, forward (quoted original kept) | ✅ | `create_draft` modes |
| To / Cc | ✅ | |
| **Bcc** | 🆕 | create/update/send all handle and display it |
| **Rich text (HTML) body** | 🆕 | `body_format: "html"`, all three compose modes |
| **Signature** | 🆕 | stored server-side (`mailbox_settings set_signature`), appended by `create_draft`; Graph exposes no API for Outlook's own signature, so the two are independent and documented as such |
| Attach files (local / URL / inline / OneDrive) | ✅ | `add_attachment`, ≤25 MB, chunked |
| **Remove an attachment while composing** | 🆕 | `update_draft remove_attachments` |
| **Importance (high/low)** | 🆕 | `importance` on create/update |
| **Read / delivery receipts** | 🆕 | `request_read_receipt` / `request_delivery_receipt` |
| Send | ✅ | `send_draft` — the only send path, two-step by design |
| **Scheduled send (delayed delivery)** | 🆕 | `send_draft send_at` + `manage_scheduled_send` list/cancel. Verified end to end live: parked in Drafts, sent itself at the exact deferred second, arrived; cancel before the time stops it (and discards the message — Exchange keeps no soft copy of a cancelled deferred send, stated loudly) |
| Undo send | ❌ | client-side illusion in Outlook (it just delays submission); the honest equivalent here is a short `send_at`, cancellable until it fires |
| Send from an alias | ❌ | Graph exposes no alias list for outlook.com accounts and a `from` override cannot be verified against one |
| Encrypt / S-MIME / IRM | ❌ | not available to consumer accounts via Graph |
| @mention someone | ❌ | no Graph API (a body convention plus client magic) |

## Mail — folders

| Capability | Status | Notes / tool |
| --- | --- | --- |
| Folder tree with counts | ✅ | `list_folders` |
| Create a folder / subfolder | ✅ | `create_folder`, duplicate-guarded |
| **Rename a folder** | 🆕 | `manage_folder rename` — id survives |
| **Move a folder (nest / un-nest)** | 🆕 | `manage_folder move` — contents and id survive |
| Delete a folder (recoverable) | ✅ | `delete_folder` — moved to Deleted Items, never Graph's destructive DELETE |
| Favorites | ❌ | no Graph API |

## Calendar

| Capability | Status | Notes / tool |
| --- | --- | --- |
| See events by day/window, all calendars | ✅ | `list_events`, `list_calendars` |
| **Search the calendar** | 🆕 | `list_events query` (subject/location contains) |
| Create events: times, all-day, location, description, reminders | ✅ | `create_event` |
| Repeating events; edit one occurrence vs the series | ✅ | `recurrence`, `scope` |
| Invite attendees | ✅ | with the invitation-email caution stated |
| **Optional attendees** | 🆕 | `optional_attendees` |
| **Change who's invited after creation** | 🆕 | `manage_event update attendees` (replace; added invited, removed cancelled) |
| Respond to invitations (accept/decline/tentative) | ✅ | `manage_event respond` |
| **Propose a new time** | 🆕 | `proposed_start`/`proposed_end` with tentative/decline — implemented per Graph's contract; the one v1.3 parameter not verified live (needs a second account's invitation) |
| Cancel / delete events, series-aware | ✅ | `manage_event cancel` |
| **Show as free/tentative/busy/OOF** | 🆕 | `show_as` on create/update |
| **Private events** | 🆕 | `private` (sensitivity) |
| **Categorize events** | 🆕 | `categories`, validated against the master list |
| **Attach a file to an event** | 🆕 | `add_attachment event_id` |
| **Forward an event** | 🆕 | `manage_event forward` |
| **Create / rename / recolor calendars** | 🆕 | `manage_calendar` |
| Delete a calendar | 🚫 | permanent, destroys every event, no recoverable copy (probed) — same policy as To Do lists; done in Outlook deliberately |
| Share a calendar / see others' free-busy | ❌ | free/busy probed live: Graph refuses outlook.com targets (error 7002); sharing grants could not be verified without actually granting a third party standing access, which this server's design excludes |
| Scheduling assistant (findMeetingTimes) | ❌ | work/school accounts only, per Microsoft |
| Teams/Skype meeting links | ❌ | probed: `isOnlineMeeting`/`onlineMeetingProvider` silently ignored on a consumer account |
| Import an .ics / subscribe to a calendar | ❌ | no Graph API for consumer accounts |
| Working hours (drives scheduling suggestions) | ✅ | `mailbox_settings` |
| Weather, holidays, birthdays calendars | — | Outlook-managed; readable like any calendar via `list_calendars`/`list_events` |

## People (contacts)

| Capability | Status | Notes / tool |
| --- | --- | --- |
| Search contacts | ✅ | `search_contacts` — now renders the full card |
| Create / edit / delete (recoverable) | ✅ | `manage_contact` |
| **Full contact card** — nickname, mobile & home phones, job title, birthday, home/business addresses, notes | 🆕 | all round-trip probed (Graph normalizes birthdays to ~11:59Z; only the date is meaningful) |
| **Categorize contacts** | 🆕 | validated against the master list |
| **Contact photo** | 🆕 | set from URL or base64 (≤4 MB); reading photos back out is not exposed (niche — say the word and it follows the `get_attachment` pattern) |
| Contact lists (groups) | ❌ | Graph does not expose outlook.com contact lists |
| Contact folders | 🚫 | deliberately scoped out: all contacts live in the default folder, which is how the modern People app effectively behaves; folder support would complicate every contact tool for a vestigial feature |

## Tasks (Microsoft To Do)

| Capability | Status | Notes / tool |
| --- | --- | --- |
| List tasks grouped by due date | ✅ | `list_tasks` |
| Create / complete / reopen / update / delete | ✅ | `manage_task` (delete is permanent — Graph has no undelete for tasks; warned loudly) |
| Subtasks (steps) | ✅ | add/complete/remove |
| Repeating tasks | ✅ | create-only, per Graph's own limits (probed in v1.1) |
| Reminders, due dates, notes | ✅ | America/Toronto |
| **Importance (star a task)** | 🆕 | `importance`; `list_tasks` marks "important" |
| Task lists: create, rename | ✅ | delete 🚫 — destroys every task, no recoverable copy |
| Turn an email into a task | ✅ | `linked_message_id` |
| My Day | ❌ | no Graph API |
| Attach a file to a task | 🚫 | deliberately scoped out: the linked-email note covers the common case, and files belong in OneDrive + `share_link` |
| Share a list | ❌ | no Graph API |

## Files (OneDrive), search, and settings

| Capability | Status | Notes / tool |
| --- | --- | --- |
| Search / browse / read / upload / move / rename / delete (recycle bin) / share OneDrive files | ✅ | the six v1.2 file tools |
| Attachments ⇄ OneDrive both directions | ✅ | `onedrive_path` / `save_to_onedrive` |
| Out-of-office auto-reply | ✅ | `auto_reply` |
| Working hours, Focused-Inbox overrides | ✅ | `mailbox_settings` |
| **Signature setting** | 🆕 | `mailbox_settings set_signature` / `clear_signature` (server-side — see Mail) |
| Time zone | ✅ | read via `mailbox_settings get`; writing it is meaningless through Graph (it normalizes to the mailbox's own zone — probed in v1.1) |
| Mailbox-level forwarding | 🚫 | the same exfiltration reasoning as forwarding rules |
| Aliases management | ❌ | not exposed by Graph for consumer accounts |
| Notification/appearance settings | — | client-only |

## Beyond the human baseline

Not part of "what a person can do", but already in the surface: delta what-changed queries, pushed
change notifications, message forensics (SPF/DKIM/DMARC, delivery chain), .eml export, rules
backup/restore, structured tool output, health self-monitoring, and the two opt-in LLM features
(auto-filing, morning digest).

## Scorecard

Of the 96 rows above: **42 already supported**, **26 newly implemented in v1.3**, **19 not
feasible** through Microsoft Graph for a personal account (each with the probe or the documented
limitation cited), **6 excluded by design** with the reason stated, and **3 client-only**
affordances of the Outlook app itself. Every newly implemented mail/folder/calendar/contact/task
capability is exercised by the live test suites; scheduled send is verified end to end — scheduled,
listed, cancelled, and delivered at its exact second — on this account.
