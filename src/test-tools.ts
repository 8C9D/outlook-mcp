// Test harness: exercises the tool handlers directly (bypassing the MCP
// transport) against the real account using the cached token, plus a stdio
// protocol smoke test of the server itself. Every test cleans up after itself;
// a final check verifies no "[MCP TEST]" artifacts remain in the account
// (messages, drafts, folders, events, calendars, contacts, inbox rules, categories,
// To Do tasks, temp files) and that mailbox settings (auto-reply) are restored exactly.
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { GraphError, callGraphServer, graphRequestLog } from "./core/graph.js";
import { PROJECT_ROOT } from "./project-root.js";
import { searchMailHandler } from "./tools/search-mail.js";
import { readThreadHandler } from "./tools/read-thread.js";
import { readMessageHandler } from "./tools/read-message.js";
import { createDraftHandler } from "./tools/create-draft.js";
import { updateDraftHandler } from "./tools/update-draft.js";
import { sendDraftHandler } from "./tools/send-draft.js";
import { manageMessageHandler } from "./tools/manage-message.js";
import { listFoldersHandler } from "./tools/list-folders.js";
import { listEventsHandler, torontoToday, addDays } from "./tools/list-events.js";
import { listCalendarsHandler } from "./tools/list-calendars.js";
import { createEventHandler } from "./tools/create-event.js";
import { manageEventHandler } from "./tools/manage-event.js";
import { searchContactsHandler } from "./tools/search-contacts.js";
import { manageContactHandler } from "./tools/manage-contact.js";
import { autoReplyHandler } from "./tools/auto-reply.js";
import { addAttachmentHandler } from "./tools/add-attachment.js";
import { getAttachmentHandler } from "./tools/get-attachment.js";
import { manageRulesHandler } from "./tools/manage-rules.js";
import { createFolderHandler } from "./tools/create-folder.js";
import { manageCategoriesHandler } from "./tools/manage-categories.js";
import { listTasksHandler, resolveTaskList } from "./tools/list-tasks.js";
import { manageTaskHandler } from "./tools/manage-task.js";
import { manageSendersHandler } from "./tools/manage-senders.js";
import { mailboxSettingsHandler } from "./tools/mailbox-settings.js";
import { exportMessageHandler } from "./tools/export-message.js";
import { SAVE_DIR } from "./tools/save-local.js";
import { checkNewMailHandler } from "./tools/check-new-mail.js";
import { getMailboxActivityHandler } from "./tools/get-mailbox-activity.js";
import {
  ACTIVITY_CAP,
  appendActivity,
  handleNotificationRequest,
  readActivity,
} from "./core/notifications.js";
import {
  ensureMailSubscription,
  renewalDecision,
  SUBSCRIPTION_RESOURCE,
  type SubscriptionRecord,
} from "./core/subscriptions.js";
import { STATE_ACTIVITY, STATE_SUBSCRIPTION, deltaKey, downloadKey } from "./core/kv-keys.js";
import { DOWNLOAD_ROUTE_PREFIX, readDownload } from "./core/downloads.js";
import { createMemoryStateStore, runWithStateStore, writeJson } from "./core/state.js";
import { installFileStateStore } from "./state-file.js";
import { FOLDERS_URI, RECENT_INBOX_URI } from "./core/resources.js";
import type { ToolResult } from "./tools/common.js";
import { installMsalTokenProvider } from "./auth.js";

// Tool handlers reach Graph through core/token.js; in Node the token comes from
// MSAL and the disk cache, exactly as under the stdio server.
installMsalTokenProvider();

// check_new_mail remembers its delta position in the state store. The harness
// installs the real file-backed store, but pointed at a throwaway file so a run
// never disturbs (or is disturbed by) the server's own position; the final
// sweep deletes it.
const TEST_STATE_FILE = path.join(
  os.tmpdir(),
  `mcp-test-state-${createHash("sha256").update(String(process.pid)).digest("hex").slice(0, 8)}.json`
);
installFileStateStore(TEST_STATE_FILE);

const TEST_PREFIX = "[MCP TEST]";

type Outcome = { name: string; passed: boolean; detail?: string };
const outcomes: Outcome[] = [];

async function test(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    outcomes.push({ name, passed: true });
    console.log(`PASS  ${name}`);
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    outcomes.push({ name, passed: false, detail });
    console.log(`FAIL  ${name}\n      ${detail.split("\n").join("\n      ")}`);
  }
}

function assert(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

function toolText(result: ToolResult, context: string): string {
  const text = result.content.map((c) => c.text).join("\n");
  assert(!result.isError, `${context} returned isError: ${text}`);
  return text;
}

function expectError(result: ToolResult, context: string): string {
  const text = result.content.map((c) => c.text).join("\n");
  assert(result.isError, `${context} should have returned isError but succeeded: ${text}`);
  return text;
}

async function expect404(pathToGet: string, what: string): Promise<void> {
  try {
    await callGraphServer(pathToGet);
    throw new Error(`${what} still exists after deletion`);
  } catch (err) {
    if (err instanceof GraphError && err.status === 404) return;
    throw err;
  }
}

async function poll<T>(
  what: string,
  timeoutMs: number,
  fn: () => Promise<T | undefined>
): Promise<T> {
  const deadline = Date.now() + timeoutMs;
  for (;;) {
    const value = await fn();
    if (value !== undefined) return value;
    if (Date.now() > deadline) throw new Error(`Timed out after ${timeoutMs}ms waiting for ${what}`);
    await new Promise((r) => setTimeout(r, 5000));
  }
}

/** TEST-ONLY: permanently remove [MCP TEST] messages (incl. soft-deleted copies).
 * The tools themselves never purge — deletes in the tool surface are always soft. */
async function purgeTestMessages(): Promise<void> {
  const data = await callGraphServer(
    `/me/messages?$filter=${encodeURIComponent(`startswith(subject,'${TEST_PREFIX}')`)}&$select=id,subject`
  );
  for (const msg of data?.value ?? []) {
    await callGraphServer(`/me/messages/${encodeURIComponent(msg.id)}/permanentDelete`, {
      method: "POST",
    });
  }
}

/** TEST-ONLY: permanently remove any [MCP TEST] mail folders (top level and in Deleted Items). */
async function purgeTestFolders(): Promise<void> {
  const candidates: any[] = [];
  const top = await callGraphServer("/me/mailFolders?$top=100&$select=id,displayName");
  candidates.push(...(top?.value ?? []));
  const deleted = await callGraphServer(
    "/me/mailFolders/deleteditems/childFolders?$top=100&$select=id,displayName"
  );
  candidates.push(...(deleted?.value ?? []));
  for (const folder of candidates) {
    if (String(folder.displayName ?? "").startsWith(TEST_PREFIX)) {
      await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/permanentDelete`, {
        method: "POST",
      });
    }
  }
}

/** TEST-ONLY: remove any [MCP TEST] inbox rules. */
async function purgeTestRules(): Promise<void> {
  const data = await callGraphServer("/me/mailFolders/inbox/messageRules");
  for (const rule of data?.value ?? []) {
    if (String(rule.displayName ?? "").startsWith(TEST_PREFIX)) {
      await callGraphServer(`/me/mailFolders/inbox/messageRules/${encodeURIComponent(rule.id)}`, {
        method: "DELETE",
      });
    }
  }
}

/** TEST-ONLY: remove any [MCP TEST] master categories. */
async function purgeTestCategories(): Promise<void> {
  const data = await callGraphServer("/me/outlook/masterCategories?$top=100");
  for (const category of data?.value ?? []) {
    if (String(category.displayName ?? "").startsWith(TEST_PREFIX)) {
      await callGraphServer(`/me/outlook/masterCategories/${encodeURIComponent(category.id)}`, {
        method: "DELETE",
      });
    }
  }
}

/** TEST-ONLY: remove any [MCP TEST] calendars, and test events in the ones that stay. */
async function purgeTestCalendars(): Promise<void> {
  const calendars = await callGraphServer("/me/calendars?$select=id,name&$top=100");
  for (const calendar of calendars?.value ?? []) {
    if (String(calendar.name ?? "").startsWith(TEST_PREFIX)) {
      await callGraphServer(`/me/calendars/${encodeURIComponent(calendar.id)}`, {
        method: "DELETE",
      });
      continue;
    }
    const events = await callGraphServer(
      `/me/calendars/${encodeURIComponent(calendar.id)}/events?$filter=${encodeURIComponent(
        `startswith(subject,'${TEST_PREFIX}')`
      )}&$select=id,subject&$top=100`
    );
    for (const event of events?.value ?? []) {
      await callGraphServer(`/me/events/${encodeURIComponent(event.id)}`, { method: "DELETE" });
    }
  }
}

/** TEST-ONLY: remove any [MCP TEST] To Do lists, tasks and all.
 * manage_task deliberately cannot delete a list (it would destroy the tasks in
 * it with no recoverable copy), so the harness cleans up its own lists with a
 * raw Graph DELETE rather than through the tool surface. */
async function purgeTestTaskLists(): Promise<void> {
  const lists = await callGraphServer("/me/todo/lists?$top=100");
  for (const list of lists?.value ?? []) {
    if (String(list.displayName ?? "").startsWith(TEST_PREFIX)) {
      await callGraphServer(`/me/todo/lists/${encodeURIComponent(list.id)}`, { method: "DELETE" });
    }
  }
}

/** TEST-ONLY: remove any Focused-Inbox override this harness created. */
async function purgeTestFocusOverrides(): Promise<void> {
  const data = await callGraphServer("/me/inferenceClassification/overrides?$top=100");
  for (const override of data?.value ?? []) {
    const address = String(override.senderEmailAddress?.address ?? "");
    if (address === FOCUS_TEST_SENDER || address.startsWith("mcp-test-")) {
      await callGraphServer(
        `/me/inferenceClassification/overrides/${encodeURIComponent(override.id)}`,
        { method: "DELETE" }
      );
    }
  }
}

/** TEST-ONLY: remove any [MCP TEST] To Do tasks, across every list. */
async function purgeTestTasks(): Promise<void> {
  const lists = await callGraphServer("/me/todo/lists?$top=100");
  for (const list of lists?.value ?? []) {
    const tasks = await callGraphServer(
      `/me/todo/lists/${encodeURIComponent(list.id)}/tasks?$top=100`
    );
    for (const task of tasks?.value ?? []) {
      if (String(task.title ?? "").startsWith(TEST_PREFIX)) {
        await callGraphServer(
          `/me/todo/lists/${encodeURIComponent(list.id)}/tasks/${encodeURIComponent(task.id)}`,
          { method: "DELETE" }
        );
      }
    }
  }
}

// ---- shared fixtures ----------------------------------------------------

const me = await callGraphServer("/me?$select=mail,userPrincipalName");
const ownAddress: string = me.mail ?? me.userPrincipalName;
console.log(`Running against account: ${ownAddress}\n`);

const latestInbox = await callGraphServer(
  "/me/mailFolders/inbox/messages?$top=1&$select=id,subject,conversationId"
);
const latestMessage = latestInbox?.value?.[0];

// Cross-test state for the v2/v3 lifecycle tests.
const state: {
  receivedId?: string;
  sentId?: string;
  testFolderId?: string;
  savedAutoReply?: any;
  savedWorkingHours?: any;
  exportedEmlPath?: string;
  tempDir?: string;
} = {};

/** Create the run's temp dir on first use; every temp file for tests lives here. */
async function tempDir(): Promise<string> {
  if (!state.tempDir) state.tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "mcp-test-"));
  return state.tempDir;
}

/** Extract "Draft id: …" from a create_draft output. */
function extractDraftId(createText: string): string {
  const draftId = createText.match(/Draft id: (\S+)/)?.[1];
  assert(draftId, `Could not extract draft id from output: ${createText}`);
  return draftId;
}

// ---- v1: search_mail ----------------------------------------------------

await test("v1a. search_mail (term from latest inbox subject)", async () => {
  assert(latestMessage, "Inbox is empty; cannot derive a guaranteed search term");
  const subject: string = latestMessage.subject ?? "";
  const tokens = subject.match(/[A-Za-z0-9]{3,}/g) ?? [];
  const term = tokens.sort((a, b) => b.length - a.length)[0] ?? subject;
  assert(term, `Could not derive a search term from subject ${JSON.stringify(subject)}`);
  const text = toolText(await searchMailHandler({ query: term }), "search_mail");
  assert(/result\(s\)/.test(text), `Expected results for ${JSON.stringify(term)}, got: ${text}`);
  assert(text.includes("Conversation id:"), "Output missing conversation ids");
});

// ---- v1: read_thread ----------------------------------------------------

await test("v1b. read_thread (latest inbox conversation)", async () => {
  assert(latestMessage?.conversationId, "No conversationId available from the latest inbox message");
  const text = toolText(
    await readThreadHandler({ conversation_id: latestMessage.conversationId }),
    "read_thread"
  );
  assert(text.startsWith("Thread:"), `Unexpected output: ${text.slice(0, 200)}`);
  assert(text.includes("From:") && text.includes("Date:"), "Output missing sender/date lines");
});

// ---- v1: list_events ----------------------------------------------------

await test("v1d. list_events (next 7 days)", async () => {
  const text = toolText(await listEventsHandler({}), "list_events");
  assert(
    text.startsWith("Events ") || text.startsWith("No events in this window"),
    `Unexpected output: ${text.slice(0, 200)}`
  );
});

// ---- a. read_message + attachment inventory ------------------------------

await test("a. read_message (latest inbox message, attachment inventory)", async () => {
  assert(latestMessage, "Inbox is empty; nothing to read");
  const text = toolText(
    await readMessageHandler({ message_id: latestMessage.id }),
    "read_message"
  );
  assert(text.includes("Subject:"), "Output missing Subject header");
  assert(text.includes("From:") && text.includes("To:"), "Output missing From/To headers");
  assert(text.includes("Date:"), "Output missing Date header");
  assert(text.includes("Message id:"), "Output missing message id");
  // Attachment shape assertion: either an inventory with ids or an explicit "none".
  assert(
    /Attachments \(\d+\):/.test(text) || text.includes("Attachments: none"),
    "Output missing attachment inventory section"
  );
  if (/Attachments \(\d+\):/.test(text)) {
    assert(text.includes("Attachment id:"), "Attachment inventory missing attachment ids");
  }
});

// ---- b. draft lifecycle: create → update → send → arrive -----------------

await test("b. draft lifecycle (create → update_draft → send_draft → arrives in inbox)", async () => {
  const subject = `${TEST_PREFIX} v2`;
  const createText = toolText(
    await createDraftHandler({
      to: [ownAddress],
      subject,
      body: "Original body from the v2 test harness.",
    }),
    "create_draft"
  );
  const draftId = createText.match(/Draft id: (\S+)/)?.[1];
  assert(draftId, `Could not extract draft id from output: ${createText}`);

  const updatedBody = "Updated body from the v2 test harness. Safe to delete.";
  const updateText = toolText(
    await updateDraftHandler({ draft_id: draftId, body: updatedBody }),
    "update_draft"
  );
  assert(updateText.includes("Draft updated"), `Unexpected update output: ${updateText}`);

  const sendText = toolText(await sendDraftHandler({ draft_id: draftId }), "send_draft");
  assert(sendText.includes("Draft sent"), `Unexpected send output: ${sendText}`);
  assert(sendText.includes(ownAddress), "Send confirmation missing recipient");

  // Poll the inbox (up to 60s) for the arrived copy and verify the updated body made it.
  const arrived = await poll("test message to arrive in inbox", 60000, async () => {
    const found = await callGraphServer(
      `/me/mailFolders/inbox/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$select=id,subject,bodyPreview,isRead`
    );
    return found?.value?.[0];
  });
  assert(
    String(arrived.bodyPreview ?? "").includes("Updated body"),
    `Arrived message does not carry the updated body: ${arrived.bodyPreview}`
  );
  state.receivedId = arrived.id;

  const sent = await callGraphServer(
    `/me/mailFolders/sentitems/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$select=id`
  );
  assert(sent?.value?.length, "Sent copy not found in Sent Items");
  state.sentId = sent.value[0].id;
});

// ---- b2. send_draft refuses non-drafts -----------------------------------

await test("b2. send_draft rejects a non-draft id", async () => {
  assert(state.receivedId, "No received test message from test b");
  const text = expectError(
    await sendDraftHandler({ draft_id: state.receivedId }),
    "send_draft(non-draft)"
  );
  assert(/not a draft/i.test(text), `Unexpected error text: ${text}`);
});

// ---- c. manage_message on the received test message ----------------------

await test("c. manage_message (mark_unread, flag, move to test folder, cleanup)", async () => {
  assert(state.receivedId, "No received test message from test b");
  let id = state.receivedId;

  toolText(
    await manageMessageHandler({ message_ids: [id], action: "mark_unread" }),
    "manage_message mark_unread"
  );
  let msg = await callGraphServer(`/me/messages/${encodeURIComponent(id)}?$select=isRead`);
  assert(msg.isRead === false, "Message is still marked read");

  toolText(await manageMessageHandler({ message_ids: [id], action: "flag" }), "manage_message flag");
  msg = await callGraphServer(`/me/messages/${encodeURIComponent(id)}?$select=flag`);
  assert(msg.flag?.flagStatus === "flagged", `Flag not set: ${JSON.stringify(msg.flag)}`);

  const folder = await callGraphServer("/me/mailFolders", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: `${TEST_PREFIX} folder` }),
  });
  state.testFolderId = folder.id;

  const moveText = toolText(
    await manageMessageHandler({
      message_ids: [id],
      action: "move",
      destination_folder: folder.id,
    }),
    "manage_message move"
  );
  const newId = moveText.match(/new id: (\S+?)\)/)?.[1];
  assert(newId, `Move output missing the new message id: ${moveText}`);
  id = newId;
  msg = await callGraphServer(`/me/messages/${encodeURIComponent(id)}?$select=parentFolderId`);
  assert(msg.parentFolderId === folder.id, "Message is not in the test folder after move");

  // Cleanup: soft-delete the received copy and the sent copy via the tool, then
  // remove the test folder (permanent, test-only).
  const delText = toolText(
    await manageMessageHandler({ message_ids: [id, state.sentId!], action: "delete" }),
    "manage_message delete"
  );
  assert(delText.includes("2/2"), `Expected both deletions to succeed: ${delText}`);
  await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/permanentDelete`, {
    method: "POST",
  });
  state.testFolderId = undefined;
});

// ---- d. list_folders -----------------------------------------------------

await test("d. list_folders (inbox/drafts/sent items with counts)", async () => {
  const text = toolText(await listFoldersHandler({}), "list_folders");
  for (const name of ["Inbox", "Drafts", "Sent Items"]) {
    assert(text.includes(name), `Folder list missing ${name}`);
  }
  assert(/\d+ unread \/ \d+ total/.test(text), "Folder list missing unread/total counts");
  assert(text.includes("id: "), "Folder list missing folder ids");
});

// ---- e. event lifecycle: create → update → cancel ------------------------

await test("e. event lifecycle (create → shift 30 min → cancel → gone)", async () => {
  const tomorrow = addDays(torontoToday(), 1);
  const createText = toolText(
    await createEventHandler({
      subject: `${TEST_PREFIX} v2 event`,
      start: `${tomorrow}T09:00`,
      end: `${tomorrow}T09:30`,
    }),
    "create_event"
  );
  const id = createText.match(/Event id: (\S+)/)?.[1];
  assert(id, `Could not extract event id from output: ${createText}`);

  const updateText = toolText(
    await manageEventHandler({
      event_id: id,
      action: "update",
      start: `${tomorrow}T09:30`,
      end: `${tomorrow}T10:00`,
    }),
    "manage_event update"
  );
  assert(updateText.includes("Event updated"), `Unexpected update output: ${updateText}`);
  const event = await callGraphServer(`/me/events/${encodeURIComponent(id)}?$select=start`, {
    headers: { Prefer: 'outlook.timezone="America/Toronto"' },
  });
  assert(
    String(event.start?.dateTime ?? "").startsWith(`${tomorrow}T09:30`),
    `Event start not shifted: ${JSON.stringify(event.start)}`
  );

  const cancelText = toolText(
    await manageEventHandler({ event_id: id, action: "cancel" }),
    "manage_event cancel"
  );
  assert(/removed|cancelled/i.test(cancelText), `Unexpected cancel output: ${cancelText}`);
  await expect404(`/me/events/${encodeURIComponent(id)}?$select=id`, "Event");
});

// ---- f. contact lifecycle ------------------------------------------------

await test("f. contact lifecycle (create → search → update phone → delete → gone)", async () => {
  const createText = toolText(
    await manageContactHandler({
      action: "create",
      given_name: TEST_PREFIX,
      surname: "Contact",
      emails: ["mcp-test@example.invalid"],
    }),
    "manage_contact create"
  );
  const id = createText.match(/Contact id: (\S+)/)?.[1];
  assert(id, `Could not extract contact id from output: ${createText}`);

  const searchText = toolText(
    await searchContactsHandler({ query: TEST_PREFIX }),
    "search_contacts"
  );
  assert(searchText.includes(id), `Search did not find the new contact: ${searchText}`);

  const updateText = toolText(
    await manageContactHandler({ action: "update", contact_id: id, phones: ["555-0100"] }),
    "manage_contact update"
  );
  assert(updateText.includes("555-0100"), `Phone not updated: ${updateText}`);

  toolText(
    await manageContactHandler({ action: "delete", contact_id: id }),
    "manage_contact delete"
  );
  const afterText = toolText(
    await searchContactsHandler({ query: TEST_PREFIX }),
    "search_contacts after delete"
  );
  assert(
    afterText.startsWith("No contacts matching"),
    `Deleted contact still found: ${afterText}`
  );
});

// ---- g. auto_reply save → set → verify → restore → verify ----------------

await test("g. auto_reply (save state → set test message → restore exactly)", async () => {
  const before = await callGraphServer("/me/mailboxSettings?$select=automaticRepliesSetting");
  state.savedAutoReply = before?.automaticRepliesSetting;
  assert(state.savedAutoReply, "Could not read current automaticRepliesSetting");

  const setText = toolText(
    await autoReplyHandler({ action: "set", message: `${TEST_PREFIX} auto-reply test` }),
    "auto_reply set"
  );
  assert(setText.includes("Auto-reply enabled"), `Unexpected set output: ${setText}`);

  const getText = toolText(await autoReplyHandler({ action: "get" }), "auto_reply get");
  assert(
    getText.includes(`${TEST_PREFIX} auto-reply test`),
    `Set message not visible in get: ${getText}`
  );

  // Restore the exact saved object (not just "clear"), then verify.
  await callGraphServer("/me/mailboxSettings", {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ automaticRepliesSetting: state.savedAutoReply }),
  });
  const after = await callGraphServer("/me/mailboxSettings?$select=automaticRepliesSetting");
  const restored = after?.automaticRepliesSetting;
  assert(restored?.status === state.savedAutoReply.status, "Auto-reply status not restored");
  assert(
    (restored?.internalReplyMessage ?? "") === (state.savedAutoReply.internalReplyMessage ?? ""),
    "Internal reply message not restored"
  );
  assert(
    (restored?.externalReplyMessage ?? "") === (state.savedAutoReply.externalReplyMessage ?? ""),
    "External reply message not restored"
  );
});

// ---- v3a. search_mail get_latest mode ------------------------------------

await test("v3a. search_mail get_latest (no query → newest-first, matches direct $orderby)", async () => {
  assert(latestMessage, "Inbox is empty; nothing to list");
  const text = toolText(await searchMailHandler({ max_results: 5 }), "search_mail no-query");
  assert(text.includes("(newest first)"), `Output not labeled newest-first: ${text.slice(0, 200)}`);
  const ids = [...text.matchAll(/Message id: (\S+)/g)].map((m) => m[1]!);
  assert(ids.length >= 1, `No message ids in output: ${text.slice(0, 300)}`);

  // The first result must match a direct $orderby=receivedDateTime desc&$top=1 call.
  const direct = await callGraphServer(
    `/me/mailFolders/inbox/messages?$orderby=${encodeURIComponent("receivedDateTime desc")}&$top=1&$select=id,receivedDateTime`
  );
  assert(direct?.value?.[0]?.id === ids[0], "First result does not match direct $orderby call");

  if (ids.length >= 2) {
    const [first, second] = await Promise.all(
      ids.slice(0, 2).map((id) =>
        callGraphServer(`/me/messages/${encodeURIComponent(id)}?$select=receivedDateTime`)
      )
    );
    assert(
      new Date(first.receivedDateTime).getTime() >= new Date(second.receivedDateTime).getTime(),
      `Results not newest-first: ${first.receivedDateTime} < ${second.receivedDateTime}`
    );
  }
});

// ---- v3b. manage_message via $batch --------------------------------------

await test("v3b. manage_message $batch (3 drafts mark_read in ONE batch request)", async () => {
  const draftIds: string[] = [];
  for (let i = 1; i <= 3; i++) {
    const createText = toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject: `${TEST_PREFIX} batch ${i}`,
        body: "Batch test draft. Safe to delete.",
      }),
      `create_draft batch ${i}`
    );
    draftIds.push(extractDraftId(createText));
  }

  const logStart = graphRequestLog.length;
  const text = toolText(
    await manageMessageHandler({ message_ids: draftIds, action: "mark_read" }),
    "manage_message batch mark_read"
  );
  const issued = graphRequestLog.slice(logStart);
  const batchCalls = issued.filter((r) => r.path === "/$batch");
  assert(
    batchCalls.length === 1,
    `Expected exactly one /$batch request, saw ${batchCalls.length} (all requests: ${JSON.stringify(issued)})`
  );
  assert(
    issued.length === batchCalls.length,
    `manage_message issued non-batch requests: ${JSON.stringify(issued)}`
  );
  assert(text.includes("3/3"), `Expected 3/3 succeeded: ${text}`);
  for (const id of draftIds) {
    assert(text.includes(`OK      ${id}`), `Missing OK line for ${id}: ${text}`);
  }

  const delText = toolText(
    await manageMessageHandler({ message_ids: draftIds, action: "delete" }),
    "manage_message batch delete"
  );
  assert(delText.includes("3/3"), `Expected all three deletions to succeed: ${delText}`);
});

// ---- v3c. add_attachment small file --------------------------------------

await test("v3c. add_attachment small (~100 KB, single POST, visible in read_message)", async () => {
  const dir = await tempDir();
  const filePath = path.join(dir, "mcp-test-small.txt");
  await fs.writeFile(filePath, "The quick brown fox jumps over the lazy dog. ".repeat(2300)); // ~101 KB

  const createText = toolText(
    await createDraftHandler({
      to: [ownAddress],
      subject: `${TEST_PREFIX} attach`,
      body: "Attachment test draft. Safe to delete.",
    }),
    "create_draft attach"
  );
  const draftId = extractDraftId(createText);
  try {
    const attachText = toolText(
      await addAttachmentHandler({ draft_id: draftId, file_path: filePath }),
      "add_attachment small"
    );
    assert(attachText.includes("mcp-test-small.txt"), `Missing name: ${attachText}`);
    assert(attachText.includes(`${TEST_PREFIX} attach`), `Missing draft subject: ${attachText}`);

    const readText = toolText(await readMessageHandler({ message_id: draftId }), "read_message");
    assert(/Attachments \(1\):/.test(readText), `Inventory missing: ${readText.slice(-500)}`);
    assert(readText.includes("mcp-test-small.txt"), `Attachment name missing: ${readText.slice(-500)}`);
    assert(/1\d{2}\.\d KB/.test(readText), `Expected ~100 KB size in inventory: ${readText.slice(-500)}`);
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
    await fs.rm(filePath, { force: true });
  }
});

// ---- v3d. add_attachment large file (upload session) ----------------------

await test("v3d. add_attachment large (~4 MB via upload session, bytes verified)", async () => {
  const dir = await tempDir();
  const filePath = path.join(dir, "mcp-test-large.bin");
  // ~4.5 MB of deterministic bytes: exercises the session path (>3 MB) with two chunks (>4 MB).
  const size = Math.round(4.5 * 1024 * 1024);
  const buffer = Buffer.alloc(size);
  for (let i = 0; i < size; i++) buffer[i] = i % 251;
  await fs.writeFile(filePath, buffer);
  const localHash = createHash("sha256").update(buffer).digest("hex");

  const createText = toolText(
    await createDraftHandler({
      to: [ownAddress],
      subject: `${TEST_PREFIX} attach large`,
      body: "Large-attachment test draft. Safe to delete.",
    }),
    "create_draft attach large"
  );
  const draftId = extractDraftId(createText);
  try {
    const attachText = toolText(
      await addAttachmentHandler({ draft_id: draftId, file_path: filePath }),
      "add_attachment large"
    );
    assert(attachText.includes("mcp-test-large.bin"), `Missing name: ${attachText}`);
    assert(attachText.includes("4.5 MB"), `Missing size: ${attachText}`);

    // Verify the uploaded bytes really match the file: fetch and hash them.
    const listed = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name,size`
    );
    const att = (listed?.value ?? []).find((a: any) => a.name === "mcp-test-large.bin");
    assert(att, `Uploaded attachment not listed: ${JSON.stringify(listed?.value)}`);
    const full = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments/${encodeURIComponent(att.id)}`
    );
    const uploaded = Buffer.from(full.contentBytes ?? "", "base64");
    assert(
      uploaded.length === size,
      `Uploaded size ${uploaded.length} does not match file size ${size}`
    );
    const uploadedHash = createHash("sha256").update(uploaded).digest("hex");
    assert(uploadedHash === localHash, "Uploaded bytes differ from the local file (sha256 mismatch)");
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
    await fs.rm(filePath, { force: true });
  }
});

// ---- v3e. add_attachment guards ------------------------------------------

await test("v3e. add_attachment guards (non-draft, missing file, oversize)", async () => {
  const dir = await tempDir();

  // Non-draft target: any received inbox message.
  const tinyPath = path.join(dir, "mcp-test-tiny.txt");
  await fs.writeFile(tinyPath, "tiny");
  assert(latestMessage, "Inbox is empty; no non-draft id available");
  const nonDraft = expectError(
    await addAttachmentHandler({ draft_id: latestMessage.id, file_path: tinyPath }),
    "add_attachment(non-draft)"
  );
  assert(/not a draft/i.test(nonDraft), `Unexpected error text: ${nonDraft}`);

  // Missing file.
  const missing = expectError(
    await addAttachmentHandler({
      draft_id: latestMessage.id,
      file_path: path.join(dir, "does-not-exist.txt"),
    }),
    "add_attachment(missing file)"
  );
  assert(/not found|unreadable/i.test(missing), `Unexpected error text: ${missing}`);

  // Oversize: a sparse 26 MB file — truthful stat.size without writing 26 MB of data.
  const bigPath = path.join(dir, "mcp-test-oversize.bin");
  await fs.writeFile(bigPath, "");
  await fs.truncate(bigPath, 26 * 1024 * 1024);
  try {
    const oversize = expectError(
      await addAttachmentHandler({ draft_id: latestMessage.id, file_path: bigPath }),
      "add_attachment(oversize)"
    );
    assert(/25 MB/.test(oversize), `Unexpected error text: ${oversize}`);
  } finally {
    await fs.rm(bigPath, { force: true });
    await fs.rm(tinyPath, { force: true });
  }
});

// ---- v3f. manage_rules lifecycle -----------------------------------------

await test("v3f. manage_rules lifecycle (create → list shows summary → delete → gone)", async () => {
  const folder = await callGraphServer("/me/mailFolders", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: `${TEST_PREFIX} rules folder` }),
  });
  state.testFolderId = folder.id;
  let ruleId: string | undefined;
  try {
    const createText = toolText(
      await manageRulesHandler({
        action: "create",
        display_name: `${TEST_PREFIX} rule`,
        conditions: { from_addresses: ["mcptest@example.invalid"] },
        actions: { move_to_folder: folder.id },
      }),
      "manage_rules create"
    );
    ruleId = createText.match(/Rule id: (\S+)/)?.[1];
    assert(ruleId, `Could not extract rule id from output: ${createText}`);

    const listText = toolText(await manageRulesHandler({ action: "list" }), "manage_rules list");
    assert(listText.includes(`${TEST_PREFIX} rule`), `List missing the rule: ${listText}`);
    assert(listText.includes(ruleId), `List missing the rule id: ${listText}`);
    assert(
      listText.includes("from mcptest@example.invalid"),
      `Summary missing the condition: ${listText}`
    );
    assert(
      listText.includes(`move to "${TEST_PREFIX} rules folder"`),
      `Summary missing the action: ${listText}`
    );

    const deleteText = toolText(
      await manageRulesHandler({ action: "delete", rule_id: ruleId }),
      "manage_rules delete"
    );
    assert(deleteText.includes("deleted"), `Unexpected delete output: ${deleteText}`);
    ruleId = undefined;

    const afterText = toolText(await manageRulesHandler({ action: "list" }), "manage_rules list after");
    assert(
      !afterText.includes(`${TEST_PREFIX} rule`),
      `Deleted rule still listed: ${afterText}`
    );
  } finally {
    if (ruleId) {
      await callGraphServer(`/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`, {
        method: "DELETE",
      }).catch(() => {});
    }
    await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/permanentDelete`, {
      method: "POST",
    });
    state.testFolderId = undefined;
  }
});

// ---- v4a. create_folder ---------------------------------------------------

await test("v4a. create_folder (root folder → duplicate rejected → subfolder → bad parent)", async () => {
  const name = `${TEST_PREFIX} v4 folder`;
  const createText = toolText(await createFolderHandler({ name }), "create_folder");
  const id = createText.match(/Folder id: (\S+)/)?.[1];
  assert(id, `Could not extract folder id from output: ${createText}`);
  state.testFolderId = id;
  try {
    const folder = await callGraphServer(
      `/me/mailFolders/${encodeURIComponent(id)}?$select=displayName,childFolderCount`
    );
    assert(folder.displayName === name, `Folder name is ${folder.displayName}, expected ${name}`);

    // Same level, different case: rejected, and the error names the existing id.
    const dupText = expectError(
      await createFolderHandler({ name: name.toUpperCase() }),
      "create_folder(duplicate)"
    );
    assert(/already exists/i.test(dupText), `Unexpected duplicate error: ${dupText}`);
    assert(dupText.includes(id), `Duplicate error does not name the existing folder id: ${dupText}`);

    // A subfolder of the new folder.
    const childName = `${TEST_PREFIX} v4 child`;
    const childText = toolText(
      await createFolderHandler({ name: childName, parent_folder: id }),
      "create_folder(subfolder)"
    );
    const childId = childText.match(/Folder id: (\S+)/)?.[1];
    assert(childId, `Could not extract subfolder id from output: ${childText}`);
    assert(childText.includes(name), `Subfolder output does not name its parent: ${childText}`);
    const child = await callGraphServer(
      `/me/mailFolders/${encodeURIComponent(childId)}?$select=displayName,parentFolderId`
    );
    assert(
      child.parentFolderId === id,
      `Subfolder parent is ${child.parentFolderId}, expected ${id}`
    );
    // The same name is free again one level down.
    const nestedText = toolText(
      await createFolderHandler({ name, parent_folder: childId }),
      "create_folder(same name, deeper level)"
    );
    assert(nestedText.includes("Folder created"), `Unexpected nested output: ${nestedText}`);

    const badParent = expectError(
      await createFolderHandler({ name: `${TEST_PREFIX} orphan`, parent_folder: "no-such-folder" }),
      "create_folder(bad parent)"
    );
    assert(/does not exist/i.test(badParent), `Unexpected bad-parent error: ${badParent}`);
  } finally {
    // Removing the parent takes the whole test subtree with it (test-only purge).
    await callGraphServer(`/me/mailFolders/${encodeURIComponent(id)}/permanentDelete`, {
      method: "POST",
    });
    state.testFolderId = undefined;
  }
  await expect404(`/me/mailFolders/${encodeURIComponent(id)}?$select=id`, "Test folder");
});

// ---- v4b. manage_rules update in place ------------------------------------

await test("v4b. manage_rules update (patch conditions + exceptions in place, same rule id)", async () => {
  const folder = await callGraphServer("/me/mailFolders", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: `${TEST_PREFIX} v4 rules folder` }),
  });
  state.testFolderId = folder.id;
  let ruleId: string | undefined;
  try {
    const createText = toolText(
      await manageRulesHandler({
        action: "create",
        display_name: `${TEST_PREFIX} v4 rule`,
        conditions: { from_addresses: ["mcpv4@example.invalid"] },
        actions: { move_to_folder: folder.id },
      }),
      "manage_rules create"
    );
    ruleId = createText.match(/Rule id: (\S+)/)?.[1];
    assert(ruleId, `Could not extract rule id from output: ${createText}`);
    const before = await callGraphServer(
      `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`
    );

    const updateText = toolText(
      await manageRulesHandler({
        action: "update",
        rule_id: ruleId,
        display_name: `${TEST_PREFIX} v4 rule updated`,
        conditions: { subject_contains: ["mcpv4subject"] },
        exceptions: { sender_contains: ["mcpv4exempt"] },
      }),
      "manage_rules update"
    );
    assert(/updated in place/i.test(updateText), `Unexpected update output: ${updateText}`);
    assert(updateText.includes(ruleId), `Update output lost the rule id: ${updateText}`);

    // Patched in place: same id and position; conditions replaced, exceptions
    // added, and the action left alone because the call did not mention it.
    const after = await callGraphServer(
      `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`
    );
    assert(after.id === ruleId, `Rule id changed: ${before.id} → ${after.id}`);
    assert(
      after.sequence === before.sequence,
      `Rule sequence changed: ${before.sequence} → ${after.sequence}`
    );
    assert(
      after.displayName === `${TEST_PREFIX} v4 rule updated`,
      `Rule not renamed: ${after.displayName}`
    );
    assert(
      !after.conditions?.fromAddresses?.length,
      `Old condition survived the replace: ${JSON.stringify(after.conditions)}`
    );
    assert(
      after.conditions?.subjectContains?.length === 1,
      `New condition missing: ${JSON.stringify(after.conditions)}`
    );
    assert(
      after.exceptions?.senderContains?.length === 1,
      `Exception missing: ${JSON.stringify(after.exceptions)}`
    );
    assert(
      after.actions?.moveToFolder === folder.id,
      `Untouched action was clobbered: ${JSON.stringify(after.actions)}`
    );

    // list surfaces the exceptions in the human-readable summary.
    const listText = toolText(await manageRulesHandler({ action: "list" }), "manage_rules list");
    assert(
      listText.includes(`${TEST_PREFIX} v4 rule updated`),
      `List missing the renamed rule: ${listText}`
    );
    assert(/EXCEPT sender contains mcpv4exempt/i.test(listText), `List missing the exception: ${listText}`);
    assert(/subject contains mcpv4subject/i.test(listText), `List missing the condition: ${listText}`);

    // An update may also clear the exceptions and toggle the rule off.
    const clearText = toolText(
      await manageRulesHandler({ action: "update", rule_id: ruleId, exceptions: {}, enabled: false }),
      "manage_rules update(clear exceptions)"
    );
    assert(clearText.includes("DISABLED"), `Rule not reported disabled: ${clearText}`);
    const cleared = await callGraphServer(
      `/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`
    );
    assert(
      !cleared.exceptions?.senderContains?.length,
      `Exceptions not cleared: ${JSON.stringify(cleared.exceptions)}`
    );
    assert(cleared.isEnabled === false, "Rule not disabled");

    // Guards: no rule_id, nothing to change, conditions emptied.
    const noId = expectError(
      await manageRulesHandler({ action: "update", display_name: "x" }),
      "manage_rules update(no rule_id)"
    );
    assert(/requires rule_id/i.test(noId), `Unexpected error: ${noId}`);
    const noFields = expectError(
      await manageRulesHandler({ action: "update", rule_id: ruleId }),
      "manage_rules update(no fields)"
    );
    assert(/at least one of/i.test(noFields), `Unexpected error: ${noFields}`);
    const emptyConditions = expectError(
      await manageRulesHandler({ action: "update", rule_id: ruleId, conditions: {} }),
      "manage_rules update(empty conditions)"
    );
    assert(/at least one condition/i.test(emptyConditions), `Unexpected error: ${emptyConditions}`);

    const deleteText = toolText(
      await manageRulesHandler({ action: "delete", rule_id: ruleId }),
      "manage_rules delete"
    );
    assert(deleteText.includes("deleted"), `Unexpected delete output: ${deleteText}`);
    ruleId = undefined;
  } finally {
    if (ruleId) {
      await callGraphServer(`/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`, {
        method: "DELETE",
      }).catch(() => {});
    }
    await callGraphServer(`/me/mailFolders/${encodeURIComponent(folder.id)}/permanentDelete`, {
      method: "POST",
    });
    state.testFolderId = undefined;
  }
});

// ---- v4c. categories + manage_message categorize --------------------------

await test("v4c. category lifecycle (create → categorize a message → verify → clear → delete)", async () => {
  const categoryName = `${TEST_PREFIX} category`;
  const createText = toolText(
    await manageCategoriesHandler({
      action: "create",
      display_name: categoryName,
      color: "preset5",
    }),
    "manage_categories create"
  );
  const categoryId = createText.match(/Category id: (\S+)/)?.[1];
  assert(categoryId, `Could not extract category id from output: ${createText}`);
  assert(createText.includes("teal"), `Colour not reported: ${createText}`);

  let draftId: string | undefined;
  try {
    const listText = toolText(
      await manageCategoriesHandler({ action: "list" }),
      "manage_categories list"
    );
    assert(listText.includes(categoryName), `List missing the category: ${listText}`);
    assert(listText.includes(categoryId), `List missing the category id: ${listText}`);

    const dupText = expectError(
      await manageCategoriesHandler({
        action: "create",
        display_name: categoryName.toUpperCase(),
        color: "preset6",
      }),
      "manage_categories create(duplicate)"
    );
    assert(dupText.includes(categoryId), `Duplicate error does not name the existing id: ${dupText}`);

    // Categorize a throwaway draft: names are matched case-insensitively but
    // written back with the master list's exact spelling.
    draftId = extractDraftId(
      toolText(
        await createDraftHandler({
          to: [ownAddress],
          subject: `${TEST_PREFIX} categorize`,
          body: "Category test draft. Safe to delete.",
        }),
        "create_draft categorize"
      )
    );

    const setText = toolText(
      await manageMessageHandler({
        message_ids: [draftId],
        action: "categorize",
        categories: [categoryName.toLowerCase()],
      }),
      "manage_message categorize"
    );
    assert(setText.includes("1/1"), `Expected 1/1 to succeed: ${setText}`);
    assert(setText.includes("categories set to"), `Unexpected categorize output: ${setText}`);
    let msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}?$select=categories`
    );
    assert(
      JSON.stringify(msg.categories) === JSON.stringify([categoryName]),
      `Message categories are ${JSON.stringify(msg.categories)}, expected ["${categoryName}"]`
    );

    const unknown = expectError(
      await manageMessageHandler({
        message_ids: [draftId],
        action: "categorize",
        categories: [`${TEST_PREFIX} nonexistent category`],
      }),
      "manage_message categorize(unknown category)"
    );
    assert(/Unknown categor/i.test(unknown), `Unexpected error: ${unknown}`);
    const missing = expectError(
      await manageMessageHandler({ message_ids: [draftId], action: "categorize" }),
      "manage_message categorize(no categories)"
    );
    assert(/requires categories/i.test(missing), `Unexpected error: ${missing}`);

    // Uncategorize is a replace with the empty list.
    const clearText = toolText(
      await manageMessageHandler({
        message_ids: [draftId],
        action: "categorize",
        categories: [],
      }),
      "manage_message categorize(clear)"
    );
    assert(clearText.includes("categories cleared"), `Unexpected clear output: ${clearText}`);
    msg = await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}?$select=categories`);
    assert(
      (msg.categories ?? []).length === 0,
      `Categories not cleared: ${JSON.stringify(msg.categories)}`
    );

    const deleteText = toolText(
      await manageCategoriesHandler({ action: "delete", category_id: categoryId }),
      "manage_categories delete"
    );
    assert(deleteText.includes(categoryName), `Delete output does not name the category: ${deleteText}`);
    const afterText = toolText(
      await manageCategoriesHandler({ action: "list" }),
      "manage_categories list after delete"
    );
    assert(!afterText.includes(categoryName), `Deleted category still listed: ${afterText}`);
    const goneText = expectError(
      await manageCategoriesHandler({ action: "delete", category_id: categoryId }),
      "manage_categories delete(gone)"
    );
    assert(/No category with id/i.test(goneText), `Unexpected error: ${goneText}`);
  } finally {
    if (draftId) {
      await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, {
        method: "DELETE",
      }).catch(() => {});
    }
    await purgeTestCategories();
  }
});

// ---- v4d. To Do task lifecycle -------------------------------------------

await test("v4d. task lifecycle (create from a message → complete → reopen → update → delete)", async () => {
  assert(latestMessage, "Inbox is empty; no message to turn into a task");
  const title = `${TEST_PREFIX} task`;
  const today = torontoToday();
  const dueDate = addDays(today, 2);
  const reminder = `${addDays(today, 1)}T09:00`;
  const taskList = await resolveTaskList(undefined);

  const createText = toolText(
    await manageTaskHandler({
      action: "create",
      title,
      due_date: dueDate,
      body: "Created by the v4 test harness.",
      reminder,
      linked_message_id: latestMessage.id,
    }),
    "manage_task create"
  );
  let taskId = createText.match(/Task id: (\S+)/)?.[1];
  assert(taskId, `Could not extract task id from output: ${createText}`);
  const taskPath = `/me/todo/lists/${encodeURIComponent(taskList.id)}/tasks/${encodeURIComponent(taskId)}`;
  try {
    assert(
      createText.includes("Linked email details"),
      `Create output does not mention the linked email: ${createText}`
    );

    // The linked mail really made it into the task notes, alongside the user's own.
    const raw = await callGraphServer(taskPath);
    const notes = String(raw.body?.content ?? "");
    assert(notes.includes("Created by the v4 test harness."), `Task notes lost the body: ${notes}`);
    assert(
      notes.includes(String(latestMessage.subject ?? "")),
      `Task notes missing the linked subject: ${notes}`
    );
    assert(
      /Open in Outlook: https:\/\//.test(notes),
      `Task notes missing the message web link: ${notes}`
    );

    const listText = toolText(await listTasksHandler({}), "list_tasks");
    assert(listText.includes(title), `list_tasks missing the task: ${listText}`);
    assert(listText.includes(taskId), `list_tasks missing the task id: ${listText}`);
    assert(listText.includes("Upcoming"), `Task not grouped as upcoming: ${listText}`);
    assert(listText.includes(`due ${dueDate}`), `Due date not shown: ${listText}`);
    assert(listText.includes(`reminder ${reminder.replace("T", " ")}`), `Reminder not shown: ${listText}`);

    // due_within_days narrower than the due date drops it.
    const narrowText = toolText(
      await listTasksHandler({ due_within_days: 0 }),
      "list_tasks due_within_days 0"
    );
    assert(
      !narrowText.includes(taskId),
      `due_within_days 0 should exclude a task due ${dueDate}: ${narrowText}`
    );

    const completeText = toolText(
      await manageTaskHandler({ action: "complete", task_id: taskId }),
      "manage_task complete"
    );
    assert(/Status: completed/.test(completeText), `Unexpected complete output: ${completeText}`);
    const openText = toolText(await listTasksHandler({}), "list_tasks after complete");
    assert(!openText.includes(taskId), `Completed task still listed as open: ${openText}`);
    const allText = toolText(
      await listTasksHandler({ include_completed: true }),
      "list_tasks include_completed"
    );
    assert(allText.includes(taskId), `include_completed did not surface the task: ${allText}`);

    const reopenText = toolText(
      await manageTaskHandler({ action: "reopen", task_id: taskId }),
      "manage_task reopen"
    );
    assert(/Status: notStarted/.test(reopenText), `Unexpected reopen output: ${reopenText}`);
    const reopened = toolText(await listTasksHandler({}), "list_tasks after reopen");
    assert(reopened.includes(taskId), `Reopened task is not listed as open: ${reopened}`);

    const updateText = toolText(
      await manageTaskHandler({ action: "update", task_id: taskId, title: `${title} updated` }),
      "manage_task update"
    );
    assert(updateText.includes(`${title} updated`), `Title not updated: ${updateText}`);

    const noId = expectError(
      await manageTaskHandler({ action: "complete" }),
      "manage_task complete(no task_id)"
    );
    assert(/requires task_id/i.test(noId), `Unexpected error: ${noId}`);

    // Delete is permanent — the output has to say so, and the task is really gone.
    const deleteText = toolText(
      await manageTaskHandler({ action: "delete", task_id: taskId }),
      "manage_task delete"
    );
    assert(/permanent/i.test(deleteText), `Delete output does not flag permanence: ${deleteText}`);
    assert(deleteText.includes(`${title} updated`), `Delete output does not name the task: ${deleteText}`);
    const goneText = expectError(
      await manageTaskHandler({ action: "complete", task_id: taskId }),
      "manage_task complete(deleted task)"
    );
    assert(/No task/i.test(goneText), `Unexpected error for a deleted task: ${goneText}`);
    taskId = undefined;
  } finally {
    if (taskId) await callGraphServer(taskPath, { method: "DELETE" }).catch(() => {});
  }
});

// ---- v5a. check_new_mail delta lifecycle ---------------------------------

await test("v5a. check_new_mail delta (baseline → send-to-self → sees exactly it → advances)", async () => {
  const subject = `${TEST_PREFIX} v5 delta`;

  const baseline = toolText(
    await checkNewMailHandler({ folder: "inbox", reset: true }),
    "check_new_mail baseline"
  );
  assert(
    /Starting position recorded for inbox/.test(baseline),
    `Unexpected baseline output: ${baseline}`
  );
  assert(
    !baseline.includes(TEST_PREFIX),
    `Baseline call should list nothing, but named a test artifact: ${baseline}`
  );

  // Immediately after a baseline there is nothing of ours to report. Real mail
  // may legitimately arrive mid-run, so only our own artifacts are asserted on.
  const quiet = toolText(await checkNewMailHandler({}), "check_new_mail quiet");
  assert(!quiet.includes(TEST_PREFIX), `Unexpected test artifact before sending: ${quiet}`);

  const draftId = extractDraftId(
    toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject,
        body: "Delta-query probe from the v5 test harness. Safe to delete.",
      }),
      "create_draft (delta probe)"
    )
  );
  toolText(await sendDraftHandler({ draft_id: draftId }), "send_draft (delta probe)");

  const arrived = await poll("the delta probe message to reach the inbox", 90000, async () => {
    const found = await callGraphServer(
      `/me/mailFolders/inbox/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$select=id,subject`
    );
    return found?.value?.[0];
  });

  // Delta lags the folder listing slightly; poll it rather than assume.
  const changes = await poll("the delta query to report the probe message", 90000, async () => {
    const text = toolText(await checkNewMailHandler({}), "check_new_mail after send");
    return text.includes(subject) ? text : undefined;
  });
  const ourLines = changes.split("\n").filter((line) => line.includes(TEST_PREFIX));
  assert(
    ourLines.length === 1,
    `Expected exactly one test message in the delta, got ${ourLines.length}:\n${changes}`
  );
  assert(
    changes.includes(arrived.id),
    `Delta output does not carry the arrived message id:\n${changes}`
  );

  // The position must have advanced: the same change is never reported twice.
  const after = toolText(await checkNewMailHandler({}), "check_new_mail after reporting");
  assert(
    !after.includes(subject),
    `The probe message was reported a second time — the delta position did not advance:\n${after}`
  );

  // reset discards the stored position and says so.
  const reset = toolText(
    await checkNewMailHandler({ reset: true }),
    "check_new_mail reset"
  );
  assert(
    /Previous position discarded/.test(reset),
    `reset did not report discarding the stored position: ${reset}`
  );

  // Housekeeping for the sweep: the sent copy and the received copy both carry
  // the test prefix, so purgeTestMessages removes them.
});

// ---- v5b. change-notification validation handshake -----------------------

await test("v5b. webhook handshake (validationToken echoed verbatim, wrong method refused)", async () => {
  const store = createMemoryStateStore("remote");
  const token = "Validation: Testing unicode ✅ and spaces";
  const response = await handleNotificationRequest(
    new Request(`https://example.invalid/notifications?validationToken=${encodeURIComponent(token)}`, {
      method: "POST",
      body: "",
    }),
    { store }
  );
  assert(response.status === 200, `handshake answered HTTP ${response.status}, expected 200`);
  const contentType = response.headers.get("content-type") ?? "";
  assert(
    contentType.startsWith("text/plain"),
    `handshake content-type is ${contentType}, Graph requires text/plain`
  );
  const body = await response.text();
  assert(body === token, `handshake echoed ${JSON.stringify(body)}, expected ${JSON.stringify(token)}`);

  // The handshake must not need any stored state — a brand-new store is enough,
  // which is what makes creating the very first subscription possible.
  assert(
    (await store.get(STATE_SUBSCRIPTION)) === null,
    "the handshake wrote state it should not need"
  );

  const wrongMethod = await handleNotificationRequest(
    new Request("https://example.invalid/notifications", { method: "GET" }),
    { store }
  );
  assert(wrongMethod.status === 405, `GET without a token returned ${wrongMethod.status}, expected 405`);
});

// ---- v5c. notification ingest, clientState, ring buffer ------------------

await test("v5c. notification ingest (clientState enforced, enriched, ring capped at 50)", async () => {
  const store = createMemoryStateStore("remote");
  const clientState = "secret-client-state-for-the-test";
  await writeJson(store, STATE_SUBSCRIPTION, {
    id: "sub-1",
    clientState,
    expirationDateTime: new Date(Date.now() + 3600_000).toISOString(),
    notificationUrl: "https://example.invalid/notifications",
    resource: SUBSCRIPTION_RESOURCE,
    createdAt: new Date().toISOString(),
  } satisfies SubscriptionRecord);

  const deliver = (items: unknown[]) =>
    handleNotificationRequest(
      new Request("https://example.invalid/notifications", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ value: items }),
      }),
      {
        store,
        enrich: async (id) => ({
          subject: `${TEST_PREFIX} notified ${id}`,
          from: "someone@example.invalid",
          receivedDateTime: "2026-08-18T12:00:00Z",
        }),
      }
    );

  const forged = await deliver([
    { subscriptionId: "sub-1", clientState: "not-the-secret", changeType: "created", resourceData: { id: "forged" } },
  ]);
  assert(forged.status === 202, `forged delivery answered ${forged.status}, expected 202`);
  assert(
    JSON.stringify(await forged.json()) === JSON.stringify({ accepted: 0, discarded: 1 }),
    "a delivery with the wrong clientState was not discarded"
  );
  assert(
    (await readActivity(store)).length === 0,
    "a delivery with the wrong clientState reached the ring buffer"
  );

  const genuine = await deliver([
    { subscriptionId: "sub-1", clientState, changeType: "created", resourceData: { id: "msg-1" } },
  ]);
  assert(genuine.status === 202, `genuine delivery answered ${genuine.status}`);
  const entries = await readActivity(store);
  assert(entries.length === 1, `expected 1 buffered entry, got ${entries.length}`);
  assert(entries[0]!.messageId === "msg-1", `wrong message id: ${entries[0]!.messageId}`);
  assert(
    entries[0]!.subject === `${TEST_PREFIX} notified msg-1`,
    `entry was not enriched: ${JSON.stringify(entries[0])}`
  );
  assert(Date.parse(entries[0]!.at) > 0, `entry has no usable timestamp: ${entries[0]!.at}`);

  // get_mailbox_activity reads that buffer through the installed store.
  const activity = toolText(
    await runWithStateStore(store, () => getMailboxActivityHandler({ since_hours: 1 })),
    "get_mailbox_activity"
  );
  assert(activity.includes("msg-1"), `activity output missing the entry: ${activity}`);
  assert(
    activity.includes(`${TEST_PREFIX} notified msg-1`),
    `activity output missing the subject: ${activity}`
  );

  // Anything older than the window is filtered out rather than shown.
  await writeJson(store, STATE_ACTIVITY, [
    { at: new Date(Date.now() - 48 * 3600_000).toISOString(), changeType: "created", messageId: "old" },
  ]);
  const stale = toolText(
    await runWithStateStore(store, () => getMailboxActivityHandler({ since_hours: 6 })),
    "get_mailbox_activity (stale)"
  );
  assert(/No mail notifications/.test(stale), `stale entry was reported as recent: ${stale}`);

  // The ring buffer is capped and newest-first.
  await writeJson(store, STATE_ACTIVITY, []);
  for (let i = 0; i < ACTIVITY_CAP + 10; i++) {
    await appendActivity(store, [
      { at: new Date(Date.now() + i).toISOString(), changeType: "created", messageId: `m${i}` },
    ]);
  }
  const ring = await readActivity(store);
  assert(ring.length === ACTIVITY_CAP, `ring buffer holds ${ring.length}, expected ${ACTIVITY_CAP}`);
  assert(
    ring[0]!.messageId === `m${ACTIVITY_CAP + 9}`,
    `newest entry is ${ring[0]!.messageId}, expected the last appended`
  );
  assert(
    !ring.some((entry) => entry.messageId === "m0"),
    "the oldest entry was not dropped when the buffer filled"
  );
});

// ---- v5d. subscription renewal (the cron handler's logic) ----------------

/**
 * An in-memory stand-in for Graph's /subscriptions surface. Mirrors the two
 * behaviours the renewal logic depends on: subscriptions persist across calls,
 * and list responses never reveal the clientState secret (verified live —
 * Graph returns it as null).
 */
function memorySubscriptionGraph() {
  const subs = new Map<string, any>();
  const calls: { method: string; path: string; body: any }[] = [];
  let nextId = 1;
  let patchFails = false;

  const graph = async (path: string, init?: RequestInit) => {
    const method = (init?.method ?? "GET").toUpperCase();
    const body = init?.body ? JSON.parse(String(init.body)) : undefined;
    calls.push({ method, path, body });
    if (method === "GET" && path === "/subscriptions") {
      return { value: [...subs.values()].map((sub) => ({ ...sub, clientState: null })) };
    }
    if (method === "POST" && path === "/subscriptions") {
      const id = `sub-${nextId++}`;
      subs.set(id, {
        id,
        resource: body.resource,
        notificationUrl: body.notificationUrl,
        expirationDateTime: body.expirationDateTime,
      });
      return { id, expirationDateTime: body.expirationDateTime };
    }
    const id = decodeURIComponent(path.split("/").pop()!);
    if (method === "PATCH") {
      if (patchFails || !subs.has(id)) {
        throw new Error("ErrorItemNotFound: subscription no longer exists");
      }
      subs.get(id)!.expirationDateTime = body.expirationDateTime;
      return { id, expirationDateTime: body.expirationDateTime };
    }
    if (method === "DELETE") {
      subs.delete(id);
      return undefined;
    }
    throw new Error(`unexpected ${method} ${path}`);
  };

  return { graph, subs, calls, setPatchFails: (v: boolean) => (patchFails = v) };
}

await test("v5d. cron renewal (create → keep → renew → recreate when Graph forgot it)", async () => {
  const store = createMemoryStateStore("remote");
  const url = "https://example.invalid/notifications";
  const { graph, subs, calls, setPatchFails } = memorySubscriptionGraph();

  const start = new Date("2026-08-18T00:00:00Z");
  const created = await ensureMailSubscription(store, url, { now: start, graph });
  assert(created.action === "created", `first run did ${created.action}, expected created`);
  // Graph is consulted before anything is created: list first, then POST.
  assert(
    calls.map((c) => c.method).join(",") === "GET,POST",
    `creation made ${calls.map((c) => c.method).join(",")}, expected GET,POST`
  );
  const posted = calls.find((c) => c.method === "POST")!.body;
  assert(posted.resource === SUBSCRIPTION_RESOURCE, `wrong resource: ${posted.resource}`);
  assert(posted.notificationUrl === url, `wrong notificationUrl: ${posted.notificationUrl}`);
  assert(posted.changeType === "created", `wrong changeType: ${posted.changeType}`);
  assert(
    typeof posted.clientState === "string" && posted.clientState.length >= 32,
    `clientState is not a long random secret: ${posted.clientState}`
  );
  assert(
    Date.parse(posted.expirationDateTime) - start.getTime() <= 4230 * 60_000,
    `requested expiry exceeds the Graph maximum: ${posted.expirationDateTime}`
  );
  const firstClientState: string = posted.clientState;

  // Healthy: no Graph traffic at all.
  calls.length = 0;
  const kept = await ensureMailSubscription(store, url, {
    now: new Date(start.getTime() + 3600_000),
    graph,
  });
  assert(kept.action === "kept", `healthy subscription was ${kept.action}, expected kept`);
  assert(calls.length === 0, `keeping the subscription made ${calls.length} Graph call(s)`);

  // Inside the renewal window: list, then PATCH the same id, same clientState.
  const nearExpiry = new Date(Date.parse(created.record.expirationDateTime) - 6 * 3600_000);
  const renewed = await ensureMailSubscription(store, url, { now: nearExpiry, graph });
  assert(renewed.action === "renewed", `near-expiry run did ${renewed.action}, expected renewed`);
  const patchCall = calls.find((call) => call.method === "PATCH");
  assert(patchCall, `renewal made no PATCH: ${JSON.stringify(calls.map((c) => c.method))}`);
  const strayCalls = calls.filter((call) => call.method !== "PATCH" && call.method !== "GET");
  assert(
    strayCalls.length === 0,
    `renewal made unexpected Graph calls: ${JSON.stringify(strayCalls.map((c) => c.method))}`
  );
  assert(
    patchCall.path.endsWith(created.record.id),
    `renewal patched ${patchCall.path}, expected ${created.record.id}`
  );
  assert(
    renewed.record.clientState === firstClientState,
    "renewal changed the clientState, which would silently break every delivery"
  );
  assert(
    Date.parse(renewed.record.expirationDateTime) > Date.parse(created.record.expirationDateTime),
    "renewal did not push the expiry out"
  );

  // Graph forgot the subscription entirely (gone from the list too): the next
  // upkeep re-creates it with a fresh secret, without attempting a PATCH.
  calls.length = 0;
  subs.clear();
  const stillNear = new Date(Date.parse(renewed.record.expirationDateTime) - 6 * 3600_000);
  const recreated = await ensureMailSubscription(store, url, { now: stillNear, graph });
  assert(recreated.action === "recreated", `expected recreated, got ${recreated.action}`);
  assert(
    calls.map((c) => c.method).join(",") === "GET,POST",
    `recreation made ${calls.map((c) => c.method).join(",")}, expected GET,POST`
  );
  assert(recreated.record.id !== created.record.id, "recreation reused the dead subscription id");
  assert(
    recreated.record.clientState !== firstClientState,
    "recreation reused the old clientState"
  );

  // The narrower race: still listed, but the PATCH itself fails (the
  // subscription died between the list and the renewal). Same recovery.
  calls.length = 0;
  setPatchFails(true);
  const nearAgain = new Date(Date.parse(recreated.record.expirationDateTime) - 6 * 3600_000);
  const patchRaced = await ensureMailSubscription(store, url, { now: nearAgain, graph });
  setPatchFails(false);
  assert(patchRaced.action === "recreated", `expected recreated, got ${patchRaced.action}`);
  assert(
    calls.some((call) => call.method === "PATCH") && calls.some((call) => call.method === "POST"),
    "recovery did not try PATCH first and then POST"
  );

  // A lapsed record, or one pointing at a different endpoint, must be rebuilt.
  const lapsed: SubscriptionRecord = {
    ...recreated.record,
    expirationDateTime: new Date(stillNear.getTime() - 60_000).toISOString(),
  };
  assert(
    renewalDecision(lapsed, url, stillNear) === "create",
    "a lapsed subscription was not scheduled for re-creation"
  );
  assert(
    renewalDecision(recreated.record, "https://elsewhere.invalid/notifications", stillNear) ===
      "create",
    "a moved notification URL did not force re-creation"
  );
  assert(renewalDecision(null, url, stillNear) === "create", "an absent record was not created");
});

// ---- v5f. concurrent upkeep is idempotent --------------------------------

await test("v5f. two racing ensure calls leave exactly one subscription (stale-KV race)", async () => {
  const url = "https://example.invalid/notifications";
  const { graph, subs } = memorySubscriptionGraph();
  const start = new Date("2026-08-18T00:00:00Z");

  // The production incident: upkeep B runs with a stale KV view (its store
  // never saw upkeep A's write) after A has already created the subscription.
  // Before the fix, B blindly created a duplicate; now it must consult Graph
  // and converge — B cannot reuse A's subscription (Graph withholds the
  // clientState), so it replaces it, but exactly one survives either way.
  const storeA = createMemoryStateStore("remote");
  const storeB = createMemoryStateStore("remote");
  const first = await ensureMailSubscription(storeA, url, { now: start, graph });
  assert(first.action === "created", `A did ${first.action}, expected created`);
  const second = await ensureMailSubscription(storeB, url, { now: start, graph });
  assert(
    subs.size === 1,
    `after the stale-KV race Graph holds ${subs.size} subscriptions, expected exactly 1`
  );
  assert(
    subs.has(second.record.id),
    "the surviving subscription is not the one the last KV write describes — " +
      "its deliveries could never be validated"
  );

  // Truly simultaneous bootstrap: both list before either creates, so both
  // create — and the next upkeep of either sweeps the loser back down to one.
  subs.clear();
  const storeC = createMemoryStateStore("remote");
  const storeD = createMemoryStateStore("remote");
  const [c, d] = await Promise.all([
    ensureMailSubscription(storeC, url, { now: start, graph }),
    ensureMailSubscription(storeD, url, { now: start, graph }),
  ]);
  assert(c.action === "created" && d.action === "created", "both bootstraps should create");
  const nearExpiry = new Date(Date.parse(d.record.expirationDateTime) - 6 * 3600_000);
  const converged = await ensureMailSubscription(storeD, url, { now: nearExpiry, graph });
  assert(converged.action === "renewed", `convergence run did ${converged.action}`);
  assert(
    subs.size === 1 && subs.has(d.record.id),
    `the renewal sweep left ${subs.size} subscription(s), expected only ${d.record.id}`
  );
});

// ---- v5e. get_mailbox_activity is remote-only ----------------------------

await test("v5e. get_mailbox_activity refuses to pretend on the local stdio server", async () => {
  // The harness has the file-backed (local) store installed, exactly as the
  // stdio server does.
  const text = expectError(
    await getMailboxActivityHandler({}),
    "get_mailbox_activity (local mode)"
  );
  assert(/remote/i.test(text), `Error does not explain the remote-only limitation: ${text}`);
  assert(/check_new_mail/.test(text), `Error does not point at the local alternative: ${text}`);

  // With a remote-mode store but no subscription yet, it says so rather than
  // reporting an empty mailbox.
  const noSubscription = expectError(
    await runWithStateStore(createMemoryStateStore("remote"), () => getMailboxActivityHandler({})),
    "get_mailbox_activity (no subscription)"
  );
  assert(
    /no change-notification subscription/i.test(noSubscription),
    `Unexpected error with no subscription: ${noSubscription}`
  );
});

// ---- v7a. add_attachment from an https URL --------------------------------

// A small, public, always-on https resource this project owns: the deployed
// Worker's health endpoint. Attaching it exercises the url source end to end
// (fetch, size cap, content-type from the response, name from the URL path).
const PROBE_URL = "https://outlook-mcp.arthur-yuhao-zhang.workers.dev/health";

await test("v7a. add_attachment url (https fetch, name and type from the response, guards)", async () => {
  const expected = await fetch(PROBE_URL);
  assert(expected.ok, `the probe URL is not serving: HTTP ${expected.status}`);
  const expectedBody = await expected.text();

  const draftId = extractDraftId(
    toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject: `${TEST_PREFIX} attach url`,
        body: "URL-attachment test draft. Safe to delete.",
      }),
      "create_draft attach url"
    )
  );
  try {
    const attachText = toolText(
      await addAttachmentHandler({ draft_id: draftId, url: PROBE_URL }),
      "add_attachment url"
    );
    assert(/Name: health/.test(attachText), `Name not taken from the URL path: ${attachText}`);
    assert(
      /Type: application\/json/.test(attachText),
      `Content type not taken from the response: ${attachText}`
    );

    const listed = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name,contentType`
    );
    const att = (listed?.value ?? []).find((a: any) => a.name === "health");
    assert(att, `URL attachment not in the inventory: ${JSON.stringify(listed?.value)}`);
    const full = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments/${encodeURIComponent(att.id)}`
    );
    assert(
      Buffer.from(full.contentBytes ?? "", "base64").toString("utf8") === expectedBody,
      "the attached bytes differ from what the URL served"
    );

    // attachment_name overrides the URL-derived name, and drives the type.
    const named = toolText(
      await addAttachmentHandler({
        draft_id: draftId,
        url: PROBE_URL,
        attachment_name: "health-check.json",
      }),
      "add_attachment url (named)"
    );
    assert(/Name: health-check\.json/.test(named), `attachment_name ignored: ${named}`);

    // Guards: scheme, unparseable URL, and a URL that does not resolve.
    const plaintext = expectError(
      await addAttachmentHandler({ draft_id: draftId, url: "http://example.com/x.txt" }),
      "add_attachment(http url)"
    );
    assert(/https/i.test(plaintext), `Unexpected scheme error: ${plaintext}`);
    const nonsense = expectError(
      await addAttachmentHandler({ draft_id: draftId, url: "not a url" }),
      "add_attachment(bad url)"
    );
    assert(/not a valid url/i.test(nonsense), `Unexpected parse error: ${nonsense}`);
    const unreachable = expectError(
      await addAttachmentHandler({
        draft_id: draftId,
        url: "https://mcp-test.invalid/nothing.txt",
      }),
      "add_attachment(unreachable url)"
    );
    assert(
      /download|fetch|failed/i.test(unreachable),
      `Unexpected fetch failure text: ${unreachable}`
    );
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
  }
});

// ---- v7b. add_attachment from inline base64 ------------------------------

await test("v7b. add_attachment content_base64 (round-trip, source-count and size guards)", async () => {
  const payload = Buffer.from("id,name\n1,[MCP TEST] inline attachment\n", "utf8");

  const draftId = extractDraftId(
    toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject: `${TEST_PREFIX} attach base64`,
        body: "Inline-attachment test draft. Safe to delete.",
      }),
      "create_draft attach base64"
    )
  );
  try {
    const attachText = toolText(
      await addAttachmentHandler({
        draft_id: draftId,
        content_base64: payload.toString("base64"),
        attachment_name: "inline.csv",
      }),
      "add_attachment content_base64"
    );
    assert(/Type: text\/csv/.test(attachText), `Type not derived from the name: ${attachText}`);

    const listed = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name`
    );
    const att = (listed?.value ?? []).find((a: any) => a.name === "inline.csv");
    assert(att, `Inline attachment not in the inventory: ${JSON.stringify(listed?.value)}`);
    const full = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments/${encodeURIComponent(att.id)}`
    );
    assert(
      Buffer.from(full.contentBytes ?? "", "base64").equals(payload),
      "the attached bytes differ from the base64 that was supplied"
    );

    // Guards: no source, two sources, malformed base64, and over the 3 MB cap.
    const none = expectError(
      await addAttachmentHandler({ draft_id: draftId }),
      "add_attachment(no source)"
    );
    assert(/exactly one source/i.test(none), `Unexpected no-source error: ${none}`);
    const both = expectError(
      await addAttachmentHandler({
        draft_id: draftId,
        url: PROBE_URL,
        content_base64: payload.toString("base64"),
      }),
      "add_attachment(two sources)"
    );
    assert(/exactly one source/i.test(both), `Unexpected two-source error: ${both}`);
    const malformed = expectError(
      await addAttachmentHandler({ draft_id: draftId, content_base64: "not base64!!" }),
      "add_attachment(bad base64)"
    );
    assert(/not valid base64/i.test(malformed), `Unexpected base64 error: ${malformed}`);
    const oversize = expectError(
      await addAttachmentHandler({
        draft_id: draftId,
        content_base64: Buffer.alloc(3 * 1024 * 1024 + 16).toString("base64"),
        attachment_name: "big.bin",
      }),
      "add_attachment(oversize base64)"
    );
    assert(/3 MB/.test(oversize), `Unexpected oversize error: ${oversize}`);

    // file_path is refused when the server has no filesystem to read from.
    const remoteRefusal = expectError(
      await runWithStateStore(createMemoryStateStore("remote"), () =>
        addAttachmentHandler({ draft_id: draftId, file_path: "/etc/hosts" })
      ),
      "add_attachment(file_path on a remote server)"
    );
    assert(
      /hosted server has no access|file_path works only on the local/i.test(remoteRefusal),
      `Unexpected remote file_path error: ${remoteRefusal}`
    );
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
  }
});

// ---- v7c. get_attachment on both transports -------------------------------

await test("v7c. get_attachment (local file save; remote inline text, TTL link, link expiry)", async () => {
  const text = Buffer.from("[MCP TEST] attachment body\nline two\n", "utf8");
  const binary = Buffer.alloc(2048);
  for (let i = 0; i < binary.length; i++) binary[i] = (i * 7) % 256;

  const draftId = extractDraftId(
    toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject: `${TEST_PREFIX} get attach`,
        body: "get_attachment test draft. Safe to delete.",
      }),
      "create_draft get attach"
    )
  );
  const saveDir = path.join(os.homedir(), "Downloads", "outlook-mcp-attachments");
  const saveDirExisted = await fs
    .access(saveDir)
    .then(() => true)
    .catch(() => false);
  let savedPath: string | undefined;
  try {
    toolText(
      await addAttachmentHandler({
        draft_id: draftId,
        content_base64: text.toString("base64"),
        attachment_name: "note.txt",
      }),
      "add_attachment note.txt"
    );
    toolText(
      await addAttachmentHandler({
        draft_id: draftId,
        content_base64: binary.toString("base64"),
        attachment_name: "blob.bin",
      }),
      "add_attachment blob.bin"
    );
    const listed = await callGraphServer(
      `/me/messages/${encodeURIComponent(draftId)}/attachments?$select=id,name`
    );
    const idOf = (name: string) => {
      const found = (listed?.value ?? []).find((a: any) => a.name === name);
      assert(found, `${name} missing from the inventory`);
      return found.id as string;
    };

    // Local mode: saved to disk, with the text also inline.
    const localText = toolText(
      await getAttachmentHandler({ message_id: draftId, attachment_id: idOf("note.txt") }),
      "get_attachment (local)"
    );
    savedPath = localText.match(/Saved to: (.+)/)?.[1];
    assert(savedPath, `No save path in the local output: ${localText}`);
    assert(
      (await fs.readFile(savedPath, "utf8")) === text.toString("utf8"),
      "the saved file does not match the attachment"
    );
    assert(localText.includes("line two"), `Text attachment was not inlined: ${localText}`);

    // Remote mode: text inline with no filesystem, binary behind a link.
    const store = createMemoryStateStore("remote", "https://mcp-test.invalid");
    const remoteText = toolText(
      await runWithStateStore(store, () =>
        getAttachmentHandler({ message_id: draftId, attachment_id: idOf("note.txt") })
      ),
      "get_attachment (remote, text)"
    );
    assert(/Delivered inline/.test(remoteText), `Text was not delivered inline: ${remoteText}`);
    assert(!/Saved to:/.test(remoteText), `The hosted server claimed to save a file: ${remoteText}`);
    assert(remoteText.includes("line two"), `Text body missing: ${remoteText}`);

    const remoteBinary = toolText(
      await runWithStateStore(store, () =>
        getAttachmentHandler({
          message_id: draftId,
          attachment_id: idOf("blob.bin"),
          link_ttl_minutes: 1,
        })
      ),
      "get_attachment (remote, binary)"
    );
    const link = remoteBinary.match(
      new RegExp(
        `Download: https://mcp-test\\.invalid${DOWNLOAD_ROUTE_PREFIX}([0-9a-f]{64})`
      )
    )?.[1];
    assert(link, `No download link with an unguessable id: ${remoteBinary}`);
    assert(/Link expires:/.test(remoteBinary), `No expiry in the output: ${remoteBinary}`);

    const record = await readDownload(store, link);
    assert(record, "the download record was not parked in the state store");
    assert(
      Buffer.from(record.base64, "base64").equals(binary),
      "the parked bytes differ from the attachment"
    );
    assert(record.name === "blob.bin", `parked under the wrong name: ${record.name}`);

    // An unknown or malformed id is never served.
    assert((await readDownload(store, "f".repeat(64))) === null, "an unknown id was served");
    assert((await readDownload(store, "short")) === null, "a malformed id was served");

    // Expiry is enforced from inside the record, not just by the store's TTL:
    // backdate it and the bytes must be refused and dropped.
    await store.put(
      downloadKey(link),
      JSON.stringify({ ...record, expiresAt: new Date(Date.now() - 1000).toISOString() })
    );
    assert((await readDownload(store, link)) === null, "an expired link still served the bytes");
    assert(
      (await store.get(downloadKey(link))) === null,
      "the expired record was not dropped from the store"
    );
  } finally {
    if (savedPath) await fs.rm(savedPath, { force: true });
    if (!saveDirExisted) await fs.rm(saveDir, { recursive: true, force: true });
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}`, { method: "DELETE" });
  }
});

// ---- v7d. recurring event lifecycle --------------------------------------

await test("v7d. recurring events (weekly ×3 → move one occurrence → cancel the series)", async () => {
  const first = addDays(torontoToday(), 1);
  const createText = toolText(
    await createEventHandler({
      subject: `${TEST_PREFIX} weekly standup`,
      start: `${first}T09:00`,
      end: `${first}T09:30`,
      recurrence: { frequency: "weekly", count: 3 },
    }),
    "create_event (weekly series)"
  );
  const seriesId = createText.match(/Event id: (\S+)/)?.[1];
  assert(seriesId, `Could not extract the series id: ${createText}`);
  assert(/Repeats week\(s\) on \w+, 3 occurrence\(s\)/.test(createText), `Unexpected recurrence summary: ${createText}`);

  try {
    // calendarView expands the series into its three dates.
    const view = () =>
      callGraphServer(
        `/me/calendarView?startDateTime=${first}T00:00:00&endDateTime=${addDays(first, 30)}T00:00:00` +
          `&$select=id,subject,start,type&$orderby=start/dateTime&$top=50`,
        { headers: { Prefer: 'outlook.timezone="America/Toronto"' } }
      );
    const mine = (data: any) =>
      (data?.value ?? []).filter((e: any) => String(e.subject ?? "").startsWith(TEST_PREFIX));

    const occurrences = mine(await view());
    assert(occurrences.length === 3, `Expected 3 occurrences, got ${occurrences.length}`);
    assert(
      occurrences.every((o: any) => String(o.start.dateTime).slice(11, 16) === "09:00"),
      `Occurrences are not all at 09:00: ${occurrences.map((o: any) => o.start.dateTime)}`
    );

    // list_events must surface occurrence ids, flagged as part of a series.
    const listText = toolText(
      await listEventsHandler({ start_date: first, days: 2, include_ids: true }),
      "list_events include_ids"
    );
    assert(
      listText.includes(occurrences[0].id),
      `list_events did not print the occurrence id:\n${listText}`
    );
    assert(
      /one occurrence of a repeating event/.test(listText),
      `list_events did not flag the occurrence:\n${listText}`
    );

    // A repeat rule cannot be changed one date at a time.
    const refusal = expectError(
      await manageEventHandler({
        event_id: occurrences[1].id,
        action: "update",
        recurrence: { frequency: "daily" },
      }),
      "manage_event(recurrence on one occurrence)"
    );
    assert(/whole series/i.test(refusal), `Unexpected refusal text: ${refusal}`);

    // Move only the second occurrence: it becomes an exception, the others stay.
    const second = occurrences[1];
    const day = String(second.start.dateTime).slice(0, 10);
    const moved = toolText(
      await manageEventHandler({
        event_id: second.id,
        action: "update",
        start: `${day}T11:00`,
        end: `${day}T11:30`,
        scope: "this_event_only",
      }),
      "manage_event (move one occurrence)"
    );
    assert(/this occurrence only/.test(moved), `Scope not reported: ${moved}`);

    const afterMove = mine(await view());
    assert(afterMove.length === 3, `Expected 3 dates after the move, got ${afterMove.length}`);
    const times = afterMove.map((e: any) => String(e.start.dateTime).slice(11, 16));
    assert(
      JSON.stringify(times) === JSON.stringify(["09:00", "11:00", "09:00"]),
      `Only the second date should have moved, got ${JSON.stringify(times)}`
    );
    assert(
      afterMove[1].type === "exception",
      `The moved date is ${afterMove[1].type}, expected an exception`
    );

    // Series-wide edits are reachable from an occurrence id.
    const renamed = toolText(
      await manageEventHandler({
        event_id: afterMove[2].id,
        action: "update",
        subject: `${TEST_PREFIX} weekly standup (renamed)`,
        reminder_minutes: 20,
        scope: "entire_series",
      }),
      "manage_event (rename the series)"
    );
    assert(/the entire series/.test(renamed), `Series scope not reported: ${renamed}`);
    const master = await callGraphServer(
      `/me/events/${encodeURIComponent(seriesId)}?$select=subject,type,isReminderOn,reminderMinutesBeforeStart`
    );
    assert(
      master.subject === `${TEST_PREFIX} weekly standup (renamed)` && master.type === "seriesMaster",
      `The series master was not renamed: ${JSON.stringify(master)}`
    );
    assert(
      master.isReminderOn === true && master.reminderMinutesBeforeStart === 20,
      `Reminder not applied to the series: ${JSON.stringify(master)}`
    );

    // this_event_only on the series id is refused rather than guessed at.
    const wrongScope = expectError(
      await manageEventHandler({
        event_id: seriesId,
        action: "cancel",
        scope: "this_event_only",
      }),
      "manage_event(this_event_only on a series id)"
    );
    assert(/include_ids/.test(wrongScope), `Refusal does not say how to fix it: ${wrongScope}`);

    // Cancelling the series from one of its occurrences removes every date.
    const cancelled = toolText(
      await manageEventHandler({
        event_id: mine(await view())[0].id,
        action: "cancel",
        scope: "entire_series",
      }),
      "manage_event (cancel the series)"
    );
    assert(/entire series/.test(cancelled), `Cancel did not report the scope: ${cancelled}`);
    assert(mine(await view()).length === 0, "occurrences survived the series cancellation");
    await expect404(`/me/events/${encodeURIComponent(seriesId)}?$select=id`, "Series master");
  } catch (err) {
    await callGraphServer(`/me/events/${encodeURIComponent(seriesId)}`, { method: "DELETE" }).catch(
      () => {}
    );
    throw err;
  }
});

// ---- v7e. reminders, named calendars, list_calendars ----------------------

await test("v7e. calendars (list_calendars → create in a named calendar with a reminder → read back)", async () => {
  const calendarName = `${TEST_PREFIX} calendar`;
  const calendar = await callGraphServer("/me/calendars", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: calendarName }),
  });
  try {
    const calendarsText = toolText(await listCalendarsHandler({}), "list_calendars");
    assert(calendarsText.includes(calendarName), `New calendar not listed: ${calendarsText}`);
    assert(calendarsText.includes(calendar.id), `Calendar ids missing: ${calendarsText}`);
    assert(/— default/.test(calendarsText), `The default calendar is not marked: ${calendarsText}`);

    const day = addDays(torontoToday(), 2);
    const createText = toolText(
      await createEventHandler({
        subject: `${TEST_PREFIX} offsite`,
        start: `${day}T13:00`,
        end: `${day}T14:00`,
        calendar: calendarName,
        reminder_minutes: 45,
      }),
      "create_event (named calendar)"
    );
    assert(createText.includes(`Calendar: ${calendarName}`), `Calendar not reported: ${createText}`);
    assert(/Reminder: 45 min before/.test(createText), `Reminder not reported: ${createText}`);
    const eventId = createText.match(/Event id: (\S+)/)?.[1];
    assert(eventId, `Could not extract the event id: ${createText}`);

    const stored = await callGraphServer(
      `/me/events/${encodeURIComponent(eventId)}?$select=subject,isReminderOn,reminderMinutesBeforeStart`
    );
    assert(
      stored.isReminderOn === true && stored.reminderMinutesBeforeStart === 45,
      `Reminder not stored: ${JSON.stringify(stored)}`
    );

    // The event is in the named calendar and nowhere else.
    const inCalendar = toolText(
      await listEventsHandler({ start_date: day, days: 1, calendar: calendarName, include_ids: true }),
      "list_events (named calendar)"
    );
    assert(inCalendar.includes(`${TEST_PREFIX} offsite`), `Event missing from its calendar: ${inCalendar}`);
    assert(inCalendar.includes(eventId), `list_events did not print the event id: ${inCalendar}`);
    assert(inCalendar.includes(`in ${calendarName}`), `Calendar not named in the header: ${inCalendar}`);
    const inDefault = toolText(
      await listEventsHandler({ start_date: day, days: 1 }),
      "list_events (default calendar)"
    );
    assert(
      !inDefault.includes(`${TEST_PREFIX} offsite`),
      `The event leaked into the default calendar: ${inDefault}`
    );

    // Turning the reminder off, and an unknown calendar name reporting the real ones.
    const off = toolText(
      await manageEventHandler({ event_id: eventId, action: "update", reminder_minutes: -1 }),
      "manage_event (reminder off)"
    );
    assert(/Reminder: off/.test(off), `Reminder-off not reported: ${off}`);
    const afterOff = await callGraphServer(
      `/me/events/${encodeURIComponent(eventId)}?$select=isReminderOn`
    );
    assert(afterOff.isReminderOn === false, "the reminder was not turned off");

    const unknown = expectError(
      await listEventsHandler({ calendar: "[MCP TEST] no such calendar" }),
      "list_events(unknown calendar)"
    );
    assert(/Available calendars:/.test(unknown), `Unknown calendar error is unhelpful: ${unknown}`);

    toolText(
      await manageEventHandler({ event_id: eventId, action: "cancel" }),
      "manage_event (cancel offsite)"
    );
  } finally {
    await callGraphServer(`/me/calendars/${encodeURIComponent(calendar.id)}`, {
      method: "DELETE",
    }).catch(() => {});
  }
});

// ---- v8a. To Do depth: lists, recurrence, subtasks ------------------------

await test("v8a. task depth (create list → repeating task → subtasks add/complete/remove → rename list)", async () => {
  const listName = `${TEST_PREFIX} v8 list`;
  const renamedList = `${TEST_PREFIX} v8 list renamed`;
  const dueDate = addDays(torontoToday(), 3);

  const createListText = toolText(
    await manageTaskHandler({ action: "create_list", list_name: listName }),
    "manage_task create_list"
  );
  const listId = createListText.match(/List id: (\S+)/)?.[1];
  assert(listId, `no list id in the output: ${createListText}`);

  try {
    // A second list of the same name is refused rather than silently duplicated.
    const dup = expectError(
      await manageTaskHandler({ action: "create_list", list_name: listName }),
      "manage_task create_list (duplicate)"
    );
    assert(/already exists/i.test(dup), `Unexpected duplicate-list error: ${dup}`);

    const createText = toolText(
      await manageTaskHandler({
        action: "create",
        task_list: listName,
        title: `${TEST_PREFIX} weekly task`,
        due_date: dueDate,
        recurrence: { frequency: "weekly", weekdays: ["monday"] },
      }),
      "manage_task create (recurring)"
    );
    assert(/Repeats week\(s\) on monday/.test(createText), `Repeat not reported: ${createText}`);
    const taskId = createText.match(/Task id: (\S+)/)?.[1];
    assert(taskId, `no task id in the output: ${createText}`);

    // Graph's own view: the task really is in the new list and really repeats.
    const raw = await callGraphServer(
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`
    );
    assert(
      raw?.recurrence?.pattern?.type === "weekly",
      `Graph does not show a weekly recurrence: ${JSON.stringify(raw?.recurrence)}`
    );

    // A repeating task needs a due date, and recurrence cannot be changed later.
    const noDue = expectError(
      await manageTaskHandler({
        action: "create",
        task_list: listName,
        title: `${TEST_PREFIX} undated repeat`,
        recurrence: { frequency: "daily" },
      }),
      "manage_task create (recurrence without due_date)"
    );
    assert(/needs due_date/i.test(noDue), `Unexpected error: ${noDue}`);
    const changeRepeat = expectError(
      await manageTaskHandler({
        action: "update",
        task_list: listName,
        task_id: taskId,
        recurrence: { frequency: "daily" },
      }),
      "manage_task update (recurrence)"
    );
    assert(
      /only be set when the task is created/i.test(changeRepeat),
      `Unexpected error: ${changeRepeat}`
    );

    // Subtasks: add two, complete one by text, remove one by id.
    const addOne = toolText(
      await manageTaskHandler({
        action: "add_subtask",
        task_list: listName,
        task_id: taskId,
        subtask: "first step",
      }),
      "manage_task add_subtask"
    );
    const subtaskId = addOne.match(/Subtask id: (\S+)/)?.[1];
    assert(subtaskId, `no subtask id in the output: ${addOne}`);
    toolText(
      await manageTaskHandler({
        action: "add_subtask",
        task_list: listName,
        task_id: taskId,
        subtask: "second step",
      }),
      "manage_task add_subtask (2)"
    );

    const completed = toolText(
      await manageTaskHandler({
        action: "complete_subtask",
        task_list: listName,
        task_id: taskId,
        subtask: "first step",
      }),
      "manage_task complete_subtask"
    );
    assert(/Subtasks \(1\/2 done\)/.test(completed), `Checklist not rendered: ${completed}`);
    assert(/\[x\] first step/.test(completed), `Completed item not ticked: ${completed}`);

    const listed = toolText(
      await listTasksHandler({ task_list: listName, include_subtasks: true }),
      "list_tasks include_subtasks"
    );
    assert(/repeating/.test(listed), `list_tasks does not flag the repeat: ${listed}`);
    assert(/1\/2 subtasks done/.test(listed), `list_tasks does not count subtasks: ${listed}`);
    assert(
      listed.includes(`subtask id: ${subtaskId}`),
      `list_tasks does not carry subtask ids: ${listed}`
    );

    const removed = toolText(
      await manageTaskHandler({
        action: "remove_subtask",
        task_list: listName,
        task_id: taskId,
        subtask_id: subtaskId,
      }),
      "manage_task remove_subtask"
    );
    assert(/Subtasks \(0\/1 done\)/.test(removed), `Checklist after removal is wrong: ${removed}`);
    const remaining = await callGraphServer(
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}/checklistItems`
    );
    assert(
      (remaining?.value ?? []).length === 1 &&
        remaining.value[0].displayName === "second step",
      `Graph shows the wrong checklist: ${JSON.stringify(remaining?.value)}`
    );

    const unknownSubtask = expectError(
      await manageTaskHandler({
        action: "complete_subtask",
        task_list: listName,
        task_id: taskId,
        subtask: "no such step",
      }),
      "manage_task complete_subtask (unknown)"
    );
    assert(/No subtask reading/i.test(unknownSubtask), `Unexpected error: ${unknownSubtask}`);

    // Stopping the repeat is the one recurrence change To Do accepts.
    const stopped = toolText(
      await manageTaskHandler({
        action: "update",
        task_list: listName,
        task_id: taskId,
        clear_recurrence: true,
      }),
      "manage_task update (clear_recurrence)"
    );
    assert(/Repeat: off/.test(stopped), `Repeat-off not reported: ${stopped}`);
    const afterStop = await callGraphServer(
      `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`
    );
    assert(!afterStop?.recurrence, `the task still repeats: ${JSON.stringify(afterStop?.recurrence)}`);

    // Renaming keeps the same list (and its tasks).
    const renameText = toolText(
      await manageTaskHandler({
        action: "rename_list",
        task_list: listName,
        list_name: renamedList,
      }),
      "manage_task rename_list"
    );
    assert(renameText.includes(renamedList), `Rename output is wrong: ${renameText}`);
    const listAfter = await callGraphServer(`/me/todo/lists/${encodeURIComponent(listId)}`);
    assert(
      listAfter?.displayName === renamedList,
      `Graph still shows "${listAfter?.displayName}" after the rename`
    );

    toolText(
      await manageTaskHandler({ action: "delete", task_list: renamedList, task_id: taskId }),
      "manage_task delete"
    );
  } finally {
    // TEST-ONLY raw DELETE: manage_task deliberately cannot delete a list
    // (deleting one destroys its tasks irrecoverably), so the harness removes
    // its own list straight through Graph rather than through the tool surface.
    await callGraphServer(`/me/todo/lists/${encodeURIComponent(listId!)}`, {
      method: "DELETE",
    }).catch(() => {});
  }
});

// ---- v8b. junk: block and unblock a sender -------------------------------

await test("v8b. manage_senders (block the sender of a real message, then unblock it again)", async () => {
  assert(latestMessage, "Inbox is empty; no message whose sender could be blocked");
  const before = await callGraphServer(
    `/me/messages/${encodeURIComponent(latestMessage.id)}?$select=id,from,parentFolderId`
  );
  const senderAddress = before?.from?.emailAddress?.address;
  assert(senderAddress, "The newest inbox message has no sender address");

  // move_message: false throughout — the block/unblock pair is the thing under
  // test, and nothing in the mailbox should move because of a test run.
  let blocked = false;
  try {
    const blockText = toolText(
      await manageSendersHandler({
        action: "block_sender",
        message_id: latestMessage.id,
        move_message: false,
      }),
      "manage_senders block_sender"
    );
    blocked = true;
    assert(blockText.includes(senderAddress), `The block did not name the sender: ${blockText}`);
    assert(/left where it was/i.test(blockText), `The message was moved: ${blockText}`);
    assert(
      /cannot read the blocked-senders list back/i.test(blockText),
      `The output hides the platform limit: ${blockText}`
    );

    const stillThere = await callGraphServer(
      `/me/messages/${encodeURIComponent(latestMessage.id)}?$select=id,parentFolderId`
    );
    assert(
      stillThere.parentFolderId === before.parentFolderId,
      "block_sender moved the message even though move_message was false"
    );

    const unblockText = toolText(
      await manageSendersHandler({
        action: "unblock_sender",
        message_id: latestMessage.id,
        move_message: false,
      }),
      "manage_senders unblock_sender"
    );
    blocked = false;
    assert(
      unblockText.includes(senderAddress) && /delivered normally again/i.test(unblockText),
      `Unexpected unblock output: ${unblockText}`
    );

    // A draft has no sender, so there is nothing to block: say so, don't guess.
    const draftText = toolText(
      await createDraftHandler({
        to: [ownAddress],
        subject: `${TEST_PREFIX} junk guard`,
        body: "Sender-block guard probe. Safe to delete.",
      }),
      "create_draft (junk guard)"
    );
    const draftId = extractDraftId(draftText);
    try {
      const refused = expectError(
        await manageSendersHandler({ action: "block_sender", message_id: draftId }),
        "manage_senders block_sender (draft)"
      );
      assert(/no sender address/i.test(refused), `Unexpected draft refusal: ${refused}`);
    } finally {
      await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}/permanentDelete`, {
        method: "POST",
      }).catch(() => {});
    }
  } finally {
    // Never leave a real correspondent blocked because a test failed halfway.
    if (blocked) {
      await manageSendersHandler({
        action: "unblock_sender",
        message_id: latestMessage.id,
        move_message: false,
      });
    }
  }
});

// ---- v8c. focused inbox overrides ----------------------------------------

const FOCUS_TEST_SENDER = "mcp-test-focus@example.com";

await test("v8c. mailbox_settings focus override (set → visible in get → clear → gone)", async () => {
  const setText = toolText(
    await mailboxSettingsHandler({
      action: "set_focus_override",
      sender: FOCUS_TEST_SENDER,
      classify_as: "other",
    }),
    "mailbox_settings set_focus_override"
  );
  assert(/→ other/.test(setText), `Unexpected set output: ${setText}`);

  try {
    const getText = toolText(
      await mailboxSettingsHandler({ action: "get" }),
      "mailbox_settings get"
    );
    assert(
      getText.includes(`${FOCUS_TEST_SENDER} → other`),
      `The override is not in the settings read-out: ${getText}`
    );

    // Setting the same sender again updates the one override rather than
    // colliding with it (Graph refuses a duplicate).
    const again = toolText(
      await mailboxSettingsHandler({
        action: "set_focus_override",
        sender: FOCUS_TEST_SENDER,
        classify_as: "focused",
      }),
      "mailbox_settings set_focus_override (again)"
    );
    assert(/updated/i.test(again) && /was other/.test(again), `Unexpected re-set output: ${again}`);
    const overrides = await callGraphServer("/me/inferenceClassification/overrides?$top=100");
    const mine = (overrides?.value ?? []).filter(
      (o: any) => o.senderEmailAddress?.address === FOCUS_TEST_SENDER
    );
    assert(mine.length === 1, `Graph holds ${mine.length} overrides for the test sender`);
    assert(mine[0].classifyAs === "focused", `The override was not updated: ${mine[0].classifyAs}`);
  } finally {
    await mailboxSettingsHandler({
      action: "clear_focus_override",
      sender: FOCUS_TEST_SENDER,
    });
  }

  const gone = expectError(
    await mailboxSettingsHandler({ action: "clear_focus_override", sender: FOCUS_TEST_SENDER }),
    "mailbox_settings clear_focus_override (twice)"
  );
  assert(/no Focused-Inbox override/i.test(gone), `Unexpected error: ${gone}`);
});

// ---- v8d. working hours --------------------------------------------------

await test("v8d. mailbox_settings working hours (save → set → verify → restore exactly)", async () => {
  const before = await callGraphServer("/me/mailboxSettings?$select=workingHours");
  state.savedWorkingHours = before?.workingHours;
  assert(state.savedWorkingHours, "Could not read the current workingHours");

  const getText = toolText(await mailboxSettingsHandler({ action: "get" }), "mailbox_settings get");
  assert(/Working hours: /.test(getText), `Working hours missing from get: ${getText}`);

  try {
    const setText = toolText(
      await mailboxSettingsHandler({
        action: "set_working_hours",
        days: ["tuesday", "thursday"],
        start_time: "10:15",
        end_time: "16:45",
      }),
      "mailbox_settings set_working_hours"
    );
    assert(
      /After: +tuesday, thursday 10:15–16:45/.test(setText),
      `Unexpected set output: ${setText}`
    );
    assert(/visible to anyone who schedules/i.test(setText), `No visibility warning: ${setText}`);

    const applied = (await callGraphServer("/me/mailboxSettings?$select=workingHours"))
      ?.workingHours;
    assert(
      JSON.stringify(applied.daysOfWeek) === JSON.stringify(["tuesday", "thursday"]) &&
        applied.startTime.startsWith("10:15") &&
        applied.endTime.startsWith("16:45"),
      `Graph shows different working hours: ${JSON.stringify(applied)}`
    );

    // A day that ends before it starts is refused rather than sent to Graph.
    const backwards = expectError(
      await mailboxSettingsHandler({
        action: "set_working_hours",
        start_time: "18:00",
        end_time: "09:00",
      }),
      "mailbox_settings set_working_hours (backwards)"
    );
    assert(/must be later than/i.test(backwards), `Unexpected error: ${backwards}`);
    const empty = expectError(
      await mailboxSettingsHandler({ action: "set_working_hours" }),
      "mailbox_settings set_working_hours (nothing to change)"
    );
    assert(/at least one of days/i.test(empty), `Unexpected error: ${empty}`);
  } finally {
    await callGraphServer("/me/mailboxSettings", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ workingHours: state.savedWorkingHours }),
    });
  }

  const restored = (await callGraphServer("/me/mailboxSettings?$select=workingHours"))
    ?.workingHours;
  assert(
    JSON.stringify(restored) === JSON.stringify(state.savedWorkingHours),
    `Working hours not restored exactly: ${JSON.stringify(restored)}`
  );
});

// ---- v8e. message forensics: internet headers ----------------------------

await test("v8e. read_message include_headers (auth results, delivery chain, reply-to check)", async () => {
  assert(latestMessage, "Inbox is empty; no message whose headers could be read");
  const text = toolText(
    await readMessageHandler({
      message_id: latestMessage.id,
      include_headers: true,
      include_attachments_list: false,
    }),
    "read_message include_headers"
  );

  assert(/\nInternet headers:/.test(text), `No header section: ${text.slice(-400)}`);
  assert(
    /Authentication-Results: (SPF|DKIM|DMARC)/.test(text),
    `No Authentication-Results verdict — inbox mail should carry one: ${text.slice(-800)}`
  );
  assert(
    /Received chain \(\d+ hop\(s\), oldest first/.test(text),
    `No Received chain: ${text.slice(-800)}`
  );
  // Compact rendering, not raw headers: every hop is one "from … by …" line.
  const hops = text.match(/^ {4}\d+\. from .+ by /gm) ?? [];
  assert(hops.length >= 1, `The delivery chain is not compactly rendered: ${text.slice(-800)}`);
  assert(/ {2}Reply-To: /.test(text), `No reply-to verdict: ${text.slice(-800)}`);

  // Without the flag the answer stays as small as it was.
  const plain = toolText(
    await readMessageHandler({ message_id: latestMessage.id, include_attachments_list: false }),
    "read_message (no headers)"
  );
  assert(!/Internet headers:/.test(plain), "Headers leaked into the default read_message output");

  // A draft has no internet headers at all: say so rather than inventing any.
  const draftText = toolText(
    await createDraftHandler({
      to: [ownAddress],
      subject: `${TEST_PREFIX} header probe`,
      body: "Header-rendering probe. Safe to delete.",
    }),
    "create_draft (header probe)"
  );
  const draftId = extractDraftId(draftText);
  try {
    const draftHeaders = toolText(
      await readMessageHandler({
        message_id: draftId,
        include_headers: true,
        include_attachments_list: false,
      }),
      "read_message include_headers (draft)"
    );
    assert(
      /\(none — Microsoft Graph carries internet headers only/.test(draftHeaders),
      `A draft's empty header set is not explained: ${draftHeaders.slice(-400)}`
    );
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(draftId)}/permanentDelete`, {
      method: "POST",
    }).catch(() => {});
  }
});

// ---- v8f. EML export (local transport) -----------------------------------

await test("v8f. export_message (saves a .eml to disk that parses as RFC 822)", async () => {
  assert(latestMessage, "Inbox is empty; no message to export");
  const text = toolText(
    await exportMessageHandler({ message_id: latestMessage.id }),
    "export_message"
  );
  const savedPath = text.match(/Saved to: (.+)$/m)?.[1];
  assert(savedPath, `No saved path in the output: ${text}`);
  assert(savedPath.endsWith(".eml"), `The export is not a .eml: ${savedPath}`);
  assert(!/Download: /.test(text), `The local server handed out a link: ${text}`);
  state.exportedEmlPath = savedPath;

  const raw = await fs.readFile(savedPath, "utf8");
  // RFC 822 shape: a header block of "Name: value" lines, then a blank line,
  // then the body — and the headers Graph reported for the same message.
  const separator = raw.indexOf("\r\n\r\n");
  assert(separator > 0, "No header/body separator in the exported MIME");
  const headerBlock = raw.slice(0, separator);
  const headerNames = headerBlock
    .split("\r\n")
    .filter((line) => /^[!-9;-~]+:/.test(line))
    .map((line) => line.slice(0, line.indexOf(":")).toLowerCase());
  for (const required of ["from", "to", "subject", "date", "received"]) {
    assert(headerNames.includes(required), `The exported MIME has no ${required} header`);
  }
  assert(raw.length > separator + 4, "The exported MIME has no body");

  const graphView = await callGraphServer(
    `/me/messages/${encodeURIComponent(latestMessage.id)}?$select=internetMessageId`
  );
  assert(
    !graphView.internetMessageId || raw.includes(graphView.internetMessageId),
    "The exported MIME is not this message (Message-ID does not match)"
  );

  const missing = expectError(
    await exportMessageHandler({ message_id: "AAAAnotarealmessageid=" }),
    "export_message (bad id)"
  );
  assert(/No message /.test(missing), `Unexpected error: ${missing}`);

  await fs.rm(savedPath, { force: true });
  state.exportedEmlPath = undefined;
});

// ---- h. stdio protocol smoke test ---------------------------------------

await test("h. stdio smoke test (initialize + tools/prompts/resources lists, clean stdout)", async () => {
  const tsxBin = path.join(PROJECT_ROOT, "node_modules", ".bin", "tsx");
  const child = spawn(tsxBin, ["src/server.ts"], {
    cwd: PROJECT_ROOT,
    stdio: ["pipe", "pipe", "pipe"],
  });
  let stdout = "";
  child.stdout.on("data", (chunk) => (stdout += chunk));
  child.stderr.on("data", () => {}); // stderr is allowed to carry logs
  try {
    const send = (msg: object) => child.stdin.write(JSON.stringify(msg) + "\n");
    send({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2025-06-18",
        capabilities: {},
        clientInfo: { name: "test-harness", version: "0.0.0" },
      },
    });

    const waitForResponse = (id: number, timeoutMs = 15000) =>
      new Promise<any>((resolve, reject) => {
        const timer = setTimeout(
          () => reject(new Error(`Timed out waiting for response id ${id}`)),
          timeoutMs
        );
        const check = () => {
          for (const line of stdout.split("\n")) {
            if (!line.trim()) continue;
            try {
              const msg = JSON.parse(line);
              if (msg.id === id) {
                clearTimeout(timer);
                clearInterval(poller);
                resolve(msg);
                return;
              }
            } catch {
              // partial line still buffering
            }
          }
        };
        const poller = setInterval(check, 100);
      });

    const initResponse = await waitForResponse(1);
    assert(initResponse.result?.serverInfo?.name === "outlook", "Server did not identify as 'outlook'");
    send({ jsonrpc: "2.0", method: "notifications/initialized" });
    send({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
    const listResponse = await waitForResponse(2);

    const tools = listResponse.result?.tools ?? [];
    const names = tools.map((t: any) => t.name).sort();
    const expected = [
      "add_attachment",
      "auto_reply",
      "check_new_mail",
      "create_draft",
      "create_event",
      "create_folder",
      "export_message",
      "get_attachment",
      "get_mailbox_activity",
      "list_calendars",
      "list_events",
      "list_folders",
      "list_tasks",
      "mailbox_settings",
      "manage_categories",
      "manage_contact",
      "manage_event",
      "manage_message",
      "manage_rules",
      "manage_senders",
      "manage_task",
      "read_message",
      "read_thread",
      "search_contacts",
      "search_mail",
      "send_draft",
      "update_draft",
    ];
    assert(tools.length === 27, `Expected 27 tools, got ${tools.length}`);
    assert(
      JSON.stringify(names) === JSON.stringify(expected),
      `Expected tools ${expected.join(", ")}; got ${names.join(", ")}`
    );
    for (const tool of tools) {
      assert(tool.description?.length > 20, `Tool ${tool.name} lacks a description`);
      assert(
        tool.inputSchema?.type === "object" &&
          Object.keys(tool.inputSchema.properties ?? {}).length > 0,
        `Tool ${tool.name} lacks a valid object input schema`
      );
    }

    // Prompts must be advertised over the same connection.
    send({ jsonrpc: "2.0", id: 3, method: "prompts/list", params: {} });
    const promptsResponse = await waitForResponse(3);
    const prompts = promptsResponse.result?.prompts ?? [];
    const promptNames = prompts.map((p: any) => p.name).sort();
    assert(
      JSON.stringify(promptNames) === JSON.stringify(["morning_brief", "triage_inbox"]),
      `Expected prompts morning_brief, triage_inbox; got ${promptNames.join(", ") || "(none)"}`
    );
    for (const prompt of prompts) {
      assert(
        prompt.description?.length > 20,
        `Prompt ${prompt.name} lacks a description`
      );
    }

    // Each prompt renders to a non-trivial user message naming the tools to use.
    const promptChecks: [string, string[]][] = [
      ["triage_inbox", ["list_folders", "manage_rules", "search_mail"]],
      ["morning_brief", ["search_mail", "list_events", "list_tasks"]],
    ];
    for (const [index, [name, mustMention]] of promptChecks.entries()) {
      const requestId = 10 + index;
      send({ jsonrpc: "2.0", id: requestId, method: "prompts/get", params: { name } });
      const got = await waitForResponse(requestId);
      const messages = got.result?.messages ?? [];
      assert(messages.length >= 1, `Prompt ${name} returned no messages`);
      const text = messages.map((m: any) => m.content?.text ?? "").join("\n");
      assert(text.length > 200, `Prompt ${name} text is suspiciously short: ${text}`);
      for (const tool of mustMention) {
        assert(text.includes(tool), `Prompt ${name} does not mention ${tool}`);
      }
    }

    // Resources are advertised over the same connection and readable.
    send({ jsonrpc: "2.0", id: 20, method: "resources/list", params: {} });
    const resourcesResponse = await waitForResponse(20);
    const resources = resourcesResponse.result?.resources ?? [];
    const uris = resources.map((r: any) => r.uri).sort();
    assert(
      JSON.stringify(uris) === JSON.stringify([FOLDERS_URI, RECENT_INBOX_URI].sort()),
      `Expected resources ${FOLDERS_URI}, ${RECENT_INBOX_URI}; got ${uris.join(", ") || "(none)"}`
    );
    for (const resource of resources) {
      assert(
        resource.description?.length > 20,
        `Resource ${resource.uri} lacks a description`
      );
      assert(resource.mimeType === "text/plain", `Resource ${resource.uri} has no mime type`);
    }

    send({
      jsonrpc: "2.0",
      id: 21,
      method: "resources/read",
      params: { uri: FOLDERS_URI },
    });
    const readResponse = await waitForResponse(21);
    const contents = readResponse.result?.contents ?? [];
    assert(contents.length === 1, `resources/read returned ${contents.length} contents`);
    assert(
      contents[0].uri === FOLDERS_URI && typeof contents[0].text === "string",
      `Unexpected resource content: ${JSON.stringify(contents[0]).slice(0, 200)}`
    );
    assert(
      /Mail folders \(\d+ top-level\)/.test(contents[0].text),
      `Folder resource does not carry the folder tree: ${contents[0].text.slice(0, 200)}`
    );

    // stdout must contain nothing but protocol JSON.
    for (const line of stdout.split("\n")) {
      if (!line.trim()) continue;
      let parsed: any;
      try {
        parsed = JSON.parse(line);
      } catch {
        throw new Error(`Non-JSON content on stdout: ${JSON.stringify(line.slice(0, 200))}`);
      }
      assert(parsed.jsonrpc === "2.0", `Non-JSON-RPC message on stdout: ${line.slice(0, 200)}`);
    }
  } finally {
    child.kill();
  }
});

// ---- i. final artifact sweep --------------------------------------------

await test("i. final sweep: no [MCP TEST] artifacts anywhere, auto-reply restored", async () => {
  // Purge soft-deleted test artifacts (test-only), then verify nothing remains.
  await purgeTestMessages();
  await purgeTestFolders();
  await purgeTestRules();
  await purgeTestCategories();
  await purgeTestCalendars();
  await purgeTestTasks();
  await purgeTestTaskLists();
  await purgeTestFocusOverrides();

  const messages = await callGraphServer(
    `/me/messages?$filter=${encodeURIComponent(`startswith(subject,'${TEST_PREFIX}')`)}&$select=id,subject,parentFolderId`
  );
  assert(
    (messages?.value ?? []).length === 0,
    `Leftover test messages: ${JSON.stringify(messages.value)}`
  );

  const events = await callGraphServer(
    `/me/events?$filter=${encodeURIComponent(`startswith(subject,'${TEST_PREFIX}')`)}&$select=id,subject`
  );
  assert(
    (events?.value ?? []).length === 0,
    `Leftover test events: ${JSON.stringify(events.value)}`
  );

  // Every calendar, not just the default one: create_event can target any of them.
  const calendars = await callGraphServer("/me/calendars?$select=id,name&$top=100");
  const testCalendars = (calendars?.value ?? []).filter((c: any) =>
    String(c.name ?? "").startsWith(TEST_PREFIX)
  );
  assert(testCalendars.length === 0, `Leftover test calendars: ${JSON.stringify(testCalendars)}`);
  for (const calendar of calendars?.value ?? []) {
    const calendarEvents = await callGraphServer(
      `/me/calendars/${encodeURIComponent(calendar.id)}/events?$filter=${encodeURIComponent(
        `startswith(subject,'${TEST_PREFIX}')`
      )}&$select=id,subject&$top=100`
    );
    assert(
      (calendarEvents?.value ?? []).length === 0,
      `Leftover test events in "${calendar.name}": ${JSON.stringify(calendarEvents.value)}`
    );
  }

  const contacts = await callGraphServer(
    `/me/contacts?$filter=${encodeURIComponent(`startswith(displayName,'${TEST_PREFIX}')`)}&$select=id,displayName`
  );
  assert(
    (contacts?.value ?? []).length === 0,
    `Leftover test contacts: ${JSON.stringify(contacts.value)}`
  );

  const folders = await callGraphServer("/me/mailFolders?$top=100&$select=id,displayName");
  const testFolders = (folders?.value ?? []).filter((f: any) =>
    String(f.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(testFolders.length === 0, `Leftover test folders: ${JSON.stringify(testFolders)}`);

  const rules = await callGraphServer("/me/mailFolders/inbox/messageRules");
  const testRules = (rules?.value ?? []).filter((r: any) =>
    String(r.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(testRules.length === 0, `Leftover test rules: ${JSON.stringify(testRules)}`);

  // Subfolders too: create_folder can nest, so check one level below every
  // top-level folder as well as the root listing above.
  for (const parent of folders?.value ?? []) {
    if (!parent.id) continue;
    const children = await callGraphServer(
      `/me/mailFolders/${encodeURIComponent(parent.id)}/childFolders?$top=100&$select=id,displayName`
    );
    const testChildren = (children?.value ?? []).filter((f: any) =>
      String(f.displayName ?? "").startsWith(TEST_PREFIX)
    );
    assert(
      testChildren.length === 0,
      `Leftover test subfolders under ${parent.displayName}: ${JSON.stringify(testChildren)}`
    );
  }

  const categories = await callGraphServer("/me/outlook/masterCategories?$top=100");
  const testCategories = (categories?.value ?? []).filter((c: any) =>
    String(c.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(
    testCategories.length === 0,
    `Leftover test categories: ${JSON.stringify(testCategories)}`
  );

  const taskLists = await callGraphServer("/me/todo/lists?$top=100");
  const testLists = (taskLists?.value ?? []).filter((l: any) =>
    String(l.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(
    testLists.length === 0,
    `Leftover test To Do lists: ${JSON.stringify(testLists.map((l: any) => l.displayName))}`
  );
  for (const list of taskLists?.value ?? []) {
    const tasks = await callGraphServer(
      `/me/todo/lists/${encodeURIComponent(list.id)}/tasks?$top=100`
    );
    const testTasks = (tasks?.value ?? []).filter((t: any) =>
      String(t.title ?? "").startsWith(TEST_PREFIX)
    );
    assert(
      testTasks.length === 0,
      `Leftover test tasks in "${list.displayName}": ${JSON.stringify(testTasks.map((t: any) => t.title))}`
    );
  }

  // Delta state: the position really was written to the run's throwaway file
  // (and so nowhere else), and that file is now gone.
  const stateContents = await fs.readFile(TEST_STATE_FILE, "utf8").catch(() => "");
  assert(
    stateContents.includes(deltaKey("inbox")),
    `check_new_mail did not persist its position to ${TEST_STATE_FILE}`
  );
  await fs.rm(TEST_STATE_FILE, { force: true });
  const stateGone = await fs
    .access(TEST_STATE_FILE)
    .then(() => false)
    .catch(() => true);
  assert(stateGone, `Test state file still exists: ${TEST_STATE_FILE}`);

  // Temp files: remove the run's temp dir and verify it is gone.
  if (state.tempDir) {
    await fs.rm(state.tempDir, { recursive: true, force: true });
    const gone = await fs
      .access(state.tempDir)
      .then(() => false)
      .catch(() => true);
    assert(gone, `Temp dir still exists: ${state.tempDir}`);
    state.tempDir = undefined;
  }

  // Focused-Inbox overrides: the harness only ever creates its own test sender's.
  const overrides = await callGraphServer("/me/inferenceClassification/overrides?$top=100");
  const testOverrides = (overrides?.value ?? []).filter((o: any) =>
    String(o.senderEmailAddress?.address ?? "").startsWith("mcp-test-")
  );
  assert(
    testOverrides.length === 0,
    `Leftover test Focused-Inbox overrides: ${JSON.stringify(testOverrides)}`
  );

  // Exported .eml files: the export test removes its own, and nothing else in
  // the local download directory may carry a test artifact.
  if (state.exportedEmlPath) {
    await fs.rm(state.exportedEmlPath, { force: true });
    state.exportedEmlPath = undefined;
  }
  const downloads = await fs.readdir(SAVE_DIR).catch(() => [] as string[]);
  const testDownloads = downloads.filter((name) => name.startsWith(TEST_PREFIX));
  assert(
    testDownloads.length === 0,
    `Leftover test files in ${SAVE_DIR}: ${testDownloads.join(", ")}`
  );

  const settings = await callGraphServer(
    "/me/mailboxSettings?$select=automaticRepliesSetting,workingHours"
  );
  if (state.savedWorkingHours) {
    assert(
      JSON.stringify(settings?.workingHours) === JSON.stringify(state.savedWorkingHours),
      `Working hours were not restored exactly: ${JSON.stringify(settings?.workingHours)}`
    );
  }
  const current = settings?.automaticRepliesSetting;
  assert(
    !String(current?.internalReplyMessage ?? "").includes(TEST_PREFIX),
    "Auto-reply still carries the test message"
  );
  if (state.savedAutoReply) {
    assert(current?.status === state.savedAutoReply.status, "Auto-reply status differs from saved state");
  }
});

// ---- summary ------------------------------------------------------------

console.log("\n=== Tool test summary ===");
const width = Math.max(...outcomes.map((o) => o.name.length));
for (const o of outcomes) {
  console.log(`${o.passed ? "PASS" : "FAIL"}  ${o.name.padEnd(width)}`);
}
const failed = outcomes.filter((o) => !o.passed);
console.log(`\n${outcomes.length - failed.length}/${outcomes.length} tests passed.`);
if (failed.length) process.exitCode = 1;
