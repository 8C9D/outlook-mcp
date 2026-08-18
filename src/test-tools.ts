// Test harness: exercises the tool handlers directly (bypassing the MCP
// transport) against the real account using the cached token, plus a stdio
// protocol smoke test of the server itself. Every test cleans up after itself;
// a final check verifies no "[MCP TEST]" artifacts remain in the account
// (messages, drafts, folders, events, contacts, inbox rules, temp files) and
// that mailbox settings (auto-reply) are restored exactly.
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { GraphError, callGraphServer, graphRequestLog } from "./graph.js";
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
import { createEventHandler } from "./tools/create-event.js";
import { manageEventHandler } from "./tools/manage-event.js";
import { searchContactsHandler } from "./tools/search-contacts.js";
import { manageContactHandler } from "./tools/manage-contact.js";
import { autoReplyHandler } from "./tools/auto-reply.js";
import { addAttachmentHandler } from "./tools/add-attachment.js";
import { manageRulesHandler } from "./tools/manage-rules.js";
import type { ToolResult } from "./tools/common.js";

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

// ---- h. stdio protocol smoke test ---------------------------------------

await test("h. stdio smoke test (initialize + full tools/list, clean stdout)", async () => {
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
      "create_draft",
      "create_event",
      "get_attachment",
      "list_events",
      "list_folders",
      "manage_contact",
      "manage_event",
      "manage_message",
      "manage_rules",
      "read_message",
      "read_thread",
      "search_contacts",
      "search_mail",
      "send_draft",
      "update_draft",
    ];
    assert(tools.length === 17, `Expected 17 tools, got ${tools.length}`);
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

  const contacts = await callGraphServer(
    `/me/contacts?$filter=${encodeURIComponent(`startswith(displayName,'${TEST_PREFIX}')`)}&$select=id,displayName`
  );
  assert(
    (contacts?.value ?? []).length === 0,
    `Leftover test contacts: ${JSON.stringify(contacts.value)}`
  );

  const folders = await callGraphServer("/me/mailFolders?$top=100&$select=displayName");
  const testFolders = (folders?.value ?? []).filter((f: any) =>
    String(f.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(testFolders.length === 0, `Leftover test folders: ${JSON.stringify(testFolders)}`);

  const rules = await callGraphServer("/me/mailFolders/inbox/messageRules");
  const testRules = (rules?.value ?? []).filter((r: any) =>
    String(r.displayName ?? "").startsWith(TEST_PREFIX)
  );
  assert(testRules.length === 0, `Leftover test rules: ${JSON.stringify(testRules)}`);

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

  const settings = await callGraphServer("/me/mailboxSettings?$select=automaticRepliesSetting");
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
