// Test harness: exercises the five tool handlers directly (bypassing the MCP
// transport) against the real account using the cached token, plus a stdio
// protocol smoke test of the server itself. Every test cleans up after itself;
// a final check verifies no "[MCP TEST]" artifacts remain in the account.
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { GraphError, callGraphServer } from "./graph.js";
import { searchMailHandler } from "./tools/search-mail.js";
import { readThreadHandler } from "./tools/read-thread.js";
import { createDraftHandler } from "./tools/create-draft.js";
import { listEventsHandler, torontoToday, addDays } from "./tools/list-events.js";
import { createEventHandler } from "./tools/create-event.js";
import type { ToolResult } from "./tools/common.js";

const PROJECT_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
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

async function expect404(pathToGet: string, what: string): Promise<void> {
  try {
    await callGraphServer(pathToGet);
    throw new Error(`${what} still exists after deletion`);
  } catch (err) {
    if (err instanceof GraphError && err.status === 404) return;
    throw err;
  }
}

/** Permanently remove any [MCP TEST] messages (including soft-deleted ones in Deleted Items). */
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

// ---- shared fixtures ----------------------------------------------------

const me = await callGraphServer("/me?$select=mail,userPrincipalName");
const ownAddress: string = me.mail ?? me.userPrincipalName;
console.log(`Running against account: ${ownAddress}\n`);

const latestInbox = await callGraphServer(
  "/me/mailFolders/inbox/messages?$top=1&$select=id,subject,conversationId"
);
const latestMessage = latestInbox?.value?.[0];

// ---- a. search_mail -----------------------------------------------------

await test("a. search_mail (term from latest inbox subject)", async () => {
  assert(latestMessage, "Inbox is empty; cannot derive a guaranteed search term");
  const subject: string = latestMessage.subject ?? "";
  const tokens = subject.match(/[A-Za-z0-9]{3,}/g) ?? [];
  const term = tokens.sort((a, b) => b.length - a.length)[0] ?? subject;
  assert(term, `Could not derive a search term from subject ${JSON.stringify(subject)}`);
  const text = toolText(await searchMailHandler({ query: term }), "search_mail");
  assert(/result\(s\)/.test(text), `Expected results for ${JSON.stringify(term)}, got: ${text}`);
  assert(text.includes("Conversation id:"), "Output missing conversation ids");
});

// ---- b. read_thread -----------------------------------------------------

await test("b. read_thread (latest inbox conversation)", async () => {
  assert(latestMessage?.conversationId, "No conversationId available from the latest inbox message");
  const text = toolText(
    await readThreadHandler({ conversation_id: latestMessage.conversationId }),
    "read_thread"
  );
  assert(text.startsWith("Thread:"), `Unexpected output: ${text.slice(0, 200)}`);
  assert(text.includes("From:") && text.includes("Date:"), "Output missing sender/date lines");
});

// ---- c. create_draft: create, verify in Drafts, delete, confirm ---------

await test("c. create_draft (new message → verify in Drafts → delete)", async () => {
  const subject = `${TEST_PREFIX} harness draft`;
  const text = toolText(
    await createDraftHandler({
      to: [ownAddress],
      subject,
      body: "Automated test draft created by the outlook-mcp test harness. Safe to delete.",
    }),
    "create_draft"
  );
  const id = text.match(/Message id: (\S+)/)?.[1];
  assert(id, `Could not extract message id from output: ${text}`);
  try {
    const draftsFolder = await callGraphServer("/me/mailFolders/drafts?$select=id");
    const msg = await callGraphServer(
      `/me/messages/${encodeURIComponent(id)}?$select=parentFolderId,subject`
    );
    assert(msg.parentFolderId === draftsFolder.id, "Draft is not in the Drafts folder");
    assert(msg.subject === subject, `Draft subject mismatch: ${msg.subject}`);
  } finally {
    await callGraphServer(`/me/messages/${encodeURIComponent(id)}`, { method: "DELETE" });
  }
  await expect404(`/me/messages/${encodeURIComponent(id)}?$select=id`, "Draft");
});

// ---- d. list_events -----------------------------------------------------

await test("d. list_events (next 7 days)", async () => {
  const text = toolText(await listEventsHandler({}), "list_events");
  assert(
    text.startsWith("Events ") || text.startsWith("No events in this window"),
    `Unexpected output: ${text.slice(0, 200)}`
  );
});

// ---- e. create_event: create, verify, delete, confirm -------------------

await test("e. create_event (tomorrow 09:00–09:30 → verify → delete)", async () => {
  const tomorrow = addDays(torontoToday(), 1);
  const text = toolText(
    await createEventHandler({
      subject: `${TEST_PREFIX} harness event`,
      start: `${tomorrow}T09:00`,
      end: `${tomorrow}T09:30`,
    }),
    "create_event"
  );
  const id = text.match(/Event id: (\S+)/)?.[1];
  assert(id, `Could not extract event id from output: ${text}`);
  try {
    const event = await callGraphServer(
      `/me/events/${encodeURIComponent(id)}?$select=subject,start,end`
    );
    assert(event.subject === `${TEST_PREFIX} harness event`, "Event subject mismatch");
    assert(
      String(event.start?.dateTime).startsWith(`${tomorrow}T09:00`) ||
        event.start?.timeZone !== "America/Toronto",
      `Unexpected event start: ${JSON.stringify(event.start)}`
    );
  } finally {
    await callGraphServer(`/me/events/${encodeURIComponent(id)}`, { method: "DELETE" });
  }
  await expect404(`/me/events/${encodeURIComponent(id)}?$select=id`, "Event");
});

// ---- f. stdio protocol smoke test ---------------------------------------

await test("f. stdio smoke test (initialize + tools/list, clean stdout)", async () => {
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
                clearInterval(poll);
                resolve(msg);
                return;
              }
            } catch {
              // partial line still buffering
            }
          }
        };
        const poll = setInterval(check, 100);
      });

    const initResponse = await waitForResponse(1);
    assert(initResponse.result?.serverInfo?.name === "outlook", "Server did not identify as 'outlook'");
    send({ jsonrpc: "2.0", method: "notifications/initialized" });
    send({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
    const listResponse = await waitForResponse(2);

    const tools = listResponse.result?.tools ?? [];
    const names = tools.map((t: any) => t.name).sort();
    const expected = ["create_draft", "create_event", "list_events", "read_thread", "search_mail"];
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

// ---- final: no leftover [MCP TEST] artifacts ----------------------------

await test("final. no leftover [MCP TEST] artifacts in the account", async () => {
  // Soft-deleted test messages land in Deleted Items; purge them so the
  // account is genuinely clean, then verify nothing remains anywhere.
  await purgeTestMessages();
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
