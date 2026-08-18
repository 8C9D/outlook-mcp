import { callGraph } from "./graph.js";

type CheckResult = { name: string; passed: boolean; detail?: string };
const results: CheckResult[] = [];

async function runCheck(name: string, fn: () => Promise<void>): Promise<void> {
  try {
    await fn();
    results.push({ name, passed: true });
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    console.error(`\n[${name}] FAILED: ${detail}`);
    results.push({ name, passed: false, detail });
  }
}

await runCheck("profile (/me)", async () => {
  const me = await callGraph("/me");
  console.log(`\nSigned in as: ${me.displayName} <${me.userPrincipalName}>`);
});

await runCheck("inbox (latest 3 messages)", async () => {
  const data = await callGraph(
    "/me/mailFolders/inbox/messages?$top=3&$select=subject,from,receivedDateTime"
  );
  console.log("\nLatest inbox messages:");
  for (const msg of data.value ?? []) {
    const sender = msg.from?.emailAddress?.address ?? "(unknown sender)";
    console.log(`  - "${msg.subject}" from ${sender} at ${msg.receivedDateTime}`);
  }
  if (!data.value?.length) console.log("  (inbox is empty)");
});

await runCheck("calendar (next 7 days)", async () => {
  const now = new Date();
  const weekOut = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  const query =
    `/me/calendarView?startDateTime=${encodeURIComponent(now.toISOString())}` +
    `&endDateTime=${encodeURIComponent(weekOut.toISOString())}` +
    `&$select=subject,start&$orderby=start/dateTime`;
  const data = await callGraph(query, {
    headers: { Prefer: 'outlook.timezone="America/Toronto"' },
  });
  console.log("\nCalendar events in the next 7 days:");
  if (!data.value?.length) {
    console.log("  no events");
    return;
  }
  for (const event of data.value) {
    console.log(`  - "${event.subject}" at ${event.start?.dateTime} (${event.start?.timeZone})`);
  }
});

console.log("\n=== Verification summary ===");
for (const r of results) {
  console.log(`${r.passed ? "PASS" : "FAIL"}  ${r.name}`);
}
if (results.some((r) => !r.passed)) process.exitCode = 1;
