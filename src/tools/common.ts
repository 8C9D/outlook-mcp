import { AuthRequiredError } from "../auth.js";
import { GraphError, callGraphServer } from "../graph.js";

export const TIMEZONE = "America/Toronto";
export const TZ_PREFER = `outlook.timezone="${TIMEZONE}"`;

export type ToolResult = {
  content: { type: "text"; text: string }[];
  isError?: boolean;
};

export function textResult(text: string): ToolResult {
  return { content: [{ type: "text", text }] };
}

export function errorResult(text: string): ToolResult {
  return { content: [{ type: "text", text }], isError: true };
}

/**
 * A caller-fixable problem with the tool's input (an unknown name, a missing
 * prerequisite). runTool surfaces the message verbatim, without the "Tool failed"
 * prefix reserved for unexpected faults.
 */
export class ToolInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ToolInputError";
  }
}

/**
 * Run a tool handler, converting every failure into an isError tool result so
 * the server process never crashes on a tool call.
 */
export async function runTool(fn: () => Promise<ToolResult>): Promise<ToolResult> {
  try {
    return await fn();
  } catch (err) {
    if (err instanceof AuthRequiredError) return errorResult(err.message);
    if (err instanceof ToolInputError) return errorResult(err.message);
    if (err instanceof GraphError) {
      let detail = err.body;
      try {
        const parsed = JSON.parse(err.body);
        if (parsed?.error?.message) detail = `${parsed.error.code}: ${parsed.error.message}`;
      } catch {
        // keep the raw body
      }
      return errorResult(
        `Microsoft Graph error (HTTP ${err.status} on ${err.requestPath}): ${detail}`
      );
    }
    return errorResult(`Tool failed: ${err instanceof Error ? err.message : String(err)}`);
  }
}

/**
 * True when a Graph failure means "there is no such item". Graph answers a
 * well-formed but unknown id with 404, and an id it cannot even parse with
 * 400 ErrorInvalidIdMalformed; both mean the same thing to a caller.
 */
export function isNotFound(err: unknown): boolean {
  return (
    err instanceof GraphError &&
    (err.status === 404 || (err.status === 400 && /ErrorInvalidIdMalformed/i.test(err.body)))
  );
}

/** GET a Graph collection, following @odata.nextLink until `cap` items are collected. */
export async function fetchPaged(
  firstPath: string,
  cap: number,
  headers?: Record<string, string>
): Promise<any[]> {
  const items: any[] = [];
  let next: string | undefined = firstPath;
  while (next && items.length < cap) {
    const page = await callGraphServer(next, headers ? { headers } : undefined);
    items.push(...(page?.value ?? []));
    next = page?.["@odata.nextLink"];
  }
  return items.slice(0, cap);
}

/** Format any Graph datetime (offset or naive) as a local America/Toronto string. */
export function formatLocal(dateTime: string | undefined): string {
  if (!dateTime) return "(no date)";
  // Naive datetimes from Prefer: outlook.timezone are already local wall-clock.
  const naive = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(dateTime) && !/[Zz]|[+-]\d{2}:?\d{2}$/.test(dateTime);
  if (naive) return `${dateTime.slice(0, 10)} ${dateTime.slice(11, 16)}`;
  const d = new Date(dateTime);
  if (Number.isNaN(d.getTime())) return dateTime;
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(d);
  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? "";
  return `${get("year")}-${get("month")}-${get("day")} ${get("hour")}:${get("minute")}`;
}

export function formatSender(from: any): string {
  const name = from?.emailAddress?.name;
  const address = from?.emailAddress?.address ?? "(unknown)";
  return name && name !== address ? `${name} <${address}>` : address;
}

export function toRecipients(addresses: string[]): { emailAddress: { address: string } }[] {
  return addresses.map((address) => ({ emailAddress: { address } }));
}

/** Escape a value for use inside a Graph $filter string literal. */
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
