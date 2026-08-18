import { z } from "zod";
import {
  ToolResult,
  escapeODataString,
  fetchPaged,
  runTool,
  textResult,
} from "./common.js";

export const searchContactsSchema = {
  query: z
    .string()
    .min(1)
    .describe("Name prefix to search for — matched against display name, given name, and surname."),
  max_results: z
    .number()
    .int()
    .min(1)
    .max(25)
    .default(10)
    .describe("Maximum number of contacts to return (default 10, max 25)."),
};

const searchContactsArgs = z.object(searchContactsSchema);

export const searchContactsDescription =
  "Search the user's saved Outlook contacts by name prefix. Returns each match's name, email addresses, phone numbers, company, and contact id (for manage_contact). Searches saved contacts only, not the organization directory.";

export function formatContact(c: any, index?: number): string {
  const emails = (c.emailAddresses ?? [])
    .map((e: any) => e.address)
    .filter(Boolean)
    .join(", ");
  const phones = [...(c.businessPhones ?? []), ...(c.homePhones ?? []), c.mobilePhone]
    .filter(Boolean)
    .join(", ");
  return [
    `${index !== undefined ? `${index + 1}. ` : ""}${c.displayName || [c.givenName, c.surname].filter(Boolean).join(" ") || "(no name)"}`,
    `   Email: ${emails || "(none)"}`,
    `   Phone: ${phones || "(none)"}`,
    ...(c.companyName ? [`   Company: ${c.companyName}`] : []),
    `   Contact id: ${c.id}`,
  ].join("\n");
}

const SELECT =
  "id,displayName,givenName,surname,emailAddresses,businessPhones,homePhones,mobilePhone,companyName";

export async function searchContactsHandler(
  input: z.input<typeof searchContactsArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { query, max_results } = searchContactsArgs.parse(input);
    const q = escapeODataString(query);
    const filter = encodeURIComponent(
      `startswith(displayName,'${q}') or startswith(givenName,'${q}') or startswith(surname,'${q}')`
    );
    const contacts = await fetchPaged(
      `/me/contacts?$filter=${filter}&$select=${SELECT}&$top=${max_results}`,
      max_results
    );
    if (contacts.length === 0) {
      return textResult(`No contacts matching ${JSON.stringify(query)}.`);
    }
    return textResult(
      `${contacts.length} contact(s) matching ${JSON.stringify(query)}:\n\n` +
        contacts.map((c, i) => formatContact(c, i)).join("\n\n")
    );
  });
}
