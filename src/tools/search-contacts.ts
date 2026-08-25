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

/** One postal address as a single line, or undefined when it is empty. */
function formatAddress(address: any): string | undefined {
  const parts = [address?.street, address?.city, address?.state, address?.postalCode, address?.countryOrRegion]
    .map((p: unknown) => String(p ?? "").trim())
    .filter(Boolean);
  return parts.length ? parts.join(", ") : undefined;
}

export function formatContact(c: any, index?: number): string {
  const emails = (c.emailAddresses ?? [])
    .map((e: any) => e.address)
    .filter(Boolean)
    .join(", ");
  const phones = [...(c.businessPhones ?? []), ...(c.homePhones ?? []), c.mobilePhone]
    .filter(Boolean)
    .join(", ");
  const home = formatAddress(c.homeAddress);
  const business = formatAddress(c.businessAddress);
  return [
    `${index !== undefined ? `${index + 1}. ` : ""}${c.displayName || [c.givenName, c.surname].filter(Boolean).join(" ") || "(no name)"}${c.nickName ? ` ("${c.nickName}")` : ""}`,
    `   Email: ${emails || "(none)"}`,
    `   Phone: ${phones || "(none)"}`,
    ...(c.companyName || c.jobTitle
      ? [`   Work: ${[c.jobTitle, c.companyName].filter(Boolean).join(", ")}`]
      : []),
    ...(c.birthday ? [`   Birthday: ${String(c.birthday).slice(0, 10)}`] : []),
    ...(home ? [`   Home address: ${home}`] : []),
    ...(business ? [`   Business address: ${business}`] : []),
    ...(c.categories?.length ? [`   Categories: ${c.categories.join(", ")}`] : []),
    ...(c.personalNotes ? [`   Notes: ${String(c.personalNotes).slice(0, 200)}`] : []),
    `   Contact id: ${c.id}`,
  ].join("\n");
}

const SELECT =
  "id,displayName,givenName,surname,nickName,emailAddresses,businessPhones,homePhones,mobilePhone," +
  "companyName,jobTitle,birthday,homeAddress,businessAddress,personalNotes,categories";

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
