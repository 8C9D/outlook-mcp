import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatContact } from "./search-contacts.js";

export const manageContactSchema = {
  action: z
    .enum(["create", "update", "delete"])
    .describe(
      "create: new contact (given_name required); update: change fields on contact_id; delete: move the contact to Deleted Items (soft delete)."
    ),
  contact_id: z
    .string()
    .optional()
    .describe("The contact to update or delete (from search_contacts). Required for update/delete."),
  given_name: z.string().optional().describe("First name. Required for create."),
  surname: z.string().optional().describe("Last name."),
  emails: z
    .array(z.string().email())
    .optional()
    .describe("Email addresses. On update this REPLACES the contact's whole email list."),
  phones: z
    .array(z.string())
    .optional()
    .describe(
      "Phone numbers (stored as business phones). On update this REPLACES the contact's business-phone list."
    ),
  company: z.string().optional().describe("Company name."),
};

const manageContactArgs = z.object(manageContactSchema);

export const manageContactDescription =
  "Create, update, or delete a saved Outlook contact. Before calling with delete, state the contact's name so the user knows who is being removed; deletion is soft (the contact goes to Deleted Items). On update, the emails and phones arrays replace the existing lists rather than appending.";

const SELECT =
  "id,displayName,givenName,surname,emailAddresses,businessPhones,homePhones,mobilePhone,companyName";

export async function manageContactHandler(
  input: z.input<typeof manageContactArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const { action, contact_id, given_name, surname, emails, phones, company } =
      manageContactArgs.parse(input);

    const fieldPatch = (): any => ({
      ...(given_name !== undefined ? { givenName: given_name } : {}),
      ...(surname !== undefined ? { surname } : {}),
      ...(emails !== undefined
        ? { emailAddresses: emails.map((address) => ({ address })) }
        : {}),
      ...(phones !== undefined ? { businessPhones: phones } : {}),
      ...(company !== undefined ? { companyName: company } : {}),
    });

    switch (action) {
      case "create": {
        if (!given_name) return errorResult('Action "create" requires given_name.');
        const created = await callGraphServer(`/me/contacts?$select=${SELECT}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(fieldPatch()),
        });
        return textResult(`Contact created.\n${formatContact(created)}`);
      }
      case "update": {
        if (!contact_id) return errorResult('Action "update" requires contact_id.');
        const patch = fieldPatch();
        if (Object.keys(patch).length === 0) {
          return errorResult(
            "Nothing to update — provide at least one of given_name, surname, emails, phones, company."
          );
        }
        const updated = await callGraphServer(
          `/me/contacts/${encodeURIComponent(contact_id)}`,
          {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(patch),
          }
        );
        return textResult(`Contact updated (${Object.keys(patch).join(", ")}).\n${formatContact(updated)}`);
      }
      case "delete": {
        if (!contact_id) return errorResult('Action "delete" requires contact_id.');
        const existing = await callGraphServer(
          `/me/contacts/${encodeURIComponent(contact_id)}?$select=displayName`
        );
        await callGraphServer(`/me/contacts/${encodeURIComponent(contact_id)}`, {
          method: "DELETE",
        });
        return textResult(
          `Contact "${existing.displayName || contact_id}" deleted (moved to Deleted Items — recoverable).`
        );
      }
    }
  });
}
