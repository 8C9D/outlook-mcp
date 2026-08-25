import { z } from "zod";
import { callGraphServer } from "../core/graph.js";
import { ToolResult, errorResult, runTool, textResult } from "./common.js";
import { formatContact } from "./search-contacts.js";
import { resolveCategoryNames } from "./manage-categories.js";
import { prepareBase64Source, prepareUrlSource } from "./file-sources.js";

/** Outlook contact photos are capped at 4 MB by Graph. */
const PHOTO_MAX = 4 * 1024 * 1024;

const addressSchema = z
  .object({
    street: z.string().optional(),
    city: z.string().optional(),
    state: z.string().optional(),
    postal_code: z.string().optional(),
    country: z.string().optional(),
  })
  .describe("A postal address; any subset of fields. On update it REPLACES that whole address.");

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
  nickname: z.string().optional().describe("Nickname."),
  emails: z
    .array(z.string().email())
    .optional()
    .describe("Email addresses. On update this REPLACES the contact's whole email list."),
  phones: z
    .array(z.string())
    .optional()
    .describe(
      "Business phone numbers. On update this REPLACES the contact's business-phone list."
    ),
  home_phones: z
    .array(z.string())
    .optional()
    .describe("Home phone numbers. On update this REPLACES the home-phone list."),
  mobile_phone: z.string().optional().describe("Mobile phone number (one; an empty string clears it)."),
  company: z.string().optional().describe("Company name."),
  job_title: z.string().optional().describe("Job title."),
  birthday: z
    .string()
    .regex(/^\d{4}-\d{2}-\d{2}$/, "ISO date (YYYY-MM-DD)")
    .optional()
    .describe("Birthday as YYYY-MM-DD. Outlook adds it to the birthday calendar."),
  home_address: addressSchema.optional(),
  business_address: addressSchema.optional(),
  personal_notes: z
    .string()
    .optional()
    .describe("Free-text notes on the contact. On update this REPLACES the existing notes."),
  categories: z
    .array(z.string().min(1))
    .optional()
    .describe(
      "Outlook category names for the contact (must exist in manage_categories' list). On update this REPLACES the contact's categories; an empty array clears them."
    ),
  photo_url: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Set the contact's photo from an https URL (JPEG/PNG, max 4 MB). Replaces any existing photo."
    ),
  photo_base64: z
    .string()
    .min(1)
    .optional()
    .describe("Set the contact's photo from base64 image bytes (max 3 MB decoded). Replaces any existing photo."),
};

const manageContactArgs = z.object(manageContactSchema);

export const manageContactDescription =
  "Create, update, or delete a saved Outlook contact — names, emails, business/home/mobile phones, company and job title, birthday, home and business addresses, notes, categories, and the contact photo (from photo_url or photo_base64). Before calling with delete, state the contact's name so the user knows who is being removed; deletion is soft (the contact goes to Deleted Items). On update, every list-shaped field (emails, phones, home_phones, categories, each address) REPLACES the existing value rather than appending.";

/** Graph physicalAddress from the tool's address shape (omitted fields are dropped). */
function toGraphAddress(address: z.infer<typeof addressSchema>): Record<string, string> {
  return {
    ...(address.street !== undefined ? { street: address.street } : {}),
    ...(address.city !== undefined ? { city: address.city } : {}),
    ...(address.state !== undefined ? { state: address.state } : {}),
    ...(address.postal_code !== undefined ? { postalCode: address.postal_code } : {}),
    ...(address.country !== undefined ? { countryOrRegion: address.country } : {}),
  };
}

const SELECT =
  "id,displayName,givenName,surname,nickName,emailAddresses,businessPhones,homePhones,mobilePhone," +
  "companyName,jobTitle,birthday,homeAddress,businessAddress,personalNotes,categories";

export async function manageContactHandler(
  input: z.input<typeof manageContactArgs>
): Promise<ToolResult> {
  return runTool(async () => {
    const args = manageContactArgs.parse(input);
    const { action, contact_id, photo_url, photo_base64 } = args;

    if (photo_url && photo_base64) {
      return errorResult("Give one photo source: photo_url or photo_base64, not both.");
    }

    const fieldPatch = async (): Promise<any> => ({
      ...(args.given_name !== undefined ? { givenName: args.given_name } : {}),
      ...(args.surname !== undefined ? { surname: args.surname } : {}),
      ...(args.nickname !== undefined ? { nickName: args.nickname } : {}),
      ...(args.emails !== undefined
        ? { emailAddresses: args.emails.map((address) => ({ address })) }
        : {}),
      ...(args.phones !== undefined ? { businessPhones: args.phones } : {}),
      ...(args.home_phones !== undefined ? { homePhones: args.home_phones } : {}),
      ...(args.mobile_phone !== undefined ? { mobilePhone: args.mobile_phone || null } : {}),
      ...(args.company !== undefined ? { companyName: args.company } : {}),
      ...(args.job_title !== undefined ? { jobTitle: args.job_title } : {}),
      // Graph stores birthdays as a UTC instant and normalizes the clock time
      // itself (verified live: it comes back as ...T11:59:00Z); only the date
      // part is meaningful.
      ...(args.birthday !== undefined ? { birthday: `${args.birthday}T00:00:00Z` } : {}),
      ...(args.home_address !== undefined ? { homeAddress: toGraphAddress(args.home_address) } : {}),
      ...(args.business_address !== undefined
        ? { businessAddress: toGraphAddress(args.business_address) }
        : {}),
      ...(args.personal_notes !== undefined ? { personalNotes: args.personal_notes } : {}),
      ...(args.categories !== undefined
        ? { categories: args.categories.length ? await resolveCategoryNames(args.categories) : [] }
        : {}),
    });

    /** PUT the photo bytes onto a contact; returns a line for the answer. */
    const applyPhoto = async (contactId: string): Promise<string> => {
      const prepared = photo_url
        ? await prepareUrlSource(photo_url, "photo")
        : prepareBase64Source(photo_base64!, "photo");
      if (!prepared.ok) throw new Error(prepared.message);
      const { buffer, contentType } = await prepared.source.read();
      if (buffer.length > PHOTO_MAX) {
        throw new Error(`The photo is ${(buffer.length / 1024 / 1024).toFixed(1)} MB — Outlook caps contact photos at 4 MB.`);
      }
      await callGraphServer(`/me/contacts/${encodeURIComponent(contactId)}/photo/$value`, {
        method: "PUT",
        headers: { "Content-Type": contentType ?? "image/jpeg" },
        body: new Uint8Array(buffer),
      });
      return `Photo: set (${(buffer.length / 1024).toFixed(0)} KB)`;
    };

    switch (action) {
      case "create": {
        if (!args.given_name) return errorResult('Action "create" requires given_name.');
        const created = await callGraphServer(`/me/contacts?$select=${SELECT}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(await fieldPatch()),
        });
        const photoLine = photo_url || photo_base64 ? `\n${await applyPhoto(created.id)}` : "";
        return textResult(`Contact created.\n${formatContact(created)}${photoLine}`);
      }
      case "update": {
        if (!contact_id) return errorResult('Action "update" requires contact_id.');
        const patch = await fieldPatch();
        if (Object.keys(patch).length === 0 && !photo_url && !photo_base64) {
          return errorResult(
            "Nothing to update — provide at least one contact field (or a photo source)."
          );
        }
        const updated = Object.keys(patch).length
          ? await callGraphServer(`/me/contacts/${encodeURIComponent(contact_id)}`, {
              method: "PATCH",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(patch),
            })
          : await callGraphServer(`/me/contacts/${encodeURIComponent(contact_id)}?$select=${SELECT}`);
        const photoLine = photo_url || photo_base64 ? `\n${await applyPhoto(contact_id)}` : "";
        const changed = [
          ...Object.keys(patch),
          ...(photo_url || photo_base64 ? ["photo"] : []),
        ];
        return textResult(`Contact updated (${changed.join(", ")}).\n${formatContact(updated)}${photoLine}`);
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
