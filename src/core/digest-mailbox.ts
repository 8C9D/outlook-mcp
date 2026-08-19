// The Graph implementation of DigestMailbox: four reads and one draft.
//
// Kept separate from core/mail-actions.js so that neither LLM feature can reach
// the other's capabilities — the classifier cannot create a draft, and the
// digest cannot move or categorize anything. Neither can send: there is no
// /sendMail and no /send here, and there never may be. The only send path in
// this codebase is the send_draft tool, driven by a human.
import { callGraphServer } from "./graph.js";
import type { DigestEvent, DigestMail, DigestMailbox, DigestTask } from "./digest.js";

const TIMEZONE = "America/Toronto";
const TZ_PREFER = `outlook.timezone="${TIMEZONE}"`;

/** At most this many To Do lists are searched for due tasks. */
const TASK_LIST_CAP = 10;

export function graphDigestMailbox(): DigestMailbox {
  return {
    async ownAddress() {
      const me = await callGraphServer("/me?$select=mail,userPrincipalName");
      const address = me?.mail ?? me?.userPrincipalName;
      if (!address) throw new Error("Graph did not report an address for this mailbox.");
      return String(address);
    },

    async unreadSince(sinceIsoUtc, cap) {
      const filter = `isRead eq false and receivedDateTime ge ${sinceIsoUtc}`;
      const data = await callGraphServer(
        `/me/mailFolders/inbox/messages?$filter=${encodeURIComponent(filter)}` +
          `&$select=subject,from,receivedDateTime,bodyPreview&$orderby=receivedDateTime desc&$top=${cap}`
      );
      return (data?.value ?? []).slice(0, cap).map(
        (message: any): DigestMail => ({
          subject: String(message?.subject ?? "(no subject)"),
          from: formatSender(message?.from),
          receivedDateTime: message?.receivedDateTime ?? undefined,
          preview: String(message?.bodyPreview ?? ""),
        })
      );
    },

    async eventsOn(torontoDate, cap) {
      const nextDay = new Date(`${torontoDate}T00:00:00Z`);
      nextDay.setUTCDate(nextDay.getUTCDate() + 1);
      const path =
        `/me/calendarView?startDateTime=${encodeURIComponent(`${torontoDate}T00:00:00`)}` +
        `&endDateTime=${encodeURIComponent(`${nextDay.toISOString().slice(0, 10)}T00:00:00`)}` +
        `&$select=subject,start,end,location,isAllDay&$orderby=start/dateTime&$top=${cap}`;
      const data = await callGraphServer(path, { headers: { Prefer: TZ_PREFER } });
      return (data?.value ?? []).slice(0, cap).map(
        (event: any): DigestEvent => ({
          subject: String(event?.subject ?? "(no subject)"),
          start: event?.isAllDay ? "all day" : clock(event?.start?.dateTime),
          end: event?.isAllDay ? undefined : clock(event?.end?.dateTime),
          location: event?.location?.displayName || undefined,
        })
      );
    },

    async tasksDueBy(isoDate, cap) {
      const lists = await callGraphServer("/me/todo/lists?$top=50");
      const due: DigestTask[] = [];
      for (const list of (lists?.value ?? []).slice(0, TASK_LIST_CAP)) {
        if (!list?.id || due.length >= cap) continue;
        const tasks = await callGraphServer(
          `/me/todo/lists/${encodeURIComponent(list.id)}/tasks` +
            `?$filter=${encodeURIComponent("status ne 'completed'")}&$top=100`
        ).catch(() => null);
        for (const task of tasks?.value ?? []) {
          // Graph's To Do dueDateTime is a date in the task's own timezone;
          // comparing the date part is what a human means by "due by".
          const dueDate = String(task?.dueDateTime?.dateTime ?? "").slice(0, 10);
          if (!dueDate || dueDate > isoDate) continue;
          due.push({ title: String(task?.title ?? "(untitled)"), due: dueDate, list: list.displayName });
          if (due.length >= cap) break;
        }
      }
      return due.sort((a, b) => (a.due ?? "").localeCompare(b.due ?? "")).slice(0, cap);
    },

    async createDraft(to, subject, body) {
      const draft = await callGraphServer("/me/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subject,
          body: { contentType: "Text", content: body },
          toRecipients: [{ emailAddress: { address: to } }],
        }),
      });
      const id = draft?.id;
      if (!id) throw new Error("Graph did not return an id for the morning-brief draft.");
      return String(id);
    },
  };
}

function formatSender(from: any): string {
  const name = from?.emailAddress?.name;
  const address = from?.emailAddress?.address ?? "(unknown)";
  return name && name !== address ? `${name} <${address}>` : String(address);
}

/** HH:MM from a naive Graph datetime already in the Prefer timezone. */
function clock(dateTime: string | undefined): string {
  return dateTime ? dateTime.slice(11, 16) : "(no time)";
}
