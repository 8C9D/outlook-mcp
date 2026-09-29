// Which job a Worker cron tick runs. Pure, so the offline tier can prove the
// dispatch (o27) without a Worker runtime; src/worker/index.ts only acts on the
// answer. Every expression here must match `triggers.crons` in wrangler.jsonc.
import { torontoHourOf } from "./auto-filing.js";

/** The hour, America/Toronto, at which the morning brief is drafted. */
export const DIGEST_HOUR_TORONTO = 7;

/**
 * The daily health-check schedule. 13:37 UTC is 09:37 Toronto in EDT and 08:37
 * in EST — always morning for the owner, after the digest, and colliding with
 * neither the 6-hourly upkeep ticks (minute 17) nor the digest ticks (11:00 and
 * 12:00 UTC). This one is dispatched on the cron expression rather than the
 * local hour, since unlike the digest it has no wall-clock meaning to preserve
 * across DST.
 */
export const HEALTH_CRON = "37 13 * * *";

/**
 * The self-alert heartbeat watchdog (core/self-alert.js), hourly at minute 47,
 * which no other tick uses. Dispatched on the expression, and before the
 * hour-based branches: it also fires at 11:47 and 12:47 UTC, one of which is
 * 07:xx in Toronto and would otherwise be taken for the digest tick.
 */
export const SELF_ALERT_WATCHDOG_CRON = "47 * * * *";

export type ScheduledJob = "health" | "self-alert-watchdog" | "digest" | "upkeep";

/**
 * The job for one tick. The exact-expression jobs come first; the remaining
 * ticks (digest and upkeep) are told apart by the hour Toronto is actually on.
 */
export function scheduledJobFor(cron: string, scheduledTime: Date): ScheduledJob {
  if (cron === HEALTH_CRON) return "health";
  if (cron === SELF_ALERT_WATCHDOG_CRON) return "self-alert-watchdog";
  if (torontoHourOf(scheduledTime) === DIGEST_HOUR_TORONTO) return "digest";
  return "upkeep";
}
