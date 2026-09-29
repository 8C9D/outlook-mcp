// Which jobs a Worker cron tick runs. Pure, so the offline tier can prove the
// dispatch (o30) without a Worker runtime; src/worker/index.ts only acts on the
// answer. Every expression here must match `triggers.crons` in wrangler.jsonc.
//
// The account is on Workers Free, which allows 5 cron triggers per account, and
// every one of them is already spoken for (this Worker's 4 plus another
// Worker's). A new job therefore rides an existing tick rather than getting a
// cron of its own — as the self-alert watchdog does.
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
 * The subscription-upkeep schedule: 00:17, 06:17, 12:17 and 18:17 UTC. Each of
 * its ticks runs the job the Toronto hour selects (upkeep, or the digest on the
 * winter 12:17 tick, which is 07:17 in EST) and then, once that job has
 * settled, the self-alert heartbeat watchdog (core/self-alert.js) — so a lapsed
 * job is noticed within 6 hours without the watchdog spending a cron trigger of
 * its own.
 */
export const UPKEEP_CRON = "17 */6 * * *";

export type ScheduledJob = "health" | "digest" | "upkeep" | "self-alert-watchdog";

/**
 * The jobs for one tick, in the order runJobsInOrder runs them. The health
 * check is matched on its exact expression and runs alone. Every other tick
 * runs the digest or upkeep, told apart by the hour Toronto is actually on, and
 * an UPKEEP_CRON tick then runs the watchdog, whichever of the two it got. The
 * digest ticks never run the watchdog, even when they fall back to upkeep.
 */
export function scheduledJobsFor(cron: string, scheduledTime: Date): ScheduledJob[] {
  if (cron === HEALTH_CRON) return ["health"];
  const job = torontoHourOf(scheduledTime) === DIGEST_HOUR_TORONTO ? "digest" : "upkeep";
  return cron === UPKEEP_CRON ? [job, "self-alert-watchdog"] : [job];
}

/**
 * Run a tick's jobs one after another: each starts only once the one before it
 * has settled, fulfilled or rejected, so the watchdog never overlaps the
 * upkeep or digest — any of them may refresh the mailbox token, which Microsoft
 * rotates on every exchange. A rejection or synchronous throw goes to onError
 * and never stops the jobs after it; the returned promise never rejects.
 */
export async function runJobsInOrder(
  jobs: readonly ScheduledJob[],
  run: (job: ScheduledJob) => Promise<void>,
  onError: (job: ScheduledJob, err: unknown) => void
): Promise<void> {
  for (const job of jobs) {
    try {
      await run(job);
    } catch (err) {
      onError(job, err);
    }
  }
}
