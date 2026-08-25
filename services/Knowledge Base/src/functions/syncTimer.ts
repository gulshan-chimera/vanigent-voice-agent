// src/functions/syncTimer.ts
//
// Scheduled sync. Without this, the index only updates when someone
// POSTs to /api/syncIndex by hand — so in production it would go stale
// the moment anyone edited a document.
//
// This also acts as the self-healing net: a file that failed on a
// transient error gets retried on the next run automatically, with no
// human noticing. Because the diff compares SharePoint against the
// index itself, correctness doesn't depend on any single run succeeding.
//
// NEVER runs with force — a scheduled wipe-and-rebuild would empty the
// index on a timer and take minutes to repopulate, during which callers
// would get "I don't have that information" for everything.

import { app, InvocationContext, Timer } from "@azure/functions";
import { runSync } from "../lib/syncRunner";

export async function syncTimer(timer: Timer, context: InvocationContext): Promise<void> {
  if (timer.isPastDue) {
    context.warn("[SYNC-TIMER] Running late — a previous scheduled run may have been missed.");
  }

  context.log("[SYNC-TIMER] Scheduled sync starting.");

  const result = await runSync({ force: false, trigger: "timer" }, context);

  if (!result.ok) {
    // Log rather than throw: a timer retry storm won't fix an upstream
    // outage, and the next scheduled run will pick things up anyway.
    context.error(`[SYNC-TIMER] Sync failed: ${result.failure.error}`);
    return;
  }

  const s = result.summary;
  const changed = s.filesQueuedNew + s.filesQueuedUpdated + s.filesDeleted;

  if (changed === 0) {
    context.log(`[SYNC-TIMER] No changes — ${s.filesUnchanged} file(s) already current.`);
  } else {
    context.log(
      `[SYNC-TIMER] Queued ${s.filesQueuedNew} new, ${s.filesQueuedUpdated} updated; deleted ${s.filesDeleted}. ${s.filesUnchanged} unchanged.`
    );
  }
}

app.timer("syncTimer", {
  // NCRONTAB: {second} {minute} {hour} {day} {month} {day-of-week}
  // "0 0 * * * *" = on the hour, every hour.
  schedule: "%KB_SYNC_SCHEDULE%",
  runOnStartup: false,
  handler: syncTimer,
});