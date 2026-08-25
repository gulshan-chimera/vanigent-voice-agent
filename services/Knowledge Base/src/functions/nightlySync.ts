// src/functions/nightlySync.ts
//
// Runs the same incremental sync as syncIndex, automatically every night
// at 2 AM, so SharePoint adds/edits/deletes get picked up without anyone
// having to call the manual endpoint. Always runs a normal (non-force)
// pass — a nightly force-wipe would defeat the point of incremental sync.
// Schedule is in the App Service's configured time zone
// (WEBSITE_TIME_ZONE / TZ) — UTC by default.

import { app, InvocationContext, Timer } from "@azure/functions";
import { runSync } from "./syncIndex";

export async function nightlySync(myTimer: Timer, context: InvocationContext): Promise<void> {
  context.log("[NIGHTLY-SYNC] Starting scheduled KB sync");

  try {
    const summary = await runSync(context, false);
    context.log(`[NIGHTLY-SYNC] Completed: ${JSON.stringify(summary)}`);
  } catch (error) {
    context.error(`[NIGHTLY-SYNC] Sync failed: ${(error as Error).message}`);
  }
}

app.timer("nightlySync", {
  schedule: "0 0 2 * * *",
  handler: nightlySync,
});
