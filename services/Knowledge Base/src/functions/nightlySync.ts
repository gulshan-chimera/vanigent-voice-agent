// src/functions/nightlySync.ts
//
// Runs the same full HR KB resync as syncIndex, automatically every night
// at 2 AM, so SharePoint adds/edits/deletes get picked up without anyone
// having to call the manual endpoint. Schedule is in the App Service's
// configured time zone (WEBSITE_TIME_ZONE / TZ) — UTC by default.

import { app, InvocationContext, Timer } from "@azure/functions";
import { runSync } from "./syncIndex";

export async function nightlySync(myTimer: Timer, context: InvocationContext): Promise<void> {
  context.log("[NIGHTLY-SYNC] Starting scheduled HR KB sync");

  try {
    const result = await runSync(context);
    context.log(`[NIGHTLY-SYNC] Completed: ${JSON.stringify(result)}`);
  } catch (error) {
    context.error(`[NIGHTLY-SYNC] Sync failed: ${(error as Error).message}`);
  }
}

app.timer("nightlySync", {
  schedule: "0 0 2 * * *",
  handler: nightlySync,
});
