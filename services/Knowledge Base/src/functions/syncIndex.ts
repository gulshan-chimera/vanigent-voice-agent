// src/functions/syncIndex.ts
//
// Manual sync trigger. The actual logic lives in lib/syncRunner so the
// scheduled timer (syncTimer) runs identical code.
//
// Returns as soon as jobs are queued, NOT when indexing finishes —
// check /api/queueStatus for progress.
//
// Pass ?force=true to wipe the index and rebuild from scratch (only
// after a schema change or a caption-prompt change).

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { runSync } from "../lib/syncRunner";

export async function syncIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[SYNC] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  const force = request.query.get("force") === "true";

  const result = await runSync({ force, trigger: "manual" }, context);

  if (!result.ok) {
    // A bad library selection is a caller mistake (400); everything else
    // is an upstream failure (502).
    const status = result.failure.availableLibraries ? 400 : 502;
    return { status, jsonBody: result.failure };
  }

  return { status: 200, jsonBody: result.summary };
}

app.http("syncIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "syncIndex",
  handler: syncIndex,
});