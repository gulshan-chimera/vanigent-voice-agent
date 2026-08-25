// src/functions/exploreSite.ts
//
// Diagnostic — lists every document library on the site with its ID and
// name. Useful for populating KB_LIBRARY_ALLOWLIST and for grabbing a
// driveId when testing the per-file diagnostics below.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { listSiteDrives } from "../lib/sharepointFiles";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"
export async function exploreSiteHandler(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[EXPLORE-SITE] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  const drives = await listSiteDrives();

  if (!drives) {
    return { status: 502, jsonBody: { error: "Failed to list SharePoint libraries" } };
  }

  return {
    status: 200,
    jsonBody: {
      count: drives.length,
      drives: drives.map((d) => ({ id: d.id, name: d.name, webUrl: d.webUrl })),
    },
  };
}

app.http("exploreSite", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "exploreSite",
  handler: exploreSiteHandler,
});