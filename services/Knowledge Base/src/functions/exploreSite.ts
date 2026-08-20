// src/functions/exploreSite.ts
//
// TEMPORARY diagnostic endpoint — returns the site ID and list of document
// libraries for VanigentPortal, so we can see where "Knowledge Base"
// actually lives before building the real endpoints.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { exploreSite } from "../lib/sharepointExplore";

export async function exploreSiteHandler(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const result = await exploreSite();

  if (!result) {
    return { status: 502, jsonBody: { error: "Failed to explore SharePoint site" } };
  }

  return { status: 200, jsonBody: result };
}

app.http("exploreSite", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "exploreSite",
  handler: exploreSiteHandler,
});