// src/functions/listIndexedFiles.ts
//
// Diagnostic — lists every file currently in the index with its
// driveItemId and driveId, so you can grab IDs for the per-file
// diagnostic endpoints without digging through Azure Portal.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getIndexedFileState } from "../lib/searchIndex";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"
export async function listIndexedFiles(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[LIST-INDEXED] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  const state = await getIndexedFileState();

  if (state === null) {
    return { status: 502, jsonBody: { error: "Failed to read index state" } };
  }

  const files = Array.from(state.values()).map((f) => ({
    fileName: f.fileName,
    library: f.library,
    driveId: f.driveId,
    itemId: f.driveItemId,
  }));

  return { status: 200, jsonBody: { count: files.length, files } };
}

app.http("listIndexedFiles", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "listIndexedFiles",
  handler: listIndexedFiles,
});