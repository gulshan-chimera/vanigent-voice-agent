// src/functions/listDriveFiles.ts
//
// Diagnostic — lists every file in a library with its itemId, including
// files we don't index (non-PDF). listIndexedFiles only shows what's
// already in the index, so this is the way to find IDs for files that
// haven't been processed yet.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { listAllFilesInDrive } from "../lib/sharepointFiles";

export async function listDriveFiles(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[LIST-DRIVE-FILES] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: { driveId?: string };

  try {
    body = (await request.json()) as { driveId?: string };
  } catch {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.driveId) {
    return { status: 400, jsonBody: { error: "driveId is required" } };
  }

  const files = await listAllFilesInDrive(body.driveId);

  if (files === null) {
    return { status: 502, jsonBody: { error: "Failed to list files from SharePoint" } };
  }

  return {
    status: 200,
    jsonBody: {
      count: files.length,
      files: files.map((f) => ({
        name: f.name,
        itemId: f.id,
        size: f.size,
        mimeType: f.file?.mimeType,
        lastModifiedDateTime: f.lastModifiedDateTime,
      })),
    },
  };
}

app.http("listDriveFiles", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "listDriveFiles",
  handler: listDriveFiles,
});
