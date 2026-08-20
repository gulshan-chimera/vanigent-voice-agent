// src/functions/getFile.ts
//
// Downloads a specific file's content from the Human Resources KB.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadHrKbFile } from "../lib/sharepointFiles";
import { KbFileDownloadRequestBody } from "../types/knowledgeBase";

export async function getFile(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: KbFileDownloadRequestBody;

  try {
    body = (await request.json()) as KbFileDownloadRequestBody;
  } catch (error) {
    context.error(`[GET-FILE] Failed to parse request body: ${(error as Error).message}`);
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.itemId) {
    return { status: 400, jsonBody: { error: "itemId is required" } };
  }

  const result = await downloadHrKbFile(body.itemId);

  if (!result) {
    return { status: 404, jsonBody: { error: "File not found or download failed" } };
  }

  return { status: 200, jsonBody: result };
}

app.http("getFile", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "getFile",
  handler: getFile,
});