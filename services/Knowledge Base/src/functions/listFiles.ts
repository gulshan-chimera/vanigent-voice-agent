// src/functions/listFiles.ts
//
// Lists files/folders in the Human Resources KB SharePoint library.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { listHrKbFiles } from "../lib/sharepointFiles";
import { KbFileListRequestBody } from "../types/knowledgeBase";

export async function listFiles(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: KbFileListRequestBody = {};

  try {
    const text = await request.text();
    if (text) {
      body = JSON.parse(text);
    }
  } catch (error) {
    context.error(`[LIST-FILES] Failed to parse request body: ${(error as Error).message}`);
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  const result = await listHrKbFiles(body.folderPath);

  if (result === null) {
    return { status: 502, jsonBody: { error: "Failed to list files from SharePoint" } };
  }

  return { status: 200, jsonBody: result };
}

app.http("listFiles", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "listFiles",
  handler: listFiles,
});