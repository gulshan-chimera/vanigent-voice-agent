// src/functions/testImageExtraction.ts
//
// TEMPORARY diagnostic endpoint. Extracts images from a given HR KB file
// and returns the raw result, so we can see the real library output
// before building the full captioning pipeline on top of it.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadDriveFile } from "../lib/sharepointFiles";
import { extractPageTexts } from "../lib/pdfPageText";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"

export async function testImageExtraction(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
    if (!isRequestAuthorized(request)) {
    context.warn("[TEST-IMAGE-EXTRACT] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  let body: { itemId?: string; driveId?: string };

  try {
    body = (await request.json()) as { itemId?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.itemId || !body.driveId) {
    return { status: 400, jsonBody: { error: "itemId and driveId are required" } };
  }

  const fileContent = await downloadDriveFile(body.driveId, body.itemId);
  if (!fileContent) {
    return { status: 404, jsonBody: { error: "File not found" } };
  }


  const images = await extractPageTexts(fileContent.base64Content);

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      imagesFound: images.length,
    },
  };
}

app.http("testImageExtraction", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testImageExtraction",
  handler: testImageExtraction,
});