// src/functions/testCaptionPage.ts
//
// TEMPORARY diagnostic — runs the caption prompt against specific pages
// of a file and returns the raw model output, so prompt changes can be
// checked without re-indexing.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadDriveFile } from "../lib/sharepointFiles";
import { extractPageTexts } from "../lib/pdfPageText";
import { renderPdfPagesAsImages } from "../lib/pdfImages";
import { generateImageCaption, SKIP_CAPTION } from "../lib/azureOpenAI";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"

export async function testCaptionPage(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
    if (!isRequestAuthorized(request)) {
    context.warn("[TEST-CAPTION] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  let body: { itemId?: string; driveId?: string; pages?: number[] };

  try {
    body = (await request.json()) as { itemId?: string; driveId?: string; pages?: number[] };
  } catch {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.itemId || !body.driveId) {
    return { status: 400, jsonBody: { error: "itemId and driveId are required" } };
  }

  const fileContent = await downloadDriveFile(body.driveId, body.itemId);
  if (!fileContent) {
    return { status: 404, jsonBody: { error: "File not found" } };
  }

  const pageTexts = await extractPageTexts(fileContent.base64Content);
  const pageImages = await renderPdfPagesAsImages(fileContent.base64Content);
  const pageCount = Math.min(pageTexts.length, pageImages.length);

  // Default to all pages if none specified.
  const requested =
    body.pages && body.pages.length > 0
      ? body.pages
      : Array.from({ length: pageCount }, (_, i) => i + 1);

  const results = [];

  for (const pageNum of requested) {
    if (pageNum < 1 || pageNum > pageCount) {
      results.push({ page: pageNum, error: `Out of range (file has ${pageCount} pages)` });
      continue;
    }

    const i = pageNum - 1;
    const caption = await generateImageCaption(pageImages[i].toString("base64"), pageTexts[i]);
    const skipped = !caption || caption.includes(SKIP_CAPTION);

    results.push({
      page: pageNum,
      skipped,
      caption: skipped ? null : caption,
      rawOutput: caption,
    });
  }

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      pagesTested: results.length,
      skippedCount: results.filter((r) => "skipped" in r && r.skipped).length,
      results,
    },
  };
}

app.http("testCaptionPage", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testCaptionPage",
  handler: testCaptionPage,
});