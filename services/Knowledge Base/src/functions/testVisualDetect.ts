// src/functions/testVisualDetect.ts
//
// TEMPORARY diagnostic — measures the RENDERED SIZE of the largest image
// on each page, and lines that up against what gpt-4o actually found
// worth describing (read from the captions already in the index).
//
// Purpose: find a size threshold that separates a header logo from a
// real screenshot or chart, so we can skip vision captioning on pages
// that carry no meaningful visual content.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadDriveFile } from "../lib/sharepointFiles";
import { detectPageVisuals, PageVisualInfo } from "../lib/pdfImageDetect";
import { getSearchClient } from "../lib/searchIndex";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"

const CAPTION_MARKER = "[Visual content on this page:";

export async function testVisualDetect(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
    if (!isRequestAuthorized(request)) {
    context.warn("[VISUAL-DETECT] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  let body: { itemId?: string; driveId?: string };

  try {
    body = (await request.json()) as { itemId?: string; driveId?: string };
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

  const detected = await detectPageVisuals(fileContent.base64Content);

  if (detected.length === 0) {
    return { status: 502, jsonBody: { error: "Failed to analyse PDF pages" } };
  }

  // What gpt-4o actually produced, from the existing index.
  // chunkIndex is 0-based; page numbers here are 1-based.
  const captionedPages = new Map<number, boolean>();
  const searchClient = getSearchClient();

  if (searchClient) {
    try {
      const results = await searchClient.search("*", {
        filter: `driveItemId eq '${body.itemId.replace(/'/g, "''")}'`,
        select: ["chunkIndex", "content"],
        top: 1000,
      });

      for await (const result of results.results) {
        captionedPages.set(
          result.document.chunkIndex,
          result.document.content.includes(CAPTION_MARKER)
        );
      }
    } catch (error) {
      context.warn(`[VISUAL-DETECT] Could not read index: ${(error as Error).message}`);
    }
  }

  if (captionedPages.size === 0) {
    context.warn(
      "[VISUAL-DETECT] No indexed pages found for this file — comparison will be empty. Has it been synced?"
    );
  }

  const comparison = detected.map((p) => ({
    page: p.page,
    imageOpCount: p.imageOpCount,
    largestImage: `${p.largestImageWidth}x${p.largestImageHeight}pt`,
    largestImageArea: p.largestImageArea,
    pathOpCount: p.pathOpCount,
    charCount: p.charCount,
    gptFoundVisual: captionedPages.get(p.page - 1) ?? null,
  }));

  // Split by what gpt-4o decided, so the size gap (if any) is visible.
  const meaningful = detected.filter((p) => captionedPages.get(p.page - 1) === true);
  const notMeaningful = detected.filter((p) => captionedPages.get(p.page - 1) === false);

  const sortedAreas = (pages: PageVisualInfo[]): number[] =>
    pages.map((p) => p.largestImageArea).sort((a, b) => a - b);

  const meaningfulAreas = sortedAreas(meaningful);
  const notMeaningfulAreas = sortedAreas(notMeaningful);

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      pageCount: detected.length,
      indexedPageCount: captionedPages.size,

      // The pages that actually matter — with their image sizes.
      meaningfulPages: meaningful.map((p) => ({
        page: p.page,
        size: `${p.largestImageWidth}x${p.largestImageHeight}pt`,
        area: p.largestImageArea,
      })),

      // The numbers that set the threshold: if the smallest meaningful
      // area sits clearly above the largest not-meaningful area, there's
      // a clean cut-off between them.
      smallestMeaningfulArea: meaningfulAreas[0] ?? null,
      largestNotMeaningfulArea: notMeaningfulAreas[notMeaningfulAreas.length - 1] ?? null,
      meaningfulAreaRange: meaningfulAreas,
      notMeaningfulAreaRange: notMeaningfulAreas,

      comparison,
    },
  };
}

app.http("testVisualDetect", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testVisualDetect",
  handler: testVisualDetect,
});