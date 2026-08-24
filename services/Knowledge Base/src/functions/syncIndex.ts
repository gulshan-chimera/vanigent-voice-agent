// src/functions/syncIndex.ts
//
// Full sync: clears the index and re-indexes all current PDFs from the
// HR KB SharePoint library. Each PAGE becomes one chunk — combining that
// page's extracted text with a context-aware caption of its visual
// elements, keeping steps and their screenshots together.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { listHrKbFiles, downloadHrKbFile } from "../lib/sharepointFiles";
import { ensureIndexExists, clearIndex, uploadDocument, HrKbDocument } from "../lib/searchIndex";
import { extractPageTexts } from "../lib/pdfPageText";
import {renderPdfPagesAsImages} from "../lib/pdfImages"
import { generateEmbedding, generateImageCaption } from "../lib/azureOpenAI";
import { DriveItem } from "../types/knowledgeBase";

export interface SyncResult {
  indexedPages: number;
  skippedPages: number;
  skippedFiles: number;
  totalPdfFiles: number;
}

/**
 * Runs the full HR KB resync (clear + re-index every PDF page). Shared by
 * the manual HTTP trigger and the nightly timer trigger. Throws on setup
 * failures (index/list) so each caller can report them its own way.
 */
export async function runSync(context: InvocationContext): Promise<SyncResult> {
  const indexReady = await ensureIndexExists();
  if (!indexReady) {
    throw new Error("Failed to prepare search index");
  }

  await clearIndex();

  const listResult = await listHrKbFiles();
  if (!listResult) {
    throw new Error("Failed to list files from SharePoint");
  }

  const pdfItems = listResult.value.filter(
    (item: DriveItem) => item.file && item.name.toLowerCase().endsWith(".pdf")
  );

  let indexedPages = 0;
  let skippedPages = 0;
  let skippedFiles = 0;

  for (const item of pdfItems) {
    context.log(`[SYNC] Processing: ${item.name}`);

    const fileContent = await downloadHrKbFile(item.id);
    if (!fileContent) {
      context.warn(`[SYNC] Failed to download: ${item.name}`);
      skippedFiles++;
      continue;
    }

    const pageTexts = await extractPageTexts(fileContent.base64Content);
    const pageImages = await renderPdfPagesAsImages(fileContent.base64Content);

    if (pageTexts.length === 0 || pageImages.length === 0) {
      context.warn(`[SYNC] No pages extracted from: ${item.name}`);
      skippedFiles++;
      continue;
    }

    if (pageTexts.length !== pageImages.length) {
      context.warn(
        `[SYNC] Page count mismatch for "${item.name}": ${pageTexts.length} text pages vs ${pageImages.length} image pages. Using the smaller count.`
      );
    }

    const pageCount = Math.min(pageTexts.length, pageImages.length);

    for (let i = 0; i < pageCount; i++) {
      const pageText = pageTexts[i];
      const pageImageBase64 = pageImages[i].toString("base64");

      const caption = await generateImageCaption(pageImageBase64, pageText);
      const captionText =
        caption && caption !== "No significant visual content."
          ? `\n\n[Visual content on this page: ${caption}]`
          : "";

      const combinedContent = `${pageText}${captionText}`.trim();

      if (combinedContent.length === 0) {
        context.warn(`[SYNC] Page ${i + 1} of "${item.name}" has no content — skipping.`);
        skippedPages++;
        continue;
      }

      const embedding = await generateEmbedding(combinedContent);
      if (!embedding) {
        context.warn(`[SYNC] Failed to embed page ${i + 1} of: ${item.name}`);
        skippedPages++;
        continue;
      }

      const doc: HrKbDocument = {
        id: Buffer.from(`${item.id}__page${i}`).toString("base64").replace(/[+/=]/g, "_"),
        fileName: item.name,
        content: combinedContent,
        webUrl: item.webUrl ?? "",
        driveItemId: item.id,
        chunkIndex: i,
        lastModifiedDateTime: item.lastModifiedDateTime ?? "",
        contentVector: embedding,
      };

      const uploaded = await uploadDocument(doc);
      if (uploaded) {
        indexedPages++;
      } else {
        skippedPages++;
      }
    }

    context.log(`[SYNC] Finished "${item.name}": ${pageCount} page(s) processed`);
  }

  return { indexedPages, skippedPages, skippedFiles, totalPdfFiles: pdfItems.length };
}

export async function syncIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const result = await runSync(context);
    return { status: 200, jsonBody: result };
  } catch (error) {
    return { status: 502, jsonBody: { error: (error as Error).message } };
  }
}

app.http("syncIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "syncIndex",
  handler: syncIndex,
});