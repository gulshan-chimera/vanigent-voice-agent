// src/functions/syncIndex.ts
//
// PROTOTYPE sync — manually triggered for now. Clears the index and
// re-indexes all current PDFs from the HR KB SharePoint library.
// PDF-only, plain keyword search — the pre-Azure-OpenAI version.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { listHrKbFiles, downloadHrKbFile } from "../lib/sharepointFiles";
import { ensureIndexExists, clearIndex, uploadDocument, HrKbDocument } from "../lib/searchIndex";
import { extractPdfText } from "../lib/pdfText";
import { DriveItem } from "../types/knowledgeBase";
import { generateEmbedding } from "../lib/azureOpenAI";
import { chunkText } from "../lib/textChunker";

export async function syncIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const indexReady = await ensureIndexExists();
  if (!indexReady) {
    return { status: 502, jsonBody: { error: "Failed to prepare search index" } };
  }

  await clearIndex();

  const listResult = await listHrKbFiles();
  if (!listResult) {
    return { status: 502, jsonBody: { error: "Failed to list files from SharePoint" } };
  }

  const pdfItems = listResult.value.filter(
    (item: DriveItem) => item.file && item.name.toLowerCase().endsWith(".pdf")
  );

  let indexed = 0;
  let skipped = 0;

  for (const item of pdfItems) {
    context.log(`[SYNC] Processing: ${item.name}`);

    const fileContent = await downloadHrKbFile(item.id);
    if (!fileContent) {
      context.warn(`[SYNC] Failed to download: ${item.name}`);
      skipped++;
      continue;
    }

    const text = await extractPdfText(fileContent.base64Content);
    if (!text) {
      context.warn(`[SYNC] No text extracted (image-based?): ${item.name}`);
      skipped++;
      continue;
    }

    const chunks = chunkText(text);
    context.log(`[SYNC] Split "${item.name}" into ${chunks.length} chunk(s)`);

    let fileIndexedChunks = 0;
    let fileSkippedChunks = 0;

    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];

      const embedding = await generateEmbedding(chunk);
      if (!embedding) {
        context.warn(`[SYNC] Failed to generate embedding for chunk ${i} of: ${item.name}`);
        fileSkippedChunks++;
        continue;
      }

      const doc: HrKbDocument = {
        id: Buffer.from(`${item.id}__chunk${i}`).toString("base64").replace(/[+/=]/g, "_"),
        fileName: item.name,
        content: chunk,
        webUrl: item.webUrl ?? "",
        driveItemId: item.id,
        chunkIndex: i,
        lastModifiedDateTime: item.lastModifiedDateTime ?? "",
        contentVector: embedding,
      };

      const uploaded = await uploadDocument(doc);
      if (uploaded) {
        fileIndexedChunks++;
      } else {
        fileSkippedChunks++;
      }
    }

    if (fileIndexedChunks > 0) {
      indexed++;
      context.log(`[SYNC] Indexed "${item.name}": ${fileIndexedChunks} chunk(s), ${fileSkippedChunks} failed`);
    } else {
      skipped++;
      context.warn(`[SYNC] All chunks failed for: ${item.name}`);
    }
  }

  return {
    status: 200,
    jsonBody: { indexed, skipped, totalPdfFiles: pdfItems.length },
  };
}

app.http("syncIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "syncIndex",
  handler: syncIndex,
});