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

    const doc: HrKbDocument = {
      id: Buffer.from(item.id).toString("base64").replace(/[+/=]/g, "_"),
      fileName: item.name,
      content: text,
      webUrl: item.webUrl ?? "",
      driveItemId: item.id,
      lastModifiedDateTime: item.lastModifiedDateTime ?? "",
    };

    const uploaded = await uploadDocument(doc);
    if (uploaded) {
      indexed++;
      context.log(`[SYNC] Indexed: ${item.name}`);
    } else {
      skipped++;
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