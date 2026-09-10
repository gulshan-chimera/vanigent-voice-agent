// src/functions/testSheetRecords.ts
//
// TEMPORARY diagnostic — shows the records generated from a spreadsheet
// WITHOUT indexing them, so we can inspect whether the model got the
// structure right before building the sync path on top.
//
// Specifically checking: are separate tables kept separate, are merged
// headers resolved into meaningful labels, and are row-specific notes
// attached to the right record.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { downloadDriveFile } from "../lib/sharepointFiles";
import { extractWorkbookGrid } from "../lib/xlsxGrid";
import { generateSheetRecords } from "../lib/azureOpenAI";

export async function testSheetRecords(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[TEST-SHEET-RECORDS] Rejected request — invalid or missing bearer token.");
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

  const grid = extractWorkbookGrid(fileContent.base64Content);

  if (!grid) {
    return {
      status: 200,
      jsonBody: { fileName: fileContent.fileName, extracted: false },
    };
  }

  const sheetResults = [];

  for (const sheet of grid.sheets) {
    const records = await generateSheetRecords(
      fileContent.fileName,
      sheet.sheetName,
      sheet.cellLines,
      sheet.mergedRanges
    );

    sheetResults.push({
      sheetName: sheet.sheetName,
      truncated: sheet.truncated,
      gridCharCount: sheet.cellLines.join("\n").length,
      recordCount: records?.length ?? 0,
      records: records ?? [],
    });
  }

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      totalSheets: grid.totalSheets,
      skippedSheets: grid.skipped,
      sheets: sheetResults,
    },
  };
}

app.http("testSheetRecords", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testSheetRecords",
  handler: testSheetRecords,
});