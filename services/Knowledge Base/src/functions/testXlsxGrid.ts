// src/functions/testXlsxGrid.ts
//
// TEMPORARY diagnostic — dumps the extracted cell grid for a spreadsheet
// so we can confirm nothing is lost or malformed BEFORE building the
// record-generation step on top of it.
//
// Specifically checking: are dates readable rather than serial numbers,
// do long cell values arrive complete rather than clipped, and are
// merged ranges reported.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { downloadDriveFile } from "../lib/sharepointFiles";
import { extractWorkbookGrid } from "../lib/xlsxGrid";

export async function testXlsxGrid(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[TEST-XLSX] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: { itemId?: string; driveId?: string; sheetIndex?: number };

  try {
    body = (await request.json()) as { itemId?: string; driveId?: string; sheetIndex?: number };
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

  // Full cell dump for one sheet if asked; otherwise just the shape of
  // each sheet, since a full workbook dump can be enormous.
  const wantsFullDump = typeof body.sheetIndex === "number";

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      extracted: true,
      totalSheets: grid.totalSheets,
      sheets: grid.sheets.map((sheet, index) => ({
        index,
        sheetName: sheet.sheetName,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        nonEmptyCells: sheet.cellLines.length,
        mergedRanges: sheet.mergedRanges,
        truncated: sheet.truncated,
        // Character count matters: it's what decides whether the grid
        // fits in a prompt alongside instructions.
        gridCharCount: sheet.cellLines.join("\n").length,
        cells:
          wantsFullDump && index === body.sheetIndex
            ? sheet.cellLines
            : sheet.cellLines.slice(0, 40),
      })),
    },
  };
}

app.http("testXlsxGrid", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testXlsxGrid",
  handler: testXlsxGrid,
});