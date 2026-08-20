// src/functions/testPageText.ts
//
// TEMPORARY diagnostic endpoint — confirms per-page text extraction
// works and aligns with the page count we already confirmed via image
// rendering, before building the full combine+caption+index pipeline.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadHrKbFile } from "../lib/sharepointFiles";
import { extractPageTexts } from "../lib/pdfPageText";

export async function testPageText(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: { itemId?: string };

  try {
    body = (await request.json()) as { itemId?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.itemId) {
    return { status: 400, jsonBody: { error: "itemId is required" } };
  }

  const fileContent = await downloadHrKbFile(body.itemId);
  if (!fileContent) {
    return { status: 404, jsonBody: { error: "File not found" } };
  }

  const pageTexts = await extractPageTexts(fileContent.base64Content);

  return {
    status: 200,
    jsonBody: {
      fileName: fileContent.fileName,
      pageCount: pageTexts.length,
      pages: pageTexts.map((text, i) => ({
        page: i + 1,
        charCount: text.length,
        preview: text.slice(0, 150),
      })),
    },
  };
}

app.http("testPageText", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testPageText",
  handler: testPageText,
});