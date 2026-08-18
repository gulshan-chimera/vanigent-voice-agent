// src/functions/testImageExtraction.ts
//
// TEMPORARY diagnostic endpoint. Extracts images from a given HR KB file
// and returns the raw result, so we can see the real library output
// before building the full captioning pipeline on top of it.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { downloadHrKbFile } from "../lib/sharepointFiles";
import {extractPdfImages} from "../lib/pdfImages"

export async function testImageExtraction(
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


  const images = await extractPdfImages(fileContent.base64Content);

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