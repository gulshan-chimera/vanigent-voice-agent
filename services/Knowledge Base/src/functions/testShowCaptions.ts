// src/functions/testShowCaptions.ts
//
// TEMPORARY diagnostic — dumps the stored caption text for each page of
// a file, so we can judge whether captions on text-heavy pages actually
// add retrievable information or just describe page furniture.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getSearchClient } from "../lib/searchIndex";
import {isRequestAuthorized} from "../lib/verifyWebhookAuth"

const CAPTION_MARKER = "[Visual content on this page:";

export async function testShowCaptions(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
    if (!isRequestAuthorized(request)) {
    context.warn("[SHOW-CAPTIONS] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  let body: { itemId?: string };

  try {
    body = (await request.json()) as { itemId?: string };
  } catch {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.itemId) {
    return { status: 400, jsonBody: { error: "itemId is required" } };
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    return { status: 502, jsonBody: { error: "Search client not configured" } };
  }

  try {
    const results = await searchClient.search("*", {
      filter: `driveItemId eq '${body.itemId.replace(/'/g, "''")}'`,
      select: ["chunkIndex", "fileName", "content"],
      orderBy: ["chunkIndex asc"],
      top: 1000,
    });

    const captioned: { page: number; caption: string }[] = [];
    let fileName = "";

    for await (const result of results.results) {
      fileName = result.document.fileName;
      const content = result.document.content;
      const markerAt = content.indexOf(CAPTION_MARKER);

      if (markerAt === -1) continue;

      captioned.push({
        page: result.document.chunkIndex + 1,
        caption: content.slice(markerAt + CAPTION_MARKER.length).replace(/\]$/, "").trim(),
      });
    }

    return {
      status: 200,
      jsonBody: { fileName, captionedPageCount: captioned.length, captions: captioned },
    };
  } catch (error) {
    context.error(`[SHOW-CAPTIONS] Failed: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Failed to read index" } };
  }
}

app.http("testShowCaptions", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "testShowCaptions",
  handler: testShowCaptions,
});