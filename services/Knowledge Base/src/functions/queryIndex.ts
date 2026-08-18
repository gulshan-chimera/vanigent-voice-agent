// src/functions/queryIndex.ts
//
// Simple keyword search against the HR KB index — for testing the
// prototype before wiring it into VAPI.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getSearchClient } from "../lib/searchIndex";

export async function queryIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: { query?: string };

  try {
    body = (await request.json()) as { query?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.query) {
    return { status: 400, jsonBody: { error: "query is required" } };
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    return { status: 502, jsonBody: { error: "Search client not configured" } };
  }

  try {
    const results = await searchClient.search(body.query, { top: 5, includeTotalCount: true });
    const matches = [];

    for await (const result of results.results) {
      matches.push({
        fileName: result.document.fileName,
        score: result.score,
        contentSnippet: result.document.content.slice(0, 300),
        webUrl: result.document.webUrl,
      });
    }

    return { status: 200, jsonBody: { matches } };
  } catch (error) {
    context.error(`[QUERY-INDEX] Search failed: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Search failed" } };
  }
}

app.http("queryIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "queryIndex",
  handler: queryIndex,
});