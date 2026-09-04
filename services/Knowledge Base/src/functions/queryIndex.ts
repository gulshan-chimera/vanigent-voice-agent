// src/functions/queryIndex.ts
//
// Real semantic search against the HR KB index — embeds the caller's
// question, then finds the closest-matching document(s) by vector
// similarity rather than exact keyword overlap.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getSearchClient } from "../lib/searchIndex";
import { generateEmbedding } from "../lib/azureOpenAI";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";

export async function queryIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {

  if (!isRequestAuthorized(request)) {
    context.warn("[QUERY-INDEX] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: { query?: string };

  try {
    body = (await request.json()) as { query?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.query) {
    return { status: 400, jsonBody: { error: "query is required" } };
  }

  const queryVector = await generateEmbedding(body.query);
  if (!queryVector) {
    return { status: 502, jsonBody: { error: "Failed to embed query" } };
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    return { status: 502, jsonBody: { error: "Search client not configured" } };
  }

  try {
    const results = await searchClient.search(body.query, {
      vectorSearchOptions: {
        queries: [
          {
            kind: "vector",
            vector: queryVector,
            fields: ["contentVector"],
            kNearestNeighborsCount: 5,
          },
        ],
      },
      queryType: "semantic",
      semanticSearchOptions: {
        configurationName: "kb-semantic-config",
        errorMode: "partial",
      },
      top: 5,
    });

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
    context.error(`[QUERY-INDEX] Vector search failed: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Search failed" } };
  }
}

app.http("queryIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "queryIndex",
  handler: queryIndex,
});