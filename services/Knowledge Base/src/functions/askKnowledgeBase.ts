// src/functions/askKnowledgeBase.ts
//
// The real end-to-end endpoint: embeds the caller's question, retrieves
// the most relevant chunks from the HR KB index, and generates a
// grounded, spoken-style answer. This is what VAPI's tool call will
// hit once we wire it up.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getSearchClient } from "../lib/searchIndex";
import { generateEmbedding, generateAnswer } from "../lib/azureOpenAI";

export async function askKnowledgeBase(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: { question?: string };

  try {
    body = (await request.json()) as { question?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.question) {
    return { status: 400, jsonBody: { error: "question is required" } };
  }

  const queryVector = await generateEmbedding(body.question);
  if (!queryVector) {
    return { status: 502, jsonBody: { error: "Failed to embed question" } };
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    return { status: 502, jsonBody: { error: "Search client not configured" } };
  }

  let contextChunks: { fileName: string; content: string }[] = [];

  try {
    const results = await searchClient.search("*", {
      vectorSearchOptions: {
        queries: [
          {
            kind: "vector",
            vector: queryVector,
            fields: ["contentVector"],
            kNearestNeighborsCount: 3,
          },
        ],
      },
    });

    for await (const result of results.results) {
      contextChunks.push({
        fileName: result.document.fileName,
        content: result.document.content,
      });
    }
  } catch (error) {
    context.error(`[ASK-KB] Search failed: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Search failed" } };
  }

  if (contextChunks.length === 0) {
    return {
      status: 200,
      jsonBody: {
        answer: "I don't have that information right now, but I can note it for the team to follow up.",
      },
    };
  }

  const answer = await generateAnswer(body.question, contextChunks);

  if (!answer) {
    return { status: 502, jsonBody: { error: "Failed to generate answer" } };
  }

  context.log(`[ASK-KB] Q: "${body.question}" -> A: "${answer}"`);

  return { status: 200, jsonBody: { answer } };
}

app.http("askKnowledgeBase", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "askKnowledgeBase",
  handler: askKnowledgeBase,
});