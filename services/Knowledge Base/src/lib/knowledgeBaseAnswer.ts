// src/lib/knowledgeBaseAnswer.ts
//
// Core logic shared by both the internal testing endpoint (askKnowledgeBase)
// and the VAPI-facing tool-call endpoint: embed the question, search the
// index, generate a grounded answer.

import { getSearchClient } from "./searchIndex";
import { generateEmbedding, generateAnswer } from "./azureOpenAI";

const FALLBACK_ANSWER =
  "I don't have that information right now, but I can note it for the team to follow up.";

export async function answerQuestion(question: string): Promise<string> {
  const queryVector = await generateEmbedding(question);
  if (!queryVector) {
    console.error("[KB-ANSWER] Failed to embed question.");
    return FALLBACK_ANSWER;
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    console.error("[KB-ANSWER] Search client not configured.");
    return FALLBACK_ANSWER;
  }

  const contextChunks: { fileName: string; content: string }[] = [];

  try {
    const results = await searchClient.search("*", {
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
    });

    for await (const result of results.results) {
      contextChunks.push({
        fileName: result.document.fileName,
        content: result.document.content,
      });
    }
  } catch (error) {
    console.error(`[KB-ANSWER] Search failed: ${(error as Error).message}`);
    return FALLBACK_ANSWER;
  }

  console.log(`[KB-ANSWER] Retrieved ${contextChunks.length} chunk(s) for question: "${question}"`);

  if (contextChunks.length === 0) {
    console.warn("[KB-ANSWER] Zero chunks returned from search — returning fallback.");
    return FALLBACK_ANSWER;
  }

  const answer = await generateAnswer(question, contextChunks);

  if (!answer) {
    console.warn("[KB-ANSWER] generateAnswer returned null — returning fallback.");
  }

  return answer ?? FALLBACK_ANSWER;
}