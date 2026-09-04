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
    // Hybrid search: combines vector similarity with BM25 full-text
    // keyword matching. Pure vector search fails when questions include
    // contextual words (e.g. "in Vanitrack") that shift the embedding
    // away from the relevant content. BM25 compensates by boosting
    // chunks whose text contains the actual question words.
    const results = await searchClient.search(question, {
      searchFields: ["content", "fileName"],
      vectorSearchOptions: {
        queries: [
          {
            kind: "vector",
            vector: queryVector,
            fields: ["contentVector"],
            kNearestNeighborsCount: 15,
          },
        ],
      },
      top: 8,
    });

    let chunkIdx = 0;
    for await (const result of results.results) {
      // With hybrid search (RRF fusion), scores are typically in the
      // 0.01–0.03 range. Filter out truly unrelated chunks so the LLM
      // doesn't see irrelevant noise.
      if (result.score !== undefined && result.score < 0.005) {
        console.log(
          `[KB-ANSWER] Skipping chunk ${chunkIdx}: score=${result.score?.toFixed(4)} file="${result.document.fileName}" (below threshold)`
        );
        chunkIdx++;
        continue;
      }

      console.log(
        `[KB-ANSWER] Chunk ${chunkIdx}: score=${result.score?.toFixed(4)} file="${result.document.fileName}" content="${result.document.content.slice(0, 80)}..."`
      );

      contextChunks.push({
        fileName: result.document.fileName,
        content: result.document.content,
      });
      chunkIdx++;
    }
  } catch (error) {
    console.error(`[KB-ANSWER] Search failed: ${(error as Error).message}`);
    return FALLBACK_ANSWER;
  }

  console.log(`[KB-ANSWER] Retrieved ${contextChunks.length} relevant chunk(s) for question: "${question}"`);

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