// src/lib/knowledgeBaseAnswer.ts
//
// Core logic shared by both the internal testing endpoint (askKnowledgeBase)
// and the VAPI-facing tool-call endpoint: embed the question, search the
// index, generate a grounded answer.

import { getSearchClient } from "./searchIndex";
import { generateEmbedding, generateAnswer, AnswerContext } from "./azureOpenAI";
import { teamNameForLibrary } from "./confidentiality";


const CURRENCY_PATTERN =
  /(\$\s?\d[\d,]*(\.\d{2})?)|(\b\d[\d,]*(\.\d{2})?\s?(dollars|usd)\b)|(\busd\s?\d[\d,]*)/i;

function containsCurrencyFigure(text: string): boolean {
  return CURRENCY_PATTERN.test(text);
}

const FALLBACK_ANSWER =
  "I don't have that information right now, but I can note it for the team to follow up.";

const ANSWER_DEADLINE_MS = 12000;

const TIMEOUT_ANSWER =
  "I'm having trouble looking that up right now — I can note it for the team to follow up.";

async function withDeadline<T>(work: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: NodeJS.Timeout | undefined;

  const deadline = new Promise<T>((resolve) => {
    timer = setTimeout(() => resolve(fallback), ms);
  });

  try {
    return await Promise.race([work, deadline]);
  } finally {
    // Without this, the process holds an open timer for the full
    // duration even when the work finished first.
    if (timer) clearTimeout(timer);
  }
}

export async function answerQuestionInner(question: string): Promise<string> {
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

  const contextChunks: AnswerContext[] = [];

  try {
    const results = await searchClient.search(question, {
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
      queryType: "semantic",
      semanticSearchOptions: {
        configurationName: "kb-semantic-config",
        errorMode: "partial",
      },
      top: 5,
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
        restricted: result.document.restricted ?? false,
        library: result.document.library,
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
    return FALLBACK_ANSWER;
  }
  const restrictedChunks = contextChunks.filter((c) => c.restricted);

  if (restrictedChunks.length > 0 && containsCurrencyFigure(answer)) {
    const team = teamNameForLibrary(restrictedChunks[0].library);

    console.warn(
      `[KB-ANSWER] BLOCKED: answer contained a currency figure with restricted sources in context. Question: "${question}"`
    );

    return `That's not something I can share over the phone. Please contact your ${team} team for that information.`;
  }

  return answer;
}

export async function answerQuestion(question: string): Promise<string> {
  const started = Date.now();

  const answer = await withDeadline(
    answerQuestionInner(question),
    ANSWER_DEADLINE_MS,
    TIMEOUT_ANSWER
  );

  if (answer === TIMEOUT_ANSWER) {
    console.error(
      `[KB-ANSWER] DEADLINE EXCEEDED after ${Date.now() - started}ms — returned fallback. Question: "${question}"`
    );
  }

  return answer;
}