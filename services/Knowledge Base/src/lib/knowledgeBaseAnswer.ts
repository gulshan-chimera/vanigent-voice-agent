// src/lib/knowledgeBaseAnswer.ts
//
// Core logic shared by both the internal testing endpoint (askKnowledgeBase)
// and the VAPI-facing tool-call endpoint: embed the question, search the
// index, generate a grounded answer.

import { getSearchClient } from "./searchIndex";
import {
  generateEmbedding,
  generateAnswer,
  AnswerContext,
} from "./azureOpenAI";
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

export interface AnswerResult {
  answer: string;
  departments: string[];
}

/**
 * Runs an async operation with a maximum deadline.
 *
 * If the operation does not complete within the specified time,
 * the fallback value is returned.
 */
async function withDeadline<T>(
  work: Promise<T>,
  ms: number,
  fallback: T
): Promise<T> {
  let timer: NodeJS.Timeout | undefined;

  const deadline = new Promise<T>((resolve) => {
    timer = setTimeout(() => resolve(fallback), ms);
  });

  try {
    return await Promise.race([work, deadline]);
  } finally {
    // Without this, the process holds an open timer for the full
    // duration even when the work finished first.
    if (timer) {
      clearTimeout(timer);
    }
  }
}

/**
 * Core Knowledge Base answer generation logic.
 *
 * This function contains the complete embedding -> search -> context ->
 * answer generation -> confidentiality flow.
 *
 * It intentionally does not apply the overall timeout. The timeout is
 * handled by answerQuestion() so that the entire operation is covered.
 */
export async function answerQuestionInner(
  question: string
): Promise<AnswerResult> {
  const queryVector = await generateEmbedding(question);

  if (!queryVector) {
    console.error("[KB-ANSWER] Failed to embed question.");

    return {
      answer: FALLBACK_ANSWER,
      departments: [],
    };
  }

  const searchClient = getSearchClient();

  if (!searchClient) {
    console.error("[KB-ANSWER] Search client not configured.");

    return {
      answer: FALLBACK_ANSWER,
      departments: [],
    };
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
          `[KB-ANSWER] Skipping chunk ${chunkIdx}: score=${result.score?.toFixed(
            4
          )} file="${result.document.fileName}" (below threshold)`
        );

        chunkIdx++;
        continue;
      }

      console.log(
        `[KB-ANSWER] Chunk ${chunkIdx}: score=${result.score?.toFixed(
          4
        )} file="${result.document.fileName}" content="${result.document.content.slice(
          0,
          80
        )}..."`
      );

      contextChunks.push({
        fileName: result.document.fileName,
        library: result.document.library,
        content: result.document.content,
        restricted: result.document.restricted ?? false,
        score: result.score,
      });

      chunkIdx++;
    }
  } catch (error) {
    console.error(
      `[KB-ANSWER] Search failed: ${(error as Error).message}`
    );

    return {
      answer: FALLBACK_ANSWER,
      departments: [],
    };
  }

  console.log(
    `[KB-ANSWER] Retrieved ${contextChunks.length} relevant chunk(s) for question: "${question}"`
  );

  if (contextChunks.length === 0) {
    console.warn(
      "[KB-ANSWER] Zero chunks returned from search — returning fallback."
    );

    return {
      answer: FALLBACK_ANSWER,
      departments: [],
    };
  }

  // Derive the department from the highest-scoring chunk only.
  // This ensures "Systems Affected" shows only the most relevant department.
  const topChunk = contextChunks.reduce((best, c) =>
    (c.score ?? 0) > (best.score ?? 0) ? c : best
  );
  const topDept = teamNameForLibrary(topChunk.library);

  // Log each chunk's department and score for debugging
  for (const c of contextChunks) {
    const dept = teamNameForLibrary(c.library);
    const isTop = c === topChunk;
    console.log(
      `[KB-ANSWER] Dept check: "${dept}" | score=${c.score?.toFixed(4) ?? "N/A"} | file="${c.fileName}" | ${isTop ? "✅ TOP (selected)" : "⏭️ SKIPPED"}`
    );
  }

  const departments = topDept.length > 0 ? [topDept] : [];

  console.log(`[KB-ANSWER] Final department: ${departments.join(", ") || "(none)"}`);

  const answer = await generateAnswer(question, contextChunks);

  if (!answer) {
    console.warn(
      "[KB-ANSWER] generateAnswer returned null — returning fallback."
    );

    return {
      answer: FALLBACK_ANSWER,
      departments,
    };
  }

  const restrictedChunks = contextChunks.filter((c) => c.restricted);

  if (restrictedChunks.length > 0 && containsCurrencyFigure(answer)) {
    const team = teamNameForLibrary(restrictedChunks[0].library);

    console.warn(
      `[KB-ANSWER] BLOCKED: answer contained a currency figure with restricted sources in context. Question: "${question}"`
    );

    return {
      answer: `That's not something I can share over the phone. Please contact your ${team} team for that information.`,
      departments,
    };
  }

  return {
    answer,
    departments,
  };
}

/**
 * Public Knowledge Base answer function.
 *
 * Applies the 12-second overall deadline to the complete
 * embedding -> search -> answer generation flow.
 */
export async function answerQuestion(
  question: string
): Promise<AnswerResult> {
  const started = Date.now();

  const timeoutResult: AnswerResult = {
    answer: TIMEOUT_ANSWER,
    departments: [],
  };

  const result = await withDeadline(
    answerQuestionInner(question),
    ANSWER_DEADLINE_MS,
    timeoutResult
  );

  if (result.answer === TIMEOUT_ANSWER) {
    console.error(
      `[KB-ANSWER] DEADLINE EXCEEDED after ${Date.now() - started
      }ms — returned fallback. Question: "${question}"`
    );
  }

  return result;
}