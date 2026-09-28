// src/functions/testConfidentiality.ts
//
// TEMPORARY diagnostic — reports which documents WOULD be marked
// restricted, without changing the index. Run this before re-indexing:
// an over-broad marker list would silently restrict documents that
// should be freely answerable.
//
// Reports occurrence counts as well as matches, because the count is
// what separates a classification LABEL (a footer repeating on every
// page) from a passing MENTION in body text. The employee handbook
// discusses confidentiality obligations at length without being
// confidential itself, and is the false positive to watch for.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { getSearchClient } from "../lib/searchIndex";
import { detectConfidentiality, teamNameForLibrary } from "../lib/confidentiality";

interface FileText {
  fileName: string;
  library: string;
  text: string;
}

interface FileVerdict {
  fileName: string;
  library: string;
  restricted: boolean;
  marker: string | null;
  occurrences: number;
}

export async function testConfidentiality(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[TEST-CONFIDENTIAL] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  const searchClient = getSearchClient();
  if (!searchClient) {
    return { status: 502, jsonBody: { error: "Search client not configured" } };
  }

  // Accumulate each file's full text before judging. Occurrence counting
  // only works across the whole document — judging chunk by chunk would
  // see a footer once per chunk and lose the distinction entirely.
  const textByFile = new Map<string, FileText>();

  try {
    let skip = 0;

    for (;;) {
      const results = await searchClient.search("*", {
        select: ["fileName", "library", "driveItemId", "content"],
        orderBy: ["id asc"],
        top: 1000,
        skip,
      });

      let count = 0;

      for await (const result of results.results) {
        count++;
        const doc = result.document;
        const existing = textByFile.get(doc.driveItemId);

        if (existing) {
          existing.text += "\n" + doc.content;
        } else {
          textByFile.set(doc.driveItemId, {
            fileName: doc.fileName,
            library: doc.library,
            text: doc.content,
          });
        }
      }

      if (count < 1000) break;
      skip += 1000;
    }
  } catch (error) {
    context.error(`[TEST-CONFIDENTIAL] Failed to read index: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Failed to read index" } };
  }

  const all: FileVerdict[] = Array.from(textByFile.values()).map((file) => {
    const verdict = detectConfidentiality(file.text);
    return {
      fileName: file.fileName,
      library: file.library,
      restricted: verdict.restricted,
      marker: verdict.marker,
      occurrences: verdict.occurrences,
    };
  });

  const restricted = all.filter((f) => f.restricted);

  // Matched a marker but fell below the threshold. This is where a
  // badly chosen threshold shows up — anything here that SHOULD be
  // restricted means the threshold is too high.
  const nearMisses = all.filter((f) => !f.restricted && f.occurrences > 0);

  const clean = all.filter((f) => f.occurrences === 0);

  context.log(
    `[TEST-CONFIDENTIAL] ${all.length} file(s): ${restricted.length} restricted, ${nearMisses.length} near miss(es).`
  );

  return {
    status: 200,
    jsonBody: {
      totalFiles: all.length,
      restrictedCount: restricted.length,

      restricted: restricted
        .sort((a, b) => b.occurrences - a.occurrences)
        .map((f) => ({
          fileName: f.fileName,
          library: f.library,
          matchedMarker: f.marker,
          occurrences: f.occurrences,
          referralTeam: teamNameForLibrary(f.library),
        })),

      nearMisses: nearMisses
        .sort((a, b) => b.occurrences - a.occurrences)
        .map((f) => ({
          fileName: f.fileName,
          library: f.library,
          matchedMarker: f.marker,
          occurrences: f.occurrences,
        })),

      unmatchedCount: clean.length,
      unmatchedFileNames: clean.map((f) => f.fileName),
    },
  };
}

app.http("testConfidentiality", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "testConfidentiality",
  handler: testConfidentiality,
});