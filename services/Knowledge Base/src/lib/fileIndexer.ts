// src/lib/fileIndexer.ts
//
// Indexes ONE file. Dispatches on file type, because PDF and DOCX need
// genuinely different handling:
//
//   PDF  — has fixed pages, so one page = one chunk, and each page can
//          be rendered to an image and captioned by the vision model.
//          This keeps a step's text together with its own screenshot.
//
//   DOCX — has no page concept (Word reflows based on printer settings),
//          so there are no boundaries to align text and images against.
//          Chunks by paragraph instead, and skips vision captioning
//          entirely — an extracted image would have no reliable
//          positional link to the surrounding text, which is precisely
//          what made the PDF captions valuable.

import { downloadDriveFile } from "./sharepointFiles";
import { deleteFileChunks, uploadDocument, KbDocument } from "./searchIndex";
import { extractPageTexts } from "./pdfPageText";
import { renderPdfPagesAsImages } from "./pdfImages";
import { extractDocxText } from "./docxText";
import { chunkText } from "./textChunker";
import { generateEmbedding, generateImageCaption, SKIP_CAPTION } from "./azureOpenAI";
import { IndexFileMessage } from "./indexQueue";

export interface IndexFileResult {
  ok: boolean;
  pagesIndexed: number;
  pagesFailed: number;
  pagesCaptioned: number;
  pagesCaptionSkipped: number;
  chunksDeleted: number;
  error?: string;
}

/** Stable, collision-proof document key. */
function buildDocId(driveId: string, itemId: string, chunkIndex: number): string {
  return Buffer.from(`${driveId}__${itemId}__page${chunkIndex}`)
    .toString("base64")
    .replace(/[+/=]/g, "_");
}

/**
 * A null caption means the API call failed — NOT that the page had
 * nothing worth describing. Callers log the difference.
 */
function isSkipCaption(caption: string | null): boolean {
  if (!caption) return true;
  if (caption.includes(SKIP_CAPTION)) return true;
  if (caption === "No significant visual content.") return true;
  return false;
}

/** Builds the index document for one chunk. Shared by both paths. */
function buildDoc(
  job: IndexFileMessage,
  chunkIndex: number,
  content: string,
  embedding: number[]
): KbDocument {
  return {
    id: buildDocId(job.driveId, job.itemId, chunkIndex),
    fileName: job.itemName,
    library: job.driveName,
    driveId: job.driveId,
    driveItemId: job.itemId,
    cTag: job.cTag,
    chunkIndex,
    content,
    webUrl: job.webUrl,
    lastModifiedDateTime: job.lastModifiedDateTime,
    contentVector: embedding,
  };
}

const EMPTY_RESULT: IndexFileResult = {
  ok: false,
  pagesIndexed: 0,
  pagesFailed: 0,
  pagesCaptioned: 0,
  pagesCaptionSkipped: 0,
  chunksDeleted: 0,
};

// ---------------------------------------------------------------------
// PDF path — page-based, with vision captioning
// ---------------------------------------------------------------------

async function indexPdfFile(
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  const pageTexts = await extractPageTexts(base64Content);
  const pageImages = await renderPdfPagesAsImages(base64Content);

  if (pageTexts.length === 0 || pageImages.length === 0) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "No pages extracted" };
  }

  if (pageTexts.length !== pageImages.length) {
    console.warn(
      `[INDEX-FILE] Page count mismatch for "${job.itemName}": ${pageTexts.length} text vs ${pageImages.length} image. Using the smaller count.`
    );
  }

  const pageCount = Math.min(pageTexts.length, pageImages.length);

  // Pages are independent, so they process in parallel batches. This
  // combines with host.json's queue batchSize — total concurrency is the
  // product of the two, which must stay inside the Azure OpenAI RPM limit.
  const PAGE_BATCH_SIZE = 3;

  interface PageOutcome {
    indexed: boolean;
    captioned: boolean;
    captionSkipped: boolean;
  }

  const processPage = async (i: number): Promise<PageOutcome> => {
    const pageText = pageTexts[i];
    const pageImageBase64 = pageImages[i].toString("base64");

    const caption = await generateImageCaption(pageImageBase64, pageText);

    if (caption === null) {
      console.warn(
        `[INDEX-FILE] Caption call FAILED for page ${i + 1} of "${job.itemName}" — indexing without caption. Visual content may be missing.`
      );
    }

    const skip = isSkipCaption(caption);

    if (!skip) {
      console.log(
        `[INDEX-FILE] Page ${i + 1} of "${job.itemName}" has visual content — captioned.`
      );
    }

    const captionText = skip ? "" : `\n\n[Visual content on this page: ${caption}]`;

    // A page with neither text nor caption (e.g. a cover image) must
    // STILL be indexed — getIndexedFileState() identifies files by their
    // chunkIndex 0 document, so a missing page 0 makes the whole file
    // invisible to the diff and it re-indexes on every run.
    const combinedContent =
      `${pageText}${captionText}`.trim() ||
      `[Page ${i + 1} of ${job.itemName} — no text content]`;

    const embedding = await generateEmbedding(combinedContent);
    if (!embedding) {
      console.warn(`[INDEX-FILE] Failed to embed page ${i + 1} of "${job.itemName}".`);
      return { indexed: false, captioned: !skip, captionSkipped: skip };
    }

    const uploaded = await uploadDocument(buildDoc(job, i, combinedContent, embedding));
    return { indexed: uploaded, captioned: !skip, captionSkipped: skip };
  };

  let pagesIndexed = 0;
  let pagesFailed = 0;
  let pagesCaptioned = 0;
  let pagesCaptionSkipped = 0;

  for (let start = 0; start < pageCount; start += PAGE_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + PAGE_BATCH_SIZE, pageCount); i++) {
      batch.push(i);
    }

    // allSettled, not all: one page throwing must not abandon the rest
    // of the batch — a partially indexed file is still useful.
    const results = await Promise.allSettled(batch.map(processPage));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Page processing threw for "${job.itemName}": ${result.reason}`
        );
        pagesFailed++;
        continue;
      }

      const outcome = result.value;
      if (outcome.indexed) pagesIndexed++;
      else pagesFailed++;
      if (outcome.captioned) pagesCaptioned++;
      if (outcome.captionSkipped) pagesCaptionSkipped++;
    }
  }

  return {
    ok: pagesIndexed > 0,
    pagesIndexed,
    pagesFailed,
    pagesCaptioned,
    pagesCaptionSkipped,
    chunksDeleted,
  };
}

// ---------------------------------------------------------------------
// DOCX path — paragraph-based, no captioning
// ---------------------------------------------------------------------

async function indexDocxFile(
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  const text = await extractDocxText(base64Content);

  if (!text) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "No text extracted from DOCX" };
  }

  const chunks = chunkText(text);

  if (chunks.length === 0) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "Text produced zero chunks" };
  }

  console.log(`[INDEX-FILE] Split "${job.itemName}" into ${chunks.length} chunk(s).`);

  // No vision calls here, so the only rate-limited call per chunk is the
  // embedding — far cheaper than the PDF path. Still batched to avoid
  // firing all chunks of a long document at once.
  const CHUNK_BATCH_SIZE = 5;

  let pagesIndexed = 0;
  let pagesFailed = 0;

  const processChunk = async (i: number): Promise<boolean> => {
    const embedding = await generateEmbedding(chunks[i]);
    if (!embedding) {
      console.warn(`[INDEX-FILE] Failed to embed chunk ${i + 1} of "${job.itemName}".`);
      return false;
    }
    return uploadDocument(buildDoc(job, i, chunks[i], embedding));
  };

  for (let start = 0; start < chunks.length; start += CHUNK_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + CHUNK_BATCH_SIZE, chunks.length); i++) {
      batch.push(i);
    }

    const results = await Promise.allSettled(batch.map(processChunk));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Chunk processing threw for "${job.itemName}": ${result.reason}`
        );
        pagesFailed++;
        continue;
      }
      if (result.value) pagesIndexed++;
      else pagesFailed++;
    }
  }

  return {
    ok: pagesIndexed > 0,
    pagesIndexed,
    pagesFailed,
    pagesCaptioned: 0,
    pagesCaptionSkipped: 0,
    chunksDeleted,
  };
}

// ---------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------

export async function indexOneFile(job: IndexFileMessage): Promise<IndexFileResult> {
  // On an update, remove old chunks FIRST. If a 5-chunk document becomes
  // 3 chunks, chunks 4 and 5 would otherwise linger as orphans and keep
  // surfacing stale content in search results.
  let chunksDeleted = 0;
  if (job.isUpdate) {
    chunksDeleted = await deleteFileChunks(job.itemId);
    console.log(`[INDEX-FILE] Removed ${chunksDeleted} old chunk(s) for "${job.itemName}".`);
  }

  const fileContent = await downloadDriveFile(job.driveId, job.itemId);
  if (!fileContent) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "Download failed" };
  }

  const name = job.itemName.toLowerCase();

  if (name.endsWith(".pdf")) {
    return indexPdfFile(job, fileContent.base64Content, chunksDeleted);
  }

  if (name.endsWith(".docx")) {
    return indexDocxFile(job, fileContent.base64Content, chunksDeleted);
  }

  // Shouldn't happen — syncRunner filters to supported types — but fail
  // loudly rather than silently indexing nothing.
  return {
    ...EMPTY_RESULT,
    chunksDeleted,
    error: `Unsupported file type: ${job.itemName}`,
  };
}