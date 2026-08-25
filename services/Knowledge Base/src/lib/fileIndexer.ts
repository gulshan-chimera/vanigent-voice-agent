// src/lib/fileIndexer.ts
//
// Indexes ONE file. Every supported format ends up going through the
// SAME page-based, vision-captioned pipeline (indexPdfFile) — PDF
// natively, and DOCX/PPTX by first asking Microsoft Graph to convert
// them to PDF server-side (downloadDriveFileAsPdf, `?format=pdf`).
//
// That conversion step is the only way any of these formats gets vision
// captioning at all: neither Word nor PowerPoint documents have a pure-JS
// rendering path of their own (no pdfjs-dist equivalent), so without
// converting first there'd be nothing to hand the vision model. Once
// converted, a Word page or a PowerPoint slide becomes one PDF page —
// same chunk-per-page model, same captioning, no format-specific
// indexing logic needed beyond the conversion call.
//
// Note for DOCX specifically: page boundaries in the converted PDF come
// from Word's print layout (page size/margins/fonts), not from anything
// the author intended as a content boundary — unlike PPTX slides, which
// are already an author-defined unit. In practice this is still a
// reasonable, consistent chunk size; it just isn't a "meaningful"
// boundary the way a slide or a PDF page in a fixed-layout document is.

import { downloadDriveFile, downloadDriveFileAsPdf } from "./sharepointFiles";
import { deleteFileChunks, uploadDocument, KbDocument } from "./searchIndex";
import { extractPageTexts } from "./pdfPageText";
import { renderPdfPagesAsImages } from "./pdfImages";
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
// Router
// ---------------------------------------------------------------------

// Formats with no native rendering path of their own — converted to PDF
// by Graph first, then handled by the exact same PDF pipeline.
const CONVERT_TO_PDF_EXTENSIONS = [".docx", ".pptx"];

export async function indexOneFile(job: IndexFileMessage): Promise<IndexFileResult> {
  // On an update, remove old chunks FIRST. If a 5-chunk document becomes
  // 3 chunks, chunks 4 and 5 would otherwise linger as orphans and keep
  // surfacing stale content in search results.
  let chunksDeleted = 0;
  if (job.isUpdate) {
    chunksDeleted = await deleteFileChunks(job.itemId);
    console.log(`[INDEX-FILE] Removed ${chunksDeleted} old chunk(s) for "${job.itemName}".`);
  }

  const name = job.itemName.toLowerCase();

  if (CONVERT_TO_PDF_EXTENSIONS.some((ext) => name.endsWith(ext))) {
    const pdfContent = await downloadDriveFileAsPdf(job.driveId, job.itemId);
    if (!pdfContent) {
      return { ...EMPTY_RESULT, chunksDeleted, error: "PDF conversion failed" };
    }
    return indexPdfFile(job, pdfContent, chunksDeleted);
  }

  if (name.endsWith(".pdf")) {
    const fileContent = await downloadDriveFile(job.driveId, job.itemId);
    if (!fileContent) {
      return { ...EMPTY_RESULT, chunksDeleted, error: "Download failed" };
    }
    return indexPdfFile(job, fileContent.base64Content, chunksDeleted);
  }

  // Shouldn't happen — syncRunner filters to supported types — but fail
  // loudly rather than silently indexing nothing.
  return {
    ...EMPTY_RESULT,
    chunksDeleted,
    error: `Unsupported file type: ${job.itemName}`,
  };
}