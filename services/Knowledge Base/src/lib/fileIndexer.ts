// src/lib/fileIndexer.ts
//
// Indexes ONE file. Dispatches on file type:
//
//   PDF  — has fixed pages, so one page = one chunk, and each page can
//          be rendered to an image and captioned by the vision model.
//          This keeps a step's text together with its own screenshot.
//
//   DOCX — extracted directly from the .docx XML via mammoth
//          (docxText.ts), chunked by paragraph (textChunker.ts), no
//          vision captioning. This is a direct reversion from routing
//          DOCX through Graph's PDF conversion + the PDF pipeline:
//          mammoth reads structured paragraph/table XML straight from
//          the document, so a table row's cells stay in document order;
//          pdfjs's text extraction from the Word→PDF-converted render
//          instead reconstructs order from on-page text positions, which
//          scrambled multi-column table rows badly enough that answers
//          about specific table entries were unreliable. Losing vision
//          captioning for DOCX is the trade-off for getting table
//          content back to something the model can actually read.
//
//   PPTX — extracted directly from the .pptx XML (pptxText.ts), one
//          chunk per SLIDE (an author-defined boundary, the same role a
//          PDF page plays), no vision captioning. Same rationale as
//          DOCX: this used to go through Graph's PDF conversion for a
//          uniform pipeline with PDF, but pdfjs's text extraction from
//          the converted PDF reconstructs reading order from on-page
//          position, which scrambled multi-column slide layouts (a step
//          list beside a screenshot) badly enough that some content
//          became unretrievable. Reading the slide XML directly avoids
//          that, at the cost of no captioning for now.
//
//   IMAGE (.png/.jpg/.jpeg) — a standalone image IS the entire document;
//          there's no text layer to extract at all, so this is the one
//          format where vision is the ONLY source of content rather
//          than a supplement to it. Uses generateImageTranscription, a
//          dedicated prompt from generateImageCaption's — that one is
//          tuned to skip content already covered by separately
//          extracted text and is heavily biased toward returning
//          nothing, which would be wrong here. One chunk for the whole
//          file (no page/slide concept for a single image).

import { downloadDriveFile } from "./sharepointFiles";
import { deleteFileChunks, uploadDocument, KbDocument } from "./searchIndex";
import { extractPageTexts } from "./pdfPageText";
import { renderPdfPagesAsImages } from "./pdfImages";
import { extractDocxText } from "./docxText";
import { extractPptxSlideTexts } from "./pptxText";
import { chunkText } from "./textChunker";
import {
  generateEmbedding,
  generateImageCaption,
  generateImageTranscription,
  SKIP_CAPTION,
} from "./azureOpenAI";
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

const SPLIT_THRESHOLD_CHARS = 1200;

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
  const PAGE_BATCH_SIZE = 3;

  /** One unit of content ready to embed, with the page it came from. */
  interface PendingChunk {
    page: number; // 1-based, for logging only
    content: string;
  }

  interface PageOutcome {
    chunks: PendingChunk[];
    captioned: boolean;
    captionSkipped: boolean;
  }

  // Phase 1: caption and assemble content per page, in parallel batches.
  // Splitting happens here but embedding does NOT — chunk indexes have to
  // be assigned in document order, which can't be done until every page
  // is known.
  const preparePage = async (i: number): Promise<PageOutcome> => {
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

    const trimmedText = pageText.trim();

    // A captioned page is never split: the whole point of page-level
    // chunking was keeping a step's text together with the caption
    // describing its screenshot. Splitting would break that pairing.
    if (!skip) {
      const combined = `${trimmedText}\n\n[Visual content on this page: ${caption}]`.trim();
      return {
        chunks: [{ page: i + 1, content: combined }],
        captioned: true,
        captionSkipped: false,
      };
    }

    // An empty page still needs a placeholder — getIndexedFileState()
    // finds files by their chunkIndex 0 document, so a file whose first
    // unit is missing becomes invisible to the diff and re-indexes on
    // every run.
    if (trimmedText.length === 0) {
      return {
        chunks: [
          { page: i + 1, content: `[Page ${i + 1} of ${job.itemName} — no text content]` },
        ],
        captioned: false,
        captionSkipped: true,
      };
    }

    // Short pages stay whole — splitting them would only fragment
    // something already focused.
    if (trimmedText.length <= SPLIT_THRESHOLD_CHARS) {
      return {
        chunks: [{ page: i + 1, content: trimmedText }],
        captioned: false,
        captionSkipped: true,
      };
    }

    // Dense text page: split on paragraph boundaries so a specific
    // clause carries its own vector instead of being averaged away.
    const parts = chunkText(trimmedText);

    if (parts.length === 0) {
      return {
        chunks: [{ page: i + 1, content: trimmedText }],
        captioned: false,
        captionSkipped: true,
      };
    }

    if (parts.length > 1) {
      console.log(
        `[INDEX-FILE] Page ${i + 1} of "${job.itemName}" split into ${parts.length} chunk(s) (${trimmedText.length} chars).`
      );
    }

    return {
      chunks: parts.map((content) => ({ page: i + 1, content })),
      captioned: false,
      captionSkipped: true,
    };
  };

  const pageOutcomes: PageOutcome[] = [];
  let pagesCaptioned = 0;
  let pagesCaptionSkipped = 0;
  let pagesFailed = 0;

  for (let start = 0; start < pageCount; start += PAGE_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + PAGE_BATCH_SIZE, pageCount); i++) {
      batch.push(i);
    }

    const results = await Promise.allSettled(batch.map(preparePage));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Page preparation threw for "${job.itemName}": ${result.reason}`
        );
        pagesFailed++;
        continue;
      }

      pageOutcomes.push(result.value);
      if (result.value.captioned) pagesCaptioned++;
      if (result.value.captionSkipped) pagesCaptionSkipped++;
    }
  }

  // Flatten in document order, assigning sequential chunk indexes.
  // NOTE: chunkIndex is no longer the page number — it's a sequence
  // number, since one page can now produce several chunks. Nothing
  // depends on it being a page number: deleteFileChunks filters on
  // driveItemId, and getIndexedFileState only needs chunk 0 to exist.
  const pending: PendingChunk[] = [];
  for (const outcome of pageOutcomes) {
    pending.push(...outcome.chunks);
  }

  if (pending.length === 0) {
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      pagesCaptioned,
      pagesCaptionSkipped,
      error: "No indexable content produced",
    };
  }

  console.log(
    `[INDEX-FILE] "${job.itemName}": ${pageCount} page(s) produced ${pending.length} chunk(s).`
  );

  // Phase 2: embed and upload. Vision calls are done, so this is only
  // embeddings — cheaper per call, so a larger batch is safe.
  const EMBED_BATCH_SIZE = 5;
  let pagesIndexed = 0;

  const embedAndUpload = async (chunkIndex: number): Promise<boolean> => {
    const chunk = pending[chunkIndex];

    const embedding = await generateEmbedding(chunk.content);
    if (!embedding) {
      console.warn(
        `[INDEX-FILE] Failed to embed chunk ${chunkIndex} (page ${chunk.page}) of "${job.itemName}".`
      );
      return false;
    }

    return uploadDocument(buildDoc(job, chunkIndex, chunk.content, embedding));
  };

  for (let start = 0; start < pending.length; start += EMBED_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + EMBED_BATCH_SIZE, pending.length); i++) {
      batch.push(i);
    }

    const results = await Promise.allSettled(batch.map(embedAndUpload));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Chunk upload threw for "${job.itemName}": ${result.reason}`
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
// PPTX path — slide-based, no captioning
// ---------------------------------------------------------------------

async function indexPptxFile(
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  const slideTexts = await extractPptxSlideTexts(base64Content);

  if (slideTexts.length === 0) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "No slides extracted" };
  }

  const SLIDE_BATCH_SIZE = 5;
  let pagesIndexed = 0;
  let pagesFailed = 0;

  const processSlide = async (i: number): Promise<boolean> => {
    // A slide with no text (e.g. a title slide that's all imagery) must
    // still be indexed — getIndexedFileState() identifies an indexed
    // file by its chunkIndex 0 document, so an empty slide 1 would make
    // the whole file invisible to the diff and it would re-index on
    // every run.
    const content =
      slideTexts[i].trim() || `[Slide ${i + 1} of ${job.itemName} — no text content]`;

    const embedding = await generateEmbedding(content);
    if (!embedding) {
      console.warn(`[INDEX-FILE] Failed to embed slide ${i + 1} of "${job.itemName}".`);
      return false;
    }
    return uploadDocument(buildDoc(job, i, content, embedding));
  };

  for (let start = 0; start < slideTexts.length; start += SLIDE_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + SLIDE_BATCH_SIZE, slideTexts.length); i++) {
      batch.push(i);
    }

    const results = await Promise.allSettled(batch.map(processSlide));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Slide processing threw for "${job.itemName}": ${result.reason}`
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
// Image path — single chunk, full transcription
// ---------------------------------------------------------------------

async function indexImageFile(
  job: IndexFileMessage,
  base64Content: string,
  mimeType: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  const transcription = await generateImageTranscription(base64Content, job.itemName, mimeType);

  if (!transcription) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "Image transcription failed" };
  }

  const embedding = await generateEmbedding(transcription);
  if (!embedding) {
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      pagesFailed: 1,
      error: "Failed to embed image transcription",
    };
  }

  const uploaded = await uploadDocument(buildDoc(job, 0, transcription, embedding));

  return {
    ok: uploaded,
    pagesIndexed: uploaded ? 1 : 0,
    pagesFailed: uploaded ? 0 : 1,
    pagesCaptioned: uploaded ? 1 : 0,
    pagesCaptionSkipped: 0,
    chunksDeleted,
  };
}

// ---------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------

const IMAGE_MIME_TYPES: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
};

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

  if (name.endsWith(".pptx")) {
    return indexPptxFile(job, fileContent.base64Content, chunksDeleted);
  }

  const imageExt = Object.keys(IMAGE_MIME_TYPES).find((ext) => name.endsWith(ext));
  if (imageExt) {
    return indexImageFile(job, fileContent.base64Content, IMAGE_MIME_TYPES[imageExt], chunksDeleted);
  }

  // Shouldn't happen — syncRunner filters to supported types — but fail
  // loudly rather than silently indexing nothing.
  return {
    ...EMPTY_RESULT,
    chunksDeleted,
    error: `Unsupported file type: ${job.itemName}`,
  };
}
