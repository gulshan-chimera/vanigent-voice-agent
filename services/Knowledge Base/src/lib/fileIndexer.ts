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
  generateSheetRecords,
} from "./azureOpenAI";
import { IndexFileMessage } from "./indexQueue";
import { extractWorkbookGrid } from "./xlsxGrid";

export interface IndexFileResult {
  ok: boolean;
  pagesIndexed: number;
  pagesFailed: number;
  pagesCaptioned: number;
  pagesCaptionSkipped: number;
  chunksDeleted: number;
  /**
   * True when the file was deliberately not indexed because it has no
   * searchable content — a logo, a decorative graphic. This is a correct
   * outcome, not a failure: the worker must NOT retry or poison-queue it,
   * or the health endpoint fills up with logos and stops being useful.
   */
  skipped?: boolean;
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
// Spreadsheet path — one chunk per logical row, via LLM structure
// inference
// ---------------------------------------------------------------------

async function indexXlsxFile(
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  const grid = extractWorkbookGrid(base64Content);

  if (!grid) {
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      error: "No usable sheets found in workbook",
    };
  }

  if (grid.skipped.length > 0) {
    console.log(
      `[INDEX-FILE] "${job.itemName}": skipped ${grid.skipped.length} sheet(s) — ${grid.skipped
        .map((s) => `${s.sheetName} (${s.reason})`)
        .join(", ")}`
    );
  }

  // Generate records sheet by sheet. One model call per sheet rather than
  // per workbook: a large workbook wouldn't fit in a single prompt, and
  // isolating the calls means one malformed sheet doesn't lose the rest.
  const allRecords: { sheetName: string; title: string; content: string }[] = [];

  for (const sheet of grid.sheets) {
    const records = await generateSheetRecords(
      job.itemName,
      sheet.sheetName,
      sheet.cellLines,
      sheet.mergedRanges
    );

    if (!records) {
      console.warn(
        `[INDEX-FILE] Sheet "${sheet.sheetName}" of "${job.itemName}" produced no records — skipping that sheet.`
      );
      continue;
    }

    if (sheet.truncated) {
      console.warn(
        `[INDEX-FILE] Sheet "${sheet.sheetName}" was truncated — records cover only the first rows.`
      );
    }

    for (const record of records) {
      allRecords.push({
        sheetName: sheet.sheetName,
        title: record.title,
        content: record.content,
      });
    }
  }

  if (allRecords.length === 0) {
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      error: "No records generated from any sheet",
    };
  }

  console.log(
    `[INDEX-FILE] "${job.itemName}": ${grid.sheets.length} sheet(s) produced ${allRecords.length} record(s).`
  );

  // Embed and upload. Only embeddings here — the chat calls are already
  // done — so a larger batch is safe.
  const RECORD_BATCH_SIZE = 5;
  let pagesIndexed = 0;
  let pagesFailed = 0;

  const embedAndUpload = async (index: number): Promise<boolean> => {
    const record = allRecords[index];

    // The sheet name and record title go INTO the embedded text, not
    // just the metadata: a question like "when does Class 5 training
    // start" needs "training" to be part of what was embedded.
    const content = `[Sheet: ${record.sheetName}] ${record.title}\n\n${record.content}`;

    const embedding = await generateEmbedding(content);
    if (!embedding) {
      console.warn(
        `[INDEX-FILE] Failed to embed record ${index} ("${record.title}") of "${job.itemName}".`
      );
      return false;
    }

    return uploadDocument(buildDoc(job, index, content, embedding));
  };

  for (let start = 0; start < allRecords.length; start += RECORD_BATCH_SIZE) {
    const batch: number[] = [];
    for (let i = start; i < Math.min(start + RECORD_BATCH_SIZE, allRecords.length); i++) {
      batch.push(i);
    }

    const results = await Promise.allSettled(batch.map(embedAndUpload));

    for (const result of results) {
      if (result.status === "rejected") {
        console.error(
          `[INDEX-FILE] Record upload threw for "${job.itemName}": ${result.reason}`
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

const IMAGE_MIME_TYPES: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
};

function getExtension(fileName: string): string {
  const dot = fileName.lastIndexOf(".");
  return dot === -1 ? "" : fileName.slice(dot).toLowerCase();
}

// ---------------------------------------------------------------------
// Image path — single chunk, full transcription
// ---------------------------------------------------------------------

async function indexImageFile(
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
): Promise<IndexFileResult> {
  // Derived here rather than passed in, so this matches the FileHandler
  // signature and can live in FILE_HANDLERS with every other format.
  const mimeType = IMAGE_MIME_TYPES[getExtension(job.itemName)] ?? "image/png";

  const transcription = await generateImageTranscription(base64Content, job.itemName, mimeType);

  if (!transcription) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "Image transcription failed" };
  }

  // A logo or decorative graphic has nothing searchable in it. Indexing
  // it would add a chunk that competes with real content in every query,
  // so it's deliberately left out of the index entirely.
  //
  // NOTE: this means the file has no chunkIndex 0 document, so
  // getIndexedFileState() won't see it and every sync will re-check it.
  // That costs one vision call per sync per skipped image — acceptable
  // for a handful of files, but worth revisiting if a library fills up
  // with decorative images.
  if (isSkipCaption(transcription)) {
    console.log(
      `[INDEX-FILE] "${job.itemName}" has no informational content (logo or decorative) — not indexed.`
    );
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      pagesCaptionSkipped: 1,
      skipped: true,
      error: "Image has no informational content",
    };
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



/** Signature every format handler shares. */
type FileHandler = (
  job: IndexFileMessage,
  base64Content: string,
  chunksDeleted: number
) => Promise<IndexFileResult>;

/**
 * Extension to handler. This is the single source of truth for which
 * formats are supported — syncRunner imports the keys rather than
 * keeping its own list, so adding a format is one edit, not two.
 */
const FILE_HANDLERS: Record<string, FileHandler> = {
  ".pdf": indexPdfFile,
  ".docx": indexDocxFile,
  ".pptx": indexPptxFile,
  // SheetJS reads all three spreadsheet formats through the same API.
  ".xlsx": indexXlsxFile,
  ".xls": indexXlsxFile,
  ".csv": indexXlsxFile,
  // The transcription IS the indexed content — there's no text layer.
  ".png": indexImageFile,
  ".jpg": indexImageFile,
  ".jpeg": indexImageFile,
};

export const SUPPORTED_EXTENSIONS = Object.keys(FILE_HANDLERS);

/** Extracts the lowercased extension, including the dot. */


export async function indexOneFile(job: IndexFileMessage): Promise<IndexFileResult> {
  // On an update, remove old chunks FIRST. If a 5-chunk document becomes
  // 3 chunks, chunks 4 and 5 would otherwise linger as orphans and keep
  // surfacing stale content in search results.
  let chunksDeleted = 0;
  if (job.isUpdate) {
    chunksDeleted = await deleteFileChunks(job.itemId);
    console.log(`[INDEX-FILE] Removed ${chunksDeleted} old chunk(s) for "${job.itemName}".`);
  }

  const handler = FILE_HANDLERS[getExtension(job.itemName)];

  // Shouldn't happen — syncRunner filters on SUPPORTED_EXTENSIONS — but
  // fail loudly rather than silently indexing nothing.
  if (!handler) {
    return {
      ...EMPTY_RESULT,
      chunksDeleted,
      error: `Unsupported file type: ${job.itemName}`,
    };
  }

  // Download AFTER the handler check: no point fetching bytes we can't
  // process. (Deliberately after deleteFileChunks, though — an update
  // whose type is no longer supported should still have its stale
  // chunks removed.)
  const fileContent = await downloadDriveFile(job.driveId, job.itemId);
  if (!fileContent) {
    return { ...EMPTY_RESULT, chunksDeleted, error: "Download failed" };
  }

  return handler(job, fileContent.base64Content, chunksDeleted);
}