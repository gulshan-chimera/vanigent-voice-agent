// src/functions/syncIndex.ts
//
// Incremental sync across all allowlisted SharePoint document libraries.
//
// Diffs what's in SharePoint against what's already in the index:
//   NEW       -> index it
//   UPDATED   -> delete that file's old chunks, then re-index
//   DELETED   -> delete that file's chunks only
//   UNCHANGED -> skip entirely (no download, no vision, no embedding)
//
// Change detection uses the driveItem's cTag, which changes on CONTENT
// edits only — so renames and metadata edits don't trigger a costly
// re-caption of every page.
//
// Pass ?force=true to wipe the whole index and rebuild from scratch
// (only needed after a schema change or a caption-prompt change).

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import {
  listSiteDrives,
  selectDrives,          // ← was filterAllowedDrives
  listAllFilesInDrive,
  downloadDriveFile,
} from "../lib/sharepointFiles";
import {
  ensureIndexExists,
  clearEntireIndex,
  getIndexedFileState,
  deleteFileChunks,
  uploadDocument,
  KbDocument,
  IndexedFileState,
} from "../lib/searchIndex";
import { extractPageTexts } from "../lib/pdfPageText";
import { renderPdfPagesAsImages } from "../lib/pdfImages";
import { generateEmbedding, generateImageCaption, SKIP_CAPTION } from "../lib/azureOpenAI";
import { DriveItem, SiteDrive } from "../types/knowledgeBase";

interface SyncSummary {
  selectionMode: string;
  librariesProcessed: string[];
  librariesExcluded: { name: string; reason: string }[];
  filesNew: number;
  filesUpdated: number;
  filesUnchanged: number;
  filesDeleted: number;
  filesFailed: number;
  pagesIndexed: number;
  pagesFailed: number;
  pagesCaptioned: number;
  pagesCaptionSkipped: number;
  chunksDeleted: number;
}

/** Builds a stable, collision-proof document key. */
function buildDocId(driveId: string, itemId: string, pageIndex: number): string {
  return Buffer.from(`${driveId}__${itemId}__page${pageIndex}`)
    .toString("base64")
    .replace(/[+/=]/g, "_");
}

/**
 * Decides whether a caption carries real information. The model is
 * instructed to return SKIP_CAPTION when a page has nothing but page
 * furniture (logos, headers, footers, formatting) — we check with
 * `includes` rather than equality in case it wraps the token in a
 * sentence, and still honour the older phrasing for safety.
 */
function isSkipCaption(caption: string | null): boolean {
  if (!caption) return true;
  if (caption.includes(SKIP_CAPTION)) return true;
  if (caption === "No significant visual content.") return true;
  return false;
}

/**
 * Downloads one PDF, renders + captions + embeds each page, and uploads
 * each page as its own chunk. Returns per-page counts.
 */
async function indexFile(
  drive: SiteDrive,
  item: DriveItem,
  context: InvocationContext
): Promise<{
  pagesIndexed: number;
  pagesFailed: number;
  pagesCaptioned: number;
  pagesCaptionSkipped: number;
  ok: boolean;
}> {
  const empty = {
    pagesIndexed: 0,
    pagesFailed: 0,
    pagesCaptioned: 0,
    pagesCaptionSkipped: 0,
    ok: false,
  };

  const fileContent = await downloadDriveFile(drive.id, item.id);
  if (!fileContent) {
    context.warn(`[SYNC] Failed to download: ${item.name}`);
    return empty;
  }

  const pageTexts = await extractPageTexts(fileContent.base64Content);
  const pageImages = await renderPdfPagesAsImages(fileContent.base64Content);

  if (pageTexts.length === 0 || pageImages.length === 0) {
    context.warn(`[SYNC] No pages extracted from: ${item.name}`);
    return empty;
  }

  if (pageTexts.length !== pageImages.length) {
    context.warn(
      `[SYNC] Page count mismatch for "${item.name}": ${pageTexts.length} text vs ${pageImages.length} image. Using the smaller count.`
    );
  }

  const pageCount = Math.min(pageTexts.length, pageImages.length);
  let pagesIndexed = 0;
  let pagesFailed = 0;
  let pagesCaptioned = 0;
  let pagesCaptionSkipped = 0;

    for (let i = 0; i < pageCount; i++) {
    const pageText = pageTexts[i];
    const pageImageBase64 = pageImages[i].toString("base64");

    const caption = await generateImageCaption(pageImageBase64, pageText);

    // A null caption means the API call FAILED — not that the page had
    // nothing worth describing. Without this warning the two are
    // indistinguishable, and a transient 500 silently drops real visual
    // content from the index.
    if (caption === null) {
      context.warn(
        `[SYNC] Caption call FAILED for page ${i + 1} of "${item.name}" — indexing without caption. Visual content may be missing.`
      );
    }

    const skip = isSkipCaption(caption);

    if (skip) {
      pagesCaptionSkipped++;
    } else {
      pagesCaptioned++;
      context.log(`[SYNC] Page ${i + 1} of "${item.name}" has visual content — captioned.`);
    }

    const captionText = skip ? "" : `\n\n[Visual content on this page: ${caption}]`;

    // A page with neither text nor caption (e.g. a cover image) must
    // STILL be indexed. getIndexedFileState() identifies indexed files by
    // their chunkIndex 0 document — so if page 1 is skipped, the whole
    // file becomes invisible to the diff and gets re-indexed on every
    // single run. Use a placeholder rather than dropping the page.
    const combinedContent =
      `${pageText}${captionText}`.trim() ||
      `[Page ${i + 1} of ${item.name} — no text content]`;

    const embedding = await generateEmbedding(combinedContent);
    if (!embedding) {
      context.warn(`[SYNC] Failed to embed page ${i + 1} of: ${item.name}`);
      pagesFailed++;
      continue;
    }

    const doc: KbDocument = {
      id: buildDocId(drive.id, item.id, i),
      fileName: item.name,
      library: drive.name,
      driveId: drive.id,
      driveItemId: item.id,
      cTag: item.cTag ?? "",
      chunkIndex: i,
      content: combinedContent,
      webUrl: item.webUrl ?? "",
      lastModifiedDateTime: item.lastModifiedDateTime ?? "",
      contentVector: embedding,
    };

    const uploaded = await uploadDocument(doc);
    if (uploaded) {
      pagesIndexed++;
    } else {
      pagesFailed++;
    }
  }

  return {
    pagesIndexed,
    pagesFailed,
    pagesCaptioned,
    pagesCaptionSkipped,
    ok: pagesIndexed > 0,
  };
}

/** Thrown for setup failures so both triggers can report them their own way. */
export class SyncSetupError extends Error {
  status: number;
  details?: Record<string, unknown>;

  constructor(message: string, status: number, details?: Record<string, unknown>) {
    super(message);
    this.status = status;
    this.details = details;
  }
}

/**
 * Runs one incremental sync pass. Shared by the manual HTTP trigger and
 * the nightly timer trigger — neither the auth check nor the `force`
 * query flag belong here, since the timer has no HttpRequest to read
 * them from; the timer always runs a normal (non-force) pass.
 */
export async function runSync(context: InvocationContext, force: boolean): Promise<SyncSummary> {
  const indexReady = await ensureIndexExists();
  if (!indexReady) {
    throw new SyncSetupError("Failed to prepare search index", 502);
  }

  const allDrives = await listSiteDrives();
  if (!allDrives) {
    throw new SyncSetupError("Failed to list SharePoint libraries", 502);
  }

  const selection = selectDrives(allDrives);
  const drives = selection.included;

  if (drives.length === 0) {
    throw new SyncSetupError(
      "No libraries selected — check KB_LIBRARY_ALLOWLIST / KB_LIBRARY_PATTERN.",
      400,
      { availableLibraries: allDrives.map((d) => d.name) }
    );
  }

  if (force) {
    context.warn("[SYNC] FORCE mode — wiping the entire index before rebuilding.");
    await clearEntireIndex();
  }

  // What's already indexed (empty after a force wipe).
  const indexedState = force
    ? new Map<string, IndexedFileState>()
    : await getIndexedFileState();

  if (indexedState === null) {
    throw new SyncSetupError("Failed to read current index state", 502);
  }

  const summary: SyncSummary = {
    selectionMode: selection.mode,
    librariesProcessed: [],
    librariesExcluded: selection.excluded,
    filesNew: 0,
    filesUpdated: 0,
    filesUnchanged: 0,
    filesDeleted: 0,
    filesFailed: 0,
    pagesIndexed: 0,
    pagesFailed: 0,
    pagesCaptioned: 0,
    pagesCaptionSkipped: 0,
    chunksDeleted: 0,
  };

  // Every item ID we saw in SharePoint this run — anything indexed but
  // NOT in here has been deleted at source.
  const seenItemIds = new Set<string>();

  // Libraries we successfully enumerated. If a library listing fails we
  // must NOT treat its indexed files as deleted, or a transient Graph
  // error would silently wipe that library from the index.
  const enumeratedDriveIds = new Set<string>();

  for (const drive of drives) {
    context.log(`[SYNC] === Library: ${drive.name} ===`);

    const files = await listAllFilesInDrive(drive.id);
    if (files === null) {
      context.error(`[SYNC] Could not enumerate "${drive.name}" — skipping it this run.`);
      continue;
    }

    enumeratedDriveIds.add(drive.id);
    summary.librariesProcessed.push(drive.name);

    const pdfItems = files.filter((f) => f.name.toLowerCase().endsWith(".pdf"));
    context.log(`[SYNC] "${drive.name}": ${pdfItems.length} PDF(s) found.`);

    for (const item of pdfItems) {
      seenItemIds.add(item.id);

      const prior = indexedState.get(item.id);
      const currentCTag = item.cTag ?? "";

      if (prior && prior.cTag === currentCTag && currentCTag !== "") {
        summary.filesUnchanged++;
        continue;
      }

      const isUpdate = Boolean(prior);

      if (isUpdate) {
        context.log(`[SYNC] UPDATED: ${item.name} — removing old chunks first.`);
        summary.chunksDeleted += await deleteFileChunks(item.id);
      } else {
        context.log(`[SYNC] NEW: ${item.name}`);
      }

      const result = await indexFile(drive, item, context);
      summary.pagesIndexed += result.pagesIndexed;
      summary.pagesFailed += result.pagesFailed;
      summary.pagesCaptioned += result.pagesCaptioned;
      summary.pagesCaptionSkipped += result.pagesCaptionSkipped;

      if (!result.ok) {
        summary.filesFailed++;
      } else if (isUpdate) {
        summary.filesUpdated++;
      } else {
        summary.filesNew++;
      }
    }
  }

  // Deletions: indexed files that no longer exist at source.
  for (const [itemId, info] of indexedState) {
    if (seenItemIds.has(itemId)) continue;
    if (!enumeratedDriveIds.has(info.driveId)) continue; // library wasn't checked

    context.log(`[SYNC] DELETED at source: ${info.fileName} (${info.library})`);
    summary.chunksDeleted += await deleteFileChunks(itemId);
    summary.filesDeleted++;
  }

  context.log(`[SYNC] Done. ${JSON.stringify(summary)}`);
  return summary;
}

export async function syncIndex(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[SYNC] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  const force = request.query.get("force") === "true";

  try {
    const summary = await runSync(context, force);
    return { status: 200, jsonBody: summary };
  } catch (error) {
    if (error instanceof SyncSetupError) {
      return {
        status: error.status,
        jsonBody: { error: error.message, ...(error.details ?? {}) },
      };
    }
    context.error(`[SYNC] Unexpected failure: ${(error as Error).message}`);
    return { status: 502, jsonBody: { error: "Sync failed unexpectedly" } };
  }
}

app.http("syncIndex", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "syncIndex",
  handler: syncIndex,
});