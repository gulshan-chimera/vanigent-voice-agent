// src/lib/syncRunner.ts
//
// The sync logic itself, extracted so both the HTTP endpoint
// (syncIndex) and the scheduled timer (syncTimer) run identical code.
//
// Diffs SharePoint against the index:
//   NEW       -> enqueue an index job
//   UPDATED   -> enqueue an index job (worker deletes old chunks first)
//   DELETED   -> delete that file's chunks inline
//   UNCHANGED -> skip entirely
//
// Change detection uses the driveItem's cTag, which changes on CONTENT
// edits only — renames and metadata edits don't trigger a re-index.

import { InvocationContext } from "@azure/functions";
import { listSiteDrives, selectDrives, listAllFilesInDrive } from "./sharepointFiles";
import {
  ensureIndexExists,
  clearEntireIndex,
  getIndexedFileState,
  deleteFileChunks,
  IndexedFileState,
} from "./searchIndex";
import { ensureQueueExists, enqueueIndexJob, IndexFileMessage } from "./indexQueue";

// File types the indexer can handle. Adding a type here makes the next
// sync pick it up automatically — including the scheduled run.
const SUPPORTED_EXTENSIONS = [".pdf", ".docx"];

function isSupported(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return SUPPORTED_EXTENSIONS.some((ext) => lower.endsWith(ext));
}

export interface SyncSummary {
  trigger: string;
  selectionMode: string;
  librariesProcessed: string[];
  librariesExcluded: { name: string; reason: string }[];
  filesQueuedNew: number;
  filesQueuedUpdated: number;
  filesUnchanged: number;
  filesDeleted: number;
  filesQueueFailed: number;
  chunksDeleted: number;
  unsupportedSkipped: { library: string; count: number; types: string[] }[];
  note: string;
}

export interface SyncFailure {
  error: string;
  availableLibraries?: string[];
}

export type SyncResult =
  | { ok: true; summary: SyncSummary }
  | { ok: false; failure: SyncFailure };

export async function runSync(
  options: { force: boolean; trigger: string },
  context: InvocationContext
): Promise<SyncResult> {
  const { force, trigger } = options;

  const indexReady = await ensureIndexExists();
  if (!indexReady) {
    return { ok: false, failure: { error: "Failed to prepare search index" } };
  }

  const queueReady = await ensureQueueExists();
  if (!queueReady) {
    return { ok: false, failure: { error: "Failed to prepare the indexing queue" } };
  }

  const allDrives = await listSiteDrives();
  if (!allDrives) {
    return { ok: false, failure: { error: "Failed to list SharePoint libraries" } };
  }

  const selection = selectDrives(allDrives);
  const drives = selection.included;

  if (drives.length === 0) {
    return {
      ok: false,
      failure: {
        error: "No libraries selected — check KB_LIBRARY_ALLOWLIST / KB_LIBRARY_PATTERN.",
        availableLibraries: allDrives.map((d) => d.name),
      },
    };
  }

  if (force) {
    context.warn("[SYNC] FORCE mode — wiping the entire index before rebuilding.");
    await clearEntireIndex();
  }

  const indexedState = force
    ? new Map<string, IndexedFileState>()
    : await getIndexedFileState();

  if (indexedState === null) {
    return { ok: false, failure: { error: "Failed to read current index state" } };
  }

  const summary: SyncSummary = {
    trigger,
    selectionMode: selection.mode,
    librariesProcessed: [],
    librariesExcluded: selection.excluded,
    filesQueuedNew: 0,
    filesQueuedUpdated: 0,
    filesUnchanged: 0,
    filesDeleted: 0,
    filesQueueFailed: 0,
    chunksDeleted: 0,
    unsupportedSkipped: [],
    note: "Files are QUEUED, not yet indexed. Check /api/queueStatus for progress and failures.",
  };

  // Every item ID seen in SharePoint this run. Anything indexed but not
  // in here has been deleted at source.
  const seenItemIds = new Set<string>();

  // Libraries we successfully enumerated. If a listing fails we must NOT
  // treat its indexed files as deleted, or one transient Graph error
  // would silently wipe that library from the index.
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

    const supportedItems = files.filter((f) => isSupported(f.name));
    const unsupportedItems = files.filter((f) => !isSupported(f.name));

    // Report what we're ignoring rather than dropping it silently — this
    // is how we discovered that two libraries had no indexable content
    // at all.
    if (unsupportedItems.length > 0) {
      const types = [
        ...new Set(
          unsupportedItems.map((f) => {
            const dot = f.name.lastIndexOf(".");
            return dot === -1 ? "(no extension)" : f.name.slice(dot).toLowerCase();
          })
        ),
      ];
      context.warn(
        `[SYNC] "${drive.name}": skipping ${unsupportedItems.length} unsupported file(s) — types: ${types.join(", ")}`
      );
      summary.unsupportedSkipped.push({
        library: drive.name,
        count: unsupportedItems.length,
        types,
      });
    }

    context.log(`[SYNC] "${drive.name}": ${supportedItems.length} indexable file(s) found.`);

    for (const item of supportedItems) {
      seenItemIds.add(item.id);

      const prior = indexedState.get(item.id);
      const currentCTag = item.cTag ?? "";

      // An empty cTag means Graph didn't give us one — re-index rather
      // than assume unchanged.
      if (prior && prior.cTag === currentCTag && currentCTag !== "") {
        summary.filesUnchanged++;
        continue;
      }

      const isUpdate = Boolean(prior);

      const job: IndexFileMessage = {
        driveId: drive.id,
        driveName: drive.name,
        itemId: item.id,
        itemName: item.name,
        cTag: currentCTag,
        webUrl: item.webUrl ?? "",
        lastModifiedDateTime: item.lastModifiedDateTime ?? "",
        isUpdate,
      };

      const queued = await enqueueIndexJob(job);

      if (!queued) {
        context.error(`[SYNC] Failed to queue: ${item.name}`);
        summary.filesQueueFailed++;
        continue;
      }

      if (isUpdate) {
        context.log(`[SYNC] QUEUED (update): ${item.name}`);
        summary.filesQueuedUpdated++;
      } else {
        context.log(`[SYNC] QUEUED (new): ${item.name}`);
        summary.filesQueuedNew++;
      }
    }
  }

  // Deletions run inline: fast (one filtered delete, no AI calls) and
  // doing them here means the index never lingers with content whose
  // source file is gone.
  for (const [itemId, info] of indexedState) {
    if (seenItemIds.has(itemId)) continue;
    if (!enumeratedDriveIds.has(info.driveId)) continue; // library wasn't checked

    context.log(`[SYNC] DELETED at source: ${info.fileName} (${info.library})`);
    summary.chunksDeleted += await deleteFileChunks(itemId);
    summary.filesDeleted++;
  }

  context.log(`[SYNC] Enqueue complete (${trigger}). ${JSON.stringify(summary)}`);
  return { ok: true, summary };
}