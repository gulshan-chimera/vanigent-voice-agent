/**
 * scripts/deltaSync.ts
 *
 * Smart delta-sync: only indexes NEW or MODIFIED documents from SharePoint.
 * Removes documents from the index that have been DELETED from SharePoint.
 * Tracks sync state in a local JSON file (.sync-state.json) to avoid
 * re-downloading and re-indexing unchanged documents.
 *
 * Run with: npm run delta-sync
 */

import fs from "node:fs";
import path from "node:path";
import pdf from "pdf-parse";
import { loadConfig, requireGraphCredentials } from "../src/config.js";
import { KnowledgeBaseSearchClient } from "../src/search/searchClient.js";
import type {
  IndexDocument,
  FolderGroupMap,
  RedactionResult,
} from "../src/search/types.js";

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------

const config = loadConfig();
requireGraphCredentials(config);

const FOLDER_GROUP_MAP: FolderGroupMap = {
  General_Policies: ["2f996595-97dc-45dd-a85a-0d20759e1682"],
  Manager_Policies: [
    "2f996595-97dc-45dd-a85a-0d20759e1682",
    "70205b81-db45-4fa4-92b7-7fb87a89c73f",
  ],
  Sales_Policies: [
    "2f996595-97dc-45dd-a85a-0d20759e1682",
    "0941d8fd-91c2-415e-8b74-e81cf00d08ea",
  ],
};

const SYNC_STATE_PATH = path.resolve(__dirname, "../../.sync-state.json");

// ---------------------------------------------------------------------------
// SYNC STATE TYPES
// ---------------------------------------------------------------------------

interface SyncedFile {
  /** SharePoint item ID */
  itemId: string;
  /** File name */
  name: string;
  /** Folder it belongs to */
  folder: string;
  /** ISO timestamp of last modification in SharePoint */
  lastModified: string;
  /** SharePoint eTag — changes whenever the file content changes */
  eTag: string;
  /** File size in bytes */
  size: number;
  /** Document ID stored in Azure AI Search index */
  indexDocId: string;
}

interface SyncState {
  /** ISO timestamp of when the last sync completed */
  lastSyncedAt: string;
  /** Map of "folder__filename" -> SyncedFile */
  files: Record<string, SyncedFile>;
}

// ---------------------------------------------------------------------------
// SYNC STATE PERSISTENCE
// ---------------------------------------------------------------------------

function loadSyncState(): SyncState {
  if (!fs.existsSync(SYNC_STATE_PATH)) {
    console.log("[SYNC-STATE] No previous state found. Starting fresh.");
    return { lastSyncedAt: "", files: {} };
  }

  try {
    const raw = fs.readFileSync(SYNC_STATE_PATH, "utf-8");
    const state = JSON.parse(raw) as SyncState;
    const fileCount = Object.keys(state.files).length;
    console.log(
      `[SYNC-STATE] Loaded previous state — ${fileCount} file(s), last synced: ${state.lastSyncedAt}`
    );
    return state;
  } catch {
    console.warn("[SYNC-STATE] Could not parse state file. Starting fresh.");
    return { lastSyncedAt: "", files: {} };
  }
}

function saveSyncState(state: SyncState): void {
  fs.writeFileSync(SYNC_STATE_PATH, JSON.stringify(state, null, 2), "utf-8");
  console.log(`[SYNC-STATE] State saved to ${SYNC_STATE_PATH}`);
}

// ---------------------------------------------------------------------------
// MICROSOFT GRAPH AUTH
// ---------------------------------------------------------------------------

async function getAccessToken(): Promise<string> {
  const res = await fetch(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: config.clientId,
        client_secret: config.clientSecret,
        scope: "https://graph.microsoft.com/.default",
      }),
    }
  );
  const data = await res.json();
  if (!data.access_token) {
    throw new Error(`Graph auth failed: ${JSON.stringify(data)}`);
  }
  return data.access_token;
}

// ---------------------------------------------------------------------------
// SHAREPOINT GRAPH HELPERS
// ---------------------------------------------------------------------------

async function getSiteAndDrive(
  token: string
): Promise<{ siteId: string; driveId: string }> {
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${config.spHostname}:${config.spSitePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const siteData = await siteRes.json();
  if (siteData.error) {
    throw new Error(`getSite failed: ${siteData.error.message}`);
  }

  const driveRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const driveData = await driveRes.json();
  if (driveData.error) {
    throw new Error(`getDrive failed: ${driveData.error.message}`);
  }

  const drive =
    driveData.value.find(
      (d: any) => d.name === "Documents" || d.driveType === "documentLibrary"
    ) || driveData.value[0];

  console.log(`[SHAREPOINT] Site: ${siteData.id}`);
  console.log(`[SHAREPOINT] Drive: "${drive.name}" (${drive.id})`);
  return { siteId: siteData.id, driveId: drive.id };
}

interface SharePointFileDetail {
  id: string;
  name: string;
  lastModifiedDateTime: string;
  eTag: string;
  size: number;
  file?: unknown;
}

async function listFilesInFolder(
  token: string,
  driveId: string,
  folderName: string
): Promise<SharePointFileDetail[]> {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderName}:/children?$select=id,name,lastModifiedDateTime,eTag,size,file`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();
  if (data.error) {
    console.warn(
      `[SHAREPOINT] Could not list "${folderName}": ${data.error.message}`
    );
    return [];
  }
  return (data.value || []).filter(
    (item: SharePointFileDetail) =>
      item.name?.toLowerCase().endsWith(".pdf") && item.file
  );
}

async function downloadFile(
  token: string,
  driveId: string,
  itemId: string
): Promise<Buffer> {
  const metaRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const metaData = await metaRes.json();
  if (metaData.error) {
    throw new Error(`Download metadata failed: ${metaData.error.message}`);
  }
  const downloadUrl = metaData["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) throw new Error("No download URL returned from Graph API");
  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) throw new Error(`Download failed: HTTP ${fileRes.status}`);
  return Buffer.from(await fileRes.arrayBuffer());
}

// ---------------------------------------------------------------------------
// PII REDACTION
// ---------------------------------------------------------------------------

const PII_PATTERNS: Array<{ name: string; pattern: RegExp }> = [
  { name: "Email", pattern: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g },
  { name: "Phone", pattern: /(\+?(\d[\s\-.]?){10,14}\d)/g },
  { name: "Aadhaar", pattern: /\b\d{4}\s?\d{4}\s?\d{4}\b/g },
  { name: "PAN", pattern: /\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b/g },
  { name: "Passport", pattern: /\b[A-Z]{1,2}[0-9]{6,7}\b/g },
  { name: "CreditCard", pattern: /\b(?:\d[ \-]?){15,16}\b/g },
  { name: "SSN", pattern: /\b\d{3}[- ]?\d{2}[- ]?\d{4}\b/g },
  { name: "DOB", pattern: /\b(DOB|Date of Birth|D\.O\.B)[:\s]+\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/gi },
  { name: "BankAccount", pattern: /\b\d{9,18}\b/g },
  { name: "IFSC", pattern: /\b[A-Z]{4}0[A-Z0-9]{6}\b/g },
];

function redactPII(text: string): RedactionResult {
  let redacted = text;
  const found: string[] = [];
  for (const { name, pattern } of PII_PATTERNS) {
    const matches = redacted.match(pattern);
    if (matches && matches.length > 0) {
      found.push(`${name} (${matches.length})`);
      redacted = redacted.replace(pattern, "[REDACTED]");
    }
  }
  return { redacted, found };
}

// ---------------------------------------------------------------------------
// DOCUMENT ID HELPER
// ---------------------------------------------------------------------------

function makeDocId(folder: string, fileName: string): string {
  return Buffer.from(`${folder}__${fileName}`)
    .toString("base64")
    .replace(/[+/=]/g, "_");
}

function makeFileKey(folder: string, fileName: string): string {
  return `${folder}__${fileName}`;
}

// ---------------------------------------------------------------------------
// DELTA DETECTION
// ---------------------------------------------------------------------------

interface DeltaResult {
  newFiles: Array<{ file: SharePointFileDetail; folder: string }>;
  modifiedFiles: Array<{ file: SharePointFileDetail; folder: string }>;
  deletedKeys: string[];
  unchangedCount: number;
}

function computeDelta(
  prevState: SyncState,
  currentFiles: Map<string, { file: SharePointFileDetail; folder: string }>
): DeltaResult {
  const newFiles: DeltaResult["newFiles"] = [];
  const modifiedFiles: DeltaResult["modifiedFiles"] = [];
  let unchangedCount = 0;

  // Check each current file against previous state
  for (const [key, { file, folder }] of currentFiles) {
    const prev = prevState.files[key];

    if (!prev) {
      // File didn't exist before → NEW
      newFiles.push({ file, folder });
    } else if (
      prev.lastModified !== file.lastModifiedDateTime ||
      prev.eTag !== file.eTag ||
      prev.size !== file.size
    ) {
      // File existed but has changed → MODIFIED
      modifiedFiles.push({ file, folder });
    } else {
      // File is identical → SKIP
      unchangedCount++;
    }
  }

  // Find deleted files: exist in previous state but NOT in current
  const deletedKeys: string[] = [];
  for (const key of Object.keys(prevState.files)) {
    if (!currentFiles.has(key)) {
      deletedKeys.push(key);
    }
  }

  return { newFiles, modifiedFiles, deletedKeys, unchangedCount };
}

// ---------------------------------------------------------------------------
// INDEX A SINGLE FILE
// ---------------------------------------------------------------------------

async function indexSingleFile(
  searchClient: KnowledgeBaseSearchClient,
  token: string,
  driveId: string,
  file: SharePointFileDetail,
  folder: string,
  groupIds: string[]
): Promise<SyncedFile | null> {
  const docId = makeDocId(folder, file.name);

  try {
    const buffer = await downloadFile(token, driveId, file.id);
    console.log(`    [DOWNLOAD] ${buffer.length} bytes`);

    const pdfData = await pdf(buffer);
    const rawText = pdfData.text;
    if (!rawText || rawText.trim().length === 0) {
      console.log(`    [SKIP] No text extracted`);
      return null;
    }
    console.log(`    [EXTRACT] ${rawText.length} characters`);

    const { redacted, found } = redactPII(rawText);
    if (found.length > 0) console.log(`    [PII REDACTED] ${found.join(" | ")}`);

    const document: IndexDocument = {
      id: docId,
      title: file.name,
      content: redacted,
      folder,
      GroupIds: groupIds,
    };

    const result = await searchClient.uploadDocuments([document]);
    const status = result.results[0];
    if (status.succeeded) {
      console.log(`    [INDEXED] ✓ "${file.name}"`);
      return {
        itemId: file.id,
        name: file.name,
        folder,
        lastModified: file.lastModifiedDateTime,
        eTag: file.eTag,
        size: file.size,
        indexDocId: docId,
      };
    } else {
      console.error(`    [FAILED] "${file.name}" — ${status.errorMessage}`);
      return null;
    }
  } catch (err: any) {
    console.error(`    [ERROR] ${file.name}: ${err.message}`);
    return null;
  }
}

// ---------------------------------------------------------------------------
// MAIN — DELTA SYNC
// ---------------------------------------------------------------------------

async function deltaSync(): Promise<void> {
  const searchClient = new KnowledgeBaseSearchClient(config);

  // ── Load previous sync state ──
  const prevState = loadSyncState();

  // ── Authenticate ──
  console.log("\n[AUTH] Getting Microsoft Graph access token...");
  const token = await getAccessToken();
  console.log("[AUTH] Token acquired.");

  const { driveId } = await getSiteAndDrive(token);

  // ── Scan all folders and build the "current" file map ──
  console.log("\n[SCAN] Scanning SharePoint folders for current files...");
  const currentFiles = new Map<
    string,
    { file: SharePointFileDetail; folder: string }
  >();

  for (const folder of Object.keys(FOLDER_GROUP_MAP)) {
    const files = await listFilesInFolder(token, driveId, folder);
    console.log(`  [FOLDER] ${folder}/ — ${files.length} PDF(s)`);
    for (const file of files) {
      const key = makeFileKey(folder, file.name);
      currentFiles.set(key, { file, folder });
    }
  }

  // ── Compute delta ──
  const delta = computeDelta(prevState, currentFiles);

  console.log("\n═══════════════════════════════════");
  console.log("  DELTA SYNC SUMMARY");
  console.log("═══════════════════════════════════");
  console.log(`  New files:       ${delta.newFiles.length}`);
  console.log(`  Modified files:  ${delta.modifiedFiles.length}`);
  console.log(`  Deleted files:   ${delta.deletedKeys.length}`);
  console.log(`  Unchanged:       ${delta.unchangedCount} (skipped)`);
  console.log("═══════════════════════════════════");

  // If nothing changed, exit early
  if (
    delta.newFiles.length === 0 &&
    delta.modifiedFiles.length === 0 &&
    delta.deletedKeys.length === 0
  ) {
    console.log("\n[SYNC] Everything is up to date. Nothing to do.");
    return;
  }

  // ── Build new sync state (start with unchanged files) ──
  const newState: SyncState = {
    lastSyncedAt: new Date().toISOString(),
    files: {},
  };

  // Carry forward unchanged files
  for (const [key, { file, folder }] of currentFiles) {
    const prev = prevState.files[key];
    if (
      prev &&
      prev.lastModified === file.lastModifiedDateTime &&
      prev.eTag === file.eTag &&
      prev.size === file.size
    ) {
      newState.files[key] = prev;
    }
  }

  // ── Process NEW files ──
  if (delta.newFiles.length > 0) {
    console.log(`\n[NEW] Indexing ${delta.newFiles.length} new file(s)...`);
    for (const { file, folder } of delta.newFiles) {
      const key = makeFileKey(folder, file.name);
      const groupIds = FOLDER_GROUP_MAP[folder];
      console.log(`\n  [NEW] ${folder}/${file.name}`);

      const synced = await indexSingleFile(
        searchClient, token, driveId, file, folder, groupIds
      );
      if (synced) newState.files[key] = synced;
    }
  }

  // ── Process MODIFIED files ──
  if (delta.modifiedFiles.length > 0) {
    console.log(
      `\n[MODIFIED] Re-indexing ${delta.modifiedFiles.length} modified file(s)...`
    );
    for (const { file, folder } of delta.modifiedFiles) {
      const key = makeFileKey(folder, file.name);
      const groupIds = FOLDER_GROUP_MAP[folder];
      const prev = prevState.files[key];
      console.log(
        `\n  [MODIFIED] ${folder}/${file.name} (last modified: ${prev?.lastModified} → ${file.lastModifiedDateTime})`
      );

      const synced = await indexSingleFile(
        searchClient, token, driveId, file, folder, groupIds
      );
      if (synced) newState.files[key] = synced;
    }
  }

  // ── Process DELETED files ──
  if (delta.deletedKeys.length > 0) {
    console.log(
      `\n[DELETED] Removing ${delta.deletedKeys.length} deleted file(s) from index...`
    );
    const idsToDelete: string[] = [];

    for (const key of delta.deletedKeys) {
      const prev = prevState.files[key];
      if (prev) {
        console.log(`  [DELETE] ${prev.folder}/${prev.name}`);
        idsToDelete.push(prev.indexDocId);
      }
    }

    if (idsToDelete.length > 0) {
      const deleteResult = await searchClient.deleteDocuments(idsToDelete);
      const failed = deleteResult.results.filter((r: any) => !r.succeeded);
      if (failed.length > 0) {
        console.warn(`  [WARN] ${failed.length} deletion(s) failed.`);
      } else {
        console.log(
          `  [DELETED] ✓ ${idsToDelete.length} document(s) removed from index.`
        );
      }
    }
  }

  // ── Save updated sync state ──
  saveSyncState(newState);

  // ── Final summary ──
  const totalFiles = Object.keys(newState.files).length;
  console.log(`\n─────────────────────────────────`);
  console.log(`[DONE] Delta sync complete.`);
  console.log(`  Indexed (new):      ${delta.newFiles.length}`);
  console.log(`  Indexed (modified): ${delta.modifiedFiles.length}`);
  console.log(`  Deleted:            ${delta.deletedKeys.length}`);
  console.log(`  Unchanged (skipped):${delta.unchangedCount}`);
  console.log(`  Total in index:     ${totalFiles}`);
  console.log(
    `  Index: "${config.searchIndexName}" on ${config.searchServiceName}.search.windows.net`
  );
}

deltaSync().catch((error) => {
  console.error("[FATAL]", error);
  process.exit(1);
});
