/**
 * scripts/syncIndex.ts
 *
 * Clears the Azure AI Search index and re-indexes all current PDFs from SharePoint.
 * Run with: npm run sync-index
 */

import pdf from "pdf-parse";
import { loadConfig, requireGraphCredentials } from "../src/config.js";
import { KnowledgeBaseSearchClient } from "../src/search/searchClient.js";
import type { IndexDocument, FolderGroupMap, RedactionResult } from "../src/search/types.js";

const config = loadConfig();
requireGraphCredentials(config);

const FOLDER_GROUP_MAP: FolderGroupMap = {
  General_Policies: ["2f996595-97dc-45dd-a85a-0d20759e1682"],
  Manager_Policies: ["70205b81-db45-4fa4-92b7-7fb87a89c73f"],
  Sales_Policies: ["0941d8fd-91c2-415e-8b74-e81cf00d08ea"],
};

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
  if (!data.access_token) throw new Error(`Auth failed: ${JSON.stringify(data)}`);
  return data.access_token;
}

async function getSiteAndDrive(token: string): Promise<{ siteId: string; driveId: string }> {
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${config.spHostname}:${config.spSitePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const siteData = await siteRes.json();
  if (siteData.error) throw new Error(`getSite failed: ${siteData.error.message}`);

  const driveRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const driveData = await driveRes.json();
  if (driveData.error) throw new Error(`getDrive failed: ${driveData.error.message}`);

  const drive = driveData.value.find(
    (d: any) => d.name === "Documents" || d.driveType === "documentLibrary"
  ) || driveData.value[0];
  return { siteId: siteData.id, driveId: drive.id };
}

interface SharePointFile { id: string; name: string; file?: unknown; }

async function listFilesInFolder(token: string, driveId: string, folderName: string): Promise<SharePointFile[]> {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderName}:/children`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await res.json();
  if (data.error) { console.warn(`  [WARN] Could not list "${folderName}": ${data.error.message}`); return []; }
  return (data.value || []).filter((item: SharePointFile) => item.name?.toLowerCase().endsWith(".pdf") && item.file);
}

async function downloadFile(token: string, driveId: string, itemId: string): Promise<Buffer> {
  const metaRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const metaData = await metaRes.json();
  if (metaData.error) throw new Error(`Download metadata failed: ${metaData.error.message}`);
  const fileRes = await fetch(metaData["@microsoft.graph.downloadUrl"]);
  if (!fileRes.ok) throw new Error(`Download failed: HTTP ${fileRes.status}`);
  return Buffer.from(await fileRes.arrayBuffer());
}

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

async function clearIndex(searchClient: KnowledgeBaseSearchClient): Promise<void> {
  console.log("\n[CLEAR] Fetching all existing documents from index...");
  const results = await searchClient.searchAll(["id"]);
  const existingIds: string[] = [];
  for await (const result of results.results) { existingIds.push(result.document.id); }

  if (existingIds.length === 0) { console.log("[CLEAR] Index is already empty."); return; }

  console.log(`[CLEAR] Deleting ${existingIds.length} existing document(s)...`);
  const deleteResult = await searchClient.deleteDocuments(existingIds);
  const failed = deleteResult.results.filter((r: any) => !r.succeeded);
  if (failed.length > 0) console.warn(`[CLEAR] ${failed.length} document(s) failed to delete.`);
  else console.log(`[CLEAR] ✓ All ${existingIds.length} old document(s) removed.`);
}

async function reindex(
  searchClient: KnowledgeBaseSearchClient, token: string, driveId: string
): Promise<{ totalIndexed: number; totalSkipped: number }> {
  const folders = Object.keys(FOLDER_GROUP_MAP);
  let totalIndexed = 0;
  let totalSkipped = 0;

  for (const folder of folders) {
    const groupIds = FOLDER_GROUP_MAP[folder];
    console.log(`\n─────────────────────────────────`);
    console.log(`[FOLDER] ${folder}/ | Groups: ${groupIds.length}`);

    const files = await listFilesInFolder(token, driveId, folder);
    if (files.length === 0) { console.log(`  No PDFs found.`); continue; }
    console.log(`  ${files.length} PDF(s) found`);

    for (const file of files) {
      console.log(`\n  [FILE] ${file.name}`);
      try {
        const buffer = await downloadFile(token, driveId, file.id);
        console.log(`  [DOWNLOAD] ${buffer.length} bytes`);

        const pdfData = await pdf(buffer);
        const rawText = pdfData.text;
        if (!rawText || rawText.trim().length === 0) {
          console.log(`  [SKIP] No text extracted`); totalSkipped++; continue;
        }
        console.log(`  [EXTRACT] ${rawText.length} characters`);

        const { redacted, found } = redactPII(rawText);
        if (found.length > 0) console.log(`  [PII REDACTED] ${found.join(" | ")}`);

        const docId = Buffer.from(`${folder}__${file.name}`).toString("base64").replace(/[+/=]/g, "_");
        const document: IndexDocument = { id: docId, title: file.name, content: redacted, folder, GroupIds: groupIds };

        const result = await searchClient.uploadDocuments([document]);
        const status = result.results[0];
        if (status.succeeded) { console.log(`  [INDEXED] ✓ "${file.name}"`); totalIndexed++; }
        else { console.error(`  [FAILED] "${file.name}" — ${status.errorMessage}`); totalSkipped++; }
      } catch (err: any) {
        console.error(`  [ERROR] ${file.name}: ${err.message}`); totalSkipped++;
      }
    }
  }
  return { totalIndexed, totalSkipped };
}

async function sync(): Promise<void> {
  const searchClient = new KnowledgeBaseSearchClient(config);
  await clearIndex(searchClient);

  console.log("\n[AUTH] Getting Microsoft Graph access token...");
  const token = await getAccessToken();
  console.log("[AUTH] Token acquired.");

  const { driveId } = await getSiteAndDrive(token);

  console.log("\n[REINDEX] Starting fresh index from SharePoint...");
  const { totalIndexed, totalSkipped } = await reindex(searchClient, token, driveId);

  console.log(`\n─────────────────────────────────`);
  console.log(`[DONE] Indexed: ${totalIndexed} | Skipped/Failed: ${totalSkipped}`);
  console.log(`Index: "${config.searchIndexName}" on ${config.searchServiceName}.search.windows.net`);
}

sync().catch((error) => { console.error("[FATAL]", error); process.exit(1); });
