/**
 * scripts/indexDocuments.ts
 *
 * Pulls PDFs from SharePoint via Microsoft Graph API, redacts PII,
 * and indexes them into Azure AI Search.
 *
 * Run with: npm run index-docs
 */

import * as fs from "fs";
import * as path from "path";
import pdf from "pdf-parse";
import {
  SearchIndexClient,
  AzureKeyCredential,
} from "@azure/search-documents";
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

// Map each SharePoint folder → Entra group IDs allowed to read those documents
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

async function getSiteId(token: string): Promise<string> {
  const url = `https://graph.microsoft.com/v1.0/sites/${config.spHostname}:${config.spSitePath}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();

  if (data.error) {
    throw new Error(`getSiteId failed: ${data.error.message}`);
  }

  console.log(`[SHAREPOINT] Site ID: ${data.id}`);
  return data.id;
}

async function getDriveId(token: string, siteId: string): Promise<string> {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const data = await res.json();

  if (data.error) {
    throw new Error(`getDriveId failed: ${data.error.message}`);
  }

  const drive =
    data.value.find(
      (d: any) => d.name === "Documents" || d.driveType === "documentLibrary"
    ) || data.value[0];

  console.log(`[SHAREPOINT] Drive: "${drive.name}" (${drive.id})`);
  return drive.id;
}

interface SharePointFile {
  id: string;
  name: string;
  file?: unknown;
}

async function listFilesInFolder(
  token: string,
  driveId: string,
  folderName: string
): Promise<SharePointFile[]> {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderName}:/children`;
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

  const pdfs = (data.value || []).filter(
    (item: SharePointFile) =>
      item.name?.toLowerCase().endsWith(".pdf") && item.file
  );

  console.log(`[SHAREPOINT] "${folderName}" — ${pdfs.length} PDF(s) found`);
  return pdfs;
}

async function downloadFile(
  token: string,
  driveId: string,
  itemId: string
): Promise<Buffer> {
  const metaUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`;
  const metaRes = await fetch(metaUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const metaData = await metaRes.json();

  if (metaData.error) {
    throw new Error(`Download metadata failed: ${metaData.error.message}`);
  }

  const downloadUrl = metaData["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) {
    throw new Error("No download URL returned from Graph API");
  }

  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) {
    throw new Error(`Download failed: HTTP ${fileRes.status}`);
  }

  const arrayBuffer = await fileRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// ---------------------------------------------------------------------------
// PII REDACTION
// ---------------------------------------------------------------------------

const PII_PATTERNS: Array<{ name: string; pattern: RegExp }> = [
  {
    name: "Email",
    pattern: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g,
  },
  { name: "Phone", pattern: /(\+?(\d[\s\-.]?){10,14}\d)/g },
  { name: "Aadhaar", pattern: /\b\d{4}\s?\d{4}\s?\d{4}\b/g },
  { name: "PAN", pattern: /\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b/g },
  { name: "Passport", pattern: /\b[A-Z]{1,2}[0-9]{6,7}\b/g },
  { name: "CreditCard", pattern: /\b(?:\d[ \-]?){15,16}\b/g },
  { name: "SSN", pattern: /\b\d{3}[- ]?\d{2}[- ]?\d{4}\b/g },
  {
    name: "DOB",
    pattern:
      /\b(DOB|Date of Birth|D\.O\.B)[:\s]+\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/gi,
  },
  { name: "BankAccount", pattern: /\b\d{9,18}\b/g },
  { name: "IFSC", pattern: /\b[A-Z]{4}0[A-Z0-9]{6}\b/g },
];

function redactPII(text: string): RedactionResult {
  let redacted = text;
  const found: string[] = [];

  for (const { name, pattern } of PII_PATTERNS) {
    const matches = redacted.match(pattern);
    if (matches && matches.length > 0) {
      found.push(
        `${name} (${matches.length} instance${matches.length > 1 ? "s" : ""})`
      );
      redacted = redacted.replace(pattern, "[REDACTED]");
    }
  }

  return { redacted, found };
}

// ---------------------------------------------------------------------------
// INDEX SCHEMA SETUP
// ---------------------------------------------------------------------------

async function ensureIndexExists(): Promise<void> {
  const indexClient = new SearchIndexClient(
    config.searchEndpoint,
    new AzureKeyCredential(config.searchApiKey)
  );

  const index = {
    name: config.searchIndexName,
    fields: [
      {
        name: "id",
        type: "Edm.String" as const,
        key: true,
        searchable: false,
        filterable: true,
      },
      {
        name: "title",
        type: "Edm.String" as const,
        key: false,
        searchable: true,
        filterable: false,
      },
      {
        name: "content",
        type: "Edm.String" as const,
        key: false,
        searchable: true,
        filterable: false,
      },
      {
        name: "folder",
        type: "Edm.String" as const,
        key: false,
        searchable: false,
        filterable: true,
      },
      {
        name: "GroupIds",
        type: "Collection(Edm.String)" as const,
        key: false,
        searchable: false,
        filterable: true,
      },
    ],
  };

  await indexClient.createOrUpdateIndex(index);
  console.log(`[INDEX] Schema created/verified: "${config.searchIndexName}"`);
}

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------

async function indexDocuments(): Promise<void> {
  const searchClient = new KnowledgeBaseSearchClient(config);

  await ensureIndexExists();

  // Authenticate with Microsoft Graph
  console.log("\n[AUTH] Getting Microsoft Graph access token...");
  const token = await getAccessToken();
  console.log("[AUTH] Token acquired.");

  // Get SharePoint site and drive
  const siteId = await getSiteId(token);
  const driveId = await getDriveId(token, siteId);

  const folders = Object.keys(FOLDER_GROUP_MAP);
  let totalIndexed = 0;
  let totalSkipped = 0;

  for (const folder of folders) {
    const groupIds = FOLDER_GROUP_MAP[folder];
    console.log(`\n─────────────────────────────────`);
    console.log(`[FOLDER] ${folder}/ | Groups: ${groupIds.length}`);

    const files = await listFilesInFolder(token, driveId, folder);

    if (files.length === 0) {
      console.log(`  [SKIP] No PDFs found in this folder.`);
      continue;
    }

    for (const file of files) {
      console.log(`\n  [FILE] ${file.name}`);

      try {
        // Download PDF from SharePoint
        console.log(`  [DOWNLOAD] Fetching from SharePoint...`);
        const buffer = await downloadFile(token, driveId, file.id);
        console.log(`  [DOWNLOAD] ${buffer.length} bytes received`);

        // Extract text
        const pdfData = await pdf(buffer);
        const rawText = pdfData.text;

        if (!rawText || rawText.trim().length === 0) {
          console.log(`  [SKIP] No text extracted from ${file.name}`);
          totalSkipped++;
          continue;
        }
        console.log(`  [EXTRACT] ${rawText.length} characters extracted`);

        // Redact PII
        const { redacted, found } = redactPII(rawText);
        if (found.length > 0) {
          console.log(`  [PII REDACTED] ${found.join(" | ")}`);
        } else {
          console.log(`  [PII] No personal data detected`);
        }

        // Build document
        const docId = Buffer.from(`${folder}__${file.name}`)
          .toString("base64")
          .replace(/[+/=]/g, "_");

        const document: IndexDocument = {
          id: docId,
          title: file.name,
          content: redacted,
          folder: folder,
          GroupIds: groupIds,
        };

        // Push to Azure AI Search
        const result = await searchClient.uploadDocuments([document]);
        const status = result.results[0];

        if (status.succeeded) {
          console.log(`  [INDEXED] ✓ "${file.name}"`);
          totalIndexed++;
        } else {
          console.error(
            `  [FAILED] "${file.name}" — ${status.errorMessage}`
          );
          totalSkipped++;
        }
      } catch (err: any) {
        console.error(`  [ERROR] ${file.name}: ${err.message}`);
        totalSkipped++;
      }
    }
  }

  console.log(`\n─────────────────────────────────`);
  console.log(
    `[DONE] Indexed: ${totalIndexed} | Skipped/Failed: ${totalSkipped}`
  );
  console.log(
    `Index: "${config.searchIndexName}" on ${config.searchServiceName}.search.windows.net`
  );
}

indexDocuments().catch((error) => {
  console.error("[FATAL]", error);
  process.exit(1);
});
