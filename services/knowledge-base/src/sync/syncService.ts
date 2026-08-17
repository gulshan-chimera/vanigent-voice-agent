import fs from "node:fs";
import path from "node:path";
import pdf from "pdf-parse";
import OpenAI from "openai";
import { loadConfig, requireGraphCredentials, KnowledgeBaseConfig } from "../config.js";
import { KnowledgeBaseSearchClient } from "../search/searchClient.js";
import type { IndexDocument, FolderGroupMap, RedactionResult } from "../search/types.js";
import { chunkText } from "../utils/chunker.js";

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

const SYNC_STATE_PATH = path.resolve(__dirname, "../../../.sync-state.json");

interface SyncedFile {
  itemId: string;
  name: string;
  folder: string;
  lastModified: string;
  eTag: string;
  size: number;
  chunkIds: string[];
}

interface SyncState {
  lastSyncedAt: string;
  files: Record<string, SyncedFile>;
}

interface SharePointFileDetail {
  id: string;
  name: string;
  lastModifiedDateTime: string;
  eTag: string;
  size: number;
  file?: unknown;
}

export class SyncService {
  private searchClient: KnowledgeBaseSearchClient;
  private openai: OpenAI;
  private config: KnowledgeBaseConfig;

  constructor() {
    this.config = loadConfig();
    requireGraphCredentials(this.config);
    this.searchClient = new KnowledgeBaseSearchClient(this.config);
    this.openai = new OpenAI({ apiKey: this.config.openaiApiKey });
  }

  public async runDeltaSync(): Promise<void> {
    console.log("[SYNC] Starting Delta Sync...");
    const prevState = this.loadSyncState();
    const token = await this.getAccessToken();
    const { driveId } = await this.getSiteAndDrive(token);
    
    const currentFiles = new Map<string, { file: SharePointFileDetail; folder: string }>();

    for (const folder of Object.keys(FOLDER_GROUP_MAP)) {
      const files = await this.listFilesInFolder(token, driveId, folder);
      for (const file of files) {
        const key = this.makeFileKey(folder, file.name);
        currentFiles.set(key, { file, folder });
      }
    }

    const delta = this.computeDelta(prevState, currentFiles);
    
    console.log(`[SYNC] Delta Summary: ${delta.newFiles.length} new, ${delta.modifiedFiles.length} modified, ${delta.deletedKeys.length} deleted, ${delta.unchangedCount} unchanged.`);

    if (delta.newFiles.length === 0 && delta.modifiedFiles.length === 0 && delta.deletedKeys.length === 0) {
       console.log("[SYNC] Everything up to date.");
       return;
    }

    const newState: SyncState = { lastSyncedAt: new Date().toISOString(), files: {} };

    for (const [key, { file }] of currentFiles) {
      const prev = prevState.files[key];
      if (prev && prev.lastModified === file.lastModifiedDateTime && prev.eTag === file.eTag && prev.size === file.size) {
        newState.files[key] = prev;
      }
    }

    // Process NEW
    for (const { file, folder } of delta.newFiles) {
      const key = this.makeFileKey(folder, file.name);
      const synced = await this.indexSingleFile(token, driveId, file, folder, FOLDER_GROUP_MAP[folder]);
      if (synced) newState.files[key] = synced;
    }

    // Process MODIFIED
    for (const { file, folder } of delta.modifiedFiles) {
      const key = this.makeFileKey(folder, file.name);
      const synced = await this.indexSingleFile(token, driveId, file, folder, FOLDER_GROUP_MAP[folder]);
      if (synced) newState.files[key] = synced;
    }

    // Process DELETED
    const idsToDelete: string[] = [];
    for (const key of delta.deletedKeys) {
      const prev = prevState.files[key];
      if (prev) {
        if (prev.chunkIds) { idsToDelete.push(...prev.chunkIds); }
        else if ((prev as any).indexDocId) { idsToDelete.push((prev as any).indexDocId); }
      }
    }
    if (idsToDelete.length > 0) {
      console.log(`[SYNC] Deleting ${idsToDelete.length} chunks from index.`);
      await this.searchClient.deleteDocuments(idsToDelete);
    }

    this.saveSyncState(newState);
    console.log("[SYNC] Delta Sync Complete.");
  }

  // --- Helpers ---

  private loadSyncState(): SyncState {
    if (!fs.existsSync(SYNC_STATE_PATH)) return { lastSyncedAt: "", files: {} };
    try { return JSON.parse(fs.readFileSync(SYNC_STATE_PATH, "utf-8")) as SyncState; } 
    catch { return { lastSyncedAt: "", files: {} }; }
  }

  private saveSyncState(state: SyncState): void {
    fs.writeFileSync(SYNC_STATE_PATH, JSON.stringify(state, null, 2), "utf-8");
  }

  private async getAccessToken(): Promise<string> {
    const res = await fetch(`https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: this.config.clientId,
        client_secret: this.config.clientSecret,
        scope: "https://graph.microsoft.com/.default",
      }),
    });
    const data = await res.json();
    if (!data.access_token) throw new Error("Auth failed");
    return data.access_token;
  }

  private async getSiteAndDrive(token: string): Promise<{ siteId: string; driveId: string }> {
    const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${this.config.spHostname}:${this.config.spSitePath}`, { headers: { Authorization: `Bearer ${token}` } });
    const siteData = await siteRes.json();
    const driveRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteData.id}/drives`, { headers: { Authorization: `Bearer ${token}` } });
    const driveData = await driveRes.json();
    const drive = driveData.value.find((d: any) => d.name === "Documents" || d.driveType === "documentLibrary") || driveData.value[0];
    return { siteId: siteData.id, driveId: drive.id };
  }

  private async listFilesInFolder(token: string, driveId: string, folderName: string): Promise<SharePointFileDetail[]> {
    const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folderName}:/children?$select=id,name,lastModifiedDateTime,eTag,size,file`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = await res.json();
    return (data.value || []).filter((item: SharePointFileDetail) => item.name?.toLowerCase().endsWith(".pdf") && item.file);
  }

  private async downloadFile(token: string, driveId: string, itemId: string): Promise<Buffer> {
    const metaRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`, { headers: { Authorization: `Bearer ${token}` } });
    const metaData = await metaRes.json();
    const downloadUrl = metaData["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) throw new Error("No download URL");
    const fileRes = await fetch(downloadUrl);
    return Buffer.from(await fileRes.arrayBuffer());
  }

  private makeDocId(folder: string, fileName: string): string {
    return Buffer.from(`${folder}__${fileName}`).toString("base64").replace(/[+/=]/g, "_");
  }

  private makeFileKey(folder: string, fileName: string): string {
    return `${folder}__${fileName}`;
  }

  private computeDelta(prevState: SyncState, currentFiles: Map<string, { file: SharePointFileDetail; folder: string }>) {
    const newFiles = [];
    const modifiedFiles = [];
    let unchangedCount = 0;
    for (const [key, { file, folder }] of currentFiles) {
      const prev = prevState.files[key];
      if (!prev) newFiles.push({ file, folder });
      else if (prev.lastModified !== file.lastModifiedDateTime || prev.eTag !== file.eTag || prev.size !== file.size) modifiedFiles.push({ file, folder });
      else unchangedCount++;
    }
    const deletedKeys = [];
    for (const key of Object.keys(prevState.files)) {
      if (!currentFiles.has(key)) deletedKeys.push(key);
    }
    return { newFiles, modifiedFiles, deletedKeys, unchangedCount };
  }

  private redactPII(text: string): RedactionResult {
    const PII_PATTERNS = [
      { name: "Email", pattern: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g },
      { name: "Phone", pattern: /(\+?(\d[\s\-.]?){10,14}\d)/g }
    ];
    let redacted = text;
    const found: string[] = [];
    for (const { name, pattern } of PII_PATTERNS) {
      const matches = redacted.match(pattern);
      if (matches && matches.length > 0) {
        found.push(`${name}`);
        redacted = redacted.replace(pattern, "[REDACTED]");
      }
    }
    return { redacted, found };
  }

  private async indexSingleFile(
    token: string,
    driveId: string,
    file: SharePointFileDetail,
    folder: string,
    groupIds: string[]
  ): Promise<SyncedFile | null> {
    const docId = this.makeDocId(folder, file.name);

    try {
      const buffer = await this.downloadFile(token, driveId, file.id);
      const pdfData = await pdf(buffer);
      const rawText = pdfData.text || "";
      const { redacted } = this.redactPII(rawText);

      // USE CUSTOM CHUNKER
      const chunks = chunkText(redacted, { chunkSize: 1000, chunkOverlap: 200 });

      const chunkIds: string[] = [];
      const documents: IndexDocument[] = [];

      for (let i = 0; i < chunks.length; i++) {
        const chunkTextContent = chunks[i];
        const chunkId = `${docId}_chunk_${i}`;
        chunkIds.push(chunkId);

        const response = await this.openai.embeddings.create({ model: "text-embedding-3-small", input: chunkTextContent });
        const vector = response.data[0].embedding;

        documents.push({
          id: chunkId,
          parentFileId: docId,
          title: `${file.name} (Part ${i+1})`,
          content: chunkTextContent,
          contentVector: vector,
          folder,
          GroupIds: groupIds,
        });
      }

      if (documents.length > 0) {
          const result = await this.searchClient.uploadDocuments(documents);
          const failed = result.results.filter((r: any) => !r.succeeded);
          if (failed.length > 0) return null;
      }

      console.log(`[INDEXED] ${file.name} (${chunks.length} chunks)`);
      return {
        itemId: file.id,
        name: file.name,
        folder,
        lastModified: file.lastModifiedDateTime,
        eTag: file.eTag,
        size: file.size,
        chunkIds,
      };
    } catch (error) {
      console.error(`[ERROR] Processing ${file.name}:`, error);
      return null;
    }
  }
}
