// src/lib/searchIndex.ts
//
// Manages the Azure AI Search index for the Human Resources KB — creating
// it if needed, clearing it, and uploading documents. Plain keyword search
// for now (no vector field) — this is the pre-Azure-OpenAI prototype.

import {
  SearchIndexClient,
  SearchClient,
  AzureKeyCredential,
  SearchIndex,
} from "@azure/search-documents";

export interface HrKbDocument {
  id: string;
  fileName: string;
  content: string;
  webUrl: string;
  driveItemId: string;
  lastModifiedDateTime: string;
}

function getConfig() {
  const serviceName = process.env.SEARCH_SERVICE_NAME;
  const apiKey = process.env.SEARCH_API_KEY;
  const indexName = process.env.SEARCH_INDEX_NAME;

  if (!serviceName || !apiKey || !indexName) {
    console.error("[SEARCH-INDEX] Missing SEARCH_SERVICE_NAME, SEARCH_API_KEY, or SEARCH_INDEX_NAME.");
    return null;
  }

  return {
    endpoint: `https://${serviceName}.search.windows.net`,
    apiKey,
    indexName,
  };
}

export function getSearchClient(): SearchClient<HrKbDocument> | null {
  const config = getConfig();
  if (!config) return null;

  return new SearchClient<HrKbDocument>(
    config.endpoint,
    config.indexName,
    new AzureKeyCredential(config.apiKey)
  );
}

function getIndexClient(): { client: SearchIndexClient; indexName: string } | null {
  const config = getConfig();
  if (!config) return null;

  return {
    client: new SearchIndexClient(config.endpoint, new AzureKeyCredential(config.apiKey)),
    indexName: config.indexName,
  };
}

/**
 * Creates the HR KB index if it doesn't already exist. Safe to call every
 * time — createOrUpdateIndex is idempotent.
 */
export async function ensureIndexExists(): Promise<boolean> {
  const result = getIndexClient();
  if (!result) return false;

  const indexDefinition: SearchIndex = {
    name: result.indexName,
    fields: [
      { name: "id", type: "Edm.String", key: true, filterable: true },
      { name: "fileName", type: "Edm.String", searchable: true, filterable: true },
      { name: "content", type: "Edm.String", searchable: true },
      { name: "webUrl", type: "Edm.String", filterable: false, searchable: false },
      { name: "driveItemId", type: "Edm.String", filterable: true, searchable: false },
      { name: "lastModifiedDateTime", type: "Edm.String", filterable: true, sortable: true },
    ],
  };

  try {
    await result.client.createOrUpdateIndex(indexDefinition);
    console.log(`[SEARCH-INDEX] Index "${result.indexName}" ready.`);
    return true;
  } catch (error) {
    console.error(`[SEARCH-INDEX] Failed to create/update index: ${(error as Error).message}`);
    return false;
  }
}

/**
 * Deletes every document currently in the index, so a fresh sync starts clean.
 */
export async function clearIndex(): Promise<boolean> {
  const searchClient = getSearchClient();
  if (!searchClient) return false;

  try {
    const existingIds: string[] = [];
    const results = await searchClient.search("*", { select: ["id"], top: 1000 });

    for await (const result of results.results) {
      existingIds.push(result.document.id);
    }

    if (existingIds.length === 0) {
      console.log("[SEARCH-INDEX] Index already empty.");
      return true;
    }

    const deleteResult = await searchClient.deleteDocuments("id", existingIds);
    const failed = deleteResult.results.filter((r) => !r.succeeded);

    if (failed.length > 0) {
      console.warn(`[SEARCH-INDEX] ${failed.length} document(s) failed to delete.`);
    }

    console.log(`[SEARCH-INDEX] Cleared ${existingIds.length - failed.length} document(s).`);
    return true;
  } catch (error) {
    console.error(`[SEARCH-INDEX] Failed to clear index: ${(error as Error).message}`);
    return false;
  }
}

export async function uploadDocument(doc: HrKbDocument): Promise<boolean> {
  const searchClient = getSearchClient();
  if (!searchClient) return false;

  try {
    const result = await searchClient.uploadDocuments([doc]);
    const status = result.results[0];

    if (!status.succeeded) {
      console.error(`[SEARCH-INDEX] Upload failed for "${doc.fileName}": ${status.errorMessage}`);
      return false;
    }

    return true;
  } catch (error) {
    console.error(`[SEARCH-INDEX] Upload error for "${doc.fileName}": ${(error as Error).message}`);
    return false;
  }
}