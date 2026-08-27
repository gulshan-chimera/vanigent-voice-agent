// src/lib/searchIndex.ts
//
// Manages the Azure AI Search index across all KB document libraries.
// Supports incremental sync: we can read back which files are currently
// indexed (and their change markers), and delete a single file's chunks
// without touching anything else.

import {
  SearchIndexClient,
  SearchClient,
  AzureKeyCredential,
  SearchIndex,
} from "@azure/search-documents";

export interface KbDocument {
  id: string;
  fileName: string;
  library: string;
  driveId: string;
  driveItemId: string;
  cTag: string;
  chunkIndex: number;
  content: string;
  webUrl: string;
  lastModifiedDateTime: string;
  contentVector: number[];
}

/** One row per indexed FILE (derived from its chunkIndex 0 page). */
export interface IndexedFileState {
  driveItemId: string;
  driveId: string;
  library: string;
  fileName: string;
  cTag: string;
}

function getConfig() {
  const serviceName = process.env.SEARCH_SERVICE_NAME;
  const apiKey = process.env.SEARCH_API_KEY;
  const indexName = process.env.SEARCH_INDEX_NAME;

  if (!serviceName || !apiKey || !indexName) {
    console.error(
      "[SEARCH-INDEX] Missing SEARCH_SERVICE_NAME, SEARCH_API_KEY, or SEARCH_INDEX_NAME."
    );
    return null;
  }

  if (serviceName.includes("://") || serviceName.includes(".")) {
    console.error(
      `[SEARCH-INDEX] SEARCH_SERVICE_NAME must be the SHORT name only (e.g. "vanigent-search"), got: "${serviceName}"`
    );
    return null;
  }

  return {
    endpoint: `https://${serviceName}.search.windows.net`,
    apiKey,
    indexName,
  };
}

export function getSearchClient(): SearchClient<KbDocument> | null {
  const config = getConfig();
  if (!config) return null;

  return new SearchClient<KbDocument>(
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

/** Escapes single quotes for safe use inside an OData filter literal. */
function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

export async function ensureIndexExists(): Promise<boolean> {
  const result = getIndexClient();
  if (!result) return false;

  const indexDefinition: SearchIndex = {
    name: result.indexName,
    fields: [
      { name: "id", type: "Edm.String", key: true, filterable: true, sortable: true },
      { name: "fileName", type: "Edm.String", searchable: true, filterable: true },
      { name: "library", type: "Edm.String", searchable: true, filterable: true, facetable: true },
      { name: "driveId", type: "Edm.String", filterable: true, searchable: false },
      { name: "driveItemId", type: "Edm.String", filterable: true, searchable: false },
      { name: "cTag", type: "Edm.String", filterable: true, searchable: false },
      { name: "chunkIndex", type: "Edm.Int32", filterable: true, sortable: true },
      { name: "content", type: "Edm.String", searchable: true },
      { name: "webUrl", type: "Edm.String", filterable: false, searchable: false },
      { name: "lastModifiedDateTime", type: "Edm.String", filterable: true, sortable: true },
      {
        name: "contentVector",
        type: "Collection(Edm.Single)",
        searchable: true,
        vectorSearchDimensions: 1536,
        vectorSearchProfileName: "kb-vector-profile",
      },
    ],
    vectorSearch: {
      algorithms: [{ name: "kb-hnsw", kind: "hnsw" }],
      profiles: [{ name: "kb-vector-profile", algorithmConfigurationName: "kb-hnsw" }],
    },
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
 * Reads back one row per indexed FILE by selecting only chunkIndex 0
 * pages — so we get the full picture of what's indexed without pulling
 * every page document. Pages via skip/top so we don't stop at 1000.
 */
export async function getIndexedFileState(): Promise<Map<string, IndexedFileState> | null> {
  const searchClient = getSearchClient();
  if (!searchClient) return null;

  const state = new Map<string, IndexedFileState>();
  const pageSize = 1000;
  let skip = 0;

  try {
    for (;;) {
      const results = await searchClient.search("*", {
        filter: "chunkIndex eq 0",
        select: ["driveItemId", "driveId", "library", "fileName", "cTag"],
        orderBy: ["id asc"],
        top: pageSize,
        skip,
      });

      let batchCount = 0;
      for await (const result of results.results) {
        const doc = result.document;
        batchCount++;
        state.set(doc.driveItemId, {
          driveItemId: doc.driveItemId,
          driveId: doc.driveId,
          library: doc.library,
          fileName: doc.fileName,
          cTag: doc.cTag ?? "",
        });
      }

      if (batchCount < pageSize) break;
      skip += pageSize;
    }

    console.log(`[SEARCH-INDEX] Index currently holds ${state.size} file(s).`);
    return state;
  } catch (error) {
    console.error(`[SEARCH-INDEX] Failed to read index state: ${(error as Error).message}`);
    return null;
  }
}

/**
 * Deletes every chunk belonging to ONE file. Loops until nothing is
 * left, so it works regardless of how many pages the file has.
 */
export async function deleteFileChunks(driveItemId: string): Promise<number> {
  const searchClient = getSearchClient();
  if (!searchClient) return 0;

  const filter = `driveItemId eq '${escapeODataValue(driveItemId)}'`;
  let totalDeleted = 0;

  try {
    for (;;) {
      const results = await searchClient.search("*", {
        filter,
        select: ["id"],
        top: 1000,
      });

      const ids: string[] = [];
      for await (const result of results.results) {
        ids.push(result.document.id);
      }

      if (ids.length === 0) break;

      const deleteResult = await searchClient.deleteDocuments("id", ids);
      const failed = deleteResult.results.filter((r) => !r.succeeded);

      if (failed.length > 0) {
        console.warn(
          `[SEARCH-INDEX] ${failed.length} chunk(s) failed to delete for item ${driveItemId} — stopping to avoid a loop.`
        );
        totalDeleted += ids.length - failed.length;
        break;
      }

      totalDeleted += ids.length;
      if (ids.length < 1000) break;
    }

    if (totalDeleted > 0) {
      console.log(`[SEARCH-INDEX] Deleted ${totalDeleted} chunk(s) for item ${driveItemId}.`);
    }
    return totalDeleted;
  } catch (error) {
    console.error(
      `[SEARCH-INDEX] Failed to delete chunks for ${driveItemId}: ${(error as Error).message}`
    );
    return totalDeleted;
  }
}

/**
 * Wipes the entire index. Only used behind an explicit force flag —
 * normal syncs must never call this.
 */
export async function clearEntireIndex(): Promise<boolean> {
  const searchClient = getSearchClient();
  if (!searchClient) return false;

  try {
    let total = 0;

    for (;;) {
      const results = await searchClient.search("*", { select: ["id"], top: 1000 });

      const ids: string[] = [];
      for await (const result of results.results) {
        ids.push(result.document.id);
      }

      if (ids.length === 0) break;

      const deleteResult = await searchClient.deleteDocuments("id", ids);
      const failed = deleteResult.results.filter((r) => !r.succeeded);
      total += ids.length - failed.length;

      if (failed.length > 0) {
        console.warn(`[SEARCH-INDEX] ${failed.length} document(s) failed to delete.`);
        break;
      }
    }

    console.log(`[SEARCH-INDEX] FORCE: cleared ${total} document(s).`);
    return true;
  } catch (error) {
    console.error(`[SEARCH-INDEX] Failed to clear index: ${(error as Error).message}`);
    return false;
  }
}

export async function uploadDocument(doc: KbDocument): Promise<boolean> {
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