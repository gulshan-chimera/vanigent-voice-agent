import {
  SearchClient,
  AzureKeyCredential,
  SearchOptions,
} from "@azure/search-documents";
import { KnowledgeBaseConfig } from "../config.js";
import { SearchServiceError } from "../errors.js";
import type { IndexDocument } from "./types.js";

// ---------------------------------------------------------------------------
// Azure AI Search Client Wrapper
// ---------------------------------------------------------------------------

/**
 * Thin wrapper around the Azure SDK SearchClient.
 * Handles client creation and provides typed access to the index.
 */
export class KnowledgeBaseSearchClient {
  private readonly client: SearchClient<IndexDocument>;

  constructor(config: KnowledgeBaseConfig) {
    const credential = new AzureKeyCredential(config.searchApiKey);

    this.client = new SearchClient<IndexDocument>(
      config.searchEndpoint,
      config.searchIndexName,
      credential
    );
  }

  /**
   * Runs a full-text search query against the Azure AI Search index.
   * Returns raw SDK results — business logic is handled in searchService.
   */
  async search(query: string, options?: SearchOptions<IndexDocument>) {
    try {
      return await this.client.search(query, options);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unknown search error";
      throw new SearchServiceError(
        `Azure AI Search query failed: ${message}`,
        error
      );
    }
  }

  /**
   * Uploads documents to the index (used by indexing scripts).
   */
  async uploadDocuments(documents: IndexDocument[]) {
    try {
      return await this.client.uploadDocuments(documents);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unknown upload error";
      throw new SearchServiceError(
        `Azure AI Search upload failed: ${message}`,
        error
      );
    }
  }

  /**
   * Deletes documents by ID (used by sync script).
   */
  async deleteDocuments(ids: string[]) {
    try {
      return await this.client.deleteDocuments("id", ids);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unknown delete error";
      throw new SearchServiceError(
        `Azure AI Search delete failed: ${message}`,
        error
      );
    }
  }

  /**
   * Searches for all documents in the index (for listing / clearing).
   */
  async searchAll(select: (keyof IndexDocument)[]) {
    try {
      return await this.client.search("*", {
        select,
        top: 1000,
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unknown search error";
      throw new SearchServiceError(
        `Azure AI Search list all failed: ${message}`,
        error
      );
    }
  }
}
