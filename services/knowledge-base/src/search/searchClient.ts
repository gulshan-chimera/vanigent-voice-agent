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
