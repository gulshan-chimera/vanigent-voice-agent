import { LRUCache } from "lru-cache";
import { createHash } from "crypto";
import { KnowledgeBaseSearchClient } from "./searchClient.js";
import { SearchQueryError } from "../errors.js";
import OpenAI from "openai";
import { loadConfig } from "../config.js";
import type {
  SearchRequest,
  SearchResponse,
  SearchResult,
  IndexDocument,
} from "./types.js";

// ---------------------------------------------------------------------------
// Knowledge Base Search Service — Business Logic
// ---------------------------------------------------------------------------

const DEFAULT_TOP = 3;
const MAX_TOP = 10;

// ── Cache Configuration ──
const CACHE_MAX_ENTRIES = 200;
const CACHE_TTL_MS = 10 * 60 * 1000; // 10 minutes

export interface SearchServiceOptions {
  cacheMaxEntries?: number;
  cacheTtlMs?: number;
}

const config = loadConfig();
const openai = new OpenAI({ apiKey: config.openaiApiKey });
export class SearchService {
  private readonly cache: LRUCache<string, SearchResponse>;

  constructor(
    private readonly client: KnowledgeBaseSearchClient,
    options?: SearchServiceOptions
  ) {
    this.cache = new LRUCache<string, SearchResponse>({
      max: options?.cacheMaxEntries ?? CACHE_MAX_ENTRIES,
      ttl: options?.cacheTtlMs ?? CACHE_TTL_MS,
    });
  }

  async search(request: SearchRequest): Promise<SearchResponse> {
    if (!request.query || request.query.trim().length === 0) {
      throw new SearchQueryError("'query' is required and must be non-empty");
    }

    const query = request.query.trim().toLowerCase();
    const top = Math.min(Math.max(request.top ?? DEFAULT_TOP, 1), MAX_TOP);

    // ── Check cache ──
    const cacheKey = this.buildCacheKey(query, request.groupIds, top);
    const cached = this.cache.get(cacheKey);

    if (cached) {
      console.log(`[CACHE] HIT for "${query}" (key: ${cacheKey.slice(0, 12)}…)`);
      return cached;
    }

    console.log(`[CACHE] MISS for "${query}" — querying Azure AI Search`);

    // ── Build filter ──
    let filter: string | undefined;

    if (request.groupIds && request.groupIds.length > 0) {
      const idList = request.groupIds
        .map((id) => id.replace(/'/g, "''"))
        .join(",");
      filter = `GroupIds/any(g: search.in(g, '${idList}'))`;
    }

    // ── Execute search ──
        // Generate embedding for semantic search
    const embeddingResponse = await openai.embeddings.create({
      model: "text-embedding-3-small",
      input: query,
    });
    const queryVector = embeddingResponse.data[0].embedding;

    // Execute Hybrid Search
    const searchResults = await this.client.search(query, {
      top,
      filter,
      searchFields: ["content", "title"],
      select: ["id", "title", "content", "folder", "GroupIds"],
      
      vectorSearchOptions: {
        queries: [
          {
            kind: "vector",
            vector: queryVector,
            kNearestNeighborsCount: top,
            fields: ["contentVector"],
          },
        ],
      },
    });

    // ── Map results ──
    const results: SearchResult[] = [];
    let totalCount = 0;

    for await (const result of searchResults.results) {
      totalCount++;
      const doc = result.document;

      results.push({
        id: doc.id,
        title: doc.title,
        content: this.truncateContent(doc.content, 1000),
        folder: doc.folder,
        score: result.score ?? 0,
      });
    }

    const response: SearchResponse = {
      ok: true,
      results,
      totalCount,
      query,
    };

    // ── Store in cache ──
    this.cache.set(cacheKey, response);

    return response;
  }

  getCacheStats(): { size: number; maxSize: number } {
    return {
      size: this.cache.size,
      maxSize: CACHE_MAX_ENTRIES,
    };
  }

  clearCache(): void {
    this.cache.clear();
    console.log("[CACHE] Cache cleared");
  }

  private buildCacheKey(
    query: string,
    groupIds: string[] | undefined,
    top: number
  ): string {
    const sortedGroups = (groupIds ?? []).slice().sort().join("|");
    const raw = `${query}::${sortedGroups}::${top}`;
    return createHash("sha256").update(raw).digest("hex");
  }

  private truncateContent(content: string, maxLength: number): string {
    if (content.length <= maxLength) {
      return content;
    }
    return content.substring(0, maxLength) + "...";
  }
}


