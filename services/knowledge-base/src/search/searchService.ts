import { LRUCache } from "lru-cache";
import { createHash } from "crypto";
import { KnowledgeBaseSearchClient } from "./searchClient.js";
import { SearchQueryError } from "../errors.js";
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
// Max 200 cached queries, each entry expires after 1 hour.
// The cache key combines query + groupIds + top so that:
//   - Different users with different groups get separate cache entries
//   - Same user asking the same question gets an instant cached response
const CACHE_MAX_ENTRIES = 200;
const CACHE_TTL_MS = 10 * 60 * 1000; // 10 minutes

export interface SearchServiceOptions {
  /** Max cached entries (default: 200) */
  cacheMaxEntries?: number;
  /** Cache TTL in milliseconds (default: 1 hour) */
  cacheTtlMs?: number;
}

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

  /**
   * Performs a knowledge base search.
   *
   * 1. Validates the request
   * 2. Checks cache for an existing result
   * 3. If cache miss, queries Azure AI Search with optional group filtering
   * 4. Caches and returns the result
   */
  async search(request: SearchRequest): Promise<SearchResponse> {
    // ── Validate ──
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
    // If groupIds are provided, only return documents the caller is allowed to see.
    // Each document has a GroupIds field (Collection(Edm.String)) set during indexing.
    let filter: string | undefined;

    if (request.groupIds && request.groupIds.length > 0) {
      const idList = request.groupIds
        .map((id) => id.replace(/'/g, "''"))
        .join(",");
      filter = `GroupIds/any(g: search.in(g, '${idList}'))`;
    }

    // ── Execute search ──
    const searchResults = await this.client.search(query, {
      top,
      filter,
      searchFields: ["content", "title"],
      select: ["id", "title", "content", "folder", "GroupIds"],
      queryType: "simple",
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

  /**
   * Returns current cache statistics.
   * Useful for the health endpoint or monitoring.
   */
  getCacheStats(): { size: number; maxSize: number } {
    return {
      size: this.cache.size,
      maxSize: CACHE_MAX_ENTRIES,
    };
  }

  /**
   * Clears the entire cache.
   * Call this after a re-index to ensure fresh results.
   */
  clearCache(): void {
    this.cache.clear();
    console.log("[CACHE] Cache cleared");
  }

  /**
   * Builds a deterministic cache key from the query parameters.
   * Uses SHA-256 hash so that the key is fixed-length and safe for any input.
   *
   * The key includes groupIds (sorted) so that two callers with different
   * permissions never share a cached result.
   */
  private buildCacheKey(
    query: string,
    groupIds: string[] | undefined,
    top: number
  ): string {
    const sortedGroups = (groupIds ?? []).slice().sort().join("|");
    const raw = `${query}::${sortedGroups}::${top}`;
    return createHash("sha256").update(raw).digest("hex");
  }

  /**
   * Truncates long document content for API responses.
   * The full content lives in the index; callers get a useful excerpt.
   */
  private truncateContent(content: string, maxLength: number): string {
    if (content.length <= maxLength) {
      return content;
    }
    return content.substring(0, maxLength) + "...";
  }
}
