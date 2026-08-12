// ---------------------------------------------------------------------------
// Knowledge Base Service — Search Types
// ---------------------------------------------------------------------------

/** Incoming search request from API consumers */
export interface SearchRequest {
  /** The user's natural-language query */
  query: string;

  /** Optional: Entra group IDs of the caller — used to filter documents by access */
  groupIds?: string[];

  /** Max results to return (default: 3) */
  top?: number;
}

/** A single search hit from Azure AI Search */
export interface SearchResult {
  /** Document ID in the index */
  id: string;

  /** Original file name (e.g. "Refund_Policy.pdf") */
  title: string;

  /** PII-redacted document content (or the matched excerpt) */
  content: string;

  /** SharePoint folder the document came from */
  folder: string;

  /** Azure AI Search relevance score */
  score: number;
}

/** Response shape returned by the search API */
export interface SearchResponse {
  ok: true;
  results: SearchResult[];
  totalCount: number;
  query: string;
}

// ---------------------------------------------------------------------------
// Indexing types (used by scripts)
// ---------------------------------------------------------------------------

/** Shape of a document stored in the Azure AI Search index */
export interface IndexDocument {
  id: string;
  title: string;
  content: string;
  folder: string;
  GroupIds: string[];
}

/** Maps SharePoint folder names to Entra group IDs that can access those docs */
export type FolderGroupMap = Record<string, string[]>;

/** PII redaction result */
export interface RedactionResult {
  redacted: string;
  found: string[];
}
