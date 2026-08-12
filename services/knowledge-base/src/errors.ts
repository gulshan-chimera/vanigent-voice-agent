// ---------------------------------------------------------------------------
// Knowledge Base Service — Error Classes
// ---------------------------------------------------------------------------

/**
 * Base error for the knowledge base service.
 * All service-specific errors extend this.
 */
export class KnowledgeBaseError extends Error {
  public readonly statusCode: number;
  public readonly code: string;

  constructor(code: string, message: string, statusCode: number = 500) {
    super(message);
    this.name = "KnowledgeBaseError";
    this.code = code;
    this.statusCode = statusCode;
  }
}

/** Missing or invalid configuration */
export class SearchConfigError extends KnowledgeBaseError {
  constructor(message: string) {
    super("SEARCH_CONFIG_ERROR", message, 500);
  }
}

/** Invalid search request (bad input) */
export class SearchQueryError extends KnowledgeBaseError {
  constructor(message: string) {
    super("SEARCH_QUERY_ERROR", message, 400);
  }
}

/** Azure AI Search connection or query failure */
export class SearchServiceError extends KnowledgeBaseError {
  constructor(message: string, public readonly details?: unknown) {
    super("SEARCH_SERVICE_ERROR", message, 502);
  }
}

/** Microsoft Graph API failure (used by indexing scripts) */
export class GraphApiError extends KnowledgeBaseError {
  constructor(message: string, public readonly details?: unknown) {
    super("GRAPH_API_ERROR", message, 502);
  }
}

/** PDF extraction failure */
export class DocumentParseError extends KnowledgeBaseError {
  constructor(fileName: string, message: string) {
    super("DOCUMENT_PARSE_ERROR", `Failed to parse "${fileName}": ${message}`, 422);
  }
}

// ---------------------------------------------------------------------------
// Error response helper
// ---------------------------------------------------------------------------

export interface ErrorResponseBody {
  ok: false;
  error: {
    code: string;
    message: string;
    statusCode: number;
  };
}

export function toErrorResponse(error: unknown): ErrorResponseBody {
  if (error instanceof KnowledgeBaseError) {
    return {
      ok: false,
      error: {
        code: error.code,
        message: error.message,
        statusCode: error.statusCode,
      },
    };
  }

  if (error instanceof Error) {
    return {
      ok: false,
      error: {
        code: "INTERNAL_ERROR",
        message: error.message,
        statusCode: 500,
      },
    };
  }

  return {
    ok: false,
    error: {
      code: "INTERNAL_ERROR",
      message: "An unexpected error occurred",
      statusCode: 500,
    },
  };
}
