// ---------------------------------------------------------------------------
// HTTP Status Codes
// ---------------------------------------------------------------------------

export enum HttpStatusCode {
  BAD_REQUEST = 400,
  UNAUTHORIZED = 401,
  FORBIDDEN = 403,
  NOT_FOUND = 404,
  REQUEST_TIMEOUT = 408,
  CONFLICT = 409,
  UNPROCESSABLE_ENTITY = 422,
  TOO_MANY_REQUESTS = 429,
  INTERNAL_SERVER_ERROR = 500,
  BAD_GATEWAY = 502,
  SERVICE_UNAVAILABLE = 503,
  GATEWAY_TIMEOUT = 504
}

// ---------------------------------------------------------------------------
// Error Codes — every error produced by the toolkit uses one of these
// ---------------------------------------------------------------------------

export enum ErrorCode {
  TOOL_NOT_FOUND = "TOOL_NOT_FOUND",
  TOOL_ALREADY_REGISTERED = "TOOL_ALREADY_REGISTERED",
  TOOL_VALIDATION_ERROR = "TOOL_VALIDATION_ERROR",
  TOOL_EXECUTION_ERROR = "TOOL_EXECUTION_ERROR",
  TOOL_TIMEOUT = "TOOL_TIMEOUT",
  TOOL_AUTHENTICATION_ERROR = "TOOL_AUTHENTICATION_ERROR",
  TOOL_AUTHORIZATION_ERROR = "TOOL_AUTHORIZATION_ERROR",
  TOOL_RATE_LIMIT_ERROR = "TOOL_RATE_LIMIT_ERROR",
  TOOL_BAD_REQUEST = "TOOL_BAD_REQUEST",
  TOOL_EXTERNAL_SERVICE_ERROR = "TOOL_EXTERNAL_SERVICE_ERROR",
  TOOL_SERVICE_UNAVAILABLE = "TOOL_SERVICE_UNAVAILABLE",
  TOOL_GATEWAY_TIMEOUT = "TOOL_GATEWAY_TIMEOUT",
  TOOL_INTERNAL_ERROR = "TOOL_INTERNAL_ERROR",
  TOOL_CONFIGURATION_ERROR = "TOOL_CONFIGURATION_ERROR",
  VAPI_WEBHOOK_ERROR = "VAPI_WEBHOOK_ERROR",
  VAPI_CALL_ERROR = "VAPI_CALL_ERROR"
}

// ---------------------------------------------------------------------------
// ErrorCode → HttpStatusCode mapping
// ---------------------------------------------------------------------------

export const ERROR_CODE_TO_STATUS: Record<ErrorCode, HttpStatusCode> = {
  [ErrorCode.TOOL_NOT_FOUND]: HttpStatusCode.NOT_FOUND,
  [ErrorCode.TOOL_ALREADY_REGISTERED]: HttpStatusCode.CONFLICT,
  [ErrorCode.TOOL_VALIDATION_ERROR]: HttpStatusCode.UNPROCESSABLE_ENTITY,
  [ErrorCode.TOOL_EXECUTION_ERROR]: HttpStatusCode.INTERNAL_SERVER_ERROR,
  [ErrorCode.TOOL_TIMEOUT]: HttpStatusCode.REQUEST_TIMEOUT,
  [ErrorCode.TOOL_AUTHENTICATION_ERROR]: HttpStatusCode.UNAUTHORIZED,
  [ErrorCode.TOOL_AUTHORIZATION_ERROR]: HttpStatusCode.FORBIDDEN,
  [ErrorCode.TOOL_RATE_LIMIT_ERROR]: HttpStatusCode.TOO_MANY_REQUESTS,
  [ErrorCode.TOOL_BAD_REQUEST]: HttpStatusCode.BAD_REQUEST,
  [ErrorCode.TOOL_EXTERNAL_SERVICE_ERROR]: HttpStatusCode.BAD_GATEWAY,
  [ErrorCode.TOOL_SERVICE_UNAVAILABLE]: HttpStatusCode.SERVICE_UNAVAILABLE,
  [ErrorCode.TOOL_GATEWAY_TIMEOUT]: HttpStatusCode.GATEWAY_TIMEOUT,
  [ErrorCode.TOOL_INTERNAL_ERROR]: HttpStatusCode.INTERNAL_SERVER_ERROR,
  [ErrorCode.TOOL_CONFIGURATION_ERROR]: HttpStatusCode.INTERNAL_SERVER_ERROR,
  [ErrorCode.VAPI_WEBHOOK_ERROR]: HttpStatusCode.BAD_REQUEST,
  [ErrorCode.VAPI_CALL_ERROR]: HttpStatusCode.INTERNAL_SERVER_ERROR
};

// Retryable error codes — callers can retry these after a backoff
const RETRYABLE_CODES: ReadonlySet<ErrorCode> = new Set([
  ErrorCode.TOOL_TIMEOUT,
  ErrorCode.TOOL_RATE_LIMIT_ERROR,
  ErrorCode.TOOL_EXTERNAL_SERVICE_ERROR,
  ErrorCode.TOOL_SERVICE_UNAVAILABLE,
  ErrorCode.TOOL_GATEWAY_TIMEOUT
]);

// ---------------------------------------------------------------------------
// Base error class
// ---------------------------------------------------------------------------

export class ToolError extends Error {
  public readonly code: ErrorCode;
  public readonly statusCode: HttpStatusCode;
  public readonly isRetryable: boolean;
  public readonly details?: unknown;

  constructor(code: ErrorCode, message: string, details?: unknown) {
    super(message);
    this.name = "ToolError";
    this.code = code;
    this.statusCode = ERROR_CODE_TO_STATUS[code] ?? HttpStatusCode.INTERNAL_SERVER_ERROR;
    this.isRetryable = RETRYABLE_CODES.has(code);
    this.details = details;
  }
}

// ---------------------------------------------------------------------------
// Specific error subclasses
// ---------------------------------------------------------------------------

/** 404 — Tool not found in the registry */
export class ToolNotFoundError extends ToolError {
  constructor(toolName: string) {
    super(ErrorCode.TOOL_NOT_FOUND, `Tool not found: ${toolName}`, { toolName });
  }
}

/** 422 — Input validation failed */
export class ToolValidationError extends ToolError {
  constructor(toolName: string, message: string, details?: unknown) {
    super(ErrorCode.TOOL_VALIDATION_ERROR, `${toolName}: ${message}`, details);
  }
}

/** 401 — Authentication failed (missing/invalid/expired token) */
export class ToolAuthenticationError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.TOOL_AUTHENTICATION_ERROR, message, details);
  }
}

/** 403 — Authenticated but insufficient permissions */
export class ToolAuthorizationError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.TOOL_AUTHORIZATION_ERROR, message, details);
  }
}

/** 408 — Tool execution timed out */
export class ToolTimeoutError extends ToolError {
  public readonly timeoutMs: number;

  constructor(toolName: string, timeoutMs: number, details?: unknown) {
    super(
      ErrorCode.TOOL_TIMEOUT,
      `Tool "${toolName}" timed out after ${timeoutMs}ms`,
      details
    );
    this.timeoutMs = timeoutMs;
  }
}

/** 429 — Rate limited by a provider (VAPI, LLM, TTS, STT, etc.) */
export class ToolRateLimitError extends ToolError {
  public readonly retryAfterMs?: number;

  constructor(message: string, retryAfterMs?: number, details?: unknown) {
    super(ErrorCode.TOOL_RATE_LIMIT_ERROR, message, {
      ...((typeof details === "object" && details !== null) ? details : {}),
      retryAfterMs
    });
    this.retryAfterMs = retryAfterMs;
  }
}

/** 400 — Malformed request structure */
export class ToolBadRequestError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.TOOL_BAD_REQUEST, message, details);
  }
}

/** 502 — Upstream service (LLM, TTS, STT, etc.) returned an error */
export class ToolExternalServiceError extends ToolError {
  public readonly serviceName: string;

  constructor(serviceName: string, message: string, details?: unknown) {
    super(
      ErrorCode.TOOL_EXTERNAL_SERVICE_ERROR,
      `[${serviceName}] ${message}`,
      details
    );
    this.serviceName = serviceName;
  }
}

/** 503 — Service temporarily unavailable */
export class ToolServiceUnavailableError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.TOOL_SERVICE_UNAVAILABLE, message, details);
  }
}

/** 504 — Upstream service timed out */
export class ToolGatewayTimeoutError extends ToolError {
  public readonly serviceName?: string;

  constructor(message: string, serviceName?: string, details?: unknown) {
    super(ErrorCode.TOOL_GATEWAY_TIMEOUT, message, details);
    this.serviceName = serviceName;
  }
}

/** 500 — Missing or invalid configuration */
export class ToolConfigurationError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.TOOL_CONFIGURATION_ERROR, message, details);
  }
}

/** 400 — VAPI webhook payload is invalid or cannot be parsed */
export class VapiWebhookError extends ToolError {
  constructor(message: string, details?: unknown) {
    super(ErrorCode.VAPI_WEBHOOK_ERROR, message, details);
  }
}

/** 500 — VAPI call ended with an error (wraps endedReason) */
export class VapiCallError extends ToolError {
  public readonly endedReason: string;

  constructor(endedReason: string, message: string, details?: unknown) {
    super(ErrorCode.VAPI_CALL_ERROR, message, {
      ...((typeof details === "object" && details !== null) ? details : {}),
      endedReason
    });
    this.endedReason = endedReason;
  }
}

// ---------------------------------------------------------------------------
// Helper utilities
// ---------------------------------------------------------------------------

/** Returns `true` if the error is an instance of ToolError and is retryable. */
export function isRetryableError(error: unknown): boolean {
  return error instanceof ToolError && error.isRetryable;
}

/** Standardized error response shape for API / webhook responses. */
export interface ErrorResponseBody {
  code: string;
  statusCode: number;
  message: string;
  isRetryable: boolean;
  details?: unknown;
}

/**
 * Converts any error into a standardized `ErrorResponseBody`.
 * Useful for returning consistent JSON error payloads from webhooks or API routes.
 */
export function toErrorResponse(error: unknown): ErrorResponseBody {
  if (error instanceof ToolError) {
    return {
      code: error.code,
      statusCode: error.statusCode,
      message: error.message,
      isRetryable: error.isRetryable,
      details: error.details
    };
  }

  if (error instanceof Error) {
    return {
      code: ErrorCode.TOOL_INTERNAL_ERROR,
      statusCode: HttpStatusCode.INTERNAL_SERVER_ERROR,
      message: error.message,
      isRetryable: false,
      details: { name: error.name, stack: error.stack }
    };
  }

  return {
    code: ErrorCode.TOOL_INTERNAL_ERROR,
    statusCode: HttpStatusCode.INTERNAL_SERVER_ERROR,
    message: "An unexpected error occurred",
    isRetryable: false,
    details: error
  };
}
