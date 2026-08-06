export class ToolError extends Error {
  public readonly code: string;
  public readonly details?: unknown;

  constructor(code: string, message: string, details?: unknown) {
    super(message);
    this.name = "ToolError";
    this.code = code;
    this.details = details;
  }
}

export class ToolNotFoundError extends ToolError {
  constructor(toolName: string) {
    super("TOOL_NOT_FOUND", `Tool not found: ${toolName}`, { toolName });
  }
}

export class ToolValidationError extends ToolError {
  constructor(toolName: string, message: string, details?: unknown) {
    super("TOOL_VALIDATION_ERROR", `${toolName}: ${message}`, details);
  }
}

export class ToolAuthenticationError extends ToolError {
  constructor(message: string, details?: unknown) {
    super("TOOL_AUTHENTICATION_ERROR", message, details);
  }
}
