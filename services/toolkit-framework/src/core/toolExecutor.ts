import { ToolError, ToolValidationError } from "./errors";
import { NoopToolLogger, ToolLogger } from "./logger";
import { ToolRegistry } from "./toolRegistry";
import {
  ToolCallFailureResponse,
  ToolCallRequest,
  ToolCallResponse,
  ToolExecutionMeta,
  ToolExecutionResult
} from "./types";

export class ToolExecutor {
  constructor(
    private readonly registry: ToolRegistry,
    private readonly logger: ToolLogger = new NoopToolLogger()
  ) {}

  async execute<TOutput>(
    request: ToolCallRequest
  ): Promise<ToolExecutionResult<TOutput>> {
    const started = Date.now();
    const startedAt = new Date(started).toISOString();
    const { requestId, toolName } = request;

    this.safeLog("info", "tool.call.started", {
      requestId,
      toolName,
      agentId: request.context.agentId,
      traceId: request.context.traceId
    });

    try {
      const tool = this.registry.get(toolName);

      let validatedInput: unknown;
      try {
        validatedInput = tool.validate(request.input);
      } catch (error) {
        throw new ToolValidationError(
          toolName,
          "Input validation failed",
          error
        );
      }

      const data = await tool.execute(validatedInput, request.context);
      const finished = Date.now();
      const finishedAt = new Date(finished).toISOString();

      const response: ToolCallResponse<TOutput> = {
        ok: true,
        requestId,
        toolName,
        output: data as TOutput,
        data: data as TOutput,
        meta: {
          toolName,
          requestId,
          agentId: request.context.agentId,
          traceId: request.context.traceId,
          durationMs: finished - started,
          startedAt,
          finishedAt
        }
      };

      this.safeLog("info", "tool.call.succeeded", {
        requestId,
        toolName,
        agentId: request.context.agentId,
        traceId: request.context.traceId,
        durationMs: response.meta.durationMs
      });

      return response;
    } catch (error) {
      const finished = Date.now();
      const response = this.toFailureResult(request, startedAt, finished, error);

      this.safeLog("error", "tool.call.failed", {
        requestId,
        toolName,
        agentId: request.context.agentId,
        traceId: request.context.traceId,
        durationMs: response.meta.durationMs,
        errorCode: response.error.code,
        errorMessage: response.error.message
      });

      return response;
    }
  }

  private toFailureResult(
    request: ToolCallRequest,
    startedAt: string,
    finished: number,
    error: unknown
  ): ToolCallFailureResponse {
    const { toolName, requestId } = request;
    const meta: ToolExecutionMeta = {
      toolName,
      requestId,
      agentId: request.context.agentId,
      traceId: request.context.traceId,
      durationMs: finished - new Date(startedAt).getTime(),
      startedAt,
      finishedAt: new Date(finished).toISOString()
    };

    if (error instanceof ToolError) {
      return {
        ok: false,
        requestId,
        toolName,
        error: {
          code: error.code,
          message: error.message,
          details: error.details
        },
        meta
      };
    }

    return {
      ok: false,
      requestId,
      toolName,
      error: {
        code: "TOOL_EXECUTION_ERROR",
        message: this.getErrorMessage(error),
        details: this.serializeError(error)
      },
      meta
    };
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }

    return "Tool execution failed";
  }

  private serializeError(error: unknown): unknown {
    if (error instanceof Error) {
      return {
        name: error.name,
        message: error.message,
        stack: error.stack
      };
    }

    return error;
  }

  private safeLog(
    level: "debug" | "info" | "warn" | "error",
    event: string,
    data: Record<string, unknown>
  ): void {
    try {
      this.logger.log({
        level,
        event,
        timestamp: new Date().toISOString(),
        data
      });
    } catch {
      return;
    }
  }
}
