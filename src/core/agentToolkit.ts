import { NoopToolLogger, ToolLogger } from "./logger";
import { ToolExecutor } from "./toolExecutor";
import { ToolRegistry } from "./toolRegistry";
import {
  ToolCallRequest,
  ToolCallResponse,
  ToolCallContext,
  ToolDefinition,
  ToolExecutionResult,
  ToolInvoker
} from "./types";

export class AgentToolkit implements ToolInvoker {
  private readonly registry: ToolRegistry;
  private readonly executor: ToolExecutor;

  constructor(
    initialTools: ToolDefinition<any, any>[] = [],
    logger: ToolLogger = new NoopToolLogger()
  ) {
    this.registry = new ToolRegistry();
    this.executor = new ToolExecutor(this.registry, logger);

    for (const tool of initialTools) {
      this.registry.register(tool);
    }
  }

  register<TInput, TOutput>(tool: ToolDefinition<TInput, TOutput>): void {
    this.registry.register(tool);
  }

  listTools(): Array<{ name: string; description: string }> {
    return this.registry.list();
  }

  async invoke<TInput, TOutput>(
    request: ToolCallRequest<TInput>
  ): Promise<ToolCallResponse<TOutput>> {
    return this.executor.execute<TOutput>(request);
  }

  async call<TInput, TOutput>(
    toolName: string,
    input: TInput,
    context: ToolCallContext
  ): Promise<ToolExecutionResult<TOutput>> {
    const requestId = context.requestId ?? this.createRequestId();
    const request: ToolCallRequest<TInput> = {
      requestId,
      toolName,
      input,
      context: {
        ...context,
        requestId
      },
      createdAt: new Date().toISOString()
    };

    return this.invoke<TInput, TOutput>(request);
  }

  private createRequestId(): string {
    return `req_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
  }
}
