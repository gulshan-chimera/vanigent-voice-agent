export interface ToolCallContext {
  agentId: string;
  requestId?: string;
  traceId?: string;
  auth?: {
    entra?: {
      tokenType: "Bearer";
      accessToken: string;
      expiresOn?: string;
      scopes?: string[];
    };
  };
  metadata?: Record<string, unknown>;
}

export interface ToolCallRequest<TInput = unknown> {
  requestId: string;
  toolName: string;
  input: TInput;
  context: ToolCallContext;
  createdAt: string;
}

export interface ToolExecutionMeta {
  toolName: string;
  requestId: string;
  agentId: string;
  traceId?: string;
  durationMs: number;
  startedAt: string;
  finishedAt: string;
}

export interface ToolCallSuccessResponse<TOutput> {
  ok: true;
  requestId: string;
  toolName: string;
  output: TOutput;
  data: TOutput;
  meta: ToolExecutionMeta;
}

export interface ToolCallFailureResponse {
  ok: false;
  requestId: string;
  toolName: string;
  error: {
    code: string;
    message: string;
    details?: unknown;
  };
  meta: ToolExecutionMeta;
}

export type ToolCallResponse<TOutput> =
  | ToolCallSuccessResponse<TOutput>
  | ToolCallFailureResponse;

export type ToolExecutionResult<TOutput> = ToolCallResponse<TOutput>;

export interface ToolDefinition<TInput, TOutput> {
  name: string;
  description: string;
  validate: (input: unknown) => TInput;
  execute: (
    input: TInput,
    context: ToolCallContext
  ) => Promise<TOutput> | TOutput;
}

export interface ToolInvoker {
  invoke<TInput, TOutput>(
    request: ToolCallRequest<TInput>
  ): Promise<ToolCallResponse<TOutput>>;

  call<TInput, TOutput>(
    toolName: string,
    input: TInput,
    context: ToolCallContext
  ): Promise<ToolCallResponse<TOutput>>;
}
