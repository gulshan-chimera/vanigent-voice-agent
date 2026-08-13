import { ToolAuthenticationError } from "../core/errors";
import {
  ToolCallContext,
  ToolCallRequest,
  ToolCallResponse,
  ToolInvoker
} from "../core/types";

export interface EntraAccessToken {
  accessToken: string;
  expiresOn?: string;
  tokenType?: "Bearer";
}

export interface EntraTokenProvider {
  getAccessToken(
    scopes: string[],
    context: ToolCallContext
  ): Promise<EntraAccessToken>;
}

export interface EntraAuthOptions {
  scopes: string[];
  tokenProvider: EntraTokenProvider;
  headerName?: string;
}

export class EntraAuthenticatedInvoker implements ToolInvoker {
  private readonly headerName: string;

  constructor(
    private readonly baseInvoker: ToolInvoker,
    private readonly options: EntraAuthOptions
  ) {
    this.headerName = options.headerName ?? "Authorization";
  }

  async invoke<TInput, TOutput>(
    request: ToolCallRequest<TInput>
  ): Promise<ToolCallResponse<TOutput>> {
    const started = Date.now();
    const startedAt = new Date(started).toISOString();

    try {
      const token = await this.options.tokenProvider.getAccessToken(
        this.options.scopes,
        request.context
      );

      const tokenType = token.tokenType ?? "Bearer";
      const withAuth = this.injectToken(request, token.accessToken, tokenType, token.expiresOn);

      return this.baseInvoker.invoke<TInput, TOutput>(withAuth);
    } catch (error) {
      const authError = new ToolAuthenticationError(
        "Failed to acquire Entra access token",
        error
      );
      const finished = Date.now();

      return {
        ok: false,
        requestId: request.requestId,
        toolName: request.toolName,
        error: {
          code: authError.code,
          message: authError.message,
          details: authError.details
        },
        meta: {
          toolName: request.toolName,
          requestId: request.requestId,
          agentId: request.context.agentId,
          traceId: request.context.traceId,
          durationMs: finished - started,
          startedAt,
          finishedAt: new Date(finished).toISOString()
        }
      };
    }
  }

  async call<TInput, TOutput>(
    toolName: string,
    input: TInput,
    context: ToolCallContext
  ): Promise<ToolCallResponse<TOutput>> {
    const requestId = context.requestId ?? this.createRequestId();

    return this.invoke<TInput, TOutput>({
      requestId,
      toolName,
      input,
      context: {
        ...context,
        requestId
      },
      createdAt: new Date().toISOString()
    });
  }

  private injectToken<TInput>(
    request: ToolCallRequest<TInput>,
    accessToken: string,
    tokenType: "Bearer",
    expiresOn?: string
  ): ToolCallRequest<TInput> {
    const rawHeaders = request.context.metadata?.headers;
    const existingHeaders =
      typeof rawHeaders === "object" && rawHeaders !== null
        ? (rawHeaders as Record<string, unknown>)
        : {};

    return {
      ...request,
      context: {
        ...request.context,
        auth: {
          ...request.context.auth,
          entra: {
            tokenType,
            accessToken,
            expiresOn,
            scopes: this.options.scopes
          }
        },
        metadata: {
          ...(request.context.metadata ?? {}),
          headers: {
            ...existingHeaders,
            [this.headerName]: `${tokenType} ${accessToken}`
          }
        }
      }
    };
  }

  private createRequestId(): string {
    return `req_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
  }
}
