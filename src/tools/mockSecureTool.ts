import { ToolValidationError } from "../core/errors";
import { ToolDefinition } from "../core/types";

type MockSecureInput = {
  payload: string;
};

type MockSecureOutput = {
  payload: string;
  authHeader: string | null;
  agentId: string;
};

export const mockSecureTool: ToolDefinition<MockSecureInput, MockSecureOutput> = {
  name: "mock.secureEcho",
  description: "Echo payload and expose auth header for test validation",
  validate(input: unknown): MockSecureInput {
    if (typeof input !== "object" || input === null) {
      throw new ToolValidationError(this.name, "Input must be an object");
    }

    const candidate = input as Partial<MockSecureInput>;
    if (typeof candidate.payload !== "string" || candidate.payload.length === 0) {
      throw new ToolValidationError(this.name, "'payload' must be a non-empty string");
    }

    return { payload: candidate.payload };
  },
  execute(input, context): MockSecureOutput {
    const headers = (context.metadata?.headers ?? {}) as Record<string, unknown>;
    const authorization = headers.Authorization;

    return {
      payload: input.payload,
      authHeader: typeof authorization === "string" ? authorization : null,
      agentId: context.agentId
    };
  }
};
