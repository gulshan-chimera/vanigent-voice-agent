import { ToolDefinition } from "../core/types";
import { ToolValidationError } from "../core/errors";

type EchoInput = {
  message: string;
};

type EchoOutput = {
  echoed: string;
};

export const echoTool: ToolDefinition<EchoInput, EchoOutput> = {
  name: "utility.echo",
  description: "Returns the same message back",
  validate(input: unknown): EchoInput {
    if (typeof input !== "object" || input === null) {
      throw new ToolValidationError(this.name, "Input must be an object");
    }

    const candidate = input as Partial<EchoInput>;
    if (typeof candidate.message !== "string" || candidate.message.length === 0) {
      throw new ToolValidationError(this.name, "'message' must be a non-empty string");
    }

    return { message: candidate.message };
  },
  execute(input: EchoInput): EchoOutput {
    return { echoed: input.message };
  }
};
