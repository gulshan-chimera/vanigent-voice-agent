import { ToolDefinition } from "../../core/types";
import { ToolValidationError } from "../../core/errors";

type AddInput = {
  a: number;
  b: number;
};

type AddOutput = {
  sum: number;
};

export const addTool: ToolDefinition<AddInput, AddOutput> = {
  name: "math.add",
  description: "Adds two numbers",
  validate(input: unknown): AddInput {
    if (typeof input !== "object" || input === null) {
      throw new ToolValidationError(this.name, "Input must be an object");
    }

    const candidate = input as Partial<AddInput>;
    if (typeof candidate.a !== "number" || Number.isNaN(candidate.a)) {
      throw new ToolValidationError(this.name, "'a' must be a valid number");
    }

    if (typeof candidate.b !== "number" || Number.isNaN(candidate.b)) {
      throw new ToolValidationError(this.name, "'b' must be a valid number");
    }

    return { a: candidate.a, b: candidate.b };
  },
  execute(input: AddInput): AddOutput {
    return { sum: input.a + input.b };
  }
};
