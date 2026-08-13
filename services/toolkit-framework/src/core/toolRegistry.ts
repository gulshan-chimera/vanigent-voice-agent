import { ToolDefinition } from "./types";
import { ToolError, ToolNotFoundError } from "./errors";

export class ToolRegistry {
  private readonly tools = new Map<string, ToolDefinition<any, any>>();

  register<TInput, TOutput>(tool: ToolDefinition<TInput, TOutput>): void {
    if (this.tools.has(tool.name)) {
      throw new ToolError(
        "TOOL_ALREADY_REGISTERED",
        `Tool already registered: ${tool.name}`
      );
    }

    this.tools.set(tool.name, tool as ToolDefinition<any, any>);
  }

  get(toolName: string): ToolDefinition<any, any> {
    const tool = this.tools.get(toolName);
    if (!tool) {
      throw new ToolNotFoundError(toolName);
    }

    return tool;
  }

  list(): Array<{ name: string; description: string }> {
    return [...this.tools.values()].map((tool) => ({
      name: tool.name,
      description: tool.description
    }));
  }
}
