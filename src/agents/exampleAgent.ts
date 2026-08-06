import { AgentToolkit } from "../core/agentToolkit";
import { ToolCallContext } from "../core/types";

export class ExampleAgent {
  constructor(private readonly toolkit: AgentToolkit) {}

  async handleUserMessage(message: string): Promise<string> {
    const context: ToolCallContext = {
      agentId: "example-agent",
      requestId: crypto.randomUUID()
    };

    if (message.startsWith("add ")) {
      const [a, b] = message
        .replace("add ", "")
        .split(" ")
        .map((v) => Number(v));

      const result = await this.toolkit.call<{ a: number; b: number }, { sum: number }>(
        "math.add",
        { a, b },
        context
      );

      if (!result.ok) {
        return `Tool error (${result.error.code}): ${result.error.message}`;
      }

      return `Sum is ${result.data.sum}`;
    }

    const echo = await this.toolkit.call<{ message: string }, { echoed: string }>(
      "utility.echo",
      { message },
      context
    );

    if (!echo.ok) {
      return `Tool error (${echo.error.code}): ${echo.error.message}`;
    }

    return `Echo: ${echo.data.echoed}`;
  }
}
