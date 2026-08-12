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

    if (message.startsWith("weather ")) {
      const parts = message.replace("weather ", "").split(" ");
      const city = parts[0];
      const units = parts[1] === "fahrenheit" ? "fahrenheit" : "celsius";

      const result = await this.toolkit.call<
        { city: string; units: string },
        { city: string; temperature: number; units: string; condition: string; humidity: number; windSpeed: number; windUnits: string; description: string }
      >("weather.get", { city, units }, context);

      if (!result.ok) {
        return `Tool error (${result.error.code}): ${result.error.message}`;
      }

      const w = result.data;
      return `Weather in ${w.city}: ${w.condition}, ${w.temperature}°${w.units === "fahrenheit" ? "F" : "C"}, Humidity: ${w.humidity}%, Wind: ${w.windSpeed} ${w.windUnits}`;
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

