export type ToolLogLevel = "debug" | "info" | "warn" | "error";

export interface ToolLogEntry {
  level: ToolLogLevel;
  event: string;
  timestamp: string;
  data: Record<string, unknown>;
}

export interface ToolLogger {
  log(entry: ToolLogEntry): void;
}

export class ConsoleToolLogger implements ToolLogger {
  log(entry: ToolLogEntry): void {
    const line = JSON.stringify(entry);
    if (entry.level === "error") {
      console.error(line);
      return;
    }

    console.log(line);
  }
}

export class NoopToolLogger implements ToolLogger {
  log(_: ToolLogEntry): void {
    return;
  }
}
