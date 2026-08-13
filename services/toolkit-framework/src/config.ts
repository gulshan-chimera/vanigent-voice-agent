import * as dotenv from "dotenv";

dotenv.config();

// ---------------------------------------------------------------------------
// Environment variable helpers
// ---------------------------------------------------------------------------

function optionalEnv(name: string, defaultValue: string): string {
  const value = process.env[name];
  return value && value.trim().length > 0 ? value.trim() : defaultValue;
}

// ---------------------------------------------------------------------------
// Service configuration
// ---------------------------------------------------------------------------

export interface ToolkitConfig {
  /** HTTP server port */
  port: number;

  /** Service name (used in logs and health check) */
  serviceName: string;

  /** CORS allowed origins (comma-separated, or "*") */
  corsOrigins: string;

  /** Rate limit: max requests per window */
  rateLimitMax: number;

  /** Rate limit: window duration in minutes */
  rateLimitWindowMinutes: number;

  /** Knowledge Base service URL (for tool calls that need KB search) */
  knowledgeBaseUrl: string;
}

export function loadConfig(): ToolkitConfig {
  return {
    port: parseInt(optionalEnv("PORT", "3000"), 10),
    serviceName: optionalEnv("SERVICE_NAME", "@vanigent/toolkit-framework"),
    corsOrigins: optionalEnv("CORS_ORIGINS", "*"),
    rateLimitMax: parseInt(optionalEnv("RATE_LIMIT_MAX", "100"), 10),
    rateLimitWindowMinutes: parseInt(optionalEnv("RATE_LIMIT_WINDOW_MINUTES", "15"), 10),
    knowledgeBaseUrl: optionalEnv("KNOWLEDGE_BASE_URL", "http://localhost:3001"),
  };
}
