import * as dotenv from "dotenv";

// Load .env file if present (ignored in production where env vars are set directly)
dotenv.config();

// ---------------------------------------------------------------------------
// Required environment variables
// ---------------------------------------------------------------------------

function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value || value.trim().length === 0) {
    throw new Error(`[CONFIG] Missing required environment variable: ${name}`);
  }
  return value.trim();
}

function optionalEnv(name: string, defaultValue: string): string {
  const value = process.env[name];
  return value && value.trim().length > 0 ? value.trim() : defaultValue;
}

// ---------------------------------------------------------------------------
// Service configuration
// ---------------------------------------------------------------------------

export interface KnowledgeBaseConfig {
  // Azure AI Search
  searchServiceName: string;
  searchApiKey: string;
  searchIndexName: string;
  searchEndpoint: string;

  // Microsoft Graph (for indexing scripts)
  tenantId: string;
  clientId: string;
  clientSecret: string;

  // SharePoint
  spHostname: string;
  spSitePath: string;

  // Server
  port: number;
}

export function loadConfig(): KnowledgeBaseConfig {
  const searchServiceName = requireEnv("SEARCH_SERVICE_NAME");

  return {
    // Azure AI Search
    searchServiceName,
    searchApiKey: requireEnv("SEARCH_API_KEY"),
    searchIndexName: optionalEnv("SEARCH_INDEX_NAME", "sharepoint-index"),
    searchEndpoint: `https://${searchServiceName}.search.windows.net`,

    // Microsoft Graph — required for indexing, optional for search-only mode
    tenantId: optionalEnv("TENANT_ID", ""),
    clientId: optionalEnv("CLIENT_ID", ""),
    clientSecret: optionalEnv("CLIENT_SECRET", ""),

    // SharePoint
    spHostname: optionalEnv("SP_HOSTNAME", "chimeratechpvtltd.sharepoint.com"),
    spSitePath: optionalEnv("SP_SITE_PATH", "/sites/Corporate-Policies"),

    // Server
    port: parseInt(optionalEnv("PORT", "3001"), 10),
  };
}

/**
 * Validates that Microsoft Graph credentials are present.
 * Required for indexing scripts but not for the search API.
 */
export function requireGraphCredentials(config: KnowledgeBaseConfig): void {
  if (!config.tenantId) {
    throw new Error("[CONFIG] TENANT_ID is required for this operation");
  }
  if (!config.clientId) {
    throw new Error("[CONFIG] CLIENT_ID is required for this operation");
  }
  if (!config.clientSecret) {
    throw new Error("[CONFIG] CLIENT_SECRET is required for this operation");
  }
}
