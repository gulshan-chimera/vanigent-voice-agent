# Vanigent Voice AI Toolkit Service

This repository is a monorepo containing the microservices that power the Vanigent Voice AI agent platform.

## Architecture

The project is structured as a set of independent, loosely-coupled microservices located in the `services/` directory.

```
Vaniget_Voice_AI_Toolkit_Service/
├── services/
│   ├── toolkit-framework/      # VAPI Webhook Handler & Tool Execution Service
│   └── knowledge-base/         # RAG & Document Retrieval Service
```

### 1. Toolkit Framework (`services/toolkit-framework`)
An Express microservice that acts as the bridge between VAPI (the voice AI provider) and our internal tools.

- **VAPI Webhook Endpoint:** Handles `assistant-request`, `function-call`, `end-of-call-report`, etc.
- **Tool Registry & Executor:** A strongly-typed generic tool-calling framework.
- **Tools:** Includes mathematical utilities and mock secure endpoints.
- **Authentication:** Uses Entra ID for secure, authenticated tool invocations.

### 2. Knowledge Base (`services/knowledge-base`)
A dedicated microservice providing RAG (Retrieval-Augmented Generation) capabilities, powered by Azure AI Search and SharePoint.

- **Search API:** Exposes endpoints to query the knowledge base.
- **Caching:** Uses an LRU cache (10-minute TTL) to optimize frequently retrieved queries.
- **Access Control:** Integrates Entra Group IDs to ensure users only see documents they have permissions for.
- **Indexing Scripts:** Includes scripts to pull PDFs from SharePoint via Microsoft Graph API, redact PII via regex, and push to Azure AI Search.

## Production Readiness

Both services are equipped with standardized production middleware:
- **CORS:** Controlled via the `CORS_ORIGINS` environment variable.
- **Rate Limiting:** Protects endpoints from abuse (Default: 100 requests / 15 minutes).
- **Request Logging:** Logs every HTTP call with a unique `x-request-id` trace ID.
- **API Versioning:** All routes are versioned (e.g., `/api/v1/`).
- **Health Checks:** Both services expose a `/api/v1/health` endpoint for monitoring.

## Getting Started

Each service is completely independent with its own `package.json`, `tsconfig.json`, and `.env.example`.

### Running the Toolkit Framework
```bash
cd services/toolkit-framework
npm install
npm run build
npm run dev
```

### Running the Knowledge Base
```bash
cd services/knowledge-base
npm install
npm run build
npm run dev
```

### Indexing SharePoint Documents
To crawl SharePoint and index documents into Azure AI Search, navigate to the `knowledge-base` service and run:
```bash
npm run build
npm run index-docs
```
*(Ensure all required Microsoft Graph API credentials are set in your `.env` file first).*
