# Vanigent Voice AI Toolkit Service

This repository is a monorepo containing the Azure Functions services that power the Vanigent Voice AI agent platform. Each service under `services/` is fully independent — its own `package.json`, `tsconfig.json`, `.env.example`, and `azure-pipelines.yml` — and is built with Node.js/TypeScript on the Azure Functions v4 programming model.

```
Vaniget_Voice_AI_Toolkit_Service/
├── services/
│   ├── Authentication/      # VAPI caller authentication gateway
│   └── Knowledge Base/      # SharePoint-backed RAG knowledge base
```

## 1. Authentication (`services/Authentication`)

Authenticates a phone caller against Microsoft Entra ID during a VAPI call.

- **`vapiWebhook`** (POST) — VAPI's webhook target. Verifies the request with a shared bearer secret (`VAPI_WEBHOOK_SECRET`, timing-safe compare) before doing anything else, then looks up the caller via Microsoft Graph.
- Graph access uses an app-only (client-credentials) token against the Entra ID app registration configured by `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET`.

## 2. Knowledge Base (`services/Knowledge Base`)

A Retrieval-Augmented Generation service backed by SharePoint documents and Azure AI Search.

**Ingestion**
- Reads one or more SharePoint document libraries (selected via `KB_LIBRARY_ALLOWLIST` or `KB_LIBRARY_PATTERN`) via Microsoft Graph, recursing into subfolders.
- Sync is incremental: each file's `cTag` (a content-only change marker) is diffed against what's already indexed. New/changed files are queued for indexing; files deleted from SharePoint have their index chunks removed; unchanged files are skipped entirely — no download, no AI calls.
- Indexing runs off an Azure Storage Queue, decoupled from the sync trigger itself, so a slow file (many pages, many vision calls) isn't bound by an HTTP request timeout. Failed jobs retry automatically and land in a poison queue after repeated failures.
- Sync runs automatically on a schedule (`syncTimer`, configured by `KB_SYNC_SCHEDULE`) and can also be triggered manually (`syncIndex`, POST, bearer-secret protected, supports `?force=true` for a full rebuild).

**Supported formats — PDF, DOCX, PPTX**
- PDF is processed natively: each page becomes one chunk, combining the page's extracted text with an Azure OpenAI vision-generated caption of its visual content (screenshots, diagrams, charts).
- DOCX and PPTX have no rendering engine of their own, so they're first converted to PDF server-side by Microsoft Graph (`?format=pdf`) and then processed through the exact same per-page pipeline as native PDFs — one Word page or one PowerPoint slide becomes one PDF page becomes one chunk, captioned the same way.
- Anything outside these three formats is skipped and reported (not silently dropped) in the sync summary.

**Retrieval**
- **`askKnowledgeBase`** (POST) — embeds a question, vector-searches the index, and generates a grounded, spoken-style answer via Azure OpenAI. Used for internal testing.
- **`vapiWebhook`** (POST) — the VAPI-facing tool-call endpoint backing the `search_hr_knowledge_base` tool; same underlying answer logic, VAPI's request/response shape.
- **`queryIndex`** (POST) — raw vector search, returns matches and scores without generating an answer.
- **`queueStatus`** (GET) — reports pending/poisoned/permanently-failed indexing jobs, so a sync's health can be checked without watching logs.

## CI/CD

Each service has its own Azure DevOps pipeline (`azure-pipelines.yml`) triggered on pushes/PRs to `develop` and `release1`, building and deploying to per-environment Azure Function Apps (dev / UAT).

## Getting Started

Each service is run independently:

```bash
cd "services/Authentication"   # or "services/Knowledge Base"
npm install
npm run build
npm run start   # builds, then runs the Azure Functions host locally (requires Azure Functions Core Tools)
```

## Configuration

Each service documents its required environment variables in its own `.env.example`. For local development, copy the values into a `local.settings.json` in that service's folder (this file is git-ignored — it holds real secrets and must never be committed). For a deployed Function App, set the same variables as Application Settings in the Azure Portal.
