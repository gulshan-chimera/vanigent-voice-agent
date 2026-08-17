import express from "express";
import cors from "cors";
import { loadConfig } from "./config.js";
import { KnowledgeBaseSearchClient } from "./search/searchClient.js";
import { SearchService } from "./search/searchService.js";
import { createSearchRouter } from "./routes/searchRoute.js";
import { SyncService } from "./sync/syncService.js";
import { createSyncRouter } from "./routes/syncRoute.js";
import { requestLogger } from "./middleware/requestLogger.js";
import { createRateLimiter } from "./middleware/rateLimiter.js";

// ---------------------------------------------------------------------------
// Knowledge Base Microservice — Entry Point
// ---------------------------------------------------------------------------

function main(): void {
  const config = loadConfig();

  // ── Build dependency graph ──
  const searchClient = new KnowledgeBaseSearchClient(config);
  const searchService = new SearchService(searchClient);
  const searchRouter = createSearchRouter(searchService);
  const syncService = new SyncService();
  const syncRouter = createSyncRouter(syncService);

  // ── Create Express app ──
  const app = express();

  app.use(express.json());
  app.use(cors({ origin: config.corsOrigins }));
  app.use(requestLogger);
  app.use(createRateLimiter(config));

  // ── Health check ──
  app.get("/api/v1/health", (_req, res) => {
    res.status(200).json({
      ok: true,
      service: "@vanigent/knowledge-base",
      version: "0.1.0",
      timestamp: new Date().toISOString(),
    });
  });

  // ── Search API ──
  app.use("/api/v1", searchRouter);
  app.use("/api/v1", syncRouter);

  // ── Start server ──
  app.listen(config.port, () => {
    console.log(`[KNOWLEDGE-BASE] Service started on port ${config.port}`);
    console.log(`[KNOWLEDGE-BASE] Search endpoint: ${config.searchEndpoint}`);
    console.log(`[KNOWLEDGE-BASE] Index: ${config.searchIndexName}`);
    console.log(`[KNOWLEDGE-BASE] Health: http://localhost:${config.port}/api/v1/health`);
    console.log(`[KNOWLEDGE-BASE] Search: POST http://localhost:${config.port}/api/v1/search`);
    console.log(`[KNOWLEDGE-BASE] Sync: POST http://localhost:${config.port}/api/v1/sync`);
  });
}

main();


