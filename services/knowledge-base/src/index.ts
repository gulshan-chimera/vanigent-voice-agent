import express from "express";
import { loadConfig } from "./config.js";
import { KnowledgeBaseSearchClient } from "./search/searchClient.js";
import { SearchService } from "./search/searchService.js";
import { createSearchRouter } from "./routes/searchRoute.js";

// ---------------------------------------------------------------------------
// Knowledge Base Microservice — Entry Point
// ---------------------------------------------------------------------------

function main(): void {
  const config = loadConfig();

  // ── Build dependency graph ──
  const searchClient = new KnowledgeBaseSearchClient(config);
  const searchService = new SearchService(searchClient);
  const searchRouter = createSearchRouter(searchService);

  // ── Create Express app ──
  const app = express();

  app.use(express.json());

  // ── Health check ──
  app.get("/api/health", (_req, res) => {
    res.status(200).json({
      ok: true,
      service: "@vanigent/knowledge-base",
      version: "0.1.0",
      timestamp: new Date().toISOString(),
    });
  });

  // ── Search API ──
  app.use("/api", searchRouter);

  // ── Start server ──
  app.listen(config.port, () => {
    console.log(`[KNOWLEDGE-BASE] Service started on port ${config.port}`);
    console.log(`[KNOWLEDGE-BASE] Search endpoint: ${config.searchEndpoint}`);
    console.log(`[KNOWLEDGE-BASE] Index: ${config.searchIndexName}`);
    console.log(`[KNOWLEDGE-BASE] Health: http://localhost:${config.port}/api/health`);
    console.log(`[KNOWLEDGE-BASE] Search: POST http://localhost:${config.port}/api/search`);
  });
}

main();
