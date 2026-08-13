import express from "express";
import cors from "cors";
import { loadConfig } from "./config.js";
import { requestLogger } from "./middleware/requestLogger.js";
import { createRateLimiter } from "./middleware/rateLimiter.js";
import { AgentToolkit } from "./core/agentToolkit.js";
import { ConsoleToolLogger } from "./core/logger.js";
import { createWebhookRouter } from "./routes/webhookRoute.js";

// ── Import tools ──
import { echoTool } from "./tools/echoTool.js";
import { addTool } from "./tools/math/addTool.js";

// ---------------------------------------------------------------------------
// Toolkit Framework Microservice — Entry Point
// ---------------------------------------------------------------------------

function main(): void {
  const config = loadConfig();

  // ── Build toolkit with registered tools ──
  const toolkit = new AgentToolkit(
    [echoTool, addTool],
    new ConsoleToolLogger()
  );

  console.log("[TOOLKIT] Registered tools:", toolkit.listTools());

  // ── Create Express app ──
  const app = express();

  // ── Global middleware ──
  app.use(express.json());
  app.use(cors({ origin: config.corsOrigins }));
  app.use(requestLogger);
  app.use(createRateLimiter(config));

  // ── Health check ──
  app.get("/api/v1/health", (_req, res) => {
    res.status(200).json({
      ok: true,
      service: config.serviceName,
      version: "0.1.0",
      tools: toolkit.listTools().map((t) => t.name),
      timestamp: new Date().toISOString(),
    });
  });

  // ── VAPI webhook ──
  const webhookRouter = createWebhookRouter(toolkit);
  app.use("/api/v1", webhookRouter);

  // ── Start server ──
  app.listen(config.port, () => {
    console.log(`[TOOLKIT] Service started on port ${config.port}`);
    console.log(`[TOOLKIT] Health:  http://localhost:${config.port}/api/v1/health`);
    console.log(`[TOOLKIT] Webhook: POST http://localhost:${config.port}/api/v1/webhook`);
  });
}

main();
