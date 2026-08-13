import rateLimit from "express-rate-limit";
import { KnowledgeBaseConfig } from "../config.js";

export function createRateLimiter(config: KnowledgeBaseConfig) {
  return rateLimit({
    windowMs: config.rateLimitWindowMinutes * 60 * 1000,
    max: config.rateLimitMax,
    standardHeaders: true,
    legacyHeaders: false,
    message: {
      ok: false,
      error: {
        code: "RATE_LIMIT_EXCEEDED",
        message: `Too many requests. Limit: ${config.rateLimitMax} per ${config.rateLimitWindowMinutes} minutes.`,
      },
    },
  });
}
