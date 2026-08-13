import rateLimit from "express-rate-limit";
import { ToolkitConfig } from "../config.js";

// ---------------------------------------------------------------------------
// Rate Limiting Middleware
// ---------------------------------------------------------------------------

/**
 * Creates a rate limiter based on service config.
 * Default: 100 requests per 15 minutes per IP.
 */
export function createRateLimiter(config: ToolkitConfig) {
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
