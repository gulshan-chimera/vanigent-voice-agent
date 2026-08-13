import { Request, Response, NextFunction } from "express";

// ---------------------------------------------------------------------------
// Request Logging Middleware
// ---------------------------------------------------------------------------

/**
 * Logs every incoming HTTP request with:
 * - Method, path, status code, response time
 * - Unique requestId (from header or auto-generated)
 *
 * Attach this before your routes.
 */
export function requestLogger(
  req: Request,
  res: Response,
  next: NextFunction
): void {
  const start = Date.now();

  // Use existing requestId header or generate one
  const requestId =
    (req.headers["x-request-id"] as string) ??
    `req_${Date.now()}_${Math.floor(Math.random() * 100000)}`;

  // Attach requestId to response headers for tracing
  res.setHeader("x-request-id", requestId);

  // Log when the response finishes
  res.on("finish", () => {
    const duration = Date.now() - start;
    const log = {
      timestamp: new Date().toISOString(),
      requestId,
      method: req.method,
      path: req.originalUrl,
      statusCode: res.statusCode,
      durationMs: duration,
      userAgent: req.headers["user-agent"] ?? "unknown",
    };

    if (res.statusCode >= 400) {
      console.error("[REQUEST]", JSON.stringify(log));
    } else {
      console.log("[REQUEST]", JSON.stringify(log));
    }
  });

  next();
}
