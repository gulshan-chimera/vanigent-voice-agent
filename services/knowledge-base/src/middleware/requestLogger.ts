import { Request, Response, NextFunction } from "express";

export function requestLogger(
  req: Request,
  res: Response,
  next: NextFunction
): void {
  const start = Date.now();

  const requestId =
    (req.headers["x-request-id"] as string) ??
    `req_${Date.now()}_${Math.floor(Math.random() * 100000)}`;

  res.setHeader("x-request-id", requestId);

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
