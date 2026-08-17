import { Request, Response, NextFunction } from "express";

export function authMiddleware(
  req: Request,
  res: Response,
  next: NextFunction
): void {
  const apiKey = req.headers["x-api-key"];
  const validKey = process.env.KB_API_KEY;

  if (!validKey) {
    console.warn(
      "[SECURITY WARNING] KB_API_KEY is not set in environment. All requests will be rejected."
    );
    res.status(500).json({ error: "Server authentication misconfigured." });
    return;
  }

  if (!apiKey || apiKey !== validKey) {
    console.warn(`[AUTH] Rejected request from IP: ${req.ip} (Invalid Key)`);
    res.status(401).json({ error: "Unauthorized. Invalid x-api-key." });
    return;
  }

  next();
}
