// src/lib/verifyWebhookAuth.ts
//
// Verifies that an incoming request to our webhook genuinely carries the
// shared secret configured on the VAPI side (via a Bearer Token credential).

import { HttpRequest } from "@azure/functions";
import { timingSafeEqual } from "crypto";

export function isRequestAuthorized(request: HttpRequest): boolean {
  const expectedSecret = process.env.VAPI_WEBHOOK_SECRET;

  if (!expectedSecret) {
    console.error("[WEBHOOK-AUTH] VAPI_WEBHOOK_SECRET is not set — rejecting all requests.");
    return false;
  }

  const authHeader = request.headers.get("authorization");

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    console.warn("[WEBHOOK-AUTH] Missing or malformed Authorization header.");
    return false;
  }

  const providedSecret = authHeader.slice("Bearer ".length);

  const expectedBuffer = Buffer.from(expectedSecret);
  const providedBuffer = Buffer.from(providedSecret);

  if (expectedBuffer.length !== providedBuffer.length) {
    return false;
  }

  return timingSafeEqual(expectedBuffer, providedBuffer);
}