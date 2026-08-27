// src/lib/verifyWebhookAuth.ts
//
// Verifies that an incoming request carries the shared secret for this
// service (Authorization: Bearer <secret>). Uses a timing-safe
// comparison so the secret can't be recovered by measuring how long
// a rejection takes.
//
// This service's secret is deliberately separate from the Authentication
// service's VAPI secret — they are separate deployments, so a leak of
// one does not compromise the other.

import { HttpRequest } from "@azure/functions";
import { timingSafeEqual } from "crypto";

export function isRequestAuthorized(request: HttpRequest): boolean {
  const expectedSecret = process.env.KB_API_SECRET;

  if (!expectedSecret) {
    console.error("[KB-AUTH] KB_API_SECRET is not set — rejecting all requests.");
    return false;
  }

  const authHeader = request.headers.get("authorization");

  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    console.warn("[KB-AUTH] Missing or malformed Authorization header.");
    return false;
  }

  const providedSecret = authHeader.slice("Bearer ".length);

  const expectedBuffer = Buffer.from(expectedSecret);
  const providedBuffer = Buffer.from(providedSecret);

  // timingSafeEqual throws on length mismatch, so check that first.
  if (expectedBuffer.length !== providedBuffer.length) {
    return false;
  }

  return timingSafeEqual(expectedBuffer, providedBuffer);
}