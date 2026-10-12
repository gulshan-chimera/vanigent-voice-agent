// src/lib/verifyRequestAuth.ts
//
// Bearer-secret check for every endpoint in this service. Uses its own
// secret (OSTICKET_SERVICE_SECRET), separate from the other services, so
// a leak of one secret doesn't open the others. Timing-safe comparison
// so the secret can't be guessed from response times.

import { HttpRequest } from "@azure/functions";
import { timingSafeEqual } from "crypto";

export function isRequestAuthorized(request: HttpRequest): boolean {
  const expectedSecret = process.env.OSTICKET_SERVICE_SECRET;

  if (!expectedSecret) {
    console.error("[OSTICKET-AUTH] OSTICKET_SERVICE_SECRET is not set — rejecting all requests.");
    return false;
  }

  const authHeader = request.headers.get("authorization");
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    console.warn("[OSTICKET-AUTH] Missing or malformed Authorization header.");
    return false;
  }

  const expectedBuffer = Buffer.from(expectedSecret);
  const providedBuffer = Buffer.from(authHeader.slice("Bearer ".length));

  // timingSafeEqual throws on different lengths, so check that first.
  if (expectedBuffer.length !== providedBuffer.length) return false;

  return timingSafeEqual(expectedBuffer, providedBuffer);
}