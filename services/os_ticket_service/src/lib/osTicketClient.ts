// src/lib/osTicketClient.ts
//
// Creates a ticket through osTicket's REST API:
//   POST {OSTICKET_API_URL}   (…/api/tickets.json)
//   Header X-API-Key: {OSTICKET_API_KEY}
//
// osTicket specifics worth knowing:
//   - Success is HTTP 201 and the body is the new ticket NUMBER as plain
//     text (e.g. "123456"), not JSON.
//   - Errors come back as plain text too (400 for validation problems,
//     401 "API key not authorized" for a bad key).
//   - An osTicket API key is tied to ONE client IP address, set when the
//     key is created in osTicket admin. Calls from any other IP get 401,
//     so local testing and the deployed Function App each need their
//     outbound IP registered (one key per IP).
//
// Fails closed: never throws, always returns a result object.

import { CreateTicketInput, CreateTicketResult, OsTicketPayload } from "../types/osTicket";

const REQUEST_TIMEOUT_MS = 15000;

function getConfig(): { url: string; apiKey: string } | null {
  const url = process.env.OSTICKET_API_URL;
  const apiKey = process.env.OSTICKET_API_KEY;
  if (!url || !apiKey) {
    console.error("[OSTICKET] OSTICKET_API_URL or OSTICKET_API_KEY is not set.");
    return null;
  }
  return { url, apiKey };
}

/**
 * osTicket reads the message as a data URL. Plain text is wrapped as
 * text/plain; anything already starting with "data:" is sent unchanged.
 */
function toMessageDataUrl(message: string): string {
  return message.startsWith("data:") ? message : `data:text/plain;charset=utf-8,${message}`;
}

export function buildPayload(input: CreateTicketInput): OsTicketPayload {
  const payload: OsTicketPayload = {
    alert: input.alert ?? true,
    autorespond: input.autorespond ?? true,
    source: "API",
    name: input.name.trim(),
    email: input.email.trim(),
    subject: input.subject.trim(),
    message: toMessageDataUrl(input.message.trim()),
  };
  if (input.phone && input.phone.trim()) payload.phone = input.phone.trim();
  if (typeof input.topicId === "number") payload.topicId = input.topicId;
  return payload;
}

export async function createTicket(input: CreateTicketInput): Promise<CreateTicketResult> {
  const config = getConfig();
  if (!config) return { ok: false, status: 500, error: "Ticket service is not configured." };

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);

  try {
    const response = await fetch(config.url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-API-Key": config.apiKey,
      },
      body: JSON.stringify(buildPayload(input)),
      signal: controller.signal,
    });

    const text = (await response.text()).trim();

    if (response.status === 201) {
      console.log(`[OSTICKET] Ticket created: #${text}`);
      return { ok: true, ticketNumber: text };
    }

    console.error(`[OSTICKET] Create failed (${response.status}): ${text.slice(0, 300)}`);
    return {
      ok: false,
      status: response.status,
      error: text.slice(0, 300) || `osTicket returned HTTP ${response.status}`,
    };
  } catch (error) {
    const aborted = (error as Error).name === "AbortError";
    console.error(
      `[OSTICKET] ${aborted ? `Timed out after ${REQUEST_TIMEOUT_MS}ms` : `Network error: ${(error as Error).message}`}`
    );
    return {
      ok: false,
      status: aborted ? 504 : 502,
      error: aborted ? "osTicket did not respond in time." : "Could not reach osTicket.",
    };
  } finally {
    clearTimeout(timer);
  }
}