// src/functions/createTicket.ts
//
// POST /api/createTicket
// Authorization: Bearer <OSTICKET_SERVICE_SECRET>
//
// Body (JSON):
//   { "name", "email", "subject", "message",                   ← required
//     "phone", "topic", "topicId", "alert", "autorespond" }    ← optional
//
// The help topic routes the ticket to its department. Give either
//   "topic"   – the topic name, e.g. "HR - Benefits" (resolved via the
//               OSTICKET_HELP_TOPICS app setting, see lib/helpTopics.ts), or
//   "topicId" – the raw osTicket help topic ID.
// With neither, osTicket uses its default help topic.
//
// 201 → { success: true,  ticketNumber }
// 4xx/5xx → { success: false, error }

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyRequestAuth";
import { createTicket } from "../lib/osTicketClient";
import { resolveTopicId } from "../lib/helpTopics";
import { CreateTicketInput } from "../types/osTicket";

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const LIMITS = { name: 128, email: 255, subject: 255, message: 20000, phone: 32 };

type ValidationResult = { input: CreateTicketInput } | { errors: string[] } | { configError: string };

function validate(body: unknown): ValidationResult {
  const b = (body ?? {}) as Record<string, unknown>;
  const errors: string[] = [];

  const str = (key: keyof typeof LIMITS, required: boolean): string | undefined => {
    const value = b[key];
    if (value === undefined || value === null || value === "") {
      if (required) errors.push(`"${key}" is required.`);
      return undefined;
    }
    if (typeof value !== "string") {
      errors.push(`"${key}" must be a string.`);
      return undefined;
    }
    if (value.trim().length === 0 && required) {
      errors.push(`"${key}" is required.`);
      return undefined;
    }
    if (value.length > LIMITS[key]) {
      errors.push(`"${key}" must be at most ${LIMITS[key]} characters.`);
      return undefined;
    }
    return value;
  };

  const name = str("name", true);
  const email = str("email", true);
  const subject = str("subject", true);
  const message = str("message", true);
  const phone = str("phone", false);

  if (email && !EMAIL_PATTERN.test(email.trim())) errors.push(`"email" is not a valid email address.`);
  if (phone && !/^[0-9+()\-.\s]+([xX]\d+)?$/.test(phone.trim())) {
    errors.push(`"phone" may contain only digits, spaces, + ( ) - . and an optional X extension.`);
  }

  for (const key of ["alert", "autorespond"] as const) {
    if (b[key] !== undefined && typeof b[key] !== "boolean") errors.push(`"${key}" must be true or false.`);
  }
  if (b.topicId !== undefined && (typeof b.topicId !== "number" || !Number.isInteger(b.topicId))) {
    errors.push(`"topicId" must be a whole number.`);
  }

  let topicId = b.topicId as number | undefined;
  if (b.topic !== undefined && b.topic !== null && b.topic !== "") {
    if (typeof b.topic !== "string") {
      errors.push(`"topic" must be a string.`);
    } else if (b.topicId !== undefined) {
      errors.push(`Send either "topic" or "topicId", not both.`);
    } else {
      const lookup = resolveTopicId(b.topic);
      if (lookup.ok) {
        topicId = lookup.topicId;
      } else if (lookup.reason === "not_configured") {
        return { configError: `"topic" is not supported: OSTICKET_HELP_TOPICS is not configured.` };
      } else {
        errors.push(`Unknown "topic" "${b.topic}". Known topics: ${lookup.knownTopics.join(", ")}.`);
      }
    }
  }

  if (errors.length > 0) return { errors };

  return {
    input: {
      name: name!,
      email: email!,
      subject: subject!,
      message: message!,
      phone,
      topicId,
      alert: b.alert as boolean | undefined,
      autorespond: b.autorespond as boolean | undefined,
    },
  };
}

export async function createTicketHandler(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[CREATE-TICKET] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { success: false, error: "Unauthorized" } };
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return { status: 400, jsonBody: { success: false, error: "Body must be valid JSON." } };
  }

  const result = validate(body);
  if ("configError" in result) {
    context.error(`[CREATE-TICKET] ${result.configError}`);
    return { status: 500, jsonBody: { success: false, error: result.configError } };
  }
  if ("errors" in result) {
    return { status: 400, jsonBody: { success: false, error: "Validation failed.", details: result.errors } };
  }

  context.log(`[CREATE-TICKET] Creating ticket — subject: "${result.input.subject}"`);
  const created = await createTicket(result.input);

  if (created.ok) {
    return { status: 201, jsonBody: { success: true, ticketNumber: created.ticketNumber } };
  }

  // Pass osTicket's own 400 (validation) through; anything else from
  // osTicket — including a 401 for its API key — is OUR config problem,
  // not the caller's, so it's reported as 502 Bad Gateway.
  const status = created.status === 400 ? 400 : created.status >= 500 ? created.status : 502;
  return { status, jsonBody: { success: false, error: created.error, osTicketStatus: created.status } };
}

app.http("createTicket", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "createTicket",
  handler: createTicketHandler,
});