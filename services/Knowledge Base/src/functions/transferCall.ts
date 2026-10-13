import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { toE164US } from "../lib/phoneNormalize";

interface VapiToolCall {
  id?: string;
  function?: { name?: string; arguments?: Record<string, unknown> | string };
}

/** Tolerates arguments arriving as an object or a JSON string. */
function readArguments(toolCall: VapiToolCall): Record<string, unknown> {
  const args = toolCall.function?.arguments;
  if (!args) return {};

  try {
    const parsed = typeof args === "string" ? JSON.parse(args) : args;
    return parsed && typeof parsed === "object" ? (parsed as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

function readDestination(args: Record<string, unknown>): string | null {
  return typeof args.destination === "string" ? args.destination : null;
}

export async function transferCall(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[TRANSFER-CALL] Rejected — missing or wrong bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }
  const rawBodyText = await request.text();

  let body: any;
  try {
    body = JSON.parse(rawBodyText);
  } catch (error) {
    context.error(`[TRANSFER-CALL] Unreadable body: ${(error as Error).message}`);
    return { status: 200, jsonBody: { error: "Sorry, I couldn't set up that transfer." } };
  }

  const message = body?.message ?? {};

  // The live handle for this call. Without it there's nothing to transfer.
  const controlUrl: string | undefined = message.call?.monitor?.controlUrl;

  // Vapi delivers the tool call under toolWithToolCallList (per the docs),
  // with toolCallList as an older fallback.
  const toolCall: VapiToolCall | undefined =
    message.toolWithToolCallList?.[0]?.toolCall ?? message.toolCallList?.[0];

  const toolCallId = toolCall?.id;

  if (!controlUrl || !toolCall) {
    context.error(
      `[TRANSFER-CALL] Missing controlUrl (${!!controlUrl}) or toolCall (${!!toolCall}).`
    );
    return {
      status: 200,
      jsonBody: {
        results: [
          {
            toolCallId,
            result: "I wasn't able to start the transfer. Let me get someone to help.",
          },
        ],
      },
    };
  }

  // ── Validate the number ──────────────────────────────────────
  const args = readArguments(toolCall);
  const raw = readDestination(args);
  const e164 = raw ? toE164US(raw) : null;

  if (!e164) {
    context.warn(`[TRANSFER-CALL] Refused — invalid destination: "${raw}"`);
    // Refusal is a normal outcome, so respond 200 and let the assistant
    // tell the caller. No POST to controlUrl means no transfer happens.
    return {
      status: 200,
      jsonBody: {
        results: [
          {
            toolCallId,
            result:
              "I'm sorry, I couldn't transfer you — I didn't get a valid number to connect to. Is there something else I can help with?",
          },
        ],
      },
    };
  }

  try {
    const response = await fetch(controlUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "transfer",
        destination: { type: "number", number: e164 },
        content: 'Transferring you now',
      }),
    });

    if (!response.ok) {
      const detail = await response.text();
      context.error(`[TRANSFER-CALL] controlUrl POST failed (${response.status}): ${detail}`);
      return {
        status: 200,
        jsonBody: {
          results: [
            { toolCallId, result: "I'm having trouble connecting the call right now." },
          ],
        },
      };
    }
    
    return {
      status: 200,
      jsonBody: {
        results: [
          {
            toolCallId,
            result:
              "TRANSFER_IN_PROGRESS. Do not say anything further.",
          },
        ],
      },
    };
  } catch (error) {
    context.error(`[TRANSFER-CALL] Transfer error: ${(error as Error).message}`);
    return {
      status: 200,
      jsonBody: {
        results: [{ toolCallId, result: "I'm having trouble connecting the call right now." }],
      },
    };
  }
}

app.http("transferCall", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "transferCall",
  handler: transferCall,
});