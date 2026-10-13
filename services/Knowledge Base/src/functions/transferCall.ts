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
function readSupportName(args: Record<string, unknown>): string | null {
  const value = args.support_name;
  if (typeof value !== "string") return null;
  const name = value
    .replace(/[\u0000-\u001f<>{}]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^the\s+/i, "")
    .slice(0, 80)
    .trim();
  return name || null;
}

const DIGIT_WORDS = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"];

/**
 * "+18444481234" → "eight, four, four. four, four, eight. one, two, three, four"
 *
 * Text-to-speech reads a raw E.164 string as one fast run of digits.
 * Spelling each digit as a word with a comma after it, and ending each
 * group (3-3-4) with a period, makes every voice provider pause between
 * digits and between groups — no SSML needed. The +1 country code is
 * dropped, as people say US numbers without it.
 */
function spokenPhoneNumber(e164: string): string {
  const digits = e164.replace(/\D/g, "").replace(/^1(?=\d{10}$)/, "");
  const say = (part: string) => part.split("").map((d) => DIGIT_WORDS[Number(d)]).join(", ");
  if (digits.length !== 10) return say(digits);
  return `${say(digits.slice(0, 3))}. ${say(digits.slice(3, 6))}. ${say(digits.slice(6))}`;
}

function transferMessage(supportName: string | null, e164: string): string {
  const who = supportName ? `the ${supportName}` : "our support team";
  return `Transferring you now to ${who} at ${spokenPhoneNumber(e164)}.`;
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

  // support_name is required on the VAPI side, but a missing label must
  // never block a transfer to a valid number — fall back to generic wording.
  const supportName = readSupportName(args);
  if (!supportName) {
    context.warn("[TRANSFER-CALL] support_name missing — using generic wording.");
  }
  const spokenMessage = transferMessage(supportName, e164);

  // ── Execute the transfer via Live Call Control ───────────────
  try {
    const response = await fetch(controlUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "transfer",
        destination: { type: "number", number: e164 },
        content: spokenMessage,
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

    context.log(`[TRANSFER-CALL] Transferred call to ${supportName ?? "support"} at ${e164}.`);
    return {
      status: 200,
      jsonBody: { results: [{ toolCallId, result: spokenMessage }] },
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