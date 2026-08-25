// src/functions/vapiWebhook.ts
//
// VAPI-facing endpoint for the "search_hr_knowledge_base" tool. Receives
// VAPI's tool-calls webhook format, extracts the caller's question from
// the tool arguments, and returns the generated answer in VAPI's
// expected results format.
//
// Error handling principle: a real caller is on the line, so once the
// request is authenticated, EVERY failure path must still return 200
// with a results array. An unanswered toolCallId leaves the assistant
// waiting with nothing to say and the caller in silence.
//
// The one exception is auth: a failed secret check returns 401, because
// a request that isn't genuinely from VAPI has no caller behind it, and
// a friendly reply would only confirm to a prober that the endpoint is
// live. Note this reads KB_API_SECRET — the Bearer Token credential
// configured on the VAPI tool must match that, NOT the Authentication
// service's separate VAPI_WEBHOOK_SECRET.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { answerQuestion } from "../lib/knowledgeBaseAnswer";

interface VapiToolCallItem {
  id: string;
  name?: string;
  arguments?: Record<string, unknown>;
  // Some VAPI versions nest name/arguments under a "function" object
  // instead of putting them at the top level — handle both shapes.
  function?: {
    name?: string;
    arguments?: Record<string, unknown> | string;
  };
}

interface VapiWebhookBody {
  message: {
    type: string;
    toolCallList?: VapiToolCallItem[];
  };
}

const NO_QUESTION_REPLY = "I didn't catch a specific question to look up.";
const LOOKUP_ERROR_REPLY = "I'm having trouble looking that up right now.";

/**
 * Pulls the question out of a tool call, tolerating both payload shapes
 * and malformed argument JSON. Returns null rather than throwing — a
 * throw here would abort the whole response and stall the call.
 */
function extractQuestion(call: VapiToolCallItem): string | null {
  const args = call.arguments ?? call.function?.arguments;

  if (!args) return null;

  try {
    const parsed = typeof args === "string" ? JSON.parse(args) : args;
    return typeof parsed?.question === "string" ? parsed.question : null;
  } catch {
    return null;
  }
}

export async function vapiWebhook(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[VAPI-WEBHOOK] Rejected request — invalid or missing webhook secret.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: VapiWebhookBody;

  try {
    body = (await request.json()) as VapiWebhookBody;
  } catch (error) {
    context.error(`[VAPI-WEBHOOK] Failed to parse request body: ${(error as Error).message}`);
    // 200 with no body: we can't know which toolCallIds to answer, and a
    // non-200 would make VAPI treat this as a server fault and retry.
    return { status: 200 };
  }

  if (body?.message?.type !== "tool-calls") {
    context.log(`[VAPI-WEBHOOK] Ignoring non tool-calls event: ${body?.message?.type}`);
    return { status: 200 };
  }

  const toolCallList = body.message.toolCallList ?? [];

  if (toolCallList.length === 0) {
    context.warn("[VAPI-WEBHOOK] tool-calls event with no toolCallList entries.");
    return { status: 200, jsonBody: { results: [] } };
  }

  // allSettled, not all: if one lookup rejects, the others must still be
  // answered. With Promise.all a single rejection discards every result.
  const settled = await Promise.allSettled(
    toolCallList.map(async (call) => {
      const question = extractQuestion(call);

      if (!question) {
        context.warn(`[VAPI-WEBHOOK] Tool call ${call.id} missing 'question' argument.`);
        return { toolCallId: call.id, result: NO_QUESTION_REPLY };
      }

      const answer = await answerQuestion(question);
      context.log(`[VAPI-WEBHOOK] Q: "${question}" -> A: "${answer}"`);

      return { toolCallId: call.id, result: answer };
    })
  );

  const results = settled.map((outcome, i) => {
    if (outcome.status === "fulfilled") {
      return outcome.value;
    }

    context.error(
      `[VAPI-WEBHOOK] Answer generation threw for tool call ${toolCallList[i].id}: ${outcome.reason}`
    );
    return { toolCallId: toolCallList[i].id, result: LOOKUP_ERROR_REPLY };
  });

  return {
    status: 200,
    jsonBody: { results },
    headers: { "Content-Type": "application/json" },
  };
}

app.http("vapiWebhook", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "vapiWebhook",
  handler: vapiWebhook,
});