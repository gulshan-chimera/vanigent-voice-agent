// src/functions/vapiWebhook.ts
//
// VAPI-facing endpoint for the "search_hr_knowledge_base" tool. Receives
// VAPI's tool-calls webhook format, extracts the caller's question from
// the tool arguments, and returns the generated answer in VAPI's
// expected results format.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { answerQuestion } from "../lib/knowledgeBaseAnswer";

interface VapiToolCallItem {
  id: string;
  name?: string;
  arguments?: Record<string, unknown>;
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

function extractQuestion(call: VapiToolCallItem): string | null {
  const args = call.arguments ?? call.function?.arguments;

  if (!args) return null;

  const parsed = typeof args === "string" ? JSON.parse(args) : args;
  return typeof parsed.question === "string" ? parsed.question : null;
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
    return { status: 200 };
  }

  if (body?.message?.type !== "tool-calls") {
    context.log(`[VAPI-WEBHOOK] Ignoring non tool-calls event: ${body?.message?.type}`);
    return { status: 200 };
  }

  const toolCallList = body.message.toolCallList ?? [];

  const results = await Promise.all(
    toolCallList.map(async (call) => {
      const question = extractQuestion(call);

      if (!question) {
        context.warn(`[VAPI-WEBHOOK] Tool call ${call.id} missing 'question' argument.`);
        return {
          toolCallId: call.id,
          result: "I didn't catch a specific question to look up.",
        };
      }

      const answer = await answerQuestion(question);
      context.log(`[VAPI-WEBHOOK] Q: "${question}" -> A: "${answer}"`);

      return { toolCallId: call.id, result: answer };
    })
  );

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