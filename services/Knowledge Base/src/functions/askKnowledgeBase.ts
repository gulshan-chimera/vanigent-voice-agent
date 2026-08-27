// src/functions/askKnowledgeBase.ts
//
// Simple internal testing endpoint — plain JSON in, plain JSON out.
// Protected by a shared Bearer secret.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { answerQuestion } from "../lib/knowledgeBaseAnswer";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";

export async function askKnowledgeBase(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[ASK-KB] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: { question?: string };

  try {
    body = (await request.json()) as { question?: string };
  } catch (error) {
    return { status: 400, jsonBody: { error: "Invalid request body" } };
  }

  if (!body.question) {
    return { status: 400, jsonBody: { error: "question is required" } };
  }

  const answer = await answerQuestion(body.question);
  context.log(`[ASK-KB] Q: "${body.question}" -> A: "${answer}"`);

  return { status: 200, jsonBody: { answer } };
}

app.http("askKnowledgeBase", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "askKnowledgeBase",
  handler: askKnowledgeBase,
});