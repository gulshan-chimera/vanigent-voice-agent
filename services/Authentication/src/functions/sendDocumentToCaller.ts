// src/functions/sendDocumentToCaller.ts
//
// Standalone webhook endpoint for the 'send_document_to_caller' VAPI tool.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { handleSendDocumentRequest } from "../lib/documentMailer";
import { VapiWebhookBody, VapiToolCallsResponse, VapiToolCallItem } from "../types/vapi";

export async function sendDocumentToCaller(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[SEND-DOC-WEBHOOK] Rejected request — invalid or missing webhook secret.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: VapiWebhookBody;

  try {
    body = (await request.json()) as VapiWebhookBody;
  } catch (error) {
    context.error(`[SEND-DOC-WEBHOOK] Failed to parse request body: ${(error as Error).message}`);
    return { status: 200 };
  }

  const messageType = body?.message?.type;

  if (messageType !== "tool-calls") {
    context.log(`[SEND-DOC-WEBHOOK] Ignoring non tool-calls event: ${messageType}`);
    return { status: 200 };
  }

  const toolCallList = body?.message?.toolCallList ?? [];

  if (toolCallList.length === 0) {
    context.warn("[SEND-DOC-WEBHOOK] tool-calls event with no toolCallList entries.");
    return jsonResponse(200, { results: [] });
  }

  const callerNumber = body?.message?.call?.customer?.number || "";
  const results = [];

  for (const call of toolCallList) {
    const toolName = call.name || call.function?.name;
    const toolArgs = call.arguments || call.function?.arguments || {};
    
    // Some VAPI versions send arguments as a JSON string
    const parsedArgs = typeof toolArgs === "string" ? JSON.parse(toolArgs) : toolArgs;

    if (toolName === "send_document_to_caller") {
      const documentQuery = parsedArgs.documentQuery;
      
      if (!documentQuery || typeof documentQuery !== "string") {
        results.push({
          toolCallId: call.id,
          result: "ERROR. Missing 'documentQuery' parameter.",
        });
        continue;
      }

      const sendResult = await handleSendDocumentRequest(callerNumber, documentQuery);
      results.push({ toolCallId: call.id, result: sendResult });
    } else {
      context.warn(`[SEND-DOC-WEBHOOK] Unknown tool call requested: ${toolName}`);
      results.push({
        toolCallId: call.id,
        result: `ERROR. Tool '${toolName}' is not supported by this server.`,
      });
    }
  }

  const response: VapiToolCallsResponse = { results };

  return jsonResponse(200, response);
}

function jsonResponse(status: number, body: unknown): HttpResponseInit {
  return {
    status,
    jsonBody: body,
    headers: { "Content-Type": "application/json" },
  };
}

app.http("sendDocumentToCaller", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "sendDocumentToCaller",
  handler: sendDocumentToCaller,
});
