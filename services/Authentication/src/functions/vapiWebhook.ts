// src/functions/vapiWebhook.ts
//
// Phase 1 (tool-call variant): The Master Agent is built and assigned
// directly in the VAPI dashboard. It answers every call, then immediately
// calls our "authenticate_caller" tool. We look up the caller's number in
// Entra ID and return a plain-text result the assistant uses to decide
// whether to proceed or reject the caller (and hang up via its own
// endCall tool, configured in the VAPI dashboard).
//
// Every incoming request is first verified against a shared secret
// (Authorization: Bearer <secret>) configured as a Custom Credential on
// the VAPI side, to ensure requests genuinely originate from VAPI.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { lookupCaller } from "../lib/callerLookup";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { VapiWebhookBody, VapiToolCallsResponse, VapiToolCallItem, EntraUser } from "../types/vapi";

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

  const messageType = body?.message?.type;

  if (messageType !== "tool-calls") {
    context.log(`[VAPI-WEBHOOK] Ignoring non tool-calls event: ${messageType}`);
    return { status: 200 };
  }

  const toolCallList = body?.message?.toolCallList ?? [];

  if (toolCallList.length === 0) {
    context.warn("[VAPI-WEBHOOK] tool-calls event with no toolCallList entries.");
    return jsonResponse(200, { results: [] });
  }

  const callerNumber = body?.message?.call?.customer?.number;

  const result = callerNumber
    ? await lookupCaller(callerNumber)
    : { isAuthenticated: false, user: null };

  if (!callerNumber) {
    context.warn("[VAPI-WEBHOOK] No caller number in payload — failing closed.");
  } else if (!result.isAuthenticated) {
    context.log(`[VAPI-WEBHOOK] REJECTED — no employee match for ${callerNumber}`);
  } else {
    context.log(
      `[VAPI-WEBHOOK] AUTHORIZED — ${callerNumber} matched employee: ${result.user?.displayName}`
    );
  }

const resultText = result.isAuthenticated && result.user
  ? `AUTHORIZED.\n${formatUserDetails(result.user)}`
  : "UNAUTHORIZED. This caller is not a recognized employee.";


  const response: VapiToolCallsResponse = {
    results: toolCallList.map((call: VapiToolCallItem) => ({
      toolCallId: call.id,
      result: resultText,
    })),
  };

  return jsonResponse(200, response);
}
function formatUserDetails(user: EntraUser): string {
  const lines: string[] = [`Caller name: ${user.displayName}`];

  if (user.jobTitle) lines.push(`Job title: ${user.jobTitle}`);
  if (user.department) lines.push(`Department: ${user.department}`);
  if (user.companyName) lines.push(`Company: ${user.companyName}`);
  if (user.officeLocation) lines.push(`Office location: ${user.officeLocation}`);
  if (user.employeeId) lines.push(`Employee ID: ${user.employeeId}`);
  if (user.employeeType) lines.push(`Employee type: ${user.employeeType}`);
  if (user.employeeHireDate) lines.push(`Hire date: ${user.employeeHireDate}`);
  if (user.userPrincipalName) lines.push(`User principal name: ${user.userPrincipalName}`);
  if (user.mail) lines.push(`Email: ${user.mail}`);
  if (user.otherMails?.length) lines.push(`Other emails: ${user.otherMails.join(", ")}`);
  if (user.mobilePhone) lines.push(`Mobile phone: ${user.mobilePhone}`);
  if (user.businessPhones?.length) lines.push(`Business phone: ${user.businessPhones.join(", ")}`);
  if (user.preferredLanguage) lines.push(`Preferred language: ${user.preferredLanguage}`);
  if (typeof user.accountEnabled === "boolean") lines.push(`Account enabled: ${user.accountEnabled}`);

  return lines.join("\n");
}

function jsonResponse(status: number, body: unknown): HttpResponseInit {
  return {
    status,
    jsonBody: body,
    headers: { "Content-Type": "application/json" },
  };
}

app.http("vapiWebhook", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "vapiWebhook",
  handler: vapiWebhook,
});
