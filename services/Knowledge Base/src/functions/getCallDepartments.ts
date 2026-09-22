// src/functions/getCallDepartments.ts
//
// HTTP endpoint that returns the departments queried during a VAPI call.
// Called by the Authentication service at end-of-call to populate
// "Systems Affected" in the follow-up email.
//
// GET /api/getCallDepartments?callId=<vapi-call-id>

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getDepartments } from "../lib/callDepartmentTracker";

export async function getCallDepartments(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const callId = request.query.get("callId");

  if (!callId) {
    return {
      status: 400,
      jsonBody: { error: "Missing callId query parameter" },
    };
  }

  const departments = await getDepartments(callId);

  context.log(
    `[GET-CALL-DEPTS] callId=${callId} -> departments: ${departments.join(", ") || "(none)"}`
  );

  return {
    status: 200,
    jsonBody: { callId, departments },
    headers: { "Content-Type": "application/json" },
  };
}

app.http("getCallDepartments", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "getCallDepartments",
  handler: getCallDepartments,
});
