// src/functions/vapiCallReport.ts
//
// Receives VAPI's end-of-call-report and emails a follow-up summary when
// the call left questions unanswered.
//
// The structured output lives at:
//   message.artifact.structuredOutputs[<uuid>].result
// We also check message.analysis.structuredOutputs as a legacy fallback.
// We iterate the values rather than hardcoding the UUID, since that ID
// changes if the structured output is recreated in the VAPI dashboard.
//
// Always returns 200 (except on auth failure): VAPI retries non-2xx
// responses, and a retry would send duplicate emails.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { timingSafeEqual } from "crypto";
import { sendFollowUpReport, FollowUpReport, CallReportContext } from "../lib/callReportMail";

interface StructuredOutputEntry {
  name?: string;
  result?: Partial<FollowUpReport>;
}

interface EndOfCallReportBody {
  message?: {
    type?: string;
    timestamp?: number;
    startedAt?: string;
    durationSeconds?: number;
    transcript?: string;
    artifact?: {
      structuredOutputs?: Record<string, StructuredOutputEntry>;
    };
    analysis?: {
      summary?: string;
      structuredOutputs?: Record<string, StructuredOutputEntry>;
    };
    call?: {
      id?: string;
      customer?: { number?: string };
    };
    customer?: { number?: string };
  };
}

/** Separate secret from the tool webhook — different caller, different credential. */
function isAuthorized(request: HttpRequest): boolean {
  const expected = process.env.VAPI_REPORT_SECRET;

  if (!expected) {
    console.error("[CALL-REPORT] VAPI_REPORT_SECRET is not set — rejecting.");
    return false;
  }

  const header = request.headers.get("authorization");
  if (!header?.startsWith("Bearer ")) return false;

  const provided = Buffer.from(header.slice("Bearer ".length));
  const expectedBuffer = Buffer.from(expected);

  if (provided.length !== expectedBuffer.length) return false;
  return timingSafeEqual(provided, expectedBuffer);
}

/**
 * Pulls the first structured output that looks like a follow-up report.
 * Iterating avoids depending on a UUID that changes if the output is
 * recreated in the dashboard.
 */
function extractReport(
  outputs: Record<string, StructuredOutputEntry> | undefined
): FollowUpReport | null {
  if (!outputs) return null;

  for (const entry of Object.values(outputs)) {
    const result = entry?.result;
    if (!result) continue;

    // followUpRequired is the field that makes this the report we want.
    if (typeof result.followUpRequired === "boolean") {
      return {
        summary: result.summary ?? "",
        urgency: result.urgency ?? "unknown",
        customerName: result.customerName ?? "",
        openQuestions: Array.isArray(result.openQuestions) ? result.openQuestions : [],
        followUpRequired: result.followUpRequired,
      };
    }
  }

  return null;
}

export async function vapiCallReport(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isAuthorized(request)) {
    context.warn("[CALL-REPORT] Rejected — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  let body: EndOfCallReportBody;

  try {
    body = (await request.json()) as EndOfCallReportBody;
  } catch (error) {
    context.error(`[CALL-REPORT] Unparseable body: ${(error as Error).message}`);
    return { status: 200 };
  }

  const message = body?.message;

  if (message?.type !== "end-of-call-report" && message?.type !== "call.analysis.completed") {
    context.log(`[CALL-REPORT] Ignoring event type: ${message?.type ?? "(none)"}`);
    return { status: 200 };
  }

  context.log(`[CALL-REPORT] Processing event: ${message.type}`);

  // VAPI puts structured outputs under artifact (current) or analysis (legacy)
  const structuredOutputs =
    message.artifact?.structuredOutputs ?? message.analysis?.structuredOutputs;

  context.log(
    `[CALL-REPORT] structuredOutputs keys: ${structuredOutputs ? Object.keys(structuredOutputs).join(", ") : "(none)"}`
  );

  const report = extractReport(structuredOutputs);

  if (!report) {
    context.warn(
      "[CALL-REPORT] No follow-up structured output found — check the assistant's artifactPlan."
    );
    return { status: 200 };
  }



  const callerNumber =
    message.call?.customer?.number ?? message.customer?.number ?? "";

  if (!callerNumber) {
    context.warn("[CALL-REPORT] No caller number in report — cannot resolve caller email.");
  }

  const reportContext: CallReportContext = {
    callerNumber,
    callId: message.call?.id ?? "(unknown)",
    startedAt: message.startedAt ?? new Date(message.timestamp ?? Date.now()).toISOString(),
    durationSeconds: message.durationSeconds ?? 0,
    transcript: message.transcript ?? "(no transcript)",
  };

  // Fetch departments from the KB service using the call ID.
  // This derives "Systems Affected" from the actual libraries queried
  // during the call, rather than relying on VAPI's structured output.
  const kbServiceUrl = process.env.KB_SERVICE_URL;
  if (kbServiceUrl && reportContext.callId && reportContext.callId !== "(unknown)") {
    try {
      const deptUrl = `${kbServiceUrl}/api/getCallDepartments?callId=${encodeURIComponent(reportContext.callId)}`;
      const deptRes = await fetch(deptUrl);
      if (deptRes.ok) {
        const deptData = (await deptRes.json()) as { departments?: string[] };
        if (deptData.departments && deptData.departments.length > 0) {
          report.systemsAffected = deptData.departments;
          console.log(
            `[CALL-REPORT] Systems affected for call ${reportContext.callId}: ${deptData.departments.join(", ")}`
          );
        }
      }
    } catch (error) {
      console.warn(
        `[CALL-REPORT] Failed to fetch departments from KB service: ${(error as Error).message}`
      );
    }
  }

  const result = await sendFollowUpReport(report, reportContext);

  context.log(
    `[CALL-REPORT] Follow-up for call ${reportContext.callId}: support=${result.supportSent}, caller=${result.callerSent}, callerEmail=${result.callerEmail ?? "none"}`
  );

  return { status: 200, jsonBody: { supportSent: result.supportSent, callerSent: result.callerSent } };
}

app.http("vapiCallReport", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "vapiCallReport",
  handler: vapiCallReport,
});