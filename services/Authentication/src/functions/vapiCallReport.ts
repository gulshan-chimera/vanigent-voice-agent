// src/functions/vapiCallReport.ts
//
// Receives VAPI's end-of-call-report and emails a follow-up summary when
// the call left questions unanswered.
//
// The structured output lives at:
//   message.analysis.structuredOutputs[<uuid>].result
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

  if (message?.type !== "end-of-call-report") {
    context.log(`[CALL-REPORT] Ignoring event type: ${message?.type ?? "(none)"}`);
    return { status: 200 };
  }

  const report = extractReport(message.analysis?.structuredOutputs);

  if (!report) {
    context.warn(
      "[CALL-REPORT] No follow-up structured output found — check the assistant's artifactPlan."
    );
    return { status: 200 };
  }

  if (!report.followUpRequired) {
    context.log(
      `[CALL-REPORT] followUpRequired is false — no email sent. Open questions: ${report.openQuestions.length}`
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