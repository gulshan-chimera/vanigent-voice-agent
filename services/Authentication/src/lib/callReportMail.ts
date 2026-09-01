// src/lib/callReportMail.ts
//
// Builds and sends the post-call follow-up report.
//
// Two recipients with different guarantees: support ALWAYS receives it,
// because that's what makes the follow-up actually happen. The caller
// receives a copy only if we can resolve a real mailbox — and if we
// can't, support is told so explicitly rather than the failure being
// silent.

import { lookupCaller } from "./callerLookup";
import { sendMail } from "./graphMail";

export interface FollowUpReport {
  summary: string;
  urgency: string;
  customerName: string;
  openQuestions: string[];
  followUpRequired: boolean;
}

export interface CallReportContext {
  callerNumber: string;
  callId: string;
  startedAt: string;
  durationSeconds: number;
  transcript: string;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildQuestionList(questions: string[]): string {
  if (questions.length === 0) {
    return "<p><em>No specific open questions were captured.</em></p>";
  }

  const items = questions
    .map((question) => `<li>${escapeHtml(question)}</li>`)
    .join("\n");

  return `<ul>\n${items}\n</ul>`;
}

function buildSupportBody(
  report: FollowUpReport,
  context: CallReportContext,
  callerEmail: string | null
): string {
  const emailLine = callerEmail
    ? `<p><strong>Caller email:</strong> ${escapeHtml(callerEmail)} (copy sent)</p>`
    : `<p style="color:#b00;"><strong>Caller email could not be resolved — no copy was sent to the caller.</strong></p>`;

  return `
<p><strong>Follow-up required from a support call.</strong></p>
<p><strong>Caller:</strong> ${escapeHtml(report.customerName || "Unknown")}<br>
<strong>Number:</strong> ${escapeHtml(context.callerNumber)}<br>
<strong>Urgency:</strong> ${escapeHtml(report.urgency)}<br>
<strong>Call time:</strong> ${escapeHtml(context.startedAt)}<br>
<strong>Duration:</strong> ${Math.round(context.durationSeconds)} seconds</p>
${emailLine}
<h3>Questions needing follow-up</h3>
${buildQuestionList(report.openQuestions)}
<h3>Call summary</h3>
<p>${escapeHtml(report.summary)}</p>
<h3>Transcript</h3>
<pre style="white-space:pre-wrap;font-family:inherit;">${escapeHtml(context.transcript)}</pre>
<p style="color:#666;font-size:12px;">Call ID: ${escapeHtml(context.callId)}</p>
`;
}

function buildCallerBody(report: FollowUpReport, context: CallReportContext): string {
  const firstName = (report.customerName || "").split(" ")[0] || "there";

  return `
<p>Hi ${escapeHtml(firstName)},</p>
<p>Thanks for calling employee support. There were a few things we couldn't
answer on the call, so we've passed them to the team to follow up with you.</p>
<h3>What we're following up on</h3>
${buildQuestionList(report.openQuestions)}
<h3>Summary of your call</h3>
<p>${escapeHtml(report.summary)}</p>
<p>Someone will be in touch. No action needed from you in the meantime.</p>
<p style="color:#666;font-size:12px;">Call reference: ${escapeHtml(context.callId)}</p>
`;
}

export async function sendFollowUpReport(
  report: FollowUpReport,
  context: CallReportContext
): Promise<{ supportSent: boolean; callerSent: boolean; callerEmail: string | null }> {
  const supportEmail = process.env.KB_SUPPORT_EMAIL;

  if (!supportEmail) {
    console.error("[CALL-REPORT-MAIL] KB_SUPPORT_EMAIL is not set — cannot send.");
    return { supportSent: false, callerSent: false, callerEmail: null };
  }

  // Resolve the caller's mailbox from their phone number. Prefer `mail`
  // (a real mailbox) over userPrincipalName (a sign-in identifier that
  // may not route mail).
  let callerEmail: string | null = null;

  try {
    const lookup = await lookupCaller(context.callerNumber);
    if (lookup.isAuthenticated && lookup.user) {
      callerEmail = lookup.user.mail ?? lookup.user.userPrincipalName ?? null;
    }
  } catch (error) {
    console.warn(
      `[CALL-REPORT-MAIL] Caller lookup failed: ${(error as Error).message}`
    );
  }

  if (!callerEmail) {
    console.warn(
      `[CALL-REPORT-MAIL] No mailbox found for ${context.callerNumber} — support only.`
    );
  }

  const subjectSuffix = report.customerName ? ` — ${report.customerName}` : "";

  const supportSent = await sendMail({
    to: [supportEmail],
    subject: `Support call follow-up (${report.urgency})${subjectSuffix}`,
    htmlBody: buildSupportBody(report, context, callerEmail),
  });

  let callerSent = false;
  if (callerEmail) {
    callerSent = await sendMail({
      to: [callerEmail],
      subject: "Following up on your support call",
      htmlBody: buildCallerBody(report, context),
    });
  }

  return { supportSent, callerSent, callerEmail };
}