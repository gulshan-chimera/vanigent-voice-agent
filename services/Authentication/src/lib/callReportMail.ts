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

function buildEmailBody(
  report: FollowUpReport,
  context: CallReportContext,
  callerEmail: string | null,
  supportEmail: string
): string {
  const caller = report.customerName || context.callerNumber;
  const durationMin = Math.round(context.durationSeconds / 60);
  const questions = report.openQuestions || [];
  const hasQuestions = questions.length > 0;

  const introText = hasQuestions
    ? `Thanks for calling the Vanigent Support Team. Please find below the call summary and the follow-up questions identified during the call.`
    : `Thanks for calling the Vanigent Support Team. Please find below the call summary.`;

  const supportTag = hasQuestions
    ? `<p><a href="mailto:${escapeHtml(supportEmail)}">@support_team</a>, please review the open questions below and try to resolve them.</p>`
    : "";

  const urgencyLabel = report.urgency
    ? report.urgency.charAt(0).toUpperCase() + report.urgency.slice(1)
    : "Unknown";

  const questionsHeader = questions.length > 1 
    ? `Open questions, Priority: ${escapeHtml(urgencyLabel)}` 
    : `Open question, Priority: ${escapeHtml(urgencyLabel)}`;

  const questionsSection = hasQuestions
    ? `<h3 style="margin:20px 0 6px">${questionsHeader}</h3>
       <ul style="line-height:1.6">
         ${questions.map((q) => `<li>${escapeHtml(q)}</li>`).join("")}
       </ul>`
    : "";

  return `
    <div style="font-family:system-ui,-apple-system,sans-serif;max-width:640px">
      <p>Hi ${escapeHtml(caller)},</p>
      <p>${introText}</p>
      ${supportTag}
      
      <h2 style="margin:0 0 4px">Call summary</h2>
      <p style="color:#555;margin:0 0 16px">
        ${escapeHtml(caller)}
        ${callerEmail ? ` &middot; ${escapeHtml(callerEmail)}` : ""}
        ${durationMin !== null ? ` &middot; ${durationMin} min` : ""}
        &middot; Call ID: ${escapeHtml(context.callId || "n/a")}
      </p>

      <p style="line-height:1.5">${escapeHtml(report.summary)}</p>

      ${questionsSection}
    </div>`;
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
  let isAuthenticated = false;

  try {
    const lookup = await lookupCaller(context.callerNumber);
    isAuthenticated = lookup.isAuthenticated;
    if (lookup.isAuthenticated && lookup.user) {
      callerEmail = lookup.user.mail ?? lookup.user.userPrincipalName ?? null;
    }
  } catch (error) {
    console.warn(
      `[CALL-REPORT-MAIL] Caller lookup failed: ${(error as Error).message}`
    );
  }

  if (!isAuthenticated) {
    console.log(
      `[CALL-REPORT-MAIL] Caller ${context.callerNumber} is not authenticated in Entra ID. Skipping follow-up email.`
    );
    return { supportSent: false, callerSent: false, callerEmail: null };
  }

  if (!callerEmail) {
    console.warn(
      `[CALL-REPORT-MAIL] No mailbox found for authenticated user ${context.callerNumber} — support only.`
    );
  }

  const subjectSuffix = report.customerName ? ` — ${report.customerName}` : "";
  const subject = `Support call follow-up${subjectSuffix}`;

  const toRecipients = [supportEmail];
  if (callerEmail) {
    toRecipients.push(callerEmail);
  }

  const sent = await sendMail({
    to: toRecipients,
    subject: subject,
    htmlBody: buildEmailBody(report, context, callerEmail, supportEmail),
  });

  return { supportSent: sent, callerSent: sent && !!callerEmail, callerEmail };
}