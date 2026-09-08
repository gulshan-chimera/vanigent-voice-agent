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
  const followUp = report.followUpRequired;
  const caller = report.customerName || context.callerNumber;
  const durationMin = Math.round(context.durationSeconds / 60);
  const questions = report.openQuestions || [];

  const banner = followUp
    ? `<div style="background:#fef3c7;border-left:4px solid #d97706;padding:10px 14px;margin:0 0 18px">
         <strong>Follow-up required</strong>
       </div>`
    : `<div style="background:#ecfdf5;border-left:4px solid #059669;padding:10px 14px;margin:0 0 18px">
         No follow-up required
       </div>`;

  return `
    <div style="font-family:system-ui,-apple-system,sans-serif;max-width:640px">
      ${banner}
      <h2 style="margin:0 0 4px">Call summary</h2>
      <p style="color:#555;margin:0 0 16px">
        ${escapeHtml(caller)}
        ${callerEmail ? ` &middot; ${escapeHtml(callerEmail)}` : ""}
        ${durationMin !== null ? ` &middot; ${durationMin} min` : ""}
      </p>

      <p style="line-height:1.5">${escapeHtml(report.summary)}</p>

      ${
        followUp && questions.length
          ? `<h3 style="margin:20px 0 6px">Open questions</h3>
             <ul style="line-height:1.6">
               ${questions.map((q) => `<li>${escapeHtml(q)}</li>`).join("")}
             </ul>`
          : ""
      }

      <p style="margin-top:24px;font-size:13px;color:#777">
        Call ID: ${escapeHtml(context.callId || "n/a")}
      </p>
    </div>`;
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