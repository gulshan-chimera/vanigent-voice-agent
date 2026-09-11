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
import { getGraphAccessToken } from "./graphAuth";

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
  supportEmail: string,
  callerDisplayName: string,
  senderDisplayName: string
): string {
  const caller = callerDisplayName;
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

      <p style="margin-top:30px">Best Regards,</p>
      <p><strong>${escapeHtml(senderDisplayName)}</strong></p>
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
  let callerDisplayName: string = report.customerName || "";
  let isAuthenticated = false;

  try {
    const lookup = await lookupCaller(context.callerNumber);
    isAuthenticated = lookup.isAuthenticated;
    if (lookup.isAuthenticated && lookup.user) {
      callerEmail = lookup.user.mail ?? lookup.user.userPrincipalName ?? null;
      // Always use the full display name from Entra ID instead of the
      // AI-extracted customerName, which can be inconsistent.
      callerDisplayName = lookup.user.displayName ?? callerDisplayName;
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

  const subjectSuffix = callerDisplayName ? ` \u2014 ${callerDisplayName}` : "";
  const subject = `Support call follow-up${subjectSuffix}`;

  const toRecipients = [supportEmail];
  if (callerEmail) {
    toRecipients.push(callerEmail);
  }

  // Fetch the display name of the sender mailbox (CALL_REPORT_SENDER_UPN)
  // to use in the "Best Regards" sign-off.
  const senderUpn = process.env.CALL_REPORT_SENDER_UPN || "";
  let senderDisplayName = "Vanigent Support Team";

  if (senderUpn) {
    try {
      const graphToken = await getGraphAccessToken();
      if (graphToken) {
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderUpn)}?$select=displayName`,
          { headers: { Authorization: `Bearer ${graphToken}` } }
        );
        if (res.ok) {
          const data = (await res.json()) as { displayName?: string };
          if (data.displayName) senderDisplayName = data.displayName;
        }
      }
    } catch {
      // Fall back to the default name if the lookup fails.
    }
  }

  const sent = await sendMail({
    to: toRecipients,
    subject: subject,
    htmlBody: buildEmailBody(report, context, callerEmail, supportEmail, callerDisplayName, senderDisplayName),
  });

  return { supportSent: sent, callerSent: sent && !!callerEmail, callerEmail };
}