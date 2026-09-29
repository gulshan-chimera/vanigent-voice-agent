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
import { sendMail, MailAttachment } from "./graphMail";
import { getGraphAccessToken } from "./graphAuth";

export interface FollowUpReport {
  summary: string;
  urgency: string;
  customerName: string;
  openQuestions: string[];
  followUpRequired: boolean;
  systemsAffected?: string[];
  transcriptRequired?: boolean;
  requestedDocumentQueries?: string[];
  requestedDocuments?: { name: string; url: string }[];
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
    ? `<p><a href="mailto:${escapeHtml(supportEmail)}">@support_team</a>, please review the follow-up questions below and try to resolve them.</p>`
    : "";

  const urgencyLabel = report.urgency
    ? report.urgency.charAt(0).toUpperCase() + report.urgency.slice(1)
    : "Unknown";

  const questionsListHtml = hasQuestions
    ? `<ul style="margin:0; padding-left:16px; line-height:1.5;">
         ${questions.map((q) => `<li>${escapeHtml(q)}</li>`).join("")}
       </ul>`
    : "<em>No follow-up needed as all questions were resolved during the call.</em>";

  const callDateObj = context.startedAt ? new Date(context.startedAt) : new Date();
  const callDate = callDateObj.toLocaleDateString('en-US', { timeZone: 'America/New_York' });
  const rawTime = callDateObj.toLocaleTimeString('en-US', { 
    timeZone: 'America/New_York', 
    hour: '2-digit', 
    minute: '2-digit'
  });
  const callTime = `${rawTime} EST`;

  const transcriptNotice = report.transcriptRequired 
    ? `<p><em>Note: The requested call transcript is attached to this email.</em></p>`
    : "";

  return `
    <div style="font-family:system-ui,-apple-system,sans-serif;max-width:640px">
      <p>Hi ${escapeHtml(caller)},</p>
      <p>${introText}</p>
      ${supportTag}
      ${transcriptNotice}
      
      ${report.requestedDocuments && report.requestedDocuments.length > 0 ? `
      <p style="margin:24px 0 8px;color:#334155">As requested during your call, please find below the document${report.requestedDocuments.length > 1 ? 's' : ''} you asked for. Click the link${report.requestedDocuments.length > 1 ? 's' : ''} to view ${report.requestedDocuments.length > 1 ? 'them' : 'it'} directly in SharePoint.</p>
      <h3 style="margin:16px 0 12px">Requested Documents</h3>
      ${report.requestedDocuments.map(doc => `
      <div style="margin:12px 0;padding:16px;background:#f8fafc;border-left:4px solid #0ea5e9;border-radius:4px">
        <p style="margin:0 0 8px;font-weight:600;color:#0f172a">📄 ${escapeHtml(doc.name)}</p>
        <p style="margin:0"><a href="${escapeHtml(doc.url)}" style="color:#0ea5e9;text-decoration:none;font-weight:500">Open in SharePoint &rarr;</a></p>
      </div>`).join('')}` : ''}

      <h2 style="margin:24px 0 12px">Call summary</h2>
      
      <table style="width:100%; border-collapse:collapse; margin-bottom:24px; font-size:14px; text-align:left; border:1px solid #e2e8f0;">
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; width:30%; font-weight:600;">Date</th>
          <td style="padding:12px; color:#0f172a;">${escapeHtml(callDate)}</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; font-weight:600;">Time Called</th>
          <td style="padding:12px; color:#0f172a;">${escapeHtml(callTime)}</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; font-weight:600;">Call ID</th>
          <td style="padding:12px; color:#0f172a;">${escapeHtml(context.callId || "Unknown")}</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; font-weight:600;">Urgency/Importance</th>
          <td style="padding:12px; color:#0f172a; font-weight:600;">${escapeHtml(urgencyLabel)}</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; font-weight:600;">Systems Affected</th>
          <td style="padding:12px; color:#0f172a;">${report.systemsAffected && report.systemsAffected.length > 0 ? escapeHtml(report.systemsAffected.join(", ")) : "<em>N/A</em>"}</td>
        </tr>
        <tr style="border-bottom:1px solid #e2e8f0;">
          <th style="padding:12px; background:#f8fafc; color:#475569; vertical-align:top; font-weight:600;">Summary</th>
          <td style="padding:12px; color:#0f172a; line-height:1.5;">${escapeHtml(report.summary)}</td>
        </tr>
        <tr>
          <th style="padding:12px; background:#f8fafc; color:#475569; vertical-align:top; font-weight:600;">Follow-up questions</th>
          <td style="padding:12px; color:#0f172a;">${questionsListHtml}</td>
        </tr>
      </table>

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

  const attachments: MailAttachment[] = [];
  if (report.transcriptRequired && context.transcript) {
    attachments.push({
      name: "transcript.txt",
      contentType: "text/plain",
      contentBytes: Buffer.from(context.transcript).toString("base64"),
    });
  }

  const sent = await sendMail({
    to: toRecipients,
    subject: subject,
    htmlBody: buildEmailBody(report, context, callerEmail, supportEmail, callerDisplayName, senderDisplayName),
    attachments: attachments.length > 0 ? attachments : undefined,
  });

  return { supportSent: sent, callerSent: sent && !!callerEmail, callerEmail };
}