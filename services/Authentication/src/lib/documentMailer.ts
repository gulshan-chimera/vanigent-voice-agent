// src/lib/documentMailer.ts
//
// Handles the send_document_to_caller tool logic.
// Looks up the caller's email, queries the KB service for the document,
// and sends an email with the SharePoint link to the caller.

import { lookupCaller } from "./callerLookup";
import { sendMail } from "./graphMail";
import { getGraphAccessToken } from "./graphAuth";

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildDocumentEmailBody(
  callerDisplayName: string,
  fileName: string,
  webUrl: string,
  senderDisplayName: string
): string {
  return `
    <div style="font-family:system-ui,-apple-system,sans-serif;max-width:640px">
      <p>Hi ${escapeHtml(callerDisplayName)},</p>
      
      <p>As requested during your support call, here is the document you asked for:</p>
      
      <div style="margin:24px 0;padding:16px;background:#f8fafc;border-left:4px solid #0ea5e9;border-radius:4px">
        <p style="margin:0 0 8px;font-weight:600;color:#0f172a">📄 ${escapeHtml(fileName)}</p>
        <p style="margin:0"><a href="${escapeHtml(webUrl)}" style="color:#0ea5e9;text-decoration:none;font-weight:500">Open in SharePoint &rarr;</a></p>
      </div>

      <p style="margin-top:30px">Best Regards,</p>
      <p><strong>${escapeHtml(senderDisplayName)}</strong></p>
    </div>`;
}

export async function searchKnowledgeBase(query: string) {
  const kbUrl = process.env.KB_SERVICE_URL;
  const secret = process.env.KB_API_SECRET;

  if (!kbUrl) {
    throw new Error("KB_SERVICE_URL is not configured.");
  }

  const response = await fetch(`${kbUrl}/api/queryIndex`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${secret}`,
    },
    body: JSON.stringify({ query }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`KB query failed (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as { matches?: { fileName: string; webUrl: string; score: number }[] };
  
  if (!data.matches || data.matches.length === 0) {
    return null;
  }

  // Use the top match
  return data.matches[0];
}

export async function handleSendDocumentRequest(
  callerNumber: string,
  documentQuery: string
): Promise<string> {
  // 1. Look up the caller in Entra ID
  let callerEmail: string | null = null;
  let callerDisplayName = "Employee";

  try {
    const lookup = await lookupCaller(callerNumber);
    if (!lookup.isAuthenticated) {
      console.warn(`[DOC-MAILER] Caller ${callerNumber} is not authenticated. Cannot send document.`);
      return "UNAUTHORIZED. Could not verify employee identity. Do not tell the caller the document was sent.";
    }

    if (lookup.user) {
      callerEmail = lookup.user.mail ?? lookup.user.userPrincipalName ?? null;
      callerDisplayName = lookup.user.displayName ?? callerDisplayName;
    }
  } catch (error) {
    console.warn(`[DOC-MAILER] Caller lookup failed: ${(error as Error).message}`);
    return "ERROR. Failed to look up employee details in Entra ID.";
  }

  if (!callerEmail) {
    console.warn(`[DOC-MAILER] No mailbox found for ${callerNumber}.`);
    return "ERROR. No email address found for this employee. Tell the caller you cannot send the email.";
  }

  // 2. Search for the document via the KB service
  let document;
  try {
    document = await searchKnowledgeBase(documentQuery);
  } catch (error) {
    console.warn(`[DOC-MAILER] Document search failed: ${(error as Error).message}`);
    return "ERROR. Failed to search the knowledge base. Tell the caller there was a system error.";
  }

  if (!document || !document.webUrl) {
    console.warn(`[DOC-MAILER] No matching document found for query: "${documentQuery}"`);
    return `NOT_FOUND. No document matched the query "${documentQuery}". Tell the caller you couldn't find the requested document.`;
  }

  // 3. Resolve sender display name
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
      // Ignore
    }
  }

  // 4. Send the email
  const subject = `Vanigent Support \u2014 ${document.fileName}`;
  const htmlBody = buildDocumentEmailBody(callerDisplayName, document.fileName, document.webUrl, senderDisplayName);

  const sent = await sendMail({
    to: [callerEmail],
    subject,
    htmlBody,
  });

  if (!sent) {
    return "ERROR. Failed to send the email via Microsoft Graph. Tell the caller you encountered an email error.";
  }

  console.log(`[DOC-MAILER] Successfully sent "${document.fileName}" to ${callerEmail}`);
  return `SUCCESS. The document "${document.fileName}" was successfully emailed to the caller. Tell the caller you have sent it to their email.`;
}
