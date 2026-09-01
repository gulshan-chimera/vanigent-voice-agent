// src/lib/graphMail.ts
//
// Sends mail via Microsoft Graph using application permissions
// (Mail.Send). With app-only auth there is no signed-in user, so the
// sending mailbox must be named explicitly — hence CALL_REPORT_SENDER_UPN.
// That mailbox must actually exist and be licensed; a UPN without a
// mailbox will fail here.

import { getGraphAccessToken } from "./graphAuth";

export interface MailMessage {
  to: string[];
  subject: string;
  htmlBody: string;
}

export async function sendMail(message: MailMessage): Promise<boolean> {
  const senderUpn = process.env.CALL_REPORT_SENDER_UPN;

  if (!senderUpn) {
    console.error("[GRAPH-MAIL] CALL_REPORT_SENDER_UPN is not set.");
    return false;
  }

  const recipients = message.to.filter((address) => Boolean(address?.trim()));

  if (recipients.length === 0) {
    console.error("[GRAPH-MAIL] No valid recipients — nothing sent.");
    return false;
  }

  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[GRAPH-MAIL] Could not acquire Graph token.");
    return false;
  }

  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(senderUpn)}/sendMail`;

  const body = {
    message: {
      subject: message.subject,
      body: { contentType: "HTML", content: message.htmlBody },
      toRecipients: recipients.map((address) => ({
        emailAddress: { address },
      })),
    },
    saveToSentItems: true,
  };

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    // sendMail returns 202 Accepted with an empty body on success.
    if (response.status === 202) {
      console.log(`[GRAPH-MAIL] Sent "${message.subject}" to ${recipients.join(", ")}`);
      return true;
    }

    const errorText = await response.text();
    console.error(`[GRAPH-MAIL] Send failed (${response.status}): ${errorText}`);
    return false;
  } catch (error) {
    console.error(`[GRAPH-MAIL] Network error: ${(error as Error).message}`);
    return false;
  }
}