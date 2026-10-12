// src/types/osTicket.ts

/** What callers of this service send us. */
export interface CreateTicketInput {
  name: string;
  email: string;
  subject: string;
  /** Plain text. A ready-made data URL ("data:text/html,...") is passed through as-is. */
  message: string;
  phone?: string;
  /** osTicket help topic ID, if tickets should be routed to a topic. */
  topicId?: number;
  /** Send staff alerts for this ticket (default true). */
  alert?: boolean;
  /** Send the auto-response email to the user (default true). */
  autorespond?: boolean;
}

/** The JSON body osTicket's /api/tickets.json expects. */
export interface OsTicketPayload {
  alert: boolean;
  autorespond: boolean;
  source: "API";
  name: string;
  email: string;
  phone?: string;
  subject: string;
  message: string;
  topicId?: number;
}

export type CreateTicketResult =
  | { ok: true; ticketNumber: string }
  | { ok: false; status: number; error: string };