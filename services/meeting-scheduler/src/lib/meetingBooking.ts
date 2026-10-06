import { invitationBody } from "./meetingInvitation";
import { createHash, createHmac, timingSafeEqual } from "node:crypto";
import { MeetingDetails } from "../types/meetingDetails";
import { checkCalendarAvailability } from "./calendarAvailability";
import { getGraphAccessToken } from "./graphAuth";

export class BookingError extends Error {
  constructor(public readonly code: "CONFIRMATION_REQUIRED" | "BOOKING_FAILED" | "BOOKING_UNKNOWN", message: string) {
    super(message);
    this.name = "BookingError";
  }
}

function bookingKey(callId: string, meeting: MeetingDetails, startDateTime: string): string {
  return createHash("sha256").update(JSON.stringify([
    process.env.TENANT_ID ?? "", callId, meeting.caller.id, meeting.caller.email.toLowerCase(),
    meeting.title, new Date(startDateTime).toISOString(), meeting.durationMinutes,
    meeting.attendees.map(user => `${user.id}:${user.email.toLowerCase()}`).sort(),
    ...(meeting.description ? [meeting.description] : []),
    ...(meeting.location ? [{ location: meeting.location }] : []),
  ])).digest("hex");
}

function signature(payload: string): string {
  const secret = process.env.MEETING_SCHEDULER_SECRET;
  if (!secret) throw new BookingError("CONFIRMATION_REQUIRED", "Scheduler authentication is not configured.");
  return createHmac("sha256", secret).update(`meeting-confirmation:${payload}`).digest("base64url");
}

export function issueConfirmation(callId: string, meeting: MeetingDetails, startDateTime: string): string {
  const payload = Buffer.from(JSON.stringify({ key: bookingKey(callId, meeting, startDateTime), expires: Date.now() + 15 * 60_000 })).toString("base64url");
  return `${payload}.${signature(payload)}`;
}

// Tool 1 checks the caller and attendees. The token binds that organizer and the exact checked meeting.
function availabilityKey(meeting: Pick<MeetingDetails, "caller" | "title" | "location" | "description" | "attendees" | "durationMinutes">, start: string): string {
  return createHash("sha256").update(JSON.stringify([
    "participant-availability-v2", process.env.TENANT_ID ?? "", meeting.caller.id, meeting.caller.email.toLowerCase(), meeting.title,
    new Date(start).toISOString(), meeting.durationMinutes,
    meeting.attendees.map(user => user.id + ":" + user.email.toLowerCase()).sort(),
    ...(meeting.description ? [meeting.description] : []),
    ...(meeting.location ? [{ location: meeting.location }] : []),
  ])).digest("hex");
}
export function issueAvailabilityConfirmation(meeting: Pick<MeetingDetails, "caller" | "title" | "location" | "description" | "attendees" | "durationMinutes">, start: string): string {
  const payload = Buffer.from(JSON.stringify({ purpose: "participant-availability-v2", key: availabilityKey(meeting, start), expires: Date.now() + 15 * 60_000 })).toString("base64url");
  return payload + "." + signature(payload);
}

function verifyConfirmation(callId: string, meeting: MeetingDetails, startDateTime: string, token: unknown): string {
  const failure = () => new BookingError("CONFIRMATION_REQUIRED", "Recheck this exact meeting with Tool 1, read back the details, and obtain confirmation before booking.");
  if (typeof token !== "string" || token.length > 2048) throw failure();
  const pieces = token.split(".");
  if (pieces.length !== 2) throw failure();
  const expected = Buffer.from(signature(pieces[0]));
  const actual = Buffer.from(pieces[1]);
  if (expected.length !== actual.length || !timingSafeEqual(expected, actual)) throw failure();
  let payload: { key?: unknown; expires?: unknown; purpose?: unknown };
  try { payload = JSON.parse(Buffer.from(pieces[0], "base64url").toString("utf8")); } catch { throw failure(); }
  const key = bookingKey(callId, meeting, startDateTime);
  const expectedKey = payload?.purpose === "participant-availability-v2" ? availabilityKey(meeting, startDateTime) : key;
  if (!payload || payload.key !== expectedKey || typeof payload.expires !== "number" || !Number.isFinite(payload.expires) || payload.expires <= Date.now()) throw failure();
  return key;
}

async function graphRequest(url: string, token: string, body?: Record<string, unknown>): Promise<Record<string, unknown>> {
  const writing = body !== undefined;
  let response: Response;
  try {
    response = await fetch(url, {
      method: writing ? "POST" : "GET", redirect: "error", signal: AbortSignal.timeout(15_000),
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: 'outlook.timezone="UTC"' },
      ...(writing ? { body: JSON.stringify(body) } : {}),
    });
  } catch {
    throw new BookingError(writing ? "BOOKING_UNKNOWN" : "BOOKING_FAILED", writing
      ? "The event request timed out or lost its connection. Its outcome is unknown. Retry the same confirmed request; do not change the details or claim it failed."
      : "Could not check for an existing booking. No new booking was attempted.");
  }
  if (!response.ok) {
    const unknown = writing && response.status >= 500;
    throw new BookingError(unknown ? "BOOKING_UNKNOWN" : "BOOKING_FAILED",
      `Microsoft Graph returned HTTP ${response.status}. ${unknown ? "Creation may have succeeded; retry the same confirmed request." : "No meeting creation was confirmed. Check permissions or retry later."}`);
  }
  try {
    const data: unknown = await response.json();
    if (!data || typeof data !== "object" || Array.isArray(data)) throw new Error();
    return data as Record<string, unknown>;
  } catch {
    throw new BookingError(writing ? "BOOKING_UNKNOWN" : "BOOKING_FAILED", "Microsoft Graph returned an unreadable result. Do not claim a booking outcome; retry the same request.");
  }
}

function onlineDetails(event: Record<string, unknown>) {
  const online = event.onlineMeeting as { joinUrl?: unknown } | null | undefined;
  const joinUrl = typeof online?.joinUrl === "string" && online.joinUrl.startsWith("https://") ? online.joinUrl : null;
  return { joinUrl, meetingLinkStatus: joinUrl ? "available" : "unavailable",
    ...(!joinUrl ? { meetingLinkMessage: "The calendar event exists, but Graph has not returned a Teams join link. Check the organizer account supports Teams meetings. Do not create a duplicate event to obtain a link." } : {}) };
}

async function existingBooking(base: string, token: string, transactionId: string, start: string, end: string): Promise<Record<string, unknown> | null> {
  const url = new URL(`${base}/calendarView`);
  url.searchParams.set("startDateTime", start);
  url.searchParams.set("endDateTime", end);
  url.searchParams.set("$select", "id,subject,start,end,webLink,transactionId,isCancelled,onlineMeeting,isOnlineMeeting,onlineMeetingProvider");
  url.searchParams.set("$top", "200");
  let next: string | undefined = url.toString();
  const seen = new Set<string>();
  for (let page = 0; next && page < 5; page++) {
    const target = new URL(next);
    if (target.origin !== url.origin || target.pathname !== url.pathname || target.username || target.password || target.hash || seen.has(target.href)) {
      throw new BookingError("BOOKING_FAILED", "Invalid calendar pagination response. No event was created.");
    }
    seen.add(target.href);
    const data = await graphRequest(target.href, token);
    if (!Array.isArray(data.value)) throw new BookingError("BOOKING_FAILED", "Could not verify prior bookings.");
    const match = data.value.find(event => event && typeof event === "object" && event.transactionId === transactionId);
    if (match) return match as Record<string, unknown>;
    const link = data["@odata.nextLink"];
    if (link !== undefined && (typeof link !== "string" || !link)) throw new BookingError("BOOKING_FAILED", "Invalid calendar pagination response.");
    next = link as string | undefined;
  }
  if (next) throw new BookingError("BOOKING_FAILED", "Too many calendar entries to safely verify prior bookings. No event was created.");
  return null;
}

export async function createConfirmedMeeting(callId: string, meeting: MeetingDetails, startDateTime: string, confirmed: unknown, confirmationToken: unknown) {
  if (confirmed !== "true") throw new BookingError("CONFIRMATION_REQUIRED", "Ask the caller to confirm the title, attendees, date, time and duration before calling schedule_meeting.");
  const key = verifyConfirmation(callId, meeting, startDateTime, confirmationToken);
  // Stable across retries and instances, even if Vapi creates a new toolCallId.
  const transactionId = `${key.slice(0, 8)}-${key.slice(8, 12)}-${key.slice(12, 16)}-${key.slice(16, 20)}-${key.slice(20, 32)}`;
  const token = await getGraphAccessToken();
  if (!token) throw new BookingError("BOOKING_FAILED", "Microsoft Graph authentication failed. No event was created.");
  const start = new Date(startDateTime).toISOString();
  const end = new Date(Date.parse(start) + meeting.durationMinutes * 60_000).toISOString();
  const base = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(meeting.caller.id)}`;
  const existing = await existingBooking(base, token, transactionId, start, end);
  const details = { ...(meeting.location ? { location: meeting.location } : {}), ...(meeting.description ? { description: meeting.description } : {}), title: meeting.title, date: meeting.date, startTime: meeting.time, durationMinutes: meeting.durationMinutes,
    startDateTime: start, endDateTime: end, timeZone: meeting.timeZone,
    organizer: { displayName: meeting.caller.displayName, email: meeting.caller.email },
    attendees: meeting.attendees.map(user => ({ displayName: user.displayName, email: user.email })),
  };
  if (existing?.isCancelled) return { status: "booking_cancelled", message: "This booking was cancelled. Confirm a new request before recreating it." };
  if (existing) {
    if (typeof existing.id !== "string" || !existing.id) throw new BookingError("BOOKING_UNKNOWN", "A prior booking was found but its details could not be verified.");
    return { status: "already_scheduled", event: { id: existing.id, title: existing.subject, start: existing.start, end: existing.end, webLink: existing.webLink, ...onlineDetails(existing) }, meeting: details,
      message: "This meeting was already created. No duplicate invitations were sent." };
  }
  const availability = await checkCalendarAvailability(meeting, start);
  if (!availability.requestedSlot.available) return { status: "availability_changed", availability,
    message: "The selected time is no longer free for everyone. Run Tool 1 and ask the caller to confirm another slot." };
  const event = await graphRequest(`${base}/events`, token, {
    subject: meeting.title,
    isOnlineMeeting: true, onlineMeetingProvider: "teamsForBusiness",
    ...(meeting.location ? { location: { displayName: meeting.location } } : {}),
    body: invitationBody(meeting.description),
    start: { dateTime: start.slice(0, 19), timeZone: "UTC" },
    end: { dateTime: end.slice(0, 19), timeZone: "UTC" },
    attendees: meeting.attendees.map(user => ({ emailAddress: { address: user.email, name: user.displayName }, type: "required" })),
    transactionId, responseRequested: true,
  });
  if (typeof event.id !== "string" || !event.id) throw new BookingError("BOOKING_UNKNOWN", "The event request returned without a usable event ID. Retry the same confirmed request to verify its outcome.");
  return { status: "scheduled", event: { id: event.id, title: event.subject, start: event.start, end: event.end, webLink: event.webLink, ...onlineDetails(event) }, meeting: details,
    message: "The meeting was created in Microsoft 365 with attendees. Microsoft 365 sends their invitations; delivery and acceptance are not verified." };
}
