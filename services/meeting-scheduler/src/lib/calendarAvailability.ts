import { getGraphAccessToken } from "./graphAuth";
import { MeetingDetails } from "../types/meetingDetails";

const MINUTE = 60_000;
const DAY = 24 * 60 * MINUTE;
const SEARCH_DAYS = 7;
const MAX_OPTIONS = 3;

export class AvailabilityError extends Error {
  constructor(
    public readonly code: "INVALID_TIME" | "AUTH_FAILED" | "GRAPH_ERROR" | "THROTTLED" | "CALENDAR_UNAVAILABLE",
    message: string,
    public readonly retryAfterSeconds?: number,
  ) {
    super(message);
    this.name = "AvailabilityError";
  }
}

interface Interval { start: number; end: number }
interface Schedule { email: string; blocked: Interval[] }

function formatter(timeZone: string): Intl.DateTimeFormat {
  try {
    return new Intl.DateTimeFormat("en-GB", {
      timeZone, year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hourCycle: "h23", weekday: "short",
    });
  } catch {
    throw new AvailabilityError("INVALID_TIME", "Provide a valid IANA timezone, for example America/New_York.");
  }
}

function localParts(date: number, format: Intl.DateTimeFormat): Record<string, string> {
  return Object.fromEntries(format.formatToParts(date).map(part => [part.type, part.value]));
}

export function zonedParts(instant: number, timeZone: string): { date: string; time: string } {
  const parts = localParts(instant, formatter(timeZone));
  return { date: `${parts.year}-${parts.month}-${parts.day}`, time: `${parts.hour}:${parts.minute}` };
}

/** Local offsets must agree with the IANA zone; Z represents an absolute UTC instant. */
export function validateRequestedTime(startDateTime: unknown, timeZone: unknown, duration: unknown, now = Date.now()): number {
  if (typeof startDateTime !== "string" || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:00(?:\.000)?(?:Z|[+-]\d{2}:\d{2})$/.test(startDateTime) ||
      typeof timeZone !== "string" || !timeZone.trim() || typeof duration !== "number" ||
      !Number.isInteger(duration) || duration < 5 || duration > 480 || duration % 5 !== 0) {
    throw new AvailabilityError("INVALID_TIME", "Provide startDateTime with a UTC offset, an IANA timezone, and a duration of 5–480 minutes in five-minute increments.");
  }
  const start = Date.parse(startDateTime);
  if (!Number.isFinite(start) || start <= now || start > now + 365 * DAY) {
    throw new AvailabilityError("INVALID_TIME", "Choose a valid future date within the next year.");
  }
  const requestedFormat = formatter(timeZone);
  const local = localParts(start, startDateTime.endsWith("Z") ? formatter("UTC") : requestedFormat);
  const rendered = `${local.year}-${local.month}-${local.day}T${local.hour}:${local.minute}:00`;
  if (rendered !== startDateTime.slice(0, 19)) {
    throw new AvailabilityError("INVALID_TIME", "The date, UTC offset and timezone disagree, or the local time does not exist. Please confirm them.");
  }
  return start;
}

function unavailable(): AvailabilityError {
  return new AvailabilityError("CALENDAR_UNAVAILABLE", "Could not establish availability for every participant. Do not treat missing calendar data as free.");
}

function graphDate(value: unknown): number {
  if (!value || typeof value !== "object") throw unavailable();
  const { dateTime, timeZone } = value as Record<string, unknown>;
  if (timeZone !== "UTC" || typeof dateTime !== "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?Z?$/.test(dateTime)) throw unavailable();
  const parsed = Date.parse(dateTime.endsWith("Z") ? dateTime : `${dateTime}Z`);
  if (!Number.isFinite(parsed)) throw unavailable();
  return parsed;
}

async function getSchedules(organizerId: string, emails: string[], start: number, end: number): Promise<Schedule[]> {
  const token = await getGraphAccessToken();
  if (!token) throw new AvailabilityError("AUTH_FAILED", "Microsoft Graph authentication failed.");
  let response: Response;
  try {
    response = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(organizerId)}/calendar/getSchedule`, {
      method: "POST", redirect: "error", signal: AbortSignal.timeout(15_000),
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: 'outlook.timezone="UTC"' },
      body: JSON.stringify({
        schedules: emails,
        startTime: { dateTime: new Date(start).toISOString().slice(0, 19), timeZone: "UTC" },
        endTime: { dateTime: new Date(end).toISOString().slice(0, 19), timeZone: "UTC" },
        availabilityViewInterval: 5,
      }),
    });
  } catch {
    throw new AvailabilityError("GRAPH_ERROR", "The calendar request failed or timed out. Please try again later.");
  }
  if (!response.ok) {
    const retry = response.headers.get("retry-after");
    const seconds = retry && /^\d+$/.test(retry) ? Number(retry) : undefined;
    throw new AvailabilityError(response.status === 429 ? "THROTTLED" : "GRAPH_ERROR",
      `Calendar access failed (HTTP ${response.status}). Availability has not been established.`, seconds);
  }
  let body: unknown;
  try { body = await response.json(); } catch { throw unavailable(); }
  const values = (body as { value?: unknown[] } | null)?.value;
  if (!Array.isArray(values) || values.length !== emails.length) throw unavailable();
  return emails.map(email => {
    const matches = values.filter(value => value && typeof value === "object" &&
      typeof (value as Record<string, unknown>).scheduleId === "string" &&
      ((value as Record<string, unknown>).scheduleId as string).toLowerCase() === email.toLowerCase());
    if (matches.length !== 1) throw unavailable();
    const schedule = matches[0] as Record<string, unknown>;
    const view = schedule.availabilityView;
    if (schedule.error || !Array.isArray(schedule.scheduleItems) || typeof view !== "string" ||
        view.length !== Math.ceil((end - start) / (5 * MINUTE)) || !/^[0-3]+$/.test(view)) throw unavailable();
    const blocked: Interval[] = [];
    // Include the coarse view as a conservative fallback when event details are hidden.
    for (let index = 0; index < view.length; index++) {
      if (view[index] !== "0") blocked.push({ start: start + index * 5 * MINUTE, end: start + (index + 1) * 5 * MINUTE });
    }
    for (const item of schedule.scheduleItems) {
      if (!item || typeof item !== "object") throw unavailable();
      const entry = item as Record<string, unknown>;
      const status = typeof entry.status === "string" ? entry.status.toLowerCase() : "unknown";
      if (!["free", "workingelsewhere", "tentative", "busy", "oof"].includes(status)) throw unavailable();
      const interval = { start: graphDate(entry.start), end: graphDate(entry.end) };
      if (interval.end <= interval.start) throw unavailable();
      if (status !== "free" && status !== "workingelsewhere") blocked.push(interval);
    }
    return { email, blocked };
  });
}

function isFree(schedules: Schedule[], start: number, end: number): boolean {
  return schedules.every(schedule => !schedule.blocked.some(busy => busy.start < end && busy.end > start));
}

export async function checkCalendarAvailability(meeting: MeetingDetails, startDateTime: string) {
  return checkParticipantAvailability([meeting.caller, ...meeting.attendees], meeting, startDateTime);
}

/** First mailbox supplies query context, not organizer identity. Checks only listed participants. */
export async function checkParticipantAvailability(participants: MeetingDetails["attendees"], meeting: Pick<MeetingDetails, "timeZone" | "durationMinutes">, startDateTime: string) {
  if (!participants.length) throw new AvailabilityError("CALENDAR_UNAVAILABLE", "Provide at least one resolved attendee.");
  const start = validateRequestedTime(startDateTime, meeting.timeZone, meeting.durationMinutes);
  const duration = meeting.durationMinutes * MINUTE;
  const end = start + duration;
  const searchEnd = start + SEARCH_DAYS * DAY;
  const emails = [...new Set(participants.map(user => user.email.toLowerCase()))];
  if (emails.length > 20) throw new AvailabilityError("CALENDAR_UNAVAILABLE", "At most 20 participants including the caller are supported.");
  const schedules = await getSchedules(participants[0].id, emails, start, searchEnd);
  const format = formatter(meeting.timeZone);
  const slot = (from: number) => ({
    startDateTime: new Date(from).toISOString(), endDateTime: new Date(from + duration).toISOString(),
    date: zonedParts(from, meeting.timeZone).date, startTime: zonedParts(from, meeting.timeZone).time,
    endDate: zonedParts(from + duration, meeting.timeZone).date, endTime: zonedParts(from + duration, meeting.timeZone).time,
    timeZone: meeting.timeZone, durationMinutes: meeting.durationMinutes,
  });
  const available = isFree(schedules, start, end);
  const alternatives: ReturnType<typeof slot>[] = [];
  if (!available) {
    for (let candidate = Math.ceil(start / (15 * MINUTE)) * 15 * MINUTE;
      candidate + duration <= searchEnd && alternatives.length < MAX_OPTIONS; candidate += 15 * MINUTE) {
      if (candidate <= Date.now()) continue;
      const from = localParts(candidate, format);
      const to = localParts(candidate + duration, format);
      const sameDate = from.year === to.year && from.month === to.month && from.day === to.day;
      const endMinute = Number(to.hour) * 60 + Number(to.minute);
      if (["Sat", "Sun"].includes(from.weekday) || Number(from.hour) < 9 || !sameDate || endMinute > 17 * 60) continue;
      if (isFree(schedules, candidate, candidate + duration)) {
        alternatives.push(slot(candidate));
        candidate += Math.ceil(duration / (15 * MINUTE)) * 15 * MINUTE - 15 * MINUTE;
      }
    }
  }
  return {
    status: available ? "available" : alternatives.length ? "alternatives_found" : "no_common_slot",
    requestedSlot: { ...slot(start), available }, alternatives,
    participants: participants.map(user => ({ id: user.id, displayName: user.displayName, email: user.email })),
    searchPolicy: { days: SEARCH_DAYS, weekdays: "Monday–Friday", localHours: "09:00–17:00", timeZone: meeting.timeZone, maximumAlternatives: MAX_OPTIONS },
    searchEndsAt: new Date(searchEnd).toISOString(),
    message: available ? "Everyone is free at the requested time. No meeting has been created."
      : alternatives.length ? "These alternative slots are free for everyone. Ask the caller to choose; no meeting has been created."
      : "No common free slot was found within the search window and working hours. Ask for another date or duration.",
  };
}
