import { AvailabilityError, validateRequestedTime, zonedParts } from "./calendarAvailability";
import { ParsedMeetingInput } from "../types/meetingDetails";

const DAY = 86_400_000;
const WEEKDAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

function invalid(message: string): never {
  throw new AvailabilityError("INVALID_TIME", message);
}

function parseDate(value: string, timeZone: string, now: number): string {
  const date = value.trim().toLowerCase().replace(/\s+/g, " ");
  const today = zonedParts(now, timeZone).date;
  const midnight = Date.parse(`${today}T00:00:00Z`);
  const offsets: Record<string, number> = { today: 0, tomorrow: 1, "day after tomorrow": 2, "the day after tomorrow": 2 };
  if (Object.hasOwn(offsets, date)) return new Date(midnight + offsets[date] * DAY).toISOString().slice(0, 10);
  const weekday = /^(?:(?:next|this|coming) )?(sunday|monday|tuesday|wednesday|thursday|friday|saturday)$/.exec(date);
  if (weekday) {
    const difference = (WEEKDAYS.indexOf(weekday[1]) - new Date(midnight).getUTCDay() + 7) % 7 || 7;
    return new Date(midnight + difference * DAY).toISOString().slice(0, 10);
  }
  const inDays = /^in (\d{1,3}) days?$/.exec(date);
  if (inDays) return new Date(midnight + Number(inDays[1]) * DAY).toISOString().slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    const parsed = Date.parse(`${date}T00:00:00Z`);
    if (Number.isFinite(parsed) && new Date(parsed).toISOString().slice(0, 10) === date) return date;
  }
  return invalid("Confirm the date as YYYY-MM-DD, today, tomorrow, next Monday (or another weekday), or in N days. Ambiguous numeric dates are not assumed.");
}

function parseTime(value: string): string {
  const time = value.trim().toLowerCase().replace(/\./g, "");
  if (time === "noon") return "12:00";
  if (time === "midnight") return "00:00";
  const clock24 = /^(\d{1,2}):(\d{2})$/.exec(time);
  if (clock24 && Number(clock24[1]) < 24 && Number(clock24[2]) < 60) return `${clock24[1].padStart(2, "0")}:${clock24[2]}`;
  const clock12 = /^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/.exec(time);
  if (clock12 && Number(clock12[1]) >= 1 && Number(clock12[1]) <= 12 && Number(clock12[2] ?? 0) < 60) {
    const hour = Number(clock12[1]) % 12 + (clock12[3] === "pm" ? 12 : 0);
    return `${String(hour).padStart(2, "0")}:${clock12[2] ?? "00"}`;
  }
  return invalid("Confirm the start time using HH:mm or an explicit AM/PM time, for example 15:00 or 3 PM.");
}

function parseDuration(value: string): number {
  const duration = value.trim().toLowerCase().replace(/^half an hour$/, "30 minutes").replace(/^an hour$/, "1 hour");
  const numeric = /^\d{1,3}$/.test(duration) ? Number(duration) : NaN;
  const units = /^(?:(\d+(?:\.\d+)?)\s*(?:hours?|hrs?|h))?\s*(?:and\s+)?(?:(\d+)\s*(?:minutes?|mins?|m))?$/.exec(duration);
  const minutes = Number.isFinite(numeric) ? numeric : units && (units[1] || units[2]) ? Number(units[1] ?? 0) * 60 + Number(units[2] ?? 0) : NaN;
  if (!Number.isInteger(minutes) || minutes < 5 || minutes > 480 || minutes % 5 !== 0) {
    return invalid("Use a duration of 5–480 minutes in five-minute increments, for example 30 minutes or 1 hour 30 minutes.");
  }
  return minutes;
}

/** Finds valid UTC instants by round-tripping local time; gaps/folds require clarification. */
function localToInstant(date: string, time: string, timeZone: string): number {
  const wall = Date.parse(`${date}T${time}:00Z`);
  const offsets = new Set<number>();
  for (const hours of [-36, -24, -12, 0, 12, 24, 36]) {
    const instant = wall + hours * 3_600_000;
    const local = zonedParts(instant, timeZone);
    offsets.add(Date.parse(`${local.date}T${local.time}:00Z`) - instant);
  }
  const matches = [...offsets].map(offset => wall - offset).filter(instant => {
    const local = zonedParts(instant, timeZone);
    return local.date === date && local.time === time;
  });
  if (matches.length !== 1) return invalid(matches.length ? "This time occurs twice during daylight saving. Choose an unambiguous time." : "This local time does not exist during daylight saving. Choose another time.");
  return matches[0];
}

export function parseMeetingInput(args: Record<string, unknown>, booking = false, now = Date.now()): ParsedMeetingInput {
  const attendeeField = booking ? "attendeeEmails" : "attendees";
  const required = ["title", attendeeField, "date", "startTime", "duration"];
  const allowed = [...required, "timeZone", "description", "location", ...(booking ? ["confirmed", "confirmationToken"] : [])];
  if (Object.keys(args).some(key => !allowed.includes(key)) || required.some(key => typeof args[key] !== "string" || !(args[key] as string).trim()) ||
      (args.timeZone !== undefined && (typeof args.timeZone !== "string" || !args.timeZone.trim()))) {
    return invalid(`Provide ${required.join(", ")} as nonempty strings. Optional timeZone must also be a string. Caller identity and attendee phone numbers are not tool arguments.`);
  }
  if (args.description !== undefined && (typeof args.description !== "string" || args.description.length > 5000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(args.description))) return invalid("Description must be plain text of at most 5000 characters.");
  if (args.location !== undefined && (typeof args.location !== "string" || args.location.length > 256 || /[\u0000-\u001f]/.test(args.location))) return invalid("Location must be text of at most 256 characters without control characters.");
  const location = typeof args.location === "string" ? args.location.trim() || undefined : undefined;
  const description = typeof args.description === "string" ? args.description.trim() || undefined : undefined;
  const title = (args.title as string).trim();
  if (title.length > 200 || /[\u0000-\u001f]/.test(title)) return invalid("Use a meeting title of 1–200 characters without control characters.");
  const rawNames = args[attendeeField] as string;
  if (rawNames.length > 5000) return invalid("Too many attendee details.");
  const attendeeNames = [...new Set(rawNames.split(/[,;\n]|\s+and\s+/i).map(name => name.trim()).filter(Boolean))];
  if (!attendeeNames.length || attendeeNames.length > 19 || attendeeNames.some(name => name.length > 256)) return invalid("Provide 1–19 comma-separated attendees.");
  if (booking && attendeeNames.some(email => !/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email))) return invalid("Tool 2 requires confirmed attendee email addresses, not names.");
  if (booking && !/^\d{4}-\d{2}-\d{2}$/.test((args.date as string).trim())) return invalid("Use the resolved YYYY-MM-DD date returned by Tool 1 when confirming a booking.");
  const timeZone = typeof args.timeZone === "string" ? args.timeZone.trim() : process.env.MEETING_TIME_ZONE?.trim() || "America/New_York";
  const date = parseDate(args.date as string, timeZone, now);
  const time = parseTime(args.startTime as string);
  const durationMinutes = parseDuration(args.duration as string);
  const start = localToInstant(date, time, timeZone);
  const startDateTime = new Date(start).toISOString();
  validateRequestedTime(startDateTime, timeZone, durationMinutes, now);
  return { title, ...(location ? { location } : {}), ...(description ? { description } : {}), attendeeNames, date, time, timeZone, durationMinutes,
    startDateTime, endDateTime: new Date(start + durationMinutes * 60_000).toISOString() };
}
