import { issueAvailabilityConfirmation } from "./meetingBooking";
import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { timingSafeEqual } from "node:crypto";
import { findAttendeeCandidates, EntraUserSearchError } from "./entraUserSearch";
import { AvailabilityError, checkParticipantAvailability } from "./calendarAvailability";
import { parseMeetingInput } from "./meetingTime";
import { EntraUser } from "../types/entraUser";
function object(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;
}
function invalid(message: string): never { throw new EntraUserSearchError("INVALID_INPUT", message); }
async function check(input: unknown) {
  const args = object(input);
  if (!args) return invalid("Provide meeting details as a JSON object.");
  const allowed = ["location", "description", "title", "attendees", "date", "time", "startTime", "duration", "durationMinutes", "timeZone"];
  if (Object.keys(args).some(key => !allowed.includes(key))) return invalid("Use title, attendees, date, time, durationMinutes and optional timeZone. No caller identity or phone number is required.");
  if ((args.time !== undefined && args.startTime !== undefined) || (args.duration !== undefined && args.durationMinutes !== undefined)) return invalid("Supply one time field and one duration field.");
  if (args.durationMinutes !== undefined && !Number.isInteger(args.durationMinutes)) return invalid("durationMinutes must be an integer.");
  if (args.duration !== undefined && typeof args.duration !== "string" && !Number.isInteger(args.duration)) return invalid("duration must be a string or integer number of minutes.");
  const array = Array.isArray(args.attendees);
  if (array && (!(args.attendees as unknown[]).length || (args.attendees as unknown[]).length > 19 ||
    (args.attendees as unknown[]).some(name => typeof name !== "string" || !name.trim() || name.length > 256 || /[\u0000-\u001f]/.test(name)))) return invalid("Provide 1-19 nonempty attendee names.");
  const duration = args.durationMinutes ?? args.duration;
  const parsed = parseMeetingInput({ title: args.title, description: args.description, location: args.location, attendees: array ? "participant" : args.attendees,
    date: args.date, startTime: args.time ?? args.startTime,
    duration: typeof duration === "number" ? String(duration) : duration, timeZone: args.timeZone });
  const names = array ? (args.attendees as string[]).map(name => name.trim()) : parsed.attendeeNames;
  const unique = new Map(names.map(name => [name.toLowerCase(), name]));
  const users = new Map<string, EntraUser>();
  const unresolved = [];
  for (const [key, name] of unique) {
    const found = await findAttendeeCandidates(name);
    const exact = found.users.filter(user => user.accountEnabled === true && user.id && user.email &&
      [user.displayName, user.givenName, user.surname, user.email, user.userPrincipalName].some(value => value?.toLowerCase() === key));
    if (!found.exact || found.users.length !== 1 || exact.length !== 1) {
      unresolved.push({ name, candidates: found.users.map(user => ({ displayName: user.displayName,
        email: user.email, jobTitle: user.jobTitle, department: user.department })) });
    } else users.set(exact[0].id, exact[0]);
  }
  const details = { title: parsed.title, ...(parsed.location ? { location: parsed.location } : {}), ...(parsed.description ? { description: parsed.description } : {}), date: parsed.date, time: parsed.time, startTime: parsed.time,
    timeZone: parsed.timeZone, durationMinutes: parsed.durationMinutes,
    startDateTime: parsed.startDateTime, endDateTime: parsed.endDateTime };
  if (unresolved.length) return { status: "needs_attendee_clarification", attendees: unresolved, meeting: details,
    message: "Clarify the attendee names and retry with the selected directory email. Calendars have not been checked." };
  const participants = [...users.values()];
  const availability = await checkParticipantAvailability(participants, parsed, parsed.startDateTime);
  const checked = { title: parsed.title, ...(parsed.location ? { location: parsed.location } : {}), ...(parsed.description ? { description: parsed.description } : {}), attendees: participants, durationMinutes: parsed.durationMinutes };
  return { ...availability,
    requestedSlot: { ...availability.requestedSlot, ...(availability.requestedSlot.available ? { confirmationToken: issueAvailabilityConfirmation(checked, parsed.startDateTime) } : {}) },
    alternatives: availability.alternatives.map(slot => ({ ...slot, confirmationToken: issueAvailabilityConfirmation(checked, slot.startDateTime) })),
    confirmationExpiresInMinutes: 15, scope: "listed_attendees", meeting: { ...details,
    attendees: participants.map(user => ({ displayName: user.displayName, email: user.email })) } };
}
export async function handleAvailability(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  const secret = process.env.MEETING_SCHEDULER_SECRET;
  const header = request.headers.get("authorization");
  const supplied = Buffer.from(header?.startsWith("Bearer ") ? header.slice(7) : "");
  if (!secret || supplied.length !== Buffer.byteLength(secret) || !timingSafeEqual(supplied, Buffer.from(secret))) return { status: 401, jsonBody: { error: "Unauthorized" } };
  let body: unknown;
  try { body = await request.json(); } catch { return { status: 400, jsonBody: { error: "Invalid JSON" } }; }
  const run = async (input: unknown) => {
    try { return await check(input); } catch (error) {
      const known = error instanceof AvailabilityError || error instanceof EntraUserSearchError;
      const code = known ? error.code : "INTERNAL_ERROR";
      context.warn("[MEETING-AVAILABILITY] " + code);
      return { status: code === "INVALID_INPUT" || code === "INVALID_TIME" ? "needs_details" : "availability_unknown",
        code, message: known ? error.message : "Unable to check availability. Please try again later.",
        ...(known && error.retryAfterSeconds !== undefined ? { retryAfterSeconds: error.retryAfterSeconds } : {}) };
    }
  };
  const root = object(body);
  if (!root) return { status: 400, jsonBody: { error: "Provide a JSON object." } };
  if (!("message" in root)) return { status: 200, jsonBody: await run(root) };
  const message = object(root.message);
  if (!message || message.type !== "tool-calls") return { status: 400, jsonBody: { error: "Expected message.type = 'tool-calls', or send meeting details directly." } };
  const calls = message.toolCallList;
  if (!Array.isArray(calls) || !calls.length || calls.length > 3 || calls.some(call => typeof object(call)?.id !== "string" || !(object(call)?.id as string).trim())) return { status: 400, jsonBody: { error: "Provide one to three tool calls with IDs." } };
  const results = [];
  for (const raw of calls) {
    const call = object(raw)!;
    const fn = object(call.function);
    let args = call.arguments ?? fn?.arguments ?? call.parameters ?? fn?.parameters;
    if (typeof args === "string") { try { args = JSON.parse(args); } catch { args = null; } }
    const result = (call.name ?? fn?.name) === "check_meeting_availability" ? await run(args) : { status: "invalid_request", message: "Unsupported tool name." };
    results.push({ toolCallId: call.id, result: JSON.stringify(result) });
  }
  return { status: 200, jsonBody: { results } };
}
