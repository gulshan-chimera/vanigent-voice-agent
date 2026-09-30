import { HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { timingSafeEqual } from "node:crypto";
import { lookupMeetingCaller, EntraUserSearchError } from "./entraUserSearch";
import { prepareMeetingDetails } from "./meetingDetails";
import { AvailabilityError, checkCalendarAvailability } from "./calendarAvailability";
import { parseMeetingInput } from "./meetingTime";
import { BookingError, createConfirmedMeeting, issueConfirmation } from "./meetingBooking";

function object(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : null;
}

function authorized(request: HttpRequest): boolean {
  const expected = process.env.MEETING_SCHEDULER_SECRET;
  const header = request.headers.get("authorization");
  if (!expected || !header?.startsWith("Bearer ")) return false;
  const supplied = Buffer.from(header.slice(7));
  const secret = Buffer.from(expected);
  return supplied.length === secret.length && timingSafeEqual(supplied, secret);
}

export async function handleMeetingTool(request: HttpRequest, context: InvocationContext, toolName: "check_meeting_availability" | "schedule_meeting"): Promise<HttpResponseInit> {
  if (!authorized(request)) return { status: 401, jsonBody: { error: "Unauthorized" } };
  let body: unknown;
  try { body = await request.json(); } catch { return { status: 400, jsonBody: { error: "Invalid JSON" } }; }
  const message = object(object(body)?.message);
  if (!message || typeof message.type !== "string" || !message.type.trim()) {
    return { status: 400, jsonBody: { error: "Invalid request format. Expected message.type = 'tool-calls' and message.toolCallList containing tool IDs, names and arguments." } };
  }
  if (message.type !== "tool-calls") return { status: 200, jsonBody: { results: [] } };
  const calls = message.toolCallList;
  if (!Array.isArray(calls) || !calls.length || calls.length > 3 || calls.some(call => typeof object(call)?.id !== "string" || !object(call)?.id)) {
    return { status: 400, jsonBody: { error: "Provide one to three tool calls with IDs." } };
  }
  const metadata = object(message.call);
  const phone = object(metadata?.customer)?.number;
  const callerEmail = object(metadata?.customer)?.email;
  let callerPromise: ReturnType<typeof lookupMeetingCaller> | undefined;
  const results = [];
  for (const rawCall of calls) {
    const call = object(rawCall)!;
    let result: unknown;
    try {
      const fn = object(call.function);
      if ((call.name ?? fn?.name) !== toolName) {
        result = { status: "invalid_request", message: "Unsupported tool name for this endpoint." };
      } else {
        let rawArguments = call.arguments ?? fn?.arguments ?? call.parameters ?? fn?.parameters;
        if (typeof rawArguments === "string") {
          try { rawArguments = JSON.parse(rawArguments); } catch { rawArguments = null; }
        }
        const args = object(rawArguments);
        if (!args) throw new AvailabilityError("INVALID_TIME", "Provide meeting details as an object of string fields.");
        const booking = toolName === "schedule_meeting";
        const parsed = parseMeetingInput(args, booking);
        if (booking && (args.confirmed !== "true" || typeof args.confirmationToken !== "string" || !args.confirmationToken)) {
          throw new BookingError("CONFIRMATION_REQUIRED", "Obtain explicit caller confirmation and pass the selected slot's confirmationToken from Tool 1.");
        }
        if (typeof phone !== "string" || typeof metadata?.id !== "string" || !metadata.id ||
            !["inboundPhoneCall", "outboundPhoneCall"].includes(String(metadata.type))) {
          result = { status: "caller_unverified", message: "A phone call with a verifiable caller number is required. Attendee phone numbers are not needed." };
        } else {
          // The protected call metadata, not model arguments, identifies the organizer.
          callerPromise ??= lookupMeetingCaller(phone);
          const caller = await callerPromise;
          if (!caller || (booking && (typeof callerEmail !== "string" || caller.email.toLowerCase() !== callerEmail.trim().toLowerCase()))) {
            result = { status: "caller_unverified", message: "Caller phone and email must identify the same enabled Entra user. Supply message.call.customer.number and message.call.customer.email." };
          } else {
            const preparation = await prepareMeetingDetails(caller, parsed, !booking);
            if (preparation.status === "needs_attendee_clarification") {
              result = {
                status: preparation.status,
                attendees: preparation.attendees.map(entry => ({ name: entry.name,
                  candidates: entry.candidates.map(user => ({ displayName: user.displayName, email: user.email, jobTitle: user.jobTitle, department: user.department })),
                })),
                meeting: { ...(parsed.location ? { location: parsed.location } : {}), ...(parsed.description ? { description: parsed.description } : {}), title: parsed.title, date: parsed.date, startTime: parsed.time, timeZone: parsed.timeZone,
                  durationMinutes: parsed.durationMinutes, startDateTime: parsed.startDateTime, endDateTime: parsed.endDateTime },
                message: "Ask the caller to identify the correct attendee, even for one fuzzy candidate. Retry Tool 1 using the selected email. Calendar availability has not been checked. Fuzzy search covers at most 1,000 directory candidates.",
              };
            } else if (preparation.status !== "ready_for_confirmation") result = preparation;
            else {
              const meeting = { ...preparation.meeting, ...(parsed.location ? { location: parsed.location } : {}), ...(parsed.description ? { description: parsed.description } : {}), title: parsed.title };
              if (booking) {
                result = await createConfirmedMeeting(metadata.id, meeting, parsed.startDateTime, args.confirmed, args.confirmationToken);
              } else {
                const availability = await checkCalendarAvailability(meeting, parsed.startDateTime);
                result = {
                  ...availability,
                  meeting: { ...(parsed.location ? { location: parsed.location } : {}), ...(parsed.description ? { description: parsed.description } : {}), title: parsed.title, date: parsed.date, startTime: parsed.time, durationMinutes: parsed.durationMinutes,
                    startDateTime: parsed.startDateTime, endDateTime: parsed.endDateTime, timeZone: parsed.timeZone,
                    organizer: { displayName: caller.displayName, email: caller.email },
                    attendees: meeting.attendees.map(user => ({ displayName: user.displayName, email: user.email, jobTitle: user.jobTitle, department: user.department })),
                  },
                  requestedSlot: { ...availability.requestedSlot,
                    ...(availability.requestedSlot.available ? { confirmationToken: issueConfirmation(metadata.id, meeting, parsed.startDateTime) } : {}),
                  },
                  alternatives: availability.alternatives.map(slot => ({ ...slot, confirmationToken: issueConfirmation(metadata.id as string, meeting, slot.startDateTime) })),
                  confirmationExpiresInMinutes: 15,
                };
              }
            }
          }
        }
      }
    } catch (error) {
      const known = error instanceof AvailabilityError || error instanceof EntraUserSearchError || error instanceof BookingError;
      const code = known ? error.code : "INTERNAL_ERROR";
      context.warn(`[MEETING-TOOL] ${code}`);
      result = {
        status: code === "INVALID_TIME" || code === "INVALID_INPUT" ? "needs_details"
          : error instanceof BookingError ? error.code.toLowerCase()
          : toolName === "schedule_meeting" ? "booking_unknown" : "availability_unknown",
        code, message: known ? error.message : "Unable to complete the request. Do not assume a booking outcome; retry the same details.",
        ...((error instanceof AvailabilityError || error instanceof EntraUserSearchError) && error.retryAfterSeconds !== undefined ? { retryAfterSeconds: error.retryAfterSeconds } : {}),
      };
    }
    results.push({ toolCallId: call.id, result: JSON.stringify(result) });
  }
  return { status: 200, jsonBody: { results } };
}
