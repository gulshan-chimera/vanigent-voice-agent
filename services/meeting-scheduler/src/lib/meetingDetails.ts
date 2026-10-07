import { searchEntraUsers, findAttendeeCandidates } from "./entraUserSearch";
import { EntraUser } from "../types/entraUser";
import {
  MeetingDetailsInput,
  MeetingPreparationResult,
} from "../types/meetingDetails";

/** Caller must come from trusted server-side identity resolution, never tool arguments.
 * This prepares details only; it does not authenticate a caller or check calendars.
 */
export async function prepareMeetingDetails(
  caller: EntraUser | null,
  input: MeetingDetailsInput,
  useFuzzySearch = false,
): Promise<MeetingPreparationResult> {
  const fields: (keyof MeetingDetailsInput | "caller")[] = [];
  if (!caller || caller.accountEnabled !== true || !caller.id || !caller.email)
    fields.push("caller");

  const names = Array.isArray(input.attendeeNames)
    ? input.attendeeNames
        .filter(
          (name): name is string => typeof name === "string" && !!name.trim(),
        )
        .map((name) => name.trim())
    : [];
  if (
    !names.length ||
    names.length !== input.attendeeNames?.length ||
    names.length > 19
  )
    fields.push("attendeeNames");

  // Preserve the caller's wording. Natural-language date/time conversion belongs
  // to the later availability tool; no timezone is inferred from the server.
  const date = typeof input.date === "string" ? input.date.trim() : "";
  const time = typeof input.time === "string" ? input.time.trim() : "";
  const timeZone =
    typeof input.timeZone === "string" ? input.timeZone.trim() : "";
  if (!date) fields.push("date");
  if (!time) fields.push("time");
  if (!timeZone) fields.push("timeZone");
  if (
    !Number.isInteger(input.durationMinutes) ||
    input.durationMinutes! <= 0 ||
    input.durationMinutes! > 1440
  )
    fields.push("durationMinutes");
  if (fields.length || !caller) return { status: "needs_details", fields };

  const attendees = new Map<string, EntraUser>();
  const unresolved: { name: string; candidates: EntraUser[] }[] = [];
  const searched = new Set<string>();
  for (const name of names) {
    const key = name.toLowerCase();
    if (searched.has(key)) continue;
    searched.add(key);
    const search = useFuzzySearch
      ? await findAttendeeCandidates(name)
      : { users: await searchEntraUsers(name), exact: true };
    const candidates = search.users;
    // One prefix match is not proof of the intended person. Only accept a
    // unique exact display-name or given-name match, then ask for confirmation.
    const exact = candidates.filter(
      (user) =>
        user.displayName.toLowerCase() === key ||
        user.givenName?.toLowerCase() === key ||
        user.surname?.toLowerCase() === key ||
        user.email.toLowerCase() === key ||
        user.userPrincipalName?.toLowerCase() === key,
    );
    if (!search.exact || candidates.length !== 1 || exact.length !== 1) {
      unresolved.push({ name, candidates });
    } else if (exact[0].id !== caller.id) {
      attendees.set(exact[0].id, exact[0]);
    }
  }
  if (unresolved.length)
    return { status: "needs_attendee_clarification", attendees: unresolved };
  if (!attendees.size)
    return { status: "needs_details", fields: ["attendeeNames"] };
  return {
    status: "ready_for_confirmation",
    meeting: {
      caller,
      attendees: [...attendees.values()],
      date,
      time,
      timeZone,
      durationMinutes: input.durationMinutes!,
    },
  };
}
