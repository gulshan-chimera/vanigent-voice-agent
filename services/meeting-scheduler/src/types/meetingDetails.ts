import { EntraUser } from "./entraUser";

export interface MeetingDetailsInput {
  attendeeNames?: string[];
  date?: string;
  time?: string;
  timeZone?: string;
  durationMinutes?: number;
}

export interface MeetingDetails {
  title?: string;
  description?: string;
  location?: string;
  caller: EntraUser;
  attendees: EntraUser[];
  date: string;
  time: string;
  timeZone: string;
  durationMinutes: number;
}

export interface ParsedMeetingInput {
  title: string;
  description?: string;
  location?: string;
  attendeeNames: string[];
  date: string;
  time: string;
  timeZone: string;
  durationMinutes: number;
  startDateTime: string;
  endDateTime: string;
}

export type MeetingPreparationResult =
  | {
      status: "needs_details";
      fields: (keyof MeetingDetailsInput | "caller")[];
    }
  | {
      status: "needs_attendee_clarification";
      attendees: { name: string; candidates: EntraUser[] }[];
    }
  | { status: "ready_for_confirmation"; meeting: MeetingDetails };
