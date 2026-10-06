## Current Tool 2 contract: email-only organizer

Tool 1's inputs, attendee lookup, output and availability behavior are unchanged.
Tool 2 requires message.call.id and a callerEmail argument.
No caller phone number, phone-call type or message.call.customer.email is required. callerEmail must resolve uniquely to an enabled
Entra user; missing, ambiguous and invalid identities fail closed. The scheduler bearer credential is still required.

Tool 2 arguments.attendeeEmails is now an array of strings, for example:
```json
{"attendeeEmails":["supportassistant@vanigent.com"],"callerEmail":"askaidev@vanigent.com"}
```
Copy attendee emails from Tool 1 meeting.attendees and supply the authenticated caller email in callerEmail. Other arguments are unchanged: title, date, startTime,
duration (string), confirmed (exact string "true"), confirmationToken, and optional timeZone,
description and location. Copy checked details exactly. The existing attendee-slot token remains unchanged;
organizer identity comes from callerEmail and all calendars are rechecked before booking.
Update the Tool 2 schema in Vapi, deploy the scheduler, and use a fresh Tool 1 token for testing.

## Meeting scheduling behavior

## Tool 2: caller-organized meeting with invitations

POST /api/scheduleMeeting accepts the Vapi tool-calls envelope and the same scheduler bearer credential.
Vapi must supply message.call.id for retry protection. Tool 2 receives callerEmail in its arguments,
resolves that email in Entra and uses the enabled directory user as organizer.
Authentication and Knowledge Base are unchanged.

After Tool 1, read back the selected free slot and get explicit confirmation. Call schedule_meeting with:
- title: exact confirmed title
- attendeeEmails: array of directory email strings from Tool 1 (exclude the organizer)
- callerEmail: authenticated caller email used to resolve the organizer
- date and startTime: the selected slot's absolute local date and HH:mm
- timeZone and duration: the selected slot's timezone and duration (string, e.g. "30")
- confirmed: "true"
- confirmationToken: the selected slot's token

Tool 1 checks listed attendees only. For booking, leave the caller out of its attendee list;
Tool 2 verifies the caller and rechecks everyone's availability, including the caller.
The selected-slot token attests to title, directory attendee identities, time and duration, not caller authentication.
Tool 2's callerEmail argument supplies organizer identity and is validated in Entra. Do not expose the scheduler bearer credential to untrusted clients.

Tool 2 posts one event to the caller's /users/{id}/events with attendees and UTC start/end.
Microsoft 365 sends invitations; attendee acceptance/calendar processing and delivery are not guaranteed by this API response.
Calendars.ReadWrite application permission and tenant consent/mailbox access are required.
A changed slot returns availability_changed without creation. Retries with the same call ID/details avoid duplicate events.
Network ambiguity returns booking_unknown: retry identical arguments instead of claiming failure.
No live meetings are created by the offline tests.

## Current availability endpoint (Postman and Vapi)

POST /api/checkMeetingAvailability accepts a plain JSON body or a Vapi tool-calls envelope.
Use Authorization: Bearer <MEETING_SCHEDULER_SECRET> and Content-Type: application/json.
No caller or attendee phone number is required. Only explicitly listed attendees are checked;
include the caller's name in attendees if their calendar should be included.
Names resolve through Entra; ambiguous or fuzzy matches require clarification before calendars are queried.
This endpoint does not create events. Free requested slots and alternatives include a signed confirmationToken valid for 15 minutes. Tool 2 requires explicit confirmation and rechecks the caller plus attendees.

Example body (use real directory names and a future date):

```json
{"attendees":["Sneha Rao","John Smith"],"date":"2030-01-07","time":"10:00","timeZone":"America/New_York","title":"Project discussion","durationMinutes":30}
```

Plain requests return a status object directly. Vapi requests return correlated results[].result JSON strings.
Legacy startTime/duration strings and comma-separated attendees remain accepted for availability.
Restart npm start after code changes; update the Vapi tool using src/check-meeting-availability.tool.json.
The first resolved mailbox is only Graph getSchedule query context, not an assigned organizer.
Live tenant permissions and mailbox access must still be verified in Postman.

## Previous caller-based booking integration (not the current availability contract)

Microsoft 365 meeting scheduler

The initial meeting request calls `check_meeting_availability`. After the caller
selects a free slot and explicitly confirms the final details, call
`schedule_meeting`. Both endpoints accept Vapi tool-call envelopes. The assistant
prompt is configured separately in Vapi.

Run from this folder:

```powershell
npm ci
npm test
npm start
```

Use Node.js 24, Azure Functions Core Tools v4 and local Azurite when
`AzureWebJobsStorage=UseDevelopmentStorage=true`. Build output and dependencies
are generated. Source files and tests are under `src/`.

Configure these settings in ignored `local.settings.json` Values locally and
Azure Function App settings when deployed:

- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`: existing app-only Graph credentials.
- `MEETING_SCHEDULER_SECRET`: shared with Vapi's custom bearer credential.
- `MEETING_TIME_ZONE`: `America/New_York`, also the code default. An explicit IANA
  `timeZone` argument overrides it for a meeting.

Application permissions require admin consent: `User.Read.All` and
`Calendars.ReadWrite` for both tools together. Check Exchange mailbox access scope.
The existing cached Graph provider is reused; no new OAuth flow was added.

Postman: Tool 1

`POST http://localhost:7071/api/checkMeetingAvailability`

Headers: `Content-Type: application/json` and
`Authorization: Bearer <MEETING_SCHEDULER_SECRET>`.

```json
{
  "message": {
    "type": "tool-calls",
    "call": {
      "id": "postman-meeting-001",
      "type": "inboundPhoneCall",
      "customer": { "number": "+14155551234" }
    },
    "toolCallList": [{
      "id": "check-001",
      "name": "check_meeting_availability",
      "arguments": {
        "title": "Project review",
        "attendees": "Sneha Rao, John Smith",
        "date": "tomorrow",
        "startTime": "15:00",
        "duration": "30 minutes"
      }
    }]
  }
}
```

Replace the example caller number with the actual organizer's Entra phone number
and names with real directory users. The phone identifies the caller from the
credential-protected Vapi metadata; attendee phone numbers are never needed.
Tool arguments cannot choose the organizer. Phone-based identification is the
existing project's assurance model, not Microsoft SSO or proof against caller-ID
spoofing. Web calls are not supported without a separate trusted identity flow.

All meeting inputs above are strings. Separate attendees with commas. Exact names
are checked first; ambiguous or fuzzy matches return names, emails, job titles
and departments. Ask the caller to select a candidate, even when only one fuzzy
suggestion exists, and retry Tool 1 with its directory email. The bounded fuzzy
fallback considers up to 1,000 directory users and returns up to 10 candidates;
it is not exhaustive for larger tenants. Disabled or email-less users are excluded.

Supported dates: YYYY-MM-DD, today, tomorrow, day after tomorrow, in N days,
and weekday names with optional next/this/coming. Weekdays resolve to the next
occurrence strictly after today; Monday on Monday means seven days later.
Relative dates use the actual server clock in the meeting timezone. Ambiguous
numeric dates and unknown wording require clarification rather than guessing.
Start time accepts HH:mm, explicit AM/PM, noon or midnight and returns HH:mm.
Duration accepts minutes or hours, including `1 hour 30 minutes`, from 5 to 480
minutes in five-minute increments. DST gaps and repeated local times require
clarification. End time is calculated by adding the duration to the UTC instant.

Tool 1 checks the caller and all attendees for the entire duration using Graph
getSchedule. If busy, it searches seven rolling days, Monday-Friday 09:00-17:00
in the meeting timezone, returning up to three non-overlapping alternatives on a
15-minute start grid. These are shared search hours, not each user's Outlook
working hours. Unknown calendar data never means free. Tentative/busy/out-of-office
block time; free/working-elsewhere do not. Five-minute free/busy buckets are
conservative when event details are hidden. At most 19 attendees plus the caller
are supported. Results contain resolved meeting/attendee details, availability,
alternatives and a 15-minute confirmationToken for each checked free slot.

Postman: Tool 2

`POST http://localhost:7071/api/scheduleMeeting`

Use the same headers and the SAME call ID. Replace
`toolCallList` in the envelope with the following, using values from the selected
Tool 1 slot and verified attendee emails:

```json
[{
  "id": "schedule-001",
  "name": "schedule_meeting",
  "arguments": {
    "title": "Project review",
    "attendeeEmails": ["sneha@company.com", "john@company.com"],
    "callerEmail": "caller@company.com",
    "date": "2030-01-08",
    "startTime": "15:00",
    "duration": "30 minutes",
    "timeZone": "America/New_York",
    "confirmed": "true",
    "confirmationToken": "COPY-FROM-SELECTED-TOOL-1-SLOT"
  }
}]
```

Use the actual selected YYYY-MM-DD date, not the illustrative date above and not
relative wording. `confirmed` is the string `true` and must only be sent after
explicit caller confirmation. The signed token binds the checked meeting details;
changed or expired details require another Tool 1 check. Tool 2 validates callerEmail independently. The assistant
is responsible for obtaining conversational confirmation; the flag does not
independently prove what the caller said.

Tool 2 re-resolves attendee emails, calculates the end time, checks for a prior
booking, rechecks availability, then POSTs an event with required attendees into
the caller's default calendar. Microsoft 365 sends the invitations automatically;
delivery and acceptance are not verified. No separate sendMail request is made.
A live successful Tool 2 test therefore creates an event and sends invitations.

Duplicate protection uses a stable Graph transactionId based on call ID and exact
meeting details, plus a calendarView lookup on retry. A timeout/5xx response can
mean the event was created: return booking_unknown and retry the SAME arguments
and call ID. Never change details merely to retry. Separate calls are separate
requests. Calendar availability is not an atomic reservation; another booking
can still race the last check. Cancelled prior events are not silently recreated.

Read `results[0].result` as JSON. Tool 1 statuses include available,
alternatives_found, no_common_slot, needs_details, needs_attendee_clarification,
caller_unverified and availability_unknown. Tool 2 adds scheduled,
already_scheduled, availability_changed, confirmation_required, booking_failed,
booking_unknown and booking_cancelled. Invalid credentials return HTTP 401;
tool outcomes use HTTP 200 with the matching toolCallId. Graph response bodies,
access tokens and secrets are not logged. The health endpoint is liveness only.

Import/update `src/check-meeting-availability.tool.json` and
`src/schedule-meeting.tool.json` in Vapi, replace their server URLs and credential
IDs, and attach both tools to the assistant. Old startDateTime/durationMinutes
arguments have been replaced by the string fields shown above. Local files do
not automatically update the Vapi dashboard or deploy the Function App.

Validation: 74 offline tests cover directory search, fuzzy suggestions, current
zoned dates, end times, common availability, both handlers, confirmation binding,
rechecks and duplicate handling. Live tenant consent and mailbox access still
require verification. No events or invitations are sent by these tests.

References: [Graph getSchedule](https://learn.microsoft.com/en-us/graph/api/calendar-getschedule?view=graph-rest-1.0),
[create event and invitations](https://learn.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0),
[event transactionId](https://learn.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0),
[Vapi function tools](https://docs.vapi.ai/tools/custom-tools).


## Optional location and description

New meetings request isOnlineMeeting=true and onlineMeetingProvider=teamsForBusiness. Existing events are not changed.
Both tools accept optional description (plain text, maximum 5000 characters) and location
(plain text, maximum 256 characters). Omit these fields when not supplied by the caller.
Copy the exact values from Tool 1 to Tool 2; changes require a fresh check and confirmation token.
Location is written to Graph location.displayName; it does not reserve a room mailbox.
When there is no join link, event.meetingLinkStatus is unavailable with an explanation; event.webLink still opens Outlook.
Update both Vapi tool definitions and restart the service before testing.

Microsoft generates the Teams invitation section and its join details when the organizer calendar supports Teams. The service returns event.joinUrl from Graph; it does not fabricate or separately retrieve meeting IDs/passcodes. Optional description and location remain supported. Account enablement and live generation must be verified separately.

New invitations use an HTML body with an escaped optional description above the Microsoft-generated Teams section. Teams supplies the real Join link, Meeting ID, Passcode and help/options links; their exact labels and styling depend on Microsoft and the mail client. No sample credentials are inserted. Existing invitations are not rewritten.

Agenda formatting: descriptions use 12pt Segoe UI/Arial. Supply a heading followed by numbered lines (1., 2., etc.) for a bold heading and spaced numbered agenda. Topic labels preceding a spaced dash are bold. Plain descriptions remain paragraphs. Text is HTML-escaped; Teams details are left to Microsoft.
