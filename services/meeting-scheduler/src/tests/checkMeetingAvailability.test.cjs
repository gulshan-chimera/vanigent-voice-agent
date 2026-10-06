const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const { HttpRequest } = require('@azure/functions');
const directory = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');
const calendar = require('../../dist/meeting-scheduler/src/lib/calendarAvailability.js');
// Legacy shared handler coverage retained for the deferred booking flow.
const { handleMeetingTool } = require('../../dist/meeting-scheduler/src/lib/vapiMeetingTool.js');
const checkMeetingAvailability = (request, context) => handleMeetingTool(request, context, 'check_meeting_availability');
const { scheduleMeeting } = require('../../dist/meeting-scheduler/src/functions/scheduleMeeting.js');
const provider = require('../../dist/meeting-scheduler/src/lib/graphAuth.js');
const saved = { lookup: directory.lookupMeetingCaller, search: directory.searchEntraUsers, find: directory.findAttendeeCandidates, check: calendar.checkCalendarAvailability, now: Date.now, secret: process.env.MEETING_SCHEDULER_SECRET, fetch: global.fetch, token: provider.getGraphAccessToken };
afterEach(() => {
  directory.lookupMeetingCaller = saved.lookup; directory.searchEntraUsers = saved.search;
  directory.findAttendeeCandidates = saved.find;
  calendar.checkCalendarAvailability = saved.check; Date.now = saved.now;
  global.fetch = saved.fetch; provider.getGraphAccessToken = saved.token;
  if (saved.secret === undefined) delete process.env.MEETING_SCHEDULER_SECRET; else process.env.MEETING_SCHEDULER_SECRET = saved.secret;
});
const user = { id: 'caller', displayName: 'Caller', givenName: 'Caller', email: 'caller@example.com', accountEnabled: true };
const attendee = { ...user, id: 'attendee', displayName: 'Sneha Rao', givenName: 'Sneha', email: 'sneha@example.com', jobTitle: 'Manager', department: 'Finance' };
const args = { title: 'Project review', attendees: 'Sneha', date: '2030-01-07', startTime: '10:00', timeZone: 'UTC', duration: '30 minutes' };
const call = { id: 'call-1', type: 'inboundPhoneCall', customer: { number: '+14155551234', email: 'caller@example.com' } };
const tool = { id: 'tool-1', name: 'check_meeting_availability', arguments: args };
const context = { warn() {} };
function setup() {
  process.env.MEETING_SCHEDULER_SECRET = 'test-secret';
  Date.now = () => Date.parse('2030-01-01T00:00:00Z');
  directory.lookupMeetingCaller = async phone => { assert.equal(phone, call.customer.number); return user; };
  directory.searchEntraUsers = async name => name.toLowerCase() === user.email ? [user] : [attendee];
  directory.findAttendeeCandidates = async name => ({ users: await directory.searchEntraUsers(name), exact: true });
  calendar.checkCalendarAvailability = async meeting => { assert.equal(meeting.caller.id, 'caller'); return { status: 'available', requestedSlot: { available: true }, alternatives: [] }; };
}
function request(tools = [tool], metadata = call, secret = 'test-secret') {
  return new HttpRequest({ method: 'POST', url: 'http://localhost/api/checkMeetingAvailability', headers: { authorization: `Bearer ${secret}`, 'content-type': 'application/json' },
    body: { string: JSON.stringify({ message: { type: 'tool-calls', call: metadata, toolCallList: tools } }) } });
}
const result = response => JSON.parse(response.jsonBody.results[0].result);

test('rejects plain meeting details and missing message type with an actionable error', async () => {
  setup();
  directory.lookupMeetingCaller = async () => assert.fail('must not call Graph');
  for (const body of [{ attendees: ['Sneha'], date: '2030-01-07', time: '10:00', title: 'Review', duration: 30 }, { message: {} }]) {
    const req = new HttpRequest({ method: 'POST', url: 'http://localhost/api/checkMeetingAvailability',
      headers: { authorization: 'Bearer test-secret', 'content-type': 'application/json' },
      body: { string: JSON.stringify(body) } });
    const response = await checkMeetingAvailability(req, context);
    assert.equal(response.status, 400);
    assert.match(response.jsonBody.error, /message.toolCallList/);
  }
});

test('requires configured bearer credential before processing any data', async () => {
  setup(); directory.lookupMeetingCaller = async () => assert.fail('must not call Graph');
  assert.equal((await checkMeetingAvailability(request([tool], call, 'bad'), context)).status, 401);
  delete process.env.MEETING_SCHEDULER_SECRET;
  assert.equal((await checkMeetingAvailability(request(), context)).status, 401);
});
test('returns Vapi results matched to toolCallId and uses metadata caller', async () => {
  setup();
  const response = await checkMeetingAvailability(request(), context);
  assert.equal(response.status, 200);
  assert.equal(response.jsonBody.results[0].toolCallId, 'tool-1');
  assert.equal(result(response).status, 'available');
});
test('supports nested function arguments and parameters as JSON or objects', async () => {
  setup();
  for (const shaped of [
    { id: 'nested', function: { name: tool.name, arguments: JSON.stringify(args) } },
    { id: 'parameters', name: tool.name, parameters: args },
  ]) assert.equal(result(await checkMeetingAvailability(request([shaped]), context)).status, 'available');
});
test('rejects model-supplied caller identity, invalid arguments, unknown tools and web callers', async () => {
  setup(); calendar.checkCalendarAvailability = async () => assert.fail('must not query calendars');
  for (const argumentsValue of [{ ...args, callerEmail: 'victim@example.com' }, {}, 'bad-json', { ...args, attendees: [] }, { ...args, timeZone: 'invalid' }]) {
    assert.equal(result(await checkMeetingAvailability(request([{ ...tool, arguments: argumentsValue }]), context)).status, 'needs_details');
  }
  assert.equal(result(await checkMeetingAvailability(request([{ ...tool, name: 'schedule_meeting' }]), context)).status, 'invalid_request');
  assert.equal(result(await checkMeetingAvailability(request([tool], { ...call, type: 'webCall' }), context)).status, 'caller_unverified');
});
test('unresolved caller and ambiguous attendees cannot query calendars', async () => {
  setup(); calendar.checkCalendarAvailability = async () => assert.fail('must not query calendars');
  directory.lookupMeetingCaller = async () => null;
  assert.equal(result(await checkMeetingAvailability(request(), context)).status, 'caller_unverified');
  directory.lookupMeetingCaller = async () => user;
  directory.searchEntraUsers = async () => [attendee, { ...attendee, id: 'other', email: 'other@example.com' }];
  const outcome = result(await checkMeetingAvailability(request(), context));
  assert.equal(outcome.status, 'needs_attendee_clarification');
  assert.equal(outcome.attendees[0].candidates.length, 2);
});
test('partial tool failure preserves all results and sanitizes unexpected errors', async () => {
  setup(); calendar.checkCalendarAvailability = async () => { throw new Error('private-token'); };
  const response = await checkMeetingAvailability(request([tool, { ...tool, id: 'tool-2', name: 'bad' }]), context);
  assert.equal(response.jsonBody.results.length, 2);
  assert.equal(result(response).status, 'availability_unknown');
  assert.ok(!JSON.stringify(response).includes('private-token'));
});

test('fuzzy suggestions include job title and department and never reach calendars', async () => {
  setup(); directory.findAttendeeCandidates = async () => ({ users: [attendee], exact: false });
  calendar.checkCalendarAvailability = async () => assert.fail('must clarify fuzzy match');
  const outcome = result(await checkMeetingAvailability(request(), context));
  assert.equal(outcome.status, 'needs_attendee_clarification');
  assert.deepEqual(outcome.attendees[0].candidates[0], { displayName: 'Sneha Rao', email: 'sneha@example.com', jobTitle: 'Manager', department: 'Finance' });
});

test('same-name users return name and department, then the selected email resolves the person', async () => {
  setup();
  const finance = { ...attendee, department: 'Finance' };
  const hr = { ...attendee, id: 'hr-user', email: 'sneha.hr@example.com', department: 'HR' };
  directory.findAttendeeCandidates = async name => ({ users: name === hr.email ? [hr] : [finance, hr], exact: true });
  let calendarCalls = 0;
  calendar.checkCalendarAvailability = async meeting => {
    calendarCalls++;
    assert.deepEqual(meeting.attendees.map(person => person.id), ['hr-user']);
    return { status: 'available', requestedSlot: { available: true }, alternatives: [] };
  };
  const ambiguous = result(await checkMeetingAvailability(request([{ ...tool, arguments: { ...args, attendees: 'Sneha Rao' } }]), context));
  assert.equal(ambiguous.status, 'needs_attendee_clarification');
  assert.deepEqual(ambiguous.attendees[0].candidates.map(person => [person.displayName, person.department]), [
    ['Sneha Rao', 'Finance'], ['Sneha Rao', 'HR'],
  ]);
  assert.equal(calendarCalls, 0);
  const clarified = result(await checkMeetingAvailability(request([{ ...tool, arguments: { ...args, attendees: hr.email } }]), context));
  assert.equal(clarified.status, 'available');
  assert.equal(calendarCalls, 1);
});

test('both endpoints complete Tool 1 then confirmed Tool 2 with mocked Graph only', async () => {
  setup();
  const checked = result(await checkMeetingAvailability(request(), context));
  assert.equal(checked.meeting.title, 'Project review');
  assert.equal(checked.meeting.startTime, '10:00');
  assert.ok(checked.requestedSlot.confirmationToken);
  const bookingArgs = { title: checked.meeting.title, attendeeEmails: checked.meeting.attendees.map(u => u.email), callerEmail: user.email,
    date: checked.meeting.date, startTime: checked.meeting.startTime, timeZone: checked.meeting.timeZone,
    duration: String(checked.meeting.durationMinutes), confirmed: 'true', confirmationToken: checked.requestedSlot.confirmationToken };
  provider.getGraphAccessToken = async () => 'fake-token';
  const posts = [];
  global.fetch = async (url, options) => {
    if (options.method === 'GET') return Response.json({ value: [] });
    posts.push(JSON.parse(options.body));
    return Response.json({ id: 'event-test' }, { status: 201 });
  };
  const booked = result(await scheduleMeeting(request([{ id: 'book-1', name: 'schedule_meeting', arguments: bookingArgs }]), context));
  assert.equal(booked.status, 'scheduled');
  assert.equal(posts.length, 1);
  assert.equal(posts[0].subject, 'Project review');
  assert.equal(posts[0].end.dateTime, '2030-01-07T10:30:00');
  assert.equal(posts[0].attendees[0].emailAddress.address, attendee.email);
});

test('Tool 2 refuses unconfirmed input and an expired or altered booking reference', async () => {
  setup();
  const checked = result(await checkMeetingAvailability(request(), context));
  const bookingArgs = { title: args.title, attendeeEmails: [attendee.email], callerEmail: user.email, date: args.date, startTime: args.startTime,
    timeZone: args.timeZone, duration: '30', confirmed: 'true', confirmationToken: checked.requestedSlot.confirmationToken };
  global.fetch = async () => assert.fail('must not reach Graph');
  for (const changes of [{ confirmed: 'false' }, { title: 'Different meeting' }, { confirmationToken: 'fake' }]) {
    const outcome = result(await scheduleMeeting(request([{ id: 'book-1', name: 'schedule_meeting', arguments: { ...bookingArgs, ...changes } }]), context));
    assert.equal(outcome.status, 'confirmation_required');
  }
});

test('busy requested time gets no confirmation token while checked alternatives do', async () => {
  setup();
  calendar.checkCalendarAvailability = async () => ({ status: 'alternatives_found', requestedSlot: { available: false }, alternatives: [{ startDateTime: '2030-01-07T11:00:00Z', date: '2030-01-07', startTime: '11:00' }] });
  const outcome = result(await checkMeetingAvailability(request(), context));
  assert.equal(outcome.requestedSlot.confirmationToken, undefined);
  assert.ok(outcome.alternatives[0].confirmationToken);
});
