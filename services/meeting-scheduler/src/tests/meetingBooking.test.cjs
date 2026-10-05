const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const provider = require('../../dist/meeting-scheduler/src/lib/graphAuth.js');
const calendar = require('../../dist/meeting-scheduler/src/lib/calendarAvailability.js');
const { issueConfirmation, createConfirmedMeeting } = require('../../dist/meeting-scheduler/src/lib/meetingBooking.js');
const saved = { fetch: global.fetch, token: provider.getGraphAccessToken, check: calendar.checkCalendarAvailability, now: Date.now, secret: process.env.MEETING_SCHEDULER_SECRET };
afterEach(() => {
  global.fetch = saved.fetch; provider.getGraphAccessToken = saved.token; calendar.checkCalendarAvailability = saved.check; Date.now = saved.now;
  if (saved.secret === undefined) delete process.env.MEETING_SCHEDULER_SECRET; else process.env.MEETING_SCHEDULER_SECRET = saved.secret;
});
const caller = { id: 'caller', email: 'caller@example.com', displayName: 'Caller' };
const attendee = { id: 'attendee', email: 'sneha@example.com', displayName: 'Sneha Rao' };
const meeting = { title: 'Review', caller, attendees: [attendee], date: '2030-01-08', time: '15:00', timeZone: 'Asia/Kolkata', durationMinutes: 30 };
const start = '2030-01-08T09:30:00.000Z';
function setup(pages = [{ value: [] }, { id: 'event-1', subject: 'Review', webLink: 'https://outlook.office.com/event/1' }]) {
  Date.now = () => Date.parse('2030-01-07T00:00:00Z');
  process.env.MEETING_SCHEDULER_SECRET = 'test-secret';
  provider.getGraphAccessToken = async () => 'test-token';
  calendar.checkCalendarAvailability = async () => ({ requestedSlot: { available: true } });
  const requests = [];
  global.fetch = async (url, options) => {
    requests.push({ url, options });
    assert.ok(pages.length, 'unexpected Graph request');
    const data = pages.shift();
    if (data instanceof Error) throw data;
    return data instanceof Response ? data : Response.json(data, { status: options.method === 'POST' ? 201 : 200 });
  };
  return { requests, token: issueConfirmation('call-1', meeting, start) };
}

test('creates organizer event with calculated end, attendees and deterministic transactionId', async () => {
  const { requests, token } = setup();
  const result = await createConfirmedMeeting('call-1', meeting, start, 'true', token);
  assert.equal(result.status, 'scheduled');
  assert.equal(result.meeting.endDateTime, '2030-01-08T10:00:00.000Z');
  const posted = JSON.parse(requests[1].options.body);
  assert.equal(requests[1].url, 'https://graph.microsoft.com/v1.0/users/caller/events');
  assert.equal(posted.subject, 'Review');
  assert.deepEqual(posted.start, { dateTime: '2030-01-08T09:30:00', timeZone: 'UTC' });
  assert.deepEqual(posted.end, { dateTime: '2030-01-08T10:00:00', timeZone: 'UTC' });
  assert.deepEqual(posted.attendees, [{ emailAddress: { address: attendee.email, name: attendee.displayName }, type: 'required' }]);
  assert.match(posted.transactionId, /^[a-f0-9-]{36}$/);
  assert.equal(requests[1].options.redirect, 'error');
});
test('confirmation is required before any Graph call', async () => {
  const { requests, token } = setup();
  for (const confirmed of [undefined, false, true, 'false', 'yes']) {
    await assert.rejects(createConfirmedMeeting('call-1', meeting, start, confirmed, token), { code: 'CONFIRMATION_REQUIRED' });
  }
  assert.equal(requests.length, 0);
});
test('signed confirmation binds title, organizer, attendees, time, duration and call', async () => {
  const { requests, token } = setup();
  for (const [callId, changed, time] of [
    ['call-2', meeting, start], ['call-1', { ...meeting, title: 'Other' }, start],
    ['call-1', { ...meeting, caller: { ...caller, id: 'other' } }, start],
    ['call-1', { ...meeting, attendees: [{ ...attendee, email: 'other@example.com' }] }, start],
    ['call-1', { ...meeting, durationMinutes: 60 }, start], ['call-1', meeting, '2030-01-08T10:30:00Z'],
  ]) await assert.rejects(createConfirmedMeeting(callId, changed, time, 'true', token), { code: 'CONFIRMATION_REQUIRED' });
  await assert.rejects(createConfirmedMeeting('call-1', meeting, start, 'true', token + 'x'), { code: 'CONFIRMATION_REQUIRED' });
  Date.now = () => Date.parse('2030-01-07T00:16:00Z');
  await assert.rejects(createConfirmedMeeting('call-1', meeting, start, 'true', token), { code: 'CONFIRMATION_REQUIRED' });
  assert.equal(requests.length, 0);
});
test('rechecks calendars and does not create when availability changed', async () => {
  const { token, requests } = setup([{ value: [] }]);
  calendar.checkCalendarAvailability = async () => ({ requestedSlot: { available: false }, alternatives: [] });
  assert.equal((await createConfirmedMeeting('call-1', meeting, start, 'true', token)).status, 'availability_changed');
  assert.equal(requests.length, 1);
});
test('retry finds prior event before availability check and sends no duplicate invitation', async () => {
  const first = setup();
  await createConfirmedMeeting('call-1', meeting, start, 'true', first.token);
  const transactionId = JSON.parse(first.requests[1].options.body).transactionId;
  const retry = setup([{ value: [{ id: 'event-1', transactionId, subject: 'Review' }] }]);
  calendar.checkCalendarAvailability = async () => assert.fail('existing event must be returned first');
  const result = await createConfirmedMeeting('call-1', meeting, start, 'true', retry.token);
  assert.equal(result.status, 'already_scheduled');
  assert.equal(retry.requests.length, 1);
});
test('duplicate transactionId stays stable even if Tool 1 issues a fresh confirmation', async () => {
  const first = setup();
  await createConfirmedMeeting('call-1', meeting, start, 'true', first.token);
  const transactionId = JSON.parse(first.requests[1].options.body).transactionId;
  const next = setup();
  Date.now = () => Date.parse('2030-01-07T00:05:00Z');
  const newToken = issueConfirmation('call-1', meeting, start);
  assert.notEqual(newToken, first.token);
  await createConfirmedMeeting('call-1', meeting, start, 'true', newToken);
  assert.equal(JSON.parse(next.requests[1].options.body).transactionId, transactionId);
});
test('POST timeout or server error is an unknown outcome, not a confirmed failure', async () => {
  for (const failure of [new Error('private-token'), new Response('sensitive', { status: 503 }), new Response('invalid JSON', { status: 201 }), {}]) {
    const { token } = setup([{ value: [] }, failure]);
    await assert.rejects(createConfirmedMeeting('call-1', meeting, start, 'true', token), error => error.code === 'BOOKING_UNKNOWN' && !error.message.includes('private-token') && !error.message.includes('sensitive'));
  }
});
test('permission or duplicate-check errors stop creation', async () => {
  for (const response of [new Response('private-data', { status: 403 }), { value: null }, { value: [], '@odata.nextLink': 'https://evil.example/calendar' }]) {
    const { requests, token } = setup([response]);
    await assert.rejects(createConfirmedMeeting('call-1', meeting, start, 'true', token), { code: 'BOOKING_FAILED' });
    assert.equal(requests.length, 1);
  }
});
test('duplicate check follows Graph pagination without changing organizer or leaking tokens', async () => {
  const { token, requests } = setup([{ value: [], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/caller/calendarView?$skiptoken=next' }, { value: [] }, { id: 'event-1' }]);
  assert.equal((await createConfirmedMeeting('call-1', meeting, start, 'true', token)).status, 'scheduled');
  assert.equal(requests.length, 3);
});


test('requests Teams with HTML body and omits location when absent', async () => {
  const {requests,token}=setup([{value:[]},{id:'event-1',onlineMeeting:{joinUrl:'https://teams.microsoft.com/l/meetup-join/test'}}]);
  const result=await createConfirmedMeeting('call-1',meeting,start,'true',token);
  const body=JSON.parse(requests[1].options.body);
  assert.equal(body.isOnlineMeeting,true);
  assert.equal(body.location,undefined);
  assert.equal(body.onlineMeetingProvider,'teamsForBusiness');
  assert.equal(body.body.contentType,'HTML');assert.ok(body.body.content.includes('<body>'));
  assert.equal(result.event.joinUrl,'https://teams.microsoft.com/l/meetup-join/test');
  assert.equal(result.event.meetingLinkStatus,'available');
});
test('description is escaped HTML and bound to confirmation, including attendee-only tokens', async () => {
  const {issueAvailabilityConfirmation}=require('../../dist/meeting-scheduler/src/lib/meetingBooking.js');
  const {requests}=setup();
  const described={...meeting,description:'Review <draft> and discuss next steps.'};
  const token=issueAvailabilityConfirmation(described,start);
  await assert.rejects(createConfirmedMeeting('call-1',{...described,description:'Changed'},start,'true',token),{code:'CONFIRMATION_REQUIRED'});
  assert.equal(requests.length,0);
  const result=await createConfirmedMeeting('call-1',described,start,'true',token);
  assert.equal(JSON.parse(requests[1].options.body).body.contentType,'HTML');assert.ok(JSON.parse(requests[1].options.body).body.content.includes('Review &lt;draft&gt; and discuss next steps.'));
  assert.equal(result.meeting.description,described.description);
  assert.equal(result.event.joinUrl,null);
  assert.equal(result.event.meetingLinkStatus,'unavailable');assert.ok(result.event.meetingLinkMessage);
  assert.equal(result.status,'scheduled');
});
test('retry returns existing Teams link without creating another invitation', async () => {
  const first=setup();await createConfirmedMeeting('call-1',meeting,start,'true',first.token);
  const transactionId=JSON.parse(first.requests[1].options.body).transactionId;
  const retry=setup([{value:[{id:'event-1',transactionId,onlineMeeting:{joinUrl:'https://teams.microsoft.com/l/meetup-join/test'}}]}]);
  const result=await createConfirmedMeeting('call-1',meeting,start,'true',retry.token);
  assert.equal(result.status,'already_scheduled');assert.equal(result.event.meetingLinkStatus,'available');
  assert.equal(retry.requests.length,1);assert.ok(retry.requests[0].url.includes('onlineMeeting'));
});


test('location is included in invitation and protected by the selected-slot confirmation', async () => {
  const {issueAvailabilityConfirmation}=require('../../dist/meeting-scheduler/src/lib/meetingBooking.js');
  const {requests}=setup();const located={...meeting,location:'Conference Room A'};
  const token=issueAvailabilityConfirmation(located,start);
  await assert.rejects(createConfirmedMeeting('call-1',{...located,location:'Room B'},start,'true',token),{code:'CONFIRMATION_REQUIRED'});
  assert.equal(requests.length,0);
  const result=await createConfirmedMeeting('call-1',located,start,'true',token);
  assert.deepEqual(JSON.parse(requests[1].options.body).location,{displayName:'Conference Room A'});
  assert.equal(result.meeting.location,'Conference Room A');
});


test('invitation HTML escapes caller text and preserves line breaks without invented Teams credentials', () => {
  const {invitationBody}=require('../../dist/meeting-scheduler/src/lib/meetingInvitation.js');
  const body=invitationBody('<script>alert("x")</script>&\nSecond line');
  assert.equal(body.contentType,'HTML');
  assert.ok(body.content.includes('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;&amp;<br>Second line'));
  assert.ok(!body.content.includes('<script>'));
  assert.ok(!body.content.includes('Passcode:'));
  assert.equal(invitationBody().contentType,'HTML');
});


test('agenda uses readable font, a bold supplied heading and spaced numbered topics', () => {
  const {invitationBody}=require('../../dist/meeting-scheduler/src/lib/meetingInvitation.js');
  const html=invitationBody('Weekly Meeting Agenda\n\n1. Sprint Progress Review ? Review sprint targets.\n\n2. Feature Demo & Feedback ? Gather feedback.').content;
  assert.ok(html.includes('font-size:12pt'));
  assert.ok(html.includes('<strong>Weekly Meeting Agenda</strong>'));
  assert.ok(html.includes('<ol start="1"'));
  assert.equal((html.match(/<li /g)||[]).length,2);
  assert.ok(html.includes('<strong>Sprint Progress Review</strong> ? Review sprint targets.'));
  assert.ok(html.includes('<strong>Feature Demo &amp; Feedback</strong>'));
});
test('plain description stays plain and list labels are escaped', () => {
  const {invitationBody}=require('../../dist/meeting-scheduler/src/lib/meetingInvitation.js');
  assert.ok(!invitationBody('Discuss project progress.').content.includes('<ol'));
  assert.ok(!invitationBody().content.includes('<p'));
  const html=invitationBody('3. <img src=x> - Description').content;
  assert.ok(html.includes('value="3"'));assert.ok(html.includes('<strong>&lt;img src=x&gt;</strong>'));
  assert.ok(!html.includes('<img'));
});
