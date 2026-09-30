const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const provider = require('../../dist/Knowledge Base/src/lib/graphAuth.js');
const { checkCalendarAvailability, validateRequestedTime } = require('../../dist/meeting-scheduler/src/lib/calendarAvailability.js');
const originalFetch = global.fetch;
const originalToken = provider.getGraphAccessToken;
const originalNow = Date.now;
afterEach(() => { global.fetch = originalFetch; provider.getGraphAccessToken = originalToken; Date.now = originalNow; });

const requested = '2030-01-07T10:00:00+00:00'; // Monday
const start = Date.parse(requested);
const caller = { id: 'caller', displayName: 'Caller', email: 'caller@example.com' };
const attendee = { id: 'attendee', displayName: 'Attendee', email: 'attendee@example.com' };
const meeting = { caller, attendees: [attendee], date: '2030-01-07', time: '10:00', timeZone: 'UTC', durationMinutes: 30 };
const item = (from, to, status = 'busy') => ({ status,
  start: { dateTime: from, timeZone: 'UTC' }, end: { dateTime: to, timeZone: 'UTC' } });
const schedule = (email, items = []) => ({ scheduleId: email, availabilityView: '0'.repeat(2016), scheduleItems: items });

function setup(values = [schedule(caller.email), schedule(attendee.email)]) {
  Date.now = () => Date.parse('2030-01-01T00:00:00Z');
  provider.getGraphAccessToken = async () => 'test-token';
  const requests = [];
  global.fetch = async (url, options) => {
    requests.push({ url, options });
    return values instanceof Response ? values : Response.json({ value: values });
  };
  return requests;
}

test('checks caller and attendee via getSchedule without creating events', async () => {
  const requests = setup();
  const result = await checkCalendarAvailability(meeting, requested);
  assert.equal(result.status, 'available');
  assert.equal(result.requestedSlot.available, true);
  assert.deepEqual(result.alternatives, []);
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, 'https://graph.microsoft.com/v1.0/users/caller/calendar/getSchedule');
  const body = JSON.parse(requests[0].options.body);
  assert.deepEqual(body.schedules, [caller.email, attendee.email]);
  assert.deepEqual(body.startTime, { dateTime: '2030-01-07T10:00:00', timeZone: 'UTC' });
  assert.equal(body.endTime.dateTime, '2030-01-14T10:00:00');
  assert.equal(body.availabilityViewInterval, 5);
  assert.equal(requests[0].options.redirect, 'error');
});

test('alternatives avoid overlapping conflicts for every participant for the full duration', async () => {
  setup([
    schedule(caller.email, [item('2030-01-07T10:00:00', '2030-01-07T11:00:00')]),
    schedule(attendee.email, [item('2030-01-07T10:45:00', '2030-01-07T11:30:00')]),
  ]);
  const result = await checkCalendarAvailability(meeting, requested);
  assert.equal(result.status, 'alternatives_found');
  assert.equal(result.requestedSlot.available, false);
  assert.equal(result.alternatives.length, 3);
  assert.equal(result.alternatives[0].startDateTime, '2030-01-07T11:30:00.000Z');
  assert.equal(result.alternatives[1].startDateTime, '2030-01-07T12:00:00.000Z');
});

test('half-open overlap boundaries allow adjacent meetings', async () => {
  setup([schedule(caller.email, [item('2030-01-07T09:00:00', '2030-01-07T10:00:00'), item('2030-01-07T10:30:00', '2030-01-07T11:00:00')]), schedule(attendee.email)]);
  assert.equal((await checkCalendarAvailability(meeting, requested)).status, 'available');
});

test('tentative and out-of-office block; free and working elsewhere do not', async () => {
  for (const status of ['tentative', 'oof', 'free', 'workingElsewhere']) {
    setup([schedule(caller.email, [item('2030-01-07T10:00:00', '2030-01-07T11:00:00', status)]), schedule(attendee.email)]);
    assert.equal((await checkCalendarAvailability(meeting, requested)).requestedSlot.available, ['free', 'workingElsewhere'].includes(status));
  }
});

test('coarse busy view still blocks when calendar details are hidden', async () => {
  const hidden = schedule(attendee.email);
  hidden.availabilityView = '2'.repeat(6) + '0'.repeat(2010);
  setup([schedule(caller.email), hidden]);
  const result = await checkCalendarAvailability(meeting, requested);
  assert.equal(result.requestedSlot.available, false);
  assert.equal(result.alternatives[0].startDateTime, '2030-01-07T10:30:00.000Z');
});

test('returns no common slot when one participant is busy throughout the window', async () => {
  setup([schedule(caller.email), schedule(attendee.email, [item('2030-01-07T00:00:00', '2030-01-15T00:00:00')])]);
  const result = await checkCalendarAvailability(meeting, requested);
  assert.equal(result.status, 'no_common_slot');
  assert.deepEqual(result.alternatives, []);
});

test('alternatives skip evenings and weekends', async () => {
  setup([schedule(caller.email, [item('2030-01-11T16:30:00', '2030-01-11T17:00:00')]), schedule(attendee.email)]);
  const result = await checkCalendarAvailability(meeting, '2030-01-11T16:30:00Z');
  assert.equal(result.alternatives[0].startDateTime, '2030-01-14T09:00:00.000Z');
});

test('alternative working hours use requested timezone, not host timezone', async () => {
  setup([schedule(caller.email, [item('2030-01-07T03:30:00', '2030-01-07T04:30:00')]), schedule(attendee.email)]);
  const result = await checkCalendarAvailability({ ...meeting, timeZone: 'Asia/Kolkata' }, '2030-01-07T09:00:00+05:30');
  assert.equal(result.alternatives[0].startDateTime, '2030-01-07T04:30:00.000Z');
  assert.equal(result.alternatives[0].startTime, '10:00');
  assert.equal(validateRequestedTime(result.alternatives[0].startDateTime, 'Asia/Kolkata', 30), Date.parse(result.alternatives[0].startDateTime));
});

test('validates date, timezone, future time, duration and DST offset', () => {
  const now = Date.parse('2030-01-01T00:00:00Z');
  for (const args of [
    ['2030-02-30T10:00:00Z', 'UTC', 30], ['2029-01-01T10:00:00Z', 'UTC', 30],
    [requested, 'bad/timezone', 30], [requested, 'Asia/Kolkata', 30],
    ['2030-01-07T10:00:00', 'UTC', 30], [requested, 'UTC', 0], [requested, 'UTC', 31],
    [requested, 'UTC', 481], ['2030-03-10T02:30:00-05:00', 'America/New_York', 30],
  ]) assert.throws(() => validateRequestedTime(...args, now), { code: 'INVALID_TIME' });
  assert.equal(validateRequestedTime('2030-11-03T01:30:00-04:00', 'America/New_York', 30, now), Date.parse('2030-11-03T05:30:00Z'));
  assert.equal(validateRequestedTime('2030-11-03T01:30:00-05:00', 'America/New_York', 30, now), Date.parse('2030-11-03T06:30:00Z'));
});

test('missing, duplicate, malformed and per-mailbox error data never means free', async () => {
  for (const values of [
    [schedule(caller.email)], [schedule(caller.email), schedule(caller.email)],
    [schedule(caller.email), { ...schedule(attendee.email), error: { code: '5006' } }],
    [schedule(caller.email), { ...schedule(attendee.email), availabilityView: '' }],
    [schedule(caller.email), { ...schedule(attendee.email), availabilityView: '4'.repeat(2016) }],
    [schedule(caller.email), { ...schedule(attendee.email), scheduleItems: null }],
    [schedule(caller.email), schedule(attendee.email, [item('bad', 'bad')])],
    [schedule(caller.email), schedule(attendee.email, [item('2030-01-07T10:00:00', '2030-01-07T11:00:00', 'unknown')])],
  ]) {
    setup(values);
    await assert.rejects(checkCalendarAvailability(meeting, requested), { code: 'CALENDAR_UNAVAILABLE' });
  }
});

test('throttling, denied access, malformed JSON, network and authentication fail safely', async () => {
  setup(new Response('sensitive response', { status: 429, headers: { 'Retry-After': '25' } }));
  await assert.rejects(checkCalendarAvailability(meeting, requested), { code: 'THROTTLED', retryAfterSeconds: 25 });
  setup(new Response('sensitive response', { status: 403 }));
  await assert.rejects(checkCalendarAvailability(meeting, requested), error => error.code === 'GRAPH_ERROR' && !error.message.includes('sensitive'));
  setup(new Response('not JSON'));
  await assert.rejects(checkCalendarAvailability(meeting, requested), { code: 'CALENDAR_UNAVAILABLE' });
  setup(); global.fetch = async () => { throw new Error('secret'); };
  await assert.rejects(checkCalendarAvailability(meeting, requested), error => error.code === 'GRAPH_ERROR' && !error.message.includes('secret'));
  setup(); provider.getGraphAccessToken = async () => null;
  await assert.rejects(checkCalendarAvailability(meeting, requested), { code: 'AUTH_FAILED' });
});
