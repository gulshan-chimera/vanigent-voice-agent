const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const directory = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');
const { prepareMeetingDetails } = require('../../dist/meeting-scheduler/src/lib/meetingDetails.js');
const originalSearch = directory.searchEntraUsers;
const originalFind = directory.findAttendeeCandidates;
afterEach(() => { directory.searchEntraUsers = originalSearch; directory.findAttendeeCandidates = originalFind; });

const caller = {
  id: 'caller', displayName: 'Alex Green', givenName: 'Alex', surname: 'Green',
  email: 'alex@company.com', mail: 'alex@company.com', userPrincipalName: 'alex@company.com',
  department: null, jobTitle: null, accountEnabled: true,
};
const attendee = { ...caller, id: 'attendee', displayName: 'Sneha Rao', givenName: 'Sneha', email: 'sneha@company.com' };
const input = { attendeeNames: ['Sneha'], date: 'tomorrow', time: '3 PM', timeZone: 'Asia/Kolkata', durationMinutes: 30 };

test('even a single fuzzy candidate requires explicit selection', async () => {
  directory.findAttendeeCandidates = async () => ({ users: [attendee], exact: false });
  const result = await prepareMeetingDetails(caller, input, true);
  assert.equal(result.status, 'needs_attendee_clarification');
  assert.equal(result.attendees[0].candidates.length, 1);
});

test('returns missing fields without calling Graph', async () => {
  directory.searchEntraUsers = async () => assert.fail('must not query Graph');
  assert.deepEqual(await prepareMeetingDetails(null, {}), {
    status: 'needs_details', fields: ['caller', 'attendeeNames', 'date', 'time', 'timeZone', 'durationMinutes'],
  });
});

test('collects caller, resolved attendees, email, date, time, timezone and duration', async () => {
  directory.searchEntraUsers = async name => { assert.equal(name, 'Sneha'); return [attendee]; };
  const result = await prepareMeetingDetails(caller, input);
  assert.equal(result.status, 'ready_for_confirmation');
  assert.deepEqual(result.meeting, { caller, attendees: [attendee], date: 'tomorrow', time: '3 PM', timeZone: 'Asia/Kolkata', durationMinutes: 30 });
});

test('requires clarification for multiple candidates, partial matches, and no matches', async () => {
  for (const candidates of [[attendee, { ...attendee, id: 'other' }], [{ ...attendee, givenName: 'Snehali', displayName: 'Snehali Rao' }], []]) {
    directory.searchEntraUsers = async () => candidates;
    assert.deepEqual(await prepareMeetingDetails(caller, input), {
      status: 'needs_attendee_clarification', attendees: [{ name: 'Sneha', candidates }],
    });
  }
});

test('requires timezone, valid duration and complete attendee names', async () => {
  directory.searchEntraUsers = async () => assert.fail('must not query Graph');
  for (const durationMinutes of [0, -10, 1.5, 1441, '30', NaN]) {
    const result = await prepareMeetingDetails(caller, { ...input, durationMinutes });
    assert.ok(result.fields.includes('durationMinutes'));
  }
  assert.deepEqual((await prepareMeetingDetails(caller, { ...input, timeZone: ' ' })).fields, ['timeZone']);
  assert.deepEqual((await prepareMeetingDetails(caller, { ...input, attendeeNames: ['Sneha', ''] })).fields, ['attendeeNames']);
});

test('deduplicates names and user IDs and omits caller from attendees', async () => {
  let searches = 0;
  directory.searchEntraUsers = async name => { searches++; return [name === 'Alex' ? caller : attendee]; };
  const result = await prepareMeetingDetails(caller, { ...input, attendeeNames: ['Sneha', 'sneha', 'Sneha Rao', 'Alex'] });
  assert.equal(searches, 3);
  assert.deepEqual(result.meeting.attendees, [attendee]);
});

test('rejects disabled caller or missing caller email', async () => {
  directory.searchEntraUsers = async () => assert.fail('must not query Graph');
  for (const invalidCaller of [{ ...caller, accountEnabled: false }, { ...caller, email: '' }]) {
    assert.deepEqual((await prepareMeetingDetails(invalidCaller, input)).fields, ['caller']);
  }
});

test('preserves Graph failures instead of treating them as missing attendees', async () => {
  const failure = new directory.EntraUserSearchError('THROTTLED', 'Retry later.', 429, 30);
  directory.searchEntraUsers = async () => { throw failure; };
  await assert.rejects(prepareMeetingDetails(caller, input), error => error === failure);
});
