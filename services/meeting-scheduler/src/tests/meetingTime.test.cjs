const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const { parseMeetingInput } = require('../../dist/meeting-scheduler/src/lib/meetingTime.js');
const { zonedParts } = require('../../dist/meeting-scheduler/src/lib/calendarAvailability.js');
const savedZone = process.env.MEETING_TIME_ZONE;
afterEach(() => { if (savedZone === undefined) delete process.env.MEETING_TIME_ZONE; else process.env.MEETING_TIME_ZONE = savedZone; });
const now = Date.parse('2030-01-06T20:00:00Z'); // Monday Jan 7 in Kolkata, Sunday in UTC
const input = { title: 'Project review', attendees: 'Sneha, John', date: 'tomorrow', startTime: '3 PM', duration: '1 hour 30 minutes' };

test('interprets tomorrow from actual zoned date and calculates end time', () => {
  process.env.MEETING_TIME_ZONE = 'Asia/Kolkata';
  const result = parseMeetingInput(input, false, now);
  assert.equal(result.date, '2030-01-08');
  assert.equal(result.time, '15:00');
  assert.equal(result.durationMinutes, 90);
  assert.equal(result.startDateTime, '2030-01-08T09:30:00.000Z');
  assert.equal(result.endDateTime, '2030-01-08T11:00:00.000Z');
  assert.deepEqual(result.attendeeNames, ['Sneha', 'John']);
});
test('upcoming weekday is strictly in the future, including next Monday on Monday', () => {
  for (const date of ['Monday', 'next Monday', 'this Monday', 'coming Monday']) {
    assert.equal(parseMeetingInput({ ...input, date, timeZone: 'Asia/Kolkata' }, false, now).date, '2030-01-14');
  }
  assert.equal(parseMeetingInput({ ...input, date: 'next Tuesday', timeZone: 'Asia/Kolkata' }, false, now).date, '2030-01-08');
});
test('accepts explicit zone, noon, midnight, decimal hours and minute strings', () => {
  assert.equal(parseMeetingInput({ ...input, startTime: '12 am', timeZone: 'UTC', duration: '30' }, false, now).time, '00:00');
  assert.equal(parseMeetingInput({ ...input, startTime: 'noon', duration: '1.5 hours' }, false, now).durationMinutes, 90);
  assert.equal(parseMeetingInput({ ...input, duration: 'half an hour' }, false, now).durationMinutes, 30);
});
test('handles month and year boundaries and cross-midnight meeting end', () => {
  const result = parseMeetingInput({ ...input, startTime: '23:30', duration: '1 hour', timeZone: 'Asia/Kolkata' }, false, Date.parse('2030-12-30T20:00:00Z'));
  assert.equal(result.date, '2031-01-01');
  assert.deepEqual(zonedParts(Date.parse(result.endDateTime), result.timeZone), { date: '2031-01-02', time: '00:30' });
});
test('rejects malformed or ambiguous dates and times, past dates, invalid strings and unknown fields', () => {
  for (const changes of [
    { date: '01/02/2030' }, { date: '2030-02-30' }, { date: '2029-12-01' }, { startTime: '3' }, { startTime: '25:00' },
    { duration: '0' }, { duration: '31 minutes' }, { duration: 30 }, { title: '' }, { attendees: [] }, { attendeePhones: '+12345678' },
    { timeZone: 'bad' }, { timeZone: '' }, { date: 'whenever' },
  ]) assert.throws(() => parseMeetingInput({ ...input, ...changes }, false, now), { code: 'INVALID_TIME' });
});
test('DST nonexistent and repeated local times ask for clarification, normal times convert', () => {
  for (const [date, startTime] of [['2030-03-10', '02:30'], ['2030-11-03', '01:30']]) {
    assert.throws(() => parseMeetingInput({ ...input, date, startTime, timeZone: 'America/New_York' }, false, now), { code: 'INVALID_TIME' });
  }
  assert.equal(parseMeetingInput({ ...input, date: '2030-03-10', startTime: '03:30', timeZone: 'America/New_York' }, false, now).startDateTime, '2030-03-10T07:30:00.000Z');
});
test('booking uses confirmed emails and absolute date to prevent relative-date drift on retries', () => {
  const booking = { title: input.title, attendeeEmails: ['sneha@example.com'], callerEmail: 'caller@example.com', date: '2030-01-08', startTime: '15:00', duration: '30', confirmed: 'true', confirmationToken: 'token' };
  assert.equal(parseMeetingInput(booking, true, now).attendeeNames[0], 'sneha@example.com');
  assert.throws(() => parseMeetingInput({ ...booking, date: 'tomorrow' }, true, now), { code: 'INVALID_TIME' });
  assert.throws(() => parseMeetingInput({ ...booking, attendeeEmails: 'Sneha' }, true, now), { code: 'INVALID_TIME' });
});


test('default Eastern timezone uses EST in winter and EDT in summer', () => {
  delete process.env.MEETING_TIME_ZONE;
  for (const [date, expected] of [['2030-01-08', '2030-01-08T20:00:00.000Z'], ['2030-07-08', '2030-07-08T19:00:00.000Z']]) {
    const result = parseMeetingInput({ ...input, date, startTime: '15:00' }, false, now);
    assert.equal(result.timeZone, 'America/New_York');
    assert.equal(result.startDateTime, expected);
  }
});


test('optional description is trimmed, bounded and allows multiline text', () => {
  assert.equal(parseMeetingInput({...input,description:'  Agenda\nNext steps  '},false,now).description,'Agenda\nNext steps');
  assert.equal(parseMeetingInput({...input,description:'  '},false,now).description,undefined);
  for(const description of [42,null,'x'.repeat(5001),String.fromCharCode(0)]) assert.throws(()=>parseMeetingInput({...input,description},false,now),{code:'INVALID_TIME'});
});


test('optional location is trimmed and validated', () => {
  assert.equal(parseMeetingInput({...input,location:' Room A '},false,now).location,'Room A');
  assert.equal(parseMeetingInput({...input,location:' '},false,now).location,undefined);
  for(const location of [42,null,'x'.repeat(257),String.fromCharCode(10)]) assert.throws(()=>parseMeetingInput({...input,location},false,now),{code:'INVALID_TIME'});
});
