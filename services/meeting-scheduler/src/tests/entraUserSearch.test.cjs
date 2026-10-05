const { test, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const provider = require('../../dist/meeting-scheduler/src/lib/graphAuth.js');
const { searchEntraUsers, EntraUserSearchError } = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');

const originalFetch = global.fetch;
const originalTokenProvider = provider.getGraphAccessToken;
const { lookupMeetingCaller } = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');
const { findAttendeeCandidates, rankAttendeeCandidates } = require('../../dist/meeting-scheduler/src/lib/entraUserSearch.js');
afterEach(() => {
  global.fetch = originalFetch;
  provider.getGraphAccessToken = originalTokenProvider;
});

const user = (changes = {}) => ({
  id: 'user-1', displayName: 'Sneha Rao', givenName: 'Sneha', surname: 'Rao',
  mail: 'sneha.rao@company.com', userPrincipalName: 'sneha@company.com',
  department: 'Finance', jobTitle: 'Manager', accountEnabled: true, ...changes,
});

function setup(pages = [{ value: [user()] }]) {
  const requests = [];
  provider.getGraphAccessToken = async () => 'secret-token-marker';
  global.fetch = async (url, options) => {
    requests.push({ url: new URL(url), options });
    assert.ok(pages.length, 'unexpected additional Graph request');
    const page = pages.shift();
    return page instanceof Response ? page : Response.json(page);
  };
  return requests;
}

test('exact names resolve before a fuzzy directory scan', async () => {
  const requests = setup();
  const result = await findAttendeeCandidates('Sneha');
  assert.equal(result.exact, true);
  assert.equal(requests.length, 1);
  assert.ok(requests[0].url.searchParams.get('$filter').includes("givenName eq 'Sneha'"));
});
test('duplicate exact names return both departments without running fuzzy search', async () => {
  const requests = setup([{ value: [
    user({ id: 'finance-user', department: 'Finance' }),
    user({ id: 'hr-user', department: 'HR', mail: 'sneha.hr@company.com' }),
  ] }]);
  const result = await findAttendeeCandidates('Sneha Rao');
  assert.equal(result.exact, true);
  assert.deepEqual(result.users.map(person => [person.displayName, person.department]), [
    ['Sneha Rao', 'Finance'], ['Sneha Rao', 'HR'],
  ]);
  assert.equal(requests.length, 1, 'exact duplicates must not trigger fuzzy search');
});
test('fuzzy fallback finds Neha/Sneha and Kai/Cai but marks them as suggestions', async () => {
  const requests = setup([{ value: [] }, { value: [] }, { value: [user()] }]);
  const result = await findAttendeeCandidates('Neha');
  assert.equal(result.exact, false);
  assert.equal(result.users[0].displayName, 'Sneha Rao');
  assert.equal(requests[2].url.searchParams.get('$top'), '200');
  assert.equal(rankAttendeeCandidates('Kai', [user({ displayName: 'Cai Smith', givenName: 'Cai' })]).length, 1);
  assert.equal(rankAttendeeCandidates('Nobody', [user()]).length, 0);
  assert.equal(rankAttendeeCandidates('X', [user()]).length, 0);
});
test('fuzzy matches support transpositions and preserve candidate profile details', () => {
  const result = rankAttendeeCandidates('Sneh a', [user(), user({ id: 'unrelated', givenName: 'Robert', displayName: 'Robert Jones', surname: 'Jones' })]);
  assert.equal(result[0].jobTitle, 'Manager');
  assert.equal(result[0].department, 'Finance');
  assert.equal(rankAttendeeCandidates('Snh ea', [user()]).length, 0);
  assert.equal(rankAttendeeCandidates('Sneh a', [user()]).length, 1);
});

test('email clarification uses exact directory email or UPN filters', async () => {
  const requests = setup();
  await searchEntraUsers('sneha.rao@company.com');
  assert.equal(requests[0].url.searchParams.get('$filter'), "accountEnabled eq true and (mail eq 'sneha.rao@company.com' or userPrincipalName eq 'sneha.rao@company.com')");
});

test('caller requires unique enabled email identity and retains country code', async () => {
  const requests = setup([{ value: [user()], '@odata.count': 1 }]);
  assert.equal((await lookupMeetingCaller('+14155551234')).id, 'user-1');
  const filter = requests[0].url.searchParams.get('$filter');
  assert.ok(filter.startsWith('accountEnabled eq true'));
  assert.ok(!filter.includes("mobilePhone eq '4155551234'"));
  for (const page of [{ value: [user()], '@odata.count': 2 }, { value: [user()] }, { value: [user({ accountEnabled: false })], '@odata.count': 1 }]) {
    setup([page]); assert.equal(await lookupMeetingCaller('+14155551234'), null);
  }
  setup(); assert.equal(await lookupMeetingCaller('4155551234'), null);
});

test('first-name prefix query selects only requested fields and uses reused auth', async () => {
  const requests = setup();
  const result = await searchEntraUsers(' Sneha ');
  assert.equal(result[0].displayName, 'Sneha Rao');
  const { url, options } = requests[0];
  assert.equal(url.origin + url.pathname, 'https://graph.microsoft.com/v1.0/users');
  assert.equal(url.searchParams.get('$filter'), "accountEnabled eq true and (startswith(displayName,'Sneha') or startswith(givenName,'Sneha') or startswith(surname,'Sneha'))");
  assert.equal(url.searchParams.get('$select'), 'id,displayName,givenName,surname,mail,userPrincipalName,department,jobTitle,accountEnabled');
  assert.equal(url.searchParams.get('$count'), 'true');
  assert.equal(url.searchParams.get('$top'), '25');
  assert.equal(url.searchParams.get('$orderby'), 'displayName');
  assert.equal(options.headers.Authorization, 'Bearer secret-token-marker');
  assert.equal(options.headers.ConsistencyLevel, 'eventual');
  assert.equal(options.redirect, 'error');
  assert.ok(options.signal instanceof AbortSignal);
});

test('full-name prefix preserves the full normalized name', async () => {
  const requests = setup();
  const result = await searchEntraUsers(' Sneha   Rao ');
  assert.equal(result.length, 1);
  assert.ok(requests[0].url.searchParams.get('$filter').includes("startswith(displayName,'Sneha Rao')"));
});

test('apostrophes and query metacharacters remain inside escaped OData literals', async () => {
  const requests = setup([{ value: [] }]);
  await searchEntraUsers("O'Neil & Sons");
  assert.ok(requests[0].url.searchParams.get('$filter').includes("'O''Neil & Sons'"));
  assert.equal([...requests[0].url.searchParams].length, 5);
});

test('returns multiple candidates without choosing one', async () => {
  setup([{ value: [user(), user({ id: 'user-2', displayName: 'Sneha Patel' })] }]);
  assert.deepEqual((await searchEntraUsers('Sneha')).map(u => u.id), ['user-1', 'user-2']);
});

test('excludes disabled and unknown account states even if Graph returns them', async () => {
  setup([{ value: [user({ accountEnabled: false }), user({ accountEnabled: null }), user({ accountEnabled: undefined }), user({ id: 'enabled' })] }]);
  assert.deepEqual((await searchEntraUsers('Sneha')).map(u => u.id), ['enabled']);
});

test('prefers trimmed mail, keeps original identity fields, and normalizes optional fields', async () => {
  setup([{ value: [user({ mail: ' sneha.rao@company.com ', department: undefined })] }]);
  const [result] = await searchEntraUsers('Sneha');
  assert.equal(result.email, 'sneha.rao@company.com');
  assert.equal(result.mail, 'sneha.rao@company.com');
  assert.equal(result.userPrincipalName, 'sneha@company.com');
  assert.equal(result.department, null);
});

test('falls back to usable UPN for missing, blank, or malformed mail', async () => {
  setup([{ value: [null, '', 'not-an-email'].map((mail, i) => user({ id: String(i), mail })) }]);
  const results = await searchEntraUsers('Sneha');
  assert.equal(results.length, 3);
  assert.ok(results.every(u => u.email === 'sneha@company.com'));
});

test('skips missing email identities and guest placeholder UPNs; accepts real guest mail', async () => {
  setup([{ value: [
    user({ mail: null, userPrincipalName: null }),
    user({ mail: '', userPrincipalName: 'guest_company.com#EXT#@tenant.onmicrosoft.com' }),
    user({ mail: 'bad address@company.com', userPrincipalName: 'bad' }),
    user({ id: 'guest-with-mail', userPrincipalName: 'guest#EXT#@tenant.onmicrosoft.com' }),
  ] }]);
  assert.deepEqual((await searchEntraUsers('Sneha')).map(u => u.id), ['guest-with-mail']);
});

test('rejects empty, non-string, excessive, and control-character input before authentication', async () => {
  provider.getGraphAccessToken = async () => assert.fail('must not acquire a token');
  for (const input of ['', '  ', null, 123, 'x'.repeat(257), 'Sneha\nRao']) {
    await assert.rejects(searchEntraUsers(input), { code: 'INVALID_INPUT' });
  }
});

test('returns an empty array for no matches or no usable candidates', async () => {
  setup([{ value: [] }, { value: [user({ mail: null, userPrincipalName: null })] }]);
  assert.deepEqual(await searchEntraUsers('Nobody'), []);
  assert.deepEqual(await searchEntraUsers('Sneha'), []);
});

test('null and thrown auth failures are sanitized and never call Graph', async () => {
  global.fetch = async () => assert.fail('must not query Graph');
  provider.getGraphAccessToken = async () => null;
  await assert.rejects(searchEntraUsers('Sneha'), { code: 'AUTH_FAILED' });
  provider.getGraphAccessToken = async () => { throw new Error('secret-token-marker'); };
  await assert.rejects(searchEntraUsers('Sneha'), error => {
    assert.equal(error.code, 'AUTH_FAILED');
    assert.ok(!String(error).includes('secret-token-marker'));
    assert.equal(error.cause, undefined);
    return true;
  });
});

test('401, 403, and server failures expose status without raw response or logs', async t => {
  const logs = [];
  for (const method of ['log', 'warn', 'error']) t.mock.method(console, method, (...args) => logs.push(args));
  for (const status of [401, 403, 500, 503]) {
    setup([new Response('secret-token-marker private Graph response', { status })]);
    await assert.rejects(searchEntraUsers('Sneha'), error => {
      assert.ok(error instanceof EntraUserSearchError);
      assert.equal(error.status, status);
      assert.equal(error.code, 'GRAPH_REQUEST_FAILED');
      assert.ok(!String(error).includes('secret-token-marker'));
      return true;
    });
  }
  assert.deepEqual(logs, []);
});

test('throttling surfaces Retry-After without immediate retries', async () => {
  const requests = setup([new Response('private response', { status: 429, headers: { 'Retry-After': '30' } })]);
  await assert.rejects(searchEntraUsers('Sneha'), { code: 'THROTTLED', status: 429, retryAfterSeconds: 30 });
  assert.equal(requests.length, 1);
});

test('supports HTTP-date Retry-After and ignores malformed values', async () => {
  const future = new Date(Date.now() + 60_000).toUTCString();
  setup([new Response('', { status: 429, headers: { 'Retry-After': future } })]);
  await assert.rejects(searchEntraUsers('Sneha'), error => error.retryAfterSeconds > 0 && error.retryAfterSeconds <= 60);
  setup([new Response('', { status: 429, headers: { 'Retry-After': 'invalid' } })]);
  await assert.rejects(searchEntraUsers('Sneha'), error => error.retryAfterSeconds === undefined);
});

test('network and timeout failures are sanitized', async () => {
  setup();
  for (const error of [new Error('secret-token-marker'), new DOMException('private-url', 'TimeoutError')]) {
    global.fetch = async () => { throw error; };
    await assert.rejects(searchEntraUsers('Sneha'), failure => {
      assert.equal(failure.code, 'GRAPH_REQUEST_FAILED');
      assert.ok(!String(failure).includes('secret-token-marker'));
      assert.equal(failure.cause, undefined);
      return true;
    });
  }
});

test('rejects malformed success responses', async () => {
  for (const response of [new Response('not json'), Response.json({}), Response.json({ value: {} }), Response.json({ value: [], '@odata.nextLink': 42 })]) {
    setup([response]);
    await assert.rejects(searchEntraUsers('Sneha'), { code: 'INVALID_RESPONSE' });
  }
});

test('follows Graph pages with headers, fills after skipped users, and deduplicates IDs', async () => {
  const requests = setup([
    { value: [user(), user({ accountEnabled: false })], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' },
    { value: [user(), user({ id: 'user-2' })] },
  ]);
  assert.equal((await searchEntraUsers('Sneha')).length, 2);
  assert.equal(requests.length, 2);
  assert.equal(requests[1].url.searchParams.get('$skiptoken'), 'abc');
  assert.equal(requests[1].options.headers.ConsistencyLevel, 'eventual');
});

test('stops at 25 results and at five pages', async () => {
  let requests = setup([{ value: Array.from({ length: 30 }, (_, i) => user({ id: String(i) })), '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=next' }]);
  assert.equal((await searchEntraUsers('Sneha')).length, 25);
  assert.equal(requests.length, 1);
  requests = setup(Array.from({ length: 5 }, (_, i) => ({ value: [], '@odata.nextLink': `https://graph.microsoft.com/v1.0/users?$skiptoken=${i}` })));
  assert.deepEqual(await searchEntraUsers('Sneha'), []);
  assert.equal(requests.length, 5);
});

test('rejects untrusted nextLinks before sending credentials', async () => {
  for (const next of ['https://example.com/users', 'http://graph.microsoft.com/v1.0/users', 'https://graph.microsoft.com/v1.0/me', 'https://name:password@graph.microsoft.com/v1.0/users', '/relative']) {
    const requests = setup([{ value: [], '@odata.nextLink': next }]);
    await assert.rejects(searchEntraUsers('Sneha'), { code: 'INVALID_RESPONSE' });
    assert.equal(requests.length, 1);
  }
});

test('does not return partial success when a later page fails', async () => {
  setup([{ value: [user()], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=next' }, new Response('', { status: 403 })]);
  await assert.rejects(searchEntraUsers('Sneha'), { status: 403 });
});

test('actual reused token provider does not log raw Entra errors', async t => {
  const logs = [];
  t.mock.method(console, 'error', (...args) => logs.push(args.join(' ')));
  t.mock.method(console, 'log', () => {});
  const saved = Object.fromEntries(['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'].map(key => [key, process.env[key]]));
  try {
    for (const key of Object.keys(saved)) process.env[key] = 'test-value';
    provider.clearTokenCache();
    global.fetch = async () => Response.json({ error: 'secret-token-marker', error_description: 'private response' }, { status: 400 });
    assert.equal(await originalTokenProvider(), null);
    assert.equal(logs.length, 1);
    assert.ok(logs[0].includes('HTTP 400'));
    assert.ok(!logs.join('').includes('secret-token-marker'));
    assert.ok(!logs.join('').includes('private response'));
  } finally {
    for (const [key, value] of Object.entries(saved)) {
      if (value === undefined) delete process.env[key]; else process.env[key] = value;
    }
    provider.clearTokenCache();
  }
});
