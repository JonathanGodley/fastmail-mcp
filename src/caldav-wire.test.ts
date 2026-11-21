// The CalDAV client against tsdav's real request and parse code, behind a fake server answering
// raw multistatus XML (src/testing/caldav-fake-server.ts). caldav-client.test.ts stubs what tsdav
// returns; these tests cover what tsdav makes of the bytes a server sends.
import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { CalDAVCalendarClient } from './caldav-client.js';
import {
  dataProp,
  etagProp,
  makeFakeCalDAVServer,
  makeFakeServerDAVClient,
  multistatus,
  resp,
  type FakeOverride,
  type FakeRequest,
  type FakeResource,
} from './testing/caldav-fake-server.js';

const ORIGIN = 'https://caldav.example.invalid';
const PERSONAL = `${ORIGIN}/dav/calendars/user/probe/personal/`;
const WORK = `${ORIGIN}/dav/calendars/user/probe/work/`;
const CALENDARS = [{ displayName: 'Personal', url: PERSONAL }, { displayName: 'Work', url: WORK }];

function ical(uid: string, summary = 'Meeting', extra: string[] = []): string {
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'BEGIN:VEVENT', `UID:${uid}`, 'DTSTART:20260401T100000Z',
    'DTEND:20260401T110000Z', `SUMMARY:${summary}`, ...extra, 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
}

function setup(
  store: Record<string, FakeResource[]>,
  override?: (request: FakeRequest, n: number) => FakeOverride | undefined,
) {
  const server = makeFakeCalDAVServer(store, override);
  const dav = makeFakeServerDAVClient(CALENDARS, server, {
    updateCalendarObject: mock.fn(async (_p: unknown) => ({ ok: true, status: 204, headers: new Headers() })),
    deleteCalendarObject: mock.fn(async (_p: unknown) => ({ ok: true, status: 204, headers: new Headers() })),
  });
  const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
  (client as any).client = dav;
  return { client, dav, server };
}

const listWindow = (client: CalDAVCalendarClient) => client.getCalendarEvents(undefined, 50, '2026-03-01', '2026-05-30');

describe('CalDAV reads through tsdav\'s own parse (#239)', () => {
  it('lists a well-formed answer through tsdav\'s own parse', async () => {
    const { client, server } = setup({
      [PERSONAL]: [{ href: `${PERSONAL}a.ics`, data: ical('a@fm', 'Alpha'), etag: '"ea"' }],
      [WORK]: [{ href: `${WORK}b.ics`, data: ical('b@fm', 'Beta'), etag: '"eb"' }],
    });
    const result = await listWindow(client);
    assert.deepEqual(result.events.map(e => [e.title, e.url]).sort(), [
      ['Alpha', `${PERSONAL}a.ics`],
      ['Beta', `${WORK}b.ics`],
    ]);
    assert.ok(server.requests.some(r => r.kind === 'query' && r.expand), 'the listing reached the fake server');
  });

  it('reads an empty multistatus, self-closed or not, as an empty window', async () => {
    for (const body of [
      '<?xml version="1.0" encoding="utf-8"?><d:multistatus xmlns:d="DAV:"/>',
      '<?xml version="1.0" encoding="utf-8"?><d:multistatus xmlns:d="DAV:"></d:multistatus>',
    ]) {
      const { client, server } = setup({}, () => ({ body }));
      const result = await listWindow(client);
      assert.deepEqual(result.events, [], body);
      assert.equal(result.total, 0, body);
      assert.equal(server.requests.length, 2, body);
    }
  });

  it('routes both legs of a UID read through the client\'s fetch', async () => {
    const { client, server } = setup({
      [PERSONAL]: [{ href: `${PERSONAL}solo.ics`, data: ical('solo@fm', 'Solo'), etag: '"e1"' }],
    });
    const result = await client.getCalendarEventById('solo@fm');
    assert.equal(result.event.title, 'Solo');
    assert.deepEqual(server.requests.map(r => [r.kind, r.url]), [
      ['query', PERSONAL],
      ['multiget', PERSONAL],
      ['query', WORK],
    ]);
  });
});

const UNREADABLE_ANSWER = { message: /^Calendar "Personal": the calendar server answered a read with a success status but a body that is not a complete WebDAV multistatus, or was not sent as XML, so what this calendar holds is unknown\. This call fails rather than read that answer as an empty calendar\. Nothing was changed\./ };

const XML = 'application/xml; charset=utf-8';
const NOT_XML: FakeOverride = { body: 'this is not xml', contentType: XML };
const GARBLED = (request: FakeRequest) => (request.url === PERSONAL ? NOT_XML : undefined);

// One complete answer for two listed events, cut at chosen points below.
const TWO_EVENTS = multistatus(
  resp('/dav/calendars/user/probe/personal/a.ics', etagProp('"ea"') + dataProp(ical('a@fm', 'Alpha')))
  + resp('/dav/calendars/user/probe/personal/b.ics', etagProp('"eb"') + dataProp(ical('b@fm', 'Beta', ['DESCRIPTION:see </d:multistatus> here']))),
);
const cutAfter = (marker: string, nth = 1) => {
  let at = -1;
  for (let i = 0; i < nth; i++) at = TWO_EVENTS.indexOf(marker, at + 1);
  return TWO_EVENTS.slice(0, at + marker.length);
};

describe('a success answer that is not a complete multistatus is an error (#239)', () => {
  const rows: Array<[string, FakeOverride]> = [
    ['not XML', NOT_XML],
    ['a mismatched end tag', { body: '<?xml version="1.0"?><d:multistatus xmlns:d="DAV:"><d:response></d:propstat></d:multistatus>', contentType: XML }],
    ['an undefined entity', { body: '<?xml version="1.0"?><d:multistatus xmlns:d="DAV:"><d:response><d:href>/a&b.ics</d:href></d:response></d:multistatus>', contentType: XML }],
    ['an empty body', { body: '', contentType: XML }],
    ['a 204 with no body', { status: 204 }],
    ['a complete multistatus sent as text/plain', { body: TWO_EVENTS, contentType: 'text/plain' }],
    ['a complete multistatus with no content type', { body: TWO_EVENTS, contentType: null }],
    ['a well-formed root that is not a multistatus', { body: '<?xml version="1.0"?><d:error xmlns:d="DAV:"/>', contentType: XML }],
    ['a body cut after a complete response', { body: cutAfter('</d:response>'), contentType: XML }],
    ['a body cut inside calendar-data', { body: cutAfter('SUMMARY:Be'), contentType: XML }],
    ['a body cut right after a nested element named like the root', {
      body: `${cutAfter('</d:response>')}<d:response><d:href>/dav/calendars/user/probe/personal/b.ics</d:href><d:error><d:multistatus></d:multistatus>`,
      contentType: XML,
    }],
    ['a body cut inside a CDATA right after the text of the root\'s end tag', { body: cutAfter('see </d:multistatus>'), contentType: XML }],
  ];
  for (const [label, answer] of rows) {
    it(`fails a listing whose answer is ${label}`, async () => {
      const { client } = setup(
        { [PERSONAL]: [{ href: `${PERSONAL}a.ics`, data: ical('a@fm', 'Alpha'), etag: '"ea"' }] },
        (request) => (request.url === PERSONAL && request.kind === 'query' ? answer : undefined),
      );
      await assert.rejects(listWindow(client), UNREADABLE_ANSWER);
    });
  }

  it('refuses a delete when another calendar\'s UID-query leg does not parse', async () => {
    const { client, dav } = setup(
      {
        [PERSONAL]: [],
        [WORK]: [{ href: `${WORK}solo.ics`, data: ical('solo@fm'), etag: '"e2"' }],
      },
      (request) => (request.url === PERSONAL && request.kind === 'query' ? NOT_XML : undefined),
    );
    await assert.rejects(client.deleteCalendarEvent('solo@fm'), UNREADABLE_ANSWER);
    assert.equal(dav.deleteCalendarObject.mock.callCount(), 0);
  });

  it('get and delete by url fail on an unparseable multiget', async () => {
    const url = `${PERSONAL}real.ics`;
    const store = { [PERSONAL]: [{ href: url, data: ical('real@fm'), etag: '"ex"' }] };
    const garbledFetch = (request: FakeRequest) => (request.url === PERSONAL && request.kind === 'multiget' ? NOT_XML : undefined);
    const forGet = setup(store, garbledFetch);
    await assert.rejects(forGet.client.getCalendarEventById(url), UNREADABLE_ANSWER);
    const forDelete = setup(store, garbledFetch);
    await assert.rejects(forDelete.client.deleteCalendarEvent(url), UNREADABLE_ANSWER);
    assert.equal(forDelete.dav.deleteCalendarObject.mock.callCount(), 0);
  });

  it('settle fails with the unparseable-answer error', async () => {
    const { client, server } = setup(
      { [PERSONAL]: [{ href: `${PERSONAL}a.ics`, data: ical('a@fm', 'Alpha'), etag: '"ea"' }] },
      (request) => (request.url === PERSONAL && request.kind === 'multiget' ? NOT_XML : undefined),
    );
    await assert.rejects(listWindow(client), UNREADABLE_ANSWER);
    assert.ok(server.requests.some(r => r.url === PERSONAL && r.kind === 'multiget'), 'the settle read was sent');
  });

  // A daily series from 1 April; every expansion holds its first instance and the 3 April one.
  const SERIES_URL = `${PERSONAL}series.ics`;
  const instance = (date: string, overridden: boolean) => ['BEGIN:VEVENT', 'UID:series@fm',
    ...(overridden ? [`RECURRENCE-ID:${date}T100000Z`] : []), `DTSTART:${date}T100000Z`, `DTEND:${date}T110000Z`,
    'SUMMARY:Series', 'END:VEVENT'];
  const SERIES: FakeResource = {
    href: SERIES_URL,
    data: ical('series@fm', 'Series', ['RRULE:FREQ=DAILY;COUNT=10']),
    expandedData: ['BEGIN:VCALENDAR', 'VERSION:2.0', ...instance('20260401', false), ...instance('20260403', true), 'END:VCALENDAR'].join('\r\n'),
    etag: '"es"',
  };
  const breakExpansion = (which: number) => {
    let expansions = 0;
    return (request: FakeRequest) => (request.kind === 'query' && request.expand && ++expansions === which ? NOT_XML : undefined);
  };

  it('occurrence delete refuses before writing on an unparseable expansion', async () => {
    const { client, dav } = setup({ [PERSONAL]: [SERIES] }, breakExpansion(1));
    await assert.rejects(client.deleteCalendarEvent('series@fm', '2026-04-03T10:00:00Z'), UNREADABLE_ANSWER);
    assert.equal(dav.updateCalendarObject.mock.callCount(), 0);
  });

  it('reports NOT CONFIRMED when the confirm read does not parse', async () => {
    const { client, dav } = setup({ [PERSONAL]: [SERIES] }, breakExpansion(3));
    const result = await client.deleteCalendarEvent('series@fm', '2026-04-03T10:00:00Z');
    assert.equal(dav.updateCalendarObject.mock.callCount(), 1);
    assert.match(result.occurrence?.failure ?? '', /^NOT CONFIRMED: .* but the read that would confirm it failed, /);
  });

  it('still classifies a deleted url as not found', async () => {
    for (const status of [404, 410]) {
      const { client } = setup(
        { [PERSONAL]: [] },
        (request) => (request.kind === 'multiget' ? { status, body: 'Gone', contentType: 'text/plain' } : undefined),
      );
      await assert.rejects(client.getCalendarEventById(`${PERSONAL}deleted.ics`), { message: /^Calendar event not found: / }, String(status));
    }
  });

  it('accepts trailing whitespace, comments and processing instructions after the root', async () => {
    for (const body of [
      `${TWO_EVENTS}\r\n  \n`,
      `${TWO_EVENTS}<!-- trailing --> <?trailing pi?>\n`,
      TWO_EVENTS.replace(/<\/d:multistatus>$/, '</d:multistatus  >'),
    ]) {
      const { client } = setup(
        { [PERSONAL]: [{ href: `${PERSONAL}a.ics`, data: ical('a@fm', 'Alpha'), etag: '"ea"' }, { href: `${PERSONAL}b.ics`, data: ical('b@fm', 'Beta'), etag: '"eb"' }] },
        (request) => (request.url === PERSONAL && request.kind === 'query' ? { body, contentType: XML } : undefined),
      );
      const result = await listWindow(client);
      assert.deepEqual(result.events.map(e => e.title).sort(), ['Alpha', 'Beta'], JSON.stringify(body.slice(-30)));
    }
  });
});

describe('every calendar-object read passes the checked fetch (#239)', () => {
  const source = readFileSync(join(dirname(fileURLToPath(import.meta.url)), 'caldav-client.ts'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, block => block.replace(/[^\n]/g, ' '))
    .split('\n').map(line => line.replace(/(^|[^:])\/\/.*$/, '$1')).join('\n');

  /** Each call's argument list, from its open parenthesis to the one that closes it. */
  const calls = [...source.matchAll(/\b(fetchCalendarObjects|calendarMultiGet)\s*\(/g)].map((match) => {
    let depth = 0;
    let end = match.index + match[0].length - 1;
    for (; end < source.length; end++) {
      if (source[end] === '(') depth++;
      else if (source[end] === ')' && --depth === 0) break;
    }
    return source.slice(match.index, end + 1);
  });

  it('finds the reads it guards', () => {
    assert.ok(calls.length > 0, 'the scan found no fetchCalendarObjects or calendarMultiGet call');
  });

  // The listing, the UID query, the url fetch, settleAmbiguousRecurrence and the occurrence expansion.
  it('finds exactly the five read sites', () => {
    assert.equal(calls.length, 5, calls.join('\n\n'));
  });

  it('passes the checked fetch at every one', () => {
    assert.deepEqual(calls.filter(call => !/\bfetch: checkedReadFetch\(/.test(call)), []);
  });
});
