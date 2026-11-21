// The CalDAV client against tsdav's real request and parse code, behind a fake server answering
// raw multistatus XML (src/testing/caldav-fake-server.ts). caldav-client.test.ts stubs what tsdav
// returns; these tests cover what tsdav makes of the bytes a server sends.
import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { CalDAVCalendarClient } from './caldav-client.js';
import { buildUnusableCopiesNote, calendarEventBody, formatCalendarEventList } from './response-formatters.js';
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
  calendars: Array<{ displayName?: string; url: string }> = CALENDARS,
) {
  const server = makeFakeCalDAVServer(store, override);
  const dav = makeFakeServerDAVClient(calendars, server, {
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

const UNREADABLE_ANSWER = { name: 'UnreadableDavAnswerError', message: /^Calendar "Personal": the calendar server answered a read with a success status but a body that is not a complete WebDAV multistatus, or was not sent as XML, so what this calendar holds is unknown\. This call fails rather than read that answer as an empty calendar\. Nothing was changed\./ };

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

  // A nameless calendar is labelled by its url, and these two differ only past 60 characters.
  it('names a nameless calendar by its whole url', async () => {
    const base = `${ORIGIN}/dav/calendars/user/probe/a-long-collection-name-shared-by-both-calendars-`;
    const first = `${base}first/`;
    const second = `${base}second/`;
    const { client } = setup(
      {},
      (request) => (request.url === second ? NOT_XML : undefined),
      [{ url: first }, { url: second }],
    );
    await assert.rejects(listWindow(client), { message: new RegExp(`^Calendar "${second.replace(/[.]/g, '\\.')}": the calendar server`) });
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

describe('the end of a success answer is read whole before it is trusted (#239)', () => {
  const STORE = {
    [PERSONAL]: [
      { href: `${PERSONAL}a.ics`, data: ical('a@fm', 'Alpha'), etag: '"ea"' },
      { href: `${PERSONAL}b.ics`, data: ical('b@fm', 'Beta'), etag: '"eb"' },
    ],
  };
  const answering = (answer: FakeOverride) => setup(STORE,
    (request) => (request.url === PERSONAL && request.kind === 'query' ? answer : undefined)).client;
  const rootWith = (before: string, attrs = '') => TWO_EVENTS.replace('<d:multistatus ', `${before}<d:multistatus ${attrs}`);
  const ROOT_END = /<\/d:multistatus>$/;

  // Every body below except the first parses in tsdav, so only the end-of-body scan decides it.
  const rows: Array<[string, string, 'accepted' | 'rejected']> = [
    ['non-whitespace text after the root', `${TWO_EVENTS} trailing`, 'rejected'],
    ['a CDATA section after the root', `${TWO_EVENTS}<![CDATA[x]]>`, 'rejected'],
    ['a second element after the root', `${TWO_EVENTS}<d:other/>`, 'rejected'],
    ['an unterminated comment after the root', `${TWO_EVENTS}<!-- cut`, 'rejected'],
    ['a comment cut to <!--> after the root', `${TWO_EVENTS}<!-->`, 'rejected'],
    ['an unterminated processing instruction after the root', `${TWO_EVENTS}<?pi cut`, 'rejected'],
    ['a processing instruction cut to <?> after the root', `${TWO_EVENTS}<?>`, 'rejected'],
    ['a body cut inside a tag name', cutAfter('<d:resp', 2), 'rejected'],
    ['a body cut before the root end tag\'s >', TWO_EVENTS.slice(0, -1), 'rejected'],
    ['a root end tag with a space after </', TWO_EVENTS.replace(ROOT_END, '</ d:multistatus>'), 'rejected'],
    // tsdav strips the trailing underscore and reads it as a multistatus.
    ['a root named multistatus_', TWO_EVENTS.replace('<d:multistatus ', '<d:multistatus_ ').replace(ROOT_END, '</d:multistatus_>'), 'rejected'],
    ['a markup declaration inside the root', TWO_EVENTS.replace('<d:response>', '<!foo><d:response>'), 'rejected'],
    ['a root attribute holding /> in double quotes', rootWith('', 'x="/>" '), 'accepted'],
    ['a root attribute holding /> in single quotes', rootWith('', 'x=\'/>\' '), 'accepted'],
    ['a root attribute holding a double quote and /> in single quotes', rootWith('', 'x=\'"/>\' '), 'accepted'],
    ['a root attribute holding a single quote and /> in double quotes', rootWith('', 'x="\'/>" '), 'accepted'],
    ['a DOCTYPE before the root', rootWith('<!DOCTYPE d:multistatus>'), 'accepted'],
    ['a DOCTYPE with an internal subset before the root', rootWith('<!DOCTYPE d:multistatus [<!ENTITY x "y">]>'), 'accepted'],
  ];
  for (const [label, body, outcome] of rows) {
    it(`${outcome === 'accepted' ? 'lists' : 'fails'} a listing whose answer has ${label}`, async () => {
      const client = answering({ body, contentType: XML });
      if (outcome === 'rejected') {
        await assert.rejects(listWindow(client), UNREADABLE_ANSWER);
      } else {
        assert.deepEqual((await listWindow(client)).events.map(e => e.title).sort(), ['Alpha', 'Beta']);
      }
    });
  }

  // davRequest throws, rather than returning its stand-in, when a status line is an element
  // whose own toString is not a function.
  it('fails a listing whose answer makes tsdav\'s own parse throw', async () => {
    const body = multistatus('<d:response><d:href>/dav/calendars/user/probe/personal/a.ics</d:href>'
      + '<d:status><d:toString/></d:status></d:response>');
    await assert.rejects(listWindow(answering({ body, contentType: XML })), UNREADABLE_ANSWER);
  });

  it('checks every 2xx status and leaves the rest to tsdav', async () => {
    assert.deepEqual((await listWindow(answering({ body: TWO_EVENTS, status: 200, contentType: XML }))).events.map(e => e.title).sort(), ['Alpha', 'Beta']);
    for (const status of [200, 299]) {
      await assert.rejects(listWindow(answering({ ...NOT_XML, status })), UNREADABLE_ANSWER, String(status));
    }
    for (const status of [404, 500]) {
      await assert.rejects(listWindow(answering({ ...NOT_XML, status })), {
        name: 'Error',
        message: `CalDAV request failed: "Collection query failed: ${status} . Raw response: this is not xml"`,
      }, String(status));
    }
  });
});

const UNREADABLE_NOTE = '\n\nNote: an event marked `unreadable` is a record the calendar server reported in or near '
  + 'this window whose event data this server could not read, so only its `url` is known; its title, times, '
  + 'free/busy and whether it repeats are not. A retry helps only if the server failed to send the data.';

describe('a listed record that came back unreadable is listed with a warning (#239)', () => {
  const failedData = '<d:propstat><d:prop><c:calendar-data/></d:prop><d:status>HTTP/1.1 500 Internal Server Error</d:status></d:propstat>';
  const rows: Array<[string, string, string]> = [
    ['a calendar-data propstat that failed', etagProp('"e"'), failedData],
    ['an empty calendar-data element', etagProp('"e"') + '<c:calendar-data/>', ''],
    ['calendar-data that tsdav reads as a number', etagProp('"e"') + '<c:calendar-data>12345</c:calendar-data>', ''],
    // Three layouts that read the same on every retry (docs/reference/calendar.md).
    ['calendar-data split across two CDATA sections', etagProp('"e"')
      + `<c:calendar-data><![CDATA[${ical('split@fm').slice(0, 40)}]]><![CDATA[${ical('split@fm').slice(40)}]]></c:calendar-data>`, ''],
    ['calendar-data mixing text with a CDATA section', etagProp('"e"')
      + `<c:calendar-data>BEGIN:VCALENDAR<![CDATA[${ical('mixed@fm').slice('BEGIN:VCALENDAR'.length)}]]></c:calendar-data>`, ''],
    ['its only BEGIN:VEVENT line legally folded', etagProp('"e"') + dataProp(ical('folded@fm').replace('BEGIN:VEVENT', 'BEGIN:VEV\r\n ENT')), ''],
  ];
  for (const [label, props, extraPropstat] of rows) {
    it(`marks a listed record with ${label} unreadable, with one note`, async () => {
      const path = '/dav/calendars/user/probe/personal/';
      const body = multistatus(
        resp(`${path}broken-1.ics`, props, '200 OK', extraPropstat)
        + resp(`${path}fine.ics`, etagProp('"ef"') + dataProp(ical('fine@fm', 'Fine')))
        + resp(`${path}broken-2.ics`, props, '200 OK', extraPropstat),
      );
      const { client } = setup(
        { [PERSONAL]: [{ href: `${PERSONAL}fine.ics`, data: ical('fine@fm', 'Fine'), etag: '"ef"' }] },
        (request) => (request.url === PERSONAL && request.kind === 'query' ? { body, contentType: XML } : undefined),
      );
      const result = await listWindow(client);
      for (const name of ['broken-1.ics', 'broken-2.ics']) {
        const url = `${PERSONAL}${name}`;
        assert.deepEqual(result.events.find(e => e.url === url), { id: url, url, title: 'Untitled', unreadable: true });
      }
      const text = formatCalendarEventList(result);
      assert.equal(text.split(UNREADABLE_NOTE).length - 1, 1, text);
    });
  }

  it('a readable row carries no unreadable key and no note', async () => {
    const { client } = setup({ [PERSONAL]: [{ href: `${PERSONAL}fine.ics`, data: ical('fine@fm', 'Fine'), etag: '"ef"' }] });
    const result = await listWindow(client);
    assert.equal(result.events.length, 1);
    assert.equal('unreadable' in result.events[0], false);
    assert.equal(formatCalendarEventList(result).includes('unreadable'), false);
  });
});

describe('get_calendar_event names a record it matched and could not use (#239)', () => {
  const unusable = (id: string, copies: string) =>
    `The event id "${id}" matched a record this server could not use as this event: ${copies}. `
    + 'A retry helps only if the server failed to send the record whole.';

  it('get names the record it could not read instead of answering not found', async () => {
    const byUid = setup({ [WORK]: [{ href: `${WORK}solo.ics`, data: ical('solo@fm'), etag: '"e2"', query: 'nodata' }] });
    await assert.rejects(byUid.client.getCalendarEventById('solo@fm'), {
      name: 'Error',
      message: unusable('solo@fm', `"Work" ("${WORK}solo.ics")`),
    });
    const url = `${PERSONAL}real.ics`;
    const byUrl = setup({ [PERSONAL]: [{ href: url, data: ical('real@fm'), etag: '"ex"', fetch: 'nodata' }] });
    await assert.rejects(byUrl.client.getCalendarEventById(url), {
      name: 'Error',
      message: unusable(url, `"Personal" ("${url}")`),
    });
  });

  it('answers with the copy it could use and discloses the one it could not', async () => {
    const { client } = setup({
      [PERSONAL]: [{ href: `${PERSONAL}solo.ics`, data: ical('solo@fm', 'Personal copy'), etag: '"e1"' }],
      [WORK]: [{ href: `${WORK}solo.ics`, data: ical('solo@fm', 'Work copy'), etag: '"e2"', query: 'nodata' }],
    });
    const result = await client.getCalendarEventById('solo@fm');
    assert.equal(result.event.title, 'Personal copy');
    assert.equal(result.otherCopies, undefined);
    assert.deepEqual(result.unreadableCopies, [{ calendar: 'Work', url: `${WORK}solo.ics` }]);
    assert.deepEqual(calendarEventBody(result.event, result.otherCopies, result.unreadableCopies), {
      ...result.event,
      unreadableCopies: [{ calendar: 'Work', url: `${WORK}solo.ics` }],
    });
    assert.equal(
      buildUnusableCopiesNote(result.unreadableCopies),
      `\n\nNote: this event id also matched a record this server could not use as this event: "Work" ("${WORK}solo.ics"). `
      + 'The event above is a record it could use. update_calendar_event and delete_calendar_event refuse this id while '
      + 'that record cannot be used.',
    );
  });

  // Matched by its url and answered, but with no UID it fails the writes' rule as well.
  it('does not disclose the answered record as an unusable copy of itself', async () => {
    const url = `${PERSONAL}nouid.ics`;
    const noUid = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'BEGIN:VEVENT', 'DTSTART:20260401T100000Z', 'SUMMARY:No UID',
      'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
    const { client } = setup({ [PERSONAL]: [{ href: url, data: noUid, etag: '"en"' }] });
    const result = await client.getCalendarEventById(url);
    assert.equal(result.event.title, 'No UID');
    assert.equal(result.unreadableCopies, undefined);
  });

  // A response whose href is the calendar's own names the collection, not a record.
  it('still answers not found when the only hit is the calendar itself', async () => {
    const url = `${PERSONAL}real.ics`;
    const { client } = setup({}, (request) => (request.url === PERSONAL && request.kind === 'multiget'
      ? { body: multistatus(resp('/dav/calendars/user/probe/personal/', etagProp('"ec"'))), contentType: XML }
      : undefined));
    await assert.rejects(client.getCalendarEventById(url), { message: /^Calendar event not found: / });
  });

  // The server matches UIDs case-insensitively; a record read whole under another UID is not a copy.
  it('does not name a record whose own UID differs from the id', async () => {
    const otherCase: FakeResource = { href: `${PERSONAL}upper.ics`, data: ical('SOLO@fm'), etag: '"eu"', query: 'noetag' };
    const alone = setup({ [PERSONAL]: [otherCase] });
    await assert.rejects(alone.client.getCalendarEventById('solo@fm'), { message: /^Calendar event not found: "solo@fm"/ });
    const beside = setup({
      [PERSONAL]: [otherCase],
      [WORK]: [{ href: `${WORK}solo.ics`, data: ical('solo@fm', 'Work copy'), etag: '"e2"' }],
    });
    const result = await beside.client.getCalendarEventById('solo@fm');
    assert.equal(result.event.title, 'Work copy');
    assert.equal(result.unreadableCopies, undefined);
    // The writes keep refusing on it: they cannot rule out a copy they could not write safely.
    await assert.rejects(beside.client.deleteCalendarEvent('solo@fm'));
    assert.equal(beside.dav.deleteCalendarObject.mock.callCount(), 0);
  });

  // A record the caller's url addresses is named whatever UID it holds.
  it('names the addressed record that came back without its entity tag', async () => {
    const url = `${PERSONAL}real.ics`;
    const { client } = setup({ [PERSONAL]: [{ href: url, data: ical('real@fm'), etag: '"ex"', fetch: 'noetag' }] });
    await assert.rejects(client.getCalendarEventById(url), { message: unusable(url, `"Personal" ("${url}")`) });
  });

  // A url the server put in its answer that lies outside the calendar asked is not named.
  it('names no url outside the calendar the record was read from', async () => {
    for (const stray of ['/dav/calendars/user/probe/personal/stray.ics', 'https://elsewhere.invalid/dav/calendars/user/probe/work/x.ics']) {
      const strayAnswer = (request: FakeRequest) => (request.url === WORK && request.kind === 'multiget'
        ? { body: multistatus(resp(stray, etagProp('"es"'))), contentType: XML }
        : undefined);
      const workCopy = { href: `${WORK}solo.ics`, data: ical('solo@fm', 'Work copy'), etag: '"e2"' };
      const alone = setup({ [WORK]: [workCopy] }, strayAnswer);
      await assert.rejects(alone.client.getCalendarEventById('solo@fm'), {
        message: unusable('solo@fm', '"Work" (at an address outside that calendar)'),
      }, stray);
      const beside = setup({
        [PERSONAL]: [{ href: `${PERSONAL}solo.ics`, data: ical('solo@fm', 'Personal copy'), etag: '"e1"' }],
        [WORK]: [workCopy],
      }, strayAnswer);
      const result = await beside.client.getCalendarEventById('solo@fm');
      assert.deepEqual(result.unreadableCopies, [{ calendar: 'Work' }], stray);
      assert.equal(
        buildUnusableCopiesNote(result.unreadableCopies).includes('"Work" (at an address outside that calendar)'),
        true,
        stray,
      );
    }
  });

  it('has no note when every match could be used', () => {
    assert.equal(buildUnusableCopiesNote(undefined), '');
    assert.equal(buildUnusableCopiesNote([]), '');
    const event = { id: 'solo@fm', title: 'Solo' } as Parameters<typeof calendarEventBody>[0];
    assert.deepEqual(calendarEventBody(event, undefined, []), event);
  });

  it('writes two copies as records in the note', () => {
    assert.equal(
      buildUnusableCopiesNote([{ calendar: 'Personal', url: `${PERSONAL}solo.ics` }, { calendar: 'Work' }]),
      `\n\nNote: this event id also matched 2 records this server could not use as this event: "Personal" ("${PERSONAL}solo.ics"), `
      + '"Work" (at an address outside that calendar). The event above is a record it could use. update_calendar_event and '
      + 'delete_calendar_event refuse this id while those records cannot be used.',
    );
  });

  it('names a copy matched by an id that came with surrounding whitespace', async () => {
    const { client } = setup({ [WORK]: [{ href: `${WORK}solo.ics`, data: ical('solo@fm'), etag: '"e2"', query: 'noetag' }] });
    await assert.rejects(client.getCalendarEventById('  solo@fm  '), { message: unusable('solo@fm', `"Work" ("${WORK}solo.ics")`) });
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
