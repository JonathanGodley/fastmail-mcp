// The CalDAV client against tsdav's real request and parse code, behind a fake server answering
// raw multistatus XML (src/testing/caldav-fake-server.ts). caldav-client.test.ts stubs what tsdav
// returns; these tests cover what tsdav makes of the bytes a server sends.
import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { CalDAVCalendarClient } from './caldav-client.js';
import {
  makeFakeCalDAVServer,
  makeFakeServerDAVClient,
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
