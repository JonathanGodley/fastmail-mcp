import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { listCalendarEventsTool } from './calendar-list-handler.js';
import type { CalendarListClient } from './calendar-list-handler.js';
import type { CalendarEventQueryResult } from './caldav-client.js';
import { buildBrokenCollectionNote } from './response-formatters.js';

function stubClient(result: CalendarEventQueryResult) {
  const getCalendarEvents = mock.fn(async (..._args: Parameters<CalendarListClient['getCalendarEvents']>) => result);
  return { client: { getCalendarEvents } satisfies CalendarListClient, getCalendarEvents };
}

const EVENTS = [
  { id: 'a@fm', url: '/cal/work/a.ics', title: 'A', start: '2027-03-02T09:00:00Z' },
  { id: 'b@fm', url: '/cal/work/b.ics', title: 'B', start: '2027-03-03T09:00:00Z' },
];

describe('listCalendarEventsTool', () => {
  // Every page re-reads every calendar, so a bad offset must cost no read at all.
  it('rejects an unusable position before any CalDAV call', async () => {
    for (const position of [-1, '-3', 1.5, '1.5', 'abc', [2], true]) {
      const { client, getCalendarEvents } = stubClient({ events: [], total: 0 });
      await assert.rejects(
        () => listCalendarEventsTool({ position }, 50, client),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError', String(position));
          assert.match(err.message, /^position /);
          // This tool has no `ascending` parameter to point at.
          assert.doesNotMatch(err.message, /ascending/);
          return true;
        },
        `expected rejection for ${JSON.stringify(position)}`,
      );
      assert.equal(getCalendarEvents.mock.callCount(), 0);
    }
  });

  it('passes the arguments, the clamped limit and the coerced position to the client', async () => {
    const { client, getCalendarEvents } = stubClient({ events: [], total: 0, position: 40 });
    await listCalendarEventsTool(
      { calendarId: 'Work', startDate: '2027-03-01', endDate: '2027-03-10', position: ' 40 ' },
      20,
      client,
    );
    assert.deepEqual(getCalendarEvents.mock.calls[0].arguments, ['Work', 20, '2027-03-01', '2027-03-10', 40]);
  });

  // A broken collection means the answer is incomplete; the 'write' or 'create' sentence
  // would tell the caller something else was or was not checked.
  it('discloses a broken collection with the read consequence', async () => {
    const paths = ['/cal/broken/'];
    const { client } = stubClient({ events: EVENTS, total: 2, brokenCollections: paths });
    const [content] = await listCalendarEventsTool({}, 50, client);
    assert.ok(content.text.endsWith(buildBrokenCollectionNote(paths, 'read')), content.text);
  });

  // The switch passes `{}`, but the handler tolerates undefined args, as the contacts handlers do.
  it('reads the first page when called with no arguments', async () => {
    const { client, getCalendarEvents } = stubClient({ events: EVENTS, total: 2, position: 0 });
    const [content] = await listCalendarEventsTool(undefined, 50, client);
    assert.deepEqual(getCalendarEvents.mock.calls[0].arguments, [undefined, 50, undefined, undefined, undefined]);
    assert.equal(content.text.split('\n')[0], 'Showing 2 of 2 results.');
  });

  it('renders the page as a paged listing', async () => {
    const { client } = stubClient({ events: EVENTS, total: 5, position: 0 });
    const [content] = await listCalendarEventsTool({}, 2, client);
    assert.equal(content.type, 'text');
    assert.equal(content.text.split('\n')[0], 'Showing 2 of 5 results. nextPosition: 2 (pass position:2 for the next page).');
  });
});
