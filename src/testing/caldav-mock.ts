// The DAVClient stand-in shared by the CalDAV client's test files.
//
// This module is test-only, and lives in src/testing/ for the reasons given in
// mock-calls.ts: outside the build, and outside the top level of src/.

import { mock } from 'node:test';

// A stand-in for tsdav's DAVClient with discovery already stubbed, for every test that
// reaches calendar discovery.
//
// The collection set is a required argument and is deliberately NEVER defaulted. A default
// would let a test that never said what its account holds start exercising some other
// discovery outcome the day discovery changes, and then pass or fail for a reason it does
// not name anywhere in its own text. Each caller states its own collections.
//
// Everything else the test needs from the client is passed in `rest` and returned untouched,
// so an assertion about what the client was asked to do — a call count, a recorded argument —
// reads the very mock function the test handed in.

// The stored master a mocked server hands back for a resource the expanded listing could not
// decide about (#155): a VEVENT with no RRULE, no RDATE and no RECURRENCE-ID, which is what a
// genuine one-off event's stored form looks like. Supplied by makeMockDAVClient below and by
// caldav-client.test.ts's makeHomeListingDAVClient because every listing row gets this
// question asked of it, so a fixture that omits it describes a server that does not exist. A
// test that is ABOUT the follow-up read passes its own `calendarMultiGet` and overrides this.
const STORED_NON_RECURRING = [
  'BEGIN:VCALENDAR', 'BEGIN:VEVENT', 'UID:stub@fixture.invalid',
  'DTSTART:20260101T000000Z', 'SUMMARY:Stub', 'END:VEVENT', 'END:VCALENDAR',
].join('\r\n');

export function defaultCalendarMultiGet() {
  return mock.fn(async (params: { objectUrls?: string[] }) =>
    (params.objectUrls ?? []).map(href => ({
      href,
      props: { getetag: '"fixture-etag"', calendarData: { _cdata: STORED_NON_RECURRING } },
    })));
}

export function makeMockDAVClient<Calendar, Rest extends object & { login?: never; fetchCalendars?: never }>(calendars: Calendar[], rest: Rest) {
  return {
    login: mock.fn(async () => {}),
    fetchCalendars: mock.fn(async () => calendars),
    calendarMultiGet: defaultCalendarMultiGet(),
    ...rest,
  };
}
