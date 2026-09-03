import { after, before, describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import {
  extractVEvent,
  resolveDisplayName,
  parseICalValue,
  findValueBoundary,
  parseAllICalProperties,
  hasICalProperty,
  parseAttendee,
  parseICalDuration,
  formatICalDate,
  parseCalendarObject,
  parseCalendarObjects,
  eventIntersectsWindow,
  escapeICalText,
  unescapeICalText,
  validateAndFormatICalDate,
  parseICalDateAsUTC,
  normalizeMasterVEventFirst,
  toICalUTC,
  foldICalLine,
  detectLineEnding,
  replaceICalProperty,
  removeAllICalProperties,
  removeOrphanedVTimezones,
  removeExceptionVEvents,
  insertBeforeEndVEvent,
  validateAttendeeEmail,
  quoteParamValue,
  sortEventsByStart,
  CalDAVCalendarClient,
  describeCreateCalendarEventResult,
  describeUpdateCalendarEventResult,
  unwrapDisplayName,
  CALENDAR_MAX_OCCURRENCES_PER_SERIES,
} from './caldav-client.js';
// A value import, not `import type`: the redirect test below stubs a method on the
// prototype, which needs the class itself. It still serves the type positions.
import { DAVClient } from 'tsdav';
import { callArguments } from './testing/mock-calls.js';
// The calendar window's local-day resolution reads the deployment's configured zone from
// the single place it is stored, so a test that asserts on a window has to pin that zone.
import { setDefaultTimezone } from './email-formatter.js';
// For the timeZone/endTimeZone serialisation check: `toolJson` is the seam get_calendar_event
// renders through, `formatQueryResult` is the one list_calendar_events renders through.
import { toolJson, isUsableTimezone, resolveCalendarInstantMs, InvalidInputError } from './coerce.js';
import { formatQueryResult } from './response-formatters.js';

// The mocked DAVClient methods below declare these parameter lists rather than
// taking no arguments. A `mock.fn(async () => …)` stub records its arguments as an
// empty tuple, so every assertion about what the client sent tsdav would be reading
// an element the type system knows cannot exist; taking the shapes from tsdav's own
// signatures also means a tsdav upgrade that changes a request shape shows up here
// instead of in a passing test asserting against the old one.
type FetchObjectsParams = Parameters<DAVClient['fetchCalendarObjects']>[0];
type UpdateObjectParams = Parameters<DAVClient['updateCalendarObject']>[0];
type CreateObjectParams = Parameters<DAVClient['createCalendarObject']>[0];
type DeleteObjectParams = Parameters<DAVClient['deleteCalendarObject']>[0];

// requireNonEmpty / validateClearFields are owned by src/coerce.ts and covered
// there (including the InvalidInputError class the calendar tools depend on for
// their InvalidParams mapping), so they are not re-tested against this module.

describe('extractVEvent', () => {
  it('extracts VEVENT block from iCalendar data', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'SUMMARY:Test Event',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    // Extraction returns null when it finds no VEVENT, and every assertion below
    // reads the block: say so here, so a regression that stops finding the block at
    // all reports that instead of a null dereference.
    assert.ok(vevent, 'expected a VEVENT block to be extracted');
    assert.ok(vevent.includes('SUMMARY:Test Event'));
    assert.ok(vevent.includes('DTSTART;TZID=Europe/Rome:20260320T083000'));
    assert.ok(!vevent.includes('VTIMEZONE'));
    assert.ok(!vevent.includes('TZID:Europe/Rome'));
  });

  it('returns null when no VEVENT block found', () => {
    const data = 'no vevent here';
    assert.equal(extractVEvent(data), null);
  });

  it('ignores VTIMEZONE DTSTART when extracting VEVENT', () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'END:STANDARD',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'SUMMARY:Meeting',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const vevent = extractVEvent(ical);
    assert.ok(vevent, 'expected a VEVENT block to be extracted');
    // Should only have the VEVENT DTSTART, not the VTIMEZONE one
    const dtstartMatches = vevent.match(/DTSTART/g);
    assert.equal(dtstartMatches?.length, 1);
    assert.ok(vevent.includes('20260320T083000'));
  });
});

describe('parseICalValue', () => {
  it('handles simple KEY:value format', () => {
    const vevent = 'SUMMARY:Test Event\nDTSTART:20260320T083000Z';
    assert.equal(parseICalValue(vevent, 'SUMMARY'), 'Test Event');
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000Z');
  });

  it('handles parameterized KEY;TZID=...:value format', () => {
    const vevent = 'DTSTART;TZID=Europe/Rome:20260320T083000\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260320T083000');
  });

  it('handles VALUE=DATE format', () => {
    const vevent = 'DTSTART;VALUE=DATE:20260324\nDTEND;VALUE=DATE:20260325';
    assert.equal(parseICalValue(vevent, 'DTSTART'), '20260324');
    assert.equal(parseICalValue(vevent, 'DTEND'), '20260325');
  });

  it('returns undefined for missing keys', () => {
    const vevent = 'SUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'LOCATION'), undefined);
  });

  // #102: RFC 5545 whitespace inside a property VALUE is significant, so a free-text
  // property must read back exactly as stored rather than silently losing padding.
  it('does not trim a free-text value', () => {
    const vevent = 'SUMMARY:  padded  ';
    assert.equal(parseICalValue(vevent, 'SUMMARY'), '  padded  ');
  });

  it('handles line folding (continuation lines)', () => {
    const vevent = 'DESCRIPTION:This is a long\n description that wraps\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });
});

describe('formatICalDate', () => {
  it('formats datetime without timezone', () => {
    assert.equal(formatICalDate('20260320T083000'), '2026-03-20T08:30:00');
  });

  it('formats datetime with Z suffix', () => {
    assert.equal(formatICalDate('20260320T083000Z'), '2026-03-20T08:30:00Z');
  });

  it('formats all-day date', () => {
    assert.equal(formatICalDate('20260324'), '2026-03-24');
  });

  it('returns undefined for undefined input', () => {
    assert.equal(formatICalDate(undefined), undefined);
  });

  it('returns cleaned string for unrecognized formats', () => {
    assert.equal(formatICalDate('something-else'), 'something-else');
  });

  it('strips carriage returns', () => {
    assert.equal(formatICalDate('20260320T083000\r'), '2026-03-20T08:30:00');
  });
});

describe('parseCalendarObject', () => {
  it('parses a full calendar object with VTIMEZONE + VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'BEGIN:STANDARD',
      'DTSTART:19701025T030000',
      'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
      'TZOFFSETFROM:+0200',
      'TZOFFSETTO:+0100',
      'END:STANDARD',
      'BEGIN:DAYLIGHT',
      'DTSTART:19700329T020000',
      'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
      'TZOFFSETFROM:+0100',
      'TZOFFSETTO:+0200',
      'END:DAYLIGHT',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:abc123@fastmail',
      'DTSTART;TZID=Europe/Rome:20260320T083000',
      'DTEND;TZID=Europe/Rome:20260320T093000',
      'SUMMARY:Morning Meeting',
      'DESCRIPTION:Discuss project\\nSecond line',
      'LOCATION:Room A\\, Building 1',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: 'https://caldav.example.com/cal/abc.ics' });

    assert.equal(event.id, 'abc123@fastmail');
    assert.equal(event.url, 'https://caldav.example.com/cal/abc.ics');
    assert.equal(event.title, 'Morning Meeting');
    assert.equal(event.description, 'Discuss project\nSecond line');
    assert.equal(event.location, 'Room A, Building 1');
    // Should get the VEVENT DTSTART, not the VTIMEZONE one
    assert.equal(event.start, '2026-03-20T08:30:00');
    assert.equal(event.end, '2026-03-20T09:30:00');
  });

  it('parses an all-day event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday1@fastmail',
      'DTSTART;VALUE=DATE:20260324',
      'DTEND;VALUE=DATE:20260325',
      'SUMMARY:All Day Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-24');
    assert.equal(event.end, '2026-03-25');
    assert.equal(event.title, 'All Day Event');
  });

  it('parses a UTC event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:utc1@fastmail',
      'DTSTART:20260320T083000Z',
      'DTEND:20260320T093000Z',
      'SUMMARY:UTC Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-03-20T08:30:00Z');
    assert.equal(event.end, '2026-03-20T09:30:00Z');
  });

  it('defaults title to Untitled when SUMMARY is missing', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:notitle@fastmail',
      'DTSTART:20260320T083000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.title, 'Untitled');
  });

  it('handles missing optional fields', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:minimal@fastmail',
      'DTSTART:20260320T083000Z',
      'SUMMARY:Minimal',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.description, undefined);
    assert.equal(event.location, undefined);
    assert.equal(event.end, undefined);
  });
});

// The emit matrix for `timeZone`/`endTimeZone` (#139) — see the field comments on
// `CalendarEvent` in src/caldav-client.ts and the "Calendar times: a name, never an offset"
// section of docs/conventions.md for the rule these tests hold to. `configuredZone` is passed
// directly rather than through `setDefaultTimezone`, so these are deterministic regardless of
// the test host's own zone.
describe('timeZone / endTimeZone (#139)', () => {
  // NOT this test host's own zone (Australia/Sydney) — every fixture below passes CONFIGURED
  // straight to parseCalendarObject as an explicit option, but a regression that silently
  // ignored that option and fell back to resolveUsableTimezone(undefined) (the host zone)
  // would still pass every omit-when-same assertion here if CONFIGURED happened to equal the
  // host zone, exactly the unfalsifiable-pin trap the end-to-end wiring test below documents.
  const CONFIGURED = 'Europe/London';

  it('emits timeZone with the name when the TZID differs from the configured zone', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:diff@fm',
      'DTSTART;TZID=Pacific/Auckland:20260320T083000',
      'DTEND;TZID=Pacific/Auckland:20260320T093000',
      'SUMMARY:NZ meeting',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, 'Pacific/Auckland');
  });

  it('omits timeZone when the TZID matches the configured zone, including a different case', () => {
    const exact = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:same@fm',
      'DTSTART;TZID=Europe/London:20260320T083000',
      'SUMMARY:Standup',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    assert.equal(parseCalendarObject({ data: exact, url: '' }, { configuredZone: CONFIGURED }).timeZone, undefined);

    const differentCase = exact.replace('TZID=Europe/London', 'TZID=europe/london');
    assert.equal(
      parseCalendarObject({ data: differentCase, url: '' }, { configuredZone: CONFIGURED }).timeZone,
      undefined,
    );
  });

  it('emits timeZone: null for a genuinely floating start (no TZID, no Z)', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:floating@fm',
      'DTSTART:20260320T083000',
      'SUMMARY:Floating',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, null);
  });

  it('omits timeZone for a Z-designated instant', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:utc@fm',
      'DTSTART:20260320T083000Z',
      'SUMMARY:UTC',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, undefined);
  });

  it('omits timeZone for a date-only (all-day) value', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday@fm',
      'DTSTART;VALUE=DATE:20260320',
      'SUMMARY:All day',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, undefined);
  });

  it('emits endTimeZone only when end is in a different TZID from start', () => {
    const flight = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:flight@fm',
      'DTSTART;TZID=Europe/London:20260320T083000',
      'DTEND;TZID=Pacific/Auckland:20260320T113000',
      'SUMMARY:Flight to Auckland',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const flightEvent = parseCalendarObject({ data: flight, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(flightEvent.timeZone, undefined, 'start matches the configured zone');
    assert.equal(flightEvent.endTimeZone, 'Pacific/Auckland');

    const sameZonePair = flight.replace('DTEND;TZID=Pacific/Auckland', 'DTEND;TZID=Europe/London');
    const sameZoneEvent = parseCalendarObject({ data: sameZonePair, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(sameZoneEvent.endTimeZone, undefined);
  });

  it('omits endTimeZone when end was computed from DURATION rather than a stored DTEND', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur@fm',
      'DTSTART;TZID=Pacific/Auckland:20260320T083000',
      'DURATION:PT1H',
      'SUMMARY:Duration event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    // start's own zone still comes through...
    assert.equal(event.timeZone, 'Pacific/Auckland');
    // ...and end was computed from start, so there is no independent end zone to report.
    assert.equal(event.endTimeZone, undefined);
  });

  it('omits endTimeZone on the DURATION path even when a malformed, empty-valued DTEND line is present', () => {
    // A VEVENT this malformed shouldn't exist, but `!rawEnd && rawStart` takes the DURATION
    // branch for it anyway (parseICalValue reads the empty value as falsy) — and
    // describeDateProperty classifies an empty value carrying a TZID parameter as `zoned`
    // regardless of the value being unused, so calling the shared end-classifier on this raw
    // line would report a zone (Europe/Paris) that never contributed to the computed `end`.
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:durMalformed@fm',
      'DTSTART;TZID=Pacific/Auckland:20260320T083000',
      'DURATION:PT1H',
      'DTEND;TZID=Europe/Paris:',
      'SUMMARY:Malformed duration + empty DTEND',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, 'Pacific/Auckland');
    assert.equal(event.endTimeZone, undefined, 'the unused DTEND TZID must not leak into endTimeZone');
  });

  it('unquotes a quoted TZID value', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:quoted@fm',
      'DTSTART;TZID="Pacific/Auckland":20260320T083000',
      'SUMMARY:Quoted TZID',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, 'Pacific/Auckland');
  });

  it('treats a leading-slash TZID (RFC 5545 §3.2.19) as the same zone, but emits it unstripped when different', () => {
    // libical strips exactly one leading '/' before comparing, so this matches the platform.
    // The simple '/Zone/Name' form is a globally-registered alias for the plain IANA name.
    const matching = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:slashMatch@fm',
      'DTSTART;TZID=/Europe/London:20260320T083000',
      'SUMMARY:Slash-prefixed, same zone',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    assert.equal(parseCalendarObject({ data: matching, url: '' }, { configuredZone: CONFIGURED }).timeZone, undefined);

    const differing = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:slashDiff@fm',
      'DTSTART;TZID=/Pacific/Auckland:20260320T083000',
      'SUMMARY:Slash-prefixed, different zone',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data: differing, url: '' }, { configuredZone: CONFIGURED });
    // Emitted with the stored spelling, slash and all — comparison is normalized, the value
    // reported to the caller is not.
    assert.equal(event.timeZone, '/Pacific/Auckland');

    // A vendor-registered form keeps comparing unequal even after stripping one slash, which
    // is the safe direction: an extra field, never a false "same zone" claim.
    const vendorPrefixed = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:vendor@fm',
      'DTSTART;TZID=/vendor.example/20050126_1/Australia/Sydney:20260320T083000',
      'SUMMARY:Vendor-prefixed',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const vendorEvent = parseCalendarObject({ data: vendorPrefixed, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(vendorEvent.timeZone, '/vendor.example/20050126_1/Australia/Sydney');
  });

  it('compares trimmed but emits the trimmed form, so a padded differing TZID still names a resolvable zone', () => {
    // isUsableTimezone("Australia/Sydney ") is false — an untrimmed emit here would poison
    // both the field itself and anything downstream that resolves it (sortEventsByStart).
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:padded@fm',
      'DTSTART;TZID=Pacific/Auckland :20260320T083000',
      'SUMMARY:Padded TZID',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, 'Pacific/Auckland');
    assert.equal(isUsableTimezone(event.timeZone!), true);
  });

  it('passes a non-IANA (Windows) TZID through verbatim, neither rejecting nor resolving it', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:windows@fm',
      'DTSTART;TZID=AUS Eastern Standard Time:20260320T083000',
      'SUMMARY:Third-party invite',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, 'AUS Eastern Standard Time');
  });

  it('compares endTimeZone against the configured zone when start is absent entirely', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noStart@fm',
      'DTEND;TZID=Pacific/Auckland:20260320T093000',
      'SUMMARY:No DTSTART',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, undefined);
    assert.equal(event.endTimeZone, 'Pacific/Auckland');
  });

  it('survives the toolJson (get_calendar_event) and formatQueryResult (list_calendar_events) serialisation seams', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:floating@fm',
      'DTSTART:20260320T083000',
      'SUMMARY:Floating',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const event = parseCalendarObject({ data, url: '' }, { configuredZone: CONFIGURED });
    assert.equal(event.timeZone, null);

    // toolJson: JSON.stringify keeps a null key; a lossy formatter would drop it like undefined.
    const single = JSON.parse(toolJson(event));
    assert.equal(single.timeZone, null);
    assert.ok('timeZone' in single, 'timeZone key must survive toolJson, not just its value');

    // formatQueryResult: summary line followed by a JSON array of items.
    const rendered = formatQueryResult({ items: [event], total: 1 });
    const jsonLine = rendered.split('\n')[1];
    const parsed = JSON.parse(jsonLine);
    assert.equal(parsed[0].timeZone, null);
    assert.ok('timeZone' in parsed[0], 'timeZone key must survive formatQueryResult, not just its value');
  });
});

describe('escapeICalText', () => {
  it('escapes backslashes', () => {
    assert.equal(escapeICalText('path\\to\\file'), 'path\\\\to\\\\file');
  });

  it('escapes semicolons', () => {
    assert.equal(escapeICalText('a;b;c'), 'a\\;b\\;c');
  });

  it('escapes commas', () => {
    assert.equal(escapeICalText('Room A, Building 1'), 'Room A\\, Building 1');
  });

  it('escapes newlines', () => {
    assert.equal(escapeICalText('line1\nline2'), 'line1\\nline2');
    assert.equal(escapeICalText('line1\r\nline2'), 'line1\\nline2');
  });

  it('leaves plain text unchanged', () => {
    assert.equal(escapeICalText('Team Standup'), 'Team Standup');
  });

  it('prevents ICS property injection via CRLF', () => {
    const malicious = 'Meeting\r\nATTENDEE:mailto:attacker@evil.com';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected ATTENDEE stays inside the text value,
    // not on its own ICS property line
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    assert.equal(escaped, 'Meeting\\nATTENDEE:mailto:attacker@evil.com');
  });

  it('prevents injection of extra VEVENT components', () => {
    const malicious = 'Meeting\r\nEND:VEVENT\r\nBEGIN:VEVENT\r\nSUMMARY:Injected';
    const escaped = escapeICalText(malicious);
    // No literal newlines means the injected properties stay inside the SUMMARY value
    assert.ok(!escaped.includes('\n'), 'escaped text must not contain literal newlines');
    assert.ok(!escaped.includes('\r'), 'escaped text must not contain literal carriage returns');
    // The entire payload is on one logical ICS line, so END:VEVENT can't terminate the block
    assert.ok(escaped.startsWith('Meeting\\n'), 'newlines should be escaped, not literal');
  });
});

describe('unescapeICalText', () => {
  it('unescapes newlines', () => {
    assert.equal(unescapeICalText('line1\\nline2'), 'line1\nline2');
    assert.equal(unescapeICalText('line1\\Nline2'), 'line1\nline2');
  });

  it('unescapes semicolons and commas', () => {
    assert.equal(unescapeICalText('a\\;b\\;c'), 'a;b;c');
    assert.equal(unescapeICalText('Room A\\, Building 1'), 'Room A, Building 1');
  });

  it('unescapes backslashes', () => {
    assert.equal(unescapeICalText('path\\\\to\\\\file'), 'path\\to\\file');
  });

  it('decodes an escaped backslash followed by "n" as a literal backslash + n, not a newline', () => {
    // "\\n" in iCal = an escaped backslash ("\\") followed by a literal "n".
    // A chained-replace implementation corrupts this into "\<newline>".
    assert.equal(unescapeICalText('\\\\n'), '\\n');
    assert.equal(unescapeICalText('C:\\\\next'), 'C:\\next');
  });

  it('handles literal backslash followed by comma', () => {
    assert.equal(unescapeICalText('\\\\,'), '\\,');
  });

  it('round-trips with escapeICalText', () => {
    for (const original of ['plain', 'a;b,c', 'line1\nline2', 'back\\slash', 'C:\\next', 'mix\n;,\\end']) {
      assert.equal(unescapeICalText(escapeICalText(original)), original);
    }
  });

  it('leaves plain text unchanged', () => {
    assert.equal(unescapeICalText('Team Standup'), 'Team Standup');
  });
});

describe('iCal create/parse round-trip', () => {
  it('round-trips special characters through create and parse', () => {
    const title = 'Meeting; discuss, plan';
    const description = 'Line one\nLine two\nPath\\to\\file; note, important';
    const location = 'Room A, Building 1; Floor 2';

    const uid = 'test-roundtrip@fastmail-mcp';
    const now = '20260407T000000Z';
    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `DTSTART:${toICalUTC('2026-04-07T18:45:00+10:00')}`,
      `DTEND:${toICalUTC('2026-04-07T20:00:00+10:00')}`,
      foldICalLine(`SUMMARY:${escapeICalText(title)}`),
      foldICalLine(`DESCRIPTION:${escapeICalText(description)}`),
      foldICalLine(`LOCATION:${escapeICalText(location)}`),
      'END:VEVENT',
      'END:VCALENDAR',
    ].filter(Boolean).join('\r\n');

    const event = parseCalendarObject({ data: ical, url: 'https://example.com/cal/test.ics' });
    assert.equal(event.title, title);
    assert.equal(event.description, description);
    assert.equal(event.location, location);
    assert.equal(event.start, '2026-04-07T08:45:00Z');
    assert.equal(event.end, '2026-04-07T10:00:00Z');
  });
});

describe('CalDAVCalendarClient.getCalendarEvents', () => {
  function makeIcal(uid: string, summary: string, dtstart: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTART:${dtstart}`,
      `SUMMARY:${summary}`,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClient(calendarObjects: Array<{ data: string; url: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    // Override the private getClient method to return a mock DAVClient
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('sorts events by start date ascending', async () => {
    const objects = [
      { data: makeIcal('c@fm', 'Evening', '20260325T200000Z'), url: '/c.ics' },
      { data: makeIcal('a@fm', 'Morning', '20260325T080000Z'), url: '/a.ics' },
      { data: makeIcal('b@fm', 'Afternoon', '20260325T140000Z'), url: '/b.ics' },
    ];
    const { client } = createMockedClient(objects);
    // The window is named, and named as INSTANTS, so this test is about ordering and nothing
    // else. Left bounds-free it would get the default window (today plus a month, #142) and
    // silently become a test that March 2026 is in the past; written as dates it would depend
    // on whatever zone the host happens to be in.
    const { events, total } = await client.getCalendarEvents(
      undefined, 50, '2026-03-25T00:00:00Z', '2026-03-26T00:00:00Z',
    );

    assert.equal(events.length, 3);
    assert.equal(total, 3);
    assert.equal(events[0].title, 'Morning');
    assert.equal(events[1].title, 'Afternoon');
    assert.equal(events[2].title, 'Evening');
  });

  it('passes timeRange to fetchCalendarObjects when startDate/endDate provided', async () => {
    const objects = [
      { data: makeIcal('a@fm', 'Event', '20260325T100000Z'), url: '/a.ics' },
    ];
    const { client, mockDAVClient } = createMockedClient(objects);
    await client.getCalendarEvents(undefined, 50, '2026-03-25T00:00:00Z', '2026-03-26T00:00:00Z');

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    // WIDENED BY FOURTEEN HOURS AT EACH EDGE (#162). The range the caller asked for is
    // 25 March 00:00Z .. 26 March 00:00Z; the range SENT runs from the 24th at 10:00Z to the
    // 26th at 14:00Z, because the server withholds all-day and floating events from a window
    // that only touches their UTC day. What comes back is then filtered exactly against the
    // caller's own window.
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-03-24T10:00:00Z',
      end: '2026-03-26T14:00:00Z',
    });
    // Expansion rides with the time range: without it a recurring series comes back as its
    // master, dated wherever the series began, and no in-window occurrence exists to report.
    assert.equal(callArgs.expand, true);
  });

  it('sends a 31-day window from local today, expanded, when no dates are provided (#142)', async () => {
    // A call naming neither bound used to go out with NO time range, and therefore with no
    // `expand` either — tsdav drops `<C:expand>` without one. So the call most likely to be
    // asked "what is on?" was the one call answering with series masters at their original
    // DTSTART, and the only alternative was an open-ended expansion nobody can bound.
    //
    // The clock is INJECTED and the zone PINNED, because a default window computed from the
    // real clock could only be asserted against a value this test recomputed the same way.
    setDefaultTimezone('Australia/Sydney');
    try {
      const client = new CalDAVCalendarClient({
        username: 'test',
        password: 'test',
        // Noon on 24 August in Sydney (+10). The local day is the 24th, whose midnight is
        // 2026-08-23T14:00:00Z — a different UTC day, which is the point of the pin.
        now: () => Date.parse('2026-08-24T02:00:00Z'),
      });
      const mockDAVClient = {
        login: mock.fn(async () => {}),
        fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
        fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
      };
      (client as any).client = mockDAVClient;

      const { windowClamp } = await client.getCalendarEvents(undefined, 50);

      const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
      // Local midnight today .. 31 fixed days later, widened by fourteen hours at each edge
      // like any other window.
      assert.deepEqual(callArgs.timeRange, {
        start: '2026-08-23T00:00:00Z',
        end: '2026-09-24T04:00:00Z',
      });
      // Expansion rides with the range here exactly as it does for a caller-named window: a
      // default window that reported series masters would answer the wrong question.
      assert.equal(callArgs.expand, true);
      // And it is disclosed, for the same reason every other narrowing is: a caller who is
      // not told reads "nothing" as an empty calendar rather than as "not the days I meant".
      assert.ok(windowClamp, 'an invented window must be disclosed');
      assert.equal(windowClamp!.invented, 'both');
      assert.equal(windowClamp!.start, '2026-08-23T14:00:00Z');
      assert.equal(windowClamp!.end, '2026-09-23T14:00:00Z');
    } finally {
      setDefaultTimezone(undefined);
    }
  });
});

describe('validateAndFormatICalDate', () => {
  it('accepts and formats date-only', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18', 'start'), '20260418');
  });

  it('accepts and formats UTC datetime', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00Z', 'start'), '20260418T100000Z');
  });

  it('accepts and normalizes positive offset to UTC', () => {
    // 2026-04-18T10:00:00+02:00 = 2026-04-18T08:00:00Z
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00+02:00', 'start'), '20260418T080000Z');
  });

  it('accepts and normalizes negative offset to UTC', () => {
    // 2026-04-18T10:00:00-05:00 = 2026-04-18T15:00:00Z
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00-05:00', 'start'), '20260418T150000Z');
  });

  it('accepts floating datetime (no zone)', () => {
    assert.equal(validateAndFormatICalDate('2026-04-18T10:00:00', 'start'), '20260418T100000');
  });

  it('rejects CRLF injection attempt', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\r\nATTENDEE:mailto:attacker@example.com', 'start'),
      /control characters/,
    );
  });

  it('rejects bare LF injection attempt', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\nATTENDEE:mailto:attacker@example.com', 'start'),
      /control characters/,
    );
  });

  it('rejects null byte', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z\0', 'start'),
      /control characters/,
    );
  });

  it('rejects malformed date', () => {
    assert.throws(
      () => validateAndFormatICalDate('not-a-date', 'start'),
      /must be ISO-8601/,
    );
  });

  it('rejects extra trailing content', () => {
    assert.throws(
      () => validateAndFormatICalDate('2026-04-18T10:00:00Z bonus', 'start'),
      /must be ISO-8601/,
    );
  });

  it('rejects non-string input', () => {
    assert.throws(
      () => validateAndFormatICalDate(undefined as any, 'start'),
      /must be a string/,
    );
  });

  it('throws with field name in error', () => {
    try {
      validateAndFormatICalDate('garbage', 'event.end');
      assert.fail('should have thrown');
    } catch (e) {
      assert.match((e as Error).message, /event\.end/);
    }
  });

  it('rejects a day its month does not have, in either form', () => {
    // new Date() rolls 2026-02-31 forward to 3 March, so accepting it would create the
    // event on a day the caller never named and report success.
    assert.throws(() => validateAndFormatICalDate('2026-02-31', 'start'), /not a real calendar date/);
    assert.throws(() => validateAndFormatICalDate('2026-02-31T10:00:00Z', 'start'), /not a real calendar date/);
    assert.throws(() => validateAndFormatICalDate('2026-06-31T10:00:00', 'start'), /not a real calendar date/);
  });

  it('classes every rejection as caller-fixable input', () => {
    for (const bad of ['garbage', '2026-02-31', '2026-04-18T10:00:00Z\r\nX', 42 as any]) {
      assert.throws(
        () => validateAndFormatICalDate(bad, 'start'),
        (err: unknown) => err instanceof Error && err.name === 'InvalidInputError',
      );
    }
  });

  it('trims surrounding whitespace before deciding the form', () => {
    // Untrimmed, ' 2026-04-18 ' misses the date-only shape and would be read as an
    // instant instead of an all-day value.
    assert.equal(validateAndFormatICalDate('  2026-04-18  ', 'start'), '20260418');
    assert.equal(validateAndFormatICalDate(' 2026-04-18T10:00:00Z ', 'start'), '20260418T100000Z');
  });
});

describe('CalDAVCalendarClient.updateCalendarEvent', () => {
  function makeFullIcal(uid: string, summary: string, dtstart: string, dtend: string, description?: string, location?: string): string {
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:20260401T000000Z`,
      `DTSTART:${dtstart}`,
      `DTEND:${dtend}`,
      `SUMMARY:${escapeICalText(summary)}`,
    ];
    if (description) lines.push(`DESCRIPTION:${escapeICalText(description)}`);
    if (location) lines.push(`LOCATION:${escapeICalText(location)}`);
    lines.push('END:VEVENT', 'END:VCALENDAR');
    return lines.join('\r\n');
  }

  function createMockedClientWithUpdateDelete(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
      deleteCalendarObject: mock.fn(async (_params: DeleteObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('updates only the title, preserving other fields', async () => {
    const ical = makeFullIcal('evt1@fm', 'Original Title', '20260401T100000Z', '20260401T110000Z', 'My description', 'Room A');
    const objects = [{ data: ical, url: '/cal/evt1.ics', etag: '"etag1"' }];
    const { client, mockDAVClient } = createMockedClientWithUpdateDelete(objects);

    const result = await client.updateCalendarEvent('evt1@fm', { title: 'New Title' });

    assert.equal(result.eventId, 'evt1@fm');
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedObj = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject;
    assert.ok(updatedObj.data.includes('SUMMARY:New Title'));
    assert.ok(updatedObj.data.includes('DESCRIPTION:My description'));
    assert.ok(updatedObj.data.includes('LOCATION:Room A'));
    assert.ok(updatedObj.data.includes('DTSTART:20260401T100000Z'));
    assert.ok(updatedObj.data.includes('UID:evt1@fm'));
  });

  it('updates start and end times', async () => {
    const ical = makeFullIcal('evt2@fm', 'Meeting', '20260401T100000Z', '20260401T110000Z');
    const objects = [{ data: ical, url: '/cal/evt2.ics' }];
    const { client, mockDAVClient } = createMockedClientWithUpdateDelete(objects);

    await client.updateCalendarEvent('evt2@fm', {
      start: '2026-04-02T14:00:00Z',
      end: '2026-04-02T15:00:00Z',
    });

    const updatedObj = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject;
    assert.ok(updatedObj.data.includes('DTSTART:20260402T140000Z'));
    assert.ok(updatedObj.data.includes('DTEND:20260402T150000Z'));
    assert.ok(updatedObj.data.includes('SUMMARY:Meeting'));
  });

  it('throws when event not found', async () => {
    const { client } = createMockedClientWithUpdateDelete([]);
    await assert.rejects(
      () => client.updateCalendarEvent('nonexistent@fm', { title: 'X' }),
      /Calendar event not found: nonexistent@fm/
    );
  });

  it('classes a not-found id as caller-fixable input, not a server fault', async () => {
    // The message assertion above passes under either class. A wrong event id is
    // something the caller corrects and re-sends, so the class has to say so:
    // InvalidInputError maps to InvalidParams, a plain Error to InternalError.
    const { client } = createMockedClientWithUpdateDelete([]);
    await assert.rejects(
      () => client.updateCalendarEvent('nonexistent@fm', { title: 'X' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Calendar event not found: nonexistent@fm/);
        return true;
      },
    );
  });

  it('classes a malformed start value as caller-fixable input', async () => {
    const ical = makeFullIcal('evt9@fm', 'Meeting', '20260401T100000Z', '20260401T110000Z');
    const { client } = createMockedClientWithUpdateDelete([{ data: ical, url: '/cal/evt9.ics' }]);
    await assert.rejects(
      () => client.updateCalendarEvent('evt9@fm', { start: 'tomorrow-ish' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Invalid start date format/);
        return true;
      },
    );
  });
});

describe('CalDAVCalendarClient.deleteCalendarEvent', () => {
  function createMockedClientWithDelete(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
      deleteCalendarObject: mock.fn(async (_params: DeleteObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('deletes an event by UID', async () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:del1@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:To Delete',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: ical, url: '/cal/del1.ics', etag: '"etag1"' }];
    const { client, mockDAVClient } = createMockedClientWithDelete(objects);

    await client.deleteCalendarEvent('del1@fm');

    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 1);
    const deletedObj = callArguments(mockDAVClient.deleteCalendarObject)[0].calendarObject;
    assert.equal(deletedObj.url, '/cal/del1.ics');
  });

  it('throws when event not found', async () => {
    const { client } = createMockedClientWithDelete([]);
    await assert.rejects(
      () => client.deleteCalendarEvent('nonexistent@fm'),
      /Calendar event not found: nonexistent@fm/
    );
  });

  it('classes a not-found id as caller-fixable input, not a server fault', async () => {
    // Pins the class, which the message assertion above cannot: InvalidInputError maps
    // to InvalidParams ("fix the id and re-send"), a plain Error to InternalError.
    const { client } = createMockedClientWithDelete([]);
    await assert.rejects(
      () => client.deleteCalendarEvent('nonexistent@fm'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Calendar event not found: nonexistent@fm/);
        return true;
      },
    );
  });
});

describe('CalDAVCalendarClient.getCalendarEventById', () => {
  function createMockedClientWithObjects(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('returns the parsed event when the UID exists', async () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:get1@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:Findable',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = createMockedClientWithObjects([{ data: ical, url: '/cal/get1.ics', etag: '"etag1"' }]);

    const event = await client.getCalendarEventById('get1@fm');
    assert.equal(event.id, 'get1@fm');
    assert.equal(event.title, 'Findable');
  });

  it('returns the STORED start and its zone exactly as written, unexpanded', async () => {
    // This is the call that reads a resource rather than a window over one, so it is the only
    // one that can promise the stored property back verbatim. `list_calendar_events` used to
    // share that promise on a bounds-free call; it no longer has one (#142), and its rows are
    // expansion output, which the server normalises to UTC instants and strips the RRULE from.
    setDefaultTimezone('America/New_York');
    try {
      const ical = [
        'BEGIN:VCALENDAR',
        'BEGIN:VEVENT',
        'UID:stored@fm',
        'DTSTART;TZID=Pacific/Auckland:20260320T083000',
        'DTEND;TZID=Pacific/Auckland:20260320T093000',
        'RRULE:FREQ=WEEKLY',
        'SUMMARY:Weekly sync',
        'END:VEVENT',
        'END:VCALENDAR',
      ].join('\r\n');
      const { client } = createMockedClientWithObjects([{ data: ical, url: '/cal/stored.ics' }]);

      const event = await client.getCalendarEventById('stored@fm');
      // The wall clock as stored — not converted, not stamped with an offset — and the TZID
      // alongside it, because it differs from the configured zone.
      assert.equal(event.start, '2026-03-20T08:30:00');
      assert.equal(event.timeZone, 'Pacific/Auckland');
      // And the series master, at its own DTSTART, with the rule intact: the unexpanded form
      // a listing row cannot show.
      assert.equal(event.recurrenceRule, 'FREQ=WEEKLY');
      assert.equal(event.recurrenceId, undefined);
    } finally {
      setDefaultTimezone(undefined);
    }
  });

  it('throws instead of returning null when the event does not exist', async () => {
    const { client } = createMockedClientWithObjects([]);
    await assert.rejects(
      () => client.getCalendarEventById('nonexistent@fm'),
      /Calendar event not found: nonexistent@fm/
    );
  });

  it('throws InvalidInputError for a not-found id so the boundary maps it to InvalidParams', async () => {
    // A wrong event id is caller-fixable, so it must not surface as InternalError
    // ("server-side, a bare retry might work"). The message assertion above passes
    // under either class; this one pins the class.
    const { client } = createMockedClientWithObjects([]);
    await assert.rejects(
      () => client.getCalendarEventById('nonexistent@fm'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });
});

// ============================================================
// New tests for calendar attendee support & non-destructive updates
// ============================================================

describe('findValueBoundary', () => {
  it('finds colon in simple property', () => {
    assert.equal(findValueBoundary('SUMMARY:Test'), 7);
  });

  it('skips colons inside quoted parameter values', () => {
    const line = 'ATTENDEE;DELEGATED-FROM="mailto:boss@example.com";CN="Smith, John":mailto:john@example.com';
    const idx = findValueBoundary(line);
    assert.equal(line.substring(idx + 1), 'mailto:john@example.com');
  });

  it('handles ALTREP with URL', () => {
    const line = 'DESCRIPTION;ALTREP="http://example.com/desc":Plain text';
    const idx = findValueBoundary(line);
    assert.equal(line.substring(idx + 1), 'Plain text');
  });

  it('returns -1 when no colon found', () => {
    assert.equal(findValueBoundary('NOCOLON'), -1);
  });
});

describe('parseAllICalProperties', () => {
  it('returns multiple ATTENDEEs', () => {
    const vevent = [
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'ATTENDEE;CN=Bob:mailto:bob@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
    ].join('\n');
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 2);
    assert.ok(results[0].includes('alice@'));
    assert.ok(results[1].includes('bob@'));
  });

  it('returns empty array when none found', () => {
    const vevent = 'BEGIN:VEVENT\nSUMMARY:Test\nEND:VEVENT';
    assert.deepEqual(parseAllICalProperties(vevent, 'ATTENDEE'), []);
  });

  it('handles folded ATTENDEE lines', () => {
    const vevent = [
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Very Long Name;PARTSTAT=ACCEPTED:mailto:long',
      ' name@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
    ].join('\n');
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.ok(results[0].includes('longname@example.com'));
  });

  it('handles CRLF input', () => {
    const vevent = 'BEGIN:VEVENT\r\nATTENDEE;CN=Alice:mailto:a@b.example\r\nEND:VEVENT';
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.ok(!results[0].includes('\r'));
  });

  it('does not match partial property names', () => {
    const vevent = 'BEGIN:VEVENT\nATTENDEE-X:foo\nATTENDEE:bar\nEND:VEVENT';
    const results = parseAllICalProperties(vevent, 'ATTENDEE');
    assert.equal(results.length, 1);
    assert.equal(results[0], 'ATTENDEE:bar');
  });
});

describe('parseAttendee', () => {
  it('parses simple ATTENDEE with CN', () => {
    const result = parseAttendee('ATTENDEE;CN=Alice:mailto:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
    assert.equal(result.name, 'Alice');
  });

  it('parses quoted CN with comma', () => {
    const result = parseAttendee('ATTENDEE;CN="Doe, John":mailto:john@example.com');
    assert.equal(result.name, 'Doe, John');
    assert.equal(result.email, 'john@example.com');
  });

  it('parses PARTSTAT, ROLE, CUTYPE, RSVP', () => {
    const result = parseAttendee('ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT;CUTYPE=INDIVIDUAL;RSVP=TRUE:mailto:alice@example.com');
    assert.equal(result.status, 'ACCEPTED');
    assert.equal(result.role, 'REQ-PARTICIPANT');
    assert.equal(result.cutype, 'INDIVIDUAL');
    assert.equal(result.rsvp, true);
  });

  it('converts RSVP=FALSE to boolean false', () => {
    const result = parseAttendee('ATTENDEE;RSVP=FALSE:mailto:alice@example.com');
    assert.equal(result.rsvp, false);
  });

  it('parses ORGANIZER lines', () => {
    const result = parseAttendee('ORGANIZER;CN=Boss:mailto:boss@example.com');
    assert.equal(result.email, 'boss@example.com');
    assert.equal(result.name, 'Boss');
  });

  it('handles missing CN', () => {
    const result = parseAttendee('ATTENDEE:mailto:anon@example.com');
    assert.equal(result.email, 'anon@example.com');
    assert.equal(result.name, undefined);
  });

  it('handles non-mailto URI', () => {
    const result = parseAttendee('ATTENDEE:urn:uuid:550e8400-e29b-41d4-a716-446655440000');
    assert.equal(result.email, 'urn:uuid:550e8400-e29b-41d4-a716-446655440000');
  });

  it('handles bare email without mailto', () => {
    const result = parseAttendee('ATTENDEE:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
  });

  it('handles DELEGATED-FROM with quoted mailto (colon inside quotes)', () => {
    const result = parseAttendee('ATTENDEE;DELEGATED-FROM="mailto:boss@example.com";CN=Alice:mailto:alice@example.com');
    assert.equal(result.email, 'alice@example.com');
    assert.equal(result.name, 'Alice');
  });

  it('omits empty CN', () => {
    const result = parseAttendee('ATTENDEE;CN=:mailto:alice@example.com');
    assert.equal(result.name, undefined);
  });

  it('handles CN with literal DQUOTE character', () => {
    // CN value with embedded quote — parseAttendee should store the raw value
    const result = parseAttendee('ATTENDEE;CN="John \'Doc\' Smith":mailto:john@example.com');
    assert.equal(result.name, "John 'Doc' Smith");
    assert.equal(result.email, 'john@example.com');
  });
});

describe('parseICalDuration', () => {
  it('parses PT2H', () => {
    assert.equal(parseICalDuration('PT2H', '2026-04-01T10:00:00Z'), '2026-04-01T12:00:00Z');
  });

  it('parses P1D', () => {
    assert.equal(parseICalDuration('P1D', '2026-04-01T10:00:00Z'), '2026-04-02T10:00:00Z');
  });

  it('parses P1W', () => {
    assert.equal(parseICalDuration('P1W', '2026-04-01T10:00:00Z'), '2026-04-08T10:00:00Z');
  });

  it('parses P1DT2H30M', () => {
    assert.equal(parseICalDuration('P1DT2H30M', '2026-04-01T10:00:00Z'), '2026-04-02T12:30:00Z');
  });

  it('parses PT90M', () => {
    assert.equal(parseICalDuration('PT90M', '2026-04-01T10:00:00Z'), '2026-04-01T11:30:00Z');
  });

  it('parses PT0S (zero duration)', () => {
    assert.equal(parseICalDuration('PT0S', '2026-04-01T10:00:00Z'), '2026-04-01T10:00:00Z');
  });

  it('parses P1DT0H0M0S (verbose)', () => {
    assert.equal(parseICalDuration('P1DT0H0M0S', '2026-04-01T10:00:00Z'), '2026-04-02T10:00:00Z');
  });

  it('returns undefined for malformed input', () => {
    assert.equal(parseICalDuration('2H', '2026-04-01T10:00:00Z'), undefined);
    assert.equal(parseICalDuration('PTXYZ', '2026-04-01T10:00:00Z'), undefined);
  });

  it('rejects bare P', () => {
    assert.equal(parseICalDuration('P', '2026-04-01T10:00:00Z'), undefined);
  });

  it('rejects P1DT (T with no time components)', () => {
    assert.equal(parseICalDuration('P1DT', '2026-04-01T10:00:00Z'), undefined);
  });

  it('handles date-only start', () => {
    assert.equal(parseICalDuration('P1D', '2026-04-01'), '2026-04-02');
  });

  it('returns floating time for floating start (no Z)', () => {
    const result = parseICalDuration('PT2H', '2026-04-01T10:00:00');
    assert.equal(result, '2026-04-01T12:00:00');
    assert.ok(!result!.includes('Z'), 'floating start should produce floating end');
  });
});

describe('parseCalendarObject with participants', () => {
  it('parses ATTENDEE and ORGANIZER', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:evt@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:Meeting',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT:mailto:alice@example.com',
      'ATTENDEE;CN=Bob;PARTSTAT=TENTATIVE:mailto:bob@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' }, { includeParticipants: true });
    assert.equal(event.organizer?.email, 'boss@example.com');
    assert.equal(event.organizer?.name, 'Boss');
    assert.equal(event.participants?.length, 2);
    assert.equal(event.participants?.[0].email, 'alice@example.com');
    assert.equal(event.participants?.[0].status, 'ACCEPTED');
    assert.equal(event.participants?.[1].email, 'bob@example.com');
  });

  it('omits participants when includeParticipants is false', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:evt@fm',
      'DTSTART:20260401T100000Z',
      'SUMMARY:Meeting',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.participants, undefined);
    assert.equal(event.organizer, undefined);
  });

  it('computes end from DURATION', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const event = parseCalendarObject({ data, url: '' });
    assert.equal(event.start, '2026-04-01T10:00:00Z');
    assert.equal(event.end, '2026-04-01T12:00:00Z');
  });

  it('returns minimal event when no VEVENT found', () => {
    const data = 'BEGIN:VCALENDAR\r\nEND:VCALENDAR';
    const event = parseCalendarObject({ data, url: '/test.ics' });
    assert.equal(event.title, 'Untitled');
    assert.equal(event.url, '/test.ics');
  });
});

describe('replaceICalProperty', () => {
  const simpleEvent = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'UID:test@fm',
    'SUMMARY:Original',
    'DTSTART:20260401T100000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\n');

  it('replaces an existing property', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(result.includes('SUMMARY:Updated'));
    assert.ok(!result.includes('SUMMARY:Original'));
  });

  it('preserves other properties when replacing', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(result.includes('UID:test@fm'));
    assert.ok(result.includes('DTSTART:20260401T100000Z'));
  });

  it('patches parameterized DTSTART (TZID form)', () => {
    const event = simpleEvent.replace('DTSTART:20260401T100000Z', 'DTSTART;TZID=Europe/Rome:20260401T100000');
    const result = replaceICalProperty(event, 'DTSTART', 'DTSTART;TZID=Europe/Rome:20260402T090000');
    assert.ok(result.includes('DTSTART;TZID=Europe/Rome:20260402T090000'));
    assert.ok(!result.includes('20260401'));
  });

  it('inserts a new property before END:VEVENT', () => {
    const result = replaceICalProperty(simpleEvent, 'LOCATION', 'LOCATION:Room A');
    assert.ok(result.includes('LOCATION:Room A'));
    const lines = result.split('\n');
    const locIdx = lines.findIndex(l => l === 'LOCATION:Room A');
    const endIdx = lines.findIndex(l => l === 'END:VEVENT');
    assert.ok(locIdx < endIdx);
  });

  it('removes a property when newLine is null', () => {
    const result = replaceICalProperty(simpleEvent, 'SUMMARY', null);
    assert.ok(!result.includes('SUMMARY'));
  });

  it('only patches first VEVENT in multi-VEVENT iCal', () => {
    const multiVevent = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'RECURRENCE-ID:20260401T100000Z',
      'SUMMARY:Exception',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(multiVevent, 'SUMMARY', 'SUMMARY:Updated Master');
    assert.ok(result.includes('SUMMARY:Updated Master'));
    assert.ok(result.includes('SUMMARY:Exception'));
  });

  it('skips properties inside VALARM', () => {
    const eventWithValarm = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:test@fm',
      'DESCRIPTION:Event description',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(eventWithValarm, 'DESCRIPTION', 'DESCRIPTION:Updated event');
    assert.ok(result.includes('DESCRIPTION:Updated event'));
    assert.ok(result.includes('DESCRIPTION:Reminder')); // VALARM DESCRIPTION preserved
  });

  it('does not touch VTIMEZONE DTSTART', () => {
    const eventWithTz = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'DTSTART:19700101T000000',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'SUMMARY:Test',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(eventWithTz, 'DTSTART', 'DTSTART;TZID=Europe/Rome:20260402T090000');
    assert.ok(result.includes('DTSTART:19700101T000000')); // VTIMEZONE preserved
    assert.ok(result.includes('DTSTART;TZID=Europe/Rome:20260402T090000'));
  });

  it('throws on missing BEGIN:VEVENT', () => {
    assert.throws(() => replaceICalProperty('no vevent', 'SUMMARY', 'SUMMARY:x'), /BEGIN:VEVENT not found/);
  });

  it('throws on empty input', () => {
    assert.throws(() => replaceICalProperty('', 'SUMMARY', 'SUMMARY:x'), /empty input/);
  });

  it('handles folded property lines', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'SUMMARY:Very long title that was',
      ' folded across two lines',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:Short');
    assert.ok(result.includes('SUMMARY:Short'));
    assert.ok(!result.includes('folded across'));
  });
});

describe('removeAllICalProperties', () => {
  it('removes multiple ATTENDEEs', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:test@fm',
      'SUMMARY:Test',
      'ATTENDEE;CN=Alice:mailto:alice@example.com',
      'ATTENDEE;CN=Bob:mailto:bob@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('ATTENDEE'));
    assert.ok(result.includes('ORGANIZER'));
    assert.ok(result.includes('SUMMARY:Test'));
  });

  it('handles folded ATTENDEE lines', () => {
    const event = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'ATTENDEE;CN=Very Long Name;PARTSTAT=ACCEPTED:mailto:long',
      ' name@example.com',
      'SUMMARY:Test',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('ATTENDEE'));
    assert.ok(!result.includes('longname'));
    assert.ok(result.includes('SUMMARY:Test'));
  });

  it('preserves CRLF line endings', () => {
    const event = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nATTENDEE:mailto:a@b.example\r\nSUMMARY:Test\r\nEND:VEVENT\r\nEND:VCALENDAR';
    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(result.includes('\r\n'));
    assert.ok(!result.includes('ATTENDEE'));
  });

  it('preserves LF-only line endings', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nATTENDEE:mailto:a@b.example\nSUMMARY:Test\nEND:VEVENT\nEND:VCALENDAR';
    const result = removeAllICalProperties(event, 'ATTENDEE');
    assert.ok(!result.includes('\r\n'));
    assert.ok(!result.includes('ATTENDEE'));
  });
});

describe('foldICalLine', () => {
  it('returns short lines unchanged', () => {
    assert.equal(foldICalLine('SUMMARY:Short'), 'SUMMARY:Short');
  });

  it('folds lines longer than 75 octets', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(Buffer.byteLength(lines[0], 'utf8') <= 75);
    assert.ok(lines[1].startsWith(' '));
  });

  it('folds very long lines into multiple segments', () => {
    const long = 'DESCRIPTION:' + 'y'.repeat(200);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(lines.length >= 3);
    for (let i = 1; i < lines.length; i++) {
      assert.ok(lines[i].startsWith(' '));
    }
  });

  it('keeps every segment within 75 octets', () => {
    const long = 'DESCRIPTION:' + 'z'.repeat(200);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    for (const line of lines) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75);
    }
  });

  it('folds multi-byte characters without exceeding 75 octets', () => {
    const long = 'LOCATION:' + '📍'.repeat(20);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(lines.length >= 2);
    for (const line of lines) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75,
        `Line exceeds 75 octets: ${Buffer.byteLength(line, 'utf8')} bytes`);
    }
  });
});

describe('foldICalLine with custom line ending', () => {
  it('uses LF when specified', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long, '\n');
    assert.ok(!folded.includes('\r'));
    assert.ok(folded.includes('\n'));
  });

  it('defaults to CRLF', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long);
    assert.ok(folded.includes('\r\n'));
  });
});

describe('removeOrphanedVTimezones', () => {
  it('removes VTIMEZONE with no references', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART:20260401T100000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeOrphanedVTimezones(data);
    assert.ok(!result.includes('VTIMEZONE'));
    assert.ok(!result.includes('Europe/Rome'));
  });

  it('preserves VTIMEZONE with references', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeOrphanedVTimezones(data);
    assert.ok(result.includes('VTIMEZONE'));
    assert.ok(result.includes('Europe/Rome'));
  });
});

describe('removeExceptionVEvents', () => {
  it('removes only orphaned exception VEVENTs', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'DTSTART:20260401T100000Z',
      'RRULE:FREQ=WEEKLY',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'RECURRENCE-ID:20260408T100000Z',
      'SUMMARY:Exception 1',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'RECURRENCE-ID:20260415T100000Z',
      'SUMMARY:Exception 2',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    // Only orphan the April 8 exception
    const orphaned = [new Date('2026-04-08T10:00:00Z')];
    const result = removeExceptionVEvents(data, orphaned);
    assert.ok(result.includes('SUMMARY:Master'));
    assert.ok(!result.includes('Exception 1'));
    assert.ok(result.includes('Exception 2'));
  });

  it('never touches master VEVENT', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:rec@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const result = removeExceptionVEvents(data, [new Date()]);
    assert.ok(result.includes('SUMMARY:Master'));
  });
});

describe('validateAttendeeEmail', () => {
  it('accepts valid email', () => {
    assert.doesNotThrow(() => validateAttendeeEmail('alice@example.com'));
  });

  it('rejects empty email', () => {
    assert.throws(() => validateAttendeeEmail(''), /required/);
  });

  it('rejects email without @', () => {
    assert.throws(() => validateAttendeeEmail('notanemail'), /Invalid/);
  });

  it('rejects email with newline', () => {
    assert.throws(() => validateAttendeeEmail('a@b.example\r\nX-INJECT:true'), /illegal/i);
  });

  it('rejects email with colon', () => {
    assert.throws(() => validateAttendeeEmail('a:b@example.com'), /illegal/i);
  });

  it('rejects email with semicolon', () => {
    assert.throws(() => validateAttendeeEmail('a;b@example.com'), /illegal/i);
  });

  it('rejects email with double quote', () => {
    assert.throws(() => validateAttendeeEmail('a"b@example.com'), /illegal/i);
  });

  it('rejects email with backslash', () => {
    assert.throws(() => validateAttendeeEmail('a\\b@example.com'), /illegal/i);
  });

  it('rejects email with whitespace', () => {
    assert.throws(() => validateAttendeeEmail('a b@example.com'), /illegal/i);
    assert.throws(() => validateAttendeeEmail('a\tb@example.com'), /illegal/i);
  });

  // #102: the single-`@` shape plus the old reject list let a bare address carry angle
  // brackets, a second address after a comma, or an invisible Unicode format character —
  // none of those is a single bare addr-spec.
  it('rejects an address wrapped in angle brackets', () => {
    assert.throws(() => validateAttendeeEmail('<a@b>'), /illegal/i);
  });

  it('rejects a second address appended after a comma', () => {
    assert.throws(() => validateAttendeeEmail('a@b,c'), /illegal/i);
  });

  it('rejects an address with a trailing bracketed fragment', () => {
    assert.throws(() => validateAttendeeEmail('a@b<c>'), /illegal/i);
  });

  it('rejects an address carrying a trailing U+200E left-to-right mark', () => {
    assert.throws(() => validateAttendeeEmail('a@b.example‎'), /illegal/i);
  });

  it('still accepts a plain dotted address', () => {
    assert.doesNotThrow(() => validateAttendeeEmail('alice.smith@example.com'));
  });

  it('still accepts a plus-tagged local part', () => {
    assert.doesNotThrow(() => validateAttendeeEmail('alice+tag@example.com'));
  });

  it('still accepts a hyphenated domain', () => {
    assert.doesNotThrow(() => validateAttendeeEmail('alice@my-domain.example.com')); // allowlist-secret: synthetic fixture on example.com, not a real address
  });
});

describe('quoteParamValue', () => {
  it('returns unquoted for simple values', () => {
    assert.equal(quoteParamValue('Alice'), 'Alice');
  });

  it('quotes values with comma', () => {
    assert.equal(quoteParamValue('Doe, John'), '"Doe, John"');
  });

  it('quotes values with semicolon', () => {
    assert.equal(quoteParamValue('A;B'), '"A;B"');
  });

  it('replaces double quotes with single quotes', () => {
    assert.equal(quoteParamValue('He said "hi"'), '"He said \'hi\'"');
  });

  it('strips newlines to prevent injection', () => {
    // Colon in result triggers DQUOTE quoting — that's correct
    assert.equal(quoteParamValue('Alice\r\nX-EVIL:payload'), '"Alice X-EVIL:payload"');
  });

  it('strips the remaining control characters and the C1 range', () => {
    assert.equal(quoteParamValue('Ali\x00ce\x07'), 'Alice');
    assert.equal(quoteParamValue('Ali\x1Bce'), 'Alice');
    assert.equal(quoteParamValue('Ali\x7Fce\x9B'), 'Alice');
  });

  it('strips the bidi overrides and isolates that would reorder the rendered name', () => {
    assert.equal(quoteParamValue('\u202EAlice'), 'Alice');
    assert.equal(quoteParamValue('Ali\u2066ce\u2069'), 'Alice');
  });

  it('keeps the plain left/right marks, which appear in real Arabic and Hebrew names', () => {
    // LRM/RLM carry no nesting scope and cannot reorder the text around them,
    // so stripping them would corrupt a legitimate name for no security gain.
    assert.equal(quoteParamValue('\u05D3\u05D5\u05D3\u200E'), '\u05D3\u05D5\u05D3\u200E');
    assert.equal(quoteParamValue('\u200F\u0645\u062D\u0645\u062F'), '\u200F\u0645\u062D\u0645\u062F');
  });

  it('keeps horizontal tab, which is legal in a quoted parameter value', () => {
    assert.equal(quoteParamValue('Ali\tce'), 'Ali\tce');
  });
});

describe('parseICalValue with CRLF and folded lines', () => {
  it('handles CRLF input with folded lines', () => {
    const vevent = 'DESCRIPTION:This is a long\r\n description that wraps\r\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), 'This is a longdescription that wraps');
  });

  it('uses quote-aware colon detection', () => {
    const vevent = 'ATTENDEE;DELEGATED-FROM="mailto:boss@example.com":mailto:alice@example.com\nSUMMARY:Test';
    assert.equal(parseICalValue(vevent, 'ATTENDEE'), 'mailto:alice@example.com');
  });
});

describe('toICalUTC', () => {
  it('converts timezone offset to UTC', () => {
    assert.equal(toICalUTC('2026-04-07T18:45:00+10:00'), '20260407T084500Z');
  });

  it('converts negative timezone offset to UTC', () => {
    assert.equal(toICalUTC('2026-04-07T08:45:00-05:00'), '20260407T134500Z');
  });

  it('handles UTC input (Z suffix)', () => {
    assert.equal(toICalUTC('2026-04-07T08:45:00Z'), '20260407T084500Z');
  });

  it('preserves floating time (no offset) without converting to UTC', () => {
    assert.equal(toICalUTC('2026-04-07T18:45:00'), '20260407T184500');
  });

  it('throws on invalid date input', () => {
    assert.throws(() => toICalUTC('not-a-date'), /Invalid date: not-a-date/);
  });

  it('handles midnight boundary crossing', () => {
    assert.equal(toICalUTC('2026-04-07T23:55:00+12:00'), '20260407T115500Z');
  });

  it('throws on date-only input', () => {
    assert.throws(() => toICalUTC('2026-04-01'), /date-only input must be handled by caller/);
  });
});

describe('CalDAVCalendarClient.updateCalendarEvent (patch-based)', () => {
  function makeRichIcal(uid: string, extra: string[] = []): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Google Inc//Google Calendar//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTAMP:20260401T000000Z',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DTEND;TZID=Europe/Rome:20260401T110000',
      'SUMMARY:Original Title',
      'DESCRIPTION:Original description',
      'LOCATION:Room A',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED:mailto:alice@example.com',
      'ATTENDEE;CN=Bob;PARTSTAT=TENTATIVE:mailto:bob@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'X-GOOGLE-CONFERENCE:https://meet.google.com/abc-def-ghi',
      'SEQUENCE:2',
      ...extra,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedPatchClient(calendarObjects: Array<{ data: string; url: string; etag?: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('preserves unknown properties when updating title only', async () => {
    const ical = makeRichIcal('evt1@fm');
    const objects = [{ data: ical, url: '/cal/evt1.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt1@fm', { title: 'New Title' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:New Title'));
    // Preserved properties
    assert.ok(updatedData.includes('ATTENDEE;CN=Alice'));
    assert.ok(updatedData.includes('ATTENDEE;CN=Bob'));
    assert.ok(updatedData.includes('ORGANIZER;CN=Boss'));
    assert.ok(updatedData.includes('VALARM'));
    assert.ok(updatedData.includes('X-GOOGLE-CONFERENCE'));
    assert.ok(updatedData.includes('DESCRIPTION:Original description'));
    assert.ok(updatedData.includes('LOCATION:Room A'));
  });

  it('does NOT re-emit DTSTART when only title changes (timezone preservation)', async () => {
    const ical = makeRichIcal('evt2@fm');
    const objects = [{ data: ical, url: '/cal/evt2.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt2@fm', { title: 'New Title' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    // Original DTSTART with TZID should be preserved exactly
    assert.ok(updatedData.includes('DTSTART;TZID=Europe/Rome:20260401T100000'));
  });

  it('increments SEQUENCE for location change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('evt3@fm');
    const objects = [{ data: ical, url: '/cal/evt3.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt3@fm', { location: 'New Room' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3')); // Was 2, now 3
  });

  it('does NOT increment SEQUENCE for title-only changes', async () => {
    const ical = makeRichIcal('evt4@fm');
    const objects = [{ data: ical, url: '/cal/evt4.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt4@fm', { title: 'New Title' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:2')); // Unchanged
  });

  it('does NOT increment SEQUENCE when no ATTENDEEs exist', async () => {
    const noAttendeeIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:solo@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Solo Event',
      'SEQUENCE:0',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noAttendeeIcal, url: '/cal/solo.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    // Stays before the untouched DTEND (20260401T110000Z) — a start change that
    // would jump past it is rejected on its own merits, which is not what this
    // test is about.
    await client.updateCalendarEvent('solo@fm', { start: '2026-04-01T10:30:00Z' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:0'));
  });

  it('removes DURATION when setting end', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: durationIcal, url: '/cal/dur.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('dur@fm', { end: '2026-04-01T13:00:00Z' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DTEND:20260401T130000Z'));
    assert.ok(!updatedData.includes('DURATION'));
  });

  it('does NOT remove DURATION when only setting start', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur2@fm',
      'DTSTART:20260401T100000Z',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: durationIcal, url: '/cal/dur2.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('dur2@fm', { start: '2026-04-02T10:00:00Z' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DURATION:PT2H'));
  });

  it('participants: [] removes all ATTENDEEs and the now-orphaned ORGANIZER', async () => {
    const ical = makeRichIcal('evt5@fm');
    const objects = [{ data: ical, url: '/cal/evt5.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt5@fm', { participants: [] });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(!updatedData.includes('ATTENDEE'));
    // An ORGANIZER with no ATTENDEEs is a malformed scheduling VEVENT, so it is stripped too.
    assert.ok(!updatedData.includes('ORGANIZER'));
  });

  it('rejects empty/whitespace/null title, description, location without writing', async () => {
    for (const fields of [
      { title: '' },
      { title: '   ' },
      { title: null as any },
      { description: '' },
      { description: '  ' },
      { location: '' },
    ]) {
      const ical = makeRichIcal('evtA@fm');
      const objects = [{ data: ical, url: '/cal/evtA.ics' }];
      const { client, mockDAVClient } = createMockedPatchClient(objects);
      await assert.rejects(() => client.updateCalendarEvent('evtA@fm', fields), /cannot be empty/);
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    }
  });

  it('empty title throws InvalidInputError so the index maps calendar input to InvalidParams (#41 collateral)', async () => {
    // The calendar tools share requireNonEmpty from coerce.ts, so the #41 reclassification
    // reaches them for free — pin it so it can't silently regress back to InternalError.
    // The message-regex assertion above passes under either error class; this one is what
    // actually holds the validators to coerce.ts rather than a local plain-Error copy.
    const ical = makeRichIcal('evtA@fm');
    const { client, mockDAVClient } = createMockedPatchClient([{ data: ical, url: '/cal/evtA.ics' }]);
    await assert.rejects(
      () => client.updateCalendarEvent('evtA@fm', { title: '' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('updating only participants preserves SUMMARY/DESCRIPTION/LOCATION', async () => {
    const ical = makeRichIcal('evtB@fm');
    const objects = [{ data: ical, url: '/cal/evtB.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evtB@fm', { participants: [{ email: 'carol@example.com' }] });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SUMMARY:Original Title'));
    assert.ok(updatedData.includes('DESCRIPTION:Original description'));
    assert.ok(updatedData.includes('LOCATION:Room A'));
  });

  it('clearFields: ["location"] removes the LOCATION line', async () => {
    const ical = makeRichIcal('evtC@fm');
    const objects = [{ data: ical, url: '/cal/evtC.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evtC@fm', { clearFields: ['location'] });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(!updatedData.includes('LOCATION:'));
    // Other content untouched
    assert.ok(updatedData.includes('SUMMARY:Original Title'));
    assert.ok(updatedData.includes('DESCRIPTION:Original description'));
  });

  it('clearFields rejects a non-clearable field (title) without writing', async () => {
    const ical = makeRichIcal('evtD@fm');
    const objects = [{ data: ical, url: '/cal/evtD.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtD@fm', { clearFields: ['title'] }),
      /Cannot clear "title"/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('rejects setting and clearing the same field', async () => {
    const ical = makeRichIcal('evtE@fm');
    const objects = [{ data: ical, url: '/cal/evtE.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtE@fm', { description: 'x', clearFields: ['description'] }),
      /cannot both set and clear description/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('floating end time preserves DTEND TZID', async () => {
    const ical = makeRichIcal('evt6@fm');
    const objects = [{ data: ical, url: '/cal/evt6.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt6@fm', { end: '2026-04-01T12:00:00' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DTEND;TZID=Europe/Rome:20260401T120000'));
  });

  it('date-only start emits VALUE=DATE', async () => {
    const allDayIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday@fm',
      'DTSTART;VALUE=DATE:20260401',
      // Runs to April 6 so the new April 5 start still precedes the untouched DTEND.
      'DTEND;VALUE=DATE:20260406',
      'SUMMARY:All Day',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: allDayIcal, url: '/cal/allday.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('allday@fm', { start: '2026-04-05' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DTSTART;VALUE=DATE:20260405'));
  });

  it('throws on event with no VEVENT', async () => {
    // Simulate an object found by URL but with no VEVENT block
    const noVeventIcal = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR';
    const objects = [{ data: noVeventIcal, url: '/cal/bad.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('/cal/bad.ics', { title: 'X' }),
      /no iCal data|not found/
    );
  });

  it('throws when VEVENT block is malformed (BEGIN without END)', async () => {
    // Has BEGIN:VEVENT (passes string check) but no END:VEVENT (extractVEvent returns null)
    const brokenIcal = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:broken@fm\r\nEND:VCALENDAR';
    const objects = [{ data: brokenIcal, url: '/cal/broken.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('/cal/broken.ics', { title: 'X' }),
      /not found|no VEVENT/
    );
  });

  it('throws on malformed start date format', async () => {
    const ical = makeRichIcal('evtbad@fm');
    const objects = [{ data: ical, url: '/cal/evtbad.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtbad@fm', { start: 'not-a-date' }),
      /Invalid start date format/
    );
  });

  it('throws on malformed end date format', async () => {
    const ical = makeRichIcal('evtbad2@fm');
    const objects = [{ data: ical, url: '/cal/evtbad2.ics' }];
    const { client } = createMockedPatchClient(objects);

    await assert.rejects(
      () => client.updateCalendarEvent('evtbad2@fm', { end: 'garbage' }),
      /Invalid end date format/
    );
  });

  it('throws when adding participants with non-email CalDAV username', async () => {
    const noOrganizerIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorg2@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Simple Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noOrganizerIcal, url: '/cal/noorg2.ics' }];
    // Use non-email username
    const client = new CalDAVCalendarClient({ username: 'not-an-email', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => objects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.updateCalendarEvent('noorg2@fm', {
        participants: [{ email: 'alice@example.com' }],
      }),
      /Invalid participant email: not-an-email/
    );
  });

  it('rejects a CalDAV username that would inject into the ORGANIZER line', async () => {
    // The update path validates the username with the same strict addr-spec check
    // as the create path. A bare .includes('@') test would accept this value and
    // emit the injected CRLF straight into the ORGANIZER line.
    const noOrganizerIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorg3@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Simple Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noOrganizerIcal, url: '/cal/noorg3.ics' }];
    const client = new CalDAVCalendarClient({
      username: 'me@example.com\r\nX-EVIL:1',
      password: 'test',
    });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => objects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.updateCalendarEvent('noorg3@fm', {
        participants: [{ email: 'alice@example.com' }],
      }),
      /Invalid participant email/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('adds ORGANIZER when adding participants to event with no existing ORGANIZER', async () => {
    const noOrganizerIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:noorg@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Simple Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const objects = [{ data: noOrganizerIcal, url: '/cal/noorg.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('noorg@fm', {
      participants: [{ email: 'alice@example.com', name: 'Alice' }],
    });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('ORGANIZER'));
    assert.ok(updatedData.includes('mailto:test@fastmail.com'));
    assert.ok(updatedData.includes('ATTENDEE'));
  });

  // The ORGANIZER-insertion gate is a PRESENCE test over the STORED payload, and a stored
  // payload can be authored by whoever sent the invitation. A `/m`-anchored test read the text
  // after a U+2028 inside a SUMMARY as a line of its own, so a forged ORGANIZER satisfied the
  // gate and the event was written with ATTENDEEs and no real ORGANIZER — a malformed
  // scheduling VEVENT (RFC 5545 §3.8.4.1) that no attendee's client can reply to.
  it('inserts a real ORGANIZER when the only ORGANIZER text is forged inside a SUMMARY', async () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:forgedorg@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Standup\u2028ORGANIZER;CN=Nobody:mailto:nobody@example.invalid',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedPatchClient([{ data: ical, url: '/cal/forgedorg.ics' }]);

    await client.updateCalendarEvent('forgedorg@fm', { participants: [{ email: 'carol@example.com' }] });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.match(updatedData, /(^|\r\n)ORGANIZER;CN=[^\r\n]*:mailto:test@fastmail\.com(\r\n|$)/);
    assert.ok(updatedData.includes('ATTENDEE;CN=' ) || updatedData.includes('ATTENDEE:mailto:carol@example.com'));
  });

  // The SEQUENCE-increment gate is the same shape of presence test, read the other way round:
  // a forged ATTENDEE made a solo event look like a scheduled one, so an ordinary location
  // edit bumped SEQUENCE and told every (non-existent) attendee's client the invitation had
  // been revised.
  it('does not increment SEQUENCE when the only ATTENDEE text is forged inside a DESCRIPTION', async () => {
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:forgedatt@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Solo Event',
      'DESCRIPTION:notes\u2028ATTENDEE;CN=Nobody:mailto:nobody@example.invalid',
      'SEQUENCE:0',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedPatchClient([{ data: ical, url: '/cal/forgedatt.ics' }]);

    await client.updateCalendarEvent('forgedatt@fm', { location: 'New Room' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:0'), updatedData);
    assert.ok(!updatedData.includes('SEQUENCE:1'), updatedData);
  });

  // RFC 5545 folding is deterministic at 75 octets, so an event's author can put a chosen fold
  // at a chosen point in a DESCRIPTION — including immediately before the text "BEGIN:". The
  // scan that picks where to INSERT a new property read that continuation head as a
  // sub-component and spliced LOCATION into the middle of the DESCRIPTION: the description
  // lost its tail and the inserted LOCATION swallowed it, two records damaged in one write
  // with nothing reported. Driven through the real write path, because the insert branch is
  // the ordinary case for adding a location to an event that has none.
  it('does not insert a property into the middle of a folded DESCRIPTION', async () => {
    const description = `${'A'.repeat(63)}BEGIN:phase two of the agenda`;
    const ical = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:folded@fm',
      'DTSTART:20260401T100000Z',
      'DTEND:20260401T110000Z',
      'SUMMARY:Planning',
      foldICalLine(`DESCRIPTION:${description}`),
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    // The fixture only bites if the fold really lands with BEGIN: at the head of the
    // continuation — that alignment is the whole test, so it is asserted rather than assumed.
    assert.ok(ical.includes('\r\n BEGIN:phase two of the agenda'), ical);

    const { client, mockDAVClient } = createMockedPatchClient([{ data: ical, url: '/cal/folded.ics' }]);
    await client.updateCalendarEvent('folded@fm', { location: 'Room 4' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    const vevent = extractVEvent(updatedData)!;
    // BOTH records: the stored one survives whole, and the new one carries only what it was
    // given. The corruption swapped a tail from the first into the second.
    assert.equal(parseICalValue(vevent, 'DESCRIPTION'), description);
    assert.equal(parseICalValue(vevent, 'LOCATION'), 'Room 4');
  });

  it('updates description only', async () => {
    const ical = makeRichIcal('evtdesc@fm');
    const objects = [{ data: ical, url: '/cal/evtdesc.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evtdesc@fm', { description: 'New description' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DESCRIPTION:New description'));
    assert.ok(updatedData.includes('SUMMARY:Original Title')); // Other fields preserved
  });

  it('updates DTSTAMP and LAST-MODIFIED', async () => {
    const ical = makeRichIcal('evt7@fm');
    const objects = [{ data: ical, url: '/cal/evt7.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt7@fm', { title: 'X' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('LAST-MODIFIED:'));
    // DTSTAMP should be updated (not the original)
    assert.ok(!updatedData.includes('DTSTAMP:20260401T000000Z'));
  });

  it('preserves Google PRODID (does not stamp ours)', async () => {
    const ical = makeRichIcal('evt8@fm');
    const objects = [{ data: ical, url: '/cal/evt8.ics' }];
    const { client, mockDAVClient } = createMockedPatchClient(objects);

    await client.updateCalendarEvent('evt8@fm', { title: 'X' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('PRODID:-//Google Inc//Google Calendar//EN'));
  });
});

describe('update_calendar_event / delete_calendar_event refuse a recurring series', () => {
  // A repeating event is refused OUTRIGHT by both write tools, with no override parameter.
  // The reason is the create surface: create_calendar_event has no RRULE parameter, so this
  // server cannot make a repeating event, and under the project rule that a destroy must not
  // remove what this server cannot recreate it must not destroy or rewrite one either. A
  // delete takes every occurrence past and future and mails every attendee a cancellation; an
  // update moves every occurrence, and for a series carrying RECURRENCE-ID overrides RFC 5545
  // does not settle what should happen to them. Restoring the capability properly is #146.
  function makeRecurringIcal(exceptions: Array<{ recurrenceId: string; summary: string }> = []): string {
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VEVENT',
      'UID:recur@fm',
      'DTSTAMP:20260401T000000Z',
      'DTSTART:20260406T100000Z',
      'DTEND:20260406T110000Z',
      'RRULE:FREQ=WEEKLY;COUNT=10',
      'SUMMARY:Weekly Meeting',
      'SEQUENCE:0',
      'END:VEVENT',
    ];
    for (const exc of exceptions) {
      lines.push(
        'BEGIN:VEVENT',
        'UID:recur@fm',
        `RECURRENCE-ID:${exc.recurrenceId}`,
        'DTSTAMP:20260401T000000Z',
        `DTSTART:${exc.recurrenceId}`,
        `DTEND:${exc.recurrenceId.replace('T10', 'T11')}`,
        `SUMMARY:${exc.summary}`,
        'END:VEVENT',
      );
    }
    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }

  // A series whose master has been removed: every block is a detached override. Still one
  // series, and deleting the file still takes every occurrence with it — which is why the
  // detector cannot key on RRULE alone.
  const ALL_OVERRIDES_ICAL = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'BEGIN:VEVENT',
    'UID:orphaned@fm',
    'RECURRENCE-ID:20260406T100000Z',
    'DTSTAMP:20260401T000000Z',
    'DTSTART:20260406T100000Z',
    'DTEND:20260406T110000Z',
    'SUMMARY:Standup',
    'END:VEVENT',
    'BEGIN:VEVENT',
    'UID:orphaned@fm',
    'RECURRENCE-ID:20260413T100000Z',
    'DTSTAMP:20260401T000000Z',
    'DTSTART:20260413T110000Z',
    'DTEND:20260413T120000Z',
    'SUMMARY:Standup',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  const SINGLE_ICAL = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'BEGIN:VEVENT',
    'UID:single@fm',
    'DTSTAMP:20260401T000000Z',
    'DTSTART:20260406T100000Z',
    'DTEND:20260406T110000Z',
    'SUMMARY:One-off Meeting',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  function createMockedRecurringClient(ical: string, url = '/cal/recur.ics') {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const objects = [{ data: ical, url }];
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => objects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
      deleteCalendarObject: mock.fn(async (_params: any) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  const isInvalidInput = (err: unknown) => err instanceof Error && err.name === 'InvalidInputError';

  it('refuses to update a series, naming it and denying that any flag exists', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(makeRecurringIcal());

    await assert.rejects(
      () => client.updateCalendarEvent('recur@fm', {
        start: '2026-04-07T10:00:00Z',
        end: '2026-04-07T11:00:00Z',
      }),
      (err: unknown) => {
        // InvalidInputError, not a plain Error: the caller fixes this by not making the call,
        // which is InvalidParams territory. InternalError would say a bare retry might work.
        assert.ok(isInvalidInput(err), `expected InvalidInputError, got ${err}`);
        const message = (err as Error).message;
        assert.match(message, /"Weekly Meeting"/);
        assert.match(message, /repeating event/);
        assert.match(message, /no parameter, flag or confirmation that overrides this/i);
        assert.match(message, /Fastmail web interface/);
        return true;
      },
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('refuses to delete a series without writing anything', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(makeRecurringIcal());

    await assert.rejects(
      () => client.deleteCalendarEvent('recur@fm'),
      (err: unknown) => {
        assert.ok(isInvalidInput(err), `expected InvalidInputError, got ${err}`);
        const message = (err as Error).message;
        assert.match(message, /"Weekly Meeting"/);
        assert.match(message, /every occurrence, past and future/);
        assert.match(message, /cancellation to every attendee/);
        return true;
      },
    );
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  // findCalendarObjectByUID matches `obj.url` interchangeably with the UID, and `url` is on
  // every row this server serialises — so a guard keyed on the argument rather than on the
  // resolved resource would be bypassed by passing the row's url instead of its id.
  it('refuses the same series when the id given is the row url', async () => {
    for (const eventId of ['recur@fm', '/cal/recur.ics']) {
      const { client, mockDAVClient } = createMockedRecurringClient(makeRecurringIcal());
      await assert.rejects(() => client.deleteCalendarEvent(eventId), isInvalidInput);
      await assert.rejects(() => client.updateCalendarEvent(eventId, { title: 'New' }), isInvalidInput);
      assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    }
  });

  // No RRULE anywhere in this resource, and it is still a series.
  it('refuses a resource that is entirely RECURRENCE-ID overrides with no master', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(ALL_OVERRIDES_ICAL, '/cal/orphaned.ics');
    await assert.rejects(() => client.deleteCalendarEvent('orphaned@fm'), isInvalidInput);
    await assert.rejects(() => client.updateCalendarEvent('orphaned@fm', { title: 'New' }), isInvalidInput);
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  // The refusal covers every field, not just the ones that move dates: patching the master's
  // title rewrites the title of every occurrence.
  it('refuses a title-only update, not just a time change', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(
      makeRecurringIcal([{ recurrenceId: '20260413T100000Z', summary: 'Exception Week 2' }]),
    );
    await assert.rejects(() => client.updateCalendarEvent('recur@fm', { title: 'New Title' }), isInvalidInput);
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  // Raised before argument validation, because no argument can make the call legal: reporting
  // the bad date first would imply that fixing it would help.
  it('refuses before it complains about a malformed date', async () => {
    const { client } = createMockedRecurringClient(makeRecurringIcal());
    await assert.rejects(
      () => client.updateCalendarEvent('recur@fm', { start: 'not-a-date' }),
      (err: unknown) => {
        assert.match((err as Error).message, /repeating event/);
        return true;
      },
    );
  });

  // A single block whose only recurrence marker is a RECURRENCE-ID: an overridden occurrence
  // with no rule beside it. Malformed as a whole resource, and still a recurrence — the
  // detector fails CLOSED because it gates an irreversible write.
  it('refuses a lone RECURRENCE-ID block with no RRULE anywhere', async () => {
    const loneOverride = [
      'BEGIN:VCALENDAR', 'VERSION:2.0',
      'BEGIN:VEVENT',
      'UID:lone@fm', 'RECURRENCE-ID:20260413T100000Z', 'DTSTAMP:20260401T000000Z',
      'DTSTART:20260413T100000Z', 'DTEND:20260413T110000Z', 'SUMMARY:Detached',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedRecurringClient(loneOverride, '/cal/lone.ics');
    await assert.rejects(() => client.deleteCalendarEvent('lone@fm'), isInvalidInput);
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  // A series that LISTS its occurrences as RDATEs instead of stating a rule (RFC 5545
  // §3.8.5.2). One block, no RRULE, no RECURRENCE-ID — so an RRULE-only detector read it as an
  // ordinary one-off and delete_calendar_event destroyed all four occurrences and mailed every
  // attendee a cancellation. The read path already calls it recurring and emits its
  // recurrenceDates, and create_calendar_event has no RDATE parameter, so it is exactly the
  // kind of record this refusal exists for (#162).
  const RDATE_ONLY_ICAL = [
    'BEGIN:VCALENDAR', 'VERSION:2.0',
    'BEGIN:VEVENT',
    'UID:rdates@fm', 'DTSTAMP:20260401T000000Z',
    'DTSTART:20260406T100000Z', 'DTEND:20260406T110000Z',
    'RDATE:20260413T100000Z,20260420T100000Z',
    'RDATE:20260427T100000Z',
    'SUMMARY:Irregular Catch-up',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  it('refuses to DELETE a series that recurs only by RDATE', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(RDATE_ONLY_ICAL, '/cal/rdates.ics');
    await assert.rejects(
      () => client.deleteCalendarEvent('rdates@fm'),
      (err: unknown) => {
        assert.ok(isInvalidInput(err), `expected InvalidInputError, got ${err}`);
        const message = (err as Error).message;
        assert.match(message, /"Irregular Catch-up"/);
        assert.match(message, /every occurrence, past and future/);
        assert.match(message, /cancellation to every attendee/);
        return true;
      },
    );
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  it('refuses to UPDATE a series that recurs only by RDATE', async () => {
    const { client, mockDAVClient } = createMockedRecurringClient(RDATE_ONLY_ICAL, '/cal/rdates.ics');
    await assert.rejects(
      () => client.updateCalendarEvent('rdates@fm', { title: 'New Title' }),
      (err: unknown) => {
        assert.ok(isInvalidInput(err), `expected InvalidInputError, got ${err}`);
        assert.match((err as Error).message, /repeating event/);
        return true;
      },
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  // Two VEVENT blocks in one resource can only be a recurrence — one CalDAV resource is one
  // UID — so the count alone refuses, even with no RRULE and no RECURRENCE-ID to read. This is
  // the same test the read path uses to set isRecurring, on purpose.
  it('refuses a resource holding more than one VEVENT, with no recurrence markers at all', async () => {
    const twoBlocks = [
      'BEGIN:VCALENDAR', 'VERSION:2.0',
      'BEGIN:VEVENT',
      'UID:two@fm', 'DTSTAMP:20260401T000000Z',
      'DTSTART:20260406T100000Z', 'DTEND:20260406T110000Z', 'SUMMARY:Standup',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:two@fm', 'DTSTAMP:20260401T000000Z',
      'DTSTART:20260413T100000Z', 'DTEND:20260413T110000Z', 'SUMMARY:Standup',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedRecurringClient(twoBlocks, '/cal/two.ics');
    await assert.rejects(() => client.deleteCalendarEvent('two@fm'), isInvalidInput);
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  // The detector reads whole content lines, never a /m-anchored regex over the raw payload.
  // U+2028, U+2029 and a bare CR all satisfy JavaScript's /m anchors and are legal, unescaped
  // characters inside an iCalendar TEXT value — so a caller-authored SUMMARY carrying one,
  // followed by something that reads as an RRULE line, would turn a one-off event into a
  // permanent refusal under a regex. Calendar text is attacker-authored here: anyone who can
  // send an invitation writes it.
  it('does not treat an RRULE forged inside a SUMMARY as a recurrence', async () => {
    for (const sep of ['\u2028', '\u2029', '\r']) {
      const forged = SINGLE_ICAL.replace(
        'SUMMARY:One-off Meeting',
        `SUMMARY:One-off Meeting${sep}RRULE:FREQ=WEEKLY`,
      );
      const { client, mockDAVClient } = createMockedRecurringClient(forged, '/cal/single.ics');
      await client.updateCalendarEvent('single@fm', { title: 'Renamed' });
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1, `separator ${JSON.stringify(sep)}`);
    }
    // …and the escaped form an honest client would write, which is not a line break at all.
    const escaped = SINGLE_ICAL.replace(
      'SUMMARY:One-off Meeting',
      'SUMMARY:One-off Meeting\\nRRULE:FREQ=WEEKLY',
    );
    const { client, mockDAVClient } = createMockedRecurringClient(escaped, '/cal/single.ics');
    await client.updateCalendarEvent('single@fm', { title: 'Renamed' });
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
  });

  // The same forgery for RDATE, the second recurrence marker isRecurringSeriesResource reads
  // (#162, RECURRENCE-ID being the third), driven through both write tools: a forged RDATE must
  // not refuse a legitimate edit or delete of a one-off event, and the delete is the irreversible
  // half this gate fronts. Separators written as escapes on purpose — a literal U+2028 or U+2029
  // here is invisible in an editor — and the braced code-point form below is deliberate, not a
  // drifted spelling of the plain form used above.
  const FORGED_RDATE_SEPARATORS = ['\u{2028}', '\u{2029}', '\r'];
  const forgedRdateIcal = (sep: string) => SINGLE_ICAL.replace(
    'SUMMARY:One-off Meeting',
    `SUMMARY:One-off Meeting${sep}RDATE:20260413T100000Z`,
  );

  it('does not treat an RDATE forged inside a SUMMARY as a recurrence', async () => {
    for (const sep of FORGED_RDATE_SEPARATORS) {
      const { client, mockDAVClient } = createMockedRecurringClient(forgedRdateIcal(sep), '/cal/single.ics');
      await client.updateCalendarEvent('single@fm', { title: 'Renamed' });
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1, `separator ${JSON.stringify(sep)}`);
    }
    // …and the escaped form an honest client would write, which is not a line break at all.
    const { client, mockDAVClient } = createMockedRecurringClient(forgedRdateIcal('\\n'), '/cal/single.ics');
    await client.updateCalendarEvent('single@fm', { title: 'Renamed' });
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
  });

  it('deletes a one-off event whose SUMMARY carries a forged RDATE', async () => {
    for (const sep of FORGED_RDATE_SEPARATORS) {
      const { client, mockDAVClient } = createMockedRecurringClient(forgedRdateIcal(sep), '/cal/single.ics');
      await client.deleteCalendarEvent('single@fm');
      assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 1, `separator ${JSON.stringify(sep)}`);
    }
  });

  // RFC 5545 §3.1 makes property names case-INSENSITIVE. libical re-serialises them upper
  // case, so this shape does not come from Fastmail's own client — but it is legal, a third
  // party can PUT it, and reading it as "no rule" would let the delete destroy the series.
  it('refuses a series whose RRULE is spelled in lower case', async () => {
    const lower = makeRecurringIcal().replace('RRULE:FREQ=WEEKLY;COUNT=10', 'rrule:FREQ=WEEKLY;COUNT=10');
    assert.ok(!lower.includes('RRULE:'), 'the premise: no upper-case RRULE is left');
    const { client, mockDAVClient } = createMockedRecurringClient(lower);
    await assert.rejects(() => client.deleteCalendarEvent('recur@fm'), isInvalidInput);
    await assert.rejects(() => client.updateCalendarEvent('recur@fm', { title: 'New' }), isInvalidInput);
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  // A fold placed inside the property NAME. Legal after RFC 5545 unfolding, and never emitted
  // by libical (it folds at a fixed octet count, not at offset 3), so this takes a hand-crafted
  // PUT — but the gate must unfold rather than skip the continuation, or the rule hides and the
  // delete proceeds. The direction is what matters: skipping fails OPEN in front of a destroy.
  it('refuses a series whose RRULE name is split across a fold', async () => {
    const folded = makeRecurringIcal().replace('RRULE:FREQ=WEEKLY;COUNT=10', 'RRU\r\n LE:FREQ=WEEKLY;COUNT=10');
    assert.ok(!folded.includes('RRULE:'), 'the premise: the name is not contiguous');
    const { client, mockDAVClient } = createMockedRecurringClient(folded);
    await assert.rejects(() => client.deleteCalendarEvent('recur@fm'), isInvalidInput);
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  // Only hasICalProperty was made case-insensitive; the structural scan and parseICalValue
  // were not, and this is why that is safe rather than half a fix. A payload whose STRUCTURAL
  // keywords are lower-cased yields no VEVENT blocks, so findCalendarObjectByUID skips the
  // object before it reads a single value: the event is invisible to every tool rather than
  // reachable through a mis-read. Fail-closed, and pinned so the reasoning stays checkable.
  it('cannot reach an event whose BEGIN:VEVENT is lower-cased at all', async () => {
    const lower = [
      'begin:vcalendar', 'version:2.0',
      'begin:vevent',
      'uid:low@fm', 'dtstamp:20260401T000000Z',
      'dtstart:20260406T100000Z', 'dtend:20260406T110000Z',
      'rrule:FREQ=WEEKLY;COUNT=10', 'summary:Weekly',
      'end:vevent',
      'end:vcalendar',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedRecurringClient(lower, '/cal/low.ics');
    for (const eventId of ['low@fm', '/cal/low.ics']) {
      await assert.rejects(
        () => client.deleteCalendarEvent(eventId),
        (err: unknown) => {
          assert.ok(isInvalidInput(err), `expected InvalidInputError, got ${err}`);
          assert.match((err as Error).message, /not found/);
          return true;
        },
      );
    }
    assert.equal(mockDAVClient.deleteCalendarObject.mock.calls.length, 0);
  });

  // Unfolding must not re-open the hole that skipping continuations was closing: a property
  // forged at the head of a continuation line is joined to the line ABOVE it, so it can never
  // begin a logical line and cannot fabricate a recurrence on a one-off event.
  it('does not read an RRULE forged at the head of a continuation line', async () => {
    const injected = SINGLE_ICAL.replace(
      'SUMMARY:One-off Meeting',
      'SUMMARY:One-off Meeting\r\n RRULE:FREQ=WEEKLY',
    );
    const { client, mockDAVClient } = createMockedRecurringClient(injected, '/cal/single.ics');
    await client.updateCalendarEvent('single@fm', { title: 'Renamed' });
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 1);
  });

  it('leaves single events alone: update and delete both proceed', async () => {
    const forUpdate = createMockedRecurringClient(SINGLE_ICAL, '/cal/single.ics');
    await forUpdate.client.updateCalendarEvent('single@fm', {
      start: '2026-04-07T10:00:00Z',
      end: '2026-04-07T11:00:00Z',
    });
    assert.equal(forUpdate.mockDAVClient.updateCalendarObject.mock.calls.length, 1);
    const updatedData = callArguments(forUpdate.mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('DTSTART:20260407T100000Z'));

    const forDelete = createMockedRecurringClient(SINGLE_ICAL, '/cal/single.ics');
    await forDelete.client.deleteCalendarEvent('single@fm');
    assert.equal(forDelete.mockDAVClient.deleteCalendarObject.mock.calls.length, 1);
  });

  // The resource is what decides, so a series with no SUMMARY still refuses — it just names
  // the id the caller passed instead of a title.
  it('falls back to the event id when the series has no title', async () => {
    const untitled = makeRecurringIcal().replace('SUMMARY:Weekly Meeting\r\n', '');
    const { client } = createMockedRecurringClient(untitled);
    await assert.rejects(
      () => client.deleteCalendarEvent('recur@fm'),
      (err: unknown) => {
        assert.match((err as Error).message, /"recur@fm" is a repeating event/);
        return true;
      },
    );
  });
});

describe('CalDAVCalendarClient.createCalendarEvent with participants', () => {
  function createMockedCreateClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('includes ATTENDEE and ORGANIZER lines with mailto URI', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [
        { email: 'alice@example.com', name: 'Alice' },
        { email: 'bob@example.com', name: 'Bob' },
      ],
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes(':mailto:me@fastmail.com'), 'ORGANIZER should have mailto URI');
    assert.ok(/ORGANIZER;CN=.+:mailto:me@fastmail.com/.test(ical), 'ORGANIZER should have CN parameter');
    assert.ok(ical.includes('ATTENDEE;CN=Alice:mailto:alice@example.com'));
    assert.ok(ical.includes('ATTENDEE;CN=Bob:mailto:bob@example.com'));
  });

  it('classes an unknown calendarId as caller-fixable input', async () => {
    // The caller fixes this by re-sending with an id or name from list_calendars, so it
    // must arrive as InvalidParams rather than the InternalError a plain Error maps to.
    const { client, mockDAVClient } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'No Such Calendar',
        title: 'Meeting',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
      }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Calendar not found: "No Such Calendar"/);
        return true;
      },
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
  });

  it('classes a participant email the caller supplied as caller-fixable input', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Meeting',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
        participants: [{ email: 'not-an-email' }],
      }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Invalid participant email/);
        return true;
      },
    );
  });

  it('keeps an unusable configured CalDAV username a server-side failure', async () => {
    // The ORGANIZER address is server configuration, not part of the tool call, so it
    // reuses the participant address rules but NOT the caller-fixable class: telling the
    // caller to re-form their arguments would send them after something they cannot change.
    const client = new CalDAVCalendarClient({ username: 'not-an-email', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Meeting',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
        participants: [{ email: 'alice@example.com' }],
      }),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /Invalid participant email: not-an-email/);
        return true;
      },
    );
  });

  it('does not include ATTENDEE/ORGANIZER when no participants', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Solo Event',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(!ical.includes('ATTENDEE'));
    assert.ok(!ical.includes('ORGANIZER'));
  });

  it('does not emit RSVP=TRUE', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [{ email: 'alice@example.com' }],
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(!ical.includes('RSVP'));
  });

  it('rejects email with injection attempt', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Meeting',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
        participants: [{ email: 'a@b.example\r\nX-INJECT:true' }],
      }),
      /illegal/i
    );
  });

  it('uses DQUOTE quoting for CN, not backslash escaping', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [{ email: 'alice@example.com', name: 'Doe, Alice' }],
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    // Should use DQUOTE quoting: CN="Doe, Alice"
    assert.ok(ical.includes('CN="Doe, Alice"'));
    // Should NOT use backslash escaping: CN=Doe\, Alice
    assert.ok(!ical.includes('CN=Doe\\, Alice'));
  });

  it('handles date-only start/end (all-day event)', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'All Day',
      start: '2026-04-07',
      end: '2026-04-08',
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('DTSTART;VALUE=DATE:20260407'));
    assert.ok(ical.includes('DTEND;VALUE=DATE:20260408'));
  });

  it('rejects date-only start and end with same value', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'Personal',
        title: 'Bad',
        start: '2026-04-07',
        end: '2026-04-07',
      }),
      /DTEND is exclusive/
    );
  });

  it('ends with trailing CRLF', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Test',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.endsWith('\r\n'));
  });
});

describe('CRLF vs LF line ending preservation', () => {
  it('replaceICalProperty preserves LF-only endings', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(!result.includes('\r\n'));
    assert.ok(result.includes('SUMMARY:New'));
  });

  it('replaceICalProperty preserves CRLF endings', () => {
    const event = 'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nSUMMARY:Old\r\nEND:VEVENT\r\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(result.includes('\r\n'));
    assert.ok(result.includes('SUMMARY:New'));
  });

  it('no mixed line endings when patching LF-only input', () => {
    const event = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(event, 'SUMMARY', 'SUMMARY:New');
    assert.ok(!result.includes('\r\n'), 'Should not contain CRLF in LF-only input');
  });

  it('foldICalLine with wrong lineEnding produces consistent output', () => {
    // Simulate a caller using CRLF fold on what will be inserted into LF document
    const folded = foldICalLine('DESCRIPTION:' + 'x'.repeat(80), '\r\n');
    // The fold itself should use CRLF consistently
    assert.ok(folded.includes('\r\n'));
    // When replaceICalProperty re-splits and re-joins with LF, CRLF folds are preserved
    // inside the replacement line — this is the caller's responsibility to match
    const lfEvent = 'BEGIN:VCALENDAR\nBEGIN:VEVENT\nSUMMARY:Old\nEND:VEVENT\nEND:VCALENDAR';
    const result = replaceICalProperty(lfEvent, 'SUMMARY', foldICalLine('SUMMARY:Short', '\n'));
    assert.ok(!result.includes('\r\n'), 'Caller using correct lineEnding prevents mixing');
  });
});

describe('VEVENT extraction consistency', () => {
  it('extractVEvent and replaceICalProperty agree on VEVENT boundaries', () => {
    const multiVevent = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Europe/Rome',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'SUMMARY:Master',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:master@fm',
      'RECURRENCE-ID:20260408T100000Z',
      'SUMMARY:Exception',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');

    const extracted = extractVEvent(multiVevent)!;
    assert.ok(extracted.includes('SUMMARY:Master'));
    assert.ok(!extracted.includes('SUMMARY:Exception'));

    // replaceICalProperty should operate on the same first VEVENT
    const patched = replaceICalProperty(multiVevent, 'SUMMARY', 'SUMMARY:Updated');
    assert.ok(patched.includes('SUMMARY:Updated'));
    assert.ok(patched.includes('SUMMARY:Exception')); // Second VEVENT untouched
  });
});

describe('Additional plan-required updateCalendarEvent tests', () => {
  function makeRichIcal(uid: string, extra: string[] = []): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Google Inc//Google Calendar//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTAMP:20260401T000000Z',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DTEND;TZID=Europe/Rome:20260401T110000',
      'SUMMARY:Original Title',
      'DESCRIPTION:Original description',
      'LOCATION:Room A',
      'ATTENDEE;CN=Alice;PARTSTAT=ACCEPTED:mailto:alice@example.com',
      'ORGANIZER;CN=Boss:mailto:boss@example.com',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'ACTION:DISPLAY',
      'DESCRIPTION:Reminder',
      'END:VALARM',
      'SEQUENCE:2',
      ...extra,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  function createMockedClient(calendarObjects: Array<{ data: string; url: string }>) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => calendarObjects),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('increments SEQUENCE for start change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqstart@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqstart.ics' }]);

    // Same day as the untouched DTEND (Europe/Rome 11:00) and before it.
    await client.updateCalendarEvent('seqstart@fm', { start: '2026-04-01T10:30:00' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('increments SEQUENCE for end change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqend@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqend.ics' }]);

    await client.updateCalendarEvent('seqend@fm', { end: '2026-04-01T12:00:00' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('increments SEQUENCE for participants change when ATTENDEEs exist', async () => {
    const ical = makeRichIcal('seqpart@fm');
    const { client, mockDAVClient } = createMockedClient([{ data: ical, url: '/cal/seqpart.ics' }]);

    await client.updateCalendarEvent('seqpart@fm', {
      participants: [{ email: 'new@example.com', name: 'New' }],
    });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(updatedData.includes('SEQUENCE:3'));
  });

  it('floating end time falls back to DTSTART TZID when DTEND was DURATION-computed', async () => {
    const durationIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:durtz@fm',
      'DTSTART;TZID=Europe/Rome:20260401T100000',
      'DURATION:PT2H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = createMockedClient([{ data: durationIcal, url: '/cal/durtz.ics' }]);

    await client.updateCalendarEvent('durtz@fm', { end: '2026-04-01T13:00:00' });

    const updatedData = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    // Should fall back to DTSTART's TZID since there was no DTEND
    assert.ok(updatedData.includes('DTEND;TZID=Europe/Rome:20260401T130000'));
  });

  it('rejects date-only start and end with same value on update', async () => {
    const allDayIcal = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday2@fm',
      'DTSTART;VALUE=DATE:20260401',
      'DTEND;VALUE=DATE:20260402',
      'SUMMARY:All Day',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = createMockedClient([{ data: allDayIcal, url: '/cal/allday2.ics' }]);

    await assert.rejects(
      () => client.updateCalendarEvent('allday2@fm', { start: '2026-04-05', end: '2026-04-05' }),
      /DTEND is exclusive/
    );
  });
});

// ---------- v1.11.0 review fixes ----------

describe('escapeICalText control-character hardening', () => {
  it('escapes a bare CR as \\n instead of passing it through', () => {
    assert.equal(escapeICalText('Standup\rATTENDEE:mailto:x@example.com'),
      'Standup\\nATTENDEE:mailto:x@example.com');
  });

  it('still escapes CRLF and LF as \\n', () => {
    assert.equal(escapeICalText('a\r\nb\nc'), 'a\\nb\\nc');
  });

  it('strips other control characters', () => {
    assert.equal(escapeICalText('a\x00b\x08c\x7Fd'), 'abcd');
  });

  it('keeps horizontal tabs (legal in iCal TEXT)', () => {
    assert.equal(escapeICalText('a\tb'), 'a\tb');
  });
});

describe('parseICalDateAsUTC', () => {
  it('interprets naive datetimes as UTC regardless of process TZ', () => {
    const d = parseICalDateAsUTC('2026-03-20T09:30:00');
    assert.equal(d.getTime(), Date.UTC(2026, 2, 20, 9, 30, 0));
  });

  it('handles explicit Z', () => {
    assert.equal(parseICalDateAsUTC('2026-03-20T09:30:00Z').getTime(), Date.UTC(2026, 2, 20, 9, 30, 0));
  });

  it('handles offsets', () => {
    assert.equal(parseICalDateAsUTC('2026-03-20T10:30:00+01:00').getTime(), Date.UTC(2026, 2, 20, 9, 30, 0));
  });

  it('handles date-only as UTC midnight', () => {
    assert.equal(parseICalDateAsUTC('2026-03-20').getTime(), Date.UTC(2026, 2, 20));
  });
});

describe('normalizeMasterVEventFirst', () => {
  const exception = 'BEGIN:VEVENT\nUID:u1\nRECURRENCE-ID:20260327T093000Z\nDTSTART:20260327T110000Z\nSUMMARY:Moved instance\nEND:VEVENT';
  const master = 'BEGIN:VEVENT\nUID:u1\nDTSTART:20260320T093000Z\nRRULE:FREQ=WEEKLY\nSUMMARY:Weekly\nEND:VEVENT';

  it('moves the master VEVENT ahead of an exception-first ordering', () => {
    const data = `BEGIN:VCALENDAR\n${exception}\n${master}\nEND:VCALENDAR`;
    const out = normalizeMasterVEventFirst(data);
    const firstVevent = out.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/)?.[0] || '';
    assert.ok(/^RRULE/m.test(firstVevent), 'master (RRULE, no RECURRENCE-ID) should now be first');
    assert.ok(out.includes('Moved instance'), 'exception must be preserved');
  });

  it('leaves master-first payloads untouched', () => {
    const data = `BEGIN:VCALENDAR\n${master}\n${exception}\nEND:VCALENDAR`;
    assert.equal(normalizeMasterVEventFirst(data), data);
  });

  it('leaves single-VEVENT payloads untouched', () => {
    const data = `BEGIN:VCALENDAR\n${master}\nEND:VCALENDAR`;
    assert.equal(normalizeMasterVEventFirst(data), data);
  });
});

describe('replaceICalProperty insert position with VALARM', () => {
  it('inserts a new property before the first sub-component, not after it', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:u1',
      'DTSTART:20260320T093000Z',
      'BEGIN:VALARM',
      'TRIGGER:-PT15M',
      'END:VALARM',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');
    const out = replaceICalProperty(data, 'DESCRIPTION', 'DESCRIPTION:hello');
    const descIdx = out.indexOf('DESCRIPTION:hello');
    const alarmIdx = out.indexOf('BEGIN:VALARM');
    assert.ok(descIdx !== -1 && alarmIdx !== -1);
    assert.ok(descIdx < alarmIdx, 'property must precede VALARM per RFC 5545 ABNF');
  });
});

describe('removeOrphanedVTimezones quoted/folded references', () => {
  it('keeps a VTIMEZONE referenced via a quoted TZID parameter', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Custom/Zone',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:u1',
      'DTSTART;TZID="Custom/Zone":20260320T093000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');
    const out = removeOrphanedVTimezones(data);
    assert.ok(out.includes('BEGIN:VTIMEZONE'), 'referenced VTIMEZONE must not be removed');
  });

  it('keeps a VTIMEZONE whose reference is split across a folded line', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:America/Argentina/ComodRivadavia',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:u1',
      'DTSTART;TZID=America/Argentina/Comod',
      ' Rivadavia:20260320T093000',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');
    const out = removeOrphanedVTimezones(data);
    assert.ok(out.includes('BEGIN:VTIMEZONE'), 'folded reference must still count');
  });

  it('still removes a genuinely orphaned VTIMEZONE', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Unused/Zone',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:u1',
      'DTSTART:20260320T093000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');
    const out = removeOrphanedVTimezones(data);
    assert.ok(!out.includes('BEGIN:VTIMEZONE'));
  });
});

// ---------- v1.11.1 security fixes ----------

describe('updateCalendarEvent — a hostile recurrence rule is never expanded', () => {
  function mockClient(icalData: string) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: icalData, url: '/cal/dos.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 207 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('refuses a sub-daily RRULE with a distant exception instead of expanding it', async () => {
    // FREQ=SECONDLY plus an exception decades out is the shape that forced billions of
    // occurrence iterations when this path expanded the rule to find orphaned exceptions.
    // Calendar events can be authored by third parties (an invitation), so the input is
    // hostile-capable. Refusing every repeating event retires the whole class: the rule is
    // never expanded on the write path at all, and the refusal is raised from a line-model
    // read that does no arithmetic.
    const ical = [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//t//t//EN',
      'BEGIN:VEVENT', 'UID:dos@fm', 'DTSTAMP:20260101T000000Z',
      'DTSTART:20260101T090000Z', 'DTEND:20260101T093000Z',
      'RRULE:FREQ=SECONDLY;INTERVAL=1', 'SUMMARY:Rapid',
      'END:VEVENT',
      'BEGIN:VEVENT', 'UID:dos@fm', 'RECURRENCE-ID:20990101T090000Z',
      'DTSTART:20990101T090000Z', 'DTEND:20990101T093000Z', 'SUMMARY:Far exception',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = mockClient(ical);
    const started = Date.now();
    await assert.rejects(
      () => client.updateCalendarEvent('dos@fm', { start: '2026-01-01T09:10:00Z' }),
      (err: unknown) => {
        assert.ok(err instanceof Error && err.name === 'InvalidInputError', `got ${err}`);
        assert.match(err.message, /repeating event/);
        return true;
      },
    );
    assert.ok(Date.now() - started < 3000, 'the rule must not be expanded at all');
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });
});

describe('CalDAV write status checking (assertDavOk)', () => {
  function mockClientWithStatus(status: number) {
    const ical = [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//t//t//EN',
      'BEGIN:VEVENT', 'UID:s@fm', 'DTSTAMP:20260101T000000Z',
      'DTSTART:20260101T090000Z', 'DTEND:20260101T093000Z', 'SUMMARY:S',
      'END:VEVENT', 'END:VCALENDAR',
    ].join('\r\n');
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: ical, url: '/cal/s.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status })),
      deleteCalendarObject: mock.fn(async (_params: DeleteObjectParams) => ({ status })),
    };
    return client;
  }

  it('throws when the server returns a 4xx/5xx on update', async () => {
    const client = mockClientWithStatus(500);
    await assert.rejects(
      () => client.updateCalendarEvent('s@fm', { title: 'X' }),
      /Failed to update calendar event: server returned 500/,
    );
  });

  it('throws when the server returns a 4xx on delete', async () => {
    const client = mockClientWithStatus(403);
    await assert.rejects(
      () => client.deleteCalendarEvent('s@fm'),
      /Failed to delete calendar event: server returned 403/,
    );
  });

  it('succeeds on a 2xx status', async () => {
    const client = mockClientWithStatus(204);
    await client.updateCalendarEvent('s@fm', { title: 'X' });
  });

  // #102: a response that never carried a numeric status used to be read as success (the
  // "older tsdav shapes" reading) — including a bare `{ ok: true }` with no status at all. A
  // write whose outcome nothing confirmed must not be reported as success.
  function mockClientWithResponse(response: unknown) {
    const ical = [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//t//t//EN',
      'BEGIN:VEVENT', 'UID:s@fm', 'DTSTAMP:20260101T000000Z',
      'DTSTART:20260101T090000Z', 'DTEND:20260101T093000Z', 'SUMMARY:S',
      'END:VEVENT', 'END:VCALENDAR',
    ].join('\r\n');
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: ical, url: '/cal/s.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => response),
      deleteCalendarObject: mock.fn(async (_params: DeleteObjectParams) => response),
    };
    return client;
  }

  it('throws on a response with no status at all', async () => {
    const client = mockClientWithResponse({});
    await assert.rejects(
      () => client.updateCalendarEvent('s@fm', { title: 'X' }),
      /Failed to update calendar event: server returned no status/,
    );
  });

  it('throws on a bare `{ ok: true }` response with no status', async () => {
    const client = mockClientWithResponse({ ok: true });
    await assert.rejects(
      () => client.deleteCalendarEvent('s@fm'),
      /Failed to delete calendar event: server returned no status/,
    );
  });
});

describe('resolveDisplayName', () => {
  it('uses the env value when it is a real string', () => {
    assert.equal(resolveDisplayName('Jeremy G', 'fallback@example.com'), 'Jeremy G');
  });
  it('falls back when unset or blank', () => {
    assert.equal(resolveDisplayName(undefined, 'fb'), 'fb');
    assert.equal(resolveDisplayName('   ', 'fb'), 'fb');
  });
  it('falls back on an unresolved DXT config placeholder', () => {
    assert.equal(resolveDisplayName('${user_config.fastmail_caldav_display_name}', 'fb'), 'fb');
  });
});

// The display name is supplied through the constructor config, not read from the
// environment inside this module, so it travels the same path as every other setting
// the server resolves for its clients. These cover both places an ORGANIZER line is
// emitted: building a new event, and adding participants to an event that has none.
describe('ORGANIZER display name comes from the client config', () => {
  function mockedCreateClient(displayName?: string) {
    const client = new CalDAVCalendarClient({
      username: 'me@fastmail.com',
      password: 'test',
      displayName,
    });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function createWithParticipant(client: CalDAVCalendarClient) {
    return client.createCalendarEvent({
      calendarId: 'Personal',
      title: 'Meeting',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
      participants: [{ email: 'alice@example.com' }],
    });
  }

  const NO_ORGANIZER_ICAL = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'UID:noorg-cn@fm',
    'DTSTART:20260401T100000Z',
    'DTEND:20260401T110000Z',
    'SUMMARY:Simple Event',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  function mockedPatchClient(displayName?: string) {
    const client = new CalDAVCalendarClient({
      username: 'me@fastmail.com',
      password: 'test',
      displayName,
    });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: NO_ORGANIZER_ICAL, url: '/cal/noorg-cn.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function addParticipant(client: CalDAVCalendarClient) {
    return client.updateCalendarEvent('noorg-cn@fm', {
      participants: [{ email: 'alice@example.com' }],
    });
  }

  it('uses the configured display name as the ORGANIZER CN on a created event', async () => {
    const { client, mockDAVClient } = mockedCreateClient('Jeremy G');
    await createWithParticipant(client);

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('ORGANIZER;CN=Jeremy G:mailto:me@fastmail.com'), ical);
  });

  it('falls back to the CalDAV username when no display name is configured', async () => {
    const { client, mockDAVClient } = mockedCreateClient();
    await createWithParticipant(client);

    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('ORGANIZER;CN=me@fastmail.com:mailto:me@fastmail.com'), ical);
  });

  it('falls back when the configured display name is blank or an unresolved placeholder', async () => {
    for (const configured of ['   ', '${user_config.fastmail_caldav_display_name}']) {
      const { client, mockDAVClient } = mockedCreateClient(configured);
      await createWithParticipant(client);

      const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
      assert.ok(ical.includes('ORGANIZER;CN=me@fastmail.com:mailto:me@fastmail.com'), ical);
    }
  });

  it('uses the configured display name when adding an ORGANIZER on update', async () => {
    const { client, mockDAVClient } = mockedPatchClient('Jeremy G');
    await addParticipant(client);

    const data = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(data.includes('ORGANIZER;CN=Jeremy G:mailto:me@fastmail.com'), data);
  });

  it('falls back to the username when adding an ORGANIZER on update with none configured', async () => {
    const { client, mockDAVClient } = mockedPatchClient();
    await addParticipant(client);

    const data = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(data.includes('ORGANIZER;CN=me@fastmail.com:mailto:me@fastmail.com'), data);
  });
});

// ---------- DTSTART/DTEND time-frame and ordering agreement ----------

describe('createCalendarEvent start/end frame and ordering agreement', () => {
  // create defaults a designator-less value to the configured zone (#157) — pinned so the
  // 'names both values and both forms' and 'writes a designator-less pair in the configured
  // zone' tests below see a deterministic zone name regardless of the machine/CI environment
  // running them. Pinned to America/New_York specifically, not just "a fixed zone": this
  // machine's own host zone (Intl.DateTimeFormat().resolvedOptions().timeZone) is
  // Australia/Sydney, so pinning to Australia/Sydney here could not tell "read the configured
  // zone" apart from "silently fell back to the host zone" — a real regression to the host-zone
  // fallback would leave these assertions green.
  before(() => setDefaultTimezone('America/New_York'));
  after(() => setDefaultTimezone(undefined));

  function createMockedCreateClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function create(client: CalDAVCalendarClient, start: string, end: string) {
    return client.createCalendarEvent({ calendarId: 'Personal', title: 'T', start, end });
  }

  it('rejects a datetime end that is before the start, without writing', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T10:00:00Z', '2026-03-20T09:00:00Z'),
      /DTEND must be later than DTSTART/
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
  });

  it('rejects a datetime end equal to the start (zero-length event)', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T10:00:00Z', '2026-03-20T10:00:00Z'),
      /DTEND must be later than DTSTART/
    );
  });

  it('rejects a backwards end expressed with a UTC offset', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T20:00:00+10:00', '2026-03-20T09:00:00Z'),
      /DTEND must be later than DTSTART/
    );
  });

  it('rejects a floating start against a UTC end', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T09:30:00', '2026-03-20T10:30:00Z'),
      /same date\/time form/
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
  });

  it('rejects a UTC start against a floating end', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T09:30:00Z', '2026-03-20T10:30:00'),
      /same date\/time form/
    );
  });

  it('rejects a date-only start against a datetime end', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20', '2026-03-20T10:30:00Z'),
      /same date\/time form/
    );
  });

  it('rejects a datetime start against a date-only end', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T09:30:00Z', '2026-03-21'),
      /same date\/time form/
    );
  });

  it('names both values and both forms so the caller can fix the call', async () => {
    const { client } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T09:30:00', '2026-03-20T10:30:00Z'),
      (err: Error) => {
        assert.ok(err.message.includes("'2026-03-20T09:30:00'"), 'names the start value');
        assert.ok(err.message.includes("'2026-03-20T10:30:00Z'"), 'names the end value');
        // The designator-less start was defaulted to the configured zone (#157) rather than
        // left floating, so the error must say THAT, not "no time zone" — the caller named no
        // zone, but this server still wrote one, and saying otherwise would be false.
        assert.ok(
          err.message.includes("the account's configured time zone (America/New_York), applied because you named none"),
          'names the start form as the defaulted zone, not as floating'
        );
        assert.ok(err.message.includes('UTC date-time'), 'names the end form');
        return true;
      }
    );
  });

  it('throws InvalidInputError so the index maps it to InvalidParams', async () => {
    const { client } = createMockedCreateClient();
    for (const [start, end] of [
      ['2026-03-20T10:00:00Z', '2026-03-20T09:00:00Z'],
      ['2026-03-20T09:30:00', '2026-03-20T10:30:00Z'],
    ]) {
      await assert.rejects(
        () => create(client, start, end),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          return true;
        }
      );
    }
  });

  it('accepts an all-day event (both date-only, exclusive end)', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await create(client, '2026-03-20', '2026-03-21');
    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('DTSTART;VALUE=DATE:20260320'));
    assert.ok(ical.includes('DTEND;VALUE=DATE:20260321'));
  });

  it('accepts a both-UTC event', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await create(client, '2026-03-20T08:30:00Z', '2026-03-20T09:30:00Z');
    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('DTSTART:20260320T083000Z'));
    assert.ok(ical.includes('DTEND:20260320T093000Z'));
  });

  it('writes a designator-less pair in the configured zone, not floating (#157 create default)', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await create(client, '2026-03-20T08:30:00', '2026-03-20T09:30:00');
    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('DTSTART;TZID=America/New_York:20260320T083000'));
    assert.ok(ical.includes('DTEND;TZID=America/New_York:20260320T093000'));
    // Never the old bare-floating write — that behaviour change is the point of #157.
    assert.ok(!ical.includes('DTSTART:20260320T083000'));
    assert.ok(!ical.includes('DTEND:20260320T093000'));
  });
});

// The serialization path used to parse the caller's start/end itself and hand anything it
// did not recognise to `new Date()`, whose legacy fallback parser accepts a great deal and
// resolves it against the SERVER's time zone. A create call therefore had no single
// meaning: the same arguments produced a different day depending on where the server ran,
// and an impossible day was rolled quietly into the next month. Every case below reaches
// the write path — none of them is stopped by the shape guard the update path applies
// before it — so they are the evidence for validating the value in one place.
describe('createCalendarEvent rejects date spellings that would be resolved by guesswork', () => {
  function createMockedCreateClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function create(client: CalDAVCalendarClient, start: string, end: string) {
    return client.createCalendarEvent({ calendarId: 'Personal', title: 'T', start, end });
  }

  it('rejects a non-ISO date spelling instead of resolving it in the server time zone', async () => {
    for (const [start, end] of [
      ['2026/04/07', '2026/04/08'],
      ['April 7 2026', 'April 8 2026'],
      ['7 Apr 2026 14:00', '7 Apr 2026 15:00'],
    ]) {
      const { client, mockDAVClient } = createMockedCreateClient();
      await assert.rejects(
        () => create(client, start, end),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, /must be ISO-8601/);
          return true;
        },
      );
      assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0, `wrote for ${start}`);
    }
  });

  it('rejects a day its month does not have instead of rolling it into the next one', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-02-31T10:00:00Z', '2026-02-31T11:00:00Z'),
      /not a real calendar date/,
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
  });

  it('rejects an out-of-range wall-clock time instead of emitting it verbatim', async () => {
    // A floating value was previously copied straight into the property line without ever
    // being parsed, so DTSTART:20260320T253000 could be written to the server.
    const { client, mockDAVClient } = createMockedCreateClient();
    await assert.rejects(
      () => create(client, '2026-03-20T25:30:00', '2026-03-20T26:30:00'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
  });

  it('keeps a padded date-only value an all-day event rather than an instant', async () => {
    const { client, mockDAVClient } = createMockedCreateClient();
    await create(client, ' 2026-03-20 ', ' 2026-03-21 ');
    const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
    assert.ok(ical.includes('DTSTART;VALUE=DATE:20260320'), ical);
    assert.ok(ical.includes('DTEND;VALUE=DATE:20260321'), ical);
  });
});

describe('updateCalendarEvent start/end frame and ordering agreement', () => {
  function mockClient(icalData: string) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: icalData, url: '/cal/e.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function stored(uid: string, dtstart: string, dtend?: string): string {
    return [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      'DTSTAMP:20260301T000000Z',
      dtstart,
      ...(dtend ? [dtend] : []),
      'SUMMARY:Stored',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  const UTC_EVENT = stored('utc@fm', 'DTSTART:20260320T083000Z', 'DTEND:20260320T093000Z');
  const FLOATING_EVENT = stored('flt@fm', 'DTSTART:20260320T083000', 'DTEND:20260320T093000');
  const ZONED_EVENT = stored(
    'tz@fm',
    'DTSTART;TZID=Europe/Rome:20260320T083000',
    'DTEND;TZID=Europe/Rome:20260320T093000'
  );
  const ALLDAY_EVENT = stored('day@fm', 'DTSTART;VALUE=DATE:20260320', 'DTEND;VALUE=DATE:20260325');

  it('rejects a backwards datetime pair supplied together', async () => {
    const { client, mockDAVClient } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', {
        start: '2026-03-20T10:00:00Z',
        end: '2026-03-20T09:00:00Z',
      }),
      /DTEND must be later than DTSTART/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('rejects an equal datetime start and end', async () => {
    const { client } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', {
        start: '2026-03-20T10:00:00Z',
        end: '2026-03-20T10:00:00Z',
      }),
      /DTEND must be later than DTSTART/
    );
  });

  it('rejects a start alone that would land after the stored end', async () => {
    const { client, mockDAVClient } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', { start: '2026-03-20T11:00:00Z' }),
      /DTEND must be later than DTSTART/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('rejects an end alone that would land before the stored start', async () => {
    const { client } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', { end: '2026-03-20T07:00:00Z' }),
      /DTEND must be later than DTSTART/
    );
  });

  it('rejects a floating start against a stored UTC end', async () => {
    const { client, mockDAVClient } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', { start: '2026-03-20T09:30:00' }),
      /same date\/time form/
    );
    assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
  });

  it('rejects a floating end against a stored UTC start', async () => {
    const { client } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', { end: '2026-03-20T09:30:00' }),
      /same date\/time form/
    );
  });

  it('rejects a UTC start against a stored floating end', async () => {
    const { client } = mockClient(FLOATING_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('flt@fm', { start: '2026-03-20T08:00:00Z' }),
      /same date\/time form/
    );
  });

  it('rejects a UTC end against a stored floating start', async () => {
    const { client } = mockClient(FLOATING_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('flt@fm', { end: '2026-03-20T10:00:00Z' }),
      /same date\/time form/
    );
  });

  it('rejects a UTC start against a stored TZID-bearing end', async () => {
    const { client } = mockClient(ZONED_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('tz@fm', { start: '2026-03-20T07:00:00Z' }),
      /same date\/time form/
    );
  });

  it('rejects a date-only start against a stored datetime end', async () => {
    const { client } = mockClient(UTC_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('utc@fm', { start: '2026-03-19' }),
      /same date\/time form/
    );
  });

  it('rejects a datetime end against a stored date-only start', async () => {
    const { client } = mockClient(ALLDAY_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('day@fm', { end: '2026-03-25T09:00:00Z' }),
      /same date\/time form/
    );
  });

  it('rejects a date-only start that would land on or after the stored date-only end', async () => {
    const { client } = mockClient(ALLDAY_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('day@fm', { start: '2026-03-25' }),
      /DTEND is exclusive/
    );
  });

  it('throws InvalidInputError so the index maps it to InvalidParams', async () => {
    for (const fields of [
      { start: '2026-03-20T09:30:00' },
      { start: '2026-03-20T11:00:00Z' },
    ]) {
      const { client } = mockClient(UTC_EVENT);
      await assert.rejects(
        () => client.updateCalendarEvent('utc@fm', fields),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          return true;
        }
      );
    }
  });

  it('accepts a start alone that stays before the stored end (same frame)', async () => {
    const { client, mockDAVClient } = mockClient(UTC_EVENT);
    await client.updateCalendarEvent('utc@fm', { start: '2026-03-20T09:00:00Z' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTSTART:20260320T090000Z'));
    assert.ok(written.includes('DTEND:20260320T093000Z'));
  });

  it('accepts an end alone that stays after the stored start (same frame)', async () => {
    const { client, mockDAVClient } = mockClient(FLOATING_EVENT);
    await client.updateCalendarEvent('flt@fm', { end: '2026-03-20T10:30:00' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTEND:20260320T103000'));
  });

  it('accepts a date-only start alone that stays before the stored date-only end', async () => {
    const { client, mockDAVClient } = mockClient(ALLDAY_EVENT);
    await client.updateCalendarEvent('day@fm', { start: '2026-03-22' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTSTART;VALUE=DATE:20260322'));
    assert.ok(written.includes('DTEND;VALUE=DATE:20260325'));
  });

  it('accepts a floating start on a TZID-bearing event, keeping the stored timezone', async () => {
    const { client, mockDAVClient } = mockClient(ZONED_EVENT);
    await client.updateCalendarEvent('tz@fm', { start: '2026-03-20T09:00:00' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTSTART;TZID=Europe/Rome:20260320T090000'));
    assert.ok(written.includes('DTEND;TZID=Europe/Rome:20260320T093000'));
  });

  it('accepts a floating end on a TZID-bearing event, keeping the stored timezone', async () => {
    const { client, mockDAVClient } = mockClient(ZONED_EVENT);
    await client.updateCalendarEvent('tz@fm', { end: '2026-03-20T11:00:00' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTEND;TZID=Europe/Rome:20260320T110000'));
  });

  it('rejects a floating start that would land after the stored end in the same timezone', async () => {
    const { client } = mockClient(ZONED_EVENT);
    await assert.rejects(
      () => client.updateCalendarEvent('tz@fm', { start: '2026-03-20T10:00:00' }),
      /DTEND must be later than DTSTART/
    );
  });

  it('accepts a cross-timezone event, skipping an ordering comparison it cannot make', async () => {
    // Departs Europe/Rome, lands America/New_York. The wall clocks read backwards
    // but the instants do not; resolving that needs a timezone database, so the
    // frames agree and the ordering check stands down rather than guess.
    const flight = stored(
      'fly@fm',
      'DTSTART;TZID=Europe/Rome:20260320T100000',
      'DTEND;TZID=America/New_York:20260320T083000'
    );
    const { client, mockDAVClient } = mockClient(flight);
    await client.updateCalendarEvent('fly@fm', { start: '2026-03-20T11:00:00' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTSTART;TZID=Europe/Rome:20260320T110000'));
  });

  it('accepts a start change on a DURATION-based event with no stored DTEND', async () => {
    const durationEvent = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:dur3@fm',
      'DTSTART:20260320T083000Z',
      'DURATION:PT1H',
      'SUMMARY:Duration Event',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = mockClient(durationEvent);
    await client.updateCalendarEvent('dur3@fm', { start: '2026-03-25T08:30:00Z' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('DTSTART:20260325T083000Z'));
    assert.ok(written.includes('DURATION:PT1H'));
  });

  it('does not block a non-time edit on an event whose stored dates are already inconsistent', async () => {
    // The check exists to stop us WRITING a broken pair, not to hold a title
    // edit hostage to an inconsistency a third-party client left behind.
    const broken = stored('bad@fm', 'DTSTART:20260320T093000', 'DTEND:20260320T093000Z');
    const { client, mockDAVClient } = mockClient(broken);
    await client.updateCalendarEvent('bad@fm', { title: 'Renamed' });
    const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
    assert.ok(written.includes('SUMMARY:Renamed'));
  });
});

// The timeZone parameter on create_calendar_event/update_calendar_event (#157). The offset-shape
// rejection gate (isUsableTimezone) and its Etc/GMT-10 carve-out are already covered directly
// against validateCallerTimezone in coerce.test.ts; these are the create/update INTEGRATION
// tests — that the tools actually call that gate, and the write-path rules that only exist at
// this layer: standing down the cross-zone ordering check on a same-zone differently-spelled
// pair, rejecting a stranded single-sided zone change, rejecting timeZone combined with a value
// that already carries Z/an offset or is date-only, timeZone's precedence over a stored TZID,
// and the create/update default split.
describe('timeZone parameter (#157)', () => {
  function createClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function updateClient(icalData: string) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: icalData, url: '/cal/e.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function storedEvent(uid: string, dtstart: string, dtend: string): string {
    return [
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'BEGIN:VEVENT',
      `UID:${uid}`, 'DTSTAMP:20260301T000000Z',
      dtstart, dtend, 'SUMMARY:Stored',
      'END:VEVENT', 'END:VCALENDAR',
    ].join('\r\n');
  }

  describe('create_calendar_event', () => {
    before(() => setDefaultTimezone('Australia/Sydney'));
    after(() => setDefaultTimezone(undefined));

    it('an explicit timeZone overrides the configured default', async () => {
      const { client, mockDAVClient } = createClient();
      await client.createCalendarEvent({
        calendarId: 'Personal', title: 'T',
        start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
        timeZone: 'America/New_York',
      });
      const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
      assert.ok(ical.includes('DTSTART;TZID=America/New_York:20260320T083000'));
      assert.ok(ical.includes('DTEND;TZID=America/New_York:20260320T093000'));
    });

    it('rejects timeZone combined with a Z-designated start', async () => {
      const { client, mockDAVClient } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20T08:30:00Z', end: '2026-03-20T09:30:00Z',
          timeZone: 'America/New_York',
        }),
        /timeZone cannot be combined with a start that already carries Z or a UTC offset/
      );
      assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
    });

    it('rejects timeZone combined with an offset-designated end', async () => {
      const { client } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00+10:00',
          timeZone: 'America/New_York',
        }),
        /timeZone cannot be combined with a end that already carries Z or a UTC offset/
      );
    });

    it('rejects timeZone combined with a date-only start/end', async () => {
      const { client } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20', end: '2026-03-21',
          timeZone: 'America/New_York',
        }),
        /timeZone cannot be combined with a date-only start/
      );
    });

    it('rejects a null timeZone (no way to force a floating write)', async () => {
      const { client } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
          timeZone: null,
        }),
        /cannot be null, empty, or whitespace-only/
      );
    });

    it('rejects an empty-string timeZone', async () => {
      const { client } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
          timeZone: '   ',
        }),
        /cannot be null, empty, or whitespace-only/
      );
    });

    it('rejects an offset-shaped timeZone reaching create through the same gate', async () => {
      const { client } = createClient();
      await assert.rejects(
        () => client.createCalendarEvent({
          calendarId: 'Personal', title: 'T',
          start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
          timeZone: '+10:00',
        }),
        /not a fixed UTC offset/
      );
    });

    it('accepts Etc/GMT-10 as a timeZone reaching create through the same gate', async () => {
      const { client, mockDAVClient } = createClient();
      await client.createCalendarEvent({
        calendarId: 'Personal', title: 'T',
        start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
        timeZone: 'Etc/GMT-10',
      });
      const ical = callArguments(mockDAVClient.createCalendarObject)[0].iCalString;
      assert.ok(ical.includes('DTSTART;TZID=Etc/GMT-10:20260320T083000'));
    });
  });

  describe('update_calendar_event', () => {
    // Pinned to a configured zone that is neither this repo's dev host's own zone
    // (Australia/Sydney) nor the ZONED fixture's stored TZID below — see the inheritance
    // test further down for why a stray fallback to either one would otherwise go
    // unnoticed.
    before(() => setDefaultTimezone('America/New_York'));
    after(() => setDefaultTimezone(undefined));

    // End is stored a day after start (not the same evening) so that updating start alone to
    // the next morning (as the tests below do) still leaves a forward-ordered pair.
    const ZONED = storedEvent('tz@fm', 'DTSTART;TZID=Australia/Sydney:20260320T190000', 'DTEND;TZID=Australia/Sydney:20260321T200000');

    it('an explicit timeZone beats the stored TZID when both sides are re-sent', async () => {
      const { client, mockDAVClient } = updateClient(ZONED);
      await client.updateCalendarEvent('tz@fm', {
        start: '2026-03-21T09:00:00', end: '2026-03-21T10:00:00',
        timeZone: 'America/New_York',
      });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=America/New_York:20260321T090000'));
      assert.ok(written.includes('DTEND;TZID=America/New_York:20260321T100000'));
    });

    it('omitting timeZone never defaults — a designator-less value still inherits the stored TZID, not the configured default', async () => {
      // A dedicated fixture, not the shared ZONED above: this describe block previously ran
      // with no configured-zone pin at all, so the configured default fell back to whatever
      // zone the test host itself is in — which, on this repo's dev host, is Australia/Sydney,
      // the same spelling ZONED's stored TZID uses. "Inherits the stored TZID" and "falls back
      // to the configured/host default" then wrote the identical TZID, so this test could not
      // tell the two apart (confirmed by temporarily making update default an omitted zone to
      // the configured zone, the way create does: 8 other tests failed, and this was not one
      // of them). Pinning the block to America/New_York and storing this fixture in
      // Europe/London — a third zone, matching neither — closes that gap: either wrong
      // fallback now writes a TZID this assertion does not expect.
      const inheritZoned = storedEvent('tz-inherit@fm', 'DTSTART;TZID=Europe/London:20260321T090000', 'DTEND;TZID=Europe/London:20260321T100000');
      const { client, mockDAVClient } = updateClient(inheritZoned);
      await client.updateCalendarEvent('tz-inherit@fm', { start: '2026-03-21T09:30:00' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=Europe/London:20260321T093000'));
    });

    it('rejects timeZone with neither start nor end, naming the way to still reach it', async () => {
      const { client, mockDAVClient } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { timeZone: 'America/New_York' }),
        /neither start nor end was.*[Rr]e-send start and\/or end/s
      );
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    });

    it('rejects timeZone combined with a Z-designated start', async () => {
      const { client } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { start: '2026-03-21T09:00:00Z', timeZone: 'America/New_York' }),
        /timeZone cannot be combined with a start that already carries Z or a UTC offset/
      );
    });

    it('rejects timeZone combined with a date-only start', async () => {
      const { client } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { start: '2026-03-21', timeZone: 'America/New_York' }),
        /timeZone cannot be combined with a date-only start/
      );
    });

    it('rejects a null timeZone', async () => {
      const { client } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { start: '2026-03-21T09:00:00', timeZone: null }),
        /cannot be null, empty, or whitespace-only/
      );
    });

    it('rejects timeZone on start alone when the stored end is a DIFFERENT named zone', async () => {
      const { client, mockDAVClient } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { start: '2026-03-21T09:00:00', timeZone: 'America/New_York' }),
        /would rewrite start into 'America\/New_York' while the stored end stays in 'Australia\/Sydney'/
      );
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    });

    it('rejects timeZone on end alone when the stored start is a DIFFERENT named zone', async () => {
      const { client } = updateClient(ZONED);
      await assert.rejects(
        () => client.updateCalendarEvent('tz@fm', { end: '2026-03-21T10:00:00', timeZone: 'America/New_York' }),
        /would rewrite end into 'America\/New_York' while the stored start stays in 'Australia\/Sydney'/
      );
    });

    it('does not fire when the untouched side is stored in the SAME zone, differently spelled', async () => {
      // 'australia/sydney' (lowercase) names the same zone as the stored 'Australia/Sydney'
      // (zoneNamesEqual is case-insensitive), so this is not the two-zone shape the stranding
      // check exists to catch. What actually lands on the wire is the CANONICAL name ICU
      // resolves 'australia/sydney' to, not the caller's lowercase spelling — canonicalization
      // happens before the write, so asserting the raw input here would be asserting a value
      // this code path no longer produces.
      const { client, mockDAVClient } = updateClient(ZONED);
      await client.updateCalendarEvent('tz@fm', { start: '2026-03-21T09:00:00', timeZone: 'australia/sydney' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=Australia/Sydney:20260321T090000'));
    });

    it('a differently-spelled same zone does NOT stand down the ordering check (backwards pair rejected)', async () => {
      // Before the fix this compared TZIDs with raw !==, so 'Australia/Sydney' vs
      // 'australia/sydney' read as two DIFFERENT zones and stood the ordering check down —
      // silently accepting a backwards pair as a "flight lands elsewhere" shape it is not.
      const { client, mockDAVClient } = updateClient(ZONED);
      await assert.rejects(
        // Stored start is 19:00; an end of 08:00 the same day, in the "same" zone under a
        // different spelling, is backwards and must be rejected as an ordinary bad pair.
        () => client.updateCalendarEvent('tz@fm', { end: '2026-03-20T08:00:00', timeZone: 'australia/sydney' }),
        /DTEND must be later than DTSTART/
      );
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    });

    it('the same differently-spelled zone accepts a genuinely forward pair', async () => {
      // See the comment above: the written TZID is the canonical 'Australia/Sydney', not the
      // caller's lowercase 'australia/sydney' spelling.
      const { client, mockDAVClient } = updateClient(ZONED);
      await client.updateCalendarEvent('tz@fm', { end: '2026-03-20T21:00:00', timeZone: 'australia/sydney' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTEND;TZID=Australia/Sydney:20260320T210000'));
    });

    it('does not strand a single-sided update against a DURATION-based event with no stored DTEND', async () => {
      // rejectStrandedZoneMismatch reads the untouched side's line to compare zones; a
      // DURATION-based event has no DTEND line to read at all, and the function returns before
      // reaching the zone comparison. The call goes on to succeed: with no stored DTEND,
      // validateDateConsistency also has nothing to compare the new start against and skips.
      const durationEvent = storedEvent('dur-tz@fm', 'DTSTART:20260320T083000Z', 'DURATION:PT1H');
      const { client, mockDAVClient } = updateClient(durationEvent);
      await client.updateCalendarEvent('dur-tz@fm', { start: '2026-03-25T08:30:00', timeZone: 'America/New_York' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=America/New_York:20260325T083000'));
      assert.ok(written.includes('DURATION:PT1H'));
    });

    it('stands down when the untouched side is stored NON-zoned — the frame-consistency check rejects it instead, with its own message', async () => {
      // The stranding check only ever fires for a stored 'zoned' untouched side (see its
      // doc comment): a stored UTC or floating value already produces a frame mismatch against
      // the newly-zoned side, and that is validateDateConsistency's job, not this function's.
      // This asserts the DISTINCTION — the call is still rejected, but by the frame-consistency
      // message, never by the "would rewrite ... while the stored ... stays in" wording.
      const utcEnd = storedEvent('utc-end@fm', 'DTSTART:20260320T190000Z', 'DTEND:20260321T200000Z');
      const { client, mockDAVClient } = updateClient(utcEnd);
      await assert.rejects(
        () => client.updateCalendarEvent('utc-end@fm', { start: '2026-03-21T09:00:00', timeZone: 'America/New_York' }),
        (err: Error) => {
          assert.match(err.message, /DTSTART and DTEND must use the same date\/time form/);
          assert.doesNotMatch(err.message, /would rewrite/);
          return true;
        },
      );
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    });

    it('recognises the stranded side as the SAME zone through a stored leading-slash TZID (RFC 5545 §3.2.19)', async () => {
      // zoneNamesEqual strips exactly one leading slash and lowercases before comparing, so a
      // stored '/Australia/Sydney' (the vendor/registry-qualified form the RFC allows) must
      // still read as the same zone as a caller's un-prefixed 'Australia/Sydney' — not two
      // different zones that would trip the stranding rejection.
      const slashZoned = storedEvent(
        'slash-tz@fm',
        'DTSTART;TZID=/Australia/Sydney:20260320T190000',
        'DTEND;TZID=/Australia/Sydney:20260321T200000',
      );
      const { client, mockDAVClient } = updateClient(slashZoned);
      await client.updateCalendarEvent('slash-tz@fm', { start: '2026-03-21T09:00:00', timeZone: 'Australia/Sydney' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=Australia/Sydney:20260321T090000'));
    });

    // zoneNamesEqual used to compare TZIDs with raw string equality, so a stored TZID that
    // named the same zone through an ICU link/alias spelling — not just a different case —
    // read as a DIFFERENT zone from a caller's re-affirmed spelling. That produced a false
    // "stranded two-zone event" rejection on an ordinary read-modify-write: reading a stored
    // 'NZ' TZID off this server (the read side emits it verbatim — #139 — since it was written
    // by some other client) and then re-zoning the touched side to the equivalent canonical
    // name. The caller cannot pass 'NZ' itself here any more (#157 amendment rejects bare
    // shorthand on write), so this exercises the canonical spelling a caller is now required to
    // send — 'Pacific/Auckland' — against a STORED side that still carries the raw alias.
    it('recognises the stranded side as the SAME zone through a link/alias spelling ("NZ" == "Pacific/Auckland")', async () => {
      const nzZoned = storedEvent('nz-tz@fm', 'DTSTART;TZID=NZ:20260320T190000', 'DTEND;TZID=NZ:20260321T200000');
      const { client, mockDAVClient } = updateClient(nzZoned);
      await client.updateCalendarEvent('nz-tz@fm', { start: '2026-03-21T09:00:00', timeZone: 'Pacific/Auckland' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=Pacific/Auckland:20260321T090000'));
    });

    it('recognises the stranded side as the SAME zone through a link/alias spelling ("US/Pacific" == "America/Los_Angeles")', async () => {
      const usPacificZoned = storedEvent('uspac-tz@fm', 'DTSTART;TZID=US/Pacific:20260320T190000', 'DTEND;TZID=US/Pacific:20260321T200000');
      const { client, mockDAVClient } = updateClient(usPacificZoned);
      await client.updateCalendarEvent('uspac-tz@fm', { start: '2026-03-21T09:00:00', timeZone: 'US/Pacific' });
      const written = callArguments(mockDAVClient.updateCalendarObject)[0].calendarObject.data;
      assert.ok(written.includes('DTSTART;TZID=America/Los_Angeles:20260321T090000'));
    });

    // Recognising 'NZ' and 'Pacific/Auckland' as the same zone is not a one-way relaxation:
    // validateDateConsistency's own "different zones, flight lands elsewhere" allowance
    // (#140) reads through the identical zoneNamesEqual, so a pair that used to look
    // cross-zone (and so skipped ordering entirely) is now ordering-checked like any
    // same-zone pair — and a genuinely backwards alias pair is still rejected, not silently
    // written. More checking, not less.
    it('an alias pair that clears the stranding check is still ordering-checked, and a backwards one is rejected', async () => {
      const nzZoned = storedEvent('nz-order@fm', 'DTSTART;TZID=NZ:20260320T190000', 'DTEND;TZID=NZ:20260321T080000');
      const { client, mockDAVClient } = updateClient(nzZoned);
      await assert.rejects(
        // Stored end is 08:00 on the 21st; a new start of 09:00 the same day, in the "same"
        // zone under the alias spelling 'Pacific/Auckland', is backwards.
        () => client.updateCalendarEvent('nz-order@fm', { start: '2026-03-21T09:00:00', timeZone: 'Pacific/Auckland' }),
        /DTEND must be later than DTSTART/
      );
      assert.equal(mockDAVClient.updateCalendarObject.mock.calls.length, 0);
    });
  });
});

// classifyWrittenLine — the function describeCreateCalendarEventResult/describeUpdateCalendarEventResult
// build their sentences from — is exercised elsewhere only via hand-built CalendarZoneWriteInfo
// literals passed straight to those two describe functions. That proves the sentence wording is
// right given a classification, but never proves createCalendarEvent/updateCalendarEvent
// actually PRODUCE that classification from a real call: a regression in classifyWrittenLine's
// own line-reading, or in what formatDateTimeProperty hands it, could pass every test above
// while still misreporting what got written. These drive the real methods end to end.
describe('calendar write result classification, driven from real create/update calls (#157)', () => {
  function createClient() {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  function updateClient(icalData: string) {
    const client = new CalDAVCalendarClient({ username: 'test@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Personal', url: '/cal/personal/' }]),
      fetchCalendarObjects: mock.fn(async (_params: FetchObjectsParams) => [{ data: icalData, url: '/cal/e.ics' }]),
      updateCalendarObject: mock.fn(async (_params: UpdateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  describe('create_calendar_event', () => {
    // A zone other than this host's own, for the same reason as the re-pin above: a regression
    // to reading the host zone instead of the configured one must not leave this green.
    before(() => setDefaultTimezone('America/New_York'));
    after(() => setDefaultTimezone(undefined));

    it('classifies a designator-less pair as zoned, in the configured zone', async () => {
      const { client } = createClient();
      const result = await client.createCalendarEvent({
        calendarId: 'Personal', title: 'T',
        start: '2026-03-20T08:30:00', end: '2026-03-20T09:30:00',
      });
      assert.deepEqual(result.start, { kind: 'zoned', zone: 'America/New_York' });
      assert.deepEqual(result.end, { kind: 'zoned', zone: 'America/New_York' });
    });

    it('classifies a Z-designated pair as utc', async () => {
      const { client } = createClient();
      const result = await client.createCalendarEvent({
        calendarId: 'Personal', title: 'T',
        start: '2026-03-20T08:30:00Z', end: '2026-03-20T09:30:00Z',
      });
      assert.deepEqual(result.start, { kind: 'utc' });
      assert.deepEqual(result.end, { kind: 'utc' });
    });

    it('classifies a date-only pair as allday', async () => {
      const { client } = createClient();
      const result = await client.createCalendarEvent({
        calendarId: 'Personal', title: 'T',
        start: '2026-03-20', end: '2026-03-21',
      });
      assert.deepEqual(result.start, { kind: 'allday' });
      assert.deepEqual(result.end, { kind: 'allday' });
    });

    // create always defaults a designator-less value to the configured zone (#157), so
    // 'floating' is not reachable from createCalendarEvent — only from updateCalendarEvent,
    // which never defaults an omitted timeZone. See that case below.
  });

  describe('update_calendar_event', () => {
    function stored(uid: string, dtstart: string, dtend: string): string {
      return [
        'BEGIN:VCALENDAR', 'BEGIN:VEVENT',
        `UID:${uid}`, 'DTSTAMP:20260301T000000Z',
        dtstart, dtend, 'SUMMARY:Stored',
        'END:VEVENT', 'END:VCALENDAR',
      ].join('\r\n');
    }

    it('classifies a re-zoned pair as zoned, in the caller-named zone', async () => {
      const icalData = stored(
        'u1@fm',
        'DTSTART;TZID=Australia/Sydney:20260320T190000',
        'DTEND;TZID=Australia/Sydney:20260321T200000',
      );
      const { client } = updateClient(icalData);
      const result = await client.updateCalendarEvent('u1@fm', {
        start: '2026-03-21T09:00:00', end: '2026-03-21T10:00:00', timeZone: 'America/New_York',
      });
      assert.deepEqual(result.start, { kind: 'zoned', zone: 'America/New_York' });
      assert.deepEqual(result.end, { kind: 'zoned', zone: 'America/New_York' });
    });

    it('classifies a Z-designated pair as utc', async () => {
      const icalData = stored('u2@fm', 'DTSTART:20260320T083000Z', 'DTEND:20260320T093000Z');
      const { client } = updateClient(icalData);
      const result = await client.updateCalendarEvent('u2@fm', {
        start: '2026-03-25T08:30:00Z', end: '2026-03-25T09:30:00Z',
      });
      assert.deepEqual(result.start, { kind: 'utc' });
      assert.deepEqual(result.end, { kind: 'utc' });
    });

    it('classifies a date-only pair as allday', async () => {
      const icalData = stored('u3@fm', 'DTSTART;VALUE=DATE:20260320', 'DTEND;VALUE=DATE:20260321');
      const { client } = updateClient(icalData);
      const result = await client.updateCalendarEvent('u3@fm', { start: '2026-03-25', end: '2026-03-26' });
      assert.deepEqual(result.start, { kind: 'allday' });
      assert.deepEqual(result.end, { kind: 'allday' });
    });

    it('classifies a floating pair as floating — the one kind create can never write', async () => {
      const icalData = stored('u4@fm', 'DTSTART:20260320T083000', 'DTEND:20260320T093000');
      const { client } = updateClient(icalData);
      const result = await client.updateCalendarEvent('u4@fm', {
        start: '2026-03-25T08:30:00', end: '2026-03-25T09:30:00',
      });
      assert.deepEqual(result.start, { kind: 'floating' });
      assert.deepEqual(result.end, { kind: 'floating' });
    });

    it('reports neither side when a non-time edit touches neither start nor end', async () => {
      const icalData = stored('u5@fm', 'DTSTART:20260320T083000Z', 'DTEND:20260320T093000Z');
      const { client } = updateClient(icalData);
      const result = await client.updateCalendarEvent('u5@fm', { title: 'Renamed' });
      assert.equal(result.start, undefined);
      assert.equal(result.end, undefined);
    });
  });
});

describe('calendar write result sentences (#157)', () => {
  it('create: reports a single zone when both sides land in it', () => {
    const msg = describeCreateCalendarEventResult({
      eventId: 'e1',
      start: { kind: 'zoned', zone: 'Australia/Sydney' },
      end: { kind: 'zoned', zone: 'Australia/Sydney' },
    });
    assert.equal(msg, ' Written in zone Australia/Sydney.');
  });

  it('create: reports both zones separately on a legitimate cross-zone (flight) pair', () => {
    const msg = describeCreateCalendarEventResult({
      eventId: 'e1',
      start: { kind: 'zoned', zone: 'Australia/Sydney' },
      end: { kind: 'zoned', zone: 'America/New_York' },
    });
    assert.equal(msg, ' Start written in zone Australia/Sydney, end written in zone America/New_York.');
  });

  it('create: reports UTC', () => {
    const msg = describeCreateCalendarEventResult({
      eventId: 'e1',
      start: { kind: 'utc' },
      end: { kind: 'utc' },
    });
    assert.equal(msg, ' Written in UTC.');
  });

  it('create: reports all-day', () => {
    const msg = describeCreateCalendarEventResult({
      eventId: 'e1',
      start: { kind: 'allday' },
      end: { kind: 'allday' },
    });
    assert.equal(msg, ' Written in all-day (no time component).');
  });

  it('update: reports only the sides actually written', () => {
    assert.equal(
      describeUpdateCalendarEventResult({ eventId: 'e1', start: { kind: 'zoned', zone: 'Australia/Sydney' } }),
      ' (start zone Australia/Sydney)'
    );
    assert.equal(
      describeUpdateCalendarEventResult({ eventId: 'e1', end: { kind: 'utc' } }),
      ' (end UTC)'
    );
    assert.equal(
      describeUpdateCalendarEventResult({
        eventId: 'e1',
        start: { kind: 'floating' },
        end: { kind: 'zoned', zone: 'Australia/Sydney' },
      }),
      ' (start floating (no zone), end zone Australia/Sydney)'
    );
  });

  it('update: reports nothing extra when neither side was touched', () => {
    assert.equal(describeUpdateCalendarEventResult({ eventId: 'e1' }), '');
  });
});

// Every CalDAV request carries the account's Basic credential, so following a
// redirect would replay that credential at whatever host the response names. The
// client is built with redirect: 'error' to make that impossible.
//
// Both halves of that are asserted, because either one alone can go green while the
// protection is gone: the option being set proves nothing if tsdav stops carrying it
// onto requests, and a request carrying it proves nothing about the client the
// production path builds. So the test takes the real client out of getClient(), then
// drives a real tsdav method through it and inspects the init that reaches fetch.
//
// The spy is installed as the client's own fetch override rather than by patching
// globalThis: tsdav resolves globalThis.fetch when its module is first evaluated, so
// a later patch is never consulted.
describe('CalDAV requests refuse to follow redirects', () => {
  it("carries redirect: 'error' from the constructed client onto the requests it makes", async () => {
    const realLogin = DAVClient.prototype.login;
    // login() is the one call getClient() makes that touches the network. Stubbing it
    // on the prototype leaves the constructor - the part under test - untouched.
    DAVClient.prototype.login = (async function (this: DAVClient) {}) as typeof realLogin;

    try {
      const wrapper = new CalDAVCalendarClient({ username: 'test@example.com', password: 'pw' });
      // getClient() is private and is where the option is wired in; going through a
      // public method would make this depend on that method's behaviour as well.
      const dav = (await (wrapper as any).getClient()) as DAVClient;

      assert.deepEqual(
        dav.fetchOptions,
        { redirect: 'error' },
        "the DAV client was built without redirect: 'error', so a redirect would replay the CalDAV credential at the host that sent it"
      );

      const seen: RequestInit[] = [];
      dav.fetchOverride = (async (_url: any, init?: RequestInit) => {
        seen.push(init ?? {});
        return new Response('', { status: 207 });
      }) as typeof globalThis.fetch;

      // createObject is the cheapest real request method: it hands the response back
      // untouched, so nothing here depends on parsing a plausible DAV body.
      await dav.createObject({ url: 'https://caldav.example.com/x', data: 'x', headers: {} });

      assert.equal(seen.length, 1, 'expected the DAV method to issue exactly one request');
      assert.equal(
        seen[0].redirect,
        'error',
        "the request reached fetch without redirect: 'error' - tsdav is no longer carrying fetchOptions onto requests, so the credential would follow a redirect"
      );
    } finally {
      DAVClient.prototype.login = realLogin;
    }
  });
});

// A failed CalDAV login used to still cache the unauthenticated client (getClient()
// assigned `this.client` before awaiting login()), so every later call took the
// `if (this.client)` fast path and failed downstream inside tsdav with a bare "no
// account for fetchCalendars" instead of the real auth error (#143). getClient() now
// only caches the client once login() has resolved.
describe('CalDAV login failure is not cached (#143)', () => {
  it('does not cache the client after a failed login, so a second call retries the login and surfaces the auth error again', async () => {
    const realLogin = DAVClient.prototype.login;
    const loginMock = mock.fn(async () => {
      throw new Error('Invalid credentials');
    });
    DAVClient.prototype.login = loginMock as unknown as typeof realLogin;

    try {
      const wrapper = new CalDAVCalendarClient({ username: 'test@example.com', password: 'wrong' });

      const isAuthError = (err: unknown) =>
        err instanceof Error && err.message.includes('CalDAV login failed') && err.message.includes('app password');

      // getClient() is private; going through a public method would make this
      // depend on that method's own behaviour as well as getClient()'s.
      await assert.rejects((wrapper as any).getClient(), isAuthError);
      assert.equal(loginMock.mock.calls.length, 1, 'expected the first call to attempt exactly one login');

      // The bug: this second call used to return the client cached by the first
      // (failed) call instead of retrying, so it never reached login() again and
      // instead failed later, downstream, with a different and less useful message.
      await assert.rejects(
        (wrapper as any).getClient(),
        isAuthError,
        'expected the second call to surface the same auth error, not a different downstream failure from a stale cached client'
      );
      assert.equal(
        loginMock.mock.calls.length,
        2,
        'expected getClient() to attempt a fresh login on the second call rather than returning a cached, unauthenticated client'
      );
    } finally {
      DAVClient.prototype.login = realLogin;
    }
  });
});

describe('CalDAV login success is still cached', () => {
  it('logs in once and reuses the same client across two getClient() calls', async () => {
    const realLogin = DAVClient.prototype.login;
    const loginMock = mock.fn(async function (this: DAVClient) {});
    DAVClient.prototype.login = loginMock as unknown as typeof realLogin;

    try {
      const wrapper = new CalDAVCalendarClient({ username: 'test@example.com', password: 'pw' });

      const first = await (wrapper as any).getClient();
      const second = await (wrapper as any).getClient();

      assert.equal(first, second, 'expected the same DAVClient instance to be reused rather than re-constructed');
      assert.equal(loginMock.mock.calls.length, 1, 'expected only one login attempt across two getClient() calls');
    } finally {
      DAVClient.prototype.login = realLogin;
    }
  });
});

// ============================================================
// Recurrence expansion and window scoping (#64), and the read
// path's silent under-reporting (#100).
// ============================================================

describe('parseCalendarObjects', () => {
  it('returns one event per VEVENT when the server expanded the recurrence', () => {
    // The shape a time-range query with expand actually returns: several occurrences inside
    // ONE calendar-data blob, RRULE stripped, RECURRENCE-ID set to the real date. Reading
    // only the first block here would report one event where the window holds three.
    const occurrence = (date: string) => [
      'BEGIN:VEVENT',
      'UID:payday@fm',
      `DTSTART;VALUE=DATE:${date}`,
      `RECURRENCE-ID;VALUE=DATE:${date}`,
      'SUMMARY:Pay Day',
      'END:VEVENT',
    ].join('\r\n');
    const data = [
      'BEGIN:VCALENDAR',
      occurrence('20270305'),
      occurrence('20270319'),
      occurrence('20270402'),
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/payday.ics' }, { expanded: true });

    // The same blob read WITHOUT the flag also yields three: no block lacks a RECURRENCE-ID,
    // so the unexpanded path sees a resource of detached overrides and emits each of them.
    assert.equal(parseCalendarObjects({ data, url: '/cal/payday.ics' }).length, 3);

    assert.equal(events.length, 3);
    assert.deepEqual(events.map(e => e.start), ['2027-03-05', '2027-03-19', '2027-04-02']);
    assert.deepEqual(events.map(e => e.recurrenceId), ['2027-03-05', '2027-03-19', '2027-04-02']);
    for (const e of events) {
      assert.equal(e.isRecurring, true);
      // An expanded occurrence has no rule of its own; the server already applied it.
      assert.equal(e.recurrenceRule, undefined);
      assert.equal(e.title, 'Pay Day');
    }
  });

  it('emits every occurrence when the first instance carries no RECURRENCE-ID', () => {
    // The shape Fastmail ACTUALLY returns when the requested window contains the series'
    // original DTSTART. Cyrus's expansion sets a RECURRENCE-ID only on instances after the
    // first, so block 0 arrives with its RRULE stripped and no RECURRENCE-ID at all. A
    // parser that looked for "the block without a RECURRENCE-ID" found one, called it the
    // series master, and threw the other four away — a five-year window over a yearly
    // birthday reported ONE event, and said nothing about the four it dropped.
    const firstInstance = [
      'BEGIN:VEVENT',
      'UID:birthday@fm',
      'DTSTART;VALUE=DATE:19940612',
      'SUMMARY:Birthday',
      'END:VEVENT',
    ].join('\r\n');
    const occurrence = (date: string) => [
      'BEGIN:VEVENT',
      'UID:birthday@fm',
      `DTSTART;VALUE=DATE:${date}`,
      `RECURRENCE-ID;VALUE=DATE:${date}`,
      'SUMMARY:Birthday',
      'END:VEVENT',
    ].join('\r\n');
    const data = [
      'BEGIN:VCALENDAR',
      firstInstance,
      occurrence('19950612'),
      occurrence('19960612'),
      occurrence('19970612'),
      occurrence('19980612'),
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/birthday.ics' }, { expanded: true });

    assert.equal(events.length, 5);
    assert.deepEqual(
      events.map(e => e.start),
      ['1994-06-12', '1995-06-12', '1996-06-12', '1997-06-12', '1998-06-12'],
    );
    // The first instance has no recurrenceId of its own — the server never set one — but it
    // is still an occurrence of a repeating series, so reporting it as a one-off would be
    // the same lie the dropped siblings told, one row smaller.
    assert.equal(events[0].recurrenceId, undefined);
    assert.equal(events[0].recurrenceRule, undefined);
    for (const e of events) assert.equal(e.isRecurring, true);
  });

  it('leaves recurrence fields off a lone expanded block, which a one-off and a series start share', () => {
    // An expanded blob holding exactly one block with no RECURRENCE-ID is genuinely
    // ambiguous: Cyrus emits a one-off event and a series' first-and-only in-window instance
    // identically. Nothing is claimed in either direction; the tool description points at
    // get_calendar_event to settle it.
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:once@fm',
      'DTSTART:20270305T093000Z',
      'SUMMARY:One off',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/once.ics' }, { expanded: true });

    assert.equal(events.length, 1);
    assert.equal(events[0].isRecurring, undefined);
  });

  it('still returns only the master for the same blob when expansion was NOT requested', () => {
    // The shape decision comes from the CALLER, never from the payload. Without `expanded`
    // this is a series master followed by an override, and only the master is reported.
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'DTSTART;VALUE=DATE:20190101',
      'RRULE:FREQ=WEEKLY;BYDAY=MO',
      'SUMMARY:Standup',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'RECURRENCE-ID;VALUE=DATE:20190408',
      'DTSTART;VALUE=DATE:20190409',
      'SUMMARY:Standup moved',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    assert.equal(parseCalendarObjects({ data, url: '/cal/s.ics' }).length, 1);
    assert.equal(parseCalendarObjects({ data, url: '/cal/s.ics' }, { expanded: true }).length, 2);
  });

  it('returns only the master for an unexpanded series, carrying its RRULE', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'DTSTART;VALUE=DATE:20190101',
      'RRULE:FREQ=WEEKLY;BYDAY=MO',
      'SUMMARY:Standup',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'RECURRENCE-ID;VALUE=DATE:20190408',
      'DTSTART;VALUE=DATE:20190409',
      'SUMMARY:Standup moved',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/standup.ics' });

    assert.equal(events.length, 1);
    assert.equal(events[0].title, 'Standup');
    assert.equal(events[0].start, '2019-01-01');
    assert.equal(events[0].isRecurring, true);
    assert.equal(events[0].recurrenceRule, 'FREQ=WEEKLY;BYDAY=MO');
    // No recurrenceId: this is the series shown at its original start, and the absence of
    // that field is what says so.
    assert.equal(events[0].recurrenceId, undefined);
  });

  it('picks the master even when an exception is serialised first', () => {
    // RFC 5545 fixes no component order, so another client can write the override ahead of
    // the master. Taking the first block would report the override as though it were the event.
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'RECURRENCE-ID;VALUE=DATE:20190408',
      'DTSTART;VALUE=DATE:20190409',
      'SUMMARY:Standup moved',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'DTSTART;VALUE=DATE:20190101',
      'RRULE:FREQ=WEEKLY;BYDAY=MO',
      'SUMMARY:Standup',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/standup.ics' });

    assert.equal(events.length, 1);
    assert.equal(events[0].title, 'Standup');
    assert.equal(events[0].recurrenceRule, 'FREQ=WEEKLY;BYDAY=MO');
  });

  it('falls back to a single minimal event when the payload has no VEVENT', () => {
    const data = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR';
    const events = parseCalendarObjects({ data, url: '/cal/empty.ics' });

    assert.equal(events.length, 1);
    assert.deepEqual(events[0], { id: '/cal/empty.ics', url: '/cal/empty.ics', title: 'Untitled' });
  });

  it('parses timed occurrences as datetimes and all-day occurrences as dates', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:mixed@fm',
      'DTSTART:20270305T093000Z',
      'DTEND:20270305T103000Z',
      'RECURRENCE-ID:20270305T093000Z',
      'SUMMARY:Timed',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/mixed.ics' });
    assert.equal(events[0].start, '2027-03-05T09:30:00Z');
    assert.equal(events[0].end, '2027-03-05T10:30:00Z');
    assert.equal(events[0].recurrenceId, '2027-03-05T09:30:00Z');
  });

  it('leaves recurrence fields off a one-off event', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:once@fm',
      'DTSTART:20270305T093000Z',
      'SUMMARY:One off',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const events = parseCalendarObjects({ data, url: '/cal/once.ics' });
    assert.equal(events.length, 1);
    assert.equal(events[0].isRecurring, undefined);
    assert.equal(events[0].recurrenceId, undefined);
    assert.equal(events[0].recurrenceRule, undefined);
  });

  it('parseCalendarObject returns the master, matching the single-event callers', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'RECURRENCE-ID;VALUE=DATE:20190408',
      'SUMMARY:Standup moved',
      'END:VEVENT',
      'BEGIN:VEVENT',
      'UID:standup@fm',
      'DTSTART;VALUE=DATE:20190101',
      'RRULE:FREQ=WEEKLY',
      'SUMMARY:Standup',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    assert.equal(parseCalendarObject({ data, url: '/cal/s.ics' }).title, 'Standup');
  });
});

describe('eventIntersectsWindow', () => {
  const WINDOW_START = Date.parse('2027-03-01T00:00:00Z');
  const WINDOW_END = Date.parse('2027-03-10T00:00:00Z');

  // The zone is passed explicitly on every call, never left to the host: it is a required
  // parameter precisely so that a zone-less value cannot be placed in whatever zone the
  // machine running the tests happens to be in. 'UTC' keeps the cases below about what they
  // are about; the zone behaviour itself is asserted separately, in zones that are not UTC.
  it('excludes a series master dated years before the window', () => {
    // The exact failure #64 reported: a ten-day window in March 2027 answered with an
    // August 2020 date.
    assert.equal(
      eventIntersectsWindow({ start: '2020-08-28', end: '2020-08-29' }, WINDOW_START, WINDOW_END, 'UTC'),
      false,
    );
  });

  it('includes an occurrence inside the window', () => {
    assert.equal(
      eventIntersectsWindow({ start: '2027-03-05', end: '2027-03-06' }, WINDOW_START, WINDOW_END, 'UTC'),
      true,
    );
  });

  it('includes an all-day event on the last day the window covers', () => {
    // The window end is exclusive, so an all-day event on 9 March runs [09, 10) and must
    // survive; getting this wrong silently loses the last day of every query.
    assert.equal(
      eventIntersectsWindow({ start: '2027-03-09', end: '2027-03-10' }, WINDOW_START, WINDOW_END, 'UTC'),
      true,
    );
  });

  it('excludes an event that starts after the window ends', () => {
    assert.equal(
      eventIntersectsWindow({ start: '2027-04-05T09:00:00Z', end: '2027-04-05T10:00:00Z' }, WINDOW_START, WINDOW_END, 'UTC'),
      false,
    );
  });

  it('DROPS a zone-free value just outside the window, which the old margin kept (#162)', () => {
    // THIS ASSERTION USED TO BE `true`. The filter granted any designator-less value fourteen
    // hours of slack on both edges, so a value one morning past the window survived it and the
    // caller saw a row outside the days they asked about. That margin now widens the range
    // REQUESTED of the server instead, where it does work a filter cannot do, and what comes
    // back is judged exactly: resolved in the configured zone, 10 March 08:00 is past a window
    // that ends at midnight on the 10th.
    assert.equal(
      eventIntersectsWindow({ start: '2027-03-10T08:00:00', end: '2027-03-10T09:00:00' }, WINDOW_START, WINDOW_END, 'UTC'),
      false,
    );
  });

  it('keeps an event with no readable dates rather than dropping it silently', () => {
    assert.equal(eventIntersectsWindow({}, WINDOW_START, WINDOW_END, 'UTC'), true);
  });

  it('keeps a multi-day date span queried on its MIDDLE day (#162)', () => {
    // A date-only DTEND is ALREADY exclusive in iCalendar, so a three-day all-day event is
    // stored as 05 .. 08 and covers the 5th, 6th and 7th. Adding a further day to the end (or
    // collapsing the span to its start) is what would make a query on the 6th miss it.
    const midDay = Date.parse('2027-03-06T00:00:00Z');
    const nextDay = Date.parse('2027-03-07T00:00:00Z');
    assert.equal(
      eventIntersectsWindow({ start: '2027-03-05', end: '2027-03-08' }, midDay, nextDay, 'UTC'),
      true,
    );
  });

  it('keeps an all-day event with no end when the window starts mid-morning (#162)', () => {
    // A missing end used to collapse the event to a zero-width instant at local midnight, so
    // any window that started later that morning contained no part of it and the whole day's
    // entry vanished. An all-day value covers its whole LOCAL day.
    const midMorning = Date.parse('2027-03-05T09:00:00Z');
    const midday = Date.parse('2027-03-05T12:00:00Z');
    assert.equal(
      eventIntersectsWindow({ start: '2027-03-05' }, midMorning, midday, 'UTC'),
      true,
    );
  });

  it('resolves a naive datetime in the CONFIGURED zone, not as UTC (#162)', () => {
    // The floating-evening bug from the other end. 20:00 with no designator is 20:00 where the
    // account lives, so on a +10 account it falls inside that evening's window and on a -05
    // one it does not. Read as UTC — which is what the old fixed frame did — the same value
    // answered both the same way, and the +10 account's evening came back empty.
    const eveningStart = Date.parse('2027-03-05T09:00:00Z'); // 20:00 in Sydney, 04:00 in New York
    const eveningEnd = Date.parse('2027-03-05T12:00:00Z');
    const floating = { start: '2027-03-05T20:00:00', end: '2027-03-05T21:00:00' };
    assert.equal(eventIntersectsWindow(floating, eveningStart, eveningEnd, 'Australia/Sydney'), true);
    assert.equal(eventIntersectsWindow(floating, eveningStart, eveningEnd, 'America/New_York'), false);
  });

  it('resolves a wall clock carrying its own resolvable TZID in ITS zone (#162)', () => {
    // The configured zone is the FALLBACK, not an override. With no slack left, reading a
    // New York wall clock as though it were Sydney time is fifteen hours out — far enough to
    // drop a real event, which is the failure this whole issue is about.
    const window = { start: Date.parse('2027-03-05T14:00:00Z'), end: Date.parse('2027-03-05T16:00:00Z') };
    const newYorkEvent = {
      start: '2027-03-05T09:30:00',
      end: '2027-03-05T10:30:00',
      timeZone: 'America/New_York',
    };
    // 09:30 in New York is 14:30Z, inside the window.
    assert.equal(eventIntersectsWindow(newYorkEvent, window.start, window.end, 'Australia/Sydney'), true);
    // The control: the same wall clock with no TZID falls back to the configured zone, where
    // 09:30 Sydney time is 22:30Z the previous day and nowhere near the window.
    assert.equal(
      eventIntersectsWindow({ start: newYorkEvent.start, end: newYorkEvent.end }, window.start, window.end, 'Australia/Sydney'),
      false,
    );
  });

  it('resolves a TZID written in the RFC 5545 global (leading-slash) form (#162)', () => {
    // The stored spelling is emitted verbatim, leading '/' and all, so it reaches the filter
    // as '/America/New_York' — which ICU rejects outright. Testing the raw string sent a real
    // New York event to the configured zone and DROPPED it: a missing event, arriving only for
    // calendars whose TZID happens to be written this way. The name is normalised first.
    const windowStart = Date.parse('2027-03-05T14:00:00Z');
    const windowEnd = Date.parse('2027-03-05T16:00:00Z');
    assert.equal(
      eventIntersectsWindow(
        { start: '2027-03-05T09:30:00', end: '2027-03-05T10:30:00', timeZone: '/America/New_York' },
        windowStart,
        windowEnd,
        'Australia/Sydney',
      ),
      true,
    );
    // The control, and the boundary of the normalisation: a vendor-prefixed TZID names a zone
    // out of somebody else's registry, so it still does not resolve and still falls back to
    // the configured zone, where 09:30 Sydney time is nowhere near this window.
    assert.equal(
      eventIntersectsWindow(
        {
          start: '2027-03-05T09:30:00',
          end: '2027-03-05T10:30:00',
          timeZone: '/vendor.example/20050126_1/America/New_York',
        },
        windowStart,
        windowEnd,
        'Australia/Sydney',
      ),
      false,
    );
  });

  it('inherits START\'s zone for an end carrying no endTimeZone of its own', () => {
    // `endTimeZone` is emitted ONLY when it differs from start's, so undefined means "same as
    // start" and must NOT fall back to the configured zone. Reading this end as Sydney time
    // puts it before its own start; Math.max then collapses the event to a zero-width instant
    // at 14:30Z, which a window covering only its TAIL does not contain.
    const windowStart = Date.parse('2027-03-05T15:00:00Z');
    const windowEnd = Date.parse('2027-03-05T16:00:00Z');
    assert.equal(
      eventIntersectsWindow(
        // 09:30-10:30 New York on 5 March 2027 is 14:30Z-15:30Z (EST; DST starts on the 14th).
        { start: '2027-03-05T09:30:00', end: '2027-03-05T10:30:00', timeZone: 'America/New_York' },
        windowStart,
        windowEnd,
        'Australia/Sydney',
      ),
      true,
    );
  });

  it('places an unresolvable zone name in the CONFIGURED zone, never the host\'s', () => {
    // A Windows zone id is passed through verbatim rather than rejected, and `zoneOffsetMsAt`
    // silently resolves a name it cannot parse against the HOST zone — so dropping the
    // usability guard would place this one event in whichever zone the deployment runs in.
    // Asserted from two different configured zones so that neither answer can be the host's
    // by coincidence: whichever machine runs this, at most one of them is the host.
    const event = {
      start: '2027-03-05T09:30:00',
      end: '2027-03-05T10:30:00',
      timeZone: 'AUS Eastern Standard Time',
    };
    // 09:30-10:30 New York on 5 March 2027 (EST, -05:00).
    const newYorkWindow = [Date.parse('2027-03-05T14:00:00Z'), Date.parse('2027-03-05T16:00:00Z')] as const;
    // 09:30-10:30 Sydney on the same date (AEDT, +11:00) is the evening of the 4th in UTC.
    const sydneyWindow = [Date.parse('2027-03-04T22:00:00Z'), Date.parse('2027-03-05T00:00:00Z')] as const;
    assert.equal(eventIntersectsWindow(event, newYorkWindow[0], newYorkWindow[1], 'America/New_York'), true);
    assert.equal(eventIntersectsWindow(event, sydneyWindow[0], sydneyWindow[1], 'Australia/Sydney'), true);
    // And the cross pairs are false, so neither assertion above can pass on a window wide
    // enough to contain both readings.
    assert.equal(eventIntersectsWindow(event, newYorkWindow[0], newYorkWindow[1], 'Australia/Sydney'), false);
    assert.equal(eventIntersectsWindow(event, sydneyWindow[0], sydneyWindow[1], 'America/New_York'), false);
  });

  it('rolls a two-digit-year all-day date over to the RIGHT next day', () => {
    // `Date.UTC(26, 7, 13)` is 1926, not the year 26, so the day after an all-day event in
    // year 0026 has to be reached by stepping over a whole Gregorian cycle and back. Get that
    // wrong and the event's end lands nineteen centuries away, making it overlap windows it
    // has nothing to do with. The window bounds resolve through the same machinery the real
    // caller's do, so both sides stay on one scale.
    const dayOf = resolveCalendarInstantMs('0026-08-12', 'UTC');
    const dayAfter = resolveCalendarInstantMs('0026-08-13', 'UTC');
    const twoDaysAfter = resolveCalendarInstantMs('0026-08-14', 'UTC');
    assert.equal(eventIntersectsWindow({ start: '0026-08-12' }, dayOf, dayAfter, 'UTC'), true);
    assert.equal(eventIntersectsWindow({ start: '0026-08-12' }, dayAfter, twoDaysAfter, 'UTC'), false);
  });

  it('still covers a whole day for an all-day event on the last representable date (#162)', () => {
    // The day after '9999-12-31' is '10000-01-01', which is not a four-digit-year date and so
    // does not resolve. That NaN used to collapse the event to local midnight, so a window
    // later that morning answered "nothing on" for an event that covers the whole day.
    assert.equal(
      eventIntersectsWindow(
        { start: '9999-12-31' },
        Date.parse('9999-12-31T09:00:00Z'),
        Date.parse('9999-12-31T12:00:00Z'),
        'UTC',
      ),
      true,
    );
  });
});

describe('CalDAVCalendarClient.getCalendarEvents caps how dense one series may be (#142)', () => {
  // Every assertion here is about which UTC instants came back, so the zone is pinned rather
  // than left to the machine.
  before(() => setDefaultTimezone('UTC'));
  after(() => setDefaultTimezone(undefined));

  const WINDOW_START = '2027-03-01T00:00:00Z';
  const WINDOW_END = '2027-03-10T00:00:00Z';

  /**
   * An expanded blob: one VEVENT block per occurrence, a minute apart, RRULE stripped.
   *
   * `uid` or `summary` given as undefined OMITS that property, which is how the refusal
   * message's fallback arms are reached: neither is required by iCalendar's grammar, so a
   * real resource can arrive without either.
   */
  function expandedBlob(uid: string | undefined, summary: string | undefined, count: number): string {
    const base = Date.parse('2027-03-01T00:10:00Z');
    const blocks: string[] = [];
    for (let i = 0; i < count; i++) {
      const at = new Date(base + i * 60000).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
      blocks.push([
        'BEGIN:VEVENT',
        ...(uid === undefined ? [] : [`UID:${uid}`]),
        `DTSTART:${at}`,
        ...(summary === undefined ? [] : [`SUMMARY:${summary}`]),
        'END:VEVENT',
      ].join('\r\n'));
    }
    return ['BEGIN:VCALENDAR', ...blocks, 'END:VCALENDAR'].join('\r\n');
  }

  function oneEvent(uid: string, summary: string, dtstart: string): string {
    return [
      'BEGIN:VCALENDAR', 'BEGIN:VEVENT', `UID:${uid}`, `DTSTART:${dtstart}`,
      `SUMMARY:${summary}`, 'END:VEVENT', 'END:VCALENDAR',
    ].join('\r\n');
  }

  function clientOver(byCalendar: Record<string, Array<{ data: string; url: string }>>) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => Object.keys(byCalendar).map(url => ({
        displayName: url === '/cal/work/' ? 'Work' : 'Personal', url,
      }))),
      fetchCalendarObjects: mock.fn(async (p: FetchObjectsParams) =>
        byCalendar[(p as any).calendar.url] ?? []),
    };
    return client;
  }

  it('rejects the call with InvalidInputError naming the series', async () => {
    // The whole listing fails rather than answering without the series. A series left out and
    // disclosed in a trailing note put the one thing the caller most needed to know at the
    // bottom of a response that otherwise looked complete; an error cannot be skimmed past.
    // The caller chooses the SPAN of a window; an attacker chooses the DENSITY inside it, and
    // anyone who can send an invitation can put a FREQ=MINUTELY series in the account.
    const client = clientOver({
      '/cal/work/': [
        { data: expandedBlob('dense@fm', 'Every minute', CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '/w-dense.ics' },
        { data: oneEvent('real@fm', 'Real meeting', '20270302T090000Z'), url: '/w-real.ics' },
      ],
      '/cal/personal/': [
        { data: oneEvent('other@fm', 'Dentist', '20270303T090000Z'), url: '/p-other.ics' },
      ],
    });

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        // InvalidInputError, not a bare Error: the handler maps that tag to InvalidParams,
        // because narrowing the window is the CALLER's action rather than a server fault.
        assert.ok(err instanceof InvalidInputError, `expected InvalidInputError, got ${err}`);
        const message = (err as Error).message;
        // Everything the caller needs to find the series and narrow around it.
        assert.match(message, /"Every minute"/);
        assert.match(message, /id dense@fm/);
        assert.match(message, new RegExp(`${CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1} occurrences`));
        assert.match(message, /calendar Work/);
        // THE LIMIT ITSELF. This message is the only place a caller is ever told the number —
        // the tool description and README deliberately stopped carrying it — so the figure is
        // pinned here, interpolated from the constant rather than written as a literal so a
        // deliberate change to the cap moves the assertion with it.
        assert.match(
          message,
          new RegExp(`more than the ${CALENDAR_MAX_OCCURRENCES_PER_SERIES} this server will materialise`),
        );
        // The cap is a judgement call, so the message says whose it is and how to contest it
        // rather than reading as a platform limit the caller can do nothing about.
        assert.match(message, /deliberate limit/);
        assert.match(message, /open an issue at https:\/\/github\.com\/JonathanGodley\/fastmail-mcp\/issues/);
        return true;
      },
    );
  });

  it('falls back to the URL when the dense series sits in a nameless calendar', async () => {
    // The message names the calendar so the caller can narrow to it. An empty
    // `<displayname/>` parses to `{}`, and stringifying that put "[object Object]" in the
    // text — truthy, so the url fallback written beside it never ran and the one field that
    // makes the refusal actionable was a marker.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: {}, url: '/cal/nameless/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => [
        { data: expandedBlob('dense@fm', 'Every minute', CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '/n-dense.ics' },
      ]),
    };

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        assert.match((err as Error).message, /calendar \/cal\/nameless\//);
        return true;
      },
    );
  });

  it('calls a dense series with no SUMMARY "Untitled"', async () => {
    // SUMMARY is optional in iCalendar, so a resource can arrive without one. The message
    // opens by quoting the title, and an empty pair of quotes there reads as a rendering
    // fault rather than as "this series has no name".
    const client = clientOver({
      '/cal/work/': [
        { data: expandedBlob('dense@fm', undefined, CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '/w-dense.ics' },
      ],
    });

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        assert.match((err as Error).message, /repeating event "Untitled"/);
        return true;
      },
    );
  });

  it('leaves the id empty when the dense resource has neither a UID nor a url', async () => {
    // The floor under the id: UID first, the resource url as the handle that normally exists
    // when there is not, and an empty string under both. Nothing may be invented there — a
    // placeholder would be echoed back as an id the caller could go looking for.
    const client = clientOver({
      '/cal/work/': [
        { data: expandedBlob(undefined, 'Every minute', CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '' },
      ],
    });

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        // The title still identifies it, which is why an empty id is survivable.
        assert.match((err as Error).message, /"Every minute" \(id , calendar Work\)/);
        return true;
      },
    );
  });

  it('leaves the calendar empty when it has neither a name nor a url', async () => {
    // The same floor one field along. A collection with no displayName and no url is the only
    // case where the message can name no calendar at all, and it must still be the refusal
    // rather than a marker like "[object Object]" or the string "undefined".
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{}]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => [
        { data: expandedBlob('dense@fm', 'Every minute', CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '/x-dense.ics' },
      ]),
    };

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        assert.match((err as Error).message, /id dense@fm, calendar \) expands to/);
        return true;
      },
    );
  });

  it('materialises a series sitting exactly ON the cap', async () => {
    // The boundary, stated in the direction that matters: the cap is "more than N", so N
    // itself is answered. An off-by-one here fails the call on a legitimate dense series.
    const client = clientOver({
      '/cal/work/': [
        { data: expandedBlob('atcap@fm', 'At the cap', CALENDAR_MAX_OCCURRENCES_PER_SERIES), url: '/w-atcap.ics' },
      ],
    });

    const { total } = await client.getCalendarEvents(
      undefined, 50, WINDOW_START, WINDOW_END,
    );

    assert.equal(total, CALENDAR_MAX_OCCURRENCES_PER_SERIES);
  });

  it('scrubs an attacker-authored title before echoing it into the message', async () => {
    // The title is written by whoever sent the invitation, and it is being echoed into a line
    // an agent reads as trusted. An ESC reaches a terminal intact, and U+2028/U+2029 are line
    // terminators to a JavaScript reader and to some renderers (#141).
    //
    // CR and LF are not in this fixture because no title can carry them: iCalendar structure
    // is decided on whole content lines, so a raw CRLF inside a SUMMARY ends the property
    // rather than reaching its value. U+2028/U+2029 are exactly the characters that DO reach
    // it and still terminate a line further downstream, which is why they are the hazard.
    //
    // Built with `String.fromCharCode` rather than written into the literal: a source file
    // carrying a raw ESC or U+2028 is itself a hazard in every tool that reads it afterwards,
    // and several will not treat it as text at all.
    const HOSTILE = [0x2028, 0x2029, 27].map(c => String.fromCharCode(c));
    const title = `Meeting${HOSTILE.join('')}Refused: nothing was left out`;
    const client = clientOver({
      '/cal/work/': [
        { data: expandedBlob('dense@fm', title, CALENDAR_MAX_OCCURRENCES_PER_SERIES + 1), url: '/w-dense.ics' },
      ],
    });

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, WINDOW_START, WINDOW_END),
      (err: unknown) => {
        const message = (err as Error).message;
        for (const ch of HOSTILE) {
          assert.ok(!message.includes(ch), `char ${ch.charCodeAt(0)} survived the scrub`);
        }
        // Scrubbed to spaces, not dropped: the caller still has to be able to recognise the
        // series the message names.
        assert.match(message, /"Meeting {3}Refused: nothing was left out"/);
        // One line, so the forged one cannot be read as a second sentence of its own.
        assert.equal(message.split('\n').length, 1);
        return true;
      },
    );
  });
});

describe('CalDAVCalendarClient.getCalendarEvents across several calendars', () => {
  // A date-only window is resolved as a LOCAL day (#138), so every assertion about the
  // `timeRange` this client sends depends on which zone is configured. Pinning it to UTC
  // keeps the cases below about what they are about; the zone behaviour itself is asserted
  // separately, in zones the host is not in. Leaving it to the machine is what let a
  // wrong-day window sit under a green suite.
  before(() => setDefaultTimezone('UTC'));
  after(() => setDefaultTimezone(undefined));

  function makeIcal(uid: string, summary: string, dtstart: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTART:${dtstart}`,
      `SUMMARY:${summary}`,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  it('queries every calendar before slicing, so limit is a genuine earliest-N', async () => {
    // The first calendar alone satisfies the limit. Under the old early break the later
    // calendars were never read at all, so an earlier event in one of them could not appear
    // and nothing said so.
    const byCalendar: Record<string, Array<{ data: string; url: string }>> = {
      '/cal/a/': [
        { data: makeIcal('a1@fm', 'A late', '20260325T200000Z'), url: '/a1.ics' },
        { data: makeIcal('a2@fm', 'A later', '20260325T210000Z'), url: '/a2.ics' },
      ],
      '/cal/b/': [
        { data: makeIcal('b1@fm', 'B earliest', '20260325T060000Z'), url: '/b1.ics' },
      ],
    };
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'A', url: '/cal/a/' },
        { displayName: 'B', url: '/cal/b/' },
      ]),
      fetchCalendarObjects: mock.fn(async (params: FetchObjectsParams) =>
        byCalendar[(params as any).calendar.url] ?? []),
    };
    (client as any).client = mockDAVClient;

    // The window is named so the subject stays "earliest N across every calendar": a
    // bounds-free call now gets today plus a month (#142), which these 2026 fixtures sit
    // outside of.
    const { events, total } = await client.getCalendarEvents(
      undefined, 2, '2026-03-25T00:00:00Z', '2026-03-26T00:00:00Z',
    );

    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 2);
    assert.equal(events.length, 2);
    assert.equal(events[0].title, 'B earliest');
    // The total reports what matched before the limit trimmed it, so a caller can see that
    // something was cut rather than reading the page as the whole answer.
    assert.equal(total, 3);
  });

  it('rejects a backwards window as caller-fixable input', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, '2027-03-10', '2027-03-01'),
      (err: Error) => {
        // InvalidInputError, not a plain Error: tsdav would throw one that reaches the tool
        // boundary as InternalError, telling the model a bare retry might work.
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('ends a date-only endDate at the end of that day, so a single-day window is not empty', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    await client.getCalendarEvents(undefined, 50, '2027-03-10', '2027-03-10');

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    // The caller's window is the 10th, 00:00Z .. 11th 00:00Z; the request carries the #162
    // margin of fourteen hours at each edge on top of it.
    assert.deepEqual(callArgs.timeRange, {
      start: '2027-03-09T10:00:00Z',
      end: '2027-03-11T14:00:00Z',
    });
  });

  it('rejects an impossible calendar date instead of rolling it into the next month', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, '2027-02-30', '2027-03-10'),
      /startDate is not a real calendar date/,
    );
  });

  it('reports every occurrence of a series whose first instance is inside the window', async () => {
    // End-to-end over the client: the server expanded a yearly series whose original DTSTART
    // falls in the window, so the first block carries no RECURRENCE-ID. Collapsing that blob
    // to its "master" answered a five-year question with one event and no sign anything was
    // missing — the total is computed after the drop, so nothing disclosed it.
    const occurrence = (date: string, withRecurrenceId: boolean) => [
      'BEGIN:VEVENT',
      'UID:birthday@fm',
      `DTSTART;VALUE=DATE:${date}`,
      ...(withRecurrenceId ? [`RECURRENCE-ID;VALUE=DATE:${date}`] : []),
      'SUMMARY:Birthday',
      'END:VEVENT',
    ].join('\r\n');
    const data = [
      'BEGIN:VCALENDAR',
      occurrence('19940612', false),
      occurrence('19950612', true),
      occurrence('19960612', true),
      occurrence('19970612', true),
      occurrence('19980612', true),
      'END:VCALENDAR',
    ].join('\r\n');

    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => [{ data, url: '/b.ics' }]),
    };

    const { events, total } = await client.getCalendarEvents(undefined, 50, '1994-01-01', '1998-12-31');

    assert.equal(events.length, 5);
    assert.equal(total, 5);
    assert.deepEqual(events.map(e => e.start), [
      '1994-06-12', '1995-06-12', '1996-06-12', '1997-06-12', '1998-06-12',
    ]);
    assert.ok(events.every(e => e.isRecurring === true));
  });

  it('bounds a window given only a startDate, and says so', async () => {
    // With `expand` on, the missing half is the range the SERVER materialises occurrences
    // over — a 2099 default asked Fastmail to generate every occurrence of every repeating
    // event for 73 years, and `limit` bounds none of that work. The invented half is one
    // month: 2027-03-01 plus 31 fixed days is 2027-04-01, and the request range is that
    // window widened by fourteen hours at each edge.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    const { windowClamp } = await client.getCalendarEvents(undefined, 50, '2027-03-01');

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    assert.deepEqual(callArgs.timeRange, {
      start: '2027-02-28T10:00:00Z',
      end: '2027-04-01T14:00:00Z',
    });
    // Never silently narrowed: a caller handed a shorter window than it asked for has to be
    // told, or "nothing after that date" reads as an empty calendar.
    //
    // AND THE CLAMP REPORTS THE CALLER'S WINDOW, not the widened request range (#162). The
    // margin is a fact about what is asked of the server; describing the caller's window as
    // fourteen hours wider than they asked for would be a false disclosure.
    assert.ok(windowClamp, 'a clamped window must be disclosed');
    assert.equal(windowClamp!.invented, 'endDate');
    assert.equal(windowClamp!.start, '2027-03-01T00:00:00Z');
    assert.equal(windowClamp!.end, '2027-04-01T00:00:00Z');
    assert.equal(windowClamp!.saturated, undefined);
  });

  it('bounds a window given only an endDate, backwards from that bound', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    const { windowClamp } = await client.getCalendarEvents(undefined, 50, undefined, '2027-03-10');

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    // A date-only endDate still covers the whole of the 10th, so the exclusive end is the
    // following midnight and the clamp counts back a month from there: 2027-03-11 less 31
    // fixed days is 2027-02-08.
    assert.deepEqual(callArgs.timeRange, {
      start: '2027-02-07T10:00:00Z',
      end: '2027-03-11T14:00:00Z',
    });
    assert.ok(windowClamp, 'a clamped window must be disclosed');
    assert.equal(windowClamp!.invented, 'startDate');
    assert.equal(windowClamp!.start, '2027-02-08T00:00:00Z');
    assert.equal(windowClamp!.end, '2027-03-11T00:00:00Z');
  });

  it('leaves a fully specified window alone and discloses nothing', async () => {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    // Deliberately far longer than the clamp: a span the caller named is the caller's own
    // decision, and the bound applies only to a half the caller did not give.
    const { windowClamp } = await client.getCalendarEvents(undefined, 50, '2020-01-01', '2029-12-31');

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    assert.deepEqual(callArgs.timeRange, {
      start: '2019-12-31T10:00:00Z',
      end: '2030-01-01T14:00:00Z',
    });
    // The request margin is not a clamp: the caller's own window was honoured exactly, so
    // there is nothing to disclose.
    assert.equal(windowClamp, undefined);
  });

  it('asks the server for the caller\'s LOCAL day, not the UTC day (#138)', async () => {
    // The reported failure, end to end. On a +10:00 account, `2026-08-12` used to be sent
    // as 12 Aug 00:00Z .. 13 Aug 00:00Z — which is 12 Aug 10:00 to 13 Aug 10:00 where the
    // user lives, so an 08:00 appointment on the 12th was outside the window. Two of that
    // day's three appointments were missing and the day read as a quiet one.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    setDefaultTimezone('Australia/Sydney');
    try {
      await client.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-12');
    } finally {
      setDefaultTimezone('UTC');
    }

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    // The local day is 11 Aug 14:00Z .. 12 Aug 14:00Z, and the #162 margin puts fourteen
    // hours either side of it on the wire.
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-08-11T00:00:00Z',
      end: '2026-08-13T04:00:00Z',
    });
  });

  it('resolves the local day the other way for a zone behind UTC (#138)', async () => {
    // The mirror case, so a sign error cannot pass both.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    setDefaultTimezone('America/New_York');
    try {
      await client.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-12');
    } finally {
      setDefaultTimezone('UTC');
    }

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-08-11T14:00:00Z',
      end: '2026-08-13T18:00:00Z',
    });
  });

  it('leaves a zone-designated window untouched whatever the configured zone is (#138)', async () => {
    // A caller that named an instant named an instant. This is the escape hatch from the
    // local-day rule, so it has to hold in a zone that is not UTC.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    setDefaultTimezone('Australia/Sydney');
    try {
      await client.getCalendarEvents(undefined, 50, '2026-08-11T14:00:00Z', '2026-08-12T14:00:00Z');
    } finally {
      setDefaultTimezone('UTC');
    }

    const callArgs = callArguments(mockDAVClient.fetchCalendarObjects)[0];
    // The instants the caller named are taken exactly as written and the #162 margin is added
    // to the REQUEST around them — the margin is about what the server will withhold, not
    // about how the caller's bounds are read, so it applies to a Z-designated window too.
    assert.deepEqual(callArgs.timeRange, {
      start: '2026-08-11T00:00:00Z',
      end: '2026-08-13T04:00:00Z',
    });
  });

  it('names the zone when rejecting a backwards window, so the resolved range reads right', async () => {
    // Without the zone stated, a caller who passed two plain dates and got back two UTC
    // instants offset from midnight cannot tell a correct local-day reading from a bug.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    setDefaultTimezone('Australia/Sydney');
    try {
      await assert.rejects(
        () => client.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-10'),
        /Dates are read as whole days in Australia\/Sydney/,
      );
    } finally {
      setDefaultTimezone('UTC');
    }
  });

  it('rejects a calendarId that matches nothing instead of answering "you are free"', async () => {
    // An unresolvable calendarId used to leave the target list empty, so the fetch loop never
    // ran and the tool returned "Showing 0 of 0 results." — a typo answered as an empty
    // calendar. The write path always threw here; both now raise the same error.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Work', url: '/cal/work/' },
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.getCalendarEvents('work', 50, '2027-03-01', '2027-03-10'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Calendar not found: "work"/);
        // The available names are listed so a case or spacing miss is fixable in one retry
        // rather than needing a separate list_calendars call.
        assert.match(err.message, /"Work"/);
        assert.match(err.message, /"Personal"/);
        return true;
      },
    );
    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 0);
  });

  it('echoes the caller\'s own values when rejecting a backwards window', async () => {
    // Reporting the post-coercion bound quoted "endDate 2026-08-11T00:00:00Z" at someone who
    // passed 2026-08-10, so the sentence explaining the whole-day rule cited a value they
    // could not find in their own call.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-10'),
      (err: Error) => {
        assert.match(err.message, /startDate "2026-08-12"/);
        assert.match(err.message, /endDate "2026-08-10"/);
        // The resolved range is still shown, so the whole-day rule stays explained.
        assert.match(err.message, /2026-08-11T00:00:00Z/);
        return true;
      },
    );
  });

  it('drops an out-of-window event the server returned anyway', async () => {
    // A resource the server matched on an occurrence but returned whole. Nothing in the
    // payload places it in the window, so it must not be listed as though it were.
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'A', url: '/cal/a/' }]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => [
        { data: makeIcal('old@fm', 'Pay Day', '20200828T000000Z'), url: '/old.ics' },
        { data: makeIcal('now@fm', 'In window', '20270305T090000Z'), url: '/now.ics' },
      ]),
    };

    const { events, total } = await client.getCalendarEvents(undefined, 50, '2027-03-01', '2027-03-10');

    assert.equal(events.length, 1);
    assert.equal(events[0].title, 'In window');
    assert.equal(total, 1);
  });
});

describe('CalDAVCalendarClient calendar discovery failures', () => {
  function clientWithEmptyDiscovery(propfindResponses: unknown) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      account: { homeUrl: 'https://caldav.example.com/dav/calendars/user/' },
      // tsdav filters a non-2xx PROPFIND out on resourcetype and returns [] without throwing.
      fetchCalendars: mock.fn(async () => []),
      propfind: mock.fn(async (_p: unknown) => propfindResponses),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('throws instead of reporting an empty list when the discovery PROPFIND failed', async () => {
    const { client } = clientWithEmptyDiscovery([{ status: 503, statusText: 'Service Unavailable' }]);
    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50),
      /discover calendars: server returned 503/,
    );
  });

  it('throws when discovery succeeds but lists no calendars', async () => {
    // An empty-but-successful discovery is still an error: answering an availability
    // question with nothing reads as free time.
    const { client } = clientWithEmptyDiscovery([{ status: 207 }]);
    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50),
      /Calendar discovery returned no calendars/,
    );
  });

  it('does not cache an empty discovery result', async () => {
    // `[]` is truthy, so the previous `if (!this.calendars)` guard cached a failed discovery
    // on this long-lived client and repeated it until the process restarted.
    const { client, mockDAVClient } = clientWithEmptyDiscovery([{ status: 207 }]);
    await assert.rejects(() => client.getCalendarEvents(undefined, 50));
    await assert.rejects(() => client.getCalendarEvents(undefined, 50));
    assert.equal(mockDAVClient.fetchCalendars.mock.callCount(), 2);
  });

  it('does not blame the caller event id when discovery failed', async () => {
    // The old path returned null from the lookup and raised InvalidInputError, telling the
    // model to fix an id that was never the problem.
    const { client } = clientWithEmptyDiscovery([{ status: 503, statusText: 'Service Unavailable' }]);
    await assert.rejects(
      () => client.getCalendarEventById('real-id@fm'),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /discover calendars/);
        return true;
      },
    );
  });
});

// Component splitting runs on FOLDED text — RFC 5545 unfolding happens later, per property —
// so the markers that bound a VEVENT have to be line-anchored or the payload can name its own
// boundaries. Calendar content here is authored by anyone who can send an invitation.
describe('VEVENT splitting is line-anchored against folded content', () => {
  // A DESCRIPTION folded so the continuation line begins with the literal END:VEVENT text.
  // Deterministic to construct: libical folds at a fixed octet count, so a description padded
  // to the right length puts the fold exactly there.
  // A DESCRIPTION whose folded continuation lines carry the two component markers. The
  // property order is the attacker's to choose, so the real SUMMARY and DTSTART sit AFTER
  // them: unanchored, the payload's own text ends the component early and starts a second
  // one, and the real event's properties land in the phantom.
  const FOLDED_TERMINATOR = [
    'BEGIN:VCALENDAR',
    'BEGIN:VEVENT',
    'DESCRIPTION:hello',
    ' END:VEVENT',
    ' BEGIN:VEVENT',
    'UID:real@fm',
    'DTSTART:20260325T080000Z',
    'SUMMARY:Lunch',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');

  it('does not split one component in two at a folded END:VEVENT', () => {
    const events = parseCalendarObjects({ data: FOLDED_TERMINATOR, url: '/u.ics' } as any, { expanded: true });
    assert.equal(events.length, 1, JSON.stringify(events));
    assert.equal(events[0].id, 'real@fm');
    assert.equal(events[0].title, 'Lunch');
    assert.equal(events[0].start, '2026-03-25T08:00:00Z');
    // The markers are description TEXT once the property is unfolded, which is all they ever
    // were.
    assert.equal(events[0].description, 'helloEND:VEVENTBEGIN:VEVENT');
    // The phantom used to arrive as a row of its own — an "Untitled" event with no dates,
    // which the window filter keeps because it has nothing to judge — AND to mark this
    // genuine one-off as recurring, because two blocks "prove" a series.
    assert.equal(events[0].isRecurring, undefined);
  });

  it('reads the whole property when extracting a single VEVENT from folded data', () => {
    const vevent = extractVEvent(FOLDED_TERMINATOR);
    assert.ok(vevent);
    assert.equal(parseICalValue(vevent!, 'SUMMARY'), 'Lunch');
    assert.equal(parseICalValue(vevent!, 'UID'), 'real@fm');
  });

  it('does not open a block at a BEGIN:VEVENT written inside another property', () => {
    // The mirror image: text BEFORE the first real component claiming to start one. In a
    // VTIMEZONE's TZNAME it opens the block ahead of the zone rule's own DTSTART, and that
    // is the DTSTART the parser then reports — the event came back dated 1970 and the window
    // filter dropped it, answering "you are free" for a day that was not.
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VTIMEZONE',
      'TZID:Fake/Zone',
      'BEGIN:STANDARD',
      'TZNAME:xBEGIN:VEVENT',
      'DTSTART:19700101T000000',
      'END:STANDARD',
      'END:VTIMEZONE',
      'BEGIN:VEVENT',
      'UID:real@fm',
      'DTSTART:20260325T080000Z',
      'SUMMARY:Lunch',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const events = parseCalendarObjects({ data, url: '/u.ics' } as any, { expanded: true });
    assert.equal(events.length, 1, JSON.stringify(events));
    assert.equal(events[0].title, 'Lunch');
    assert.equal(events[0].start, '2026-03-25T08:00:00Z');
  });
});

// RFC 5545 knows one line break: CRLF (and a bare LF, tolerated). JavaScript's `/m` anchors
// know three more — U+2028, U+2029 and a lone CR — and none of those three has to be escaped
// inside an iCalendar TEXT value, so all three survive verbatim into a SUMMARY or DESCRIPTION
// the event's author chose. Every one of these payloads is ONE VEVENT by the RFC's reading.
//
// These tests use the code points EXPLICITLY because the fold tests above pass either way:
// a `/m`-anchored parser and a line-based one both handle a folded terminator correctly, so
// nothing already in this file would notice a regression back to `/m`.
describe('non-RFC line terminators inside a value are text, not structure', () => {
  // Written as escapes on purpose: a literal U+2028 or U+2029 in this source is invisible in
  // an editor and indistinguishable from a space in review, which is most of why the class is
  // worth a test at all.
  const TERMINATORS: Array<[string, string]> = [
    ['U+2028 LINE SEPARATOR', '\u2028'],
    ['U+2029 PARAGRAPH SEPARATOR', '\u2029'],
    ['a bare CR', '\r'],
  ];

  // The UID/DTSTART forgery, which is the severe one: this resource is the ATTACKER'S, and
  // the properties smuggled into its SUMMARY name a DIFFERENT record. `delete_calendar_event`
  // and `update_calendar_event` resolve an id through findCalendarObjectByUID, which returns
  // the first object in the ACCOUNT whose UID matches — so a caller told "this id names the
  // series, act on it" destroys a record it never named and mails its attendees a cancellation.
  function spoofingPayload(sep: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      // The forged property lines come FIRST, because property order is the author's to
      // choose and every read here takes the FIRST match in the block.
      `SUMMARY:Lunch${sep}UID:board-meeting@victim.example${sep}DTSTART:20260812T090000Z${sep}DTEND:20260812T093000Z`,
      'UID:attacker-own@example.invalid',
      'DTSTART:20260901T100000Z',
      'DTEND:20260901T110000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  // The component-forging one: the value carries a whole END/BEGIN pair, so the block is cut
  // in two and the REAL DTSTART/DTEND land in the discarded tail.
  function splittingPayload(sep: string): string {
    return [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:real@fm',
      `DESCRIPTION:hello${sep}END:VEVENT${sep}BEGIN:VEVENT${sep}UID:phantom@fm${sep}DTSTART:20200101T000000Z`,
      'DTSTART:20260901T100000Z',
      'DTEND:20260901T110000Z',
      'SUMMARY:Real',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
  }

  for (const [name, sep] of TERMINATORS) {
    it(`does not let ${name} in a SUMMARY forge another resource's UID`, () => {
      const data = spoofingPayload(sep);
      const events = parseCalendarObjects({ data, url: '/cal/attacker.ics' } as any, {});
      assert.equal(events.length, 1, JSON.stringify(events));
      assert.equal(events[0].id, 'attacker-own@example.invalid');
      // The dates are the resource's own, not the ones smuggled into the summary.
      assert.equal(events[0].start, '2026-09-01T10:00:00Z');
      assert.equal(events[0].end, '2026-09-01T11:00:00Z');
      // Read directly too: parseICalValue takes the FIRST match in the block, which is where
      // the spoof landed.
      assert.equal(parseICalValue(extractVEvent(data)!, 'UID'), 'attacker-own@example.invalid');
      assert.equal(parseICalValue(extractVEvent(data)!, 'DTSTART'), '20260901T100000Z');
    });

    it(`does not let ${name} in a DESCRIPTION split one component into two`, () => {
      const data = splittingPayload(sep);
      const events = parseCalendarObjects({ data, url: '/u.ics' } as any, { expanded: true });
      assert.equal(events.length, 1, JSON.stringify(events));
      assert.equal(events[0].id, 'real@fm');
      assert.equal(events[0].title, 'Real');
      // Start and end used to vanish entirely — they sat in the discarded tail — and a row
      // with no dates survives the window filter, so it displayed as an undated event inside
      // a window nothing had placed it in.
      assert.equal(events[0].start, '2026-09-01T10:00:00Z');
      assert.equal(events[0].end, '2026-09-01T11:00:00Z');
      // And the second, wholly attacker-authored row is gone with it, so the real event is
      // no longer marked recurring by "two blocks prove a series".
      assert.equal(events[0].isRecurring, undefined);
    });

    it(`does not let ${name} forge a RECURRENCE-ID that reorders the write path`, () => {
      // normalizeMasterVEventFirst swaps blocks when block 0 looks like an override. A forged
      // RECURRENCE-ID inside block 0's SUMMARY made it look like one, so every in-place patch
      // helper — which all target the FIRST VEVENT — was aimed at the wrong component.
      const data = [
        'BEGIN:VCALENDAR',
        'BEGIN:VEVENT',
        'UID:evt@fm',
        `SUMMARY:Standup${sep}RECURRENCE-ID:20260401T100000Z`,
        'RRULE:FREQ=WEEKLY',
        'DTSTART:20260401T100000Z',
        'END:VEVENT',
        'BEGIN:VEVENT',
        'UID:evt@fm',
        'DTSTART:20260408T100000Z',
        'END:VEVENT',
        'END:VCALENDAR',
      ].join('\r\n');
      assert.equal(normalizeMasterVEventFirst(data), data, 'block 0 is the master and must stay put');
    });
  }

  // The presence test itself, asserted directly rather than through a reader that was never
  // `/m`-based. The RRULE / RDATE / ORGANIZER / ATTENDEE gates on the write path all go through
  // hasICalProperty, and it is the ONLY thing they go through \u2014 an earlier version of this
  // test drove parseAllICalProperties instead, which passed identically before and after the
  // fix and left every one of those gates unpinned. All four are also driven end to end through
  // updateCalendarEvent: ORGANIZER and ATTENDEE in the patch-based suite, because they steer the
  // patch, and the recurrence markers in the refusal suite ("does not treat an RRULE forged
  // inside a SUMMARY as a recurrence", and its RDATE pair), because a marker the gate reads as
  // present short-circuits into the REFUSAL, which is itself observable. RDATE is here for the same
  // reason it is there — it fronts the same irreversible destroy as RRULE:
  // isRecurringSeriesResource takes it as a recurrence marker too, so a forged RDATE in a TEXT
  // value reaching that gate would refuse a legitimate edit or delete of a one-off event.
  for (const [name, sep] of TERMINATORS) {
    it(`treats a property forged after ${name} inside a value as absent, not present`, () => {
      const vevent = [
        'BEGIN:VEVENT',
        'UID:e@fm',
        `SUMMARY:Standup${sep}RRULE:FREQ=WEEKLY${sep}ORGANIZER:mailto:nobody@example.invalid`,
        `LOCATION:Room 1${sep}RDATE:20260408T100000Z`,
        `DESCRIPTION:notes${sep}ATTENDEE;CN=Nobody:mailto:nobody@example.invalid`,
        'END:VEVENT',
      ].join('\r\n');
      for (const key of ['RRULE', 'RDATE', 'ORGANIZER', 'ATTENDEE']) {
        assert.equal(hasICalProperty(vevent, key), false, `${key} was read out of a value`);
      }
      // \u2026and a real one on its own line is still found, so the guard cannot be "fixed" into
      // never matching.
      const real = vevent.replace(
        'UID:e@fm',
        'UID:e@fm\r\nRRULE:FREQ=DAILY\r\nRDATE:20260415T100000Z',
      );
      assert.equal(hasICalProperty(real, 'RRULE'), true);
      assert.equal(hasICalProperty(real, 'RDATE'), true);
    });
  }

  it('reads no ATTENDEE from a payload whose only ATTENDEE text sits in a DESCRIPTION', () => {
    const vevent = [
      'BEGIN:VEVENT',
      'UID:e@fm',
      'DESCRIPTION:notes\u2028ATTENDEE;CN=Nobody:mailto:nobody@example.invalid\u2028ORGANIZER:mailto:nobody@example.invalid',
      'END:VEVENT',
    ].join('\r\n');
    assert.deepEqual(parseAllICalProperties(vevent, 'ATTENDEE'), []);
    assert.deepEqual(parseAllICalProperties(vevent, 'ORGANIZER'), []);
  });

  it('still splits on a bare LF, which real servers emit even though the RFC says CRLF', () => {
    const data = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:lf@fm',
      'SUMMARY:Lunch',
      'DTSTART:20260325T080000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\n');
    const events = parseCalendarObjects({ data, url: '/u.ics' } as any, {});
    assert.equal(events.length, 1);
    assert.equal(events[0].title, 'Lunch');
    assert.equal(events[0].start, '2026-03-25T08:00:00Z');
  });
});

// `limit` slices this list, so the ORDER decides which events a caller never sees.
describe('sortEventsByStart orders by the instant, not the spelling', () => {
  it('places a TZID-stripped wall clock against a UTC value correctly', () => {
    // The two spellings a real calendar mixes: Fastmail stores a user-created event against a
    // named zone (whose TZID this read drops, leaving bare digits) and an external invitation
    // usually arrives in UTC. On an account well ahead of UTC the bare 08:30 is hours EARLIER
    // than the 08:00Z beside it; compared as strings it sorted after it.
    const events = [
      { id: 'utc', title: 'UTC', start: '2026-03-25T08:00:00Z' },
      { id: 'local', title: 'Local', start: '2026-03-25T08:30:00' },
    ] as any[];
    sortEventsByStart(events, 'Australia/Sydney');
    assert.deepEqual(events.map(e => e.id), ['local', 'utc']);
    // And the other way in a zone behind UTC, so this cannot pass on a sign error.
    sortEventsByStart(events, 'America/New_York');
    assert.deepEqual(events.map(e => e.id), ['utc', 'local']);
  });

  it('places an all-day date at local midnight, on the same scale as the rest', () => {
    const events = [
      { id: 'timed', title: 'Timed', start: '2026-03-25T00:30:00Z' },
      { id: 'allday', title: 'All day', start: '2026-03-25' },
    ] as any[];
    // Sydney midnight on the 25th is 2026-03-24T13:00Z, before the 00:30Z value.
    sortEventsByStart(events, 'Australia/Sydney');
    assert.deepEqual(events.map(e => e.id), ['allday', 'timed']);
  });

  it('keeps an unreadable start where the limit cannot cut it off', () => {
    // It cannot be placed; placing it last would put it exactly where the slice truncates,
    // turning "cannot order this" into "dropped this".
    const events = [
      { id: 'good', title: 'Good', start: '2026-03-25T08:00:00Z' },
      { id: 'bad', title: 'Bad' },
    ] as any[];
    sortEventsByStart(events, 'Australia/Sydney');
    assert.deepEqual(events.map(e => e.id), ['bad', 'good']);
  });

  it('orders by an event\'s OWN named zone (#139), not the configured zone passed in', () => {
    // Both starts read as the same bare wall-clock digits, but one event carries its own
    // timeZone naming a zone hours behind the one passed as the fallback. Sorting on the
    // fallback alone would treat them as simultaneous; sorting on each event's own zone
    // does not.
    const events = [
      { id: 'auckland', title: 'NZ', start: '2026-03-25T08:00:00', timeZone: 'Pacific/Auckland' },
      { id: 'sydney', title: 'AU', start: '2026-03-25T08:00:00' },
    ] as any[];
    sortEventsByStart(events, 'Australia/Sydney');
    // Auckland is ahead of Sydney, so the same wall-clock digits there are an EARLIER instant.
    assert.deepEqual(events.map(e => e.id), ['auckland', 'sydney']);
  });

  it('falls back to the configured zone for an event whose own timeZone cannot be resolved', () => {
    // zoneOffsetMsAt silently resolves an unusable name against the HOST zone, not the
    // configured one — sorting a Windows TZID straight through it would place the event in
    // whichever zone the test happens to run on. The isUsableTimezone guard is what keeps
    // this event pinned to the configured (Sydney) fallback instead.
    const events = [
      { id: 'windows', title: 'Third-party', start: '2026-03-25T08:00:00', timeZone: 'AUS Eastern Standard Time' },
      { id: 'sydney', title: 'AU', start: '2026-03-25T08:00:00' },
    ] as any[];
    sortEventsByStart(events, 'Australia/Sydney');
    // Read in the same (Sydney) fallback zone, identical wall-clock digits are simultaneous,
    // so the pre-existing stable order is preserved rather than either one jumping ahead.
    assert.deepEqual(events.map(e => e.id), ['windows', 'sydney']);
  });

  it('reads a leading-slash TZID the same way the window filter does (#162)', () => {
    // The sort and the filter share one zone rule (`zoneForValue`) rather than each stating
    // its own, because a read that ordered an event in one zone and filtered it in another
    // would put it at the wrong place in a list `limit` then truncates. This is the case that
    // separated them: the RFC 5545 global form of a TZID, which ICU rejects until the leading
    // slash is stripped.
    const events = [
      { id: 'newyork', title: 'US', start: '2026-03-25T08:00:00', timeZone: '/America/New_York' },
      { id: 'sydney', title: 'AU', start: '2026-03-25T08:00:00' },
    ] as any[];
    sortEventsByStart(events, 'Australia/Sydney');
    // New York is behind Sydney, so the same wall-clock digits there are a LATER instant.
    // Falling back to Sydney would make the two simultaneous and leave the input order.
    assert.deepEqual(events.map(e => e.id), ['sydney', 'newyork']);
  });
});

describe('CalDAVCalendarClient.getCalendarEvents argument and bound edges', () => {
  before(() => setDefaultTimezone('Australia/Sydney'));
  after(() => setDefaultTimezone(undefined));

  function mockedClient(objects: Array<{ data: string; url: string }> = []) {
    const client = new CalDAVCalendarClient({ username: 'test', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Work', url: '/cal/work/' },
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => objects),
    };
    (client as any).client = mockDAVClient;
    return { client, mockDAVClient };
  }

  it('rejects an empty calendarId rather than widening the query to every calendar', async () => {
    // `calendarId` narrows what the call touches, so it fails CLOSED. Under a truthiness test
    // an empty string skipped the filter and silently read the whole account, while a
    // whitespace-only one was rejected — the same mistake answered from two different
    // calendars.
    const { client, mockDAVClient } = mockedClient();
    for (const value of ['', '   ']) {
      await assert.rejects(
        () => client.getCalendarEvents(value, 50, '2027-03-01', '2027-03-10'),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, /Calendar not found/);
          return true;
        },
        `expected rejection for ${JSON.stringify(value)}`,
      );
    }
    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 0);
  });

  it('matches a calendarId with surrounding whitespace trimmed', async () => {
    const { client, mockDAVClient } = mockedClient();
    await client.getCalendarEvents('  Work  ', 50, '2027-03-01', '2027-03-10');
    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 1);
    assert.equal(callArguments(mockDAVClient.fetchCalendarObjects)[0].calendar.url, '/cal/work/');
  });

  it('says a same-instant window is zero-length instead of printing one range twice', async () => {
    const { client } = mockedClient();
    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, '2026-08-12T09:00:00Z', '2026-08-12T09:00:00Z'),
      (err: Error) => {
        assert.match(err.message, /both resolve to the same instant, 2026-08-12T09:00:00Z/);
        assert.match(err.message, /zero-length window/);
        return true;
      },
    );
  });

  it('saturates the invented half of the window at the end of the representable range', async () => {
    // Shifting a year past 9999 produces the expanded ISO form (+010000-…), which tsdav
    // rejects with a plain Error — an InternalError raised over a caller-fixable argument.
    const { client, mockDAVClient } = mockedClient();
    await client.getCalendarEvents(undefined, 50, '9999-12-30');
    const range = callArguments(mockDAVClient.fetchCalendarObjects)[0].timeRange!;
    assert.equal(range.end, '9999-12-31T23:59:59Z');
    assert.ok(!range.end.startsWith('+'), range.end);
  });

  it('saturates a CALLER-NAMED bound too, and says that it did', async () => {
    // The saturation used to cover the invented half only. A caller bound is not immune: it
    // resolves through a zone, so an offset alone pushes `9999-12-31` over the end of the
    // four-digit-year range and tsdav answered a caller-fixable argument with a plain Error.
    setDefaultTimezone('America/New_York');
    try {
      const { client, mockDAVClient } = mockedClient();
      const { windowClamp } = await client.getCalendarEvents(undefined, 50, '2026-08-12', '9999-12-31');
      const range = callArguments(mockDAVClient.fetchCalendarObjects)[0].timeRange!;
      assert.equal(range.end, '9999-12-31T23:59:59Z');
      assert.ok(!range.end.startsWith('+'), range.end);
      // Named, not silent: unlike the invented half this is a bound the caller DID choose.
      assert.ok(windowClamp, 'a saturated caller bound must be disclosed');
      assert.deepEqual(windowClamp!.saturated, [{ bound: 'endDate', edge: 'latest' }]);
      assert.equal(windowClamp!.invented, undefined);
    } finally {
      setDefaultTimezone('Australia/Sydney');
    }
  });

  // The other end, which had no coverage: a zone AHEAD of UTC pushes an early date off the
  // bottom the same way a zone behind it pushes a late one off the top. The clamp carries the
  // EDGE because the disclosure is an opposite statement at each end — knowing only the top
  // one, the note told this caller their startDate had "resolved past the LAST date this
  // server can express".
  it('saturates a caller-named bound at the EARLIEST edge, and says which edge that was', async () => {
    const { client, mockDAVClient } = mockedClient();
    const { windowClamp } = await client.getCalendarEvents(undefined, 50, '0000-01-01', '0001-01-01');
    const range = callArguments(mockDAVClient.fetchCalendarObjects)[0].timeRange!;
    assert.equal(range.start, '0000-01-01T00:00:00Z');
    assert.ok(!range.start.startsWith('-'), range.start);
    assert.ok(windowClamp, 'a saturated caller bound must be disclosed');
    assert.deepEqual(windowClamp!.saturated, [{ bound: 'startDate', edge: 'earliest' }]);
    assert.equal(windowClamp!.invented, undefined);
  });

  it('rejects a one-sided window that saturation collapses to zero length', async () => {
    // The inversion check used to be the `else if` alternative to the clamp, so a one-sided
    // window never reached it — and the comment beside it claimed a single bound "is clamped
    // above, never inverted". Saturation makes that false: a startDate on the last
    // representable instant leaves the invented month nowhere to go, and tsdav answered
    // with a plain Error (InternalError) over what is a caller-fixable bound.
    const { client, mockDAVClient } = mockedClient();
    await assert.rejects(
      () => client.getCalendarEvents(undefined, 50, '9999-12-31T23:59:59Z'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /zero-length window/);
        // The one-sided arm gets its own advice: "put startDate before endDate" is
        // unfollowable when only one was passed.
        assert.match(err.message, /Pass both startDate and endDate/);
        // And the omitted bound is named as omitted, not quoted as "undefined".
        assert.match(err.message, /endDate \(omitted\)/);
        return true;
      },
    );
    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 0);
  });

  it('keeps an unexpanded series master the window filter would otherwise drop', async () => {
    // The windowed path runs the residue filter over blocks the server was asked to EXPAND.
    // If a server ever declines to expand one, the master arrives carrying its RRULE and its
    // ORIGINAL DTSTART — years before the window for a long-running weekly event — and judging
    // that date deleted the row. The filter's whole rule is "drop only what provably cannot
    // intersect"; a recurrence rule is proof that DTSTART is not the only date the event has.
    const master = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:weekly@fm',
      'SUMMARY:Standup',
      'RRULE:FREQ=WEEKLY;BYDAY=MO',
      'DTSTART:20200106T090000Z',
      'DTEND:20200106T093000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = mockedClient([{ data: master, url: '/cal/weekly.ics' }]);
    const { events, total } = await client.getCalendarEvents(undefined, 50, '2027-03-01', '2027-03-10');
    // Two target calendars in the mock, both returning the same object.
    assert.equal(total, 2, JSON.stringify(events));
    assert.equal(events[0].id, 'weekly@fm');
    assert.equal(events[0].recurrenceRule, 'FREQ=WEEKLY;BYDAY=MO');

    // The control: strip the RRULE and the same out-of-window row IS dropped, so this is the
    // recurrence guard doing the work rather than the filter having been switched off.
    const oneOff = master.replace('RRULE:FREQ=WEEKLY;BYDAY=MO\r\n', '');
    const { client: c2 } = mockedClient([{ data: oneOff, url: '/cal/oneoff.ics' }]);
    const { total: t2 } = await c2.getCalendarEvents(undefined, 50, '2027-03-01', '2027-03-10');
    assert.equal(t2, 0);
  });

  it('keeps an unexpanded master that recurs by RDATE alone, which carries no rule (#162)', async () => {
    // A series may list its occurrences individually instead of stating a rule. Such a master
    // has no RRULE at all, so an RRULE-only guard read it as an ordinary one-off, judged it on
    // its original DTSTART years before the window, and deleted the row — the missing-event
    // direction the guard exists to prevent.
    const master = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:listed@fm',
      'SUMMARY:Board meeting',
      'DTSTART:20200106T090000Z',
      'DTEND:20200106T103000Z',
      'RDATE:20270302T090000Z,20270309T090000Z',
      'RDATE:20270316T090000Z',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = mockedClient([{ data: master, url: '/cal/listed.ics' }]);
    const { events, total } = await client.getCalendarEvents(undefined, 50, '2027-03-01', '2027-03-10');

    // Two target calendars in the mock, both returning the same object.
    assert.equal(total, 2, JSON.stringify(events));
    assert.equal(events[0].id, 'listed@fm');
    // EVERY RDATE line, not the first: the property may repeat, and a first-match read would
    // report two of the three dates this event has.
    assert.equal(events[0].recurrenceDates, '20270302T090000Z,20270309T090000Z,20270316T090000Z');
    assert.equal(events[0].isRecurring, true);
    assert.equal(events[0].recurrenceRule, undefined);

    // The control: strip the RDATEs and the same out-of-window row IS dropped, so this is the
    // recurrence guard doing the work rather than the filter having been switched off.
    const oneOff = master
      .replace('RDATE:20270302T090000Z,20270309T090000Z\r\n', '')
      .replace('RDATE:20270316T090000Z\r\n', '');
    const { client: c2 } = mockedClient([{ data: oneOff, url: '/cal/oneoff.ics' }]);
    const { total: t2 } = await c2.getCalendarEvents(undefined, 50, '2027-03-01', '2027-03-10');
    assert.equal(t2, 0);
  });

  it('widens the REQUEST so the server cannot withhold an all-day event from a sub-day window (#162)', async () => {
    // The measured bug. Cyrus matches a date-only value on its UTC day and <C:expand> emits
    // nothing for a window that merely touches that day, so "the morning of the 12th" on a
    // +10 account came back with no occurrence of the 12th's all-day event at all. No filter
    // can keep what the server never sent, which is why the margin is on the request.
    const allDay = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:allday@fm',
      'SUMMARY:Public holiday',
      'DTSTART;VALUE=DATE:20260812',
      'DTEND;VALUE=DATE:20260813',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client, mockDAVClient } = mockedClient([{ data: allDay, url: '/cal/allday.ics' }]);
    // 09:00 to 12:00 on the 12th, local: 11 Aug 23:00Z .. 12 Aug 02:00Z.
    const { events } = await client.getCalendarEvents(undefined, 50, '2026-08-12T09:00:00', '2026-08-12T12:00:00');

    const range = callArguments(mockDAVClient.fetchCalendarObjects)[0].timeRange!;
    assert.deepEqual(range, { start: '2026-08-11T09:00:00Z', end: '2026-08-12T16:00:00Z' });

    // And the RE-FILTER still uses the caller's own window, not the widened one: the all-day
    // event runs 11 Aug 14:00Z .. 12 Aug 14:00Z locally, which genuinely covers that morning.
    assert.equal(events.length, 2, JSON.stringify(events));
    assert.equal(events[0].id, 'allday@fm');
  });

  it('drops the neighbouring-day all-day row the old margin left behind (#162)', async () => {
    // The residue, measured: a Sydney account asking about the 12th was shown the all-day
    // event on the 11th, because the server matched it on its UTC day and fourteen hours of
    // client-side slack then reached back far enough to keep it. The exact filter resolves it
    // as the LOCAL day it is — 11 Aug 14:00Z .. 12 Aug 14:00Z — which ends exactly where the
    // caller's window begins, and a half-open interval does not overlap there.
    const dayBefore = [
      'BEGIN:VCALENDAR',
      'BEGIN:VEVENT',
      'UID:neighbour@fm',
      'SUMMARY:Yesterday all day',
      'DTSTART;VALUE=DATE:20260811',
      'DTEND;VALUE=DATE:20260812',
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');
    const { client } = mockedClient([{ data: dayBefore, url: '/cal/neighbour.ics' }]);
    const { events, total } = await client.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-12');
    assert.equal(total, 0, JSON.stringify(events));

    // The control: the same event on the day the caller DID ask about is kept, so this is the
    // window edge being read correctly rather than all-day events being dropped wholesale.
    const sameDay = dayBefore
      .replace('DTSTART;VALUE=DATE:20260811', 'DTSTART;VALUE=DATE:20260812')
      .replace('DTEND;VALUE=DATE:20260812', 'DTEND;VALUE=DATE:20260813');
    const { client: c2 } = mockedClient([{ data: sameDay, url: '/cal/today.ics' }]);
    const { total: t2 } = await c2.getCalendarEvents(undefined, 50, '2026-08-12', '2026-08-12');
    assert.equal(t2, 2);
  });

  it('saturates the WIDENED bounds without disclosing them as a clamp (#162)', async () => {
    // Fourteen hours past 9999-12-31 runs off the four-digit year, `toISOString` answers with
    // the expanded +010000-… form and tsdav throws a plain Error over it — an InternalError
    // raised for a caller-fixable argument. Both bounds here sit INSIDE the representable
    // range and only the margin added to them runs off the end, so this is the widening's own
    // saturation and nothing else's.
    const { client, mockDAVClient } = mockedClient();
    const { windowClamp } = await client.getCalendarEvents(
      undefined, 50, '0000-01-01T05:00:00Z', '9999-12-31T20:00:00Z',
    );

    const range = callArguments(mockDAVClient.fetchCalendarObjects)[0].timeRange!;
    assert.equal(range.start, '0000-01-01T00:00:00Z');
    assert.equal(range.end, '9999-12-31T23:59:59Z');
    assert.ok(!range.start.startsWith('-'), range.start);
    assert.ok(!range.end.startsWith('+'), range.end);

    // AND NOTHING IS DISCLOSED. The clamp array reports what happened to bounds the CALLER
    // named, and both of theirs were honoured exactly; the shortfall is in the margin, which
    // only ever reaches rows the exact filter drops anyway. Reporting it would tell the caller
    // their own window had been moved when it had not.
    assert.equal(windowClamp, undefined);
  });

  // End-to-end wiring check for #139: the configured zone must reach the parser as an injected
  // parameter and drive the omit-when-same rule for real, not just when `configuredZone` is
  // passed directly to `parseCalendarObject`. This test overrides the describe block's own
  // `Australia/Sydney` pin to `America/New_York` for its own duration: on the machine this was
  // written on, `Intl.DateTimeFormat().resolvedOptions().timeZone` (the host's own zone) is
  // ALSO `Australia/Sydney`, so a pin of `Australia/Sydney` here could not tell "read the
  // configured zone" apart from "silently fell back to the host zone" — a real production
  // regression to the host-zone fallback would leave this test green. `America/New_York` is
  // not this host's zone, so the omit-when-same fixture only omits if the configured value was
  // genuinely read.
  it('wires the pinned configured zone through to timeZone omit-when-same and emit-when-different', async () => {
    setDefaultTimezone('America/New_York');
    try {
      const sameZone = [
        'BEGIN:VCALENDAR',
        'BEGIN:VEVENT',
        'UID:same@fm',
        'DTSTART;TZID=America/New_York:20260320T083000',
        'DTEND;TZID=America/New_York:20260320T093000',
        'SUMMARY:Standup',
        'END:VEVENT',
        'END:VCALENDAR',
      ].join('\r\n');
      const differentZone = [
        'BEGIN:VCALENDAR',
        'BEGIN:VEVENT',
        'UID:different@fm',
        'DTSTART;TZID=Pacific/Auckland:20260320T083000',
        'DTEND;TZID=Pacific/Auckland:20260320T093000',
        'SUMMARY:Standup NZ',
        'END:VEVENT',
        'END:VCALENDAR',
      ].join('\r\n');
      const { client } = mockedClient([
        { data: sameZone, url: '/cal/same.ics' },
        { data: differentZone, url: '/cal/different.ics' },
      ]);
      // The window is named as instants and wide enough to hold both fixtures whichever zone
      // each is written in. It is here because there is no bounds-free listing any more
      // (#142) and the default window would put these 2026 fixtures out of range; the subject
      // is still which ZONE the parser was handed, not which days were searched.
      const { events } = await client.getCalendarEvents(
        undefined, 50, '2026-03-19T00:00:00Z', '2026-03-22T00:00:00Z',
      );
      const same = events.find(e => e.id === 'same@fm')!;
      const different = events.find(e => e.id === 'different@fm')!;
      assert.equal(same.timeZone, undefined);
      assert.equal(different.timeZone, 'Pacific/Auckland');
    } finally {
      // Restore the describe block's own pin — later tests in this block rely on it.
      setDefaultTimezone('Australia/Sydney');
    }
  });

  it('will not write into, or delete from, a calendar the read path refuses to read', async () => {
    // The read path selected from `selectableCalendars` and the write path from the raw list,
    // so an event could be created in — and later destroyed from — the hidden task collection
    // that `list_calendars` never shows and `list_calendar_events` answers "not found" for.
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
        { displayName: 'DEFAULT_TASK_CALENDAR_NAME', url: '/cal/tasks/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;

    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'DEFAULT_TASK_CALENDAR_NAME',
        title: 'T',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
      }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Calendar not found/);
        return true;
      },
    );
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 0);
    // And the lookup behind get/update/delete never reads that collection either.
    await assert.rejects(() => client.getCalendarEventById('anything@fm'), /Calendar event not found/);
    assert.equal(mockDAVClient.fetchCalendarObjects.mock.callCount(), 1, 'only the selectable calendar is read');
    assert.equal(callArguments(mockDAVClient.fetchCalendarObjects)[0].calendar.url, '/cal/personal/');
  });

  it('trims a calendarId on the write path, as the read path already did', async () => {
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [{ displayName: 'Work', url: '/cal/work/' }]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };
    (client as any).client = mockDAVClient;
    await client.createCalendarEvent({
      calendarId: '  Work  ',
      title: 'T',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });
    assert.equal(mockDAVClient.createCalendarObject.mock.calls.length, 1);
  });
});

describe('calendarNotFoundError lists only calendars a caller can name', () => {
  it('does not advertise the hidden task collection on the create path', async () => {
    // The read path filtered it out before calling the shared helper and the write path did
    // not, so a mistyped id on a create named a calendar list_calendars never shows.
    const client = new CalDAVCalendarClient({ username: 'me@fastmail.com', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
        { displayName: 'DEFAULT_TASK_CALENDAR_NAME', url: '/cal/tasks/' },
      ]),
      createCalendarObject: mock.fn(async (_params: CreateObjectParams) => ({ status: 200 })),
    };

    await assert.rejects(
      () => client.createCalendarEvent({
        calendarId: 'personal',
        title: 'T',
        start: '2026-04-07T14:00:00Z',
        end: '2026-04-07T15:00:00Z',
      }),
      (err: Error) => {
        assert.match(err.message, /"Personal"/);
        assert.ok(!err.message.includes('DEFAULT_TASK_CALENDAR_NAME'), err.message);
        return true;
      },
    );
  });
});

describe('unwrapDisplayName', () => {
  // The shapes below are the ones tsdav 2.3.1 can hand back for a calendar's `displayname`,
  // measured against xml-js under tsdav's own parser options rather than assumed. tsdav reads
  // the property as `props?.displayname?._cdata ?? props?.displayname` and passes it through
  // with no string guard, so its `displayName: string` typing is a claim, not a fact.

  it('returns a plain string name, trimmed', () => {
    assert.equal(unwrapDisplayName('Work'), 'Work');
    assert.equal(unwrapDisplayName('  Work  '), 'Work');
  });

  it('returns undefined for the empty-element shape', () => {
    // `<displayname/>` and `<displayname></displayname>` both parse to a bare `{}`: xml-js
    // compact mode, and tsdav's `textFn` only fires when the element HAS text. A
    // whitespace-only body is the same case, because tsdav parses with `trim: true`.
    assert.equal(unwrapDisplayName({}), undefined);
  });

  it('returns undefined for the attribute-typed shape', () => {
    // `<D:displayname xml:lang="en"/>` — no text, so the compact form keeps only its
    // attributes. tsdav's `attributesFn` strips a bare `xmlns` but not a prefixed one.
    assert.equal(unwrapDisplayName({ _attributes: { 'xml:lang': 'en' } }), undefined);
    assert.equal(unwrapDisplayName({ _attributes: { 'xmlns:d': 'DAV:' } }), undefined);
  });

  it('returns undefined for an absent property', () => {
    assert.equal(unwrapDisplayName(undefined), undefined);
    assert.equal(unwrapDisplayName(null), undefined);
  });

  it('drops a duplicated displayname rather than joining it into a name', () => {
    // Two `<displayname>` elements parse to an array. That collection is malformed and has no
    // name; the old code rendered "A,B" as though it did. The fallback is the degrade.
    assert.equal(unwrapDisplayName(['A', 'B']), undefined);
  });

  it('keeps a name tsdav typed away from string', () => {
    // tsdav's `nativeType` coerces element text that looks numeric or reads true/false, so a
    // calendar called "2026" arrives as the NUMBER 2026. It is still its name.
    assert.equal(unwrapDisplayName(2026), '2026');
    assert.equal(unwrapDisplayName(true), 'true');
    // Lossy BEFORE we see it: `<displayname>1e3</displayname>` is already the number 1000 by
    // the time tsdav hands it over, so "1000" is the only name that can be reported. What
    // survives is the invariant that matters — the reported name resolves as a calendarId.
    assert.equal(unwrapDisplayName(1000), '1000');
  });

  // DEFENSIVE ONLY — these two shapes cannot arrive from tsdav 2.3.1. It flattens `_cdata`
  // itself (`props?.displayname?._cdata ?? props?.displayname`) and its `textFn` replaces an
  // element with its text, so neither key survives to reach this function. They are covered
  // because the helper handles them if that read ever changes, NOT as evidence that a server
  // can produce them. Building a fixture on these believing them reachable is the mistake
  // that made the first version of this fix claim a bug that cannot happen.
  it('unwraps the CDATA and text-node shapes if tsdav ever stops flattening them', () => {
    assert.equal(unwrapDisplayName({ _attributes: { 'xmlns:d': 'DAV:' }, _cdata: 'Shared' }), 'Shared');
    assert.equal(unwrapDisplayName({ _text: 'Shared' }), 'Shared');
    // Empty text is a name that is not there, not a name of "".
    assert.equal(unwrapDisplayName({ _cdata: '   ' }), undefined);
  });
});

describe('calendar display names that are not strings', () => {
  it('never renders [object Object] as a calendar name, and lets the fallback fire', async () => {
    // The defect: `String({})` is "[object Object]" — a TRUTHY string, so the `|| 'Unnamed'`
    // fallback was dead on the exact input it existed for, and the marker reached the caller.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    const mockDAVClient = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
        // An empty <displayname/>, as parsed.
        { displayName: {}, url: '/cal/nameless/' },
        // An attribute-typed one.
        { displayName: { _attributes: { 'xml:lang': 'en' } }, url: '/cal/typed/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };
    (client as any).client = mockDAVClient;

    const calendars = await client.getCalendars();

    assert.equal(calendars.length, 3);
    for (const cal of calendars) {
      assert.ok(!cal.displayName.includes('[object Object]'), `leaked marker: ${cal.displayName}`);
    }
    assert.equal(calendars.find(c => c.id === '/cal/personal/')!.displayName, 'Personal');
    assert.equal(calendars.find(c => c.id === '/cal/nameless/')!.displayName, 'Unnamed');
    assert.equal(calendars.find(c => c.id === '/cal/typed/')!.displayName, 'Unnamed');
  });

  it('omits a nameless calendar\'s colour rather than emitting an object for it', async () => {
    // The identical defect one field over: tsdav guards `description` with a string check but
    // passes `calendarColor` through raw, so an empty `<calendar-color/>` parses to `{}` —
    // truthy, so the `|| undefined` beside it never fired and an object went out as a colour.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/', calendarColor: {} },
        { displayName: 'Work', url: '/cal/work/', calendarColor: '#aabbcc' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    const calendars = await client.getCalendars();

    assert.equal(calendars.find(c => c.id === '/cal/personal/')!.color, undefined);
    assert.equal(calendars.find(c => c.id === '/cal/work/')!.color, '#aabbcc');
  });

  it('still hides the task collection, which the unwrap must not disturb', async () => {
    // NOT a bug fix, and deliberately not written as one. There is no reachable input where
    // the raw comparison and the unwrapped one disagree here: the hidden name arrives as a
    // plain string, and every other shape tsdav can produce ({}, {_attributes}, a number, a
    // boolean, an array) fails to equal it either way. A padded name is not reachable either
    // — xml-js parses with `trim: true`, which trims plain text and CDATA alike. The unwrap
    // at this filter buys consistency with every other read of the field; this test is the
    // regression guard that it did not cost the behaviour that was already correct.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: '/cal/personal/' },
        { displayName: 'DEFAULT_TASK_CALENDAR_NAME', url: '/cal/tasks/' },
        // A nameless collection is NOT the hidden one and must survive the filter.
        { displayName: {}, url: '/cal/nameless/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    const calendars = await client.getCalendars();

    assert.deepEqual(calendars.map(c => c.id), ['/cal/personal/', '/cal/nameless/']);
  });

  it('matches a numeric calendarId against a calendar tsdav typed as a number', async () => {
    // Normalising only the STORED side broke this: `<displayname>2026</displayname>` arrives
    // as the number 2026, a caller passing the JSON number 2026 matched it by raw equality
    // before, and `"2026" === 2026` does not. Both sides go through the one normaliser now.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    const fetchCalendarObjects = mock.fn(async (_p: FetchObjectsParams) => []);
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 2026, url: '/cal/year/' },
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects,
    };

    const result = await client.getCalendarEvents(2026 as unknown as string, 50, '2026-04-01', '2026-04-30');

    assert.equal(result.events.length, 0);
    assert.equal(fetchCalendarObjects.mock.calls.length, 1);
    const [params] = callArguments(fetchCalendarObjects, 0);
    assert.equal(params.calendar.url, '/cal/year/');
    // And the name it reports back is one that resolves, which is the invariant the
    // stringifying half of the helper exists to keep.
    const listed = await client.getCalendars();
    assert.equal(listed.find(c => c.id === '/cal/year/')!.displayName, '2026');
  });

  it('finds a numerically-named calendar on the write path too', async () => {
    // The create/find path resolves `calendarId` by its own copy of the comparison, so it
    // needed the same normalisation. Diverging here would put an event in the wrong
    // collection, or refuse a calendar the read path accepts.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    const createCalendarObject = mock.fn(async (_p: CreateObjectParams) => ({ ok: true, status: 200 }));
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 2026, url: '/cal/year/' },
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
      createCalendarObject,
    };

    await client.createCalendarEvent({
      calendarId: 2026 as unknown as string,
      title: 'Planning',
      start: '2026-04-07T14:00:00Z',
      end: '2026-04-07T15:00:00Z',
    });

    assert.equal(createCalendarObject.mock.calls.length, 1);
    const [params] = callArguments(createCalendarObject, 0);
    assert.equal(params.calendar.url, '/cal/year/');
  });

  it('names a nameless calendar by its URL in the not-found error, WHOLE', async () => {
    // Two defects, and the second only shows on a realistic URL. The name list first dropped
    // anything unwrapping to undefined, so the message under-reported what the caller could
    // name. Listing the URL then echoed it through the default 60-character limit, which
    // truncates every real Fastmail collection URL — the prefix alone is 47 characters — and
    // a truncated URL pasted back earns the same error again with nothing saying it was cut.
    // The whole point of listing it is that it is a calendarId that RESOLVES, so the fixture
    // is a full-length URL rather than the short synthetic path that cannot reach the bug.
    const namelessUrl = 'https://caldav.fastmail.com/dav/calendars/user/user@example.invalid/a1b2c3d4e5f6/';
    assert.ok(namelessUrl.length > 60, 'fixture must exceed the default echo limit to test it');

    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: 'Personal', url: 'https://caldav.fastmail.com/dav/calendars/user/user@example.invalid/personal/' },
        { displayName: {}, url: namelessUrl },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    await assert.rejects(
      () => client.getCalendarEvents('Nope', 50, '2026-04-01', '2026-04-30'),
      (err: Error) => {
        assert.match(err.message, /"Personal"/);
        // Whole and unelided — `echoCallerText` marks a truncation with an ellipsis.
        assert.ok(err.message.includes(`"${namelessUrl}"`), err.message);
        assert.ok(!err.message.includes('…'), `URL was truncated: ${err.message}`);
        assert.ok(!err.message.includes('[object Object]'), err.message);
        return true;
      },
    );
  });

  it('bounds an offered URL at 200 and a name at 60, and echoes a rejected URL whole', async () => {
    // Pins the two limits against each other, because they are easy to collapse into one and
    // the whole point is that they differ. A URL is offered to be PASTED BACK, so it gets
    // room; a name is offered to be RECOGNISED, so it keeps the shared default. And the
    // REJECTED value gets the URL bound too — the caller who lands here is often one who
    // mistyped a URL, and 60 characters cuts inside the fixed prefix, where two different
    // wrong URLs render identically.
    const prefix = 'https://caldav.fastmail.com/dav/calendars/user/user@example.invalid/';
    const overlongUrl = `${prefix}${'a'.repeat(260 - prefix.length)}/`;
    const overlongName = `Team ${'x'.repeat(80)}`;
    const rejectedUrl = `${prefix}typo-but-realistic-collection-id/`;
    assert.ok(overlongUrl.length > 200, 'URL fixture must exceed the URL bound');
    assert.ok(overlongName.length > 60, 'name fixture must exceed the default bound');
    assert.ok(rejectedUrl.length > 60, 'rejected fixture must exceed the default bound');

    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: overlongName, url: '/cal/long-name/' },
        { displayName: {}, url: overlongUrl },
      ]),
      fetchCalendarObjects: mock.fn(async (_p: FetchObjectsParams) => []),
    };

    await assert.rejects(
      () => client.getCalendarEvents(rejectedUrl, 50, '2026-04-01', '2026-04-30'),
      (err: Error) => {
        // The rejected value survives intact, so the caller can see WHICH URL missed.
        assert.ok(err.message.includes(`Calendar not found: "${rejectedUrl}"`), err.message);
        // An offered URL past the bound is cut AT the bound, and says so with the ellipsis.
        assert.ok(err.message.includes(`"${overlongUrl.slice(0, 200)}…"`), err.message);
        assert.ok(!err.message.includes(overlongUrl), 'a 260-char URL must not appear whole');
        // A name past 60 keeps the shared default rather than inheriting the URL bound.
        assert.ok(err.message.includes(`"${overlongName.slice(0, 60)}…"`), err.message);
        return true;
      },
    );
  });

  it('matches a calendarId of "true" against a calendar tsdav typed as a boolean', async () => {
    // The number is not the only coercion `nativeType` performs: `<displayname>true</displayname>`
    // arrives as the BOOLEAN true. Same regression, same fix, and worth its own case because
    // the boolean branch of the helper is otherwise only unit-tested.
    const client = new CalDAVCalendarClient({ username: 'me@example.invalid', password: 'test' });
    const fetchCalendarObjects = mock.fn(async (_p: FetchObjectsParams) => []);
    (client as any).client = {
      login: mock.fn(async () => {}),
      fetchCalendars: mock.fn(async () => [
        { displayName: true, url: '/cal/boolish/' },
        { displayName: 'Personal', url: '/cal/personal/' },
      ]),
      fetchCalendarObjects,
    };

    const result = await client.getCalendarEvents('true', 50, '2026-04-01', '2026-04-30');

    assert.equal(result.events.length, 0);
    assert.equal(fetchCalendarObjects.mock.calls.length, 1);
    const [params] = callArguments(fetchCalendarObjects, 0);
    assert.equal(params.calendar.url, '/cal/boolish/');
    // And it is listed under the name that resolves it.
    const listed = await client.getCalendars();
    assert.equal(listed.find(c => c.id === '/cal/boolish/')!.displayName, 'true');
  });
});
