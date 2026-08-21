// What this probe settles
// -----------------------
// `list_calendar_events` reports a recurring event at its ORIGINAL DTSTART and returns
// events outside the requested window (#64). The diagnosis is that the read path never
// asks the server to expand recurrences, so there is no in-window instance to report or
// to filter on. The fix depends on one platform fact that cannot be read off the RFC:
//
//   what does Fastmail's CalDAV server actually return when tsdav is passed
//   `expand: true` alongside a `timeRange`?
//
// RFC 4791 says a <C:expand> request returns each in-window occurrence as its own VEVENT
// with the recurrence applied and RRULE stripped. What matters for our code is the SHAPE
// it arrives in: whether those occurrences come back as several VEVENTs inside ONE
// calendar-data blob per resource (which `parseCalendarObject` would silently reduce to
// the first, changing the symptom without fixing the bug) or as separate objects.
//
// This probe answers that by querying the same window both ways and printing the
// structure of what comes back. It is READ-ONLY: it creates nothing. That is deliberate
// beyond ordinary tidiness — creating an event with participants makes the server send
// real iTIP invitations (see the README), so a calendar probe that writes is not a safe
// default.
//
// SECOND PASS: does the FIRST instance of a series carry a RECURRENCE-ID?
// -----------------------------------------------------------------------
// The pass above queries one fixed window and says nothing about a window that contains a
// series' ORIGINAL DTSTART — and that turned out to be the case that matters. Cyrus's
// `expand_cb` sets a RECURRENCE-ID only on instances after the first, so such a window
// returns [first-instance-with-no-RECURRENCE-ID, occurrence, occurrence, …]. Any parser
// that identifies "the block without a RECURRENCE-ID" as a series master reads block 0 as
// one and drops every sibling; a five-year window over a yearly series reported ONE event.
// The default window here contains no series start, which is exactly why the original pass
// never saw it, so the second pass finds a real recurring series in the account, builds a
// window from its own DTSTART, and asserts the shape directly.
//
// Run: python scripts/probes/run-probe.py calendar-expand.probe.mjs [startDate endDate]

import { DAVClient } from 'tsdav';
// The shared PASS/FAIL harness. This probe talks to tsdav rather than to the MCP server, so
// it uses nothing else from probelib — but it takes its `check` from there so the two
// calendar probes cannot disagree about the argument order, which they did.
import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-expand.probe.mjs');
  process.exit(1);
}

// A window far enough out that anything landing in it is there because a recurrence put
// it there, not because a one-off event happens to sit in the range. Overridable so the
// probe can be re-pointed at a window whose contents are known.
const START = process.argv[2] || '2027-03-01T00:00:00Z';
const END = process.argv[3] || '2027-03-10T00:00:00Z';

const VEVENT_RE = /BEGIN:VEVENT[\s\S]*?END:VEVENT/g;

function icalValue(block, key) {
  const m = block.match(new RegExp(`^${key}[;:].*$`, 'm'));
  return m ? m[0].replace(/\r$/, '') : undefined;
}

/** Structural summary of one calendar-data blob: how many VEVENTs, and what each carries. */
function describeBlob(data) {
  const blocks = data.match(VEVENT_RE) || [];
  return blocks.map(b => ({
    dtstart: icalValue(b, 'DTSTART'),
    rrule: icalValue(b, 'RRULE'),
    recurrenceId: icalValue(b, 'RECURRENCE-ID'),
  }));
}

function inWindow(dtstartLine, start, end) {
  if (!dtstartLine) return null;
  const raw = dtstartLine.slice(dtstartLine.lastIndexOf(':') + 1).trim();
  const iso = raw.length === 8
    ? `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}T00:00:00Z`
    : raw.replace(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})/, '$1-$2-$3T$4:$5:$6');
  const t = Date.parse(iso);
  if (Number.isNaN(t)) return null;
  return t >= Date.parse(start) && t < Date.parse(end);
}

const client = new DAVClient({
  serverUrl: 'https://caldav.fastmail.com/dav/',
  credentials: { username: USERNAME, password: PASSWORD },
  authMethod: 'Basic',
  defaultAccountType: 'caldav',
});

const { check, failures } = makeChecker();

// Series masters seen in the plain (unexpanded) pass, used to build the first-instance
// window below: a master carries the RRULE and the series' ORIGINAL DTSTART, which is
// precisely the date the expanded pass needs to include.
const seriesCandidates = [];

await client.login();
const calendars = await client.fetchCalendars();
console.log(`\nWindow: ${START} .. ${END}`);
console.log(`Calendars discovered: ${calendars.length}\n`);
check('calendar discovery returned at least one calendar', calendars.length > 0);

for (const cal of calendars) {
  const name = cal.displayName || cal.url;
  console.log(`\n=== ${name} ===`);

  const plain = await client.fetchCalendarObjects({
    calendar: cal,
    timeRange: { start: START, end: END },
  });
  const expanded = await client.fetchCalendarObjects({
    calendar: cal,
    timeRange: { start: START, end: END },
    expand: true,
  });

  console.log(`  objects without expand: ${plain.length}`);
  console.log(`  objects with expand:    ${expanded.length}`);

  for (const obj of plain) {
    const parts = describeBlob(obj.data || '');
    if (parts.length === 0) continue;
    const first = parts[0];
    const win = inWindow(first.dtstart, START, END);
    console.log(`  [plain]    veventsInBlob=${parts.length} firstDTSTART=${first.dtstart} rrule=${first.rrule ? 'yes' : 'no'} firstInWindow=${win}`);
  }

  for (const obj of expanded) {
    const parts = describeBlob(obj.data || '');
    if (parts.length === 0) continue;
    const allIn = parts.every(p => inWindow(p.dtstart, START, END) !== false);
    const anyRrule = parts.some(p => p.rrule);
    console.log(`  [expanded] veventsInBlob=${parts.length} allInWindow=${allIn} anyRRULE=${anyRrule ? 'yes' : 'no'}`);
    for (const p of parts) {
      console.log(`               DTSTART=${p.dtstart}${p.recurrenceId ? ` RECURRENCE-ID=${p.recurrenceId}` : ''}`);
    }

    // The two facts the fix depends on.
    check(`[${name}] expanded occurrences carry no RRULE (server applied the recurrence)`, !anyRrule);
    check(`[${name}] every expanded occurrence falls inside the requested window`, allIn);
    if (parts.length > 1) {
      console.log(`  NOTE: ${parts.length} VEVENTs arrived in ONE blob — a first-match parser would drop ${parts.length - 1}.`);
    }
  }

  // Remember any series master seen here, for the first-instance pass below.
  for (const obj of plain) {
    for (const b of describeBlob(obj.data || '')) {
      if (b.rrule && b.dtstart && !b.recurrenceId) seriesCandidates.push({ calendar: cal, name, ...b });
    }
  }
}

// ---------------------------------------------------------------------------
// Second pass: a window that CONTAINS the series' first instance.
// ---------------------------------------------------------------------------

/** The ISO instant a DTSTART line names, read as UTC (TZID names are not resolved here). */
function dtstartInstant(dtstartLine) {
  const raw = dtstartLine.slice(dtstartLine.lastIndexOf(':') + 1).trim();
  const iso = raw.length === 8
    ? `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}T00:00:00Z`
    : `${raw.replace(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2}).*$/, '$1-$2-$3T$4:$5:$6')}Z`;
  const t = Date.parse(iso);
  return Number.isNaN(t) ? null : t;
}

// How far past the series' start to look, per FREQ, so the window holds a handful of
// occurrences and no more. Sub-daily frequencies are SKIPPED rather than sized: asking the
// server to expand a FREQ=MINUTELY series over any useful span is the denial-of-service the
// window clamp exists to prevent, and this probe must stay safe to run unattended.
const SPAN_DAYS_BY_FREQ = { YEARLY: 5 * 366, MONTHLY: 5 * 31, WEEKLY: 5 * 7, DAILY: 5 };

function spanDaysFor(rruleLine) {
  const freq = /FREQ=([A-Z]+)/.exec(rruleLine || '')?.[1];
  return SPAN_DAYS_BY_FREQ[freq] ?? null;
}

console.log('\n\n=== First-instance shape: a window containing a series DTSTART ===');

const usable = seriesCandidates
  .map(c => ({ ...c, spanDays: spanDaysFor(c.rrule), startMs: dtstartInstant(c.dtstart) }))
  .filter(c => c.spanDays !== null && c.startMs !== null);

if (usable.length === 0) {
  // Not a failure: the account may hold no recurring series in the queried window, or only
  // sub-daily ones. Say so plainly rather than passing a check that never ran.
  console.log('  SKIP  no daily-or-coarser recurring series was found in the window above,');
  console.log('        so there was no series start to build a window from. Re-run with a');
  console.log('        window that contains one, e.g. calendar-expand.probe.mjs 2020-01-01 2021-01-01.');
} else {
  const probe = usable[0];
  const winStart = new Date(probe.startMs).toISOString().replace(/\.\d{3}Z$/, 'Z');
  const winEnd = new Date(probe.startMs + probe.spanDays * 86400000).toISOString().replace(/\.\d{3}Z$/, 'Z');
  console.log(`  series:  ${probe.dtstart}  ${probe.rrule}`);
  console.log(`  window:  ${winStart} .. ${winEnd}  (from the series' own DTSTART)\n`);

  const expanded = await client.fetchCalendarObjects({
    calendar: probe.calendar,
    timeRange: { start: winStart, end: winEnd },
    expand: true,
  });

  // Match the resource by its DTSTART: the same instant the master reported, which after
  // expansion is the first instance's DTSTART.
  let matched = null;
  for (const obj of expanded) {
    const parts = describeBlob(obj.data || '');
    if (parts.some(p => p.dtstart && dtstartInstant(p.dtstart) === probe.startMs)) matched = parts;
  }

  if (!matched) {
    check('the expanded query returned the series whose start the window was built from', false);
  } else {
    for (const p of matched) {
      console.log(`    DTSTART=${p.dtstart}${p.recurrenceId ? ` RECURRENCE-ID=${p.recurrenceId}` : '  (no RECURRENCE-ID)'}`);
    }
    const bare = matched.filter(p => !p.recurrenceId);
    check('the window holds more than one occurrence of the series', matched.length > 1, `veventsInBlob=${matched.length}`);
    // THE FACT THIS PASS EXISTS FOR. A parser that treats "the block with no RECURRENCE-ID"
    // as a series master finds exactly one here and discards every sibling.
    check(
      'exactly ONE expanded block carries no RECURRENCE-ID — the series FIRST instance',
      bare.length === 1,
      `bare=${bare.length} of ${matched.length}`,
    );
    check(
      "that bare block is the series' original DTSTART, not an unexpanded master",
      bare.length === 1 && dtstartInstant(bare[0].dtstart) === probe.startMs,
    );
    check('no expanded block carries an RRULE, including the first instance', !matched.some(p => p.rrule));
    console.log(`\n  A "block without a RECURRENCE-ID is the master" parser would report ${matched.length} occurrences as 1.`);
  }
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
