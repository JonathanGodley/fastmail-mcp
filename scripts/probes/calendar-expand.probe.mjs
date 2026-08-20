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
// Run: python scripts/probes/run-probe.py calendar-expand.probe.mjs [startDate endDate]

import { DAVClient } from 'tsdav';

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

let failures = 0;
const check = (ok, label, detail) => {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` — ${detail}` : ''}`);
  if (!ok) failures++;
};

await client.login();
const calendars = await client.fetchCalendars();
console.log(`\nWindow: ${START} .. ${END}`);
console.log(`Calendars discovered: ${calendars.length}\n`);
check(calendars.length > 0, 'calendar discovery returned at least one calendar');

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
    check(!anyRrule, `[${name}] expanded occurrences carry no RRULE (server applied the recurrence)`);
    check(allIn, `[${name}] every expanded occurrence falls inside the requested window`);
    if (parts.length > 1) {
      console.log(`  NOTE: ${parts.length} VEVENTs arrived in ONE blob — a first-match parser would drop ${parts.length - 1}.`);
    }
  }
}

console.log(`\n${failures === 0 ? 'ALL CHECKS PASSED' : `${failures} CHECK(S) FAILED`}`);
process.exit(failures === 0 ? 0 : 1);
