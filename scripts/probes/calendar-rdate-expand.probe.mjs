// What this probe settles
// -----------------------
// `docs/conventions.md` and the #162 close-out both record that Fastmail's CalDAV server
// (Cyrus) strips `RDATE` from an `<C:expand>`-ed block the way `calendar-expand.probe.mjs`
// MEASURED that it strips `RRULE`. The RRULE half is observed; the RDATE half was read off the
// Cyrus source and never observed — nothing in `scripts/probes/` looked at `RDATE` under
// `<C:expand>` at all. This probe is that observation (#165).
//
// It is load-bearing in two places:
//
//   - the recurrence guard in `eventIntersectsWindow`'s caller (`getCalendarEvents`) keeps any
//     block still carrying `RRULE` *or* `RDATE`, on the grounds that such a block is an
//     unexpanded master whose original DTSTART must not be judged against the window. If the
//     server does NOT strip `RDATE` under expand, that guard fires on every expanded
//     occurrence of an RDATE-listed series and the window filter stops filtering them at all;
//   - `update_calendar_event` / `delete_calendar_event` refuse a series that recurs only by
//     `RDATE` (`isRecurringSeriesResource`). That refusal reads the UNEXPANDED master, so it is
//     unaffected by the strip either way — but the read path's `recurrenceDates` field is
//     documented as "expected to be absent on the ordinary listing path" on the strength of
//     the same unmeasured claim.
//
// WHAT IS MEASURED, and how to read a result. Four queries against a temporary collection
// holding three synthetic series:
//
//   Q1  wide window, NO expand   — the stored form. Baseline: RDATE and RRULE arrive intact.
//   Q2  wide window, WITH expand — the shipped shape. Per resource: how many VEVENT blocks are
//                                  in the blob, whether any carries an RDATE line, whether any
//                                  carries an RRULE line, which carry a RECURRENCE-ID and at
//                                  what value, and whether every RDATE occurrence was emitted.
//   Q3  narrow window over ONE RDATE occurrence, WITH expand — is that single occurrence
//                                  emitted, and in what RECURRENCE-ID form (TZID or Z)?
//   Q4  the same narrow window, NO expand — does the server's time-range FILTER evaluate RDATE
//                                  occurrences at all, or only DTSTART? (The frames probe
//                                  showed the filter and the expansion can disagree, so the two
//                                  halves are asked separately.)
//
// TWO RESULTS, and the second was not what the repo had recorded.
//
//   THE STRIP IS CONFIRMED (Q2). An `<C:expand>`-ed RDATE-only series comes back as one VEVENT
//   per occurrence with no RDATE line anywhere, RECURRENCE-ID set on every block after the
//   first and absent on the first — the same shape `calendar-expand.probe.mjs` measured for
//   RRULE. `docs/conventions.md` no longer calls that half derived.
//
//   THE TIME-RANGE FILTER DOES NOT WALK RDATEs (Q3/Q4, and the diagnostic windows). A window
//   that covers an RDATE occurrence but NOT the series DTSTART matches the resource not at all,
//   with or without expand — while the RRULE control's occurrence at the IDENTICAL instant is
//   returned by the very same request. The two diagnostic windows say which span the filter
//   does index: `narrowDtstart` (covers DTSTART only) matches, `betweenOccurrences` (covers no
//   occurrence at all, but lies between DTSTART and the last RDATE) does not. So the indexed
//   span for an RDATE-only resource is DTSTART..DTSTART+DURATION, and the RDATEs are invisible
//   to the filter. The repo had assumed the filter walks every occurrence.
//
//   The Cyrus source says why, and the two halves genuinely use different walkers:
//   `<C:expand>` (http_caldav.c, `expand_cb`) runs Cyrus's own `icalcomponent_myforeach`
//   (ical_support.c), which adds RDATEs explicitly; the time-range filter
//   (http_caldav.c, `apply_comp_timerange`) and the indexed span (caldav_db.c →
//   ical_support.c, `icalrecurrenceset_get_utc_timespan`) both use libical's
//   `icalcomponent_foreach_recurrence`, which in Fastmail's build does not reach them.
//
// BOTH RDATE SERIALISATIONS ARE MEASURED, because "two RDATE lines" and "one comma-joined
// RDATE line" are the same property set to a parser but not necessarily to an indexer: the
// fixtures carry one of each and every check below runs per form, so a form-specific result
// reports as a divergence between the two rather than as a single ambiguous number.
//
// Every check below asserts the platform behaviour this probe MEASURED, the way
// `calendar-window-frames.probe.mjs` asserts the UTC-day fact. PASS means the platform still
// behaves as recorded here; FAIL means it has CHANGED, at which point the guard design resting
// on these facts needs re-deciding rather than the probe re-tuned.
//
// WHY RAW CALDAV. This measures the platform, not our parsing, so it talks HTTP directly and
// builds the same XML tsdav builds for `fetchCalendarObjects` with `timeRange` + `expand` — a
// calendar-query REPORT at Depth 1 asking for d:getetag and a c:calendar-data carrying
// c:expand, filtered VCALENDAR > VEVENT > c:time-range, both ranges written as the same UTC
// basic-format instants. Raw PUT is likewise not a shortcut: `create_calendar_event` has no
// RDATE parameter, so this server cannot author the RDATE-only fixture at all.
//
// FIXTURES AND CLEANUP. Three synthetic series — an RDATE-only one written as a single
// comma-joined RDATE line, the same series written as one RDATE property per line, and an RRULE
// control that reproduces the known RRULE measurement in the same run — are written into a
// temporary calendar created by MKCALENDAR, out in July 2027 so nothing real shares the window
// (the frames probe uses June 2027; these deliberately do not collide). None carries an ATTENDEE
// or an ORGANIZER: a participant would make the server send real iTIP mail (see the README).
// If MKCALENDAR fails this probe STOPS — it does not fall back to an existing calendar, because
// the fixtures are writes into a live personal account. The finally block deletes the whole
// collection, which removes every fixture in one request, then PROPFINDs to confirm it is gone.
// Each run mints its own collection name, so a re-run never needs a manual delete first.
//
// Run: python scripts/probes/run-probe.py calendar-rdate-expand.probe.mjs

import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-rdate-expand.probe.mjs');
  process.exit(1);
}

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');

// Every CalDAV href on this server embeds the account's own address, and this probe prints
// whole calendar-data blobs. Nothing printed is allowed to carry a real address: probe output
// gets pasted into issues. The second rule catches the fixtures' own `@probe.invalid` UIDs too,
// which is a legibility cost worth paying — each blob is printed under a label that says which
// fixture it is.
const redact = s => String(s)
  .split(USERNAME).join('<account>')
  .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');

// The zone the fixtures are written in and the zone this deployment runs in. Hard-coded rather
// than read from the environment because the fixtures' VTIMEZONE has to match it.
const ZONE = 'Australia/Sydney';

// ---------------------------------------------------------------------------
// Wall clock -> instant, in ZONE. Two passes: sample the offset at the naive guess, then
// re-check at the instant that lands on. Enough for dates nowhere near a transition.
// ---------------------------------------------------------------------------
function offsetMsAt(zone, utcMs) {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: zone, hour12: false,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(new Date(utcMs)).map(p => [p.type, p.value]),
  );
  return Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour % 24, +parts.minute, +parts.second) - utcMs;
}

function wallToUtcMs(y, mo, d, h, mi = 0, s = 0) {
  const guess = Date.UTC(y, mo - 1, d, h, mi, s);
  const once = guess - offsetMsAt(ZONE, guess);
  return guess - offsetMsAt(ZONE, once);
}

const wallToUtcIso = (...a) => new Date(wallToUtcMs(...a)).toISOString().replace(/\.\d{3}Z$/, 'Z');

// tsdav's spelling of a time-range bound: UTC basic format, seconds precision.
const davInstant = iso => `${new Date(iso).toISOString().slice(0, 19).replace(/[-:.]/g, '')}Z`;

// ---------------------------------------------------------------------------
// Minimal HTTP + XML plumbing. Regex parsing is enough for a probe against one known server;
// every matcher tolerates any namespace prefix because Cyrus picks its own.
// ---------------------------------------------------------------------------
async function dav(method, url, { body, headers = {} } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: AUTH,
      ...(body !== undefined ? { 'Content-Type': headers['Content-Type'] ?? 'text/xml;charset=UTF-8' } : {}),
      ...headers,
    },
    body,
  });
  return { status: res.status, statusText: res.statusText, text: await res.text() };
}

const el = (xml, name) => {
  const m = xml.match(new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`));
  return m ? m[1] : undefined;
};
const elAll = (xml, name) =>
  [...xml.matchAll(new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`, 'g'))].map(m => m[1]);
const responses = xml => elAll(xml, 'response');
const abs = href => new URL(href, ROOT).href;

const unescapeXml = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');

const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

// The exact request tsdav sends for fetchCalendarObjects({ timeRange, expand }); `expand` false
// drops only the c:expand child, which is the one difference between the two calls.
const calendarQueryXml = (startIso, endIso, expand) =>
  `<?xml version="1.0" encoding="utf-8"?>\n` +
  `<c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:cs="http://calendarserver.org/ns/" ` +
  `xmlns:ca="http://apple.com/ns/ical/" xmlns:d="DAV:">` +
  `<d:prop><d:getetag/>` +
  (expand
    ? `<c:calendar-data><c:expand start="${davInstant(startIso)}" end="${davInstant(endIso)}"/></c:calendar-data>`
    : `<c:calendar-data/>`) +
  `</d:prop>` +
  `<c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT">` +
  `<c:time-range start="${davInstant(startIso)}" end="${davInstant(endIso)}"/>` +
  `</c:comp-filter></c:comp-filter></c:filter></c:calendar-query>`;

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const STAMP = Date.now();
const uid = kind => `probe-165-${kind}-${STAMP}@probe.invalid`;

const ics = lines => lines.join('\r\n') + '\r\n';

const VTIMEZONE_SYDNEY = [
  'BEGIN:VTIMEZONE',
  'TZID:Australia/Sydney',
  'BEGIN:STANDARD',
  'DTSTART:19700405T030000',
  'RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU',
  'TZOFFSETFROM:+1100',
  'TZOFFSETTO:+1000',
  'TZNAME:AEST',
  'END:STANDARD',
  'BEGIN:DAYLIGHT',
  'DTSTART:19701004T020000',
  'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU',
  'TZOFFSETFROM:+1000',
  'TZOFFSETTO:+1100',
  'TZNAME:AEDT',
  'END:DAYLIGHT',
  'END:VTIMEZONE',
];

const DTSTAMP = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '') + 'Z';

// July/August 2027: Sydney is on standard time (+10:00) throughout, with no DST transition
// anywhere near, so every instant below is DTSTART-local-hour minus ten and checkable by eye.
const SERIES_START = { y: 2027, mo: 7, d: 14, h: 9 };            // Wed 14 Jul 2027, 09:00 local
const RDATE_OCCURRENCES = [                                       // in the order the RDATEs list
  { y: 2027, mo: 7, d: 14, h: 9, label: '14 Jul (the DTSTART itself)' },
  { y: 2027, mo: 7, d: 21, h: 9, label: '21 Jul (first RDATE)' },
  { y: 2027, mo: 8, d: 4, h: 9, label: '4 Aug (second RDATE)' },
];
const RRULE_OCCURRENCES = [                                       // FREQ=WEEKLY;COUNT=3
  { y: 2027, mo: 7, d: 14, h: 9, label: '14 Jul' },
  { y: 2027, mo: 7, d: 21, h: 9, label: '21 Jul' },
  { y: 2027, mo: 7, d: 28, h: 9, label: '28 Jul' },
];

const FIXTURES = [
  {
    kind: 'rdate',
    label: 'RDATE-only series (no RRULE; three occurrences listed)',
    uid: uid('rdate'),
    occurrences: RDATE_OCCURRENCES,
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-165//EN', 'CALSCALE:GREGORIAN',
      ...VTIMEZONE_SYDNEY,
      'BEGIN:VEVENT', `UID:${uid('rdate')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART;TZID=${ZONE}:20270714T090000`,
      'DURATION:PT1H',
      `RDATE;TZID=${ZONE}:20270721T090000,20270804T090000`,
      'SUMMARY:probe-165 RDATE-only series', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
  {
    kind: 'rdateLines',
    label: 'RDATE-only series, one RDATE property per line (same three occurrences)',
    uid: uid('rdatelines'),
    occurrences: RDATE_OCCURRENCES,
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-165//EN', 'CALSCALE:GREGORIAN',
      ...VTIMEZONE_SYDNEY,
      'BEGIN:VEVENT', `UID:${uid('rdatelines')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART;TZID=${ZONE}:20270714T090000`,
      'DURATION:PT1H',
      // The one difference from the fixture above: RFC 5545 §3.8.5.2 allows either spelling,
      // and a parser sees the same property set either way — but an INDEXER need not.
      `RDATE;TZID=${ZONE}:20270721T090000`,
      `RDATE;TZID=${ZONE}:20270804T090000`,
      'SUMMARY:probe-165 RDATE-only series, one per line', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
  {
    kind: 'rrule',
    label: 'RRULE control (FREQ=WEEKLY;COUNT=3)',
    uid: uid('rrule'),
    occurrences: RRULE_OCCURRENCES,
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-165//EN', 'CALSCALE:GREGORIAN',
      ...VTIMEZONE_SYDNEY,
      'BEGIN:VEVENT', `UID:${uid('rrule')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART;TZID=${ZONE}:20270714T090000`,
      'DURATION:PT1H',
      'RRULE:FREQ=WEEKLY;COUNT=3',
      'SUMMARY:probe-165 RRULE control series', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
];

// The two RDATE serialisations, which every check runs over so a form-specific result reports
// as a divergence between them rather than as one ambiguous number.
const RDATE_KINDS = ['rdate', 'rdateLines'];

// ---------------------------------------------------------------------------
// Windows. Each names the question it stands for.
// ---------------------------------------------------------------------------
const WINDOWS = {
  wide: {
    label: `wide: local 2027-07-01 .. 2027-08-31 in ${ZONE} (covers every occurrence of both series)`,
    start: wallToUtcIso(2027, 7, 1, 0),
    end: wallToUtcIso(2027, 8, 31, 0),
  },
  narrowSecondRdate: {
    label: `narrow: local 2027-07-21 08:00..10:00 in ${ZONE} (only the FIRST RDATE occurrence)`,
    start: wallToUtcIso(2027, 7, 21, 8),
    end: wallToUtcIso(2027, 7, 21, 10),
  },
  // Two windows that isolate WHAT SPAN the time-range filter indexes for an RDATE-only
  // resource, once the narrow window above turns out not to match it.
  narrowDtstart: {
    label: `narrow: local 2027-07-14 08:00..10:00 in ${ZONE} (only the DTSTART occurrence)`,
    start: wallToUtcIso(2027, 7, 14, 8),
    end: wallToUtcIso(2027, 7, 14, 10),
  },
  betweenOccurrences: {
    label: `narrow: local 2027-07-17 08:00..10:00 in ${ZONE} (no occurrence, but inside DTSTART..last RDATE)`,
    start: wallToUtcIso(2027, 7, 17, 8),
    end: wallToUtcIso(2027, 7, 17, 10),
  },
};

// ---------------------------------------------------------------------------
// iCalendar reading. Whole content lines only, the same discipline src/caldav-client.ts uses:
// U+2028/U+2029/bare CR are inert characters inside a value, not line breaks.
// ---------------------------------------------------------------------------
const contentLines = blob => blob.split(/\r\n|\n/).map(l => l.replace(/\r$/, ''));

/** Split a VCALENDAR blob into its VEVENT blocks (arrays of content lines). */
function veventBlocks(blob) {
  const out = [];
  let cur = null;
  for (const line of contentLines(blob)) {
    if (line === 'BEGIN:VEVENT') { cur = []; continue; }
    if (line === 'END:VEVENT') { if (cur) out.push(cur); cur = null; continue; }
    if (cur) cur.push(line);
  }
  return out;
}

const propLines = (block, key) => block.filter(l => new RegExp(`^${key}[;:]`).test(l));
const propValue = (block, key) => {
  const line = propLines(block, key)[0];
  return line === undefined ? undefined : line.slice(line.indexOf(':') + 1);
};
const propParams = (block, key) => {
  const line = propLines(block, key)[0];
  if (line === undefined) return undefined;
  return line.slice(key.length, line.indexOf(':'));
};

/**
 * A date-time content line -> UTC ms, so occurrences can be compared whatever form the server
 * hands them back in. Three forms are possible and the frames probe measured which one Cyrus
 * actually emits after expansion (a bare Z instant); the other two are read anyway so a change
 * in the platform reports as a mismatch rather than as an unparseable line.
 */
function lineToUtcMs(line) {
  if (!line) return undefined;
  const value = line.slice(line.indexOf(':') + 1);
  const params = line.slice(0, line.indexOf(':'));
  const m = value.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (!m) return undefined;
  const [, y, mo, d, h, mi, s, z] = m;
  if (z === 'Z') return Date.UTC(+y, +mo - 1, +d, +h, +mi, +s);
  const tzid = params.match(/TZID=([^;:]+)/)?.[1];
  // A TZID we authored, or a floating value the platform resolves as UTC (frames probe).
  if (tzid === ZONE) return wallToUtcMs(+y, +mo, +d, +h, +mi, +s);
  if (!tzid) return Date.UTC(+y, +mo - 1, +d, +h, +mi, +s);
  return undefined;
}

const occurrenceMs = o => wallToUtcMs(o.y, o.mo, o.d, o.h);
const asIso = ms => (ms === undefined ? '(unparsed)' : new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z'));

/** Everything one expanded blob says about a resource, in the terms the checks below need. */
function describe(blob) {
  const blocks = veventBlocks(blob);
  return {
    blockCount: blocks.length,
    anyRdate: blocks.some(b => propLines(b, 'RDATE').length > 0),
    anyRrule: blocks.some(b => propLines(b, 'RRULE').length > 0),
    rdateLines: blocks.flatMap(b => propLines(b, 'RDATE')),
    rruleLines: blocks.flatMap(b => propLines(b, 'RRULE')),
    blocks: blocks.map(b => ({
      dtstartLine: propLines(b, 'DTSTART')[0],
      dtstartMs: lineToUtcMs(propLines(b, 'DTSTART')[0]),
      recurrenceIdLine: propLines(b, 'RECURRENCE-ID')[0],
      recurrenceIdMs: lineToUtcMs(propLines(b, 'RECURRENCE-ID')[0]),
      recurrenceIdParams: propParams(b, 'RECURRENCE-ID'),
      duration: propValue(b, 'DURATION'),
    })),
  };
}

// ---------------------------------------------------------------------------
const { check, failures } = makeChecker();

let tempCalendarUrl = null;   // set only once MKCALENDAR has succeeded

async function discover() {
  const rootPf = await dav('PROPFIND', ROOT, { body: PROPFIND(['d:current-user-principal']), headers: { Depth: '0' } });
  const principal = el(el(rootPf.text, 'current-user-principal') ?? '', 'href');
  if (!principal) throw new Error(`could not read current-user-principal (HTTP ${rootPf.status})`);

  const prinPf = await dav('PROPFIND', abs(principal.trim()), {
    body: PROPFIND(['c:calendar-home-set']), headers: { Depth: '0' },
  });
  const home = el(el(prinPf.text, 'calendar-home-set') ?? '', 'href');
  if (!home) throw new Error(`could not read calendar-home-set (HTTP ${prinPf.status})`);
  return abs(home.trim());
}

/** Run one calendar-query. Returns the calendar-data per fixture kind, plus which were matched. */
async function query(calendarUrl, win, expand) {
  const res = await dav('REPORT', calendarUrl, {
    body: calendarQueryXml(win.start, win.end, expand),
    headers: { Depth: '1' },
  });
  const seen = new Map();   // fixture kind -> the calendar-data blob carrying its UID
  const hrefs = new Set();  // fixture kind -> the resource was named in the multistatus at all
  for (const block of responses(res.text)) {
    const href = (el(block, 'href') ?? '').trim();
    const data = unescapeXml(el(block, 'calendar-data') ?? '');
    for (const f of FIXTURES) {
      if (f.url && href && decodeURIComponent(new URL(href, ROOT).pathname) === decodeURIComponent(new URL(f.url).pathname)) {
        hrefs.add(f.kind);
      }
      if (data.includes(f.uid)) seen.set(f.kind, data);
    }
  }
  return { status: res.status, seen, hrefs, raw: res.text };
}

/**
 * A query that matched nothing is only informative if you can see what came back — an error
 * multistatus and an empty one look identical through `seen`. Print the raw body whenever a
 * query returns no fixture at all.
 */
function dumpIfEmpty(label, q) {
  if (q.seen.size > 0) return;
  console.log(`  (${label} matched no fixture; raw multistatus, HTTP ${q.status})`);
  console.log(redact(q.raw.trim()).split('\n').map(l => `    ${l}`).join('\n'));
}

/** One printed account of what an expanded (or plain) blob carries. */
function report(prefix, d) {
  console.log(`${prefix} veventsInBlob=${d.blockCount} anyRDATE=${d.anyRdate ? 'yes' : 'no'} anyRRULE=${d.anyRrule ? 'yes' : 'no'}`);
  for (const line of d.rruleLines) console.log(`${prefix}   RRULE line: ${line}`);
  for (const line of d.rdateLines) console.log(`${prefix}   RDATE line: ${line}`);
  for (const b of d.blocks) {
    console.log(
      `${prefix}   ${b.dtstartLine ?? '(no DTSTART)'}  ->  ${asIso(b.dtstartMs)}` +
      (b.recurrenceIdLine ? `   ${b.recurrenceIdLine} -> ${asIso(b.recurrenceIdMs)}` : '   (no RECURRENCE-ID)'),
    );
  }
}

try {
  const homeUrl = await discover();
  console.log(`\nCalendar home: ${redact(homeUrl)}`);
  console.log(`Zone the fixtures are written in: ${ZONE} (host zone reports ${Intl.DateTimeFormat().resolvedOptions().timeZone})`);
  console.log('\nWindows sent (UTC instants, exactly as the shipped tool spells them):');
  for (const [k, w] of Object.entries(WINDOWS)) {
    console.log(`  ${k.padEnd(18)} ${davInstant(w.start)} .. ${davInstant(w.end)}   ${w.label}`);
  }
  console.log('\nExpected occurrence instants:');
  for (const f of FIXTURES) {
    console.log(`  ${f.kind}: ${f.occurrences.map(o => `${asIso(occurrenceMs(o))} (${o.label})`).join(', ')}`);
  }

  // --- MKCALENDAR. No fallback: these are writes into a live personal account. ----------
  const candidateUrl = new URL(`probe-165-${STAMP}/`, homeUrl).href;
  const mk = await dav('MKCALENDAR', candidateUrl, {
    body: `<?xml version="1.0" encoding="utf-8"?>\n` +
      `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
      `<d:displayname>probe-165 temporary</d:displayname>` +
      `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
      `</d:prop></d:set></c:mkcalendar>`,
  });
  if (mk.status < 200 || mk.status >= 300) {
    throw new Error(
      `MKCALENDAR failed (HTTP ${mk.status} ${mk.statusText}) — stopping rather than writing fixtures ` +
      `into an existing calendar.`,
    );
  }
  tempCalendarUrl = candidateUrl;
  const calendarUrl = candidateUrl;
  console.log(`\nMKCALENDAR ${mk.status}: temporary collection created (deleted in finally).`);

  // --- PUT the fixtures ------------------------------------------------------------------
  console.log('\n--- fixtures ---');
  for (const f of FIXTURES) {
    const url = new URL(`${f.uid.split('@')[0]}.ics`, calendarUrl).href;
    const put = await dav('PUT', url, { body: f.body, headers: { 'Content-Type': 'text/calendar; charset=utf-8' } });
    const ok = put.status >= 200 && put.status < 300;
    f.url = ok ? url : null;
    check(`PUT accepted: ${f.label}`, ok, `HTTP ${put.status} ${put.statusText}`);
  }
  if (FIXTURES.some(f => !f.url)) throw new Error('a fixture PUT was rejected; the measurement cannot be made');

  // --- Q1: the stored form, no expand -----------------------------------------------------
  console.log('\n--- Q1: wide window, NO expand (the stored form) ---');
  const q1 = await query(calendarUrl, WINDOWS.wide, false);
  const d1 = Object.fromEntries(FIXTURES.map(f => [f.kind, q1.seen.get(f.kind) && describe(q1.seen.get(f.kind))]));
  for (const f of FIXTURES) if (d1[f.kind]) report(`  [${f.kind}]`, d1[f.kind]);
  for (const kind of RDATE_KINDS) {
    check(`Q1 [${kind}]: the RDATE-only series is returned unexpanded`, !!d1[kind]);
    check(`Q1 [${kind}]: its RDATE value(s) survive storage intact`, !!d1[kind]?.anyRdate, d1[kind]?.rdateLines.join(' | ') ?? '(no blob)');
    check(`Q1 [${kind}]: it carries no RRULE (it recurs only by RDATE)`, d1[kind] ? !d1[kind].anyRrule : false);
  }
  check('Q1 [rrule]: the RRULE control is returned with its RRULE intact', !!d1.rrule?.anyRrule, d1.rrule?.rruleLines.join(' | ') ?? '(no blob)');

  // --- Q2: the shipped shape --------------------------------------------------------------
  console.log('\n--- Q2: wide window, WITH expand (the shape the shipped tool receives) ---');
  const q2 = await query(calendarUrl, WINDOWS.wide, true);
  const d2 = Object.fromEntries(FIXTURES.map(f => [f.kind, q2.seen.get(f.kind) && describe(q2.seen.get(f.kind))]));
  for (const f of FIXTURES) if (d2[f.kind]) report(`  [${f.kind}]`, d2[f.kind]);

  // THE MEASUREMENT #165 EXISTS FOR, and it is CONFIRMED: Cyrus strips RDATE from an expanded
  // block exactly as calendar-expand.probe.mjs measured it stripping RRULE.
  for (const kind of RDATE_KINDS) {
    check(`Q2 [${kind}]: the RDATE-only series is returned under expand`, !!d2[kind]);
    check(
      `Q2 [${kind}]: no expanded block carries an RDATE line — the strip is real, not derived`,
      d2[kind] ? !d2[kind].anyRdate : false,
      d2[kind] ? `blocks=${d2[kind].blockCount} rdateLines=[${d2[kind].rdateLines.join(' | ')}]` : '(no blob)',
    );
  }
  // The known RRULE measurement, reproduced in the same run so the two halves are comparable.
  check(
    'Q2 [rrule]: no expanded block of the RRULE control carries an RRULE line (the already-measured half)',
    d2.rrule ? !d2.rrule.anyRrule : false,
    d2.rrule ? `blocks=${d2.rrule.blockCount} rruleLines=[${d2.rrule.rruleLines.join(' | ')}]` : '(no blob)',
  );
  for (const f of FIXTURES) {
    const want = f.occurrences.map(occurrenceMs);
    const got = (d2[f.kind]?.blocks ?? []).map(b => b.dtstartMs);
    check(
      `Q2 [${f.kind}]: every occurrence is emitted (${want.length} expected)`,
      want.every(ms => got.includes(ms)) && got.length === want.length,
      `expected=[${want.map(asIso).join(', ')}] got=[${got.map(asIso).join(', ')}]`,
    );
    // The shape calendar-expand.probe.mjs measured for RRULE: the FIRST instance carries none.
    const d = d2[f.kind];
    const withId = (d?.blocks ?? []).filter(b => b.recurrenceIdLine);
    check(
      `Q2 [${f.kind}]: a RECURRENCE-ID on every block after the first, and none on the first`,
      !!d && d.blockCount > 0 && !d.blocks[0].recurrenceIdLine && withId.length === d.blockCount - 1,
      `first=${d?.blocks[0]?.recurrenceIdLine ?? '(none)'} withId=${withId.length}/${d?.blockCount ?? 0}`,
    );
  }

  // --- Q3/Q4: the filter does NOT walk RDATEs -----------------------------------------------
  // MEASURED, and it is not what the repo had recorded. The window below covers the FIRST RDATE
  // occurrence and nothing else. The RRULE control has an occurrence at the IDENTICAL instant,
  // so it is the discriminator: it separates "the server does not walk RDATEs when filtering"
  // from "this window is aimed wrong". The assertions encode the measurement, so a future run
  // FAILING here means Fastmail's build CHANGED — at which point the read path's window
  // handling is worth re-deciding, not this probe re-tuned.
  console.log('\n--- Q3: narrow window over ONE RDATE occurrence, WITH expand ---');
  const q3 = await query(calendarUrl, WINDOWS.narrowSecondRdate, true);
  const d3 = Object.fromEntries(FIXTURES.map(f => [f.kind, q3.seen.get(f.kind) && describe(q3.seen.get(f.kind))]));
  for (const f of FIXTURES) if (d3[f.kind]) report(`  [${f.kind}]`, d3[f.kind]);
  dumpIfEmpty('Q3', q3);
  const wantNarrow = occurrenceMs(RDATE_OCCURRENCES[1]);
  for (const kind of RDATE_KINDS) {
    check(
      `Q3 [${kind}]: the window over an RDATE occurrence emits NOTHING for the RDATE-only series`,
      !d3[kind] && !q3.hrefs.has(kind),
      `expected the occurrence at ${asIso(wantNarrow)}; matched=${q3.hrefs.has(kind)} blocks=${d3[kind]?.blockCount ?? 0}`,
    );
  }
  check(
    'Q3 [rrule]: the discriminator — the RRULE control IS emitted, at that same instant',
    d3.rrule?.blockCount === 1 && d3.rrule.blocks[0].dtstartMs === wantNarrow,
    `blocks=${d3.rrule?.blockCount ?? 0} [${(d3.rrule?.blocks ?? []).map(b => asIso(b.dtstartMs)).join(', ')}]`,
  );
  if (d3.rrule?.blocks[0]) {
    const b = d3.rrule.blocks[0];
    console.log(`  [rrule] RECURRENCE-ID form: ${b.recurrenceIdLine ?? '(absent)'}  params=${JSON.stringify(b.recurrenceIdParams ?? '')}`);
    check(
      'Q3 [rrule]: that occurrence names itself with a RECURRENCE-ID at the occurrence instant, in the bare-Z form',
      b.recurrenceIdMs === wantNarrow && /^RECURRENCE-ID:\d{8}T\d{6}Z$/.test(b.recurrenceIdLine ?? ''),
      `${b.recurrenceIdLine ?? '(absent)'} -> ${asIso(b.recurrenceIdMs)}`,
    );
  }

  console.log('\n--- Q4: the same narrow window, NO expand (does the time-range filter walk RDATEs?) ---');
  const q4 = await query(calendarUrl, WINDOWS.narrowSecondRdate, false);
  const d4 = Object.fromEntries(FIXTURES.map(f => [f.kind, q4.seen.get(f.kind) && describe(q4.seen.get(f.kind))]));
  for (const f of FIXTURES) if (d4[f.kind]) report(`  [${f.kind}]`, d4[f.kind]);
  dumpIfEmpty('Q4', q4);
  for (const kind of RDATE_KINDS) {
    check(
      `Q4 [${kind}]: the time-range FILTER does not match the resource either — the RDATEs are invisible to it`,
      !q4.hrefs.has(kind) && !d4[kind],
      `matched=${q4.hrefs.has(kind)} blobReturned=${!!d4[kind]}`,
    );
  }
  check(
    'Q4 [rrule]: the discriminator — the filter DOES match the RRULE control on the same window',
    q4.hrefs.has('rrule') && !!d4.rrule,
    `matched=${q4.hrefs.has('rrule')} blobReturned=${!!d4.rrule}`,
  );

  // --- Which span does the filter index for an RDATE-only resource? -----------------------
  // The two windows below carry NO assertion of their own beyond the span conclusion: each
  // exists to discriminate one alternative explanation of Q3/Q4.
  //   narrowDtstart      covers the series DTSTART and no other occurrence. If the resource
  //                      matches here but not on an RDATE occurrence, the filter is reading
  //                      DTSTART rather than being blind to the resource altogether.
  //   betweenOccurrences covers NO occurrence at all, but lies between DTSTART and the last
  //                      RDATE. If the resource matched here, the filter would be indexing the
  //                      whole DTSTART..last-RDATE span; it does not, so the indexed span is
  //                      DTSTART..DTSTART+DURATION and stops there.
  console.log('\n--- what span the time-range filter indexes for an RDATE-only resource ---');
  const span = {};
  for (const key of ['narrowDtstart', 'betweenOccurrences']) {
    span[key] = {};
    for (const expand of [true, false]) {
      const q = await query(calendarUrl, WINDOWS[key], expand);
      span[key][expand ? 'expand' : 'plain'] = q;
      const kinds = FIXTURES.filter(f => q.hrefs.has(f.kind)).map(f => f.kind).join(', ') || '(none)';
      const d = q.seen.get('rdate') && describe(q.seen.get('rdate'));
      console.log(
        `  ${key.padEnd(19)} ${expand ? 'expand ' : 'plain  '} ` +
        `${davInstant(WINDOWS[key].start)}..${davInstant(WINDOWS[key].end)}  matched=[${kinds}]` +
        (d ? `  rdateBlocks=${d.blockCount} [${d.blocks.map(b => asIso(b.dtstartMs)).join(', ')}]` : ''),
      );
    }
  }
  for (const kind of RDATE_KINDS) {
    check(
      `span [${kind}]: the DTSTART window DOES match — the filter reads DTSTART, it is not blind to the resource`,
      span.narrowDtstart.expand.hrefs.has(kind) && span.narrowDtstart.plain.hrefs.has(kind),
      `expand=${span.narrowDtstart.expand.hrefs.has(kind)} plain=${span.narrowDtstart.plain.hrefs.has(kind)}`,
    );
    check(
      `span [${kind}]: a window between DTSTART and the last RDATE matches nothing — the indexed span is DTSTART..DTSTART+DURATION`,
      !span.betweenOccurrences.expand.hrefs.has(kind) && !span.betweenOccurrences.plain.hrefs.has(kind),
      `expand=${span.betweenOccurrences.expand.hrefs.has(kind)} plain=${span.betweenOccurrences.plain.hrefs.has(kind)}`,
    );
  }
  // If the two serialisations ever diverge, that is the headline, so say it in one line.
  const formsAgree = RDATE_KINDS.every(k =>
    (!!d3[k] === !!d3[RDATE_KINDS[0]]) && (q4.hrefs.has(k) === q4.hrefs.has(RDATE_KINDS[0])));
  check(
    'the two RDATE serialisations behave IDENTICALLY (comma-joined vs one property per line)',
    formsAgree,
    RDATE_KINDS.map(k => `${k}: q3blob=${!!d3[k]} q4matched=${q4.hrefs.has(k)}`).join('; '),
  );

  // --- The bytes, so the report carries them ----------------------------------------------
  console.log('\n--- expanded calendar-data blobs, verbatim (redacted) ---');
  for (const f of FIXTURES) {
    const blob = q2.seen.get(f.kind);
    console.log(`\n[Q2 expanded] ${f.label}`);
    console.log(blob ? redact(blob.replace(/\r\n/g, '\n').trim()) : '(not returned)');
  }
  for (const f of FIXTURES) {
    for (const [label, q] of [['Q3 expanded', q3], ['Q4 unexpanded', q4]]) {
      const blob = q.seen.get(f.kind);
      console.log(`\n[${label}, narrow window over an RDATE occurrence] ${f.label}`);
      console.log(blob ? redact(blob.replace(/\r\n/g, '\n').trim()) : '(not returned)');
    }
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
} finally {
  // Deleting the collection removes both fixtures in one request. The PROPFIND afterwards is
  // the only thing that proves it: a DELETE that returns 2xx and leaves the collection standing
  // would otherwise go unnoticed, and this probe writes into a live personal account.
  if (tempCalendarUrl) {
    const del = await dav('DELETE', tempCalendarUrl);
    console.log(`\nCleanup: DELETE temporary collection -> HTTP ${del.status} ${del.statusText}`);
    const after = await dav('PROPFIND', tempCalendarUrl, { body: PROPFIND(['d:resourcetype']), headers: { Depth: '0' } });
    check(
      'cleanup: the temporary collection is gone (PROPFIND no longer finds it)',
      after.status === 404 || after.status === 410,
      `DELETE ${del.status}, PROPFIND ${after.status} ${after.statusText}`,
    );
    if (after.status !== 404 && after.status !== 410) {
      console.log(`  ⚠ the temporary collection may still exist at ${redact(tempCalendarUrl)} — remove it by hand`);
    }
  }
}

console.log('');
process.exit(failures() > 0 ? 1 : 0);
