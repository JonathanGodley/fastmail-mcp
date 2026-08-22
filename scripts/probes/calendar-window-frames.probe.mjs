// What this probe settles
// -----------------------
// #162 records two MISSING-EVENT bugs in the calendar read path that were derived from the
// Cyrus source and from RFC 4791, and never observed. Both say the server never returns a
// resource the caller should have seen, so no client-side change can fix either one; both
// therefore decide whether the request window this server sends has to be widened before it
// goes over the wire. A decision that changes what goes on the wire deserves an observation
// rather than a derivation, and this probe is that observation.
//
//   Bug 1 - a sub-day window loses all-day events. Cyrus matches a date-only (VALUE=DATE)
//   value on its UTC day and never shifts it into a collection timezone. A caller in a +10
//   zone asking for "the morning of the 16th" (local 00:00..10:00) sends the UTC instants
//   15 Jun 14:00Z .. 16 Jun 00:00Z, which cannot intersect the all-day event's UTC day of
//   [16 Jun 00:00Z, 17 Jun 00:00Z). Predicted: the resource is not returned at all.
//
//   Bug 2 - a genuinely floating timed value (no TZID, no Z) is judged as UTC. RFC 4791
//   section 9.9 says such a value is resolved in the collection's calendar timezone, but the
//   prediction is that Fastmail sets no calendar-timezone on a collection, so the filter
//   falls back to UTC. A floating 20:00 event then sits at 20:00Z: outside a +10 caller's
//   evening window (local 19:00..21:00 = 09:00Z..11:00Z) and inside a window over 20:00Z.
//
// Two supporting facts the same run settles:
//   (a) whether Fastmail exposes CALDAV:calendar-timezone on a collection or on the calendar
//       home at all - the property whose absence is what makes the UTC fallback happen;
//   (b) what shape a floating event comes back in after <C:expand>, i.e. whether expansion
//       stamps it with a trailing Z and destroys the floating marker (#162 item 5).
//
// WHAT THE FIRST RUN FOUND, so the checks below read as assertions of a measured platform
// rather than as a rerun of the prediction. Both bugs are real in their caller-visible
// outcome; one of them is real for a different reason than the issue records.
//
//   Bug 2 landed exactly as derived. The floating event is not even MATCHED by the
//   local-evening window and is returned by the 19:00Z..21:00Z one; the TZID control does the
//   opposite, which is what proves the two windows are aimed where this probe says they are.
//   No calendar-timezone exists at either level (both 404), so RFC 4791 section 9.9 has
//   nothing to resolve against and UTC is what is left.
//
//   Bug 1's OUTCOME is real and its MECHANISM is not. The prediction was that the server
//   never returns the resource. What actually happens on the boundary-touching window
//   (…14:00Z .. 16 Jun 00:00Z, whose exclusive end is exactly the start of the all-day
//   event's UTC day) is that the time-range FILTER matches the resource - so an unexpanded
//   query returns the whole thing - while <C:expand> in the same request emits zero VEVENTs.
//   The shipped tool always sends expand, so it receives a matched resource carrying no
//   occurrence and reports nothing: same missing event, one layer further in. Pull the window
//   an hour back so its end no longer touches midnight UTC and the resource is not matched at
//   all, which is what identifies the first result as the filter treating the window's
//   exclusive end as inclusive rather than as a wider stored span.
//
//   Side effect, and it closes the libical question #162's second comment left open: an
//   all-day event survives expansion as `DTSTART;VALUE=DATE:20270616`, unshifted and with no
//   Z. The expansion path does not honour a zone for a date value either, so a date-only
//   value is UTC-framed on every path and the deviation from section 9.9 is total.
//
// HOW TO READ THE RESULT. Each check asserts an observed platform behaviour, so PASS means
// the platform still behaves as measured here and FAIL means it has changed - at which point
// the redesign resting on these facts needs re-deciding, not the probe re-tuned.
//
// WHY RAW CALDAV, AND WHY RAW PUT. This measures the platform, not our parsing, so it talks
// HTTP directly and builds the same XML tsdav builds for `fetchCalendarObjects` with
// `timeRange` + `expand` (a calendar-query REPORT at Depth 1, asking for d:getetag and a
// c:calendar-data carrying c:expand, filtered VCALENDAR > VEVENT > c:time-range, with both
// ranges written as the same UTC basic-format instants). Raw PUT is not a shortcut either:
// this server's own create path always writes a TZID since #157, so it cannot author the
// floating fixture at all. If Fastmail rejects or rewrites the floating PUT, that is itself
// the answer to bug 2 and is reported as such rather than worked around.
//
// FIXTURES AND CLEANUP. Three synthetic events (an all-day, a floating timed, and a
// TZID-stamped control that acts as the discriminator proving the windows are aimed where
// this probe says they are) are written into a temporary calendar created by MKCALENDAR, far
// out in June 2027 so nothing real shares the window. None carries an ATTENDEE or ORGANIZER:
// a participant would make the server send real iTIP mail (see the README). The finally block
// deletes the whole temporary collection, which removes every fixture atomically; if
// MKCALENDAR is unavailable the probe falls back to an existing calendar and deletes each
// resource it PUT, individually, in the same finally. Each run mints its own collection name,
// so a re-run never needs a manual delete first.
//
// Run: python scripts/probes/run-probe.py calendar-window-frames.probe.mjs

import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-window-frames.probe.mjs');
  process.exit(1);
}

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');

// Every CalDAV href on this server embeds the account's own address. Nothing printed by this
// probe is allowed to carry it: probe output gets pasted into issues.
const redact = s => String(s).split(USERNAME).join('<account>');

// The zone the derivation in #162 is written against, and the zone this deployment runs in.
// Hard-coded rather than read from the environment because the fixture's VTIMEZONE has to
// match it, and because the windows below are only a discriminating test in a zone whose
// offset is far enough from UTC to separate the frames.
const ZONE = 'Australia/Sydney';
// June, deliberately: Sydney is on standard time (+10:00) with no DST transition anywhere
// near, so the offsets below are stable and the arithmetic in this header is checkable by eye.
const DAY = { y: 2027, mo: 6, d: 16 };

// ---------------------------------------------------------------------------
// Wall clock -> instant, in ZONE. Two passes: sample the offset at the naive guess, then
// re-check at the instant that lands on. Enough for a date nowhere near a transition.
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

function wallToUtcIso(y, mo, d, h, mi = 0, s = 0) {
  const guess = Date.UTC(y, mo - 1, d, h, mi, s);
  const once = guess - offsetMsAt(ZONE, guess);
  const twice = guess - offsetMsAt(ZONE, once);
  return new Date(twice).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

const utcIso = (y, mo, d, h, mi = 0) =>
  new Date(Date.UTC(y, mo - 1, d, h, mi, 0)).toISOString().replace(/\.\d{3}Z$/, 'Z');

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
const hasEmptyEl = (xml, name) => new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?\\s*\\/>`).test(xml);
const responses = xml => elAll(xml, 'response');
const abs = href => new URL(href, ROOT).href;

const unescapeXml = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');

const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

// The exact request tsdav sends for fetchCalendarObjects({ timeRange, expand }); `expand`
// false drops only the c:expand child, which is the one difference between the two calls.
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
const uid = kind => `probe-162-${kind}-${STAMP}@probe.invalid`;

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

const pad = n => String(n).padStart(2, '0');
const YMD = `${DAY.y}${pad(DAY.mo)}${pad(DAY.d)}`;
const YMD_NEXT = `${DAY.y}${pad(DAY.mo)}${pad(DAY.d + 1)}`;
const DTSTAMP = new Date().toISOString().slice(0, 19).replace(/[-:]/g, '') + 'Z';

const FIXTURES = [
  {
    kind: 'allday',
    label: 'all-day (VALUE=DATE)',
    uid: uid('allday'),
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-162//EN', 'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT', `UID:${uid('allday')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART;VALUE=DATE:${YMD}`, `DTEND;VALUE=DATE:${YMD_NEXT}`,
      'SUMMARY:probe-162 all-day fixture', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
  {
    kind: 'floating',
    label: 'floating timed (no TZID, no Z)',
    uid: uid('floating'),
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-162//EN', 'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT', `UID:${uid('floating')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART:${YMD}T200000`, `DTEND:${YMD}T210000`,
      'SUMMARY:probe-162 floating fixture', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
  {
    kind: 'zoned',
    label: `TZID control (${ZONE} 20:00)`,
    uid: uid('zoned'),
    body: ics([
      'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//fastmail-mcp//probe-162//EN', 'CALSCALE:GREGORIAN',
      ...VTIMEZONE_SYDNEY,
      'BEGIN:VEVENT', `UID:${uid('zoned')}`, `DTSTAMP:${DTSTAMP}`,
      `DTSTART;TZID=${ZONE}:${YMD}T200000`, `DTEND;TZID=${ZONE}:${YMD}T210000`,
      'SUMMARY:probe-162 zoned control fixture', 'END:VEVENT', 'END:VCALENDAR',
    ]),
  },
];

// ---------------------------------------------------------------------------
// Windows. Each names the caller-side question it stands for.
// ---------------------------------------------------------------------------
const WINDOWS = {
  subDay: {
    label: `sub-day: local ${DAY.y}-${pad(DAY.mo)}-${pad(DAY.d)} 00:00..10:00 in ${ZONE}`,
    start: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 0),
    end: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 10),
  },
  // Same shape as subDay but ending an hour EARLIER, so its exclusive end (23:00Z) does not
  // touch the start of the all-day event's UTC day. It separates "the prefilter matches a
  // wider span" from "the prefilter treats the window's exclusive end as inclusive".
  subDayEarly: {
    label: `sub-day, ending clear of midnight UTC: local 00:00..09:00 in ${ZONE}`,
    start: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 0),
    end: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 9),
  },
  wholeLocalDay: {
    label: `control: the whole local day in ${ZONE}`,
    start: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 0),
    end: wallToUtcIso(DAY.y, DAY.mo, DAY.d + 1, 0),
  },
  utcDay: {
    label: 'the same date written as a UTC day',
    start: utcIso(DAY.y, DAY.mo, DAY.d, 0),
    end: utcIso(DAY.y, DAY.mo, DAY.d + 1, 0),
  },
  eveningLocal: {
    label: `evening: local 19:00..21:00 in ${ZONE}`,
    start: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 19),
    end: wallToUtcIso(DAY.y, DAY.mo, DAY.d, 21),
  },
  eveningUtc: {
    label: 'the same wall-clock hours read as UTC (19:00Z..21:00Z)',
    start: utcIso(DAY.y, DAY.mo, DAY.d, 19),
    end: utcIso(DAY.y, DAY.mo, DAY.d, 21),
  },
};

// ---------------------------------------------------------------------------
const { check, failures } = makeChecker();

let tempCalendarUrl = null;     // set only when MKCALENDAR succeeded
const putResourceUrls = [];     // fallback cleanup list

async function discover() {
  const rootPf = await dav('PROPFIND', ROOT, { body: PROPFIND(['d:current-user-principal']), headers: { Depth: '0' } });
  const principal = el(el(rootPf.text, 'current-user-principal') ?? '', 'href');
  if (!principal) throw new Error(`could not read current-user-principal (HTTP ${rootPf.status})`);

  const prinPf = await dav('PROPFIND', abs(principal.trim()), {
    body: PROPFIND(['c:calendar-home-set']), headers: { Depth: '0' },
  });
  const home = el(el(prinPf.text, 'calendar-home-set') ?? '', 'href');
  if (!home) throw new Error(`could not read calendar-home-set (HTTP ${prinPf.status})`);
  return { principalUrl: abs(principal.trim()), homeUrl: abs(home.trim()) };
}

/** Which of the requested props came back 200 and which 404, for one PROPFIND response block. */
function propStatus(responseXml, names) {
  const out = {};
  for (const stat of elAll(responseXml, 'propstat')) {
    const code = (el(stat, 'status') ?? '').trim();
    const prop = el(stat, 'prop') ?? '';
    for (const n of names) {
      const present = new RegExp(`<(?:[\\w-]+:)?${n}[\\s>/]`).test(prop);
      if (present) out[n] = { code, value: el(prop, n) };
    }
  }
  for (const n of names) if (!(n in out)) out[n] = { code: '(not mentioned)', value: undefined };
  return out;
}

const TZ_PROPS = ['calendar-timezone', 'calendar-timezone-id'];

async function reportTimezoneProps(label, url) {
  const pf = await dav('PROPFIND', url, {
    body: PROPFIND(['d:displayname', 'c:calendar-timezone', 'c:calendar-timezone-id']),
    headers: { Depth: '0' },
  });
  const block = responses(pf.text)[0] ?? pf.text;
  const status = propStatus(block, TZ_PROPS);
  console.log(`  ${label}`);
  for (const n of TZ_PROPS) {
    const v = status[n].value;
    const shown = v && v.trim() ? ` value=${JSON.stringify(unescapeXml(v).trim().slice(0, 120))}` : '';
    console.log(`    ${n}: ${status[n].code}${shown}`);
  }
  return status;
}

/** Run one calendar-query and return { fixtures seen, raw calendar-data per fixture }. */
async function query(calendarUrl, win, expand) {
  const res = await dav('REPORT', calendarUrl, {
    body: calendarQueryXml(win.start, win.end, expand),
    headers: { Depth: '1' },
  });
  const seen = new Map();   // fixture kind -> the calendar-data that actually carried a VEVENT
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

const dtLines = data =>
  (data.match(/^(?:DTSTART|DTEND|RECURRENCE-ID)[;:].*$/gm) ?? []).map(l => l.replace(/\r$/, ''));

try {
  const { principalUrl, homeUrl } = await discover();
  console.log(`\nCalendar home: ${redact(homeUrl)}`);
  console.log(`Configured zone assumed by this probe: ${ZONE} (host zone reports ${Intl.DateTimeFormat().resolvedOptions().timeZone})`);
  console.log('');
  console.log('Windows sent (as UTC instants, exactly as the tool would spell them):');
  for (const [k, w] of Object.entries(WINDOWS)) {
    console.log(`  ${k.padEnd(14)} ${davInstant(w.start)} .. ${davInstant(w.end)}   ${w.label}`);
  }

  // --- MKCALENDAR, or fall back to an existing calendar -------------------------------
  const candidateUrl = new URL(`probe-162-${STAMP}/`, homeUrl).href;
  const mk = await dav('MKCALENDAR', candidateUrl, {
    body: `<?xml version="1.0" encoding="utf-8"?>\n` +
      `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
      `<d:displayname>probe-162 temporary</d:displayname>` +
      `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
      `</d:prop></d:set></c:mkcalendar>`,
  });
  let calendarUrl;
  if (mk.status >= 200 && mk.status < 300) {
    tempCalendarUrl = candidateUrl;
    calendarUrl = candidateUrl;
    console.log(`\nMKCALENDAR ${mk.status}: temporary collection created (deleted in finally).`);
  } else {
    console.log(`\nMKCALENDAR ${mk.status} ${mk.statusText} - falling back to an existing calendar; each resource is deleted individually.`);
    const homePf = await dav('PROPFIND', homeUrl, {
      body: PROPFIND(['d:resourcetype', 'd:displayname', 'c:supported-calendar-component-set']),
      headers: { Depth: '1' },
    });
    const cal = responses(homePf.text).find(r => {
      const rt = el(r, 'resourcetype') ?? '';
      const comps = el(r, 'supported-calendar-component-set') ?? '';
      return hasEmptyEl(rt, 'calendar') && (!comps || /name="VEVENT"/.test(comps));
    });
    const href = cal && el(cal, 'href');
    if (!href) throw new Error('no writable VEVENT calendar found for the fallback path');
    calendarUrl = abs(href.trim());
  }
  check('a target calendar is available for the fixtures', !!calendarUrl, redact(calendarUrl));

  // --- Supporting fact (a): is there a collection timezone at all? --------------------
  console.log('\n--- CALDAV:calendar-timezone (the property whose absence forces the UTC fallback) ---');
  const calTz = await reportTimezoneProps('target calendar', calendarUrl);
  const homeTz = await reportTimezoneProps('calendar home', homeUrl);
  const absent = s => !s['calendar-timezone'].value?.trim() && !s['calendar-timezone-id'].value?.trim();
  check(
    'derivation confirmed: Fastmail exposes no calendar-timezone on the collection',
    absent(calTz),
    `collection reports ${calTz['calendar-timezone'].code} / ${calTz['calendar-timezone-id'].code}`,
  );
  check(
    'derivation confirmed: no calendar-timezone on the calendar home either (the Cyrus fallback level)',
    absent(homeTz),
    `home reports ${homeTz['calendar-timezone'].code} / ${homeTz['calendar-timezone-id'].code}`,
  );

  // --- PUT the fixtures ---------------------------------------------------------------
  console.log('\n--- fixtures ---');
  for (const f of FIXTURES) {
    const url = new URL(`${f.uid.split('@')[0]}.ics`, calendarUrl).href;
    const put = await dav('PUT', url, { body: f.body, headers: { 'Content-Type': 'text/calendar; charset=utf-8' } });
    const ok = put.status >= 200 && put.status < 300;
    if (ok) putResourceUrls.push(url);
    check(`PUT accepted: ${f.label}`, ok, `HTTP ${put.status} ${put.statusText}`);
    f.url = ok ? url : null;
  }

  // --- Did the server keep the floating value floating? -------------------------------
  const floating = FIXTURES.find(f => f.kind === 'floating');
  let storedFloating = '';
  if (floating.url) {
    const got = await dav('GET', floating.url);
    storedFloating = got.text;
    const stored = dtLines(storedFloating);
    console.log(`  stored form of the floating fixture: ${stored.join(' | ') || '(no DTSTART found)'}`);
    check(
      'the floating PUT survives storage unrewritten (no TZID, no Z on DTSTART)',
      stored.some(l => /^DTSTART:\d{8}T\d{6}$/.test(l)),
      stored.join(' | '),
    );
  }

  // --- The queries --------------------------------------------------------------------
  const results = {};
  console.log('\n--- what each window returns (expanded = the shape the shipped tool sends) ---');
  for (const [key, win] of Object.entries(WINDOWS)) {
    const expanded = await query(calendarUrl, win, true);
    const plain = await query(calendarUrl, win, false);
    results[key] = { expanded, plain };
    // Two different things, and this probe found they can disagree: `matched` is the resource
    // being named in the multistatus at all (the server's time-range filter said yes), `events`
    // is a VEVENT actually arriving for it.
    const ev = r => FIXTURES.filter(f => r.seen.has(f.kind)).map(f => f.kind).join(', ') || '(none)';
    const hr = r => FIXTURES.filter(f => r.hrefs.has(f.kind)).map(f => f.kind).join(', ') || '(none)';
    console.log(`  ${key}`);
    console.log(`    ${davInstant(win.start)}..${davInstant(win.end)}`);
    console.log(`    expanded: events=[${ev(expanded)}] matched-resources=[${hr(expanded)}]`);
    console.log(`    plain:    events=[${ev(plain)}] matched-resources=[${hr(plain)}]`);
  }

  const inWin = (key, kind, mode = 'expanded') => results[key][mode].seen.has(kind);
  const matched = (key, kind, mode = 'expanded') => results[key][mode].hrefs.has(kind);

  // --- Bug 1 --------------------------------------------------------------------------
  console.log('\n--- Bug 1: a sub-day window loses the all-day event ---');
  check(
    'bug 1: a sub-day local-morning window yields NO all-day occurrence on the shipped path',
    !inWin('subDay', 'allday'),
    `window ${davInstant(WINDOWS.subDay.start)}..${davInstant(WINDOWS.subDay.end)}`,
  );
  check(
    'control: the same all-day event IS returned for the whole-local-day window',
    inWin('wholeLocalDay', 'allday'),
  );
  check(
    'a date-only value is matched on its UTC day, not on the caller local day',
    inWin('utcDay', 'allday'),
  );
  // The mechanism, and it is not the one #162 derived. The filter and the expansion disagree.
  check(
    'the time-range filter MATCHES the resource when the window end touches the UTC day start, while expand emits nothing',
    matched('subDay', 'allday') && inWin('subDay', 'allday', 'plain') && !inWin('subDay', 'allday'),
    `matched=${matched('subDay', 'allday')} plainEvents=${inWin('subDay', 'allday', 'plain')} expandedEvents=${inWin('subDay', 'allday')}`,
  );
  check(
    'that match is only the touching boundary: a sub-day window ending clear of midnight UTC matches nothing, either mode',
    !matched('subDayEarly', 'allday') && !inWin('subDayEarly', 'allday', 'plain'),
    `window ${davInstant(WINDOWS.subDayEarly.start)}..${davInstant(WINDOWS.subDayEarly.end)}`,
  );
  check(
    'residue: the all-day event comes back for ANY window inside its UTC day, including a two-hour one',
    inWin('eveningLocal', 'allday') && inWin('eveningUtc', 'allday'),
  );

  // --- Bug 2 --------------------------------------------------------------------------
  console.log('\n--- Bug 2: a floating time is judged as UTC ---');
  check(
    'bug 2 (half a): the floating 20:00 event is NOT returned for the local-evening window',
    !inWin('eveningLocal', 'floating') && !matched('eveningLocal', 'floating'),
    `window ${davInstant(WINDOWS.eveningLocal.start)}..${davInstant(WINDOWS.eveningLocal.end)}`,
  );
  check(
    'bug 2 (half b): it IS returned for the window covering 20:00Z',
    inWin('eveningUtc', 'floating'),
    `window ${davInstant(WINDOWS.eveningUtc.start)}..${davInstant(WINDOWS.eveningUtc.end)}`,
  );
  check(
    'discriminator: the TZID control IS returned for the local-evening window',
    inWin('eveningLocal', 'zoned'),
  );
  check(
    'discriminator: the TZID control is NOT returned for the 19:00Z..21:00Z window',
    !inWin('eveningUtc', 'zoned'),
  );

  // --- Supporting fact (b): what shape does expansion hand back? ----------------------
  console.log('\n--- expansion output shapes ---');
  const floatData = results.eveningUtc.expanded.seen.get('floating');
  if (floatData) {
    const lines = dtLines(floatData);
    console.log(`  floating, expanded: ${lines.join(' | ')}`);
    check(
      'residual confirmed: expansion rewrites the floating value to UTC, destroying the floating marker',
      lines.some(l => /^DTSTART[^:]*:\d{8}T\d{6}Z$/.test(l)),
      lines.join(' | '),
    );
  } else {
    check('a floating occurrence was available to inspect after expansion', false, 'not returned by any window');
  }
  const alldayData = results.wholeLocalDay.expanded.seen.get('allday');
  if (alldayData) {
    const lines = dtLines(alldayData);
    console.log(`  all-day, expanded:  ${lines.join(' | ')}`);
    // Settles the libical question left open in #162's second comment: the recurrence-walk /
    // expansion path does NOT shift a date value into any zone, so a date-only value is
    // UTC-framed everywhere and the deviation from RFC 4791 section 9.9 is total.
    check(
      'expansion leaves a date-only value as a DATE, unshifted and un-Z-stamped',
      lines.some(l => /^DTSTART;VALUE=DATE:\d{8}$/.test(l)),
      lines.join(' | '),
    );
  }
  const zonedData = results.eveningLocal.expanded.seen.get('zoned');
  if (zonedData) {
    const lines = dtLines(zonedData);
    console.log(`  zoned, expanded:    ${lines.join(' | ')}`);
    check(
      'expansion rewrites a TZID value to the UTC instant it names and drops the TZID',
      lines.some(l => /^DTSTART:\d{8}T\d{6}Z$/.test(l)),
      lines.join(' | '),
    );
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
} finally {
  // Removing the collection removes every fixture inside it in one request; the per-resource
  // loop is only for the fallback path, where the fixtures live in a calendar that must stay.
  if (tempCalendarUrl) {
    const del = await dav('DELETE', tempCalendarUrl);
    console.log(`\nCleanup: DELETE temporary collection -> HTTP ${del.status} ${del.statusText}`);
    if (del.status >= 300) {
      console.log('  collection delete failed; deleting each fixture resource individually');
      for (const url of putResourceUrls) {
        const r = await dav('DELETE', url);
        console.log(`    DELETE ${redact(url)} -> HTTP ${r.status}`);
      }
    }
  } else {
    for (const url of putResourceUrls) {
      const r = await dav('DELETE', url);
      console.log(`Cleanup: DELETE ${redact(url)} -> HTTP ${r.status}`);
    }
  }
}

console.log('');
process.exit(failures() > 0 ? 1 : 0);
