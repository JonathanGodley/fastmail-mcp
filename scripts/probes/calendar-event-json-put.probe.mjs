// What this probe settles
// -----------------------
// This server writes calendar events over CalDAV as `text/calendar` (the JMAP session advertises
// no calendars capability), and on a `text/calendar` PUT Fastmail's server (Cyrus) attaches no
// VTIMEZONE for a TZID the event uses (#166). Cyrus registers a second body type for a CalDAV PUT,
// though: `application/event+json`, a JSCalendar event (RFC 8984). Its conversion path adds a
// VTIMEZONE itself:
//
//   imap/http_caldav.c:285-290   the type is registered in caldav_mime_types, under #ifdef WITH_JMAP
//   imap/http_dav.c:7225-7232    meth_put picks the mime entry by Content-Type; an unregistered
//                                type is refused 403 with the CALDAV:supported-calendar-data
//                                precondition (not 415)
//   imap/http_dav.c:7345         mime->to_object(body)
//   imap/jmap_ical.c:8258        jevent_string_as_icalcomponent -> jmapical_toical (:8273)
//   imap/jmap_ical.c:8190        icalcomponent_add_required_timezones (ical_support.c:3086): one
//                                VTIMEZONE per referenced zone, truncated to the event's span,
//                                with TZUNTIL set to the span end
//
// Whether Fastmail's deployment is compiled WITH_JMAP and accepts that type on a CalDAV PUT cannot
// be read off the source, so this probe measures it. Four conditions, each reported separately:
//
//   1. a PUT of an `application/event+json` event into a calendar collection is accepted (2xx);
//   2. a raw GET of the stored resource as `text/calendar` holds exactly one VTIMEZONE, whose TZID
//      is the event's zone;
//   3. that block carries a TZUNTIL;
//   4. DTSTART keeps the intended zone and wall time, and the end survives as the intended instant.
//
// A refused PUT prints its status and the DAV:error precondition element name only; conditions 2-4
// are then reported as not reached. A JSCalendar body the converter rejects comes back 403 with
// CALDAV:valid-calendar-data (caldav_put, http_caldav.c:3966), which is how "the type is accepted
// but this body is not" differs from "the type is not registered".
//
// THE FIXTURE is minimal JSCalendar, checked against what jmap_ical.c requires: `uid` is mandatory
// (jmapical_toical, :8168) and so is `start`, as a local date-time with seconds and no zone
// designator (startend_to_ical, :4617; parse_datetime, :1322). `@type` is optional but must be
// "Event" when present (:7697). `timeZone` must resolve through libical's builtin zones (:4562).
// The event is timed in Australia/Sydney, starting at local 01:30 on 2026-10-04 for PT2H, so its
// span crosses that morning's spring-forward (02:00 AEST -> 03:00 AEDT). The two hours are elapsed
// time: the event ends at 04:30 AEDT. With no end-zone location Cyrus writes DURATION, not DTEND
// (:4722-4733), so condition 4 accepts either end form and checks the instant it names.
//
// WHY A STORED-RESOURCE GET SHOWS THE ANSWER. caldav_store_resource strips VTIMEZONEs on write only
// under ALLOW_CAL_NOTZ (caldav_util.c:1080), and calendar-tzdist.probe.mjs measured that this
// deployment does not advertise calendar-no-timezone, so a block added by the converter survives
// storage and a plain GET returns it.
//
// SAFETY. Raw CalDAV over bare `fetch`, with no tsdav and not the built server. The fixture goes into
// a temporary collection minted by MKCALENDAR; if MKCALENDAR fails the probe stops rather than
// writing into a real calendar. The finally block deletes the whole collection and PROPFINDs to
// confirm it is gone. No participants, so nothing is mailed. Output is PASS/FAIL, status codes,
// zone names and the stored VTIMEZONE (timezone data, not account data): no collection URL, UID,
// event title or account address is printed, so a run can be quoted verbatim into a public issue.
//
// RESULT, 25 Sep 2026: condition 1 FAILS: the PUT is refused 403 with CALDAV:supported-calendar-data,
// and conditions 2-4 are not reached. On the PUT path that precondition has one emitter
// (http_dav.c:7231: no entry in the mime table matched the Content-Type), so this deployment does
// not register `application/event+json` for CalDAV PUT. The platform will not supply a VTIMEZONE
// by this route either.
//
// Run: python scripts/probes/run-probe.py calendar-event-json-put.probe.mjs

import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-event-json-put.probe.mjs');
  process.exit(1);
}

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');

// Applied to every server-derived string that reaches stdout.
const redact = s => String(s)
  .split(USERNAME).join('<account>')
  .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');

const ZONE = 'Australia/Sydney';
const START_LOCAL = '2026-10-04T01:30:00';
const DURATION = 'PT2H';
const STAMP = Date.now();
const UID = `probe-jscal-put-${STAMP}@probe.invalid`;
// The collection's own path segment is the probe's constant, not account data, so it is the one
// thing printed if a manual delete is ever needed.
const SEGMENT = `probe-jscal-put-${STAMP}`;

// Independent of Cyrus: the instants the fixture names, via Intl.
function offsetMsAt(utcMs) {
  const p = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: ZONE, hour12: false,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(new Date(utcMs)).map(x => [x.type, x.value]),
  );
  return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second) - utcMs;
}
// Two passes, because both wall clocks here sit beside a transition (neither is in the gap).
function wallToUtcMs(wall) {
  const y = +wall.slice(0, 4), mo = +wall.slice(4, 6), d = +wall.slice(6, 8);
  const h = +wall.slice(9, 11), mi = +wall.slice(11, 13), s = +wall.slice(13, 15);
  const naive = Date.UTC(y, mo - 1, d, h, mi, s);
  return naive - offsetMsAt(naive - offsetMsAt(naive));
}
function utcMsToWall(ms) {
  const local = new Date(ms + offsetMsAt(ms));
  return local.toISOString().slice(0, 19).replace(/[-:]/g, '');
}
const basicUtc = ms => new Date(ms).toISOString().slice(0, 19).replace(/[-:]/g, '') + 'Z';

const WANT_START_WALL = START_LOCAL.replace(/[-:]/g, '');                  // 20261004T013000
const WANT_START_MS = wallToUtcMs(WANT_START_WALL);
const WANT_END_MS = WANT_START_MS + 2 * 3600 * 1000;
const WANT_END_WALL = utcMsToWall(WANT_END_MS);                            // 20261004T043000

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
const abs = href => new URL(href, ROOT).href;

const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

async function discoverHome() {
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

// The local name of the first child of a DAV:error body: the precondition Cyrus reports. It is
// one of Cyrus's fixed element names, never account data; it is still passed through redact.
function errorElementName(body) {
  const inner = el(body, 'error');
  if (inner === undefined) return '(no DAV:error element)';
  const m = inner.match(/<(?:[\w-]+:)?([\w-]+)[\s/>]/);
  return m ? redact(m[1]) : '(empty DAV:error)';
}

const unfold = ics => ics.replace(/\r?\n[ \t]/g, '');
const lines = ics => unfold(ics).split(/\r?\n/);
const paramOf = (line, name) => line.slice(0, line.indexOf(':')).match(new RegExp(`;${name}=("?)([^;:"]+)\\1`))?.[2];
const valueOf = line => line.slice(line.indexOf(':') + 1);

// RFC 5545 dur-time for the simple forms Cyrus emits (no weeks/days needed here, but read anyway).
function durationMs(v) {
  const m = v.match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return undefined;
  const [, sign, w, d, h, mi, s] = m;
  const ms = (((+(w ?? 0) * 7 + +(d ?? 0)) * 24 + +(h ?? 0)) * 60 + +(mi ?? 0)) * 60000 + +(s ?? 0) * 1000;
  return sign === '-' ? -ms : ms;
}

const { check, failures } = makeChecker();
const notReached = n => console.log(`NOT REACHED condition ${n}`);

let tempCalendarUrl = null;

try {
  const homeUrl = await discoverHome();

  const candidateUrl = new URL(`${SEGMENT}/`, homeUrl).href;
  const mk = await dav('MKCALENDAR', candidateUrl, {
    body: `<?xml version="1.0" encoding="utf-8"?>\n` +
      `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
      `<d:displayname>probe jscalendar-put temporary</d:displayname>` +
      `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
      `</d:prop></d:set></c:mkcalendar>`,
  });
  const mkOk = mk.status >= 200 && mk.status < 300;
  check('MKCALENDAR accepted (never falls back to a real calendar)', mkOk, `HTTP ${mk.status}`);
  if (!mkOk) throw new Error('stopping: no isolated collection to write into');
  tempCalendarUrl = candidateUrl;

  console.log(`Fixture: timed event, timeZone=${ZONE}, start=${START_LOCAL} (local), duration=${DURATION}`);
  console.log(`Expected: start ${WANT_START_WALL} local = ${basicUtc(WANT_START_MS)}, end ${WANT_END_WALL} local = ${basicUtc(WANT_END_MS)}`);
  console.log(`Offsets at start/end (Intl): ${offsetMsAt(WANT_START_MS) / 3600000}h / ${offsetMsAt(WANT_END_MS) / 3600000}h`);

  // --- 1. PUT the JSCalendar event ----------------------------------------------------------
  const jsevent = {
    '@type': 'Event',
    uid: UID,
    title: 'probe jscalendar-put fixture',
    start: START_LOCAL,
    timeZone: ZONE,
    duration: DURATION,
  };
  const resourceUrl = new URL(`${SEGMENT}.ics`, candidateUrl).href;
  const put = await dav('PUT', resourceUrl, {
    body: JSON.stringify(jsevent),
    headers: { 'Content-Type': 'application/event+json; charset=utf-8' },
  });
  const putOk = put.status >= 200 && put.status < 300;
  check(
    'condition 1: CalDAV PUT of an application/event+json JSCalendar event is accepted',
    putOk,
    putOk ? `HTTP ${put.status}` : `HTTP ${put.status}, DAV:error element: ${errorElementName(put.text)}`,
  );

  if (!putOk) {
    notReached(2); notReached(3); notReached(4);
  } else {
    // --- raw GET as text/calendar -------------------------------------------------------------
    const got = await dav('GET', resourceUrl, { headers: { Accept: 'text/calendar' } });
    console.log(`GET stored resource as text/calendar -> HTTP ${got.status}`);
    const raw = got.status === 200 ? got.text : '';

    // --- 2. one VTIMEZONE, TZID = the event's zone --------------------------------------------
    const vtz = [...raw.matchAll(/BEGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/g)].map(m => m[0]);
    const tzids = vtz.map(b => lines(b).find(l => l.startsWith('TZID:'))?.slice(5) ?? '(none)');
    check(
      'condition 2: exactly one VTIMEZONE, whose TZID is the event\'s zone',
      vtz.length === 1 && tzids[0] === ZONE,
      `VTIMEZONE count=${vtz.length} TZIDs=[${tzids.map(redact).join(', ')}]`,
    );

    // --- 3. TZUNTIL ---------------------------------------------------------------------------
    const block = vtz[0] ?? '';
    const tzuntil = lines(block).find(l => l.startsWith('TZUNTIL'));
    check(
      'condition 3: that VTIMEZONE carries a TZUNTIL',
      !!tzuntil,
      tzuntil
        ? `${redact(tzuntil)} (event end in UTC is ${basicUtc(WANT_END_MS)}; ${valueOf(tzuntil) === basicUtc(WANT_END_MS) ? 'equal' : 'differs'})`
        : (vtz.length ? 'no TZUNTIL line' : 'no VTIMEZONE to read'),
    );
    if (block) {
      const obs = lines(block).filter(l => /^BEGIN:(STANDARD|DAYLIGHT)$/.test(l)).map(l => l.slice(6));
      console.log(`  observances in the block: ${obs.length} [${obs.join(', ')}]`);
      const blockLines = lines(block);
      console.log('  stored VTIMEZONE, verbatim (timezone data; capped at 30 lines):');
      for (const l of blockLines.slice(0, 30)) console.log(`    ${redact(l)}`);
      if (blockLines.length > 30) console.log(`    ... ${blockLines.length - 30} more line(s)`);
    }

    // --- 4. DTSTART / end survive ------------------------------------------------------------
    const vevent = (raw.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/) ?? [''])[0];
    const ev = lines(vevent);
    const dtstart = ev.find(l => /^DTSTART[;:]/.test(l));
    const dtend = ev.find(l => /^DTEND[;:]/.test(l));
    const dur = ev.find(l => /^DURATION[;:]/.test(l));
    console.log(`  VEVENT timing lines: ${[dtstart, dtend, dur].filter(Boolean).map(redact).join(' | ') || '(none)'}`);

    const startZone = dtstart && paramOf(dtstart, 'TZID');
    const startWall = dtstart && valueOf(dtstart);
    check(
      'condition 4a: DTSTART keeps the intended zone and wall time',
      startZone === ZONE && startWall === WANT_START_WALL,
      `TZID=${redact(startZone ?? '(none)')} wall=${redact(startWall ?? '(none)')} want TZID=${ZONE} wall=${WANT_START_WALL}`,
    );

    let endMs, endForm;
    if (dtend) {
      const z = paramOf(dtend, 'TZID');
      const v = valueOf(dtend);
      endForm = `DTEND TZID=${redact(z ?? '(none)')} wall=${redact(v)}`;
      if (z === ZONE && /^\d{8}T\d{6}$/.test(v)) endMs = wallToUtcMs(v);
      else if (!z && /^\d{8}T\d{6}Z$/.test(v)) endMs = Date.parse(`${v.slice(0, 4)}-${v.slice(4, 6)}-${v.slice(6, 8)}T${v.slice(9, 11)}:${v.slice(11, 13)}:${v.slice(13, 15)}Z`);
    } else if (dur) {
      const d = durationMs(valueOf(dur));
      endForm = `DURATION ${redact(valueOf(dur))}`;
      if (d !== undefined && startZone === ZONE && startWall) endMs = wallToUtcMs(startWall) + d;
    } else {
      endForm = 'no DTEND and no DURATION';
    }
    check(
      'condition 4b: the end survives as the intended instant (DTEND or DURATION)',
      endMs === WANT_END_MS,
      `${endForm} -> ${endMs === undefined ? '(unreadable)' : `${basicUtc(endMs)} = ${utcMsToWall(endMs)} local`}; want ${basicUtc(WANT_END_MS)} = ${WANT_END_WALL} local`,
    );
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
} finally {
  // Guarded, so a network failure here cannot replace the error that brought us here or skip the
  // line naming what to remove by hand.
  try {
    if (tempCalendarUrl) {
      const del = await dav('DELETE', tempCalendarUrl);
      console.log(`Cleanup: DELETE temporary collection -> HTTP ${del.status}`);
      const after = await dav('PROPFIND', tempCalendarUrl, { body: PROPFIND(['d:resourcetype']), headers: { Depth: '0' } });
      const gone = after.status === 404 || after.status === 410;
      check('cleanup: the temporary collection is gone (PROPFIND no longer finds it)', gone, `DELETE ${del.status}, PROPFIND ${after.status}`);
      if (!gone) console.log(`  DELETE MANUALLY: the calendar collection named ${SEGMENT} in the calendar home`);
    }
  } catch (err) {
    console.log(`Cleanup FAILED: ${redact(err?.message ?? String(err))}`);
    console.log(`  DELETE MANUALLY: the calendar collection named ${SEGMENT} in the calendar home`);
    process.exit(1);
  }
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
