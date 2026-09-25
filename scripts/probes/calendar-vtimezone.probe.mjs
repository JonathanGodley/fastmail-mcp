// What this probe settles
// -----------------------
// #166: create_calendar_event and update_calendar_event generate a VTIMEZONE component
// (RFC 5545 §3.6.5) for every zone a written DTSTART/DTEND references, since Fastmail's own
// server never supplies or repairs one on this server's CalDAV write path (measured in
// docs/conventions.md's "VTIMEZONE residual" and docs/fastmail-action-availability.md). The
// generator itself (src/vtimezone.ts) is unit-tested against Node's own ICU data;
// what a unit test cannot prove is that the WIRED-UP tool actually puts that generated block on
// the wire, in a resource this account's real CalDAV server accepts and stores unchanged. That
// is this probe.
//
// It creates ONE timed event through the BUILT server (dist/index.js, via
// scripts/mcp-harness.mjs) in Australia/Sydney — a zone with DST, so the block is a real
// STANDARD/DAYLIGHT observance rather than the single-offset case — sitting directly ON Sydney's
// own October 2026 spring-forward transition (clocks jump from 02:00 AEST straight to 03:00
// AEDT), so the event's own short span crosses it and carries exactly two observances to check.
// (The generator's transition-finding and its DAYLIGHT-vs-STANDARD classification across a
// transition are what src/vtimezone.test.ts proves; this probe is about the wire, not the
// arithmetic.) It then fetches the stored resource back RAW over CalDAV — bare `fetch`, no
// tsdav, nothing this server's own parser touches — and checks the bytes themselves, computing
// its own expected offsets independently via Intl rather than by importing src/vtimezone.ts, so
// a shared bug in both would not agree with itself.
//
// The fixture goes into a temporary collection minted by MKCALENDAR, the same provenance
// discipline calendar-window-frames.probe.mjs uses: this runs against a live personal account,
// so nothing is ever written into a real calendar. If MKCALENDAR fails the probe stops rather
// than falling back to one. The whole collection is removed in a `finally`, which takes the one
// fixture with it in a single request. No participants, so nothing is mailed.
//
// Output is PASS/FAIL and counts/offsets only: no collection URL, UID, event title or other
// account-derived value is printed, so a run can be quoted verbatim into a public issue or
// commit.
//
// Run: python scripts/probes/run-probe.py calendar-vtimezone.probe.mjs
// Requires FASTMAIL_API_TOKEN plus FASTMAIL_CALDAV_USERNAME/PASSWORD; the launcher injects all
// three from the local MCP client config. Build first: npm run build.

import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-vtimezone.probe.mjs');
  process.exit(1);
}

const { check, failures } = makeChecker();

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');
// Every CalDAV href and every account-derived value is redacted before it is ever printed or
// compared into a message — this account's own address included.
const redact = s => String(s).split(USERNAME).join('<account>');

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

// The zone Intl reports for Australia/Sydney at `utcMs`, as milliseconds — computed
// independently of src/vtimezone.ts (and of src/coerce.ts's zoneOffsetMsAt), so this probe is
// not just re-checking the generator against itself.
function offsetMsAt(utcMs) {
  const p = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: 'Australia/Sydney', hour12: false,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(new Date(utcMs)).map(p => [p.type, p.value]),
  );
  return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second) - utcMs;
}

function offsetToken(offsetMs) {
  const sign = offsetMs < 0 ? '-' : '+';
  const totalMin = Math.round(Math.abs(offsetMs) / 60000);
  const hh = String(Math.floor(totalMin / 60)).padStart(2, '0');
  const mm = String(totalMin % 60).padStart(2, '0');
  return `${sign}${hh}${mm}`;
}

async function discoverHome() {
  const PROPFIND = props =>
    `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
    `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;
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

// Every STANDARD/DAYLIGHT sub-component in a VTIMEZONE block, unfolded first (RFC 5545 §3.1
// continuation lines start with a space or tab).
function observances(block) {
  const unfolded = block.replace(/\r\n[ \t]/g, '').split(/\r?\n/);
  const out = [];
  let cur = null;
  for (const line of unfolded) {
    if (line === 'BEGIN:STANDARD' || line === 'BEGIN:DAYLIGHT') cur = { kind: line.slice(6) };
    else if (cur && line.startsWith('TZOFFSETFROM:')) cur.from = line.slice(13);
    else if (cur && line.startsWith('TZOFFSETTO:')) cur.to = line.slice(11);
    else if (cur && (line === 'END:STANDARD' || line === 'END:DAYLIGHT')) { out.push(cur); cur = null; }
  }
  return out;
}

let tempCalendarUrl = null;

try {
  const client = createClient({ env: process.env });
  await client.init();
  try {
    const homeUrl = await discoverHome();

    // --- mint a temporary, isolated collection ---------------------------------------
    const stamp = Date.now();
    const candidateUrl = new URL(`probe-166-${stamp}/`, homeUrl).href;
    const mk = await dav('MKCALENDAR', candidateUrl, {
      body: `<?xml version="1.0" encoding="utf-8"?>\n` +
        `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
        `<d:displayname>probe-166 temporary</d:displayname>` +
        `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
        `</d:prop></d:set></c:mkcalendar>`,
    });
    check('MKCALENDAR accepted (never falls back to a real calendar)', mk.status >= 200 && mk.status < 300, `HTTP ${mk.status} ${mk.statusText}`);
    if (!(mk.status >= 200 && mk.status < 300)) throw new Error('stopping: no isolated collection to write into');
    tempCalendarUrl = candidateUrl;

    // --- create one timed event through the built server ------------------------------
    // Straddling Sydney's own 2026-10-04 spring-forward: 02:00 AEST jumps straight to 03:00
    // AEDT, so a 01:00-04:00 local span sits on both sides and this short event carries exactly
    // two observances — one STANDARD (pre-transition), one DAYLIGHT (post-transition).
    const createRes = await client.call('create_calendar_event', {
      calendarId: candidateUrl,
      title: 'probe-166 fixture',
      start: '2026-10-04T01:00:00',
      end: '2026-10-04T04:00:00',
      timeZone: 'Australia/Sydney',
    });
    const createBody = text(createRes);
    // The id never contains '.' or whitespace (`${Date.now()}-${random}@fastmail-mcp`); the
    // response sentence ends it with a literal period, which a bare \S+ would swallow. Only
    // whether an id was parsed is ever printed below — the response text itself carries the
    // event's UID and is never put on stdout.
    const eventId = /Event ID: ([^\s.]+)\.?/.exec(createBody)?.[1];
    check('create_calendar_event returned an event id', !!eventId, eventId ? 'an id was parsed' : 'no id was parsed');
    if (!eventId) throw new Error('stopping: no event id to fetch back');

    // --- fetch the resource back RAW over CalDAV, exactly as written ------------------
    const resourceUrl = new URL(`${eventId}.ics`, candidateUrl).href;
    const got = await dav('GET', resourceUrl);
    check('the created resource is fetchable raw over CalDAV', got.status === 200, `HTTP ${got.status}`);
    const raw = got.text;

    const vtzBlocks = [...raw.matchAll(/BEGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/g)].map(m => m[0]);
    check('exactly one VTIMEZONE block is present', vtzBlocks.length === 1, `found ${vtzBlocks.length}`);
    const block = vtzBlocks[0] ?? '';

    check('the block carries TZID:Australia/Sydney', /^TZID:Australia\/Sydney\r?$/m.test(block));

    // Scoped to the VEVENT, not the whole resource: a STANDARD/DAYLIGHT sub-component inside
    // the VTIMEZONE block above it carries its own bare `DTSTART:<onset>` line (RFC 5545
    // §3.6.5), which a whole-file search would find first and misread as the event's own.
    const veventBlock = (raw.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/) ?? [''])[0];
    check('a VEVENT block is present to read DTSTART/DTEND from', veventBlock.length > 0);
    const dtstartLine = (veventBlock.match(/^DTSTART[;:].*$/m) ?? [''])[0];
    const dtendLine = (veventBlock.match(/^DTEND[;:].*$/m) ?? [''])[0];
    const wallClockOf = line => /:(\d{8}T\d{6})/.exec(line)?.[1];
    const startWall = wallClockOf(dtstartLine);
    const endWall = wallClockOf(dtendLine);
    check('DTSTART/DTEND are still zoned wall clocks, not rewritten', !!startWall && !!endWall, `DTSTART=${dtstartLine.split(':')[0]} DTEND=${dtendLine.split(':')[0]}`);

    // The instant each wall clock names, via the same offsetMsAt used above. Both wall clocks
    // sit deliberately close to the transition, so a single `naive - offsetMsAt(naive)` can read
    // the offset off the wrong side of it; a second pass, off the first pass's own corrected
    // instant, converges (neither wall clock falls in the skipped 02:00-03:00 local gap itself).
    const wallToUtcMs = wall => {
      const y = +wall.slice(0, 4), mo = +wall.slice(4, 6), d = +wall.slice(6, 8);
      const h = +wall.slice(9, 11), mi = +wall.slice(11, 13), s = +wall.slice(13, 15);
      const naive = Date.UTC(y, mo - 1, d, h, mi, s);
      const firstPass = naive - offsetMsAt(naive);
      return naive - offsetMsAt(firstPass);
    };
    const startMs = startWall ? wallToUtcMs(startWall) : NaN;
    const endMs = endWall ? wallToUtcMs(endWall) : NaN;

    const obs = observances(block);
    check('the block carries exactly two observances (span crosses the October transition)', obs.length === 2, `found ${obs.length}`);
    const standardObs = obs.find(o => o.kind === 'STANDARD');
    const daylightObs = obs.find(o => o.kind === 'DAYLIGHT');
    check('one observance is STANDARD and the other DAYLIGHT', !!standardObs && !!daylightObs, `kinds=${obs.map(o => o.kind).join(',')}`);
    const expectedOffset = offsetToken(offsetMsAt(startMs));
    const expectedOffsetAtEnd = offsetToken(offsetMsAt(endMs));
    check(
      'independently-computed offsets at DTSTART and DTEND differ (the span crosses the transition)',
      expectedOffset !== expectedOffsetAtEnd,
      `start=${expectedOffset} end=${expectedOffsetAtEnd}`,
    );
    check(
      "the STANDARD observance's TZOFFSETTO matches Intl's independently-computed pre-transition offset",
      standardObs?.to === expectedOffset,
      `block=${standardObs?.to} intl=${expectedOffset}`,
    );
    check(
      "the DAYLIGHT observance's TZOFFSETTO matches Intl's independently-computed post-transition offset",
      daylightObs?.to === expectedOffsetAtEnd,
      `block=${daylightObs?.to} intl=${expectedOffsetAtEnd}`,
    );

    const tzuntilLine = (block.match(/^TZUNTIL:(\d{8}T\d{6}Z)$/m) ?? [])[1];
    const expectedTzuntil = Number.isFinite(endMs)
      ? new Date(endMs).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '')
      : undefined;
    check('TZUNTIL is present and equals the event\'s own DTEND, in UTC', !!tzuntilLine && tzuntilLine === expectedTzuntil, `block=${tzuntilLine} expected=${expectedTzuntil}`);
  } finally {
    client.close();
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
} finally {
  if (tempCalendarUrl) {
    const del = await dav('DELETE', tempCalendarUrl);
    console.log(`Cleanup: DELETE temporary collection -> HTTP ${del.status} ${del.statusText}`);
  }
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
