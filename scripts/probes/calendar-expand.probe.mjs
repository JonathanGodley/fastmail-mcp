// What this probe settles
// -----------------------
// `list_calendar_events` reported a recurring event at its ORIGINAL DTSTART and returned events
// outside the requested window (#64), because the read path never asked the server to expand
// recurrences. The fix rests on what Fastmail's CalDAV server returns for a `calendar-query` with
// `<C:expand>` inside a `time-range`, which cannot be read off RFC 4791:
//
// PASS 1, the shape (#64). Expanded occurrences arrive with RRULE stripped, inside the window, and
// several of them as multiple VEVENTs in ONE calendar-data blob per resource (a first-match parser
// would drop all but one).
//
// PASS 2, the first instance (#64). Cyrus's `expand_cb` sets a RECURRENCE-ID only on instances
// after the first, so a window holding a series' original DTSTART returns one block with no
// RECURRENCE-ID beside its siblings, and "the block without one is the master" is a false read.
// Each daily-or-coarser series that a plain query over DISCOVER_START..DISCOVER_END returns, and
// whose DTSTART resolves to an instant, is expanded over a window built from that DTSTART, up to
// MAX_SERIES of them.
//
// PASS 2 also records the FORM of each expanded RECURRENCE-ID and DTSTART line (#146): whether it
// carries `TZID=`, `VALUE=DATE`, a trailing `Z`, or none of those (`floating`), or is `absent`. It
// is tallied separately for a generated instance of a timed series whose master has a TZID, for an
// override instance (one whose RECURRENCE-ID matches an override block in the stored resource), for
// an instance of an all-day series, and for the block that lacks a RECURRENCE-ID in a series whose
// first instance stands, is overridden, or is excluded by an EXDATE. Instants are compared after resolving a TZID with Intl,
// so a stored `TZID=` value matches an expanded `Z` value naming the same instant.
//
// ETAGS (#146). For every resource an expand REPORT returns, its `getetag` is compared byte for
// byte (the raw XML text, and again after XML unescaping) with the `getetag` a Depth 1 PROPFIND on
// the collection returns for the same href, and for pass 2's series with a `calendar-multiget`
// REPORT that asks for `getetag` only. The same href is compared again on each expand that
// returns it, so the comparison count exceeds the resource count, which is printed beside it.
//
// READ-ONLY: every request goes through `dav()`, which refuses any method but PROPFIND and REPORT.
// Redirects are not followed, and credentials go only to ROOT's origin.
//
// OUTPUT IS SAFE TO QUOTE PUBLICLY. Hrefs, UIDs, summaries, calendar names, zone names, dates,
// times and etags are held in memory only. It prints counts, booleans, the form tokens in FORMS
// (anything else prints as `<other>`, and is counted), PASS/FAIL lines, the fixed `REFUSED` lines
// of calendar-component-set.probe.mjs plus `REFUSED  <step>: method not allowed`, and for an
// unexpected error its class and code only.
//
// LIMITS: a TZID that Intl cannot resolve leaves that value's instant unknown (counted); a floating
// value is read as UTC, as the server's own time-range matching does
// (calendar-window-frames.probe.mjs); at most MAX_SERIES series are expanded.
//
// Run: python scripts/probes/run-probe.py calendar-expand.probe.mjs [startDate endDate]

import { el, elAll, resolveHref, sameResource, directChildren } from './calendar-component-set.probe.mjs';
import { makeChecker } from './probelib.mjs';

const ROOT = 'https://caldav.fastmail.com/dav/';
const ALLOWED_METHODS = new Set(['PROPFIND', 'REPORT']);

// Pass 1's window: far enough out that anything in it is there because a recurrence put it there.
const START = process.argv[2] || '2027-03-01T00:00:00Z';
const END = process.argv[3] || '2027-03-10T00:00:00Z';
// Pass 2 finds series by a plain (unexpanded) query over this span.
const DISCOVER_START = '2000-01-01T00:00:00Z';
const DISCOVER_END = '2035-01-01T00:00:00Z';
const MAX_SERIES = 60;
const DAY = 86400000;
// Days per occurrence, per FREQ. Sub-daily series are skipped, never expanded: expanding a
// FREQ=MINUTELY series over a useful span is the denial of service the window clamp exists for.
const DAYS_PER = { YEARLY: 366, MONTHLY: 31, WEEKLY: 7, DAILY: 1 };
const MIN_OCCURRENCES = 5;
const MAX_OCCURRENCES = 400;

export const FORMS = new Set(['TZID', 'VALUE=DATE', 'Z', 'floating', 'absent']);
let otherForms = 0;
const allowForm = f => {
  if (FORMS.has(f)) return f;
  otherForms++;
  return '<other>';
};

// ---- iCalendar ----

const unfold = s => s.replace(/\r?\n[ \t]/g, '');
const VEVENT_RE = /BEGIN:VEVENT[\s\S]*?END:VEVENT/g;

/** Every content line named `name` in a block: { params: {KEY: value}, value }. */
export function lines(block, name) {
  const out = [];
  for (const raw of block.split(/\r?\n/)) {
    const colon = raw.search(/:(?=(?:[^"]*"[^"]*")*[^"]*$)/);
    if (colon < 0) continue;
    const [head, ...paramParts] = raw.slice(0, colon).split(';');
    if (head.toUpperCase() !== name) continue;
    const params = {};
    for (const p of paramParts) {
      const eq = p.indexOf('=');
      if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
    }
    out.push({ params, value: raw.slice(colon + 1).replace(/\r$/, '') });
  }
  return out;
}

/** The form token of a date or date-time line, or `absent`. Mixed shapes return a non-token. */
export function formOf(line) {
  if (!line) return 'absent';
  const isDate = /^\d{8}$/.test(line.value);
  const isUtc = /^\d{8}T\d{6}Z$/.test(line.value);
  const isLocal = /^\d{8}T\d{6}$/.test(line.value);
  if (line.params.VALUE?.toUpperCase() === 'DATE') return isDate && !line.params.TZID ? 'VALUE=DATE' : 'mixed';
  if (line.params.TZID) return isLocal ? 'TZID' : 'mixed';
  if (isUtc) return 'Z';
  if (isLocal) return 'floating';
  return 'mixed';
}

function zoneOffset(tz, utcMs) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', {
    timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(new Date(utcMs)).map(p => [p.type, p.value]));
  return Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute, +parts.second) - utcMs;
}

let unresolvedZones = 0;

/** The instant one date or date-time value names, or null. */
export function instantOf(value, tzid) {
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/.exec(value);
  if (!m) return null;
  const wall = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0));
  if (!m[4] || m[7] || !tzid) return wall;
  try {
    const guess = wall - zoneOffset(tzid, wall);
    return wall - zoneOffset(tzid, guess);
  } catch {
    unresolvedZones++;
    return null;
  }
}

const lineInstant = l => (l ? instantOf(l.value, l.params.TZID) : null);

function describeBlob(data) {
  return (unfold(data).match(VEVENT_RE) || []).map(b => ({
    dtstart: lines(b, 'DTSTART')[0],
    recurrenceId: lines(b, 'RECURRENCE-ID')[0],
    rrule: lines(b, 'RRULE')[0],
    exdates: lines(b, 'EXDATE').flatMap(l => l.value.split(',').map(v => instantOf(v, l.params.TZID))),
  }));
}

const utcStamp = ms => new Date(ms).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

// ---- WebDAV ----

const unescapeXml = s => s
  .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
  .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16))).replace(/&amp;/g, '&');

/** Each `<response>`: its resolved href, and from its 200 propstats the raw getetag and data. */
function parseMultistatus(xml, resolve) {
  return elAll(xml, 'response').map(r => {
    const out = { url: resolve(el(r, 'href') ?? ''), etagRaw: undefined, data: undefined, types: [] };
    for (const stat of elAll(r, 'propstat')) {
      if (!/\s200\s/.test(` ${(el(stat, 'status') ?? '').trim()} `)) continue;
      const prop = el(stat, 'prop') ?? '';
      const etag = el(prop, 'getetag');
      if (etag !== undefined) out.etagRaw = etag;
      const data = el(prop, 'calendar-data');
      if (data !== undefined) out.data = unescapeXml(data);
      const rt = el(prop, 'resourcetype');
      if (rt !== undefined) out.types = directChildren(rt);
    }
    return out;
  });
}

class Refused extends Error {}

function refuse(line) {
  console.log(`REFUSED  ${line}`);
  throw new Refused();
}

function calendarQuery(start, end, expand) {
  const s = utcStamp(Date.parse(start));
  const e = utcStamp(Date.parse(end));
  return '<?xml version="1.0" encoding="utf-8"?>\n'
    + '<c:calendar-query xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop><d:getetag/>'
    + (expand ? `<c:calendar-data><c:expand start="${s}" end="${e}"/></c:calendar-data>` : '<c:calendar-data/>')
    + '</d:prop><c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT">'
    + `<c:time-range start="${s}" end="${e}"/></c:comp-filter></c:comp-filter></c:filter></c:calendar-query>`;
}

async function main() {
  const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
  const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;
  if (!USERNAME || !PASSWORD) {
    console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
    console.error('      Run through: python scripts/probes/run-probe.py calendar-expand.probe.mjs');
    process.exit(1);
  }
  const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');
  const resolve = href => {
    const r = resolveHref(href);
    if (r.refused) refuse(r.refused);
    return r.url;
  };
  const methodsSent = new Set();

  async function dav(step, method, url, depth, body) {
    if (!ALLOWED_METHODS.has(method)) refuse(`${step}: method not allowed`);
    methodsSent.add(method);
    const res = await fetch(url, {
      method,
      redirect: 'manual',
      headers: { Authorization: AUTH, 'Content-Type': 'text/xml;charset=UTF-8', Depth: depth },
      body,
    });
    if (res.status !== 207) refuse(`${step}: HTTP ${res.status}`);
    return res.text();
  }
  const propfind = (step, url, props, depth) => dav(step, 'PROPFIND', url, depth,
    '<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">'
    + `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`);
  const query = async (step, calUrl, start, end, expand) =>
    parseMultistatus(await dav(step, 'REPORT', calUrl, '1', calendarQuery(start, end, expand)), resolve);

  const { check, failures } = makeChecker();

  // ---- discovery ----
  const rootXml = await propfind('current-user-principal', new URL(ROOT), ['d:current-user-principal'], '0');
  const principal = el(el(rootXml, 'current-user-principal') ?? '', 'href');
  if (!principal) refuse('current-user-principal: absent');
  const prinXml = await propfind('calendar-home-set', resolve(principal), ['c:calendar-home-set'], '0');
  const homeHref = el(el(prinXml, 'calendar-home-set') ?? '', 'href');
  if (!homeHref) refuse('calendar-home-set: absent');
  const homeUrl = resolve(homeHref);
  const listing = parseMultistatus(await propfind('calendar home listing', homeUrl, ['d:resourcetype'], '1'), resolve);
  const calendars = listing.filter(e => !sameResource(e.url, homeUrl) && e.types.includes('calendar')).map(e => e.url);
  console.log(`calendar collections: ${calendars.length}`);
  check('calendar discovery returned at least one calendar', calendars.length > 0);

  // Etags as a Depth 1 PROPFIND reports them, per collection, fetched once.
  const propfindEtags = new Map();
  async function etagsFor(calUrl) {
    const key = calUrl.href;
    if (!propfindEtags.has(key)) {
      const rows = parseMultistatus(await propfind('etag listing', calUrl, ['d:getetag'], '1'), resolve);
      propfindEtags.set(key, rows.filter(r => !sameResource(r.url, calUrl)));
    }
    return propfindEtags.get(key);
  }
  const etag = { compared: 0, nonEmpty: 0, distinct: new Set(), resources: new Set(), rawEqual: 0,decodedEqual: 0, missingInPropfind: 0, missingInReport: 0 };
  async function compareEtags(calUrl, reportRows) {
    const listed = await etagsFor(calUrl);
    for (const r of reportRows) {
      if (r.etagRaw === undefined) { etag.missingInReport++; continue; }
      const p = listed.find(x => sameResource(x.url, r.url));
      if (!p || p.etagRaw === undefined) { etag.missingInPropfind++; continue; }
      etag.resources.add(r.url.href);
      etag.compared++;
      if (r.etagRaw.trim()) etag.nonEmpty++;
      etag.distinct.add(r.etagRaw);
      if (p.etagRaw === r.etagRaw) etag.rawEqual++;
      if (unescapeXml(p.etagRaw).trim() === unescapeXml(r.etagRaw).trim()) etag.decodedEqual++;
    }
  }

  // ---- pass 1: shape ----
  console.log('\n=== Pass 1: expanded shape in a fixed window ===');
  let p1Blobs = 0;
  let p1Multi = 0;
  let p1Rrule = 0;
  let p1Outside = 0;
  for (const [i, cal] of calendars.entries()) {
    const plain = await query('plain query', cal, START, END, false);
    const expanded = await query('expanded query', cal, START, END, true);
    console.log(`#${i + 1}  objects without expand: ${plain.length}  with expand: ${expanded.length}`);
    await compareEtags(cal, expanded);
    for (const obj of expanded) {
      const parts = describeBlob(obj.data || '');
      if (parts.length === 0) continue;
      p1Blobs++;
      if (parts.length > 1) p1Multi++;
      if (parts.some(p => p.rrule)) p1Rrule++;
      if (parts.some(p => {
        const t = lineInstant(p.dtstart);
        return t !== null && (t < Date.parse(START) || t >= Date.parse(END));
      })) p1Outside++;
    }
  }
  console.log(`expanded blobs: ${p1Blobs}  with more than one VEVENT: ${p1Multi}  carrying RRULE: ${p1Rrule}  with a DTSTART outside the window: ${p1Outside}`);
  check('expanded occurrences carry no RRULE (the server applied the recurrence)', p1Rrule === 0);
  check('every expanded occurrence starts inside the requested window', p1Outside === 0);

  // ---- pass 2: series, first instance, forms ----
  console.log('\n=== Pass 2: every series, expanded from its own DTSTART ===');
  const series = [];
  let skippedSubDaily = 0;
  for (const cal of calendars) {
    for (const obj of await query('discovery query', cal, DISCOVER_START, DISCOVER_END, false)) {
      const blocks = describeBlob(obj.data || '');
      const master = blocks.find(b => b.rrule && !b.recurrenceId);
      if (!master) continue;
      const rule = master.rrule.value;
      const freq = /FREQ=([A-Z]+)/.exec(rule)?.[1];
      const interval = Math.max(1, +(/INTERVAL=(\d+)/.exec(rule)?.[1] ?? 1));
      if (!DAYS_PER[freq]) { skippedSubDaily++; continue; }
      const startMs = lineInstant(master.dtstart);
      if (startMs === null) continue;
      series.push({
        cal, url: obj.url, master, startMs, daysPer: DAYS_PER[freq] * interval,
        overrides: blocks.filter(b => b.recurrenceId).map(b => lineInstant(b.recurrenceId)),
      });
    }
  }
  const shapeOf = s => formOf(s.master.dtstart);
  console.log(`series found: ${series.length}  sub-daily skipped: ${skippedSubDaily}  expanded: ${Math.min(series.length, MAX_SERIES)}`);

  const tally = { a: {}, b: {}, c: {}, d: {}, first: {}, firstOverride: {} };
  const firstOverridden = { series: 0, bareCounts: {} };
  const bump = (bucket, rid, dts) => {
    const k = `RECURRENCE-ID=${allowForm(rid)} DTSTART=${allowForm(dts)}`;
    bucket[k] = (bucket[k] ?? 0) + 1;
  };
  const shapes = {};
  let notMatched = 0;
  let firstSeries = 0;
  let firstOk = 0;
  let anyRrule = 0;
  let multi = 0;
  const excluded = { series: 0, blocks: 0, withBlocks: 0, bareCounts: {},bareIsEarliest: 0, bareIsOverride: 0, bareIsMasterStart: 0 };
  const multigetRows = new Map();

  for (const s of series.slice(0, MAX_SERIES)) {
    const shape = allowForm(shapeOf(s));
    shapes[shape] = (shapes[shape] ?? 0) + 1;
    const winStart = s.startMs - DAY;
    const lastOverride = Math.max(0, ...s.overrides.filter(t => t !== null));
    const winEnd = Math.min(
      Math.max(s.startMs + MIN_OCCURRENCES * s.daysPer * DAY, lastOverride + 2 * DAY),
      s.startMs + MAX_OCCURRENCES * s.daysPer * DAY,
    );
    const rows = await query('series expand', s.cal, new Date(winStart).toISOString(), new Date(winEnd).toISOString(), true);
    await compareEtags(s.cal, rows);
    const row = rows.find(r => sameResource(r.url, s.url));
    if (!row) { notMatched++; continue; }
    if (!multigetRows.has(s.cal.href)) multigetRows.set(s.cal.href, { cal: s.cal, rows: [] });
    multigetRows.get(s.cal.href).rows.push(row);

    const parts = describeBlob(row.data || '');
    if (parts.length > 1) multi++;
    if (parts.some(p => p.rrule)) anyRrule++;
    const isOverride = p => {
      const t = lineInstant(p.recurrenceId);
      return t !== null && s.overrides.includes(t);
    };
    for (const p of parts) {
      if (!p.recurrenceId) continue;
      if (isOverride(p)) bump(tally.b, formOf(p.recurrenceId), formOf(p.dtstart));
      else if (shapeOf(s) === 'TZID') bump(tally.a, formOf(p.recurrenceId), formOf(p.dtstart));
      else if (shapeOf(s) === 'VALUE=DATE') bump(tally.c, formOf(p.recurrenceId), formOf(p.dtstart));
    }
    const bare = parts.filter(p => !p.recurrenceId);
    const firstExcluded = s.master.exdates.includes(s.startMs);
    if (firstExcluded) {
      excluded.series++;
      excluded.blocks += parts.length;
      if (parts.length > 0) excluded.withBlocks++;
      excluded.bareCounts[bare.length] = (excluded.bareCounts[bare.length] ?? 0) + 1;
      const starts = parts.map(p => lineInstant(p.dtstart)).filter(t => t !== null);
      for (const b of bare) {
        bump(tally.d, 'absent', formOf(b.dtstart));
        const t = lineInstant(b.dtstart);
        if (t !== null && t === Math.min(...starts)) excluded.bareIsEarliest++;
        if (t !== null && t === s.startMs) excluded.bareIsMasterStart++;
        if (t !== null && s.overrides.includes(t)) excluded.bareIsOverride++;
      }
    } else if (s.overrides.includes(s.startMs)) {
      firstOverridden.series++;
      firstOverridden.bareCounts[bare.length] = (firstOverridden.bareCounts[bare.length] ?? 0) + 1;
      for (const b of bare) bump(tally.firstOverride, 'absent', formOf(b.dtstart));
    } else {
      firstSeries++;
      for (const b of bare) bump(tally.first, 'absent', formOf(b.dtstart));
      if (bare.length === 1 && lineInstant(bare[0].dtstart) === s.startMs) firstOk++;
    }
  }

  console.log(`master DTSTART form, per series: ${Object.entries(shapes).map(([k, v]) => `${k}=${v}`).join('  ') || 'none'}`);
  console.log(`series not found in their own expand window: ${notMatched}  with more than one VEVENT: ${multi}  carrying RRULE: ${anyRrule}`);
  const show = (label, bucket) => {
    const entries = Object.entries(bucket);
    console.log(`${label}: ${entries.length ? entries.map(([k, v]) => `${k} x${v}`).join('; ') : 'no such block'}`);
  };
  show('(a) generated instance, timed series with a TZID master', tally.a);
  show('(b) override instance', tally.b);
  show('(c) generated instance, all-day series', tally.c);
  show('(d) block lacking RECURRENCE-ID, series whose first instance is excluded', tally.d);
  show('block lacking RECURRENCE-ID, series whose first instance stands', tally.first);
  show('block lacking RECURRENCE-ID, series whose first instance is overridden', tally.firstOverride);
  console.log(`series with the first instance overridden: ${firstOverridden.series}`
    + `  bare blocks per series: ${Object.entries(firstOverridden.bareCounts).map(([k, v]) => `${k}x${v}`).join(' ') || 'n/a'}`);
  console.log(`(d) series with the first instance excluded: ${excluded.series}`
    + `  blocks emitted: ${excluded.blocks}  series emitting at least one block: ${excluded.withBlocks}`
    + `  bare blocks per series: ${Object.entries(excluded.bareCounts).map(([k, v]) => `${k}x${v}`).join(' ') || 'n/a'}`
    + `  bare block is the earliest emitted: ${excluded.bareIsEarliest}`
    + `  is the excluded DTSTART: ${excluded.bareIsMasterStart}`
    + `  is an override: ${excluded.bareIsOverride}`);
  console.log(`series with the first instance neither excluded nor overridden: ${firstSeries}`
    + `  with exactly one bare block, at the series DTSTART: ${firstOk}`);
  check('every series was returned by an expand over its own window', notMatched === 0);
  check('no expanded block of any series carries an RRULE', anyRrule === 0);
  check('a series whose first instance stands returns exactly one block without a RECURRENCE-ID, the first instance',
    firstSeries > 0 && firstOk === firstSeries, `${firstOk} of ${firstSeries}`);

  // ---- etags ----
  console.log('\n=== Etags: expand REPORT against non-expanded reads ===');
  const mg = { compared: 0, rawEqual: 0, decodedEqual: 0, missing: 0 };
  for (const { cal, rows } of multigetRows.values()) {
    const body = '<?xml version="1.0" encoding="utf-8"?>\n'
      + '<c:calendar-multiget xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop><d:getetag/></d:prop>'
      + rows.map(r => `<d:href>${r.url.pathname.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</d:href>`).join('') + '</c:calendar-multiget>';
    const got = parseMultistatus(await dav('multiget', 'REPORT', cal, '1', body), resolve);
    for (const r of rows) {
      const g = got.find(x => sameResource(x.url, r.url));
      if (!g || g.etagRaw === undefined || r.etagRaw === undefined) { mg.missing++; continue; }
      mg.compared++;
      if (g.etagRaw === r.etagRaw) mg.rawEqual++;
      if (unescapeXml(g.etagRaw).trim() === unescapeXml(r.etagRaw).trim()) mg.decodedEqual++;
    }
  }
  console.log(`comparisons: ${etag.compared} over ${etag.resources.size} distinct resources;`
    + ` every compared etag non-empty: ${etag.nonEmpty === etag.compared}; more than one distinct etag seen: ${etag.distinct.size > 1}`);
  console.log(`vs Depth 1 PROPFIND: compared ${etag.compared}  raw-equal ${etag.rawEqual}  equal after unescaping ${etag.decodedEqual}`
    + `  not-equal ${etag.compared - etag.rawEqual}  no etag in REPORT ${etag.missingInReport}  href absent from PROPFIND ${etag.missingInPropfind}`);
  console.log(`vs calendar-multiget (series only): compared ${mg.compared}  raw-equal ${mg.rawEqual}`
    + `  equal after unescaping ${mg.decodedEqual}  not-equal ${mg.compared - mg.rawEqual}  missing ${mg.missing}`);
  check('every expand-REPORT etag equals the PROPFIND etag byte for byte',
    etag.compared > 0 && etag.rawEqual === etag.compared && etag.missingInPropfind === 0);
  check('the compared etags are non-empty and not one constant value',
    etag.nonEmpty === etag.compared && etag.distinct.size > 1);
  check('every series expand-REPORT etag equals the multiget etag byte for byte',
    mg.compared > 0 && mg.rawEqual === mg.compared && mg.missing === 0);

  console.log(`\nTZID values Intl could not resolve: ${unresolvedZones}  forms outside the allowlist, printed as <other>: ${otherForms}`);
  console.log(`HTTP methods sent: ${[...methodsSent].filter(m => ALLOWED_METHODS.has(m)).sort().join(',')}`);
  console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
  if (failures() > 0) process.exitCode = 1;
}

main().catch(e => {
  if (!(e instanceof Refused)) console.log(`FAIL  unexpected error: ${e?.name ?? 'Error'}${e?.code ? ` ${e.code}` : ''}`);
  process.exit(1);
});
