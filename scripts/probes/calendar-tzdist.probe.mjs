// What this probe settles
// -----------------------
// #166: every timed event Fastmail's own client authors is stored with a `VTIMEZONE` (added by
// Cyrus's JMAP-to-iCalendar converter; measured in docs/fastmail-action-availability.md), so
// matching that shape means embedding one. Where would the block come from?
//
// The cheap source is RFC 7808 timezone data distribution: ask the server for a zone by name
// and get its `VTIMEZONE` back, optionally truncated to a span. The expensive alternative is
// bundling or generating daylight-saving rules ourselves. Cyrus implements tzdist, but the
// service is gated on a per-deployment config switch that no source reading can settle, so
// this probe measures it; its FAIL is as useful an answer as its PASS.
//
// Three conditions, reported separately so a partial result is legible:
//
//   1. The service answers at all — discovered, not guessed (see "Discovery" below).
//   2. A named zone returns a parseable `VTIMEZONE` carrying the `TZID` asked for. Two zones
//      are asked for: one the account plausibly uses and one clearly foreign.
//   3. Truncation is honoured. A zone bounded by `start`/`end` must come back actually
//      truncated AND carrying the `TZUNTIL` the server adds for the `end` bound. A service
//      that IGNORES truncation cannot reproduce that stored shape, so it FAILS here.
//
// Discovery. The base URL is asked for, never assumed. Four routes, all read off the Cyrus
// source (imap/http_tzdist.c, imap/http_caldav.c, imap/httpd.c) rather than the RFC alone:
//
//   A. OPTIONS on the calendar home — the `DAV:` response header carries the
//      `calendar-no-timezone` feature token if and only if the tzdist namespace is enabled.
//      Says the service exists; does not say where.
//   B. PROPFIND for `CALDAV:timezone-service-set` (RFC 7809) on the calendar home — the
//      server names its own tzdist prefix in an href. Authoritative, and account-scoped.
//   C. `GET /.well-known/timezone` (RFC 7808) with redirects unfollowed — the 301's
//      `Location` is the prefix.
//   D. `GET /.well-known/` — Cyrus serves an HTML index of its enabled well-known URLs.
//
// Route B's request also carries the other RFC 7809 property, reported but not gated (see there).
//
// If no route names a base, the run still measures Cyrus's compiled-in default prefix
// (`/tzdist`) so the result can distinguish "not advertised and not there" from "not
// advertised but serving". That measurement never turns condition 1 into a pass on its own:
// condition 1 is the service ANSWERING, and the discovery result is reported beside it.
//
// Reading a 404 correctly. The host sits behind an edge proxy that routes only some path
// prefixes to the CalDAV backend, so a proxy 404 says nothing about the service. Condition 1
// names the tier that answered and asks again under the DAV root, which reaches the backend.
// Routes A and B settle it either way: both are served by the backend and emitted only when
// the timezone namespace is enabled.
//
// READ-ONLY: it creates and deletes nothing.
//
// OUTPUT DISCIPLINE. PASS/FAIL, counts and zone names only; account-derived strings and server
// error text (which echoes the request on some failures) are redacted. ONE truncated
// `VTIMEZONE` block (timezone data, not account data) is printed, capped at 25 lines, as the
// evidence for condition 3.
//
// Raw CalDAV/HTTP over bare `fetch`, not the built server and not tsdav — the question is what
// the platform serves, with none of our parsing in the way.
//
// Run: python scripts/probes/run-probe.py calendar-tzdist.probe.mjs

import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-tzdist.probe.mjs');
  process.exit(1);
}

const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');
const ROOT = 'https://caldav.fastmail.com/dav/';
const ORIGIN = new URL(ROOT).origin;

// Cyrus's compiled-in tzdist prefix (imap/http_tzdist.c, `namespace_tzdist`). Used ONLY as the
// last-ditch measurement when every discovery route comes back empty, and labelled as such
// wherever it is used, so a run never reports a guessed URL as a discovered one.
const CYRUS_DEFAULT_PREFIX = '/tzdist';

// One the account plausibly uses, one clearly foreign. Probe constants, so safe to print.
const ZONES = ['Australia/Sydney', 'Asia/Hong_Kong'];

// The truncation spans. Fixed constants, chosen so the result is directly comparable to that
// stored shape: docs/fastmail-action-availability.md records a client-authored event whose
// stored VTIMEZONE carried `TZUNTIL:20260923T145959Z`, so asking for that exact end instant
// makes a match visible byte for byte. The year span crosses both Sydney DST
// transitions, so its observance count sits between the event span's and the untruncated
// block's — which is what shows truncation tracking the span rather than being a fixed trim.
const SPAN_EVENT = { start: '2026-08-22T00:00:00Z', end: '2026-09-23T14:59:59Z' };
const SPAN_YEAR = { start: '2026-01-01T00:00:00Z', end: '2027-01-01T00:00:00Z' };

// libical serialises a UTC date-time in iCalendar basic form, which is what `TZUNTIL` carries.
const icalUtc = iso => `${new Date(iso).toISOString().slice(0, 19).replace(/[-:.]/g, '')}Z`;

const { check, failures } = makeChecker();

// Every account-derived string that must never reach the log, pushed as it is read.
const secrets = [];
const redact = s => {
  const substringPass = secrets
    .filter(Boolean)
    .reduce((acc, v) => acc.split(v).join('<redacted>'), String(s ?? ''));
  return substringPass
    .split(USERNAME).join('<account>')
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');
};

/** One HTTP request. `redirect: manual` so a well-known 301 can be READ rather than followed. */
async function http(method, url, { body, headers = {} } = {}) {
  const res = await fetch(url, {
    method,
    redirect: 'manual',
    headers: {
      Authorization: AUTH,
      ...(body !== undefined ? { 'Content-Type': 'text/xml;charset=UTF-8' } : {}),
      ...headers,
    },
    body,
  });
  return { status: res.status, headers: res.headers, text: await res.text() };
}

const el = (xml, name) => {
  const m = xml.match(new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`));
  return m ? m[1] : undefined;
};
const abs = href => new URL(href, ROOT).href;

/**
 * The DAV status a multistatus reported for one named property.
 *
 * A 207 carries one propstat PER STATUS, so asking for two properties and finding one returns two
 * blocks — reading a single `<d:status>` out of the whole response would hand one block's status to
 * every property in it. The name is matched up to a delimiter, so an empty `<c:foo/>` and a valued
 * `<c:foo>x</c:foo>` both hit and `calendar-timezone` cannot match `calendar-timezone-id`.
 */
function propStatus(xml, name) {
  const blocks = [...String(xml ?? '').matchAll(
    /<(?:[\w-]+:)?propstat(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w-]+:)?propstat>/g,
  )].map(m => m[1]);
  for (const block of blocks) {
    if (new RegExp(`<(?:[\\w-]+:)?${name}[\\s/>]`).test(block)) return el(block, 'status')?.trim();
  }
  return undefined;
}
const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

/**
 * Which tier answered a FAILED request: the CalDAV backend, or the edge proxy in front of it.
 *
 * Both send the same `Server:` header. The backend's error pages carry a `color-scheme` style
 * attribute on `<html>`; the proxy's stock page has none, and a body with no marker is
 * attributed to the proxy (the conservative reading). A 200 body carries no marker either, so
 * it gets the empty string rather than a misreported proxy.
 */
const tier = res => (res.status === 200
  ? ''
  : ` (answered by the ${/color-scheme:\s*dark light/.test(res.text) ? 'CalDAV backend' : 'edge proxy'})`);

/** Unfold RFC 5545 continuation lines, so a wrapped property reads as one value. */
const unfold = s => String(s ?? '').replace(/\r?\n[ \t]/g, '');

/**
 * The VTIMEZONE blocks in an iCalendar body, reduced to the facts the conditions turn on: the
 * caller prints from this, and the raw body is printed only once, under a cap.
 */
function timezonesIn(body) {
  const text = unfold(body);
  const blocks = text.match(/BEGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/g) || [];
  return blocks.map(b => ({
    tzid: b.match(/^TZID:(.*)$/m)?.[1]?.trim(),
    tzuntil: b.match(/^TZUNTIL:(.*)$/m)?.[1]?.trim(),
    tzurl: b.match(/^TZURL:(.*)$/m)?.[1]?.trim(),
    observances: (b.match(/^BEGIN:(?:STANDARD|DAYLIGHT)$/gm) || []).length,
    rrules: (b.match(/^RRULE:/gm) || []).length,
    lines: b.split(/\r?\n/).length,
  }));
}

/**
 * GET one zone from the tzdist service, with optional truncation.
 *
 * RFC 7808 percent-encodes the tzid's `/`; Cyrus's path parser also accepts the literal slash.
 * Both are tried and the one that answered is reported. `start`/`end` go UNENCODED, as in RFC
 * 7808's own examples.
 */
async function getZone(base, zone, span) {
  const query = span ? `?start=${span.start}&end=${span.end}` : '';
  const attempts = [
    { label: 'percent-encoded', path: encodeURIComponent(zone) },
    { label: 'literal slash', path: zone },
  ];
  let last = null;
  for (const attempt of attempts) {
    const res = await http('GET', `${base}/zones/${attempt.path}${query}`);
    // `encoding` names the spelling that WORKED, so it is null when none did — naming the last
    // one tried would read as "this spelling was accepted" on a run where neither was.
    last = { ...res, encoding: res.status === 200 ? attempt.label : null };
    if (res.status === 200) return last;
  }
  return last;
}

console.log('\nRFC 7808 timezone data distribution on this account (#166)');
console.log('PASS/FAIL, counts and zone names only — no account value is printed.\n');

// Wrapped so an unguarded rejection lands as one redacted FAIL line rather than a raw stack.
try {
  console.log('=== Discovery: where does this account say its timezone service lives? ===');

  const rootPf = await http('PROPFIND', ROOT, {
    body: PROPFIND(['d:current-user-principal']),
    headers: { Depth: '0' },
  });
  const principalHref = el(el(rootPf.text, 'current-user-principal') ?? '', 'href')?.trim();
  if (principalHref) secrets.push(principalHref, abs(principalHref));
  const prinPf = principalHref
    ? await http('PROPFIND', abs(principalHref), {
      body: PROPFIND(['c:calendar-home-set']),
      headers: { Depth: '0' },
    })
    : null;
  const homeHref = prinPf ? el(el(prinPf.text, 'calendar-home-set') ?? '', 'href')?.trim() : undefined;
  const home = homeHref ? abs(homeHref) : null;
  if (homeHref) secrets.push(homeHref, home, new URL(home).pathname);
  check('the calendar home was discovered (so the routes below have somewhere to ask)', home !== null);

  const discovered = [];

  // Route A: the CalDAV feature token. Present iff the tzdist namespace is enabled.
  let routeA = false;
  if (home) {
    const opts = await http('OPTIONS', home);
    const davHeader = opts.headers.get('dav') ?? '';
    routeA = /calendar-no-timezone/i.test(davHeader);
    console.log(`  A. OPTIONS DAV header advertises calendar-no-timezone: ${routeA ? 'yes' : 'no'}`);
  } else {
    console.log('  A. OPTIONS DAV header: not asked (no calendar home)');
  }

  // Route B: CALDAV:timezone-service-set. The authoritative one — the server names the prefix.
  //
  // CALDAV:calendar-timezone-id rides along REPORTED AND NOT GATED: a deployment may serve the
  // service with no zone id set on any collection, so a condition over it would fail a working
  // service. It separates "the whole time-zones-by-reference family is off" from "the service
  // is off".
  if (home) {
    const tzsPf = await http('PROPFIND', home, {
      body: PROPFIND(['c:timezone-service-set', 'c:calendar-timezone-id']),
      headers: { Depth: '0' },
    });
    const set = el(tzsPf.text, 'timezone-service-set');
    const href = set ? el(set, 'href')?.trim() : undefined;
    if (href) discovered.push({ route: 'B (CALDAV:timezone-service-set)', base: abs(href).replace(/\/$/, '') });
    console.log(`  B. CALDAV:timezone-service-set names a service: ${href ? 'yes' : 'no'}`);

    const tzIdStatus = propStatus(tzsPf.text, 'calendar-timezone-id');
    const tzIdValue = el(tzsPf.text, 'calendar-timezone-id')?.trim();
    console.log(`  B'. CALDAV:calendar-timezone-id, asked in the same request (reported, not gated): ${tzIdStatus ?? '(no propstat)'}${tzIdValue ? ', carries a value' : ', no value'}`);
  } else {
    console.log('  B. CALDAV:timezone-service-set: not asked (no calendar home)');
    console.log("  B'. CALDAV:calendar-timezone-id: not asked (no calendar home)");
  }

  // Route C: the RFC 7808 well-known URI, redirects unfollowed so the Location IS the answer.
  const wk = await http('GET', `${ORIGIN}/.well-known/timezone`);
  const wkLocation = wk.headers.get('location');
  if (wk.status >= 300 && wk.status < 400 && wkLocation) {
    discovered.push({ route: 'C (/.well-known/timezone)', base: abs(wkLocation).replace(/\/$/, '') });
  }
  console.log(`  C. GET /.well-known/timezone: HTTP ${wk.status}${wkLocation ? ' -> Location present' : ''}`);

  // Route D: Cyrus's HTML index of the well-known URLs it has enabled.
  const wkIndex = await http('GET', `${ORIGIN}/.well-known/`);
  const indexNamesTimezone = wkIndex.status === 200 && /\.well-known\/timezone/.test(wkIndex.text);
  if (indexNamesTimezone) {
    const href = wkIndex.text.match(/href="([^"]*)"[^<]*>[^<]*\/\.well-known\/timezone/)?.[1];
    if (href) discovered.push({ route: 'D (/.well-known/ index)', base: abs(href).replace(/\/$/, '') });
  }
  console.log(`  D. GET /.well-known/ index lists a timezone entry: ${indexNamesTimezone ? 'yes' : 'no'} (HTTP ${wkIndex.status})`);

  console.log(`  Routes that named a base URL: ${discovered.length}`);
  for (const d of discovered) console.log(`    - ${d.route}`);

  const guessed = discovered.length === 0;
  const base = guessed ? `${ORIGIN}${CYRUS_DEFAULT_PREFIX}` : discovered[0].base;
  console.log(
    guessed
      ? `  No route named a base. Falling back to Cyrus's compiled-in default prefix (${CYRUS_DEFAULT_PREFIX}) as a MEASUREMENT, so an unadvertised-but-serving deployment is still visible.`
      : `  Using the base named by route ${discovered[0].route}.`,
  );

  // `capabilities` also lists the actions offered, so a server with no `get` is caught here.
  console.log('\n=== Condition 1: does the timezone service answer? ===');
  const capa = await http('GET', `${base}/capabilities`);
  console.log(`  GET <base>/capabilities: HTTP ${capa.status}${tier(capa)}`);
  if (capa.status !== 200) {
    console.log(`    body: ${redact(capa.text).replace(/\s+/g, ' ').slice(0, 200)}`);
  }

  // See "Reading a 404 correctly" in the header.
  const viaDav = await http('GET', `${ROOT}tzdist/capabilities`);
  console.log(`  GET <dav root>tzdist/capabilities: HTTP ${viaDav.status}${tier(viaDav)}`);

  let capaJson = null;
  try { capaJson = JSON.parse(capa.text); } catch { /* not JSON: reported by the checks below */ }
  const actions = Array.isArray(capaJson?.actions)
    ? capaJson.actions.map(a => a?.name).filter(Boolean)
    : [];
  if (capa.status === 200) {
    console.log(`  version=${capaJson?.version ?? '(none)'}  actions=${actions.length ? actions.join(',') : '(none)'}`);
  }

  check('condition 1: the timezone service answered a capabilities request with HTTP 200', capa.status === 200, `status=${capa.status}`);
  check('condition 1: the capabilities body parses as JSON naming a "get" action', actions.includes('get'), `actions=${actions.length}`);
  check(
    'condition 1: the service was DISCOVERED rather than guessed at',
    !guessed,
    guessed
      ? `no discovery route named a base URL; DAV feature token ${routeA ? 'present' : 'absent'}`
      : discovered[0].route,
  );

  console.log('\n=== Condition 2: does a named zone return a parseable VTIMEZONE? ===');
  const full = new Map();
  for (const zone of ZONES) {
    const res = await getZone(base, zone);
    const type = res.headers.get('content-type') ?? '';
    console.log(`  ${zone}: HTTP ${res.status} (${res.encoding ? `${res.encoding} path accepted` : 'neither path spelling accepted'}), content-type=${type.split(';')[0] || '(none)'}`);
    if (res.status !== 200) {
      console.log(`    body: ${redact(res.text).replace(/\s+/g, ' ').slice(0, 200)}`);
      check(`condition 2 [${zone}]: the zone was served`, false, `status=${res.status}`);
      continue;
    }
    const blocks = timezonesIn(res.text);
    full.set(zone, { blocks, body: res.text });
    console.log(`    VTIMEZONE blocks=${blocks.length}, observances=${blocks[0]?.observances ?? 0}, RRULEs=${blocks[0]?.rrules ?? 0}, lines=${blocks[0]?.lines ?? 0}`);
    check(`condition 2 [${zone}]: exactly one VTIMEZONE came back`, blocks.length === 1, `count=${blocks.length}`);
    check(`condition 2 [${zone}]: its TZID is the zone asked for`, blocks[0]?.tzid === zone, `tzid=${blocks[0]?.tzid ?? '(none)'}`);
    check(`condition 2 [${zone}]: it carries at least one STANDARD/DAYLIGHT observance`, (blocks[0]?.observances ?? 0) >= 1);
  }

  // The untruncated block carrying NO TZUNTIL is checked too: otherwise a TZUNTIL present in
  // the zone file all along would read as proof of truncation.
  console.log('\n=== Condition 3: is start/end truncation honoured? ===');
  let printed = false;
  for (const zone of ZONES) {
    const fullEntry = full.get(zone);
    if (!fullEntry) {
      check(`condition 3 [${zone}]: truncation could be measured`, false, 'the untruncated fetch did not succeed');
      continue;
    }
    const fullBlock = fullEntry.blocks[0];

    const year = await getZone(base, zone, SPAN_YEAR);
    const event = await getZone(base, zone, SPAN_EVENT);
    if (event.status !== 200 || year.status !== 200) {
      console.log(`  ${zone}: truncated fetch HTTP ${event.status} / ${year.status}`);
      console.log(`    body: ${redact(event.text).replace(/\s+/g, ' ').slice(0, 200)}`);
      check(`condition 3 [${zone}]: the truncated request was served`, false, `status=${event.status}`);
      continue;
    }

    const yearBlock = timezonesIn(year.text)[0];
    const eventBlock = timezonesIn(event.text)[0];
    const wantUntil = icalUtc(SPAN_EVENT.end);

    console.log(`  ${zone}: observances  full=${fullBlock?.observances ?? 0}  year-span=${yearBlock?.observances ?? 0}  event-span=${eventBlock?.observances ?? 0}`);
    console.log(`    TZUNTIL  full=${fullBlock?.tzuntil ?? '(none)'}  event-span=${eventBlock?.tzuntil ?? '(none)'}  (asked for end=${SPAN_EVENT.end})`);
    console.log(`    TZURL echoes the truncation query: ${eventBlock?.tzurl?.includes('start=') ? 'yes' : 'no'}`);

    check(
      `condition 3 [${zone}]: the truncated block carries TZUNTIL equal to the requested end`,
      eventBlock?.tzuntil === wantUntil,
      `tzuntil=${eventBlock?.tzuntil ?? '(none)'} want=${wantUntil}`,
    );
    check(
      `condition 3 [${zone}]: the UNtruncated block carries no TZUNTIL, so the property came from truncation`,
      fullBlock?.tzuntil === undefined,
      `tzuntil=${fullBlock?.tzuntil ?? '(none)'}`,
    );
    check(
      `condition 3 [${zone}]: the truncated block has FEWER observances than the full one`,
      (eventBlock?.observances ?? 0) < (fullBlock?.observances ?? 0),
      `event-span=${eventBlock?.observances ?? 0} full=${fullBlock?.observances ?? 0}`,
    );

    // Only the first zone's block: the rest add no information.
    if (!printed && eventBlock) {
      printed = true;
      const lines = unfold(event.text).match(/BEGIN:VTIMEZONE[\s\S]*?END:VTIMEZONE/)?.[0].split(/\r?\n/) ?? [];
      console.log(`\n  Truncated ${zone} block as served (first 25 lines of ${lines.length}):`);
      for (const line of lines.slice(0, 25)) console.log(`    ${redact(line)}`);
    }
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
