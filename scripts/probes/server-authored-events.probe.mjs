// What this probe settles
// -----------------------
// What THIS server's own create path puts in front of a human (#164). It is the exact
// inverse of `client-authored-events.probe.mjs`, whose measurement of Fastmail's own
// client writes is recorded in docs/fastmail-action-availability.md: there the client
// authored and we read the bytes; here we author the bytes and the client renders them.
// The question a byte dump cannot answer is whether the Fastmail client AGREES with what
// this server says it wrote — whether a designator-less start really lands on the wall
// clock the response promised, whether an explicitly-zoned event is annotated the way the
// client annotates its own zone-picker events, and whether an exclusive DTEND draws the
// all-day band over the days the caller meant. That is a human looking at a screen, so
// this probe is two phases with a person in between.
//
// Usage (via the token launcher, which injects the API token and the CalDAV app password):
//
//   python scripts/probes/run-probe.py server-authored-events.probe.mjs create  [calendarName]
//   ... the operator opens the Fastmail client and looks at the four events ...
//   python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup [calendarName]
//
// `calendarName` defaults to "MCP probe calendar" and must be the SAME in both phases.
//
// WHY THE FIXTURES PERSIST. Every other probe here deletes its fixtures in a finally.
// This one deliberately does not: the whole point is that a human views them in the
// client, which cannot happen inside a script run. `create` leaves the four events in
// place and exits; `cleanup` removes the whole collection later. Nothing else sweeps
// them, so an un-run `cleanup` leaves a stray calendar in the account.
//
// MKCALENDAR OR STOP. The fixtures go into a temporary collection minted by MKCALENDAR.
// If MKCALENDAR fails the probe prints the status and exits non-zero — it NEVER falls back
// to writing into an existing calendar. This runs against a live personal account and the
// cleanup phase deletes a whole collection; writing into a real calendar would put that
// delete over the operator's own events. If a collection with the requested display name
// already exists it is reused rather than a second one minted, so a re-run after a partial
// failure does not litter.
//
// NO PARTICIPANTS. None of the four events carries one. An attendee makes the server's
// scheduling layer send a real iTIP invitation from the account under test, and the later
// delete sends the matching cancellation (see the README, "Creating a calendar event with
// a participant sends real mail"). Nothing here needs the scheduling hop.
//
// WHAT IS PRINTED. For each event: the create_calendar_event response verbatim (its
// statement of the zone / all-day form it actually wrote, and the id and url it returned),
// then — after all four are written — every resource in the collection fetched back raw
// over CalDAV with a calendar-query REPORT and NO expand, printed as stored. That raw
// iCalendar is what the client is rendering, so it is the thing to compare the screen
// against. Each event then gets a "Look for in the client:" line naming what to check.
// Output redacts the account name and every email-shaped string.
//
// The writes go through the BUILT server over the MCP harness (so the create path under
// test is the shipped one); the collection plumbing and the read-back are raw CalDAV over
// bare fetch, so none of our parsing sits between the store and the dump.

import { makeChecker, text } from './probelib.mjs';
import { createClient } from '../mcp-harness.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py server-authored-events.probe.mjs <create|cleanup>');
  process.exit(1);
}

const MODE = process.argv[2];
if (MODE !== 'create' && MODE !== 'cleanup') {
  console.error('FAIL  first argument must be "create" or "cleanup".');
  console.error('      python scripts/probes/run-probe.py server-authored-events.probe.mjs create  [calendarName]');
  console.error('      python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup [calendarName]');
  process.exit(1);
}
const CAL_NAME = process.argv[3] || 'MCP probe calendar';

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');

// Every CalDAV href on this server embeds the account's own address, and an event's
// ORGANIZER carries it too. Nothing printed here may: probe output gets pasted into issues.
const redact = s => String(s)
  .split(USERNAME).join('<account>')
  .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');

// ---------------------------------------------------------------------------
// Minimal HTTP + XML plumbing. Regex parsing is enough for a probe against one known
// server; every matcher tolerates any namespace prefix because Cyrus picks its own.
// ---------------------------------------------------------------------------
async function dav(method, url, { body, headers = {} } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: AUTH,
      ...(body !== undefined ? { 'Content-Type': 'text/xml;charset=UTF-8' } : {}),
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

// Cyrus returns calendar-data as a CDATA section, so the raw iCalendar arrives wrapped and
// NOT entity-escaped. Unwrap first and skip the entity pass on that content: a dump that
// still says `<![CDATA[BEGIN:VCALENDAR` is not the stored bytes.
const unescapeXml = s => {
  const cdata = s.match(/^\s*<!\[CDATA\[([\s\S]*?)\]\]>\s*$/);
  if (cdata) return cdata[1];
  return s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
};
const escapeXml = s => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

// Every VEVENT in the collection, stored form. No c:expand and no time-range: the point is
// the bytes as written, and a window would silently drop a fixture that fell outside it.
const ALL_EVENTS_QUERY =
  `<?xml version="1.0" encoding="utf-8"?>\n` +
  `<c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:d="DAV:">` +
  `<d:prop><d:getetag/><c:calendar-data/></d:prop>` +
  `<c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT"/></c:comp-filter></c:filter>` +
  `</c:calendar-query>`;

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

/** Every calendar collection under the home, with its display name. */
async function listCollections(homeUrl) {
  const pf = await dav('PROPFIND', homeUrl, {
    body: PROPFIND(['d:displayname', 'd:resourcetype']), headers: { Depth: '1' },
  });
  const out = [];
  for (const block of responses(pf.text)) {
    const href = (el(block, 'href') ?? '').trim();
    if (!href || abs(href) === homeUrl) continue;
    if (!/calendar[\s/>]/.test(el(block, 'resourcetype') ?? '')) continue;
    out.push({ url: abs(href), name: unescapeXml(el(block, 'displayname') ?? '(unnamed)') });
  }
  return out;
}

/** Every stored resource in one collection: href + calendar-data, verbatim. */
async function fetchResources(calendarUrl) {
  const res = await dav('REPORT', calendarUrl, { body: ALL_EVENTS_QUERY, headers: { Depth: '1' } });
  const out = [];
  for (const block of responses(res.text)) {
    const data = unescapeXml(el(block, 'calendar-data') ?? '');
    if (!data) continue;
    out.push({ href: (el(block, 'href') ?? '').trim(), data: data.replace(/\r\n/g, '\n').trim() });
  }
  return { status: res.status, statusText: res.statusText, resources: out };
}

const summaryOf = data => (data.match(/^SUMMARY:(.*)$/m)?.[1] ?? '(no SUMMARY)').trim();

// ---------------------------------------------------------------------------
// The four reference shapes. Deliberately in one contiguous stretch of days so a single
// week view of the client shows all four at once. `timeZone` is omitted on the first
// (create writes the configured zone) and named on the second, which is the pair that
// tells you whether the client annotates a non-default zone the way it does its own.
// DTEND is exclusive, so the three-day band ends on the 31st to cover 28/29/30.
// ---------------------------------------------------------------------------
const EVENTS = [
  {
    title: 'MCP-164 timed default zone',
    args: { start: '2026-08-26T10:00:00', end: '2026-08-26T11:00:00' },
    lookFor: 'renders at 10:00-11:00 local with no unexpected zone annotation',
  },
  {
    title: 'MCP-164 timed Hong Kong',
    args: { start: '2026-08-26T10:00:00', end: '2026-08-26T11:00:00', timeZone: 'Asia/Hong_Kong' },
    lookFor: 'the client shows Asia/Hong_Kong the way it does for its own zone-picker events',
  },
  {
    title: 'MCP-164 all-day one day',
    args: { start: '2026-08-27', end: '2026-08-28' },
    lookFor: '"All day" on exactly 27 Aug',
  },
  {
    title: 'MCP-164 all-day three days',
    args: { start: '2026-08-28', end: '2026-08-31' },
    lookFor: 'spans 28-30 Aug with no off-by-one at the exclusive end',
  },
];

// ---------------------------------------------------------------------------
const { check, failures } = makeChecker();

async function runCreate() {
  const homeUrl = await discoverHome();
  console.log(`Calendar home: ${redact(homeUrl)}`);
  console.log(`Temporary calendar display name: ${JSON.stringify(CAL_NAME)}`);

  // --- the collection: reuse an existing one by display name, else MKCALENDAR ----------
  const existing = (await listCollections(homeUrl)).find(c => c.name === CAL_NAME);
  let calendarUrl;
  if (existing) {
    calendarUrl = existing.url;
    console.log(`\nA collection with that display name already exists - reusing it rather than minting a second.`);
    console.log(`  ${redact(calendarUrl)}`);
    check('collection available (reused existing)', true);
  } else {
    const candidateUrl = new URL(`mcp-164-${Date.now()}/`, homeUrl).href;
    const mk = await dav('MKCALENDAR', candidateUrl, {
      body: `<?xml version="1.0" encoding="utf-8"?>\n` +
        `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
        `<d:displayname>${escapeXml(CAL_NAME)}</d:displayname>` +
        `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
        `</d:prop></d:set></c:mkcalendar>`,
    });
    if (!(mk.status >= 200 && mk.status < 300)) {
      // Deliberately terminal. Falling back to an existing calendar would put the cleanup
      // phase's whole-collection DELETE over the operator's own events.
      check(`MKCALENDAR ${redact(candidateUrl)}`, false, `HTTP ${mk.status} ${mk.statusText}`);
      console.error('\nStopping: this probe never writes into an existing calendar.');
      return;
    }
    calendarUrl = candidateUrl;
    check(`MKCALENDAR created the temporary collection`, true, `HTTP ${mk.status}`);
    console.log(`  ${redact(calendarUrl)}`);
  }

  // --- write the four events through the built server ----------------------------------
  const client = createClient({ env: process.env });
  try {
    await client.init();

    // The collection has to be visible to the JMAP/CalDAV surface the server reads before
    // create_calendar_event can address it by display name. Bounded retry, not a poll loop.
    let visible = false;
    for (let attempt = 1; attempt <= 5 && !visible; attempt++) {
      const body = text(await client.call('list_calendars', {}));
      visible = body.includes(CAL_NAME);
      if (!visible) await new Promise(r => setTimeout(r, 1000 * attempt));
      console.log(`list_calendars attempt ${attempt}: ${visible ? 'calendar visible' : 'not yet visible'}`);
    }
    check('the temporary calendar is addressable by display name in list_calendars', visible);
    if (!visible) return;

    console.log('');
    for (const ev of EVENTS) {
      const res = await client.call('create_calendar_event', {
        calendarId: CAL_NAME,
        title: ev.title,
        ...ev.args,
      });
      const body = text(res);
      console.log('-'.repeat(70));
      console.log(`create_calendar_event  ${ev.title}`);
      console.log(`  args: ${JSON.stringify(ev.args)}`);
      console.log(redact(body).split('\n').map(l => '  ' + l).join('\n'));
      check(`created "${ev.title}"`, !res.isError && /error/i.test(body) === false);
    }
  } finally {
    client.close();
  }

  // --- read every stored resource back, raw --------------------------------------------
  const { status, statusText, resources } = await fetchResources(calendarUrl);
  console.log(`\n${'='.repeat(70)}`);
  console.log(`Stored resources, fetched back raw over CalDAV (REPORT ${status} ${statusText})`);
  console.log(`This is what the client is rendering.`);
  check(`all ${EVENTS.length} events are stored in the collection`, resources.length === EVENTS.length,
    `found ${resources.length}`);
  for (const r of resources) {
    console.log('='.repeat(70));
    console.log(`event:    ${redact(summaryOf(r.data))}`);
    console.log(`resource: ${redact(r.href)}`);
    console.log('-'.repeat(70));
    console.log(redact(r.data));
    console.log('');
  }

  console.log('='.repeat(70));
  console.log('Look for in the client:');
  for (const ev of EVENTS) console.log(`  ${ev.title}\n      ${ev.lookFor}`);
  console.log('');
  console.log(`The events are LEFT IN PLACE for viewing. When you are done, remove them with:`);
  console.log(`  python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup ${JSON.stringify(CAL_NAME)}`);
}

async function runCleanup() {
  const homeUrl = await discoverHome();
  console.log(`Calendar home: ${redact(homeUrl)}`);
  console.log(`Looking for the collection named ${JSON.stringify(CAL_NAME)}.`);

  const target = (await listCollections(homeUrl)).find(c => c.name === CAL_NAME);
  if (!target) {
    console.log('\nNo collection with that display name - nothing to clean up.');
    return;
  }
  console.log(`  ${redact(target.url)}`);

  const { status, statusText, resources } = await fetchResources(target.url);
  console.log(`\nResources in the collection (REPORT ${status} ${statusText}): ${resources.length}`);
  for (const r of resources) console.log(`  - ${redact(summaryOf(r.data))}`);

  // The whole collection, in one request: it takes every resource inside with it, and
  // nothing outside it. No other DELETE is issued on any path.
  const del = await dav('DELETE', target.url);
  console.log(`\nDELETE ${redact(target.url)} -> HTTP ${del.status} ${del.statusText}`);
  check('DELETE of the temporary collection succeeded', del.status >= 200 && del.status < 300,
    `HTTP ${del.status} ${del.statusText}`);

  const stillThere = (await listCollections(homeUrl)).some(c => c.name === CAL_NAME);
  check('the collection is gone from the calendar home', !stillThere);
}

try {
  if (MODE === 'create') await runCreate();
  else await runCleanup();
} catch (err) {
  check(`${MODE} completed without throwing`, false, redact(err?.message ?? String(err)));
}

if (failures() > 0) {
  console.error(`\n${failures()} check(s) FAILED.`);
  if (MODE === 'create') {
    console.error('Nothing was auto-deleted: what got written is listed above, and removing it is your call.');
    console.error(`Run: python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup ${JSON.stringify(CAL_NAME)}`);
  }
  process.exitCode = 1;
}
