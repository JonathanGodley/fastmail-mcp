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
// WHY THE FIXTURES PERSIST. A human has to view them in the client, so `create` leaves the
// four events in place and `cleanup` removes the whole collection later. Nothing else sweeps
// them, so an un-run `cleanup` leaves a stray calendar in the account.
//
// NEVER TOUCH A COLLECTION IT DID NOT MINT. `cleanup` deletes a WHOLE COLLECTION, so a display
// name alone cannot identify a target: `cleanup Personal` must not wipe a real calendar. The
// provenance test is the PATH: `create` mints a collection whose last segment starts with
// `mcp-164-`, and both phases match on the display name AND that segment. If MKCALENDAR fails,
// or a same-named collection exists on any other path, both phases refuse loudly and non-zero.
// The writes are addressed by the minted URL, never by the name.
//
// EVERY RUN MINTS ITS OWN COLLECTION, so its resource count is meaningful and one run's
// `cleanup` cannot take another's fixtures unasked. `cleanup` deletes EVERY minted collection
// under the name, so nothing is stranded.
//
// NO PARTICIPANTS: an attendee makes the server send real iTIP mail (see the README).
//
// WHAT IS PRINTED. Each create_calendar_event response verbatim, then every stored resource
// fetched back raw (REPORT, NO expand) — what the client is rendering — then a "Look for in
// the client:" line per event. Output redacts the account name and every email-shaped string.
//
// The writes go through the BUILT server over the MCP harness; the collection plumbing and the
// read-back are raw CalDAV, so none of our parsing sits between the store and the dump.

import { makeChecker, text, jsonOf } from './probelib.mjs';
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

// CalDAV hrefs and an event's ORGANIZER carry the account's address; output gets pasted into issues.
const redact = s => String(s)
  .split(USERNAME).join('<account>')
  .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');

// Every XML matcher tolerates any namespace prefix because Cyrus picks its own.
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

// Cyrus returned calendar-data as a CDATA section on one live run (23 Aug 2026), wrapped and
// NOT entity-escaped. EVERY section is SPLICED IN PLACE: a literal `]]>` splits a payload into
// two adjacent sections, and a MIXED payload has escaped text either side, so only the text
// outside the sections gets the entity pass.
const ENTITIES = t => t
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
const unescapeXml = s => {
  let out = '';
  let at = 0;
  for (const m of s.matchAll(/<!\[CDATA\[([\s\S]*?)\]\]>/g)) {
    out += ENTITIES(s.slice(at, m.index)) + m[1];
    at = m.index + m[0].length;
  }
  return out + ENTITIES(s.slice(at));
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

/** Is THIS run's collection — the one at `url` — present in a list_calendars response?
 *
 *  The URL, never the display name: an earlier run's same-named collection is already listed,
 *  so a name test would pass before the server had seen this one. Trailing slashes are
 *  ignored because the two sides need not agree on them. */
const sameCalendarUrl = (a, b) => a.replace(/\/+$/, '') === b.replace(/\/+$/, '');

function calendarIsListed(body, url) {
  try {
    const payload = jsonOf(body);
    const urls = [];
    const walk = v => {
      if (Array.isArray(v)) v.forEach(walk);
      else if (v && typeof v === 'object') for (const [k, val] of Object.entries(v)) {
        if (typeof val === 'string' && /^(url|href)$/i.test(k)) urls.push(val);
        else walk(val);
      }
    };
    walk(payload);
    return urls.some(u => sameCalendarUrl(u, url));
  } catch {
    // Not JSON (or no payload): the URL as a plain substring. Still unambiguous — the
    // stamped path segment appears nowhere else.
    const bare = url.replace(/\/+$/, '');
    return body.includes(bare);
  }
}

// PROVENANCE: only a collection on a minted path is eligible to be written into or deleted.
const MINTED_PREFIX = 'mcp-164-';
const mintedSegment = () => `${MINTED_PREFIX}${Date.now()}`;
const isMinted = url => {
  const segments = new URL(url).pathname.split('/').filter(Boolean);
  return (segments.at(-1) ?? '').startsWith(MINTED_PREFIX);
};

// The four reference shapes, dated from TODAY: anchored on the next Wednesday at least two
// days out, clear of today's real entries and all in one week view.
//
// `timeZone` is omitted on the first (create writes the configured zone) and named on the
// second: that pair shows whether the client annotates a non-default zone the way it does its
// own. DTEND is exclusive, so the three-day band ends on the fourth day.
const WEDNESDAY = 3;
const MIN_DAYS_AHEAD = 2;

/** Local calendar date, YYYY-MM-DD, offset by whole days. Local, not UTC: the fixture dates
 *  are the days the operator will open in the client, which are their own local days. */
function localDate(base, addDays = 0) {
  const d = new Date(base.getFullYear(), base.getMonth(), base.getDate() + addDays);
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
/** "Wed 26 Aug 2026", for the look-for lines: the operator needs the day, not an ISO string. */
const humanDate = iso => {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('en-AU',
    { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
};

const TODAY = new Date();
// Days from today to the next Wednesday that is at least MIN_DAYS_AHEAD away.
const ANCHOR_OFFSET = (() => {
  let offset = (WEDNESDAY - TODAY.getDay() + 7) % 7;
  while (offset < MIN_DAYS_AHEAD) offset += 7;
  return offset;
})();
const DAY_TIMED = localDate(TODAY, ANCHOR_OFFSET);          // both timed events
const DAY_ALLDAY_1 = localDate(TODAY, ANCHOR_OFFSET + 1);   // the single all-day
const DAY_ALLDAY_1_END = localDate(TODAY, ANCHOR_OFFSET + 2);
const DAY_ALLDAY_3 = localDate(TODAY, ANCHOR_OFFSET + 2);   // the three-day band starts here
const DAY_ALLDAY_3_LAST = localDate(TODAY, ANCHOR_OFFSET + 4);
const DAY_ALLDAY_3_END = localDate(TODAY, ANCHOR_OFFSET + 5); // exclusive

const EVENTS = [
  {
    title: 'MCP-164 timed default zone',
    args: { start: `${DAY_TIMED}T10:00:00`, end: `${DAY_TIMED}T11:00:00` },
    lookFor: `on ${humanDate(DAY_TIMED)}: renders at 10:00-11:00 local with no unexpected zone annotation`,
  },
  {
    title: 'MCP-164 timed Hong Kong',
    args: { start: `${DAY_TIMED}T10:00:00`, end: `${DAY_TIMED}T11:00:00`, timeZone: 'Asia/Hong_Kong' },
    lookFor: `on ${humanDate(DAY_TIMED)}: the client shows Asia/Hong_Kong the way it does for its own zone-picker events`,
  },
  {
    title: 'MCP-164 all-day one day',
    args: { start: DAY_ALLDAY_1, end: DAY_ALLDAY_1_END },
    lookFor: `"All day" on exactly ${humanDate(DAY_ALLDAY_1)}`,
  },
  {
    title: 'MCP-164 all-day three days',
    args: { start: DAY_ALLDAY_3, end: DAY_ALLDAY_3_END },
    lookFor: `spans ${humanDate(DAY_ALLDAY_3)} to ${humanDate(DAY_ALLDAY_3_LAST)} inclusive, ` +
      `with no off-by-one at the exclusive end (${humanDate(DAY_ALLDAY_3_END)} must be clear)`,
  },
];

const { check, failures } = makeChecker();

async function runCreate() {
  const homeUrl = await discoverHome();
  console.log(`Calendar home: ${redact(homeUrl)}`);
  console.log(`Temporary calendar display name: ${JSON.stringify(CAL_NAME)}`);

  // --- the collection: ALWAYS mint a fresh one (see the header) ------------------------
  // A foreign same-named collection means STOP even when a minted one exists:
  // src/caldav-client.ts resolves a `calendarId` name to whichever collection discovery
  // lists first, which can be the operator's real calendar.
  const sameName = (await listCollections(homeUrl)).filter(c => c.name === CAL_NAME);
  const earlier = sameName.filter(c => isMinted(c.url));
  const foreign = sameName.find(c => !isMinted(c.url));
  if (foreign) {
    check('no collection outside this probe carries the display name', false,
      `a calendar named ${JSON.stringify(CAL_NAME)} exists and is not this probe's`);
    console.error(`\nStopping: ${redact(foreign.url)} carries that display name but was not minted by`);
    console.error(`this probe (its path does not start with "${MINTED_PREFIX}"), so it is a real calendar.`);
    console.error('Nothing was written. Choose another name and re-run:');
    console.error(`  python scripts/probes/run-probe.py server-authored-events.probe.mjs create "MCP probe calendar 2"`);
    return { wroteAnything: false };
  }
  if (earlier.length) {
    console.log(`\n${earlier.length} earlier minted collection(s) under this name exist; \`cleanup\` will delete them all.`);
    for (const c of earlier) console.log(`  ${redact(c.url)}`);
    console.log('This run mints its own and leaves those untouched.');
  }
  const calendarUrl = new URL(`${mintedSegment()}/`, homeUrl).href;
  const mk = await dav('MKCALENDAR', calendarUrl, {
    body: `<?xml version="1.0" encoding="utf-8"?>\n` +
      `<c:mkcalendar xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:set><d:prop>` +
      `<d:displayname>${escapeXml(CAL_NAME)}</d:displayname>` +
      `<c:supported-calendar-component-set><c:comp name="VEVENT"/></c:supported-calendar-component-set>` +
      `</d:prop></d:set></c:mkcalendar>`,
  });
  if (!(mk.status >= 200 && mk.status < 300)) {
    // Deliberately terminal. Falling back to any existing calendar would put the cleanup
    // phase's whole-collection DELETE over events this run did not write.
    check(`MKCALENDAR ${redact(calendarUrl)}`, false, `HTTP ${mk.status} ${mk.statusText}`);
    console.error('\nStopping: this probe never writes into a calendar it did not mint.');
    console.error('Nothing was written.');
    return { wroteAnything: false };
  }
  check(`MKCALENDAR created the temporary collection`, true, `HTTP ${mk.status}`);
  console.log(`  ${redact(calendarUrl)}`);

  // --- write the four events through the built server ----------------------------------
  // Tracked so the failure banner can tell "nothing was written" from "some of it was":
  // set the moment a create is ATTEMPTED, since a call that errors may still have landed.
  let wroteAnything = false;
  const client = createClient({ env: process.env });
  try {
    await client.init();

    // The minted collection has to reach the server's discovery before anything is written.
    // Bounded retry; how long it takes is itself reported.
    const ATTEMPTS = 5;
    let visible = false;
    for (let attempt = 1; attempt <= ATTEMPTS && !visible; attempt++) {
      visible = calendarIsListed(text(await client.call('list_calendars', {})), calendarUrl);
      console.log(`list_calendars attempt ${attempt}/${ATTEMPTS}: ${visible ? 'minted collection listed' : 'not yet listed'}`);
      if (!visible && attempt < ATTEMPTS) await new Promise(r => setTimeout(r, 1000 * attempt));
    }
    check('the freshly minted collection is listed by list_calendars', visible);
    if (!visible) return { wroteAnything: false };

    console.log('');
    console.log(`Addressing the writes by URL: ${redact(calendarUrl)}`);
    for (const ev of EVENTS) {
      wroteAnything = true;
      console.log('-'.repeat(70));
      console.log(`create_calendar_event  ${ev.title}`);
      console.log(`  args: ${JSON.stringify(ev.args)}`);
      try {
        const res = await client.call('create_calendar_event', {
          // BY URL, never by display name: only the URL names exactly one collection.
          calendarId: calendarUrl,
          title: ev.title,
          ...ev.args,
        });
        const body = text(res);
        console.log(redact(body).split('\n').map(l => '  ' + l).join('\n'));
        check(`created "${ev.title}"`, !res.isError);
      } catch (err) {
        check(`created "${ev.title}"`, false, redact(err.message));
        continue;
      }
    }
  } finally {
    client.close();
  }

  // --- read every stored resource back, raw --------------------------------------------
  const { status, statusText, resources } = await fetchResources(calendarUrl);
  console.log(`\n${'='.repeat(70)}`);
  console.log(`Stored resources, fetched back raw over CalDAV (REPORT ${status} ${statusText})`);
  console.log(`This is what the client is rendering.`);
  // Nothing else has ever written to this collection, so any other count is a real discrepancy.
  check(`exactly the ${EVENTS.length} events this run wrote are stored in the collection it minted`,
    resources.length === EVENTS.length,
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
  console.log(`The events are LEFT IN PLACE for viewing. When you are done, remove them with the`);
  console.log(`command below, which deletes EVERY collection this probe has minted under this name,`);
  console.log(`this run's and any earlier run's:`);
  console.log(`  python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup ${JSON.stringify(CAL_NAME)}`);
  return { wroteAnything };
}

async function runCleanup() {
  const homeUrl = await discoverHome();
  console.log(`Calendar home: ${redact(homeUrl)}`);
  console.log(`Looking for the collection named ${JSON.stringify(CAL_NAME)}.`);

  const sameName = (await listCollections(homeUrl)).filter(c => c.name === CAL_NAME);
  // ALL of them, not the first: a crashed `create` leaves one behind and the next mints
  // another, and deleting only the first would strand the rest.
  const targets = sameName.filter(c => isMinted(c.url));
  if (!targets.length) {
    const foreign = sameName.find(c => !isMinted(c.url));
    if (foreign) {
      check('the collection to delete was made by this probe', false,
        `a collection named ${JSON.stringify(CAL_NAME)} exists but was not made by this probe - refusing to delete it`);
      console.error(`\n${redact(foreign.url)} carries that display name, but its path does not start`);
      console.error(`with "${MINTED_PREFIX}", so this probe did not create it. Nothing was deleted.`);
      return;
    }
    console.log('\nNo collection of this probe\'s with that display name - nothing to clean up.');
    return;
  }
  console.log(`Found ${targets.length} collection(s) minted by this probe under that name:`);
  for (const t of targets) console.log(`  ${redact(t.url)}`);

  for (const target of targets) {
    const { status, statusText, resources } = await fetchResources(target.url);
    console.log(`\n${redact(target.url)}`);
    console.log(`Resources in the collection (REPORT ${status} ${statusText}): ${resources.length}`);
    for (const r of resources) console.log(`  - ${redact(summaryOf(r.data))}`);

    // The whole collection, in one request: it takes every resource inside with it, and
    // nothing outside it. No DELETE is issued against anything not in `targets`.
    const del = await dav('DELETE', target.url);
    console.log(`DELETE -> HTTP ${del.status} ${del.statusText}`);
    check(`DELETE of ${redact(target.url)} succeeded`, del.status >= 200 && del.status < 300,
      `HTTP ${del.status} ${del.statusText}`);
  }

  // Scoped to MINTED collections: a foreign same-named one surviving is correct.
  const stillThere = (await listCollections(homeUrl))
    .filter(c => c.name === CAL_NAME && isMinted(c.url));
  check('every minted collection under that name is gone from the calendar home',
    stillThere.length === 0, stillThere.length ? `${stillThere.length} remain(s)` : '');
}

// A throw leaves it unknown whether a create landed, so the banner below assumes it might
// have: the honest default for a phase that deliberately leaves fixtures behind.
let outcome = { wroteAnything: MODE === 'create' };
try {
  outcome = (MODE === 'create' ? await runCreate() : await runCleanup()) ?? outcome;
} catch (err) {
  check(`${MODE} completed without throwing`, false, redact(err?.message ?? String(err)));
}

if (failures() > 0) {
  console.error(`\n${failures()} check(s) FAILED.`);
  // Only when a write was attempted: otherwise there are no fixtures to look for.
  if (MODE === 'create' && outcome.wroteAnything) {
    console.error('Nothing was auto-deleted: what got written is listed above, and removing it is your call.');
    console.error(`Run: python scripts/probes/run-probe.py server-authored-events.probe.mjs cleanup ${JSON.stringify(CAL_NAME)}`);
  }
  process.exitCode = 1;
}
