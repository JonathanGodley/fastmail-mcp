// What this probe settles
// -----------------------
// #197: this server tells a task collection from an event calendar by its display name alone
// (`DEFAULT_TASK_CALENDAR_NAME`), so a renamed or second task list is offered as a calendar and
// an event calendar carrying that name is hidden. RFC 4791 section 5.2.3's
// `CALDAV:supported-calendar-component-set` is the protocol's own signal, but nothing records
// what Fastmail returns for it. Cyrus reports every component type for a collection that stores
// no set, so a task list created without one would advertise VEVENT and the property would not
// tell it apart. Only a measurement settles which.
//
// For each calendar collection in the calendar home it reports: whether the property came back
// in a 200 propstat, the component names it lists, and whether the display name equals the
// literal `DEFAULT_TASK_CALENDAR_NAME`. For each non-calendar entry it reports the local names of
// its `resourcetype` children and the same two facts. It counts the display names it read and
// runs the name comparison on a synthetic entry, so `taskNamed=false` cannot be a parse failure,
// and it counts the calendar homes the principal names, which bounds what it looked at. Raw
// PROPFIND over bare `fetch`, so it sees the answer before tsdav's own component filter or its
// flattening of absent and empty into `[]`.
//
// READ-ONLY: PROPFIND only (principal, calendar home, then the home at Depth 1). Nothing is
// created, changed or deleted.
//
// OUTPUT IS SAFE TO QUOTE PUBLICLY. Names, URLs and hrefs are held in memory only. It prints
// counts, an index per collection, booleans, the component names (protocol constants; any name
// that is not a plain token prints as `<non-token>`), and on a refused request the HTTP status.
//
// Run: python scripts/probes/run-probe.py calendar-component-set.probe.mjs

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-component-set.probe.mjs');
  process.exit(1);
}

const ROOT = 'https://caldav.fastmail.com/dav/';
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');
const TASK_NAME = 'DEFAULT_TASK_CALENDAR_NAME';

async function propfind(url, props, depth) {
  const body = `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">`
    + `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;
  const res = await fetch(url, {
    method: 'PROPFIND',
    headers: { Authorization: AUTH, 'Content-Type': 'text/xml;charset=UTF-8', Depth: depth },
    body,
  });
  return { status: res.status, text: await res.text() };
}

// Any namespace prefix: Cyrus picks its own.
const elRe = name => new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`, 'g');
const el = (xml, name) => [...xml.matchAll(elRe(name))][0]?.[1];
const elAll = (xml, name) => [...xml.matchAll(elRe(name))].map(m => m[1]);
const has = (xml, name) => new RegExp(`<(?:[\\w-]+:)?${name}[\\s>/]`).test(xml);
const abs = href => new URL(href.trim(), ROOT).href;
const unescapeXml = s => s
  .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');

function refused(step, status) {
  console.log(`REFUSED  ${step}: HTTP ${status}`);
  process.exit(1);
}

const TOKEN = /^[A-Za-z][A-Za-z0-9-]{0,40}$/;
const token = s => (TOKEN.test(s) ? s : '<non-token>');

/** One `<response>` block, read from its 200 propstats only. Nothing here is printed raw. */
function parseEntry(response) {
  const entry = { types: [], name: undefined, setPresent: false, comps: [] };
  for (const stat of elAll(response, 'propstat')) {
    if (!/\s200\s/.test(` ${(el(stat, 'status') ?? '').trim()} `)) continue;
    const prop = el(stat, 'prop') ?? '';
    const rt = el(prop, 'resourcetype');
    if (rt !== undefined) {
      entry.types = [...rt.matchAll(/<(?:[\w-]+:)?([\w-]+)\b[^>]*?\/?>/g)].map(m => token(m[1]));
    }
    const dn = el(prop, 'displayname');
    if (dn !== undefined) entry.name = unescapeXml(dn).trim();
    if (has(prop, 'supported-calendar-component-set')) {
      entry.setPresent = true;
      const set = el(prop, 'supported-calendar-component-set') ?? '';
      entry.comps = [...set.matchAll(/<(?:[\w-]+:)?comp\b[^>]*?\bname\s*=\s*"([^"]*)"/g)].map(m => token(m[1]));
    }
  }
  entry.isCalendar = entry.types.includes('calendar');
  entry.taskNamed = entry.name === TASK_NAME;
  return entry;
}

// The comparison, proved on a synthetic block shaped like the server's (a prefixed displayname
// in a 200 propstat), so a false below means the name differs rather than that it went unread.
const selfCheck = parseEntry(
  `<D:response><D:href>/x/</D:href><D:propstat><D:prop><D:displayname>${TASK_NAME}</D:displayname>`
  + `</D:prop><D:status>HTTP/1.1 200 OK</D:status></D:propstat></D:response>`,
).taskNamed;

const rootPf = await propfind(ROOT, ['d:current-user-principal'], '0');
const principal = el(el(rootPf.text, 'current-user-principal') ?? '', 'href');
if (!principal) refused('current-user-principal', rootPf.status);

const prinPf = await propfind(abs(principal), ['c:calendar-home-set'], '0');
const homeHrefs = [...new Set(elAll(el(prinPf.text, 'calendar-home-set') ?? '', 'href').map(abs))];
if (homeHrefs.length === 0) refused('calendar-home-set', prinPf.status);
const homeUrl = homeHrefs[0];

const homePf = await propfind(homeUrl, ['d:resourcetype', 'd:displayname', 'c:supported-calendar-component-set'], '1');
if (homePf.status !== 207) refused('calendar home listing', homePf.status);

const entries = [];
for (const response of elAll(homePf.text, 'response')) {
  const href = el(response, 'href');
  if (!href || abs(href) === homeUrl) continue;
  entries.push(parseEntry(response));
}
const calendars = entries.filter(e => e.isCalendar);
const others = entries.filter(e => !e.isCalendar);
const hasVevent = e => e.comps.some(c => c.toUpperCase() === 'VEVENT');

console.log(`calendar homes named by the principal: ${homeHrefs.length}; listed: 1, by PROPFIND Depth 1 on that home only`);
console.log(`self-check: the comparison matches a synthetic ${TASK_NAME} entry: ${selfCheck}`);
console.log(`display names read: ${entries.filter(e => e.name).length} of ${entries.length}`
  + ` (calendars ${calendars.filter(e => e.name).length} of ${calendars.length},`
  + ` other entries ${others.filter(e => e.name).length} of ${others.length})`);
console.log(`calendar collections: ${calendars.length} (other non-calendar entries in the home: ${others.length})`);
calendars.forEach((e, i) => {
  console.log(`#${i + 1}  taskNamed=${e.taskNamed}  setPresent=${e.setPresent}  components=[${e.comps.join(',')}]`);
});
others.forEach((e, i) => {
  console.log(`other #${i + 1}  resourcetype=[${e.types.join(',')}]  taskNamed=${e.taskNamed}`
    + `  setPresent=${e.setPresent}${e.setPresent ? `  components=[${e.comps.join(',')}]` : ''}`);
});
console.log(`with VEVENT: ${calendars.filter(e => e.setPresent && hasVevent(e)).length}`
  + `  set without VEVENT: ${calendars.filter(e => e.setPresent && !hasVevent(e)).length}`
  + `  no set: ${calendars.filter(e => !e.setPresent).length}`
  + `  task-named (all entries): ${entries.filter(e => e.taskNamed).length}`);
