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
// literal `DEFAULT_TASK_CALENDAR_NAME`. Raw PROPFIND over bare `fetch`, so it sees the answer
// before tsdav's own component filter or its flattening of absent and empty into `[]`.
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

const rootPf = await propfind(ROOT, ['d:current-user-principal'], '0');
const principal = el(el(rootPf.text, 'current-user-principal') ?? '', 'href');
if (!principal) refused('current-user-principal', rootPf.status);

const prinPf = await propfind(abs(principal), ['c:calendar-home-set'], '0');
const home = el(el(prinPf.text, 'calendar-home-set') ?? '', 'href');
if (!home) refused('calendar-home-set', prinPf.status);
const homeUrl = abs(home);

const homePf = await propfind(homeUrl, ['d:resourcetype', 'd:displayname', 'c:supported-calendar-component-set'], '1');
if (homePf.status !== 207) refused('calendar home listing', homePf.status);

const rows = [];
let otherCollections = 0;
for (const response of elAll(homePf.text, 'response')) {
  const href = el(response, 'href');
  if (!href || abs(href) === homeUrl) continue;
  let isCalendar = false;
  let name;
  let setPresent = false;
  let comps = [];
  for (const stat of elAll(response, 'propstat')) {
    const ok = /\s200\s/.test(` ${(el(stat, 'status') ?? '').trim()} `);
    const prop = el(stat, 'prop') ?? '';
    if (!ok) continue;
    const rt = el(prop, 'resourcetype');
    if (rt !== undefined && has(rt, 'calendar')) isCalendar = true;
    const dn = el(prop, 'displayname');
    if (dn !== undefined) name = unescapeXml(dn).trim();
    if (has(prop, 'supported-calendar-component-set')) {
      setPresent = true;
      const set = el(prop, 'supported-calendar-component-set') ?? '';
      comps = [...set.matchAll(/<(?:[\w-]+:)?comp\b[^>]*?\bname\s*=\s*"([^"]*)"/g)]
        .map(m => (/^[A-Za-z][A-Za-z0-9-]{0,40}$/.test(m[1]) ? m[1] : '<non-token>'));
    }
  }
  if (!isCalendar) { otherCollections++; continue; }
  rows.push({ taskNamed: name === TASK_NAME, setPresent, comps });
}

const hasVevent = r => r.comps.some(c => c.toUpperCase() === 'VEVENT');
console.log(`calendar collections: ${rows.length} (other non-calendar entries in the home: ${otherCollections})`);
rows.forEach((r, i) => {
  console.log(`#${i + 1}  taskNamed=${r.taskNamed}  setPresent=${r.setPresent}  components=[${r.comps.join(',')}]`);
});
console.log(`with VEVENT: ${rows.filter(r => r.setPresent && hasVevent(r)).length}`
  + `  set without VEVENT: ${rows.filter(r => r.setPresent && !hasVevent(r)).length}`
  + `  no set: ${rows.filter(r => !r.setPresent).length}`
  + `  task-named: ${rows.filter(r => r.taskNamed).length}`);
