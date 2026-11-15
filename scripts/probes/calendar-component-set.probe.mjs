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
// literal `DEFAULT_TASK_CALENDAR_NAME`. For each non-calendar entry it reports the names of
// `resourcetype`'s direct children and the same two facts. It counts the display names it read
// and runs the name comparison on a synthetic entry: together those show a false `taskNamed` is
// a name that differs, not a name that went unread or a comparison that cannot match. It counts
// the calendar homes the principal names, which bounds what it looked at. Raw PROPFIND over bare
// `fetch`, so it sees the answer before tsdav's own component filter or its flattening of absent
// and empty into `[]`.
//
// READ-ONLY: PROPFIND only (principal, calendar home, then the home at Depth 1). Nothing is
// created, changed or deleted. Redirects are not followed, and credentials go only to ROOT's
// origin: a 3xx or an href on another origin is refused.
//
// OUTPUT IS SAFE TO QUOTE PUBLICLY. Names, URLs and hrefs are held in memory only. It prints
// counts, an index per entry, booleans, component and resourcetype names from the allowlists
// below (anything else prints as `<other>`, and is counted), the HTTP status of a refused
// request, and for an unexpected error its name and code, never its message.
//
// LIMITS: elements are matched by local name whatever their namespace, and hex character
// references (`&#x…;`) in a display name are not decoded, so a name written that way would not
// compare equal.
//
// Exits non-zero on a refused request, when the self-check fails, or when no display name is read.
//
// Run: python scripts/probes/run-probe.py calendar-component-set.probe.mjs

import { pathToFileURL } from 'node:url';
import { makeChecker } from './probelib.mjs';

const ROOT = 'https://caldav.fastmail.com/dav/';
const ORIGIN = new URL(ROOT).origin;
const TASK_NAME = 'DEFAULT_TASK_CALENDAR_NAME';

// RFC 5545 section 3.6, RFC 7953 (VAVAILABILITY), and VPOLL from the VPOLL draft that Cyrus
// implements.
export const COMPONENTS = new Set(['VEVENT', 'VTODO', 'VJOURNAL', 'VFREEBUSY', 'VTIMEZONE', 'VAVAILABILITY', 'VPOLL']);
// RFC 4918 (collection), RFC 4791 (calendar), RFC 6638 (schedule-inbox, schedule-outbox),
// RFC 3744 (principal), RFC 6352 (addressbook), and CalendarServer's notification collection.
export const RESOURCE_TYPES = new Set(['collection', 'calendar', 'schedule-inbox', 'schedule-outbox', 'principal', 'addressbook', 'notification']);

let otherNames = 0;
export const allow = (set, name) => {
  if (set.has(name)) return name;
  otherNames++;
  return '<other>';
};

// Any namespace prefix: Cyrus picks its own.
const elRe = name => new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`, 'g');
export const el = (xml, name) => [...xml.matchAll(elRe(name))][0]?.[1];
export const elAll = (xml, name) => [...xml.matchAll(elRe(name))].map(m => m[1]);
export const has = (xml, name) => new RegExp(`<(?:[\\w-]+:)?${name}[\\s>/]`).test(xml);
const unescapeXml = s => s
  .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');

/** Local names of the direct child elements of an element's content. */
export function directChildren(xml) {
  const names = [];
  let depth = 0;
  for (const m of xml.matchAll(/<(\/?)(?:[\w-]+:)?([\w-]+)\b[^>]*?(\/?)>/g)) {
    const [, closing, name, selfClosing] = m;
    if (closing) { depth--; continue; }
    if (depth === 0) names.push(name);
    if (!selfClosing) depth++;
  }
  return names;
}

/** The `name` attribute of each `comp` element, single- or double-quoted, matched exactly. */
export function compNames(setXml) {
  return [...setXml.matchAll(/<(?:[\w-]+:)?comp\b([^>]*)>/g)]
    .map(m => m[1].match(/(?:^|\s)name\s*=\s*(?:"([^"]*)"|'([^']*)')/))
    .filter(Boolean)
    .map(a => a[1] ?? a[2]);
}

/** An absolute URL from an href, or a thrown error whose code alone is printed. */
export function abs(href) {
  return new URL(href.trim(), ROOT);
}

/** An href credentials may be sent to, or the line to refuse it with; neither carries the href. */
export function resolveHref(href) {
  let url;
  try {
    url = abs(href);
  } catch (e) {
    return { refused: `href: ${e?.code ?? 'unparseable'}` };
  }
  return url.origin === ORIGIN ? { url } : { refused: 'href: on another origin' };
}

/** Origin, percent-decoded path, no trailing slash: two spellings of one collection compare equal. */
export function sameResource(a, b) {
  const key = u => {
    let path = u.pathname;
    try { path = decodeURIComponent(path); } catch { /* keep it encoded */ }
    return `${u.origin}${path.replace(/\/+$/, '')}`;
  };
  return key(a) === key(b);
}

/** One `<response>` block, read from its 200 propstats only. Nothing here is printed raw. */
export function parseEntry(response) {
  const entry = { types: [], name: undefined, setPresent: false, comps: [] };
  for (const stat of elAll(response, 'propstat')) {
    if (!/\s200\s/.test(` ${(el(stat, 'status') ?? '').trim()} `)) continue;
    const prop = el(stat, 'prop') ?? '';
    const rt = el(prop, 'resourcetype');
    if (rt !== undefined) entry.types = directChildren(rt);
    const dn = el(prop, 'displayname');
    if (dn !== undefined) entry.name = unescapeXml(dn).trim();
    if (has(prop, 'supported-calendar-component-set')) {
      entry.setPresent = true;
      entry.comps = compNames(el(prop, 'supported-calendar-component-set') ?? '');
    }
  }
  entry.isCalendar = entry.types.includes('calendar');
  entry.taskNamed = entry.name === TASK_NAME;
  return entry;
}

class Refused extends Error {}

function refuse(line) {
  console.log(`REFUSED  ${line}`);
  throw new Refused();
}

async function main() {
  const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
  const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;
  if (!USERNAME || !PASSWORD) {
    console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
    console.error('      Run through: python scripts/probes/run-probe.py calendar-component-set.probe.mjs');
    process.exit(1);
  }
  const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');

  const resolve = href => {
    const r = resolveHref(href);
    if (r.refused) refuse(r.refused);
    return r.url;
  };

  async function propfind(step, url, props, depth) {
    const body = `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">`
      + `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;
    const res = await fetch(url, {
      method: 'PROPFIND',
      redirect: 'manual',
      headers: { Authorization: AUTH, 'Content-Type': 'text/xml;charset=UTF-8', Depth: depth },
      body,
    });
    if (res.status !== 207) refuse(`${step}: HTTP ${res.status}`);
    return res.text();
  }

  const { check, failures } = makeChecker();

  const selfCheck = parseEntry(
    `<D:response><D:href>/x/</D:href><D:propstat><D:prop><D:displayname>${TASK_NAME}</D:displayname>`
    + `</D:prop><D:status>HTTP/1.1 200 OK</D:status></D:propstat></D:response>`,
  ).taskNamed;

  const rootXml = await propfind('current-user-principal', new URL(ROOT), ['d:current-user-principal'], '0');
  const principal = el(el(rootXml, 'current-user-principal') ?? '', 'href');
  if (!principal) refuse('current-user-principal: absent');

  const prinXml = await propfind('calendar-home-set', resolve(principal), ['c:calendar-home-set'], '0');
  const homes = [];
  for (const href of elAll(el(prinXml, 'calendar-home-set') ?? '', 'href')) {
    const url = resolve(href);
    if (!homes.some(h => sameResource(h, url))) homes.push(url);
  }
  if (homes.length === 0) refuse('calendar-home-set: absent');
  const homeUrl = homes[0];

  const homeXml = await propfind('calendar home listing', homeUrl,
    ['d:resourcetype', 'd:displayname', 'c:supported-calendar-component-set'], '1');

  const entries = [];
  for (const response of elAll(homeXml, 'response')) {
    const href = el(response, 'href');
    if (!href || sameResource(resolve(href), homeUrl)) continue;
    entries.push(parseEntry(response));
  }
  const calendars = entries.filter(e => e.isCalendar);
  const others = entries.filter(e => !e.isCalendar);
  const hasVevent = e => e.comps.some(c => c.toUpperCase() === 'VEVENT');
  const comps = e => e.comps.map(c => allow(COMPONENTS, c)).join(',');
  const types = e => e.types.map(t => allow(RESOURCE_TYPES, t)).join(',');
  const namesRead = entries.filter(e => e.name).length;

  console.log(`calendar homes named by the principal: ${homes.length}; listed: 1, by PROPFIND Depth 1 on that home only`);
  console.log(`self-check: the comparison matches a synthetic ${TASK_NAME} entry: ${selfCheck}`);
  console.log(`display names read: ${namesRead} of ${entries.length}`
    + ` (calendars ${calendars.filter(e => e.name).length} of ${calendars.length},`
    + ` other entries ${others.filter(e => e.name).length} of ${others.length})`);
  console.log(`calendar collections: ${calendars.length} (other non-calendar entries in the home: ${others.length})`);
  calendars.forEach((e, i) => {
    console.log(`#${i + 1}  taskNamed=${e.taskNamed}  setPresent=${e.setPresent}  components=[${comps(e)}]`);
  });
  others.forEach((e, i) => {
    console.log(`other #${i + 1}  resourcetype=[${types(e)}]  taskNamed=${e.taskNamed}`
      + `  setPresent=${e.setPresent}${e.setPresent ? `  components=[${comps(e)}]` : ''}`);
  });
  console.log(`with VEVENT: ${calendars.filter(e => e.setPresent && hasVevent(e)).length}`
    + `  set without VEVENT: ${calendars.filter(e => e.setPresent && !hasVevent(e)).length}`
    + `  no set: ${calendars.filter(e => !e.setPresent).length}`
    + `  task-named (all entries): ${entries.filter(e => e.taskNamed).length}`);
  console.log(`names outside the allowlists, printed as <other>: ${otherNames}`);

  check('the display-name comparison matches a synthetic task-named entry', selfCheck);
  check('at least one display name was read', namesRead > 0);
  if (failures() > 0) process.exitCode = 1;
}

// Imported for an offline check, it runs nothing.
if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  main().catch(e => {
    if (!(e instanceof Refused)) console.log(`FAIL  unexpected error: ${e?.name ?? 'Error'}${e?.code ? ` ${e.code}` : ''}`);
    process.exit(1);
  });
}
