// What this probe settles
// -----------------------
// What Fastmail's OWN clients write on the wire when a user authors a calendar event —
// the raw stored iCalendar, fetched back untouched. docs/fastmail-action-availability.md
// is extended by measurement, never by inference from a role's name, and #165 lists the
// shapes still unmeasured (UNTIL and BYDAY recurrence forms, a DST-spanning all-day
// event, a whole-series edit). The workflow: the operator authors events in a Fastmail
// client with distinctive titles, then this dumps each matching resource's calendar-data
// exactly as stored, so the client-authored byte-shape can be recorded.
//
// Usage (via the token launcher, which injects the CalDAV app password):
//
//   python scripts/probes/run-probe.py client-authored-events.probe.mjs [title-substring ...]
//
// Any event whose SUMMARY contains one of the given substrings is dumped; with no
// arguments it matches the reference-event set authored 22 Aug 2026 ("A timed event
// where you ..."). The query window is relative — 30 days back to 120 days ahead — so
// freshly authored events are always in range.
//
// Raw CalDAV over bare fetch (PROPFIND discovery → calendar-query REPORT), not the built
// server and not tsdav: the point is what the client wrote, with none of our parsing in
// the way. Read-only; creates and deletes nothing. Output redacts the account name and
// every email-shaped string, and unwraps the CDATA the server wraps calendar-data and
// display names in, so the bytes printed are the bytes stored.

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;
if (!USERNAME || !PASSWORD) {
  console.error('FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  process.exit(1);
}
const AUTH = 'Basic ' + Buffer.from(`${USERNAME}:${PASSWORD}`).toString('base64');
const ROOT = 'https://caldav.fastmail.com/dav/';

// Redact the account name and any email-shaped string in everything printed.
const redact = s => String(s)
  .split(USERNAME).join('<account>')
  .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');

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
  return { status: res.status, text: await res.text() };
}

const el = (xml, name) => {
  const m = xml.match(new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`));
  return m ? m[1] : undefined;
};
const elAll = (xml, name) =>
  [...xml.matchAll(new RegExp(`<(?:[\\w-]+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w-]+:)?${name}>`, 'g'))].map(m => m[1]);
const responses = xml => elAll(xml, 'response');
const abs = href => new URL(href, ROOT).href;
const unescapeXml = s => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');

// XML text that may carry CDATA sections. Cyrus wraps calendar-data and display names in them,
// so a plain unescape leaves the `<![CDATA[` … `]]>` markers sitting in the printed bytes.
//
// The unwrap has to be POSITIONAL. Inside a section the text is literal (no entities); outside
// it, the ordinary escaping applies. So each section's contents are spliced back in exactly
// where it sat and only the text around it is unescaped. Returning just the sections' contents
// would silently drop the non-CDATA text of a mixed payload, and unescaping the whole string
// would corrupt a literal `&amp;` that the CDATA existed to protect.
const decodeXmlText = s => {
  let out = '';
  let i = 0;
  for (;;) {
    const start = s.indexOf('<![CDATA[', i);
    if (start < 0) return out + unescapeXml(s.slice(i));
    const end = s.indexOf(']]>', start + 9);
    // Unterminated section: nothing reliable left to splice, so treat the rest as ordinary text
    // rather than dropping it.
    if (end < 0) return out + unescapeXml(s.slice(i));
    out += unescapeXml(s.slice(i, start)) + s.slice(start + 9, end);
    i = end + 3;
  }
};

const PROPFIND = props =>
  `<?xml version="1.0" encoding="utf-8"?>\n<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">` +
  `<d:prop>${props.map(p => `<${p}/>`).join('')}</d:prop></d:propfind>`;

const davInstant = iso => `${new Date(iso).toISOString().slice(0, 19).replace(/[-:.]/g, '')}Z`;
const queryXml = (s, e) =>
  `<?xml version="1.0" encoding="utf-8"?>\n` +
  `<c:calendar-query xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:d="DAV:">` +
  `<d:prop><d:getetag/><c:calendar-data/></d:prop>` +
  `<c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT">` +
  `<c:time-range start="${davInstant(s)}" end="${davInstant(e)}"/>` +
  `</c:comp-filter></c:comp-filter></c:filter></c:calendar-query>`;

// SUMMARY substrings to match: CLI arguments, or the 22 Aug 2026 reference set.
const WANTED = process.argv.slice(2).length ? process.argv.slice(2) : ['A timed event where you'];

// Window: 30 days back to 120 days ahead, so freshly authored events are in range.
const DAY = 24 * 60 * 60 * 1000;
const windowStart = new Date(Date.now() - 30 * DAY).toISOString();
const windowEnd = new Date(Date.now() + 120 * DAY).toISOString();

// Discover calendar home.
const rootPf = await dav('PROPFIND', ROOT, { body: PROPFIND(['d:current-user-principal']), headers: { Depth: '0' } });
const principal = el(el(rootPf.text, 'current-user-principal') ?? '', 'href');
const prinPf = await dav('PROPFIND', abs(principal.trim()), { body: PROPFIND(['c:calendar-home-set']), headers: { Depth: '0' } });
const home = abs(el(el(prinPf.text, 'calendar-home-set') ?? '', 'href').trim());

// Enumerate collections with displaynames.
const homePf = await dav('PROPFIND', home, {
  body: PROPFIND(['d:displayname', 'd:resourcetype']),
  headers: { Depth: '1' },
});
const collections = [];
for (const block of responses(homePf.text)) {
  const href = (el(block, 'href') ?? '').trim();
  if (!href || abs(href) === home) continue;
  if (!/calendar[\s/>]/.test(el(block, 'resourcetype') ?? '')) continue;
  collections.push({ url: abs(href), name: decodeXmlText(el(block, 'displayname') ?? '(unnamed)') });
}
console.log(`Calendar collections: ${collections.length}`);
for (const c of collections) console.log(`  - ${redact(c.name)}  ${redact(c.url)}`);
console.log(`Window: ${windowStart} .. ${windowEnd}`);
console.log(`Matching SUMMARY substrings: ${WANTED.map(w => JSON.stringify(w)).join(', ')}`);

// Query each collection and keep events whose SUMMARY matches.
const found = [];
for (const c of collections) {
  const res = await dav('REPORT', c.url, { body: queryXml(windowStart, windowEnd), headers: { Depth: '1' } });
  for (const block of responses(res.text)) {
    const data = decodeXmlText(el(block, 'calendar-data') ?? '');
    if (!data) continue;
    const summary = (data.match(/^SUMMARY.*$/m) ?? [''])[0];
    if (WANTED.some(w => summary.includes(w))) {
      found.push({ cal: c.name, href: (el(block, 'href') ?? '').trim(), data });
    }
  }
}

console.log(`\nMatched resources: ${found.length}\n`);
for (const f of found) {
  console.log('='.repeat(70));
  console.log(`calendar: ${redact(f.cal)}`);
  console.log(`resource: ${redact(f.href)}`);
  console.log('-'.repeat(70));
  console.log(redact(f.data.replace(/\r\n/g, '\n').trim()));
  console.log('');
}
