// Which Email/query sort properties does Fastmail accept (#223)? RFC 8621 section 4.4.2
// lists `receivedAt`, `size`, `from`, `to`, `subject` and `sentAt`, but support is per
// property: the #35 snooze probe got `unsupportedSort` for `snoozedUntil`. For each property,
// ascending and descending:
//
//   (a) accepted, or the error type;
//   (b) the order is what the property says, checked in memory over the first page
//       (receivedAt, sentAt and size exactly; from, to and subject against candidate keys,
//       reported as which key the order agrees with);
//   (c) `position` paging is stable: the same query twice returns the same ids, and a
//       second page at position 10 equals ids 10-19 of a 20-id first query.
//
// Cyrus's `_email_buildsort` (imap/jmap_mail.c) maps `from`/`to` to the RFC 5957 display
// sorts (display name, else address, under a server collation the candidate keys here only
// approximate) and appends receivedAt descending, then id descending, as tie-breaks; the
// in-memory key check is therefore informational for the address and subject sorts.
//
// Also measured: a two-key sort (property, then receivedAt) as a tie-break, and a control
// that names a property RFC 8621 does not define, to record the refusal's error type.
//
// Raw JMAP, not the built server. READ-ONLY: Email/query (limit 20 at most) and Email/get
// restricted to the one property under test. Values are compared in memory and never
// printed: output is accepted-or-error-type, counts and booleans only, so a run can be
// quoted into a public issue or commit.

import { getSession, jmap } from './jmaplib.mjs';
import { makeChecker } from './probelib.mjs';

const { check, failures } = makeChecker();
const s = await getSession();
const PROPS = ['receivedAt', 'size', 'from', 'to', 'subject', 'sentAt'];

async function query(sort, { position = 0, limit = 20, withGet = null } = {}) {
  const calls = [['Email/query', { accountId: s.accountId, sort, position, limit, calculateTotal: true }, 'q']];
  if (withGet) calls.push(['Email/get', { accountId: s.accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: ['id', withGet] }, 'g']);
  const body = await jmap(s, calls);
  const [name, args] = body.methodResponses[0];
  if (name === 'error') return { ok: false, type: args.type }; // the TYPE only
  const got = withGet ? body.methodResponses[1] : null;
  const list = got && got[0] === 'Email/get' ? got[1].list : null;
  return { ok: true, ids: args.ids, total: args.total, position: args.position, list };
}

const sameIds = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
const firstAddr = v => (Array.isArray(v) && v.length ? v[0] : null);
// Cyrus's comparator (imap/index.c): `get_displayname` takes the FIRST address's display
// name in RFC 5255 search form (uppercased, diacritics stripped), else MAILBOX@DOMAIN
// uppercased, and `strcmpsafe` compares the results byte-wise. The search form is
// approximated here (NFD, combining marks dropped, whitespace collapsed, uppercased).
const searchForm = s => s.normalize('NFD').replace(/\p{M}/gu, '').replace(/\s+/g, ' ').trim().toUpperCase();
const displayKey = v => {
  const a = firstAddr(v);
  if (!a) return '';
  return a.name && a.name.trim() ? searchForm(a.name) : (a.email ?? '').toUpperCase();
};
// RFC 5256 base subject (index_extract_subject): leading Re:/Fwd:/[tag] and trailing
// "(fwd)" stripped, then uppercased.
const baseSubject = v => (v ?? '').replace(/\s+/g, ' ').trim()
  .replace(/(\s*\(fwd\))+$/i, '')
  .replace(/^((re|fwd?|fw)\s*(\[[^\]]*\])?\s*:\s*|\[[^\]]*\]\s*)+/i, '').trim().toUpperCase();
const ADDR_KEYS = { 'display name (search form), else ADDRESS': displayKey };
const SUBJECT_KEYS = { 'base subject, uppercased': baseSubject };
const EXACT = { receivedAt: v => Date.parse(v), sentAt: v => Date.parse(v), size: v => v };

const byteCmp = (a, b) => {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return Buffer.compare(x, y);
};
function ordered(values, key, asc) {
  const k = values.map(key);
  for (let i = 1; i < k.length; i++) {
    const c = typeof k[i] === 'number' ? Math.sign(k[i - 1] - k[i]) : byteCmp(k[i - 1], k[i]);
    if (asc ? c > 0 : c < 0) return false;
  }
  return true;
}
// Collation-independent: once a key has been left, it never comes back (equal keys sit
// together), which is the "group by" property a caller relies on.
function contiguous(values, key) {
  const seen = new Set();
  let prev;
  for (const k of values.map(key)) {
    if (k !== prev && seen.has(k)) return false;
    seen.add(k);
    prev = k;
  }
  return true;
}

const accepted = [];
for (const prop of PROPS) {
  for (const asc of [true, false]) {
    const dir = asc ? 'ascending' : 'descending';
    const sort = [{ property: prop, isAscending: asc }];
    const first = await query(sort, { withGet: prop });
    check(`${prop} ${dir}: accepted`, first.ok, first.ok ? `${first.ids.length} ids, total ${first.total}` : `error type ${first.type}`);
    if (!first.ok) continue;
    accepted.push(`${prop} ${dir}`);

    const byId = new Map((first.list ?? []).map(e => [e.id, e[prop]]));
    const values = first.ids.map(id => byId.get(id));
    check(`${prop} ${dir}: Email/get returned every id`, values.length === first.ids.length && first.ids.every(id => byId.has(id)));
    if (EXACT[prop]) {
      check(`${prop} ${dir}: ordering verified`, ordered(values, EXACT[prop], asc));
    } else {
      const [[keyName, key]] = Object.entries(prop === 'subject' ? SUBJECT_KEYS : ADDR_KEYS);
      // A page of one repeated key verifies nothing, so look further into the set.
      const pages = [{ at: 0, values }];
      for (const at of [Math.floor(first.total / 3), Math.floor((2 * first.total) / 3)]) {
        const more = await query(sort, { position: at, withGet: prop });
        if (!more.ok || !more.list) continue;
        const m = new Map(more.list.map(e => [e.id, e[prop]]));
        pages.push({ at, values: more.ids.map(id => m.get(id)) });
      }
      for (const p of pages) {
        const distinct = new Set(p.values.map(key)).size;
        check(`${prop} ${dir} at position ${p.at}: ordering verified against ${keyName}`, ordered(p.values, key, asc), `${distinct} distinct keys in ${p.values.length}`);
        check(`${prop} ${dir} at position ${p.at}: equal keys are contiguous`, contiguous(p.values, key));
      }
    }

    const again = await query(sort);
    check(`${prop} ${dir}: same query twice, same ids`, again.ok && sameIds(again.ids, first.ids));
    const page2 = await query(sort, { position: 10, limit: 10 });
    check(`${prop} ${dir}: page at position 10 = ids 10-19 of the first query`,
      page2.ok && page2.position === 10 && sameIds(page2.ids, first.ids.slice(10, 20)));
  }
}

// Ties: a two-key sort with receivedAt as the tie-break.
for (const prop of ['size', 'from', 'subject']) {
  const two = await query([{ property: prop, isAscending: false }, { property: 'receivedAt', isAscending: false }]);
  check(`two-key sort [${prop} desc, receivedAt desc]: accepted`, two.ok, two.ok ? `${two.ids.length} ids` : `error type ${two.type}`);
}

// Control: a property RFC 8621 does not define.
const bogus = await query([{ property: 'notARealSortProperty', isAscending: true }]);
console.log(`INFO undefined sort property: ${bogus.ok ? 'accepted (unexpected)' : `error type ${bogus.type}`}`);

console.log(`accepted: ${accepted.length}/${PROPS.length * 2}`);
console.log(failures() ? `SORT-PROPERTIES: ${failures()} FAILED` : 'SORT-PROPERTIES: ALL PASS');
process.exit(failures() ? 1 : 0);
