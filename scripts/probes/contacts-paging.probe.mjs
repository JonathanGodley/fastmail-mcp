// `list_contacts` and `search_contacts` send `sort: [name/given, name/surname, uid]`, all
// ascending, and page with `position` (#94). Cyrus's source accepts those three comparators and compares them byte-wise
// with an absent component read as the empty string, which would put nameless cards first
// and make the order case-sensitive. Whether Fastmail runs that code cannot be read off the
// source, so this measures it:
//
//   (a) the sort is accepted: no method error, no unsupportedSort;
//   (b) `position` is honoured: each page starts where the last ended, pages do not
//       overlap, and their union equals one large query's ids in the same order;
//   (c) where cards with no given name sit in that order;
//   (d) whether the order is byte-wise (case-sensitive) or case-insensitive.
//
// Raw JMAP, not the built server, so it measures the platform rather than our handlers.
// READ-ONLY: ContactCard/query and ContactCard/get only (names and uids, compared in
// memory). It prints counts, booleans and positions only - never a name, address, uid or
// id - so a run can be quoted into a public issue or commit.

import { getSession, jmap } from './jmaplib.mjs';
import { makeChecker } from './probelib.mjs';

const CONTACTS = 'urn:ietf:params:jmap:contacts';
const SORT = [
  { property: 'name/given', isAscending: true },
  { property: 'name/surname', isAscending: true },
  { property: 'uid', isAscending: true },
];
const PAGE = 50;
const LARGE = 5000;
const { check, failures } = makeChecker();

const s = await getSession();
const accountId = s.contactsAccountId;
if (!accountId) {
  console.error('The session has no contacts primary account. Stopping.');
  process.exit(2);
}

// The error TYPE only: a server description could quote account data.
function only(body) {
  const [name, args] = body.methodResponses[0];
  if (name === 'error') return { error: args.type ?? 'unknown' };
  return args;
}

async function query(position, limit) {
  const args = { accountId, sort: SORT, limit, calculateTotal: true };
  if (position !== undefined) args.position = position;
  return only(await jmap(s, [['ContactCard/query', args, 'q']], [CONTACTS]));
}

// (a)
const large = await query(undefined, LARGE);
check('(a) the sort is accepted', !large.error, large.error ? `error type ${large.error}` : '');
if (large.error) process.exit(1);
const total = large.total;
console.log(`total ${total}; large query returned ${large.ids.length} ids at position ${large.position}`);
check('(a) one large query reaches every card', large.ids.length === total);

const again = await query(undefined, LARGE);
check('(b) two identical large queries return the same order',
  !again.error && again.ids.length === large.ids.length && again.ids.every((id, i) => id === large.ids[i]));

// (b)
const union = [];
let pageCount = 0;
let positionMismatches = 0;
for (let position = 0; position < total; position += PAGE) {
  const page = await query(position, PAGE);
  if (page.error) {
    check(`(b) page at position ${position}`, false, `error type ${page.error}`);
    break;
  }
  pageCount++;
  if (page.position !== position) positionMismatches++;
  console.log(`page ${pageCount}: requested position ${position}, served position ${page.position}, ${page.ids.length} ids`);
  union.push(...page.ids);
}
const overlap = union.length - new Set(union).size;
check('(b) every page is served at the position requested', positionMismatches === 0, `${positionMismatches} mismatches`);
check('(b) pages do not overlap', overlap === 0, `${overlap} repeated ids`);
check('(b) the union of the pages equals the large query, in order',
  union.length === large.ids.length && union.every((id, i) => id === large.ids[i]),
  `${union.length} paged ids vs ${large.ids.length}`);

// (c) and (d): the names, compared in memory in the order the server returned.
const cards = new Map();
for (let i = 0; i < large.ids.length; i += 500) {
  const got = only(await jmap(s, [['ContactCard/get', {
    accountId, ids: large.ids.slice(i, i + 500), properties: ['name', 'uid'],
  }, 'g']], [CONTACTS]));
  if (got.error) {
    check('(c) ContactCard/get', false, `error type ${got.error}`);
    process.exit(1);
  }
  for (const card of got.list) cards.set(card.id, card);
}

// A name component as Cyrus's sort reads it: sortAs first, else the matching components
// joined by the separator.
function component(card, kind) {
  const name = card.name;
  if (!name) return '';
  if (typeof name.sortAs?.[kind] === 'string') return name.sortAs[kind];
  const defsep = typeof name.defaultSeparator === 'string' ? name.defaultSeparator : ' ';
  let sep = defsep;
  let out = '';
  for (const c of name.components ?? []) {
    if (c.kind === 'separator') { sep = c.value; continue; }
    if (c.kind === kind) out += (out ? sep : '') + c.value;
    sep = defsep;
  }
  return out;
}

const keys = large.ids.map((id) => {
  const card = cards.get(id) ?? {};
  return [component(card, 'given'), component(card, 'surname'), card.uid ?? ''];
});
const bytes = (a, b) => Buffer.compare(Buffer.from(a, 'utf8'), Buffer.from(b, 'utf8'));
const tuple = (a, b, cmp) => {
  for (let i = 0; i < 3; i++) {
    const r = cmp(a[i], b[i]);
    if (r) return r;
  }
  return 0;
};
const folded = (a, b) => bytes(a.toLowerCase(), b.toLowerCase());

const nameless = keys.map((k, i) => (k[0] === '' ? i : -1)).filter((i) => i >= 0);
const contiguous = nameless.every((p, i) => i === 0 || p === nameless[i - 1] + 1);
console.log(`(c) ${nameless.length} of ${keys.length} cards have no given name`);
console.log(`(c) nameless cards contiguous: ${contiguous}; at the start: ${nameless.length > 0 && contiguous && nameless[0] === 0}; at the end: ${nameless.length > 0 && contiguous && nameless[nameless.length - 1] === keys.length - 1}`);
if (nameless.length) console.log(`(c) nameless cards span positions ${nameless[0]} to ${nameless[nameless.length - 1]}`);

let byteOutOfOrder = 0;
let byteOnly = 0;
let foldedOnly = 0;
for (let i = 1; i < keys.length; i++) {
  const b = tuple(keys[i - 1], keys[i], bytes);
  const f = tuple(keys[i - 1], keys[i], folded);
  if (b > 0) byteOutOfOrder++;
  if (b <= 0 && f > 0) byteOnly++;
  if (f <= 0 && b > 0) foldedOnly++;
}
console.log(`(d) adjacent pairs: ${keys.length - 1}`);
console.log(`(d) pairs out of byte order: ${byteOutOfOrder}`);
console.log(`(d) pairs in byte order but out of case-insensitive order: ${byteOnly}`);
console.log(`(d) pairs in case-insensitive order but out of byte order: ${foldedOnly}`);
check('(d) the order is byte-wise over given name, surname, uid', byteOutOfOrder === 0);

console.log(failures() ? `${failures()} check(s) failed` : 'all checks passed');
process.exit(failures() ? 1 : 0);
