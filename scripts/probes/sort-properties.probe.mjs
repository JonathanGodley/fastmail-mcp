// Which Email/query sort properties does Fastmail accept (#223), and does the two-key form
// the server will send (the property, then receivedAt descending) order and page as
// intended? Support is per property: the #35 snooze probe got `unsupportedSort` for
// `snoozedUntil`. For each of receivedAt, size, from, to, subject and sentAt, ascending and
// descending:
//
//   (a) accepted, or the error type;
//   (b) the order is what the property says, checked in memory: exactly for the date and
//       size sorts; for from, to and subject against the key Cyrus's comparator builds
//       (imap/index.c: `get_displayname` gives the first address's display name in RFC 5255
//       search form, else MAILBOX@DOMAIN uppercased; `index_extract_subject` gives the base
//       subject with ASCII letters uppercased; both compared byte-wise), approximated here
//       and gated on, at
//       three positions across the set, with the distinct-key count printed. A message with
//       no value for the key is counted and reported, and the order is checked across the
//       values that are present;
//   (c) `position` paging is stable: the same query twice returns the same ids, and a
//       page at position 10 equals ids 10-19 of a 20-id query.
//
// Then, for every property but receivedAt, both directions, the two-key form: within each
// run of equal primary keys receivedAt is non-increasing, and (c) holds. A case whose
// sampled pages hold no ties says so rather than reporting the tie-break verified. The same
// two-key form is then run under four filter shapes the tools send (see below). Where the
// subject key rejects an adjacent pair, a DIAG line gives booleans about that pair only.
//
// Also a control: a property RFC 8621 does not define, to record the refusal's error type.
//
// Raw JMAP, not the built server. READ-ONLY: Email/query (limit 20 at most) and Email/get
// restricted to the property under test and receivedAt. Values are compared in memory and
// never printed, and a refused request reports only its HTTP status and error type: output
// is accepted-or-error-type, counts and booleans only, so a run can be quoted into a public
// issue or commit.

import { getSession, AUTH } from './jmaplib.mjs';
import { makeChecker } from './probelib.mjs';

const { check, failures } = makeChecker();
const s = await getSession();
const PROPS = ['receivedAt', 'size', 'from', 'to', 'subject', 'sentAt'];

// Not jmaplib's jmap(): on a non-OK status that throws with the response body in the message.
async function call(methodCalls) {
  const res = await fetch(s.apiUrl, {
    method: 'POST',
    headers: { ...AUTH, 'Content-Type': 'application/json' },
    body: JSON.stringify({ using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'], methodCalls }),
  });
  if (res.status === 401) { console.error('AUTH FAILED: 401 from the JMAP API. Stopping.'); process.exit(3); }
  let body = null;
  try { body = await res.json(); } catch { /* reported as a status below */ }
  if (!res.ok || !body) return { ok: false, type: `HTTP ${res.status}${typeof body?.type === 'string' ? ` ${body.type}` : ''}` };
  return { ok: true, body };
}

async function query(sort, { position = 0, limit = 20, get = null, filter = null } = {}) {
  const q = { accountId: s.accountId, sort, position, limit, calculateTotal: true };
  if (filter) q.filter = filter;
  const calls = [['Email/query', q, 'q']];
  if (get) calls.push(['Email/get', { accountId: s.accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: ['id', ...get] }, 'g']);
  const r = await call(calls);
  if (!r.ok) return r;
  const [name, args] = r.body.methodResponses[0];
  if (name === 'error') return { ok: false, type: args.type }; // the TYPE only
  const got = get ? r.body.methodResponses[1] : null;
  let rows = null;
  if (got && got[0] === 'Email/get') {
    const byId = new Map(got[1].list.map(e => [e.id, e]));
    rows = args.ids.map(id => byId.get(id) ?? null);
  }
  return { ok: true, ids: args.ids, total: args.total, position: args.position, rows };
}

const sameIds = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
const firstAddr = v => (Array.isArray(v) && v.length ? v[0] : null);
const searchForm = t => t.normalize('NFD').replace(/\p{M}/gu, '').replace(/\s+/g, ' ').trim().toUpperCase();
// Cyrus compares a missing address as the empty string, so '' is a real key, not a gap.
const displayKey = v => {
  const a = firstAddr(v);
  if (!a) return '';
  return a.name && a.name.trim() ? searchForm(a.name) : (a.email ?? '').toUpperCase();
};
// Cyrus uppercases with C toupper, byte by byte: ASCII letters only.
const asciiUpper = t => t.replace(/[a-z]/g, c => c.toUpperCase());
const baseSubject = v => asciiUpper((v ?? '').replace(/\s+/g, ' ').trim()
  .replace(/(\s*\(fwd\))+$/i, '')
  .replace(/^((re|fwd?|fw)\s*(\[[^\]]*\])?\s*:\s*|\[[^\]]*\]\s*)+/i, '').trim());
const PREFIX = /^((re|fwd?|fw)\s*(\[[^\]]*\])?\s*:|\[)/i;
// For each adjacent pair the subject key rejects, booleans only: a non-ASCII character in
// either subject, a reply/forward/bracket prefix on either, and whether raw bytes of the whole
// subjects are in order. Never a subject or any part of one.
function subjectDiagnosis(label, rows, asc) {
  let n = 0;
  for (let i = 1; i < rows.length; i++) {
    const a = rows[i - 1]?.subject ?? '', b = rows[i]?.subject ?? '';
    const c = byteCmp(baseSubject(a), baseSubject(b));
    if (!(asc ? c > 0 : c < 0)) continue;
    n++;
    const raw = byteCmp(a, b);
    const sf = byteCmp(searchForm(baseSubject(a)), searchForm(baseSubject(b)));
    const lead = t => /^[^\x00-\x7f]/.test(baseSubject(t));
    console.log(`DIAG ${label} pair ${i - 1}-${i}: nonAscii ${/[^\x00-\x7f]/.test(a) || /[^\x00-\x7f]/.test(b)}, prefix ${PREFIX.test(a.trim()) || PREFIX.test(b.trim())}, rawBytesInOrder ${asc ? raw <= 0 : raw >= 0}, searchFormInOrder ${asc ? sf <= 0 : sf >= 0}, leadingNonAscii ${lead(a) || lead(b)}`);
  }
  return n;
}
const dateKey = v => (typeof v === 'string' ? Date.parse(v) : NaN);
const KEYS = {
  receivedAt: dateKey,
  sentAt: dateKey,
  size: v => (typeof v === 'number' ? v : NaN),
  from: displayKey,
  to: displayKey,
  subject: baseSubject,
};
const isGap = k => k === null || k === undefined || (typeof k === 'number' && Number.isNaN(k));

const byteCmp = (a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b));
const cmp = (a, b) => (typeof a === 'number' ? Math.sign(a - b) : byteCmp(a, b));

// Order across the keys that are present; gaps are counted, never compared across.
function ordered(keys, asc) {
  const present = keys.filter(k => !isGap(k));
  for (let i = 1; i < present.length; i++) {
    const c = cmp(present[i - 1], present[i]);
    if (asc ? c > 0 : c < 0) return { ok: false, gaps: keys.length - present.length };
  }
  return { ok: true, gaps: keys.length - present.length };
}
// Collation-independent: once a key has been left, it never comes back.
function contiguous(keys) {
  const seen = new Set();
  let prev;
  for (const k of keys) {
    if (k !== prev && seen.has(k)) return false;
    seen.add(k);
    prev = k;
  }
  return true;
}
// Within each run of equal primary keys, receivedAt non-increasing. Returns the number of
// adjacent tied pairs examined, so a page with none is visible.
function tieBreak(primary, received) {
  let tiedPairs = 0, ok = true;
  for (let i = 1; i < primary.length; i++) {
    if (isGap(primary[i]) || isGap(primary[i - 1]) || cmp(primary[i - 1], primary[i]) !== 0) continue;
    tiedPairs++;
    if (isGap(received[i]) || isGap(received[i - 1]) || received[i - 1] < received[i]) ok = false;
  }
  return { ok, tiedPairs };
}
const keysOf = (rows, prop) => rows.map(r => (r ? KEYS[prop](r[prop]) : null));

async function stablePaging(label, sort, first, filter = null) {
  const again = await query(sort, { filter });
  check(`${label}: same query twice, same ids`, again.ok && sameIds(again.ids, first.ids));
  const page2 = await query(sort, { position: 10, limit: 10, filter });
  check(`${label}: page at position 10 = ids 10-19 of the first query`,
    page2.ok && page2.position === 10 && sameIds(page2.ids, first.ids.slice(10, 20)));
}

// Pages to sample: the start, and two points further into the set, since the start of an
// address or subject sort is often one repeated key.
async function samples(sort, get, first) {
  const out = [{ at: 0, rows: first.rows }];
  for (const at of [Math.floor(first.total / 3), Math.floor((2 * first.total) / 3)]) {
    const more = await query(sort, { position: at, get });
    if (more.ok && more.rows) out.push({ at, rows: more.rows });
  }
  return out;
}

let accepted = 0;
for (const prop of PROPS) {
  for (const asc of [true, false]) {
    const dir = asc ? 'ascending' : 'descending';
    const sort = [{ property: prop, isAscending: asc }];
    const first = await query(sort, { get: [prop] });
    check(`${prop} ${dir}: accepted`, first.ok, first.ok ? `${first.ids.length} ids, total ${first.total}` : `error type ${first.type}`);
    if (!first.ok) continue;
    accepted++;
    check(`${prop} ${dir}: Email/get returned every id`, !!first.rows && first.rows.every(r => r !== null));
    for (const p of await samples(sort, [prop], first)) {
      const keys = keysOf(p.rows, prop);
      const o = ordered(keys, asc);
      const distinct = new Set(keys.filter(k => !isGap(k))).size;
      check(`${prop} ${dir} at position ${p.at}: ordering verified`, o.ok,
        `${distinct} distinct keys in ${keys.length}${o.gaps ? `, ${o.gaps} with no value` : ''}`);
      if (!o.ok && prop === 'subject') subjectDiagnosis(`${prop} ${dir} at position ${p.at}`, p.rows, asc);
      if (typeof keys.find(k => !isGap(k)) === 'string') check(`${prop} ${dir} at position ${p.at}: equal keys are contiguous`, contiguous(keys));
    }
    await stablePaging(`${prop} ${dir}`, sort, first);
  }
}

// The two-key form the build sends: the property, then receivedAt descending whatever the
// primary direction.
for (const prop of PROPS.filter(p => p !== 'receivedAt')) {
  for (const asc of [true, false]) {
    const label = `two-key [${prop} ${asc ? 'asc' : 'desc'}, receivedAt desc]`;
    const sort = [{ property: prop, isAscending: asc }, { property: 'receivedAt', isAscending: false }];
    const get = [prop, 'receivedAt'];
    const first = await query(sort, { get });
    check(`${label}: accepted`, first.ok, first.ok ? `${first.ids.length} ids` : `error type ${first.type}`);
    if (!first.ok) continue;
    let tied = 0, tieOk = true, primaryOk = true;
    for (const p of await samples(sort, get, first)) {
      const primary = keysOf(p.rows, prop);
      if (!ordered(primary, asc).ok) {
        primaryOk = false;
        if (prop === 'subject') subjectDiagnosis(`${label} at position ${p.at}`, p.rows, asc);
      }
      const t = tieBreak(primary, keysOf(p.rows, 'receivedAt'));
      tied += t.tiedPairs;
      if (!t.ok) tieOk = false;
    }
    check(`${label}: primary ordering verified`, primaryOk);
    if (tied === 0) console.log(`INFO ${label}: no ties on the sampled pages, tie-break not exercised`);
    else check(`${label}: receivedAt non-increasing within equal primary keys`, tieOk, `${tied} tied adjacent pairs`);
    await stablePaging(label, sort, first);
  }
}

// The same two-key form under the filter shapes the tools send. Mailbox ids come from the
// mailboxIds of the newest 20 messages (held in memory, never printed), so no Mailbox/get is
// needed: the exclusion shape uses two of them in place of the real Trash and Spam ids, which
// gives the same `inMailboxOtherThan` structure, not the same membership.
const seed = await query([{ property: 'receivedAt', isAscending: false }], { get: ['mailboxIds'] });
const counts = new Map();
for (const r of seed.rows ?? []) for (const id of Object.keys(r?.mailboxIds ?? {})) counts.set(id, (counts.get(id) ?? 0) + 1);
const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1]).map(([id]) => id);
console.log(`INFO filtered cases: ${ranked.length} distinct mailboxes among the newest ${seed.rows?.length ?? 0} messages`);
if (ranked.length < 2) {
  check('filtered cases: at least two mailboxes to build the filter shapes from', false);
} else {
  const [busiest, other] = ranked;
  const SHAPES = {
    // runFilteredQuery's base when nothing else is set: the default exclusion alone.
    'default exclusion (inMailboxOtherThan, 2 ids)': { inMailboxOtherThan: [other, ranked[2] ?? other] },
    'inMailbox': { inMailbox: busiest },
    // A generic common word, not taken from the mailbox.
    'text "the"': { text: 'the' },
    // searchEmails' AND: the base condition, then one keyword condition.
    'AND(exclusion, hasKeyword $seen)': { operator: 'AND', conditions: [{ inMailboxOtherThan: [other] }, { hasKeyword: '$seen' }] },
  };
  for (const [shape, filter] of Object.entries(SHAPES)) {
    for (const prop of PROPS.filter(p => p !== 'receivedAt')) {
      for (const asc of [true, false]) {
        const label = `filtered [${shape}] [${prop} ${asc ? 'asc' : 'desc'}, receivedAt desc]`;
        const sort = [{ property: prop, isAscending: asc }, { property: 'receivedAt', isAscending: false }];
        const first = await query(sort, { get: [prop, 'receivedAt'], filter });
        check(`${label}: accepted`, first.ok, first.ok ? `${first.ids.length} ids, total ${first.total}` : `error type ${first.type}`);
        if (!first.ok) continue;
        const rows = first.rows ?? [];
        const primary = keysOf(rows, prop);
        const o = ordered(primary, asc);
        check(`${label}: ordering verified`, !!first.rows && o.ok,
          `${new Set(primary.filter(k => !isGap(k))).size} distinct keys in ${primary.length}${o.gaps ? `, ${o.gaps} with no value` : ''}`);
        if (!o.ok && prop === 'subject') subjectDiagnosis(label, rows, asc);
        const t = tieBreak(primary, keysOf(rows, 'receivedAt'));
        if (t.tiedPairs === 0) console.log(`INFO ${label}: no ties on the page, tie-break not exercised`);
        else check(`${label}: receivedAt non-increasing within equal primary keys`, t.ok, `${t.tiedPairs} tied adjacent pairs`);
        if (first.total > 10) await stablePaging(label, sort, first, filter);
        else console.log(`INFO ${label}: ${first.total} matches, too few for a second page`);
      }
    }
  }
}

const bogus = await query([{ property: 'notARealSortProperty', isAscending: true }]);
console.log(`INFO undefined sort property: ${bogus.ok ? 'accepted (unexpected)' : `error type ${bogus.type}`}`);

console.log(`accepted: ${accepted}/${PROPS.length * 2}`);
console.log(failures() ? `SORT-PROPERTIES: ${failures()} FAILED` : 'SORT-PROPERTIES: ALL PASS');
process.exit(failures() ? 1 : 0);
