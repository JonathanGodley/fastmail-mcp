// Does Fastmail's Email/query accept `hasKeyword`/`notKeyword` conditions on `$answered`
// and `$forwarded` (#222), under `urn:ietf:params:jmap:mail`, the capability this server
// already sends? Each is measured alone, AND-ed with each other, AND-ed with the `$seen` and
// `$flagged` conditions `isUnread`/`isPinned` already send, and inside the operator shape
// `runFilteredQuery` builds (a base FilterCondition plus one keyword condition each).
//
// Also checked: has + not partition the account (their totals add up to the unfiltered
// total), so a filter is honoured rather than ignored, and a narrower AND never counts more
// than its wider half.
//
// Raw JMAP, not the built server. READ-ONLY: Email/query only, limit 1, totals read.
// Output is accepted-or-error-type (a refused request: HTTP status and error type, never
// the body) and counts only - no id, subject, address or mailbox
// name - so a run can be quoted into a public issue or commit.

import { getSession, AUTH } from './jmaplib.mjs';
import { makeChecker } from './probelib.mjs';

const { check, failures } = makeChecker();
const s = await getSession();

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

// A base condition of the kind runFilteredQuery puts first in the AND (it holds the text,
// date and mailbox keys); a date floor matches every message without naming a mailbox.
const BASE = { after: '1990-01-01T00:00:00Z' };

// Returns { ok, total, type }. The error TYPE only: a description could quote account data.
async function total(filter) {
  const r = await call([['Email/query', { accountId: s.accountId, filter, limit: 1, calculateTotal: true }, 'q']]);
  if (!r.ok) return r;
  const [name, args] = r.body.methodResponses[0];
  if (name === 'error') return { ok: false, type: args.type };
  return { ok: true, total: args.total };
}

const and = (...conditions) => ({ operator: 'AND', conditions });
const has = k => ({ hasKeyword: k });
const not = k => ({ notKeyword: k });

const all = await total(null);
check('unfiltered query accepted', all.ok, all.ok ? `total ${all.total}` : all.type);
if (!all.ok) process.exit(1);

const cases = [
  ['hasKeyword $answered', has('$answered')],
  ['notKeyword $answered', not('$answered')],
  ['hasKeyword $forwarded', has('$forwarded')],
  ['notKeyword $forwarded', not('$forwarded')],
  ['AND(hasKeyword $answered, hasKeyword $forwarded)', and(has('$answered'), has('$forwarded'))],
  ['AND(hasKeyword $answered, notKeyword $forwarded)', and(has('$answered'), not('$forwarded'))],
  ['AND(notKeyword $answered, notKeyword $forwarded)', and(not('$answered'), not('$forwarded'))],
  ['AND(notKeyword $answered, notKeyword $seen)', and(not('$answered'), not('$seen'))],
  ['AND(hasKeyword $answered, hasKeyword $seen)', and(has('$answered'), has('$seen'))],
  ['AND(hasKeyword $forwarded, hasKeyword $flagged)', and(has('$forwarded'), has('$flagged'))],
  ['AND(notKeyword $forwarded, notKeyword $flagged)', and(not('$forwarded'), not('$flagged'))],
  ['AND(base, notKeyword $answered)', and(BASE, not('$answered'))],
  ['AND(base, notKeyword $seen, hasKeyword $flagged, notKeyword $answered, hasKeyword $forwarded, notKeyword $draft)',
    and(BASE, not('$seen'), has('$flagged'), not('$answered'), has('$forwarded'), not('$draft'))],
];
const r = {};
for (const [label, filter] of cases) {
  r[label] = await total(filter);
  const x = r[label];
  check(`accepted: ${label}`, x.ok, x.ok ? `total ${x.total}` : `error type ${x.type}`);
}

// A control: two keyword keys in ONE FilterCondition is what runFilteredQuery's comment
// says cannot be sent. Informational: recorded, not gated.
const two = await total({ hasKeyword: '$answered', notKeyword: '$forwarded' });
console.log(`INFO two keyword keys in one FilterCondition: ${two.ok ? `accepted, total ${two.total}` : `error type ${two.type}`}`);

const t = label => r[label]?.ok ? r[label].total : NaN;
check('$answered: has + not = unfiltered total',
  t('hasKeyword $answered') + t('notKeyword $answered') === all.total);
check('$forwarded: has + not = unfiltered total',
  t('hasKeyword $forwarded') + t('notKeyword $forwarded') === all.total);
check('$answered is set on at least one message (the filter is exercised, not vacuous)', t('hasKeyword $answered') > 0);
check('$forwarded is set on at least one message', t('hasKeyword $forwarded') > 0,
  t('hasKeyword $forwarded') > 0 ? '' : 'none set: the partition above is vacuous for $forwarded');
check('AND narrows: answered AND forwarded <= answered',
  t('AND(hasKeyword $answered, hasKeyword $forwarded)') <= t('hasKeyword $answered'));
check('AND partitions: (answered AND forwarded) + (answered AND NOT forwarded) = answered',
  t('AND(hasKeyword $answered, hasKeyword $forwarded)') + t('AND(hasKeyword $answered, notKeyword $forwarded)') === t('hasKeyword $answered'));
check('AND with the base: base AND not-answered = not-answered (base matches everything)',
  t('AND(base, notKeyword $answered)') === t('notKeyword $answered'));

console.log(failures() ? `KEYWORD-FILTER: ${failures()} FAILED` : 'KEYWORD-FILTER: ALL PASS');
process.exit(failures() ? 1 : 0);
