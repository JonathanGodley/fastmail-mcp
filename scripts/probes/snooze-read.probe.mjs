// Does Fastmail serve a snoozed message's `snoozed` object (wake time, destination,
// keyword changes) over JMAP, and under which capability? The mechanism, the IETF
// draft and the 2026-10-28 result (refused at every name to an API token) are on #35
// and in the README row.
//
// Raw JMAP, not the built server. READ-ONLY: Mailbox/get, Email/query, Email/get.
// Output: capability names, counts, wake times and the object's shape; subjects are
// truncated to 24 characters.

// FASTMAIL_WEB_TOKEN, when set, is the bearer token Fastmail's own web client sends,
// so the run measures a first-party session instead. jmaplib reads FASTMAIL_API_TOKEN
// at import, so the override must land before the dynamic import.
const WEB = process.env.FASTMAIL_WEB_TOKEN;
if (WEB) process.env.FASTMAIL_API_TOKEN = WEB;
const { getSession, jmap } = await import('./jmaplib.mjs');
import { makeChecker } from './probelib.mjs';

const CYRUS = 'https://cyrusimap.org/ns/jmap/mail';
const IETF = 'urn:ietf:params:jmap:mail:snooze';
const FM_DEV_MAIL = 'https://www.fastmail.com/dev/mail'; // exists at the API layer (403, not 400)
const STANDARD = ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission', 'urn:ietf:params:jmap:contacts'];
const { check, failures } = makeChecker();

console.log('credential:', WEB ? 'FASTMAIL_WEB_TOKEN (web client session)' : 'FASTMAIL_API_TOKEN');
const TOKEN = process.env.FASTMAIL_API_TOKEN;
const sessionRes = await fetch('https://api.fastmail.com/jmap/session', { headers: { Authorization: `Bearer ${TOKEN}` } });
if (sessionRes.status === 401) { console.error('AUTH FAILED: 401 from /jmap/session (a web session token may have expired). Stopping.'); process.exit(3); }
const session = await sessionRes.json();
const s = await getSession();
const sessionCaps = Object.keys(session.capabilities ?? {});
const accountCaps = Object.keys(session.accounts?.[s.accountId]?.accountCapabilities ?? {});
console.log('session capabilities:', sessionCaps.join(', '));
console.log('account capabilities:', accountCaps.join(', '));
// Every non-standard capability the session advertises is tried too, so a first-party
// name not guessed here is still measured.
const advertised = [...new Set([...sessionCaps, ...accountCaps])].filter(c => !STANDARD.includes(c));
const hasCyrus = sessionCaps.includes(CYRUS) || accountCaps.includes(CYRUS);
const hasIetf = sessionCaps.includes(IETF) || accountCaps.includes(IETF);
check('advertises the Cyrus vendor capability', hasCyrus);
check('advertises the IETF snooze capability', hasIetf, hasIetf ? '' : 'absent (informational)');

function res(body, i) {
  const entry = body.methodResponses[i];
  return entry;
}

const mb = await jmap(s, [['Mailbox/get', { accountId: s.accountId, properties: ['id', 'name', 'role'] }, 'm']]);
const mailboxes = res(mb, 0)[1].list;
const byId = Object.fromEntries(mailboxes.map(m => [m.id, m]));
const snoozedMb = mailboxes.find(m => m.role === 'snoozed');
check('account has a role=snoozed mailbox', !!snoozedMb, snoozedMb ? `name "${snoozedMb.name}"` : '');
if (!snoozedMb) process.exit(1);
const label = id => { const m = byId[id]; return m ? (m.role ? `role:${m.role}` : `label:"${m.name}"`) : `unknown:${id}`; };

const PROPS = ['id', 'subject', 'mailboxIds', 'keywords', 'snoozed', 'receivedAt'];

// 1. Read `snoozed` with no extension in `using`: does the standard capability alone serve it?
const plain = await jmap(s, [
  ['Email/query', { accountId: s.accountId, filter: { inMailbox: snoozedMb.id }, limit: 50 }, 'q'],
  ['Email/get', { accountId: s.accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: PROPS }, 'g'],
]);
const qPlain = res(plain, 0);
const gPlain = res(plain, 1);
console.log('Email/query (no extension):', qPlain[0], qPlain[0] === 'error' ? JSON.stringify(qPlain[1]) : `${qPlain[1].ids.length} ids, total ${qPlain[1].total}`);
console.log('Email/get snoozed (no extension):', gPlain[0], gPlain[0] === 'error' ? JSON.stringify(gPlain[1]) : `${gPlain[1].list.length} emails`);
check('Email/get with `snoozed` under the standard mail capability alone', gPlain[0] === 'Email/get', gPlain[0] === 'error' ? `error type ${gPlain[1].type}` : 'served (informational: tells us whether `using` must change)');

// 2. The same read under each candidate capability, sorted by snoozedUntil. jmaplib
// throws on a non-200, the expected answer for a refused name, so each attempt is caught.
let gExt = null, servedCap = null;
for (const cap of [...new Set([CYRUS, IETF, FM_DEV_MAIL, ...advertised])]) {
  let body;
  try {
    body = await jmap(s, [
      ['Email/query', { accountId: s.accountId, filter: { inMailbox: snoozedMb.id }, sort: [{ property: 'snoozedUntil', mailboxId: snoozedMb.id, isAscending: true }], limit: 50 }, 'q'],
      ['Email/get', { accountId: s.accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: PROPS }, 'g'],
    ], [cap]);
  } catch (e) {
    check(`\`using\` ${cap} is accepted`, false, e.message.replace(/"trace_id":"[^"]*",?/, ''));
    continue;
  }
  const q = res(body, 0), g = res(body, 1);
  check(`Email/query sorts by snoozedUntil under ${cap}`, q[0] === 'Email/query', q[0] === 'error' ? JSON.stringify(q[1]) : `${q[1].ids.length} ids`);
  check(`Email/get serves \`snoozed\` under ${cap}`, g[0] === 'Email/get', g[0] === 'error' ? JSON.stringify(g[1]) : '');
  if (g[0] === 'Email/get') { gExt = g; servedCap = cap; break; }
}
if (!gExt) {
  console.log('\nNo capability served `snoozed`; the message dump and the null-control read did not run.');
  process.exit(1);
}

const list = gExt[1].list;
console.log(`\n${list.length} message(s) in the snoozed mailbox:`);
let withObj = 0, sortedOk = true, prev = '';
for (const e of list) {
  const sn = e.snoozed;
  if (sn && typeof sn === 'object') withObj++;
  const until = sn?.until ?? '(none)';
  if (sn?.until) { if (prev && sn.until < prev) sortedOk = false; prev = sn.until; }
  const dest = sn?.moveToMailboxId ? label(sn.moveToMailboxId) : '(absent -> Inbox)';
  const kw = sn?.setKeywords ? JSON.stringify(sn.setKeywords) : '(absent)';
  const extra = sn ? Object.keys(sn).filter(k => !['until', 'moveToMailboxId', 'setKeywords'].includes(k)) : [];
  console.log(`  subject "${(e.subject ?? '').slice(0, 24)}" | in ${Object.keys(e.mailboxIds).map(label).join(' + ')} | keywords ${JSON.stringify(e.keywords)} | received ${e.receivedAt}`);
  console.log(`    snoozed.until ${until} | moveToMailboxId ${dest} | setKeywords ${kw}${extra.length ? ' | OTHER FIELDS ' + extra.join(',') : ''}`);
}
check('every message in the snoozed mailbox carries a snoozed object', withObj === list.length, `${withObj}/${list.length}`);
check('snoozedUntil sort order is ascending by until', sortedOk);

// 3. A message NOT in the snoozed mailbox reads snoozed: null (the property is served, not omitted).
const inbox = mailboxes.find(m => m.role === 'inbox');
const ctl = await jmap(s, [
  ['Email/query', { accountId: s.accountId, filter: { inMailbox: inbox.id }, limit: 1 }, 'q'],
  ['Email/get', { accountId: s.accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: ['id', 'snoozed'] }, 'g'],
], [servedCap]).catch(() => ({ methodResponses: [[], ['error', {}]] }));
const ctlList = res(ctl, 1)[1]?.list ?? [];
check('an un-snoozed message reads snoozed: null', ctlList.length === 1 && 'snoozed' in ctlList[0] && ctlList[0].snoozed === null, JSON.stringify(ctlList[0]?.snoozed));

process.exit(failures() ? 1 : 0);
