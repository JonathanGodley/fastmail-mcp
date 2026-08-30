// Live probe: exact-instance thread-state marking (docs/conventions.md "Draft
// provenance"). Verifies against a real account that replying to one stored copy
// of a duplicated message marks exactly that copy $answered, not its twin — the
// case the Message-ID fallback would skip as ambiguous. This is the un-mockable
// external path; run it before a release that touches the provenance/marking code.
//
// Run: FASTMAIL_API_TOKEN and FASTMAIL_PROBE_TEST_ADDR (an address authorized to
// receive test mail) in the environment — see scripts/mcp-harness.mjs header for
// sourcing the token safely — then `node scripts/probes/probe-exact-instance.mjs`.
// Writes it makes: one reply draft (sent to the authorized test address), one
// $answered keyword — both reverted/trashed in cleanup on the way out.
//
// 1. Find a real duplicated pair (two stored copies of one Message-ID, none $answered).
// 2. draft_email mode:'reply' to the FILED (non-Sent) copy, addressed to the authorized
//    test address, subject overridden and no {{quote}} placed, so no real content or
//    subject leaves the account.
// 3. Assert the draft carries X-Fastmail-MCP-Source-Id = the filed copy's id.
// 4. send_draft; assert the result reports the original marked (not the old
//    ambiguous skip), then assert $answered landed on the filed copy ONLY.
// 5. Restore the filed copy's prior keywords and trash the sent probe reply.
import { createClient } from '../mcp-harness.mjs';

const TOKEN = process.env.FASTMAIL_API_TOKEN;
const AUTH = { Authorization: `Bearer ${TOKEN}`, 'Content-Type': 'application/json' };
// The probe transmits one real (content-free) reply, so the recipient must be an
// address you are authorized to send test mail to. Deliberately not defaulted.
const TEST_ADDR = process.env.FASTMAIL_PROBE_TEST_ADDR;
if (!TEST_ADDR) {
  console.log('Set FASTMAIL_PROBE_TEST_ADDR to an address authorized to receive test mail.');
  process.exit(1);
}
const SUBJECT = 'MCP exact-instance probe (ignore)';
const SRC_PROP = 'header:X-Fastmail-MCP-Source-Id:asText';

const session = await (await fetch('https://api.fastmail.com/jmap/session', { headers: AUTH })).json();
const accountId = session.primaryAccounts['urn:ietf:params:jmap:mail'];
async function jmap(methodCalls) {
  const res = await fetch(session.apiUrl, {
    method: 'POST', headers: AUTH,
    body: JSON.stringify({ using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'], methodCalls }),
  });
  return (await res.json()).methodResponses;
}

const c = createClient({ env: process.env });
const text = (r) => r?.content?.[0]?.text ?? JSON.stringify(r);
const parseItems = (t) => JSON.parse(t.slice(t.indexOf('['), t.lastIndexOf(']') + 1));
let fail = 0;
const check = (ok, label) => { console.log(`${ok ? 'PASS' : 'FAIL'} ${label}`); if (!ok) fail++; };

try {
  await c.init();

  // Mailbox roles, to tell the Sent copy from the filed one.
  const [[, mb]] = await jmap([['Mailbox/get', { accountId, properties: ['id', 'role', 'name'] }, 'm']]);
  const roleOf = new Map(mb.list.map((m) => [m.id, m.role ?? m.name]));
  const sentId = mb.list.find((m) => m.role === 'sent')?.id;
  const trashId = mb.list.find((m) => m.role === 'trash')?.id;

  // Find a duplicate pair: two copies of one Message-ID, none $answered,
  // one in Sent, the other filed elsewhere (not Trash).
  const resp = await jmap([
    ['Email/query', { accountId, sort: [{ property: 'receivedAt', isAscending: false }], limit: 500 }, 'q'],
    ['Email/get', { accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: ['id', 'messageId', 'keywords', 'mailboxIds', 'subject'] }, 'g'],
  ]);
  const emails = resp[1][1].list;
  const byMsgId = new Map();
  for (const e of emails) {
    const mid = e.messageId?.[0];
    if (!mid) continue;
    if (!byMsgId.has(mid)) byMsgId.set(mid, []);
    byMsgId.get(mid).push(e);
  }
  let filed = null, twin = null;
  for (const copies of byMsgId.values()) {
    if (copies.length !== 2) continue;
    if (copies.some((e) => e.keywords?.$answered || e.keywords?.$draft)) continue;
    const inSent = copies.find((e) => e.mailboxIds?.[sentId]);
    const other = copies.find((e) => !e.mailboxIds?.[sentId] && !e.mailboxIds?.[trashId]);
    if (inSent && other && inSent !== other) { filed = other; twin = inSent; break; }
  }
  if (!filed) throw new Error('no suitable duplicate pair found in the newest 500 messages');
  const filedLoc = Object.keys(filed.mailboxIds).map((id) => roleOf.get(id)).join('+');
  console.log(`pair: filed=${filed.id} (${filedLoc}) twin=${twin.id} (sent); msgId=${filed.messageId[0]}`);
  const priorKeywords = { ...(filed.keywords ?? {}) };

  // Reply to the FILED copy — the two-step Message-ID lookup would see BOTH copies and skip.
  const replyRes = text(await c.call('draft_email', {
    mode: 'reply', originalEmailId: filed.id, to: TEST_ADDR, subject: SUBJECT,
    htmlBody: '<p>Automated probe of exact-instance thread-state marking; please ignore.</p>',
  }));
  const m = replyRes.match(/Email ID: ([^\s).,]+)/);
  if (!m) throw new Error(`no draft id in: ${replyRes.slice(0, 200)}`);
  const draftId = m[1];

  // The draft must record the exact instance.
  const dgResp = await jmap([['Email/get', { accountId, ids: [draftId], properties: ['id', SRC_PROP] }, 'd']]);
  const dg = dgResp[0][1];
  if (!dg?.list) throw new Error(`draft Email/get failed: ${JSON.stringify(dgResp).slice(0, 400)}`);
  const stamped = (dg.list[0]?.[SRC_PROP] ?? '').trim();
  check(stamped === filed.id, `draft records the filed copy (${SRC_PROP}=${stamped || '(absent)'})`);

  const sendRes = text(await c.call('send_draft', { emailId: draftId }));
  console.log(`send_draft => ${sendRes}`);
  check(/answered/i.test(sendRes) && !/not marked|skip/i.test(sendRes), 'result reports the original marked (no ambiguous skip)');

  const [[, after]] = await jmap([['Email/get', { accountId, ids: [filed.id, twin.id], properties: ['id', 'keywords'] }, 'a']]);
  const kw = new Map(after.list.map((e) => [e.id, e.keywords ?? {}]));
  check(kw.get(filed.id)?.$answered === true, `filed copy ${filed.id} has $answered`);
  check(!kw.get(twin.id)?.$answered, `sent twin ${twin.id} untouched`);

  // ---- cleanup ----
  await jmap([['Email/set', { accountId, update: { [filed.id]: { keywords: priorKeywords } } }, 's']]);
  const [[, restored]] = await jmap([['Email/get', { accountId, ids: [filed.id], properties: ['keywords'] }, 'r']]);
  console.log(`restored filed-copy keywords: ${JSON.stringify(restored.list[0]?.keywords)}`);
  const artifacts = [];
  for (const mailbox of ['sent', 'drafts']) {
    const hits = parseItems(text(await c.call('search_emails', {
      subject: SUBJECT, mailbox, limit: 10, fields: ['id', 'subject'],
    })));
    artifacts.push(...hits.map((e) => e.id));
  }
  for (const id of artifacts) {
    try { await c.call('delete_email', { emailId: id }); } catch (err) { console.log(`cleanup skip ${id}: ${err.message}`); }
  }
  console.log(`cleanup: trashed ${artifacts.length} probe artifact(s)`);

  console.log(fail ? `\n${fail} FAILURE(S)` : '\nALL PASS');
  process.exitCode = fail ? 1 : 0;
} catch (err) {
  console.log(`PROBE ERROR: ${err?.message ?? err}`);
  process.exitCode = 1;
} finally {
  c.close();
}
