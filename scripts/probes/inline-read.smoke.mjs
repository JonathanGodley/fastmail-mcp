// Read-side inline-image surfacing through the real server: isInline/cid in
// get_email, raw purity, get_email_attachments shape, download by cid: form
// (an @-bearing cid, the real-world foreign shape), compact lists unchanged.
// Self-fixturing: creates one related[html, inline image] message in Inbox and
// trashes it on exit. Token via run-probe.py child env; never printed.
import { unlinkSync } from 'node:fs';
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf } from './probelib.mjs';
import { getSession, jmap, upload, makePng } from './jmaplib.mjs';

const CID = 'read-fix@probe.example';
const session = await getSession();
const mb = await jmap(session, [['Mailbox/get', { accountId: session.accountId, ids: null, properties: ['id', 'role'] }, 'm']]);
const inbox = mb.methodResponses.find(r => r[2] === 'm')[1].list.find(x => x.role === 'inbox');
const png = await upload(session, makePng(40, 40, 200), 'image/png');
const pngSize = makePng(40, 40, 200).length;

const created = await jmap(session, [['Email/set', {
  accountId: session.accountId,
  create: { f: {
    mailboxIds: { [inbox.id]: true }, keywords: { $seen: true },
    from: [{ name: 'Probe Fixture', email: 'probe@invalid.example' }],
    subject: 'inline-read probe fixture',
    bodyStructure: { type: 'multipart/related', subParts: [
      { partId: 'h', type: 'text/html' },
      { blobId: png.blobId, type: 'image/png', name: 'readfix.png', cid: CID, disposition: 'inline' },
    ] },
    bodyValues: { h: { value: `<div>Read fixture.<img src="cid:${CID}" alt="fix logo"> tail</div>` } },
  } },
}, 'c']]);
const FIX = created.methodResponses.find(x => x[2] === 'c')[1].created?.f?.id;
if (!FIX) { console.error('FIXTURE CREATE FAILED'); process.exit(1); }

const c = createClient({ env: { ...process.env } });
await c.init();
const { check, failures } = makeChecker();

try {
  // 1. get_email: union entry with isInline + verbatim cid
  let r = await c.call('get_email', { emailId: FIX });
  const email = jsonOf(text(r));
  const att = email.attachments?.find(a => a.cid === CID);
  check('get_email union entry with cid', !!att, JSON.stringify(att ?? email.attachments ?? null).slice(0, 200));
  check('get_email isInline true', att?.isInline === true);

  // 2. get_email raw: pure JMAP, no derived isInline anywhere
  r = await c.call('get_email', { emailId: FIX, raw: true });
  const raw = jsonOf(text(r));
  const rawAtts = raw.attachments ?? [];
  check('raw attachments carry no isInline key', rawAtts.length > 0 && rawAtts.every(a => !('isInline' in a)));

  // 3. get_email_attachments: cid present, raw-shaped (no isInline)
  r = await c.call('get_email_attachments', { emailId: FIX });
  const listing = jsonOf(text(r));
  const entries = Array.isArray(listing) ? listing : listing.attachments ?? [];
  const entry = entries.find(a => a.cid === CID);
  check('get_email_attachments cid entry', !!entry, JSON.stringify(entry ?? listing).slice(0, 200));
  check('get_email_attachments entry has no isInline', entry ? !('isInline' in entry) : false);

  // 4. get_email_attachments raw:true — first item parses as pure JSON
  r = await c.call('get_email_attachments', { emailId: FIX, raw: true });
  let parsed = null;
  try { parsed = JSON.parse(r.content[0].text); } catch { /* fail below */ }
  check('attachments raw first item is pure JSON', parsed !== null && typeof parsed === 'object', `items=${r.content.length}`);

  // 5. download by prefixed cid (@-bearing, the foreign shape)
  r = await c.call('download_attachment', { emailId: FIX, attachmentId: `cid:${CID}`, path: 'probe-readfix.png' });
  check('download by cid: succeeds', !r.isError, text(r).slice(0, 150));
  const saved = text(r).match(/Saved to: (.+?) \((\d+) bytes\)/);
  check('downloaded byte count matches the fixture PNG', saved?.[2] === String(pngSize), saved?.[1] ?? 'no path in response');
  if (saved?.[1]) { try { unlinkSync(saved[1]); } catch { /* leave it */ } }

  // 6. compact path unchanged: list entries carry no attachments array
  r = await c.call('list_emails', { limit: 3 });
  const list = jsonOf(text(r));
  const msgs = list.emails ?? list.messages ?? list;
  check('list_emails compact: no attachments key', Array.isArray(msgs) && msgs.every(m => !('attachments' in m)), `n=${Array.isArray(msgs) ? msgs.length : typeof msgs}`);
} finally {
  try { await c.call('delete_email', { emailId: FIX }); console.log('fixture trashed'); } catch (e) { console.log('CLEANUP FAILED for', FIX, String(e).slice(0, 100)); }
  await c.close();
}

console.log(failures() === 0 ? '\nINLINE-READ: ALL PASS' : `\nINLINE-READ: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
