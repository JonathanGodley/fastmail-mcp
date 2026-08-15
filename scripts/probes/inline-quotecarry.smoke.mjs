// Reply/forward quote-carry live smoke: minted ii-...@inline.invalid cids,
// keep-rebuild reuse (stable identifiers), text-only-reply drop note, the
// unsupported-reference-form drop count, flag-off forwards still carrying body
// images, asAttachment untouched, and the send_draft transmit receipt (one
// send-to-self, swept afterwards). Fixtures are received-like messages created
// in Inbox via raw JMAP; every artifact is trashed on exit, including the
// sent/arrived copies of the send-to-self.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf, idOf, rawBodies } from './probelib.mjs';
import { getSession, jmap, upload, makePng } from './jmaplib.mjs';

const MARK = `qcprobe-${Math.random().toString(36).slice(2, 8)}`;
const CID_A = 'orig-a@fix.example';
const MINT_RE = /ii-[0-9a-f]{32}@inline\.invalid/i;

// Client first: a tool-surface error must not leave orphaned fixtures behind.
const c = createClient({ env: { ...process.env } });
await c.init();
const { check, failures } = makeChecker();
const trash = [];

// The send-to-self target: the account's own identity.
const idRes = await c.call('list_identities', {});
const SELF = (text(idRes).match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/) ?? [])[0];
if (!SELF) { console.error('Could not resolve an identity address for send-to-self.'); await c.close(); process.exit(1); }

const session = await getSession();
const mb = await jmap(session, [['Mailbox/get', { accountId: session.accountId, ids: null, properties: ['id', 'role'] }, 'm']]);
const inbox = mb.methodResponses.find(r => r[2] === 'm')[1].list.find(x => x.role === 'inbox');
const png = await upload(session, makePng(200, 40, 40), 'image/png');
const zip = await upload(session, Buffer.from('PK\x05\x06' + '\x00'.repeat(18), 'binary'), 'application/zip');

const mkEmail = async (subject, bodyStructure, bodyValues) => {
  const r = await jmap(session, [['Email/set', {
    accountId: session.accountId,
    create: { f: {
      mailboxIds: { [inbox.id]: true }, keywords: { $seen: true },
      from: [{ name: 'Probe Fixture', email: SELF }], to: [{ email: SELF }],
      subject, bodyStructure, bodyValues,
    } },
  }, 'c']]);
  const set = r.methodResponses.find(x => x[2] === 'c')[1];
  if (!set.created?.f) throw new Error('FIXTURE CREATE FAILED: ' + JSON.stringify(set.notCreated ?? set).slice(0, 400));
  trash.push(set.created.f.id);
  return set.created.f.id;
};

try {
  const FIX_A = await mkEmail(`${MARK} image original`,
    { type: 'multipart/related', subParts: [
      { partId: 'h', type: 'text/html' },
      { blobId: png.blobId, type: 'image/png', name: 'fixa.png', cid: CID_A, disposition: 'inline' },
    ] },
    { h: { value: `<div>Original body text here.<br><img src="cid:${CID_A}" alt="fixA logo"><br>tail line</div>` } });

  const FIX_B = await mkEmail(`${MARK} relative srcs`,
    { partId: 'h', type: 'text/html' },
    { h: { value: '<div>Rel refs.<img src="/logo.png"><img src="//cdn.example.com/a.png"> end</div>' } });

  const FIX_C = await mkEmail(`${MARK} image plus zip`,
    { type: 'multipart/mixed', subParts: [
      { type: 'multipart/related', subParts: [
        { partId: 'h', type: 'text/html' },
        { blobId: png.blobId, type: 'image/png', name: 'fixc.png', cid: 'orig-c@fix.example', disposition: 'inline' },
      ] },
      { blobId: zip.blobId, type: 'application/zip', name: 'data.zip', disposition: 'attachment' },
    ] },
    { h: { value: '<div>C body.<img src="cid:orig-c@fix.example" alt="c logo"> tail</div>' } });

  // 1. html reply to the image original: embed + mint
  let r = await c.call('reply_email', { originalEmailId: FIX_A, htmlBody: '<p>Thanks for the picture.</p>' });
  check('reply: draft created', !r.isError, text(r).slice(0, 250));
  check('reply: embed note', /This draft embeds 1 image\(s\) from the quoted message \(/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d1 = idOf(text(r)); if (d1) trash.push(d1);
  r = await c.call('get_email', { emailId: d1 });
  let em = jsonOf(text(r));
  const minted1 = (em.attachments ?? []).map(a => a.cid).find(x => MINT_RE.test(x ?? ''));
  check('reply: minted reserved-shape cid, isInline', !!minted1 && (em.attachments ?? []).find(a => a.cid === minted1)?.isInline === true, JSON.stringify(em.attachments ?? []));
  let st = await rawBodies(c, d1);
  check('reply: quote html references the mint', !!minted1 && st.html.includes(`cid:${minted1}`), st.html.slice(0, 200));
  check('reply: original cid nowhere in draft html', !st.html.includes(CID_A));
  check('reply: derived text has no cid leak', !st.txt.includes('cid:'), JSON.stringify(st.txt.slice(0, 120)));

  // 2. keep-rebuild edit of that reply: the minted cid is REUSED (stable)
  r = await c.call('edit_draft', { emailId: d1, htmlBody: '<p>Edited note, quote kept.</p>', originalEmailId: FIX_A });
  check('edit keep-rebuild: succeeded', !r.isError, text(r).slice(0, 300));
  const d2 = idOf(text(r)); if (d2) trash.push(d2);
  r = await c.call('get_email', { emailId: d2 });
  em = jsonOf(text(r));
  const minted2 = (em.attachments ?? []).map(a => a.cid).find(x => MINT_RE.test(x ?? ''));
  check('edit keep-rebuild: minted cid reused verbatim', !!minted2 && minted2 === minted1, `${minted1} vs ${minted2}`);

  // 3. text-only reply: no mint, drop reported
  r = await c.call('reply_email', { originalEmailId: FIX_A, textBody: 'Plain text reply only.' });
  check('text reply: draft created', !r.isError, text(r).slice(0, 250));
  check('text reply: dropped note', /image\(s\) from the quoted message were dropped/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d3 = idOf(text(r)); if (d3) trash.push(d3);
  r = await c.call('get_email', { emailId: d3 });
  em = jsonOf(text(r));
  check('text reply: no attachments (nothing pooled/minted)', (em.attachments ?? []).length === 0, JSON.stringify(em.attachments ?? []));

  // 4. reply to relative-src original: unsupported-form drops counted + noted
  r = await c.call('reply_email', { originalEmailId: FIX_B, htmlBody: '<p>About that page.</p>' });
  check('rel reply: draft created', !r.isError, text(r).slice(0, 250));
  check('rel reply: unsupported-form note, count 2', /2 image\(s\) in the quoted message used a reference form this server cannot carry into a quote and were dropped/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d4 = idOf(text(r)); if (d4) trash.push(d4);
  st = await rawBodies(c, d4);
  check('rel reply: quote text survives the drops', st.html.includes('Rel refs.'), st.html.slice(0, 200));

  // 5. forward: embed + mint, "from the original" wording
  r = await c.call('forward_email', { originalEmailId: FIX_A, to: SELF });
  check('forward: draft created', !r.isError, text(r).slice(0, 250));
  check('forward: embed note (original wording)', /This draft embeds 1 image\(s\) from the original \(/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d5 = idOf(text(r)); if (d5) trash.push(d5);
  r = await c.call('get_email', { emailId: d5 });
  em = jsonOf(text(r));
  check('forward: minted cid on draft', (em.attachments ?? []).some(a => MINT_RE.test(a.cid ?? '')), JSON.stringify((em.attachments ?? []).map(a => a.cid)));

  // 6. forward flag-off: zip excluded + counted, body image still carried
  r = await c.call('forward_email', { originalEmailId: FIX_C, to: SELF, includeOriginalAttachments: false });
  check('flag-off forward: draft created', !r.isError, text(r).slice(0, 250));
  check('flag-off forward: exclusion note + carried sentence', /not included because includeOriginalAttachments is false/.test(text(r)) && /Body-embedded images were still carried/.test(text(r)), text(r).split('\n').filter(l => /includ|image/i.test(l)).join(' | ').slice(0, 300));
  const d6 = idOf(text(r)); if (d6) trash.push(d6);
  r = await c.call('get_email', { emailId: d6 });
  em = jsonOf(text(r));
  check('flag-off forward: image carried, zip absent', (em.attachments ?? []).some(a => MINT_RE.test(a.cid ?? '')) && !(em.attachments ?? []).some(a => /zip/i.test(a.name ?? '') || a.type === 'application/zip'), JSON.stringify((em.attachments ?? []).map(a => ({ n: a.name, t: a.type }))));

  // 7. asAttachment forward untouched: .eml, no embed note
  r = await c.call('forward_email', { originalEmailId: FIX_A, to: SELF, asAttachment: true });
  check('asAttachment forward: draft created', !r.isError, text(r).slice(0, 250));
  check('asAttachment forward: no embed note', !/This draft embeds/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d7 = idOf(text(r)); if (d7) trash.push(d7);
  r = await c.call('get_email', { emailId: d7 });
  em = jsonOf(text(r));
  check('asAttachment forward: .eml attached', (em.attachments ?? []).some(a => /\.eml$/i.test(a.name ?? '')), JSON.stringify((em.attachments ?? []).map(a => a.name)));

  // 8. send_draft transmit receipt on the edited reply (send-to-self)
  r = await c.call('send_draft', { emailId: d2 });
  check('send_draft: sent', !r.isError, text(r).slice(0, 300));
  check('send_draft: transmit receipt', /Sent with 1 embedded image\(s\) \(/.test(text(r)), text(r).split('\n').filter(l => /Sent|image/i.test(l)).join(' | '));
} finally {
  // Trash known artifacts, then sweep the send-to-self pair (and anything a
  // mid-run failure left behind) by the unique subject mark.
  for (const id of trash) { try { await c.call('delete_email', { emailId: id }); } catch { /* swept below */ } }
  await new Promise(res => setTimeout(res, 3000));
  try {
    const r = await c.call('search_emails', { query: MARK, limit: 50 });
    const ids = [...text(r).matchAll(/"id": "([A-Za-z0-9_-]+)"/g)].map(m => m[1]).filter(id => !trash.includes(id));
    for (const id of ids) { try { await c.call('delete_email', { emailId: id }); trash.push(id); } catch { /* report count below */ } }
  } catch (e) { console.log('cleanup sweep failed: ' + String(e).slice(0, 120)); }
  console.log(`trashed ${trash.length} artifacts (fixtures, drafts, sent/arrived copies)`);
  await c.close();
}

console.log(failures() === 0 ? '\nQUOTE-CARRY: ALL PASS' : `\nQUOTE-CARRY: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
