// Quote/forward-carry live smoke over draft_email's {{quote}} and {{forward}} tokens:
// minted ii-...@inline.invalid cids,
// their durability across an edit that hands the quote back, text-only-reply drop note, the
// unsupported-reference-form drop count, flag-off forwards still carrying body
// images, asAttachment untouched, and the send_draft transmit receipt (one
// send-to-self, swept afterwards). Fixtures are received-like messages created
// in Inbox via raw JMAP; every artifact is trashed on exit, including the copy the
// transmit delivers back, and the cleanup itself is CHECKED - a sweep that quietly
// removes nothing is the one failure this probe cannot afford.
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
// Set once step 8 transmits. The delivered copy arrives with an id nothing here has recorded,
// so once this is true the sweep below has work it MUST find (see the note on it).
let transmitted = false;

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
  let r = await c.call('draft_email', { mode: 'reply', originalEmailId: FIX_A, htmlBody: '<p>Thanks for the picture.</p>{{quote}}' });
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

  // 2. edit of that reply, handing the quote back: the minted cid is DURABLE (unchanged)
  // The draft's html is handed back whole, minted reference and all, which is what an edit
  // that keeps the quote looks like now that nothing is preserved for the caller. The cid
  // still names a part this draft carries, so the reserved-shape guard permits it; what is
  // being measured is that the identifier survives the recreate rather than being re-minted,
  // because a caller holding that html across two edits would otherwise reference nothing.
  r = await c.call('get_email', { emailId: d1, fields: ['bodyText', 'bodyHtml', 'bodyHash'] });
  const hash1 = jsonOf(text(r)).bodyHash;
  check('edit: the draft read issued a bodyHash', typeof hash1 === 'string' && hash1.length > 0, text(r).slice(0, 200));
  r = await c.call('edit_draft', { emailId: d1, bodyHash: hash1, htmlBody: `<p>Edited note, quote kept.</p>${st.html}` });
  check('edit: succeeded', !r.isError, text(r).slice(0, 300));
  const d2 = idOf(text(r)); if (d2) trash.push(d2);
  r = await c.call('get_email', { emailId: d2 });
  em = jsonOf(text(r));
  const minted2 = (em.attachments ?? []).map(a => a.cid).find(x => MINT_RE.test(x ?? ''));
  check('edit: minted cid durable across the edit', !!minted2 && minted2 === minted1, `${minted1} vs ${minted2}`);

  // 3. text-only reply: no mint, drop reported
  r = await c.call('draft_email', { mode: 'reply', originalEmailId: FIX_A, textBody: 'Plain text reply only.\n\n{{quote}}' });
  check('text reply: draft created', !r.isError, text(r).slice(0, 250));
  check('text reply: dropped note', /image\(s\) from the quoted message were dropped/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d3 = idOf(text(r)); if (d3) trash.push(d3);
  r = await c.call('get_email', { emailId: d3 });
  em = jsonOf(text(r));
  check('text reply: no attachments (nothing pooled/minted)', (em.attachments ?? []).length === 0, JSON.stringify(em.attachments ?? []));

  // 4. reply to relative-src original: unsupported-form drops counted + noted
  r = await c.call('draft_email', { mode: 'reply', originalEmailId: FIX_B, htmlBody: '<p>About that page.</p>{{quote}}' });
  check('rel reply: draft created', !r.isError, text(r).slice(0, 250));
  check('rel reply: unsupported-form note, count 2', /2 image\(s\) in the quoted message used a reference form this server cannot carry into a quote and were dropped/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d4 = idOf(text(r)); if (d4) trash.push(d4);
  st = await rawBodies(c, d4);
  check('rel reply: quote text survives the drops', st.html.includes('Rel refs.'), st.html.slice(0, 200));

  // 5. forward: embed + mint, "from the original" wording
  r = await c.call('draft_email', { mode: 'forward', originalEmailId: FIX_A, to: SELF, htmlBody: '{{forward}}' });
  check('forward: draft created', !r.isError, text(r).slice(0, 250));
  check('forward: embed note (original wording)', /This draft embeds 1 image\(s\) from the original \(/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d5 = idOf(text(r)); if (d5) trash.push(d5);
  r = await c.call('get_email', { emailId: d5 });
  em = jsonOf(text(r));
  check('forward: minted cid on draft', (em.attachments ?? []).some(a => MINT_RE.test(a.cid ?? '')), JSON.stringify((em.attachments ?? []).map(a => a.cid)));

  // 6. forward flag-off: zip excluded + counted, body image still carried
  r = await c.call('draft_email', { mode: 'forward', originalEmailId: FIX_C, to: SELF, includeOriginalAttachments: false, htmlBody: '{{forward}}' });
  check('flag-off forward: draft created', !r.isError, text(r).slice(0, 250));
  check('flag-off forward: exclusion note + carried sentence', /not included because includeOriginalAttachments is false/.test(text(r)) && /Body-embedded images were still carried/.test(text(r)), text(r).split('\n').filter(l => /includ|image/i.test(l)).join(' | ').slice(0, 300));
  const d6 = idOf(text(r)); if (d6) trash.push(d6);
  r = await c.call('get_email', { emailId: d6 });
  em = jsonOf(text(r));
  check('flag-off forward: image carried, zip absent', (em.attachments ?? []).some(a => MINT_RE.test(a.cid ?? '')) && !(em.attachments ?? []).some(a => /zip/i.test(a.name ?? '') || a.type === 'application/zip'), JSON.stringify((em.attachments ?? []).map(a => ({ n: a.name, t: a.type }))));

  // 7. asAttachment forward untouched: .eml, no embed note
  r = await c.call('draft_email', { mode: 'forward', originalEmailId: FIX_A, to: SELF, asAttachment: true });
  check('asAttachment forward: draft created', !r.isError, text(r).slice(0, 250));
  check('asAttachment forward: no embed note', !/This draft embeds/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  const d7 = idOf(text(r)); if (d7) trash.push(d7);
  r = await c.call('get_email', { emailId: d7 });
  em = jsonOf(text(r));
  check('asAttachment forward: .eml attached', (em.attachments ?? []).some(a => /\.eml$/i.test(a.name ?? '')), JSON.stringify((em.attachments ?? []).map(a => a.name)));

  // 8. send_draft transmit receipt on the edited reply (send-to-self)
  r = await c.call('send_draft', { emailId: d2 });
  transmitted = !r.isError;
  check('send_draft: sent', !r.isError, text(r).slice(0, 300));
  check('send_draft: transmit receipt', /Sent with 1 embedded image\(s\) \(/.test(text(r)), text(r).split('\n').filter(l => /Sent|image/i.test(l)).join(' | '));
} finally {
  // Trash known artifacts, then sweep the send-to-self pair (and anything a
  // mid-run failure left behind) by the unique subject mark.
  for (const id of trash) { try { await c.call('delete_email', { emailId: id }); } catch { /* swept below */ } }
  await new Promise(res => setTimeout(res, 3000));

  // Ids out of the PARSED payload, never a regex over the rendered text. Result payloads are
  // serialised compact, so the pattern this replaced - written against pretty-printed output,
  // with a space after the colon - matched nothing the moment that landed. It failed silently:
  // the sweep found no ids, deleted nothing, and the run still printed a count taken from
  // `trash` and passed, leaving two real messages live in the mailbox on every run. jsonOf is
  // whitespace-agnostic, so it cannot rot the same way.
  //
  // search_emails' default scope excludes Trash, so this reads as "still live in the mailbox"
  // both before the sweep (what to remove) and after it (what the sweep failed to remove).
  const stillLive = async () => {
    const r = await c.call('search_emails', { query: MARK, limit: 50 });
    if (r.isError) throw new Error(text(r).slice(0, 200));
    return jsonOf(text(r)).map(e => e?.id).filter(Boolean);
  };
  // What the sweep is for, precisely: the ARRIVED copy of the send-to-self. The Sent copy is
  // not a second message - EmailSubmission's onSuccessUpdateEmail patches the draft itself
  // into Sent, so it keeps d2's id and the tracked loop above already trashed it. Delivery is
  // asynchronous, so poll for the arrival instead of assuming one 3s wait covered it.
  const DELIVERY_PASSES = 6;
  let swept = 0;
  try {
    for (let pass = 0; pass < DELIVERY_PASSES; pass++) {
      for (const id of await stillLive()) {
        if (trash.includes(id)) continue;
        try { await c.call('delete_email', { emailId: id }); trash.push(id); swept++; } catch { /* re-reported as still live below */ }
      }
      if (!transmitted || swept > 0) break;
      await new Promise(res => setTimeout(res, 3000));
    }
    await new Promise(res => setTimeout(res, 3000));
    // The whole point of asserting on the cleanup: a sweep nobody checks is a sweep that can
    // stop working without anyone noticing, which is exactly what happened here.
    const remaining = await stillLive();
    check('cleanup: nothing matching the probe mark is still live', remaining.length === 0,
      remaining.length ? `${remaining.length} message(s) left in the mailbox: ${remaining.join(', ')} — search_emails query "${MARK}"` : '');
    if (transmitted) {
      check('cleanup: the delivered copy of the send-to-self was swept', swept > 0,
        `sweep removed ${swept} untracked message(s) over ${DELIVERY_PASSES} passes; the transmit delivers a copy back, so it must find at least one`);
    }
  } catch (e) {
    check('cleanup: the sweep ran', false, `sweep failed, artifacts may still be live under "${MARK}": ${String(e).slice(0, 160)}`);
  }
  console.log(`trashed ${trash.length} artifacts (${trash.length - swept} tracked, ${swept} swept by subject mark)`);
  await c.close();
}

console.log(failures() === 0 ? '\nQUOTE-CARRY: ALL PASS' : `\nQUOTE-CARRY: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
