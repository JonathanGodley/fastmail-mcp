// Edit round-trip of a foreign-shape draft: alternative[text, related[html,
// inline image]] with a real @-bearing Content-ID — the structure Fastmail's
// own web client saves for a draft with an embedded image. Verifies the
// recreate-based edit pipeline carries the part verbatim through a metadata
// edit, a body edit that keeps the reference, and a ref-dropping edit (loud
// degrade, part not lost). All drafts trashed on exit.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf, idOf, rawBodies } from './probelib.mjs';
import { getSession, jmap, upload, makePng } from './jmaplib.mjs';

const CID = `${Date.now()}.${Math.floor(Math.random() * 1e15)}@content.messagingengine.example`;
const c = createClient({ env: { ...process.env } });
await c.init();
const { check, failures } = makeChecker();
const trash = [];

const session = await getSession();
const png = await upload(session, makePng(30, 30, 220), 'image/png');
const mbResp = await jmap(session, [['Mailbox/get', { accountId: session.accountId, ids: null, properties: ['id', 'role'] }, 'mb']]);
const drafts = mbResp.methodResponses.find(r => r[2] === 'mb')[1].list.find(m => m.role === 'drafts');

const html = `<div>Foreign-shape round-trip fixture.<br><img src="cid:${CID}"><br>after image</div>`;
const resp = await jmap(session, [['Email/set', {
  accountId: session.accountId,
  create: { f: {
    mailboxIds: { [drafts.id]: true },
    keywords: { $draft: true, $seen: true },
    subject: 'Foreign-shape roundtrip fixture',
    bodyStructure: {
      type: 'multipart/alternative',
      subParts: [
        { partId: 'p1', type: 'text/plain' },
        { type: 'multipart/related', subParts: [
          { partId: 'p2', type: 'text/html' },
          { blobId: png.blobId, type: 'image/png', name: 'fixture.png', cid: CID, disposition: 'inline' },
        ] },
      ],
    },
    bodyValues: {
      p1: { value: 'Foreign-shape round-trip fixture.\nafter image' },
      p2: { value: html },
    },
  } },
}, 'c']]);
const created = resp.methodResponses.find(r => r[2] === 'c')[1];
if (!created.created?.f) { console.error('CREATE FAILED:', JSON.stringify(created.notCreated ?? created).slice(0, 500)); await c.close(); process.exit(1); }
const D0 = created.created.f.id;
trash.push(D0);

const findImg = em => (em.attachments ?? []).find(a => a.cid === CID);

try {
  // Edit 1: metadata-only — the carry must be byte-invariant on the body
  let r = await c.call('edit_draft', { emailId: D0, subject: 'Foreign-shape roundtrip fixture (edited subject)' });
  check('metadata edit succeeded', !r.isError, text(r).slice(0, 300));
  const d1 = idOf(text(r)); if (d1) trash.push(d1);
  // verbose:true so the read returns BOTH stored parts: this draft carries a text/plain and
  // a text/html part, and get_email issues a bodyHash only for a read that showed the whole
  // stored body. Without it the read withholds and the body edits below cannot be made.
  r = await c.call('get_email', { emailId: d1, verbose: true });
  let em = jsonOf(text(r));
  check('metadata edit: read issues a bodyHash', typeof em.bodyHash === 'string', JSON.stringify(em.bodyHashWithheld ?? em.bodyHash));
  let st = await rawBodies(c, d1);
  check('metadata edit: cid verbatim + isInline', findImg(em)?.isInline === true, JSON.stringify(findImg(em) ?? em.attachments));
  check('metadata edit: html still references cid', st.html.includes(`cid:${CID}`), st.html.slice(0, 150));
  check('metadata edit: subject carried', em.subject === 'Foreign-shape roundtrip fixture (edited subject)', JSON.stringify(em.subject));
  check('metadata edit: text part preserved', st.txt.includes('Foreign-shape round-trip fixture'), JSON.stringify(st.txt.slice(0, 80)));

  // Edit 2: body-touching, KEEPING the reference — full pipeline runs against
  // the real @-bearing foreign cid; no false un-recreatable/broken-draft reject
  const newHtml = `<div><p>Edited body, image kept below.</p><img src="cid:${CID}"><p>After image.</p></div>`;
  r = await c.call('edit_draft', { emailId: d1, htmlBody: newHtml, bodyHash: em.bodyHash });
  check('body-touching edit succeeded', !r.isError, text(r).slice(0, 300));
  const d2 = idOf(text(r)); if (d2) trash.push(d2);
  r = await c.call('get_email', { emailId: d2, verbose: true });
  em = jsonOf(text(r));
  check('body edit: read issues a bodyHash', typeof em.bodyHash === 'string', JSON.stringify(em.bodyHashWithheld ?? em.bodyHash));
  st = await rawBodies(c, d2);
  check('body edit: cid part survives verbatim', !!findImg(em), JSON.stringify((em.attachments ?? []).map(a => a.cid)));
  check('body edit: isInline true', findImg(em)?.isInline === true);
  check('body edit: html references cid', st.html.includes(`cid:${CID}`), st.html.slice(0, 150));
  check('body edit: derived text has no cid leak', !st.txt.includes('cid:'), JSON.stringify(st.txt.slice(0, 120)));

  // Edit 3: body edit that DROPS the reference — loud degrade, part carried
  r = await c.call('edit_draft', { emailId: d2, htmlBody: '<p>Image reference removed entirely.</p>', bodyHash: em.bodyHash });
  check('ref-dropping edit succeeded', !r.isError, text(r).slice(0, 300));
  const d3 = idOf(text(r)); if (d3) trash.push(d3);
  check('ref-dropping edit emitted the degrade note', /became regular attachments/.test(text(r)), text(r).split('\n').filter(l => /image|attachment/i.test(l)).join(' | ').slice(0, 300));
  r = await c.call('get_email', { emailId: d3 });
  em = jsonOf(text(r));
  check('after drop: part carried as regular attachment (not lost)', !!findImg(em) && findImg(em).isInline !== true, JSON.stringify(findImg(em) ?? em.attachments));
} finally {
  // Each edit already trashed its predecessor; trashing the rest is idempotent.
  for (const id of trash) { try { await c.call('delete_email', { emailId: id }); } catch { /* already in Trash */ } }
  console.log(`trashed ${trash.length} draft generations`);
  await c.close();
}

console.log(failures() === 0 ? '\nFOREIGN ROUND-TRIP: ALL PASS' : `\nFOREIGN ROUND-TRIP: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
