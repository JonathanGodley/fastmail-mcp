// Authoring live smoke: create_draft with a cid-bearing attachment + html
// referencing it → embed note + isInline read-back; lenient cid spellings;
// text-only degrade; dangling-ref and bad-cid rejects (JSON-RPC InvalidParams);
// [image] text derivation for a cid-image-only body. Writes its fixture PNG to
// the OS temp dir and points FASTMAIL_ATTACH_DIR there. Drafts trashed on exit.
import { writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf, idOf } from './probelib.mjs';
import { makePng } from './jmaplib.mjs';

const SELF = 'probe-author@invalid.example';
const ATTACH_DIR = tmpdir();
writeFileSync(join(ATTACH_DIR, 'probe-author.png'), makePng(40, 200, 40));

const c = createClient({ env: { ...process.env, FASTMAIL_ATTACH_DIR: ATTACH_DIR } });
await c.init();
const { check, failures } = makeChecker();
const trash = [];

try {
  // 1. Embed: html references the cid → inline part, embed note
  let r = await c.call('create_draft', { to: SELF, subject: 'author probe embed', htmlBody: '<p>logo:</p><img src="cid:logo" alt="green square"><p>end</p>', attachments: [{ path: 'probe-author.png', cid: 'logo' }] });
  check('embed create succeeded', !r.isError, text(r).slice(0, 250));
  check('embed note emitted', /This draft embeds 1 image/.test(text(r)), text(r).split('\n').filter(l => /image/i.test(l)).join(' | '));
  let id = idOf(text(r));
  if (id) trash.push(id);
  r = await c.call('get_email', { emailId: id });
  const att = (jsonOf(text(r)).attachments ?? []).find(a => a.cid === 'logo');
  check('read-back: isInline + cid', att?.isInline === true && att?.cid === 'logo', JSON.stringify(att ?? null));

  // 2. Lenient spelling: <cid:x> coerces (one <> strip first, then one cid: strip)
  r = await c.call('create_draft', { to: SELF, subject: 'author probe spelling', htmlBody: '<img src="cid:logo2">', attachments: [{ path: 'probe-author.png', cid: '<cid:logo2>' }] });
  check('lenient <cid:x> spelling accepted + linked', !r.isError && /This draft embeds 1 image/.test(text(r)), text(r).slice(0, 200));
  id = idOf(text(r)); if (id) trash.push(id);

  // 3. Degrade: text-only body + cid spec → regular attachment + truthful note
  r = await c.call('create_draft', { to: SELF, subject: 'author probe degrade', textBody: 'plain only', attachments: [{ path: 'probe-author.png', cid: 'ghost' }] });
  check('text-only degrade succeeded (no server reject)', !r.isError, text(r).slice(0, 250));
  check('degrade note truthful', /became regular attachments \(nothing in the body displays them\)/.test(text(r)), text(r).split('\n').filter(l => /image|attach/i.test(l)).join(' | ').slice(0, 250));
  id = idOf(text(r)); if (id) trash.push(id);

  // 4. Reject: dangling ref, before upload (surfaces as JSON-RPC InvalidParams)
  try {
    r = await c.call('create_draft', { to: SELF, subject: 'author probe dangle', htmlBody: '<img src="cid:missing">', attachments: [{ path: 'probe-author.png' }] });
    check('dangling ref rejected', r.isError === true && /references cid "missing"/.test(text(r)), text(r).slice(0, 200));
  } catch (e) {
    check('dangling ref rejected', /references cid \\?"missing\\?" but no attachment supplies it/.test(String(e)), String(e).slice(0, 200));
  }

  // 5. Reject: unusable cid at coerce, with the real item index
  try {
    r = await c.call('create_draft', { to: SELF, subject: 'author probe badcid', htmlBody: '<p>x</p>', attachments: [{ path: 'probe-author.png', cid: 'has space' }] });
    check('bad cid rejected with index', r.isError === true && /attachments\[0\]\.cid/.test(text(r)), text(r).slice(0, 200));
  } catch (e) {
    check('bad cid rejected with index', /attachments\[0\]\.cid/.test(String(e)), String(e).slice(0, 200));
  }

  // 6. cid-image-only body composes and derives [image] (no html-only reject)
  r = await c.call('create_draft', { to: SELF, subject: 'author probe imageonly', htmlBody: '<img src="cid:solo">', attachments: [{ path: 'probe-author.png', cid: 'solo' }] });
  check('cid-image-only body composable', !r.isError, text(r).slice(0, 250));
  id = idOf(text(r)); if (id) trash.push(id);
  if (id) {
    r = await c.call('get_email', { emailId: id, raw: true });
    const raw = jsonOf(text(r));
    const txt = (raw.textBody ?? []).map(p => raw.bodyValues?.[p.partId]?.value ?? '').join('');
    check('derived text is [image]', txt.trim() === '[image]', JSON.stringify(txt.slice(0, 60)));
  }
} finally {
  for (const t of trash) { try { await c.call('delete_email', { emailId: t }); } catch { /* report below */ } }
  console.log(`trashed ${trash.length} probe drafts`);
  await c.close();
}

console.log(failures() === 0 ? '\nINLINE-AUTHOR: ALL PASS' : `\nINLINE-AUTHOR: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
