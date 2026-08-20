// Identity-signature live smoke (#33). Settles the three things unit tests cannot:
//
//  1. The marker survives the real round trip. The signature block is recognised by its
//     class attribute, and nothing in the unit suite proves Fastmail stores and returns a
//     class on a body part rather than re-serializing it away. If it did, every edit-time
//     preservation would silently stop working against the real server.
//  2. An htmlBody-alone edit — the edit that would otherwise drop the signature, because
//     the merge replaces the whole html body — leaves EXACTLY ONE signature block, still
//     above the quoted original, on the draft as actually stored.
//  3. appendSignature:false really takes it off, on a stored draft rather than in a mock.
//
// Also checks that the plain-text part of a signed html draft is DERIVED from the html
// signature (not the configured textSignature), which is only observable on the assembled
// message the server hands back.
//
// Drafts only: this probe never calls send_draft and creates no calendar events, so it
// transmits nothing. Every artifact is trashed on exit, including on failure.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf, idOf, rawBodies } from './probelib.mjs';
import { getSession, jmap } from './jmaplib.mjs';

const MARK = `sigprobe-${Math.random().toString(36).slice(2, 8)}`;
const SIG_CLASS = 'fm-mcp-signature';
const countSig = (html) => (html.match(new RegExp(SIG_CLASS, 'g')) ?? []).length;
// A distinctive line of the configured html signature, with tags removed, so the derived
// text part can be located without assuming how the html was converted.
const lastTextLine = (html) => html
  .replace(/<[^>]+>/g, '\n').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
  .split('\n').map(s => s.trim()).filter(Boolean).pop() ?? '';

const c = createClient({ env: { ...process.env } });
await c.init();
const { check, failures } = makeChecker();
const trash = [];

// The account's own identity, and the signature it is configured with. Without one there
// is nothing to prove, so say that plainly rather than passing vacuously.
const idRes = await c.call('list_identities', {});
const identities = jsonOf(text(idRes));
const identity = identities.find(i => i.mayDelete === false) ?? identities[0];
const SELF = identity?.email;
const HTML_SIG = identity?.htmlSignature;
if (!SELF || !HTML_SIG) {
  console.error('The default identity has no htmlSignature configured; this probe has nothing to measure.');
  await c.close();
  process.exit(1);
}

const session = await getSession();
const mb = await jmap(session, [['Mailbox/get', { accountId: session.accountId, ids: null, properties: ['id', 'role'] }, 'm']]);
const inbox = mb.methodResponses.find(r => r[2] === 'm')[1].list.find(x => x.role === 'inbox');

try {
  // A received-like fixture to reply to, so the reply carries a real quote.
  const created = await jmap(session, [['Email/set', {
    accountId: session.accountId,
    create: { f: {
      mailboxIds: { [inbox.id]: true }, keywords: { $seen: true },
      from: [{ name: 'Probe Fixture', email: SELF }], to: [{ email: SELF }],
      subject: `${MARK} original`,
      bodyStructure: { partId: 'h', type: 'text/html' },
      bodyValues: { h: { value: '<div>Original body of the probe fixture.</div>' } },
    } },
  }, 'c']]);
  const setRes = created.methodResponses.find(x => x[2] === 'c')[1];
  if (!setRes.created?.f) throw new Error('FIXTURE CREATE FAILED: ' + JSON.stringify(setRes.notCreated ?? setRes).slice(0, 400));
  const FIX = setRes.created.f.id;
  trash.push(FIX);

  // 1. create_draft: the marker survives create -> store -> fetch
  let r = await c.call('create_draft', { to: SELF, subject: `${MARK} plain compose`, htmlBody: '<p>Body of the compose test.</p>', appendSignature: true });
  check('compose: draft created', !r.isError, text(r).slice(0, 250));
  const d1 = idOf(text(r)); if (d1) trash.push(d1);
  let st = await rawBodies(c, d1);
  check('compose: exactly one signature block survives the round trip', countSig(st.html) === 1, st.html.slice(0, 400));
  check('compose: the block carries the configured html signature', st.html.includes(HTML_SIG), st.html.slice(0, 400));
  // The text part is derived from the html signature, not copied from textSignature. On a
  // Fastmail account the two are normally kept in sync, so this check can only prove the
  // derived text is PRESENT; the unit suite is what separates derived from verbatim.
  const SIG_TAIL = lastTextLine(HTML_SIG);
  check('compose: text part carries the derived signature', SIG_TAIL.length > 0 && st.txt.includes(SIG_TAIL), `tail=${JSON.stringify(SIG_TAIL)} txt=${st.txt.slice(-200)}`);

  // 2. an unsigned compose stays unsigned (the default is off)
  r = await c.call('create_draft', { to: SELF, subject: `${MARK} unsigned`, htmlBody: '<p>No signature here.</p>' });
  const d2 = idOf(text(r)); if (d2) trash.push(d2);
  st = await rawBodies(c, d2);
  check('compose: nothing is appended when the flag is absent', countSig(st.html) === 0, st.html.slice(0, 300));

  // 3. reply: the signature sits above the quote on the stored draft
  r = await c.call('reply_email', { originalEmailId: FIX, htmlBody: '<p>Replying with a sign-off.</p>', appendSignature: true });
  check('reply: draft created', !r.isError, text(r).slice(0, 250));
  const d3 = idOf(text(r)); if (d3) trash.push(d3);
  st = await rawBodies(c, d3);
  check('reply: exactly one signature block', countSig(st.html) === 1, st.html.slice(0, 500));
  check('reply: signature sits ABOVE the quote', st.html.indexOf(SIG_CLASS) > 0 && st.html.indexOf(SIG_CLASS) < st.html.indexOf('<blockquote'), st.html.slice(0, 500));
  check('reply: text part signs above the attribution', st.txt.includes(SIG_TAIL) && st.txt.indexOf('wrote:') > 0 && st.txt.indexOf(SIG_TAIL) < st.txt.indexOf('wrote:'), st.txt.slice(0, 400));

  // 4. THE EDIT THAT WOULD LOSE IT: htmlBody alone, no appendSignature at all
  r = await c.call('edit_draft', { emailId: d3, htmlBody: '<p>Rewritten body, no signature written by me.</p>', originalEmailId: FIX });
  check('edit: accepted', !r.isError, text(r).slice(0, 300));
  check('edit: announces the re-append', /signature was re-appended/i.test(text(r)), text(r).slice(0, 400));
  const d4 = idOf(text(r)); if (d4) trash.push(d4);
  st = await rawBodies(c, d4);
  check('edit: EXACTLY ONE signature block survives', countSig(st.html) === 1, st.html.slice(0, 500));
  check('edit: still above the quote', st.html.indexOf(SIG_CLASS) < st.html.indexOf('<blockquote'), st.html.slice(0, 500));
  check('edit: the quote itself survived', st.html.includes('Original body of the probe fixture'), st.html.slice(0, 500));

  // 5. appendSignature:false takes it off a stored draft
  r = await c.call('edit_draft', { emailId: d4, htmlBody: '<p>Now with no sign-off at all.</p>', originalEmailId: FIX, appendSignature: false });
  check('drop: accepted', !r.isError, text(r).slice(0, 300));
  check('drop: says nothing about a signature', !/signature/i.test(text(r)), text(r).slice(0, 300));
  const d5 = idOf(text(r)); if (d5) trash.push(d5);
  st = await rawBodies(c, d5);
  check('drop: no signature block remains', countSig(st.html) === 0, st.html.slice(0, 400));
  check('drop: the quote still survived', st.html.includes('Original body of the probe fixture'), st.html.slice(0, 400));
} finally {
  for (const id of trash) { try { await c.call('delete_email', { emailId: id }); } catch { /* swept below */ } }
  await new Promise(res => setTimeout(res, 3000));
  try {
    const r = await c.call('search_emails', { query: MARK, limit: 50 });
    const ids = [...text(r).matchAll(/"id": "([A-Za-z0-9_-]+)"/g)].map(m => m[1]).filter(id => !trash.includes(id));
    for (const id of ids) { try { await c.call('delete_email', { emailId: id }); trash.push(id); } catch { /* report count below */ } }
  } catch (e) { console.log('cleanup sweep failed: ' + String(e).slice(0, 120)); }
  console.log(`trashed ${trash.length} artifacts (fixture + drafts)`);
  await c.close();
}

console.log(failures() === 0 ? '\nSIGNATURE: ALL PASS' : `\nSIGNATURE: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
