// Identity-signature live smoke (#33). Settles what unit tests cannot:
//
//  1. The marker survives the real round trip. The signature block is recognised by its
//     class attribute, and nothing in the unit suite proves Fastmail stores and returns a
//     class on a body part rather than re-serializing it away. If it did, every edit-time
//     preservation would silently stop working against the real server.
//  2. An htmlBody-alone edit — the edit that would otherwise drop the signature, because
//     the merge replaces the whole html body — leaves EXACTLY ONE signature block, still
//     above the quoted original, on the draft as actually stored.
//  3. appendSignature:false really takes it off, on a stored draft rather than in a mock.
//  4. A PLAIN-TEXT draft, re-signed from the body it hands back, still carries exactly one
//     sign-off. That draft has no class marker for preservation to find, so the documented
//     recovery is appendSignature:true on the edit — which makes the recovery the place a
//     signature could stack. The idempotence protecting it compares against the text the
//     server actually stored, so it is worth measuring against the real store.
//  5. The CROSS-FORM case of the same thing: an HTML draft converted to plain text with
//     clearFields:['htmlBody'], handed back the text part the server stored. That text is
//     the form DERIVED from the html signature, while a text-only call writes the CONFIGURED
//     textSignature — two different strings on any identity where the two differ — so a
//     check against only the outgoing form appended a second sign-off here. Measured live
//     because what the loop feeds back is the text Fastmail actually returns, whitespace,
//     line endings and all, and those are exactly what a byte-exact match dies on. Note
//     what the run can prove depends on the account: Fastmail normally keeps the two forms
//     in sync, and where they agree this step measures the round trip but cannot separate
//     the two forms. It prints which case it took, and the unit suite is what pins the
//     separation on an identity whose forms deliberately differ.
//  6. The same loop on a plain-text REPLY draft, which is the shape the recovery is actually
//     recommended for. A reply keeps its quoted original inside the body, so the text handed
//     back carries the sign-off in the MIDDLE and the quote at the end — and a check anchored
//     only at the end of the body read that as unsigned and stacked a second sign-off below
//     the quote. What settles it is the attribution line and quote text the server really
//     stored, so it is measured here rather than only against a hand-written fixture.
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

// Bodies read by MIME TYPE. rawBodies() joins whatever is in the htmlBody list, which for a
// TEXT-ONLY message is the text/plain part itself (RFC 8621 §4.1.4 lets a server list one
// part in both lists), so it reports a plain-text draft as having html and no text. The
// text-only steps below need the two told apart, so they key on the part's type.
const typedBodies = async (client, id) => {
  const raw = jsonOf(text(await client.call('get_email', { emailId: id, raw: true })));
  const join = (list, type) => (list ?? [])
    .filter(p => p?.type === type)
    .map(p => raw.bodyValues?.[p.partId]?.value ?? '')
    .join('');
  return { txt: join(raw.textBody, 'text/plain'), html: join(raw.htmlBody, 'text/html') };
};

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

  // 6. THE PLAIN-TEXT ROUND TRIP. A text-only draft carries no class marker, so preservation
  //    cannot reach it and appendSignature:true on the edit is the documented recovery. That
  //    makes the recovery the duplication risk: the natural loop reads the stored body back,
  //    changes the words, and sends the WHOLE text (sign-off included) with the flag still
  //    set. Measured on the stored draft rather than in a mock, because what the loop feeds
  //    back is the text Fastmail actually returns, whitespace and all.
  r = await c.call('create_draft', { to: SELF, subject: `${MARK} plain text`, textBody: 'Hi there.', appendSignature: true });
  check('text: draft created', !r.isError, text(r).slice(0, 250));
  const d6 = idOf(text(r)); if (d6) trash.push(d6);
  let tb = await typedBodies(c, d6);
  const countTail = (s) => (SIG_TAIL && s ? s.split(SIG_TAIL).length - 1 : -1);
  check('text: exactly one sign-off in the text-only draft', countTail(tb.txt) === 1, JSON.stringify(tb.txt.slice(-300)));
  check('text: no html part was fabricated for it', !tb.html, tb.html.slice(0, 200));
  // The nit that only shows on the wire: an images-only html signature derives to nothing,
  // and joining that used to leave the text part ending in bare whitespace.
  check('text: the body does not end in dangling whitespace', tb.txt === tb.txt.trimEnd(), JSON.stringify(tb.txt.slice(-40)));

  r = await c.call('edit_draft', {
    emailId: d6,
    textBody: tb.txt.replace('Hi there.', 'Hi there, one more thing.'),
    appendSignature: true,
  });
  check('text: re-signing edit accepted', !r.isError, text(r).slice(0, 300));
  const d7 = idOf(text(r)); if (d7) trash.push(d7);
  tb = await typedBodies(c, d7);
  check('text: STILL exactly one sign-off after re-signing the body it read back', countTail(tb.txt) === 1, JSON.stringify(tb.txt.slice(-300)));
  check('text: the edited words survived', tb.txt.includes('one more thing'), JSON.stringify(tb.txt.slice(0, 200)));

  // 7. THE CROSS-FORM ROUND TRIP. Convert the signed HTML draft from step 1 to plain text
  //    and hand back the text part the server stored — which is the form DERIVED from the
  //    html signature, while a text-only call writes the CONFIGURED textSignature. Where an
  //    identity's two forms differ these are different strings, and matching only the
  //    outgoing one stacked a second sign-off on exactly this edit.
  const TEXT_SIG = identity?.textSignature ?? '';
  const derivedTb = await typedBodies(c, d1);
  console.log(`cross-form: the identity's two signature forms ${TEXT_SIG && !derivedTb.txt.includes(TEXT_SIG) ? 'DIFFER' : 'agree (or no textSignature is set)'}`);
  r = await c.call('edit_draft', {
    emailId: d1,
    textBody: derivedTb.txt.replace('Body of the compose test.', 'Body of the compose test, revised.'),
    clearFields: ['htmlBody'],
    appendSignature: true,
  });
  check('cross-form: the html-to-text conversion was accepted', !r.isError, text(r).slice(0, 300));
  const d8 = idOf(text(r)); if (d8) trash.push(d8);
  tb = await typedBodies(c, d8);
  check('cross-form: exactly one sign-off after the conversion', countTail(tb.txt) === 1, JSON.stringify(tb.txt.slice(-300)));
  check('cross-form: the configured text signature was not appended alongside the derived one',
    !TEXT_SIG || tb.txt.split(TEXT_SIG).length - 1 <= 1, JSON.stringify(tb.txt.slice(-300)));
  check('cross-form: the edited words survived', tb.txt.includes('revised'), JSON.stringify(tb.txt.slice(0, 200)));
  check('cross-form: no html part survived the conversion', !tb.html, tb.html.slice(0, 200));

  // 8. THE PLAIN-TEXT REPLY ROUND TRIP — step 6's loop on the draft shape it is actually
  //    recommended for. A plain-text reply keeps the quoted original IN its body, so the text
  //    the store hands back has the sign-off in the MIDDLE and the quote at the end. A check
  //    that only looked at the end of the body read that as unsigned and stacked a second
  //    sign-off below the quote, every time round the loop. Measured live because what
  //    decides it is the attribution line and quote text Fastmail actually stored —
  //    whitespace, line endings and wrapping included — which is the half a mock cannot
  //    reproduce.
  r = await c.call('reply_email', { originalEmailId: FIX, textBody: 'Replying in plain text.', appendSignature: true });
  check('text reply: draft created', !r.isError, text(r).slice(0, 250));
  const d9 = idOf(text(r)); if (d9) trash.push(d9);
  tb = await typedBodies(c, d9);
  check('text reply: exactly one sign-off, above the quote',
    countTail(tb.txt) === 1 && tb.txt.indexOf(SIG_TAIL) < tb.txt.indexOf('wrote:'),
    JSON.stringify(tb.txt.slice(0, 400)));
  check('text reply: no html part was fabricated for it', !tb.html, tb.html.slice(0, 200));

  // noQuote, not originalEmailId, and the difference matters. The loop hands back the WHOLE
  // stored text, quote included, so the quote is preserved by the body itself; originalEmailId
  // asks for it to be REBUILT and appended, which on this body appends a second copy of a
  // quote that is already there. Measured: with originalEmailId the stored text came back
  // carrying the attribution line and the quoted original twice. The sign-off check below
  // passes either way, which is exactly why the quote is counted here too.
  r = await c.call('edit_draft', {
    emailId: d9,
    textBody: tb.txt.replace('Replying in plain text.', 'Replying in plain text, one more thing.'),
    noQuote: true,
    appendSignature: true,
  });
  check('text reply: re-signing edit accepted', !r.isError, text(r).slice(0, 300));
  check('text reply: the edit reports the body already carried a sign-off',
    /already carries a signature/i.test(text(r)), text(r).slice(-400));
  const d10 = idOf(text(r)); if (d10) trash.push(d10);
  tb = await typedBodies(c, d10);
  check('text reply: STILL exactly one sign-off after re-signing a body whose quote is below it',
    countTail(tb.txt) === 1, JSON.stringify(tb.txt.slice(0, 500)));
  check('text reply: it is still above the quote, not stacked underneath',
    tb.txt.indexOf(SIG_TAIL) < tb.txt.indexOf('wrote:'), JSON.stringify(tb.txt.slice(0, 500)));
  check('text reply: the edited words survived', tb.txt.includes('one more thing'), JSON.stringify(tb.txt.slice(0, 200)));
  check('text reply: the quote itself survived, exactly once',
    tb.txt.split('Original body of the probe fixture').length - 1 === 1,
    JSON.stringify(tb.txt.slice(-300)));
} finally {
  // The sweep by run marker catches anything a mid-run failure left behind, so it is what
  // keeps a failed run from leaving drafts in a live mailbox. It runs BEFORE the known-id
  // deletes, because search_emails excludes Trash by default: once an artifact is trashed
  // the sweep can no longer see it, and a sweep that can never find anything is a sweep
  // that cannot be observed working.
  //
  // Parse with jsonOf rather than a regex over the rendered text. The response is compact
  // JSON ({"id":"M1",...}), so the pretty-printed /"id": "…"/ this used to scan for matched
  // ZERO — and because nothing checked the parse, the probe printed a trashed count taken
  // from the known-id list and reported ALL PASS while sweeping nothing at all.
  const sweepIds = async (label) => {
    const r = await c.call('search_emails', { query: MARK, limit: 50 });
    const items = jsonOf(text(r));
    // Asserted, not assumed: this is the check that fails loudly if the response shape
    // moves under the parser again, and it does not depend on search having indexed
    // anything yet (a just-created draft may not be findable for a few seconds).
    check(`cleanup: the ${label} marker search parsed as a list`, Array.isArray(items), typeof items);
    return (Array.isArray(items) ? items : []).map(e => e?.id).filter(id => typeof id === 'string');
  };

  let found = [];
  try { found = await sweepIds('pre-delete'); }
  catch (e) { check('cleanup: the pre-delete marker search parsed as a list', false, String(e).slice(0, 160)); }
  console.log(`cleanup: marker sweep found ${found.length} artifact(s) outside Trash`);

  const targets = [...new Set([...trash, ...found])];
  const failed = [];
  let deleted = 0;
  for (const id of targets) {
    try { await c.call('delete_email', { emailId: id }); deleted++; }
    catch (e) { failed.push(`${id} (${String(e.message ?? e).slice(0, 60)})`); }
  }
  check('cleanup: every artifact was trashed', failed.length === 0, failed.join('; '));

  await new Promise(res => setTimeout(res, 3000));
  try {
    const left = await sweepIds('verification');
    check('cleanup: nothing matching the run marker is left outside Trash', left.length === 0, left.join(', '));
  } catch (e) { check('cleanup: the verification marker search parsed as a list', false, String(e).slice(0, 160)); }

  console.log(`trashed ${deleted} of ${targets.length} artifacts (fixture + drafts)`);
  await c.close();
}

console.log(failures() === 0 ? '\nSIGNATURE: ALL PASS' : `\nSIGNATURE: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
