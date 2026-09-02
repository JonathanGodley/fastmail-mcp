// Body-token round trip against a live account, DRAFTS ONLY: this probe never sends, and
// every artifact it makes is a draft it trashes on exit. It settles what only a real
// account can answer about {{signature}} - that the identity's CONFIGURED sign-off is what
// lands, and that it lands ONCE, at write, so a body read back carries no token to expand
// again. Everything else here is the surface around that: the escape ships bare braces,
// the four compose refusals fire before anything is stored, and edit_draft's bodyHash gate
// refuses a body edit with no hash and with a stale one, then hands back the hash the NEXT
// edit needs. The unit suite covers all of that against a mock; what it cannot cover is a
// signature this server did not invent and a hash taken over bytes a real server stored.
//
// Recipients are addressed into invalid.example (RFC 2606 reserves it and it publishes no
// MX), so nothing here is deliverable even by accident. No send_draft call, no identity
// address in a To line, no calendar participant.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf, idOf } from './probelib.mjs';

const TO = 'probe-tokens@invalid.example';
const SUBJ = `token probe ${Math.random().toString(36).slice(2, 8)}`;
const TOKEN_RE = /\{\{signature\}\}/;

// Tag-stripped, whitespace-collapsed form of some html, so "the sign-off landed" is asked
// of the words rather than of markup this server may wrap or reflow.
const words = s => (s ?? '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();

// The receipt draft_email prints after its summary. Sliced at the marker rather than handed
// to jsonOf whole: the notes above it are prose, and several of them contain braces.
const tokensOf = t => { const i = t.indexOf('Tokens: '); return i < 0 ? undefined : jsonOf(t.slice(i + 8)); };
const partOf = (receipt, part) => (receipt?.parts ?? []).find(p => p.part === part);

const c = createClient({ env: { ...process.env } });
await c.init();
const { check, failures } = makeChecker();
const trash = [];
const keep = r => { const id = idOf(text(r)); if (id) trash.push(id); return id; };

// BOTH bodies and the hash in one read. The default get_email omits bodyHtml whenever a
// text part exists (it reports bodyHtmlSize instead), and a hash is issued only to a read
// that showed the caller every stored byte it covers - so this projection is what makes
// the html assertions and the edits below possible at all.
const readDraft = async id => {
  const r = await c.call('get_email', { emailId: id, fields: ['bodyText', 'bodyHtml', 'bodyHash'] });
  if (r.isError) throw new Error(`read of ${id} failed: ${text(r).slice(0, 200)}`);
  const em = jsonOf(text(r));
  return { html: em.bodyHtml ?? '', txt: em.bodyText ?? '', hash: em.bodyHash, withheld: em.bodyHashWithheld };
};

// A call that must be refused. Both shapes are handled because these refusals are
// McpError(InvalidParams), which the harness may surface as a thrown error or as isError.
// A refusal that unexpectedly SUCCEEDS has stored a draft, so track it before failing.
const mustRefuse = async (label, tool, args, re) => {
  let body;
  try {
    const r = await c.call(tool, args);
    if (!r.isError) { keep(r); check(label, false, `call succeeded, expected a refusal: ${text(r).slice(0, 180)}`); return; }
    body = text(r);
  } catch (e) { body = String(e); }
  check(label, re.test(body), body.slice(0, 240));
};

try {
  // ---- 0. The identity this probe composes as, and the sign-off it carries -------------
  // `from` is passed explicitly on every compose below so the identity the server signs
  // with is the one read here. Without it the server picks the account default and every
  // expectation would be about an identity the probe never looked at.
  let r = await c.call('list_identities', {});
  check('identities: readable', !r.isError, text(r).slice(0, 200));
  const identities = jsonOf(text(r));
  const signed = identities.find(i => words(i.htmlSignature) || (i.textSignature ?? '').trim());
  const identity = signed ?? identities[0];
  check('identities: one to compose as', !!identity?.email, JSON.stringify(identities).slice(0, 200));
  const FROM = identity.email;
  const SIG_WORDS = words(identity?.htmlSignature);
  console.log(`composing as ${FROM}: ${signed ? 'signature configured' : 'NO signature configured, so the no-signature branch is what runs'}`);

  // ---- 1. {{signature}} expands once, at write ----------------------------------------
  r = await c.call('draft_email', {
    mode: 'new', from: FROM, to: TO, subject: SUBJ,
    htmlBody: '<p>Body above.</p><div>{{signature}}</div>',
    textBody: 'Body above.\n\n{{signature}}',
  });
  check('signature: draft created', !r.isError, text(r).slice(0, 250));
  const created = text(r);
  const d1 = keep(r);
  const receipt = tokensOf(created);
  check('signature: the receipt names both parts', !!partOf(receipt, 'htmlBody') && !!partOf(receipt, 'textBody'), JSON.stringify(receipt));

  const stored = await readDraft(d1);
  // The property the design rests on, and it holds whether or not a signature exists: the
  // token is consumed at write, so nothing read back can expand a second time.
  check('signature: no {{signature}} survives in the stored html', !TOKEN_RE.test(stored.html), stored.html.slice(0, 200));
  check('signature: no {{signature}} survives in the stored text', !TOKEN_RE.test(stored.txt), stored.txt.slice(0, 200));

  // The receipt and the stored body have to tell the same story, so a receipt claiming an
  // expansion that did not happen fails here instead of reading as a pass.
  const expandedIn = part => (partOf(receipt, part)?.expanded ?? []).some(e => e.token === 'signature' && e.count === 1);
  const removedIn = part => (partOf(receipt, part)?.removed ?? []).some(x => x.token === 'signature' && x.cause === 'no-signature');
  if (signed) {
    check('signature: the receipt reports one expansion per part', expandedIn('htmlBody') && expandedIn('textBody'), JSON.stringify(receipt));
    if (SIG_WORDS) {
      check('signature: the identity\'s configured sign-off is in the stored html',
        words(stored.html).includes(SIG_WORDS), `looking for "${SIG_WORDS.slice(0, 60)}" in ${words(stored.html).slice(0, 200)}`);
    }
  } else {
    check('no signature configured: the receipt records the removal and its cause', removedIn('htmlBody'), JSON.stringify(receipt));
    check('no signature configured: the removal is stated in a note',
      /\{\{signature\}\} in htmlBody was removed: the sending identity has no signature configured/.test(created),
      created.slice(0, 400));
  }

  // ---- 2. The escape ships the braces as text -----------------------------------------
  // Html alone, so the text alternative is derived from it: the bare token has to survive
  // that derivation too, or the text part says something the html did not.
  r = await c.call('draft_email', {
    mode: 'new', from: FROM, to: TO, subject: `${SUBJ} escaped`,
    htmlBody: '<p>Ships as text: \\{{signature}}</p>',
  });
  check('escape: draft created', !r.isError, text(r).slice(0, 250));
  const d2 = keep(r);
  check('escape: no token receipt, because an escape is not a token', tokensOf(text(r)) === undefined, text(r).slice(-200));
  const escaped = await readDraft(d2);
  check('escape: the bare token is stored in the html', TOKEN_RE.test(escaped.html), escaped.html.slice(0, 200));
  check('escape: the backslash was consumed', !/\\\{\{signature\}\}/.test(escaped.html), escaped.html.slice(0, 200));
  check('escape: the bare token reaches the derived text part', TOKEN_RE.test(escaped.txt), escaped.txt.slice(0, 200));

  // ---- 3. The compose refusals, each of which fires before anything is stored ----------
  await mustRefuse('refusal: a near-miss spelling', 'draft_email',
    { mode: 'new', from: FROM, to: TO, subject: SUBJ, htmlBody: '<p>{{Signature}}</p>' },
    /is not a token[\s\S]*the exact spelling is \{\{signature\}\}/);
  await mustRefuse('refusal: a history token this mode has no block for', 'draft_email',
    { mode: 'new', from: FROM, to: TO, subject: SUBJ, htmlBody: '<p>{{quote}}</p>' },
    /\{\{quote\}\} does not apply to mode:'new'/);
  await mustRefuse('refusal: the same token twice in one part', 'draft_email',
    { mode: 'new', from: FROM, to: TO, subject: SUBJ, htmlBody: '<p>{{signature}}</p><p>{{signature}}</p>' },
    /\{\{signature\}\} appears 2 times in htmlBody/);
  await mustRefuse('refusal: a token in one supplied part but not the other', 'draft_email',
    { mode: 'new', from: FROM, to: TO, subject: SUBJ, htmlBody: '<p>{{signature}}</p>', textBody: 'no token here' },
    /\{\{signature\}\} is in htmlBody but not in textBody/);

  // ---- 4. The bodyHash gate on edit_draft ---------------------------------------------
  check('bodyHash: the read that shows every stored byte issues one', typeof stored.hash === 'string' && stored.hash.length > 0,
    stored.withheld ?? '(neither a hash nor a withheld reason came back)');

  await mustRefuse('bodyHash: a body edit carrying none is refused', 'edit_draft',
    { emailId: d1, htmlBody: '<p>no hash</p>' },
    /needs bodyHash/);
  await mustRefuse('bodyHash: a stale one is refused', 'edit_draft',
    { emailId: d1, bodyHash: 'not-this-drafts-hash', htmlBody: '<p>stale hash</p>' },
    /not this draft's current one/);

  // An unflagged edit stores the body byte for byte, so a {{signature}} written into it
  // ships with its braces showing - and the result has to say so rather than let it pass.
  r = await c.call('edit_draft', {
    emailId: d1, bodyHash: stored.hash,
    htmlBody: '<p>Edited.</p><div>{{signature}}</div>',
    textBody: 'Edited.\n\n{{signature}}',
  });
  check('unflagged edit: accepted with the current hash', !r.isError, text(r).slice(0, 300));
  const d3 = keep(r);
  check('unflagged edit: the stored-as-written note fired',
    /\{\{signature\}\} tokens? the stored body did not/.test(text(r)), text(r).slice(0, 400));
  check('unflagged edit: a hash for the next edit came back',
    /Body hash for your next edit of this draft: \S+/.test(text(r)), text(r).slice(-200));
  const unflagged = await readDraft(d3);
  check('unflagged edit: the token is stored, not expanded',
    TOKEN_RE.test(unflagged.html) && TOKEN_RE.test(unflagged.txt),
    `${unflagged.html.slice(0, 120)} | ${unflagged.txt.slice(0, 120)}`);

  // ---- 5. The flag is the only thing that expands on an edit --------------------------
  await mustRefuse('flagged edit: the flag with no token to expand is refused', 'edit_draft',
    { emailId: d3, bodyHash: unflagged.hash, expandSignature: true, htmlBody: '<p>no token</p>', textBody: 'no token' },
    /carries no \{\{signature\}\}/);

  r = await c.call('edit_draft', {
    emailId: d3, bodyHash: unflagged.hash, expandSignature: true,
    htmlBody: '<p>Signed edit.</p><div>{{signature}}</div>',
    textBody: 'Signed edit.\n\n{{signature}}',
  });
  check('flagged edit: accepted', !r.isError, text(r).slice(0, 300));
  const d4 = keep(r);
  const flagged = await readDraft(d4);
  check('flagged edit: no token survives the expansion',
    !TOKEN_RE.test(flagged.html) && !TOKEN_RE.test(flagged.txt),
    `${flagged.html.slice(0, 120)} | ${flagged.txt.slice(0, 120)}`);
  if (signed && SIG_WORDS) {
    check('flagged edit: the same configured sign-off landed',
      words(flagged.html).includes(SIG_WORDS), `looking for "${SIG_WORDS.slice(0, 60)}" in ${words(flagged.html).slice(0, 200)}`);
  }
} finally {
  // Every artifact is a draft this probe created; an edit already moved each superseded
  // copy to Trash, so a repeat delete is a no-op. The failures are COUNTED rather than
  // swallowed: a cleanup nobody checks is one that can stop working unnoticed.
  let undeleted = 0;
  for (const id of trash) {
    try { const r = await c.call('delete_email', { emailId: id }); if (r.isError) undeleted++; } catch { undeleted++; }
  }
  check('cleanup: every draft this probe created was trashed', undeleted === 0,
    undeleted ? `${undeleted} of ${trash.length} could not be deleted; search Drafts for "${SUBJ}"` : '');
  console.log(`trashed ${trash.length - undeleted} of ${trash.length} probe drafts`);
  await c.close();
}

console.log(failures() === 0 ? '\nBODY-TOKENS: ALL PASS' : `\nBODY-TOKENS: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
