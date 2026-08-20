// The sending identity's signature on the compose surface (#33).
//
// Four things are worth pinning here, and each fails silently on its own:
//
//  1. WHERE the block lands. A sign-off below the quoted history reads as part of the
//     quote, so every path has to put it above — including the two branches where the reply
//     builder returns the caller's bodies untouched and never reaches the quote code at all.
//  2. WHICH form the text part gets. It is derived from the html signature whenever html
//     ships, because that is what the first html-only edit will regenerate; a verbatim
//     textSignature would look correct and then change by itself.
//  3. That an edit PRESERVES a signature the draft already carries. The merge rule drops the
//     unwritten partner on any body write, so an htmlBody-alone edit is exactly where a
//     signature disappears without anyone noticing — which is why the re-append is keyed on
//     the stored draft rather than on the flag, and why it is announced.
//  4. That nothing happens when nobody asked. Default off, and an identity with no signature
//     configured appends nothing and says nothing.

import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import {
  applySignature, hasSignatureMarker, signatureHtmlBlock, signatureTextBlock,
  buildReplyBodies, buildForwardBodies, NOTE_SIGNATURE_REAPPENDED,
} from './reply-quote.js';
import { resolveSignature, selectIdentity, signatureOf } from './identity.js';
import { composeDraft } from './compose-handler.js';
import type { ComposeClient } from './compose-handler.js';
import { composeReply } from './reply-handler.js';
import type { ReplyClient } from './reply-handler.js';
import { composeForward } from './forward-handler.js';
import type { ForwardClient } from './forward-handler.js';
import { JmapClient } from './jmap-client.js';
import type { JmapRequest } from './jmap-client.js';
import { FastmailAuth } from './auth.js';

// The account's default identity, signed. The html and text forms deliberately DIFFER in
// wording so a test can tell which one a body got: the html form says "Kind regards", the
// configured text form says "Regards". Real identities keep the two in sync (Fastmail writes
// both), which is exactly why a bug here would be invisible against a matched pair.
const SIGNED_IDENTITY = {
  id: 'id-1',
  name: 'Test User',
  email: 'me@example.com',
  mayDelete: false,
  textSignature: 'Regards,\nTest User',
  htmlSignature: '<div>Kind regards,</div><div>Test User</div>',
};
const SIG = signatureOf(SIGNED_IDENTITY)!;
const UNSIGNED_IDENTITY = { id: 'id-2', name: 'Test User', email: 'me@example.com', mayDelete: false };

// What the html signature derives to. Every "the text part is derived" assertion compares
// against this rather than against the configured textSignature.
const DERIVED_TEXT = 'Kind regards,\nTest User';

// ---------------------------------------------------------------------------
// resolveSignature / identity selection
// ---------------------------------------------------------------------------

describe('reading the sign-off off an identity', () => {
  it('returns both configured forms', () => {
    assert.deepEqual(signatureOf(SIGNED_IDENTITY), {
      html: SIGNED_IDENTITY.htmlSignature,
      text: SIGNED_IDENTITY.textSignature,
    });
  });

  it('treats an identity with no signature as signature-less', () => {
    assert.equal(signatureOf(UNSIGNED_IDENTITY), undefined);
  });

  it('treats a blank signature as no signature', () => {
    assert.equal(signatureOf({ textSignature: '   ', htmlSignature: '' }), undefined);
  });

  it('keeps a half-configured identity, whichever half it is', () => {
    assert.deepEqual(signatureOf({ htmlSignature: '<div>Bye</div>' }), { html: '<div>Bye</div>' });
    assert.deepEqual(signatureOf({ textSignature: 'Bye' }), { text: 'Bye' });
  });

  it('selects the identity matching an explicit from, not the default', () => {
    const other = { id: 'id-9', email: 'other@example.com', mayDelete: true, textSignature: 'Other' };
    assert.equal(selectIdentity([SIGNED_IDENTITY, other], 'other@example.com'), other);
    assert.equal(resolveSignature([SIGNED_IDENTITY, other], 'OTHER@example.com')?.text, 'Other');
  });

  it('falls back to the identity that cannot be deleted when no from is given', () => {
    const first = { id: 'id-0', email: 'first@example.com', mayDelete: true };
    assert.equal(selectIdentity([first, SIGNED_IDENTITY]), SIGNED_IDENTITY);
  });

  it('honours a wildcard identity', () => {
    const wild = { id: 'id-w', email: '*@example.com', mayDelete: true, textSignature: 'Wild' };
    assert.equal(resolveSignature([wild], 'anything@example.com')?.text, 'Wild');
  });

  // createDraft raises the real "not verified for sending" refusal a moment later; a
  // signature lookup that threw its own version first would replace an accurate message
  // with an oblique one.
  it('resolves to no signature — not an error — when from names nothing verified', () => {
    assert.equal(resolveSignature([SIGNED_IDENTITY], 'stranger@elsewhere.example'), undefined);
  });
});

// ---------------------------------------------------------------------------
// applySignature — the insertion itself
// ---------------------------------------------------------------------------

describe('applySignature', () => {
  it('appends a marked block to an html body', () => {
    const out = applySignature({ htmlBody: '<p>Hi</p>' }, SIG);
    assert.equal(out.htmlBody, '<p>Hi</p><div><br></div><div class="fm-mcp-signature"><div>Kind regards,</div><div>Test User</div></div>');
    assert.ok(hasSignatureMarker(out.htmlBody));
    assert.equal(out.textBody, undefined); // no text part to sign; it is derived downstream
  });

  it('DERIVES the text part from the html signature when html ships', () => {
    const out = applySignature({ textBody: 'Hi', htmlBody: '<p>Hi</p>' }, SIG);
    // Not the configured textSignature ("Regards,"): an html-only edit regenerates the text
    // part from the html, so a verbatim value would silently change on that edit.
    assert.equal(out.textBody, `Hi\n\n${DERIVED_TEXT}`);
    assert.doesNotMatch(out.textBody!, /^Regards,/m);
  });

  it('uses the configured text signature when no html ships', () => {
    const out = applySignature({ textBody: 'Hi' }, SIG);
    assert.equal(out.textBody, 'Hi\n\nRegards,\nTest User');
  });

  it('treats a present-but-blank html body as no html at all', () => {
    // buildBodyParts drops a blank body, so signing it would put the sign-off in a part no
    // recipient ever sees while the text part went unsigned.
    const out = applySignature({ textBody: 'Hi', htmlBody: '   ' }, SIG);
    assert.equal(out.textBody, 'Hi\n\nRegards,\nTest User');
    assert.equal(out.htmlBody, '   ');
  });

  it('escapes a text-only identity into the html body rather than dropping it', () => {
    const out = applySignature({ htmlBody: '<p>Hi</p>' }, { text: 'Bye <boss>\nMe' });
    assert.equal(out.htmlBody, '<p>Hi</p><div><br></div><div class="fm-mcp-signature">Bye &lt;boss&gt;<br>Me</div>');
  });

  it('derives text from an html-only identity when no html ships', () => {
    const out = applySignature({ textBody: 'Hi' }, { html: '<div>Bye</div><div>Me</div>' });
    assert.equal(out.textBody, 'Hi\n\nBye\nMe');
  });

  it('leaves a body that already carries the marker alone', () => {
    const already = '<p>Hi</p><div class="fm-mcp-signature"><div>Mine</div></div>';
    const out = applySignature({ textBody: 'Hi', htmlBody: already }, SIG);
    assert.equal(out.htmlBody, already);
    assert.equal(out.textBody, 'Hi');
  });

  it('changes nothing when the identity has no signature', () => {
    const bodies = { textBody: 'Hi', htmlBody: '<p>Hi</p>' };
    assert.deepEqual(applySignature(bodies, undefined), bodies);
  });

  it('preserves definedness exactly, so no path gains a body it did not have', () => {
    const out = applySignature({ htmlBody: '<p>Hi</p>' }, SIG);
    assert.ok(!('textBody' in out) || out.textBody === undefined);
    assert.deepEqual(applySignature({}, SIG), {});
  });

  it('recognises the marker through a re-serialized attribute and a shared class list', () => {
    assert.ok(hasSignatureMarker("<div class='fm-mcp-signature'>x</div>"));
    assert.ok(hasSignatureMarker('<div class=fm-mcp-signature>x</div>'));
    assert.ok(hasSignatureMarker('<div id="s" class="sig fm-mcp-signature">x</div>'));
    // A longer class that merely starts with ours over-matches, the same way the
    // gmail_quote arm of hasQuoteMarker does. Recorded rather than fixed because it can
    // only ever suppress an append (the body is read as already signed), never cause a
    // loss — and narrowing it would cost the tolerance the round trip needs.
    assert.ok(hasSignatureMarker('<div class="fm-mcp-signature-ish">x</div>'));
    assert.ok(!hasSignatureMarker('<div class="notfm-mcp-signature">x</div>'));
    assert.ok(!hasSignatureMarker('<div>Kind regards</div>'));
    assert.ok(!hasSignatureMarker('<blockquote class="fm-mcp-signature">x</blockquote>'));
  });

  it('exposes the two block forms on their own for a body that has nothing to append to', () => {
    assert.equal(signatureHtmlBlock(undefined), undefined);
    assert.equal(signatureTextBlock(undefined, false), undefined);
    assert.equal(signatureTextBlock(SIG, true), DERIVED_TEXT);
    assert.equal(signatureTextBlock(SIG, false), 'Regards,\nTest User');
  });
});

// ---------------------------------------------------------------------------
// Placement: above the quoted / forwarded history
// ---------------------------------------------------------------------------

const ORIGINAL = {
  id: 'orig-1',
  messageId: ['<orig@example.com>'],
  from: [{ name: 'Alice', email: 'alice@example.com' }],
  sentAt: '2026-01-02T03:04:05Z',
  subject: 'Project update',
  textBody: [{ partId: 't', type: 'text/plain' }],
  htmlBody: [{ partId: 'h', type: 'text/html' }],
  bodyValues: { t: { value: 'Original text.' }, h: { value: '<p>Original html.</p>' } },
};

describe('signature placement in a reply', () => {
  it('sits between the reply body and the quote in html', () => {
    const out = buildReplyBodies({
      original: ORIGINAL, htmlBody: '<p>Thanks.</p>', quoteOriginal: true, signature: SIG,
    });
    const sig = out.htmlBody!.indexOf('fm-mcp-signature');
    const quote = out.htmlBody!.indexOf('<blockquote');
    assert.ok(sig > 0 && quote > 0, out.htmlBody);
    assert.ok(sig < quote, `signature must precede the quote: ${out.htmlBody}`);
  });

  it('sits between the reply body and the quote in text', () => {
    const out = buildReplyBodies({
      original: ORIGINAL, textBody: 'Thanks.', quoteOriginal: true, signature: SIG,
    });
    assert.match(out.textBody!, /^Thanks\.\n\nRegards,\nTest User\n\nOn .* wrote:\n> Original text\./);
  });

  it('derives the text signature when the reply also ships html', () => {
    const out = buildReplyBodies({
      original: ORIGINAL, textBody: 'Thanks.', htmlBody: '<p>Thanks.</p>',
      quoteOriginal: true, signature: SIG,
    });
    assert.ok(out.textBody!.startsWith(`Thanks.\n\n${DERIVED_TEXT}\n\nOn `), out.textBody);
  });

  // The two branches that return the caller's bodies through passthrough() and never reach
  // the quote assembly. A signature hooked into the quoting branch alone would be silently
  // missing from both.
  it('still signs when quoteOriginal is false', () => {
    const out = buildReplyBodies({
      original: ORIGINAL, htmlBody: '<p>Thanks.</p>', quoteOriginal: false, signature: SIG,
    });
    assert.ok(hasSignatureMarker(out.htmlBody), out.htmlBody);
    assert.doesNotMatch(out.htmlBody!, /<blockquote/);
  });

  it('still signs when the original has nothing quotable', () => {
    const empty = { ...ORIGINAL, textBody: [], htmlBody: [], bodyValues: {} };
    const out = buildReplyBodies({
      original: empty, htmlBody: '<p>Thanks.</p>', quoteOriginal: true, signature: SIG,
    });
    assert.ok(hasSignatureMarker(out.htmlBody), out.htmlBody);
    assert.doesNotMatch(out.htmlBody!, /<blockquote/);
  });

  it('appends nothing when no signature is passed', () => {
    const out = buildReplyBodies({ original: ORIGINAL, htmlBody: '<p>Thanks.</p>', quoteOriginal: true });
    assert.ok(!hasSignatureMarker(out.htmlBody));
  });
});

describe('signature placement in a forward', () => {
  it('sits between the note and the forwarded-message block in html', () => {
    const out = buildForwardBodies({ original: ORIGINAL, htmlBody: '<p>FYI.</p>', signature: SIG });
    const sig = out.htmlBody!.indexOf('fm-mcp-signature');
    const block = out.htmlBody!.indexOf('----- Original message -----');
    assert.ok(sig > 0 && block > 0, out.htmlBody);
    assert.ok(sig < block, `signature must precede the forwarded block: ${out.htmlBody}`);
  });

  it('sits between the note and the forwarded-message block in text', () => {
    const out = buildForwardBodies({ original: ORIGINAL, textBody: 'FYI.', signature: SIG });
    const sig = out.textBody!.indexOf('Regards,');
    const block = out.textBody!.indexOf('----- Original message -----');
    assert.ok(sig > 0 && block > 0 && sig < block, out.textBody);
  });

  // A bare FYI forward has no note for the signature to hang off, so the block becomes the
  // whole of the content above the forwarded message.
  it('signs a forward that carries no note at all', () => {
    const out = buildForwardBodies({ original: ORIGINAL, signature: SIG });
    assert.ok(hasSignatureMarker(out.htmlBody), out.htmlBody);
    assert.ok(out.htmlBody!.indexOf('fm-mcp-signature') < out.htmlBody!.indexOf('Original message'));
  });

  it('leaves a note-less forward unsigned when nothing asked', () => {
    const out = buildForwardBodies({ original: ORIGINAL });
    assert.ok(!hasSignatureMarker(out.htmlBody));
  });
});

// ---------------------------------------------------------------------------
// The compose orchestrations
// ---------------------------------------------------------------------------

function composeClient(identities: any[] = [SIGNED_IDENTITY]) {
  const calls: any = { identityLookups: 0 };
  const client: ComposeClient = {
    getIdentities: async () => { calls.identityLookups += 1; return identities; },
    uploadAttachments: async () => [],
    createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    getEmailById: async (id) => ({ id, attachments: [] }),
  };
  return { client, calls };
}

describe('create_draft signs on request', () => {
  it('appends the signature to the html body', async () => {
    const { client, calls } = composeClient();
    await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: true }, client, undefined, false);
    assert.ok(hasSignatureMarker(calls.draft.htmlBody), calls.draft.htmlBody);
  });

  it('accepts the string form a lenient client sends', async () => {
    const { client, calls } = composeClient();
    await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: 'true' }, client, undefined, false);
    assert.ok(hasSignatureMarker(calls.draft.htmlBody));
  });

  it('reads "false" as false rather than as a truthy string', async () => {
    const { client, calls } = composeClient();
    await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: 'false' }, client, undefined, false);
    assert.ok(!hasSignatureMarker(calls.draft.htmlBody));
    assert.equal(calls.identityLookups, 0);
  });

  it('costs no identity lookup when the flag is absent', async () => {
    const { client, calls } = composeClient();
    await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>' }, client, undefined, false);
    assert.equal(calls.identityLookups, 0);
    assert.equal(calls.draft.htmlBody, '<p>Hi</p>');
  });

  it('appends nothing, and says nothing, for an identity with no signature', async () => {
    const { client, calls } = composeClient([UNSIGNED_IDENTITY]);
    const r = await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: true }, client, undefined, false);
    assert.equal(calls.draft.htmlBody, '<p>Hi</p>');
    assert.equal(r.notes, undefined);
  });

  it('signs with the identity named by from, not the default', async () => {
    const other = { id: 'id-9', email: 'other@example.com', mayDelete: true, htmlSignature: '<div>From Other</div>' };
    const { client, calls } = composeClient([SIGNED_IDENTITY, other]);
    await composeDraft(
      { to: ['bob@example.com'], from: 'other@example.com', htmlBody: '<p>Hi</p>', appendSignature: true },
      client, undefined, false,
    );
    assert.match(calls.draft.htmlBody, /From Other/);
  });

  // A signature is a sign-off on a message, not a message: it must not rescue a call that
  // supplied no content of its own.
  it('does not let a signature turn a contentless call into a draft', async () => {
    const { client } = composeClient();
    await assert.rejects(
      () => composeDraft({ appendSignature: true }, client, undefined, false),
      /At least one of to, subject, textBody, htmlBody, or attachments/,
    );
  });
});

describe('reply_email and forward_email sign on request', () => {
  function replyClient() {
    const calls: any = { identityLookups: 0 };
    const client: ReplyClient = {
      getEmailById: async (id) => (id === 'draft-9' ? { id, attachments: [] } : ORIGINAL),
      getIdentities: async () => { calls.identityLookups += 1; return [SIGNED_IDENTITY]; },
      uploadAttachments: async () => [],
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    };
    return { client, calls };
  }

  function forwardClient() {
    const calls: any = { identityLookups: 0 };
    const client: ForwardClient = {
      getEmailById: async (id) => (id === 'draft-7' ? { id, attachments: [] } : { ...ORIGINAL, blobId: 'blob-eml' }),
      getIdentities: async () => { calls.identityLookups += 1; return [SIGNED_IDENTITY]; },
      uploadAttachments: async () => [],
      createDraft: async (p) => { calls.draft = p; return 'draft-7'; },
    };
    return { client, calls };
  }

  it('puts the reply signature above the quote', async () => {
    const { client, calls } = replyClient();
    await composeReply({ originalEmailId: 'orig-1', htmlBody: '<p>Thanks.</p>', appendSignature: true }, client, undefined, false);
    assert.ok(calls.draft.htmlBody.indexOf('fm-mcp-signature') < calls.draft.htmlBody.indexOf('<blockquote'));
  });

  it('leaves an unsigned reply unsigned and skips the lookup', async () => {
    const { client, calls } = replyClient();
    await composeReply({ originalEmailId: 'orig-1', htmlBody: '<p>Thanks.</p>' }, client, undefined, false);
    assert.ok(!hasSignatureMarker(calls.draft.htmlBody));
    assert.equal(calls.identityLookups, 0);
  });

  it('puts the forward signature above the forwarded-message block', async () => {
    const { client, calls } = forwardClient();
    await composeForward({ originalEmailId: 'orig-1', to: ['bob@example.com'], htmlBody: '<p>FYI.</p>', appendSignature: true }, client, undefined, false);
    assert.ok(calls.draft.htmlBody.indexOf('fm-mcp-signature') < calls.draft.htmlBody.indexOf('Original message'));
  });

  // asAttachment has no forwarded block to sit above: the note IS the body, and it has its
  // own assembly path that bypasses the block builder entirely.
  it('signs the note on an asAttachment forward', async () => {
    const { client, calls } = forwardClient();
    await composeForward(
      { originalEmailId: 'orig-1', to: ['bob@example.com'], htmlBody: '<p>FYI.</p>', asAttachment: true, appendSignature: true },
      client, undefined, false,
    );
    assert.ok(hasSignatureMarker(calls.draft.htmlBody), calls.draft.htmlBody);
  });

  it('signs the filler body of a note-less asAttachment forward', async () => {
    const { client, calls } = forwardClient();
    await composeForward(
      { originalEmailId: 'orig-1', to: ['bob@example.com'], asAttachment: true, appendSignature: true },
      client, undefined, false,
    );
    assert.equal(calls.draft.textBody, 'Forwarded message attached.\n\nRegards,\nTest User');
  });
});

// ---------------------------------------------------------------------------
// edit_draft: preservation, and the announcement
// ---------------------------------------------------------------------------

// A local harness over the client's single outbound seam, deliberately small: these tests
// only need Email/get to hand back a fixture and Email/set to accept the recreate.
const ACCOUNT_ID = 'acct-123';
const MAILBOXES = [
  { id: 'mb-drafts', name: 'Drafts', role: 'drafts' },
  { id: 'mb-trash', name: 'Trash', role: 'trash' },
];

function makeClient(identities: any[] = [SIGNED_IDENTITY]): JmapClient {
  const client = new JmapClient(new FastmailAuth({ apiToken: 'fake-token' }));
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/', accountId: ACCOUNT_ID, capabilities: {},
  }));
  mock.method(client, 'getIdentities', async () => identities);
  mock.method(client, 'getMailboxes', async () => MAILBOXES);
  return client;
}

/** Serves the draft fixture for its own id and ORIGINAL for the quoted message's. */
function mockUpdate(client: JmapClient, fixture: any) {
  const seen: any = {};
  mock.method(client, 'makeRequest', async (request: JmapRequest) => {
    const [method, params] = request.methodCalls[0] as [string, any, string];
    if (method === 'Email/get') {
      const wanted = params.ids?.[0];
      return { methodResponses: [['Email/get', { list: [wanted === 'orig-1' ? ORIGINAL : fixture] }, 'getEmail']] };
    }
    if (params.create) {
      seen.created = params.create.draft;
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
    }
    return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
  });
  return seen;
}

const SIGNED_BODY = '<p>Hello.</p><div><br></div><div class="fm-mcp-signature"><div>Kind regards,</div><div>Test User</div></div>';

function signedDraft(over: any = {}) {
  return {
    id: 'draft-1',
    subject: 'Old Subject',
    from: [{ email: 'me@example.com' }],
    to: [{ email: 'bob@example.com' }],
    cc: [],
    bcc: [],
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: `Hello.\n\n${DERIVED_TEXT}` }, h: { value: SIGNED_BODY } },
    mailboxIds: { 'mb-drafts': true },
    keywords: { $draft: true },
    ...over,
  };
}

describe('edit_draft preserves a signature the draft already carries', () => {
  it('re-appends it on an htmlBody-alone edit and says so', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.ok(hasSignatureMarker(seen.created.bodyValues.html.value), seen.created.bodyValues.html.value);
    assert.deepEqual(result.notes, [NOTE_SIGNATURE_REAPPENDED]);
    // And the regenerated text fallback carries it too, derived from the signed html.
    assert.match(seen.created.bodyValues.text.value, /Kind regards,\nTest User$/);
  });

  it('drops it when the edit says appendSignature:false, and says nothing', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>', appendSignature: false });
    assert.equal(seen.created.bodyValues.html.value, '<p>Rewritten.</p>');
    assert.equal(result.notes, undefined);
  });

  it('does not add a second block when the new body already carries one', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Rewritten.</p><div class="fm-mcp-signature"><div>Mine</div></div>',
    });
    assert.equal(seen.created.bodyValues.html.value.match(/fm-mcp-signature/g)!.length, 1);
    assert.equal(result.notes, undefined);
  });

  it('leaves a draft that never carried one alone', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft({
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { t: { value: 'Hello.' }, h: { value: '<p>Hello.</p>' } },
    }));

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.equal(seen.created.bodyValues.html.value, '<p>Rewritten.</p>');
    assert.equal(result.notes, undefined);
  });

  // The flag is also how a signature gets ONTO a draft that never had one. That is a request
  // made in this call, so it earns no announcement — the caller already knows.
  it('adds one on request to a draft that never had one, without a note', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft({
      bodyValues: { t: { value: 'Hello.' }, h: { value: '<p>Hello.</p>' } },
    }));

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>', appendSignature: true });
    assert.ok(hasSignatureMarker(seen.created.bodyValues.html.value));
    assert.equal(result.notes, undefined);
  });

  it('appends nothing when the identity no longer has a signature configured', async () => {
    const client = makeClient([UNSIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.equal(seen.created.bodyValues.html.value, '<p>Rewritten.</p>');
    assert.equal(result.notes, undefined);
  });

  it('leaves both bodies untouched on a metadata-only edit', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(seen.created.bodyValues.html.value, SIGNED_BODY);
    assert.equal(result.notes, undefined);
  });

  it('signs the text body when the edit converts the draft to plain text', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    await client.updateDraft('draft-1', { textBody: 'Rewritten.', clearFields: ['htmlBody'] });
    // No html ships now, so the configured text form is the right one to write.
    assert.equal(seen.created.bodyValues.text.value, 'Rewritten.\n\nRegards,\nTest User');
    assert.equal(seen.created.bodyValues.html, undefined);
  });

  // The ordering trap: after the quote rebuild updates.htmlBody already carries the quoted
  // original, so a signature appended there would land underneath it.
  it('keeps the re-appended signature above a rebuilt quote', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft({
      inReplyTo: ['<orig@example.com>'],
      bodyValues: {
        t: { value: `Hello.\n\n${DERIVED_TEXT}\n\nOn Fri, Alice wrote:\n> Original text.` },
        h: { value: `${SIGNED_BODY}<blockquote type="cite"><p>Original html.</p></blockquote>` },
      },
    }));

    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Rewritten.</p>', originalEmailId: 'orig-1',
    });
    const html = seen.created.bodyValues.html.value;
    assert.ok(html.indexOf('fm-mcp-signature') < html.indexOf('<blockquote'), html);
    assert.deepEqual(result.notes, [NOTE_SIGNATURE_REAPPENDED]);
  });
});
