// The sending identity, and the sign-off read off it (#33).
//
// This file covers ONE step of the signature story: turning a list of identities and an
// optional `from` address into the two configured sign-off forms, or into nothing. That is
// all `identity.ts` does, and it is the step every compose path starts from.
//
// The rest of the story lives where the code does. Which FORM a part gets, and what a
// sign-off looks like once it is a block, are `signatureTextBlock` / `signatureHtmlBlock` /
// `signatureBlock` in `reply-quote.ts`, pinned in `reply-quote.test.ts`. WHERE the block
// lands is the caller's — `{{signature}}` is placed by whoever writes the body — and what
// `draft_email` does with it is pinned in `draft-email-handler.test.ts`.
//
// The one thing worth stating here, because it is the trap in this step rather than in
// either of the others: a `from` that names nothing verified resolves to NO SIGNATURE and
// not to an error. `createDraft` raises the real "not verified for sending" refusal a moment
// later, and a signature lookup that threw its own version first would replace an accurate
// message with an oblique one.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { selectIdentity, signatureOf } from './identity.js';

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
const UNSIGNED_IDENTITY = { id: 'id-2', name: 'Test User', email: 'me@example.com', mayDelete: false };

// What every compose path does: pick the identity, then read its sign-off. Written out
// rather than wrapped in a helper because that is the shape production uses — the caller
// needs the identity OBJECT too, for the address the "nothing was appended" note names.
const sigFor = (identities: any[], from?: string) => signatureOf(selectIdentity(identities, from));

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
    assert.equal(sigFor([SIGNED_IDENTITY, other], 'OTHER@example.com')?.text, 'Other');
  });

  it('falls back to the identity that cannot be deleted when no from is given', () => {
    const first = { id: 'id-0', email: 'first@example.com', mayDelete: true };
    assert.equal(selectIdentity([first, SIGNED_IDENTITY]), SIGNED_IDENTITY);
  });

  it('honours a wildcard identity', () => {
    const wild = { id: 'id-w', email: '*@example.com', mayDelete: true, textSignature: 'Wild' };
    assert.equal(sigFor([wild], 'anything@example.com')?.text, 'Wild');
  });

  it('resolves to no signature — not an error — when from names nothing verified', () => {
    assert.equal(sigFor([SIGNED_IDENTITY], 'stranger@elsewhere.example'), undefined);
  });
});
