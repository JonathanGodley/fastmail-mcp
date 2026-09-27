// The sending identity, and the sign-off read off it (#33): identities plus an optional
// `from` in, the two configured sign-off forms (or nothing) out. The block a sign-off
// becomes is pinned in `reply-quote.test.ts`; where `draft_email` places it, in
// `draft-email-handler.test.ts`.
//
// A `from` that names nothing verified resolves to NO SIGNATURE, not an error. `createDraft`
// raises the real "not verified for sending" refusal a moment later, and a signature lookup
// that threw its own version first would replace an accurate message with an oblique one.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { matchesIdentity, selectIdentity, signatureOf } from './identity.js';

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

// What every compose path does: pick the identity, then read its sign-off.
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

  // The positive criterion: each half is printable and non-space, with no control or format
  // character and none of the characters that bracket, comment or quote in an address.
  it('refuses control, format and bracketing characters in an address a wildcard identity is asked to verify', () => {
    for (const addr of [
      'a\u0000b@example.com', 'a<b@example.com', 'a>b@example.com', 'a@example.com\u0000',
      'a\u0001b@example.com', 'a\u007fb@example.com', 'a\u202eb@example.com', 'a\u200bb@example.com',
      'a(c)@example.com', 'a\\b@example.com', 'a@exa\u00admple.com',
    ]) {
      assert.equal(matchesIdentity('*@example.com', addr), false, JSON.stringify(addr));
    }
    assert.equal(matchesIdentity('*@example.com', 'a.b+c@example.com'), true);
  });

  it('prefers an exact-address identity to a wildcard listed before it', () => {
    const wild = { id: 'id-w', email: '*@example.com', mayDelete: true, textSignature: 'Wild' };
    const exact = { id: 'id-e', email: 'ops@example.com', mayDelete: true, textSignature: 'Ops' };
    assert.equal(selectIdentity([wild, exact], 'OPS@example.com'), exact);
    assert.equal(selectIdentity([exact, wild], 'ops@example.com'), exact);
    assert.equal(selectIdentity([wild, exact], 'other@example.com'), wild);
  });

  it('resolves to no signature — not an error — when from names nothing verified', () => {
    assert.equal(sigFor([SIGNED_IDENTITY], 'stranger@elsewhere.example'), undefined);
  });
});
