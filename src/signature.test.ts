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
  buildReplyBodies, buildForwardBodies, NOTE_SIGNATURE_REAPPENDED, signatureSkipReason,
  noteSignatureNotAppended, noteSignatureUnavailableOnEdit,
} from './reply-quote.js';
import { selectIdentity, signatureOf } from './identity.js';
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

// An identity whose html sign-off carries a QUOTATION — a quote-of-the-day signature, which
// is an ordinary thing for a person to configure. It matters far out of proportion to how
// exotic it looks: html-to-text renders a <blockquote> as '> '-prefixed lines, so this is the
// only fixture in this file whose derived text block contains a line that looks like reply
// quoting. Every other signature here is a plain two-line sign-off or an <img>, and a guard
// that treated a '> ' line in the BODY as quoted history while leaving it in the BLOCK it was
// comparing against therefore passed the whole suite while duplicating this identity's
// sign-off on every pass of the read-modify-write loop, with nothing reported.
const QUOTING_IDENTITY = {
  id: 'id-q',
  name: 'Test User',
  email: 'me@example.com',
  mayDelete: false,
  textSignature: 'Regards,\nTest User\n> Per aspera ad astra',
  htmlSignature: '<div>Kind regards,</div><blockquote>Per aspera ad astra</blockquote><div>Test User</div>',
};
const QUOTING_SIG = signatureOf(QUOTING_IDENTITY)!;

// ---------------------------------------------------------------------------
// identity selection and the sign-off read off it
// ---------------------------------------------------------------------------

// What every handler does: pick the identity, then read its sign-off. Written out here
// rather than wrapped in a helper because that is the shape production uses — the handlers
// need the identity OBJECT too, for the address the "nothing was appended" note names.
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

  // createDraft raises the real "not verified for sending" refusal a moment later; a
  // signature lookup that threw its own version first would replace an accurate message
  // with an oblique one.
  it('resolves to no signature — not an error — when from names nothing verified', () => {
    assert.equal(sigFor([SIGNED_IDENTITY], 'stranger@elsewhere.example'), undefined);
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

  // A plain-text body carries no class, so the html marker cannot protect it. Without an
  // idempotence rule of its own, the read-modify-write loop the plain-text residual
  // prescribes as the recovery (read the draft, edit the text, send it back with
  // appendSignature:true) stacks a second copy every time round.
  it('does not sign a plain-text body twice when it already ends with the block', () => {
    const once = applySignature({ textBody: 'Hi Bob' }, SIG).textBody!;
    assert.equal(once, 'Hi Bob\n\nRegards,\nTest User');
    // The loop: read that body back, change the words above the sign-off, send it in again.
    const edited = once.replace('Hi Bob', 'Hi Bob, one more thing');
    const twice = applySignature({ textBody: edited }, SIG);
    assert.equal(twice.textBody, edited);
    assert.equal(twice.textBody!.match(/Regards,/g)!.length, 1);
  });

  it('reports the already-signed text body rather than silently doing nothing', () => {
    assert.equal(signatureSkipReason({ textBody: 'Hi\n\nRegards,\nTest User' }, SIG), 'already-signed');
    assert.equal(signatureSkipReason({ textBody: 'Hi' }, SIG), undefined);
  });

  // The cross-form hole. Which text block a previous call left in the body depends on
  // whether THAT call shipped html — an html-shipping one writes the form derived from
  // htmlSignature, a text-only one writes the configured textSignature — and for an identity
  // that configures both (the normal case, and what SIGNED_IDENTITY models) those differ.
  // Testing only the form THIS call would write signed a converted draft twice.
  it('does not sign twice when the body carries the OTHER form of the same signature', () => {
    // An html draft's derived text part, handed back to a text-only call.
    const converted = `Hello, one more thing.\n\n${DERIVED_TEXT}`;
    assert.equal(applySignature({ textBody: converted }, SIG).textBody, converted);
    assert.equal(signatureSkipReason({ textBody: converted }, SIG), 'already-signed');

    // …and the mirror: a text-only draft's configured sign-off, handed to a call that also
    // ships html, where the derived form is what would otherwise be appended.
    const textForm = 'Hello.\n\nRegards,\nTest User';
    const out = applySignature({ textBody: textForm, htmlBody: '<p>Hello.</p>' }, SIG);
    assert.equal(out.textBody, textForm);
    assert.ok(hasSignatureMarker(out.htmlBody)); // the unsigned html still gets one
  });

  // The edit path hands planSignature the caller's WHOLE text body, quote included, so the
  // sign-off it is looking for is not at the end — the quote is. An end-of-body test alone
  // read this as unsigned and stacked a second block underneath the quote, which is the exact
  // duplication the flag's own documentation promises cannot happen, on the one path that
  // documentation points callers at.
  it('does not sign twice when the sign-off sits above a quoted original', () => {
    const withQuote = 'Thanks.\n\nRegards,\nTest User\n\nOn 1 Jan Alice wrote:\n> original';
    assert.equal(applySignature({ textBody: withQuote }, SIG).textBody, withQuote);
    assert.equal(signatureSkipReason({ textBody: withQuote }, SIG), 'already-signed');

    // The other form above the same quote (an html draft's derived text, converted).
    const derivedAbove = `Thanks.\n\n${DERIVED_TEXT}\n\nOn 1 Jan Alice wrote:\n> original`;
    assert.equal(applySignature({ textBody: derivedAbove }, SIG).textBody, derivedAbove);
  });

  // Nothing here parses the attribution line, so the shapes that used to need special
  // handling are just ordinary bodies: the quote's own "> " prefix is what identifies it, and
  // whatever sits between the sign-off and the quote is left in the haystack rather than being
  // located and cut off. A hard-wrapped Gmail attribution — the shape that broke the previous
  // rule — and one written tight against the sign-off with no blank line are the same case.
  it('does not sign twice whatever shape the attribution above the quote takes', () => {
    const wrapped = [
      'Thanks for that.',
      '',
      'Regards,',
      'Test User',
      '',
      'On Mon, Jan 1, 2026 at 9:00 AM Alice Example <alice@example.com>',
      'wrote:',
      '',
      '> the original message',
    ].join('\n');
    assert.equal(signatureSkipReason({ textBody: wrapped }, SIG), 'already-signed');
    assert.equal(applySignature({ textBody: wrapped }, SIG).textBody, wrapped);

    // The same body with the attribution on one line, which is the shape this server writes.
    const single = wrapped.replace('<alice@example.com>\nwrote:', '<alice@example.com> wrote:');
    assert.equal(signatureSkipReason({ textBody: single }, SIG), 'already-signed');

    // …and one written directly above the attribution with no blank line between them.
    const tight = 'Thanks.\nRegards,\nTest User\nOn 1 Jan Alice wrote:\n> original';
    assert.equal(signatureSkipReason({ textBody: tight }, SIG), 'already-signed');
  });

  // A genuinely unsigned body above a quote is still signed, however its own prose runs into
  // the attribution: nothing about the attribution is read, and no line is treated specially.
  it('signs an unsigned body whose last line runs into the attribution', () => {
    const unsigned = 'Thanks.\nSee my reply below.\nOn 1 Jan Alice wrote:\n> original';
    assert.equal(signatureSkipReason({ textBody: unsigned }, SIG), undefined);
    const out = applySignature({ textBody: unsigned }, SIG).textBody!;
    assert.equal(out.match(/Regards,/g)!.length, 1, out);
  });

  // Quote depth changes nothing: the block's lines carry no '>' at all, so neither '> ' nor
  // '>> ' in front of them can equal one. A nested block carrying the sign-off is not ours.
  it('does not read a sign-off inside a nested >> quote as the draft\'s own', () => {
    const nested = [
      'Thanks.',
      '',
      'On 1 Jan Alice wrote:',
      '> On 31 Dec Test User wrote:',
      '>> Here it is.',
      '>>',
      '>> Regards,',
      '>> Test User',
    ].join('\n');
    assert.equal(signatureSkipReason({ textBody: nested }, SIG), undefined);
    const out = applySignature({ textBody: nested }, SIG).textBody!;
    assert.equal(out.match(/^Regards,$/gm)!.length, 1, out);
  });

  // A sign-off that CONTAINS a quotation. html-to-text renders <blockquote> as '> '-prefixed
  // lines, so this identity's derived text block has a line that looks exactly like reply
  // quoting — and a guard that filtered such lines out of the body while leaving them in the
  // block it compared against could never match this signature at all. The result was a
  // second sign-off on every pass of the loop, with nothing reported.
  it('recognises a sign-off whose own text contains a quoted line', () => {
    const derived = signatureTextBlock(QUOTING_SIG, true)!;
    assert.match(derived, /^> Per aspera ad astra$/m, derived); // the premise

    for (const block of [derived, signatureTextBlock(QUOTING_SIG, false)!]) {
      const body = `Hello.\n\n${block}`;
      assert.equal(applySignature({ textBody: body }, QUOTING_SIG).textBody, body, block);
      assert.equal(signatureSkipReason({ textBody: body }, QUOTING_SIG), 'already-signed', block);
    }
  });

  // …and the same signature above a real reply quote, which is where the two kinds of
  // '> ' line sit in one body at once: the block's own quotation, and the quoted original.
  it('recognises a quoting sign-off sitting above a quoted original', () => {
    const block = signatureTextBlock(QUOTING_SIG, false)!;
    const body = `Thanks.\n\n${block}\n\nOn 1 Jan Alice wrote:\n> original`;
    assert.equal(applySignature({ textBody: body }, QUOTING_SIG).textBody, body);
    assert.equal(signatureSkipReason({ textBody: body }, QUOTING_SIG), 'already-signed');
  });

  // The quoted original still does not read as this draft's own sign-off. That is the job the
  // removed line filter looked like it was doing; the whole-line comparison does it, because
  // '> Regards,' is not 'Regards,' — including when the identity itself quotes.
  it('still ignores a sign-off that is only present inside the quote', () => {
    const body = 'Thanks.\n\nOn 1 Jan Alice wrote:\n> Regards,\n> Test User\n> > Per aspera ad astra';
    assert.equal(signatureSkipReason({ textBody: body }, QUOTING_SIG), undefined);
    assert.equal(signatureSkipReason({ textBody: body }, SIG), undefined);
  });

  // A body that IS the sign-off and nothing else — the match starts at line 0, the one end of
  // the run the other fixtures never exercise (they all have text above the block).
  it('recognises a body whose every line is the sign-off', () => {
    assert.equal(signatureSkipReason({ textBody: 'Regards,\nTest User' }, SIG), 'already-signed');
    assert.equal(applySignature({ textBody: 'Regards,\nTest User' }, SIG).textBody, 'Regards,\nTest User');
  });

  // A configured signature that opens or closes with a blank line: the block is compared at
  // its non-blank extent, because the body need not carry a matching blank line on that side.
  it('is not defeated by a signature configured with blank lines around it', () => {
    const padded = { text: '\nRegards,\nTest User\n' };
    assert.equal(signatureSkipReason({ textBody: 'Hi\n\nRegards,\nTest User' }, padded), 'already-signed');
    // Idempotent through its own append, too — the blank lines it writes change nothing.
    const once = applySignature({ textBody: 'Hi' }, padded).textBody!;
    assert.equal(applySignature({ textBody: once }, padded).textBody, once);
    assert.equal(once.match(/Regards,/g)!.length, 1, once);
  });

  it('does not sign twice when the sign-off sits above a forwarded-message block', () => {
    const withBlock = 'FYI.\n\nRegards,\nTest User\n\n----- Original message -----\nFrom: Alice\n\nbody';
    assert.equal(applySignature({ textBody: withBlock }, SIG).textBody, withBlock);
    assert.equal(signatureSkipReason({ textBody: withBlock }, SIG), 'already-signed');
  });

  // …and the forwarded block is removed entirely rather than searched. A forward reproduces
  // the forwarded message's text UNPREFIXED, so a message that itself carries this identity's
  // sign-off would read as this draft's own and the caller's note would go out bare —
  // wherever in the forwarded block that sign-off sits, including at its very end.
  it('still signs a note above a forwarded message that carries the same sign-off', () => {
    const above = 'Please see below.\n\n----- Original message -----\nFrom: me\n\nRegards,\nTest User\n\nHi';
    assert.equal(signatureSkipReason({ textBody: above }, SIG), undefined);
    // WHERE the block lands on a caller body that already contains history is a separate
    // question this does not pin; what it pins is that the note is signed at all.
    assert.equal(applySignature({ textBody: above }, SIG).textBody!.match(/Regards,/g)!.length, 2);

    // Gmail's separator, and the sign-off as the LAST thing in the forwarded message.
    const atEnd = 'Please see below.\n\n---------- Forwarded message ----------\nFrom: me\n\nHi\n\nRegards,\nTest User';
    assert.equal(signatureSkipReason({ textBody: atEnd }, SIG), undefined);
    assert.equal(applySignature({ textBody: atEnd }, SIG).textBody!.match(/Regards,/g)!.length, 2);
  });

  // The residual, pinned so it reads as a decision rather than a surprise. A forwarded block
  // in a shape no separator names — Outlook's From:/Sent:/To:/Subject: header block — is not
  // removed, so a forwarded message carrying this identity's own sign-off reads as this
  // draft's own and the note goes out bare. That is the CHEAP error by construction: it is
  // announced as already-signed, and the caller can write the sign-off into the note. The
  // expensive error is the other one — two sign-offs shipped in silence. Extending separator
  // recognition is issue #144.
  it('reports already-signed for a forward whose separator this server does not know', () => {
    const body = [
      'Please see below.',
      '',
      'From: Alice <alice@example.com>',
      'Sent: Monday, 1 January 2026 09:00',
      'To: Test User <me@example.com>',
      'Subject: Project update',
      '',
      'Hi',
      '',
      'Regards,',
      'Test User',
    ].join('\n');
    assert.equal(signatureSkipReason({ textBody: body }, SIG), 'already-signed');
    assert.equal(applySignature({ textBody: body }, SIG).textBody, body);
  });

  // A body that already contains its own quote gets the block appended at the END, below the
  // quote, so the next round of the same loop has to find it there — the quoted lines drop
  // out of the haystack and the appended block sits in what is left.
  it('does not stack when a previous append landed below the history', () => {
    const once = applySignature(
      { textBody: 'Thanks.\n\nOn 1 Jan Alice wrote:\n> original' },
      SIG,
    ).textBody!;
    assert.equal(once.match(/Regards,/g)!.length, 1, once);
    assert.equal(applySignature({ textBody: once }, SIG).textBody, once);
    assert.equal(signatureSkipReason({ textBody: once }, SIG), 'already-signed');
  });

  // Idempotence has to survive a round trip through a mail store, which is free to change
  // line endings and (because "-- " ends in a space) to strip per-line trailing whitespace.
  it('is not defeated by CRLF line endings or stripped trailing whitespace', () => {
    const once = applySignature({ textBody: 'Hi' }, SIG).textBody!;
    assert.equal(applySignature({ textBody: once.replace(/\n/g, '\r\n') }, SIG).textBody, once.replace(/\n/g, '\r\n'));

    // RFC 3676: the delimiter ends in a space, and a mail store is free to strip it. The
    // identity configures an html form too — the ordinary Fastmail shape — so the DERIVED
    // form (which loses the trailing space on its way through html) cannot stand in for the
    // configured one here: the only thing that matches the stripped body is the per-line
    // normalisation applied to both sides.
    const delimited = { text: 'Cheers,\n-- \nTest User', html: '<div>Kind regards,</div><div>Test User</div>' };
    const signed = applySignature({ textBody: 'Hi' }, delimited).textBody!;
    assert.match(signed, /\n-- \n/); // the premise: the space is there to lose
    const stripped = signed.split('\n').map((l) => l.replace(/[ \t]+$/, '')).join('\n');
    assert.equal(applySignature({ textBody: stripped }, delimited).textBody, stripped);
    assert.equal(signatureSkipReason({ textBody: stripped }, delimited), 'already-signed');
  });

  // Three more ways the same round trip changes a body. Each of these, left unhandled, reports
  // the body unsigned and appends a SECOND sign-off with no note — the silent failure this
  // guard exists to prevent, so each is pinned rather than left to the two above.
  it('is not defeated by bare CR endings, an NBSP delimiter, or space-stuffed lines', () => {
    const delimited = { text: 'Cheers,\n-- \nTest User', html: '<div>Kind regards,</div><div>Test User</div>' };
    const signed = applySignature({ textBody: 'Hi' }, delimited).textBody!;

    // A store that ends lines with a bare CR. Without it the whole body is one line.
    const bareCr = signed.replace(/\n/g, '\r');
    assert.equal(applySignature({ textBody: bareCr }, delimited).textBody, bareCr);
    assert.equal(signatureSkipReason({ textBody: bareCr }, delimited), 'already-signed');

    // The delimiter's trailing space turned into a non-breaking space in transit.
        const nbsp = signed.replace('-- ', '--\u00a0');
    assert.equal(applySignature({ textBody: nbsp }, delimited).textBody, nbsp);
    assert.equal(signatureSkipReason({ textBody: nbsp }, delimited), 'already-signed');

    // RFC 3676 space-stuffing, applied by a format=flowed sender to every line.
    const stuffed = signed.split('\n').map((l) => ' ' + l).join('\n');
    assert.equal(applySignature({ textBody: stuffed }, delimited).textBody, stuffed);
    assert.equal(signatureSkipReason({ textBody: stuffed }, delimited), 'already-signed');
  });

  // Same rule for the derived text form on a message that also ships html: the html side is
  // protected by its marker, the text side by the equality test.
  it('does not sign the derived text part twice either', () => {
    const out = applySignature(
      { textBody: `Hi\n\n${DERIVED_TEXT}`, htmlBody: '<p>Hi</p>' },
      SIG,
    );
    assert.equal(out.textBody, `Hi\n\n${DERIVED_TEXT}`);
    assert.ok(hasSignatureMarker(out.htmlBody)); // the html was unsigned, so it still gets one
  });

  // An html signature that is nothing but a REMOTE logo derives to '' under every image
  // policy (a remote image never writes a placeholder — see ImagePlaceholderPolicy), and
  // joining that appended two newlines and no sign-off — a text part that gained trailing
  // whitespace and nothing else, while the html part was correctly signed.
  const IMAGE_ONLY_SIG = { html: '<img src="https://example.com/logo.png">' };

  it('leaves the text part alone when the html signature has no text form', () => {
    assert.equal(signatureTextBlock(IMAGE_ONLY_SIG, true), ''); // the premise
    const out = applySignature({ textBody: 'hello', htmlBody: '<p>hello</p>' }, IMAGE_ONLY_SIG);
    assert.equal(out.textBody, 'hello');
    assert.ok(hasSignatureMarker(out.htmlBody));
  });

  // …and says so. Leaving it alone is right — there is nothing to write — but the html half
  // succeeded, so this is a PARTIAL: a recipient rendering the html sees the sign-off and one
  // reading the text alternative does not. A success report over a half-signed message is the
  // silent half-result the whole reporting rule exists to prevent.
  it('reports the text part it could not sign on a message whose html it did', () => {
    assert.equal(
      signatureSkipReason({ textBody: 'hello', htmlBody: '<p>hello</p>' }, IMAGE_ONLY_SIG),
      'text-part-unsigned',
    );
  });

  // The same outcome with NO text part supplied, which used to be silent. The text
  // alternative of an html-only message is derived downstream from the signed html, so it
  // derives exactly what the branch above would have written — nothing. The recipient reading
  // it sees no sign-off either way, so the report cannot depend on which bodies the caller
  // happened to pass.
  it('reports the derived text alternative it could not sign either', () => {
    assert.equal(signatureSkipReason({ htmlBody: '<p>hello</p>' }, IMAGE_ONLY_SIG), 'text-part-unsigned');
    // The html half still landed: this is a partial, not a refusal.
    assert.ok(hasSignatureMarker(applySignature({ htmlBody: '<p>hello</p>' }, IMAGE_ONLY_SIG).htmlBody));
  });

  // An EMBEDDED image is the other spelling of "my signature is a logo", and it is NOT the
  // same outcome: the image ships with the html, so the text alternative gets the "[image]"
  // placeholder the unconditional policy exists for, and the sign-off is present rather than
  // missing. Pinned with a cid: fixture because the https one above never produced a
  // placeholder in the first place and so could not tell the two branches apart.
  it('writes the embedded-image placeholder into the text part when the html ships', () => {
    const embedded = { html: '<img src="cid:logo">' };
    assert.equal(signatureTextBlock(embedded, true), '[image]'); // the premise
    const out = applySignature({ textBody: 'hello', htmlBody: '<p>hello</p>' }, embedded);
    assert.equal(out.textBody, 'hello\n\n[image]');
    assert.equal(signatureSkipReason({ textBody: 'hello', htmlBody: '<p>hello</p>' }, embedded), undefined);
    // …and it agrees with what the derivation downstream would make of the signed html, which
    // is the whole reason this branch does not suppress: the two must not disagree.
    assert.equal(signatureSkipReason({ htmlBody: '<p>hello</p>' }, embedded), undefined);
  });

  it('reports a text-only message it cannot sign at all', () => {
    assert.deepEqual(applySignature({ textBody: 'hello' }, IMAGE_ONLY_SIG), { textBody: 'hello' });
    assert.equal(signatureSkipReason({ textBody: 'hello' }, IMAGE_ONLY_SIG), 'no-text-form');
  });

  // An EMBEDDED image derives a "[image]" placeholder under the unconditional policy, which
  // is right for the text alternative of a body that ships the image — and wrong here. No
  // html ships, so no image ships, and the recipient's entire sign-off would be a line
  // describing something no part of the message carries. Same rule the quote builders apply.
  it('never ships a bare image placeholder as the sign-off of a text-only message', () => {
    const embedded = { html: '<img src="cid:logo">' };
    assert.equal(signatureTextBlock(embedded, false), '');
    assert.deepEqual(applySignature({ textBody: 't' }, embedded), { textBody: 't' });
    assert.equal(signatureSkipReason({ textBody: 't' }, embedded), 'no-text-form');
  });

  it('still derives the alt text of an image signature that has some', () => {
    const alted = { html: '<img src="cid:logo" alt="Test User, Example Ltd">' };
    assert.equal(signatureTextBlock(alted, false), 'Test User, Example Ltd');
    assert.equal(applySignature({ textBody: 't' }, alted).textBody, 't\n\nTest User, Example Ltd');
  });

  // A blank text body is no body, exactly as a blank html body is. It used to fall through to
  // the join, whose blank-body arm made the signature the ENTIRE body — the caller's text
  // discarded, a signature-only message stored, and no note, while every wording on every
  // surface says a blank body counts as none.
  it('treats a present-but-blank text body as no body at all', () => {
    assert.equal(signatureSkipReason({ textBody: '   ' }, SIG), 'no-body');
    assert.equal(signatureSkipReason({ textBody: '\u200b' }, SIG), 'no-body');
    assert.deepEqual(applySignature({ textBody: '   ' }, SIG), { textBody: '   ' });
    assert.deepEqual(applySignature({ textBody: '\u200b' }, SIG), { textBody: '\u200b' });
  });

  // The same rule on the other branch: a blank text alternative beside a real html body ships
  // no text part (buildBodyParts drops it), so writing the block into it would make the
  // message's whole plain-text alternative a bare sign-off under an html body carrying the
  // actual message. Left alone, the alternative is derived from the signed html instead.
  it('does not turn a blank text alternative into a signature-only text part', () => {
    const out = applySignature({ textBody: '  ', htmlBody: '<p>Hi</p>' }, SIG);
    assert.equal(out.textBody, '  ');
    assert.ok(hasSignatureMarker(out.htmlBody));
    assert.equal(signatureSkipReason({ textBody: '  ', htmlBody: '<p>Hi</p>' }, SIG), undefined);
  });

  it('reports the reason an append landed nowhere', () => {
    assert.equal(signatureSkipReason({ textBody: 'Hi' }, undefined), 'no-signature');
    assert.equal(signatureSkipReason({}, SIG), 'no-body');
    assert.equal(signatureSkipReason({ htmlBody: '' }, SIG), 'no-body');
    assert.equal(
      signatureSkipReason({ htmlBody: '<p>Hi</p><div class="fm-mcp-signature">x</div>' }, SIG),
      'already-signed',
    );
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
    assert.ok(hasSignatureMarker('<div class="fm-mcp-signature trailing">x</div>'));
    assert.ok(hasSignatureMarker('<div class = "fm-mcp-signature">x</div>'));
  });

  // The class must be a WHOLE token of the class list. A longer name that merely starts with
  // ours is somebody else's class, and matching it is not the harmless over-suppression it
  // was once recorded as: the same predicate reads the STORED body in updateDraft, where a
  // false positive makes an edit that said nothing about signatures append one nobody asked
  // for and announce that the draft already carried one.
  it('does not match a longer class name that merely starts with the marker', () => {
    assert.ok(!hasSignatureMarker('<div class="fm-mcp-signature-ish">x</div>'));
    assert.ok(!hasSignatureMarker('<div class="a fm-mcp-signature-ish b">x</div>'));
    assert.ok(!hasSignatureMarker('<div class="notfm-mcp-signature">x</div>'));
    assert.ok(!hasSignatureMarker('<div>Kind regards</div>'));
    assert.ok(!hasSignatureMarker('<blockquote class="fm-mcp-signature">x</blockquote>'));
  });

  // The other half of the same predicate: the ATTRIBUTE NAME needs a real boundary too, and
  // `\b` is not one — `-` is a word boundary, so `data-class` ended in `class` and matched.
  it('does not match a longer attribute name that merely ends in class', () => {
    assert.ok(!hasSignatureMarker('<div data-class="fm-mcp-signature">x</div>'));
    assert.ok(!hasSignatureMarker('<div myclass="fm-mcp-signature">x</div>'));
    assert.ok(hasSignatureMarker('<div data-x="1" class="fm-mcp-signature">x</div>'));
  });

  // A quoted attribute VALUE is not a place attributes live. Accepting a quote character as
  // the boundary in front of `class` let any string containing `class=fm-mcp-signature` claim
  // the marker — and this predicate reads agent-authored html straight off the stored draft,
  // where the sanitizer never ran.
  it('does not match class= written inside another attribute value', () => {
    assert.ok(!hasSignatureMarker('<div data-x=" class=fm-mcp-signature ">x</div>'));
    assert.ok(!hasSignatureMarker(`<div title="a class='fm-mcp-signature' b">x</div>`));
    assert.ok(!hasSignatureMarker('<div alt="class = fm-mcp-signature">x</div>'));
  });

  // The mirror failure, which costs a signature rather than inventing one: a `>` inside an
  // earlier attribute's value ended the tag as far as a `[^>]*` scan was concerned, so a
  // genuine marker behind one went unseen and the preserve path silently stopped firing.
  it('finds the marker behind an attribute value containing a closing bracket', () => {
    assert.ok(hasSignatureMarker('<div title="a > b" class="fm-mcp-signature">x</div>'));
    assert.ok(hasSignatureMarker(`<div title='x > y' class=fm-mcp-signature>x</div>`));
    assert.ok(hasSignatureMarker('<p>before</p><div title="1>2"><div class="fm-mcp-signature">x</div></div>'));
  });

  it('survives tags it cannot parse without claiming a marker', () => {
    assert.ok(!hasSignatureMarker('<div title="unterminated class="fm-mcp-signature">x'));
    assert.ok(!hasSignatureMarker('<div'));
    assert.ok(!hasSignatureMarker('<div class>x</div>'));
    assert.ok(hasSignatureMarker('<div hidden class="fm-mcp-signature">x</div>'));
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

  // Load-bearing and, until this test, unasserted: the marker means "this draft's OWN
  // signature", and edit_draft's preservation reads it off the stored body. If a signature
  // quoted back inside a reply could false-trip it, an edit of that reply would treat
  // somebody else's sign-off as the draft's own. What makes it hold lives in ANOTHER module
  // — QUOTE_ALLOWED_ATTRIBUTES in src/inline-images.ts has no global '*' entry, so class= is
  // stripped from every quoted element — so this drives the real quote path rather than
  // restating the rule.
  it('cannot be false-tripped by a signature quoted back from the original', () => {
    const signedOriginal = {
      ...ORIGINAL,
      bodyValues: {
        t: { value: 'Original text.' },
        h: { value: `<p>Original html.</p>${signatureHtmlBlock(SIG)}` },
      },
    };
    assert.ok(hasSignatureMarker(signedOriginal.bodyValues.h.value)); // the source really is signed
    const out = buildReplyBodies({
      original: signedOriginal, htmlBody: '<p>Thanks.</p>', quoteOriginal: true,
    });
    assert.match(out.htmlBody!, /Kind regards,/); // the quoted content survives…
    assert.ok(!hasSignatureMarker(out.htmlBody), out.htmlBody); // …but not as a marker
  });

  it('still detects the draft\'s own signature above a quote carrying somebody else\'s', () => {
    const signedOriginal = {
      ...ORIGINAL,
      bodyValues: {
        t: { value: 'Original text.' },
        h: { value: `<p>Original html.</p>${signatureHtmlBlock(SIG)}` },
      },
    };
    const out = buildReplyBodies({
      original: signedOriginal, htmlBody: '<p>Thanks.</p>', quoteOriginal: true, signature: SIG,
    });
    assert.ok(hasSignatureMarker(out.htmlBody));
    assert.ok(out.htmlBody!.indexOf('fm-mcp-signature') < out.htmlBody!.indexOf('<blockquote'));
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

  // The note-less arm on a TEXT-ONLY original takes the derived text form, so an images-only
  // signature made the forward's entire visible content a "[image]" placeholder describing
  // an image no part of that forward carries. It now ships nothing and reports why.
  it('does not make a bare image placeholder the whole content of a text-only forward', () => {
    const textOnly = { ...ORIGINAL, htmlBody: [], bodyValues: { t: { value: 'Original text.' } } };
    const out = buildForwardBodies({ original: textOnly, signature: { html: '<img src="cid:logo">' } });
    assert.ok(out.textBody!.startsWith('----- Original message -----'), out.textBody);
    assert.equal(out.signatureSkip, 'no-text-form');
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

  // appendSignature is an input the caller cannot verify without re-reading the draft, so an
  // append that lands nowhere is reported rather than passed over in silence — and the note
  // names WHICH reason, because the fix differs per reason.
  it('appends nothing for an identity with no signature, and says so', async () => {
    const { client, calls } = composeClient([UNSIGNED_IDENTITY]);
    const r = await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: true }, client, undefined, false);
    assert.equal(calls.draft.htmlBody, '<p>Hi</p>');
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-signature', 'me@example.com')]);
  });

  it('says so when the call writes no body for the signature to sign', async () => {
    const { client, calls } = composeClient();
    const r = await composeDraft({ to: ['bob@example.com'], subject: 'Hi', appendSignature: true }, client, undefined, false);
    assert.equal(calls.draft.htmlBody, undefined);
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-body', 'me@example.com')]);
  });

  // Html the caller pasted in (from anywhere) that happens to carry the marker suppresses
  // the append. Suppression is right — one signature, not two — but silence is not.
  it('says so when the supplied html already carries a signature block', async () => {
    const { client } = composeClient();
    const r = await composeDraft(
      { to: ['bob@example.com'], htmlBody: '<p>Hi</p><div class="fm-mcp-signature">Someone</div>', appendSignature: true },
      client, undefined, false,
    );
    assert.deepEqual(r.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  it('says nothing when the append succeeded', async () => {
    const { client } = composeClient();
    const r = await composeDraft({ to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: true }, client, undefined, false);
    assert.equal(r.notes, undefined);
  });

  // A lookup that failed after the upload would leave freshly written blobs in the account
  // with no draft referencing them — the same ordering rule the embedded-image plan has.
  it('resolves the identity BEFORE any attachment is uploaded', async () => {
    const order: string[] = [];
    const client: ComposeClient = {
      getIdentities: async () => { order.push('identities'); return [SIGNED_IDENTITY]; },
      uploadAttachments: async () => { order.push('upload'); return []; },
      createDraft: async () => 'draft-9',
      getEmailById: async (id) => ({ id, attachments: [] }),
    };
    await composeDraft(
      { to: ['bob@example.com'], htmlBody: '<p>Hi</p>', appendSignature: true, attachments: [{ blobId: 'b', name: 'f.txt' }] },
      client, undefined, true,
    );
    assert.deepEqual(order, ['identities', 'upload']);
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

  // A body-less reply is left unsigned, and DELIBERATELY so — unlike a forward, where a
  // note-less "FYI, see below" is the normal shape and the builder makes the signature the
  // whole content above the block, a reply whose entire content is a sign-off over a quote
  // is not a message. The asymmetry is a behaviour decision, not an oversight; what this
  // pins is that the caller is told rather than left to discover it by re-reading the draft.
  it('reports a body-less reply that asked for a signature', async () => {
    const { client, calls } = replyClient();
    const r = await composeReply({ originalEmailId: 'orig-1', appendSignature: true }, client, undefined, false);
    assert.ok(!hasSignatureMarker(calls.draft.htmlBody));
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-body', 'me@example.com')]);
  });

  // htmlShips is false for a blank body (buildBodyParts drops it), textBody is undefined, so
  // nothing is signed — while the quote path still builds an html body around it.
  it('reports a reply whose only body is a blank htmlBody', async () => {
    const { client } = replyClient();
    const r = await composeReply(
      { originalEmailId: 'orig-1', htmlBody: '', appendSignature: true }, client, undefined, false,
    );
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-body', 'me@example.com')]);
  });

  it('reports a reply whose identity has no signature configured', async () => {
    const calls: any = {};
    const client: ReplyClient = {
      getEmailById: async (id) => (id === 'draft-9' ? { id, attachments: [] } : ORIGINAL),
      getIdentities: async () => [UNSIGNED_IDENTITY],
      uploadAttachments: async () => [],
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    };
    const r = await composeReply(
      { originalEmailId: 'orig-1', htmlBody: '<p>Thanks.</p>', appendSignature: true }, client, undefined, false,
    );
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-signature', 'me@example.com')]);
  });

  it('reports a forward whose identity has no signature configured', async () => {
    const client: ForwardClient = {
      getEmailById: async (id) => (id === 'draft-7' ? { id, attachments: [] } : { ...ORIGINAL, blobId: 'blob-eml' }),
      getIdentities: async () => [UNSIGNED_IDENTITY],
      uploadAttachments: async () => [],
      createDraft: async () => 'draft-7',
    };
    const r = await composeForward(
      { originalEmailId: 'orig-1', to: ['bob@example.com'], htmlBody: '<p>FYI.</p>', appendSignature: true },
      client, undefined, false,
    );
    assert.deepEqual(r.notes, [noteSignatureNotAppended('no-signature', 'me@example.com')]);
  });

  // The forward's own no-note shape signs, so it must NOT be reported as an empty append.
  it('says nothing about a note-less forward that the signature became the body of', async () => {
    const { client } = forwardClient();
    const r = await composeForward(
      { originalEmailId: 'orig-1', to: ['bob@example.com'], appendSignature: true }, client, undefined, false,
    );
    assert.equal(r.notes, undefined);
  });

  it('says nothing about an ordinary unsigned forward', async () => {
    const { client } = forwardClient();
    const r = await composeForward(
      { originalEmailId: 'orig-1', to: ['bob@example.com'], htmlBody: '<p>FYI.</p>' }, client, undefined, false,
    );
    assert.equal(r.notes, undefined);
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

  it('leaves the caller\'s own updates object untouched when it appends the signature', async () => {
    // The signed body is the server's output, and an argument object belongs to the caller: a
    // call that wrote its output back over the html it was handed would silently change a value
    // the caller can still read, and would hand any second call made with the same object a body
    // that already ends in the sign-off.
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const args = { htmlBody: '<p>Rewritten.</p>' };
    const before = structuredClone(args);
    await client.updateDraft('draft-1', args);
    // The append has to have actually run, or there is no signed body to write back and the
    // assertion below would hold for the wrong reason.
    assert.ok(hasSignatureMarker(seen.created.bodyValues.html.value), seen.created.bodyValues.html.value);
    assert.deepEqual(args, before);
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

  // The preserve branch is armed by the STORED draft, so the omitted flag meant "keep it".
  // When the keep cannot be honoured the draft loses a block it was carrying, and a silent
  // loss is exactly what the branch exists to prevent.
  it('says so when the draft loses its signature because the identity no longer has one', async () => {
    const client = makeClient([UNSIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.equal(seen.created.bodyValues.html.value, '<p>Rewritten.</p>');
    assert.deepEqual(result.notes, [noteSignatureUnavailableOnEdit('me@example.com')]);
  });

  it('says so when appendSignature:true has no signature to append', async () => {
    const client = makeClient([UNSIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft({
      bodyValues: { t: { value: 'Hello.' }, h: { value: '<p>Hello.</p>' } },
    }));

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>', appendSignature: true });
    assert.equal(seen.created.bodyValues.html.value, '<p>Rewritten.</p>');
    assert.deepEqual(result.notes, [noteSignatureNotAppended('no-signature', 'me@example.com')]);
  });

  // Nothing is lost when no body is written: both bodies (and the signature in them) are
  // carried through verbatim, so the loss note must not fire.
  it('says nothing about a metadata-only edit on an identity with no signature', async () => {
    const client = makeClient([UNSIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(seen.created.bodyValues.html.value, SIGNED_BODY);
    assert.equal(result.notes, undefined);
  });

  // The draft's stored `from` matches no verified identity — an address removed from the
  // account, or a draft made elsewhere. `selectedIdentity` falls back to the account default
  // for the recreate, but the address WRITTEN into `from` stays the stored one, so signing
  // from that fallback would put one identity's sign-off under another's address.
  it('does not sign with the default identity when the written from matches none', async () => {
    const client = makeClient([SIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft({ from: [{ email: 'gone@elsewhere.example' }] }));

    const result = await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.equal(seen.created.from[0].email, 'gone@elsewhere.example');
    assert.ok(!hasSignatureMarker(seen.created.bodyValues.html.value));
    assert.deepEqual(result.notes, [noteSignatureUnavailableOnEdit('gone@elsewhere.example')]);
  });

  it('signs with the identity named by an explicit from on the edit', async () => {
    const other = { id: 'id-9', email: 'other@example.com', mayDelete: true, htmlSignature: '<div>From Other</div>' };
    const client = makeClient([SIGNED_IDENTITY, other]);
    const seen = mockUpdate(client, signedDraft());

    await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>', from: 'other@example.com' });
    assert.equal(seen.created.from[0].email, 'other@example.com');
    assert.match(seen.created.bodyValues.html.value, /From Other/);
  });

  // The plain-text residual's own recovery loop. A text-only draft carries no marker, so an
  // edit that rewrites its body must be given appendSignature:true to keep the sign-off —
  // and the natural agent loop reads the draft back, edits the words, and sends the WHOLE
  // text (signature included) with the flag still set. That must not stack.
  it('does not stack the signature on a re-signed plain-text draft', async () => {
    const client = makeClient();
    const textOnly = {
      id: 'draft-1',
      subject: 'Old Subject',
      from: [{ email: 'me@example.com' }],
      to: [{ email: 'bob@example.com' }],
      cc: [], bcc: [],
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [],
      bodyValues: { t: { value: 'Hi Bob\n\nRegards,\nTest User' } },
      mailboxIds: { 'mb-drafts': true },
      keywords: { $draft: true },
    };
    const seen = mockUpdate(client, textOnly);

    await client.updateDraft('draft-1', {
      textBody: 'Hi Bob, one more thing\n\nRegards,\nTest User',
      appendSignature: true,
    });
    const stored = seen.created.bodyValues.text.value;
    assert.equal(stored.match(/Regards,/g)!.length, 1, stored);
    assert.equal(stored, 'Hi Bob, one more thing\n\nRegards,\nTest User');
  });

  // The same recovery loop on a REPLY draft, which is the shape it is actually recommended
  // for: a plain-text reply keeps the quoted original in its body, so the text handed back
  // carries the sign-off in the MIDDLE and the quote at the end. The fixture above has no
  // quote in it and so could not exhibit this; run through the real updateDraft it stacked a
  // second sign-off below the quote every time the loop went round.
  it('does not stack the signature on a re-signed plain-text reply draft that carries a quote', async () => {
    const client = makeClient();
    const QUOTE = 'On 1 Jan 2026, at 03:04, Alice <alice@example.com> wrote:\n> Original text.';
    const textOnly = {
      id: 'draft-1',
      subject: 'Re: Project update',
      from: [{ email: 'me@example.com' }],
      to: [{ email: 'alice@example.com' }],
      cc: [], bcc: [],
      inReplyTo: ['<orig@example.com>'],
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [],
      bodyValues: { t: { value: `Thanks.\n\nRegards,\nTest User\n\n${QUOTE}` } },
      mailboxIds: { 'mb-drafts': true },
      keywords: { $draft: true },
    };
    const seen = mockUpdate(client, textOnly);

    const edited = `Thanks, and one more thing.\n\nRegards,\nTest User\n\n${QUOTE}`;
    const result = await client.updateDraft('draft-1', {
      textBody: edited,
      noQuote: true,
      appendSignature: true,
    });
    const stored = seen.created.bodyValues.text.value;
    assert.equal(stored.match(/Regards,/g)!.length, 1, stored);
    assert.match(stored, /^Thanks, and one more thing\.\n\nRegards,\nTest User\n/, stored);
    // The quote came back in the body, so it must be stored exactly as handed over — once.
    assert.equal(stored, edited, stored);
    // The flag was passed and nothing was appended, so it is reported rather than silent.
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  /** A plain-text reply draft storing `value`, replying to ORIGINAL. */
  const textReplyDraft = (value: string) => ({
    id: 'draft-1',
    subject: 'Re: Project update',
    from: [{ email: 'me@example.com' }],
    to: [{ email: 'alice@example.com' }],
    cc: [], bcc: [],
    inReplyTo: ['<orig@example.com>'],
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [],
    bodyValues: { t: { value } },
    mailboxIds: { 'mb-drafts': true },
    keywords: { $draft: true },
  });

  /**
   * The recovery loop: hand the draft's own text back with the flag still set.
   *
   * noQuote, not originalEmailId: the text handed back already carries the quoted original, and
   * originalEmailId REBUILDS the quote underneath the body supplied, so it would store the quote
   * twice (and is refused for that reason, #145). noQuote is what the loop is documented to
   * pass, and it stores the body as written — which is the premise every assertion here makes.
   */
  async function reSign(
    body: string,
    edit: (s: string) => string = (s) => s,
    identities?: any[],
  ) {
    const client = makeClient(identities);
    const seen = mockUpdate(client, textReplyDraft(body));
    const result = await client.updateDraft('draft-1', {
      textBody: edit(body), noQuote: true, appendSignature: true,
    });
    return { stored: seen.created.bodyValues.text.value as string, result };
  }

  // Sign-offs the DRAFT carries, counted at line start and unquoted, so a "Regards," inside a
  // quote is not mistaken for one of the draft's own.
  const ownSignOffs = (s: string) => (s.match(/^Regards,$/gm) ?? []).length;
  // The draft's OWN attribution line — "… wrote:" at line start and NOT inside a quote, so a
  // nested attribution stays the quoted message's rather than counting as the draft's.
  const ownAttributions = (s: string) => (s.match(/^(?!>).*\bwrote:[ \t]*$/gm) ?? []).length;
  // The quoted original's own content, at whatever nesting depth it sits. Counting the QUOTE, not
  // just the sign-off, is what catches a rebuild landing a second copy of it underneath the body:
  // these tests all passed while storing it twice, because a sign-off count cannot see that.
  const quotedOriginals = (s: string) => (s.match(/^>+ Original text\.$/gm) ?? []).length;

  // The identity whose own sign-off contains a quotation, driven through the real edit on the
  // loop the docs prescribe. This is the shape that duplicated silently while every pure-helper
  // test in the file passed, because no fixture here had a '> ' line inside its signature.
  it('does not stack a sign-off that itself contains a quoted line', async () => {
    const block = signatureTextBlock(QUOTING_SIG, false)!;
    const body = `Thanks.\n\n${block}\n\nOn 1 Jan 2026, Alice <alice@example.com> wrote:\n> Original text.`;
    const { stored, result } = await reSign(
      body,
      (s) => s.replace('Thanks.', 'Thanks again.'),
      [QUOTING_IDENTITY],
    );
    assert.equal(ownSignOffs(stored), 1, stored);
    assert.equal(stored.match(/Per aspera ad astra/g)!.length, 1, stored);
    assert.equal(ownAttributions(stored), 1, stored);
    assert.equal(quotedOriginals(stored), 1, stored);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // …and the html-derived form of the same signature, which is what a draft converted out of
  // html hands back. That form's quotation comes from html-to-text's <blockquote> rendering.
  it('does not stack the derived form of a quoting sign-off on a converted draft', async () => {
    const client = makeClient([QUOTING_IDENTITY]);
    const derived = signatureTextBlock(QUOTING_SIG, true)!;
    const seen = mockUpdate(client, signedDraft({
      bodyValues: {
        t: { value: `Hello.\n\n${derived}` },
        h: { value: `<p>Hello.</p><div class="fm-mcp-signature">${QUOTING_IDENTITY.htmlSignature}</div>` },
      },
    }));

    const result = await client.updateDraft('draft-1', {
      textBody: `Hello, one more thing.\n\n${derived}`,
      clearFields: ['htmlBody'],
      appendSignature: true,
    });
    const stored = seen.created.bodyValues.text.value;
    assert.equal(stored, `Hello, one more thing.\n\n${derived}`, stored);
    assert.equal(stored.match(/Per aspera ad astra/g)!.length, 1, stored);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // The property the removed line filter was supposed to protect, re-checked without it: a
  // forward separator INSIDE a reply quote must not cut the body short, or a sign-off below
  // the quote would fall outside the haystack and be appended a second time. It cannot,
  // because FORWARD_SEPARATOR_LINE anchors at the dashes and the '> ' defeats that anchor.
  it('does not let a quoted forward separator cut the body short', async () => {
    const body = [
      'Thanks.',
      '',
      'On 1 Jan 2026, Alice <alice@example.com> wrote:',
      '> ----- Original message -----',
      '> From: Bob',
      '>',
      '> Original text.',
      '',
      'Regards,',
      'Test User',
    ].join('\n');
    const { stored, result } = await reSign(body, (s) => s.replace('Thanks.', 'Thanks again.'));
    assert.equal(ownSignOffs(stored), 1, stored);
    assert.equal(ownAttributions(stored), 1, stored);
    assert.equal(quotedOriginals(stored), 1, stored);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // The attribution shapes other clients write. Nothing in the guard parses an attribution,
  // so a hard-wrapped Gmail attribution is an ordinary body. Driven through the real edit
  // because that is where every earlier version of this stacked: the pure builder is handed a
  // body with no quote in it.
  it('does not stack when the Gmail attribution above the quote is hard-wrapped', async () => {
    const body = [
      'Thanks.',
      '',
      'Regards,',
      'Test User',
      '',
      'On Mon, Jan 1, 2026 at 9:00 AM Alice Example <alice@example.com>',
      'wrote:',
      '',
      '> Original text.',
    ].join('\n');
    const { stored, result } = await reSign(body, (s) => s.replace('Thanks.', 'Thanks again.'));
    assert.equal(ownSignOffs(stored), 1, stored);
    assert.equal(ownAttributions(stored), 1, stored);
    assert.equal(quotedOriginals(stored), 1, stored);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // Nested quoting is the same test — '>' is the first non-whitespace character either way.
  it('does not stack on a re-signed reply draft quoting a quote', async () => {
    const body = [
      'Thanks.',
      '',
      'Regards,',
      'Test User',
      '',
      'On 1 Jan 2026, Alice <alice@example.com> wrote:',
      '> On 31 Dec 2025, Test User <me@example.com> wrote:',
      '>> Original text.',
      '>>',
      '>> Regards,',
      '>> Test User',
    ].join('\n');
    const { stored, result } = await reSign(body, (s) => s.replace('Thanks.', 'Thanks again.'));
    assert.equal(ownSignOffs(stored), 1, stored);
    assert.equal(ownAttributions(stored), 1, stored);
    assert.equal(quotedOriginals(stored), 1, stored);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // The mirror: an UNSIGNED note above a quote whose quoted text carries this identity's own
  // sign-off (the ordinary thread where an earlier message of ours was signed). The quoted
  // lines are removed, so the note reads as unsigned and IS signed — one sign-off of the
  // draft's own. Reading through the quote instead would refuse to sign every reply in a
  // thread the sender had ever signed.
  it('signs a reply whose quote carries the sign-off but whose note does not', async () => {
    const body = [
      'Thanks.',
      '',
      'On 1 Jan 2026, Alice <alice@example.com> wrote:',
      '> Original text.',
      '>',
      '> Regards,',
      '> Test User',
    ].join('\n');
    const { stored, result } = await reSign(body);
    assert.equal(ownSignOffs(stored), 1, stored);
    assert.equal(ownAttributions(stored), 1, stored);
    assert.equal(quotedOriginals(stored), 1, stored);
    assert.equal(result.notes, undefined);
  });

  // A forward whose forwarded message carries this identity's sign-off, with the caller's own
  // note unsigned: the separator removes the block, so the note gets signed rather than the
  // forwarded sign-off being read as the note's.
  it('signs the note on a forward whose forwarded block carries the same sign-off', async () => {
    const body = [
      'Please see below.',
      '',
      '----- Original message -----',
      'From: Test User <me@example.com>',
      'Subject: Project update',
      '',
      'Here it is.',
      '',
      'Regards,',
      'Test User',
    ].join('\n');
    const client = makeClient();
    const seen = mockUpdate(client, textReplyDraft(body));
    const result = await client.updateDraft('draft-1', {
      textBody: body, noQuote: true, appendSignature: true,
    });
    const stored = seen.created.bodyValues.text.value;
    assert.equal(ownSignOffs(stored), 2, stored); // the forwarded one, and the note's new one
    assert.match(stored, /Test User$/, stored);
    assert.equal(result.notes, undefined);
  });

  // hasSignatureMarker reads the STORED body here, so a false positive is not the harmless
  // over-suppression it is on the compose side: it makes an edit that said nothing about
  // signatures append one nobody asked for, and announce that the draft already carried one.
  it('does not treat a lookalike class on the stored body as a signature', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft({
      bodyValues: {
        t: { value: 'Hello.' },
        h: { value: '<p>Hello.</p><div class="fm-mcp-signature-ish">Not ours</div>' },
      },
    }));

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

  // The cross-form duplication, driven through the real edit rather than the pure builder.
  // clearFields:['htmlBody'] is the permitted html→text conversion, and the natural loop
  // hands back the text the draft was storing — which is the DERIVED form, while a text-only
  // call writes the CONFIGURED one. Matching only the outgoing form stacked a second sign-off
  // here, on the explicit recovery the schema recommends for exactly this draft shape.
  it('does not stack a second sign-off when a signed html draft is converted to plain text', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', {
      textBody: `Hello, one more thing.\n\n${DERIVED_TEXT}`,
      clearFields: ['htmlBody'],
      appendSignature: true,
    });
    const stored = seen.created.bodyValues.text.value;
    assert.equal(stored, `Hello, one more thing.\n\n${DERIVED_TEXT}`, stored);
    assert.equal(stored.match(/regards,/gi)!.length, 1, stored);
    // The flag was passed and nothing was appended, so it is reported — the body already
    // carried the sign-off.
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  // Same conversion with the flag OMITTED — the preserve path, where the caller asked for
  // nothing signature-related. This used to duplicate AND announce a re-append that had not
  // happened, which is worse than the duplication: the note asserted the body it was handed
  // carried no signature when it plainly did.
  it('neither stacks nor announces a re-append on an omitted-flag conversion', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', {
      textBody: `Hello, one more thing.\n\n${DERIVED_TEXT}`,
      clearFields: ['htmlBody'],
    });
    assert.equal(seen.created.bodyValues.text.value, `Hello, one more thing.\n\n${DERIVED_TEXT}`);
    assert.equal(result.notes, undefined);
  });

  // An images-only html signature is not signature-LESS, so the old gate (keyed on the
  // identity having no signature at all) armed nothing for it — and the conversion below
  // dropped a sign-off the draft was carrying in silence, including on the explicit
  // appendSignature:true recovery.
  const IMAGE_ONLY_IDENTITY = {
    id: 'id-1', name: 'Test User', email: 'me@example.com', mayDelete: false,
    htmlSignature: '<img src="https://example.com/logo.png">',
  };
  const imageSignedDraft = () => signedDraft({
    bodyValues: {
      t: { value: 'Hello.' },
      h: { value: '<p>Hello.</p><div class="fm-mcp-signature"><img src="https://example.com/logo.png"></div>' },
    },
  });

  it('says so when a requested append has no plain-text form to write', async () => {
    const client = makeClient([IMAGE_ONLY_IDENTITY]);
    const seen = mockUpdate(client, imageSignedDraft());

    const result = await client.updateDraft('draft-1', {
      textBody: 'Rewritten.', clearFields: ['htmlBody'], appendSignature: true,
    });
    assert.equal(seen.created.bodyValues.text.value, 'Rewritten.');
    assert.deepEqual(result.notes, [noteSignatureNotAppended('no-text-form', 'me@example.com')]);
  });

  it('says so when a PRESERVED signature has no plain-text form to carry over', async () => {
    const client = makeClient([IMAGE_ONLY_IDENTITY]);
    const seen = mockUpdate(client, imageSignedDraft());

    const result = await client.updateDraft('draft-1', {
      textBody: 'Rewritten.', clearFields: ['htmlBody'],
    });
    assert.equal(seen.created.bodyValues.text.value, 'Rewritten.');
    // The loss wording, not the "you asked" wording: nothing in this call asked.
    assert.deepEqual(result.notes, [noteSignatureUnavailableOnEdit('me@example.com', 'no-text-form')]);
  });

  // The other three compose surfaces report a requested append that landed nowhere; this one
  // used to stay silent for two of the reasons. The flag is an input the caller cannot verify
  // without re-reading the draft, so it reports on every surface or on none.
  it('says so when appendSignature:true lands on a body that already ends with one', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft());

    const result = await client.updateDraft('draft-1', { htmlBody: SIGNED_BODY, appendSignature: true });
    assert.equal(seen.created.bodyValues.html.value.match(/fm-mcp-signature/g)!.length, 1);
    assert.deepEqual(result.notes, [noteSignatureNotAppended('already-signed', 'me@example.com')]);
  });

  it('says so when appendSignature:true is passed by an edit that writes no body', async () => {
    const client = makeClient();
    const seen = mockUpdate(client, signedDraft({
      bodyValues: { t: { value: 'Hello.' }, h: { value: '<p>Hello.</p>' } },
    }));

    const result = await client.updateDraft('draft-1', { subject: 'New Subject', appendSignature: true });
    assert.equal(seen.created.bodyValues.html.value, '<p>Hello.</p>'); // body-invariant, as documented
    assert.deepEqual(result.notes, [noteSignatureNotAppended('no-body', 'me@example.com')]);
  });

  // The address hoist that keeps the sign-off with the sender has to carry the DISPLAY NAME
  // too. Signing was already resolved against the written address; the name was still read
  // off the account default, so a draft whose stored from matches no verified identity came
  // back with the default identity's name in front of a foreign address.
  it('does not put the default identity\'s display name against a foreign address', async () => {
    const client = makeClient([SIGNED_IDENTITY]);
    const seen = mockUpdate(client, signedDraft({ from: [{ name: 'Old Name', email: 'gone@elsewhere.example' }] }));

    await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>' });
    assert.deepEqual(seen.created.from, [{ name: 'Old Name', email: 'gone@elsewhere.example' }]);
  });

  it('takes the display name from the identity an explicit from names', async () => {
    const other = { id: 'id-9', name: 'Other Person', email: 'other@example.com', mayDelete: true };
    const client = makeClient([SIGNED_IDENTITY, other]);
    const seen = mockUpdate(client, signedDraft());

    await client.updateDraft('draft-1', { htmlBody: '<p>Rewritten.</p>', from: 'other@example.com' });
    assert.deepEqual(seen.created.from, [{ name: 'Other Person', email: 'other@example.com' }]);
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
