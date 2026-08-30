import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildForwardBlocks, buildQuoteBlocks } from './reply-quote.js';

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

// A raw-JMAP-shaped original carrying whichever formats the test asks for.
function makeOriginal(opts: {
  text?: string; html?: string; name?: string; email?: string;
  sentAt?: string; receivedAt?: string; aliasType?: string;
}) {
  const { text, html, name, email = 'jon@example.com', sentAt, receivedAt, aliasType } = opts;
  const bodyValues: Record<string, any> = {};
  const textBody: any[] = [];
  const htmlBody: any[] = [];
  if (text !== undefined) { bodyValues.t = { value: text }; textBody.push({ partId: 't', type: aliasType ?? 'text/plain' }); }
  if (html !== undefined) { bodyValues.h = { value: html }; htmlBody.push({ partId: 'h', type: 'text/html' }); }
  return {
    from: [{ ...(name !== undefined && { name }), email }],
    ...(sentAt && { sentAt }), ...(receivedAt && { receivedAt }),
    textBody, htmlBody, bodyValues,
  };
}

const TZ = 'Australia/Sydney';

// A raw-JMAP-shaped original with both bodies, Cc, and a sent date.
function fwdOriginal(over: any = {}) {
  return {
    from: [{ name: 'Ada Lovelace', email: 'ada@example.com' }],
    to: [{ name: 'Bob', email: 'bob@example.com' }],
    cc: [{ email: 'carol@example.com' }],
    subject: 'Original subject',
    sentAt: '2026-07-01T09:14:00-04:00',
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>original html</p>' } },
    ...over,
  };
}
const attachmentOnlyOriginal = () => fwdOriginal({ textBody: undefined, htmlBody: undefined, bodyValues: {} });

// ---------------------------------------------------------------------------
// Where a block begins
// ---------------------------------------------------------------------------
//
// A block is handed to whoever placed the token, and it has to arrive with no leading
// separator of its own: the caller decides what sits between their body and the history,
// and a block that opened with its own blank line or <br> would put spacing there that the
// caller neither wrote nor can remove.

describe('buildQuoteBlocks — a block starts at its attribution line', () => {
  it('opens at the attribution, carrying no separator of its own', () => {
    const blocks = buildQuoteBlocks({ original: fwdOriginal(), htmlShips: true });
    assert.ok(blocks.textBlock!.startsWith('On '), blocks.textBlock);
    assert.ok(blocks.htmlBlock!.startsWith('<div>On '), blocks.htmlBlock);
  });

  it('an original quotable in no format yields no block at all (no orphan attribution)', () => {
    const blocks = buildQuoteBlocks({ original: attachmentOnlyOriginal(), htmlShips: true });
    assert.equal(blocks.textBlock, undefined);
    assert.equal(blocks.htmlBlock, undefined);
    assert.equal(blocks.images.htmlQuoteShips, false);
  });
});

describe('buildForwardBlocks — a block starts at its header line', () => {
  it('does NOT open with a <br>: that was a separator sitting inside the block\'s own div', () => {
    const { htmlBlock } = buildForwardBlocks({ original: fwdOriginal(), htmlShips: true });
    assert.doesNotMatch(htmlBlock, /^<div><br>/);
    assert.ok(htmlBlock.startsWith('<div>----- Original message -----<br>'), htmlBlock);
  });

  it('the text block opens at the marker line, carrying no blank lines of its own', () => {
    const { textBlock } = buildForwardBlocks({ original: fwdOriginal(), htmlShips: false });
    assert.ok(textBlock.startsWith('----- Original message -----'), textBlock);
  });
});

// ---------------------------------------------------------------------------
// Block CONTENT, pinned on the block builders themselves
// ---------------------------------------------------------------------------
//
// What a quoted or forwarded block SAYS — the attribution line, the "> " prefixing, the
// header block, the html shapes it wraps the original in — is produced by buildQuoteBlocks /
// buildForwardBlocks and is a property of the block alone. Whatever joins the block to a body
// cannot change any of it, so these read the builder's return value directly.
//
// Deliberately NOT here, even though the same builders produce it: anything whose subject is
// what the STORED message contains. The quote sanitiser and the forward header's escaping of
// hostile fields both answer "can this construct reach a recipient", and asserting them over
// a block would answer a strictly narrower question in the same words — the block could be
// clean and the join still ship the construct. Those live over the stored parts, in
// draft-email-handler.test.ts.

// Late import, beside the suites that use it (same convention as the suites above).
import { signatureBlock, signatureHtmlBlock, signatureTextBlock } from './reply-quote.js';

// A signature as signatureOf hands it over: either form may be absent.
const HTML_ONLY_SIG = { html: '<div>Kind regards,</div><div>Test User</div>' };
// A quote-of-the-day sign-off. html-to-text renders its <blockquote> as '> '-prefixed lines,
// which is the only derived signature block in this repo that looks like reply quoting.
const QUOTING_SIG = {
  text: 'Regards,\nTest User\n> Per aspera ad astra',
  html: '<div>Kind regards,</div><blockquote>Per aspera ad astra</blockquote><div>Test User</div>',
};

describe('buildQuoteBlocks — the attribution line', () => {
  const textBlockFor = (opts: Parameters<typeof makeOriginal>[0]) =>
    buildQuoteBlocks({ original: makeOriginal(opts), htmlShips: false, timezone: TZ }).textBlock!;

  it('renders the exact captured Fastmail attribution (local time, ASCII-spaced)', () => {
    const block = textBlockFor({ text: 'orig', name: 'Jonathan Godley', sentAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On Mon, Jun 15, 2026, at 1:29 PM, Jonathan Godley wrote:\n/);
  });

  it('uses sentAt over receivedAt', () => {
    const block = textBlockFor({
      text: 'orig', name: 'Jon', sentAt: '2026-06-15T03:29:02Z', receivedAt: '2026-06-15T09:00:00Z',
    });
    assert.match(block, /at 1:29 PM, Jon wrote:/); // 1:29 PM = sentAt, not the 7 PM receivedAt
  });

  it('falls back to receivedAt when sentAt is absent', () => {
    const block = textBlockFor({ text: 'orig', name: 'Jon', receivedAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On Mon, Jun 15, 2026, at 1:29 PM, Jon wrote:\n/);
  });

  it('omits the date entirely (never "Invalid Date") when no timestamp is present', () => {
    const block = textBlockFor({ text: 'orig', name: 'Jon' }); // no sentAt/receivedAt
    assert.match(block, /^Jon wrote:\n/);          // exactly "Jon wrote:", no "On "
    assert.doesNotMatch(block, /Invalid Date/);
    assert.doesNotMatch(block, /On .*wrote:/);
  });

  it('collapses a newline in the sender display name', () => {
    const block = textBlockFor({ text: 'orig', name: 'Jon\nGodley', sentAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On .*, Jon Godley wrote:\n/);
  });

  it('falls back to the email when there is no display name', () => {
    const block = textBlockFor({ text: 'orig', email: 'jon@example.com', sentAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /jon@example\.com wrote:/);
  });

  it('collapses U+0085 in the display name (NEL is not in ECMAScript backslash-s)', () => {
    const nel = String.fromCharCode(0x85);
    const original = fwdOriginal({ from: [{ name: 'Eve' + nel + 'Impostor', email: 'e@x.example' }] });
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false });
    const attribution = textBlock!.split('\n').find((l) => l.includes('wrote:'))!;
    assert.equal(attribution.includes(nel), false);
    assert.match(attribution, /Eve Impostor wrote:/);
  });
});

describe('buildQuoteBlocks — the text form of the quote', () => {
  it('prefixes every quoted line (incl. blank lines) with "> "', () => {
    const original = makeOriginal({ text: 'line one\n\nline three', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /^On .*wrote:\n> line one\n> \n> line three$/);
  });

  it('quotes an html-only original via htmlToText when the message ships no html', () => {
    const original = makeOriginal({ html: '<p>Hello <b>world</b></p>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /> Hello world/);
  });
});

describe('buildQuoteBlocks — the html form of the quote', () => {
  it('wraps the quote in a cite blockquote with the portable quote-bar style and escapes the attribution', () => {
    const original = makeOriginal({ html: '<p>original <b>body</b></p>', name: 'Jon & Co', sentAt: '2026-06-15T03:29:02Z' });
    const { htmlBlock } = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(htmlBlock!, /<blockquote type="cite" style="margin:0 0 0 \.8ex;border-left:1px solid #ccc;padding-left:1ex">/);
    assert.match(htmlBlock!, /Jon &amp; Co wrote:/);           // attribution html-escaped
    assert.match(htmlBlock!, /<p>original <b>body<\/b><\/p>/); // formatting preserved
  });

  it('quotes a text-only original via an escaped html block', () => {
    const original = makeOriginal({ text: 'plain <b>not bold</b>\nsecond', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const { htmlBlock } = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(htmlBlock!, /plain &lt;b&gt;not bold&lt;\/b&gt;<br>second/); // escaped + <br>
  });

  it('quotes each format from its matching original part', () => {
    const original = makeOriginal({ text: 'orig text', html: '<p>orig html</p>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const blocks = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(blocks.textBlock!, /> orig text/);
    assert.match(blocks.htmlBlock!, /<p>orig html<\/p>/);
  });
});

describe('buildQuoteBlocks — what it will and will not read', () => {
  it('quotes an original body part that has no type (matching extractBody leniency)', () => {
    // A single-format original whose part is untyped; the reader must still read it.
    const original = {
      from: [{ name: 'Jon' }], sentAt: '2026-06-15T03:29:02Z',
      textBody: [{ partId: 't' }], htmlBody: [{ partId: 't' }],
      bodyValues: { t: { value: 'untyped body' } },
    };
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /> untyped body/);
  });

  it('yields no html block for a cid-image-only original (content-based, not string trim)', () => {
    // No orphan "On … wrote:" over an empty blockquote: the attribution goes with the quote.
    const original = makeOriginal({ html: '<div><img src="cid:logo@x"></div>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const blocks = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.equal(blocks.htmlBlock, undefined);
    assert.equal(blocks.textBlock, undefined);
  });
});

describe('buildForwardBlocks — the header block (canonical Fastmail shape)', () => {
  const textBlockFor = (over: any = {}) =>
    buildForwardBlocks({ original: fwdOriginal(over), htmlShips: false }).textBlock;

  it('emits the dashed marker + From/To/Cc/Subject/Date lines with a verbatim ISO date', () => {
    assert.match(
      textBlockFor(),
      /^----- Original message -----\nFrom: Ada Lovelace <ada@example\.com>\nTo: Bob <bob@example\.com>\nCc: carol@example\.com\nSubject: Original subject\nDate: 2026-07-01T09:14:00-04:00\n\noriginal text/,
    );
  });

  it('omits the Cc line when the original has no Cc', () => {
    assert.doesNotMatch(textBlockFor({ cc: [] }), /\nCc:/);
  });

  it('omits the whole Date line when sentAt and receivedAt are both absent (line-omission rule)', () => {
    assert.doesNotMatch(textBlockFor({ sentAt: undefined }), /\nDate:/);
  });

  it('falls back to receivedAt for the Date line', () => {
    assert.match(textBlockFor({ sentAt: undefined, receivedAt: '2026-07-02T00:00:00Z' }), /\nDate: 2026-07-02T00:00:00Z/);
  });

  it('omits the To line for a Bcc-only-received original, and the Subject line when empty', () => {
    const block = textBlockFor({ to: [], cc: [], subject: '' });
    assert.doesNotMatch(block, /\nTo:/);
    assert.doesNotMatch(block, /\nSubject:/);
  });
});

describe('signatureTextBlock — the form a sign-off takes in a text part', () => {
  it('derives the text part from the html form when html ships', () => {
    // The text part of an html message is a derived fallback, regenerated from the html on
    // the first html-only edit. Writing the configured text form verbatim here would look
    // right and then change by itself on that edit, with nothing reporting it.
    assert.equal(signatureTextBlock(BOTH_FORMS_SIG, true), 'Kind regards,\nTest User');
  });

  it('uses the configured text form, verbatim, when no html ships', () => {
    assert.equal(signatureTextBlock(BOTH_FORMS_SIG, false), 'Regards,\nTest User');
  });

  it('suppresses the image placeholder when no html ships, leaving nothing', () => {
    // No html ships, so no image ships either. '[image]' would make the recipient's entire
    // sign-off a description of something no part of the message carries.
    assert.equal(signatureTextBlock(IMAGE_ONLY_SIG, false), '');
  });

  it('derives text from an html-only identity when no html ships', () => {
    // The html body it was written for is not shipping, but the WORDS are the user's
    // sign-off; dropping them would lose a signature they really have.
    assert.equal(signatureTextBlock(HTML_ONLY_SIG, false), 'Kind regards,\nTest User');
  });

  it("renders a sign-off's own <blockquote> as '> '-prefixed lines", () => {
    // Not cosmetic: this is the only derived signature block that looks like reply quoting,
    // so anything comparing a body's lines against this block has to survive it.
    const derived = signatureTextBlock(QUOTING_SIG, true)!;
    assert.match(derived, /^> Per aspera ad astra$/m, derived);
  });

  it('writes the embedded-image placeholder when the html ships', () => {
    // The image ships with the html, so "[image]" describes something the message carries.
    assert.equal(signatureTextBlock({ html: '<img src="cid:logo">' }, true), '[image]');
  });

  it('derives the alt text of an image signature that has some', () => {
    const alted = { html: '<img src="cid:logo" alt="Test User, Example Ltd">' };
    assert.equal(signatureTextBlock(alted, false), 'Test User, Example Ltd');
  });
});

// The two forms differ in WORDING on purpose — the html says "Kind regards", the configured
// text says "Regards" — so a test can tell which one a part was given. Real identities keep
// the two in sync (Fastmail writes both), which is exactly why a bug here would be invisible
// against a matched pair.
const BOTH_FORMS_SIG = { ...HTML_ONLY_SIG, text: 'Regards,\nTest User' };
// An html sign-off that is nothing but an embedded image.
const IMAGE_ONLY_SIG = { html: '<img src="cid:logo">' };
// An html sign-off that is real markup and renders to no words at all.
const MARKUP_ONLY_SIG = { html: '<div><br></div>' };

describe('signatureHtmlBlock — the form a sign-off takes in an html part', () => {
  it('escapes a text-only identity into html rather than dropping it', () => {
    // No html form was configured, but the WORDS are the user's sign-off. They are escaped
    // and line-broken into html, never emitted raw and never silently skipped.
    const block = signatureHtmlBlock({ text: 'Regards,\nTest & User' });
    assert.equal(block, '<div>Regards,<br>Test &amp; User</div>');
  });

  it('neither block function invents a sign-off for an identity that has none', () => {
    assert.equal(signatureHtmlBlock(undefined), undefined);
    assert.equal(signatureTextBlock(undefined, true), undefined);
    assert.equal(signatureTextBlock(undefined, false), undefined);
  });
});

describe('signatureBlock — which form a part gets, and when it gets none', () => {
  it('gives a text part the DERIVED form when the message ships html', () => {
    const block = signatureBlock(BOTH_FORMS_SIG, 'textBody', true);
    assert.deepEqual(block, { available: true, content: 'Kind regards,\nTest User' });
  });

  it('gives a text part the CONFIGURED form when the message ships no html', () => {
    // The third argument is the MESSAGE question, not "does this part carry a token": a
    // message whose html part is blank ships no html, so its text part is the whole message
    // and gets the text form its owner actually wrote.
    const block = signatureBlock(BOTH_FORMS_SIG, 'textBody', false);
    assert.deepEqual(block, { available: true, content: 'Regards,\nTest User' });
  });

  it('reports no-text-form rather than shipping a bare image placeholder', () => {
    // Having a signature with no form this part can hold is a different sentence from having
    // none at all, and the two causes read differently to the caller.
    assert.deepEqual(
      signatureBlock(IMAGE_ONLY_SIG, 'textBody', false),
      { available: false, cause: 'no-text-form' },
    );
    assert.deepEqual(
      signatureBlock(undefined, 'textBody', false),
      { available: false, cause: 'no-signature' },
    );
  });

  it('reports no-text-form for an html sign-off that renders to whitespace', () => {
    // The derived form here is not undefined, it is "\n" — real markup carrying no words.
    // The guard is isBlank, so this does not become a text part whose sign-off is a newline.
    assert.equal(signatureTextBlock(MARKUP_ONLY_SIG, true), '\n');
    assert.deepEqual(
      signatureBlock(MARKUP_ONLY_SIG, 'textBody', true),
      { available: false, cause: 'no-text-form' },
    );
    // The html part still gets it: the markup is what that part was configured with.
    assert.deepEqual(
      signatureBlock(MARKUP_ONLY_SIG, 'htmlBody', true),
      { available: true, content: '<div><div><br></div></div>' },
    );
  });
});
