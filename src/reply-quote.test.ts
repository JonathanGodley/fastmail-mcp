import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildForwardBlocks, buildQuoteBlocks, signatureCidRefs } from './reply-quote.js';

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
// What a block SAYS is a property of the block alone, which no join can change, so these read
// the builder's return value directly.
//
// Deliberately NOT here, even though the same builders produce it: anything whose subject is
// what the STORED message contains. The quote sanitiser and the forward header's escaping of
// hostile fields both answer "can this construct reach a recipient", and asserting them over
// a block would answer a strictly narrower question in the same words — the block could be
// clean and the join still ship the construct. Those live over the stored parts, in
// draft-email-handler.test.ts.

// Late import, beside the suites that use it.
import { emptyQuoteImages, signatureBlock, signatureHtmlBlock, signatureTextBlock } from './reply-quote.js';

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
    const block = textBlockFor({ text: 'orig', name: 'Alex Example', sentAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On Mon, Jun 15, 2026, at 1:29 PM, Alex Example wrote:\n/);
  });

  it('uses sentAt over receivedAt', () => {
    const block = textBlockFor({
      text: 'orig', name: 'Alex', sentAt: '2026-06-15T03:29:02Z', receivedAt: '2026-06-15T09:00:00Z',
    });
    assert.match(block, /at 1:29 PM, Alex wrote:/); // 1:29 PM = sentAt, not the 7 PM receivedAt
  });

  it('falls back to receivedAt when sentAt is absent', () => {
    const block = textBlockFor({ text: 'orig', name: 'Alex', receivedAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On Mon, Jun 15, 2026, at 1:29 PM, Alex wrote:\n/);
  });

  it('omits the date entirely (never "Invalid Date") when no timestamp is present', () => {
    const block = textBlockFor({ text: 'orig', name: 'Alex' });
    assert.match(block, /^Alex wrote:\n/);
    assert.doesNotMatch(block, /Invalid Date/);
    assert.doesNotMatch(block, /On .*wrote:/);
  });

  it('trims the sender display name', () => {
    assert.match(textBlockFor({ text: 'orig', name: '  Alex  ' }), /^Alex wrote:\n/);
  });

  it('names the sender "unknown" when the original has no usable sender', () => {
    const fromCases = [[], undefined, null, [{ name: ' \n ', email: '' }], [{ name: null, email: null }]];
    for (const from of fromCases) {
      const original = { ...makeOriginal({ text: 'orig', sentAt: '2026-06-15T03:29:02Z' }), from };
      const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
      assert.equal(textBlock, 'On Mon, Jun 15, 2026, at 1:29 PM, unknown wrote:\n> orig', JSON.stringify(from));
      const undated = buildQuoteBlocks({ original: { ...makeOriginal({ text: 'orig' }), from }, htmlShips: false });
      assert.equal(undated.textBlock, 'unknown wrote:\n> orig', JSON.stringify(from));
    }
  });

  it('falls back to the email when the display name is whitespace only', () => {
    assert.match(textBlockFor({ text: 'orig', name: ' \t ', email: 'jon@example.com' }), /^jon@example\.com wrote:\n/);
  });

  it('treats a name or email of only invisible characters as missing', () => {
    assert.match(textBlockFor({ text: 'orig', name: '​‍', email: 'jon@example.com' }), /^jon@example\.com wrote:\n/);
    assert.match(textBlockFor({ text: 'orig', name: '​', email: '﻿ ­' }), /^unknown wrote:\n/);
  });

  it('collapses a newline in the sender display name', () => {
    const block = textBlockFor({ text: 'orig', name: 'Alex\nExample', sentAt: '2026-06-15T03:29:02Z' });
    assert.match(block, /^On .*, Alex Example wrote:\n/);
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
    const original = makeOriginal({ text: 'line one\n\nline three', name: 'Alex', sentAt: '2026-06-15T03:29:02Z' });
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /^On .*wrote:\n> line one\n> \n> line three$/);
  });

  it('quotes an html-only original via htmlToText when the message ships no html', () => {
    const original = makeOriginal({ html: '<p>Hello <b>world</b></p>', name: 'Alex', sentAt: '2026-06-15T03:29:02Z' });
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /> Hello world/);
  });
});

describe('buildQuoteBlocks — the html form of the quote', () => {
  it('wraps the quote in a cite blockquote with the portable quote-bar style and escapes the attribution', () => {
    const original = makeOriginal({ html: '<p>original <b>body</b></p>', name: 'Alex & Co', sentAt: '2026-06-15T03:29:02Z' });
    const { htmlBlock } = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(htmlBlock!, /<blockquote type="cite" style="margin:0 0 0 \.8ex;border-left:1px solid #ccc;padding-left:1ex">/);
    assert.match(htmlBlock!, /Alex &amp; Co wrote:/);
    assert.match(htmlBlock!, /<p>original <b>body<\/b><\/p>/);
  });

  it('quotes a text-only original via an escaped html block', () => {
    const original = makeOriginal({ text: `plain <b>not bold</b> "q" it's\nsecond`, name: 'Alex', sentAt: '2026-06-15T03:29:02Z' });
    const { htmlBlock } = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(htmlBlock!, /plain &lt;b&gt;not bold&lt;\/b&gt; &quot;q&quot; it&#39;s<br>second/);
  });

  it('quotes each format from its matching original part', () => {
    const original = makeOriginal({ text: 'orig text', html: '<p>orig html</p>', name: 'Alex', sentAt: '2026-06-15T03:29:02Z' });
    const blocks = buildQuoteBlocks({ original, htmlShips: true, timezone: TZ });
    assert.match(blocks.textBlock!, /> orig text/);
    assert.match(blocks.htmlBlock!, /<p>orig html<\/p>/);
  });
});

describe('buildQuoteBlocks — what it will and will not read', () => {
  it('quotes an original body part that has no type (matching extractBody leniency)', () => {
    const original = {
      from: [{ name: 'Alex' }], sentAt: '2026-06-15T03:29:02Z',
      textBody: [{ partId: 't' }], htmlBody: [{ partId: 't' }],
      bodyValues: { t: { value: 'untyped body' } },
    };
    const { textBlock } = buildQuoteBlocks({ original, htmlShips: false, timezone: TZ });
    assert.match(textBlock!, /> untyped body/);
  });

  const withParts = (textBody: any[], htmlBody: any[], bodyValues: Record<string, any>) =>
    ({ from: [{ name: 'Alex', email: 'alex@example.com' }], textBody, htmlBody, bodyValues });

  it('skips a part of the other body type rather than quoting its raw markup', () => {
    // JMAP puts the html part in textBody when an original has no text alternative.
    const html = [{ partId: 'h', type: 'text/html' }];
    const original = withParts(html, html, { h: { value: '<p>Hello</p>' } });
    assert.equal(buildQuoteBlocks({ original, htmlShips: false }).textBlock, 'Alex wrote:\n> Hello');
  });

  it('joins several parts with a newline and skips a part with no body value', () => {
    const parts = ['a', 'gone', 'b'].map((partId) => ({ partId, type: 'text/plain' }));
    const original = withParts(parts, [], { a: { value: 'one' }, b: { value: 'two' } });
    assert.equal(buildQuoteBlocks({ original, htmlShips: false }).textBlock, 'Alex wrote:\n> one\n> two');
  });

  it('marks a truncated part at the end of each form', () => {
    const original = withParts(
      [{ partId: 't', type: 'text/plain' }], [{ partId: 'h', type: 'text/html' }],
      { t: { value: 'text', isTruncated: true }, h: { value: '<p>html</p>', isTruncated: true } },
    );
    const quote = buildQuoteBlocks({ original, htmlShips: true });
    assert.equal(quote.textBlock, 'Alex wrote:\n> text\n> […]');
    assert.match(quote.htmlBlock!, /<p>html<\/p><div>\[…\]<\/div><\/blockquote>$/);
    const forward = buildForwardBlocks({ original, htmlShips: true });
    assert.ok(forward.textBlock.endsWith('\n\ntext\n[…]'), forward.textBlock);
    assert.match(forward.htmlBlock, /<p>html<\/p><div>\[…\]<\/div><\/div>$/);
  });

  it('strips a sentinel the body value already carries', () => {
    const original = makeOriginal({ text: 'one[body truncated] two[encoding issues detected]', name: 'Alex' });
    assert.equal(buildQuoteBlocks({ original, htmlShips: false }).textBlock, 'Alex wrote:\n> one two');
  });

  it('quotes the html when the text part is only whitespace', () => {
    const original = makeOriginal({ text: '  \n ', html: '<p>from html</p>', name: 'Alex' });
    assert.equal(buildQuoteBlocks({ original, htmlShips: false }).textBlock, 'Alex wrote:\n> from html');
  });

  it('quotes an original whose only content is a remote image', () => {
    const original = makeOriginal({ html: '<img alt="" src="https://img.example/a.png">', name: 'Alex' });
    const { htmlBlock } = buildQuoteBlocks({ original, htmlShips: true });
    assert.match(htmlBlock!, /<img alt="" src="https:\/\/img\.example\/a\.png" \/><\/blockquote>$/);
  });

  it('does not quote an original whose only content is a link around a dropped image', () => {
    for (const img of ['<img src="foo.png">', '<img src="cid:gone@x.example">']) {
      const original = makeOriginal({ html: `<a href="https://x.example/">${img}</a>`, name: 'Alex' });
      const quoteImages = { sourceParts: [] };
      assert.equal(buildQuoteBlocks({ original, htmlShips: true, quoteImages }).htmlBlock, undefined, img);
      assert.equal(buildForwardBlocks({ original, htmlShips: true, quoteImages }).htmlQuotable, false, img);
    }
    const linked = makeOriginal({ html: '<a href="https://x.example/">site</a>', name: 'Alex' });
    assert.notEqual(buildQuoteBlocks({ original: linked, htmlShips: true }).htmlBlock, undefined);
  });

  it('does not quote an original whose only image is one the quote would drop', () => {
    for (const src of ['foo.png', './a.png', '/a.png', 'mailto:a@example.com']) {
      const original = makeOriginal({ html: `<img src="${src}">`, name: 'Alex' });
      for (const htmlShips of [true, false]) {
        const quoteImages = { sourceParts: [] };
        const quote = buildQuoteBlocks({ original, htmlShips, quoteImages });
        assert.equal(quote.htmlBlock, undefined, src);
        assert.equal(buildForwardBlocks({ original, htmlShips, quoteImages }).htmlQuotable, false, src);
      }
    }
  });

  it('builds nothing for a missing original', () => {
    assert.deepEqual(buildQuoteBlocks({ original: undefined, htmlShips: true }), { images: emptyQuoteImages() });
    const forward = buildForwardBlocks({ original: undefined, htmlShips: true });
    assert.equal(forward.textBlock, '----- Original message -----');
    assert.equal(forward.htmlBlock, '<div>----- Original message -----<br></div>');
  });

  it('yields no html block for a cid-image-only original (content-based, not string trim)', () => {
    // No orphan "On … wrote:" over an empty blockquote: the attribution goes with the quote.
    const original = makeOriginal({ html: '<div><img src="cid:logo@x"></div>', name: 'Alex', sentAt: '2026-06-15T03:29:02Z' });
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

  it('omits the From line when there is no sender, and the Subject line when there is no subject', () => {
    const block = textBlockFor({ from: [], subject: undefined });
    assert.doesNotMatch(block, /\nFrom:/);
    assert.doesNotMatch(block, /\nSubject:/);
  });

  it('joins addresses with ", " and skips an entry with nothing to show', () => {
    const to = [null, {}, { email: ' ' }, { name: 'Bob', email: 'bob@example.com' }, { email: 'dee@example.com' }];
    assert.match(textBlockFor({ to }), /\nTo: Bob <bob@example\.com>, dee@example\.com\n/);
  });

  it('shows an address entry with a name and no email as the name alone', () => {
    assert.match(textBlockFor({ to: [{ name: 'Bob' }] }), /\nTo: Bob\n/);
  });

  it('puts a text-only original into the html block as escaped text', () => {
    const { htmlBlock } = buildForwardBlocks({ original: fwdOriginal({ htmlBody: [] }), htmlShips: true });
    assert.ok(htmlBlock.endsWith('<div type="cite">original text</div>'), htmlBlock);
  });

  it('is the header block alone over an attachment-only original', () => {
    const { textBlock, htmlBlock } = buildForwardBlocks({ original: attachmentOnlyOriginal(), htmlShips: true });
    assert.ok(textBlock.endsWith('\nDate: 2026-07-01T09:14:00-04:00'), textBlock);
    assert.ok(htmlBlock.endsWith('Date: 2026-07-01T09:14:00-04:00<br></div>'), htmlBlock);
  });
});

describe('the image outcome the builders report', () => {
  const PART = { cid: 'logo@x.example', blobId: 'B1', type: 'image/png' };
  const original = (html: string) => fwdOriginal({ textBody: [], sentAt: undefined, bodyValues: { h: { value: html } } });
  const mint = () => 'minted@x.example';

  it('emptyQuoteImages carries nothing', () => {
    assert.deepEqual(emptyQuoteImages(), {
      minted: [], mappings: [], resolvedParts: [], unresolvedRefs: [],
      droppedDataImages: 0, droppedUnsupportedImages: 0, htmlQuoteShips: false,
    });
  });

  it('reports nothing when no image channel is given', () => {
    const textOnly = makeOriginal({ text: 'orig', name: 'Alex' });
    assert.deepEqual(buildQuoteBlocks({ original: textOnly, htmlShips: true }).images, emptyQuoteImages());
    assert.deepEqual(buildForwardBlocks({ original: fwdOriginal(), htmlShips: false }).images, emptyQuoteImages());
  });

  it('mints with the injected mint', () => {
    const quoteImages = { sourceParts: [PART], mint };
    const html = '<p>x</p><img src="cid:logo@x.example">';
    for (const built of [
      buildQuoteBlocks({ original: original(html), htmlShips: true, quoteImages }),
      buildForwardBlocks({ original: original(html), htmlShips: true, quoteImages }),
    ]) {
      assert.equal(built.images.minted[0]?.cid, 'minted@x.example');
      assert.match(built.htmlBlock!, /<img src="cid:minted@x\.example" \/>/);
    }
  });

  it('counts an unsupported image when html ships, with or without an image channel', () => {
    const html = original('<p>x</p><img src="ftp://img.example/a.png">');
    for (const quoteImages of [{ sourceParts: [] }, undefined]) {
      for (const [htmlShips, count] of [[true, 1], [false, 0]] as const) {
        assert.equal(buildQuoteBlocks({ original: html, htmlShips, quoteImages }).images.droppedUnsupportedImages, count);
        assert.equal(buildForwardBlocks({ original: html, htmlShips, quoteImages }).images.droppedUnsupportedImages, count);
      }
    }
  });

  it('reports an image dropped from html that is not quoted, when html ships', () => {
    const html = fwdOriginal({ sentAt: undefined, bodyValues: { t: { value: 'plain body' }, h: { value: '<img src="foo.png">' } } });
    const quoteImages = { sourceParts: [] };
    for (const [htmlShips, count] of [[true, 1], [false, 0]] as const) {
      assert.equal(buildQuoteBlocks({ original: html, htmlShips, quoteImages }).images.droppedUnsupportedImages, count);
      assert.equal(buildForwardBlocks({ original: html, htmlShips, quoteImages }).images.droppedUnsupportedImages, count);
    }
  });

  it('counts a dropped data: image only when html ships', () => {
    const html = original('<p>x</p><img src="data:image/png;base64,AA">');
    for (const quoteImages of [{ sourceParts: [] }, undefined]) {
      for (const [htmlShips, count] of [[true, 1], [false, 0]] as const) {
        assert.equal(buildQuoteBlocks({ original: html, htmlShips, quoteImages }).images.droppedDataImages, count);
        assert.equal(buildForwardBlocks({ original: html, htmlShips, quoteImages }).images.droppedDataImages, count);
      }
    }
  });

  it('reports no dropped image when nothing of the original is quoted', () => {
    for (const img of ['<img src="/logo.png">', '<img src="//cdn.example.com/a.png">', '<img src="data:image/png;base64,AA">']) {
      const nothing = original(img);
      const quoteImages = { sourceParts: [] };
      for (const { images } of [
        buildQuoteBlocks({ original: nothing, htmlShips: true, quoteImages }),
        buildForwardBlocks({ original: nothing, htmlShips: true, quoteImages }),
      ]) {
        assert.deepEqual([images.droppedDataImages, images.droppedUnsupportedImages], [0, 0], img);
      }
    }
  });

  it('writes no placeholder in the text form for an image the message does not carry', () => {
    const html = original('<p>x</p><img src="cid:gone@x.example">');
    const quoteImages = { sourceParts: [] };
    for (const htmlShips of [true, false]) {
      assert.equal(buildQuoteBlocks({ original: html, htmlShips, quoteImages }).textBlock, 'Ada Lovelace wrote:\n> x');
      assert.ok(buildForwardBlocks({ original: html, htmlShips, quoteImages }).textBlock.endsWith('\n\nx'));
    }
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
    // No html form was configured, but the WORDS are the user's sign-off.
    const block = signatureHtmlBlock({ text: 'Regards,\nTest & User' });
    assert.equal(block, '<div>Regards,<br>Test &amp; User</div>');
  });

  it('neither block function invents a sign-off for an identity that has none', () => {
    assert.equal(signatureHtmlBlock(undefined), undefined);
    assert.equal(signatureTextBlock(undefined, true), undefined);
    assert.equal(signatureTextBlock(undefined, false), undefined);
    assert.equal(signatureHtmlBlock({}), undefined);
    assert.equal(signatureTextBlock({}, true), undefined);
    assert.equal(signatureTextBlock({}, false), undefined);
    assert.deepEqual(signatureBlock({}, 'htmlBody', true), { available: false, cause: 'no-signature' });
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

describe('signatureCidRefs', () => {
  it('reads the embedded-image references out of the html signature', () => {
    assert.deepEqual(signatureCidRefs({ html: '<div>R</div><img src="cid:logo">' }), ['logo']);
  });

  it('is empty for no signature, and for a text-only one', () => {
    assert.deepEqual(signatureCidRefs(undefined), []);
    assert.deepEqual(signatureCidRefs({ text: 'Regards' }), []);
  });
});
