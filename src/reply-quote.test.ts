import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyBodies } from './reply-quote.js';

// Build a raw-JMAP-shaped original. Single-format inputs alias their one part into both
// lists (matching Fastmail), which the alias-safe reader must handle.
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

describe('buildReplyBodies — passthrough', () => {
  it('returns caller bodies unchanged when quoteOriginal is false', () => {
    const r = buildReplyBodies({ original: makeOriginal({ text: 'hi' }), textBody: 'my reply', htmlBody: '<p>my reply</p>', quoteOriginal: false });
    assert.deepEqual(r, { textBody: 'my reply', htmlBody: '<p>my reply</p>' });
  });
  it('returns only the formats the caller supplied', () => {
    const r = buildReplyBodies({ original: makeOriginal({ text: 'hi' }), htmlBody: '<p>only html</p>', quoteOriginal: false });
    assert.deepEqual(r, { htmlBody: '<p>only html</p>' });
  });
});

describe('buildReplyBodies — attribution', () => {
  it('renders the exact captured Fastmail attribution (local time, ASCII-spaced)', () => {
    const original = makeOriginal({ text: 'orig', name: 'Jonathan Godley', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /On Mon, Jun 15, 2026, at 1:29 PM, Jonathan Godley wrote:/);
  });
  it('uses sentAt over receivedAt', () => {
    const original = makeOriginal({ text: 'orig', name: 'Jon', sentAt: '2026-06-15T03:29:02Z', receivedAt: '2026-06-15T09:00:00Z' });
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /at 1:29 PM, Jon wrote:/); // 1:29 PM = sentAt, not the 7 PM receivedAt
  });
  it('falls back to receivedAt when sentAt is absent', () => {
    const original = makeOriginal({ text: 'orig', name: 'Jon', receivedAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /On Mon, Jun 15, 2026, at 1:29 PM, Jon wrote:/);
  });
  it('omits the date entirely (never "Invalid Date") when no timestamp is present', () => {
    const original = makeOriginal({ text: 'orig', name: 'Jon' }); // no sentAt/receivedAt
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /\nJon wrote:\n/);          // exactly "Jon wrote:", no "On "
    assert.doesNotMatch(r.textBody!, /Invalid Date/);
    assert.doesNotMatch(r.textBody!, /On .*wrote:/);
  });
  it('collapses a newline in the sender display name', () => {
    const original = makeOriginal({ text: 'orig', name: 'Jon\nGodley', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /Jon Godley wrote:/);
  });
  it('falls back to the email when there is no display name', () => {
    const original = makeOriginal({ text: 'orig', email: 'jon@example.com', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /jon@example\.com wrote:/);
  });
});

describe('buildReplyBodies — text quote', () => {
  it('prefixes every quoted line (incl. blank lines) with "> "', () => {
    const original = makeOriginal({ text: 'line one\n\nline three', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'my reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /^my reply\n\nOn .*wrote:\n> line one\n> \n> line three$/);
  });
  it('quotes an html-only original via htmlToText for a text-caller reply', () => {
    const original = makeOriginal({ html: '<p>Hello <b>world</b></p>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'my reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /> Hello world/);
  });
});

describe('buildReplyBodies — html quote', () => {
  it('wraps the quote in a cite blockquote with the portable quote-bar style and escapes the attribution', () => {
    const original = makeOriginal({ html: '<p>original <b>body</b></p>', name: 'Jon & Co', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, htmlBody: '<p>my reply</p>', quoteOriginal: true, timezone: TZ });
    assert.match(r.htmlBody!, /<blockquote type="cite" style="margin:0 0 0 \.8ex;border-left:1px solid #ccc;padding-left:1ex">/);
    assert.match(r.htmlBody!, /Jon &amp; Co wrote:/);     // attribution html-escaped
    assert.match(r.htmlBody!, /<p>original <b>body<\/b><\/p>/); // formatting preserved
  });
  it('quotes a text-only original via an escaped html block for an html-caller reply', () => {
    const original = makeOriginal({ text: 'plain <b>not bold</b>\nsecond', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, htmlBody: '<p>reply</p>', quoteOriginal: true, timezone: TZ });
    assert.match(r.htmlBody!, /plain &lt;b&gt;not bold&lt;\/b&gt;<br>second/); // escaped + <br>
  });
});

describe('buildReplyBodies — both formats', () => {
  it('quotes each format from its matching original part', () => {
    const original = makeOriginal({ text: 'orig text', html: '<p>orig html</p>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, textBody: 'rt', htmlBody: '<p>rh</p>', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /> orig text/);
    assert.match(r.htmlBody!, /<p>orig html<\/p>/);
  });
});

describe('buildReplyBodies — alias-safe reader', () => {
  it('quotes an original body part that has no type (matching extractBody leniency)', () => {
    // A single-format original whose part is untyped; the reader must still read it.
    const original = {
      from: [{ name: 'Jon' }], sentAt: '2026-06-15T03:29:02Z',
      textBody: [{ partId: 't' }], htmlBody: [{ partId: 't' }],
      bodyValues: { t: { value: 'untyped body' } },
    };
    const r = buildReplyBodies({ original, textBody: 'reply', quoteOriginal: true, timezone: TZ });
    assert.match(r.textBody!, /> untyped body/);
  });
});

describe('buildReplyBodies — no quotable original', () => {
  it('skips quote AND attribution for an attachment-only original (no throw)', () => {
    const original = { from: [{ name: 'Jon' }], sentAt: '2026-06-15T03:29:02Z', textBody: [], htmlBody: [], bodyValues: {} };
    const r = buildReplyBodies({ original, textBody: 'reply', htmlBody: '<p>reply</p>', quoteOriginal: true, timezone: TZ });
    assert.deepEqual(r, { textBody: 'reply', htmlBody: '<p>reply</p>' }); // unchanged, no "wrote:"
  });
  it('skips quote AND attribution for a cid-image-only original (content-based, not string trim)', () => {
    const original = makeOriginal({ html: '<div><img src="cid:logo@x"></div>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' });
    const r = buildReplyBodies({ original, htmlBody: '<p>reply</p>', quoteOriginal: true, timezone: TZ });
    assert.equal(r.htmlBody, '<p>reply</p>'); // no orphan "On … wrote:" over an empty blockquote
    assert.doesNotMatch(r.htmlBody!, /wrote:/);
  });
});

describe('buildReplyBodies — sanitizeForQuote (via html quote output)', () => {
  const quote = (html: string) =>
    buildReplyBodies({ original: makeOriginal({ html, name: 'Jon', sentAt: '2026-06-15T03:29:02Z' }), htmlBody: '<p>r</p>', quoteOriginal: true, timezone: TZ }).htmlBody!;

  it('strips a full document down to its body content', () => {
    const out = quote('<!DOCTYPE html><html><head><style>p{color:red}</style></head><body><p>kept</p></body></html>');
    assert.match(out, /<p>kept<\/p>/);
    assert.doesNotMatch(out, /<style>|<head>|DOCTYPE|color:red/i);
  });
  it('strips script and event handlers but keeps formatting', () => {
    const out = quote('<p onclick="evil()">hi <b>bold</b> <a href="http://x.com">link</a></p><script>steal()</script>');
    assert.doesNotMatch(out, /onclick|script|steal/i);
    assert.match(out, /<b>bold<\/b>/);
    assert.match(out, /<a href="http:\/\/x\.com">link<\/a>/);
  });
  it('drops a style attribute on a kept tag (no global "*" allowance)', () => {
    const out = quote('<p style="background:url(evil)">x</p>');
    assert.doesNotMatch(out, /style="background/);
    assert.match(out, /<p>x<\/p>/);
  });
  it('keeps a real http(s) image but drops a cid:/data: image entirely (no broken placeholder)', () => {
    const out = quote('<p><img src="https://cdn/x.png" alt="real"> <img src="cid:logo@x"> <img src="data:image/png;base64,AAAA"></p>');
    assert.match(out, /<img src="https:\/\/cdn\/x\.png" alt="real"/);
    assert.doesNotMatch(out, /cid:/);
    assert.doesNotMatch(out, /data:image/);
  });
  it('handles a no-<body> fragment robustly (no regex extraction)', () => {
    const out = quote('<p>hi</p>');
    assert.match(out, /<p>hi<\/p>/);
  });
});
