import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyBodies, hasQuoteMarker, hasTextQuoteMarker } from './reply-quote.js';

describe('hasQuoteMarker (#37 reply-quote detection)', () => {
  it('detects the marker buildReplyBodies emits', () => {
    const html = buildReplyBodies({
      original: { from: [{ email: 'a@b.com' }], textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [], bodyValues: { t: { value: 'hi' } } },
      htmlBody: '<p>reply</p>', quoteOriginal: true,
    }).htmlBody!;
    assert.equal(hasQuoteMarker(html), true);
  });
  it('is tolerant of quote style and attribute order', () => {
    assert.equal(hasQuoteMarker('<blockquote type="cite">x</blockquote>'), true);
    assert.equal(hasQuoteMarker("<blockquote type='cite'>x</blockquote>"), true);
    assert.equal(hasQuoteMarker('<blockquote type=cite>x</blockquote>'), true);
    assert.equal(hasQuoteMarker('<blockquote class="q" type="cite">x</blockquote>'), true);
    assert.equal(hasQuoteMarker('<blockquote  TYPE = "cite">x</blockquote>'), true);
  });
  it('recognizes Gmail\'s class="gmail_quote" shape (no type="cite")', () => {
    assert.equal(hasQuoteMarker('<blockquote class="gmail_quote">x</blockquote>'), true);
    assert.equal(hasQuoteMarker('<blockquote class="gmail_quote" style="margin:0 0 0 .8ex">x</blockquote>'), true);
    assert.equal(hasQuoteMarker('<blockquote class="foo gmail_quote">x</blockquote>'), true); // multi-class
    assert.equal(hasQuoteMarker("<blockquote class='gmail_quote'>x</blockquote>"), true);
    // A different class is not a marker (only gmail_quote / type=cite are machine-emitted reply quotes).
    assert.equal(hasQuoteMarker('<blockquote class="pullquote">x</blockquote>'), false);
  });
  it('returns false for plain html and empty/nullish input', () => {
    assert.equal(hasQuoteMarker('<p>just a reply, no quote</p>'), false);
    assert.equal(hasQuoteMarker('<blockquote>not a cite</blockquote>'), false);
    assert.equal(hasQuoteMarker(''), false);
    assert.equal(hasQuoteMarker(null), false);
    assert.equal(hasQuoteMarker(undefined), false);
  });
});

describe('hasTextQuoteMarker (#42 text reply-quote detection)', () => {
  // Pin against the RAW text shapes Fastmail returns (captured from a live store/fetch round-
  // trip 2026-06-28), NOT just our buildReplyBodies output — the runtime guard reads Fastmail's
  // re-serialized bodyValues. Two shapes occur: a caller-supplied text body comes back as
  // "wrote:\n> " (one newline); the html-DERIVED text fallback comes back as "wrote:\n\n> "
  // (a blank line). The blank-line tolerance is load-bearing for the derived case.
  const RAW_DIRECT_TEXT = 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, PlanningAlerts wrote:\n> 2/2 Rowe St Eastwood NSW 2122: Change of Use and Fitout of Pilates Studio\n> \n> Contact us if you have questions.';
  const RAW_DERIVED_TEXT = 'my reply\n\n\n\nOn Sun, Jun 28, 2026, at 12:46 AM, PlanningAlerts wrote:\n\n> 1 new planning application near 6/30-32 Doomben Ave\n> 2/2 Rowe St Eastwood NSW 2122';

  it('matches the raw caller-supplied text shape ("wrote:\\n> ")', () => {
    assert.equal(hasTextQuoteMarker(RAW_DIRECT_TEXT), true);
  });
  it('matches the raw html-derived fallback shape ("wrote:\\n\\n> ", blank line)', () => {
    assert.equal(hasTextQuoteMarker(RAW_DERIVED_TEXT), true);
  });

  // Generation-side pin: the live buildReplyBodies output must keep matching, so a future
  // change to the attribution/quote format fails CI here (direct text + html-derived text).
  it('matches live buildReplyBodies output (direct text quote)', () => {
    const r = buildReplyBodies({
      original: makeOriginal({ text: 'orig line', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' }),
      textBody: 'my reply', quoteOriginal: true, timezone: TZ,
    });
    assert.equal(hasTextQuoteMarker(r.textBody!), true);
  });
  it('matches live buildReplyBodies output (html-derived text quote)', () => {
    const r = buildReplyBodies({
      original: makeOriginal({ html: '<p>orig <b>html</b></p>', name: 'Jon', sentAt: '2026-06-15T03:29:02Z' }),
      textBody: 'my reply', quoteOriginal: true, timezone: TZ,
    });
    assert.equal(hasTextQuoteMarker(r.textBody!), true);
  });

  it('matches a dated and an undated attribution', () => {
    assert.equal(hasTextQuoteMarker('reply\n\nOn Mon, Jun 15, 2026, at 1:29 PM, Jon wrote:\n> quoted'), true);
    assert.equal(hasTextQuoteMarker('reply\n\nJon wrote:\n> quoted'), true);
  });
  it('tolerates CRLF line endings', () => {
    assert.equal(hasTextQuoteMarker('reply\r\n\r\nJon wrote:\r\n> quoted'), true);
  });
  it('returns false for plain text, prose ending in "wrote:", and empty/nullish input', () => {
    assert.equal(hasTextQuoteMarker('just my reply, no quote here'), false);
    // Prose that merely ends with "wrote:" but has no following "> " quote line must NOT match
    // (the old over-loose new-body scan false-positived on exactly this).
    assert.equal(hasTextQuoteMarker('As I wrote: please review the attached document.'), false);
    assert.equal(hasTextQuoteMarker('She wrote:\nthen continued without quoting'), false);
    assert.equal(hasTextQuoteMarker(''), false);
    assert.equal(hasTextQuoteMarker(null), false);
    assert.equal(hasTextQuoteMarker(undefined), false);
  });
});

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
