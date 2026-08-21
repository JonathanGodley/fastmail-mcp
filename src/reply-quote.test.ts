import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyBodies, hasQuoteMarker, hasTextQuoteMarker } from './reply-quote.js';

describe('hasQuoteMarker (#37 reply-quote detection)', () => {
  it('detects the marker buildReplyBodies emits', () => {
    const html = buildReplyBodies({
      original: { from: [{ email: 'a@b.example' }], textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [], bodyValues: { t: { value: 'hi' } } },
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
  // Only the marker region (the attribution line and the newlines before "> ") is byte-exact from
  // that capture; the quoted lines carry synthetic content, since the matcher never reads them.
  // Keep it that way — a fixture must not reproduce a real message's contents.
  const RAW_DIRECT_TEXT = 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:\n> 2/2 Example St Sampleton NSW 2000: Change of Use and Fitout of a Studio\n> \n> Contact us if you have questions.';
  const RAW_DERIVED_TEXT = 'my reply\n\n\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:\n\n> 1 new planning application near 1 Example Street\n> 2/2 Example St Sampleton NSW 2000';

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

describe('buildReplyBodies — sanitizeQuoteHtml (via html quote output)', () => {
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

// ---------------------------------------------------------------------------
// Forward support (#30)
// ---------------------------------------------------------------------------

// Late imports for the forward suites (kept beside them rather than merged into the
// header so the reply suites above stay byte-stable).
import { buildForwardBodies, hasForwardMarker, hasTextForwardMarker } from './reply-quote.js';
import { htmlToText } from './body-format.js';

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
const textOnlyOriginal = () => fwdOriginal({ htmlBody: undefined, bodyValues: { t: { value: 'plain original' } } });
const htmlOnlyOriginal = () => fwdOriginal({ textBody: undefined, bodyValues: { h: { value: '<p>rich original</p>' } } });
const attachmentOnlyOriginal = () => fwdOriginal({ textBody: undefined, htmlBody: undefined, bodyValues: {} });

describe('buildForwardBodies — header block (canonical Fastmail shape)', () => {
  it('emits the dashed marker + From/To/Cc/Subject/Date lines with a verbatim ISO date', () => {
    const { textBody } = buildForwardBodies({ original: fwdOriginal(), textBody: 'note' });
    assert.match(textBody!, /note\n\n\n----- Original message -----\nFrom: Ada Lovelace <ada@example\.com>\nTo: Bob <bob@example\.com>\nCc: carol@example\.com\nSubject: Original subject\nDate: 2026-07-01T09:14:00-04:00\n\noriginal text/);
  });
  it('omits the Cc line when the original has no Cc', () => {
    const { textBody } = buildForwardBodies({ original: fwdOriginal({ cc: [] }), textBody: 'n' });
    assert.doesNotMatch(textBody!, /\nCc:/);
  });
  it('omits the whole Date line when sentAt and receivedAt are both absent (line-omission rule)', () => {
    const { textBody } = buildForwardBodies({ original: fwdOriginal({ sentAt: undefined }), textBody: 'n' });
    assert.doesNotMatch(textBody!, /\nDate:/);
  });
  it('falls back to receivedAt for the Date line', () => {
    const { textBody } = buildForwardBodies({ original: fwdOriginal({ sentAt: undefined, receivedAt: '2026-07-02T00:00:00Z' }), textBody: 'n' });
    assert.match(textBody!, /\nDate: 2026-07-02T00:00:00Z/);
  });
  it('omits the To line for a Bcc-only-received original, and the Subject line when empty', () => {
    const { textBody } = buildForwardBodies({ original: fwdOriginal({ to: [], cc: [], subject: '' }), textBody: 'n' });
    assert.doesNotMatch(textBody!, /\nTo:/);
    assert.doesNotMatch(textBody!, /\nSubject:/);
  });
});

describe('buildForwardBodies — format emission matrix', () => {
  it('caller-neither with an html original emits an HTML forward only', () => {
    const out = buildForwardBodies({ original: htmlOnlyOriginal() });
    assert.equal(out.textBody, undefined);
    assert.match(out.htmlBody!, /<div type="cite"><p>rich original<\/p><\/div>/);
  });
  it('caller-neither with a text-only original emits a TEXT forward only (never fabricates html)', () => {
    const out = buildForwardBodies({ original: textOnlyOriginal() });
    assert.equal(out.htmlBody, undefined);
    assert.match(out.textBody!, /----- Original message -----[\s\S]*plain original/);
  });
  it('caller html only: html = note + block + sanitized original; text side left to the downstream fallback', () => {
    const out = buildForwardBodies({ original: fwdOriginal(), htmlBody: '<p>see below</p>' });
    assert.equal(out.textBody, undefined);
    assert.match(out.htmlBody!, /^<p>see below<\/p><div><br>----- Original message -----/);
  });
  it('caller html only over a text-only original: original text escaped into the html (caller chose html)', () => {
    const out = buildForwardBodies({ original: textOnlyOriginal(), htmlBody: '<p>note</p>' });
    assert.match(out.htmlBody!, /<div type="cite">plain original<\/div>/);
  });
  it('caller text only over an html original: text = note + block + htmlToText(original html)', () => {
    const out = buildForwardBodies({ original: htmlOnlyOriginal(), textBody: 'note' });
    assert.equal(out.htmlBody, undefined);
    assert.match(out.textBody!, /----- Original message -----[\s\S]*rich original/);
  });
  it('caller BOTH: both emitted, each with its own note; the custom text alternative is never replaced', () => {
    const out = buildForwardBodies({ original: fwdOriginal(), textBody: 'my custom text', htmlBody: '<p>my html</p>' });
    assert.match(out.textBody!, /^my custom text\n\n\n----- Original message -----/);
    assert.match(out.textBody!, /original text/);
    assert.match(out.htmlBody!, /^<p>my html<\/p><div><br>----- Original message -----/);
  });
  it('attachment-only original: the header block stands alone (no empty cite shell)', () => {
    const out = buildForwardBodies({ original: attachmentOnlyOriginal() });
    assert.equal(out.htmlBody, undefined);
    assert.match(out.textBody!, /----- Original message -----/);
    const htmlOut = buildForwardBodies({ original: attachmentOnlyOriginal(), htmlBody: '<p>fyi</p>' });
    assert.doesNotMatch(htmlOut.htmlBody!, /<div type="cite">/);
  });
});

describe('buildForwardBodies — hostile interpolated fields (re-sent under the user From)', () => {
  it('escapes every header field in the HTML block (subject + display names)', () => {
    const original = fwdOriginal({
      subject: '<img src=x onerror=alert(1)>',
      from: [{ name: '"><script>alert(2)</script>', email: 'evil@example.com' }],
    });
    const { htmlBody } = buildForwardBodies({ original, htmlBody: '<p>n</p>' });
    assert.doesNotMatch(htmlBody!, /<img src=x onerror/);
    assert.doesNotMatch(htmlBody!, /<script>/);
    assert.match(htmlBody!, /&lt;img src=x onerror=alert\(1\)&gt;/);
  });
  it('collapses line-splitting whitespace in text-block fields — incl. U+2028/U+2029/U+0085 (NEL is NOT in ECMAScript backslash-s)', () => {
    // Invisible characters are built from char codes so no raw control char lives in
    // this source file. The NEL assertion checks ACTUAL collapse, not just class behavior.
    const nel = String.fromCharCode(0x85);
    const ls = String.fromCharCode(0x2028);
    const ps = String.fromCharCode(0x2029);
    const original = fwdOriginal({
      from: [{ name: 'Fake' + nel + 'To: victim@example.com' + ls + 'X:' + ps + 'Y', email: 'evil@example.com' }],
      subject: 'line1\r\nline2',
    });
    const { textBody } = buildForwardBodies({ original, textBody: 'n' });
    const fromLine = textBody!.split('\n').find((l) => l.startsWith('From: '))!;
    assert.equal(fromLine.includes(nel), false);
    assert.equal(fromLine.includes(ls), false);
    assert.equal(fromLine.includes(ps), false);
    assert.match(fromLine, /^From: Fake To: victim@example\.com X: Y <evil@example\.com>$/);
    const subjectLine = textBody!.split('\n').find((l) => l.startsWith('Subject: '))!;
    assert.equal(subjectLine, 'Subject: line1 line2');
  });
  it('sanitizes the reproduced original html (script stripped, http img kept, cid img dropped)', () => {
    const original = fwdOriginal({
      bodyValues: {
        t: { value: 'x' },
        h: { value: '<p onclick="x()">hi</p><script>evil()</script><img src="https://x.com/a.png"><img src="cid:img1">' },
      },
    });
    const { htmlBody } = buildForwardBodies({ original, htmlBody: '' });
    assert.doesNotMatch(htmlBody!, /<script|onclick/);
    assert.match(htmlBody!, /<img src="https:\/\/x\.com\/a\.png" \/>/);
    assert.doesNotMatch(htmlBody!, /cid:img1/);
  });
});

describe('hasForwardMarker / hasTextForwardMarker (#30 forward detection)', () => {
  it('detects what buildForwardBodies emits (html + text) and tolerates attribute variants', () => {
    const out = buildForwardBodies({ original: fwdOriginal(), textBody: 'n', htmlBody: '<p>n</p>' });
    assert.equal(hasForwardMarker(out.htmlBody), true);
    assert.equal(hasTextForwardMarker(out.textBody), true);
    assert.equal(hasForwardMarker('<div type=cite>x</div>'), true);
    assert.equal(hasForwardMarker("<div class='q' type='cite'>x</div>"), true);
    assert.equal(hasForwardMarker('<DIV TYPE="cite">x</DIV>'), true);
  });
  it("recognizes Gmail's dashed text line (mirroring hasQuoteMarker's dual recognition)", () => {
    assert.equal(hasTextForwardMarker('---------- Forwarded message ---------'), true);
    assert.equal(hasTextForwardMarker('note\n---------- Forwarded message ----------\nFrom: x'), true);
  });
  it('must match htmlToText(<the html block>) — the derived text fallback of an html-only forward (load-bearing)', () => {
    const { htmlBody } = buildForwardBodies({ original: htmlOnlyOriginal(), htmlBody: '<p>note</p>' });
    const derived = htmlToText(htmlBody!);
    assert.equal(hasTextForwardMarker(derived), true);
  });
  it('rejects a "> "-quoted dashed line (a reply QUOTING a forward must dispatch to the reply variant)', () => {
    assert.equal(hasTextForwardMarker('On x, y wrote:\n> ----- Original message -----\n> From: a'), false);
    assert.equal(hasTextForwardMarker('> ---------- Forwarded message ----------'), false);
  });
  it('rejects prose, plain dashes, the asAttachment filler, and nullish input', () => {
    assert.equal(hasTextForwardMarker('----- meeting notes -----'), false);
    assert.equal(hasTextForwardMarker('the original message was lost'), false);
    assert.equal(hasTextForwardMarker('Forwarded message attached.'), false);
    assert.equal(hasTextForwardMarker('-----'), false);
    assert.equal(hasTextForwardMarker(''), false);
    assert.equal(hasTextForwardMarker(null), false);
    assert.equal(hasForwardMarker('<div class="cite">x</div>'), false);
    assert.equal(hasForwardMarker('<p>type="cite" mentioned in prose</p>'), false);
    assert.equal(hasForwardMarker(null), false);
  });
});

describe('cross-predicate disjointness (forward vs reply markers)', () => {
  it('buildForwardBodies html output fails hasQuoteMarker (div, not blockquote)', () => {
    const { htmlBody } = buildForwardBodies({ original: fwdOriginal(), htmlBody: '<p>n</p>' });
    assert.equal(hasQuoteMarker(htmlBody), false);
  });
  it('buildReplyBodies html output fails hasForwardMarker (blockquote, not div)', () => {
    const reply = buildReplyBodies({ original: fwdOriginal(), htmlBody: '<p>r</p>', quoteOriginal: true });
    assert.equal(hasForwardMarker(reply.htmlBody), false);
  });
  it('forwarding a REPLY cannot false-trip hasQuoteMarker (sanitizer strips the embedded type=cite)', () => {
    const replyOriginal = fwdOriginal({
      bodyValues: { t: { value: 'x' }, h: { value: '<p>top</p><blockquote type="cite">quoted</blockquote>' } },
    });
    const fwd = buildForwardBodies({ original: replyOriginal, htmlBody: '<p>n</p>' });
    assert.equal(hasQuoteMarker(fwd.htmlBody), false);
    assert.match(fwd.htmlBody!, /<blockquote>quoted<\/blockquote>/); // content kept, attribute stripped
  });
  it('a reply QUOTING a forward cannot false-trip hasForwardMarker (embedded div loses type=cite)', () => {
    const forwardOriginal = fwdOriginal({
      bodyValues: { t: { value: 'x' }, h: { value: '<div><br>----- Original message -----<br>From: a</div><div type="cite">fwd body</div>' } },
    });
    const reply = buildReplyBodies({ original: forwardOriginal, htmlBody: '<p>r</p>', quoteOriginal: true });
    assert.equal(hasForwardMarker(reply.htmlBody), false);
  });
});

describe('normalizeName NEL widening also protects the reply attribution line', () => {
  it('collapses U+0085 in a reply attribution display name', () => {
    const nel = String.fromCharCode(0x85);
    const original = fwdOriginal({ from: [{ name: 'Eve' + nel + 'Impostor', email: 'e@x.example' }] });
    const { textBody } = buildReplyBodies({ original, textBody: 'r', quoteOriginal: true });
    const attribution = textBody!.split('\n').find((l) => l.includes('wrote:'))!;
    assert.equal(attribution.includes(nel), false);
    assert.match(attribution, /Eve Impostor wrote:/);
  });
});
