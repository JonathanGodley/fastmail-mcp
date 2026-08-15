import sanitizeHtml from 'sanitize-html';
import { htmlToText, isBlank } from './body-format.js';
import { formatAddress, formatReplyDate } from './email-formatter.js';
import { sanitizeQuoteHtml } from './inline-images.js';

// Build the reply bodies (caller's new text + an attributed, top-posted quote of the
// original), matching the Fastmail web client with a portable quote-bar. createDraft adds
// the auto text/plain fallback downstream for an html-only caller reply, so this
// function only quotes the formats the caller actually supplied (no double-quoting).

// Escape the five HTML-significant characters for safe interpolation into quote markup.
function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Collapse internal whitespace runs (incl. newlines) to single spaces, so a display name
// containing a newline can't split the attribution line. The class is \s plus U+0085
// (NEL): ECMAScript \s covers \r, U+2028, and U+2029 but NOT U+0085, which IS a mandatory
// line break per UAX #14 — verified empirically 2026-07-05, so don't "simplify" back to \s.
function normalizeName(s: string): string {
  return s.replace(/[\s\u0085]+/g, ' ').trim();
}

// Strip our own truncation/encoding sentinels defensively before quoting (the raw reader
// below doesn't add them, but an upstream value might already carry one).
function stripSentinels(s: string): string {
  return s.replace(/\n?\[body truncated\]/g, '').replace(/\n?\[encoding issues detected\]/g, '');
}

// Sanitize an original's html for the HTML reply quote. Allow formatting tags; drop
// script/style/handlers/wrappers and ALL unscoped attributes (no global '*' key, so
// style=/class=/on*= are removed — style is the classic CSS-exfil/mXSS vector); pin
// schemes; and DROP an <img> whose src isn't a usable http(s) URL (cid:/data: get
// scheme-stripped to an empty src, which we remove entirely so the quote never carries a
// broken-image placeholder). This is purely a safety floor — we re-send under the user's
// From — matching what mainstream clients emit; it is not a tracker-pixel filter.
//
// A cidMap turns this into the mapping pass instead: each embedded-image reference the map
// resolves is rewritten to the Content-ID the caller is attaching, so the quote displays the
// image rather than losing it. Callers that pass no map (every compose path today) get the
// shipped behaviour byte for byte — the two sanitizers apply the same tag/attribute floor,
// and the map-less configuration admits no cid scheme at all.
function sanitizeForQuote(html: string, cidMap?: Map<string, string>): string {
  if (cidMap) return sanitizeQuoteHtml(html, { mode: 'map', cidMap }).html;
  return sanitizeHtml(html, {
    allowedTags: [
      'p', 'div', 'span', 'br', 'b', 'i', 'strong', 'em', 'u', 'a', 'ul', 'ol', 'li',
      'blockquote', 'pre', 'code', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
      'table', 'thead', 'tbody', 'tr', 'td', 'th', 'img',
    ],
    allowedAttributes: { a: ['href'], img: ['src', 'alt'] },
    allowedSchemes: ['http', 'https', 'mailto'],
    allowProtocolRelative: false,
    exclusiveFilter: (frame) => frame.tag === 'img' && !frame.attribs.src,
  });
}

// "Quotable" = the sanitized html has real visible content: non-empty text OR a surviving
// (http/https) <img>. Content-based, NOT a string trim: a cid:-image-only original
// sanitizes to e.g. <div></div> (non-empty as a string, visually empty), which must NOT
// count as quotable or we'd emit an orphan "On … wrote:" over an empty blockquote.
// Placeholder-suppressed on purpose: this reads the SANITIZED html, where an embedded image
// that was not mapped has already been dropped. A placeholder here would count an image the
// quote does not carry as quotable content, and produce an attribution line over a quote
// that shows nothing.
function isQuotable(sanitized: string): boolean {
  if (!isBlank(htmlToText(sanitized, 'suppress'))) return true;
  return /<img\b[^>]*\bsrc\s*=/i.test(sanitized);
}

// Plain text → escaped html block with <br> line breaks (for quoting a text-only original).
function textToHtmlBlock(s: string): string {
  return escapeHtml(s).replace(/\n/g, '<br>');
}

// Prefix each line (incl. blank lines) with "> " for a plain-text quote. (Fastmail does
// not emit format=flowed — verified live 2026-06-24 — so uniform "> " is correct.)
function quoteText(s: string): string {
  return s.split('\n').map((l) => '> ' + l).join('\n');
}

// Trim-based pick: an empty-but-present '' must fall through to the fallback (?? would not).
function pick(a: string | null | undefined, b: string | null | undefined): string {
  return a && a.trim() ? a : (b ?? '');
}

// Read all parts of `mimeType` from a JMAP body list, joined with \n. Alias-safe: accepts
// an untyped part (matching extractBody — strict equality would drop a typeless part the
// user just saw quoted), skips only a mismatched type, strips our sentinels, and appends
// `truncMarker` if any contributing part reports isTruncated. Returns '' when nothing matches.
function readBodyList(
  parts: any[] | undefined | null,
  bodyValues: any,
  mimeType: string,
  truncMarker: string,
): string {
  if (!parts?.length || !bodyValues) return '';
  const chunks: string[] = [];
  let truncated = false;
  for (const part of parts) {
    if (part.type && part.type !== mimeType) continue; // accept untyped, skip mismatched
    const bv = bodyValues[part.partId];
    if (!bv?.value) continue;
    chunks.push(stripSentinels(bv.value));
    if (bv.isTruncated) truncated = true;
  }
  if (chunks.length === 0) return '';
  return chunks.join('\n') + (truncated ? truncMarker : '');
}

const QUOTE_OPEN = '<blockquote type="cite" style="margin:0 0 0 .8ex;border-left:1px solid #ccc;padding-left:1ex">';

// True if html carries a reply-quote blockquote. Recognizes two machine-emitted shapes:
// `type="cite"` (what buildReplyBodies emits, also Apple Mail and the Fastmail web client) and
// Gmail's `class="gmail_quote"`. Both are tool-generated on a <blockquote>, so neither
// false-positives on a hand-written prose blockquote — a bare <blockquote> is deliberately NOT
// a marker. Tolerant of attribute order and quote style ("..." / '...' / bare). This is a
// PRESENCE check, not a content check: any such blockquote counts and an empty shell passes —
// edit_draft's guard treats originalEmailId as the authoritative way to keep/regenerate the
// quote, so a loose marker here only governs whether the guard fires. Other clients (e.g.
// Outlook's div-based quoting) aren't recognized; see the recognition residual in
// docs/email-bodies.md.
export function hasQuoteMarker(html: string | null | undefined): boolean {
  if (!html) return false;
  return /<blockquote[^>]*\b(?:type\s*=\s*["']?cite|class\s*=\s*["'][^"']*\bgmail_quote\b)/i.test(html);
}

// True if plain text carries our reply quote: an attribution line ("… wrote:") immediately
// followed (allowing blank lines between) by a "> "-prefixed quote line. buildReplyBodies
// emits exactly `${attribution}\n${quoteText(...)}`, so the runtime form is `wrote:\n> `;
// the blank-line / CRLF tolerance is belt-and-suspenders for how a store/fetch round-trip or
// a future format tweak might re-serialize it. Used ONLY on the OLD (stored) text, never on
// caller input — see edit_draft's guard. Like hasQuoteMarker this is a PRESENCE check that
// only governs whether the guard fires; originalEmailId is the authoritative keep path.
// NOTE: each `([ \t]*\r?\n)*` iteration consumes a mandatory `\r?\n` over a class disjoint
// from `\n` (no zero-width match), so this can't catastrophically backtrack — do NOT relax it
// into a `\s*` / nested-quantifier form that could.
export function hasTextQuoteMarker(text: string | null | undefined): boolean {
  if (!text) return false;
  return /\bwrote:[ \t]*\r?\n([ \t]*\r?\n)*[ \t]*>/.test(text);
}

// The original's html as the quote builders read it. Exported so a caller that has to decide
// which embedded images the quote will display — edit_draft, which resolves that against the
// draft's own surviving parts before the quote is rebuilt — reads exactly the same string
// these builders will quote, instead of reimplementing the body-list read.
export function readQuotableHtml(original: any): string {
  return readBodyList(original?.htmlBody, original?.bodyValues || {}, 'text/html', '<div>[…]</div>');
}

export function buildReplyBodies(input: {
  original: any;            // raw JMAP email from getEmailById (textBody/htmlBody arrays + bodyValues + date)
  textBody?: string;        // caller's new text
  htmlBody?: string;        // caller's new html
  quoteOriginal: boolean;
  timezone?: string;
  // Embedded-image reference -> the Content-ID the rebuilt quote should emit for it. Omitted
  // on every compose path (no draft exists yet to resolve references against), in which case
  // the quote is built exactly as it always was.
  cidMap?: Map<string, string>;
}): { textBody?: string; htmlBody?: string } {
  const { original, textBody, htmlBody, quoteOriginal, timezone, cidMap } = input;

  // Return only the formats the caller supplied (createDraft adds the text fallback later).
  const passthrough = () => ({
    ...(textBody !== undefined && { textBody }),
    ...(htmlBody !== undefined && { htmlBody }),
  });

  if (!quoteOriginal) return passthrough();

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // Determine quotable content (content-based, not raw presence).
  const sanitizedHtml = origHtml ? sanitizeForQuote(origHtml, cidMap) : '';
  const htmlQuotable = sanitizedHtml ? isQuotable(sanitizedHtml) : false;
  const textQuotable = !isBlank(origText);

  // No quotable original (attachment-only / cid-image-only / ICS-only): skip the quote AND
  // the attribution — no orphan "On … wrote:" over an empty quote.
  if (!htmlQuotable && !textQuotable) return passthrough();

  // Attribution in LOCAL time; the date is omitted (never "Invalid Date") when the original
  // has no usable sentAt/receivedAt, and the line drops the leading "On " + comma in that case.
  const senderRaw = original?.from?.[0]?.name || original?.from?.[0]?.email || '';
  const name = normalizeName(senderRaw);
  const date = formatReplyDate(original?.sentAt ?? original?.receivedAt, timezone);
  const attribution = date ? `On ${date}, ${name} wrote:` : `${name} wrote:`;

  const out: { textBody?: string; htmlBody?: string } = {};

  // Whether this reply ships an html quote at all — the text side's image policy turns on
  // it. When it does, an embedded image the quote carries is something the reader can look
  // at, so the text alternative may say an image is there; when it does not, nothing carries
  // the image and a placeholder would describe an absent thing.
  const htmlQuoteShips = htmlBody !== undefined;

  if (textBody !== undefined) {
    // text quote source: the original's text, else a readable conversion of its html. The
    // conversion runs over the RAW original html, exactly as it always has — deriving from
    // the sanitized output instead would silently change the derived text of every html
    // original, embedded images or not, because the quote floor drops tags that carry text.
    const textSource = pick(
      origText,
      htmlToText(origHtml, htmlQuoteShips ? 'resolve' : 'suppress', cidMap),
    );
    out.textBody = `${textBody ?? ''}\n\n${attribution}\n${quoteText(textSource)}`;
  }

  if (htmlBody !== undefined) {
    // rich quote: prefer the sanitized html; else a text-only original → escaped block.
    const htmlSource = htmlQuotable ? sanitizedHtml : textToHtmlBlock(origText);
    out.htmlBody = `${htmlBody ?? ''}<div><br></div><div>${escapeHtml(attribution)}</div>${QUOTE_OPEN}${htmlSource}</blockquote>`;
  }

  return out;
}

// ---------------------------------------------------------------------------
// Forward support (forward_email + edit_draft's forward guard)
// ---------------------------------------------------------------------------

// The forwarded-message block matches the canonical Fastmail shape (probed live
// 2026-07-05 against the official Fastmail client's forward): a dashed marker
// line, then From/To/Cc/Subject/Date header lines (Cc only when present), then
// the original below a blank line. The HTML wrapper is the platform's own
// <div type="cite">. Reply markers are <blockquote>-anchored (hasQuoteMarker),
// so the two marker families are disjoint by tag name, and sanitizeForQuote
// strips type= from embedded divs, so a forward quoted inside a reply (or
// pasted sanitized forward HTML) can't false-trip hasForwardMarker.
const FORWARD_MARKER_LINE = '----- Original message -----';
const FORWARD_OPEN = '<div type="cite">';

// Header lines for the forwarded-message block, unescaped (text form; the HTML
// form escapes each line). Every field is attacker-controlled content re-sent
// under the user's identity: normalizeName collapses line-splitting whitespace
// across the WHOLE composed address (defensively covering the email portion,
// not just the display name). Line-omission rule: a field with no usable value
// drops its whole line — no bare "To:" for a Bcc-only-received original, no
// "Date:" when sentAt/receivedAt are both absent.
function forwardHeaderLines(original: any): string[] {
  const joinAddrs = (list: any[] | undefined | null): string =>
    (list ?? [])
      .filter((a: any) => a && (a.email || a.name))
      .map((a: any) => normalizeName(formatAddress(a)))
      .filter(Boolean)
      .join(', ');
  const lines: string[] = [FORWARD_MARKER_LINE];
  const from = joinAddrs(original?.from);
  if (from) lines.push(`From: ${from}`);
  const to = joinAddrs(original?.to);
  if (to) lines.push(`To: ${to}`);
  const cc = joinAddrs(original?.cc);
  if (cc) lines.push(`Cc: ${cc}`);
  const subject = normalizeName(original?.subject ?? '');
  if (subject) lines.push(`Subject: ${subject}`);
  // The Date line is the JMAP sentAt/receivedAt string verbatim (ISO 8601 with
  // offset) — the platform's own forward block uses this form, deliberately not
  // the humanized formatReplyDate shape.
  const date = normalizeName(original?.sentAt ?? original?.receivedAt ?? '');
  if (date) lines.push(`Date: ${date}`);
  return lines;
}

// Build the forward bodies: the caller's note (optional) above a forwarded-
// message header block, with the original reproduced verbatim below it (no "> "
// prefixing — that is reply quoting). Unlike buildReplyBodies there is no
// passthrough: the block IS the forward's content, so this always emits it.
export function buildForwardBodies(input: {
  original: any;      // raw JMAP email from getEmailById (body lists + bodyValues + addresses)
  textBody?: string;  // caller's note, placed above the block
  htmlBody?: string;
  // See buildReplyBodies: absent on every compose path, supplied by edit_draft's rebuild.
  cidMap?: Map<string, string>;
}): { textBody?: string; htmlBody?: string } {
  const { original, textBody, htmlBody, cidMap } = input;

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');
  const sanitizedHtml = origHtml ? sanitizeForQuote(origHtml, cidMap) : '';
  const htmlQuotable = sanitizedHtml ? isQuotable(sanitizedHtml) : false;
  const textQuotable = !isBlank(origText);

  const lines = forwardHeaderLines(original);
  const textBlock = lines.join('\n');
  const htmlBlock = `<div><br>${lines.map(escapeHtml).join('<br>')}<br></div>`;

  // Content below the block. Text form: the original's own text, else a
  // readable conversion of its html; may be blank (attachment-only original),
  // in which case the block stands alone. HTML form: the sanitized original
  // html, else the original text escaped — never fabricated from nothing.
  // Which formats this forward emits is decided at the bottom of this function; the text
  // side's image policy needs the html half of that decision up front, for the same reason
  // the reply builder does — a placeholder must not describe an image no format carries.
  const htmlBlockShips = htmlBody !== undefined || (textBody === undefined && htmlQuotable);
  const textSource = pick(
    origText,
    htmlToText(origHtml, htmlBlockShips ? 'resolve' : 'suppress', cidMap),
  );
  const htmlSource = htmlQuotable ? sanitizedHtml : (textQuotable ? textToHtmlBlock(origText) : '');

  const composeText = (note: string | undefined): string => {
    const prefix = note && !isBlank(note) ? `${note}\n\n\n` : '';
    const below = !isBlank(textSource) ? `\n\n${textSource}` : '';
    return `${prefix}${textBlock}${below}`;
  };
  const composeHtml = (note: string | undefined): string => {
    const cite = htmlSource ? `${FORWARD_OPEN}${htmlSource}</div>` : '';
    return `${note ?? ''}${htmlBlock}${cite}`;
  };

  // Which formats are emitted:
  //   - caller supplied a format → that format, with the note on top (both
  //     supplied → both, each carrying the caller's own text — a custom text
  //     alternative is never silently replaced by a derived fallback);
  //   - caller supplied neither → html only when the original has quotable
  //     html; a text-only original yields a TEXT forward (the body model's
  //     "never fabricate HTML from plain text" holds for the tool's own
  //     default choice — see docs/email-bodies.md).
  const out: { textBody?: string; htmlBody?: string } = {};
  if (htmlBody !== undefined) out.htmlBody = composeHtml(htmlBody);
  if (textBody !== undefined) out.textBody = composeText(textBody);
  if (htmlBody === undefined && textBody === undefined) {
    if (htmlQuotable) out.htmlBody = composeHtml(undefined);
    else out.textBody = composeText(undefined);
  }
  return out;
}

// True if html carries a forwarded-message wrapper: a <div type="cite"> — the
// canonical Fastmail forward wrapper, which buildForwardBodies also emits.
// Attribute-keyed ONLY, never text-keyed: sanitizeForQuote strips type= from
// embedded content, so a reply quoting a forward (or pasted sanitized forward
// HTML) loses the attribute and cannot false-trip this. Disjoint from
// hasQuoteMarker by tag name (div vs blockquote) — the official Fastmail client
// uses <blockquote type="cite"> for replies and <div type="cite"> only for
// forwards (probed live 2026-07-05). Like hasQuoteMarker, a PRESENCE check that
// only governs whether edit_draft's guard fires; originalEmailId is the
// authoritative keep path.
export function hasForwardMarker(html: string | null | undefined): boolean {
  if (!html) return false;
  return /<div\b[^>]*\btype\s*=\s*["']?cite\b/i.test(html);
}

// True if plain text carries a forwarded-message attribution line: a dashed
// marker in the canonical Fastmail shape ("----- Original message -----") or
// Gmail's ("---------- Forwarded message ----------"), anchored at line start
// (mirroring hasQuoteMarker's recognition of both our own and Gmail's shape).
// The anchor is load-bearing: a reply draft QUOTING a forwarded message carries
// "> ----- Original message -----", and the "[ \t]*" prefix cannot consume the
// ">", so that draft dispatches to the reply variant. This must also match
// htmlToText(<the html block>) — an html-only forward stores a derived
// text/plain alternative (pinned by test). Matching pasted forwarded content of
// the same conventional shape is an accepted, documented cost: the guard
// over-asks loudly and noQuote resolves it in one step. Linear-time: every
// quantifier consumes from classes disjoint from its neighbors.
export function hasTextForwardMarker(text: string | null | undefined): boolean {
  if (!text) return false;
  return /(^|\n)[ \t]*-{3,}[ \t]*(Original|Forwarded) message[ \t]*-{2,}/i.test(text);
}
