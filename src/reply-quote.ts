import sanitizeHtml from 'sanitize-html';
import { htmlToText, isBlank } from './body-format.js';
import { formatReplyDate } from './email-formatter.js';

// Build the reply bodies (caller's new text + an attributed, top-posted quote of the
// original), matching the Fastmail web client with a portable quote-bar. The body-format
// law (applied downstream by createDraft/sendEmail) adds the text/plain fallback for an
// html-only caller reply, so the cross-format matrix collapses here.

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
// containing a newline can't split the attribution line.
function normalizeName(s: string): string {
  return s.replace(/\s+/g, ' ').trim();
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
function sanitizeForQuote(html: string): string {
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
function isQuotable(sanitized: string): boolean {
  if (!isBlank(htmlToText(sanitized))) return true;
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

export function buildReplyBodies(input: {
  original: any;            // raw JMAP email from getEmailById (textBody/htmlBody arrays + bodyValues + date)
  textBody?: string;        // caller's new text
  htmlBody?: string;        // caller's new html
  quoteOriginal: boolean;
  timezone?: string;
}): { textBody?: string; htmlBody?: string } {
  const { original, textBody, htmlBody, quoteOriginal, timezone } = input;

  // Return only the formats the caller supplied (the law adds the text fallback later).
  const passthrough = () => ({
    ...(textBody !== undefined && { textBody }),
    ...(htmlBody !== undefined && { htmlBody }),
  });

  if (!quoteOriginal) return passthrough();

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // Determine quotable content (content-based, not raw presence).
  const sanitizedHtml = origHtml ? sanitizeForQuote(origHtml) : '';
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

  if (textBody !== undefined) {
    // text quote source: the original's text, else a readable conversion of its html.
    const textSource = pick(origText, htmlToText(origHtml));
    out.textBody = `${textBody ?? ''}\n\n${attribution}\n${quoteText(textSource)}`;
  }

  if (htmlBody !== undefined) {
    // rich quote: prefer the sanitized html; else a text-only original → escaped block.
    const htmlSource = htmlQuotable ? sanitizedHtml : textToHtmlBlock(origText);
    out.htmlBody = `${htmlBody ?? ''}<div><br></div><div>${escapeHtml(attribution)}</div>${QUOTE_OPEN}${htmlSource}</blockquote>`;
  }

  return out;
}
