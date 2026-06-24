import { convert } from 'html-to-text';

// Body-format rule applied across every compose path: never ship an HTML body without a
// readable text/plain alternative. The text part is a DERIVED fallback, auto-generated
// from the HTML when the caller does not supply one; an explicitly-supplied textBody is
// stored verbatim. text/plain-only mail is legitimate and left untouched (we never
// fabricate HTML). Degrade gracefully: when the HTML yields no derivable text (an
// image-only message), ship it HTML-only rather than reject; only a genuinely no-body
// send (no text and no visible content) is refused. These helpers implement that rule;
// the per-function comments below describe exactly what each does.

// Zero-width / invisible characters that bare trim() leaves behind but that render as
// blank: ZWSP (U+200B), ZWNJ (U+200C), ZWJ (U+200D), BOM/ZWNBSP (U+FEFF), soft hyphen
// (U+00AD). A '&zwnj;&#8203;'-only body decodes to a "non-empty" string that is visually
// empty, so the emptiness test must strip these in addition to trim().
const ZERO_WIDTH = /[\u200B\u200C\u200D\uFEFF\u00AD]/g;

// The single emptiness predicate shared by normalizeBodies and every emit gate, so '' /
// whitespace / zero-width-only all read as "absent" consistently everywhere.
export function isBlank(s: string | undefined | null): boolean {
  return !s || s.replace(ZERO_WIDTH, '').trim() === '';
}

// Custom <img> formatter: emit the alt text only (nothing when there is no alt). The
// stock html-to-text image formatter falls back to the src/filename (e.g. "[logo.png]")
// when alt is absent, which would (a) emit junk as the "fallback" and (b) make an
// image-only, no-alt newsletter convert to non-empty text — defeating the html-only
// degrade path. Alt-only keeps accessible text where it exists and yields '' otherwise.
const HTML_TO_TEXT_OPTS = {
  wordwrap: false as const,
  formatters: {
    imgAltOnly: (elem: any, _walk: any, builder: any) => {
      const alt = elem?.attribs?.alt;
      if (alt && alt.trim()) builder.addInline(alt);
    },
  },
  selectors: [
    { selector: 'a', options: { hideLinkHrefIfSameAsText: true } },
    { selector: 'img', format: 'imgAltOnly' },
  ],
};

// Convert HTML to a readable plain-text fallback. NEVER throws — on a converter
// failure, fall back to a minimal tag-strip so a send is never blocked. May
// legitimately return '' for image-only / empty HTML. The emptiness checks elsewhere
// run on whatever this returns, INCLUDING the catch-path output.
export function htmlToText(html: string): string {
  try {
    return convert(html, HTML_TO_TEXT_OPTS);
  } catch {
    return html
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/\s+/g, ' ')
      .trim();
  }
}

// Does this HTML render anything a recipient would see? True if it converts to
// non-empty text OR carries any visible-media element (an image-only newsletter often
// renders via <img>, CSS background-image, <svg>, <video>/<picture>, <object>/<embed>).
// This is a reject gate that ERRS TOWARD SHIPPING (a false positive sends an arguably
// thin email; a false negative would block a real one), so an imperfect scan is
// safe-by-direction. Comments + CDATA are stripped first so a commented-out tag or
// prose mention doesn't trip it.
export function htmlHasVisibleContent(html: string): boolean {
  if (!isBlank(htmlToText(html))) return true;
  const stripped = html
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, ' ');
  if (/<(img|image|svg|video|picture|object|embed)[\s/>]/i.test(stripped)) return true;
  // background-image as an actual CSS value (ignore `background-image: none`).
  if (/background-image\s*:\s*(?!\s*none\b)[^;}"']+/i.test(stripped)) return true;
  return false;
}

// Derive the text/plain fallback (degrade-gracefully). If htmlBody is present and textBody
// is absent, derive the text fallback from the HTML; if that derives to empty, leave text absent
// and flag htmlOnly (an INTERNAL signal the authoring guard consumes — NOT a reject by
// itself, and not surfaced to the consumer). text-only and both-supplied pass through
// untouched (distinct content preserved). Presence uses the shared isBlank predicate.
export function normalizeBodies(input: { textBody?: string; htmlBody?: string }): {
  textBody?: string; htmlBody?: string; htmlOnly?: boolean;
} {
  const text = !isBlank(input.textBody) ? input.textBody : undefined;
  const html = !isBlank(input.htmlBody) ? input.htmlBody : undefined;
  if (html && !text) {
    const derived = htmlToText(html);
    if (isBlank(derived)) return { htmlBody: html, htmlOnly: true };
    return { textBody: derived, htmlBody: html };
  }
  return { ...(text !== undefined && { textBody: text }), ...(html !== undefined && { htmlBody: html }) };
}

// Pure shaping — NO fallback derivation (that is normalizeBodies' job). Build the JMAP
// body-part arrays + bodyValues keyed by the literal partIds 'text'/'html' (must match
// the part-array partIds). Accepts strings
// only (callers extract from JMAP part arrays first). Drops a blank body via the shared
// predicate so a cleared/empty body never emits a part.
export function buildBodyParts(input: { textBody?: string; htmlBody?: string }): {
  textBody?: Array<{ partId: string; type: string }>;
  htmlBody?: Array<{ partId: string; type: string }>;
  bodyValues?: Record<string, { value: string }>;
} {
  const text = !isBlank(input.textBody) ? input.textBody! : undefined;
  const html = !isBlank(input.htmlBody) ? input.htmlBody! : undefined;
  const out: {
    textBody?: Array<{ partId: string; type: string }>;
    htmlBody?: Array<{ partId: string; type: string }>;
    bodyValues?: Record<string, { value: string }>;
  } = {};
  if (text !== undefined) out.textBody = [{ partId: 'text', type: 'text/plain' }];
  if (html !== undefined) out.htmlBody = [{ partId: 'html', type: 'text/html' }];
  if (text !== undefined || html !== undefined) {
    out.bodyValues = {
      ...(text !== undefined && { text: { value: text } }),
      ...(html !== undefined && { html: { value: html } }),
    };
  }
  return out;
}
