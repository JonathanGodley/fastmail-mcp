import { convert } from 'html-to-text';
import { InvalidInputError } from './coerce.js';
import { classifyImgSrc } from './inline-images.js';

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

// ---------------------------------------------------------------------------
// Caller-supplied body validation (#62, #71/#77, #78)
// ---------------------------------------------------------------------------
// Three malformed-body shapes that the compose paths used to accept silently, each
// reaching a recipient before anyone noticed. Applied to the CALLER's own textBody /
// htmlBody at every compose seam, BEFORE a reply quote or forwarded-message block is
// merged in — a merged body is partly server-generated and partly the quoted original's
// content, and validating that would reject legitimate mail (e.g. replying to a message
// that quotes an XML snippet).

// An escaped tag run (`&lt;p&gt;`), and any real element.
//
// The escaped test is deliberately narrow, because it only ever fires on a body with NO
// real markup — i.e. on prose, where escaped angle brackets are ordinary content. A loose
// "&lt; anything &gt;" test rejected real messages: "Hi &lt;name&gt;, see attached.",
// "mail me at &lt;a@b.example&gt;", "reply with &lt;approve&gt; or &lt;reject&gt;". So the
// escaped tag NAME must be a known HTML element, and the lookahead requires a genuine tag
// delimiter after it (whitespace, `/`, or the closing `&gt;`) — which is what separates
// `&lt;a href=…&gt;` (an anchor) from `&lt;a@b.example&gt;` (an email address).
const ESCAPED_TAG = /&lt;\/?(p|br|div|span|a|b|i|u|em|strong|ul|ol|li|h[1-6]|table|thead|tbody|tr|td|th|img|pre|code|blockquote|hr|body|html|head|style|font|sub|sup|small|big|center)(?=\s|\/|&gt;)/i;
const REAL_TAG = /<[a-z][^>]*>/i;
// Case-insensitive: an HTML parser treats `<!` as the start of a markup declaration
// regardless of the case that follows, so a lowercase spelling is no safer.
const CDATA_OPEN = /<!\[CDATA\[/i;
const CDATA_START = /^<!\[CDATA\[/i;

// Reject a present-but-non-string body (#62). `undefined` AND `null` both mean "omitted":
// null is how several lenient clients spell an unset optional field, and every downstream
// body check (isBlank, the falsy guards) already reads it as absent — so accepting it here
// keeps the existing behaviour rather than turning a working call into an error.
function requireBodyString(name: string, value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    const got = Array.isArray(value) ? 'array' : typeof value;
    throw new InvalidInputError(`${name} must be a string; received ${got}. Pass the message body as a plain string.`);
  }
  return value;
}

// Validate the caller's body parameters. Reads only textBody / htmlBody, so the whole
// tool-args object can be passed straight in. Throws InvalidInputError (mapped to
// InvalidParams at the tool boundary) on:
//
//  - a present, non-string body (#62);
//  - an htmlBody that is entirely HTML-ESCAPED markup — escaped element tags and zero
//    real elements (#71/#77). Rejecting beats unescaping, which would guess at intent.
//    htmlBody only: literal `<p>` characters in a plain-text body are ordinary content,
//    and escaped markup inside real tags (`<pre>&lt;p&gt;</pre>`) is legitimate HTML;
//  - a CDATA section (#78). Asymmetric by format, because the damage is:
//      htmlBody — rejected wherever `<![CDATA[` appears. The two parsers that read the body
//        disagree, and both outcomes are bad. This server's html-to-text derivation
//        (htmlparser2) recognizes the section and consumes everything from the opening token
//        to the next `]]>`, or to the end of the body when unclosed — measured: a CDATA-
//        wrapped body derives to '', and one wrapping only the prose the caller wrote above
//        a `{{quote}}` derives to the quoted original alone, the new message silently gone
//        under someone else's words. A browser instead handles
//        `<![CDATA[` as a bogus comment that ends at the first `>`, so it drops the opening
//        token and renders the trailing `]]>` as visible text. Mid-body sections do the same
//        damage, hence "anywhere" rather than "at the start". To show a literal CDATA token
//        in HTML it must be escaped anyway (`&lt;![CDATA[`), which passes.
//      textBody — rejected only when the body STARTS with `<![CDATA[`, i.e. the caller
//        wrapped the whole body. A plain-text part is never markup-parsed, so an embedded
//        CDATA token is inert, and mail that quotes an XML snippet is perfectly legitimate
//        content that must keep working. A bare `]]>` is left alone in BOTH formats for
//        the same reason: without an opening token it renders as literal text and survives
//        the text derivation intact, so rejecting it would only block real prose.
//        The textBody refusal names HTMLBODY as the place to fix it, because the caller
//        that trips it most often never typed the token: an htmlBody may legally carry an
//        escaped `&lt;![CDATA[`, that escape unescapes when the plain-text alternative is
//        derived, and handing the derived text back on a later edit is then refused. A
//        remedy pointing at the text part would be unactionable there — the part is
//        regenerated from the markup, so only the markup can change it. It also says to
//        OMIT textBody, because fixing the html and handing the derived text back beside it
//        trips this same refusal: the check reads the caller's arguments, not the draft.
//        The sentence is CONDITIONAL ("if you did not write that token") rather than split
//        per surface the way `rejectReservedCidRef` is, deliberately. That one split because
//        its diagnosis clause was FALSE on one of its two surfaces, which is a correctness
//        problem. Every clause here is TRUE on both: a compose caller did type the token,
//        and the sentence tells it so before offering the other reading. Matching the shape
//        of a neighbouring message for its own sake would be churn.
export function assertBodyInputs(bodies: { textBody?: unknown; htmlBody?: unknown }): void {
  const text = requireBodyString('textBody', bodies?.textBody);
  const html = requireBodyString('htmlBody', bodies?.htmlBody);

  if (text !== undefined && CDATA_START.test(text.trimStart())) {
    throw new InvalidInputError(
      'textBody is wrapped in a CDATA section. Pass the message body as a plain string with no <![CDATA[ ... ]]> wrapper. If you did not write that token, it was derived rather than typed: the plain-text part of an HTML message is generated from the markup, and an escaped &lt;![CDATA[ in htmlBody unescapes into it. In that case change htmlBody and omit textBody entirely — the text part is regenerated from the markup, so editing textBody alone cannot clear this, and handing the derived text back alongside your corrected htmlBody trips this same refusal again.',
    );
  }

  if (html === undefined) return;

  if (CDATA_OPEN.test(html)) {
    throw new InvalidInputError(
      'htmlBody contains a CDATA section (<![CDATA[), which is not valid in an HTML email body: the plain-text alternative is derived with an HTML parser that drops the section and everything inside it, so the message would be lost from it, while the rendered HTML shows a stray ]]>. Pass the body as plain markup, or escape the token as &lt;![CDATA[ to show it literally.',
    );
  }

  if (ESCAPED_TAG.test(html) && !REAL_TAG.test(html)) {
    throw new InvalidInputError(
      'htmlBody appears to be HTML-escaped: it contains escaped tag sequences (&lt;p&gt;) and no actual HTML elements, so recipients would see the tags as text. Pass real markup (<p>...</p>), or use textBody for a plain-text message.',
    );
  }
}

/**
 * What the plain-text derivation writes for an image that carries no alt text.
 *
 * Alt text always wins and is emitted verbatim in every mode — this only decides what
 * happens in its ABSENCE, and only for images the message actually embeds. A remote
 * (http/https) image with no alt writes nothing under every mode: it is decoration by
 * default, and inventing a placeholder for it would put "[image]" into the text
 * alternative of every tracking pixel and spacer gif on the web.
 *
 *  - `suppress`      — write nothing. The historical behaviour, and what a text-only
 *                      branch wants: if no html ships, no image ships either, so a
 *                      placeholder would describe something the recipient never gets.
 *  - `unconditional` — write `[image]` for ANY embedded (cid:) image. For the contexts
 *                      that derive the text alternative of a body that is about to ship:
 *                      an embedded image IS content, and deriving '' for a body whose
 *                      only content is one would leave the message with no readable
 *                      text part at all.
 *  - `resolve`       — write `[image]` only for embedded images the caller-supplied map
 *                      resolves to a part that ships. For quote derivations, where a
 *                      reference that was dropped rather than carried must not be
 *                      described as an image the reader can look at.
 *
 * A literal `cid:` reference is NEVER written into the derived text under any mode — it
 * is a MIME-internal handle and means nothing to a person reading the plain-text part.
 */
export type ImagePlaceholderPolicy = 'suppress' | 'unconditional' | 'resolve';

// The placeholder itself: bare, and deliberately not the filename. The stock html-to-text
// image formatter falls back to the src/filename (e.g. "[logo.png]") when alt is absent,
// which would (a) emit junk as the "fallback", (b) leak a cid: handle into readable text,
// and (c) make an image-only, no-alt newsletter convert to non-empty text — defeating the
// html-only degrade path for remote images.
const IMAGE_PLACEHOLDER = '[image]';

// Custom <img> formatter: alt text when there is any, else the policy's answer for this
// image (see ImagePlaceholderPolicy).
function imageFormatter(policy: ImagePlaceholderPolicy, cidMap?: ReadonlyMap<string, string>) {
  return (elem: any, _walk: any, builder: any) => {
    const alt = elem?.attribs?.alt;
    if (alt && alt.trim()) {
      builder.addInline(alt);
      return;
    }
    if (policy === 'suppress') return;
    const classified = classifyImgSrc(elem?.attribs?.src);
    if (classified.kind !== 'cid') return;
    if (policy === 'resolve' && !cidMap?.get(classified.key)) return;
    builder.addInline(IMAGE_PLACEHOLDER);
  };
}

function htmlToTextOptions(policy: ImagePlaceholderPolicy, cidMap?: ReadonlyMap<string, string>) {
  return {
    wordwrap: false as const,
    formatters: {
      imgAltOrPlaceholder: imageFormatter(policy, cidMap),
    },
    selectors: [
      { selector: 'a', options: { hideLinkHrefIfSameAsText: true } },
      { selector: 'img', format: 'imgAltOrPlaceholder' },
    ],
  };
}

// Convert HTML to a readable plain-text fallback. NEVER throws — on a converter
// failure, fall back to a minimal tag-strip so a send is never blocked. May
// legitimately return '' for image-only / empty HTML. The emptiness checks elsewhere
// run on whatever this returns, INCLUDING the catch-path output.
//
// The default policy is the historical one, so a caller that expresses no opinion gets
// exactly the text this function has always produced; every production call site states
// its policy explicitly. NOTE the catch path emits no placeholder at all — it is a bare
// tag strip with no element formatting — so on a converter failure a body whose only
// content is an embedded image still derives '', and the no-readable-body gates that the
// unconditional policy normally keeps out of reach are reached after all. An accepted
// degrade: the fallback exists so a send is never blocked by a converter fault, and it
// cannot be made image-aware without becoming a second HTML parser.
export function htmlToText(
  html: string,
  policy: ImagePlaceholderPolicy = 'suppress',
  cidMap?: ReadonlyMap<string, string>,
): string {
  try {
    return convert(html, htmlToTextOptions(policy, cidMap));
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
//
// The unconditional image policy changes no answer here — the element scan below already
// returns true for any <img>, so an embedded-image-only body was always visible. It is
// stated anyway so this gate and the derivation it guards read the same body the same way.
export function htmlHasVisibleContent(html: string): boolean {
  if (!isBlank(htmlToText(html, 'unconditional'))) return true;
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
//
// The derivation is UNCONDITIONAL about embedded images: this is the body that is about to
// ship, and a body whose only content is an embedded image has real content. Deriving ''
// for it would flag the message html-only and hand it to the no-readable-body gates, which
// would refuse a message that displays perfectly well.
export function normalizeBodies(input: { textBody?: string; htmlBody?: string }): {
  textBody?: string; htmlBody?: string; htmlOnly?: boolean;
} {
  const text = !isBlank(input.textBody) ? input.textBody : undefined;
  const html = !isBlank(input.htmlBody) ? input.htmlBody : undefined;
  if (html && !text) {
    const derived = htmlToText(html, 'unconditional');
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
