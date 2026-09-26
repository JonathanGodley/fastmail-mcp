import { convert } from 'html-to-text';
import { InvalidInputError } from './coerce.js';
import { classifyImgSrc } from './inline-images.js';

// The body-format model (HTML as source of truth, text/plain as a derived fallback, an
// image-only body shipped HTML-only rather than refused) is in docs/email-bodies.md.

// Invisible characters trim() leaves behind: a '&zwnj;&#8203;'-only body is visually empty.
const ZERO_WIDTH = /[\u200B\u200C\u200D\uFEFF\u00AD]/g;

// The one emptiness predicate every emit gate shares.
export function isBlank(s: string | undefined | null): boolean {
  return !s || s.replace(ZERO_WIDTH, '').trim() === '';
}

// ---------------------------------------------------------------------------
// Caller-supplied body validation (#62, #71/#77, #78)
// ---------------------------------------------------------------------------
// Applied to the CALLER's own textBody / htmlBody, BEFORE a reply quote or forwarded block is
// merged in: validating a merged body would reject legitimate mail (a reply to a message that
// quotes an XML snippet).

// The escaped test is deliberately narrow, because it fires only on prose, where escaped
// angle brackets are ordinary content ("Hi &lt;name&gt;", "mail me at &lt;a@b.example&gt;").
// The name must be a known element and the lookahead a genuine tag delimiter, which is what
// separates `&lt;a href=…&gt;` from `&lt;a@b.example&gt;`.
const ESCAPED_TAG = /&lt;\/?(p|br|div|span|a|b|i|u|em|strong|ul|ol|li|h[1-6]|table|thead|tbody|tr|td|th|img|pre|code|blockquote|hr|body|html|head|style|font|sub|sup|small|big|center)(?=\s|\/|&gt;)/i;
const REAL_TAG = /<[a-z][^>]*>/i;
// Case-insensitive: a parser treats `<!` as a markup declaration whatever case follows.
const CDATA_OPEN = /<!\[CDATA\[/i;
const CDATA_START = /^<!\[CDATA\[/i;

// Reject a present-but-non-string body (#62). `null` is how several lenient clients spell an
// unset optional field, so it means "omitted" like `undefined`.
function requireBodyString(name: string, value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    const got = Array.isArray(value) ? 'array' : typeof value;
    throw new InvalidInputError(`${name} must be a string; received ${got}. Pass the message body as a plain string.`);
  }
  return value;
}

// Validate the caller's body parameters (#62, #71/#77, #78). Rejecting escaped markup beats
// unescaping it, which would guess at intent; that check is htmlBody only, since escaped
// markup is ordinary content in a text part or inside real tags.
//
// CDATA is asymmetric by format, because the damage is:
//   htmlBody is rejected wherever `<![CDATA[` appears. The html-to-text derivation consumes
//     everything up to the next `]]>` (a body wrapping only the prose above a `{{quote}}`
//     derives to the quoted original alone), while a browser renders a stray `]]>`.
//   textBody is rejected only when it STARTS with the token: a text part is never
//     markup-parsed, and mail quoting an XML snippet must keep working. A bare `]]>` is left
//     alone in both formats. The textBody refusal points at htmlBody and says to omit
//     textBody, because the usual trigger is an escaped `&lt;![CDATA[` in htmlBody that
//     unescaped into the derived text handed back on a later edit.
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
 * What the plain-text derivation writes for an EMBEDDED image with no alt text. Alt text
 * always wins; a remote image with no alt writes nothing in every mode, or every tracking
 * pixel would become "[image]". A `cid:` reference is never written.
 *
 *  - `suppress`      write nothing: for a text-only branch, where no image ships either.
 *  - `unconditional` write `[image]` for any embedded image: for a body about to ship, where
 *                    deriving '' for an image-only body would leave no readable text part.
 *  - `resolve`       write `[image]` only where `cidMap` resolves to a part that ships: for
 *                    quotes, where a dropped reference must not be described as viewable.
 */
export type ImagePlaceholderPolicy = 'suppress' | 'unconditional' | 'resolve';

// Deliberately not the filename, which the stock formatter falls back to: that would leak a
// cid: handle and make a no-alt remote-image newsletter derive non-empty text.
const IMAGE_PLACEHOLDER = '[image]';

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

// NEVER throws: on a converter failure it falls back to a tag strip so a send is never
// blocked. May return '' for image-only HTML. The catch path writes no image placeholder, so
// there an embedded-image-only body derives '' under every policy; an accepted degrade, since
// making it image-aware would mean a second HTML parser.
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

// Does this HTML render anything a recipient would see? A reject gate that ERRS TOWARD
// SHIPPING (a false negative would block a real message), so an imperfect scan is safe by
// direction.
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

// Derive the text/plain fallback when only HTML is supplied. `htmlOnly` is an INTERNAL signal
// for the authoring guard, not a reject and not surfaced to the consumer.
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

// Pure shaping, NO fallback derivation (that is normalizeBodies'). The bodyValues keys must
// match the part-array partIds.
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
