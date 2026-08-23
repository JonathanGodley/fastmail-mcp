import { htmlToText, isBlank } from './body-format.js';
import { formatAddress, formatReplyDate } from './email-formatter.js';
import { buildCidMap, resolveCidRefs, sanitizeQuoteHtml } from './inline-images.js';
import type { CidMapping, CidPart, MintedInlinePart } from './inline-images.js';
import type { ResolvedSignature } from './identity.js';

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

// Both quote builders sanitize the original's html through the shared two-pass sanitizer in
// src/inline-images.ts. Its safety floor is the same one this file has always applied —
// formatting tags kept, script/style/handlers and ALL unscoped attributes dropped (no global
// '*' key; style is the classic CSS-exfil/mXSS vector), schemes pinned, and any <img> whose
// src did not survive removed entirely so a quote never shows a broken-image placeholder.
// The floor is what makes it safe to re-send someone else's markup under the user's own
// From; it is not a tracker-pixel filter.
//
// What the two passes add is WHICH embedded images the quote can show:
//
//   collect — reports the references the original's body makes and rewrites nothing. Its
//             html is exactly what the sanitizer alone would emit (no cid scheme is
//             admitted), so it is the input the quotability check reads.
//   map     — rewrites each reference that resolved to a part into the Content-ID this
//             draft attaches for it, so the quote displays the image instead of losing it.
//
// The passes run in that order for a reason: minting an identifier commits this call to
// attaching a part, and a part nothing references is a stray file on the finished message.
// So pass one decides whether an html quote ships at all, and pass two runs only on a branch
// that really ships one.

// "Quotable" = the sanitized html has real visible content: non-empty text OR a surviving
// <img>. Content-based, NOT a string trim: an original whose only content is an embedded
// image sanitizes in COLLECT mode to e.g. <div></div> (non-empty as a string, visually
// empty), which must not count on its own or we would emit an orphan "On … wrote:" over an
// empty blockquote. Such an original becomes quotable through the separate resolvability
// test below — because a reference that resolves to a carriable part is content the quote
// really will show. Placeholder-suppressed on purpose: this reads sanitized html, where an
// image that was not mapped has already been dropped, and a placeholder would count an image
// the quote does not carry as quotable content.
function isQuotable(sanitized: string): boolean {
  if (!isBlank(htmlToText(sanitized, 'suppress'))) return true;
  return /<img\b[^>]*\bsrc\s*=/i.test(sanitized);
}

/**
 * What a compose path gives a quote builder so it can resolve the original's embedded images
 * itself.
 *
 * Present on the compose paths, where no draft exists yet and the builder owns both passes.
 * Absent on the edit path, which resolves references against the draft's own surviving parts
 * before calling in and passes the finished `cidMap` instead — there the builder only
 * rewrites, and never mints.
 */
export interface QuoteImageInput {
  /** The original's parts (the gated union). Their Content-IDs are compared literally. */
  sourceParts: CidPart[];
  /** Injected so callers' tests are deterministic. */
  mint?: () => string;
}

/** What the builder decided about the original's embedded images. */
export interface QuoteImageOutcome {
  /**
   * Parts to attach so the quote can display them, each under a freshly minted Content-ID.
   * Empty whenever no html quote ships — a text-only branch mints nothing.
   */
  minted: MintedInlinePart[];
  /** One entry per reference the quote embeds, carrying the part it came from. */
  mappings: CidMapping[];
  /** Distinct parts the references resolved to, in first-reference order. */
  resolvedParts: CidPart[];
  /** References that matched no part at all. Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  /** `data:`-URI images in the original's body, which are not carried. */
  droppedDataImages: number;
  /**
   * Images the shipped quote dropped because their src was neither an embedded-image
   * reference nor an http(s) URL (a relative path, a protocol-relative URL, an exotic
   * scheme). Non-zero only on a branch that actually rewrote the quote's html.
   */
  droppedUnsupportedImages: number;
  /** Whether a quote reproducing the original's own html ships. */
  htmlQuoteShips: boolean;
}

const NO_QUOTE_IMAGES: QuoteImageOutcome = {
  minted: [], mappings: [], resolvedParts: [], unresolvedRefs: [],
  droppedDataImages: 0, droppedUnsupportedImages: 0, htmlQuoteShips: false,
};

/** An outcome that carried nothing. Exported as the fallback for a caller that always asks. */
export function emptyQuoteImages(): QuoteImageOutcome {
  return { ...NO_QUOTE_IMAGES, minted: [], mappings: [], resolvedParts: [], unresolvedRefs: [] };
}

/**
 * Pass one over an original's html: what it references, what those references resolve to,
 * and whether the html is worth quoting at all.
 *
 * `quotable` is the disjunct that makes an image-only original quotable: either the html has
 * visible content of its own, or at least one of its references would really embed. The
 * second half deliberately tests EMBEDDABILITY rather than mere resolution — a reference
 * whose Content-ID names two parts, or whose part has no blob, cannot be carried, so counting
 * it would produce an attribution over a quote showing nothing and leave the shortfall
 * sentence describing a quote that never shipped.
 */
function collectQuoteRefs(
  origHtml: string,
  images: QuoteImageInput | undefined,
  cidMap: Map<string, string> | undefined,
): { html: string; refs: string[]; droppedDataImages: number; quotable: boolean; resolvedParts: CidPart[]; unresolvedRefs: string[] } {
  if (!origHtml) {
    return { html: '', refs: [], droppedDataImages: 0, quotable: false, resolvedParts: [], unresolvedRefs: [] };
  }
  const collected = sanitizeQuoteHtml(origHtml, { mode: 'collect' });
  const resolution = images ? resolveCidRefs(collected.refs, images.sourceParts ?? []) : null;
  // On the edit path the resolution already happened elsewhere: the map holds exactly the
  // references that resolved to a carriable part, so membership answers the same question.
  const resolvesSomething = resolution
    ? resolution.embeddableRefs.length > 0
    : collected.refs.some((r) => cidMap?.has(r) === true);
  return {
    html: collected.html,
    refs: collected.refs,
    droppedDataImages: collected.droppedDataImages,
    quotable: isQuotable(collected.html) || resolvesSomething,
    resolvedParts: resolution?.resolvedParts ?? [],
    unresolvedRefs: resolution?.unresolvedRefs ?? [],
  };
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
// quote, so a loose marker here only governs whether the guard fires — and, on the keep path,
// whether a caller-supplied body is refused for already carrying the quote (#145); it can only
// ever refuse on caller input, never exempt. Other clients (e.g.
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
// a future format tweak might re-serialize it. Mostly used on the OLD (stored) text, which is
// what decides whether edit_draft's guard fires. The one caller-input use is that guard's keep
// path (#145), where a supplied body already carrying the quote originalEmailId would rebuild
// is REFUSED — read-only in the refuse direction, so caller input can never use this to make an
// edit pass. Like hasQuoteMarker this is a PRESENCE check that only governs whether the guard
// fires; originalEmailId is the authoritative keep path.
// NOTE: each `([ \t]*\r?\n)*` iteration consumes a mandatory `\r?\n` over a class disjoint
// from `\n` (no zero-width match), so this can't catastrophically backtrack — do NOT relax it
// into a `\s*` / nested-quantifier form that could.
export function hasTextQuoteMarker(text: string | null | undefined): boolean {
  if (!text) return false;
  return /\bwrote:[ \t]*\r?\n([ \t]*\r?\n)*[ \t]*>/.test(text);
}

// ---------------------------------------------------------------------------
// The sending identity's signature (#33)
// ---------------------------------------------------------------------------
//
// A third marked block, beside the reply quote and the forwarded-message block above. Same
// job as those markers: let a LATER edit tell what this server put in the body apart from
// what the caller wrote, so a body rewrite does not silently throw it away. What differs is
// the recovery — a dropped quote is challenged and rebuilt from the named original, while a
// dropped signature is simply re-appended from the identity, because the identity is where
// it came from and nothing about it depends on another message.
//
// The block is placed ABOVE any quoted or forwarded history, which is why insertion lives in
// this file at all: by the time a compose path reaches createDraft the quote is already
// concatenated onto the body, and appending there would put the sign-off underneath it.
//
// Detection is by CLASS, which is HTML-only by construction. A text-only draft carries no
// marker, so nothing PRESERVES its signature across an edit that rewrites the body — see the
// signature section of docs/email-bodies.md for that residual, which applies to a signature
// this server wrote just as much as to a hand-typed one. Not to be confused with
// idempotence, which the text path does have: applySignature declines to append a block a
// body already ends with, so the edit-time flag that recovers the residual cannot stack.
export const SIGNATURE_CLASS = 'fm-mcp-signature';

// Where a `<div` tag opens. Everything after it is read by the attribute walk below rather
// than by a regex: the tag has to be understood, not pattern-matched. Anchored on <div> so
// the marker is disjoint from both quote-marker families (blockquote type=cite, div
// type=cite) — a signature inside a quote is caught by neither of those, and a quote inside
// a signature is not a shape anything here emits.
const DIV_TAG_OPEN_RE = /<div\b/gi;

/**
 * Whether one `<div>` tag's attributes carry the marker class, reading the tag the way a
 * parser would: attribute by attribute, with quoted values consumed whole.
 *
 * A single regex over the raw tag got this wrong in both directions, and both directions
 * matter because the same predicate reads the STORED body in updateDraft, where agent-authored
 * html arrives unsanitised:
 *
 *  - FALSE POSITIVE. Any pattern that accepts a quote character as the boundary in front of
 *    `class` matches `class=` sitting INSIDE another attribute's value —
 *    `<div data-x=" class=fm-mcp-signature ">` and `<div title="a class='fm-mcp-signature'">`
 *    both did. The cost is the one the marker exists to avoid: an edit that said nothing
 *    about signatures appends one nobody asked for and announces that the draft already
 *    carried one.
 *  - FALSE NEGATIVE. `[^>]*` before the attribute cannot cross a `>` inside a quoted value,
 *    so `<div title="a > b" class="fm-mcp-signature">` — a perfectly ordinary tag — MISSED a
 *    genuine marker, which silently drops the block on the preserve path.
 *
 * Walking the attributes settles both: a quoted value is skipped as a unit whatever is in
 * it, and the name is whatever sits between the delimiters, so `data-class` and `myclass` are
 * simply different names rather than near-misses of a boundary rule.
 *
 * The class must also be a WHOLE token of the class list, which is why the value is split on
 * whitespace instead of searched: a longer name that merely starts with ours
 * (`fm-mcp-signature-ish`) is somebody else's class.
 *
 * Returns the index just past the tag so the caller can carry on scanning.
 */
function divTagCarriesMarker(html: string, from: number): { found: boolean; next: number } {
  let i = from;
  while (i < html.length) {
    while (i < html.length && /\s/.test(html[i])) i++;
    if (i >= html.length || html[i] === '>') break;
    if (html[i] === '/') { i++; continue; }
    const nameStart = i;
    while (i < html.length && !/[\s=>/]/.test(html[i])) i++;
    const name = html.slice(nameStart, i).toLowerCase();
    // Whitespace is allowed on both sides of `=` (`class = "…"`), and a valueless attribute
    // is legal, so the `=` is looked for past any run of spaces and its absence is not a
    // parse failure.
    let j = i;
    while (j < html.length && /\s/.test(html[j])) j++;
    let value = '';
    if (html[j] === '=') {
      j++;
      while (j < html.length && /\s/.test(html[j])) j++;
      const quote = html[j];
      if (quote === '"' || quote === "'") {
        const end = html.indexOf(quote, j + 1);
        // An unterminated quoted value means the rest of the document is inside it; there is
        // no further attribute to read on this tag or any other.
        if (end < 0) return { found: false, next: html.length };
        value = html.slice(j + 1, end);
        i = end + 1;
      } else {
        const valueStart = j;
        while (j < html.length && !/[\s>]/.test(html[j])) j++;
        value = html.slice(valueStart, j);
        i = j;
      }
    } else {
      i = j;
    }
    if (name === 'class' && value.split(/\s+/).includes(SIGNATURE_CLASS)) {
      return { found: true, next: i };
    }
  }
  return { found: false, next: Math.min(i + 1, html.length) };
}

// The blank line between the body and the signature, matching the spacer the reply quote
// uses below so the two read as one consistently spaced message.
const SIGNATURE_SPACER = '<div><br></div>';

/**
 * True if html carries a signature block this server wrote.
 *
 * Whole-document, so it does not care where in the body the block sits — which is why the
 * html side of the edit path needs no equivalent of the text side's history-aware search.
 *
 * Note the sanitizer strips class= from quoted content (src/inline-images.ts), so our own
 * signature quoted back inside someone's reply CANNOT false-trip this. That is a feature:
 * the marker means "this draft's own signature", never "somebody once signed something".
 * The mechanism lives in another module, so signature.test.ts pins it end to end.
 */
export function hasSignatureMarker(html: string | null | undefined): boolean {
  if (!html) return false;
  DIV_TAG_OPEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = DIV_TAG_OPEN_RE.exec(html)) !== null) {
    const tag = divTagCarriesMarker(html, m.index + m[0].length);
    if (tag.found) return true;
    DIV_TAG_OPEN_RE.lastIndex = Math.max(tag.next, DIV_TAG_OPEN_RE.lastIndex);
  }
  return false;
}

/**
 * The signature as an html block, marker and all. Undefined when the identity has none.
 *
 * An identity configured with ONLY a text signature is escaped into html rather than
 * skipped: the html body already exists, so this is not the body-model's forbidden
 * "fabricate html from a plain-text message" — it is putting a sign-off into a body that is
 * html either way, and skipping it would silently drop a signature the user really has.
 */
export function signatureHtmlBlock(signature: ResolvedSignature | undefined): string | undefined {
  if (!signature) return undefined;
  const inner = signature.html
    ?? (signature.text !== undefined ? textToHtmlBlock(signature.text) : undefined);
  if (inner === undefined) return undefined;
  return `<div class="${SIGNATURE_CLASS}">${inner}</div>`;
}

/**
 * The signature as plain text. Undefined when the identity has none.
 *
 * `htmlShips` decides WHICH form, and the answer is not "whichever was configured":
 *
 *  - html ships → derive from the html form. The text part of an html message is a derived
 *    fallback (docs/email-bodies.md), regenerated from the html on the first html-only edit.
 *    Writing the configured `textSignature` verbatim here would look right and then change
 *    by itself on that edit, with nothing reporting it. Deriving means the text part reads
 *    the same before and after.
 *  - no html ships → the configured `textSignature`, verbatim. That is the one place the
 *    configured text form is the answer. An identity that configured ONLY an html signature
 *    still gets a text form derived from it — the html body it was written for is not
 *    shipping, but the WORDS in it are the user's sign-off and dropping them would lose a
 *    signature they really have.
 *
 * The derivation on that second branch suppresses image placeholders, and that is
 * load-bearing rather than cosmetic: no html ships, so no image ships either, and writing
 * `[image]` would make the recipient's entire sign-off a description of something no part
 * of the message carries. It is the same rule `buildReplyBodies` applies to quoted images.
 * An images-only html signature therefore yields '' here, which `planSignature` reports as
 * `no-text-form` rather than shipping the placeholder.
 *
 * The FIRST branch deliberately does NOT suppress, and the asymmetry is the point. There the
 * html ships, so the embedded image ships with it, and `[image]` is what
 * `ImagePlaceholderPolicy.unconditional` is for: the text alternative of a body that is about
 * to carry the image. It also has to agree with the derivation downstream — an html-only
 * draft's text fallback is derived from the WHOLE signed html under that same policy, so
 * suppressing here would make a caller who supplies both bodies get a different text sign-off
 * from one who supplies html alone, which is the by-itself drift this function's html/text
 * split exists to prevent. An images-only html signature therefore yields '[image]' here when
 * the image is embedded, and '' when it is remote (a remote image writes no placeholder under
 * any policy — see ImagePlaceholderPolicy). Only the second of those is a missing sign-off,
 * and `planSignature` reports it as `text-part-unsigned`.
 */
export function signatureTextBlock(
  signature: ResolvedSignature | undefined,
  htmlShips: boolean,
): string | undefined {
  if (!signature) return undefined;
  if (htmlShips) {
    const block = signatureHtmlBlock(signature);
    return block === undefined ? undefined : htmlToText(block, 'unconditional');
  }
  return signature.text
    ?? (signature.html !== undefined ? htmlToText(signature.html, 'suppress') : undefined);
}

/**
 * Why an append did not land where it was asked to. Every one of these is a call that ASKED
 * for a signature and did not fully get one, so each is reportable: the caller cannot see the
 * stored body from here, and a flag that silently no-ops is indistinguishable from one that
 * worked.
 *
 *  - `no-signature`        the identity has none configured (or names no verified identity).
 *  - `no-body`             this call writes no body for the sign-off to sit under.
 *  - `already-signed`      the body supplied already carries a sign-off this call would
 *                          otherwise duplicate — the html marker, or a text body carrying one
 *                          of the two block forms among its OWN lines (quoted and forwarded
 *                          history excluded; see textAlreadySigned).
 *  - `no-text-form`        the message ships no html, and the signature has no plain-text
 *                          form to write into the text part (an images-only html signature).
 *  - `text-part-unsigned`  the html body WAS signed, but the signature has no text form for
 *                          the message's plain-text alternative — whether the caller supplied
 *                          that alternative or it is derived from the html downstream. The
 *                          only partial outcome here: a recipient rendering the html sees the
 *                          sign-off, one reading the text alternative does not.
 */
export type SignatureSkipReason =
  | 'no-signature' | 'no-body' | 'already-signed' | 'no-text-form' | 'text-part-unsigned';

/**
 * The one decision function behind both `applySignature` and `signatureSkipReason`, so the
 * bodies that come back and the reason reported for an empty append cannot disagree.
 */
function planSignature(
  bodies: { textBody?: string; htmlBody?: string },
  signature: ResolvedSignature | undefined,
): { textBody?: string; htmlBody?: string; skipped?: SignatureSkipReason } {
  const { textBody, htmlBody } = bodies;
  if (!signature) return { ...bodies, skipped: 'no-signature' };
  // Which form applies follows what the message SHIPS, not what it declares: a
  // present-but-blank htmlBody emits no html part (buildBodyParts drops it), so treating it
  // as an html message would sign a part no recipient ever sees.
  const htmlShips = !isBlank(htmlBody);
  if (htmlShips && hasSignatureMarker(htmlBody)) return { ...bodies, skipped: 'already-signed' };

  // Every call site below has already established the body is non-blank — `htmlShips` for
  // the html, the `isBlank` gates for the text — so this only ever concatenates. It used to
  // fall back to the block alone for a blank body, which is how a present-but-blank textBody
  // came to be REPLACED by the signature instead of counting as no body at all.
  const join = (body: string, block: string, spacer: string) => `${body}${spacer}${block}`;

  const out = { ...bodies };
  if (htmlShips) {
    const block = signatureHtmlBlock(signature);
    if (block === undefined) return { ...bodies, skipped: 'no-signature' };
    out.htmlBody = join(htmlBody!, block, SIGNATURE_SPACER);
    // What the plain-text alternative will carry — asked BEFORE looking at whether the caller
    // supplied that alternative, because the outcome is the same either way. A caller who
    // supplies a text part gets this written into it; a caller who supplies none gets a text
    // part derived downstream from the (now signed) html, under the same `unconditional`
    // policy this is derived under. So when this is blank the recipient reading the text
    // alternative sees no sign-off on EITHER path, and reporting only the first one made the
    // html-only call the single place in this feature where a flag landed nowhere in silence.
    //
    // Blank, not merely defined: an html signature made only of a REMOTE logo derives to '',
    // and joining that would append two newlines and no sign-off to the text part. The html
    // half succeeded, so this is a PARTIAL rather than a no-op, and it is reported as one.
    const text = signatureTextBlock(signature, true);
    if (text === undefined || isBlank(text)) return { ...out, skipped: 'text-part-unsigned' };
    // A blank text part is no text part — buildBodyParts drops it, and the alternative is
    // derived from the signed html instead. Writing the block into it would turn a body the
    // caller left empty into a text alternative whose ENTIRE content is the sign-off, while
    // the html carried the real message. Same rule as `htmlShips` above, read off the other
    // body: what the message SHIPS decides, not what it declares.
    if (!isBlank(textBody) && !textAlreadySigned(textBody, signature)) {
      out.textBody = join(textBody!, text, '\n\n');
    }
    return out;
  }
  // Blank counts as none, matching `htmlShips` on the other body and the wording every
  // surface uses for this reason. Falling through with a blank text body made the signature
  // the WHOLE body — the caller's own (blank) text discarded, a signature-only message
  // stored, and nothing reported, because `join` treats a blank body as "nothing to sit
  // under" rather than as a body it must not replace.
  if (isBlank(textBody)) return { ...bodies, skipped: 'no-body' };
  if (textAlreadySigned(textBody, signature)) return { ...bodies, skipped: 'already-signed' };
  const text = signatureTextBlock(signature, false);
  if (text === undefined || isBlank(text)) return { ...bodies, skipped: 'no-text-form' };
  out.textBody = join(textBody!, text, '\n\n');
  return out;
}

// The forward separator, as a whole-line test: this server's '----- Original message -----'
// (FORWARD_MARKER_LINE) and Gmail's '---------- Forwarded message ----------'. Deliberately
// the same shape hasTextForwardMarker matches, so a body that guard calls forwarded is a body
// this one agrees is forwarded. Anchored at the dashes, so a '> '-quoted separator inside a
// reply quote is quoted content and does not cut the body short — the same answer
// hasTextForwardMarker gives, whose line anchor the '>' likewise defeats.
const FORWARD_SEPARATOR_LINE = /^[ \t]*-{3,}[ \t]*(?:Original|Forwarded) message[ \t]*-{2,}/i;

/**
 * A plain-text body as normalised lines: any line ending → LF, then each line trimmed at BOTH
 * ends. Applied identically to the body and to the signature block, so it can only ever make
 * a match MORE likely — which is the announced-error direction (see textAlreadySigned).
 *
 * Every part of it answers a way a body changes on a round trip through a mail store, and
 * every one of those, left unhandled, produces a SECOND sign-off with nothing reported:
 *
 *  - `\r\n` and a bare `\r` are both line endings here. RFC 5322 says CRLF; a store that
 *    emits bare CR would otherwise leave the whole body as one line and match nothing.
 *  - trailing whitespace goes because `-- `, the RFC 3676 signature delimiter, ENDS IN A
 *    SPACE, and stripping it is an ordinary thing for a store to do. `\s` rather than
 *    `[ \t]` so a space replaced by NBSP (U+00A0) on the trip is covered too — `\s` includes
 *    it, along with the other Unicode spaces.
 *  - LEADING whitespace goes because of RFC 3676 space-stuffing: a format=flowed sender
 *    prepends a space to any line starting with a space, '>' or 'From '. Fastmail does not
 *    emit format=flowed (verified live — see quoteText), but `edit_draft` receives drafts
 *    composed by other clients, and one that does would otherwise be signed twice.
 */
function normalizeTextLines(s: string): string[] {
  return s.replace(/\r\n?/g, '\n').split('\n').map((line) => line.trim());
}

/**
 * The lines of a plain-text body that can be the draft's OWN text: everything above a forward
 * separator. See textAlreadySigned for the design and for why nothing here excludes a
 * reply-quoted line.
 */
function draftOwnTextLines(body: string): string[] {
  const own: string[] = [];
  for (const line of normalizeTextLines(body)) {
    if (FORWARD_SEPARATOR_LINE.test(line)) break;
    own.push(line);
  }
  return own;
}

/**
 * Whether `needle` appears in `hay` as a contiguous run of WHOLE lines.
 *
 * The empty-needle arm is defensive rather than reachable from textAlreadySigned — its
 * `isBlank(block)` gate already excludes a block that would trim to nothing — but an empty
 * needle matching everything is the wrong answer for a general helper, so it is stated here
 * rather than left to the one caller.
 */
function containsLineRun(hay: string[], needle: string[]): boolean {
  if (needle.length === 0 || needle.length > hay.length) return false;
  for (let i = 0; i + needle.length <= hay.length; i++) {
    if (needle.every((line, j) => hay[i + j] === line)) return true;
  }
  return false;
}

/**
 * Whether a plain-text body already carries one of this identity's sign-off blocks in its own
 * text — as opposed to inside quoted or forwarded history it happens to be carrying along.
 *
 * This is the text part's whole idempotence story: the html side is protected by the CLASS
 * marker, which a plain-text body cannot carry, so without this the read-modify-write loop
 * docs/email-bodies.md prescribes as the recovery (read the draft, edit the words, send the
 * body back with appendSignature:true) appends a second copy every time round.
 *
 * The rule is SUBTRACTIVE: cut the body at the forward separator, then look for the block as a
 * run of WHOLE LINES in what is above it (see draftOwnTextLines). The separator is a thing that
 * IS a line — distinctive, machine-emitted, anchored at line start — rather than prose.
 *
 * Why that matters: four earlier versions of this guard tried to locate where history BEGINS
 * by recognising the attribution line above the quote ("On <date>, <name> wrote:"). Attribution
 * lines wrap, localise and vary per client, so each version was a fresh guess about one client
 * with an unbounded supply of others left to be wrong about — a hard-wrapped Gmail attribution
 * is what broke the last one. A separator line needs no parsing: it matches or it does not.
 *
 * A REPLY QUOTE NEEDS NO REMOVAL, and trying to remove one was a bug. The whole-line comparison
 * already keeps quoted content out: a quoted line is '> ' + text, and `'> Regards,'` is not
 * `'Regards,'`, so a thread quoting an earlier signed message of ours cannot read as this
 * draft's own sign-off. Dropping '>'-prefixed lines from the body as well looked like belt and
 * braces and was the opposite, because the block on the other side of the comparison is NOT
 * filtered: `signatureTextBlock` derives the text form through html-to-text, whose blockquote
 * output IS '> '-prefixed, so any identity whose html signature contains a <blockquote> had a
 * needle line that no surviving body line could ever equal — and every pass of the
 * read-modify-write loop appended another sign-off, silently. Removing the filter cannot
 * introduce a wrong match either: a '> '-prefixed BODY line can only equal a '> '-prefixed
 * NEEDLE line, which is the signature's own quoted content and the match we want.
 *
 * THE TWO ERRORS ARE NOT SYMMETRIC, and this is tuned for that. A false positive (read as
 * signed when it is not) appends nothing AND reports `already-signed`, so the caller sees it
 * and can write the sign-off themselves. A false negative (read as unsigned when it is signed)
 * ships two sign-offs and says nothing at all. Only the second is invisible to the caller, so
 * wherever this cannot be certain it must fall on the side of NOT appending. Do not "improve"
 * this into a fuzzy or best-effort match: that trades the announced error for the silent one.
 *
 * BOTH forms are tested, not just the one this call would write. Which form a previous call
 * left there depends on whether THAT call shipped html: an html-shipping call writes the
 * form derived from `htmlSignature`, a text-only call writes the configured `textSignature`,
 * and for the ordinary identity that configures both, those are different strings. So an
 * html draft converted to plain text — `clearFields:['htmlBody']` with the text the draft
 * just handed back — hands this the derived form while the call is about to write the
 * configured one. Testing only the outgoing form stacked a second sign-off there, which is
 * exactly the duplication this function exists to prevent.
 *
 * RESIDUAL: a forwarded block in a shape neither separator names — an Outlook
 * `From:/Sent:/To:/Subject:` header block, say — is not removed, so a forwarded message whose
 * own text carries this identity's sign-off reads as this draft's own and the caller's note
 * goes out bare. That is the CHEAP error by construction: announced as `already-signed`, and
 * recovered by writing the sign-off into the note. Extending separator recognition is tracked
 * as issue #144; the residual is written down in docs/email-bodies.md.
 */
function textAlreadySigned(
  body: string | undefined,
  signature: ResolvedSignature,
): boolean {
  if (body === undefined) return false;
  const hay = draftOwnTextLines(body);
  for (const block of [signatureTextBlock(signature, true), signatureTextBlock(signature, false)]) {
    if (block === undefined || isBlank(block)) continue;
    // Trimmed to the block's non-blank extent. A configured signature that opens or closes
    // with a blank line would otherwise demand a matching blank line in the body on that side,
    // and there need not be one — the append writes the block against whatever the body
    // already ended with, and the OTHER form of the same signature (see above) may carry the
    // blank where this one does not. Interior blank lines are kept: they are part of the run.
    const needle = normalizeTextLines(block);
    while (needle.length > 0 && isBlank(needle[0])) needle.shift();
    while (needle.length > 0 && isBlank(needle[needle.length - 1])) needle.pop();
    if (containsLineRun(hay, needle)) return true;
  }
  return false;
}

/**
 * Append the signature to a caller's bodies, returning new ones. Pure: the identity lookup
 * is the caller's job (the handlers own the client), and this takes the resolved strings.
 *
 * Definedness is preserved exactly — a body that came in undefined goes out undefined — so
 * every downstream `!== undefined` test that decides which formats a message emits reads the
 * same answer with a signature as without one.
 *
 * A body already carrying the signature is returned untouched: an html body is recognised by
 * the marker class, a plain-text one by already carrying EITHER of the identity's two block
 * forms among its own lines — the body with its reply-quoted lines and any forwarded block
 * removed (see textAlreadySigned — which form is there depends on whether the call that put
 * it there shipped html). A caller that supplied a signature gets theirs, not two.
 */
export function applySignature(
  bodies: { textBody?: string; htmlBody?: string },
  signature: ResolvedSignature | undefined,
): { textBody?: string; htmlBody?: string } {
  // `skipped` is destructured off rather than deleted so the returned object keeps exactly
  // the body keys it came in with — every downstream `!== undefined` test depends on that.
  const { skipped, ...out } = planSignature(bodies, signature);
  return out;
}

/**
 * What `applySignature` would decline to do to these bodies, or undefined when it appends
 * something. The orchestrations use this to report a requested signature that landed
 * nowhere; it runs the same plan rather than re-deriving the rules.
 */
export function signatureSkipReason(
  bodies: { textBody?: string; htmlBody?: string },
  signature: ResolvedSignature | undefined,
): SignatureSkipReason | undefined {
  return planSignature(bodies, signature).skipped;
}

/**
 * What edit_draft says when it re-appended a signature nobody asked about in that call.
 *
 * The re-append is preservation of an EARLIER decision, so it is the one signature outcome
 * that happens without being requested in the same breath — and therefore the one the
 * result has to announce.
 */
export const NOTE_SIGNATURE_REAPPENDED =
  "The identity's configured signature was re-appended: this draft already carried one and the body you supplied did not. Pass appendSignature:false to drop it instead.";

/**
 * What a compose call says when appendSignature:true appended nothing.
 *
 * The flag is an input the caller cannot verify without re-reading the draft, so an append
 * that lands nowhere has to be announced — the same rule that makes a dropped enrichment
 * surface its degradation rather than vanish. Each reason names WHICH one it was, because
 * the fix differs: configure a signature, supply a body, or leave the one you already wrote.
 *
 * `text-part-unsigned` is the one arm that does not open with "nothing was appended", because
 * on that arm something was: the html carries the sign-off and only its text alternative
 * does not. Saying "nothing was appended" there would be a plainly false report.
 */
export function noteSignatureNotAppended(
  reason: SignatureSkipReason,
  identityAddress: string | undefined,
): string {
  const who = identityAddress
    ? `the identity this message sends as (${identityAddress})`
    : 'the identity this message sends as';
  // The no-signature arm names an ADDRESS rather than an identity, and says both ways it can
  // fail. `signatureOf(undefined)` is what reaches this arm, and "undefined" covers two
  // different things: a verified identity with the field blank, and an address that is not a
  // verified identity at all. edit_draft reaches the second on a stored `from` it never
  // vetted (only an EXPLICIT from is rejected), where "set one in Fastmail" is both false and
  // unactionable. Same wording as noteSignatureUnavailableOnEdit's default arm, which had it
  // right; this one is not a different fact.
  const whoAddress = identityAddress
    ? `the address this message sends as (${identityAddress})`
    : 'the address this message sends as';
  const head = 'appendSignature was requested but nothing was appended: ';
  switch (reason) {
    case 'no-signature':
      return `${head}no signature is available for ${whoAddress} — it has none configured in Fastmail, or it is not one of your verified identities. Set one there, or write the sign-off into the body yourself.`;
    case 'no-body':
      return `${head}this call wrote no body for the signature to sign (a blank body counts as none). Supply textBody or htmlBody in the call that should carry the sign-off.`;
    case 'already-signed':
      return `${head}the body you supplied already carries a signature — at the end, or above a quoted or forwarded original — so it was left as you wrote it rather than signed twice.`;
    case 'no-text-form':
      return `${head}this message ships no HTML body, and ${who} has no plain-text form of its signature to write into the text part.`;
    case 'text-part-unsigned':
      return `appendSignature was requested and the HTML body was signed, but its plain-text alternative was not: ${who} has an HTML signature with no readable text in it (images only), so there is nothing to write into the text part. A recipient reading the plain-text alternative sees no sign-off.`;
  }
}

/**
 * What edit_draft says when a draft that WAS carrying a signature comes out of the edit
 * without one (or without it everywhere), because the keep could not be honoured.
 *
 * The preserve branch is armed by the stored draft, so this is the caller's OMITTED flag
 * meaning "keep it" and the keep failing. Silence here would drop a block the draft was
 * carrying with nothing said — see the never-silently-drop rule in CLAUDE.md. It is a
 * separate function from `noteSignatureNotAppended` for exactly that reason: nothing in this
 * call asked, so a note opening "appendSignature was requested" would be untrue.
 *
 * The reason is passed in rather than assumed. `no-signature` is not the only way a keep
 * fails: an identity whose signature is images-only html has nothing to put in a draft this
 * edit converts to plain text, and that loses the sign-off just as completely.
 */
export function noteSignatureUnavailableOnEdit(
  fromAddress: string | undefined,
  reason: SignatureSkipReason = 'no-signature',
): string {
  const who = fromAddress ? ` (${fromAddress})` : '';
  const head = "This draft carried the sending identity's signature, but ";
  const tail = ' Write the sign-off into the body if you want it back.';
  switch (reason) {
    case 'text-part-unsigned':
      return `${head}only its HTML body carries it now: the signature configured for the address this edit sends as${who} is images only, so there is no plain-text form to write into the text alternative you supplied. A recipient reading that alternative sees no sign-off.${tail}`;
    case 'no-text-form':
      return `${head}this edit leaves it as a plain-text message and the signature configured for the address this edit sends as${who} is images-only HTML, so there is no plain-text form to carry over — the edited draft carries none.${tail}`;
    default:
      return `${head}no signature is available for the address this edit sends as${who} — it has none configured, or it is not one of your verified identities — so the edited draft carries none.${tail}`;
  }
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
  // Embedded-image reference -> the Content-ID the rebuilt quote should emit for it. The
  // EDIT path's channel: it resolved the references against the draft's own surviving parts
  // and hands the finished map in, so this builder only rewrites.
  cidMap?: Map<string, string>;
  // The COMPOSE path's channel: no draft exists yet, so the builder runs both passes itself
  // and reports what it decided on `quoteImages` in the result.
  quoteImages?: QuoteImageInput;
  // The sending identity's sign-off, already resolved by the caller (this builder does no
  // I/O). Appended to the caller's own body BEFORE anything below runs, which is what puts
  // it above the quote — and what gets it onto the two passthrough branches too.
  signature?: ResolvedSignature;
}): { textBody?: string; htmlBody?: string; quoteImages?: QuoteImageOutcome } {
  const {
    original, textBody: callerText, htmlBody: callerHtml, quoteOriginal, timezone, cidMap,
    quoteImages, signature,
  } = input;
  const { textBody, htmlBody } = applySignature(
    { textBody: callerText, htmlBody: callerHtml },
    signature,
  );

  // Return only the formats the caller supplied (createDraft adds the text fallback later).
  // The outcome rides along only for a caller that asked this builder to resolve images, so
  // the result shape is unchanged for every other caller.
  const passthrough = (images: QuoteImageOutcome = emptyQuoteImages()) => ({
    ...(textBody !== undefined && { textBody }),
    ...(htmlBody !== undefined && { htmlBody }),
    ...(quoteImages && { quoteImages: images }),
  });

  // quoteOriginal:false drops the whole quote, and with it every image the quote would have
  // carried. There is no quote-text-without-images setting: the images ARE the quoted body.
  if (!quoteOriginal) return passthrough();

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // PASS 1 — collect. Decides quotable content (content-based, not raw presence) and, for a
  // compose path, which references resolve. Nothing is minted here.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
  const textQuotable = !isBlank(origText);

  const carriedNothing = (): QuoteImageOutcome => ({
    minted: [],
    mappings: [],
    resolvedParts: collected.resolvedParts,
    unresolvedRefs: collected.unresolvedRefs,
    droppedDataImages: collected.droppedDataImages,
    // No html quote ships on this branch, so nothing was rewritten and nothing dropped.
    droppedUnsupportedImages: 0,
    htmlQuoteShips: false,
  });

  // No quotable original (attachment-only / ICS-only, or an image-only original whose images
  // cannot be carried): skip the quote AND the attribution — no orphan "On … wrote:" over an
  // empty quote. Whatever the references pointed at is reported as dropped by the caller.
  if (!htmlQuotable && !textQuotable) return passthrough(carriedNothing());

  // Attribution in LOCAL time; the date is omitted (never "Invalid Date") when the original
  // has no usable sentAt/receivedAt, and the line drops the leading "On " + comma in that case.
  const senderRaw = original?.from?.[0]?.name || original?.from?.[0]?.email || '';
  const name = normalizeName(senderRaw);
  const date = formatReplyDate(original?.sentAt ?? original?.receivedAt, timezone);
  const attribution = date ? `On ${date}, ${name} wrote:` : `${name} wrote:`;

  const out: { textBody?: string; htmlBody?: string; quoteImages?: QuoteImageOutcome } = {};

  // Whether a quote reproducing the original's own html ships. Two things turn on it: only
  // such a quote can display an embedded image, and the text side's image policy follows —
  // when one ships, an image the quote carries is something the reader can look at, so the
  // text alternative may say an image is there; when none does, a placeholder would describe
  // an absent thing.
  const htmlQuoteShips = htmlBody !== undefined && htmlQuotable;

  // PASS 2 — map. Runs ONLY on a branch that ships an html quote, so an identifier is minted
  // only when a body exists to reference it.
  const resolved = htmlQuoteShips && quoteImages
    ? buildCidMap({
        refs: collected.refs,
        sourceParts: quoteImages.sourceParts ?? [],
        ...(quoteImages.mint && { mint: quoteImages.mint }),
      })
    : null;
  const quoteMap = resolved ? resolved.cidMap : cidMap;
  const mapped = htmlQuotable && quoteMap
    ? sanitizeQuoteHtml(origHtml, { mode: 'map', cidMap: quoteMap })
    : null;
  const sanitizedHtml = htmlQuotable ? (mapped ? mapped.html : collected.html) : '';

  if (textBody !== undefined) {
    // text quote source: the original's text, else a readable conversion of its html. The
    // conversion runs over the RAW original html, exactly as it always has — deriving from
    // the sanitized output instead would silently change the derived text of every html
    // original, embedded images or not, because the quote floor drops tags that carry text.
    const textSource = pick(
      origText,
      htmlToText(origHtml, htmlQuoteShips ? 'resolve' : 'suppress', quoteMap),
    );
    // A blank source gets no attribution and no quote. An "On … wrote:" over an empty "> "
    // line describes a quote that is not there, and it would arm this server's own text
    // quote marker, so the next edit of the draft would be challenged over a quote it has
    // never had. (The forward builder has always had this gate; the reply builder gains it.)
    out.textBody = isBlank(textSource)
      ? textBody
      : `${textBody ?? ''}\n\n${attribution}\n${quoteText(textSource)}`;
  }

  if (htmlBody !== undefined) {
    // rich quote: prefer the sanitized html; else a text-only original → escaped block.
    const htmlSource = htmlQuotable ? sanitizedHtml : textToHtmlBlock(origText);
    out.htmlBody = `${htmlBody ?? ''}<div><br></div><div>${escapeHtml(attribution)}</div>${QUOTE_OPEN}${htmlSource}</blockquote>`;
  }

  // Read off the pass that produced the html actually shipped: only that pass drops a
  // reference form it cannot carry, and only when its output is what ships.
  const droppedUnsupportedImages = htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0;

  // Reported when the caller asked about the original's images, and ALSO whenever there is a
  // loss only this pass can see. The edit path supplies a Content-ID map rather than parts —
  // it resolved its own images already — so it asks for no outcome; without this second arm
  // its quote could drop an image and say nothing.
  if (quoteImages || droppedUnsupportedImages > 0) {
    out.quoteImages = {
      minted: resolved?.minted ?? [],
      mappings: resolved?.mappings ?? [],
      resolvedParts: collected.resolvedParts,
      unresolvedRefs: collected.unresolvedRefs,
      droppedDataImages: collected.droppedDataImages,
      droppedUnsupportedImages,
      htmlQuoteShips,
    };
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
// so the two marker families are disjoint by tag name, and the quote sanitizer
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
  // See buildReplyBodies: the edit path's rewrite-only channel.
  cidMap?: Map<string, string>;
  // See buildReplyBodies: the compose path's channel, where this builder runs both passes.
  quoteImages?: QuoteImageInput;
  // See buildReplyBodies: the resolved sign-off, appended to the caller's note so it lands
  // above the forwarded-message block rather than under the message being forwarded.
  signature?: ResolvedSignature;
}): {
  textBody?: string; htmlBody?: string; quoteImages?: QuoteImageOutcome;
  /** Why a requested signature landed nowhere; see BuiltForward.signatureSkip. */
  signatureSkip?: SignatureSkipReason;
} {
  const { original, textBody: callerText, htmlBody: callerHtml, cidMap, quoteImages, signature } = input;
  const noNote = callerText === undefined && callerHtml === undefined;
  // A noted forward signs the note; a note-less one has the signature become the whole
  // content above the block, on the two arms at the bottom of this function. The skip for
  // that second shape is decided down there, where those arms are, so this asks about the
  // note only when there is a note.
  let signatureSkip = noNote
    ? undefined
    : signatureSkipReason({ textBody: callerText, htmlBody: callerHtml }, signature);
  const { textBody, htmlBody } = applySignature(
    { textBody: callerText, htmlBody: callerHtml },
    signature,
  );

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // PASS 1 — collect (see buildReplyBodies). An original whose body is nothing but embedded
  // images becomes quotable here, which is also what flips this tool's no-caller-body default
  // from a text forward to an html one for such a message.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
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
  const htmlQuoteShips = htmlQuotable && (htmlBody !== undefined || textBody === undefined);

  // PASS 2 — map, on a branch that ships the original's html and nowhere else.
  const resolved = htmlQuoteShips && quoteImages
    ? buildCidMap({
        refs: collected.refs,
        sourceParts: quoteImages.sourceParts ?? [],
        ...(quoteImages.mint && { mint: quoteImages.mint }),
      })
    : null;
  const quoteMap = resolved ? resolved.cidMap : cidMap;
  const mapped = htmlQuotable && quoteMap
    ? sanitizeQuoteHtml(origHtml, { mode: 'map', cidMap: quoteMap })
    : null;
  const sanitizedHtml = htmlQuotable ? (mapped ? mapped.html : collected.html) : '';

  const textSource = pick(
    origText,
    htmlToText(origHtml, htmlQuoteShips ? 'resolve' : 'suppress', quoteMap),
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
  const out: {
    textBody?: string; htmlBody?: string; quoteImages?: QuoteImageOutcome;
    signatureSkip?: SignatureSkipReason;
  } = {};
  if (htmlBody !== undefined) out.htmlBody = composeHtml(htmlBody);
  if (textBody !== undefined) out.textBody = composeText(textBody);
  if (htmlBody === undefined && textBody === undefined) {
    // No note to append to, but a requested signature still ships: it becomes the whole of
    // the content above the forwarded-message block. (Both arms pass undefined when no
    // signature was asked for, which is exactly the bare-FYI forward this has always made.)
    const block = htmlQuotable ? signatureHtmlBlock(signature) : signatureTextBlock(signature, false);
    if (htmlQuotable) out.htmlBody = composeHtml(block);
    else out.textBody = composeText(block);
    // Read off the block this shape actually produced, so the reported reason follows the
    // arm taken rather than a rule restated here. An html-quotable original wants the html
    // block, a text one the derived text form, and either can come back empty.
    if (signature === undefined) signatureSkip = 'no-signature';
    else if (block === undefined || isBlank(block)) signatureSkip = htmlQuotable ? 'no-signature' : 'no-text-form';
  }
  if (signatureSkip) out.signatureSkip = signatureSkip;
  // Only the rewriting pass drops a reference form it cannot carry, and only its output
  // ships an html forwarded block. Reported on the same two arms as the reply builder's —
  // see the comment there for why a loss is reported even when nothing asked.
  const droppedUnsupportedImages = htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0;
  if (quoteImages || droppedUnsupportedImages > 0) {
    out.quoteImages = {
      minted: resolved?.minted ?? [],
      mappings: resolved?.mappings ?? [],
      resolvedParts: collected.resolvedParts,
      unresolvedRefs: collected.unresolvedRefs,
      droppedDataImages: collected.droppedDataImages,
      droppedUnsupportedImages,
      htmlQuoteShips,
    };
  }
  return out;
}

// True if html carries a forwarded-message wrapper: a <div type="cite"> — the
// canonical Fastmail forward wrapper, which buildForwardBodies also emits.
// Attribute-keyed ONLY, never text-keyed: the quote sanitizer strips type= from
// embedded content, so a reply quoting a forward (or pasted sanitized forward
// HTML) loses the attribute and cannot false-trip this. Disjoint from
// hasQuoteMarker by tag name (div vs blockquote) — the official Fastmail client
// uses <blockquote type="cite"> for replies and <div type="cite"> only for
// forwards (probed live 2026-07-05). Like hasQuoteMarker, a PRESENCE check that
// governs whether edit_draft's guard fires and — on the keep path — whether a
// caller-supplied body is refused for already carrying the block (#145, refuse
// only, never exempt); originalEmailId is the authoritative keep path.
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
// over-asks loudly and noQuote resolves it in one step — which is also the
// answer when this fires on a caller-supplied keep body (#145). Linear-time: every
// quantifier consumes from classes disjoint from its neighbors.
export function hasTextForwardMarker(text: string | null | undefined): boolean {
  if (!text) return false;
  return /(^|\n)[ \t]*-{3,}[ \t]*(Original|Forwarded) message[ \t]*-{2,}/i.test(text);
}
