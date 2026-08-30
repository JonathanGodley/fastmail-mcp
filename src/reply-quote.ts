import { htmlToText, isBlank } from './body-format.js';
import { formatAddress, formatReplyDate } from './email-formatter.js';
import { buildCidMap, resolveCidRefs, sanitizeQuoteHtml } from './inline-images.js';
import type { CidMapping, CidPart, MintedInlinePart } from './inline-images.js';
import type { ResolvedSignature } from './identity.js';
import type { BodyBlock } from './body-tokens.js';

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

// Both BLOCK builders — buildQuoteBlocks and buildForwardBlocks, which is where the two
// passes below run — sanitize the original's html through the shared two-pass sanitizer in
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

// ---------------------------------------------------------------------------
// The sending identity's signature (#33)
// ---------------------------------------------------------------------------
//
// A block, beside the reply quote and the forwarded-message block above, and — unlike those
// two — one whose content comes from the account rather than from another message. It lives
// in this file because the same three builders feed one expansion: `draft_email` and
// `edit_draft` both substitute `{{signature}}`, and the rule for WHICH form a part gets is
// one rule (see signatureTextBlock).
//
// The block carries NO marker class. It used to: a `fm-mcp-signature` class let a later edit
// recognise a sign-off this server had written, which is what the automatic append needed to
// avoid signing a body twice. Nothing appends automatically any more — a signature lands
// where the caller wrote `{{signature}}` and nowhere else — so there is nothing for a marker
// to protect, and writing an identifying class into every signed body bought a reader
// nothing while claiming the block was ours to manage.

/**
 * The signature as an html block. Undefined when the identity has none.
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
  return `<div>${inner}</div>`;
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
 * of the message carries. It is the same rule `buildQuoteBlocks` applies to quoted images.
 * An images-only html signature therefore yields '' here, which `signatureBlock` reports as
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
 * and `signatureBlock` reports it as `no-text-form` on the part that could not hold it.
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
 * The signature as a `{{signature}}` block for one part, carrying the cause when the identity
 * offers nothing this part can hold.
 *
 * Here rather than in either handler because BOTH tools expand the token — `draft_email` on
 * the body it composes, `edit_draft` on a flagged edit — and the rule for which form a part
 * gets is one rule. `messageShipsHtml` is the MESSAGE question (a non-blank html part
 * exists), never "does this part carry a token"; see signatureTextBlock for why the answer
 * is not simply "whichever form was configured".
 */
export function signatureBlock(
  signature: ResolvedSignature | undefined,
  part: 'textBody' | 'htmlBody',
  messageShipsHtml: boolean,
): BodyBlock {
  if (!signature) return { available: false, cause: 'no-signature' };
  const content = part === 'htmlBody'
    ? signatureHtmlBlock(signature)
    : signatureTextBlock(signature, messageShipsHtml);
  if (content === undefined || isBlank(content)) {
    // An identity that has a signature but no form this part can carry (an images-only html
    // signature, in a text part) is a different sentence from having none at all.
    return { available: false, cause: part === 'htmlBody' ? 'no-signature' : 'no-text-form' };
  }
  return { available: true, content };
}

/**
 * The attributed reply quote, one block per body format.
 *
 * A BLOCK STARTS AT ITS ATTRIBUTION LINE. The separator between the caller's own body and the
 * block — `\n\n` in text, `<div><br></div>` in html — is NOT part of it. The separator belongs
 * to the concatenation and is deleted with it, so a caller that places a block somewhere else
 * in a body it wrote itself is not handed our spacing along with the block.
 */
export interface QuoteBlocks {
  /** The attributed text quote. Undefined when there is nothing to put in one. */
  textBlock?: string;
  /** The attributed html quote. Undefined when the original is quotable in no format at all. */
  htmlBlock?: string;
  /** What the two passes decided about the original's embedded images. */
  images: QuoteImageOutcome;
}

/**
 * Build the reply quote's blocks, running both image passes (see the pass-ordering note at the
 * top of this file, which is what this function's PASS 1 / PASS 2 comments refer to).
 *
 * `htmlShips` is an INPUT rather than a decision made here: it is a fact about the message
 * being composed, not about the original, and the two callers do not answer it the same way —
 * a forward with no caller body ships html where a reply would not.
 */
export function buildQuoteBlocks(input: {
  original: any;            // raw JMAP email from getEmailById (textBody/htmlBody arrays + bodyValues + date)
  /** Whether the message being composed ships an html part at all. */
  htmlShips: boolean;
  timezone?: string;
  // See QuoteImageInput: the edit path's rewrite-only channel.
  cidMap?: Map<string, string>;
  // See QuoteImageInput: the compose path's channel, where this builder runs both passes.
  quoteImages?: QuoteImageInput;
}): QuoteBlocks {
  const { original, htmlShips, timezone, cidMap, quoteImages } = input;

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // PASS 1 — collect. Decides quotable content (content-based, not raw presence) and, for a
  // compose path, which references resolve. Nothing is minted here.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
  const textQuotable = !isBlank(origText);

  // Whether a quote reproducing the original's own html ships. Two things turn on it: only
  // such a quote can display an embedded image, and the text side's image policy follows —
  // when one ships, an image the quote carries is something the reader can look at, so the
  // text alternative may say an image is there; when none does, a placeholder would describe
  // an absent thing.
  const htmlQuoteShips = htmlShips && htmlQuotable;

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

  const images: QuoteImageOutcome = {
    minted: resolved?.minted ?? [],
    mappings: resolved?.mappings ?? [],
    resolvedParts: collected.resolvedParts,
    unresolvedRefs: collected.unresolvedRefs,
    droppedDataImages: collected.droppedDataImages,
    // Read off the pass that produced the html actually shipped: only that pass drops a
    // reference form it cannot carry, and only when its output is what ships.
    droppedUnsupportedImages: htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0,
    htmlQuoteShips,
  };

  // No quotable original (attachment-only / ICS-only, or an image-only original whose images
  // cannot be carried): no block at all, in either format — the attribution goes with the
  // quote, so there is no orphan "On … wrote:" over an empty one. Whatever the references
  // pointed at is reported as dropped by the caller.
  if (!htmlQuotable && !textQuotable) return { images };

  // Attribution in LOCAL time; the date is omitted (never "Invalid Date") when the original
  // has no usable sentAt/receivedAt, and the line drops the leading "On " + comma in that case.
  const senderRaw = original?.from?.[0]?.name || original?.from?.[0]?.email || '';
  const name = normalizeName(senderRaw);
  const date = formatReplyDate(original?.sentAt ?? original?.receivedAt, timezone);
  const attribution = date ? `On ${date}, ${name} wrote:` : `${name} wrote:`;

  const blocks: QuoteBlocks = { images };

  // text quote source: the original's text, else a readable conversion of its html. The
  // conversion runs over the RAW original html, exactly as it always has — deriving from
  // the sanitized output instead would silently change the derived text of every html
  // original, embedded images or not, because the quote floor drops tags that carry text.
  const textSource = pick(
    origText,
    htmlToText(origHtml, htmlQuoteShips ? 'resolve' : 'suppress', quoteMap),
  );
  // A blank source gets no attribution and no quote. An "On … wrote:" over an empty "> "
  // line describes a quote that is not there: the draft opens by telling its reader that
  // the original follows, and nothing does. It is also text the caller then has to delete
  // by hand on every edit, since an edit stores the body it is handed and this server
  // recognises nothing in it. (The forward builder has always had this gate; the reply
  // builder gains it.)
  if (!isBlank(textSource)) blocks.textBlock = `${attribution}\n${quoteText(textSource)}`;

  // rich quote: prefer the sanitized html; else a text-only original → escaped block.
  const htmlSource = htmlQuotable ? sanitizedHtml : textToHtmlBlock(origText);
  blocks.htmlBlock = `<div>${escapeHtml(attribution)}</div>${QUOTE_OPEN}${htmlSource}</blockquote>`;

  return blocks;
}

// ---------------------------------------------------------------------------
// Forward support (draft_email's mode:'forward' + edit_draft's forward guard)
// ---------------------------------------------------------------------------

// The forwarded-message block matches the canonical Fastmail shape (probed live
// 2026-07-05 against the official Fastmail client's forward): a dashed marker
// line, then From/To/Cc/Subject/Date header lines (Cc only when present), then
// the original below a blank line. The HTML wrapper is the platform's own
// <div type="cite">, where a quoted reply is <blockquote>-anchored, so the two
// block shapes stay distinguishable by tag for a human reading the source.
// Nothing in this server reads either shape back: the predicates that used to
// detect an already-present forward or quote block are gone with the automatic
// placement they served — a block lands where the caller wrote its token, so
// there is no "did we already put one here?" question left to answer.
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

/**
 * The forwarded-message block, one per body format.
 *
 * Same rule as QuoteBlocks: A BLOCK STARTS AT ITS HEADER LINE, and the separator between the
 * caller's note and the block belongs to the concatenation — `\n\n\n` in text, and in html
 * nothing at all.
 *
 * The html join is empty because the block used to OPEN with a `<br>`, sitting inside the
 * block's own div, separating it from the caller's note. Carving the block out DROPS that
 * `<br>`; it is not moved to the join. A block has to read the same wherever a caller places
 * it, and a leading blank line is only ever right directly under a note — under a `<div>`-per-
 * paragraph body it is a stray gap, and at the top of a body it opens the message with one. So
 * a forwarded message's stored html no longer opens with it: a deliberate behaviour change,
 * pinned by a test.
 *
 * Both blocks are always built. Unlike a reply quote there is no "nothing to quote" case: the
 * header block IS the forward's content, and it stands alone over an attachment-only original.
 */
export interface ForwardBlocks {
  /** The header block, with the original's text below it when there is any. */
  textBlock: string;
  /** The header block, with the original's html (or its escaped text) cited below it. */
  htmlBlock: string;
  /**
   * Whether the original's html is worth reproducing. Exposed because the forward builder's
   * no-note arm reads it to choose which format a bare forward emits.
   */
  htmlQuotable: boolean;
  /** What the two passes decided about the original's embedded images. */
  images: QuoteImageOutcome;
}

/**
 * Build the forwarded-message blocks, running both image passes (see the pass-ordering note at
 * the top of this file). `htmlShips` is an input for the same reason it is on buildQuoteBlocks
 * — and here the answer is not the reply's: a forward with no caller body at all still ships
 * html when the original has quotable html.
 */
export function buildForwardBlocks(input: {
  original: any;      // raw JMAP email from getEmailById (body lists + bodyValues + addresses)
  /** Whether the message being composed ships an html part at all. */
  htmlShips: boolean;
  // See QuoteImageInput: the edit path's rewrite-only channel.
  cidMap?: Map<string, string>;
  // See QuoteImageInput: the compose path's channel, where this builder runs both passes.
  quoteImages?: QuoteImageInput;
}): ForwardBlocks {
  const { original, htmlShips, cidMap, quoteImages } = input;

  const bodyValues = original?.bodyValues || {};
  const origText = readBodyList(original?.textBody, bodyValues, 'text/plain', '\n[…]');
  const origHtml = readBodyList(original?.htmlBody, bodyValues, 'text/html', '<div>[…]</div>');

  // PASS 1 — collect (see buildQuoteBlocks). An original whose body is nothing but embedded
  // images becomes quotable here, which is also what flips this tool's no-caller-body default
  // from a text forward to an html one for such a message.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
  const textQuotable = !isBlank(origText);

  const lines = forwardHeaderLines(original);
  const headerText = lines.join('\n');
  // No leading <br>: it was the separator under the caller's note, sitting inside the block's
  // own div, and carving the block out drops it rather than moving it to the join — see
  // ForwardBlocks for why a block cannot carry its own leading blank line.
  const headerHtml = `<div>${lines.map(escapeHtml).join('<br>')}<br></div>`;

  // Content below the block. Text form: the original's own text, else a
  // readable conversion of its html; may be blank (attachment-only original),
  // in which case the block stands alone. HTML form: the sanitized original
  // html, else the original text escaped — never fabricated from nothing.
  // The text side's image policy needs the html half of the emission decision up front, for
  // the same reason the reply builder does — a placeholder must not describe an image no
  // format carries.
  const htmlQuoteShips = htmlQuotable && htmlShips;

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
  const below = !isBlank(textSource) ? `\n\n${textSource}` : '';
  const cite = htmlSource ? `${FORWARD_OPEN}${htmlSource}</div>` : '';

  return {
    textBlock: `${headerText}${below}`,
    htmlBlock: `${headerHtml}${cite}`,
    htmlQuotable,
    images: {
      minted: resolved?.minted ?? [],
      mappings: resolved?.mappings ?? [],
      resolvedParts: collected.resolvedParts,
      unresolvedRefs: collected.unresolvedRefs,
      droppedDataImages: collected.droppedDataImages,
      // Only the rewriting pass drops a reference form it cannot carry, and only its output
      // ships an html forwarded block.
      droppedUnsupportedImages: htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0,
      htmlQuoteShips,
    },
  };
}
