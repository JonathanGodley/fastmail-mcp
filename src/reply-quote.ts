import { htmlToText, isBlank } from './body-format.js';
import { formatAddress, formatReplyDate } from './email-formatter.js';
import { buildCidMap, describePart, resolveCidRefs, sanitizeQuoteHtml } from './inline-images.js';
import type { CidMapping, CidPart, MintedInlinePart } from './inline-images.js';
import type { ResolvedSignature } from './identity.js';
import type { BodyBlock } from './body-tokens.js';

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Collapse whitespace runs so a display name with a newline can't split the attribution
// line. ECMAScript \s does NOT cover U+0085 (NEL), a mandatory line break per UAX #14, so
// don't "simplify" back to \s.
function normalizeName(s: string): string {
  return s.replace(/[\s\u0085]+/g, ' ').trim();
}

// Defensive: the raw reader below adds no sentinels, but an upstream value might carry one.
function stripSentinels(s: string): string {
  return s.replace(/\n?\[body truncated\]/g, '').replace(/\n?\[encoding issues detected\]/g, '');
}

// Both block builders run the original's html through the two-pass sanitizer in
// src/inline-images.ts (the posture is in docs/conventions.md). `collect` reports references
// and rewrites nothing; `map` rewrites resolved references to the Content-IDs this draft
// attaches. The order matters: minting an identifier commits the call to attaching a part,
// so pass one decides whether an html quote ships and pass two runs only when one does.

// Content-based, NOT a string trim: an embedded-image-only original collects to e.g.
// <div></div>, which must not count or an orphan "On … wrote:" ships over an empty quote
// (it becomes quotable through the resolvability test instead). Placeholders are suppressed
// because an unmapped image has already been dropped from sanitized html.
function isQuotable(sanitized: string): boolean {
  if (!isBlank(htmlToText(sanitized, 'suppress'))) return true;
  return /<img\b[^>]*\bsrc\s*=/i.test(sanitized);
}

/**
 * What a compose path gives a quote builder so it can resolve the original's embedded images
 * itself. Absent on the edit path, which passes a finished `cidMap` instead, so the builder
 * only rewrites and never mints.
 */
export interface QuoteImageInput {
  /** The original's parts (the gated union). Their Content-IDs are compared literally. */
  sourceParts: CidPart[];
  /** Injected so callers' tests are deterministic. */
  mint?: () => string;
}

/** What the builder decided about the original's embedded images. */
export interface QuoteImageOutcome {
  /** Empty whenever no html quote ships. */
  minted: MintedInlinePart[];
  mappings: CidMapping[];
  /** Distinct parts the references resolved to, in first-reference order. */
  resolvedParts: CidPart[];
  /** Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  droppedDataImages: number;
  /**
   * Images whose src was neither a cid reference nor http(s). Non-zero only on a branch that
   * actually rewrote the quote's html.
   */
  droppedUnsupportedImages: number;
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
 * `quotable` tests EMBEDDABILITY rather than mere resolution: a reference whose Content-ID
 * names two parts, or whose part has no blob, cannot be carried, and counting it would put an
 * attribution over a quote showing nothing.
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

function textToHtmlBlock(s: string): string {
  return escapeHtml(s).replace(/\n/g, '<br>');
}

// Fastmail does not emit format=flowed (verified live 2026-06-24), so uniform "> " is correct.
function quoteText(s: string): string {
  return s.split('\n').map((l) => '> ' + l).join('\n');
}

// Trim-based pick: an empty-but-present '' must fall through to the fallback (?? would not).
function pick(a: string | null | undefined, b: string | null | undefined): string {
  return a && a.trim() ? a : (b ?? '');
}

// Accepts an untyped part, matching extractBody: strict equality would drop a typeless part
// the user just saw.
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
// Here beside the quote and forward builders because `draft_email` and `edit_draft` both
// expand these three tokens, and the rule for which form a part gets is one rule.
//
// The block carries NO marker class, deliberately: a signature lands only where the caller
// wrote `{{signature}}`, so there is nothing for one to protect. The removed class,
// `fm-mcp-signature`, is named here and in docs/email-bodies.md as the record of that; a
// sweep for its last occurrences should leave both.

/**
 * The signature as an html block. Undefined when the identity has none. A text-only
 * signature is escaped into html rather than skipped: the body is html either way, so this
 * is not fabricating html from a plain-text message.
 */
export function signatureHtmlBlock(signature: ResolvedSignature | undefined): string | undefined {
  if (!signature) return undefined;
  const inner = signature.html
    ?? (signature.text !== undefined ? textToHtmlBlock(signature.text) : undefined);
  if (inner === undefined) return undefined;
  return `<div>${inner}</div>`;
}

/**
 * The embedded-image (cid:) references the html signature makes. An identity's signature is
 * a string, so no part carries the image any of these names.
 */
export function signatureCidRefs(signature: ResolvedSignature | undefined): string[] {
  if (signature?.html === undefined) return [];
  return sanitizeQuoteHtml(signature.html, { mode: 'collect' }).refs;
}

export function rejectSignatureEmbeddedImage(ref: string): string {
  return (
    `The sending identity's signature displays an embedded image "${describePart(ref)}", ` +
    'and nothing in this call supplies it: the identity holds the signature\'s html but not ' +
    'the image. Write the sign-off into htmlBody yourself in place of {{signature}}, or remove ' +
    'the embedded image from the identity\'s signature in Fastmail\'s settings.'
  );
}

/**
 * The signature as plain text. Undefined when the identity has none.
 *
 * `htmlShips` decides WHICH form, and the answer is not "whichever was configured":
 *
 *  - html ships: derive from the html form, under the `unconditional` image policy. The text
 *    part is a derived fallback regenerated on the first html-only edit, so a verbatim
 *    `textSignature` would change by itself then; and the policy must match that downstream
 *    derivation, or supplying both bodies and supplying html alone would sign differently.
 *  - no html ships: the configured `textSignature`, else a derivation from the html form with
 *    placeholders SUPPRESSED, since no image ships either. An images-only html signature then
 *    yields '', which `signatureBlock` reports as `no-text-form`.
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
 * offers nothing this part can hold. `messageShipsHtml` is the MESSAGE question (a non-blank
 * html part exists), never "does this part carry a token".
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
    // A signature with no form this part can carry is a different sentence from having none.
    return { available: false, cause: part === 'htmlBody' ? 'no-signature' : 'no-text-form' };
  }
  return { available: true, content };
}

/**
 * The attributed reply quote, one block per body format.
 *
 * A BLOCK STARTS AT ITS ATTRIBUTION LINE: no separator from the caller's body is part of it,
 * so a caller placing the block is not handed our spacing along with it.
 */
export interface QuoteBlocks {
  /** Undefined when there is nothing to put in one. */
  textBlock?: string;
  /** Undefined when the original is quotable in no format at all. */
  htmlBlock?: string;
  images: QuoteImageOutcome;
}

/**
 * Build the reply quote's blocks, running both image passes (the pass-ordering note is at the
 * top of this file).
 *
 * `htmlShips` is an INPUT because it is a fact about the message being composed, and the two
 * callers answer it differently.
 */
export function buildQuoteBlocks(input: {
  original: any;            // raw JMAP email from getEmailById (textBody/htmlBody arrays + bodyValues + date)
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

  // PASS 1: collect. Nothing is minted here.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
  const textQuotable = !isBlank(origText);

  // The text side's image policy follows this: without an html quote, a placeholder would
  // describe an absent image.
  const htmlQuoteShips = htmlShips && htmlQuotable;

  // PASS 2: map, ONLY when an html quote ships.
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
    // Only the rewriting pass drops a reference form it cannot carry, and only its output ships.
    droppedUnsupportedImages: htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0,
    htmlQuoteShips,
  };

  // No block in either format, so no orphan "On … wrote:" over an empty quote.
  if (!htmlQuotable && !textQuotable) return { images };

  const senderRaw = original?.from?.[0]?.name || original?.from?.[0]?.email || '';
  const name = normalizeName(senderRaw);
  const date = formatReplyDate(original?.sentAt ?? original?.receivedAt, timezone);
  const attribution = date ? `On ${date}, ${name} wrote:` : `${name} wrote:`;

  const blocks: QuoteBlocks = { images };

  // Converts the RAW original html, not the sanitized output: the quote floor drops tags that
  // carry text.
  const textSource = pick(
    origText,
    htmlToText(origHtml, htmlQuoteShips ? 'resolve' : 'suppress', quoteMap),
  );
  // No "On … wrote:" over an empty "> " line.
  if (!isBlank(textSource)) blocks.textBlock = `${attribution}\n${quoteText(textSource)}`;

  const htmlSource = htmlQuotable ? sanitizedHtml : textToHtmlBlock(origText);
  blocks.htmlBlock = `<div>${escapeHtml(attribution)}</div>${QUOTE_OPEN}${htmlSource}</blockquote>`;

  return blocks;
}

// ---------------------------------------------------------------------------
// Forward support (draft_email's mode:'forward')
// ---------------------------------------------------------------------------

// Matches the Fastmail client's own forward block (probed live 2026-07-05), including its
// <div type="cite"> wrapper where a reply quote uses <blockquote>. Nothing in this server
// reads either shape back.
const FORWARD_MARKER_LINE = '----- Original message -----';
const FORWARD_OPEN = '<div type="cite">';

// Unescaped; the HTML form escapes each line. Every field is attacker-controlled content
// re-sent under the user's identity, so normalizeName covers the WHOLE composed address, not
// just the display name. A field with no usable value drops its whole line.
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
  // Verbatim ISO 8601, as the platform's own forward block writes it, deliberately not the
  // humanized formatReplyDate shape.
  const date = normalizeName(original?.sentAt ?? original?.receivedAt ?? '');
  if (date) lines.push(`Date: ${date}`);
  return lines;
}

/**
 * The forwarded-message block, one per body format.
 *
 * Same rule as QuoteBlocks: A BLOCK STARTS AT ITS HEADER LINE. The block opens with no `<br>`,
 * deliberately (pinned by a test): it has to read the same wherever a caller places it, and a
 * leading blank line is right only directly under a note.
 *
 * Both blocks are always built: the header block IS the forward's content, and it stands
 * alone over an attachment-only original.
 */
export interface ForwardBlocks {
  textBlock: string;
  htmlBlock: string;
  /**
   * Whether the original has html worth reproducing. draft_email reads it to tell a caller
   * whose {{forward}} ships only in the text form that the formatting was lost.
   */
  htmlQuotable: boolean;
  images: QuoteImageOutcome;
}

/**
 * Build the forwarded-message blocks, running both image passes (the pass-ordering note is at
 * the top of this file). Each block is substituted where the caller placed {{forward}}, so
 * html ships only in a caller-supplied htmlBody.
 */
export function buildForwardBlocks(input: {
  original: any;      // raw JMAP email from getEmailById (body lists + bodyValues + addresses)
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

  // PASS 1: collect. An image-only original becomes quotable here, so an html block over it
  // shows the picture rather than the header block alone.
  const collected = collectQuoteRefs(origHtml, quoteImages, cidMap);
  const htmlQuotable = collected.quotable;
  const textQuotable = !isBlank(origText);

  const lines = forwardHeaderLines(original);
  const headerText = lines.join('\n');
  // No leading <br>: see ForwardBlocks.
  const headerHtml = `<div>${lines.map(escapeHtml).join('<br>')}<br></div>`;

  const htmlQuoteShips = htmlQuotable && htmlShips;

  // PASS 2: map, only when the original's html ships.
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
      droppedUnsupportedImages: htmlQuoteShips && mapped ? mapped.droppedUnsupportedImages : 0,
      htmlQuoteShips,
    },
  };
}
