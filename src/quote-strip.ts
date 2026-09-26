import { InvalidInputError } from './coerce.js';

// Plain-text quote stripping for the READ path (#73). Unlike reply-quote.ts, which writes
// our own quote, the input here is a FOREIGN client's and a match DELETES text, so every
// match must be one of a small set of conventional, machine-emitted shapes.
//
// Recognise confidently or not at all: an unrecognised shape passes through byte-identical
// with `quotedBytesStripped` 0. Prefer under-strip at every fork. Over-strip is real (a
// leading ">" is also markdown or a shell prompt); both failure modes are in
// docs/email-bodies.md.

export interface QuoteStripResult {
  text: string;
  // Includes the attribution/marker lines and framing whitespace. 0 means no marker matched.
  quotedBytesStripped: number;
}

const QUOTE_LINE = /^[ \t]{0,3}>/;
const BLANK_LINE = /^\s*$/;

// Recognised only directly above a quote block, never on its own. Gmail wraps a long
// attribution, so the "On" opener may be up to two lines above the "wrote:" line.
const ATTRIBUTION_END = /\bwrote:\s*$/i;
const ATTRIBUTION_START = /^[ \t]*On\b/;
const ATTRIBUTION_MAX_WRAPPED_LINES = 3;

// Anchored as a whole line so the phrase inside a sentence can't match.
const ORIGINAL_MESSAGE = /^[ \t]*-{2,}[ \t]*Original Message[ \t]*-{2,}\s*$/i;

// Outlook's unprefixed header block, recognised as a BLOCK: an addressed "From:" line, then
// at least two more header lines. The ADDRESS requirement matters because this marker cuts
// to the END of the message: pasted "From: The Hiring Team" over To:/Subject: lines would
// otherwise take the reader's own text below it. A bare display-name From: under-strips.
const HEADER_FROM = /^[ \t]*From:[ \t]*(\S.*)$/;
const ADDRESS_TOKEN = /@|<[^<>]*>/;
const HEADER_SIBLING = /^[ \t]*(Sent|Date|To|Cc|Bcc|Subject|Reply-To):/i;
const HEADER_BLOCK_LOOKAHEAD = 6;
const HEADER_BLOCK_MIN_SIBLINGS = 2;

// A rule directly above a to-end marker is part of the separator, not content.
const SEPARATOR_RULE = /^[ \t]*[_-]{3,}\s*$/;

function byteLength(s: string): number {
  return Buffer.byteLength(s, 'utf8');
}

// A quoted line some converters wrap without its ">" prefix (#181). Two conditions tell it
// from a person's INLINE REPLY, which eating would delete the sender's own writing, and both
// must hold:
//
//   1. It does not begin at the left margin; a typed line is flush left.
//   2. It is glued to the quote: a quote line directly above with no blank between, and one
//      again within QUOTE_GAP_MAX_LINES below. An inline reply is set off by a blank line.
//
// Each alone would eat a shape the other saves: a flush-left "Yes." under a quoted question
// (kept by 1), an indented block pasted between quoted paragraphs (kept by 2). A flush-left
// continuation is a documented under-strip.
const UNPREFIXED_CONTINUATION = /^[ \t]+\S/;
// A wrap of a wrap yields two; past that the gap is a block of text, not a broken line.
const QUOTE_GAP_MAX_LINES = 2;

// The quote line that resumes the run after the unprefixed continuation at `j`, or -1.
function continuationResumesQuote(lines: string[], j: number): number {
  if (j === 0 || !QUOTE_LINE.test(lines[j - 1])) return -1;
  for (let k = j; k < lines.length && k - j < QUOTE_GAP_MAX_LINES; k++) {
    if (!UNPREFIXED_CONTINUATION.test(lines[k])) return -1;
    if (k + 1 < lines.length && QUOTE_LINE.test(lines[k + 1])) return k + 1;
  }
  return -1;
}

// Blank lines inside the run are tolerated (clients drop the "> " from an empty quoted
// line), but the run ends at the LAST quoted line, so a trailing blank separator is kept.
function quoteRunEnd(lines: string[], start: number): number {
  let last = start;
  for (let j = start; j < lines.length; j++) {
    if (QUOTE_LINE.test(lines[j])) { last = j; continue; }
    if (BLANK_LINE.test(lines[j])) continue;
    const resume = continuationResumesQuote(lines, j);
    if (resume < 0) break;
    // Resume ON the quote line, so `last` advances only to quote lines.
    j = resume - 1;
  }
  return last;
}

// Start of the region to remove for a quote run beginning at `q`: the blank separator
// above it, plus an attribution line (and its wrapped continuation) when one is there.
function quoteRegionStart(lines: string[], q: number): number {
  let i = q - 1;
  while (i >= 0 && BLANK_LINE.test(lines[i])) i--;
  const blankRunStart = i + 1;
  if (i < 0 || !ATTRIBUTION_END.test(lines[i])) return blankRunStart;

  let a = i;
  if (!ATTRIBUTION_START.test(lines[a])) {
    // Wrapped attribution. With no opener found, the "wrote:" line strips only itself.
    for (let k = a - 1; k >= 0 && a - k < ATTRIBUTION_MAX_WRAPPED_LINES; k--) {
      if (BLANK_LINE.test(lines[k])) break;
      if (ATTRIBUTION_START.test(lines[k])) { a = k; break; }
    }
  }
  let b = a - 1;
  while (b >= 0 && BLANK_LINE.test(lines[b])) b--;
  return b + 1;
}

// Start of the region for a to-end-of-message marker at `i`: the marker line, plus any
// horizontal rule and blank lines immediately above it.
function markerRegionStart(lines: string[], i: number): number {
  let k = i - 1;
  while (k >= 0 && (BLANK_LINE.test(lines[k]) || SEPARATOR_RULE.test(lines[k]))) k--;
  return k + 1;
}

function isHeaderBlock(lines: string[], i: number): boolean {
  const from = HEADER_FROM.exec(lines[i]);
  if (!from || !ADDRESS_TOKEN.test(from[1])) return false;
  let siblings = 0;
  for (let j = i + 1; j < lines.length && j <= i + HEADER_BLOCK_LOOKAHEAD; j++) {
    if (BLANK_LINE.test(lines[j])) continue;
    if (HEADER_SIBLING.test(lines[j])) { siblings++; continue; }
    break;
  }
  return siblings >= HEADER_BLOCK_MIN_SIBLINGS;
}

// REGIONS, not a single boundary: a "below the first marker" cut would destroy bottom-posted
// and inline replies. Only the two markers with no end delimiter (the Outlook header block,
// "Original Message") run to the end, so stripping a FORWARD leaves the covering note only.
export function stripQuotedText(text: string): QuoteStripResult {
  if (!text) return { text: text ?? '', quotedBytesStripped: 0 };

  const lines = text.split('\n');
  const remove = new Array<boolean>(lines.length).fill(false);
  const mark = (from: number, to: number) => {
    for (let k = Math.max(0, from); k <= to; k++) remove[k] = true;
  };
  let matched = false;

  for (let i = 0; i < lines.length; i++) {
    if (ORIGINAL_MESSAGE.test(lines[i]) || isHeaderBlock(lines, i)) {
      mark(markerRegionStart(lines, i), lines.length - 1);
      matched = true;
      break;
    }
    if (QUOTE_LINE.test(lines[i])) {
      const end = quoteRunEnd(lines, i);
      mark(quoteRegionStart(lines, i), end);
      matched = true;
      i = end;
    }
  }

  if (!matched) return { text, quotedBytesStripped: 0 };

  const kept = lines.filter((_, idx) => !remove[idx]);
  while (kept.length > 0 && BLANK_LINE.test(kept[0])) kept.shift();
  while (kept.length > 0 && BLANK_LINE.test(kept[kept.length - 1])) kept.pop();

  // Otherwise a CRLF body ends on a lone CR whose LF went with the quote.
  const stripped = kept.join('\n').replace(/\s+$/, '');
  return { text: stripped, quotedBytesStripped: byteLength(text) - byteLength(stripped) };
}

// Rejected rather than silently ignoring one flag, which would leave the caller believing
// the response was stripped.
export function assertStripQuotedNotRaw(stripQuoted: boolean, raw: boolean): void {
  if (stripQuoted && raw) {
    throw new InvalidInputError(
      'stripQuoted cannot be combined with raw: raw returns the JMAP response unmodified. ' +
      'Drop raw for a stripped bodyText, or drop stripQuoted for verbatim JMAP.',
    );
  }
}
