import { InvalidInputError } from './coerce.js';

// Plain-text quote stripping for the READ path (#73): given a message's text/plain body,
// remove the quoted correspondence that a reply carries below (or above) the new content,
// and report how many bytes went. A long thread re-sends the same reply chain at ever
// deeper quote depths, so the quoted tail dominates the tokens while the new information
// per message stays constant.
//
// This is a different job from the marker checks in reply-quote.ts, which is why it lives
// in its own module: hasQuoteMarker/hasTextQuoteMarker recognise OUR OWN emitted quote on a
// draft we are about to edit, and only decide whether a guard fires. Here the input is
// whatever a FOREIGN client produced, and a match DELETES text from the output — so the
// rules below are deliberately narrower and every match must be one of a small set of
// conventional, machine-emitted shapes.
//
// Posture: recognise confidently or not at all. When no marker matches, the body is
// returned byte-identical and `quotedBytesStripped` is 0 — an unrecognised shape passes
// through unchanged BY DESIGN rather than being guessed at. Callers can therefore treat a
// 0 as "this body is verbatim" and re-read nothing.

export interface QuoteStripResult {
  // The body with every recognised quoted region removed. Byte-identical to the input
  // when nothing matched.
  text: string;
  // UTF-8 bytes removed: the quoted lines themselves plus the attribution/marker lines
  // and the separator whitespace that framed them. 0 means no marker matched.
  quotedBytesStripped: number;
}

// A quoted line: a leading ">" run, optionally indented a little (clients occasionally
// pad). Nesting (">>", "> >") is the same shape and needs no separate rule.
const QUOTE_LINE = /^[ \t]{0,3}>/;
const BLANK_LINE = /^\s*$/;

// Attribution ("On <date>, <someone> wrote:") — recognised only as the last line of the
// run directly above a quote block, never on its own. Gmail wraps a long attribution, so
// a line that merely ENDS with "wrote:" can be joined to up to two preceding lines when
// one of them opens with "On".
const ATTRIBUTION_END = /\bwrote:\s*$/i;
const ATTRIBUTION_START = /^[ \t]*On\b/;
const ATTRIBUTION_MAX_WRAPPED_LINES = 3;

// "-----Original Message-----" (Outlook's separator; the dash count and inner spacing
// vary between clients, and Fastmail's own forward block uses the same words). Anchored
// as a whole line so a mention of the phrase inside a sentence can't match.
const ORIGINAL_MESSAGE = /^[ \t]*-{2,}[ \t]*Original Message[ \t]*-{2,}\s*$/i;

// Outlook's header block, which opens a quoted section with no ">" prefixing at all.
// Confidence comes from the BLOCK, not the single line: a "From:" line with a value, then
// at least two more header lines within the next few lines, with nothing but headers and
// blanks in between. One stray "From:" in prose therefore cannot trigger it.
const HEADER_FROM = /^[ \t]*From:[ \t]*\S/;
const HEADER_SIBLING = /^[ \t]*(Sent|Date|To|Cc|Bcc|Subject|Reply-To):/i;
const HEADER_BLOCK_LOOKAHEAD = 6;
const HEADER_BLOCK_MIN_SIBLINGS = 2;

// A horizontal rule ("________" / "-------") directly above a header block or an
// "Original Message" marker is part of the separator, not content.
const SEPARATOR_RULE = /^[ \t]*[_-]{3,}\s*$/;

function byteLength(s: string): number {
  return Buffer.byteLength(s, 'utf8');
}

// End of a run of quoted lines: blank lines inside the run are tolerated (clients often
// drop the "> " from an empty quoted line), but the run ends at the LAST quoted line, so
// a blank separator before following content is never swallowed. An unquoted content line
// ends the run — text below a quote block is the reader's own writing and is kept.
function quoteRunEnd(lines: string[], start: number): number {
  let last = start;
  for (let j = start; j < lines.length; j++) {
    if (QUOTE_LINE.test(lines[j])) { last = j; continue; }
    if (BLANK_LINE.test(lines[j])) continue;
    break;
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
    // Wrapped attribution: walk back over consecutive non-blank lines looking for the
    // "On ..." opener. Bounded, and abandoned entirely if it isn't found — a "wrote:"
    // line with no recognisable opener strips only itself.
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
  if (!HEADER_FROM.test(lines[i])) return false;
  let siblings = 0;
  for (let j = i + 1; j < lines.length && j <= i + HEADER_BLOCK_LOOKAHEAD; j++) {
    if (BLANK_LINE.test(lines[j])) continue;
    if (HEADER_SIBLING.test(lines[j])) { siblings++; continue; }
    break;
  }
  return siblings >= HEADER_BLOCK_MIN_SIBLINGS;
}

// Strip every recognised quoted region from a plain-text body.
//
// REGIONS, not a single boundary. A top-posted reply is the common case (new text, then
// one quote block running to the end) and falls out of this as the degenerate case, but
// scanning for regions also handles the two shapes a single "everything below the first
// marker" cut would destroy: a bottom-posted reply (quote first, new text under it) and an
// inline reply interleaved between quoted paragraphs. Unquoted text is never removed for
// being *positioned* after a quote.
//
// The two marker classes that have no end delimiter — an Outlook header block and
// "-----Original Message-----" — do run to the end of the message, because that is what
// they mean. Note the consequence for a FORWARD: a forwarded message's content sits below
// exactly that marker, so stripping a forward leaves the covering note only. That is
// reported through quotedBytesStripped rather than guessed around; see README/docs.
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

  // Nothing recognised: hand back the caller's bytes untouched.
  if (!matched) return { text, quotedBytesStripped: 0 };

  const kept = lines.filter((_, idx) => !remove[idx]);
  // Removing a region can leave the body opening or closing on blank lines; trim those
  // (they are counted in quotedBytesStripped like any other removed byte).
  while (kept.length > 0 && BLANK_LINE.test(kept[0])) kept.shift();
  while (kept.length > 0 && BLANK_LINE.test(kept[kept.length - 1])) kept.pop();

  // Trailing whitespace on the final kept line goes too — otherwise a CRLF body ends on
  // a stray lone CR whose LF was removed with the quote.
  const stripped = kept.join('\n').replace(/\s+$/, '');
  return { text: stripped, quotedBytesStripped: byteLength(text) - byteLength(stripped) };
}

// stripQuoted rewrites the simplified body; `raw` is the pure-JMAP escape valve that
// external clients JSON.parse wholesale. Rather than silently ignoring one of the two
// (leaving the caller to believe the response was stripped when it was not), the
// combination is rejected with the two ways forward.
export function assertStripQuotedNotRaw(stripQuoted: boolean, raw: boolean): void {
  if (stripQuoted && raw) {
    throw new InvalidInputError(
      'stripQuoted cannot be combined with raw: raw returns the JMAP response unmodified. ' +
      'Drop raw for a stripped bodyText, or drop stripQuoted for verbatim JMAP.',
    );
  }
}
