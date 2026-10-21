/**
 * The join around each block this server inserted at a body token: the caller's text either
 * side of it and the block's own first and last lines, so a bad join (a closing line over the
 * signature's own, doubled spacing, history above the sign-off) is visible in the write
 * result. It reports; it never edits the body.
 *
 * Every string a join carries is display text, already redacted, neutralised and bounded.
 * Redaction runs over the whole text before the block, the block, and the whole text after
 * it, each in one piece, before any line is cut out: a credential can span two lines.
 */
import type { BodyTokenName, ExpandedTokenSite } from './body-tokens.js';
import { htmlToText } from './body-format.js';
import { redactBearerTokens } from './coerce.js';
import { describePart } from './inline-images.js';

export type JoinPart = 'textBody' | 'htmlBody';

/**
 * A history block's text, shown as a count and never echoed: it was built from an original
 * someone else wrote, and the caller may not have read it.
 */
export interface HistoryLabel {
  label: 'quoted original' | 'forwarded message';
  lines: number;
}

export type JoinLine = string | HistoryLabel;

export interface TokenJoin {
  token: BodyTokenName;
  part: JoinPart;
  /** Up to 2 lines, top to bottom; empty at the start of the body. */
  above: JoinLine[];
  /** Text parts: blank lines between `above` and the block, wherever they fall. */
  blankAbove?: number;
  /** Html parts: the tags, comments and whitespace on the caller's side of the token. */
  markupAbove?: string;
  /** The first 2 and last 2 lines, or all of them when there are 4 or fewer. Empty: renders no text. */
  block: JoinLine[];
  /** Lines left out between the block's second and third shown lines. */
  blockOmitted?: number;
  markupBelow?: string;
  blankBelow?: number;
  /** Up to 2 lines; empty at the end of the body. */
  below: JoinLine[];
}

/** Where one inserted block sits in the stored part. */
export interface InsertedBlock {
  token: BodyTokenName;
  start: number;
  end: number;
}

/** A long sign-off or attribution line fits; at most 8 lines a join, so a join stays small. */
export const JOIN_LINE_ECHO_LIMIT = 120;
/** Two block wrappers with inline styles fit; #176's doubled spacer is 36. */
export const JOIN_MARKUP_ECHO_LIMIT = 80;

const WINDOW = 2;

const HISTORY: Partial<Record<BodyTokenName, HistoryLabel['label']>> = {
  quote: 'quoted original',
  forward: 'forwarded message',
};

/** The blocks an expansion inserted, with their output offsets moved by `shift`. */
export function insertedBlocks(sites: readonly ExpandedTokenSite[], shift = 0): InsertedBlock[] {
  return sites
    .filter((s) => s.output !== undefined)
    .map((s) => ({ token: s.name, start: s.output!.start + shift, end: s.output!.end + shift }));
}

/** One line for display: neutralised, bounded, trimmed. Redact before calling. */
export function echoJoinLine(line: string): string {
  return describePart(line, JOIN_LINE_ECHO_LIMIT).trim();
}

/** The joins of every block inserted into one stored part, in body order. */
export function buildJoins(part: JoinPart, stored: string, blocks: readonly InsertedBlock[]): TokenJoin[] {
  const sorted = [...blocks].sort((a, b) => a.start - b.start);
  // Caller text and blocks alternate, caller text first and last (possibly empty), so block
  // k is piece 2k+1.
  const pieces: Piece[] = [];
  let cursor = 0;
  for (const b of sorted) {
    pieces.push({ text: stored.slice(cursor, b.start) });
    pieces.push({ text: stored.slice(b.start, b.end), token: b.token });
    cursor = b.end;
  }
  pieces.push({ text: stored.slice(cursor) });
  return sorted.map((_, k) => (part === 'htmlBody' ? htmlJoin(pieces, 2 * k + 1) : textJoin(pieces, 2 * k + 1)));
}

interface Piece {
  text: string;
  /** Set on a block; absent on the caller's text. */
  token?: BodyTokenName;
}

const historyOf = (p: Piece): HistoryLabel['label'] | undefined => (p.token ? HISTORY[p.token] : undefined);

// ---------------------------------------------------------------------------
// Text parts: physical lines, with blank runs counted
// ---------------------------------------------------------------------------

interface Range {
  start: number;
  end: number;
}

interface Item extends Range {
  line: JoinLine;
  blank: boolean;
}

/** Does a line (content only, terminator excluded) hold a character of `r`? */
function touches(line: Range, r: Range): boolean {
  if (line.start === line.end) return r.start <= line.start && line.start < r.end;
  return line.start < r.end && line.end > r.start;
}

function textJoin(pieces: Piece[], at: number): TokenJoin {
  // Rebuilt with each region redacted whole; a history block is kept as it is, since its
  // text is only counted, never shown.
  let text = '';
  let group = '';
  const history: (Range & { label: HistoryLabel['label'] })[] = [];
  let block!: Range;
  pieces.forEach((p, i) => {
    const label = historyOf(p);
    if (i !== at && !label) {
      group += p.text;
      return;
    }
    text += redactBearerTokens(group);
    group = '';
    const start = text.length;
    text += label ? p.text : redactBearerTokens(p.text);
    if (label) history.push({ start, end: text.length, label });
    if (i === at) block = { start, end: text.length };
  });
  text += redactBearerTokens(group);

  const items: Item[] = [];
  const breaks = /\r\n|\n|\r/g;
  let lineStart = 0;
  for (;;) {
    const m = breaks.exec(text);
    const line = { start: lineStart, end: m ? m.index : text.length };
    const h = history.find((r) => touches(line, r));
    const prev = items[items.length - 1];
    if (h && prev && typeof prev.line === 'object') {
      prev.line.lines++;
      prev.end = line.end;
    } else if (h) {
      items.push({ ...line, line: { label: h.label, lines: 1 }, blank: false });
    } else {
      const shown = echoJoinLine(text.slice(line.start, line.end));
      items.push({ ...line, line: shown, blank: shown === '' });
    }
    if (!m) break;
    lineStart = m.index + m[0].length;
  }

  const inBlock = items.map((it) => touches(it, block));
  const shown = items.map((it, i) => inBlock[i] && !it.blank);
  let first = shown.indexOf(true);
  let last = shown.lastIndexOf(true);
  if (first === -1) {
    // The block renders nothing; the join sits where its lines would be.
    first = items.findIndex((it, i) => inBlock[i] || it.start >= block.end);
    last = first - 1;
  }

  let up = first - 1;
  let blankAbove = 0;
  while (up >= 0 && items[up].blank) { blankAbove++; up--; }
  let down = last + 1;
  let blankBelow = 0;
  while (down < items.length && items[down].blank) { blankBelow++; down++; }

  const visible = (list: Item[]) => list.filter((it) => !it.blank).map((it) => it.line);
  return {
    token: pieces[at].token!,
    part: 'textBody',
    above: visible(items.slice(0, up + 1)).slice(-WINDOW),
    blankAbove,
    ...windowBlock(visible(items.slice(first, last + 1))),
    blankBelow,
    below: visible(items.slice(down)).slice(0, WINDOW),
  };
}

function windowBlock(lines: JoinLine[]): Pick<TokenJoin, 'block' | 'blockOmitted'> {
  if (lines.length <= 2 * WINDOW) return { block: lines };
  return {
    block: [...lines.slice(0, WINDOW), ...lines.slice(-WINDOW)],
    blockOmitted: lines.length - 2 * WINDOW,
  };
}

// ---------------------------------------------------------------------------
// Html parts: each piece rendered on its own, spacing shown as the markup at the join
// ---------------------------------------------------------------------------

/** The image policy the shipped text alternative uses, so an image-only signature shows as [image]. */
const render = (html: string): string => htmlToText(html, 'unconditional');

const shownLines = (rendered: string): string[] =>
  rendered.split(/\r\n|\n|\r/).map(echoJoinLine).filter((l) => l !== '');

/** The lines of a run of pieces: caller text and signatures redacted together, history as labels. */
function htmlLines(pieces: Piece[]): JoinLine[] {
  const out: JoinLine[] = [];
  let group: string[] = [];
  const flush = () => {
    out.push(...shownLines(redactBearerTokens(group.join('\n'))));
    group = [];
  };
  for (const p of pieces) {
    const label = historyOf(p);
    if (!label) {
      group.push(render(p.text));
      continue;
    }
    flush();
    out.push({ label, lines: render(p.text).split(/\r\n|\n|\r/).filter((l) => l.trim() !== '').length });
  }
  flush();
  return out;
}

function htmlJoin(pieces: Piece[], at: number): TokenJoin {
  const before = pieces[at - 1].text;
  const after = pieces[at + 1].text;
  const inComment = before.lastIndexOf('<!--') > before.lastIndexOf('-->');
  const lt = before.lastIndexOf('<');
  const inTag = !inComment && lt > before.lastIndexOf('>') && TAG_START.test(before.charAt(lt + 1));
  return {
    token: pieces[at].token!,
    part: 'htmlBody',
    above: htmlLines(pieces.slice(0, at)).slice(-WINDOW),
    markupAbove: echoMarkup(markupBefore(before, inComment, inTag), true),
    ...windowBlock(htmlLines([pieces[at]])),
    markupBelow: echoMarkup(markupAfter(after, inComment, inTag), false),
    below: htmlLines(pieces.slice(at + 1)).slice(0, WINDOW),
  };
}

const TAG_START = /[A-Za-z/!?]/;

/** The run of tags, comments and whitespace ending where `seg` ends (at the token). */
function markupBefore(seg: string, inComment: boolean, inTag: boolean): string {
  let p = seg.length;
  if (inComment) p = seg.lastIndexOf('<!--');
  else if (inTag) p = seg.lastIndexOf('<');
  for (;;) {
    if (/\s/.test(seg.charAt(p - 1))) { p--; continue; }
    if (seg.endsWith('-->', p)) {
      const q = seg.lastIndexOf('<!--', p - 3);
      if (q === -1) break;
      p = q;
      continue;
    }
    if (seg.charAt(p - 1) === '>') {
      const q = seg.lastIndexOf('<', p - 2);
      if (q === -1 || !TAG_START.test(seg.charAt(q + 1))) break;
      p = q;
      continue;
    }
    break;
  }
  return seg.slice(p);
}

/** The run of tags, comments and whitespace starting where `seg` starts (at the token). */
function markupAfter(seg: string, inComment: boolean, inTag: boolean): string {
  let p = 0;
  if (inComment || inTag) {
    const close = inComment ? '-->' : '>';
    const q = seg.indexOf(close);
    p = q === -1 ? seg.length : q + close.length;
  }
  for (;;) {
    if (/\s/.test(seg.charAt(p))) { p++; continue; }
    if (seg.startsWith('<!--', p)) {
      const q = seg.indexOf('-->', p + 4);
      if (q === -1) break;
      p = q + 3;
      continue;
    }
    if (seg.charAt(p) === '<' && TAG_START.test(seg.charAt(p + 1))) {
      const q = seg.indexOf('>', p + 1);
      if (q === -1) break;
      p = q + 1;
      continue;
    }
    break;
  }
  return seg.slice(0, p);
}

/** Redacted whole, cut on the side away from the token (`keepEnd` above it), then neutralised. */
function echoMarkup(run: string, keepEnd: boolean): string {
  const points = [...redactBearerTokens(run)];
  if (points.length <= JOIN_MARKUP_ECHO_LIMIT) return describePart(points.join(''), JOIN_MARKUP_ECHO_LIMIT);
  const kept = describePart(
    (keepEnd ? points.slice(-JOIN_MARKUP_ECHO_LIMIT) : points.slice(0, JOIN_MARKUP_ECHO_LIMIT)).join(''),
    JOIN_MARKUP_ECHO_LIMIT,
  );
  return keepEnd ? `…${kept}` : `${kept}…`;
}
