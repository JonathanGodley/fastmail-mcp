/**
 * The body tokens `{{signature}}`, `{{quote}}` and `{{forward}}`: a caller writes one into
 * the body it is composing to say WHERE this server's generated block belongs, instead of
 * accepting the fixed placement the builders would otherwise impose.
 *
 * This module scans and substitutes; the refusals and the receipt belong to the handler, so
 * that one grammar serves compose and edit alike.
 *
 * THE SECURITY RULE. Substitution is a SINGLE PASS: one alternation over all three tokens
 * with a FUNCTION replacer, run on the caller's own authored body only, before any fetched
 * content is joined in. Consecutive `replace` calls would let a later pass scan a block an
 * earlier one inserted, and the quote and forward blocks are built from an attacker-authored
 * original, so a `{{signature}}` inside somebody's email would be expanded.
 * `String.prototype.replace` never rescans what a replacer returns; a string replacement would
 * interpret `$&` and friends inside the block, which is why the replacer stays a function.
 *
 * THE SCAN IS A RAW STRING SCAN. No HTML parsing, no entity decoding, now or later. A token
 * split by a tag (`{{sig<b></b>nature}}`) or spelled as entities is NOT a token. Decoding
 * would create a second scan surface that disagrees per part (html decodes, text does not)
 * and would make `\{{` unspellable next to a literal `&#123;`.
 */

/** The three tokens. Nothing else is a token, in any spelling. */
export type BodyTokenName = 'signature' | 'quote' | 'forward';

const TOKEN_NAMES: readonly BodyTokenName[] = ['signature', 'quote', 'forward'];

/**
 * The one alternation. Every branch of the grammar is a branch of THIS regex, never a pass
 * before or after it, so an escaped `{{SIGNATURE}}` is consumed where it stands and reaches no
 * near-miss report.
 *
 *  1. The escape. A backslash before a spelling branch 2 would match as a token or near-miss
 *     is consumed and the braces ship literal; `classify` applies branch 2's run-length test.
 *     A backslash before prose (`\{signature}`) escapes nothing and ships. Exactly ONE raw
 *     backslash is consumed, so there is no spelling for a literal backslash before a real token.
 *  2. A brace run on each side of one of the names; `classify` decides token, near-miss or prose.
 *  3. Any other `{{…}}`, reported as left unexpanded. The `[^{]` MUST NOT be widened: crossing a
 *     `{` would let an outer prose spelling swallow a real token (`{{a{{signature}}`) and hide it
 *     from the refusal. The price, `{{a{b}}}` reported nowhere, is the right way round.
 *
 * The `(?!\{)` makes brace runs maximal; without it `{{{signature}}}` backtracks into the exact
 * token. The `(?<!\{)` on branch 2 is a complexity bound: without it a long brace run costs
 * quadratic time (about 18 seconds for 100,000 braces). It must NOT be hoisted onto the whole
 * alternation, because branch 3 has to start inside a run (`{{{a}}` is the prose `{{a}}`).
 *
 * `\s` is exactly the class `String.prototype.trim()` strips, so NBSP, U+202F and U+FEFF
 * inside the braces still make a token; U+200B is not in it, so `{{<U+200B>signature}}` is prose.
 */
const BODY_TOKEN_RE = new RegExp(
  [
    String.raw`\\(\{+)(?!\{)\s*(?:${TOKEN_NAMES.join('|')})\s*(\}+)`,
    String.raw`(?<!\{)(\{+)(?!\{)\s*(${TOKEN_NAMES.join('|')})\s*(\}+)`,
    String.raw`\{\{[^{]*?\}\}`,
  ].join('|'),
  'gi',
);

/** Where a spelling sits in the part, and the literal text the caller wrote there. */
export interface BodyTokenSite {
  /** The token this spelling is (a near-miss carries the name it near-missed). */
  name: BodyTokenName;
  /** Index of the spelling in the part, so a caller can report landing order. */
  index: number;
  /** The literal text the caller wrote, verbatim — what a refusal quotes back. */
  text: string;
}

/** A `{{…}}` spelling that is neither a token nor a near-miss, so it names no token. */
export interface BodySpellingSite {
  index: number;
  text: string;
}

/** What one body part says about the tokens. Facts only — no decision, no refusal. */
export interface BodyTokenScan {
  /** Every exact token, in the order it appears in the part. */
  tokens: BodyTokenSite[];
  /** How many exact tokens of each name the part carries. All three keys are always present. */
  counts: Record<BodyTokenName, number>;
  /** Every unescaped near-miss spelling, with the literal text so a handler can quote it back. */
  nearMisses: BodyTokenSite[];
  /** Every other `{{…}}` whose contents carry no further `{` (branch 3's bound, see BODY_TOKEN_RE). */
  otherSpellings: BodySpellingSite[];
  /**
   * Every ESCAPED spelling, backslash included. Not a defect: it exists for `edit_draft`, which
   * stores an unflagged body byte for byte, so an escape there ships WITH its backslash.
   */
  escapes: BodySpellingSite[];
}

/**
 * Why a block has nothing to expand to; each is a distinct sentence a handler owes the caller.
 *
 *  - `no-signature`                the identity has no signature at all.
 *  - `no-text-form`                  the signature exists but has no form this part can carry
 *                                    (an images-only html signature, in a text part).
 *  - `nothing-quotable`              the original has nothing quotable in ANY form —
 *                                    attachment-only, or images that cannot be carried.
 *  - `nothing-quotable-in-this-form` the original is quotable, but not in this part's form.
 */
export type BlockUnavailableCause =
  | 'no-signature'
  | 'no-text-form'
  | 'nothing-quotable'
  | 'nothing-quotable-in-this-form';

/**
 * One block, for one part. An unavailable block carries WHY, so a token that expands to
 * nothing never vanishes in silence.
 *
 * `'as-written'` leaves the spelling exactly as typed, where `undefined` removes it. It is for
 * `edit_draft`, which stores an unflagged body as written: a `{{quote}}` handed back there is
 * text, and deleting it would silently edit the caller's body.
 */
export type BodyBlock =
  | { available: true; content: string }
  | { available: false; cause: BlockUnavailableCause }
  | { available: 'as-written' };

/**
 * The blocks for one part. All three keys are REQUIRED so no token reaches expansion
 * unconsidered; `undefined` is for a token this call does not offer (no `{{forward}}` on a reply).
 */
export type BodyBlocks = Record<BodyTokenName, BodyBlock | undefined>;

/** One token site after expansion. */
export interface ExpandedTokenSite extends BodyTokenSite {
  /** False when the token was removed or, with `asWritten`, left exactly as typed. */
  expanded: boolean;
  asWritten?: true;
  /**
   * Why nothing was substituted. Absent on an unexpanded, non-as-written site means `blocks` had
   * NO entry for that token: the handler's own bug, not a fact about the message.
   */
  cause?: BlockUnavailableCause;
}

/** What one expansion did, in enough detail to build a receipt naming tokens, counts and order. */
export interface BodyTokenExpansion {
  text: string;
  tokens: ExpandedTokenSite[];
  counts: Record<BodyTokenName, number>;
  /** Left in the text verbatim. The refusal is decided from `scanBodyTokens`, before expansion. */
  nearMisses: BodyTokenSite[];
  otherSpellings: BodySpellingSite[];
}

/** No token of any name. A fresh object each call — callers increment it. */
function zeroCounts(): Record<BodyTokenName, number> {
  return { signature: 0, quote: 0, forward: 0 };
}

type Classified =
  | { kind: 'escape'; literal: string }
  | { kind: 'token'; name: BodyTokenName }
  | { kind: 'near-miss'; name: BodyTokenName }
  | { kind: 'prose' }
  | { kind: 'other' };

/**
 * Which branch of the grammar this match is. The case and run-length tests live here rather
 * than in the regex because splitting them into separate patterns would reintroduce a second
 * pass; the escape shares branch 2's run-length test so the two cannot drift apart.
 */
function classify(
  whole: string,
  escapedLeft: string | undefined,
  escapedRight: string | undefined,
  left: string | undefined,
  name: string | undefined,
  right: string | undefined,
): Classified {
  if (escapedLeft !== undefined && escapedRight !== undefined) {
    // Slices rather than reading a capture: exactly ONE backslash is dropped, and the capture
    // is only the brace run.
    if (escapedLeft.length >= 2 || escapedRight.length >= 2) {
      return { kind: 'escape', literal: whole.slice(1) };
    }
    return { kind: 'prose' };
  }
  if (left === undefined || name === undefined || right === undefined) return { kind: 'other' };
  const lower = name.toLowerCase() as BodyTokenName;
  if (left.length === 2 && right.length === 2 && name === lower) return { kind: 'token', name: lower };
  if (left.length >= 2 || right.length >= 2) return { kind: 'near-miss', name: lower };
  return { kind: 'prose' };
}

/**
 * What one body part says about the tokens. REPORTS ONLY: refusals belong to the handler, and
 * a scanned part keeps its text exactly (an escape's backslash is consumed only by expansion).
 */
export function scanBodyTokens(part: string): BodyTokenScan {
  const scan: BodyTokenScan = { tokens: [], counts: zeroCounts(), nearMisses: [], otherSpellings: [], escapes: [] };
  // `matchAll` honours a module-level `g` pattern's `lastIndex` (unlike `replace`), so a non-zero
  // value would silently skip every token before it. Nothing here calls `test` or `exec` today,
  // so this is 0 already; the reset keeps it that way if someone does.
  BODY_TOKEN_RE.lastIndex = 0;
  for (const m of part.matchAll(BODY_TOKEN_RE)) {
    const index = m.index ?? 0;
    const text = m[0];
    const c = classify(text, m[1], m[2], m[3], m[4], m[5]);
    switch (c.kind) {
      case 'token':
        scan.tokens.push({ name: c.name, index, text });
        scan.counts[c.name]++;
        break;
      case 'near-miss':
        scan.nearMisses.push({ name: c.name, index, text });
        break;
      case 'other':
        scan.otherSpellings.push({ index, text });
        break;
      case 'escape':
        scan.escapes.push({ index, text });
        break;
      case 'prose':
        break;
    }
  }
  return scan;
}

/**
 * Substitute the blocks into the caller's authored part: THE single pass of the module
 * header's security rule. `authored` must never already have fetched content joined in.
 *
 * A block lands at the token's position with no spacer or newline added; spacing at the join
 * is the caller's, who wrote the body around the token.
 */
export function expandBodyTokens(authored: string, blocks: BodyBlocks): BodyTokenExpansion {
  const tokens: ExpandedTokenSite[] = [];
  const counts = zeroCounts();
  const nearMisses: BodyTokenSite[] = [];
  const otherSpellings: BodySpellingSite[] = [];

  const text = authored.replace(
    BODY_TOKEN_RE,
    (
      whole: string,
      escapedLeft: string | undefined,
      escapedRight: string | undefined,
      left: string | undefined,
      name: string | undefined,
      right: string | undefined,
      index: number,
    ): string => {
      const c = classify(whole, escapedLeft, escapedRight, left, name, right);
      switch (c.kind) {
        case 'escape': {
          // An as-written token's escape is as-written too; it may be someone else's text.
          const escaped = whole.replace(/[\\{}\s]/g, '').toLowerCase() as BodyTokenName;
          return blocks[escaped]?.available === 'as-written' ? whole : c.literal;
        }
        case 'near-miss':
          nearMisses.push({ name: c.name, index, text: whole });
          return whole;
        case 'other':
          otherSpellings.push({ index, text: whole });
          return whole;
        case 'prose':
          return whole;
        case 'token': {
          counts[c.name]++;
          const block = blocks[c.name];
          if (block?.available === 'as-written') {
            tokens.push({ name: c.name, index, text: whole, expanded: false, asWritten: true });
            return whole;
          }
          const site: ExpandedTokenSite = {
            name: c.name,
            index,
            text: whole,
            expanded: block?.available === true,
          };
          if (block?.available === false) site.cause = block.cause;
          tokens.push(site);
          return block?.available === true ? block.content : '';
        }
      }
    },
  );

  return { text, tokens, counts, nearMisses, otherSpellings };
}
