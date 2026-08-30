/**
 * The body tokens `{{signature}}`, `{{quote}}` and `{{forward}}`: a caller writes one into
 * the body it is composing to say WHERE this server's generated block belongs, instead of
 * accepting the fixed placement the builders would otherwise impose.
 *
 * This module is the whole grammar and nothing else. It scans and it substitutes; it does
 * not decide, refuse, or compose a sentence for anyone. A handler reads a scan, refuses what
 * it will not accept, builds the blocks, and calls the expander — the refusal wording and the
 * receipt are the handler's, so that the same grammar can serve compose and edit alike.
 *
 * ---------------------------------------------------------------------------
 * THE SECURITY RULE. Substitution is a SINGLE PASS: one alternation over all three tokens
 * with a FUNCTION replacer, run on the caller's own body only, before any fetched content is
 * joined in.
 *
 * Three consecutive `replace` calls would let a later pass scan the block an earlier one
 * inserted — and the quote and forward blocks are built from an attacker-authored original.
 * A `{{signature}}` sitting inside somebody's email would then land on the REPLACEMENT side
 * of a substitution and be expanded, which is a stranger choosing what goes into the user's
 * outgoing message. One pass makes that impossible: `String.prototype.replace` never rescans
 * what a replacer returns, so a block's own content is inert by construction. The replacer
 * must stay a FUNCTION for the same reason a second pass is forbidden — a string replacement
 * would interpret `$&` and friends inside the block.
 *
 * The fetched original never enters the substitution as INPUT either. Only the caller's
 * authored part does.
 * ---------------------------------------------------------------------------
 *
 * THE SCAN IS A RAW STRING SCAN. No HTML parsing, no entity decoding — not here, and not in
 * any later edit of this file. A token split by a tag (`{{sig<b></b>nature}}`) or spelled as
 * entities (`&#123;&#123;signature&#125;&#125;`) is NOT a token and is stored exactly as the
 * caller wrote it. Decoding would create a second scan surface that disagrees per part (an
 * html part decodes, a text part does not), and it would make `\{{` unspellable next to a
 * literal `&#123;`, because the escape would then have to know which of the two layers it
 * was escaping.
 *
 * Nothing in a stored part is rewritten after expansion.
 */

/** The three tokens. Nothing else is a token, in any spelling. */
export type BodyTokenName = 'signature' | 'quote' | 'forward';

const TOKEN_NAMES: readonly BodyTokenName[] = ['signature', 'quote', 'forward'];

/**
 * The one alternation. Every branch of the grammar is a branch of THIS regex — the escape,
 * the exact tokens, the near-miss spellings and every other `{{…}}` — never a pass before or
 * after it. That is what lets an escaped `{{SIGNATURE}}` reach no near-miss report at all:
 * the escape branch consumes it where it stands, so nothing downstream ever sees it, and a
 * body quoting another templating system's token can be written.
 *
 * Branch by branch:
 *
 *  1. `\\(\{+)(?!\{)\s*(name)\s*(\}+)` — the escape. A backslash immediately before a spelling
 *     branch 2 would match AS A TOKEN OR A NEAR-MISS is consumed, and the braces pass through
 *     literal. Which of those a spelling is — or whether it is the single-brace-each-side prose
 *     branch 2 leaves alone — is decided in the replacer from the run lengths, the same test
 *     branch 2 uses, written once. A backslash before PROSE escapes nothing and is not eaten:
 *     `\{signature}` ships both characters, so `use \{quote} in your template` keeps its
 *     backslash. It is placed FIRST, per the grammar; the ordering is belt-and-braces here,
 *     since the backslash sits one character to the left of anything branch 2 could start on
 *     and the engine reaches it first regardless. The rule is over RAW CHARACTERS and consumes
 *     exactly ONE backslash — a run of them is not special — so `\\{{signature}}` written with
 *     two raw backslashes ships one literal backslash followed by the literal token text, and
 *     there is no spelling for a literal backslash directly before a REAL token. A backslash
 *     before any other `{{` is prose and stays.
 *  2. `(?<!\{)(\{+)(?!\{)\s*(name)\s*(\}+)` — an unescaped brace run on each side of one of the
 *     three names. Whether that is the exact token or a near-miss is decided in the replacer,
 *     from the run lengths and the case: exactly two braces on each side and an exactly-lowercase
 *     name is the token; one or more on each side with two or more on at least one is a
 *     near-miss; a single brace on each side is ordinary prose and template syntax
 *     (`{signature}`), left alone and reported nowhere.
 *  3. `\{\{[^{]*?\}\}` — any other `{{…}}`. Prose (`{{ item }}`, `{{#if}}`), passed through
 *     untouched and reported so a caller can be told what was left unexpanded. The `[^{]` stops
 *     it at the next `{` and MUST NOT be widened: a branch 3 that could cross a `{` would let an
 *     outer prose spelling swallow a real token spelling inside it (`{{a{{signature}}` would
 *     match whole), hiding it from the report and from the refusal that reads the report. The
 *     price is that a spelling containing a `{` (`{{a{b}}}`) is reported nowhere, which is the
 *     right way round: unreported prose is a smaller failure than an unreported token.
 *
 * THE BRACE RUNS ARE MATCHED MAXIMALLY, which the `(?!\{)` enforces. Without it the engine
 * backtracks the run and `{{{signature}}}` matches as the exact token with a stray brace left
 * on each side — expanding a spelling the grammar calls a near-miss.
 *
 * THE `(?<!\{)` ON BRANCH 2 IS A COMPLEXITY BOUND, and it belongs to branch 2 alone. The
 * lookahead above makes a run atomic within ONE attempt, but the engine still begins a FRESH
 * attempt at every position, so branch 2 re-walked the whole remaining run from each brace in
 * it: a body carrying a long run of braces cost quadratic time, measured here at 733ms for
 * 20,000 braces and about 18 seconds for 100,000, off one authored body. The lookbehind refuses
 * to START branch 2 inside a run, which forfeits nothing — a start inside a run could never
 * match, because the `(?!\{)` already required the run to be maximal — and over 46,000 generated
 * spellings it changes no classification at all. It must NOT be hoisted onto the whole
 * alternation: branch 3 has to be able to start inside a run (`{{{a}}` is the prose `{{a}}` at
 * index 1), and branch 1 is anchored by its own `\\` and was never a driver.
 *
 * `\s` is used for the trimming, and that is exact rather than approximate: ECMAScript's `\s`
 * is the same character class `String.prototype.trim()` strips. So `{{ signature }}` is the
 * token, and so is one carrying an NBSP, a U+202F or a U+FEFF inside the braces — none of
 * those can make a token fall through as prose. U+200B is NOT in that class and is not
 * trimmed, so `{{<U+200B>signature}}` really is prose (it falls to branch 3).
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
  /**
   * Every unescaped near-miss spelling, each carrying the literal text written, so a handler
   * can quote it back rather than describe it. An ESCAPED spelling is not here: the escape
   * branch consumed it and the caller meant it as text.
   */
  nearMisses: BodyTokenSite[];
  /**
   * Every other `{{…}}` in the part WHOSE CONTENTS CARRY NO FURTHER `{`, so a handler can
   * report what it left unexpanded. The bound is branch 3's, and deliberate: `{{a{b}}}` is
   * reported nowhere, because a branch 3 that could cross a `{` would let an outer prose
   * spelling swallow a real token inside it and hide it from this very list.
   */
  otherSpellings: BodySpellingSite[];
}

/**
 * Why a block has nothing to expand to. Each is a distinct sentence a handler owes the
 * caller, which is why they are distinguished here rather than collapsed into one "empty".
 *
 *  - `no-signature`                  the identity has no signature at all.
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
 * One block, for one part. An unavailable block carries WHY, because the handler has to name
 * the cause to the caller — a token that expands to nothing must never vanish in silence.
 */
export type BodyBlock =
  | { available: true; content: string }
  | { available: false; cause: BlockUnavailableCause };

/**
 * The blocks for one part, one per token.
 *
 * All three keys are REQUIRED so that reaching expansion with a token nobody considered is
 * impossible; `undefined` is the honest value for a token this call does not offer at all
 * (there is no `{{forward}}` on a reply), and a handler that does not offer a token refuses
 * it from the scan long before it gets here.
 */
export type BodyBlocks = Record<BodyTokenName, BodyBlock | undefined>;

/** One token site after expansion. */
export interface ExpandedTokenSite extends BodyTokenSite {
  /** True when the block's content was substituted here. False when the token was removed. */
  expanded: boolean;
  /**
   * Why nothing was substituted, on an unexpanded site whose block named a cause.
   *
   * Absent on an unexpanded site means `blocks` carried NO entry for that token — the handler
   * reached expansion with a token it never authorised. That is the handler's own bug to
   * report, not a fact about the message, and it is reported rather than hidden so it cannot
   * pass for an ordinary empty block.
   */
  cause?: BlockUnavailableCause;
}

/** What one expansion did, in enough detail to build a receipt naming tokens, counts and order. */
export interface BodyTokenExpansion {
  /** The part, with every exact token replaced by its block or removed. */
  text: string;
  /** Every exact token that was in the part, in order, and what happened to it. */
  tokens: ExpandedTokenSite[];
  /** How many exact tokens of each name the part carried. All three keys are always present. */
  counts: Record<BodyTokenName, number>;
  /**
   * Near-miss spellings, left in the text verbatim. The same single pass sees them, so they
   * are reported here too — but the REFUSAL is decided from `scanBodyTokens` before any
   * expansion runs. This function refuses nothing.
   */
  nearMisses: BodyTokenSite[];
  /** Other `{{…}}` spellings, left in the text verbatim. Same bound as on the scan. */
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
 * Which branch of the grammar this match is, decided from the capture groups alone.
 *
 * The case test and the run-length test both live here rather than in the regex, so that one
 * alternation can serve the token, the near-miss and the escape at once — splitting them into
 * separate patterns is what would reintroduce a second pass. The escape reads the SAME
 * run-length test as branch 2 for the same reason: there is one rule about what a brace run
 * spells, and an escape branch carrying its own copy of it in regex syntax is how the two
 * drift apart.
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
    // An escape needs something to escape. A single brace on each side is prose branch 2 would
    // have left alone, so the backslash before it is escaping nothing and ships as written:
    // `\{signature}` stays `\{signature}`. Exactly ONE backslash is dropped from the front,
    // which is why this slices rather than reading a capture — the capture is the brace run.
    if (escapedLeft.length >= 2 || escapedRight.length >= 2) {
      return { kind: 'escape', literal: whole.slice(1) };
    }
    return { kind: 'prose' };
  }
  if (left === undefined || name === undefined || right === undefined) return { kind: 'other' };
  const lower = name.toLowerCase() as BodyTokenName;
  if (left.length === 2 && right.length === 2 && name === lower) return { kind: 'token', name: lower };
  // One brace on each side is ordinary template and prose syntax, not a near-miss of anything.
  if (left.length >= 2 || right.length >= 2) return { kind: 'near-miss', name: lower };
  return { kind: 'prose' };
}

/**
 * What one body part says about the tokens: the exact tokens present and where, the near-miss
 * spellings present as the caller wrote them, and every other `{{…}}` in the part.
 *
 * REPORTS ONLY. Which near-miss is a refusal, and whether an unexpandable token is an error,
 * belong to the handler that reads this. The scan runs on every written part; the backslash of
 * an escape is consumed only where substitution runs, so a scanned part keeps its text exactly.
 */
export function scanBodyTokens(part: string): BodyTokenScan {
  const scan: BodyTokenScan = { tokens: [], counts: zeroCounts(), nearMisses: [], otherSpellings: [] };
  // The pattern is module-level and `g`, so it carries `lastIndex` between calls. `matchAll`
  // and `replace` both reset it themselves today; the reset is here so that a future read
  // through `exec` or `test` cannot make this function's answer depend on the call before it.
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
      // An escape means the caller asked for that text; prose was never a token. Neither is
      // reported, because neither is something the caller has to be told about.
      case 'escape':
      case 'prose':
        break;
    }
  }
  return scan;
}

/**
 * Substitute the blocks into the caller's authored part. THE single pass — see the security
 * rule in the module header before changing anything about how this runs.
 *
 * A token whose block is available is replaced by the block's content AT THE TOKEN'S POSITION,
 * with nothing added around it: no spacer, no newline. Spacing at the join is the caller's,
 * because the caller is the one who wrote the body around the token.
 *
 * A token with nothing to expand to is REMOVED, and the result says so per site, with the
 * cause its block named.
 *
 * `authored` must be the caller's own body and nothing else — never a part that has already
 * had fetched content joined into it.
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
        case 'escape':
          // The braces ship as text, the backslash does not.
          return c.literal;
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
          const site: ExpandedTokenSite = {
            name: c.name,
            index,
            text: whole,
            expanded: block?.available === true,
          };
          if (block && !block.available) site.cause = block.cause;
          tokens.push(site);
          // The block's own content is returned from a FUNCTION replacer, so it is never
          // rescanned and never reinterpreted: a `{{signature}}` inside a quoted attacker
          // email lands here inert, as literal text.
          return block?.available === true ? block.content : '';
        }
      }
    },
  );

  return { text, tokens, counts, nearMisses, otherSpellings };
}
