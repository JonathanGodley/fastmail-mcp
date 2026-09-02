import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { expandBodyTokens, scanBodyTokens } from './body-tokens.js';
import type { BodyBlock, BodyBlocks, BodyTokenName } from './body-tokens.js';

// Two writing conventions in this file, both load-bearing:
//
//  - every literal carrying a backslash is written with String.raw, so what the test says is
//    what the caller typed. The escape character of this grammar IS a backslash, and a '\\'
//    in an ordinary literal reads as two and means one.
//  - every invisible character is built from its char code, so no raw invisible character
//    lives in this source file and each one names the code point it is about (the same
//    convention reply-quote.test.ts uses for U+0085).

const available = (content: string): BodyBlock => ({ available: true, content });

/** All three tokens available, each expanding to a distinguishable block. */
const BLOCKS: BodyBlocks = {
  signature: available('[SIG]'),
  quote: available('[QUOTE]'),
  forward: available('[FWD]'),
};

/** All three tokens unavailable, each with a different cause. */
const NO_BLOCKS: BodyBlocks = {
  signature: { available: false, cause: 'no-signature' },
  quote: { available: false, cause: 'nothing-quotable' },
  forward: { available: false, cause: 'nothing-quotable-in-this-form' },
};

const NAMES: BodyTokenName[] = ['signature', 'quote', 'forward'];

const FEFF = String.fromCharCode(0xfeff);   // trimmed by trim(), and in the ECMAScript whitespace class
const ZWSP = String.fromCharCode(0x200b);   // NOT trimmed, NOT in that class
const NBSP = String.fromCharCode(0x00a0);   // no-break space
const NNBSP = String.fromCharCode(0x202f);  // narrow no-break space

describe('body tokens — the exact token spellings', () => {
  it('expands all three literal tokens', () => {
    const out = expandBodyTokens('a {{signature}} b {{quote}} c {{forward}} d', BLOCKS);
    assert.equal(out.text, 'a [SIG] b [QUOTE] c [FWD] d');
    assert.deepEqual(out.tokens.map((t) => t.name), ['signature', 'quote', 'forward']);
    assert.deepEqual(out.tokens.map((t) => t.expanded), [true, true, true]);
    assert.deepEqual(out.counts, { signature: 1, quote: 1, forward: 1 });
  });

  it('scans all three literal tokens with their positions, in landing order', () => {
    const part = 'x{{quote}}y{{signature}}';
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.tokens.map((t) => t.name), ['quote', 'signature']);
    assert.deepEqual(scan.tokens.map((t) => t.index), [part.indexOf('{{quote}}'), part.indexOf('{{signature}}')]);
    assert.deepEqual(scan.tokens.map((t) => t.text), ['{{quote}}', '{{signature}}']);
    assert.deepEqual(scan.counts, { signature: 1, quote: 1, forward: 0 });
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, []);
  });

  it("{{ signature }} is the token (the braces are trimmed with trim()'s own class)", () => {
    assert.equal(expandBodyTokens('a {{ signature }} b', BLOCKS).text, 'a [SIG] b');
    assert.equal(scanBodyTokens('a {{ signature }} b').counts.signature, 1);
    assert.deepEqual(scanBodyTokens('a {{ signature }} b').nearMisses, []);
  });

  it('U+FEFF inside the braces is trimmed, so {{<U+FEFF>signature}} IS the token', () => {
    const part = `a {{${FEFF}signature}} b`;
    assert.equal(expandBodyTokens(part, BLOCKS).text, 'a [SIG] b');
    const scan = scanBodyTokens(part);
    assert.equal(scan.counts.signature, 1);
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, []);
  });

  it('U+200B is NOT trimmed, so {{<U+200B>signature}} is prose and ships unchanged', () => {
    const part = `a {{${ZWSP}signature}} b`;
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part);
    assert.deepEqual(out.tokens, []);
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.nearMisses, []);
    // Prose, but a {{…}} spelling — reported so a caller can be told it was left unexpanded.
    assert.deepEqual(scan.otherSpellings.map((s) => s.text), [`{{${ZWSP}signature}}`]);
  });

  it('an NBSP and a U+202F inside the braces are trimmed too', () => {
    assert.equal(expandBodyTokens(`{{${NBSP}signature${NBSP}}}`, BLOCKS).text, '[SIG]');
    assert.equal(expandBodyTokens(`{{${NNBSP}quote${NNBSP}}}`, BLOCKS).text, '[QUOTE]');
  });

  it('counts a token repeated within one part', () => {
    const part = '{{quote}} then {{quote}} then {{quote}}';
    assert.equal(scanBodyTokens(part).counts.quote, 3);
    assert.equal(scanBodyTokens(part).tokens.length, 3);
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.counts.quote, 3);
    assert.equal(out.text, '[QUOTE] then [QUOTE] then [QUOTE]');
  });
});

describe('body tokens — near-miss spellings', () => {
  // The one refusal criterion: an unescaped run of braces on each side, two or more on at
  // least one side, whose contents trim case-insensitively to one of the three names.
  const NEAR_MISSES = ['{{Signature}}', '{{{signature}}}', '{{signature}}}', '{{signature}', '{signature}}'];

  for (const spelling of NEAR_MISSES) {
    it(`reports ${spelling} as a near-miss, naming the spelling the caller wrote`, () => {
      const part = `before ${spelling} after`;
      const scan = scanBodyTokens(part);
      assert.deepEqual(scan.tokens, [], `${spelling} must not expand`);
      assert.deepEqual(scan.counts, { signature: 0, quote: 0, forward: 0 });
      assert.equal(scan.nearMisses.length, 1);
      assert.equal(scan.nearMisses[0].text, spelling);
      assert.equal(scan.nearMisses[0].name, 'signature');
      assert.equal(scan.nearMisses[0].index, part.indexOf(spelling));
      // Nothing else claims it: a near-miss is not also an "other" spelling.
      assert.deepEqual(scan.otherSpellings, []);
      // Expansion refuses nothing and rewrites nothing here — the handler refuses upstream.
      assert.equal(expandBodyTokens(part, BLOCKS).text, part);
    });
  }

  it('names the token a near-miss of {{quote}} / {{forward}} near-missed', () => {
    assert.equal(scanBodyTokens('{{QUOTE}}').nearMisses[0].name, 'quote');
    assert.equal(scanBodyTokens('{{{forward}}}').nearMisses[0].name, 'forward');
  });

  it('{signature} is prose — a single brace each side is ordinary template syntax', () => {
    const part = 'the {signature} placeholder';
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, []);
    assert.equal(expandBodyTokens(part, BLOCKS).text, part);
  });

  it('{{signature}} beside {signature} still expands, and the {signature} is untouched', () => {
    const out = expandBodyTokens('{{signature}} and {signature}', BLOCKS);
    assert.equal(out.text, '[SIG] and {signature}');
    assert.equal(out.counts.signature, 1);
    assert.deepEqual(scanBodyTokens('{{signature}} and {signature}').nearMisses, []);
  });
});

describe('body tokens — the escape', () => {
  it(String.raw`\{{{signature}}} ships as {{{signature}}} with the backslash consumed`, () => {
    const out = expandBodyTokens(String.raw`see \{{{signature}}} there`, BLOCKS);
    assert.equal(out.text, 'see {{{signature}}} there');
    assert.deepEqual(out.tokens, []);
    assert.deepEqual(out.nearMisses, []);
  });

  it('an escaped {{signature}} ships as literal token text and expands nothing', () => {
    const part = String.raw`write \{{signature}} to mean the token`;
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, 'write {{signature}} to mean the token');
    assert.deepEqual(out.tokens, []);
    assert.deepEqual(scanBodyTokens(part).tokens, []);
  });

  it('an escaped {{SIGNATURE}} reaches NO near-miss report at all (one alternation, escape branch wins)', () => {
    const part = String.raw`\{{SIGNATURE}}`;
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.nearMisses, [], 'the escape branch consumed it before any refusal could see it');
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.otherSpellings, []);
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, '{{SIGNATURE}}');
    assert.deepEqual(out.nearMisses, []);
  });

  it(String.raw`\{{ item }} ships unchanged, backslash included (the escape is only for this grammar's spellings)`, () => {
    const part = String.raw`\{{ item }}`;
    assert.equal(expandBodyTokens(part, BLOCKS).text, part);
  });

  it('a backslash before single-brace PROSE escapes nothing, and is not eaten', () => {
    // One brace on each side was never a token and never a near-miss, so there is nothing here
    // for a backslash to escape. Both characters are the caller's and both ship.
    const part = String.raw`use \{quote} in your template`;
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part);
    assert.deepEqual(out.tokens, []);
    assert.deepEqual(out.nearMisses, []);
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.otherSpellings, []);
    assert.deepEqual(scan.nearMisses, []);
  });

  it('the same for spaced prose, while a near-miss on either side IS escapable', () => {
    assert.equal(expandBodyTokens(String.raw`\{ forward }`, BLOCKS).text, String.raw`\{ forward }`);
    assert.equal(expandBodyTokens(String.raw`\{signature}`, BLOCKS).text, String.raw`\{signature}`);
    // Two braces on ONE side makes it a near-miss spelling, which is something to escape: the
    // backslash goes, the braces stay, and no near-miss is reported.
    const out = expandBodyTokens(String.raw`\{signature}}`, BLOCKS);
    assert.equal(out.text, '{signature}}');
    assert.deepEqual(out.nearMisses, []);
    assert.equal(expandBodyTokens(String.raw`\{{signature}`, BLOCKS).text, '{{signature}');
  });

  it('two raw backslashes give one literal backslash followed by the literal token text', () => {
    // String.raw`\\…` is TWO raw backslashes. One is consumed by the escape; the other is
    // ordinary text, because a RUN of backslashes is not special.
    const out = expandBodyTokens(String.raw`\\{{signature}}`, BLOCKS);
    assert.equal(out.text, String.raw`\{{signature}}`);
    assert.deepEqual(out.tokens, [], 'the token text is literal, not a token');
  });
});

describe('body tokens — other {{…}} spellings are prose', () => {
  it('{{ item }} and {{#if}} pass through and are listed among the unexpanded spellings', () => {
    const part = 'hi {{ item }} and {{#if}} bye';
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part);
    assert.deepEqual(out.otherSpellings.map((s) => s.text), ['{{ item }}', '{{#if}}']);
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.otherSpellings.map((s) => s.text), ['{{ item }}', '{{#if}}']);
    assert.deepEqual(scan.otherSpellings.map((s) => s.index), [part.indexOf('{{ item }}'), part.indexOf('{{#if}}')]);
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.nearMisses, []);
  });

  it("a real token beside another system's spelling still expands", () => {
    const out = expandBodyTokens('{{#if x}}{{signature}}{{/if}}', BLOCKS);
    assert.equal(out.text, '{{#if x}}[SIG]{{/if}}');
    assert.equal(out.counts.signature, 1);
  });
});

describe('body tokens — substitution places the block and nothing else', () => {
  it('substitutes at the token position with no spacer added on either side', () => {
    assert.equal(expandBodyTokens('A{{signature}}B', BLOCKS).text, 'A[SIG]B');
    assert.equal(expandBodyTokens('{{signature}}', BLOCKS).text, '[SIG]');
  });

  it('a $-sequence in the block is inert (a function replacer, never a string replacement)', () => {
    const dollars = "$& $` $' $1";
    const out = expandBodyTokens('A{{signature}}B', { ...BLOCKS, signature: available(dollars) });
    assert.equal(out.text, `A${dollars}B`);
  });

  it('removes a token with nothing to expand to and reports the cause per site', () => {
    const out = expandBodyTokens('a {{signature}} b {{quote}} c {{forward}} d', NO_BLOCKS);
    assert.equal(out.text, 'a  b  c  d');
    assert.deepEqual(out.tokens.map((t) => t.expanded), [false, false, false]);
    assert.deepEqual(out.tokens.map((t) => t.cause), [
      'no-signature', 'nothing-quotable', 'nothing-quotable-in-this-form',
    ]);
    assert.deepEqual(out.counts, { signature: 1, quote: 1, forward: 1 });
  });

  it('reports the no-text-form cause verbatim (the fourth vocabulary word)', () => {
    const blocks: BodyBlocks = { ...BLOCKS, signature: { available: false, cause: 'no-text-form' } };
    const out = expandBodyTokens('note{{signature}}', blocks);
    assert.equal(out.text, 'note');
    assert.equal(out.tokens[0].cause, 'no-text-form');
  });

  it('removes a token whose block entry is absent, and marks it with no cause', () => {
    const out = expandBodyTokens('note {{forward}}', { ...BLOCKS, forward: undefined });
    assert.equal(out.text, 'note ');
    assert.equal(out.tokens.length, 1);
    assert.equal(out.tokens[0].expanded, false);
    assert.equal(out.tokens[0].cause, undefined);
  });

  it("reports each site's index in the AUTHORED part, so landing order survives expansion", () => {
    const part = 'one {{quote}} two {{signature}}';
    const out = expandBodyTokens(part, BLOCKS);
    assert.deepEqual(out.tokens.map((t) => t.index), [part.indexOf('{{quote}}'), part.indexOf('{{signature}}')]);
    assert.deepEqual(out.tokens.map((t) => t.text), ['{{quote}}', '{{signature}}']);
  });
});

describe('body tokens — the security rule: a block is never rescanned', () => {
  // The quote and forward blocks are built from an attacker-authored original. If expansion
  // ran as three passes, a {{signature}} inside somebody's email would land on the
  // REPLACEMENT side of the first pass and be expanded by the second — a stranger choosing
  // what goes into the user's outgoing message. One pass makes it inert.
  for (const outer of NAMES) {
    for (const inner of NAMES) {
      it(`expanding {{${outer}}} with a block containing the literal text {{${inner}}} leaves it unexpanded`, () => {
        const hostile = `hostile <${inner}> {{${inner}}} tail`;
        const blocks: BodyBlocks = { ...BLOCKS, [outer]: available(hostile) };
        const out = expandBodyTokens(`note {{${outer}}} end`, blocks);
        assert.equal(out.text, `note ${hostile} end`);
        assert.ok(out.text.includes(`{{${inner}}}`), 'the literal token text survives in the output');
        // Only the outer token was ever a token; the block's own text was never a site.
        assert.deepEqual(out.tokens.map((t) => t.name), [outer]);
        assert.equal(out.counts[inner], inner === outer ? 1 : 0);
      });
    }
  }

  it('a block carrying a near-miss or an escape is equally inert', () => {
    const hostile = String.raw`{{Signature}} and \{{signature}}`;
    const out = expandBodyTokens('{{quote}}', { ...BLOCKS, quote: available(hostile) });
    assert.equal(out.text, hostile);
    assert.deepEqual(out.nearMisses, [], 'the block is not scanned, so it reports nothing either');
  });
});

describe('body tokens — the scan is raw text: no tag stripping, no entity decoding', () => {
  it('a token split by a tag is not a token, and is reported as an unexpanded spelling', () => {
    const part = '<div>{{sig<b></b>nature}}</div>';
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part, 'stored exactly as the caller wrote it');
    assert.deepEqual(out.tokens, []);
    assert.deepEqual(out.otherSpellings.map((s) => s.text), ['{{sig<b></b>nature}}']);
    assert.deepEqual(scanBodyTokens(part).tokens, []);
  });

  it('a token spelled as entities is not a token and is reported nowhere', () => {
    const part = '&#123;&#123;signature&#125;&#125;';
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part);
    assert.deepEqual(out.tokens, []);
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, [], 'it carries no brace at all, so no branch of the grammar sees it');
  });

  it('a spelling containing another brace is reported nowhere — the bound that stops a swallow', () => {
    const part = 'hi {{a{b}}} bye';
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.text, part);
    assert.deepEqual(out.otherSpellings, [], 'the other-spelling report stops at the inner {');
    // Why that bound is worth an unreported piece of prose: an other-spelling branch that could
    // cross a `{` would match `{{a{{signature}}` whole and swallow the real token inside it.
    const nested = expandBodyTokens('{{a{{signature}}', BLOCKS);
    assert.equal(nested.text, '{{a[SIG]', 'the inner token is seen and expanded, not hidden inside an outer match');
    assert.equal(nested.counts.signature, 1);
  });
});

describe('body tokens — a long run of braces is scanned in linear time', () => {
  it('100,000 unmatched braces scan promptly and match nothing', () => {
    // Branch 2's `(?<!\{)` is the bound this pins. Without it the engine restarts that branch
    // at every brace in the run and re-walks the remainder, which is quadratic: measured on
    // this machine at 7ms for 2,000 braces, 111ms for 8,000, 733ms for 20,000, and about 18
    // SECONDS for the 100,000 below — off one authored body. It is now under a millisecond.
    //
    // Wall-clock is the only thing that observes this: the classification is identical either
    // way, which is exactly why the defect could sit here unnoticed. So the budget is
    // deliberately enormous — a thousand times the measured cost, and still an order of
    // magnitude below the defect, so a slow or loaded machine cannot make it flaky while a
    // reintroduced quadratic cannot slip under it.
    const part = '{'.repeat(100_000);
    const started = Date.now();
    const scan = scanBodyTokens(part);
    const elapsed = Date.now() - started;
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, []);
    assert.ok(elapsed < 2000, `scanning 100,000 braces took ${elapsed}ms`);
  });

  it('a real token after a long run of braces is still found and expanded', () => {
    // The bound must not cost a match: it only refuses to START inside a run, and a token
    // after the run does not start inside it.
    const part = `${'{'.repeat(50_000)}\n{{signature}}`;
    const out = expandBodyTokens(part, BLOCKS);
    assert.equal(out.counts.signature, 1);
    assert.ok(out.text.endsWith('\n[SIG]'), out.text.slice(-20));
  });
});

// The two shapes edit_draft needs and no compose surface does: a scan that reports what the
// caller escaped, and a block form that leaves a token exactly as typed.
//
// Both exist because edit_draft stores an unflagged body byte for byte. It has to be able to
// tell the caller what the stored part will contain, and it must never delete a spelling it
// was not asked to expand.
describe('body tokens — escapes are reported as facts, not defects', () => {
  it('lists each escaped spelling with the backslash the caller typed', () => {
    const part = String.raw`a \{{signature}} and a \{{quote}} here`;
    const scan = scanBodyTokens(part);
    assert.deepEqual(scan.escapes.map((e) => e.text), [String.raw`\{{signature}}`, String.raw`\{{quote}}`]);
    // An escape is neither a token nor a near-miss nor an unexpanded spelling.
    assert.deepEqual(scan.tokens, []);
    assert.deepEqual(scan.nearMisses, []);
    assert.deepEqual(scan.otherSpellings, []);
    assert.equal(scan.counts.signature, 0);
  });

  it('records an escaped NEAR-MISS spelling too, where no name is available to report', () => {
    // Branch 1 does not capture a name, which is why `escapes` is a site list and not a
    // count: the only thing that can be quoted back is the literal text.
    const scan = scanBodyTokens(String.raw`\{{SIGNATURE}}`);
    assert.deepEqual(scan.escapes.map((e) => e.text), [String.raw`\{{SIGNATURE}}`]);
    assert.deepEqual(scan.nearMisses, []);
  });

  it('reports nothing for an unescaped token or for prose braces', () => {
    assert.deepEqual(scanBodyTokens('{{signature}}').escapes, []);
    assert.deepEqual(scanBodyTokens(String.raw`use \{quote} in a template`).escapes, []);
  });
});

describe('body tokens — the as-written block leaves a token exactly as typed', () => {
  const AS_WRITTEN: BodyBlocks = {
    signature: available('[SIG]'),
    quote: { available: 'as-written' },
    forward: { available: 'as-written' },
  };

  it('keeps the spelling in the text and marks the site asWritten', () => {
    const out = expandBodyTokens('sign here {{signature}} and quote here {{quote}}', AS_WRITTEN);
    assert.equal(out.text, 'sign here [SIG] and quote here {{quote}}');
    const quoteSite = out.tokens.find((t) => t.name === 'quote')!;
    assert.equal(quoteSite.asWritten, true);
    assert.equal(quoteSite.expanded, false);
    assert.equal(quoteSite.cause, undefined); // nothing was unavailable, so nothing to explain
    assert.equal(quoteSite.text, '{{quote}}');
  });

  it('still counts the site, so a caller can say the token expanded nowhere', () => {
    const out = expandBodyTokens('{{forward}} {{forward}}', AS_WRITTEN);
    assert.equal(out.counts.forward, 2);
    assert.equal(out.tokens.filter((t) => t.asWritten).length, 2);
    assert.equal(out.text, '{{forward}} {{forward}}');
  });

  // The distinction that matters: an unavailable block REMOVES the token (and says why),
  // while as-written keeps it. A handler that stores the body byte for byte needs the second
  // one — removal would be this server silently editing the caller's body.
  it('is not the same answer as an unavailable block', () => {
    const removed = expandBodyTokens('x {{quote}} y', NO_BLOCKS);
    assert.equal(removed.text, 'x  y');
    assert.equal(removed.tokens[0].cause, 'nothing-quotable');
    assert.equal(removed.tokens[0].asWritten, undefined);
  });

  // And not the same as no block at all, which also removes — the form edit_draft must never
  // reach for on a token it does not offer.
  it('is not the same answer as an absent block', () => {
    const out = expandBodyTokens('x {{quote}} y', { signature: available('[SIG]'), quote: undefined, forward: undefined });
    assert.equal(out.text, 'x  y');
    assert.equal(out.tokens[0].asWritten, undefined);
  });
});
