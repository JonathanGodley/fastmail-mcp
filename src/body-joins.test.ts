import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildJoins, echoJoinLine, insertedBlocks, JOIN_LINE_ECHO_LIMIT, JOIN_MARKUP_ECHO_LIMIT } from './body-joins.js';
import type { InsertedBlock } from './body-joins.js';
import { expandBodyTokens } from './body-tokens.js';

/** Expand `authored` with the given blocks and return the stored text and where each landed. */
function expand(authored: string, signature = 'SIG', quote?: string): { text: string; blocks: InsertedBlock[] } {
  const out = expandBodyTokens(authored, {
    signature: { available: true, content: signature },
    quote: quote === undefined ? undefined : { available: true, content: quote },
    forward: undefined,
  });
  return { text: out.text, blocks: insertedBlocks(out.tokens) };
}

describe('body joins — text parts', () => {
  it('reports the start and end of the body, with blank runs at each', () => {
    const { text, blocks } = expand('{{signature}}\n\n');
    assert.deepEqual(buildJoins('textBody', text, blocks), [{
      token: 'signature', part: 'textBody',
      above: [], blankAbove: 0, block: ['SIG'], blankBelow: 2, below: [],
    }]);
  });

  it('counts a blank run that ends inside the block', () => {
    const { text, blocks } = expand('Hi\n{{signature}}', '\n\nSIG\n');
    const [join] = buildJoins('textBody', text, blocks);
    assert.equal(join.blankAbove, 2);
    assert.deepEqual(join.block, ['SIG']);
  });

  it('shifts every block by the offset it is given', () => {
    assert.deepEqual(
      insertedBlocks([{ name: 'signature', index: 0, text: '{{signature}}', expanded: true, output: { start: 2, end: 5 } }], 10),
      [{ token: 'signature', start: 12, end: 15 }],
    );
  });

  it('keeps a history block\'s line count to its own lines, and the caller line after it', () => {
    const { text, blocks } = expand('{{signature}}\n{{quote}}\nPS', 'SIG', 'q1\nq2\nq3');
    const [sig, quote] = buildJoins('textBody', text, blocks);
    assert.deepEqual(sig.below, [{ label: 'quoted original', lines: 3 }, 'PS']);
    assert.deepEqual(quote.block, [{ label: 'quoted original', lines: 3 }]);
    assert.deepEqual(quote.below, ['PS']);
  });

  it('redacts the text before the block whole, so a credential split across lines is caught', () => {
    const { text, blocks } = expand('auth Bearer\n\nabcdefghijklmnopqrstuvwxyz\n{{signature}}');
    assert.deepEqual(buildJoins('textBody', text, blocks)[0].above, ['auth Bearer [REDACTED]']);
  });

  it('bounds an echoed line at the named limit', () => {
    const line = echoJoinLine('y'.repeat(500));
    assert.equal([...line].length, JOIN_LINE_ECHO_LIMIT + 1);
    assert.ok(line.endsWith('…'));
  });
});

describe('body joins — html parts', () => {
  it('shows the markup either side (line breaks dropped), and the lines from each piece rendered on its own', () => {
    const { text, blocks } = expand('<p>One</p><p>Two</p><p>Three</p>\n<div>{{signature}}</div>\n<p>Four</p>', '<b>SIG</b>');
    assert.deepEqual(buildJoins('htmlBody', text, blocks), [{
      token: 'signature', part: 'htmlBody',
      above: ['Two', 'Three'], markupAbove: '</p><div>',
      block: ['SIG'],
      markupBelow: '</div><p>', below: ['Four'],
    }]);
  });

  it('runs to the opening of a comment the token sits in', () => {
    const { text, blocks } = expand('<p>Hi</p><!-- {{signature}} --><p>After</p>');
    const [join] = buildJoins('htmlBody', text, blocks);
    assert.equal(join.markupAbove, '</p><!-- ');
    assert.equal(join.markupBelow, ' --><p>');
  });

  it('stops a run at text, and at a bare < that opens no tag', () => {
    const { text, blocks } = expand('a < b {{signature}} c > d');
    const [join] = buildJoins('htmlBody', text, blocks);
    assert.equal(join.markupBelow, ' ');
    const { text: t2, blocks: b2 } = expand('x <p>{{signature}}< y');
    assert.equal(buildJoins('htmlBody', t2, b2)[0].markupBelow, '');
  });

  it('cuts a long run on the side away from the token, and redacts it first', () => {
    const wide = `<div style="${'a'.repeat(200)}">`;
    const { text, blocks } = expand(`<p>x</p>${wide}{{signature}}${wide}<p>y</p>`);
    const [join] = buildJoins('htmlBody', text, blocks);
    assert.ok(join.markupAbove!.startsWith('…') && join.markupAbove!.endsWith("a'>"), join.markupAbove);
    assert.equal([...join.markupAbove!].length, JOIN_MARKUP_ECHO_LIMIT + 1);
    assert.ok(join.markupBelow!.startsWith('<div') && join.markupBelow!.endsWith('…'), join.markupBelow);

    const secret = '<a title="Bearer not-a-token">';
    const { text: t2, blocks: b2 } = expand(`${secret}{{signature}}`);
    assert.equal(buildJoins('htmlBody', t2, b2)[0].markupAbove, "<a title='Bearer [REDACTED]");
  });

  it('says a block that renders no text renders none, and labels a history block by its rendered lines', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}{{quote}}', '<div><img src="https://example.com/x.png"></div>', '<p>q1</p><p>q2</p>');
    const [sig, quote] = buildJoins('htmlBody', text, blocks);
    assert.deepEqual(sig.block, []);
    assert.deepEqual(sig.below, [{ label: 'quoted original', lines: 2 }]);
    assert.equal(sig.markupBelow, '');
    assert.deepEqual(quote.above, ['Hi']);
  });

  it('shows an embedded image in the signature as [image]', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}', '<div><img src="cid:logo"></div>');
    assert.deepEqual(buildJoins('htmlBody', text, blocks)[0].block, ['[image]']);
  });
});
