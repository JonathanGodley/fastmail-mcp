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

  it('shows a history block\'s lines like any block\'s, beside a join and as its own', () => {
    const { text, blocks } = expand('{{signature}}\n{{quote}}\nPS', 'SIG', 'q1\nq2\nq3');
    const [sig, quote] = buildJoins('textBody', text, blocks);
    assert.deepEqual(sig.below, ['q1', 'q2']);
    assert.deepEqual(quote.above, ['SIG']);
    assert.deepEqual(quote.block, ['q1', 'q2', 'q3']);
    assert.deepEqual(quote.below, ['PS']);
  });

  it('shows only a long history block\'s edges, with the middle counted', () => {
    const { text, blocks } = expand('{{signature}}\n{{quote}}', 'SIG', 'q1\nq2\nq3\nq4\nq5\nq6');
    const [, quote] = buildJoins('textBody', text, blocks);
    assert.deepEqual(quote.block, ['q1', 'q2', 'q5', 'q6']);
    assert.equal(quote.blockOmitted, 2);
  });

  it('redacts, neutralises and bounds a history line, beside a join and in its own', () => {
    const hostile = 'A' + String.fromCharCode(0x202e) + '"B' + 'x'.repeat(200);
    const { text, blocks } = expand('{{signature}}\n{{quote}}', 'SIG', `token Bearer\nabcdefghijklmnopqrstuvwxyz\n${hostile}`);
    const [sig, quote] = buildJoins('textBody', text, blocks);
    assert.equal(sig.below[0], 'token Bearer [REDACTED]');
    assert.equal(sig.below[1], quote.block[quote.block.length - 1]);
    const last = quote.block[quote.block.length - 1];
    assert.ok(last.startsWith("A'B") && last.endsWith('…') && [...last].length === JOIN_LINE_ECHO_LIMIT + 1, last);
  });

  it('redacts the text before the block whole, so a credential split across lines is caught', () => {
    const { text, blocks } = expand('auth Bearer\n\nabcdefghijklmnopqrstuvwxyz\n{{signature}}');
    assert.deepEqual(buildJoins('textBody', text, blocks)[0].above, ['auth Bearer [REDACTED]']);
  });

  it('counts a blank run at the very start of the body', () => {
    const { text, blocks } = expand('\n\n{{signature}}');
    assert.equal(buildJoins('textBody', text, blocks)[0].blankAbove, 2);
  });

  it('places a line that ends where the block starts above it, and one that starts where it ends below it', () => {
    const { text, blocks } = expand('Before{{signature}}After', '\nSIG\n');
    assert.deepEqual(buildJoins('textBody', text, blocks), [{
      token: 'signature', part: 'textBody',
      above: ['Before'], blankAbove: 0, block: ['SIG'], blankBelow: 0, below: ['After'],
    }]);
  });

  it('counts an empty line at a history block\'s first character in the blank run beside it', () => {
    const { text, blocks } = expand('{{signature}}\n{{quote}}', 'SIG', '\nq1\n\nq2');
    const [sig] = buildJoins('textBody', text, blocks);
    assert.equal(sig.blankBelow, 1);
    assert.deepEqual(sig.below, ['q1', 'q2']);
  });

  it('shows no whitespace-only line of a history block', () => {
    const { text, blocks } = expand('{{quote}}', 'SIG', 'q1\n   \nq2');
    assert.deepEqual(buildJoins('textBody', text, blocks)[0].block, ['q1', 'q2']);
  });

  it('leaves an empty line starting where a history block ends to the caller', () => {
    const { text, blocks } = expand('{{quote}}\nPS', 'SIG', 'q1\n');
    const [quote] = buildJoins('textBody', text, blocks);
    assert.deepEqual(quote.block, ['q1']);
    assert.equal(quote.blankBelow, 1);
    assert.deepEqual(quote.below, ['PS']);
  });

  it('shows at most 2 lines below, and a 4-line block whole', () => {
    const { text, blocks } = expand('{{signature}}\na\nb\nc', 'L1\nL2\nL3\nL4');
    const [join] = buildJoins('textBody', text, blocks);
    assert.deepEqual(join.block, ['L1', 'L2', 'L3', 'L4']);
    assert.equal(join.blockOmitted, undefined);
    assert.deepEqual(join.below, ['a', 'b']);
  });

  it('places a block that renders no text between the lines either side of it', () => {
    const zwsp = String.fromCharCode(0x200b);
    const { text, blocks } = expand('Hi\n{{signature}}\nBye', zwsp);
    assert.deepEqual(buildJoins('textBody', text, blocks), [{
      token: 'signature', part: 'textBody',
      above: ['Hi'], blankAbove: 0, block: [], blankBelow: 1, below: ['Bye'],
    }]);
    // A block that is only a line break touches no line's content.
    const { text: t2, blocks: b2 } = expand('Hi{{signature}}Bye', '\n');
    const [join] = buildJoins('textBody', t2, b2);
    assert.deepEqual([join.above, join.block, join.below], [['Hi'], [], ['Bye']]);
  });

  it('redacts caller text and an earlier block together, so a credential across them is caught', () => {
    const stored = 'auth Bearer\nabcdefghijklmnopqrstuvwxyz\nSIG2';
    const first = stored.indexOf('abc');
    const blocks: InsertedBlock[] = [
      { token: 'signature', start: stored.indexOf('SIG2'), end: stored.length },
      { token: 'signature', start: first, end: first + 26 },
    ];
    const joins = buildJoins('textBody', stored, blocks);
    assert.deepEqual(joins.map((j) => j.block), [['abcdefghijklmnopqrstuvwxyz'], ['SIG2']], 'joins come back in body order');
    assert.deepEqual(joins[1].above, ['auth Bearer [REDACTED]']);
  });

  it('trims an echoed line', () => {
    assert.equal(echoJoinLine('   Regards  '), 'Regards');
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
    assert.ok(join.markupBelow!.startsWith('<div') && join.markupBelow!.endsWith('a…'), join.markupBelow);
    assert.equal([...join.markupBelow!].length, JOIN_MARKUP_ECHO_LIMIT + 1);

    const secret = '<a title="Bearer not-a-token">';
    const { text: t2, blocks: b2 } = expand(`${secret}{{signature}}`);
    assert.equal(buildJoins('htmlBody', t2, b2)[0].markupAbove, "<a title='Bearer [REDACTED]");
  });

  it('says a block that renders no text renders none, and shows a history block by its rendered lines', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}{{quote}}', '<div><img src="https://example.com/x.png"></div>', '<p>q1</p><p>q2</p>');
    const [sig, quote] = buildJoins('htmlBody', text, blocks);
    assert.deepEqual(sig.block, []);
    assert.deepEqual(sig.below, ['q1', 'q2']);
    assert.deepEqual(quote.block, ['q1', 'q2']);
    assert.equal(sig.markupBelow, '');
    assert.deepEqual(quote.above, ['Hi']);
  });

  it('keeps caller text and history in body order, each piece on its own line, at most 2 below', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}<p>mid</p>{{quote}}<p>after</p>', '<b>SIG</b>', '<p>q1</p><p>&nbsp;</p><p>q2</p>');
    const [sig, quote] = buildJoins('htmlBody', text, blocks);
    assert.deepEqual(sig.below, ['mid', 'q1']);
    assert.deepEqual(quote.above, ['SIG', 'mid']);
  });

  it('runs back past a closed comment, even one holding a <', () => {
    const { text, blocks } = expand('<p>x</p><!-- a<b --><div>{{signature}}');
    assert.equal(buildJoins('htmlBody', text, blocks)[0].markupAbove, '</p><!-- a<b --><div>');
  });

  it('runs forward over a comment holding a >, and stops at one never closed', () => {
    const { text, blocks } = expand('{{signature}}<!-- a>b --><p>x</p>');
    assert.equal(buildJoins('htmlBody', text, blocks)[0].markupBelow, '<!-- a>b --><p>');
    const { text: t2, blocks: b2 } = expand('{{signature}} <!-- never closed');
    assert.equal(buildJoins('htmlBody', t2, b2)[0].markupBelow, ' ');
  });

  it('runs to the ends of a comment the token sits in, though the comment holds a >', () => {
    const { text, blocks } = expand('<p>x</p><!-- a > b {{signature}} c > d --><p>y</p>');
    const [join] = buildJoins('htmlBody', text, blocks);
    assert.equal(join.markupAbove, '</p><!-- a > b ');
    assert.equal(join.markupBelow, ' c > d --><p>');
  });

  it('stops at a stray --> or > with nothing opening it, and at a tag never closed', () => {
    for (const [authored, above, below] of [
      ['a --><div>{{signature}}', '<div>', ''],
      ['a><div>{{signature}}<p', '<div>', ''],
      ['Hello{{signature}}', '', ''],
      ['<b>x</b> y{{signature}}', '', ''],
      ['a < b><div>{{signature}}< b>x', '<div>', ''],
      ['<!-- a --><!-- b -->{{signature}}<!--c--><!--d--><p>', '<!-- a --><!-- b -->', '<!--c--><!--d--><p>'],
    ]) {
      const { text, blocks } = expand(authored);
      const [join] = buildJoins('htmlBody', text, blocks);
      assert.deepEqual([join.markupAbove, join.markupBelow], [above, below], authored);
    }
  });

  it('ends a run at a prose > or < even when a real tag lies beyond it', () => {
    for (const [authored, above, below] of [
      ['<p>a -> {{signature}}</p>', ' ', '</p>'],
      ['<p>Thanks => </p>{{signature}}', ' </p>', ''],
      ['{{signature}}<a then <b>x', '', ''],
      ['<p>don\'t => {{signature}}', ' ', ''],
      ['{{signature}}xy> z', '', ''],
    ]) {
      const { text, blocks } = expand(authored);
      const [join] = buildJoins('htmlBody', text, blocks);
      assert.deepEqual([join.markupAbove, join.markupBelow], [above, below], authored);
    }
  });

  it('keeps a tag whole when a quoted attribute value holds a > or <', () => {
    const cases = [
      ['<p>Regards</p><div title="a>b">{{signature}}</div>', "</p><div title='a>b'>", '</div>'],
      ['<p>x</p>{{signature}}<span title="x<y">z</span>', '</p>', "<span title='x<y'>"],
      ["<b data-x='1>2' title=\"it's\">{{signature}}<i data-y='3<4'>q", "<b data-x='1>2' title='it's'>", "<i data-y='3<4'>"],
    ];
    const runs = cases.map(([authored]) => {
      const { text, blocks } = expand(authored);
      const [join] = buildJoins('htmlBody', text, blocks);
      return [authored, join.markupAbove, join.markupBelow];
    });
    assert.deepEqual(runs, cases);
  });

  it('knows the token is inside a tag when a quoted attribute value holds a >, on either side', () => {
    const cases = [
      ['<p>Hi</p><p title="a>b {{signature}}">x</p>', "</p><p title='a>b ", "'>", ['Hi'], ['x']],
      ['<p>Hi</p><p title="{{signature}} a>b">x</p>', "</p><p title='", " a>b'>", ['Hi'], ['x']],
      ["<!-- don't --><p>Hi {{signature}}</p>", ' ', '</p>', ['Hi'], []],
      ['<!-- <b title="x --><p>Hi {{signature}}</p>', ' ', '</p>', ['Hi'], []],
      ["<p>Hi</p><p title='a>b {{signature}}'>x</p>", "</p><p title='a>b ", "'>", ['Hi'], ['x']],
      ['<p>Hi</p><p {{signature}} title="a>b">x</p>', '</p><p ', " title='a>b'>", ['Hi'], ['x']],
      ["<p>Hi</p><p {{signature}} title='a>b'>x</p>", '</p><p ', " title='a>b'>", ['Hi'], ['x']],
      ['<p>Hi</p><a then <b title="{{signature}}">x</b>', " <b title='", "'>", ['Hi'], ['x']],
      ['<p>Hi</p><a then < b {{signature}} c', ' ', ' ', ['Hi'], ['c']],
      ['<p>x</p><!-- {{signature}} never closed', '</p><!-- ', ' never closed', ['x'], []],
    ] as const;
    const runs = cases.map(([authored]) => {
      const { text, blocks } = expand(authored);
      const [join] = buildJoins('htmlBody', text, blocks);
      return [authored, join.markupAbove, join.markupBelow, join.above, join.below];
    });
    assert.deepEqual(runs, cases);
  });

  it('shows no part of the tag or comment the token sits in as a display line', () => {
    const { text, blocks } = expand('<p>Hi</p><p title="{{signature}}">There</p>');
    const [join] = buildJoins('htmlBody', text, blocks);
    assert.deepEqual([join.above, join.below], [['Hi'], ['There']]);
    const { text: t2, blocks: b2 } = expand('<p>Hi</p><!-- {{signature}} --><p>x</p>');
    const [j2] = buildJoins('htmlBody', t2, b2);
    assert.deepEqual([j2.above, j2.below], [['Hi'], ['x']]);
  });

  it('shows no line for an html history block that renders nothing', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}{{quote}}', 'SIG', '<div><img src="https://example.com/x.png"></div>');
    assert.deepEqual(buildJoins('htmlBody', text, blocks)[0].below, []);
  });

  it('redacts a history block\'s rendering together with the caller text beside it', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}<p>auth Bearer</p>{{quote}}', 'SIG', '<p>abcdefghijklmnopqrstuvwxyz</p>');
    assert.deepEqual(buildJoins('htmlBody', text, blocks)[0].below, ['auth Bearer [REDACTED]']);
  });

  it('shows the rest of the segment when the tag the token sits in never closes', () => {
    const { text, blocks } = expand('<p title="{{signature}} tail');
    assert.equal(buildJoins('htmlBody', text, blocks)[0].markupBelow, ' tail');
  });

  it('shows a run of exactly the limit whole', () => {
    const exact = `<b x="${'a'.repeat(JOIN_MARKUP_ECHO_LIMIT - 8)}">`;
    assert.equal([...exact].length, JOIN_MARKUP_ECHO_LIMIT);
    const { text, blocks } = expand(`x${exact}{{signature}}`);
    assert.equal(buildJoins('htmlBody', text, blocks)[0].markupAbove, exact.replace(/"/g, "'"));
  });

  it('shows an embedded image in the signature as [image]', () => {
    const { text, blocks } = expand('<p>Hi</p>{{signature}}', '<div><img src="cid:logo"></div>');
    assert.deepEqual(buildJoins('htmlBody', text, blocks)[0].block, ['[image]']);
  });
});
