import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { stripQuotedText, assertStripQuotedNotRaw } from './quote-strip.js';
import { InvalidInputError } from './coerce.js';

// The signal must always describe the transformation exactly: bytes removed = the
// difference between the input and the output measured in UTF-8.
function assertSignalMatches(input: string, result: { text: string; quotedBytesStripped: number }) {
  assert.equal(
    result.quotedBytesStripped,
    Buffer.byteLength(input, 'utf8') - Buffer.byteLength(result.text, 'utf8'),
    'quotedBytesStripped must equal the UTF-8 bytes actually removed',
  );
}

function strip(input: string) {
  const result = stripQuotedText(input);
  assertSignalMatches(input, result);
  return result;
}

describe('stripQuotedText — "> " quote runs', () => {
  it('removes a trailing quote run and its attribution from a top-posted reply', () => {
    const body = [
      'Confirmed for Thursday.',
      '',
      'On Mon, Jun 15, 2026 at 1:29 PM, Alice Example <alice@example.com> wrote:',
      '> Are we still on for Thursday?',
      '> Let me know.',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, 'Confirmed for Thursday.');
    assert.ok(quotedBytesStripped > 0);
  });

  it('removes nested quote depths in the same run', () => {
    const body = [
      'Latest answer.',
      '',
      'On Mon, Jun 15, 2026, Alice wrote:',
      '> On Sun, Jun 14, 2026, Bob wrote:',
      '>> Original question',
      '>>> Even deeper',
      '> Reply to Bob',
    ].join('\n');
    assert.equal(strip(body).text, 'Latest answer.');
  });

  it('tolerates blank lines inside a quote run without keeping them', () => {
    const body = ['New text.', '', 'On Mon, Alice wrote:', '> first', '', '> second'].join('\n');
    assert.equal(strip(body).text, 'New text.');
  });

  it('accepts a slightly indented quote marker', () => {
    const body = ['New text.', '', '  > quoted line'].join('\n');
    assert.equal(strip(body).text, 'New text.');
  });

  it('does NOT treat a ">" inside a line as a quote marker', () => {
    const body = 'The condition is a > b, always.\nStill mine.';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('keeps a bottom-posted reply written UNDER the quote', () => {
    const body = [
      'On Mon, Jun 15, 2026, Alice wrote:',
      '> Are we still on for Thursday?',
      '',
      'Yes, see you then.',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, 'Yes, see you then.');
    assert.ok(quotedBytesStripped > 0);
  });

  it('keeps every unquoted paragraph of an inline (interleaved) reply', () => {
    const body = [
      '> Can you review the draft?',
      '',
      'Done, comments inline.',
      '',
      '> And confirm the budget?',
      '',
      'Confirmed at 40k.',
    ].join('\n');
    assert.equal(strip(body).text, 'Done, comments inline.\n\nConfirmed at 40k.');
  });

  it('returns an empty body when the message is nothing but a quote', () => {
    const body = 'On Mon, Alice wrote:\n> everything here is quoted';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, '');
    assert.equal(quotedBytesStripped, Buffer.byteLength(body, 'utf8'));
  });
});

describe('stripQuotedText — attribution lines', () => {
  it('removes a wrapped attribution that spans two lines above the quote', () => {
    const body = [
      'Thanks.',
      '',
      'On Mon, Jun 15, 2026 at 1:29 PM Alice Example <alice@example.com>',
      'wrote:',
      '> original',
    ].join('\n');
    assert.equal(strip(body).text, 'Thanks.');
  });

  it('removes a bare "<name> wrote:" line when it sits directly above the quote', () => {
    const body = ['Thanks.', '', 'Alice Example wrote:', '> original'].join('\n');
    assert.equal(strip(body).text, 'Thanks.');
  });

  it('keeps an attribution-shaped line that is NOT above a quote block', () => {
    const body = 'I saw what Alice wrote:\nit was fine.';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('strips only the "wrote:" line when no "On ..." opener is within reach', () => {
    const body = [
      'Reply.',
      '',
      'This paragraph is my own writing and mentions nothing special.',
      'Alice wrote:',
      '> quoted',
    ].join('\n');
    assert.equal(
      strip(body).text,
      'Reply.\n\nThis paragraph is my own writing and mentions nothing special.',
    );
  });
});

describe('stripQuotedText — Outlook header block', () => {
  it('removes a From:/Sent:/To:/Subject: block and everything below it', () => {
    const body = [
      'See below.',
      '',
      '________________________________',
      'From: Alice Example <alice@example.com>',
      'Sent: Monday, 15 June 2026 13:29',
      'To: Jon Godley',
      'Subject: Re: Thursday',
      '',
      'Are we still on for Thursday?',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, 'See below.');
    assert.ok(quotedBytesStripped > 0);
  });

  it('needs the block, not one header line: a lone "From:" line is left alone', () => {
    const body = 'Quick note.\n\nFrom: the design review, we agreed on option B.';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('leaves a "From:" line with only one sibling header untouched', () => {
    const body = ['Note.', '', 'From: Alice', 'Subject: lunch', 'Nothing else here.'].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });
});

describe('stripQuotedText — "-----Original Message-----"', () => {
  it('removes the marker and everything below it', () => {
    const body = [
      'Forwarding my answer.',
      '',
      '-----Original Message-----',
      'From: Alice',
      'Sent: Monday',
      '',
      'The original text.',
    ].join('\n');
    assert.equal(strip(body).text, 'Forwarding my answer.');
  });

  it('matches the spacing and casing variants clients emit', () => {
    for (const marker of ['----- Original message -----', '--Original Message--', '-----ORIGINAL MESSAGE-----']) {
      const body = `Note.\n\n${marker}\nOld content.`;
      assert.equal(strip(body).text, 'Note.', `marker not recognised: ${marker}`);
    }
  });

  it('does not match the phrase inside a sentence', () => {
    const body = 'I attached the original message for reference.\nThanks.';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });
});

describe('stripQuotedText — no-marker passthrough and signal semantics', () => {
  it('returns the body byte-identical with a 0 signal when nothing matches', () => {
    const body = 'Hi Alice,\n\nThursday works for me.\n\nJon\n';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('does not trim trailing whitespace when nothing was stripped', () => {
    const body = 'Just this.\n\n\n';
    assert.equal(strip(body).text, body);
  });

  it('handles an empty body', () => {
    assert.deepEqual(stripQuotedText(''), { text: '', quotedBytesStripped: 0 });
  });

  it('counts UTF-8 bytes, not characters', () => {
    const quoted = '> café ☕ résumé';
    const body = `Réponse.\n\n${quoted}`;
    const { quotedBytesStripped } = strip(body);
    assert.equal(quotedBytesStripped, Buffer.byteLength(`\n\n${quoted}`, 'utf8'));
    assert.ok(quotedBytesStripped > `\n\n${quoted}`.length, 'multi-byte characters must count as more than one byte');
  });

  it('handles CRLF line endings', () => {
    const body = 'New text.\r\n\r\nOn Mon, Alice wrote:\r\n> quoted\r\n> more';
    assert.equal(strip(body).text, 'New text.');
  });

  it('keeps an unquoted wrap continuation of a quoted line (a known, deliberate limit)', () => {
    // Some clients wrap a long quoted line without re-prefixing it. Rather than delete an
    // unquoted line on suspicion, the fragment is kept and the surrounding quote goes.
    const body = ['Reply.', '', '> a very long quoted line that got', 'wrapped without a marker', '> next quoted line'].join('\n');
    assert.equal(strip(body).text, 'Reply.\nwrapped without a marker');
  });
});

describe('assertStripQuotedNotRaw', () => {
  it('rejects stripQuoted together with raw, naming both ways out', () => {
    assert.throws(
      () => assertStripQuotedNotRaw(true, true),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /stripQuoted cannot be combined with raw/);
        assert.match(err.message, /Drop raw/);
        return true;
      },
    );
  });

  it('allows either flag on its own', () => {
    assert.doesNotThrow(() => assertStripQuotedNotRaw(true, false));
    assert.doesNotThrow(() => assertStripQuotedNotRaw(false, true));
  });
});
