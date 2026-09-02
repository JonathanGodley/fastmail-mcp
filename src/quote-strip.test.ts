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

  it('leaves a pasted job posting alone — header-shaped lines with no address are not a quote', () => {
    // The over-strip case this rule is narrowed for: the block runs to the end of the
    // message, so firing on a paste would take the sender's own question with it.
    const body = [
      'Does this look worth applying for?',
      '',
      'From: The Hiring Team',
      'To: Candidates',
      'Subject: Senior Engineer, Remote',
      '',
      'We are looking for an engineer to join the platform team.',
      '',
      'What do you think?',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('leaves a pasted support mail whose From: carries no address', () => {
    const body = [
      'They sent me this, is it a phishing attempt?',
      '',
      'From: Account Support',
      'Sent: Yesterday',
      'Subject: Verify your account',
      '',
      'Click here to verify.',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('still fires on a bare address with no display name', () => {
    const body = ['Note.', '', 'From: alice@example.com', 'Sent: Monday', 'Subject: Re: x', '', 'quoted'].join('\n');
    assert.equal(strip(body).text, 'Note.');
  });

  it('needs the block, not one header line: a lone "From:" line is left alone', () => {
    const body = 'Quick note.\n\nFrom: the design review, we agreed on option B.';
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, body);
    assert.equal(quotedBytesStripped, 0);
  });

  it('leaves a "From:" line with only one sibling header untouched', () => {
    const body = ['Note.', '', 'From: Alice <alice@example.com>', 'Subject: lunch', 'Nothing else here.'].join('\n');
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

  it('keeps a FLUSH-LEFT wrap continuation of a quoted line (a known, deliberate limit)', () => {
    // Some clients wrap a long quoted line without re-prefixing it. A continuation that
    // starts at the left margin is indistinguishable from a terse inline reply written
    // straight under a quoted question, so the fragment is kept and the quote around it
    // goes. An INDENTED continuation is recognised and removed — see the run-continuation
    // block below.
    const body = ['Reply.', '', '> a very long quoted line that got', 'wrapped without a marker', '> next quoted line'].join('\n');
    assert.equal(strip(body).text, 'Reply.\nwrapped without a marker');
  });
});

// A quoted line whose wrap lost its ">" prefix used to end the run one line early, leaking
// the fragment into the kept output and repeating it once per quote depth (#181). The run
// now continues across such a line, but ONLY where it cannot be an inline reply: indented
// (not written at the left margin) and glued to quote lines above and below with no blank
// line either side. These tests hold that boundary from both directions.
describe('stripQuotedText — unprefixed continuations inside a quote run', () => {
  it('removes an unprefixed continuation at every quote depth it appears at', () => {
    const body = [
      'Hi,',
      '',
      'New reply text here.',
      '',
      'On Tue, Sep 1, 2026, at 9:00 AM, Someone Else wrote:',
      '> Quoted paragraph one.',
      '> ',
      ' <https://example.com/link>*Boilerplate that follows a wrapped link, unprefixed.',
      '> ',
      '> Sent from a phone',
      '>> On Aug 28, 2026, Another Person wrote:',
      '>> Quoted paragraph two.',
      '>> ',
      ' <https://example.com/link>*Boilerplate that follows a wrapped link, unprefixed.',
      '>> ',
      '>> Sent from a phone',
    ].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, 'Hi,\n\nNew reply text here.');
    assert.ok(quotedBytesStripped > 0);
  });

  it('keeps an indented paragraph that a blank line sets off from the quote', () => {
    // The reader pasted an indented note between two quoted questions. It is separated by
    // blank lines, which is what an interleaved reply looks like, so it stays.
    const body = [
      '> Can you review the draft?',
      '',
      '    Done, see the indented note pasted below.',
      '',
      '> And confirm the budget?',
      '',
      'Confirmed at 40k.',
    ].join('\n');
    assert.equal(
      strip(body).text,
      '    Done, see the indented note pasted below.\n\nConfirmed at 40k.',
    );
  });

  it('keeps a terse inline answer written flush left directly under each quoted question', () => {
    // No blank lines at all, so only the left margin distinguishes this from a wrapped
    // continuation. It is the sender's own new writing and must survive.
    const body = [
      '> Can you do Thursday?',
      'Yes.',
      '> And can you bring the figures?',
      'Yes, all of them.',
    ].join('\n');
    assert.equal(strip(body).text, 'Yes.\nYes, all of them.');
  });

  it('carries the run across a continuation that itself wrapped onto a second line', () => {
    const body = [
      'Reply.',
      '',
      'On Mon, Alice wrote:',
      '> a very long quoted line that got',
      '  wrapped without a marker, and the wrap',
      '  wrapped again',
      '> next quoted line',
    ].join('\n');
    assert.equal(strip(body).text, 'Reply.');
  });

  it('ends the run at a BLOCK of unprefixed lines — three is not a wrap', () => {
    const body = [
      'Reply.',
      '',
      '> Quoted opening.',
      '  first pasted line',
      '  second pasted line',
      '  third pasted line',
      '> Quoted close.',
    ].join('\n');
    assert.equal(
      strip(body).text,
      'Reply.\n  first pasted line\n  second pasted line\n  third pasted line',
    );
  });

  it('never absorbs an indented line with no quote line below it', () => {
    const body = ['On Mon, Alice wrote:', '> Are we still on?', '  Yes, indented for some reason.'].join('\n');
    assert.equal(strip(body).text, '  Yes, indented for some reason.');
  });
});

// These pin the OVER-strip direction: content that is not quoted correspondence but wears
// a quote marker. All are documented in docs/email-bodies.md and the README as accepted
// residuals — the caller's tell is quotedBytesStripped, and the remedy is re-reading
// without the flag. They are here so the behaviour is a pinned decision, not a surprise.
describe('stripQuotedText — documented over-strip residuals', () => {
  it('strips a markdown blockquote in the sender\'s own writing', () => {
    const body = ['The spec says:', '', '> the value MUST be a string', '', 'which we do not honour.'].join('\n');
    const { text, quotedBytesStripped } = strip(body);
    assert.equal(text, 'The spec says:\n\nwhich we do not honour.');
    assert.ok(quotedBytesStripped > 0, 'the signal is the caller\'s tell that this happened');
  });

  it('strips pasted shell/REPL output whose prompt is ">"', () => {
    const body = ['Repro:', '', '> npm test', '> 3 failing', '', 'Any ideas?'].join('\n');
    assert.equal(strip(body).text, 'Repro:\n\nAny ideas?');
  });

  it('can eat an INDENTED inline answer glued between two quoted lines', () => {
    // The residual the run-continuation rule (#181) accepts: an answer that is both
    // indented and pressed against the quote with no blank line either side is the same
    // shape as a wrapped continuation. Composers write flush left and set a reply off with
    // a blank line, so either habit alone keeps the text; this is the case with neither.
    const body = [
      '> Can you do Thursday?',
      '  Yes.',
      '> And bring the figures?',
      '',
      'Anything else?',
    ].join('\n');
    assert.equal(strip(body).text, 'Anything else?');
  });

  it('can eat a prose line pulled into a wrapped attribution above a real quote', () => {
    // The walk-back that catches Gmail's two-line attribution cannot tell this apart from
    // one. Narrowing it further would drop the wrapped-attribution case, which is common;
    // this is the accepted trade.
    const body = [
      'On Tuesday we agreed to postpone,',
      'which is not what the author wrote:',
      '> the original message',
      '',
      'Anyway.',
    ].join('\n');
    assert.equal(strip(body).text, 'Anyway.');
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
