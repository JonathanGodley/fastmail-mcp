import { createHash } from 'node:crypto';
import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { foldICalLine } from './ical-fold.js';

describe('foldICalLine', () => {
  it('returns short lines unchanged', () => {
    assert.equal(foldICalLine('SUMMARY:Short'), 'SUMMARY:Short');
  });

  it('folds lines longer than 75 octets', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(Buffer.byteLength(lines[0], 'utf8') <= 75);
    assert.ok(lines[1].startsWith(' '));
  });

  it('folds very long lines into multiple segments', () => {
    const long = 'DESCRIPTION:' + 'y'.repeat(200);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(lines.length >= 3);
    for (let i = 1; i < lines.length; i++) {
      assert.ok(lines[i].startsWith(' '));
    }
  });

  it('keeps every segment within 75 octets', () => {
    const long = 'DESCRIPTION:' + 'z'.repeat(200);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    for (const line of lines) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75);
    }
  });

  it('folds multi-byte characters without exceeding 75 octets', () => {
    const long = 'LOCATION:' + '📍'.repeat(20);
    const folded = foldICalLine(long);
    const lines = folded.split('\r\n');
    assert.ok(lines.length >= 2);
    for (const line of lines) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75,
        `Line exceeds 75 octets: ${Buffer.byteLength(line, 'utf8')} bytes`);
    }
  });

  it('does not fold a line of exactly 75 octets, but folds one of 76', () => {
    const line75 = 'X'.repeat(75);
    assert.equal(foldICalLine(line75), line75);
    const line76 = 'X'.repeat(76);
    assert.ok(foldICalLine(line76).includes('\r\n'));
  });

  it('fills each ASCII segment to exactly 75 octets', () => {
    assert.equal(foldICalLine('X'.repeat(80)), 'X'.repeat(75) + '\r\n ' + 'X'.repeat(5));
  });

  it('moves a whole surrogate pair to the next segment at both ends of the low-surrogate range', () => {
    // 72 octets of ASCII plus a lone high surrogate (3 octets) is exactly 75, so the cut lands
    // between the two halves of the pair.
    for (const ch of ['\u{10000}', '\u{10FFFF}']) {
      assert.equal(foldICalLine('X'.repeat(72) + ch), 'X'.repeat(72) + '\r\n ' + ch);
    }
  });

  it('never leaves a lone surrogate when the 75-octet cut falls inside a surrogate pair', () => {
    function unfold(folded: string): string {
      const lines = folded.split('\r\n');
      let out = lines[0];
      for (let i = 1; i < lines.length; i++) {
        out += lines[i].startsWith(' ') ? lines[i].slice(1) : lines[i];
      }
      return out;
    }
    function hasLoneSurrogate(s: string): boolean {
      for (let i = 0; i < s.length; i++) {
        const code = s.charCodeAt(i);
        if (code >= 0xD800 && code <= 0xDBFF) {
          const next = s.charCodeAt(i + 1);
          if (!(next >= 0xDC00 && next <= 0xDFFF)) return true;
          i++;
        } else if (code >= 0xDC00 && code <= 0xDFFF) {
          return true;
        }
      }
      return false;
    }
    const cases = [
      'X'.repeat(8) + '📍'.repeat(20),
      'X'.repeat(10) + '\u{10FFFF}'.repeat(15),
      'X'.repeat(10) + '🐀'.repeat(20),
      'X'.repeat(72) + '！'.repeat(5),
      // Only this prefix of 3 lands the 75-octet cut inside a surrogate pair; none of the four
      // cases above does.
      'X'.repeat(3) + '📍'.repeat(20),
    ];
    for (const body of cases) {
      const input = 'LOCATION:' + body;
      const folded = foldICalLine(input);
      const lines = folded.split('\r\n');
      for (const line of lines) {
        assert.ok(!hasLoneSurrogate(line), `lone surrogate in ${JSON.stringify(line)}`);
        assert.equal(Buffer.from(line, 'utf8').toString('utf8'), line, `piece does not round-trip: ${JSON.stringify(line)}`);
      }
      assert.equal(unfold(folded), input);
    }
  });
});

describe('foldICalLine with custom line ending', () => {
  it('uses LF when specified', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long, '\n');
    assert.ok(!folded.includes('\r'));
    assert.ok(folded.includes('\n'));
  });

  it('defaults to CRLF', () => {
    const long = 'DESCRIPTION:' + 'x'.repeat(80);
    const folded = foldICalLine(long);
    assert.ok(folded.includes('\r\n'));
  });
});

describe('foldICalLine exact output', () => {
  const pin = '\u{1F4CD}';

  it('cuts an ASCII line at 75 octets, then 74 after each continuation space', () => {
    assert.equal(foldICalLine('A'.repeat(160)),
      'A'.repeat(75) + '\r\n ' + 'A'.repeat(74) + '\r\n ' + 'A'.repeat(11));
  });

  it('cuts before a 2-octet character that would pass 75', () => {
    assert.equal(foldICalLine('X'.repeat(74) + 'é'.repeat(3)),
      'X'.repeat(74) + '\r\n ' + 'é'.repeat(3));
  });

  it('cuts before a 3-octet character that would pass 75', () => {
    assert.equal(foldICalLine('X'.repeat(73) + '！'.repeat(2)),
      'X'.repeat(73) + '\r\n ' + '！'.repeat(2));
  });

  it('moves a surrogate pair straddling 75 octets on a continuation line', () => {
    assert.equal(foldICalLine('X'.repeat(146) + pin + 'Y'),
      'X'.repeat(75) + '\r\n ' + 'X'.repeat(71) + '\r\n ' + pin + 'Y');
  });

  it('leaves 75 octets alone and folds 76', () => {
    assert.equal(foldICalLine('é'.repeat(37) + 'X'), 'é'.repeat(37) + 'X');
    assert.equal(foldICalLine('X'.repeat(76)), 'X'.repeat(75) + '\r\n X');
  });

  it('returns an empty string unchanged', () => {
    assert.equal(foldICalLine(''), '');
  });

  it('counts a lone high surrogate as 3 octets', () => {
    assert.equal(foldICalLine('X'.repeat(73) + '\uD83D' + 'YZ'),
      'X'.repeat(73) + '\r\n \uD83DYZ');
    assert.equal(foldICalLine('X'.repeat(72) + '\uD83D' + 'YZ'),
      'X'.repeat(72) + '\uD83D\r\n YZ');
  });

  it('steps back one unit when a lone low surrogate sits at the cut', () => {
    assert.equal(foldICalLine('X'.repeat(73) + '\uDC00' + 'YZ'),
      'X'.repeat(72) + '\r\n X\uDC00YZ');
    // The step back lands inside the preceding pair, so the continuation opens with an orphaned
    // low surrogate, which counts 3 octets there.
    assert.equal(foldICalLine('X'.repeat(69) + pin + '\uDC00' + 'Y'.repeat(80)),
      'X'.repeat(69) + '\uD83D\r\n \uDCCD\uDC00' + 'Y'.repeat(68) + '\r\n ' + 'Y'.repeat(12));
  });

  it('folds a long mixed line to a pinned output', () => {
    const alphabet = ['a', 'B', 'é', 'Ω', '！', '中', pin, '\u{10FFFF}'];
    let seed = 12345;
    let input = 'DESCRIPTION:';
    for (let i = 0; i < 600; i++) {
      seed = (seed * 48271) % 2147483647;
      input += alphabet[seed % alphabet.length];
    }
    const folded = foldICalLine(input);
    const lines = folded.split('\r\n');
    for (const line of lines) assert.ok(Buffer.byteLength(line, 'utf8') <= 75);
    assert.equal(lines[0] + lines.slice(1).map((l) => l.slice(1)).join(''), input);
    assert.equal(createHash('sha256').update(folded).digest('hex'), '55b7b29b23f2901b7d1eb856fc2de3381ef305cccfe3aa58c26c63de6aecd18d');
  });
});
