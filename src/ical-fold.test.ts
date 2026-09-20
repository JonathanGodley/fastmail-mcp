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
