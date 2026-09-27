import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { trimEnd, isWhitespace } from './trim-end.js';

describe('trimEnd', () => {
  it('removes only the trailing run the predicate accepts', () => {
    assert.equal(trimEnd('a b  ', (ch) => ch === ' '), 'a b');
    assert.equal(trimEnd('  a', (ch) => ch === ' '), '  a');
    assert.equal(trimEnd('', (ch) => ch === ' '), '');
  });

  it('stops at the start of the text, never asking the predicate past it', () => {
    const asked: unknown[] = [];
    const result = trimEnd('ab', (ch) => { asked.push(ch); return asked.length <= 3; });
    assert.equal(result, '');
    assert.deepEqual(asked, ['b', 'a']);
  });
});

describe('isWhitespace', () => {
  it('matches what \\s matches', () => {
    for (const ch of [' ', '\t', '\n', '\r', ' ', ' ']) assert.equal(isWhitespace(ch), true, JSON.stringify(ch));
    for (const ch of ['a', '.', '​']) assert.equal(isWhitespace(ch), false, JSON.stringify(ch));
  });
});
