import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { InvalidInputError } from './coerce.js';
import {
  NOTE_BODY_EDITS_DISCARDED_TEXT_PART, REJECT_BODY_EDITS_NO_BODY, REJECT_BODY_EDITS_WITH_BODY,
  REJECT_BODY_EDITS_WITHOUT_SIGNATURE, locateBodyEdits, noteBodyEditsSplitSignature, spliceBodyEdits,
  unmatchedSegments,
} from './body-edits.js';

/** Asserts `fn` throws `type` with a message ending in `message`. */
function refuses(fn: () => unknown, type: typeof InvalidInputError, message: string) {
  assert.throws(fn, (err: unknown) => {
    assert.ok(err instanceof type, `expected ${type.name}, got ${String(err)}`);
    assert.ok((err as Error).message.endsWith(message), `message was: ${(err as Error).message}`);
    return true;
  });
}

describe('locateBodyEdits', () => {
  it('locates a single exact occurrence', () => {
    assert.deepEqual(locateBodyEdits('Hello world', [{ find: 'world', replace: 'there' }], 'textBody'), [
      { offset: 6, matchedSize: 5 },
    ]);
  });

  it('locates a match at the very start and at the very end', () => {
    assert.deepEqual(
      locateBodyEdits('abcdef', [{ find: 'ef', replace: '' }, { find: 'ab', replace: '' }], 'textBody'),
      [{ offset: 4, matchedSize: 2 }, { offset: 0, matchedSize: 2 }],
    );
  });

  it('refuses an empty find before looking for any other op', () => {
    refuses(
      () => locateBodyEdits('abc', [{ find: 'zzz', replace: '' }, { find: '', replace: 'x' }], 'htmlBody'),
      InvalidInputError,
      'bodyEdits[1].find is empty; give the exact text to replace.',
    );
  });

  it('refuses a find that is not there, naming the op and the part', () => {
    refuses(
      () => locateBodyEdits('abc', [{ find: 'a', replace: '' }, { find: 'zzz', replace: '' }], 'htmlBody'),
      InvalidInputError,
      'bodyEdits[1].find not found; re-read the draft and copy the text exactly as its htmlBody stores it (no entity decoding or whitespace normalisation is applied).',
    );
  });

  it('reports the first failing op in caller order', () => {
    refuses(
      () => locateBodyEdits('aXa', [{ find: 'a', replace: '' }, { find: 'zzz', replace: '' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0].find occurs 2 times in textBody: ambiguous; include more context so it matches exactly once.',
    );
  });

  it('refuses two non-overlapping occurrences as ambiguous, with the count', () => {
    refuses(
      () => locateBodyEdits('one two one two one', [{ find: 'one', replace: '1' }], 'htmlBody'),
      InvalidInputError,
      'bodyEdits[0].find occurs 3 times in htmlBody: ambiguous; include more context so it matches exactly once.',
    );
  });

  it('counts overlapping self-occurrences', () => {
    refuses(
      () => locateBodyEdits('aaa', [{ find: 'aa', replace: 'b' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0].find occurs 2 times in textBody: ambiguous; include more context so it matches exactly once.',
    );
  });

  it('matches raw: an entity is not its character', () => {
    const html = '<p>Tom &amp; Jerry</p>';
    refuses(
      () => locateBodyEdits(html, [{ find: 'Tom & Jerry', replace: 'x' }], 'htmlBody'),
      InvalidInputError,
      'bodyEdits[0].find not found; re-read the draft and copy the text exactly as its htmlBody stores it (no entity decoding or whitespace normalisation is applied).',
    );
    assert.deepEqual(locateBodyEdits(html, [{ find: 'Tom &amp; Jerry', replace: 'x' }], 'htmlBody'), [
      { offset: 3, matchedSize: 15 },
    ]);
  });

  it('matches raw: whitespace is not normalised', () => {
    refuses(
      () => locateBodyEdits('see\n  you', [{ find: 'see you', replace: 'x' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0].find not found; re-read the draft and copy the text exactly as its textBody stores it (no entity decoding or whitespace normalisation is applied).',
    );
  });

  it('refuses two ops whose matches intersect, naming both', () => {
    refuses(
      () => locateBodyEdits('abcdef', [{ find: 'f', replace: '' }, { find: 'abc', replace: '' }, { find: 'cd', replace: '' }], 'htmlBody'),
      InvalidInputError,
      'bodyEdits[1] and bodyEdits[2] match overlapping text in htmlBody; combine them into one edit.',
    );
  });

  it('refuses one op whose match contains another', () => {
    refuses(
      () => locateBodyEdits('abcdef', [{ find: 'bcd', replace: '' }, { find: 'abcde', replace: '' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0] and bodyEdits[1] match overlapping text in textBody; combine them into one edit.',
    );
  });

  it('refuses identical finds as overlapping', () => {
    refuses(
      () => locateBodyEdits('abc', [{ find: 'b', replace: 'x' }, { find: 'b', replace: 'y' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0] and bodyEdits[1] match overlapping text in textBody; combine them into one edit.',
    );
  });

  it('allows matches that touch', () => {
    assert.deepEqual(
      locateBodyEdits('abcd', [{ find: 'cd', replace: '' }, { find: 'ab', replace: '' }], 'textBody'),
      [{ offset: 2, matchedSize: 2 }, { offset: 0, matchedSize: 2 }],
    );
    assert.deepEqual(
      locateBodyEdits('abcd', [{ find: 'ab', replace: '' }, { find: 'cd', replace: '' }], 'textBody'),
      [{ offset: 0, matchedSize: 2 }, { offset: 2, matchedSize: 2 }],
    );
  });

  it('refuses a match that overlaps by one unit at either edge', () => {
    refuses(
      () => locateBodyEdits('abcd', [{ find: 'abc', replace: '' }, { find: 'cd', replace: '' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0] and bodyEdits[1] match overlapping text in textBody; combine them into one edit.',
    );
    refuses(
      () => locateBodyEdits('abcd', [{ find: 'cd', replace: '' }, { find: 'abc', replace: '' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[0] and bodyEdits[1] match overlapping text in textBody; combine them into one edit.',
    );
  });

  it('locates against the original: a find that only an earlier replacement creates is not found', () => {
    refuses(
      () => locateBodyEdits('Hi Bob', [{ find: 'Bob', replace: 'Carol' }, { find: 'Carol', replace: 'Dave' }], 'textBody'),
      InvalidInputError,
      'bodyEdits[1].find not found; re-read the draft and copy the text exactly as its textBody stores it (no entity decoding or whitespace normalisation is applied).',
    );
  });

  it('locates against the original: a replacement that repeats a later find does not make it ambiguous', () => {
    assert.deepEqual(
      locateBodyEdits('Hi Bob, from Al', [{ find: 'Bob', replace: 'Bob and Al' }, { find: 'Al', replace: 'Alice' }], 'textBody'),
      [{ offset: 3, matchedSize: 3 }, { offset: 13, matchedSize: 2 }],
    );
  });

  it('measures offsets and sizes in UTF-16 code units', () => {
    const face = String.fromCodePoint(0x1f600); // two code units
    const stored = `caf${String.fromCharCode(0xe9)} ${face} end`;
    assert.deepEqual(locateBodyEdits(stored, [{ find: `${face} end`, replace: '' }], 'textBody'), [
      { offset: 5, matchedSize: 6 },
    ]);
  });
});

describe('spliceBodyEdits', () => {
  it('applies every op at its original offset, whatever the caller order', () => {
    const stored = 'Hi Bob, from Al';
    const ops = [{ find: 'Al', replace: 'Alice' }, { find: 'Bob', replace: 'Bob and Al' }];
    const located = locateBodyEdits(stored, ops, 'textBody');
    assert.equal(spliceBodyEdits(stored, located, ops.map((o) => o.replace)), 'Hi Bob and Al, from Alice');
  });

  it('deletes with an empty replacement and handles touching matches at both ends', () => {
    const stored = 'abcdef';
    const located = [{ offset: 4, matchedSize: 2 }, { offset: 0, matchedSize: 2 }, { offset: 2, matchedSize: 2 }];
    assert.equal(spliceBodyEdits(stored, located, ['', 'X', 'YY']), 'XYY');
  });

  it('splices whichever replacements it is handed', () => {
    const located = [{ offset: 1, matchedSize: 1 }];
    assert.equal(spliceBodyEdits('abc', located, ['[expanded]']), 'a[expanded]c');
  });

  it('keeps the untouched text byte for byte', () => {
    const stored = 'x &amp;  \r\n y';
    assert.equal(spliceBodyEdits(stored, [{ offset: 0, matchedSize: 1 }], ['z']), 'z &amp;  \r\n y');
  });
});

describe('bodyEdits message constants', () => {
  it('are distinct, non-empty strings', () => {
    const all = [
      REJECT_BODY_EDITS_WITH_BODY, NOTE_BODY_EDITS_DISCARDED_TEXT_PART,
      REJECT_BODY_EDITS_WITHOUT_SIGNATURE, REJECT_BODY_EDITS_NO_BODY,
    ];
    for (const message of all) assert.ok(typeof message === 'string' && message.trim() !== '');
    assert.equal(new Set(all).size, all.length);
  });
});

describe('unmatchedSegments', () => {
  it('returns the text around and between the matches, in stored order', () => {
    assert.deepEqual(
      unmatchedSegments('abcdefgh', [{ offset: 5, matchedSize: 2 }, { offset: 1, matchedSize: 2 }]),
      ['a', 'de', 'h'],
    );
  });

  it('returns an empty segment where matches touch or reach an end', () => {
    assert.deepEqual(
      unmatchedSegments('abcd', [{ offset: 0, matchedSize: 2 }, { offset: 2, matchedSize: 2 }]),
      ['', '', ''],
    );
  });
});

describe('noteBodyEditsSplitSignature', () => {
  it('says one token in the singular', () => {
    assert.equal(
      noteBodyEditsSplitSignature('htmlBody', 1),
      '1 {{signature}} token in htmlBody was formed where a bodyEdits replace meets the text beside it, ' +
        'so it is stored as literal text: only a token wholly inside a replace expands. Put the whole {{signature}} in one replace.',
    );
  });

  it('says several tokens in the plural', () => {
    assert.equal(
      noteBodyEditsSplitSignature('textBody', 2),
      '2 {{signature}} tokens in textBody were formed where a bodyEdits replace meets the text beside it, ' +
        'so they are stored as literal text: only a token wholly inside a replace expands. Put the whole {{signature}} in one replace.',
    );
  });
});
