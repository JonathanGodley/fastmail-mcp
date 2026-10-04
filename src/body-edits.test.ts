import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { McpError } from '@modelcontextprotocol/sdk/types.js';
import { InvalidInputError } from './coerce.js';
import {
  coerceBodyEdits, formatBodyEditsReceipt, locateBodyEdits, noteBodyEditsSplitSignature, spliceBodyEdits,
} from './body-edits.js';

/** Asserts `fn` throws `type` with exactly `message` (McpError prefixes its code). */
function refuses(fn: () => unknown, type: typeof McpError | typeof InvalidInputError, message: string) {
  assert.throws(fn, (err: unknown) => {
    assert.ok(err instanceof type, `expected ${type.name}, got ${String(err)}`);
    assert.ok((err as Error).message.endsWith(message), `message was: ${(err as Error).message}`);
    return true;
  });
}

describe('coerceBodyEdits', () => {
  it('reads an absent, null or blank value as omitted', () => {
    assert.equal(coerceBodyEdits(undefined), undefined);
    assert.equal(coerceBodyEdits(null), undefined);
    assert.equal(coerceBodyEdits(''), undefined);
    assert.equal(coerceBodyEdits('  \n '), undefined);
  });

  it('passes a well-formed array through', () => {
    assert.deepEqual(coerceBodyEdits([{ find: 'a', replace: 'b' }, { find: 'c', replace: '' }]), [
      { find: 'a', replace: 'b' },
      { find: 'c', replace: '' },
    ]);
  });

  it('never trims find or replace', () => {
    assert.deepEqual(coerceBodyEdits([{ find: ' a\n', replace: '\tb ' }]), [{ find: ' a\n', replace: '\tb ' }]);
    assert.deepEqual(coerceBodyEdits(['  {"find":" x ","replace":"  "}  ']), [{ find: ' x ', replace: '  ' }]);
  });

  it('accepts a JSON-encoded array, padded', () => {
    assert.deepEqual(coerceBodyEdits(' [{"find":"a","replace":"b"}] '), [{ find: 'a', replace: 'b' }]);
  });

  it('accepts JSON-encoded object elements', () => {
    assert.deepEqual(coerceBodyEdits([{ find: 'a', replace: 'b' }, '{"find":"c","replace":"d"}']), [
      { find: 'a', replace: 'b' },
      { find: 'c', replace: 'd' },
    ]);
  });

  it('refuses an empty array, in either form', () => {
    const msg = 'bodyEdits cannot be empty; omit it to leave the body unchanged.';
    refuses(() => coerceBodyEdits([]), McpError, msg);
    refuses(() => coerceBodyEdits('[]'), McpError, msg);
  });

  it('refuses a value that is not an array', () => {
    const msg = 'bodyEdits must be an array of {find, replace} objects.';
    refuses(() => coerceBodyEdits('not json'), McpError, msg);
    refuses(() => coerceBodyEdits('{"find":"a","replace":"b"}'), McpError, msg);
    refuses(() => coerceBodyEdits({ find: 'a', replace: 'b' }), McpError, msg);
    refuses(() => coerceBodyEdits(3), McpError, msg);
  });

  it('refuses a bare-string element by index', () => {
    refuses(
      () => coerceBodyEdits([{ find: 'a', replace: 'b' }, 'find a']),
      McpError,
      'bodyEdits[1] must be a {find, replace} object, not a bare string.',
    );
    refuses(() => coerceBodyEdits(['{find']), McpError, 'bodyEdits[0] must be a {find, replace} object, not a bare string.');
    refuses(() => coerceBodyEdits(['find}']), McpError, 'bodyEdits[0] must be a {find, replace} object, not a bare string.');
  });

  it('refuses a braced string element that is not JSON, by index', () => {
    refuses(
      () => coerceBodyEdits([{ find: 'a', replace: 'b' }, '{find: a}']),
      McpError,
      "bodyEdits[1] is a string that isn't valid JSON; pass a {find, replace} object.",
    );
  });

  it('refuses a non-object element by index', () => {
    for (const bad of [null, 7, ['a', 'b'], true]) {
      refuses(() => coerceBodyEdits([{ find: 'a', replace: 'b' }, bad]), McpError, 'bodyEdits[1] must be a {find, replace} object.');
    }
  });

  it('refuses unknown keys by index, naming them', () => {
    refuses(
      () => coerceBodyEdits([{ find: 'a', replace: 'b', all: true, with: 'x' }]),
      McpError,
      'bodyEdits[0] has unknown key(s): all, with. Valid: find, replace',
    );
  });

  it('refuses a missing or non-string find or replace by index', () => {
    refuses(() => coerceBodyEdits([{ replace: 'b' }]), McpError, 'bodyEdits[0].find must be a string.');
    refuses(() => coerceBodyEdits([{ find: 1, replace: 'b' }]), McpError, 'bodyEdits[0].find must be a string.');
    refuses(() => coerceBodyEdits([{ find: 'a', replace: 'b' }, { find: 'a' }]), McpError, 'bodyEdits[1].replace must be a string.');
    refuses(() => coerceBodyEdits([{ find: 'a', replace: null }]), McpError, 'bodyEdits[0].replace must be a string.');
  });
});

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

describe('noteBodyEditsSplitSignature', () => {
  it('says one token in the singular', () => {
    assert.equal(
      noteBodyEditsSplitSignature('htmlBody', 1),
      '1 {{signature}} token in htmlBody was formed where a bodyEdits replace meets the stored text beside it, ' +
        'so it is stored as literal text: only a token wholly inside a replace expands. Put the whole {{signature}} in one replace.',
    );
  });

  it('says several tokens in the plural', () => {
    assert.equal(
      noteBodyEditsSplitSignature('textBody', 2),
      '2 {{signature}} tokens in textBody were formed where a bodyEdits replace meets the stored text beside it, ' +
        'so they are stored as literal text: only a token wholly inside a replace expands. Put the whole {{signature}} in one replace.',
    );
  });
});

describe('formatBodyEditsReceipt', () => {
  it('is empty without a receipt', () => {
    assert.equal(formatBodyEditsReceipt(undefined), '');
  });

  it('renders each op on one line, in order', () => {
    assert.equal(
      formatBodyEditsReceipt({
        part: 'htmlBody',
        ops: [
          { offset: 12, matchedSize: 5, replacementSize: 9 },
          { offset: 3, matchedSize: 2, replacementSize: 0 },
        ],
      }),
      '\nbodyEdits applied to htmlBody: [0] at offset 12, 5 chars replaced with 9; [1] at offset 3, 2 chars replaced with 0.',
    );
  });
});
