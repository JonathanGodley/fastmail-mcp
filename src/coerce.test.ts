import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceRecipients, coerceBool, coercePosition, clampLimit, coerceUtcDate, redactBearerTokens, registerSecret, requireNonEmpty, validateClearFields, parseAddress, assertKnownParams, coerceAttachments, InvalidInputError } from './coerce.js';
import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

describe('coerceStringArray', () => {
  it('returns undefined for undefined input', () => {
    assert.equal(coerceStringArray(undefined), undefined);
  });

  it('returns undefined for null input', () => {
    assert.equal(coerceStringArray(null), undefined);
  });

  it('returns undefined for non-array, non-string input', () => {
    assert.equal(coerceStringArray(123), undefined);
    assert.equal(coerceStringArray({}), undefined);
    assert.equal(coerceStringArray(true), undefined);
  });

  it('returns array as-is', () => {
    assert.deepEqual(coerceStringArray(['a@b.example', 'c@d.example']), ['a@b.example', 'c@d.example']);
  });

  it('stringifies array elements', () => {
    assert.deepEqual(coerceStringArray([1, 2, 3] as any), ['1', '2', '3']);
  });

  it('parses JSON-stringified array', () => {
    assert.deepEqual(coerceStringArray('["a@b.example", "c@d.example"]'), ['a@b.example', 'c@d.example']);
  });

  it('parses JSON-stringified array with whitespace', () => {
    assert.deepEqual(coerceStringArray('  ["a@b.example"]  '), ['a@b.example']);
  });

  it('splits comma-separated string', () => {
    assert.deepEqual(coerceStringArray('a@b.example, c@d.example'), ['a@b.example', 'c@d.example']);
  });

  it('wraps single address as one-item array', () => {
    assert.deepEqual(coerceStringArray('single@example.com'), ['single@example.com']);
  });

  it('returns empty array for empty string', () => {
    assert.deepEqual(coerceStringArray(''), []);
  });

  it('trims whitespace and filters empty segments in comma-split', () => {
    assert.deepEqual(coerceStringArray('a@b.example, ,c@d.example,'), ['a@b.example', 'c@d.example']);
  });

  it('falls back to comma-split when JSON parsing fails', () => {
    assert.deepEqual(coerceStringArray('[not valid json]'), ['[not valid json]']);
  });
});

describe('coerceRecipients', () => {
  it('coerces all four fields from arrays, JSON-strings, comma-strings, and bare strings', () => {
    const result = coerceRecipients({
      to: ['a@b.example'],
      cc: '["c@d.example", "e@f.example"]',
      bcc: 'g@h.example, i@j.example',
      replyTo: 'k@l.example',
    });
    assert.deepEqual(result, {
      to: ['a@b.example'],
      cc: ['c@d.example', 'e@f.example'],
      bcc: ['g@h.example', 'i@j.example'],
      replyTo: ['k@l.example'],
    });
  });

  it('coerces empty string to empty array for each field (the accepted edit-clear path)', () => {
    assert.deepEqual(coerceRecipients({ to: '', cc: '', bcc: '', replyTo: '' }), {
      to: [],
      cc: [],
      bcc: [],
      replyTo: [],
    });
  });

  it('returns undefined for omitted fields', () => {
    assert.deepEqual(coerceRecipients({}), {
      to: undefined,
      cc: undefined,
      bcc: undefined,
      replyTo: undefined,
    });
  });

  it('returns undefined for non-string, non-array values', () => {
    assert.deepEqual(coerceRecipients({ to: 123, cc: {}, bcc: true, replyTo: null } as any), {
      to: undefined,
      cc: undefined,
      bcc: undefined,
      replyTo: undefined,
    });
  });
});

describe('coerceBool', () => {
  it('returns boolean as-is', () => {
    assert.equal(coerceBool(true), true);
    assert.equal(coerceBool(false), false);
  });

  it('coerces "true" string to true', () => {
    assert.equal(coerceBool('true'), true);
  });

  it('coerces "false" string to false', () => {
    assert.equal(coerceBool('false'), false);
  });

  it('returns undefined for unrecognized strings', () => {
    assert.equal(coerceBool('yes'), undefined);
    assert.equal(coerceBool('1'), undefined);
    assert.equal(coerceBool(''), undefined);
  });

  it('returns undefined for null/undefined', () => {
    assert.equal(coerceBool(undefined), undefined);
    assert.equal(coerceBool(null), undefined);
  });

  it('returns undefined for numbers', () => {
    assert.equal(coerceBool(1), undefined);
    assert.equal(coerceBool(0), undefined);
  });

  // Pins the edit_draft noQuote handler seam: index.ts computes
  // `coerceBool(args.noQuote) === true` before calling updateDraft (which takes a real
  // boolean). The schema admits ['boolean','string'], so a lenient client's string
  // "true" must coerce to a real drop, and NOTHING ELSE may — a garbage string
  // yielding undefined can never silently discard a quote/forwarded block.
  it('noQuote seam: string "true" becomes a real drop; anything unrecognized never does', () => {
    assert.equal(coerceBool('true') === true, true);
    assert.equal(coerceBool(true) === true, true);
    assert.equal(coerceBool('false') === true, false);
    assert.equal(coerceBool('garbage') === true, false);
    assert.equal(coerceBool(undefined) === true, false);
  });
});

describe('coercePosition (#51)', () => {
  it('returns undefined for the blank shapes, which all mean "first page"', () => {
    assert.equal(coercePosition(undefined), undefined);
    assert.equal(coercePosition(null), undefined);
    assert.equal(coercePosition(''), undefined);
    assert.equal(coercePosition('   '), undefined);
  });

  it('accepts a number', () => {
    assert.equal(coercePosition(0), 0);
    assert.equal(coercePosition(40), 40);
  });

  // Lenient values: a client that stringifies numbers must not be forced to page by hand.
  it('accepts a stringified number', () => {
    assert.equal(coercePosition('40'), 40);
    assert.equal(coercePosition(' 40 '), 40);
    assert.equal(coercePosition('0'), 0);
  });

  // JMAP reads a negative position as an offset from the END of the results, so
  // accepting -1 would quietly serve the last page instead of failing.
  it('rejects a negative position and names the alternative', () => {
    assert.throws(
      () => coercePosition(-1),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /cannot be negative/);
        assert.match(err.message, /ascending:true/);
        return true;
      },
    );
    assert.throws(() => coercePosition('-20'), InvalidInputError);
  });

  it('rejects a fraction rather than rounding to a guessed offset', () => {
    assert.throws(
      () => coercePosition(1.5),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /whole number/);
        return true;
      },
    );
    assert.throws(() => coercePosition('1.5'), InvalidInputError);
  });

  it('rejects non-numeric text, Infinity and an integer too large to be exact', () => {
    assert.throws(() => coercePosition('abc'), InvalidInputError);
    assert.throws(() => coercePosition(NaN), InvalidInputError);
    assert.throws(() => coercePosition(Infinity), InvalidInputError);
    assert.throws(() => coercePosition(Number.MAX_SAFE_INTEGER + 2), InvalidInputError);
  });

  it('rejects non-number, non-string shapes', () => {
    assert.throws(() => coercePosition([20]), InvalidInputError);
    assert.throws(() => coercePosition({ position: 20 }), InvalidInputError);
    assert.throws(() => coercePosition(true), InvalidInputError);
  });

  it('names the parameter in the rejection so the caller knows what to fix', () => {
    assert.throws(
      () => coercePosition('abc'),
      (err: Error) => {
        assert.match(err.message, /^position /);
        return true;
      },
    );
  });
});

describe('clampLimit', () => {
  it('returns a value already inside the bounds unchanged', () => {
    assert.equal(clampLimit(5, 3, 10), 5);
  });

  it('accepts a stringified number from a lenient client', () => {
    assert.equal(clampLimit('7', 3, 10), 7);
  });

  it('falls back on non-numeric input rather than passing NaN through', () => {
    // A NaN limit reaches the JMAP query as `"limit": null`, which the server
    // reads as no limit at all - an unbounded result set.
    assert.equal(clampLimit('abc', 3, 10), 3);
    assert.equal(clampLimit(undefined, 3, 10), 3);
    assert.equal(clampLimit(null, 3, 10), 3);
    assert.equal(clampLimit({}, 3, 10), 3);
  });

  it('falls back on zero and lifts negatives to 1', () => {
    assert.equal(clampLimit(0, 3, 10), 3);
    assert.equal(clampLimit(-4, 3, 10), 1);
  });

  it('clamps above the maximum, including a stringified oversize value', () => {
    assert.equal(clampLimit(1000, 3, 10), 10);
    assert.equal(clampLimit('1000', 3, 10), 10);
  });
});

describe('coerceUtcDate (#70)', () => {
  it('returns undefined for undefined/null (no date bound)', () => {
    assert.equal(coerceUtcDate(undefined, 'after'), undefined);
    assert.equal(coerceUtcDate(null, 'before'), undefined);
  });

  it('expands a date-only value to midnight UTC on that date', () => {
    assert.equal(coerceUtcDate('2026-07-20', 'after'), '2026-07-20T00:00:00Z');
    assert.equal(coerceUtcDate('2026-01-01', 'before'), '2026-01-01T00:00:00Z');
  });

  it('passes a full UTC datetime through unchanged', () => {
    assert.equal(coerceUtcDate('2026-07-20T14:30:00Z', 'after'), '2026-07-20T14:30:00Z');
  });

  it('converts an offset datetime to UTC', () => {
    assert.equal(coerceUtcDate('2026-07-20T14:30:00+01:00', 'after'), '2026-07-20T13:30:00Z');
    assert.equal(coerceUtcDate('2026-07-20T14:30:00-05:00', 'before'), '2026-07-20T19:30:00Z');
  });

  it('reads a zone-less datetime as host local time and emits the same instant in UTC', () => {
    // Asserted against the host's own conversion so the test is timezone-independent;
    // what is pinned is that the instant survives and the emitted shape is a UTCDate.
    const out = coerceUtcDate('2026-07-20T14:30:00', 'after');
    assert.equal(new Date(out!).getTime(), new Date('2026-07-20T14:30:00').getTime());
    assert.match(out!, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
  });

  it('trims milliseconds to the canonical seconds-precision UTCDate', () => {
    assert.equal(coerceUtcDate('2026-07-20T14:30:00.123Z', 'after'), '2026-07-20T14:30:00Z');
  });

  it('trims surrounding whitespace before parsing', () => {
    assert.equal(coerceUtcDate('  2026-07-20  ', 'after'), '2026-07-20T00:00:00Z');
  });

  it('rejects an unparseable value, naming the parameter and the accepted formats', () => {
    assert.throws(
      () => coerceUtcDate('last tuesday-ish', 'before'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /^before /);
        assert.match(err.message, /last tuesday-ish/);
        assert.match(err.message, /2026-07-20T14:30:00Z/);
        return true;
      },
    );
  });

  it('rejects shapes outside YYYY-MM-DD / YYYY-MM-DDThh:mm:ss rather than guessing', () => {
    // Date's legacy fallback parser would accept all of these, but reads them as
    // host-local midnight instead of the documented UTC midnight (and rolls impossible
    // days over), so the search window would silently differ from what was asked for.
    for (const value of ['2026-7-20', '2026/07/20', '20 July 2026', 'July 20 2026', '2026', '2026-07', '2026-2-31', '2026/02/30', '2026-07-20T']) {
      assert.throws(
        () => coerceUtcDate(value, 'after'),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError);
          assert.match(err.message, /after is not a (valid|real calendar) date/);
          assert.match(err.message, /2026-07-20/);
          return true;
        },
        `expected ${value} to be rejected`,
      );
    }
  });

  it('rejects an impossible calendar date instead of rolling it over', () => {
    for (const value of ['2026-02-31', '2026-02-31T09:00:00Z', '2026-04-31T09:00:00+02:00']) {
      assert.throws(
        () => coerceUtcDate(value, 'after'),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError);
          assert.match(err.message, /after is not a real calendar date/);
          return true;
        },
      );
    }
  });

  it('keeps an offset value whose UTC date differs from the written date', () => {
    // The calendar-date guard must not misfire when an offset legitimately moves the
    // instant onto the next UTC day.
    assert.equal(coerceUtcDate('2026-07-20T23:00:00-05:00', 'before'), '2026-07-21T04:00:00Z');
  });

  it('rejects an empty or whitespace-only value rather than dropping the bound', () => {
    for (const value of ['', '   ']) {
      assert.throws(
        () => coerceUtcDate(value, 'after'),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError);
          assert.match(err.message, /after cannot be empty/);
          return true;
        },
      );
    }
  });

  it('rejects a non-string value', () => {
    assert.throws(
      () => coerceUtcDate(1753000000000, 'before'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /before must be a date string, not a number/);
        return true;
      },
    );
    assert.throws(
      () => coerceUtcDate(['2026-07-20'], 'after'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /after must be a date string, not an array/);
        return true;
      },
    );
  });

  it('truncates a very long value in the rejection message', () => {
    const long = 'x'.repeat(200);
    assert.throws(
      () => coerceUtcDate(long, 'after'),
      (err: Error) => {
        assert.match(err.message, /x{60}\.\.\./);
        assert.ok(!err.message.includes('x'.repeat(61)));
        return true;
      },
    );
  });
});

describe('redactBearerTokens', () => {
  it('redacts Bearer header pattern', () => {
    const out = redactBearerTokens('Authorization: Bearer abc.def.ghi failed');
    assert.equal(out, 'Authorization: Bearer [REDACTED] failed');
  });

  it('redacts case-insensitive Bearer', () => {
    assert.equal(redactBearerTokens('bearer secret'), 'Bearer [REDACTED]');
    assert.equal(redactBearerTokens('BEARER xyz'), 'Bearer [REDACTED]');
  });

  it('redacts Fastmail token shape (fmu...)', () => {
    // Synthetic value — matches the fmuN-<hex>-<hex>-N-<hex> shape only, never a real token.
    const out = redactBearerTokens(
      'Failed: token fmu0-00000000-1111111111111111111111111111111a-0-2222222222222222222222222222222b invalid' // allowlist-secret (synthetic)
    );
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('fmu0-0000'));
  });

  it('does not redact unrelated text', () => {
    const original = 'JMAP error: invalidArguments — mailbox not found';
    assert.equal(redactBearerTokens(original), original);
  });

  it('redacts multiple tokens in one string', () => {
    const out = redactBearerTokens('Bearer one and Bearer two');
    assert.equal(out, 'Bearer [REDACTED] and Bearer [REDACTED]');
  });

  it('handles empty string', () => {
    assert.equal(redactBearerTokens(''), '');
  });

  it('redacts Basic auth credentials (CalDAV path)', () => {
    const out = redactBearerTokens('401 on Authorization: Basic dXNlcjpwYXNzd29yZA== failed'); // allowlist-secret (synthetic base64 of "user:password")
    assert.match(out, /Basic \[REDACTED\]/);
    assert.ok(!out.includes('dXNlcjpwYXNz'));
  });

  it('redacts a Fastmail token containing an underscore fully (no tail leak)', () => {
    const out = redactBearerTokens('token fmu9-abcd1234-aaaaaaaaaaaaaaaaaaaa_bbbbbbbbbb invalid'); // allowlist-secret (synthetic)
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('bbbbbbbbbb'), 'the post-underscore tail must not survive');
  });

  it('redacts an exact registered secret value even without a recognizable shape', () => {
    registerSecret('sk-pla1n-cr3dential-with-no-prefix');
    const out = redactBearerTokens('self-hosted auth failed for sk-pla1n-cr3dential-with-no-prefix here');
    assert.ok(!out.includes('sk-pla1n-cr3dential'));
    assert.match(out, /\[REDACTED\]/);
  });

  it('registerSecret ignores short/empty values (avoids over-broad matches)', () => {
    registerSecret('abc');
    registerSecret('');
    assert.equal(redactBearerTokens('the word abc should survive'), 'the word abc should survive');
  });
});

describe('requireNonEmpty', () => {
  it('returns the trimmed value for a normal string', () => {
    assert.equal(requireNonEmpty('  hello  ', 'title'), 'hello');
  });

  it('returns the value unchanged when no trimming is needed', () => {
    assert.equal(requireNonEmpty('hello', 'title'), 'hello');
  });

  it('throws for an empty string', () => {
    assert.throws(() => requireNonEmpty('', 'title'), /title cannot be empty/);
  });

  it('throws for a whitespace-only string', () => {
    assert.throws(() => requireNonEmpty('   ', 'title'), /title cannot be empty/);
  });

  it('throws for null', () => {
    assert.throws(() => requireNonEmpty(null, 'title'), /title cannot be empty/);
  });

  it('throws for undefined', () => {
    assert.throws(() => requireNonEmpty(undefined, 'title'), /title cannot be empty/);
  });

  it('names the field in the error message', () => {
    assert.throws(() => requireNonEmpty('', 'location'), /location cannot be empty; omit the field to leave it unchanged/);
  });

  it('uses a custom hint when provided', () => {
    assert.throws(
      () => requireNonEmpty('', 'subject', 'list it in clearFields to clear it'),
      /subject cannot be empty; list it in clearFields to clear it/,
    );
  });

  it('throws InvalidInputError so the index maps it to InvalidParams (#41)', () => {
    assert.throws(() => requireNonEmpty('', 'title'), InvalidInputError);
  });
});

describe('validateClearFields', () => {
  const allowed = new Set(['description', 'location']);

  it('no-ops on an empty array', () => {
    assert.doesNotThrow(() => validateClearFields([], allowed, new Set()));
  });

  it('no-ops on undefined', () => {
    assert.doesNotThrow(() => validateClearFields(undefined as any, allowed, new Set()));
  });

  it('accepts allowed fields not also being set', () => {
    assert.doesNotThrow(() => validateClearFields(['location'], allowed, new Set(['title'])));
  });

  it('throws for a field not in the allowed set', () => {
    assert.throws(() => validateClearFields(['title'], allowed, new Set()), /title/);
  });

  it('lists the allowed set in the unknown-field error', () => {
    assert.throws(() => validateClearFields(['start'], allowed, new Set()), /description, location/);
  });

  it('throws when a field is both set and cleared', () => {
    assert.throws(
      () => validateClearFields(['description'], allowed, new Set(['description'])),
      /cannot both set and clear description/
    );
  });

  it('throws InvalidInputError so the index maps it to InvalidParams (#41)', () => {
    assert.throws(() => validateClearFields(['title'], allowed, new Set()), InvalidInputError);
    assert.throws(
      () => validateClearFields(['description'], allowed, new Set(['description'])),
      InvalidInputError,
    );
  });
});

describe('parseAddress', () => {
  it('parses "Name <email>" into name + email', () => {
    assert.deepEqual(parseAddress('Alice <a@x.example>'), { name: 'Alice', email: 'a@x.example' });
  });

  it('strips surrounding double-quotes from a quoted display name', () => {
    assert.deepEqual(parseAddress('"Doe, John" <j@x.example>'), { name: 'Doe, John', email: 'j@x.example' });
  });

  it('passes a bare address through as { email }', () => {
    assert.deepEqual(parseAddress('a@x.example'), { email: 'a@x.example' });
  });

  it('omits the name when angle brackets carry no display name', () => {
    assert.deepEqual(parseAddress('<a@x.example>'), { email: 'a@x.example' });
  });

  it('trims surrounding whitespace and an empty name', () => {
    assert.deepEqual(parseAddress('   <a@x.example>  '), { email: 'a@x.example' });
  });

  it('trims whitespace around name and email', () => {
    assert.deepEqual(parseAddress('  Bob   <  b@x.example  >  '), { name: 'Bob', email: 'b@x.example' });
  });

  it('uses the last angle-bracket pair so a name may contain "<"', () => {
    assert.deepEqual(parseAddress('a<b <c@x.example>'), { name: 'a<b', email: 'c@x.example' });
  });
});

describe('assertKnownParams (#11)', () => {
  const allowed = new Set(['mailboxId', 'limit', 'raw']);

  it('passes when every key is declared', () => {
    assert.doesNotThrow(() => assertKnownParams('list_emails', { mailboxId: 'x', limit: 5 }, allowed, false));
  });

  it('throws InvalidParams for an unknown key, listing the offender and the valid keys', () => {
    try {
      assertKnownParams('list_emails', { mailbox: 'drafts' }, allowed, false);
      assert.fail('expected throw');
    } catch (e) {
      assert.ok(e instanceof McpError);
      assert.equal((e as McpError).code, ErrorCode.InvalidParams);
      assert.match((e as McpError).message, /Unknown parameter\(s\): mailbox/);
      assert.match((e as McpError).message, /Valid: mailboxId, limit, raw/);
    }
  });

  it('lists every unknown key when several are present', () => {
    assert.throws(
      () => assertKnownParams('list_emails', { mailbox: 'x', folder: 'y', limit: 5 }, allowed, false),
      /Unknown parameter\(s\): mailbox, folder/,
    );
  });

  it('bypasses entirely when additionalProperties is true (escape hatch)', () => {
    assert.doesNotThrow(() => assertKnownParams('whatever', { anything: 1, goes: 2 }, allowed, true));
  });

  it('treats null/undefined args as no-args (passes)', () => {
    assert.doesNotThrow(() => assertKnownParams('list_emails', undefined, allowed, false));
    assert.doesNotThrow(() => assertKnownParams('list_emails', null, allowed, false));
  });

  it('a param-less tool (empty allowed set) rejects any arg but accepts {}', () => {
    assert.doesNotThrow(() => assertKnownParams('ping', {}, new Set(), false));
    assert.throws(() => assertKnownParams('ping', { x: 1 }, new Set(), false), /Unknown parameter\(s\): x/);
  });

  it('does NOT reject a stringified-but-known key — key-strictness only, value-leniency is separate', () => {
    assert.doesNotThrow(() => assertKnownParams('list_emails', { limit: '20' }, allowed, false));
  });
});

describe('coerceAttachments', () => {
  it('returns undefined for undefined/null', () => {
    assert.equal(coerceAttachments(undefined), undefined);
    assert.equal(coerceAttachments(null), undefined);
  });

  it('passes through an array of well-formed specs', () => {
    const specs = [{ path: 'a.pdf' }, { path: 'b.png', name: 'pic.png', contentType: 'image/png' }];
    assert.deepEqual(coerceAttachments(specs), specs);
  });

  it('parses a JSON-stringified array (lenient client)', () => {
    assert.deepEqual(
      coerceAttachments('[{"path":"a.pdf"},{"path":"b.png","name":"pic"}]'),
      [{ path: 'a.pdf' }, { path: 'b.png', name: 'pic' }],
    );
  });

  it('parses a per-item JSON-object string', () => {
    assert.deepEqual(coerceAttachments(['{"path":"a.pdf"}']), [{ path: 'a.pdf' }]);
  });

  it('rejects (does not drop) a spec missing path, naming the index', () => {
    assert.throws(
      () => coerceAttachments([{ name: 'x' }]),
      (err: unknown) => err instanceof McpError && err.code === ErrorCode.InvalidParams && /attachments\[0\].*path/.test(err.message),
    );
  });

  it('rejects a non-object element', () => {
    assert.throws(
      () => coerceAttachments([42]),
      (err: unknown) => err instanceof McpError && /attachments\[0\]/.test(err.message),
    );
  });

  it('rejects a bare (non-JSON) string element rather than guessing it is a path', () => {
    assert.throws(
      () => coerceAttachments(['a.pdf']),
      (err: unknown) => err instanceof McpError && /not a bare string/.test(err.message),
    );
  });

  it('rejects an unexpected per-item key (index + key named)', () => {
    assert.throws(
      () => coerceAttachments([{ path: 'a.pdf', bogus: 1 }]),
      (err: unknown) => err instanceof McpError && /attachments\[0\].*bogus/.test(err.message),
    );
  });

  it('rejects a non-array value', () => {
    assert.throws(
      () => coerceAttachments(42),
      (err: unknown) => err instanceof McpError && /must be an array/.test(err.message),
    );
  });

  it('trims an accidental leading/trailing space on path', () => {
    assert.deepEqual(coerceAttachments([{ path: '  report.pdf  ' }]), [{ path: 'report.pdf' }]);
  });

  // A cid is what makes a file display inside the body, and its value ends up in a MIME
  // header, so it is vetted here rather than anywhere downstream.
  it('accepts a plain cid token', () => {
    assert.deepEqual(coerceAttachments([{ path: 'a.png', cid: 'logo' }]), [{ path: 'a.png', cid: 'logo' }]);
  });

  it('normalises the two spellings a caller realistically copies a cid from', () => {
    // Out of an html reference, out of a header, and the header-quoted reference.
    assert.deepEqual(coerceAttachments([{ path: 'a.png', cid: 'cid:logo' }]), [{ path: 'a.png', cid: 'logo' }]);
    assert.deepEqual(coerceAttachments([{ path: 'b.png', cid: '<logo>' }]), [{ path: 'b.png', cid: 'logo' }]);
    assert.deepEqual(coerceAttachments([{ path: 'c.png', cid: '<cid:logo>' }]), [{ path: 'c.png', cid: 'logo' }]);
  });

  it('rejects a spelling the normalisation deliberately does not chase', () => {
    // `cid:<logo>` strips to `<logo>`, which fails the vet: chasing it would mean
    // accepting arbitrarily nested spellings.
    assert.throws(
      () => coerceAttachments([{ path: 'a.png', cid: 'cid:<logo>' }]),
      (err: unknown) => err instanceof McpError && err.code === ErrorCode.InvalidParams,
    );
  });

  it('rejects a cid carrying a line break (header injection), naming the real index', () => {
    assert.throws(
      () => coerceAttachments([{ path: 'a.png' }, { path: 'b.png', cid: 'logo\r\nX-Evil: 1' }]),
      (err: unknown) => err instanceof McpError && /attachments\[1\]\.cid/.test(err.message),
    );
  });

  it('rejects an over-long cid and a non-string cid', () => {
    assert.throws(() => coerceAttachments([{ path: 'a.png', cid: 'x'.repeat(65) }]), McpError);
    assert.throws(() => coerceAttachments([{ path: 'a.png', cid: 42 }]), McpError);
  });
});
