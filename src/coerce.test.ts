import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceStringArrayStrict, coerceRecipients, coerceBool, coercePosition, clampLimit, coerceUtcDate, redactBearerTokens, redactedJson, registerSecret, requireNonEmpty, validateClearFields, parseAddress, assertKnownParams, coerceAttachments, coerceParticipants, coerceContactEmails, coerceContactPhones, coerceContactAddresses, coerceContactName, InvalidInputError } from './coerce.js';
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

describe('coerceStringArrayStrict', () => {
  it('accepts everything the lenient coercion accepts', () => {
    assert.deepEqual(coerceStringArrayStrict(['Archive', 'Sent'], 'requiredMailboxes'), ['Archive', 'Sent']);
    assert.deepEqual(coerceStringArrayStrict('Archive, Sent', 'requiredMailboxes'), ['Archive', 'Sent']);
    assert.deepEqual(coerceStringArrayStrict('["Archive"]', 'requiredMailboxes'), ['Archive']);
    assert.deepEqual(coerceStringArrayStrict('', 'requiredMailboxes'), []);
  });

  it('treats undefined and null as absent', () => {
    assert.equal(coerceStringArrayStrict(undefined, 'requiredMailboxes'), undefined);
    assert.equal(coerceStringArrayStrict(null, 'requiredMailboxes'), undefined);
  });

  it('trims on the array branch, so both branches agree about whitespace', () => {
    // Without this the two branches of one coercer disagree: the comma-split branch has
    // always trimmed, while the array branch passed a padded element straight through. A
    // padded id then reaches the server and comes back as a not-found — a type error
    // wearing a lookup error's clothes, which is the exact failure this variant exists to
    // stop. Deleting the trim leaves every other assertion in this file green.
    // Through the LENIENT coercer directly. The strict one delegates here for the trim, so an
    // assertion that only went through coerceStringArrayStrict would keep passing if the
    // lenient array branch stopped trimming — which is every bulk tool's emailIds.
    assert.deepEqual(coerceStringArray([' e1 ', 'e2']), ['e1', 'e2']);
    assert.deepEqual(coerceStringArray('[" e1 "]'), ['e1']);
    assert.deepEqual(coerceStringArrayStrict([' M123 ', 'M2'], 'emailIds'), ['M123', 'M2']);
    assert.deepEqual(coerceStringArrayStrict('[" M123 "]', 'emailIds'), ['M123']);
    assert.deepEqual(coerceStringArrayStrict(' M123 , M2 ', 'emailIds'), ['M123', 'M2']);
  });

  // The whole point of the strict variant: a value that is present but unusable must not
  // come back as `undefined`, because on a scoping parameter that reads as "not supplied"
  // and silently widens the query the caller passed it to narrow.
  it('rejects a present-but-uncoercible value instead of dropping it', () => {
    for (const bad of [123, {}, true]) {
      assert.throws(
        () => coerceStringArrayStrict(bad, 'excludeMailboxes'),
        InvalidInputError,
      );
    }
  });

  it('names the parameter in the rejection', () => {
    assert.throws(
      () => coerceStringArrayStrict({ mailbox: 'Archive' }, 'excludeMailboxes'),
      /excludeMailboxes/,
    );
  });

  // Without the per-element check these reach the mailbox matcher as "null", "123" and
  // "[object Object]", and come back as a not-found error telling the caller to fix a
  // spelling — when what they actually passed was the wrong type.
  it('rejects a non-string entry by index instead of stringifying it', () => {
    const cases: [unknown, RegExp][] = [
      [[null], /\[0\] must be a string; received null/],
      [[123], /\[0\] must be a string; received number/],
      [[{}], /\[0\] must be a string; received object/],
      [[true], /\[0\] must be a string; received boolean/],
      [[['Archive']], /\[0\] must be a string; received array/],
      [['Archive', 7], /\[1\] must be a string; received number/],
    ];
    for (const [value, pattern] of cases) {
      assert.throws(() => coerceStringArrayStrict(value, 'requiredMailboxes'), InvalidInputError);
      assert.throws(() => coerceStringArrayStrict(value, 'requiredMailboxes'), pattern);
    }
  });

  // An empty element is a string, so the type check above passes it, and the plain coercer
  // only drops blanks on its comma-split branch. Without this it reaches the lookup as a real
  // value and comes back as "not found" — the same type-error-wearing-a-lookup-error's-clothes
  // failure the per-element check exists to stop.
  it('rejects an empty or whitespace-only element by index', () => {
    const cases: Array<[unknown, RegExp]> = [
      [[''], /\[0\] must be a non-empty string/],
      [['   '], /\[0\] must be a non-empty string/],
      [['Archive', ''], /\[1\] must be a non-empty string/],
      // A JSON-string array is unwrapped before the element check, so a blank cannot hide in
      // one. These two are the bare string, not an array wrapping it.
      ['[""]', /\[0\] must be a non-empty string/],
      ['["Archive", "  "]', /\[1\] must be a non-empty string/],
    ];
    for (const [value, pattern] of cases) {
      assert.throws(() => coerceStringArrayStrict(value, 'emailIds'), InvalidInputError);
      assert.throws(() => coerceStringArrayStrict(value, 'emailIds'), pattern);
    }
  });

  // A blank between commas is a separator artefact rather than something a caller wrote
  // down, so the comma-split branch keeps dropping it. The two are deliberately different.
  it('still drops blanks on the comma-split branch', () => {
    assert.deepEqual(coerceStringArrayStrict('Archive,,Sent', 'requiredMailboxes'), ['Archive', 'Sent']);
    assert.deepEqual(coerceStringArrayStrict('Archive, , Sent', 'requiredMailboxes'), ['Archive', 'Sent']);
  });

  // The JSON-string array a lenient client sends is unwrapped before the element check,
  // so it cannot smuggle a non-string entry past it.
  it('checks the elements of a JSON-string array too', () => {
    assert.deepEqual(coerceStringArrayStrict('["Archive", "Sent"]', 'requiredMailboxes'), ['Archive', 'Sent']);
    assert.throws(
      () => coerceStringArrayStrict('[1, 2]', 'requiredMailboxes'),
      /requiredMailboxes\[0\] must be a string; received number/,
    );
  });

  it('leaves the non-JSON string forms comma-split as before', () => {
    assert.deepEqual(coerceStringArrayStrict('[not valid json]', 'requiredMailboxes'), ['[not valid json]']);
    assert.deepEqual(coerceStringArrayStrict('Archive', 'requiredMailboxes'), ['Archive']);
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

});

describe('redactedJson', () => {
  it('stays parseable when a string value begins with "Bearer"', () => {
    // The failure this function exists to prevent. Redacting the FINISHED document instead
    // lets BEARER_PATTERN's `\S+` run past the value and eat its closing quote and the comma
    // after it, and the item a caller is promised can be parsed stops parsing. A mailbox
    // named "Bearer Bonds" is enough; mailbox names are free text.
    const out = redactedJson({ results: [{ id: 'e1', mailboxes: ['Bearer Bonds', 'Inbox'] }] }, 2);
    const parsed = JSON.parse(out);
    assert.deepEqual(parsed.results[0].mailboxes[1], 'Inbox');
  });

  it('still redacts a real credential inside a value', () => {
    // The other half: staying parseable must not come at the cost of letting a token through.
    const out = redactedJson({ results: [{ id: 'e1', reason: { description: 'auth failed: Bearer abc123xyz' } }] });
    JSON.parse(out);
    assert.match(out, /Bearer \[REDACTED\]/);
    assert.ok(!out.includes('abc123xyz'));
  });

  it('redacts a token that ENDS a value without breaking the document', () => {
    // A server description ending in a credential is the case where redaction matters most,
    // and it is also the case where document-level redaction destroyed the report - the
    // trailing delimiters sit immediately after the token with no whitespace between.
    const out = redactedJson({ description: 'rejected by Bearer sk-0123456789' });
    const parsed = JSON.parse(out);
    assert.equal(parsed.description, 'rejected by Bearer [REDACTED]');
  });

  it('leaves non-string values alone', () => {
    const out = redactedJson({ counts: { failed: 2 }, ok: true, missing: null });
    assert.deepEqual(JSON.parse(out), { counts: { failed: 2 }, ok: true, missing: null });
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

  it('rejects (does not drop) a spec naming no source, naming the index', () => {
    assert.throws(
      () => coerceAttachments([{ name: 'x' }]),
      (err: unknown) => err instanceof McpError && err.code === ErrorCode.InvalidParams
        && /attachments\[0\] names no source/.test(err.message)
        && /path/.test(err.message) && /blobId/.test(err.message) && /attachmentId/.test(err.message),
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

  // ---- the two in-account sources: blobId, and a part of an existing message ----

  it('accepts a blobId item, which must name the file recipients will see', () => {
    assert.deepEqual(
      coerceAttachments([{ blobId: 'G1234', name: 'report.pdf' }]),
      [{ blobId: 'G1234', name: 'report.pdf' }],
    );
  });

  it('rejects a bare blobId with no name rather than inventing a filename', () => {
    assert.throws(
      () => coerceAttachments([{ blobId: 'G1234' }]),
      (err: unknown) => err instanceof McpError && /attachments\[0\] gives a blobId but no 'name'/.test(err.message),
    );
    // A blank or whitespace-only name is the same omission wearing a value.
    assert.throws(() => coerceAttachments([{ blobId: 'G1234', name: '  ' }]), McpError);
  });

  // `??` cannot tell a blank name from an absent one, so a blank would ride all the way out
  // as the filename recipients see while the schema promised the source's own default.
  // Resolved here, once, so all three sources agree on what "no name given" means.
  it('reads a blank or whitespace-only name as absent, on every source that has a default', () => {
    assert.deepEqual(coerceAttachments([{ path: 'a.pdf', name: '' }]), [{ path: 'a.pdf' }]);
    assert.deepEqual(coerceAttachments([{ path: 'a.pdf', name: '   ' }]), [{ path: 'a.pdf' }]);
    assert.deepEqual(
      coerceAttachments([{ emailId: 'M1', attachmentId: 'p2', name: ' ' }]),
      [{ emailId: 'M1', attachmentId: 'p2' }],
    );
    // A real name is kept, trimmed — a stray space must not become the filename either.
    assert.deepEqual(coerceAttachments([{ path: 'a.pdf', name: ' r.pdf ' }]), [{ path: 'a.pdf', name: 'r.pdf' }]);
    // blobId is the one source with nothing to fall back to, so there the same input is a
    // rejection rather than a default. Both readings come from one rule, applied once.
    assert.throws(() => coerceAttachments([{ blobId: 'G1', name: '' }]), McpError);
  });

  it('accepts an emailId + attachmentId pair, and trims both ids', () => {
    assert.deepEqual(
      coerceAttachments([{ emailId: ' M1 ', attachmentId: ' 2.1 ' }]),
      [{ emailId: 'M1', attachmentId: '2.1' }],
    );
  });

  it('rejects half of the emailId/attachmentId pair, in either direction', () => {
    for (const half of [{ emailId: 'M1' }, { attachmentId: '2.1' }]) {
      assert.throws(
        () => coerceAttachments([half]),
        (err: unknown) => err instanceof McpError
          && /attachments\[0\] is missing a non-empty/.test(err.message)
          && /together/.test(err.message),
      );
    }
  });

  it('rejects an item naming two sources, naming the keys it saw', () => {
    assert.throws(
      () => coerceAttachments([{ path: 'a.pdf', blobId: 'G1' }]),
      (err: unknown) => err instanceof McpError && /names more than one source \(path, blobId\)/.test(err.message),
    );
    assert.throws(
      () => coerceAttachments([{ path: 'a.pdf', emailId: 'M1', attachmentId: '2' }]),
      (err: unknown) => err instanceof McpError && /names more than one source/.test(err.message),
    );
  });

  // A key belonging to a source the item did not choose is an error, never a silent
  // ignore: `{ blobId, attachmentId }` reads as two different intentions, and attaching
  // the blob while dropping the part reference would send bytes nobody asked for.
  it('rejects a key irrelevant to the chosen source instead of ignoring it', () => {
    assert.throws(
      () => coerceAttachments([{ blobId: 'G1', name: 'r.pdf', attachmentId: '2' }]),
      (err: unknown) => err instanceof McpError && /names more than one source \(blobId, attachmentId\)/.test(err.message),
    );
  });

  // A lenient client that fills every declared key sends null for the ones it has nothing
  // to say about. Reading those as "this item named four sources" would reject every call
  // such a client makes.
  it('treats a null source key as absent, not as a second source', () => {
    assert.deepEqual(
      coerceAttachments([{ path: 'a.pdf', blobId: null, emailId: null, attachmentId: null }]),
      [{ path: 'a.pdf' }],
    );
  });

  it('carries name/contentType/cid on every source', () => {
    assert.deepEqual(
      coerceAttachments([
        { blobId: 'G1', name: 'logo.png', contentType: 'image/png', cid: 'cid:logo' },
        { emailId: 'M1', attachmentId: 'p2', name: 'chart.png', cid: '<chart>' },
      ]),
      [
        { blobId: 'G1', name: 'logo.png', contentType: 'image/png', cid: 'logo' },
        { emailId: 'M1', attachmentId: 'p2', name: 'chart.png', cid: 'chart' },
      ],
    );
  });
});

describe('coerceParticipants', () => {
  const isInvalidInput = (pattern: RegExp) => (err: unknown) =>
    err instanceof InvalidInputError && pattern.test(err.message);

  it('returns undefined for undefined/null', () => {
    assert.equal(coerceParticipants(undefined), undefined);
    assert.equal(coerceParticipants(null), undefined);
  });

  it('passes through an array of well-formed entries', () => {
    const specs = [{ email: 'a@example.com' }, { email: 'b@example.com', name: 'Bee' }];
    assert.deepEqual(coerceParticipants(specs), specs);
  });

  it('parses a JSON-stringified array (lenient client)', () => {
    assert.deepEqual(
      coerceParticipants('[{"email":"a@example.com"},{"email":"b@example.com","name":"Bee"}]'),
      [{ email: 'a@example.com' }, { email: 'b@example.com', name: 'Bee' }],
    );
  });

  it('parses a JSON-stringified empty array as the empty list, which clears attendees', () => {
    assert.deepEqual(coerceParticipants('[]'), []);
  });

  it('parses a per-item JSON-object string', () => {
    assert.deepEqual(coerceParticipants(['{"email":"a@example.com","name":"Ay"}']), [
      { email: 'a@example.com', name: 'Ay' },
    ]);
  });

  it('accepts a bare address string as the email, matching the recipient lists', () => {
    assert.deepEqual(coerceParticipants(['a@example.com', { email: 'b@example.com' }]), [
      { email: 'a@example.com' },
      { email: 'b@example.com' },
    ]);
  });

  it('trims a stray space around an address, from either entry shape', () => {
    assert.deepEqual(coerceParticipants(['  a@example.com  ', { email: ' b@example.com ' }]), [
      { email: 'a@example.com' },
      { email: 'b@example.com' },
    ]);
  });

  it('treats a blank string parameter as omitted, never as the destructive empty list', () => {
    // An empty array REMOVES every attendee on update_calendar_event, so an ambiguous
    // blank must not resolve in that direction.
    assert.equal(coerceParticipants(''), undefined);
    assert.equal(coerceParticipants('   '), undefined);
  });

  it('rejects a non-array, non-string value', () => {
    assert.throws(() => coerceParticipants(42), isInvalidInput(/must be an array/));
    assert.throws(() => coerceParticipants({ email: 'a@example.com' }), isInvalidInput(/must be an array/));
  });

  it('rejects a string that is not JSON, rather than comma-splitting it', () => {
    assert.throws(
      () => coerceParticipants('a@example.com,b@example.com'),
      isInvalidInput(/must be an array/),
    );
  });

  it('rejects a JSON string that parses to something other than an array', () => {
    assert.throws(() => coerceParticipants('{"email":"a@example.com"}'), isInvalidInput(/must be an array/));
  });

  it('rejects an entry missing a usable email, naming the index', () => {
    assert.throws(
      () => coerceParticipants([{ email: 'a@example.com' }, { name: 'Bee' }]),
      isInvalidInput(/participants\[1\].*email/),
    );
    assert.throws(
      () => coerceParticipants([{ email: '   ' }]),
      isInvalidInput(/participants\[0\].*email/),
    );
  });

  it('rejects an unknown per-item key, naming the index and the key', () => {
    // The MCP SDK does not enforce inputSchema, so an rsvp/role the tool never declared
    // would otherwise ride through into the ATTENDEE line.
    assert.throws(
      () => coerceParticipants([{ email: 'a@example.com' }, { email: 'b@example.com', rsvp: true }]),
      isInvalidInput(/participants\[1\].*rsvp/),
    );
  });

  it('rejects a non-string email, which the schema alone does not stop', () => {
    assert.throws(
      () => coerceParticipants([{ email: 42 }]),
      isInvalidInput(/participants\[0\].*email/),
    );
    assert.throws(
      () => coerceParticipants([{ email: ['a@example.com'] }]),
      isInvalidInput(/participants\[0\].*email/),
    );
  });

  it('rejects a non-string name, which the schema alone does not stop', () => {
    assert.throws(
      () => coerceParticipants([{ email: 'a@example.com', name: { first: 'Ay' } }]),
      isInvalidInput(/participants\[0\]\.name must be a string/),
    );
  });

  it('rejects a non-object, non-string element', () => {
    assert.throws(() => coerceParticipants([42]), isInvalidInput(/participants\[0\]/));
    assert.throws(() => coerceParticipants([null]), isInvalidInput(/participants\[0\]/));
    assert.throws(() => coerceParticipants([['a@example.com']]), isInvalidInput(/participants\[0\]/));
  });

  it('rejects an empty-string element', () => {
    assert.throws(() => coerceParticipants(['  ']), isInvalidInput(/participants\[0\]/));
  });

  it('rejects a brace-wrapped element that is not valid JSON', () => {
    assert.throws(() => coerceParticipants(['{oops}']), isInvalidInput(/participants\[0\].*JSON/));
  });

  it('builds fresh objects rather than passing the caller\'s through', () => {
    // A future key added to ParticipantSpec must be a conscious edit here, not something
    // that starts flowing through because the caller's object was reused.
    const input = [{ email: 'a@example.com', name: 'Ay' }];
    const out = coerceParticipants(input)!;
    assert.notEqual(out[0], input[0]);
    assert.deepEqual(out[0], { email: 'a@example.com', name: 'Ay' });
  });
});

// ---------- contact write inputs ----------

describe('contact entry coercion', () => {
  const isInvalidInput = (pattern: RegExp) => (err: unknown) =>
    err instanceof InvalidInputError && pattern.test(err.message);

  it('returns undefined for undefined, null and a blank string', () => {
    // A blank string reads as "not supplied", never as the empty array: the empty array is
    // a REJECTED shape on these parameters, so resolving a client quirk into it would turn
    // a stringification bug into a rejection the caller cannot explain.
    for (const coerce of [coerceContactEmails, coerceContactPhones, coerceContactAddresses]) {
      assert.equal(coerce(undefined), undefined);
      assert.equal(coerce(null), undefined);
      assert.equal(coerce('   '), undefined);
    }
  });

  it('accepts both the bare-value and object shapes for emails and phones', () => {
    assert.deepEqual(coerceContactEmails(['a@b.example', { address: 'c@d.example', label: 'work' }]), [
      { address: 'a@b.example' },
      { address: 'c@d.example', label: 'work' },
    ]);
    assert.deepEqual(coerceContactPhones(['+1 555 0100', { number: '+1 555 0199', label: 'work' }]), [
      { number: '+1 555 0100' },
      { number: '+1 555 0199', label: 'work' },
    ]);
  });

  it('accepts a JSON-string array from a lenient client', () => {
    assert.deepEqual(coerceContactEmails('[{"address":"a@b.example","label":"work"}]'), [
      { address: 'a@b.example', label: 'work' },
    ]);
    assert.deepEqual(coerceContactAddresses('[{"full":"1 Road"}]'), [{ full: '1 Road' }]);
  });

  it('refuses a bare string for addresses, which has no single obvious reading', () => {
    assert.throws(() => coerceContactAddresses(['1 Road']), isInvalidInput(/addresses\[0\].*bare string/));
    assert.deepEqual(coerceContactAddresses([{ full: '1 Road', label: 'home' }]), [{ full: '1 Road', label: 'home' }]);
  });

  it('rejects an unknown per-item key, naming the index', () => {
    assert.throws(() => coerceContactEmails([{ address: 'a@b.example', type: 'work' }]), isInvalidInput(/emails\[0\].*unknown key\(s\): type/));
    assert.throws(() => coerceContactPhones([{ number: '1' }, { number: '2', pref: 1 }]), isInvalidInput(/phones\[1\].*unknown key\(s\): pref/));
    assert.throws(() => coerceContactAddresses([{ full: '1 Road', country: 'GB' }]), isInvalidInput(/addresses\[0\].*unknown key\(s\): country/));
  });

  it('rejects a WRONG-TYPED value, naming the index', () => {
    // A key allowlist alone is not enough. The MCP SDK does not enforce inputSchema, and
    // these values are written into a ContactCard/set patch, so an object where a string
    // belongs would land on a real card verbatim.
    assert.throws(() => coerceContactEmails([{ address: { at: 'b' } }]), isInvalidInput(/emails\[0\]\.address must be a string/));
    assert.throws(() => coerceContactEmails([{ address: 'a@b.example', label: [] }]), isInvalidInput(/emails\[0\]\.label must be a string/));
    assert.throws(() => coerceContactPhones([{ number: 5550100 }]), isInvalidInput(/phones\[0\]\.number must be a string/));
    assert.throws(() => coerceContactAddresses([{ full: ['1 Road'] }]), isInvalidInput(/addresses\[0\]\.full must be a string/));
  });

  it('rejects a missing or blank primary value, naming the index', () => {
    assert.throws(() => coerceContactEmails([{ label: 'work' }]), isInvalidInput(/emails\[0\]\.address must be a string/));
    assert.throws(() => coerceContactEmails([{ address: '   ' }]), isInvalidInput(/emails\[0\].*non-empty 'address'/));
    assert.throws(() => coerceContactEmails(['  ']), isInvalidInput(/emails\[0\].*empty string/));
  });

  it('rejects a blank label rather than writing a property that reads as absent', () => {
    // On a stored card `label: ""` is exactly what "no label" looks like, so writing one
    // would be a change with no visible effect. It is deliberately not repurposed as a way
    // to REMOVE a label either — that would be a clearing mechanism invented in the coercer.
    assert.throws(
      () => coerceContactEmails([{ address: 'a@b.example', label: '  ' }]),
      isInvalidInput(/emails\[0\]\.label cannot be empty; omit it/),
    );
    assert.throws(
      () => coerceContactPhones([{ number: '+1 555 0100' }, { number: '+1 555 0199', label: '' }]),
      isInvalidInput(/phones\[1\]\.label cannot be empty/),
    );
  });

  it('rejects a repeated value, naming both positions', () => {
    // A repeat cannot be matched against the stored card twice, so it would silently
    // surface as an unknown addition and trip the ambiguity guard for no reason.
    assert.throws(
      () => coerceContactEmails(['a@b.example', { address: 'a@b.example', label: 'work' }]),
      isInvalidInput(/emails\[1\] repeats the address already given at emails\[0\]/),
    );
  });

  it('rejects a non-array and a non-object element', () => {
    assert.throws(() => coerceContactEmails(42), isInvalidInput(/must be an array/));
    assert.throws(() => coerceContactEmails([42]), isInvalidInput(/emails\[0\]/));
    assert.throws(() => coerceContactEmails([null]), isInvalidInput(/emails\[0\]/));
  });

  it('builds fresh objects rather than passing the caller\'s through', () => {
    // Adding a key later has to be a conscious edit in the coercer, not something that
    // starts flowing to the server because the caller's object was reused.
    const input = [{ address: 'a@b.example', label: 'work' }];
    const out = coerceContactEmails(input)!;
    assert.notEqual(out[0], input[0]);
    assert.deepEqual(out[0], { address: 'a@b.example', label: 'work' });
  });
});

describe('coerceContactName', () => {
  const isInvalidInput = (pattern: RegExp) => (err: unknown) =>
    err instanceof InvalidInputError && pattern.test(err.message);

  it('reads a bare string as the full name', () => {
    assert.deepEqual(coerceContactName('  Ada Lovelace '), { full: 'Ada Lovelace' });
  });

  it('accepts the structured form and a JSON-string of it', () => {
    assert.deepEqual(coerceContactName({ given: 'Ada', surname: 'Lovelace' }), { given: 'Ada', surname: 'Lovelace' });
    assert.deepEqual(coerceContactName('{"full":"Ada Lovelace"}'), { full: 'Ada Lovelace' });
  });

  it('returns undefined for undefined/null', () => {
    assert.equal(coerceContactName(undefined), undefined);
    assert.equal(coerceContactName(null), undefined);
  });

  it('rejects an unknown key and a wrong-typed one', () => {
    assert.throws(() => coerceContactName({ nickname: 'Ada' }), isInvalidInput(/unknown key\(s\): nickname/));
    assert.throws(() => coerceContactName({ given: 42 }), isInvalidInput(/name\.given must be a string/));
    assert.throws(() => coerceContactName(['Ada']), isInvalidInput(/full-name string or an object/));
  });

  it('rejects a blank name rather than reading it as "clear the name"', () => {
    // `name` is not clearable, so a silently dropped blank would leave the caller believing
    // a name had been removed when nothing happened.
    assert.throws(() => coerceContactName('   '), isInvalidInput(/name cannot be empty/));
    assert.throws(() => coerceContactName({ given: '  ' }), isInvalidInput(/name\.given cannot be empty/));
    assert.throws(() => coerceContactName({}), isInvalidInput(/at least one of/));
  });
});
