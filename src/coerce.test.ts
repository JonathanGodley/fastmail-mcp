import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceStringArrayStrict, coerceRecipients, coerceBool, coercePosition, clampLimit, coerceUtcDate, coerceCalendarWindowStart, coerceCalendarWindowEnd, describeTimezone, resolveUsableTimezone, isUsableTimezone, validateCallerTimezone, resolveCalendarInstantMs, redactBearerTokens, redactedJson, registerSecret, requireNonEmpty, validateClearFields, parseAddress, assertKnownParams, coerceAttachments, coerceParticipants, coerceContactEmails, coerceContactPhones, coerceContactAddresses, coerceContactName, InvalidInputError } from './coerce.js';
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
        // One ellipsis CHARACTER, not three dots: every echo in this server now goes through
        // the same helper, so the truncation marker is the same everywhere.
        assert.match(err.message, /x{60}…/);
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
    const out = redactedJson({ results: [{ id: 'e1', mailboxes: ['Bearer Bonds', 'Inbox'] }] });
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

// Every case here passes an EXPLICIT zone. A test that leaves it to the host machine
// passes under both the UTC-day reading and the local-day one whenever the host happens to
// sit in the zone being asserted — which is how a whole suite stayed green while a date-only
// window answered the wrong day (#138). 'UTC' below is a pinned zone, not "no zone".
describe('coerceCalendarWindowEnd (#64)', () => {
  it('returns undefined for undefined/null (no window bound)', () => {
    assert.equal(coerceCalendarWindowEnd(undefined, 'endDate', 'UTC'), undefined);
    assert.equal(coerceCalendarWindowEnd(null, 'endDate', 'UTC'), undefined);
  });

  it('runs a date-only value to the end of that day', () => {
    // The one deliberate divergence from coerceUtcDate. CalDAV time-range ends are
    // exclusive, so the following midnight is exactly "through the end of the 10th".
    assert.equal(coerceCalendarWindowEnd('2027-03-10', 'endDate', 'UTC'), '2027-03-11T00:00:00Z');
  });

  it('rolls a date-only value at a month end into the next month', () => {
    assert.equal(coerceCalendarWindowEnd('2027-03-31', 'endDate', 'UTC'), '2027-04-01T00:00:00Z');
    assert.equal(coerceCalendarWindowEnd('2027-12-31', 'endDate', 'UTC'), '2028-01-01T00:00:00Z');
    // A leap day is a real date and must advance to 1 March, not 29 February again.
    assert.equal(coerceCalendarWindowEnd('2028-02-28', 'endDate', 'UTC'), '2028-02-29T00:00:00Z');
  });

  it('takes a full datetime literally, like coerceUtcDate', () => {
    assert.equal(coerceCalendarWindowEnd('2027-03-10T17:00:00Z', 'endDate', 'UTC'), '2027-03-10T17:00:00Z');
    assert.equal(coerceCalendarWindowEnd('2027-03-10T17:00:00+01:00', 'endDate', 'UTC'), '2027-03-10T16:00:00Z');
  });

  it('trims surrounding whitespace on a date-only value before extending it', () => {
    assert.equal(coerceCalendarWindowEnd('  2027-03-10  ', 'endDate', 'UTC'), '2027-03-11T00:00:00Z');
  });

  it('rejects a day its month does not have rather than rolling it over', () => {
    assert.throws(
      () => coerceCalendarWindowEnd('2027-02-30', 'endDate', 'UTC'),
      /endDate is not a real calendar date/,
    );
  });

  it('rejects free text, reduced precision and non-strings, naming the parameter', () => {
    for (const value of ['next friday', '2027', '2027-03', '2027/03/10', '', '   ']) {
      assert.throws(
        () => coerceCalendarWindowEnd(value, 'endDate', 'UTC'),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, /endDate/);
          return true;
        },
        `expected rejection for ${JSON.stringify(value)}`,
      );
    }
    assert.throws(() => coerceCalendarWindowEnd(20270310, 'endDate', 'UTC'), /endDate must be a date string/);
  });
});

// A calendar window's DAY is a local day (#138). Every zone here is injected, and two of
// them are chosen so a sign error cannot hide: Sydney is ahead of UTC, New York behind it,
// so a start that should move BACKWARD in one must move FORWARD in the other.
describe('calendar window bounds resolve a date in the configured zone (#138)', () => {
  const SYDNEY = 'Australia/Sydney';   // +10:00 in August, +11:00 in January
  const NEW_YORK = 'America/New_York'; // -04:00 in August, -05:00 in January

  it('reads a date-only startDate as local midnight, not UTC midnight', () => {
    // The reported defect, pinned. Asking a +10:00 account for the 12th used to search
    // 12 Aug 10:00 to 13 Aug 10:00 local, so an 08:00 appointment ON the 12th fell outside
    // the window and the day read as quieter than it was.
    assert.equal(coerceCalendarWindowStart('2026-08-12', 'startDate', SYDNEY), '2026-08-11T14:00:00Z');
    assert.equal(coerceCalendarWindowEnd('2026-08-12', 'endDate', SYDNEY), '2026-08-12T14:00:00Z');
  });

  it('moves the bound the other way for a zone behind UTC', () => {
    assert.equal(coerceCalendarWindowStart('2026-08-12', 'startDate', NEW_YORK), '2026-08-12T04:00:00Z');
    assert.equal(coerceCalendarWindowEnd('2026-08-12', 'endDate', NEW_YORK), '2026-08-13T04:00:00Z');
  });

  it('uses the offset in force on the date asked about, not one fixed offset', () => {
    // January is +11:00 in Sydney and -05:00 in New York; August is +10:00 and -04:00. A
    // hard-coded offset would be right for half the year in each.
    assert.equal(coerceCalendarWindowStart('2026-01-15', 'startDate', SYDNEY), '2026-01-14T13:00:00Z');
    assert.equal(coerceCalendarWindowStart('2026-01-15', 'startDate', NEW_YORK), '2026-01-15T05:00:00Z');
  });

  it('spans a whole local day even when a DST change makes it 23 or 25 hours long', () => {
    // Sydney springs forward on 2026-10-04 (a 23-hour day) and falls back on 2026-04-05 (a
    // 25-hour one). Advancing the exclusive end by a fixed 24 hours would land it an hour
    // inside the day on one and an hour into the next day on the other.
    const springStart = coerceCalendarWindowStart('2026-10-04', 'startDate', SYDNEY)!;
    const springEnd = coerceCalendarWindowEnd('2026-10-04', 'endDate', SYDNEY)!;
    assert.equal((Date.parse(springEnd) - Date.parse(springStart)) / 3600000, 23);

    const fallStart = coerceCalendarWindowStart('2026-04-05', 'startDate', SYDNEY)!;
    const fallEnd = coerceCalendarWindowEnd('2026-04-05', 'endDate', SYDNEY)!;
    assert.equal((Date.parse(fallEnd) - Date.parse(fallStart)) / 3600000, 25);
  });

  it('reads a zone-less datetime in the configured zone, deliberately', () => {
    // This used to resolve in whatever zone the SERVER PROCESS ran in, by accident of how
    // `new Date()` parses a value with no designator — so the same call answered differently
    // on two machines, with nothing in the schema admitting it.
    assert.equal(coerceCalendarWindowStart('2026-08-12T09:00:00', 'startDate', SYDNEY), '2026-08-11T23:00:00Z');
    assert.equal(coerceCalendarWindowStart('2026-08-12T09:00:00', 'startDate', NEW_YORK), '2026-08-12T13:00:00Z');
    // Seconds are optional in the wall-clock form.
    assert.equal(coerceCalendarWindowStart('2026-08-12T09:00', 'startDate', SYDNEY), '2026-08-11T23:00:00Z');
  });

  it('does NOT add a day to a zone-less datetime used as the end', () => {
    // The whole-day rule belongs to a DATE. A wall clock names a time of day, so the
    // exclusive end is that time — the same as its Z-designated equivalent.
    assert.equal(coerceCalendarWindowEnd('2026-08-12T17:00:00', 'endDate', SYDNEY), '2026-08-12T07:00:00Z');
  });

  it('leaves an explicit Z or numeric offset exactly as written, in any zone', () => {
    for (const zone of [SYDNEY, NEW_YORK, 'UTC']) {
      assert.equal(coerceCalendarWindowStart('2026-08-12T09:00:00Z', 'startDate', zone), '2026-08-12T09:00:00Z');
      assert.equal(coerceCalendarWindowEnd('2026-08-12T09:00:00Z', 'endDate', zone), '2026-08-12T09:00:00Z');
      assert.equal(coerceCalendarWindowStart('2026-08-12T09:00:00+05:30', 'startDate', zone), '2026-08-12T03:30:00Z');
      // A `+00:00` offset is a zone designator too, so it is taken literally rather than
      // being mistaken for a wall clock.
      assert.equal(coerceCalendarWindowStart('2026-08-12T09:00:00+00:00', 'startDate', zone), '2026-08-12T09:00:00Z');
    }
  });

  it('falls back to the host zone rather than throwing on an unusable zone name', () => {
    // A mistyped FASTMAIL_TIMEZONE is a deployment mistake, not a caller mistake; it must
    // not turn every calendar read into an error. Same posture as toLocalIso.
    assert.equal(
      coerceCalendarWindowStart('2026-08-12', 'startDate', 'Not/AZone'),
      coerceCalendarWindowStart('2026-08-12', 'startDate', undefined),
    );
  });

  it('rejects the same bad values the UTC coercion does, naming the zone in the hint', () => {
    for (const value of ['next friday', '2027', '2027-03', '2027/03/10', '', '   ']) {
      assert.throws(
        () => coerceCalendarWindowStart(value, 'startDate', SYDNEY),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, /startDate/);
          return true;
        },
        `expected rejection for ${JSON.stringify(value)}`,
      );
    }
    assert.throws(() => coerceCalendarWindowStart('2027-02-30', 'startDate', SYDNEY), /is not a real calendar date/);
    assert.throws(() => coerceCalendarWindowStart(20270310, 'startDate', SYDNEY), /must be a date string/);
    // The accepted-shapes sentence names the zone the caller's dates will be read in, so a
    // rejection is also the one place the rule is stated back to them.
    assert.throws(() => coerceCalendarWindowStart('next friday', 'startDate', SYDNEY), /Australia\/Sydney/);
  });

  it('leaves the email search bounds on the UTC rule', () => {
    // coerceUtcDate is shared with search_emails' before/after, which compare against a
    // message's receivedAt — an instant, not a day on anybody's wall. It is deliberately
    // NOT zone-aware, and this pins that the calendar change did not leak into it.
    assert.equal(coerceUtcDate('2026-08-12', 'after'), '2026-08-12T00:00:00Z');
  });
});

describe('describeTimezone', () => {
  it('names the configured zone', () => {
    assert.equal(describeTimezone('Australia/Sydney'), 'Australia/Sydney');
  });

  it('names the host zone when none is configured, rather than saying nothing', () => {
    const host = describeTimezone(undefined);
    assert.ok(host && host.length > 0);
    assert.equal(host, Intl.DateTimeFormat().resolvedOptions().timeZone);
  });
});

// resolveUsableTimezone (#139) is the single decision point describeTimezone and every calendar
// read path share for "which zone is really in force" — these pin that it is not just an
// internal helper of describeTimezone's wording, but returns the resolved NAME itself.
describe('resolveUsableTimezone', () => {
  it('returns the configured zone when it is set and resolvable', () => {
    assert.equal(resolveUsableTimezone('Australia/Sydney'), 'Australia/Sydney');
  });

  it('falls back to the host zone when nothing is configured', () => {
    assert.equal(resolveUsableTimezone(undefined), Intl.DateTimeFormat().resolvedOptions().timeZone);
  });

  it('falls back to the host zone when the configured value is not one ICU can resolve', () => {
    // A non-IANA name (e.g. a Windows zone from an external invite) is a valid TZID to echo
    // verbatim on an event, but it is not a usable CONFIGURED zone — there is nothing to
    // compare an event's own zone against, so the host zone is what is actually in force.
    assert.equal(resolveUsableTimezone('AUS Eastern Standard Time'), Intl.DateTimeFormat().resolvedOptions().timeZone);
  });

  it('agrees with isUsableTimezone on what counts as resolvable', () => {
    assert.equal(isUsableTimezone('Australia/Sydney'), true);
    assert.equal(isUsableTimezone('Not/AZone'), false);
    assert.equal(resolveUsableTimezone('Not/AZone') === 'Not/AZone', isUsableTimezone('Not/AZone'));
  });
});

// validateCallerTimezone is the timeZone argument's gate on create_calendar_event/
// update_calendar_event (#157). Its offset-shape denylist runs BEFORE isUsableTimezone (which
// alone would wrongly accept an offset-shaped string like "+10:00" — ICU resolves it as a
// fixed-offset zone), so these pin the denylist's exact boundary: what it rejects that
// isUsableTimezone alone would have let through, and what it must still let past to
// isUsableTimezone (Etc/GMT-10 and other legitimate names that happen to contain digits).
describe('validateCallerTimezone', () => {
  it('accepts a usable IANA zone name, trimmed', () => {
    assert.equal(validateCallerTimezone('Australia/Sydney'), 'Australia/Sydney');
    assert.equal(validateCallerTimezone('  Australia/Sydney  '), 'Australia/Sydney');
  });

  // ICU resolves a zone name case-insensitively and through backward-compatibility links, but
  // Cyrus (the CalDAV server behind this account) looks a TZID up with an exact-string match
  // against its own tzdata — far stricter than ICU. So the return value is ICU's own canonical
  // spelling for whatever resolved, not an echo of what the caller typed, and these pin actual
  // resolutions on this runtime (Node's ICU) rather than assuming case alone changes.
  for (const [input, canonical] of [
    ['australia/sydney', 'Australia/Sydney'],
    ['AUSTRALIA/SYDNEY', 'Australia/Sydney'],
    ['NZ', 'Pacific/Auckland'],
    ['Zulu', 'UTC'],
    ['GMT0', 'UTC'],
  ] as const) {
    it(`canonicalises "${input}" to "${canonical}", not the caller's own spelling`, () => {
      assert.equal(validateCallerTimezone(input), canonical);
    });
  }

  // RFC 5545 §3.2.19 permits a leading '/' on a TZID; the read half (#139) emits it verbatim
  // when a stored event carries it, so a caller echoing back a timeZone this server just
  // handed them (read-modify-write) must not be rejected for that spelling.
  it('accepts a leading slash (RFC 5545 §3.2.19) and canonicalises past it', () => {
    assert.equal(validateCallerTimezone('/Australia/Sydney'), 'Australia/Sydney');
  });

  // Only the simple '/Zone/Name' form is recoverable this way — a vendor-prefixed TZID has
  // more than one slash and still fails to resolve after only the leading one is stripped,
  // which is the safe direction (a real rejection, not a guess at which registry it names).
  it('still rejects a vendor-prefixed TZID beyond the simple leading-slash form', () => {
    assert.throws(
      () => validateCallerTimezone('/vendor.example/20050126_1/Australia/Sydney'),
      /is not a time zone this server can resolve/
    );
  });

  // A non-IANA name is a genuine round-trip limit, not a gap in the ICU probe above: there is
  // no canonical IANA spelling to resolve it to.
  it('rejects a non-IANA Windows zone name', () => {
    assert.throws(
      () => validateCallerTimezone('AUS Eastern Standard Time'),
      /is not a time zone this server can resolve/
    );
  });

  it('rejects null as a request to write floating, not a floating value itself', () => {
    assert.throws(() => validateCallerTimezone(null), /cannot be null, empty, or whitespace-only/);
  });

  it('rejects an empty or whitespace-only string the same way as null', () => {
    assert.throws(() => validateCallerTimezone(''), /cannot be null, empty, or whitespace-only/);
    assert.throws(() => validateCallerTimezone('   '), /cannot be null, empty, or whitespace-only/);
  });

  it('rejects a non-string value naming its actual type', () => {
    assert.throws(() => validateCallerTimezone(42), /must be an IANA zone name string.*not number/);
  });

  // Offset-shaped strings ICU resolves as a fixed-offset zone, which is exactly what a
  // calendar time must never carry — isUsableTimezone alone would (wrongly) accept every one
  // of these, so the denylist has to run first.
  for (const offsetShaped of ['+10:00', '+1000', '+10', '-05:00', 'GMT+10', 'GMT-5', 'UTC+10', 'UT+10']) {
    it(`rejects the offset-shaped string "${offsetShaped}"`, () => {
      assert.throws(
        () => validateCallerTimezone(offsetShaped),
        /must be an IANA zone NAME.*not a fixed UTC offset/
      );
    });
  }

  it('rejects a digit-leading string even outside the GMT/UTC/UT forms', () => {
    assert.throws(() => validateCallerTimezone('5Etc/Something'), /not a fixed UTC offset/);
  });

  // The denylist must not overreach: these are legitimate IANA names that happen to fail the
  // GMT/UTC/UT/leading-digit shapes only superficially (Etc/GMT-10 embeds a POSIX-style sign
  // AFTER the name, not at the start) and must still resolve via isUsableTimezone. Most of
  // these are also unchanged by canonicalisation (an Etc/GMT offset name and UTC are already
  // ICU-canonical); EST5EDT is not, and is pinned to what it actually resolves to.
  for (const [legit, canonical] of [
    ['Etc/GMT-10', 'Etc/GMT-10'],
    ['Etc/GMT+5', 'Etc/GMT+5'],
    ['UTC', 'UTC'],
    ['EST5EDT', 'America/New_York'],
  ] as const) {
    it(`still accepts the legitimate zone name "${legit}"${legit === canonical ? '' : `, canonicalised to "${canonical}"`}`, () => {
      assert.equal(validateCallerTimezone(legit), canonical);
    });
  }

  it('rejects a string that is not a real zone name and not offset-shaped', () => {
    assert.throws(
      () => validateCallerTimezone('Not/AZone'),
      /is not a time zone this server can resolve/
    );
  });

  it('echoes the caller-supplied value in an unresolvable-zone rejection, bounded', () => {
    const long = 'Not/A'.repeat(20);
    assert.throws(() => validateCallerTimezone(long), (err: Error) => {
      // Bounded rather than reflecting the full (potentially huge) input verbatim.
      assert.ok(err.message.length < long.length + 200);
      return true;
    });
  });
});

// The window bounds read the TIME out of a caller's value, which the shape pattern alone does
// not validate — so these are the cases where "the calendar pair rejects what coerceUtcDate
// rejects" stops being true by construction and has to be asserted.
describe('calendar window bounds reject a time of day that does not exist (#138)', () => {
  const SYDNEY = 'Australia/Sydney';

  it('rejects an out-of-range wall clock instead of rolling it into another day', () => {
    // Date.UTC ROLLS these rather than failing: 25:00 became 15:00Z the same day and 99:99:99
    // became a window starting three and a half days later — silently, since only the
    // one-sided clamp says anything about a window that moved. coerceUtcDate throws on all
    // three, and so does create_calendar_event, so the family accepted on a read what it
    // refused on a write.
    for (const value of ['2026-08-12T25:00:00', '2026-08-12T12:75:00', '2026-08-12T99:99:99', '2026-08-12T23:59:60']) {
      for (const coerce of [coerceCalendarWindowStart, coerceCalendarWindowEnd]) {
        assert.throws(
          () => coerce(value, 'startDate', SYDNEY),
          (err: Error) => {
            assert.equal(err.name, 'InvalidInputError');
            assert.match(err.message, /is not a valid date/);
            return true;
          },
          `expected rejection for ${value}`,
        );
      }
      // The pinned parity: the same value, through the UTC coercion the email bounds use.
      assert.throws(() => coerceUtcDate(value, 'after'), /is not a valid date/);
    }
  });

  it('accepts 24:00:00 as the end of the day, exactly as the UTC coercion does', () => {
    // The ECMAScript Date Time String Format allows it, so `new Date()` takes it and
    // coerceUtcDate accepts it. Rejecting it here would be the same divergence pointing the
    // other way.
    assert.equal(coerceCalendarWindowStart('2026-08-12T24:00:00', 'startDate', 'UTC'), '2026-08-13T00:00:00Z');
    assert.equal(coerceUtcDate('2026-08-12T24:00:00Z', 'after'), '2026-08-13T00:00:00Z');
  });

  it('reads a year below 0100 as that year, not as the 1900s', () => {
    // Date.UTC maps years 0-99 to 1900-1999, so `0026-08-12` resolved to a window in 1926
    // while coerceUtcDate on the same value correctly returned the year 26.
    assert.equal(coerceCalendarWindowStart('0026-08-12', 'startDate', 'UTC'), '0026-08-12T00:00:00Z');
    assert.equal(coerceUtcDate('0026-08-12', 'after'), '0026-08-12T00:00:00Z');
    // In a zone with an offset it is still that year, whatever the offset of the day turns
    // out to have been.
    assert.match(coerceCalendarWindowStart('0026-08-12', 'startDate', 'Australia/Sydney')!, /^0026-/);
  });

  it('reads year 0000 as year 0, not as the year Intl calls "1 BC"', () => {
    // The offset is read back out of Intl.formatToParts, and WITHOUT an `era` in the options
    // Intl prints the ERA-RELATIVE year: proleptic year 0 formats as "1". That number went
    // back in as though it were the proleptic year, so the offset came out a whole year
    // wrong and the bound landed on the wrong DAY with nothing said — `0000-12-31` resolved
    // to 0000-01-01, a silently different window of exactly the kind this section exists for.
    assert.equal(coerceCalendarWindowStart('0000-12-31', 'startDate', 'UTC'), '0000-12-31T00:00:00Z');
    assert.equal(coerceCalendarWindowStart('0000-01-01', 'startDate', 'UTC'), '0000-01-01T00:00:00Z');
    assert.equal(coerceCalendarWindowEnd('0000-12-31', 'endDate', 'UTC'), '0001-01-01T00:00:00Z');
    // And the year either side of the era boundary, so a sign error cannot pass.
    assert.equal(coerceCalendarWindowStart('0001-01-01', 'startDate', 'UTC'), '0001-01-01T00:00:00Z');
    // In an offset zone the DAY still has to be the day asked for. Sydney was on LMT that
    // far back, so the instant carries a minutes-and-seconds offset rather than a round hour;
    // what matters is that it sits within a day of the date named, not a year away.
    const sydney = coerceCalendarWindowStart('0000-12-31', 'startDate', 'Australia/Sydney')!;
    assert.match(sydney, /^0000-12-3[01]T/, sydney);
  });
});

// Caller text quoted back inside an error message is untrusted content in a channel an agent
// reads as trusted. One helper decides how, so the policy cannot differ per message.
describe('echoCallerText is the one echo policy (#141)', () => {
  it('strips the control characters that would forge extra lines in a message', () => {
    // Measured before this converged: coerceUtcDate echoed a raw ESC straight through while
    // the calendar window's backwards-range error scrubbed the identical value.
    const withEsc = '2026-08-12T\u001B[31mBAD\u2028INJECTED';
    assert.throws(
      () => coerceUtcDate(withEsc, 'after'),
      (err: Error) => {
        assert.ok(!err.message.includes('\u001B'), err.message);
        assert.ok(!err.message.includes('\u2028'), err.message);
        return true;
      },
    );
  });

  it('trims, so the value quoted is the value that was judged', () => {
    // The coercion trims before validating, so echoing the padding back quotes a string the
    // server never looked at.
    assert.throws(
      () => coerceUtcDate('   2026-13-45   ', 'after'),
      (err: Error) => {
        assert.match(err.message, /"2026-13-45"/);
        return true;
      },
    );
  });

  it('names an unresolvable timezone with a visible truncation marker', () => {
    // Silently cutting at 40 characters printed a name the caller could neither recognise nor
    // correct.
    const long = `Not/AZone${'x'.repeat(80)}`;
    const described = describeTimezone(long);
    assert.match(described, /…/);
    assert.ok(!described.includes('x'.repeat(41)), described);
  });
});

// A wall clock a DST transition skips or repeats names no instant, or two. Both are resolved
// the way RFC 5545 and Temporal's `compatible` disambiguation resolve them.
describe('calendar window bounds resolve a DST gap forward (#138)', () => {
  it('does not drop the last hour of a day whose midnight transition skips it', () => {
    // Santiago springs forward AT MIDNIGHT on 2026-09-06, so local midnight that day does not
    // exist. Resolving it backward — to the last instant before the gap — ended a single-day
    // window at local 23:00 and an event at 23:30 on the 5th was never searched for, with
    // nothing saying so.
    const start = coerceCalendarWindowStart('2026-09-05', 'startDate', 'America/Santiago');
    const end = coerceCalendarWindowEnd('2026-09-05', 'endDate', 'America/Santiago');
    assert.equal(start, '2026-09-05T04:00:00Z');
    assert.equal(end, '2026-09-06T04:00:00Z');
    // 24 hours of UTC covering a 23-hour local day: the missing hour is the one the clocks
    // skipped, not one the window dropped.
    assert.equal((Date.parse(end!) - Date.parse(start!)) / 3600000, 24);
  });

  it('resolves a skipped wall-clock datetime forward, to the transition instant', () => {
    // Sydney 2026-10-04 02:30 does not exist (02:00 jumps to 03:00). Forward is 03:30 AEDT.
    assert.equal(coerceCalendarWindowStart('2026-10-04T02:30:00', 'startDate', 'Australia/Sydney'), '2026-10-03T16:30:00Z');
  });

  it('resolves a repeated wall clock to the earlier of its two instants', () => {
    // Sydney 2026-04-05 02:30 happens twice (03:00 falls back to 02:00). The earlier one is
    // still +11:00.
    assert.equal(coerceCalendarWindowStart('2026-04-05T02:30:00', 'startDate', 'Australia/Sydney'), '2026-04-04T15:30:00Z');
  });

  it('still covers a whole local day either side of a mid-day transition', () => {
    // The 23/25-hour cases, re-asserted here because the gap handling is the code that
    // computes them now.
    const spring = Date.parse(coerceCalendarWindowEnd('2026-10-04', 'e', 'Australia/Sydney')!) -
      Date.parse(coerceCalendarWindowStart('2026-10-04', 's', 'Australia/Sydney')!);
    const fall = Date.parse(coerceCalendarWindowEnd('2026-04-05', 'e', 'Australia/Sydney')!) -
      Date.parse(coerceCalendarWindowStart('2026-04-05', 's', 'Australia/Sydney')!);
    assert.equal(spring / 3600000, 23);
    assert.equal(fall / 3600000, 25);
  });
});

describe('resolveCalendarInstantMs', () => {
  it('reads a zone-less value in the configured zone and a designated one as written', () => {
    // The TZID is gone by the time an event start reaches here, so a bare wall clock is read
    // the way the window bounds read one. A Z value names its own instant and is untouched.
    assert.equal(
      resolveCalendarInstantMs('2026-03-25T08:30:00', 'Australia/Sydney'),
      // 25 March is still +11:00 in Sydney — the offset in force on the day, not a fixed one.
      Date.parse('2026-03-24T21:30:00Z'),
    );
    assert.equal(resolveCalendarInstantMs('2026-03-25T08:00:00Z', 'Australia/Sydney'), Date.parse('2026-03-25T08:00:00Z'));
    assert.equal(resolveCalendarInstantMs('2026-03-25T08:00:00+05:30', 'Australia/Sydney'), Date.parse('2026-03-25T02:30:00Z'));
  });

  it('places an all-day date at local midnight, where the window puts one', () => {
    assert.equal(
      resolveCalendarInstantMs('2026-03-25', 'Australia/Sydney'),
      Date.parse(coerceCalendarWindowStart('2026-03-25', 'startDate', 'Australia/Sydney')!),
    );
  });

  it('returns NaN for anything it cannot read, rather than throwing', () => {
    // It orders SERVER data, not caller input: something unreadable is for the caller to
    // place, and a throw here would fail a whole listing over one odd event.
    for (const value of [undefined, '', '   ', 'whenever', 12 as any]) {
      assert.ok(Number.isNaN(resolveCalendarInstantMs(value, 'Australia/Sydney')), `for ${String(value)}`);
    }
  });
});

describe('describeTimezone names the zone that actually resolved', () => {
  it('names the host zone AND the configured value when the configured one is unusable', () => {
    // zoneOffsetMsAt falls back to the host zone on a name ICU cannot resolve, so naming the
    // configured value alone printed a zone the dates were not read in — on the one call
    // where the caller is trying to work out why their days look wrong.
    const label = describeTimezone('Not/AZone');
    assert.match(label, new RegExp(Intl.DateTimeFormat().resolvedOptions().timeZone.replace(/[/]/g, '\\/')));
    assert.match(label, /Not\/AZone/);
    assert.match(label, /not a time zone this server can resolve/);
  });

  it('says nothing extra about a zone that resolves', () => {
    assert.equal(describeTimezone('America/New_York'), 'America/New_York');
  });
});
