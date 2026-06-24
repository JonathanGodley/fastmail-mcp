import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { coerceStringArray, coerceRecipients, coerceBool, redactBearerTokens, requireNonEmpty, validateClearFields, parseAddress, assertKnownParams } from './coerce.js';
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
    assert.deepEqual(coerceStringArray(['a@b.com', 'c@d.com']), ['a@b.com', 'c@d.com']);
  });

  it('stringifies array elements', () => {
    assert.deepEqual(coerceStringArray([1, 2, 3] as any), ['1', '2', '3']);
  });

  it('parses JSON-stringified array', () => {
    assert.deepEqual(coerceStringArray('["a@b.com", "c@d.com"]'), ['a@b.com', 'c@d.com']);
  });

  it('parses JSON-stringified array with whitespace', () => {
    assert.deepEqual(coerceStringArray('  ["a@b.com"]  '), ['a@b.com']);
  });

  it('splits comma-separated string', () => {
    assert.deepEqual(coerceStringArray('a@b.com, c@d.com'), ['a@b.com', 'c@d.com']);
  });

  it('wraps single address as one-item array', () => {
    assert.deepEqual(coerceStringArray('single@example.com'), ['single@example.com']);
  });

  it('returns empty array for empty string', () => {
    assert.deepEqual(coerceStringArray(''), []);
  });

  it('trims whitespace and filters empty segments in comma-split', () => {
    assert.deepEqual(coerceStringArray('a@b.com, ,c@d.com,'), ['a@b.com', 'c@d.com']);
  });

  it('falls back to comma-split when JSON parsing fails', () => {
    assert.deepEqual(coerceStringArray('[not valid json]'), ['[not valid json]']);
  });
});

describe('coerceRecipients', () => {
  it('coerces all four fields from arrays, JSON-strings, comma-strings, and bare strings', () => {
    const result = coerceRecipients({
      to: ['a@b.com'],
      cc: '["c@d.com", "e@f.com"]',
      bcc: 'g@h.com, i@j.com',
      replyTo: 'k@l.com',
    });
    assert.deepEqual(result, {
      to: ['a@b.com'],
      cc: ['c@d.com', 'e@f.com'],
      bcc: ['g@h.com', 'i@j.com'],
      replyTo: ['k@l.com'],
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
    const out = redactBearerTokens(
      'Failed: token fmu1-3b1e4048-036f4f86690cd04d8d05105a369ee30b-0-dbfc727af72d5e3e27dd324675869337 invalid'
    );
    assert.match(out, /fmu\[REDACTED\]/);
    assert.ok(!out.includes('fmu1-3b1e'));
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
});

describe('parseAddress', () => {
  it('parses "Name <email>" into name + email', () => {
    assert.deepEqual(parseAddress('Alice <a@x.com>'), { name: 'Alice', email: 'a@x.com' });
  });

  it('strips surrounding double-quotes from a quoted display name', () => {
    assert.deepEqual(parseAddress('"Doe, John" <j@x.com>'), { name: 'Doe, John', email: 'j@x.com' });
  });

  it('passes a bare address through as { email }', () => {
    assert.deepEqual(parseAddress('a@x.com'), { email: 'a@x.com' });
  });

  it('omits the name when angle brackets carry no display name', () => {
    assert.deepEqual(parseAddress('<a@x.com>'), { email: 'a@x.com' });
  });

  it('trims surrounding whitespace and an empty name', () => {
    assert.deepEqual(parseAddress('   <a@x.com>  '), { email: 'a@x.com' });
  });

  it('trims whitespace around name and email', () => {
    assert.deepEqual(parseAddress('  Bob   <  b@x.com  >  '), { name: 'Bob', email: 'b@x.com' });
  });

  it('uses the last angle-bracket pair so a name may contain "<"', () => {
    assert.deepEqual(parseAddress('a<b <c@x.com>'), { name: 'a<b', email: 'c@x.com' });
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
