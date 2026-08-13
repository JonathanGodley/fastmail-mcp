import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { EMAIL_FIELD_NAMES, parseEmailFields, projectEmail, wantsHtmlBody } from './field-projection.js';
import { simplifyEmail } from './email-formatter.js';
import { InvalidInputError } from './coerce.js';

// A raw JMAP email carrying one value of every field class the simplified shape
// emits: scalars, address arrays, keyword-derived flags, the location arrays, the
// non-standard keyword map, both bodies, and attachments.
function rawEmail(overrides: Record<string, any> = {}): any {
  const raw: any = {
    id: 'e1',
    subject: 'Quarterly numbers',
    from: [{ name: 'Alice', email: 'alice@example.com' }],
    to: [{ email: 'bob@example.com' }],
    cc: [{ email: 'carol@example.com' }],
    receivedAt: '2026-01-15T09:00:00Z',
    threadId: 't1',
    messageId: ['<m1@example.com>'],
    references: ['<r1@example.com>', '<r2@example.com>'],
    inReplyTo: ['<r2@example.com>'],
    preview: 'Numbers for the quarter',
    blobId: 'blob-1',
    size: 4096,
    keywords: { $seen: true, $flagged: true, $important: true },
    textBody: [{ partId: 'text1', type: 'text/plain' }],
    htmlBody: [{ partId: 'html1', type: 'text/html' }],
    bodyValues: { text1: { value: 'Numbers for the quarter' }, html1: { value: '<p>Numbers</p>' } },
    attachments: [{ partId: 'att1', name: 'q.pdf', type: 'application/pdf', size: 10, blobId: 'blob-2' }],
    ...overrides,
  };
  return raw;
}

// Attach the location info the client layer resolves, the way attachMailboxInfo does.
function withMailboxInfo(raw: any, info: { names?: string[]; roles?: string[]; unresolved?: string[] }): any {
  if (info.names !== undefined) Object.defineProperty(raw, '_mailboxNames', { value: info.names, enumerable: false, configurable: true });
  if (info.roles !== undefined) Object.defineProperty(raw, '_mailboxRoles', { value: info.roles, enumerable: false, configurable: true });
  if (info.unresolved !== undefined) Object.defineProperty(raw, '_unresolvedMailboxIds', { value: info.unresolved, enumerable: false, configurable: true });
  return raw;
}

// ---------- the field vocabulary ----------

describe('EMAIL_FIELD_NAMES', () => {
  it('covers every field simplifyEmail can emit', () => {
    // The guard against a field being added to the simplified shape but staying
    // permanently unselectable. The type of EMAIL_FIELD_MAP catches an omission at
    // compile time; this catches the runtime half from the other direction.
    const emitted = Object.keys(
      simplifyEmail(withMailboxInfo(rawEmail({ keywords: { $seen: true, $flagged: true, $draft: true, $answered: true, $forwarded: true, $important: true } }), {
        names: ['Inbox'],
        roles: ['inbox'],
        unresolved: ['mb-gone'],
      }), { includeHtml: true }),
    );
    const vocabulary = new Set(EMAIL_FIELD_NAMES);
    const missing = emitted.filter(k => !vocabulary.has(k));
    assert.deepEqual(missing, [], `simplified fields missing from the projection vocabulary: ${missing.join(', ')}`);
  });

  it('includes the size hints, which are only emitted in some modes', () => {
    assert.ok(EMAIL_FIELD_NAMES.includes('bodyHtmlSize'));
    assert.ok(EMAIL_FIELD_NAMES.includes('bodyTextSize'));
  });
});

// ---------- parseEmailFields ----------

describe('parseEmailFields', () => {
  it('returns undefined when the parameter is absent', () => {
    assert.equal(parseEmailFields(undefined), undefined);
    assert.equal(parseEmailFields(null), undefined);
  });

  it('accepts an array of valid names', () => {
    const fields = parseEmailFields(['id', 'subject', 'date']);
    assert.deepEqual([...fields!].sort(), ['date', 'id', 'subject']);
  });

  it('accepts a comma-separated string from a lenient client', () => {
    const fields = parseEmailFields('id, subject ,date');
    assert.deepEqual([...fields!].sort(), ['date', 'id', 'subject']);
  });

  it('accepts a stringified JSON array from a lenient client', () => {
    const fields = parseEmailFields('["id","threadId"]');
    assert.deepEqual([...fields!].sort(), ['id', 'threadId']);
  });

  it('trims surrounding whitespace on array entries', () => {
    const fields = parseEmailFields([' id ', 'subject']);
    assert.deepEqual([...fields!].sort(), ['id', 'subject']);
  });

  it('de-duplicates repeated names', () => {
    const fields = parseEmailFields(['id', 'id', 'subject']);
    assert.equal(fields!.size, 2);
  });

  it('rejects an unknown name and lists the valid set', () => {
    assert.throws(
      () => parseEmailFields(['id', 'sender']),
      (err: unknown) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match((err as Error).message, /"sender"/);
        assert.match((err as Error).message, /Valid fields:/);
        assert.match((err as Error).message, /threadId/);
        return true;
      },
    );
  });

  it('rejects the whole call when one name of several is unknown', () => {
    // Never partially honour a projection: a typo that silently returned the other
    // fields would read as "that field is always empty".
    assert.throws(() => parseEmailFields(['id', 'subjct', 'from']), InvalidInputError);
  });

  it('names every unknown field in one message', () => {
    assert.throws(
      () => parseEmailFields(['sender', 'body']),
      (err: unknown) => {
        assert.match((err as Error).message, /"sender"/);
        assert.match((err as Error).message, /"body"/);
        return true;
      },
    );
  });

  it('is case-sensitive (the simplified names are camelCase)', () => {
    assert.throws(() => parseEmailFields(['ThreadId']), InvalidInputError);
  });

  it('rejects an empty array rather than returning the full shape', () => {
    assert.throws(
      () => parseEmailFields([]),
      (err: unknown) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match((err as Error).message, /cannot be empty/);
        assert.match((err as Error).message, /omit the parameter/);
        return true;
      },
    );
  });

  it('rejects an empty string the same way (a lenient client sending "")', () => {
    assert.throws(() => parseEmailFields(''), InvalidInputError);
  });

  it('rejects a value that is neither an array nor a string', () => {
    assert.throws(() => parseEmailFields(42), InvalidInputError);
    assert.throws(() => parseEmailFields({ id: true }), InvalidInputError);
  });

  it('rejects fields together with raw:true instead of letting raw win silently', () => {
    assert.throws(
      () => parseEmailFields(['id'], { raw: true }),
      (err: unknown) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match((err as Error).message, /raw:true/);
        return true;
      },
    );
  });

  it('rejects raw:true even for an otherwise invalid fields value', () => {
    // The raw conflict is reported first: it is the one the caller must resolve.
    assert.throws(
      () => parseEmailFields([], { raw: true }),
      (err: unknown) => {
        assert.match((err as Error).message, /raw:true/);
        return true;
      },
    );
  });

  it('leaves raw:false unaffected', () => {
    assert.deepEqual([...parseEmailFields(['id'], { raw: false })!], ['id']);
  });

  it('does not reject raw:true when no fields were requested', () => {
    assert.equal(parseEmailFields(undefined, { raw: true }), undefined);
  });
});

// ---------- wantsHtmlBody ----------

describe('wantsHtmlBody', () => {
  it('is true only when bodyHtml is projected', () => {
    assert.equal(wantsHtmlBody(new Set(['bodyHtml'])), true);
    assert.equal(wantsHtmlBody(new Set(['bodyText', 'subject'])), false);
    assert.equal(wantsHtmlBody(undefined), false);
  });

  it('is false for the size hint alone (a hint is not the body)', () => {
    assert.equal(wantsHtmlBody(new Set(['bodyHtmlSize'])), false);
  });
});

// ---------- projectEmail ----------

describe('projectEmail', () => {
  const simplified = simplifyEmail(withMailboxInfo(rawEmail(), { names: ['Inbox'], roles: ['inbox'] }));

  it('returns the email untouched when no projection was requested', () => {
    assert.equal(projectEmail(simplified, undefined), simplified);
  });

  it('keeps only the projected fields', () => {
    const projected = projectEmail(simplified, new Set(['id', 'subject', 'date', 'threadId', 'to']));
    assert.deepEqual(Object.keys(projected).sort(), ['date', 'id', 'subject', 'threadId', 'to']);
  });

  it('preserves each field class unchanged: scalar, address array, flag, keyword map, attachments', () => {
    const projected = projectEmail(simplified, new Set(['id', 'to', 'isRead', 'keywords', 'attachments']));
    assert.equal(projected.id, 'e1');
    assert.deepEqual(projected.to, ['bob@example.com']);
    assert.equal(projected.isRead, true);
    assert.deepEqual(projected.keywords, { $important: true });
    assert.deepEqual(projected.attachments, [
      { contentType: 'application/pdf', size: 10, blobId: 'blob-2', partId: 'att1', name: 'q.pdf' },
    ]);
  });

  it('drops the threading plumbing and preview that dominate a wide listing', () => {
    const projected = projectEmail(simplified, new Set(['id', 'subject', 'date']));
    for (const dropped of ['references', 'messageId', 'inReplyTo', 'blobId', 'keywords', 'preview']) {
      assert.equal(dropped in projected, false, `${dropped} should have been projected away`);
    }
  });

  it('omits a projected field the message does not have, without inventing it', () => {
    const projected = projectEmail(simplified, new Set(['bcc']));
    assert.deepEqual(projected, {});
  });

  it('emits unresolvedMailboxIds alongside a projected mailboxes', () => {
    // The degradation half of mailboxes/roles rides along uninvited: a short
    // location array that silently omitted the ids it could not resolve would be
    // the #53 failure reached through a new door.
    const degraded = simplifyEmail(withMailboxInfo(rawEmail(), { names: ['Inbox'], roles: ['inbox'], unresolved: ['mb-gone'] }));
    const projected = projectEmail(degraded, new Set(['id', 'mailboxes']));
    assert.deepEqual(projected.mailboxes, ['Inbox']);
    assert.deepEqual(projected.unresolvedMailboxIds, ['mb-gone']);
  });

  it('emits unresolvedMailboxIds alongside a projected roles', () => {
    const degraded = simplifyEmail(withMailboxInfo(rawEmail(), { names: ['Inbox'], roles: ['inbox'], unresolved: ['mb-gone'] }));
    const projected = projectEmail(degraded, new Set(['roles']));
    assert.deepEqual(projected.unresolvedMailboxIds, ['mb-gone']);
  });

  it('does not emit unresolvedMailboxIds when no location field was projected', () => {
    const degraded = simplifyEmail(withMailboxInfo(rawEmail(), { names: ['Inbox'], unresolved: ['mb-gone'] }));
    const projected = projectEmail(degraded, new Set(['id', 'subject']));
    assert.equal('unresolvedMailboxIds' in projected, false);
  });

  it('does not duplicate unresolvedMailboxIds when it was projected explicitly', () => {
    const degraded = simplifyEmail(withMailboxInfo(rawEmail(), { names: ['Inbox'], unresolved: ['mb-gone'] }));
    const projected = projectEmail(degraded, new Set(['mailboxes', 'unresolvedMailboxIds']));
    assert.deepEqual(projected.unresolvedMailboxIds, ['mb-gone']);
    assert.equal(Object.keys(projected).filter(k => k === 'unresolvedMailboxIds').length, 1);
  });

  it('adds nothing when mailbox resolution succeeded', () => {
    const projected = projectEmail(simplified, new Set(['mailboxes']));
    assert.deepEqual(projected, { mailboxes: ['Inbox'] });
  });
});

// ---------- the get_email read path (#69) ----------

describe('get_email field projection', () => {
  // Mirrors the two lines the get_email handler runs: simplify (with the HTML body
  // turned on when it was projected), then project.
  function readEmail(raw: any, options: { verbose?: boolean; fields?: unknown }): Record<string, any> {
    const fields = parseEmailFields(options.fields);
    return projectEmail(simplifyEmail(raw, { includeHtml: !!options.verbose || wantsHtmlBody(fields) }), fields) as Record<string, any>;
  }

  it('returns the HTML body alone, with no verbose and no other field', () => {
    const result = readEmail(rawEmail(), { fields: ['bodyHtml'] });
    assert.deepEqual(result, { bodyHtml: '<p>Numbers</p>' });
  });

  it('returns the text body alone', () => {
    const result = readEmail(rawEmail(), { fields: ['bodyText'] });
    assert.deepEqual(result, { bodyText: 'Numbers for the quarter' });
  });

  it('returns the full default shape when no projection is requested', () => {
    const result = readEmail(rawEmail(), {});
    assert.equal(result.subject, 'Quarterly numbers');
    assert.equal(result.bodyText, 'Numbers for the quarter');
    assert.equal(result.bodyHtml, undefined);
    assert.equal(result.bodyHtmlSize, '<p>Numbers</p>'.length);
  });

  it('an html-only message projected to bodyText comes back empty, not silently HTML', () => {
    // simplifyEmail's html-as-fallback only fires for the whole shape; a caller that
    // asked for bodyText specifically is told there is none rather than handed markup.
    const htmlOnly = rawEmail({ textBody: [], bodyValues: { html1: { value: '<p>Numbers</p>' } } });
    assert.deepEqual(readEmail(htmlOnly, { fields: ['bodyText'] }), {});
    assert.deepEqual(readEmail(htmlOnly, { fields: ['bodyHtml'] }), { bodyHtml: '<p>Numbers</p>' });
  });

  it('verbose and fields compose: verbose supplies the HTML, fields narrows the output', () => {
    const result = readEmail(rawEmail(), { verbose: true, fields: ['id', 'bodyHtml'] });
    assert.deepEqual(result, { id: 'e1', bodyHtml: '<p>Numbers</p>' });
  });
});
