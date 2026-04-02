import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { simplifyEmail, formatAddress } from './email-formatter.js';

describe('formatAddress', () => {
  it('formats name and email', () => {
    assert.equal(formatAddress({ name: 'Alice', email: 'alice@example.com' }), 'Alice <alice@example.com>');
  });

  it('returns email only when no name', () => {
    assert.equal(formatAddress({ email: 'bob@example.com' }), 'bob@example.com');
  });

  it('returns unknown for null/undefined', () => {
    assert.equal(formatAddress(null as any), 'unknown');
    assert.equal(formatAddress(undefined as any), 'unknown');
  });
});

describe('simplifyEmail', () => {
  it('extracts all fields from a standard email', () => {
    const raw = {
      id: 'e1',
      threadId: 't1',
      messageId: ['msg-1@example.com'],
      references: ['msg-0@example.com'],
      subject: 'Hello',
      from: [{ name: 'Alice', email: 'alice@example.com' }],
      to: [{ name: 'Bob', email: 'bob@example.com' }],
      cc: [{ email: 'carol@example.com' }],
      bcc: [{ email: 'secret@example.com' }],
      receivedAt: '2026-03-01T12:00:00Z',
      inReplyTo: ['msg-0@example.com'],
      keywords: { $seen: true, $flagged: true },
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: {
        '1': { value: 'Hello world' },
        '2': { value: '<p>Hello world</p>' },
      },
      attachments: [
        { name: 'doc.pdf', type: 'application/pdf', size: 1024, blobId: 'b1' },
      ],
    };

    const result = simplifyEmail(raw);

    assert.equal(result.id, 'e1');
    assert.equal(result.threadId, 't1');
    assert.deepEqual(result.messageId, ['msg-1@example.com']);
    assert.deepEqual(result.references, ['msg-0@example.com']);
    assert.equal(result.subject, 'Hello');
    assert.equal(result.from, 'Alice <alice@example.com>');
    assert.deepEqual(result.to, ['Bob <bob@example.com>']);
    assert.deepEqual(result.cc, ['carol@example.com']);
    assert.deepEqual(result.bcc, ['secret@example.com']);
    assert.equal(result.date, '2026-03-01T12:00:00Z');
    assert.equal(result.isReply, true);
    assert.equal(result.isRead, true);
    assert.equal(result.isFlagged, true);
    assert.equal(result.bodyText, 'Hello world');
    // bodyHtml omitted by default, bodyHtmlSize provided instead
    assert.equal(result.bodyHtml, undefined);
    assert.equal(result.bodyHtmlSize, '<p>Hello world</p>'.length);
    assert.equal(result.attachments!.length, 1);
    assert.equal(result.attachments![0].name, 'doc.pdf');
    assert.equal(result.attachments![0].contentType, 'application/pdf');
    assert.equal(result.attachments![0].size, 1024);
    assert.equal(result.attachments![0].blobId, 'b1');
  });

  it('returns null bodyText for HTML-only email', () => {
    const raw = {
      id: 'e2',
      textBody: [],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<b>Bold</b>' } },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, undefined);
    assert.equal(result.bodyHtml, '<b>Bold</b>');
  });

  it('returns null bodyHtml for plain-text-only email', () => {
    const raw = {
      id: 'e3',
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [],
      bodyValues: { t: { value: 'Just text' } },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, 'Just text');
    assert.equal(result.bodyHtml, undefined);
  });

  it('omits bodies when bodyValues is missing', () => {
    const raw = {
      id: 'e4',
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, undefined);
    assert.equal(result.bodyHtml, undefined);
  });

  it('omits bodies when bodyValues is empty', () => {
    const raw = {
      id: 'e5',
      textBody: [{ partId: '1', type: 'text/plain' }],
      bodyValues: {},
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, undefined);
  });

  it('concatenates multi-part text bodies', () => {
    const raw = {
      id: 'e6',
      textBody: [
        { partId: 'a', type: 'text/plain' },
        { partId: 'b', type: 'text/plain' },
      ],
      bodyValues: {
        a: { value: 'Part one' },
        b: { value: 'Part two' },
      },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, 'Part one\nPart two');
  });

  it('surfaces isTruncated flag', () => {
    const raw = {
      id: 'e7',
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'Partial...', isTruncated: true } },
    };
    const result = simplifyEmail(raw);
    assert.match(result.bodyText!, /\[body truncated\]/);
  });

  it('surfaces isEncodingProblem flag', () => {
    const raw = {
      id: 'e8',
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'Garbled', isEncodingProblem: true } },
    };
    const result = simplifyEmail(raw);
    assert.match(result.bodyText!, /\[encoding issues detected\]/);
  });

  it('handles email with attachments', () => {
    const raw = {
      id: 'e9',
      attachments: [
        { name: 'photo.jpg', type: 'image/jpeg', size: 5000, blobId: 'b1' },
        { type: 'image/png', size: 200, blobId: 'b2' },
      ],
    };
    const result = simplifyEmail(raw);
    assert.equal(result.attachments!.length, 2);
    assert.equal(result.attachments![0].contentType, 'image/jpeg');
    assert.equal(result.attachments![0].name, 'photo.jpg');
    assert.equal(result.attachments![1].name, undefined);
    assert.equal(result.attachments![1].contentType, 'image/png');
  });

  it('omits attachments when empty', () => {
    const raw = { id: 'e10' };
    const result = simplifyEmail(raw);
    assert.equal(result.attachments, undefined);
  });

  it('detects reply emails', () => {
    const raw = {
      id: 'e11',
      inReplyTo: ['msg-123@example.com'],
    };
    assert.equal(simplifyEmail(raw).isReply, true);
  });

  it('omits noise-when-false flags but keeps isRead', () => {
    const result = simplifyEmail({ id: 'e12', inReplyTo: [] });
    assert.equal(result.isReply, undefined, 'isReply false should be omitted');
    assert.equal(result.isRead, false, 'isRead false should be preserved (unread is meaningful)');
    assert.equal(result.isFlagged, undefined, 'isFlagged false should be omitted');
    assert.equal(result.isDraft, undefined, 'isDraft false should be omitted');
  });

  it('omits null/empty fields', () => {
    const result = simplifyEmail({ id: 'e15' });
    assert.equal(result.id, 'e15');
    assert.equal(result.subject, '(no subject)');
    assert.equal(result.from, 'unknown');
    // These should all be omitted, not present as null/[]
    assert.equal(result.to, undefined);
    assert.equal(result.cc, undefined);
    assert.equal(result.bcc, undefined);
    assert.equal(result.date, undefined);
    assert.equal(result.threadId, undefined);
    assert.equal(result.messageId, undefined);
    assert.equal(result.references, undefined);
    assert.equal(result.attachments, undefined);
  });

  it('includes references as a proper field', () => {
    const raw = {
      id: 'e19',
      references: ['msg-1@example.com', 'msg-2@example.com'],
    };
    const result = simplifyEmail(raw);
    assert.deepEqual(result.references, ['msg-1@example.com', 'msg-2@example.com']);
  });

  it('includes preview when present', () => {
    const raw = {
      id: 'e20',
      preview: 'Hey, just wanted to check in about...',
    };
    const result = simplifyEmail(raw);
    assert.equal(result.preview, 'Hey, just wanted to check in about...');
  });

  it('omits preview when absent', () => {
    const result = simplifyEmail({ id: 'e21' });
    assert.equal(result.preview, undefined);
  });

  it('includes hasAttachment when true and no attachments array', () => {
    const raw = {
      id: 'e22',
      hasAttachment: true,
    };
    const result = simplifyEmail(raw);
    assert.equal(result.hasAttachment, true);
  });

  it('omits hasAttachment when false', () => {
    const raw = {
      id: 'e23',
      hasAttachment: false,
    };
    const result = simplifyEmail(raw);
    assert.equal(result.hasAttachment, undefined);
  });

  it('drops hasAttachment when attachments array is present', () => {
    const raw = {
      id: 'e24',
      hasAttachment: true,
      attachments: [{ type: 'image/png', size: 100, blobId: 'b1' }],
    };
    const result = simplifyEmail(raw);
    assert.equal(result.hasAttachment, undefined);
    assert.equal(result.attachments!.length, 1);
  });

  // --- bodyHtml omission and bodyHtmlSize ---

  it('omits bodyHtml by default, provides bodyHtmlSize', () => {
    const raw = {
      id: 'e30',
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: {
        '1': { value: 'Plain text' },
        '2': { value: '<p>HTML content that is much longer</p>' },
      },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, 'Plain text');
    assert.equal(result.bodyHtml, undefined);
    assert.equal(result.bodyHtmlSize, '<p>HTML content that is much longer</p>'.length);
  });

  it('includes bodyHtml when includeHtml is true', () => {
    const raw = {
      id: 'e31',
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: {
        '1': { value: 'Plain text' },
        '2': { value: '<p>HTML</p>' },
      },
    };
    const result = simplifyEmail(raw, { includeHtml: true });
    assert.equal(result.bodyText, 'Plain text');
    assert.equal(result.bodyHtml, '<p>HTML</p>');
    assert.equal(result.bodyHtmlSize, undefined);
  });

  it('handles real HTML-only email (both arrays point to same text/html part)', () => {
    // Real JMAP behavior: HTML-only emails have both textBody and htmlBody
    // pointing to the same text/html part with the same partId
    const raw = {
      id: 'e32',
      textBody: [{ partId: '1', type: 'text/html' }],
      htmlBody: [{ partId: '1', type: 'text/html' }],
      bodyValues: {
        '1': { value: '<html><body><h1>Newsletter</h1></body></html>' },
      },
    };
    const result = simplifyEmail(raw);
    // bodyText should be null (no text/plain parts)
    assert.equal(result.bodyText, undefined);
    // bodyHtml auto-included as fallback since bodyText is null
    assert.equal(result.bodyHtml, '<html><body><h1>Newsletter</h1></body></html>');
    assert.equal(result.bodyHtmlSize, undefined);
  });

  it('plain-text-only email has no bodyHtml or bodyHtmlSize', () => {
    const raw = {
      id: 'e33',
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '1', type: 'text/plain' }],
      bodyValues: {
        '1': { value: 'Just plain text' },
      },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, 'Just plain text');
    assert.equal(result.bodyHtml, undefined);
    assert.equal(result.bodyHtmlSize, undefined);
  });

  it('parts with no type field pass through (defensive)', () => {
    const raw = {
      id: 'e34',
      textBody: [{ partId: '1' }],
      bodyValues: {
        '1': { value: 'Content without type' },
      },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.bodyText, 'Content without type');
  });

  it('formats replyTo addresses', () => {
    const raw = {
      id: 'e35',
      replyTo: [
        { name: 'Support', email: 'support@example.com' },
        { email: 'noreply@example.com' },
      ],
    };
    const result = simplifyEmail(raw);
    assert.deepEqual(result.replyTo, ['Support <support@example.com>', 'noreply@example.com']);
  });

});

// ==========================================================================
// Functional test issues — these tests document gaps found during live testing.
// Written as TDD: tests first, then fix the code.
// ==========================================================================

// ---------- inReplyTo preservation ----------

describe('simplifyEmail inReplyTo', () => {
  it('preserves inReplyTo Message-IDs', () => {
    const raw = {
      id: 'e-reply-1',
      subject: 'Re: Test',
      from: [{ email: 'a@b.com' }],
      inReplyTo: ['<msg-123@example.com>'],
    };
    const result = simplifyEmail(raw);
    assert.deepEqual(result.inReplyTo, ['<msg-123@example.com>']);
    assert.equal(result.isReply, true);
  });

  it('omits inReplyTo when null', () => {
    const raw = {
      id: 'e-reply-2',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      inReplyTo: null,
    };
    const result = simplifyEmail(raw);
    assert.equal(result.inReplyTo, undefined);
    assert.equal(result.isReply, undefined); // false is omitted by addIf
  });
});

// ---------- non-standard keywords ----------

describe('simplifyEmail non-standard keywords', () => {
  it('surfaces non-standard keywords in verbose mode', () => {
    const raw = {
      id: 'e-kw-1',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      keywords: {
        '$seen': true,
        '$flagged': true,
        '$canunsubscribe': true,
        '$x-me-annot-2': true,
        'custom-label': true,
      },
    };
    const result = simplifyEmail(raw, { includeHtml: true });
    assert.equal(result.isRead, true);
    assert.equal(result.isFlagged, true);
    assert.ok(result.keywords !== undefined, 'should surface non-standard keywords');
    assert.equal(result.keywords['$canunsubscribe'], true);
    assert.equal(result.keywords['$x-me-annot-2'], true);
    assert.equal(result.keywords['custom-label'], true);
  });

  it('omits keywords field when no non-standard keywords exist', () => {
    const raw = {
      id: 'e-kw-2',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      keywords: { '$seen': true },
    };
    const result = simplifyEmail(raw);
    assert.equal(result.keywords, undefined);
  });

  it('surfaces non-standard keywords even in default mode', () => {
    const raw = {
      id: 'e-kw-3',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      keywords: {
        '$seen': true,
        '$canunsubscribe': true,
      },
    };
    const result = simplifyEmail(raw);
    assert.ok(result.keywords !== undefined, 'should surface non-standard keywords in default mode too');
    assert.equal(result.keywords['$canunsubscribe'], true);
  });
});

// ---------- blobId and size ----------

describe('simplifyEmail blobId and size', () => {
  it('includes blobId when present', () => {
    const raw = {
      id: 'e-blob-1',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      blobId: 'B-abc123',
    };
    const result = simplifyEmail(raw);
    assert.equal(result.blobId, 'B-abc123');
  });

  it('includes size when present', () => {
    const raw = {
      id: 'e-size-1',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      size: 48210,
    };
    const result = simplifyEmail(raw);
    assert.equal(result.size, 48210);
  });

});

// ---------- attachment partId ----------

describe('simplifyEmail attachment partId', () => {
  it('includes partId in simplified attachments', () => {
    const raw = {
      id: 'e-att-1',
      subject: 'Test',
      from: [{ email: 'a@b.com' }],
      attachments: [
        { partId: 'part-1', blobId: 'B-123', type: 'application/pdf', size: 1024, name: 'doc.pdf' },
      ],
    };
    const result = simplifyEmail(raw);
    assert.equal(result.attachments![0].partId, 'part-1');
  });
});
