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
    assert.equal(result.bodyHtml, '<p>Hello world</p>');
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

  it('omits false boolean flags', () => {
    const result = simplifyEmail({ id: 'e12', inReplyTo: [] });
    assert.equal(result.isReply, undefined);
    assert.equal(result.isRead, undefined);
    assert.equal(result.isFlagged, undefined);
    assert.equal(result.isDraft, undefined);
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
    assert.equal(result._extra, undefined);
  });

  it('includes references as a proper field', () => {
    const raw = {
      id: 'e19',
      references: ['msg-1@example.com', 'msg-2@example.com'],
    };
    const result = simplifyEmail(raw);
    assert.deepEqual(result.references, ['msg-1@example.com', 'msg-2@example.com']);
    assert.equal(result._extra, undefined);
  });

  it('captures unknown fields in _extra', () => {
    const raw = {
      id: 'e18',
      subject: 'Test',
      customField: 'surprise',
      somethingNew: 42,
    };
    const result = simplifyEmail(raw);
    assert.equal(result._extra!.customField, 'surprise');
    assert.equal(result._extra!.somethingNew, 42);
    // Known fields should NOT be in _extra
    assert.equal(result._extra!.id, undefined);
    assert.equal(result._extra!.subject, undefined);
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
    assert.equal(result._extra, undefined);
  });

  it('returns empty _extra when all fields are known', () => {
    const raw = {
      id: 'e17',
      threadId: 't1',
      messageId: ['msg@example.com'],
      references: ['ref@example.com'],
      subject: 'All known',
      from: [{ email: 'a@b.com' }],
      to: [{ email: 'c@d.com' }],
      cc: [],
      bcc: [],
      receivedAt: '2026-01-01T00:00:00Z',
      inReplyTo: null,
      keywords: { $seen: true },
      textBody: [],
      htmlBody: [],
      bodyValues: {},
      attachments: [],
      hasAttachment: false,
      preview: 'Some preview text',
    };
    const result = simplifyEmail(raw);
    assert.equal(result._extra, undefined);
  });
});
