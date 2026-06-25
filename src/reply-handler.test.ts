import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyParams, composeReply } from './reply-handler.js';
import type { ReplyClient } from './reply-handler.js';

// A raw-JMAP-shaped original (as getEmailById returns it).
function makeOriginal(over: any = {}) {
  return {
    messageId: ['orig-msg@example.com'],
    references: ['root@example.com'],
    subject: 'Project update',
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>original html</p>' } },
    ...over,
  };
}

describe('buildReplyParams — quoteOriginal wiring', () => {
  it('defaults quoteOriginal to true when omitted (quote IS appended to both bodies)', () => {
    const { quoteOriginal, replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply', htmlBody: '<p>my reply</p>' },
      makeOriginal(),
    );
    assert.equal(quoteOriginal, true);
    assert.match(replyParams.textBody!, /my reply\n\nOn .*wrote:\n> original text/);
    assert.match(replyParams.htmlBody!, /<blockquote type="cite"[^>]*>.*original html/s);
  });

  it('omits the quote when quoteOriginal is false', () => {
    const { quoteOriginal, replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply', quoteOriginal: false },
      makeOriginal(),
    );
    assert.equal(quoteOriginal, false);
    assert.equal(replyParams.textBody, 'my reply');
    assert.doesNotMatch(replyParams.textBody!, /wrote:/);
  });

  it('coerces a stringified quoteOriginal ("false") like a lenient client sends', () => {
    const { quoteOriginal, replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply', quoteOriginal: 'false' },
      makeOriginal(),
    );
    assert.equal(quoteOriginal, false);
    assert.equal(replyParams.textBody, 'my reply');
  });

  it('threads the quoted bodies through for an html-only reply (text left for the downstream fallback)', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>html reply</p>' },
      makeOriginal(),
    );
    assert.equal(replyParams.textBody, undefined); // caller gave no text; createDraft/sendEmail add the fallback later
    assert.match(replyParams.htmlBody!, /html reply.*<blockquote/s);
  });
});

describe('buildReplyParams — send flag', () => {
  it('defaults shouldSend to false (draft is the safe default)', () => {
    assert.equal(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal()).shouldSend, false);
  });
  it('send=true → shouldSend true (transmit path)', () => {
    assert.equal(buildReplyParams({ originalEmailId: 'e1', textBody: 'x', send: true }, makeOriginal()).shouldSend, true);
  });
  it('coerces a stringified send ("true")', () => {
    assert.equal(buildReplyParams({ originalEmailId: 'e1', textBody: 'x', send: 'true' }, makeOriginal()).shouldSend, true);
  });
});

describe('buildReplyParams — subject, recipients, threading', () => {
  it('prefixes the subject with Re: (and does not double-prefix)', () => {
    assert.equal(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal({ subject: 'Hello' })).replyParams.subject, 'Re: Hello');
    assert.equal(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal({ subject: 'Re: Hello' })).replyParams.subject, 'Re: Hello');
  });
  it('defaults the recipient to the original sender', () => {
    assert.deepEqual(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal()).replyParams.to, ['jon@example.com']);
  });
  it('uses an explicit to over the original sender', () => {
    assert.deepEqual(buildReplyParams({ originalEmailId: 'e1', textBody: 'x', to: ['alice@x.com'] }, makeOriginal()).replyParams.to, ['alice@x.com']);
  });
  it('builds inReplyTo and appends to references', () => {
    const { replyParams } = buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal());
    assert.deepEqual(replyParams.inReplyTo, ['orig-msg@example.com']);
    assert.deepEqual(replyParams.references, ['root@example.com', 'orig-msg@example.com']);
  });
});

describe('buildReplyParams — validation', () => {
  it('rejects a body-less send (trim/zero-width-aware)', () => {
    assert.throws(() => buildReplyParams({ originalEmailId: 'e1', htmlBody: '   ', send: true }, makeOriginal()), /Either textBody or htmlBody is required/);
    assert.throws(() => buildReplyParams({ originalEmailId: 'e1', send: true }, makeOriginal()), /Either textBody or htmlBody is required/);
  });
  it('allows a body-less DRAFT (default send=false) — does not throw', () => {
    assert.doesNotThrow(() => buildReplyParams({ originalEmailId: 'e1' }, makeOriginal()));
  });
  it('throws when the original has no Message-ID', () => {
    assert.throws(() => buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal({ messageId: undefined })), /does not have a Message-ID/);
  });
  it('throws when no recipient can be determined', () => {
    assert.throws(() => buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal({ from: [] })), /Could not determine reply recipient/);
  });
});

describe('composeReply — attachment threading into both branches', () => {
  const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];

  // A spy ReplyClient: records what each method was called with; uploadAttachments
  // returns the canned UPLOADED parts so we can assert they thread through.
  function spyClient(over: Partial<ReplyClient> = {}) {
    const calls: any = {};
    const client: ReplyClient = {
      getEmailById: async (id) => { calls.getId = id; return makeOriginal(); },
      uploadAttachments: async (specs, dir) => { calls.upload = { specs, dir }; return UPLOADED; },
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
      sendEmail: async (p) => { calls.send = p; return 'sub-9'; },
      ...over,
    };
    return { client, calls };
  }

  it('uploads with the given attachDir and threads parts into the SEND branch', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', send: true, textBody: 'hi', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.equal(r.sent, true);
    assert.equal(r.submissionId, 'sub-9');
    assert.deepEqual(calls.upload, { specs: [{ path: 'a.pdf' }], dir: '/attach/root' });
    assert.deepEqual(calls.send.attachments, UPLOADED); // threaded into sendEmail
    assert.equal(calls.draft, undefined);               // draft branch not taken
  });

  it('threads parts into the DRAFT branch (send=false)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', send: false, textBody: 'hi', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.equal(r.sent, false);
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.draft.attachments, UPLOADED); // threaded into createDraft
    assert.equal(calls.send, undefined);                 // send branch not taken
  });

  it('does not call uploadAttachments when no attachments are given', async () => {
    let uploadCalled = false;
    const { client } = spyClient({ uploadAttachments: async () => { uploadCalled = true; return []; } });
    const r = await composeReply({ originalEmailId: 'o1', send: false, textBody: 'hi' }, client, '/attach/root');
    assert.equal(r.emailId, 'draft-9');
    assert.equal(uploadCalled, false);
  });

  it('requires originalEmailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeReply({ textBody: 'hi' }, client, undefined), /originalEmailId is required/);
  });
});
