import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { sendDraftAndMaintainKeywords, selectSource } from './send-draft-handler.js';
import type { SendDraftClient } from './send-draft-handler.js';
import { InvalidInputError } from './coerce.js';

// A mock client recording what the orchestration asked of it. Defaults describe a reply
// draft whose original resolves to exactly one message; each test overrides the one piece
// it is about.
function spyClient(over: Partial<SendDraftClient> = {}) {
  const calls: any = {};
  const client: SendDraftClient = {
    sendDraft: async (id: string) => {
      calls.sent = id;
      return {
        submissionId: 'sub-1',
        sourceReferences: { inReplyTo: ['orig-msg@example.com'], forwardedMessageId: [] },
      };
    },
    findEmailIdsByMessageId: async (messageId: string) => {
      calls.lookup = messageId;
      return ['orig-1'];
    },
    addKeywords: async (emailId: string, keywords: string[]) => {
      calls.keywords = { id: emailId, kw: keywords };
    },
    ...over,
  };
  return { client, calls };
}

describe('selectSource', () => {
  it('reads a reply from In-Reply-To', () => {
    assert.deepEqual(
      selectSource({ inReplyTo: ['a@example.com'], forwardedMessageId: [] }),
      { kind: 'reply', messageId: 'a@example.com' },
    );
  });

  it('reads a forward from X-Forwarded-Message-Id', () => {
    assert.deepEqual(
      selectSource({ inReplyTo: [], forwardedMessageId: ['f@example.com'] }),
      { kind: 'forward', messageId: 'f@example.com' },
    );
  });

  it('treats a draft carrying both headers as a reply', () => {
    assert.deepEqual(
      selectSource({ inReplyTo: ['a@example.com'], forwardedMessageId: ['f@example.com'] }),
      { kind: 'reply', messageId: 'a@example.com' },
    );
  });

  it('takes the first entry of a multi-id In-Reply-To (the conventional parent)', () => {
    assert.deepEqual(
      selectSource({ inReplyTo: ['parent@example.com', 'grandparent@example.com'], forwardedMessageId: [] }),
      { kind: 'reply', messageId: 'parent@example.com' },
    );
  });

  it('ignores blank entries and falls through to the forward header', () => {
    assert.deepEqual(
      selectSource({ inReplyTo: ['', '   '], forwardedMessageId: ['f@example.com'] }),
      { kind: 'forward', messageId: 'f@example.com' },
    );
  });

  it('returns nothing when the draft references no message', () => {
    assert.equal(selectSource({ inReplyTo: [], forwardedMessageId: [] }), undefined);
  });
});

describe('sendDraftAndMaintainKeywords', () => {
  it('requires emailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => sendDraftAndMaintainKeywords({}, client), /emailId is required/);
  });

  it('marks the original answered+read for a reply draft', async () => {
    const { client, calls } = spyClient();
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(r.submissionId, 'sub-1');
    assert.equal(calls.sent, 'd1');
    assert.equal(calls.lookup, 'orig-msg@example.com');
    assert.deepEqual(calls.keywords, { id: 'orig-1', kw: ['$answered', '$seen'] });
    assert.deepEqual(r.keywordMaintenance, {
      kind: 'reply',
      messageId: 'orig-msg@example.com',
      originalEmailId: 'orig-1',
      marked: true,
    });
  });

  it('marks the original forwarded+read for a forward draft', async () => {
    const { client, calls } = spyClient({
      sendDraft: async () => ({
        submissionId: 'sub-1',
        sourceReferences: { inReplyTo: [], forwardedMessageId: ['fwd-msg@example.com'] },
      }),
      findEmailIdsByMessageId: async () => ['orig-2'],
    });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.deepEqual(calls.keywords, { id: 'orig-2', kw: ['$forwarded', '$seen'] });
    assert.equal(r.keywordMaintenance?.kind, 'forward');
    assert.equal(r.keywordMaintenance?.marked, true);
  });

  it('treats a draft carrying both headers as a reply (never crosses the two keywords)', async () => {
    const { client, calls } = spyClient({
      sendDraft: async () => ({
        submissionId: 'sub-1',
        sourceReferences: {
          inReplyTo: ['reply-msg@example.com'],
          forwardedMessageId: ['fwd-msg@example.com'],
        },
      }),
    });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(calls.lookup, 'reply-msg@example.com');
    assert.deepEqual(calls.keywords, { id: 'orig-1', kw: ['$answered', '$seen'] });
    assert.equal(r.keywordMaintenance?.kind, 'reply');
  });

  it('writes no keywords when the draft references nothing (an ordinary compose, or a forward saved with asAttachment)', async () => {
    const { client, calls } = spyClient({
      sendDraft: async () => ({
        submissionId: 'sub-1',
        sourceReferences: { inReplyTo: [], forwardedMessageId: [] },
      }),
    });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(r.submissionId, 'sub-1');
    assert.equal(calls.lookup, undefined);
    assert.equal(calls.keywords, undefined);
    assert.equal(r.keywordMaintenance, undefined);
  });

  it('skips the write when the Message-ID matches no stored message', async () => {
    const { client, calls } = spyClient({ findEmailIdsByMessageId: async () => [] });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(calls.keywords, undefined);
    assert.equal(r.keywordMaintenance?.marked, false);
    assert.equal(r.keywordMaintenance?.skipReason, 'not-found');
    assert.equal(r.keywordMaintenance?.originalEmailId, undefined);
  });

  it('skips the write when the Message-ID matches more than one message', async () => {
    const { client, calls } = spyClient({ findEmailIdsByMessageId: async () => ['a', 'b'] });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(calls.keywords, undefined);
    assert.equal(r.keywordMaintenance?.marked, false);
    assert.equal(r.keywordMaintenance?.skipReason, 'ambiguous');
  });

  it('skips the write when the lookup itself fails', async () => {
    const { client, calls } = spyClient({
      findEmailIdsByMessageId: async () => { throw new Error('boom'); },
    });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(calls.keywords, undefined);
    assert.equal(r.keywordMaintenance?.marked, false);
    assert.equal(r.keywordMaintenance?.skipReason, 'lookup-failed');
  });

  it('swallows a keyword-set failure best-effort (the draft already sent)', async () => {
    const { client } = spyClient({
      addKeywords: async () => { throw new InvalidInputError('nope'); },
    });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(r.submissionId, 'sub-1');           // send success is NOT masked
    assert.equal(r.keywordMaintenance?.marked, false);
    assert.equal(r.keywordMaintenance?.skipReason, undefined); // the original WAS identified
    assert.equal(r.keywordMaintenance?.originalEmailId, 'orig-1');
  });

  it('swallows a keyword-set failure best-effort — plain Error class', async () => {
    const { client } = spyClient({ addKeywords: async () => { throw new Error('boom'); } });
    const r = await sendDraftAndMaintainKeywords({ emailId: 'd1' }, client);
    assert.equal(r.submissionId, 'sub-1');
    assert.equal(r.keywordMaintenance?.marked, false);
  });

  it('does no keyword work when the send itself fails (an unreadable draft never sends)', async () => {
    const { client, calls } = spyClient({
      sendDraft: async () => { throw new InvalidInputError('Cannot send a non-draft email'); },
    });
    await assert.rejects(
      () => sendDraftAndMaintainKeywords({ emailId: 'd1' }, client),
      /Cannot send a non-draft email/,
    );
    assert.equal(calls.lookup, undefined);
    assert.equal(calls.keywords, undefined);
  });
});
