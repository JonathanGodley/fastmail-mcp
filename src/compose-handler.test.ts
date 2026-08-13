import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { composeSend, composeDraft } from './compose-handler.js';
import type { ComposeClient } from './compose-handler.js';
import { InvalidInputError } from './coerce.js';

const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];

// A spy ComposeClient: records what each method was called with; uploadAttachments
// returns the canned UPLOADED parts so we can assert they thread through.
function spyClient(over: Partial<ComposeClient> = {}) {
  const calls: any = {};
  const client: ComposeClient = {
    uploadAttachments: async (specs, dir) => { calls.upload = { specs, dir }; return UPLOADED; },
    createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    sendEmail: async (p) => { calls.send = p; return 'sub-9'; },
    ...over,
  };
  return { client, calls };
}

describe('composeSend — required fields and threading', () => {
  it('sends and returns the submission id', async () => {
    const { client, calls } = spyClient();
    const id = await composeSend({ to: 'a@b.com', subject: 'Hi', textBody: 'hello' }, client, undefined);
    assert.equal(id, 'sub-9');
    assert.deepEqual(calls.send.to, ['a@b.com']); // recipient string coerced to an array
    assert.equal(calls.send.subject, 'Hi');
  });

  // Pin the WHOLE outgoing object, so a field dropped from the handoff fails the suite
  // rather than silently never reaching the server.
  it('passes every supported field through to sendEmail', async () => {
    const { client, calls } = spyClient();
    await composeSend({
      to: ['a@b.com'],
      cc: ['c@d.com'],
      bcc: ['e@f.com'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: [{ path: 'a.pdf' }],
    }, client, '/attach/root');
    assert.deepEqual(calls.send, {
      to: ['a@b.com'],
      cc: ['c@d.com'],
      bcc: ['e@f.com'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: UPLOADED,
    });
  });

  it('uploads attachments with the given attachDir and threads the parts through', async () => {
    const { client, calls } = spyClient();
    await composeSend(
      { to: ['a@b.com'], subject: 'Hi', textBody: 'hello', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.upload, { specs: [{ path: 'a.pdf' }], dir: '/attach/root' });
    assert.deepEqual(calls.send.attachments, UPLOADED);
  });

  it('does not upload when no attachments are given', async () => {
    const { client, calls } = spyClient();
    await composeSend({ to: ['a@b.com'], subject: 'Hi', textBody: 'hello' }, client, '/attach/root');
    assert.equal(calls.upload, undefined);
  });

  it('requires to, subject and a body', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeSend({ subject: 'Hi', textBody: 'x' }, client, undefined), /to field is required/);
    await assert.rejects(() => composeSend({ to: ['a@b.com'], textBody: 'x' }, client, undefined), /subject is required/);
    await assert.rejects(() => composeSend({ to: ['a@b.com'], subject: 'Hi' }, client, undefined), /Either textBody or htmlBody is required/);
  });
});

describe('composeSend — malformed bodies are rejected before anything is transmitted', () => {
  it('rejects a non-string body (#62)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeSend({ to: ['a@b.com'], subject: 'Hi', htmlBody: 42 }, client, undefined),
      (e: any) => e instanceof InvalidInputError && /htmlBody must be a string/.test(e.message),
    );
    assert.equal(calls.send, undefined);
  });

  it('rejects an entirely HTML-escaped htmlBody (#71/#77)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeSend({ to: ['a@b.com'], subject: 'Hi', htmlBody: '&lt;p&gt;Hi&lt;/p&gt;' }, client, undefined),
      /htmlBody appears to be HTML-escaped/,
    );
    assert.equal(calls.send, undefined);
  });

  it('rejects a CDATA-wrapped body (#78)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeSend({ to: ['a@b.com'], subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>' }, client, undefined),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => composeSend({ to: ['a@b.com'], subject: 'Hi', textBody: '<![CDATA[Hi]]>' }, client, undefined),
      /textBody is wrapped in a CDATA section/,
    );
    assert.equal(calls.send, undefined);
  });

  it('rejects before uploading attachments (nothing hits the blob store)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeSend({ to: ['a@b.com'], subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>', attachments: [{ path: 'a.pdf' }] }, client, '/attach/root'),
      /CDATA/,
    );
    assert.equal(calls.upload, undefined);
  });
});

describe('composeDraft — creation and summary fields', () => {
  it('creates the draft and returns the fields the handler summarises', async () => {
    const { client, calls } = spyClient();
    const r = await composeDraft({ to: ['a@b.com'], cc: ['c@d.com'], subject: 'Hi', textBody: 'hello' }, client, undefined);
    assert.deepEqual(r, { emailId: 'draft-9', subject: 'Hi', to: ['a@b.com'], cc: ['c@d.com'] });
    assert.equal(calls.draft.subject, 'Hi');
  });

  it('passes every supported field through to createDraft', async () => {
    const { client, calls } = spyClient();
    await composeDraft({
      to: ['a@b.com'],
      cc: ['c@d.com'],
      bcc: ['e@f.com'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: [{ path: 'a.pdf' }],
    }, client, '/attach/root');
    assert.deepEqual(calls.draft, {
      to: ['a@b.com'],
      cc: ['c@d.com'],
      bcc: ['e@f.com'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: UPLOADED,
    });
  });

  it('allows an attachment-only draft (attachments count as content)', async () => {
    const { client, calls } = spyClient();
    const r = await composeDraft({ attachments: [{ path: 'a.pdf' }] }, client, '/attach/root');
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.draft.attachments, UPLOADED);
  });

  it('rejects a draft with nothing in it at all', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeDraft({}, client, undefined), /At least one of to, subject, textBody, htmlBody, or attachments/);
  });
});

describe('composeDraft — malformed bodies are rejected before the draft is created', () => {
  it('rejects a non-string body (#62)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', textBody: ['a', 'b'] }, client, undefined),
      (e: any) => e instanceof InvalidInputError && /textBody must be a string; received array/.test(e.message),
    );
    assert.equal(calls.draft, undefined);
  });

  it('rejects an entirely HTML-escaped htmlBody (#71/#77)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', htmlBody: '&lt;p&gt;Hi&lt;/p&gt;' }, client, undefined),
      /htmlBody appears to be HTML-escaped/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('rejects a CDATA-wrapped body (#78)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>' }, client, undefined),
      /htmlBody contains a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('still accepts legitimate bodies that look superficially similar', async () => {
    const { client, calls } = spyClient();
    await composeDraft({ subject: 'Hi', htmlBody: '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>', textBody: 'Close it with ]]> here' }, client, undefined);
    assert.equal(calls.draft.htmlBody, '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>');
  });
});
