import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyParams, composeReply } from './reply-handler.js';
import type { ReplyClient } from './reply-handler.js';
import { InvalidInputError } from './coerce.js';

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
  it('uses a caller-supplied subject verbatim, with no Re: prefixing (#68)', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'x', subject: 'Budget sign-off' },
      makeOriginal(),
    );
    assert.equal(replyParams.subject, 'Budget sign-off');
    // The override changes only the subject: the threading chain is built the same way.
    assert.deepEqual(replyParams.inReplyTo, ['orig-msg@example.com']);
    assert.deepEqual(replyParams.references, ['root@example.com', 'orig-msg@example.com']);
  });
  it('does not add Re: to an override that lacks it, even when the original had one (#68)', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'x', subject: 'New topic' },
      makeOriginal({ subject: 'Re: Project update' }),
    );
    assert.equal(replyParams.subject, 'New topic');
  });
  it('treats a supplied-but-blank subject as omitted (falls to the Re: default) (#68)', () => {
    // '​' is a zero-width space: visually empty, so it reads as blank like the rest.
    for (const blank of ['', '   ', '​']) {
      const { replyParams } = buildReplyParams(
        { originalEmailId: 'e1', textBody: 'x', subject: blank },
        makeOriginal(),
      );
      assert.equal(replyParams.subject, 'Re: Project update');
    }
  });
  it('treats a null subject as omitted, as lenient clients spell an unset field (#68)', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'x', subject: null },
      makeOriginal(),
    );
    assert.equal(replyParams.subject, 'Re: Project update');
  });
  it('rejects a non-string subject rather than silently inheriting the default (#68)', () => {
    for (const bad of [42, ['a'], { s: 1 }, true]) {
      assert.throws(
        () => buildReplyParams({ originalEmailId: 'e1', textBody: 'x', subject: bad }, makeOriginal()),
        /subject must be a string/,
      );
    }
  });
  it('defaults the recipient to the original sender, preserving the display name (#31)', () => {
    assert.deepEqual(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, makeOriginal()).replyParams.to, ['Jon Godley <jon@example.com>']);
  });
  it('defaults to a bare address when the original sender has no display name (#31)', () => {
    const orig = makeOriginal({ from: [{ email: 'noname@example.com' }] });
    assert.deepEqual(buildReplyParams({ originalEmailId: 'e1', textBody: 'x' }, orig).replyParams.to, ['noname@example.com']);
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

describe('buildReplyParams — malformed caller bodies (#62, #71/#77, #78)', () => {
  it('rejects a non-string body instead of crashing on it', () => {
    assert.throws(
      () => buildReplyParams({ originalEmailId: 'e1', textBody: 42 }, makeOriginal()),
      (e: any) => e instanceof InvalidInputError && /textBody must be a string/.test(e.message),
    );
  });

  it('rejects an entirely HTML-escaped htmlBody', () => {
    assert.throws(
      () => buildReplyParams({ originalEmailId: 'e1', htmlBody: '&lt;p&gt;Hi there&lt;/p&gt;' }, makeOriginal()),
      /htmlBody appears to be HTML-escaped/,
    );
  });

  // The quote is what makes this class of malformed body dangerous on a reply: it hides
  // the damage from every downstream check. A CDATA-wrapped htmlBody derives to an empty
  // text part on its own (which the no-readable-body reject would catch), but once the
  // quoted original is appended the merged HTML has plenty of visible content, so the
  // reply ships with the new message dropped from the plain-text alternative and only the
  // original's words quoted back at the recipient. The guard therefore runs on the
  // caller's body, before the quote is merged in.
  it('rejects a CDATA-wrapped body even though the quote would supply visible content', () => {
    assert.throws(
      () => buildReplyParams(
        { originalEmailId: 'e1', htmlBody: '<![CDATA[<p>Just checking in</p>]]>', quoteOriginal: true },
        makeOriginal(),
      ),
      /htmlBody contains a CDATA section/,
    );
    assert.throws(
      () => buildReplyParams(
        { originalEmailId: 'e1', textBody: '<![CDATA[Just checking in]]>', quoteOriginal: true },
        makeOriginal(),
      ),
      /textBody is wrapped in a CDATA section/,
    );
  });

  it('rejects on the draft path too (send defaults to false, so nothing is transmitted either way)', () => {
    assert.throws(
      () => buildReplyParams({ originalEmailId: 'e1', send: false, htmlBody: '<![CDATA[<p>x</p>]]>' }, makeOriginal()),
      /CDATA/,
    );
  });

  it('leaves a legitimate body alone: escaped markup inside real tags, and a bare ]]>', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>Write it as <code>&lt;p&gt;</code>, and end with ]]&gt;</p>' },
      makeOriginal(),
    );
    assert.match(replyParams.htmlBody!, /&lt;p&gt;/);
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
      addKeywords: async (id, kw) => { calls.keywords = { id, kw }; },
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
    // #52/#54: the send branch marks the original answered + read and reports it.
    assert.deepEqual(calls.keywords, { id: 'o1', kw: ['$answered', '$seen'] });
    assert.equal(r.markedAnswered, true);
  });

  it('carries a subject override into the SEND branch and the reported result (#68)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', send: true, textBody: 'hi', subject: 'Budget sign-off' },
      client, undefined,
    );
    assert.equal(calls.send.subject, 'Budget sign-off');
    assert.equal(r.subject, 'Budget sign-off');
  });

  it('carries a subject override into the DRAFT branch and the reported result (#68)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', send: false, textBody: 'hi', subject: 'Budget sign-off' },
      client, undefined,
    );
    assert.equal(calls.draft.subject, 'Budget sign-off');
    assert.equal(r.subject, 'Budget sign-off');
  });

  it('inherits the Re: subject in both branches when no override is given (#68)', async () => {
    const draft = spyClient();
    await composeReply({ originalEmailId: 'o1', send: false, textBody: 'hi' }, draft.client, undefined);
    assert.equal(draft.calls.draft.subject, 'Re: Project update');
    const sent = spyClient();
    await composeReply({ originalEmailId: 'o1', send: true, textBody: 'hi' }, sent.client, undefined);
    assert.equal(sent.calls.send.subject, 'Re: Project update');
  });

  it('marks the original answered+read on a bare send (no attachments)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', send: true, textBody: 'hi' }, client, undefined);
    assert.equal(r.sent, true);
    assert.deepEqual(calls.keywords, { id: 'o1', kw: ['$answered', '$seen'] });
    assert.equal(r.markedAnswered, true);
  });

  it('does NOT mark keywords on the DRAFT branch (a draft answered nothing)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', send: false, textBody: 'hi' }, client, undefined);
    assert.equal(r.sent, false);
    assert.equal(calls.keywords, undefined);
    assert.equal(r.markedAnswered, undefined);
  });

  it('swallows a keyword-set failure best-effort (reply already sent) — notFound class', async () => {
    const { client } = spyClient({ addKeywords: async () => { throw new InvalidInputError('nope'); } });
    const r = await composeReply({ originalEmailId: 'o1', send: true, textBody: 'hi' }, client, undefined);
    assert.equal(r.sent, true);          // send success is NOT masked by the keyword failure
    assert.equal(r.submissionId, 'sub-9');
    assert.equal(r.markedAnswered, false);
  });

  it('swallows a keyword-set failure best-effort — plain Error class', async () => {
    const { client } = spyClient({ addKeywords: async () => { throw new Error('boom'); } });
    const r = await composeReply({ originalEmailId: 'o1', send: true, textBody: 'hi' }, client, undefined);
    assert.equal(r.sent, true);
    assert.equal(r.markedAnswered, false);
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

  it('defaults to the DRAFT branch when send is omitted (orchestrator level)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, '/attach/root');
    assert.equal(r.sent, false);
    assert.equal(r.emailId, 'draft-9');
    assert.ok(calls.draft);               // draft branch taken without an explicit send:false
    assert.equal(calls.send, undefined);  // nothing transmitted
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

  it('rejects a malformed body before sending, uploading, or saving a draft', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeReply(
        { originalEmailId: 'o1', send: true, htmlBody: '<![CDATA[<p>hi</p>]]>', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );
    assert.equal(calls.send, undefined);
    assert.equal(calls.draft, undefined);
    assert.equal(calls.upload, undefined);
  });
});
