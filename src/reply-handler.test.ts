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
    assert.equal(replyParams.textBody, undefined); // caller gave no text; createDraft adds the fallback later
    assert.match(replyParams.htmlBody!, /html reply.*<blockquote/s);
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
  it('allows a body-less reply draft — does not throw (fill it in via edit_draft before send_draft)', () => {
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

  it('rejects a malformed body even though nothing is ever transmitted from this tool', () => {
    assert.throws(
      () => buildReplyParams({ originalEmailId: 'e1', htmlBody: '<![CDATA[<p>x</p>]]>' }, makeOriginal()),
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

describe('composeReply — draft-only orchestration', () => {
  const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];

  // A spy ReplyClient: records what each method was called with; uploadAttachments
  // returns the canned UPLOADED parts so we can assert they thread through. The
  // interface has no transmit method at all — send_draft is the only sender.
  function spyClient(over: Partial<ReplyClient> = {}, uploadResult: any[] = UPLOADED) {
    const calls: any = { gets: [] as string[] };
    const client: ReplyClient = {
      // Serves two jobs: fetching the original, and the post-save confirmation read. The
      // recorded ids are what tell the two apart.
      getEmailById: async (id) => { calls.getId = id; calls.gets.push(id); return makeOriginal(); },
      uploadAttachments: async (specs, dir, options) => { calls.upload = { specs, dir, options }; return uploadResult; },
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
      ...over,
    };
    return { client, calls };
  }

  it('saves a draft and returns its id and subject', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, undefined);
    assert.deepEqual(r, { subject: 'Re: Project update', emailId: 'draft-9' });
    assert.equal(calls.getId, 'o1');
    assert.ok(calls.draft);
  });

  it('carries a subject override into the draft and the reported result (#68)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', textBody: 'hi', subject: 'Budget sign-off' },
      client, undefined,
    );
    assert.equal(calls.draft.subject, 'Budget sign-off');
    assert.equal(r.subject, 'Budget sign-off');
  });

  it('uploads with the given attachDir and threads the parts into the draft', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', textBody: 'hi', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'a.pdf' }], dir: '/attach/root', options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments, UPLOADED); // threaded into createDraft
  });

  // The send parameter was removed when the compose surface went draft-first (#32/#66):
  // the schema now rejects it as an unknown parameter before this orchestration runs.
  // Belt-and-braces: even if a stray send flag reached this layer, it is ignored and the
  // reply is still saved as a draft — this function has no way to transmit anything.
  it('ignores a stray send flag: the reply is drafted, never transmitted', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi', send: true }, client, undefined);
    assert.equal(r.emailId, 'draft-9');
    assert.ok(calls.draft); // the only write is the draft create
  });

  it('does not call uploadAttachments when no attachments are given', async () => {
    let uploadCalled = false;
    const { client } = spyClient({ uploadAttachments: async () => { uploadCalled = true; return []; } });
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, '/attach/root');
    assert.equal(r.emailId, 'draft-9');
    assert.equal(uploadCalled, false);
  });

  it('requires originalEmailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeReply({ textBody: 'hi' }, client, undefined), /originalEmailId is required/);
  });

  it('rejects a malformed body before uploading or saving a draft', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeReply(
        { originalEmailId: 'o1', htmlBody: '<![CDATA[<p>hi</p>]]>', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
    assert.equal(calls.upload, undefined);
  });

  // Embedding an image in the note the caller writes above the quote (#13).
  describe('embedded images in the caller\'s note', () => {
    const NOTE_ARGS = {
      originalEmailId: 'o1',
      htmlBody: '<p>Compare with:</p><img src="cid:chart">',
      attachments: [{ path: 'chart.png', cid: 'chart' }],
    };
    const CHART = { blobId: 'blob-chart', type: 'image/png', name: 'chart.png', disposition: 'inline', cid: 'chart' };

    it('tells the upload which Content-ID the note displays', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeReply(NOTE_ARGS, client, '/attach/root');
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports what the saved draft embeds, after re-reading it', async () => {
      const { client, calls } = spyClient({
        getEmailById: async (id) => {
          if (id === 'o1') return makeOriginal();
          return { id, attachments: [{ blobId: 'blob-chart', cid: 'chart', size: 2048, disposition: 'inline' }] };
        },
      }, [CHART]);
      const r = await composeReply(NOTE_ARGS, client, '/attach/root');
      assert.deepEqual(r.notes, ['This draft embeds 1 image(s) (2 KB).']);
      assert.equal(calls.draft.attachments.length, 1);
    });

    // The reply wording differs from a fresh compose on purpose: a quote sits below the
    // note, and its images arrive on their own rather than being authored.
    it('rejects a note reference nothing supplies, in the reply wording', async () => {
      const { client, calls } = spyClient();
      await assert.rejects(
        () => composeReply(
          { originalEmailId: 'o1', htmlBody: '<img src="cid:chart">' },
          client, '/attach/root',
        ),
        (e: any) => /your note/.test(e.message) && /Quoted images appear inside the quote automatically/.test(e.message),
      );
      assert.equal(calls.upload, undefined);
      assert.equal(calls.draft, undefined);
    });

    // The checks read the caller's own note, never the merged body: the quote this server
    // builds carries its own identifiers and must not be held against the caller.
    it('does not re-read the draft for an ordinary reply', async () => {
      const { client, calls } = spyClient();
      await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, undefined);
      assert.deepEqual(calls.gets, ['o1']);
    });

    it('sees attachments supplied as a JSON string (lenient client)', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeReply(
        { ...NOTE_ARGS, attachments: '[{"path":"chart.png","cid":"chart"}]' },
        client, '/attach/root',
      );
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports a malformed body before an attachment problem', async () => {
      const { client } = spyClient();
      await assert.rejects(
        () => composeReply(
          { originalEmailId: 'o1', htmlBody: '<p>hi</p><img src="cid:missing"><![CDATA[x]]>' },
          client, '/attach/root',
        ),
        /htmlBody contains a CDATA section/,
      );
    });
  });
});

describe('buildReplyParams — recorded source instance', () => {
  it("records the original's JMAP id as sourceEmailId (the exact copy send_draft marks)", () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply' },
      makeOriginal({ id: 'orig-1' }),
    );
    assert.equal(replyParams.sourceEmailId, 'orig-1');
  });

  it('omits sourceEmailId when the fetched original carries no id', () => {
    const { replyParams } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply' },
      makeOriginal(),
    );
    assert.equal(replyParams.sourceEmailId, undefined);
  });
});
