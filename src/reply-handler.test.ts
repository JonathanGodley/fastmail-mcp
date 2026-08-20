import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildReplyParams, composeReply } from './reply-handler.js';
import type { ReplyClient } from './reply-handler.js';
import { hasTextQuoteMarker } from './reply-quote.js';
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
    assert.deepEqual(buildReplyParams({ originalEmailId: 'e1', textBody: 'x', to: ['alice@x.example'] }, makeOriginal()).replyParams.to, ['alice@x.example']);
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

// A sending identity with a configured sign-off, for the appendSignature cases (#33).
const SIGNED_IDENTITY = {
  id: 'id-1',
  name: 'Test User',
  email: 'me@example.com',
  mayDelete: false,
  textSignature: 'Kind regards,\nTest User',
  htmlSignature: '<div>Kind regards,</div><div>Test User</div>',
};

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
      getIdentities: async () => { calls.identityLookups = (calls.identityLookups ?? 0) + 1; return [SIGNED_IDENTITY]; },
      uploadAttachments: async (specs, dir, allowBlob, options) => { calls.upload = { specs, dir, allowBlob, options }; return uploadResult; },
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
      ...over,
    };
    return { client, calls };
  }

  it('saves a draft and returns its id and subject', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, undefined, false);
    assert.deepEqual(r, { subject: 'Re: Project update', emailId: 'draft-9' });
    assert.equal(calls.getId, 'o1');
    assert.ok(calls.draft);
  });

  it('carries a subject override into the draft and the reported result (#68)', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', textBody: 'hi', subject: 'Budget sign-off' },
      client, undefined, false,
    );
    assert.equal(calls.draft.subject, 'Budget sign-off');
    assert.equal(r.subject, 'Budget sign-off');
  });

  it('uploads with the given attachDir and threads the parts into the draft', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply(
      { originalEmailId: 'o1', textBody: 'hi', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root', false,
    );
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'a.pdf' }], dir: '/attach/root', allowBlob: false, options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments, UPLOADED); // threaded into createDraft
  });

  // The other half of the flag. Every case above passes false, so a handler that dropped
  // the argument and hardcoded "off" would still satisfy them — and the only symptom would
  // be an in-account source refused on a server configured to allow it.
  it('passes the blob opt-in through when it is on', async () => {
    const { client, calls } = spyClient();
    await composeReply(
      { originalEmailId: 'o1', textBody: 'hi', attachments: [{ emailId: 'M1', attachmentId: 'p2' }] },
      client, undefined, true,
    );
    assert.deepEqual(calls.upload, {
      specs: [{ emailId: 'M1', attachmentId: 'p2' }], dir: undefined, allowBlob: true, options: { inlineCids: new Set() },
    });
  });

  // The send parameter was removed when the compose surface went draft-first (#32/#66):
  // the schema now rejects it as an unknown parameter before this orchestration runs.
  // Belt-and-braces: even if a stray send flag reached this layer, it is ignored and the
  // reply is still saved as a draft — this function has no way to transmit anything.
  it('ignores a stray send flag: the reply is drafted, never transmitted', async () => {
    const { client, calls } = spyClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi', send: true }, client, undefined, false);
    assert.equal(r.emailId, 'draft-9');
    assert.ok(calls.draft); // the only write is the draft create
  });

  it('does not call uploadAttachments when no attachments are given', async () => {
    let uploadCalled = false;
    const { client } = spyClient({ uploadAttachments: async () => { uploadCalled = true; return []; } });
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, '/attach/root', false);
    assert.equal(r.emailId, 'draft-9');
    assert.equal(uploadCalled, false);
  });

  it('requires originalEmailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeReply({ textBody: 'hi' }, client, undefined, false), /originalEmailId is required/);
  });

  it('rejects a malformed body before uploading or saving a draft', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeReply(
        { originalEmailId: 'o1', htmlBody: '<![CDATA[<p>hi</p>]]>', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root', false,
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
      await composeReply(NOTE_ARGS, client, '/attach/root', false);
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports what the saved draft embeds, after re-reading it', async () => {
      const { client, calls } = spyClient({
        getEmailById: async (id) => {
          if (id === 'o1') return makeOriginal();
          return { id, attachments: [{ blobId: 'blob-chart', cid: 'chart', size: 2048, disposition: 'inline' }] };
        },
      }, [CHART]);
      const r = await composeReply(NOTE_ARGS, client, '/attach/root', false);
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
          client, '/attach/root', false,
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
      await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, undefined, false);
      assert.deepEqual(calls.gets, ['o1']);
    });

    it('sees attachments supplied as a JSON string (lenient client)', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeReply(
        { ...NOTE_ARGS, attachments: '[{"path":"chart.png","cid":"chart"}]' },
        client, '/attach/root', false,
      );
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports a malformed body before an attachment problem', async () => {
      const { client } = spyClient();
      await assert.rejects(
        () => composeReply(
          { originalEmailId: 'o1', htmlBody: '<p>hi</p><img src="cid:missing"><![CDATA[x]]>' },
          client, '/attach/root', false,
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

// A fixed stand-in for a minted Content-ID, so the assertions read as values rather than as
// patterns. Real mints come from the CSPRNG; the shape is pinned in inline-images.test.ts.
const MINT = 'ii-00000000000000000000000000000001@inline.invalid';

// An original whose html displays an embedded image the message really carries.
const inlinePng = {
  partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'pic.png',
  disposition: 'inline', cid: 'img-1',
};
function withInlineImage(over: any = {}) {
  return makeOriginal({
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:img-1"></p>' } },
    attachments: [inlinePng],
    ...over,
  });
}

describe('buildReplyParams — images the quote carries', () => {
  it('carries a quoted image under a minted identifier and points the quote at it', () => {
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply', htmlBody: '<p>my reply</p>' },
      withInlineImage(), undefined, () => MINT,
    );
    assert.deepEqual(quoteImages.minted, [
      { blobId: 'blob-png', type: 'image/png', name: 'pic.png', cid: MINT, disposition: 'inline' },
    ]);
    assert.match(replyParams.htmlBody!, new RegExp(`src="cid:${MINT}"`));
    assert.equal(quoteImages.htmlQuoteShips, true);
    // The builder is pure: threading the parts onto the draft is the orchestration's job.
    assert.equal(replyParams.attachments, undefined);
  });

  it('carries nothing when the whole quote is dropped', () => {
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>', quoteOriginal: false },
      withInlineImage(), undefined, () => MINT,
    );
    assert.deepEqual(quoteImages.minted, []);
    assert.equal(replyParams.htmlBody, '<p>my reply</p>');
  });

  it('mints nothing for a text-only reply, and reports what the quote lost', () => {
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply' },
      withInlineImage(), undefined, () => MINT,
    );
    assert.deepEqual(quoteImages.minted, []);
    assert.equal(quoteImages.htmlQuoteShips, false);
    assert.deepEqual(quoteImages.resolvedParts, [inlinePng]);
    assert.match(replyParams.textBody!, /On .*wrote:\n> original text/);
  });

  it('quotes an image-only original, which has no text of its own to quote', () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
    });
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>' },
      original, undefined, () => MINT,
    );
    assert.match(replyParams.htmlBody!, new RegExp(`src="cid:${MINT}"`));
    assert.match(replyParams.htmlBody!, /wrote:/);
    assert.equal(quoteImages.minted.length, 1);
  });

  // The phantom-quote row: an image-only original whose sole reference resolves to nothing
  // has no content to quote at all, so the reply must not open an attribution over an empty
  // blockquote.
  it('skips the quote entirely when an image-only original references a part it does not carry', () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:gone"></div>' } },
      attachments: [],
    });
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>' },
      original, undefined, () => MINT,
    );
    assert.equal(replyParams.htmlBody, '<p>my reply</p>');
    assert.doesNotMatch(replyParams.htmlBody!, /wrote:/);
    assert.deepEqual(quoteImages.unresolvedRefs, ['gone']);
    assert.deepEqual(quoteImages.minted, []);
  });

  // The same original, but with an image it CAN carry and a reply that ships no html. The
  // quote has nothing to put in the plain-text alternative, so an attribution there would
  // describe a quote that is not present — and would arm this server's own text quote
  // marker, so the next edit of the draft would be challenged over a quote it never had.
  it('a text-only reply to an image-only original gets no attribution and no armed marker', () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
    });
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', textBody: 'my reply' },
      original, undefined, () => MINT,
    );
    assert.equal(replyParams.textBody, 'my reply');
    assert.equal(hasTextQuoteMarker(replyParams.textBody), false);
    assert.deepEqual(quoteImages.minted, []);
    assert.deepEqual(quoteImages.resolvedParts, [inlinePng]);
  });

  it('carries an SVG the original displayed, like any other embedded image', () => {
    const svg = { ...inlinePng, blobId: 'blob-svg', type: 'image/svg+xml', name: 'logo.svg' };
    const { quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>' },
      withInlineImage({ attachments: [svg] }), undefined, () => MINT,
    );
    assert.deepEqual(quoteImages.minted, [
      { blobId: 'blob-svg', type: 'image/svg+xml', name: 'logo.svg', cid: MINT, disposition: 'inline' },
    ]);
  });

  it('carries neither of two parts that share one Content-ID, and counts both as lost', () => {
    const twin = { ...inlinePng, partId: '5', blobId: 'blob-png-2', name: 'pic2.png' };
    const { replyParams, quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>' },
      withInlineImage({ attachments: [inlinePng, twin] }), undefined, () => MINT,
    );
    assert.deepEqual(quoteImages.minted, []);
    assert.equal(quoteImages.resolvedParts.length, 2);
    // Resolved, so never also counted as a reference that matched nothing.
    assert.deepEqual(quoteImages.unresolvedRefs, []);
    assert.doesNotMatch(replyParams.htmlBody!, /<img/);
  });

  it('counts a data:-URI image in the quoted body without carrying it', () => {
    const original = withInlineImage({
      bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="data:image/png;base64,AAA="></p>' } },
      attachments: [],
    });
    const { quoteImages } = buildReplyParams(
      { originalEmailId: 'e1', htmlBody: '<p>my reply</p>' },
      original, undefined, () => MINT,
    );
    assert.equal(quoteImages.droppedDataImages, 1);
    assert.deepEqual(quoteImages.minted, []);
  });
});

describe('composeReply — threading the quote images onto the draft', () => {
  // The confirmation read returns a draft echoing back exactly what the create call
  // attached, which is what a real saved draft holds — and the only way a test can assert
  // on the Content-IDs the quote mints, since production mints them at random.
  function imageClient(original: any = withInlineImage(), uploadResult: any[] = []) {
    const calls: any = {};
    const client: ReplyClient = {
      getEmailById: async (id) => {
        calls.gets = [...(calls.gets ?? []), id];
        if (id !== 'draft-9') return original;
        return { id, attachments: (calls.draft?.attachments ?? []).map((a: any) => ({ ...a, size: 1024 })) };
      },
      getIdentities: async () => [SIGNED_IDENTITY],
      uploadAttachments: async () => uploadResult,
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    };
    return { client, calls };
  }

  it("appends the quote images BEHIND the caller's own files rather than replacing them", async () => {
    const uploaded = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];
    const { client, calls } = imageClient(withInlineImage(), uploaded);
    await composeReply(
      { originalEmailId: 'o1', htmlBody: '<p>hi</p>', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['up-1', 'blob-png']);
    assert.equal(calls.draft.attachments[1].disposition, 'inline');
  });

  it('carries the quoted image even with no attachments directory configured', async () => {
    const { client, calls } = imageClient();
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-png']);
    assert.deepEqual(r.notes, ['This draft embeds 1 image(s) from the quoted message (1 KB).']);
  });

  it('re-reads the saved draft for a reply whose only images come from the quote', async () => {
    const { client, calls } = imageClient();
    await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(calls.gets, ['o1', 'draft-9']);
  });

  it('says so when an image the quote carries is not on the saved draft', async () => {
    const calls: any = {};
    const client: ReplyClient = {
      // The read-back finds a draft with none of the parts this call attached.
      getEmailById: async (id) => (id === 'draft-9' ? { id, attachments: [] } : withInlineImage()),
      getIdentities: async () => [SIGNED_IDENTITY],
      uploadAttachments: async () => [],
      createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    };
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the quoted message (1 KB).',
      '1 embedded image(s) this call attached were not found on the saved draft.' +
      ' Open the draft to check how it renders.',
    ]);
  });

  it('does not re-read the draft for a reply that carries no images at all', async () => {
    const { client, calls } = imageClient(makeOriginal());
    await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(calls.gets, ['o1']);
  });

  it('leaves attachments unset when the reply carries nothing', async () => {
    const { client, calls } = imageClient(makeOriginal());
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.equal('attachments' in calls.draft, false);
    assert.equal(r.notes, undefined);
  });

  it('says plainly that a text-only reply dropped the images the original displayed', async () => {
    const { client } = imageClient();
    const r = await composeReply({ originalEmailId: 'o1', textBody: 'hi' }, client, undefined, false);
    assert.deepEqual(r.notes, [
      '1 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('reports a shortfall when the quote could not carry everything it referenced', async () => {
    const twin = { ...inlinePng, partId: '5', blobId: 'blob-png-2', name: 'pic2.png' };
    const { client } = imageClient(withInlineImage({ attachments: [inlinePng, twin] }));
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(r.notes, [
      'Embedded 0 of 2 image part(s) referenced by the quote (0 KB embedded); ' +
      '2 could not be embedded and are not part of this draft.',
    ]);
  });

  // A quote is re-sent from a different message, so an image referenced by a path relative to
  // the original's own origin cannot come with it. The quote still ships; the loss is said.
  it('says how many images the quote dropped for a reference form it cannot carry', async () => {
    const orig = withInlineImage({
      bodyValues: {
        t: { value: 'original text' },
        h: { value: '<p><img src="cid:img-1"><img src="/logo.png"><img src="//cdn.example.com/a.png"></p>' },
      },
    });
    const { client } = imageClient(orig);
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the quoted message (1 KB).',
      '2 image(s) in the quoted message used a reference form this server cannot carry into' +
      ' a quote and were dropped; the rest of the quote was kept.',
    ]);
  });

  it('reports a reference in the quoted body that matched no part', async () => {
    const { client } = imageClient(withInlineImage({ attachments: [] }));
    const r = await composeReply({ originalEmailId: 'o1', htmlBody: '<p>hi</p>' }, client, undefined, false);
    assert.deepEqual(r.notes, ['1 reference(s) matched no part and were skipped.']);
  });
});
