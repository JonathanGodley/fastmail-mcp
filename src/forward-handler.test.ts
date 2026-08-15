import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildForwardParams, composeForward, sanitizeEmlFilename } from './forward-handler.js';
import type { ForwardClient } from './forward-handler.js';
import { hasTextForwardMarker, hasForwardMarker } from './reply-quote.js';
import { EMAIL_BODY_PROPERTIES } from './jmap-client.js';
import { InvalidInputError } from './coerce.js';

// A raw-JMAP-shaped original (as getEmailById returns it AFTER the
// EMAIL_BODY_PROPERTIES disposition/cid addition — the attachment part shape here
// deliberately mirrors a live Fastmail response, so an under-shaped mock can't
// mask a missing property).
function makeOriginal(over: any = {}) {
  return {
    id: 'orig-1',
    blobId: 'blob-orig-raw',
    messageId: ['orig-msg@example.com'],
    subject: 'Project update',
    from: [{ name: 'Ada Lovelace', email: 'ada@example.com' }],
    to: [{ name: 'Bob', email: 'bob@example.com' }],
    sentAt: '2026-07-01T09:14:00-04:00',
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>original html</p>' } },
    attachments: [
      { partId: '3', blobId: 'blob-doc', type: 'application/pdf', size: 1234, name: 'doc.pdf', disposition: 'attachment', cid: null },
    ],
    ...over,
  };
}

// The drop predicate needs disposition+cid on every fetched part — pin the property
// list so a later "cleanup" of EMAIL_BODY_PROPERTIES can't silently disarm it.
describe('EMAIL_BODY_PROPERTIES — forward prerequisites', () => {
  it('fetches disposition and cid (the inline-drop predicate inputs)', () => {
    assert.ok((EMAIL_BODY_PROPERTIES as readonly string[]).includes('disposition'));
    assert.ok((EMAIL_BODY_PROPERTIES as readonly string[]).includes('cid'));
  });
});

describe('buildForwardParams — subject', () => {
  it("defaults to 'Fwd: <original subject>'", () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.equal(forwardParams.subject, 'Fwd: Project update');
  });
  it('does not double-prefix an existing Fwd:/fw:/FWD:', () => {
    for (const s of ['Fwd: Hello', 'fw: Hello', 'FWD: Hello', 'FW: Hello']) {
      const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ subject: s }));
      assert.equal(forwardParams.subject, s);
    }
  });
  it('uses a caller-supplied subject verbatim', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], subject: 'Custom line' }, makeOriginal());
    assert.equal(forwardParams.subject, 'Custom line');
  });
  it('treats a supplied-but-blank subject as omitted (falls to the default rule)', () => {
    // Zero-width-aware, like every other emptiness test: '​' renders as empty.
    for (const blank of ['', '   ', '​']) {
      const { forwardParams } = buildForwardParams({ to: ['x@y.example'], subject: blank }, makeOriginal());
      assert.equal(forwardParams.subject, 'Fwd: Project update');
    }
  });
  it('treats a null subject as omitted, as lenient clients spell an unset field', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], subject: null }, makeOriginal());
    assert.equal(forwardParams.subject, 'Fwd: Project update');
  });
  it('rejects a non-string subject rather than silently using the default (#68)', () => {
    for (const bad of [42, ['a'], { s: 1 }, true]) {
      assert.throws(
        () => buildForwardParams({ to: ['x@y.example'], subject: bad }, makeOriginal()),
        /subject must be a string/,
      );
    }
  });
  it('handles an original with no subject', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ subject: undefined }));
    assert.equal(forwardParams.subject, 'Fwd: ');
  });
});

describe('buildForwardParams — recipients', () => {
  it('requires to (no default recipient, unlike reply)', () => {
    assert.throws(() => buildForwardParams({}, makeOriginal()), /to is required for a forward; there is no default recipient/);
    assert.throws(() => buildForwardParams({ to: [] }, makeOriginal()), /to is required for a forward/);
  });
  it('accepts a comma-separated to string (lenient coercion)', () => {
    const { forwardParams } = buildForwardParams({ to: 'a@x.example, b@y.example' }, makeOriginal());
    assert.deepEqual(forwardParams.to, ['a@x.example', 'b@y.example']);
  });
});

describe('buildForwardParams — malformed caller notes (#62, #71/#77, #78)', () => {
  it('rejects a non-string note instead of crashing on it', () => {
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], htmlBody: 42 }, makeOriginal()),
      (e: any) => e instanceof InvalidInputError && /htmlBody must be a string/.test(e.message),
    );
  });

  it('rejects an entirely HTML-escaped note', () => {
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], htmlBody: '&lt;p&gt;FYI&lt;/p&gt;' }, makeOriginal()),
      /htmlBody appears to be HTML-escaped/,
    );
  });

  // Same shadowing as the reply quote: the forwarded-message block supplies the visible
  // content the no-readable-body reject looks for, so a CDATA-wrapped note would ship
  // with the note itself missing from the plain-text part.
  it('rejects a CDATA-wrapped note even though the forwarded block would supply content', () => {
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], htmlBody: '<![CDATA[<p>FYI</p>]]>' }, makeOriginal()),
      /htmlBody contains a CDATA section/,
    );
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], textBody: '<![CDATA[FYI]]>' }, makeOriginal()),
      /textBody is wrapped in a CDATA section/,
    );
  });

  it('rejects on the asAttachment path too (the note is the whole body there)', () => {
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], asAttachment: true, htmlBody: '<![CDATA[<p>FYI</p>]]>' }, makeOriginal()),
      /CDATA/,
    );
  });

  it('leaves a legitimate note alone: escaped markup inside real tags', () => {
    const { forwardParams } = buildForwardParams(
      { to: ['x@y.example'], htmlBody: '<p>See the <code>&lt;config&gt;</code> block below</p>' },
      makeOriginal(),
    );
    assert.match(forwardParams.htmlBody!, /&lt;config&gt;/);
  });
});

describe('buildForwardParams — forwardedMessageId (recorded source)', () => {
  it("records the original's Message-ID", () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.deepEqual(forwardParams.forwardedMessageId, ['orig-msg@example.com']);
  });
  it('omits it when the original has no Message-ID', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ messageId: undefined }));
    assert.equal(forwardParams.forwardedMessageId, undefined);
  });
  it('omits a malformed Message-ID rather than fail the create (Fastmail rejects/mangles these on SET)', () => {
    for (const bad of [
      'has space@example.com',
      'angle<bracket@example.com',
      'angle>bracket@example.com',
      'non-ascii-käse@example.com',
      'ctrlbell@example.com',
      'x'.repeat(999) + '@example.com',
      '',
    ]) {
      const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ messageId: [bad] }));
      assert.equal(forwardParams.forwardedMessageId, undefined, `should omit: ${JSON.stringify(bad.slice(0, 40))}`);
    }
  });
  it('records it on asAttachment forwards too (send_draft resolves it to mark the original)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, makeOriginal());
    assert.deepEqual(forwardParams.forwardedMessageId, ['orig-msg@example.com']);
  });
});

describe('buildForwardParams — bodies', () => {
  it('caller-neither with an html original emits an html forward (no fabricated text side)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.equal(forwardParams.textBody, undefined);
    assert.match(forwardParams.htmlBody!, /----- Original message -----/);
    assert.match(forwardParams.htmlBody!, /original html/);
  });
  it('caller-neither with a TEXT-ONLY original emits a text forward (never fabricates html)', () => {
    const orig = makeOriginal({ htmlBody: undefined, bodyValues: { t: { value: 'plain only' } } });
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, orig);
    assert.equal(forwardParams.htmlBody, undefined);
    assert.match(forwardParams.textBody!, /----- Original message -----/);
    assert.match(forwardParams.textBody!, /plain only/);
  });
  it('places the caller note above the block', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], textBody: 'FYI folks' }, makeOriginal());
    assert.match(forwardParams.textBody!, /^FYI folks\n\n\n----- Original message -----/);
  });
});

describe('buildForwardParams — asAttachment (.eml)', () => {
  it("attaches the original's own blobId as message/rfc822 with a subject-derived name", () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, makeOriginal());
    assert.deepEqual(forwardParams.attachments, [{
      blobId: 'blob-orig-raw',
      type: 'message/rfc822',
      name: 'Project update.eml',
      disposition: 'attachment',
    }]);
  });
  it('note-less: emits the non-marker filler (gives the draft a readable body, arms no guard)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, makeOriginal());
    assert.equal(forwardParams.textBody, 'Forwarded message attached.');
    assert.equal(forwardParams.htmlBody, undefined);
    assert.equal(hasTextForwardMarker(forwardParams.textBody), false);
  });
  it('with a note: passes the caller bodies through with NO forwarded-message block', () => {
    const { forwardParams } = buildForwardParams(
      { to: ['x@y.example'], asAttachment: true, textBody: 'see attached', htmlBody: '<p>see attached</p>' },
      makeOriginal(),
    );
    assert.equal(forwardParams.textBody, 'see attached');
    assert.equal(forwardParams.htmlBody, '<p>see attached</p>');
    assert.equal(hasForwardMarker(forwardParams.htmlBody), false);
    assert.equal(hasTextForwardMarker(forwardParams.textBody), false);
  });
  it('throws when the original has no blobId', () => {
    assert.throws(
      () => buildForwardParams({ to: ['x@y.example'], asAttachment: true }, makeOriginal({ blobId: undefined })),
      /no blobId/,
    );
  });
});

describe('sanitizeEmlFilename', () => {
  it('keeps an ordinary subject', () => {
    assert.equal(sanitizeEmlFilename('Quarterly report'), 'Quarterly report.eml');
  });
  it('falls back on a blank/absent subject', () => {
    assert.equal(sanitizeEmlFilename(''), 'forwarded-message.eml');
    assert.equal(sanitizeEmlFilename('   '), 'forwarded-message.eml');
    assert.equal(sanitizeEmlFilename(undefined), 'forwarded-message.eml');
  });
  it('neutralizes traversal-shaped subjects (path separators, leading dots, colon)', () => {
    assert.equal(sanitizeEmlFilename('..\\..\\Startup\\x'), '_.._Startup_x.eml');
    assert.equal(sanitizeEmlFilename('../x'), '_x.eml');
    assert.equal(sanitizeEmlFilename('C:autorun'), 'C_autorun.eml');
  });
  it('strips control and format/bidi characters', () => {
    // U+202E (RLO) and U+0085 (NEL) are built from char codes so no raw invisible
    // character lives in this source file.
    const rlo = String.fromCharCode(0x202e);
    const nel = String.fromCharCode(0x85);
    assert.equal(sanitizeEmlFilename(`abc${rlo}lme.gpj`), 'abclme.gpj.eml');
    assert.equal(sanitizeEmlFilename(`ab${nel}cd`), 'abcd.eml');
  });
  it('caps the length before appending .eml', () => {
    const name = sanitizeEmlFilename('x'.repeat(500));
    assert.ok(name.length <= 84);
    assert.ok(name.endsWith('.eml'));
  });
  it('caps by code points — never splits a surrogate pair at the boundary', () => {
    // 79 ASCII chars then an astral char: a UTF-16-unit slice(0,80) would cut the
    // pair in half and leave a lone surrogate (invalid on the wire).
    const name = sanitizeEmlFilename('x'.repeat(79) + '\u{1F600}' + 'tail');
    assert.equal(name, 'x'.repeat(79) + '\u{1F600}.eml');
    assert.ok(!/[\uD800-\uDFFF]$/.test(name.slice(0, -4)) || /[\uDC00-\uDFFF]/.test(name)); // no lone surrogate
    assert.ok([...name].every((c) => !(c.length === 1 && c.charCodeAt(0) >= 0xd800 && c.charCodeAt(0) <= 0xdfff)));
  });
});

describe('buildForwardParams — attachment carry', () => {
  const inlinePng = { partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'pic.png', disposition: 'inline', cid: 'img-1' };
  const inlineNoCid = { partId: '5', blobId: 'blob-odd', type: 'application/zip', size: 10, name: 'odd.zip', disposition: 'inline', cid: null };

  it('carries non-inline originals through the whitelist (never server-set partId/size)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
    ]);
  });
  it('drops true-inline (cid) parts and counts them in droppedInlineImages', () => {
    const orig = makeOriginal({ attachments: [makeOriginal().attachments[0], inlinePng] });
    const { forwardParams, droppedInlineImages } = buildForwardParams({ to: ['x@y.example'] }, orig);
    assert.equal(droppedInlineImages, 1);
    assert.deepEqual(forwardParams.attachments!.map((a) => a.blobId), ['blob-doc']);
  });
  it("normalizes inline-WITHOUT-cid to disposition 'attachment' (kept, not counted; keeps the draft editable)", () => {
    const orig = makeOriginal({ attachments: [inlineNoCid] });
    const { forwardParams, droppedInlineImages } = buildForwardParams({ to: ['x@y.example'] }, orig);
    assert.equal(droppedInlineImages, 0);
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-odd', type: 'application/zip', name: 'odd.zip', disposition: 'attachment' },
    ]);
  });
  it('includeOriginalAttachments:false drops the carry but still counts inline drops (the body strips their <img> either way)', () => {
    const orig = makeOriginal({ attachments: [makeOriginal().attachments[0], inlinePng] });
    const { forwardParams, droppedInlineImages } = buildForwardParams(
      { to: ['x@y.example'], includeOriginalAttachments: false }, orig,
    );
    assert.equal(forwardParams.attachments, undefined);
    assert.equal(droppedInlineImages, 1);
  });
  it('asAttachment: droppedInlineImages is ABSENT even for an inline-bearing original (the .eml embeds them losslessly)', () => {
    const orig = makeOriginal({ attachments: [inlinePng] });
    const { droppedInlineImages } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, orig);
    assert.equal(droppedInlineImages, undefined);
  });
});

describe('composeForward — draft-only orchestration', () => {
  const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'new.pdf', disposition: 'attachment' }];

  // The interface has no transmit method at all — send_draft is the only sender.
  function spyClient(over: Partial<ForwardClient> = {}, uploadResult: any[] = UPLOADED) {
    const calls: any = { gets: [] as string[] };
    const client: ForwardClient = {
      // Serves two jobs: fetching the original, and the post-save confirmation read. The
      // recorded ids are what tell the two apart.
      getEmailById: async (id) => { calls.getId = id; calls.gets.push(id); return makeOriginal(); },
      uploadAttachments: async (specs, dir, options) => { calls.upload = { specs, dir, options }; return uploadResult; },
      createDraft: async (p) => { calls.draft = p; return 'draft-7'; },
      ...over,
    };
    return { client, calls };
  }

  it('requires originalEmailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeForward({ to: ['x@y.example'] }, client, undefined), /originalEmailId is required/);
  });
  it('saves a draft and returns its id and subject', async () => {
    const { client, calls } = spyClient();
    const r = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined);
    assert.equal(r.emailId, 'draft-7');
    assert.equal(r.subject, 'Fwd: Project update');
    assert.ok(calls.draft);
  });
  // The send parameter was removed when the compose surface went draft-first (#32/#66):
  // the schema now rejects it as an unknown parameter before this orchestration runs.
  // Belt-and-braces: even if a stray send flag reached this layer, it is ignored and the
  // forward is still saved as a draft — this function has no way to transmit anything.
  it('ignores a stray send flag: the forward is drafted, never transmitted', async () => {
    const { client, calls } = spyClient();
    const r = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'], send: true }, client, undefined);
    assert.equal(r.emailId, 'draft-7');
    assert.ok(calls.draft); // the only write is the draft create
  });
  it('appends caller uploads BEHIND the carried originals', async () => {
    const { client, calls } = spyClient();
    await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], attachments: [{ path: 'new.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'new.pdf' }], dir: '/attach/root', options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-doc', 'up-1']);
  });
  it('appends caller uploads behind the .eml on asAttachment', async () => {
    const { client, calls } = spyClient();
    await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true, attachments: [{ path: 'new.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-orig-raw', 'up-1']);
    assert.equal(calls.draft.attachments[0].type, 'message/rfc822');
  });
  it('surfaces droppedInlineImages in the result', async () => {
    const inline = { partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'p.png', disposition: 'inline', cid: 'img-1' };
    const { client } = spyClient({ getEmailById: async () => makeOriginal({ attachments: [inline] }) });
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined);
    assert.equal(d.droppedInlineImages, 1);
  });
  it('does not call uploadAttachments when no new attachments are given', async () => {
    let uploadCalled = false;
    const { client } = spyClient({ uploadAttachments: async () => { uploadCalled = true; return []; } });
    await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined);
    assert.equal(uploadCalled, false);
  });
  it('rejects a malformed note before uploading or saving a draft', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeForward(
        { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<![CDATA[<p>FYI</p>]]>', attachments: [{ path: 'new.pdf' }] },
        client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
    assert.equal(calls.upload, undefined);
  });

  // Embedding an image in the note the caller writes above the forwarded block (#13).
  describe('embedded images in the caller\'s note', () => {
    const NOTE_ARGS = {
      originalEmailId: 'o1',
      to: ['x@y.example'],
      htmlBody: '<p>Context:</p><img src="cid:chart">',
      attachments: [{ path: 'chart.png', cid: 'chart' }],
    };
    const CHART = { blobId: 'blob-chart', type: 'image/png', name: 'chart.png', disposition: 'inline', cid: 'chart' };

    it('tells the upload which Content-ID the note displays', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeForward(NOTE_ARGS, client, '/attach/root');
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports what the saved draft embeds, after re-reading it', async () => {
      const { client } = spyClient({
        getEmailById: async (id) => {
          if (id === 'o1') return makeOriginal();
          return { id, attachments: [{ blobId: 'blob-chart', cid: 'chart', size: 2048, disposition: 'inline' }] };
        },
      }, [CHART]);
      const r = await composeForward(NOTE_ARGS, client, '/attach/root');
      assert.deepEqual(r.notes, ['This draft embeds 1 image(s) (2 KB).']);
    });

    it('rejects a note reference nothing supplies, in the forward wording', async () => {
      const { client, calls } = spyClient();
      await assert.rejects(
        () => composeForward(
          { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<img src="cid:chart">' },
          client, '/attach/root',
        ),
        (e: any) => /your note/.test(e.message) && /Quoted images appear inside the quote automatically/.test(e.message),
      );
      assert.equal(calls.upload, undefined);
      assert.equal(calls.draft, undefined);
    });

    it('does not re-read the draft for an ordinary forward', async () => {
      const { client, calls } = spyClient();
      await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined);
      assert.deepEqual(calls.gets, ['o1']);
    });

    it('sees attachments supplied as a JSON string (lenient client)', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeForward(
        { ...NOTE_ARGS, attachments: '[{"path":"chart.png","cid":"chart"}]' },
        client, '/attach/root',
      );
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports a malformed note before an attachment problem', async () => {
      const { client } = spyClient();
      await assert.rejects(
        () => composeForward(
          { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>FYI</p><img src="cid:missing"><![CDATA[x]]>' },
          client, '/attach/root',
        ),
        /htmlBody contains a CDATA section/,
      );
    });
  });
});

describe('buildForwardParams — recorded source instance', () => {
  it("records the original's JMAP id as sourceEmailId on an inline forward", () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.equal(forwardParams.sourceEmailId, 'orig-1');
  });

  it('records it on an asAttachment forward too', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, makeOriginal());
    assert.equal(forwardParams.sourceEmailId, 'orig-1');
  });

  it('records it even when the original has no settable Message-ID (id refines WHICH copy; the header still gates marking)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ messageId: undefined }));
    assert.equal(forwardParams.forwardedMessageId, undefined);
    assert.equal(forwardParams.sourceEmailId, 'orig-1');
  });
});
