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

// A fixed stand-in for a minted Content-ID, so the assertions read as values rather than
// as patterns. Real mints come from the CSPRNG; the shape is pinned in inline-images.test.ts.
const MINT = 'ii-00000000000000000000000000000001@inline.invalid';

describe('buildForwardParams — attachment carry', () => {
  const inlinePng = { partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'pic.png', disposition: 'inline', cid: 'img-1' };
  const inlineNoCid = { partId: '5', blobId: 'blob-odd', type: 'application/zip', size: 10, name: 'odd.zip', disposition: 'inline', cid: null };

  it('carries non-inline originals through the whitelist (never server-set partId/size)', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal());
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
    ]);
  });
  // An original whose html displays the embedded image, so the forwarded block can carry it.
  const referencing = (over: any = {}) => makeOriginal({
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:img-1"></p>' } },
    ...over,
  });

  it('carries a body-referenced image under a minted identifier and points the block at it', () => {
    const orig = referencing({ attachments: [inlinePng] });
    const { forwardParams, quoteImages, carry } = buildForwardParams({ to: ['x@y.example'] }, orig, undefined, () => MINT);
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-png', type: 'image/png', name: 'pic.png', cid: MINT, disposition: 'inline' },
    ]);
    assert.match(forwardParams.htmlBody!, new RegExp(`src="cid:${MINT}"`));
    assert.equal(quoteImages.minted.length, 1);
    assert.deepEqual(carry, { pooled: [], attached: [], notIncluded: [] });
  });
  it('pools an inline image the body never references, with its Content-ID stripped', () => {
    const orig = makeOriginal({ attachments: [makeOriginal().attachments[0], inlinePng] });
    const { forwardParams, carry } = buildForwardParams({ to: ['x@y.example'] }, orig);
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
      { blobId: 'blob-png', type: 'image/png', name: 'pic.png', disposition: 'attachment' },
    ]);
    assert.deepEqual(carry.pooled.map((p: any) => p.blobId), ['blob-png']);
    assert.deepEqual(carry.attached.map((p: any) => p.blobId), ['blob-doc']);
  });
  // A part carrying NO disposition, listed only under attachments, is the ordinary shape a
  // body-embedded image arrives in — so what makes it body content is the body referencing
  // it, not its metadata.
  it('pools a referenced image the block could not display, even with no disposition on it', () => {
    const twins = [
      { partId: '7', blobId: 'blob-a', type: 'image/png', size: 20, name: 'a.png', cid: 'a@x' },
      { partId: '8', blobId: 'blob-b', type: 'image/png', size: 30, name: 'b.png', cid: 'a@x' },
    ];
    const orig = makeOriginal({
      bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:a@x"></p>' } },
      attachments: twins,
    });
    const { quoteImages, carry } = buildForwardParams({ to: ['x@y.example'] }, orig, undefined, () => MINT);
    // The Content-ID names two parts, so nothing can be embedded under it.
    assert.deepEqual(quoteImages.minted, []);
    assert.deepEqual(carry.pooled.map((p: any) => p.blobId), ['blob-a', 'blob-b']);
    assert.deepEqual(carry.attached, []);
  });
  it('never carries a foreign Content-ID of the shape this server mints', () => {
    const forged = { partId: '9', blobId: 'blob-forged', type: 'image/png', name: 'f.png', disposition: 'inline', cid: 'ii-0123456789abcdef0123456789abcdef@inline.invalid' };
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ attachments: [forged] }));
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-forged', type: 'image/png', name: 'f.png', disposition: 'attachment' },
    ]);
  });
  it("normalizes inline-WITHOUT-cid to disposition 'attachment' (keeps the draft editable)", () => {
    const orig = makeOriginal({ attachments: [inlineNoCid] });
    const { forwardParams, carry } = buildForwardParams({ to: ['x@y.example'] }, orig);
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-odd', type: 'application/zip', name: 'odd.zip', disposition: 'attachment' },
    ]);
    // Marked inline but with no Content-ID, so nothing could have displayed it: an ordinary
    // file riding along, not a media part the block failed to show.
    assert.deepEqual(carry.pooled, []);
    assert.deepEqual(carry.attached.map((p: any) => p.blobId), ['blob-odd']);
  });
  it('carries the body-referenced image even with includeOriginalAttachments:false, and leaves the files behind', () => {
    const orig = referencing({ attachments: [makeOriginal().attachments[0], inlinePng] });
    const { forwardParams, carry } = buildForwardParams(
      { to: ['x@y.example'], includeOriginalAttachments: false }, orig, undefined, () => MINT,
    );
    assert.deepEqual(forwardParams.attachments!.map((a) => a.blobId), ['blob-png']);
    assert.equal(forwardParams.attachments![0].cid, MINT);
    assert.deepEqual(carry.notIncluded.map((p: any) => p.blobId), ['blob-doc']);
  });
  // The force-carry is bounded to parts the sender declared an image. A body can point an
  // <img> at anything; the bound is what stops a reference dragging an arbitrary file past
  // includeOriginalAttachments:false.
  it('does not force-carry a referenced part the sender did not declare an image', () => {
    const doc = { partId: '7', blobId: 'blob-doc2', type: 'application/pdf', name: 'sneaky.pdf', disposition: 'inline', cid: 'img-1' };
    const orig = referencing({ attachments: [doc] });
    const { forwardParams, carry } = buildForwardParams(
      { to: ['x@y.example'], includeOriginalAttachments: false }, orig, undefined, () => MINT,
    );
    assert.equal(forwardParams.attachments, undefined);
    assert.deepEqual(carry.notIncluded.map((p: any) => p.blobId), ['blob-doc2']);
  });
  it('unions in an embedded image the server routed into a body list rather than attachments', () => {
    const orig = referencing({
      attachments: [],
      htmlBody: [{ partId: 'h', type: 'text/html' }, inlinePng],
    });
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, orig, undefined, () => MINT);
    assert.deepEqual(forwardParams.attachments!.map((a) => a.blobId), ['blob-png']);
  });
  it('omits attachments entirely when the original carries none', () => {
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, makeOriginal({ attachments: [] }));
    assert.equal(forwardParams.attachments, undefined);
  });
  it('a text-only forward mints nothing and pools the image instead', () => {
    const orig = referencing({ attachments: [inlinePng] });
    const { forwardParams, quoteImages, carry } = buildForwardParams(
      { to: ['x@y.example'], textBody: 'see below' }, orig, undefined, () => MINT,
    );
    assert.equal(forwardParams.htmlBody, undefined);
    assert.deepEqual(quoteImages.minted, []);
    assert.equal(quoteImages.htmlQuoteShips, false);
    assert.deepEqual(carry.pooled.map((p: any) => p.blobId), ['blob-png']);
  });
  it('an original whose only content is an embedded image forwards as html, not text', () => {
    const orig = makeOriginal({
      textBody: [],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
      attachments: [inlinePng],
    });
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, orig, undefined, () => MINT);
    assert.equal(forwardParams.textBody, undefined);
    assert.match(forwardParams.htmlBody!, new RegExp(`src="cid:${MINT}"`));
  });
  it('carries an SVG the body displays, like any other embedded image', () => {
    const svg = { partId: '6', blobId: 'blob-svg', type: 'image/svg+xml', name: 'logo.svg', disposition: 'inline', cid: 'img-1' };
    const orig = referencing({ attachments: [svg] });
    const { forwardParams } = buildForwardParams({ to: ['x@y.example'] }, orig, undefined, () => MINT);
    assert.deepEqual(forwardParams.attachments, [
      { blobId: 'blob-svg', type: 'image/svg+xml', name: 'logo.svg', cid: MINT, disposition: 'inline' },
    ]);
  });
  it('asAttachment carries only the .eml and resolves no images (the .eml embeds them losslessly)', () => {
    const orig = referencing({ attachments: [inlinePng] });
    const { forwardParams, quoteImages, carry } = buildForwardParams({ to: ['x@y.example'], asAttachment: true }, orig);
    assert.deepEqual(forwardParams.attachments!.map((a) => a.blobId), ['blob-orig-raw']);
    assert.deepEqual(quoteImages.minted, []);
    assert.deepEqual(carry, { pooled: [], attached: [], notIncluded: [] });
  });
});

describe('composeForward — draft-only orchestration', () => {
  const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'new.pdf', disposition: 'attachment' }];
  // A sending identity with a configured sign-off, for the appendSignature cases (#33).
  const SIGNED_IDENTITY = {
    id: 'id-1',
    name: 'Test User',
    email: 'me@example.com',
    mayDelete: false,
    textSignature: 'Kind regards,\nTest User',
    htmlSignature: '<div>Kind regards,</div><div>Test User</div>',
  };

  // The interface has no transmit method at all — send_draft is the only sender.
  function spyClient(over: Partial<ForwardClient> = {}, uploadResult: any[] = UPLOADED) {
    const calls: any = { gets: [] as string[] };
    const client: ForwardClient = {
      // Serves two jobs: fetching the original, and the post-save confirmation read. The
      // recorded ids are what tell the two apart.
      getEmailById: async (id) => { calls.getId = id; calls.gets.push(id); return makeOriginal(); },
      getIdentities: async () => { calls.identityLookups = (calls.identityLookups ?? 0) + 1; return [SIGNED_IDENTITY]; },
      uploadAttachments: async (specs, dir, allowBlob, options) => { calls.upload = { specs, dir, allowBlob, options }; return uploadResult; },
      createDraft: async (p) => { calls.draft = p; return 'draft-7'; },
      ...over,
    };
    return { client, calls };
  }

  // A client whose confirmation read returns a draft echoing back exactly what the create
  // call attached — which is what a real saved draft holds, and the only way a test can
  // assert on the Content-IDs the forwarded block mints (they are random in production).
  function echoingClient(original: any) {
    const calls: any = { gets: [] as string[] };
    const client: ForwardClient = {
      getEmailById: async (id) => {
        calls.gets.push(id);
        if (id !== 'draft-7') return original;
        return { id, attachments: (calls.draft?.attachments ?? []).map((a: any) => ({ ...a, size: 1024 })) };
      },
      getIdentities: async () => [SIGNED_IDENTITY],
      uploadAttachments: async (specs, dir, allowBlob, options) => { calls.upload = { specs, dir, allowBlob, options }; return UPLOADED; },
      createDraft: async (p) => { calls.draft = p; return 'draft-7'; },
    };
    return { client, calls };
  }

  it('requires originalEmailId', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeForward({ to: ['x@y.example'] }, client, undefined, false), /originalEmailId is required/);
  });
  it('saves a draft and returns its id and subject', async () => {
    const { client, calls } = spyClient();
    const r = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
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
    const r = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'], send: true }, client, undefined, false);
    assert.equal(r.emailId, 'draft-7');
    assert.ok(calls.draft); // the only write is the draft create
  });
  it('appends caller uploads BEHIND the carried originals', async () => {
    const { client, calls } = spyClient();
    await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], attachments: [{ path: 'new.pdf' }] },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'new.pdf' }], dir: '/attach/root', allowBlob: false, options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-doc', 'up-1']);
  });
  // The other half of the flag. Every case above passes false, so a handler that dropped
  // the argument and hardcoded "off" would still satisfy them — and the only symptom would
  // be an in-account source refused on a server configured to allow it.
  it('passes the blob opt-in through when it is on', async () => {
    const { client, calls } = spyClient();
    await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], attachments: [{ blobId: 'G1', name: 'new.pdf' }] },
      client, undefined, true,
    );
    assert.deepEqual(calls.upload, {
      specs: [{ blobId: 'G1', name: 'new.pdf' }], dir: undefined, allowBlob: true, options: { inlineCids: new Set() },
    });
  });

  it('appends caller uploads behind the .eml on asAttachment', async () => {
    const { client, calls } = spyClient();
    await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true, attachments: [{ path: 'new.pdf' }] },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-orig-raw', 'up-1']);
    assert.equal(calls.draft.attachments[0].type, 'message/rfc822');
  });
  const inlinePng = { partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'p.png', disposition: 'inline', cid: 'img-1' };
  const referencingOriginal = (over: any = {}) => makeOriginal({
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:img-1"></p>' } },
    attachments: [inlinePng],
    ...over,
  });

  it('reports the images the forwarded block embeds, with their total size', async () => {
    const { client } = echoingClient(referencingOriginal());
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(d.notes, ['This draft embeds 1 image(s) from the original (1 KB).']);
  });
  it('re-reads the saved draft for a forward whose only images come from the original', async () => {
    const { client, calls } = echoingClient(referencingOriginal());
    await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(calls.gets, ['o1', 'draft-7']);
  });
  it('says so when a carried image is not on the saved draft', async () => {
    const { client } = spyClient({
      // The read-back finds a draft with none of the parts this call attached.
      getEmailById: async (id) => (id === 'draft-7' ? { id, attachments: [] } : referencingOriginal()),
    });
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(d.notes, [
      'This draft embeds 1 image(s) from the original (1 KB).',
      '1 embedded image(s) this call attached were not found on the saved draft.' +
      ' Open the draft to check how it renders.',
    ]);
  });
  it('says plainly when a media part could not be embedded and rode along as a file', async () => {
    const { client } = spyClient({ getEmailById: async () => makeOriginal({ attachments: [inlinePng] }) });
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(d.notes, [
      '1 media part(s) could not be embedded and were attached as regular attachments: "p.png"' +
      ' — re-run with asAttachment: true for full fidelity, then delete this draft.',
    ]);
  });
  it('says a referenced image rode along as a file even when the part declared no disposition', async () => {
    const twins = [
      { partId: '7', blobId: 'blob-a', type: 'image/png', size: 20, name: 'a.png', cid: 'a@x' },
      { partId: '8', blobId: 'blob-b', type: 'image/png', size: 30, name: 'b.png', cid: 'a@x' },
    ];
    const { client } = spyClient({
      getEmailById: async () => makeOriginal({
        bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:a@x"></p>' } },
        attachments: twins,
      }),
    });
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(d.notes, [
      '2 media part(s) could not be embedded and were attached as regular attachments:' +
      ' "a.png", "b.png" — re-run with asAttachment: true for full fidelity, then delete this draft.',
    ]);
  });
  it('reports a reference in the original that matched no part', async () => {
    const { client } = spyClient({
      getEmailById: async () => referencingOriginal({ attachments: [] }),
    });
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.deepEqual(d.notes, [
      "1 image reference(s) in the original's body had no matching part; nothing was carried for them.",
    ]);
  });
  it('counts only the EXCLUDED files in the flag note, and says the embedded images were kept anyway', async () => {
    const withFile = referencingOriginal({
      attachments: [inlinePng, { partId: '3', blobId: 'blob-doc', type: 'application/pdf', size: 1234, name: 'doc.pdf', disposition: 'attachment' }],
    });
    const { client } = echoingClient(withFile);
    const d = await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], includeOriginalAttachments: false }, client, undefined, false,
    );
    assert.deepEqual(d.notes, [
      'This draft embeds 1 image(s) from the original (1 KB).',
      '1 attachment(s), including 0 image(s), were not included because includeOriginalAttachments is false.' +
      ' Body-embedded images were still carried — they are part of the message body.',
    ]);
  });
  it('omits the carried-anyway sentence when the forward embedded nothing', async () => {
    const { client } = spyClient({ getEmailById: async () => makeOriginal() });
    const d = await composeForward(
      { originalEmailId: 'o1', to: ['x@y.example'], includeOriginalAttachments: false }, client, undefined, false,
    );
    assert.deepEqual(d.notes, [
      '1 attachment(s), including 0 image(s), were not included because includeOriginalAttachments is false.',
    ]);
  });
  it('an ordinary forward of a message with a plain attachment says nothing at all', async () => {
    const { client } = spyClient();
    const d = await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.equal(d.notes, undefined);
  });
  it('does not call uploadAttachments when no new attachments are given', async () => {
    let uploadCalled = false;
    const { client } = spyClient({ uploadAttachments: async () => { uploadCalled = true; return []; } });
    await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
    assert.equal(uploadCalled, false);
  });
  it('rejects a malformed note before uploading or saving a draft', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeForward(
        { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<![CDATA[<p>FYI</p>]]>', attachments: [{ path: 'new.pdf' }] },
        client, '/attach/root', false,
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
      await composeForward(NOTE_ARGS, client, '/attach/root', false);
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports what the saved draft embeds, after re-reading it', async () => {
      const { client } = spyClient({
        getEmailById: async (id) => {
          if (id === 'o1') return makeOriginal();
          return { id, attachments: [{ blobId: 'blob-chart', cid: 'chart', size: 2048, disposition: 'inline' }] };
        },
      }, [CHART]);
      const r = await composeForward(NOTE_ARGS, client, '/attach/root', false);
      assert.deepEqual(r.notes, ['This draft embeds 1 image(s) (2 KB).']);
    });

    it('rejects a note reference nothing supplies, in the forward wording', async () => {
      const { client, calls } = spyClient();
      await assert.rejects(
        () => composeForward(
          { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<img src="cid:chart">' },
          client, '/attach/root', false,
        ),
        (e: any) => /your note/.test(e.message) && /Quoted images appear inside the quote automatically/.test(e.message),
      );
      assert.equal(calls.upload, undefined);
      assert.equal(calls.draft, undefined);
    });

    it('does not re-read the draft for an ordinary forward', async () => {
      const { client, calls } = spyClient();
      await composeForward({ originalEmailId: 'o1', to: ['x@y.example'] }, client, undefined, false);
      assert.deepEqual(calls.gets, ['o1']);
    });

    it('sees attachments supplied as a JSON string (lenient client)', async () => {
      const { client, calls } = spyClient({}, [CHART]);
      await composeForward(
        { ...NOTE_ARGS, attachments: '[{"path":"chart.png","cid":"chart"}]' },
        client, '/attach/root', false,
      );
      assert.deepEqual(calls.upload.options, { inlineCids: new Set(['chart']) });
    });

    it('reports a malformed note before an attachment problem', async () => {
      const { client } = spyClient();
      await assert.rejects(
        () => composeForward(
          { originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>FYI</p><img src="cid:missing"><![CDATA[x]]>' },
          client, '/attach/root', false,
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
