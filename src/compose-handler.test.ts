import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { composeDraft } from './compose-handler.js';
import type { ComposeClient } from './compose-handler.js';
import { InvalidInputError } from './coerce.js';

const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];

// A sending identity with a configured sign-off, for the appendSignature cases (#33).
const SIGNED_IDENTITY = {
  id: 'id-1',
  name: 'Test User',
  email: 'me@example.com',
  mayDelete: false,
  textSignature: 'Kind regards,\nTest User',
  htmlSignature: '<div>Kind regards,</div><div>Test User</div>',
};

// A spy ComposeClient: records what each method was called with; uploadAttachments
// returns the canned UPLOADED parts so we can assert they thread through. getEmailById is
// only ever reached by the post-save confirmation read, so its call count is itself an
// assertion target: an ordinary draft must not pay for a round trip it has no use for.
function spyClient(over: Partial<ComposeClient> = {}, uploadResult: any[] = UPLOADED) {
  const calls: any = { readBacks: 0, identityLookups: 0 };
  const client: ComposeClient = {
    getIdentities: async () => { calls.identityLookups += 1; return [SIGNED_IDENTITY]; },
    uploadAttachments: async (specs, dir, allowBlob, options) => { calls.upload = { specs, dir, allowBlob, options }; return uploadResult; },
    createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    getEmailById: async (id) => { calls.readBacks += 1; calls.readId = id; return { id, attachments: [] }; },
    ...over,
  };
  return { client, calls };
}

// An uploaded part as it comes back when the message really does display it.
function inlinePart(cid: string, name = 'logo.png') {
  return { blobId: 'blob-' + cid, type: 'image/png', name, disposition: 'inline', cid };
}

describe('composeDraft — creation and summary fields', () => {
  it('creates the draft and returns the fields the handler summarises', async () => {
    const { client, calls } = spyClient();
    const r = await composeDraft({ to: ['a@b.example'], cc: ['c@d.example'], subject: 'Hi', textBody: 'hello' }, client, undefined, false);
    assert.deepEqual(r, { emailId: 'draft-9', subject: 'Hi', to: ['a@b.example'], cc: ['c@d.example'] });
    assert.equal(calls.draft.subject, 'Hi');
  });

  it('coerces a lone recipient string to an array (lenient-client input)', async () => {
    const { client, calls } = spyClient();
    await composeDraft({ to: 'a@b.example', subject: 'Hi', textBody: 'hello' }, client, undefined, false);
    assert.deepEqual(calls.draft.to, ['a@b.example']);
  });

  // Pin the WHOLE outgoing object, so a field dropped from the handoff fails the suite
  // rather than silently never reaching the server.
  it('passes every supported field through to createDraft', async () => {
    const { client, calls } = spyClient();
    await composeDraft({
      to: ['a@b.example'],
      cc: ['c@d.example'],
      bcc: ['e@f.example'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: [{ path: 'a.pdf' }],
    }, client, '/attach/root', false);
    assert.deepEqual(calls.draft, {
      to: ['a@b.example'],
      cc: ['c@d.example'],
      bcc: ['e@f.example'],
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
    await composeDraft(
      { to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root', false,
    );
    // The upload is told which Content-IDs the message displays; an ordinary attachment
    // displays none, so the set is empty and the part is dispositioned as a file.
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'a.pdf' }], dir: '/attach/root', allowBlob: false, options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments, UPLOADED);
  });

  // The other half of the flag. Every case above passes false, so a handler that dropped
  // the argument and hardcoded "off" would still satisfy them — and the only symptom would
  // be an in-account source refused on a server configured to allow it.
  it('passes the blob opt-in through when it is on', async () => {
    const { client, calls } = spyClient();
    await composeDraft(
      { to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [{ blobId: 'G1', name: 'a.pdf' }] },
      client, undefined, true,
    );
    assert.deepEqual(calls.upload, {
      specs: [{ blobId: 'G1', name: 'a.pdf' }], dir: undefined, allowBlob: true, options: { inlineCids: new Set() },
    });
  });

  it('does not upload when no attachments are given', async () => {
    const { client, calls } = spyClient();
    await composeDraft({ to: ['a@b.example'], subject: 'Hi', textBody: 'hello' }, client, '/attach/root', false);
    assert.equal(calls.upload, undefined);
  });

  it('allows an attachment-only draft (attachments count as content)', async () => {
    const { client, calls } = spyClient();
    const r = await composeDraft({ attachments: [{ path: 'a.pdf' }] }, client, '/attach/root', false);
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.draft.attachments, UPLOADED);
  });

  it('rejects a draft with nothing in it at all', async () => {
    const { client } = spyClient();
    await assert.rejects(() => composeDraft({}, client, undefined, false), /At least one of to, subject, textBody, htmlBody, or attachments/);
  });
});

describe('composeDraft — malformed bodies are rejected before the draft is created', () => {
  it('rejects a non-string body (#62)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', textBody: ['a', 'b'] }, client, undefined, false),
      (e: any) => e instanceof InvalidInputError && /textBody must be a string; received array/.test(e.message),
    );
    assert.equal(calls.draft, undefined);
  });

  it('rejects an entirely HTML-escaped htmlBody (#71/#77)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', htmlBody: '&lt;p&gt;Hi&lt;/p&gt;' }, client, undefined, false),
      /htmlBody appears to be HTML-escaped/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('rejects a CDATA-wrapped body (#78)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>' }, client, undefined, false),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', textBody: '<![CDATA[Hi]]>' }, client, undefined, false),
      /textBody is wrapped in a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('rejects before uploading attachments (nothing hits the blob store)', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({ subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>', attachments: [{ path: 'a.pdf' }] }, client, '/attach/root', false),
      /CDATA/,
    );
    assert.equal(calls.upload, undefined);
  });

  it('still accepts legitimate bodies that look superficially similar', async () => {
    const { client, calls } = spyClient();
    await composeDraft({ subject: 'Hi', htmlBody: '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>', textBody: 'Close it with ]]> here' }, client, undefined, false);
    assert.equal(calls.draft.htmlBody, '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>');
  });
});

// Coercing a helper correctly is only half the guarantee — the other half is that the
// handler actually routes the parameter through it. These cover the threading headers on
// the create_draft path, where an uncoerced string would reach JMAP as a per-character
// header list instead of one Message-ID.
describe('composeDraft — threading-header coercion', () => {
  async function draftWith(args: object): Promise<any> {
    const { client, calls } = spyClient();
    await composeDraft({ subject: 'threading probe', ...args }, client, undefined, false);
    return calls.draft;
  }

  it('leaves a real array untouched', async () => {
    const params = await draftWith({ inReplyTo: ['<a@example.com>'], references: ['<a@example.com>'] });
    assert.deepEqual(params.inReplyTo, ['<a@example.com>']);
    assert.deepEqual(params.references, ['<a@example.com>']);
  });

  it('wraps a bare single Message-ID string in an array', async () => {
    const params = await draftWith({ inReplyTo: '<a@example.com>' });
    assert.deepEqual(params.inReplyTo, ['<a@example.com>']);
  });

  it('parses a JSON-stringified array', async () => {
    const params = await draftWith({ references: '["<a@example.com>", "<b@example.com>"]' });
    assert.deepEqual(params.references, ['<a@example.com>', '<b@example.com>']);
  });

  it('splits a comma-separated string', async () => {
    const params = await draftWith({ references: '<a@example.com>, <b@example.com>' });
    assert.deepEqual(params.references, ['<a@example.com>', '<b@example.com>']);
  });

  it('leaves both headers undefined when neither is supplied', async () => {
    const params = await draftWith({});
    assert.equal(params.inReplyTo, undefined);
    assert.equal(params.references, undefined);
  });
});

describe('composeDraft — embedding the caller\'s own images (#13)', () => {
  const EMBED_ARGS = {
    to: ['a@b.example'],
    subject: 'Logo',
    htmlBody: '<p>See:</p><img src="cid:logo">',
    attachments: [{ path: 'logo.png', cid: 'logo' }],
  };

  it('tells the upload which Content-ID the body displays', async () => {
    const { client, calls } = spyClient({}, [inlinePart('logo')]);
    await composeDraft(EMBED_ARGS, client, '/attach/root', false);
    assert.deepEqual(calls.upload.options, { inlineCids: new Set(['logo']) });
  });

  it('normalises the supplied spelling before matching it to the reference', async () => {
    const { client, calls } = spyClient({}, [inlinePart('logo')]);
    await composeDraft(
      { ...EMBED_ARGS, attachments: [{ path: 'logo.png', cid: '<cid:logo>' }] },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.upload.options, { inlineCids: new Set(['logo']) });
  });

  it('reports what the saved draft embeds, with the confirmed size', async () => {
    let readBacks = 0;
    const { client } = spyClient({
      getEmailById: async (id) => {
        readBacks += 1;
        return { id, attachments: [{ blobId: 'blob-logo', cid: 'logo', size: 214000, disposition: 'inline' }] };
      },
    }, [inlinePart('logo')]);
    const r = await composeDraft(EMBED_ARGS, client, '/attach/root', false);
    assert.deepEqual(r.notes, ['This draft embeds 1 image(s) (209 KB).']);
    assert.equal(readBacks, 1);
  });

  it('does not read the draft back when nothing was embedded', async () => {
    const { client, calls } = spyClient();
    const r = await composeDraft(
      { to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root', false,
    );
    assert.equal(calls.readBacks, 0);
    assert.equal(r.notes, undefined);
  });

  it('says so when the saved draft could not be re-read', async () => {
    const { client } = spyClient({
      getEmailById: async () => { throw new Error('network'); },
    }, [inlinePart('logo')]);
    const r = await composeDraft(EMBED_ARGS, client, '/attach/root', false);
    assert.equal(r.emailId, 'draft-9');
    assert.ok(r.notes?.some((n) => /could not re-read it to confirm/.test(n)));
  });

  it('says so when the saved draft is short of an image this call attached', async () => {
    const { client } = spyClient({
      getEmailById: async (id) => ({ id, attachments: [] }),
    }, [inlinePart('logo')]);
    const r = await composeDraft(EMBED_ARGS, client, '/attach/root', false);
    assert.ok(r.notes?.some((n) => /1 embedded image\(s\) this call attached were not found/.test(n)));
  });

  it('rejects a reference no attachment supplies, before anything is uploaded', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft(
        { to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:missing">', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root', false,
      ),
      (e: any) => e instanceof InvalidInputError
        && /references cid "missing" but no attachment supplies it/.test(e.message)
        && /add an attachments item with cid: "missing"/.test(e.message),
    );
    assert.equal(calls.upload, undefined);
    assert.equal(calls.draft, undefined);
  });

  it('drops the add-an-attachment repair when attachments are disabled', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => composeDraft(
        { to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:missing">' },
        client, undefined, false,
      ),
      (e: any) => e instanceof InvalidInputError
        && /neither FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is set/.test(e.message)
        && !/add an attachments item with cid/.test(e.message),
    );
  });

  it('rejects a reference to a server-managed identifier', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => composeDraft(
        { to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:ii-' + 'a'.repeat(32) + '@inline.invalid">' },
        client, '/attach/root', false,
      ),
      (e: any) => e instanceof InvalidInputError && /server-managed identifier/.test(e.message),
    );
  });

  it('rejects two attachments sharing one Content-ID, interpolating the real count', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => composeDraft({
        to: ['a@b.example'],
        subject: 'Hi',
        htmlBody: '<img src="cid:logo">',
        // Two spellings of ONE identifier: the collision is judged on the canonical value.
        attachments: [{ path: 'a.png', cid: 'logo' }, { path: 'b.png', cid: '<logo>' }],
      }, client, '/attach/root', false),
      (e: any) => e instanceof InvalidInputError && /2 attachments/.test(e.message) && /"logo"/.test(e.message),
    );
    assert.equal(calls.upload, undefined);
  });

  // The file is the caller's; a message that cannot display it still carries it.
  it('attaches a cid-bearing file to a text-only draft and says it was not embedded', async () => {
    const degraded = [{ blobId: 'blob-logo', type: 'image/png', name: 'logo.png', disposition: 'attachment', cid: 'logo' }];
    const { client, calls } = spyClient({}, degraded);
    const r = await composeDraft(
      { to: ['a@b.example'], subject: 'Hi', textBody: 'no html here', attachments: [{ path: 'logo.png', cid: 'logo' }] },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.upload.options, { inlineCids: new Set() });
    assert.deepEqual(calls.draft.attachments, degraded);
    assert.deepEqual(r.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
    assert.equal(calls.readBacks, 0);
  });

  // The second route to the same outcome: an html body DOES ship, it just displays
  // something else. The note must not claim there was no html body.
  it('says an unreferenced image was attached even when an html body ships', async () => {
    const degraded = [{ blobId: 'blob-spare', type: 'image/png', name: 'spare.png', disposition: 'attachment', cid: 'spare' }];
    const { client, calls } = spyClient({}, degraded);
    const r = await composeDraft({
      to: ['a@b.example'],
      subject: 'Hi',
      htmlBody: '<p>text, and no image reference at all</p>',
      attachments: [{ path: 'spare.png', cid: 'spare' }],
    }, client, '/attach/root', false);
    assert.deepEqual(calls.upload.options, { inlineCids: new Set() });
    assert.deepEqual(r.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('notes reference-shaped text it could not act on, without refusing the message', async () => {
    const { client, calls } = spyClient({}, [inlinePart('logo')]);
    const r = await composeDraft({
      ...EMBED_ARGS,
      htmlBody: '<p>Reference it as cid:whatever in the css</p><img src="cid:logo">',
    }, client, '/attach/root', false);
    assert.equal(calls.draft.emailId, undefined);
    assert.ok(r.notes?.some((n) => /looks like an embedded-image \(cid:\) reference/.test(n)));
  });

  // A lenient client may send the whole array as a JSON string; the checks must see the
  // items, not an opaque string that would make every reference look unsupplied.
  it('sees attachments supplied as a JSON string (lenient client)', async () => {
    const { client, calls } = spyClient({}, [inlinePart('logo')]);
    await composeDraft(
      { ...EMBED_ARGS, attachments: '[{"path":"logo.png","cid":"logo"}]' },
      client, '/attach/root', false,
    );
    assert.deepEqual(calls.upload.options, { inlineCids: new Set(['logo']) });
  });

  it('reports a malformed body before an attachment problem', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => composeDraft(
        { to: ['a@b.example'], subject: 'Hi', htmlBody: '<p>ok</p><img src="cid:missing"><![CDATA[x]]>' },
        client, '/attach/root', false,
      ),
      /htmlBody contains a CDATA section/,
    );
  });
});
