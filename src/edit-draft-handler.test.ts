// The edit_draft orchestration, exercised through its injected client.
//
// This seam existed only inside the CallTool switch until the attachment sources grew a
// second capability gate. A gate that can only be observed by running the real server
// against a real account is not regression protection — so the orchestration moved into
// editDraft() and the branches below are covered here, with no credentials and no network.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { editDraft } from './edit-draft-handler.js';
import type { EditDraftClient } from './edit-draft-handler.js';
import { InvalidInputError, PathAccessError } from './coerce.js';
import { McpError } from '@modelcontextprotocol/sdk/types.js';

const UPLOADED: any[] = [{ blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' }];

const RESULT: any = { id: 'draft-new', replacedDraft: { id: 'draft-old' }, trashedOldDraftId: 'draft-old' };

// A spy EditDraftClient. uploadAttachments records its arguments — including the two
// capability flags, which is the whole point of the extraction — and returns canned parts
// so they can be asserted where they land on the update.
function spyClient(over: Partial<EditDraftClient> = {}) {
  const calls: any = {};
  const client: EditDraftClient = {
    uploadAttachments: async (specs, dir, allowBlob, options) => {
      calls.upload = { specs, dir, allowBlob, options };
      return UPLOADED;
    },
    updateDraft: async (emailId, updates, options) => {
      calls.update = { emailId, updates, options };
      return RESULT;
    },
    ...over,
  };
  return { client, calls };
}

describe('editDraft — coercion and delegation', () => {
  it('coerces the caller fields and passes them to updateDraft', async () => {
    const { client, calls } = spyClient();
    const r = await editDraft(
      { emailId: 'd1', to: 'a@b.example', subject: 'Hi', clearFields: 'cc', removeAttachments: 'blob-9', expandSignature: 'true' },
      client, undefined, false,
    );
    assert.equal(r, RESULT);
    assert.equal(calls.update.emailId, 'd1');
    assert.deepEqual(calls.update.updates.to, ['a@b.example']);
    assert.deepEqual(calls.update.updates.clearFields, ['cc']);
    assert.deepEqual(calls.update.updates.removeAttachments, ['blob-9']);
    // Lenient clients stringify booleans; "true" must mean true here or a caller that
    // asked for its {{signature}} to expand would silently get the braces stored.
    assert.equal(calls.update.updates.expandSignature, true);
    assert.equal(calls.upload, undefined);
  });

  it('never reads a stringified expandSignature as true unless it says true', async () => {
    const { client, calls } = spyClient();
    await editDraft({ emailId: 'd1', subject: 'Hi', expandSignature: 'garbage' }, client, undefined, false);
    assert.equal(calls.update.updates.expandSignature, false);
  });

  // The hash is NOT checked here. updateDraft owns the refusal order — the body-shape
  // coupling guards name the shape the caller has to fix, and complaining about a stale
  // read ahead of that would be no use to it. So the handler's job is to hand the value
  // through untouched, including when it is absent.
  it('passes bodyHash through to updateDraft without validating it', async () => {
    const { client, calls } = spyClient();
    await editDraft({ emailId: 'd1', textBody: 'Hi', bodyHash: 'bh1-deadbeef' }, client, undefined, false);
    assert.equal(calls.update.updates.bodyHash, 'bh1-deadbeef');
  });

  it('reaches updateDraft with no bodyHash rather than refusing a body edit itself', async () => {
    const { client, calls } = spyClient();
    await editDraft({ emailId: 'd1', textBody: 'Hi' }, client, undefined, false);
    assert.equal(calls.update.updates.bodyHash, undefined);
    assert.equal(calls.update.updates.textBody, 'Hi');
  });

  it('requires an emailId', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => editDraft({ subject: 'Hi' }, client, undefined, false),
      (e: unknown) => e instanceof McpError && /emailId is required/.test((e as Error).message),
    );
  });

  it('refuses a malformed body before any attachment is read', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => editDraft(
        { emailId: 'd1', htmlBody: '<![CDATA[<p>Hi</p>]]>', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root', false,
      ),
      /htmlBody contains a CDATA section/,
    );
    assert.equal(calls.upload, undefined);
    assert.equal(calls.update, undefined);
  });

  // Order-discriminating: the input above is defective on BOTH sides, so only the ordering
  // decides which error the caller sees. The body guard must win, because that is what the
  // compose tools do with the identical input — an edit that reported the attachments key
  // first would make the same mistake diagnosable two different ways depending on the tool.
  it('reports the body defect ahead of an attachments defect, as the compose tools do', async () => {
    const { client, calls } = spyClient();
    const args = {
      emailId: 'd1',
      htmlBody: '<![CDATA[<p>Hi</p>]]>',
      attachments: [{ bogusKey: 1 }],
    };
    await assert.rejects(
      () => editDraft(args, client, '/attach/root', false),
      (e: unknown) => {
        const m = (e as Error).message;
        assert.match(m, /htmlBody contains a CDATA section/);
        assert.doesNotMatch(m, /bogusKey/);
        return true;
      },
    );
    assert.equal(calls.upload, undefined);
    assert.equal(calls.update, undefined);
  });
});

describe('editDraft — attachment sources and their gates', () => {
  it('threads uploaded parts onto the update, with both capability flags', async () => {
    const { client, calls } = spyClient();
    await editDraft({ emailId: 'd1', attachments: [{ path: 'a.pdf' }] }, client, '/attach/root', false);
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'a.pdf' }], dir: '/attach/root', allowBlob: false, options: undefined,
    });
    assert.deepEqual(calls.update.updates.attachments, UPLOADED);
  });

  it('passes the blob opt-in through, so an in-account source is reachable from an edit', async () => {
    const { client, calls } = spyClient();
    await editDraft(
      { emailId: 'd1', attachments: [{ emailId: 'M1', attachmentId: 'p2' }] },
      client, undefined, true,
    );
    assert.equal(calls.upload.allowBlob, true);
    assert.equal(calls.upload.dir, undefined);
  });

  // The gate itself lives in uploadAttachments; what this proves is that edit_draft routes
  // to it rather than quietly attaching nothing, and that the refusal reaches the caller.
  it('surfaces the blob gate refusal and writes no update', async () => {
    const { client, calls } = spyClient({
      uploadAttachments: async () => {
        throw new InvalidInputError('attachments[0] attaches by blobId, which is disabled. Set FASTMAIL_ALLOW_BLOB_ATTACH=true');
      },
    });
    await assert.rejects(
      () => editDraft({ emailId: 'd1', attachments: [{ blobId: 'G1', name: 'r.pdf' }] }, client, '/attach/root', false),
      (e: unknown) => e instanceof InvalidInputError && /FASTMAIL_ALLOW_BLOB_ATTACH/.test((e as Error).message),
    );
    assert.equal(calls.update, undefined);
  });

  it('surfaces the attach-dir refusal for a local file and writes no update', async () => {
    const { client, calls } = spyClient({
      uploadAttachments: async () => {
        throw new PathAccessError('Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR');
      },
    });
    await assert.rejects(
      () => editDraft({ emailId: 'd1', attachments: [{ path: 'a.pdf' }] }, client, undefined, true),
      (e: unknown) => e instanceof PathAccessError && /FASTMAIL_ATTACH_DIR/.test((e as Error).message),
    );
    assert.equal(calls.update, undefined);
  });

  it('rejects a malformed attachments item by index, before the client is touched', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => editDraft({ emailId: 'd1', attachments: [{ blobId: 'G1' }] }, client, undefined, true),
      (e: unknown) => e instanceof McpError && /attachments\[0\] gives a blobId but no 'name'/.test((e as Error).message),
    );
    assert.equal(calls.upload, undefined);
    assert.equal(calls.update, undefined);
  });

  it('reads attachments sent as a JSON string (lenient client)', async () => {
    const { client, calls } = spyClient();
    await editDraft({ emailId: 'd1', attachments: '[{"path":"a.pdf"}]' }, client, '/attach/root', false);
    assert.deepEqual(calls.upload.specs, [{ path: 'a.pdf' }]);
  });
});

describe('editDraft — what a refusal may suggest', () => {
  // updateDraft's refusal over a missing embedded image offers "add an attachments item"
  // only when some source could actually supply one. EITHER gate being open is enough.
  it('reports attachments enabled when either source is open, and disabled when neither is', async () => {
    for (const [dir, allowBlob, expected] of [
      ['/attach/root', false, true],
      [undefined, true, true],
      ['/attach/root', true, true],
      [undefined, false, false],
    ] as const) {
      const { client, calls } = spyClient();
      await editDraft({ emailId: 'd1', subject: 'Hi' }, client, dir, allowBlob);
      assert.deepEqual(calls.update.options, { attachmentsEnabled: expected });
    }
  });
});
