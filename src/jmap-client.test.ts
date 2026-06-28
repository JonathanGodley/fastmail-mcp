import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { homedir } from 'os';
import { resolve, join, basename, sep } from 'path';
import { JmapClient } from './jmap-client.js';
import { FastmailAuth } from './auth.js';

// ---------- helpers ----------

const ACCOUNT_ID = 'acct-123';
const IDENTITY = { id: 'id-1', name: 'Test User', email: 'me@example.com', mayDelete: false };
const DRAFTS_MAILBOX = { id: 'mb-drafts', name: 'Drafts', role: 'drafts' };

function makeClient(): JmapClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new JmapClient(auth);

  // Stub getSession so no network call is made
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {},
  }));

  // Default stubs — tests override as needed
  mock.method(client, 'getIdentities', async () => [IDENTITY]);
  mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX]);

  return client;
}

function stubMakeRequest(client: JmapClient, response: any) {
  mock.method(client, 'makeRequest', async () => response);
}

// ---------- tests ----------

describe('createDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  // 1. Happy path
  it('returns email ID on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-42' } } }, 'createDraft'],
      ],
    });

    const id = await client.createDraft({ subject: 'Hello' });
    assert.equal(id, 'email-42');
  });

  // 2. Correct JMAP request structure
  it('sends correct JMAP request structure', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Test', textBody: 'body' });

    assert.equal(makeReq.mock.calls.length, 1);
    const request = makeReq.mock.calls[0].arguments[0];

    // capabilities
    assert.deepEqual(request.using, [
      'urn:ietf:params:jmap:core',
      'urn:ietf:params:jmap:mail',
    ]);

    // method
    assert.equal(request.methodCalls[0][0], 'Email/set');

    // accountId
    assert.equal(request.methodCalls[0][1].accountId, ACCOUNT_ID);

    // email object shape
    const emailObj = request.methodCalls[0][1].create.draft;
    assert.equal(emailObj.subject, 'Test');
    assert.deepEqual(emailObj.from, [{ name: 'Test User', email: 'me@example.com' }]);
    assert.deepEqual(emailObj.keywords, { $draft: true });
    assert.equal(emailObj.mailboxIds[DRAFTS_MAILBOX.id], true);
  });

  // 3. Bug 1 regression — JMAP method-level error throws
  it('throws on JMAP method-level error', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'unknownMethod', description: 'bad call' }, 'createDraft'],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /unknownMethod/);
        assert.match(err.message, /bad call/);
        return true;
      },
    );
  });

  // 4. Bug 2 regression — notCreated includes server type + description
  it('throws with server-provided error details from notCreated', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        [
          'Email/set',
          {
            notCreated: {
              draft: { type: 'invalidProperties', description: 'subject too long' },
            },
          },
          'createDraft',
        ],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /invalidProperties/);
        assert.match(err.message, /subject too long/);
        return true;
      },
    );
  });

  // 5. Bug 3 regression — missing created.draft.id throws
  it('throws when created.draft.id is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { created: { draft: {} } }, 'createDraft'],
      ],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /no email ID/);
        return true;
      },
    );
  });

  // 6. Validation — empty input throws
  it('throws when no meaningful fields are provided', async () => {
    await assert.rejects(
      () => client.createDraft({}),
      (err: Error) => {
        assert.match(err.message, /at least one/i);
        return true;
      },
    );
  });

  // 7. Custom from address used correctly
  it('uses custom from address when provided', async () => {
    const altIdentity = { id: 'id-2', name: 'Alias User', email: 'alias@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [IDENTITY, altIdentity]);

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-7' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Hi', from: 'alias@example.com' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Alias User', email: 'alias@example.com' }]);
  });

  // 8. Invalid from address throws
  it('throws when from address is not a verified identity', async () => {
    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: 'nobody@example.com' }),
      (err: Error) => {
        assert.match(err.message, /not verified/i);
        return true;
      },
    );
  });

  // 8b. Wildcard identity matches concrete from address
  it('matches wildcard identity for from address', async () => {
    const wildcardIdentity = { id: 'id-wild', name: 'Wild User', email: '*@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [wildcardIdentity]);

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-wild' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Hi', from: 'work@example.com' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Wild User', email: 'work@example.com' }]);
  });

  // 8c. Bare @ rejected (no local part)
  it('rejects bare @ address against wildcard identity', async () => {
    const wildcardIdentity = { id: 'id-wild', email: '*@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [wildcardIdentity]);

    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: '@example.com' }),
      (err: Error) => {
        assert.match(err.message, /not verified/i);
        return true;
      },
    );
  });

  // 8d. Wildcard identity does not match different domain
  it('rejects from address that does not match wildcard domain', async () => {
    const wildcardIdentity = { id: 'id-wild', email: '*@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [wildcardIdentity]);

    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: 'work@other.com' }),
      (err: Error) => {
        assert.match(err.message, /not verified/i);
        return true;
      },
    );
  });

  // 9. Provided mailbox (id/role/name) resolved against the mailbox list
  it('saves into the provided mailbox, resolved against the mailbox list', async () => {
    mock.method(client, 'getMailboxes', async () => [
      DRAFTS_MAILBOX,
      { id: 'mb-custom', name: 'Project X', role: null },
    ]);

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-9' } } }, 'createDraft'],
      ],
    }));

    // Resolve by name -> the custom mailbox's id.
    await client.createDraft({ subject: 'Custom', mailbox: 'Project X' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj.mailboxIds['mb-custom'], true);
  });

  it('throws InvalidInputError when the provided mailbox is unknown', async () => {
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX]);
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { created: { draft: { id: 'x' } } }, 'createDraft']],
    });

    await assert.rejects(
      () => client.createDraft({ subject: 'Custom', mailbox: 'nope' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /not found/);
        return true;
      },
    );
  });

  // 10. HTML body constructed correctly
  it('derives a text/plain fallback for an html-only draft', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-10' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Rich', htmlBody: '<p>Hello</p>' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.htmlBody, [{ partId: 'html', type: 'text/html' }]);
    // The fallback is auto-generated as a readable text/plain alternative from the html.
    assert.deepEqual(emailObj.textBody, [{ partId: 'text', type: 'text/plain' }]);
    assert.equal(emailObj.bodyValues.html.value, '<p>Hello</p>');
    assert.match(emailObj.bodyValues.text.value, /Hello/);
  });
});

// ---------- updateDraft ----------

const EXISTING_DRAFT = {
  id: 'draft-1',
  subject: 'Old Subject',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [],
  bcc: [],
  textBody: [{ partId: 'text', type: 'text/plain' }],
  htmlBody: null,
  bodyValues: { text: { value: 'Old body' } },
  mailboxIds: { 'mb-drafts': true },
  keywords: { $draft: true },
};

// A draft with every field populated, for exercising clearFields / empty-reject.
const RICH_DRAFT = {
  id: 'draft-1',
  subject: 'Old Subject',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [{ email: 'carol@example.com' }],
  bcc: [],
  replyTo: [{ email: 'reply@example.com' }],
  textBody: [{ partId: '1', type: 'text/plain' }],
  htmlBody: [{ partId: '2', type: 'text/html' }],
  bodyValues: { '1': { value: 'The text' }, '2': { value: '<p>The html</p>' } },
  mailboxIds: { 'mb-drafts': true },
  keywords: { $draft: true },
};

// Wire makeRequest for create-then-delete: Email/get returns the fixture; the create-only
// Email/set returns a created id; the destroy-only Email/set returns destroyed. Returns the mock.
function mockUpdate(client: JmapClient, fixture: any) {
  return mock.method(client, 'makeRequest', async (req: any) => {
    const [method, params] = req.methodCalls[0];
    if (method === 'Email/get') {
      return { methodResponses: [['Email/get', { list: [fixture] }, 'getEmail']] };
    }
    // Email/set — create-then-delete issues a create-only call, then a destroy-only call.
    if (params.create) {
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
    }
    return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
  });
}

// Pull the recreated draft object out of the create (second overall) call.
function draftFromCall(makeReq: ReturnType<typeof mock.method>) {
  return makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
}

describe('updateDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns new email ID on success (create-then-delete: create first, then destroy)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.orphanedOldDraftId, undefined);

    // Three calls: Email/get, then a create-ONLY Email/set, then a destroy-ONLY Email/set.
    assert.equal(makeReq.mock.calls.length, 3);
    const createCall = makeReq.mock.calls[1].arguments[0].methodCalls[0];
    assert.equal(createCall[0], 'Email/set');
    assert.equal(createCall[1].destroy, undefined); // create call must NOT also destroy
    assert.equal(createCall[1].create.draft.subject, 'New Subject');
    const destroyCall = makeReq.mock.calls[2].arguments[0].methodCalls[0];
    assert.equal(destroyCall[0], 'Email/set');
    assert.deepEqual(destroyCall[1].destroy, ['draft-1']);
    assert.equal(destroyCall[1].create, undefined); // destroy call must NOT also create
  });

  it('merges fields — preserves existing values for unspecified fields', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Updated' });

    // The create call should keep existing to address
    const makeReq = client.makeRequest as ReturnType<typeof mock.method>;
    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ email: 'bob@example.com' }]);
    assert.equal(emailObj.subject, 'Updated');
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...EXISTING_DRAFT, keywords: { $seen: true } };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [nonDraft] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.updateDraft('email-1', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /non-draft/i);
        return true;
      },
    );
  });

  it('throws when email not found', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.updateDraft('missing-id', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('throws on JMAP error during the create call', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['error', { type: 'serverFail', description: 'oops' }, 'updateDraft']] };
    });

    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      (err: Error) => {
        assert.match(err.message, /serverFail/);
        return true;
      },
    );
  });

  // Body-extraction correctness (the `|| true` bug). Fixtures mirror real Fastmail
  // shapes captured live: a single-format draft aliases its one part into BOTH the
  // textBody and htmlBody lists; a dual-format draft has two distinct typed parts.
  // Assertions are on the recreate OUTPUT, whose bodyValues are re-keyed to 'text'/'html'.

  it('preserves a single text-only body without synthesising a phantom html part', async () => {
    const aliasedDraft = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '1', type: 'text/plain' }], // server aliases the one part into both lists
      bodyValues: { '1': { value: 'Plain only' } },
    };
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [aliasedDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'New subject' });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj.htmlBody, undefined);
    assert.deepEqual(emailObj.bodyValues, { text: { value: 'Plain only' } });
  });

  it('preserves both bodies from their own parts on a subject-only edit', async () => {
    const dualDraft = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '1': { value: 'The text' }, '2': { value: '<p>The html</p>' } },
    };
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [dualDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'New subject' });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.bodyValues, {
      text: { value: 'The text' },
      html: { value: '<p>The html</p>' },
    });
  });

  // ---- one-sided guard + text-fallback regeneration on html edit ----

  it('throws when editing textBody alone on a dual-body draft (html is what recipients see)', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'NEW text' }),
      /editing textBody alone.*edit htmlBody.*clearFields:\['htmlBody'\]/s,
    );
  });

  it('regenerates the text fallback when htmlBody is edited alone on a dual-body draft', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>' });
    const draft = draftFromCall(makeReq);
    // The old "The text" is replaced by the fallback regenerated from the NEW html.
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW' }, html: { value: '<p>NEW</p>' } });
  });

  it('writes textBody and drops htmlBody when the partner is named in clearFields', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text', clearFields: ['htmlBody'] });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW text' } });
  });

  it('updates both bodies when both are supplied (no throw)', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text', htmlBody: '<p>NEW</p>' });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.bodyValues, {
      text: { value: 'NEW text' },
      html: { value: '<p>NEW</p>' },
    });
  });

  it('writes textBody on a text-only draft (no partner, stays text-only, no throw)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text' });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW text' } });
  });

  it('regenerates the text fallback when htmlBody is edited alone on a text-only draft', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>' });
    const draft = draftFromCall(makeReq);
    // The old "Old body" text is replaced by the fallback regenerated from the new html.
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW' }, html: { value: '<p>NEW</p>' } });
  });

  it('rejects clearFields:["textBody"] while htmlBody is written (text fallback is auto-managed)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>', clearFields: ['textBody'] }),
      /textBody can't be cleared on its own while htmlBody is present/,
    );
  });

  it('preserves both bodies on a subject-only edit of a dual-body draft (guard does not fire)', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { subject: 'New' });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.bodyValues, {
      text: { value: 'The text' },
      html: { value: '<p>The html</p>' },
    });
  });

  it('saves html-only when an edited htmlBody is image-only (degrade gracefully)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<div><img src="banner.jpg"></div>' });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.textBody, undefined); // no derivable text → no fallback part
    assert.deepEqual(draft.bodyValues, { html: { value: '<div><img src="banner.jpg"></div>' } });
  });

  it('rejects an edited htmlBody that has no visible content (no-body)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p></p>' }),
      /no readable body/,
    );
  });

  it('rejects clearing the only body (a draft needs a body)', async () => {
    mockUpdate(client, EXISTING_DRAFT); // text-only draft
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['textBody'] }),
      /a draft needs a body/,
    );
  });

  // ---- Layer 2: strict empty-reject ----

  it('rejects an empty subject (use clearFields to clear)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: '' }),
      /subject cannot be empty; omit to leave it unchanged, or list it in clearFields to clear it/,
    );
  });

  it('rejects a whitespace-only textBody', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: '   ' }),
      /textBody cannot be empty/,
    );
  });

  it('rejects an empty to with the clearFields hint', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { to: [] }),
      /to cannot be empty; omit to leave it unchanged, or list it in clearFields to clear it/,
    );
  });

  it('rejects an empty cc', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { cc: [] }),
      /cc cannot be empty/,
    );
  });

  it('rejects an empty replyTo', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { replyTo: [] }),
      /replyTo cannot be empty/,
    );
  });

  it('rejects an empty from (from is not clearable)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { from: '' }),
      /from cannot be empty/,
    );
  });

  // ---- Layer 2: clearFields ----

  it('clears cc via clearFields and preserves other fields', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['cc'] });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.cc, []);
    assert.deepEqual(draft.to, [{ email: 'bob@example.com' }]);
    assert.equal(draft.subject, 'Old Subject');
  });

  it('clears to via clearFields (a recipient-less draft is valid)', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['to'] });
    assert.deepEqual(draftFromCall(makeReq).to, []);
  });

  it('clears subject via clearFields', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['subject'] });
    assert.equal(draftFromCall(makeReq).subject, '');
  });

  it('rejects clearing textBody alone on a dual-body draft (text fallback is auto-managed)', async () => {
    // Was: dropped the text part. Now the text fallback is managed automatically, so
    // clearing it while htmlBody survives is rejected (use clearFields:['htmlBody'] for
    // a plain-text email instead).
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['textBody'] }),
      /textBody can't be cleared on its own while htmlBody is present/,
    );
  });

  it('clears htmlBody via clearFields and preserves textBody', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['htmlBody'] });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.deepEqual(draft.bodyValues, { text: { value: 'The text' } });
  });

  it('clears replyTo via clearFields (the spread-omit path)', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['replyTo'] });
    assert.equal(draftFromCall(makeReq).replyTo, undefined);
  });

  it('rejects clearFields:["from"] (from is not clearable)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['from'] }),
      /Cannot clear "from"/,
    );
  });

  it('rejects an unknown clearFields name', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['bogus'] }),
      /Cannot clear "bogus"/,
    );
  });

  it('rejects setting and clearing the same field', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { cc: ['x@y.com'], clearFields: ['cc'] }),
      /cannot both set and clear cc/,
    );
  });

  it('lets the set+clear conflict win over the empty-reject when cc:[] + clearFields:["cc"]', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { cc: [], clearFields: ['cc'] }),
      /cannot both set and clear cc/, // conflict check runs before the empty loop
    );
  });

  it('clearing an already-absent field still succeeds and emits the empty value', async () => {
    // EXISTING_DRAFT.cc is already [] — clear is idempotent, not state-dependent.
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['cc'] });
    assert.deepEqual(draftFromCall(makeReq).cc, []);
  });

  it('a non-empty normal edit still succeeds (regression guard)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    const result = await client.updateDraft('draft-1', { subject: 'Real new subject' });
    assert.equal(result.id, 'draft-2');
    assert.equal(draftFromCall(makeReq).subject, 'Real new subject');
  });

  // ---- faithful recreate: carry threading / attachments / keywords ----

  const DRAFT_WITH_EXTRAS = {
    ...EXISTING_DRAFT,
    keywords: { $draft: true, $flagged: true, 'custom-label': true },
    inReplyTo: ['<orig@example.com>'],
    references: ['<root@example.com>', '<orig@example.com>'],
    attachments: [
      { blobId: 'blob-att', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment', cid: null, partId: '3', size: 1234 },
    ],
  };

  it('carries inReplyTo/references/keywords/attachments through the recreate', async () => {
    const makeReq = mockUpdate(client, DRAFT_WITH_EXTRAS);
    await client.updateDraft('draft-1', { subject: 'New subject' });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.inReplyTo, ['<orig@example.com>']);
    assert.deepEqual(draft.references, ['<root@example.com>', '<orig@example.com>']);
    // keywords merged: $draft preserved alongside $flagged and the custom label
    assert.equal(draft.keywords.$draft, true);
    assert.equal(draft.keywords.$flagged, true);
    assert.equal(draft.keywords['custom-label'], true);
    // attachments carried by blobId, whitelisted fields only (NO partId/size)
    assert.deepEqual(draft.attachments, [
      { blobId: 'blob-att', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
    ]);
  });

  it('rejects editing a draft with an inline (cid:) image', async () => {
    const inlineDraft = {
      ...EXISTING_DRAFT,
      attachments: [{ blobId: 'blob-img', type: 'image/png', disposition: 'inline', cid: 'img@x', name: null, partId: '2', size: 70 }],
    };
    mockUpdate(client, inlineDraft);
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /inline images.*Recreate the draft instead/s,
    );
  });

  it('carries a regular attachment that merely has a cid (disposition not inline)', async () => {
    const cidAttachDraft = {
      ...EXISTING_DRAFT,
      attachments: [{ blobId: 'blob-logo', type: 'image/png', disposition: 'attachment', cid: 'logo@x', name: 'logo.png', partId: '2', size: 99 }],
    };
    const makeReq = mockUpdate(client, cidAttachDraft);
    await client.updateDraft('draft-1', { subject: 'X' }); // not rejected
    assert.deepEqual(draftFromCall(makeReq).attachments, [
      { blobId: 'blob-logo', type: 'image/png', name: 'logo.png', disposition: 'attachment', cid: 'logo@x' },
    ]);
  });

  it('rejects editing a draft with a non-text/non-html body part', async () => {
    const weirdDraft = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/calendar' }],
      htmlBody: null,
      bodyValues: { '1': { value: 'BEGIN:VCALENDAR' } },
    };
    mockUpdate(client, weirdDraft);
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /plain text or HTML.*Recreate the draft instead/s,
    );
  });

  it('does NOT reject a text-only draft aliased into both body lists (alias-aware)', async () => {
    const aliasedDraft = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '1', type: 'text/plain' }], // single part aliased into both lists
      bodyValues: { '1': { value: 'Plain only' } },
    };
    const makeReq = mockUpdate(client, aliasedDraft);
    await client.updateDraft('draft-1', { subject: 'New' }); // must not throw
    assert.equal(draftFromCall(makeReq).subject, 'New');
  });

  // ---- create-then-delete ordering (data-loss prevention) ----

  it('on create failure: throws, issues NO destroy, leaves the old draft untouched', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { notCreated: { draft: { type: 'invalidProperties', description: 'bad blob' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /Failed to create updated draft.*invalidProperties/s,
    );
    // Exactly 2 calls: Email/get + the failed create. The destroy must NEVER be issued.
    assert.equal(makeReq.mock.calls.length, 2);
    assert.ok(makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create);
  });

  it('on destroy failure (notDestroyed): returns new id + orphan warning, does NOT throw', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { notDestroyed: { 'draft-1': { type: 'serverFail' } } }, 'destroyDraft']] };
    });
    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.orphanedOldDraftId, 'draft-1');
  });

  it('on destroy throw (transport error after a good create): returns orphan warning, does NOT throw', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      throw new Error('network down');
    });
    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.orphanedOldDraftId, 'draft-1');
  });

  // ---- reply-quote preservation on body edit (#37, redesigned #42) ----
  //
  // The guard decides on the EXISTING (stored) body, so these fixtures use the RAW body shapes
  // Fastmail returns for reply drafts — captured from a live store/fetch round-trip (2026-06-28)
  // and trimmed of the bulk quoted body but BYTE-EXACT in the marker region the guard reads.
  // Pinning to Fastmail's re-serialized shape (not our buildReplyBodies output) is the point:
  // an html-derived text fallback comes back as "wrote:\n\n> " (blank line). The coercion of
  // noQuote ("true"/"garbage") lives at the index.ts handler seam and is pinned by coerce.test.ts
  // (coerceBool) + the live harness; updateDraft only ever sees a real boolean, so it is not
  // re-tested here.
  const RAW_HTML_QUOTE = '<p>my reply</p><div><br></div><div>On Sun, Jun 28, 2026, at 12:46 AM, PlanningAlerts wrote:</div><blockquote type="cite" style="margin:0 0 0 .8ex;border-left:1px solid #ccc;padding-left:1ex">\n  1 new planning application near 6/30-32 Doomben Ave\n</blockquote>';
  const RAW_TEXT_QUOTE = 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, PlanningAlerts wrote:\n> 2/2 Rowe St Eastwood NSW 2122: Change of Use and Fitout of Pilates Studio\n> \n> Contact us if you have questions.';
  // Quote-LESS bodies (no marker) — for the asymmetric / oldTextQuoted-precondition cells.
  const PLAIN_TEXT = 'my reply with no quote at all';
  const PLAIN_HTML = '<p>my reply with no quote at all</p>';

  const REPLY_BASE = {
    id: 'draft-1', subject: 'Re: Hello',
    from: [{ email: 'me@example.com' }], to: [{ email: 'bob@example.com' }],
    cc: [], bcc: [],
    mailboxIds: { 'mb-drafts': true }, keywords: { $draft: true },
    inReplyTo: ['orig-msg@example.com'], references: ['orig-msg@example.com'],
  };
  // dual: text/plain + text/html, both quoted (the common shape this server creates).
  const DUAL_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: RAW_TEXT_QUOTE }, h: { value: RAW_HTML_QUOTE } } };
  // text-only: the ONE text/plain part aliases into BOTH lists (so bodyValueForType('text/html')
  // is undefined → existingHtmlValue blank). This is the #42 shape.
  const TEXT_ONLY_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: RAW_TEXT_QUOTE } } };
  // html-only: the ONE text/html part aliases into both lists (a foreign-client shape — an
  // html-only reply_email is actually stored dual; included to exercise the html-only path).
  const HTML_ONLY_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { h: { value: RAW_HTML_QUOTE } } };
  // asymmetric: html present but quote-LESS; the quote lives only in the text.
  const ASYMMETRIC_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: RAW_TEXT_QUOTE }, h: { value: PLAIN_HTML } } };
  // dual where only the HTML carries the quote; the text is plain (pins the oldTextQuoted
  // precondition on the plain-text-conversion carve-out).
  const HTMLQUOTE_ONLY_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: PLAIN_TEXT }, h: { value: RAW_HTML_QUOTE } } };

  // The message the reply draft replies to (distinct id 'orig-1'; fully body-valued so the
  // regenerate path produces a real quote we can assert against).
  const ORIGINAL_FOR_REPLY = {
    id: 'orig-1',
    messageId: ['orig-msg@example.com'],
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [{ partId: 'ot', type: 'text/plain' }],
    htmlBody: [{ partId: 'oh', type: 'text/html' }],
    bodyValues: { ot: { value: 'ORIGINAL TEXT BODY' }, oh: { value: '<p>ORIGINAL HTML BODY</p>' } },
  };

  // A non-quotable original (attachment-only: no text/html parts) — buildReplyBodies skips the
  // quote AND attribution for it, so the keep path can't restore a quote and must reject loudly.
  const NONQUOTABLE_ORIGINAL = {
    id: 'orig-empty',
    messageId: ['orig-msg@example.com'],
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [], htmlBody: [], bodyValues: {},
  };

  // Dispatch Email/get BY ID — the chosen draft fixture for the draft id, the original fixture
  // for 'orig-1'. A single-fixture mock would make the regenerate test quote the DRAFT as its
  // own original and prove nothing, so id-dispatch is mandatory here. 'orig-missing' → notFound
  // (drives the not-found path). getEmailById issues Email/get + Mailbox/get; we answer only
  // Email/get (its mailbox read is defensive/optional).
  function mockReplyUpdate(c: JmapClient, draft: any = DUAL_REPLY) {
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        const id = params.ids?.[0];
        if (id === 'orig-1') return { methodResponses: [['Email/get', { list: [ORIGINAL_FOR_REPLY] }, 'email']] };
        if (id === 'orig-empty') return { methodResponses: [['Email/get', { list: [NONQUOTABLE_ORIGINAL] }, 'email']] };
        if (id === 'orig-missing') return { methodResponses: [['Email/get', { list: [], notFound: ['orig-missing'] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [draft] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
  }

  // Find the create call by predicate — the regenerate path inserts a second Email/get, so
  // the create is no longer at a fixed index.
  function createdDraft(makeReq: ReturnType<typeof mock.method>) {
    const call = makeReq.mock.calls.find((c: any) => c.arguments[0].methodCalls[0][1].create);
    return call!.arguments[0].methodCalls[0][1].create.draft;
  }

  // -- dual-body reply draft --

  it('rejects editing htmlBody on a dual reply draft without a flag', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>just my new reply</p>' }),
      /would drop the quoted original.*originalEmailId/s,
    );
  });

  it('rejects editing htmlBody even when the new html itself has a quote marker (no new-body scan)', async () => {
    // Under the redesign the decision is on the OLD body, so a caller-supplied quote in the new
    // html does NOT exempt the edit (the fork.8 #37 behavior — "new html with marker passes" —
    // is deliberately reversed: it was bypassable).
    mockReplyUpdate(client, DUAL_REPLY);
    const html = '<p>my edited reply</p><blockquote type="cite">a different quote</blockquote>';
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: html }),
      /would drop the quoted original/,
    );
  });

  it('regenerates and keeps the html quote from originalEmailId (dual)', async () => {
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>my edited reply</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    // Regenerated html carries the caller's new text AND the ORIGINAL's body (not the draft's).
    assert.match(draft.bodyValues.html.value, /my edited reply/);
    assert.match(draft.bodyValues.html.value, /ORIGINAL HTML BODY/);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);
    // A non-empty text fallback regenerates from the combined html (quote-bearing).
    assert.ok(!isBlankStr(draft.bodyValues.text.value));
    assert.match(draft.bodyValues.text.value, /ORIGINAL HTML BODY/);
  });

  it('regenerates the quote into BOTH bodies when both are written + originalEmailId (no silent text-side drop)', async () => {
    // A caller editing both a new html and a custom text alternative on the keep path must NOT
    // lose the quote on the text side: the quote is rebuilt into both formats.
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>edited html</p>', textBody: 'edited text', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.html.value, /edited html/);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);          // html quote kept
    assert.match(draft.bodyValues.html.value, /ORIGINAL HTML BODY/);
    assert.match(draft.bodyValues.text.value, /edited text/);
    assert.match(draft.bodyValues.text.value, /> ORIGINAL TEXT BODY/);              // text quote kept too
  });

  it('drops the quote from BOTH bodies on noQuote:true when both are written', async () => {
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>bare html</p>', textBody: 'bare text', noQuote: true });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>bare html</p>');
    assert.equal(draft.bodyValues.text.value, 'bare text');
  });

  it('drops the quote on noQuote:true (no second fetch)', async () => {
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>bare reply</p>', noQuote: true });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>bare reply</p>');
    // No keep → no second Email/get for an original.
    const getCalls = makeReq.mock.calls.filter((c: any) => c.arguments[0].methodCalls[0][0] === 'Email/get');
    assert.equal(getCalls.length, 1);
  });

  it('throws when originalEmailId and noQuote are both given', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-1', noQuote: true }),
      /not both/,
    );
  });

  it('throws an actionable error when originalEmailId is not found', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-missing' }),
      /originalEmailId 'orig-missing' could not be fetched/,
    );
  });

  it('rejects a self-inconsistent keep: originalEmailId names a non-quotable original', async () => {
    // Reachable only by naming a wrong/empty original (a draft naming its own original can't, by
    // immutability). The keep can't be honored, so fail loudly with an actionable error rather
    // than store a quote-less body — no caller input is lost either way.
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-empty' }),
      /has no quotable content.*noQuote/s,
    );
  });

  // -- text-only reply draft (#42) --

  it('rejects editing textBody on a text-only reply draft without a flag (#42)', async () => {
    mockReplyUpdate(client, TEXT_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'just my new reply' }),
      /would drop the quoted original.*originalEmailId/s,
    );
  });

  it('regenerates the text quote from originalEmailId and stays text-only', async () => {
    const makeReq = mockReplyUpdate(client, TEXT_ONLY_REPLY);
    await client.updateDraft('draft-1', { textBody: 'my edited reply', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.text.value, /my edited reply/);
    assert.match(draft.bodyValues.text.value, /> ORIGINAL TEXT BODY/); // regenerated "> " text quote
    assert.equal(draft.htmlBody, undefined);                            // stays text-only
  });

  it('drops the text quote on noQuote:true (text-only, stays text-only)', async () => {
    const makeReq = mockReplyUpdate(client, TEXT_ONLY_REPLY);
    await client.updateDraft('draft-1', { textBody: 'bare reply', noQuote: true });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.text.value, 'bare reply');
    assert.equal(draft.htmlBody, undefined);
  });

  it('format-flip: htmlBody + originalEmailId on a text-only reply draft becomes dual-body', async () => {
    const makeReq = mockReplyUpdate(client, TEXT_ONLY_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>now html</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    // Accepted, pinned behavior: the caller chose to add html.
    assert.match(draft.bodyValues.html.value, /now html/);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);
    assert.ok(!isBlankStr(draft.bodyValues.text.value)); // derived text fallback → dual
  });

  // -- carve-outs (quote-preserving by construction) --

  it('carve-out: a subject-only edit on a quoted reply draft preserves both bodies', async () => {
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { subject: 'Re: Hello (edited)' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.text.value, RAW_TEXT_QUOTE);
    assert.equal(draft.bodyValues.html.value, RAW_HTML_QUOTE);
  });

  it('carve-out: clearFields:["htmlBody"] on a dual reply draft keeps the "> " text quote', async () => {
    // The load-bearing carve-out the over-strict regex would have wrongly rejected.
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { clearFields: ['htmlBody'] });
    const draft = createdDraft(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.equal(draft.bodyValues.text.value, RAW_TEXT_QUOTE);
  });

  // -- guard / coupling-guard interactions --

  it('rejects clearFields:["htmlBody"] + a quote-free textBody on a dual reply draft', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['htmlBody'], textBody: 'plain reply, no quote' }),
      /would drop the quoted original/,
    );
  });

  it('regenerates a text-only quote on clearFields:["htmlBody"] + textBody + originalEmailId', async () => {
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { clearFields: ['htmlBody'], textBody: 'my reply', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.match(draft.bodyValues.text.value, /> ORIGINAL TEXT BODY/);
  });

  it('clearFields:["textBody"] on a dual reply draft hits the textBody-coupling guard, not the quote guard', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['textBody'] }),
      /textBody can't be cleared on its own while htmlBody is present/,
    );
  });

  it('regression: textBody alone on a dual reply draft still hits the textBody-coupling guard', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'new text' }),
      /editing textBody alone won't change what most recipients see/,
    );
  });

  it('precedence: textBody-alone + originalEmailId on a dual draft is owned by the coupling guard (loud reject, no data loss)', async () => {
    // A text-only edit while html survives can't change what recipients render, so the coupling
    // guard rejects regardless of originalEmailId — the keep-intent is moot because the whole
    // edit is rejected (nothing is written). The remedy ("edit htmlBody") then keeps the quote.
    // Pinned so this precedence is intended, not accidental.
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'new text', originalEmailId: 'orig-1' }),
      /editing textBody alone won't change what most recipients see/,
    );
  });

  it('asymmetric draft (quote-less html, quoted text): editing textBody alone → textBody-coupling guard', async () => {
    // Pins coupledTextEdit case (i) on an asymmetric draft, not just the symmetric one.
    mockReplyUpdate(client, ASYMMETRIC_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'new text' }),
      /editing textBody alone won't change what most recipients see/,
    );
  });

  it('clearFields:["htmlBody"] where only the html is quoted falls through to the quote guard', async () => {
    // Pins the oldTextQuoted precondition on the plain-text-conversion carve-out: the surviving
    // text is quote-LESS, so this is NOT a clean carve-out → REJECT.
    mockReplyUpdate(client, HTMLQUOTE_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['htmlBody'] }),
      /would drop the quoted original/,
    );
  });

  it('htmlBody + clearFields:["textBody"] + originalEmailId: quote regenerates, then the textBody-clear coupling guard rejects', async () => {
    // Odd-but-safe: the quote is preserved into the html, then the pre-existing clearFields-
    // textBody coupling guard (shipped behavior, independent of this feature) rejects.
    mockReplyUpdate(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>', clearFields: ['textBody'], originalEmailId: 'orig-1' }),
      /textBody can't be cleared on its own while htmlBody is present/,
    );
  });

  it('originalEmailId on a clear-only edit (nothing to regenerate into) rejects loudly', async () => {
    // dual, quote-less text, clearFields:['htmlBody'] + originalEmailId → no body is being
    // written, so the keep intent can't be honored: loud reject, NOT a silent no-op.
    mockReplyUpdate(client, HTMLQUOTE_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['htmlBody'], originalEmailId: 'orig-1' }),
      /can't regenerate a quote on a body you're not writing/,
    );
  });

  // -- html-only reply draft (foreign-client shape) --

  it('rejects editing htmlBody on an html-only reply draft without a flag', async () => {
    mockReplyUpdate(client, HTML_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>just my new reply</p>' }),
      /would drop the quoted original/,
    );
  });

  it('regenerates and keeps the quote from originalEmailId on an html-only reply draft', async () => {
    const makeReq = mockReplyUpdate(client, HTML_ONLY_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>edited html-only</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.html.value, /edited html-only/);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);
    assert.match(draft.bodyValues.html.value, /ORIGINAL HTML BODY/);
  });

  it('drops the quote on noQuote:true on an html-only reply draft', async () => {
    const makeReq = mockReplyUpdate(client, HTML_ONLY_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>bare html-only</p>', noQuote: true });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>bare html-only</p>');
  });

  // -- non-reply draft --

  it('does not fire the guard on a NON-reply draft (no inReplyTo)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT); // no inReplyTo
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>' });
    assert.equal(draftFromCall(makeReq).bodyValues.html.value, '<p>NEW</p>');
  });
});

// Local non-empty check for the regenerate-fallback assertion.
function isBlankStr(s: string | undefined): boolean {
  return !s || s.trim() === '';
}

// ---------- sendDraft ----------

const SENDABLE_DRAFT = {
  id: 'draft-1',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [{ email: 'cc@example.com' }],
  bcc: [],
  keywords: { $draft: true },
};

const SENT_MAILBOX = { id: 'mb-sent', name: 'Sent', role: 'sent' };

describe('sendDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, SENT_MAILBOX]);
  });

  it('returns submission ID on success', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const subId = await client.sendDraft('draft-1');
    assert.equal(subId, 'sub-1');

    // Verify submission call structure
    const submitCall = makeReq.mock.calls[1].arguments[0];
    assert.equal(submitCall.methodCalls[0][0], 'EmailSubmission/set');
    assert.equal(submitCall.methodCalls[0][1].create.submission.emailId, 'draft-1');
    assert.equal(submitCall.methodCalls[0][1].create.submission.identityId, IDENTITY.id);

    // Verify envelope has all recipients (to + cc)
    const rcptTo = submitCall.methodCalls[0][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 2);
    assert.deepEqual(rcptTo[0], { email: 'bob@example.com' });
    assert.deepEqual(rcptTo[1], { email: 'cc@example.com' });
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...SENDABLE_DRAFT, keywords: { $seen: true } };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [nonDraft] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('email-1'),
      (err: Error) => {
        assert.match(err.message, /non-draft/i);
        return true;
      },
    );
  });

  it('rejects draft with no recipients', async () => {
    const noRecipients = { ...SENDABLE_DRAFT, to: [], cc: [], bcc: [] };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [noRecipients] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /no recipients/i);
        return true;
      },
    );
  });

  it('throws when email not found', async () => {
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('missing-id'),
      (err: Error) => {
        assert.match(err.message, /not found/i);
        return true;
      },
    );
  });

  it('throws on JMAP submission error', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['error', { type: 'forbidden', description: 'not allowed' }, 'submitDraft']] };
    });

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /forbidden/);
        return true;
      },
    );
  });

  // ---- reject an empty body part on send ----
  // Our own tools never originate an empty part, but an externally-created draft can carry
  // one, so these fixtures hand-build the malformed shapes.

  it('rejects a draft with a real text part and an empty html part (names htmlBody)', async () => {
    const emptyHtml = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '1': { value: 'Real text' }, '2': { value: '' } },
    };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [emptyHtml] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      /empty htmlBody that would render blank/,
    );
  });

  it('rejects a draft with an empty text part and a real html part (names textBody)', async () => {
    const emptyText = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '1': { value: '   ' }, '2': { value: '<p>Real html</p>' } },
    };
    mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/get', { list: [emptyText] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      /empty textBody that would render blank/,
    );
  });

  it('sends a clean dual-body draft (both parts non-empty)', async () => {
    const dual = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '1': { value: 'Real text' }, '2': { value: '<p>Real html</p>' } },
    };
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [dual] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal(await client.sendDraft('draft-1'), 'sub-1');
  });

  it('sends a clean text-only draft (absent partner is undefined, not empty)', async () => {
    const textOnly = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '1', type: 'text/plain' }], // server aliases the one part into both lists
      bodyValues: { '1': { value: 'Real text' } },
    };
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [textOnly] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal(await client.sendDraft('draft-1'), 'sub-1');
  });

  it('sends an html-only draft with real content as-is (image-only/html-only mail is valid)', async () => {
    const htmlOnly = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '2', type: 'text/html' }], // single html part aliased into both lists
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '2': { value: '<div><img src="https://x/banner.jpg"></div>' } },
    };
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [htmlOnly] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal(await client.sendDraft('draft-1'), 'sub-1'); // not rejected — html-only is sendable
  });
});

// ---------- JMAP response validation ----------

describe('JMAP response validation', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('throws when methodResponses is missing', async () => {
    stubMakeRequest(client, { sessionState: 's1' });
    await assert.rejects(
      () => client.getEmailById('email-1'),
      (err: Error) => {
        assert.match(err.message, /missing expected method/i);
        return true;
      },
    );
  });

  it('throws when index exceeds methodResponses length', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'serverFail', description: 'oops' }, 'query'],
      ],
    });
    // The visible Email/query came back as an error tag, so getMethodResult throws.
    await assert.rejects(
      () => client.getEmails({}),
      (err: Error) => {
        assert.ok(err.message.length > 0);
        return true;
      },
    );
  });

  it('throws on malformed methodResponses entry', async () => {
    stubMakeRequest(client, {
      methodResponses: ['not-a-tuple' as any],
    });
    await assert.rejects(
      () => client.getEmailById('email-1'),
      (err: Error) => {
        assert.match(err.message, /malformed/i);
        return true;
      },
    );
  });
});

// ---------- searchEmails ----------

describe('searchEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  // makeClient() stubs getMailboxes -> [DRAFTS_MAILBOX] (no trash/junk role), so the
  // default Trash/Spam exclusion resolves nothing here: no inMailboxOtherThan key and no
  // count query are added, and getMailboxes is mocked (not via makeRequest), so the
  // Email/query batch stays makeRequest.calls[0].
  it('returns email list on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1'], total: 1 }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
      ],
    });
    const results = await client.searchEmails({ query: 'test', limit: 10 });
    assert.equal(results.items.length, 1);
    assert.equal(results.items[0].subject, 'Test');
  });

  it('returns empty array when no results', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    });
    const results = await client.searchEmails({ query: 'nonexistent' });
    assert.deepEqual(results.items, []);
  });

  it('throws on JMAP error in query', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'invalidArguments', description: 'bad filter' }, 'query'],
        ['error', { type: 'invalidArguments', description: 'bad filter' }, 'emails'],
      ],
    });
    await assert.rejects(
      () => client.searchEmails({ query: 'test' }),
      (err: Error) => {
        assert.match(err.message, /invalidArguments/);
        return true;
      },
    );
  });

  it('excludeDrafts adds a notKeyword $draft condition (AND-wrapped with the query)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails({ query: 'quarterly', limit: 10, excludeDrafts: true });
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    // text in the base, $draft as its own keyword condition, AND-wrapped.
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.text === 'quarterly'));
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$draft'));
  });

  it('includes drafts by default (no notKeyword; flat text filter)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails({ query: 'quarterly', limit: 10 });
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.text, 'quarterly');
    assert.equal(filter.notKeyword, undefined);
  });
});

// ---------- validateSavePath tests ----------

describe('validateSavePath', () => {
  const allowedDir = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  it('accepts paths within the allowed directory', () => {
    const input = join(allowedDir, 'photo.jpg');
    const result = JmapClient.validateSavePath(input);
    assert.equal(result, input);
  });

  it('accepts paths in subdirectories', () => {
    const input = join(allowedDir, 'andrew', 'assets', 'logo.png');
    const result = JmapClient.validateSavePath(input);
    assert.equal(result, input);
  });

  it('rejects paths outside the allowed directory', () => {
    assert.throws(
      () => JmapClient.validateSavePath('/tmp/evil.sh'),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects path traversal attempts', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${allowedDir}/../../../.bashrc`),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects home directory writes', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${homedir()}/.ssh/authorized_keys`),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects null bytes', () => {
    assert.throws(
      () => JmapClient.validateSavePath(`${allowedDir}/file\0.txt`),
      (err: Error) => {
        assert.match(err.message, /null bytes/);
        return true;
      },
    );
  });

  it('accepts paths within a custom download directory', () => {
    const customDir = resolve('tmp', 'my-downloads');
    const input = join(customDir, 'photo.jpg');
    const result = JmapClient.validateSavePath(input, customDir);
    assert.equal(result, input);
  });

  it('rejects paths outside a custom download directory', () => {
    const customDir = resolve('tmp', 'my-downloads');
    assert.throws(
      () => JmapClient.validateSavePath('/etc/passwd', customDir),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects traversal out of a custom download directory', () => {
    const customDir = '/tmp/my-downloads';
    assert.throws(
      () => JmapClient.validateSavePath(`${customDir}/../../etc/shadow`, customDir),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('resolves a relative savePath against the download directory', () => {
    const customDir = resolve('tmp', 'my-downloads');
    const result = JmapClient.validateSavePath('report.pdf', customDir);
    assert.equal(result, join(customDir, 'report.pdf'));
  });

  it('resolves a relative savePath with subdirectories against the download directory', () => {
    const customDir = resolve('tmp', 'my-downloads');
    const result = JmapClient.validateSavePath(join('thread', 'invoice.pdf'), customDir);
    assert.equal(result, join(customDir, 'thread', 'invoice.pdf'));
  });

  it('rejects a relative savePath that traverses out of the download directory', () => {
    const customDir = resolve('tmp', 'my-downloads');
    assert.throws(
      () => JmapClient.validateSavePath('../escape.pdf', customDir),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });

  it('rejects a Windows drive-absolute path outside the download directory', function () {
    if (sep !== '\\') return; // drive-absolute semantics are win32-only
    const customDir = resolve('C:\\Users\\me\\Downloads', 'fastmail-mcp');
    assert.throws(
      () => JmapClient.validateSavePath('C:\\Windows\\evil.exe', customDir),
      (err: Error) => {
        assert.match(err.message, /must be within/);
        return true;
      },
    );
  });
});

// ---------- safeWritePath (symlink-safe canonicalization) ----------

import { mkdtemp, symlink, rm, mkdir as fsMkdir, writeFile as fsWriteFile } from 'fs/promises';
import { tmpdir } from 'os';

describe('safeWritePath (symlink escapes)', () => {
  it('accepts a normal path inside the allowed directory', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-safe-'));
    try {
      const allowed = join(root, 'allowed');
      await fsMkdir(allowed, { recursive: true });
      const target = join(allowed, 'attachment.bin');
      const safe = await JmapClient.safeWritePath(target, allowed);
      // realpath on macOS may add /private prefix, so just check basename equality
      assert.equal(basename(safe), 'attachment.bin');
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('creates intermediate directories under the canonical allowed dir', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-safe-'));
    try {
      const allowed = join(root, 'allowed');
      await fsMkdir(allowed, { recursive: true });
      const target = join(allowed, 'sub1', 'sub2', 'file.bin');
      const safe = await JmapClient.safeWritePath(target, allowed);
      assert.ok(safe.endsWith(join('sub1', 'sub2', 'file.bin')));
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects writes via a symlink that escapes the allowed directory', async (t) => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-safe-'));
    try {
      const allowed = join(root, 'allowed');
      const outside = join(root, 'outside');
      await fsMkdir(allowed, { recursive: true });
      await fsMkdir(outside, { recursive: true });
      // Symlink inside allowed pointing to outside.
      // Symlink creation requires elevated privileges on Windows; skip where unavailable.
      try {
        await symlink(outside, join(allowed, 'escape'));
      } catch (err) {
        if ((err as any)?.code === 'EPERM' || (err as any)?.code === 'EACCES') {
          t.skip('symlink creation not permitted on this platform');
          return;
        }
        throw err;
      }
      const target = join(allowed, 'escape', 'pwned.bin');
      await assert.rejects(
        () => JmapClient.safeWritePath(target, allowed),
        /outside the allowed directory|symlink escape/i,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('refuses to overwrite a pre-existing symlink at the target path', async (t) => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-safe-'));
    try {
      const allowed = join(root, 'allowed');
      const outside = join(root, 'outside.txt');
      await fsMkdir(allowed, { recursive: true });
      await fsWriteFile(outside, 'orig');
      const target = join(allowed, 'sneaky.bin');
      // Symlink creation requires elevated privileges on Windows; skip where unavailable.
      try {
        await symlink(outside, target);
      } catch (err) {
        if ((err as any)?.code === 'EPERM' || (err as any)?.code === 'EACCES') {
          t.skip('symlink creation not permitted on this platform');
          return;
        }
        throw err;
      }
      await assert.rejects(
        () => JmapClient.safeWritePath(target, allowed),
        /existing symlink/i,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('still rejects lexical traversal even when allowed dir exists', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-safe-'));
    try {
      const allowed = join(root, 'allowed');
      await fsMkdir(allowed, { recursive: true });
      await assert.rejects(
        () => JmapClient.safeWritePath(`${allowed}/../escape.bin`, allowed),
        /must be within/,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

// ---------- sendEmail replyTo ----------

describe('sendEmail replyTo', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getMailboxes', async () => [
      DRAFTS_MAILBOX,
      { id: 'mb-sent', name: 'Sent', role: 'sent' },
    ]);
  });

  it('includes replyTo in JMAP emailObject when provided', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({
      to: ['bob@example.com'],
      subject: 'Test',
      textBody: 'Hello',
      replyTo: ['other@example.com'],
    });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'other@example.com' }]);
  });

  it('does NOT include replyTo when not provided', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({
      to: ['bob@example.com'],
      subject: 'Test',
      textBody: 'Hello',
    });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj.replyTo, undefined);
  });
});

// ---------- sendEmail envelope recipients ----------

describe('sendEmail envelope recipients', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getMailboxes', async () => [
      DRAFTS_MAILBOX,
      { id: 'mb-sent', name: 'Sent', role: 'sent' },
    ]);
  });

  it('includes to, cc, and bcc in envelope rcptTo', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({
      to: ['alice@example.com'],
      cc: ['bob@example.com'],
      bcc: ['charlie@example.com'],
      subject: 'Test',
      textBody: 'Hello',
    });

    const req = makeReq.mock.calls[0].arguments[0];
    const rcptTo = req.methodCalls[1][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 3);
    assert.deepEqual(rcptTo[0], { email: 'alice@example.com' });
    assert.deepEqual(rcptTo[1], { email: 'bob@example.com' });
    assert.deepEqual(rcptTo[2], { email: 'charlie@example.com' });
  });

  it('works with only to recipients (no cc/bcc)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({
      to: ['alice@example.com'],
      subject: 'Test',
      textBody: 'Hello',
    });

    const req = makeReq.mock.calls[0].arguments[0];
    const rcptTo = req.methodCalls[1][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 1);
    assert.deepEqual(rcptTo[0], { email: 'alice@example.com' });
  });
});

// ---------- recipient name parsing ----------

describe('recipient name parsing', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getMailboxes', async () => [
      DRAFTS_MAILBOX,
      { id: 'mb-sent', name: 'Sent', role: 'sent' },
    ]);
  });

  it('createDraft parses "Name <email>" recipients into { name, email }', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'd-1' } } }, 'createDraft']],
    }));

    await client.createDraft({ subject: 'Hi', to: ['Alice <a@x.com>'] });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ name: 'Alice', email: 'a@x.com' }]);
  });

  it('updateDraft parses "Name <email>" recipients into { name, email }', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'd-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { to: ['Alice <a@x.com>'] });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ name: 'Alice', email: 'a@x.com' }]);
  });

  it('sendEmail parses "Name <email>" recipients in the email object', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'e-1' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({ to: ['Alice <a@x.com>'], subject: 'Hi', textBody: 'Hello' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ name: 'Alice', email: 'a@x.com' }]);
  });

  it('sendEmail keeps the SMTP envelope rcptTo as a bare address (strips display name)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'e-1' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({ to: ['Alice <a@x.com>'], subject: 'Hi', textBody: 'Hello' });

    const rcptTo = makeReq.mock.calls[0].arguments[0].methodCalls[1][1].create.submission.envelope.rcptTo;
    assert.deepEqual(rcptTo, [{ email: 'a@x.com' }]);
  });
});

// ---------- createDraft replyTo ----------

describe('createDraft replyTo', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('includes replyTo in created email object when provided', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-draft' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({
      subject: 'Draft with replyTo',
      replyTo: ['noreply@example.com'],
    });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'noreply@example.com' }]);
  });
});

// ---------- updateDraft replyTo ----------

describe('updateDraft replyTo', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('overrides existing replyTo when provided in updates', async () => {
    const existingWithReplyTo = {
      ...EXISTING_DRAFT,
      replyTo: [{ email: 'old-reply@example.com' }],
    };

    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWithReplyTo] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-new' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { replyTo: ['new-reply@example.com'] });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'new-reply@example.com' }]);
  });

  it('preserves existing replyTo when not provided in updates', async () => {
    const existingWithReplyTo = {
      ...EXISTING_DRAFT,
      replyTo: [{ email: 'keep-me@example.com' }],
    };

    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWithReplyTo] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-new' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Updated subject only' });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'keep-me@example.com' }]);
  });
});

// ---------- wildcard identity ----------

const WILDCARD_IDENTITY = { id: 'id-wild', name: 'Jonathan Godley', email: '*@example.com', mayDelete: true };

describe('sendEmail wildcard identity', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [WILDCARD_IDENTITY]);
    mock.method(client, 'getMailboxes', async () => [
      DRAFTS_MAILBOX,
      { id: 'mb-sent', name: 'Sent', role: 'sent' },
    ]);
  });

  it('uses concrete from address in email and envelope, not wildcard literal', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));

    await client.sendEmail({
      to: ['bob@example.com'],
      subject: 'Test',
      textBody: 'Hello',
      from: 'work@example.com',
    });

    const req = makeReq.mock.calls[0].arguments[0];
    const emailObj = req.methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Jonathan Godley', email: 'work@example.com' }]);
    const envelope = req.methodCalls[1][1].create.submission.envelope;
    assert.deepEqual(envelope.mailFrom, { email: 'work@example.com' });
  });
});

describe('sendDraft wildcard identity', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [WILDCARD_IDENTITY]);
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, SENT_MAILBOX]);
  });

  it('matches wildcard identity when draft has concrete from address', async () => {
    const wildcardDraft = { ...SENDABLE_DRAFT, from: [{ email: 'work@example.com' }] };
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [wildcardDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const subId = await client.sendDraft('draft-1');
    assert.equal(subId, 'sub-1');

    const submitCall = makeReq.mock.calls[1].arguments[0];
    assert.equal(submitCall.methodCalls[0][1].create.submission.identityId, WILDCARD_IDENTITY.id);
    assert.deepEqual(submitCall.methodCalls[0][1].create.submission.envelope.mailFrom, { email: 'work@example.com' });
  });
});

describe('updateDraft wildcard identity', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [WILDCARD_IDENTITY]);
  });

  it('uses concrete from address when updating with wildcard identity', async () => {
    const existingWild = { ...EXISTING_DRAFT, from: [{ email: 'old@example.com' }] };
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWild] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { from: 'new@example.com' });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Jonathan Godley', email: 'new@example.com' }]);
  });

  it('preserves concrete from when updating without changing from', async () => {
    const existingWild = { ...EXISTING_DRAFT, from: [{ email: 'work@example.com' }] };
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWild] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Changed subject only' });

    const emailObj = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Jonathan Godley', email: 'work@example.com' }]);
  });
});

// ---------- version sync ----------

describe('version sync', () => {
  it('package.json, manifest.json, and index.ts all have the same version', async () => {
    const { readFileSync } = await import('fs');
    const { resolve: r } = await import('path');
    const root = r(import.meta.dirname, '..');

    const pkg = JSON.parse(readFileSync(r(root, 'package.json'), 'utf8'));
    const manifest = JSON.parse(readFileSync(r(root, 'manifest.json'), 'utf8'));
    const indexSrc = readFileSync(r(root, 'src', 'index.ts'), 'utf8');

    const indexMatch = indexSrc.match(/version:\s*'([^']+)'/);
    assert.ok(indexMatch, 'Could not find version string in index.ts');

    assert.equal(pkg.version, manifest.version, 'package.json and manifest.json versions must match');
    assert.equal(pkg.version, indexMatch[1], 'package.json and index.ts versions must match');
  });
});

// ---------- outgoing attachments: uploadBlob ----------

describe('uploadBlob', () => {
  function clientWithUpload(): JmapClient {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    const client = new JmapClient(auth);
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
      uploadUrl: 'https://up.example.com/upload/{accountId}/',
    }));
    return client;
  }

  it('POSTs raw bytes to the {accountId}-substituted uploadUrl with an explicit (non-json) Content-Type', async (t) => {
    const client = clientWithUpload();
    let captured: any;
    t.mock.method(globalThis, 'fetch', async (url: any, init: any) => {
      captured = { url, init };
      return { ok: true, json: async () => ({ accountId: ACCOUNT_ID, blobId: 'blob-9', type: 'application/pdf', size: 3 }) } as any;
    });

    const result = await client.uploadBlob(Buffer.from([1, 2, 3]), 'application/pdf');

    assert.equal(captured.url, `https://up.example.com/upload/${ACCOUNT_ID}/`);
    assert.equal(captured.init.method, 'POST');
    assert.equal(captured.init.headers['Content-Type'], 'application/pdf');
    assert.notEqual(captured.init.headers['Content-Type'], 'application/json');
    assert.equal(captured.init.headers['Authorization'], 'Bearer fake-token');
    // Body is the raw bytes (a BufferSource view), never a JSON string.
    assert.equal(typeof captured.init.body === 'string', false);
    assert.equal(captured.init.body.byteLength, 3);
    assert.equal(result.blobId, 'blob-9');
    assert.equal(result.type, 'application/pdf');
  });

  it('uses the server-returned type as authoritative even if it differs from what we sent', async (t) => {
    const client = clientWithUpload();
    t.mock.method(globalThis, 'fetch', async () => ({
      ok: true, json: async () => ({ accountId: ACCOUNT_ID, blobId: 'b', type: 'image/png', size: 1 }),
    }) as any);
    const result = await client.uploadBlob(Buffer.from([0]), 'application/octet-stream');
    assert.equal(result.type, 'image/png');
  });

  it('throws when the session has no uploadUrl', async () => {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    const client = new JmapClient(auth);
    mock.method(client, 'getSession', async () => ({ apiUrl: 'x', accountId: ACCOUNT_ID, capabilities: {} }));
    await assert.rejects(() => client.uploadBlob(Buffer.from([1]), 'text/plain'), /Upload capability not available/);
  });

  it('throws when the server response carries no blobId', async (t) => {
    const client = clientWithUpload();
    t.mock.method(globalThis, 'fetch', async () => ({ ok: true, json: async () => ({}) }) as any);
    await assert.rejects(() => client.uploadBlob(Buffer.from([1]), 'text/plain'), /no blobId/);
  });
});

// ---------- outgoing attachments: safeReadPath (read confinement) ----------

describe('safeReadPath (read confinement)', () => {
  it('throws the opt-in error (no fs touch) when attachDir is undefined', async () => {
    await assert.rejects(
      () => JmapClient.safeReadPath('anything.pdf', undefined),
      /FASTMAIL_ATTACH_DIR/,
    );
  });

  it('returns a usable handle for a regular file inside the root', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      await fsWriteFile(join(root, 'doc.pdf'), 'hello');
      const { handle, size } = await JmapClient.safeReadPath('doc.pdf', root);
      try {
        assert.equal(size, 5);
        const buf = Buffer.alloc(size);
        await handle.read(buf, 0, size, 0);
        assert.equal(buf.toString(), 'hello');
      } finally {
        await handle.close();
      }
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects a missing file', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      await assert.rejects(() => JmapClient.safeReadPath('nope.pdf', root), /File not found/);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects a directory (not a regular file)', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      await fsMkdir(join(root, 'subdir'));
      // Any rejection is acceptable (the message varies by platform) — it must NOT resolve.
      await assert.rejects(() => JmapClient.safeReadPath('subdir', root));
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects a path that escapes the root via ..', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      await assert.rejects(() => JmapClient.safeReadPath('../escape.txt', root), /must be within/);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects a symlink whose target escapes the root', async (t) => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    const outsideDir = await mkdtemp(join(tmpdir(), 'fastmail-mcp-out-'));
    try {
      const outside = join(outsideDir, 'secret.txt');
      await fsWriteFile(outside, 'secret');
      try {
        await symlink(outside, join(root, 'link.txt'));
      } catch (err) {
        if ((err as any)?.code === 'EPERM' || (err as any)?.code === 'EACCES') {
          t.skip('symlink creation not permitted on this platform');
          return;
        }
        throw err;
      }
      await assert.rejects(
        () => JmapClient.safeReadPath('link.txt', root),
        /outside the allowed directory|symlink escape/i,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
      await rm(outsideDir, { recursive: true, force: true });
    }
  });

  it('rejects the Windows escape shapes (UNC, device namespace, ADS, drive-relative, short name)', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      const bad = ['\\\\server\\share\\f.txt', '\\\\?\\C:\\f.txt', 'doc.pdf:stream', 'C:relative', 'PROGRA~1\\f.txt'];
      for (const input of bad) {
        await assert.rejects(() => JmapClient.safeReadPath(input, root), /not allowed|drive-relative/);
      }
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('does NOT reject a legitimate filename that merely contains a tilde (only 8.3 ~digit forms)', async () => {
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-read-'));
    try {
      await fsWriteFile(join(root, 'report~final.txt'), 'ok'); // tilde + letter, not an 8.3 short name
      const { handle } = await JmapClient.safeReadPath('report~final.txt', root);
      await handle.close();
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

// ---------- outgoing attachments: uploadAttachments (read + cap + part shape) ----------

describe('uploadAttachments', () => {
  function clientWithUpload(): JmapClient {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    const client = new JmapClient(auth);
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
      uploadUrl: 'https://up.example.com/upload/{accountId}/',
    }));
    return client;
  }

  it('throws the opt-in error when attachDir is undefined', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ path: 'x.pdf' }], undefined),
      /FASTMAIL_ATTACH_DIR/,
    );
  });

  it('builds a fresh 4-key part from the server type, defaulting name to the basename', async (t) => {
    const client = clientWithUpload();
    t.mock.method(globalThis, 'fetch', async () => ({
      ok: true, json: async () => ({ accountId: ACCOUNT_ID, blobId: 'blob-up', type: 'application/pdf', size: 5 }),
    }) as any);
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'report.pdf'), 'hello');
      const parts = await client.uploadAttachments([{ path: 'report.pdf' }], root);
      assert.deepEqual(parts, [
        { blobId: 'blob-up', type: 'application/pdf', name: 'report.pdf', disposition: 'attachment' },
      ]);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('rejects an invalid caller contentType before any read/upload', async () => {
    const client = clientWithUpload();
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'f.bin'), 'x');
      await assert.rejects(
        () => client.uploadAttachments([{ path: 'f.bin', contentType: 'not a mime type' }], root),
        /invalid contentType/,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('uploads multiple files and returns a part per file (two-pass, in order)', async (t) => {
    const client = clientWithUpload();
    let n = 0;
    t.mock.method(client, 'uploadBlob', async (data: Buffer, ct: string) => ({ blobId: 'blob-' + (++n), type: ct, size: data.length }));
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'a.txt'), 'aa');
      await fsWriteFile(join(root, 'b.txt'), 'bbb');
      const parts = await client.uploadAttachments([{ path: 'a.txt' }, { path: 'b.txt', contentType: 'text/plain' }], root);
      assert.deepEqual(parts.map(p => p.name), ['a.txt', 'b.txt']);
      assert.deepEqual(parts.map(p => p.blobId), ['blob-1', 'blob-2']);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('validates every file before uploading any — a later bad path uploads zero blobs (no orphans)', async (t) => {
    const client = clientWithUpload();
    let uploads = 0;
    t.mock.method(client, 'uploadBlob', async () => { uploads++; return { blobId: 'x', type: 'text/plain', size: 1 }; });
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'good.txt'), 'ok');
      await assert.rejects(
        // good.txt validates+opens in pass 1, then the escaping path rejects — pass 2 never runs.
        () => client.uploadAttachments([{ path: 'good.txt' }, { path: '../escape.txt' }], root),
        /must be within/,
      );
      assert.equal(uploads, 0);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

// ---------- outgoing attachments wired into sendEmail / createDraft ----------

describe('attachments on send/create', () => {
  it('sendEmail places attachment parts in the email object (reply send branch)', async () => {
    const client = makeClient();
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, { id: 'mb-sent', name: 'Sent', role: 'sent' }]);
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-new' } } }, 'createEmail'],
        ['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitEmail'],
      ],
    }));
    const part = { blobId: 'b1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' };
    await client.sendEmail({ to: ['bob@example.com'], subject: 'T', textBody: 'Hi', attachments: [part] });
    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.attachments, [part]);
  });

  it('createDraft places attachment parts in the email object (reply draft branch)', async () => {
    const client = makeClient();
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-42' } } }, 'createDraft']],
    }));
    const part = { blobId: 'b2', type: 'image/png', name: 'p.png', disposition: 'attachment' };
    await client.createDraft({ subject: 'Hello', attachments: [part] });
    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.attachments, [part]);
  });

  it('createDraft accepts an attachment-only draft (no to/subject/body) — attachments count as content', async () => {
    const client = makeClient();
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-43' } } }, 'createDraft']],
    }));
    const part = { blobId: 'b3', type: 'application/pdf', name: 'only.pdf', disposition: 'attachment' };
    const id = await client.createDraft({ attachments: [part] }); // must NOT throw the contentless guard
    assert.equal(id, 'email-43');
    assert.deepEqual(makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft.attachments, [part]);
  });

  it('createDraft still rejects a truly empty draft (no fields, no attachments)', async () => {
    const client = makeClient();
    await assert.rejects(
      () => client.createDraft({}),
      /At least one of to, subject, textBody, htmlBody, or attachments/,
    );
  });
});

// ---------- outgoing attachments: updateDraft append / remove / clear-all ----------

describe('updateDraft attachments', () => {
  let client: JmapClient;
  beforeEach(() => { client = makeClient(); });

  const NEW_PART = { blobId: 'new-blob', type: 'application/pdf', name: 'new.pdf', disposition: 'attachment' };

  // One carried PDF attachment (blob-att / doc.pdf), no inline parts.
  const DRAFT_ONE_ATT = {
    ...EXISTING_DRAFT,
    attachments: [{ blobId: 'blob-att', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment', cid: null, partId: '3', size: 1234 }],
  };

  it('appends new parts, keeping carried attachments', async () => {
    const makeReq = mockUpdate(client, DRAFT_ONE_ATT);
    await client.updateDraft('draft-1', { attachments: [NEW_PART] });
    assert.deepEqual(draftFromCall(makeReq).attachments, [
      { blobId: 'blob-att', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
      NEW_PART,
    ]);
  });

  it('removes a carried attachment by blobId', async () => {
    const makeReq = mockUpdate(client, DRAFT_ONE_ATT);
    await client.updateDraft('draft-1', { removeAttachments: ['blob-att'] });
    assert.equal(draftFromCall(makeReq).attachments, undefined);
  });

  it('removes a carried attachment by a unique name', async () => {
    const makeReq = mockUpdate(client, DRAFT_ONE_ATT);
    await client.updateDraft('draft-1', { removeAttachments: ['doc.pdf'] });
    assert.equal(draftFromCall(makeReq).attachments, undefined);
  });

  it('rejects a remove ref that matches nothing', async () => {
    mockUpdate(client, DRAFT_ONE_ATT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { removeAttachments: ['nonexistent'] }),
      /matched no attachment/,
    );
  });

  it('rejects an ambiguous remove-by-name (more than one match)', async () => {
    const dupDraft = {
      ...EXISTING_DRAFT,
      attachments: [
        { blobId: 'b1', type: 'application/pdf', name: 'dup.pdf', disposition: 'attachment', cid: null, partId: '3', size: 1 },
        { blobId: 'b2', type: 'application/pdf', name: 'dup.pdf', disposition: 'attachment', cid: null, partId: '4', size: 2 },
      ],
    };
    mockUpdate(client, dupDraft);
    await assert.rejects(
      () => client.updateDraft('draft-1', { removeAttachments: ['dup.pdf'] }),
      /matches 2 attachments by name/,
    );
  });

  it('never matches a null-named attachment by name (and removes the named one only)', async () => {
    const mixedDraft = {
      ...EXISTING_DRAFT,
      attachments: [
        { blobId: 'b-null', type: 'application/octet-stream', name: null, disposition: 'attachment', cid: null, partId: '3', size: 1 },
        { blobId: 'b-real', type: 'application/pdf', name: 'real.pdf', disposition: 'attachment', cid: null, partId: '4', size: 2 },
      ],
    };
    const makeReq = mockUpdate(client, mixedDraft);
    await client.updateDraft('draft-1', { removeAttachments: ['real.pdf'] });
    // The null-named one survives; the named one is gone.
    assert.deepEqual(draftFromCall(makeReq).attachments, [
      { blobId: 'b-null', type: 'application/octet-stream', disposition: 'attachment' },
    ]);
  });

  it('clears all attachments on clearFields:["attachments"]', async () => {
    const makeReq = mockUpdate(client, DRAFT_ONE_ATT);
    await client.updateDraft('draft-1', { clearFields: ['attachments'] });
    assert.equal(draftFromCall(makeReq).attachments, undefined);
  });

  it('rejects attachments + clearFields:["attachments"] together (conflict)', async () => {
    mockUpdate(client, DRAFT_ONE_ATT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { attachments: [NEW_PART], clearFields: ['attachments'] }),
      /cannot both set and clear attachments/,
    );
  });

  it('leaves a body-less draft body-invariant on an attachment-only edit (no no-body throw)', async () => {
    const bodyless = { ...EXISTING_DRAFT, textBody: null, htmlBody: null, bodyValues: {} };
    const makeReq = mockUpdate(client, bodyless);
    const result = await client.updateDraft('draft-1', { attachments: [NEW_PART] });
    assert.equal(result.id, 'draft-2');
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.attachments, [NEW_PART]);
    assert.equal(draft.textBody, undefined);
    assert.equal(draft.htmlBody, undefined);
  });

  it('still throws the no-body error when the last body is cleared alongside an attachments change', async () => {
    mockUpdate(client, DRAFT_ONE_ATT); // text-only draft with one attachment
    await assert.rejects(
      () => client.updateDraft('draft-1', { attachments: [NEW_PART], clearFields: ['textBody'] }),
      /a draft needs a body/,
    );
  });
});
