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

  // 9. Custom mailboxId used instead of auto-lookup
  it('uses provided mailboxId without looking up mailboxes', async () => {
    const getMailboxes = client.getMailboxes as ReturnType<typeof mock.method>;

    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-9' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Custom', mailboxId: 'mb-custom' });

    // getMailboxes should not have been called
    assert.equal(getMailboxes.mock.calls.length, 0);

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj.mailboxIds['mb-custom'], true);
  });

  // 10. HTML body constructed correctly
  it('constructs HTML body parts correctly', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-10' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Rich', htmlBody: '<p>Hello</p>' });

    const emailObj = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.htmlBody, [{ partId: 'html', type: 'text/html' }]);
    assert.equal(emailObj.textBody, undefined);
    assert.deepEqual(emailObj.bodyValues, { html: { value: '<p>Hello</p>' } });
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

  // ---- cross-format coupling guard (option D) ----

  it('throws when writing textBody alone on a dual-body draft (would discard htmlBody)', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'NEW text' }),
      /htmlBody.*Supply htmlBody as well.*clearFields/s,
    );
  });

  it('throws when writing htmlBody alone on a dual-body draft (parity, would discard textBody)', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>' }),
      /textBody.*Supply textBody as well.*clearFields/s,
    );
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

  it('throws when adding htmlBody alone to a text-only draft (parity: would discard textBody)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>' }),
      /textBody.*Supply textBody as well.*clearFields/s,
    );
  });

  it('adds htmlBody to a text-only draft when textBody is cleared (html-only result)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>', clearFields: ['textBody'] });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.textBody, undefined);
    assert.deepEqual(draft.bodyValues, { html: { value: '<p>NEW</p>' } });
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

  it('clears textBody via clearFields and preserves htmlBody', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['textBody'] });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.textBody, undefined);
    assert.deepEqual(draft.bodyValues, { html: { value: '<p>The html</p>' } });
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
});

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

  // ---- Change 2: reject an empty body part on send ----
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
    // getEmails uses getListResult(response, 1) but only 1 response exists
    await assert.rejects(
      () => client.getEmails(undefined, 10),
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

  it('returns email list on success', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
      ],
    });
    const results = await client.searchEmails('test', 10);
    assert.equal(results.items.length, 1);
    assert.equal(results.items[0].subject, 'Test');
  });

  it('returns empty array when no results', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    });
    const results = await client.searchEmails('nonexistent');
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
      () => client.searchEmails('test'),
      (err: Error) => {
        assert.match(err.message, /invalidArguments/);
        return true;
      },
    );
  });

  it('excludeDrafts pushes notKeyword $draft into the server-side filter', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails('quarterly', 10, false, true);
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.text, 'quarterly');
    assert.equal(filter.notKeyword, '$draft');
  });

  it('includes drafts by default (no notKeyword in the filter)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails('quarterly', 10);
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

// ---------- recipient name parsing (B3) ----------

describe('recipient name parsing (B3)', () => {
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
