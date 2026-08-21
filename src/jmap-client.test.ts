import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { homedir } from 'os';
import { resolve, join, basename, sep } from 'path';
import { JmapClient } from './jmap-client.js';
import type { JmapRequest } from './jmap-client.js';
import { FastmailAuth } from './auth.js';
import { InvalidInputError, PathAccessError } from './coerce.js';
import { callArguments, findCallArguments } from './testing/mock-calls.js';

// ---------- helpers ----------

/**
 * Install a stub over the client's single outbound seam and hand back the mock.
 *
 * The implementation parameter is declared with the request even where a caller's
 * stub ignores it. node:test types the mock it returns from the real method AND the
 * implementation, so an implementation taking no arguments records its arguments as
 * a union with the empty tuple — and every assertion that reads the request back is
 * then reading something the compiler must treat as possibly undefined.
 */
function stubRequests(client: JmapClient, impl: (request: JmapRequest) => Promise<any>) {
  return mock.method(client, 'makeRequest', impl);
}

/** A mock installed over `makeRequest`, as the read helpers below see it. */
type RequestMock = ReturnType<typeof stubRequests>;

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
  stubRequests(client, async () => response);
}

// ---------- tests ----------

// Every throwing set-error site routes through throwSingleSetError, so the same JMAP
// reason means the same MCP error code wherever it surfaces. The draft-lifecycle
// notCreated throws were the exception for a while: a create rejected for
// invalidProperties reported a server bug while the identical failure on a contact
// reported a caller-fixable one. The behavioural cases below pin the create path; the
// source check pins the other two, which need a whole draft lifecycle to reach and would
// otherwise be guarded by nothing.
describe('draft-lifecycle set errors are classified, not just described', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  function stubCreateFailure(setError: { type: string; description?: string }) {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notCreated: { draft: setError } }, 'createDraft'],
      ],
    });
  }

  it('reports a rejected property on create as the caller\'s to fix', async () => {
    stubCreateFailure({ type: 'invalidProperties', description: 'to' });

    await assert.rejects(
      () => client.createDraft({ subject: 'Hello' }),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.equal(err.message, 'Failed to create draft: invalidProperties - to');
        return true;
      },
    );
  });

  it('keeps a create refused on quota a server-side failure', async () => {
    stubCreateFailure({ type: 'overQuota' });

    await assert.rejects(
      () => client.createDraft({ subject: 'Hello' }),
      (err: Error) => {
        assert.ok(!(err instanceof InvalidInputError));
        assert.equal(err.message, 'Failed to create draft: overQuota');
        return true;
      },
    );
  });

  it('builds no set-error message outside the classifier', async () => {
    const { readFileSync } = await import('fs');
    const source = readFileSync(new URL('./jmap-client.ts', import.meta.url), 'utf-8');

    // describeSetError formats the reason; throwSingleSetError formats AND classifies it.
    // A throw that reaches for the formatter directly has skipped the classification, so
    // it always surfaces as InternalError regardless of what the server actually said.
    const unclassified = source
      .split('\n')
      .map((line, i) => ({ line, lineNo: i + 1 }))
      .filter(({ line }) => /throw new Error\(.*describeSetError\(/.test(line))
      .map(({ line, lineNo }) => `${lineNo}: ${line.trim()}`);

    assert.deepEqual(
      unclassified,
      [],
      'these throws format a set error without classifying it; route them through throwSingleSetError',
    );
  });
});

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
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-1' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Test', textBody: 'body' });

    assert.equal(makeReq.mock.calls.length, 1);
    const request = callArguments(makeReq)[0];

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
        // the rejected property came from this call, so the caller can fix it by
        // re-forming the request → InvalidInputError → InvalidParams
        assert.equal(err.name, 'InvalidInputError');
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
        // operational post-condition (not caller-fixable) stays a plain Error → InternalError;
        // pins the #41 classification boundary against over-migration.
        assert.equal(err.name, 'Error');
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

    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-7' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Hi', from: 'alias@example.com' });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
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

    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-wild' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Hi', from: 'work@example.com' });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
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

  // 8e. Composite/injection from-string rejected even though it ends in the wildcard domain
  it('rejects a composite from address against a wildcard identity', async () => {
    const wildcardIdentity = { id: 'id-wild', email: '*@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [wildcardIdentity]);

    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: 'attacker@evil.com,me@example.com' }),
      (err: Error) => {
        assert.match(err.message, /not verified/i);
        return true;
      },
    );
  });

  it('rejects a from address with CR/LF against a wildcard identity', async () => {
    const wildcardIdentity = { id: 'id-wild', email: '*@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [wildcardIdentity]);

    await assert.rejects(
      () => client.createDraft({ subject: 'Hi', from: 'a@example.com\r\nBCC:x@example.com' }),
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

    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-9' } } }, 'createDraft'],
      ],
    }));

    // Resolve by name -> the custom mailbox's id.
    await client.createDraft({ subject: 'Custom', mailbox: 'Project X' });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
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
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-10' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({ subject: 'Rich', htmlBody: '<p>Hello</p>' });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
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

// The mailbox list updateDraft reads to find where the replaced draft should go. (The
// makeClient default deliberately has no trash role — some exclusion tests depend on that —
// so the disposal tests supply their own list.)
const MAILBOXES_WITH_TRASH = [
  DRAFTS_MAILBOX,
  { id: 'mb-trash', name: 'Trash', role: 'trash' },
];

// Wire makeRequest for create-then-dispose: Email/get returns the fixture; the create-only
// Email/set returns a created id; the update-only Email/set moves the old draft to Trash.
// Returns the makeRequest mock.
function mockUpdate(client: JmapClient, fixture: any, mailboxes: any[] = MAILBOXES_WITH_TRASH) {
  mock.method(client, 'getMailboxes', async () => mailboxes);
  return stubRequests(client, async (req: any) => {
    const [method, params] = req.methodCalls[0];
    if (method === 'Email/get') {
      return { methodResponses: [['Email/get', { list: [fixture] }, 'getEmail']] };
    }
    // Email/set — create-then-dispose issues a create-only call, then an update-only call.
    if (params.create) {
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
    }
    return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
  });
}

// Pull the recreated draft object out of the create (second overall) call.
function draftFromCall(makeReq: RequestMock) {
  return callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
}

describe('updateDraft', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns new email ID on success (create first, then dispose of the old draft)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.trashedOldDraftId, 'draft-1');
    assert.equal(result.orphanedOldDraftId, undefined);

    // Three calls: Email/get, a create-ONLY Email/set, then the Trash move.
    assert.equal(makeReq.mock.calls.length, 3);
    const createCall = callArguments(makeReq, 1)[0].methodCalls[0];
    assert.equal(createCall[0], 'Email/set');
    assert.equal(createCall[1].destroy, undefined); // create call must NOT also destroy
    assert.equal(createCall[1].update, undefined);  // nor touch the old draft
    assert.equal(createCall[1].create.draft.subject, 'New Subject');
    const disposeCall = callArguments(makeReq, 2)[0].methodCalls[0];
    assert.equal(disposeCall[0], 'Email/set');
    assert.equal(disposeCall[1].create, undefined); // dispose call must NOT also create
  });

  // ---- disposal of the replaced draft: Trash, never destroy (#65) ----

  it('moves the replaced draft to Trash and never issues a destroy', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(result.trashedOldDraftId, 'draft-1');

    const disposeParams = callArguments(makeReq, 2)[0].methodCalls[0][1];
    // mailboxIds only: keywords (incl. $draft) are left untouched, so the trashed copy is
    // still a draft and can be moved back to Drafts.
    assert.deepEqual(disposeParams.update, { 'draft-1': { mailboxIds: { 'mb-trash': true } } });
    assert.equal(disposeParams.update['draft-1'].keywords, undefined);
    // No call anywhere in the exchange may destroy the replaced draft.
    for (const call of makeReq.mock.calls) {
      assert.equal(call.arguments[0].methodCalls[0][1].destroy, undefined);
    }
  });

  it('resolves Trash by exact role, not by a folder merely named like it', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT, [
      { id: 'mb-drafts', name: 'Drafts', role: 'drafts' },
      { id: 'mb-fake', name: 'Trash bin rules', role: null },
      { id: 'mb-trash', name: 'Deleted Items', role: 'Trash' }, // role casing is normalised
    ]);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.equal(result.trashedOldDraftId, 'draft-1');
    const disposeParams = callArguments(makeReq, 2)[0].methodCalls[0][1];
    assert.deepEqual(disposeParams.update, { 'draft-1': { mailboxIds: { 'mb-trash': true } } });
  });

  it('with no Trash mailbox: leaves the old draft in place with a reason, does NOT destroy it', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT, [DRAFTS_MAILBOX]);

    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.trashedOldDraftId, undefined);
    assert.equal(result.orphanedOldDraftId, 'draft-1');
    assert.match(result.orphanedOldDraftReason!, /trash role/i);
    // Email/get + create, and then nothing: no fallback destroy.
    assert.equal(makeReq.mock.calls.length, 2);
  });

  it('on a failed Trash move (notUpdated): surfaces the orphan + reason, does NOT throw', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { notUpdated: { 'draft-1': { type: 'serverFail', description: 'busy' } } }, 'trashOldDraft']] };
    });

    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.trashedOldDraftId, undefined);
    assert.equal(result.orphanedOldDraftId, 'draft-1');
    assert.match(result.orphanedOldDraftReason!, /serverFail|busy/);
  });

  it('when the server reports the move in neither updated nor notUpdated: orphaned, not assumed trashed', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      // Neither map mentions draft-1: claiming "recoverable in Trash" here would be a
      // guess about a draft that may still be sitting in Drafts.
      return { methodResponses: [['Email/set', { updated: {}, notUpdated: {} }, 'trashOldDraft']] };
    });

    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.trashedOldDraftId, undefined);
    assert.equal(result.orphanedOldDraftId, 'draft-1');
    assert.match(result.orphanedOldDraftReason!, /did not report the move/i);
  });

  it('treats a null entry in updated as a successful move (JMAP allows id -> null)', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });

    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.trashedOldDraftId, 'draft-1');
    assert.equal(result.orphanedOldDraftId, undefined);
  });

  it('when the mailbox lookup throws: surfaces the orphan + reason, does NOT throw', async () => {
    mock.method(client, 'getMailboxes', async () => { throw new Error('network down'); });
    stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });

    const result = await client.updateDraft('draft-1', { subject: 'X' });
    assert.equal(result.id, 'draft-2');
    assert.equal(result.orphanedOldDraftId, 'draft-1');
    assert.match(result.orphanedOldDraftReason!, /network down/);
  });

  // ---- echo-back of the replaced draft (staleness detection, #65) ----

  it('echoes back what the replaced draft contained', async () => {
    mockUpdate(client, RICH_DRAFT);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.deepEqual(result.replacedDraft, {
      id: 'draft-1',
      subject: 'Old Subject',              // the PRE-edit subject, not the new one
      to: ['bob@example.com'],
      cc: ['carol@example.com'],
      textBodySize: 'The text'.length,
      htmlBodySize: '<p>The html</p>'.length,
    });
  });

  it('omits echo-back fields the replaced draft did not have', async () => {
    const bare = { ...EXISTING_DRAFT, subject: undefined, cc: [] };
    mockUpdate(client, bare);

    const result = await client.updateDraft('draft-1', { subject: 'New Subject' });
    assert.deepEqual(result.replacedDraft, {
      id: 'draft-1',
      to: ['bob@example.com'],
      textBodySize: 'Old body'.length,     // html was never present on this draft
    });
  });

  it('merges fields — preserves existing values for unspecified fields', async () => {
    const makeReq = stubRequests(client, async (req) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Updated' });

    // The create call should keep existing to address
    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ email: 'bob@example.com' }]);
    assert.equal(emailObj.subject, 'Updated');
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...EXISTING_DRAFT, keywords: { $seen: true } };
    stubRequests(client, async () => ({
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
    stubRequests(client, async () => ({
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
    stubRequests(client, async (req: any) => {
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
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [aliasedDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'New subject' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
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
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [dualDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'New subject' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.bodyValues, {
      text: { value: 'The text' },
      html: { value: '<p>The html</p>' },
    });
  });

  // ---- malformed caller bodies (#62, #71/#77, #78) ----

  it('rejects a non-string body without touching the server', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: 42 as any }),
      (err: Error) => {
        assert.match(err.message, /htmlBody must be a string; received number/);
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 0); // fails before the Email/get fetch
  });

  it('rejects an entirely HTML-escaped htmlBody', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '&lt;p&gt;Rewritten&lt;/p&gt;' }),
      /htmlBody appears to be HTML-escaped/,
    );
  });

  it('rejects a CDATA-wrapped body in either format', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<![CDATA[<p>Rewritten</p>]]>' }),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: '<![CDATA[Rewritten]]>' }),
      /textBody is wrapped in a CDATA section/,
    );
  });

  it('still accepts a body that merely contains escaped markup inside real tags', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>Write <code>&lt;p&gt;</code> like this</p>' });
    assert.equal(draftFromCall(makeReq).bodyValues.html.value, '<p>Write <code>&lt;p&gt;</code> like this</p>');
  });

  // ---- one-sided guard + text-fallback regeneration on html edit ----

  it('throws when editing textBody alone on a dual-body draft (html is what recipients see)', async () => {
    mockUpdate(client, RICH_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'NEW text' }),
      (err: Error) => {
        assert.match(err.message, /editing textBody alone.*edit htmlBody.*clearFields:\['htmlBody'\]/s);
        // body-coupling reject is caller-fixable input → InvalidInputError → InvalidParams (#41)
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
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
      (err: Error) => {
        assert.match(err.message, /no readable body/);
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
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
      () => client.updateDraft('draft-1', { cc: ['x@y.example'], clearFields: ['cc'] }),
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

  it('carries an embedded (cid:) image through a metadata edit unchanged', async () => {
    const inlineDraft = {
      ...EXISTING_DRAFT,
      attachments: [{ blobId: 'blob-img', type: 'image/png', disposition: 'inline', cid: 'img@x', name: null, partId: '2', size: 70 }],
    };
    const makeReq = mockUpdate(client, inlineDraft);
    await client.updateDraft('draft-1', { subject: 'X' });
    assert.deepEqual(draftFromCall(makeReq).attachments, [
      { blobId: 'blob-img', type: 'image/png', disposition: 'inline', cid: 'img@x' },
    ]);
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

  it('rejects editing a draft whose body carries a part the recreate cannot reproduce', async () => {
    const weirdDraft = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/calendar' }],
      htmlBody: null,
      bodyValues: { '1': { value: 'BEGIN:VCALENDAR' } },
    };
    mockUpdate(client, weirdDraft);
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /body contains a part this server cannot carry.*text\/calendar.*Recreate it/s,
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

  // ---- create-before-dispose ordering (data-loss prevention) ----

  it('on create failure: throws, disposes of nothing, leaves the old draft untouched', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    const makeReq = stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      if (params.create) {
        return { methodResponses: [['Email/set', { notCreated: { draft: { type: 'invalidProperties', description: 'bad blob' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /Failed to create updated draft.*invalidProperties/s,
    );
    // Exactly 2 calls: Email/get + the failed create. The old draft must NEVER be moved.
    assert.equal(makeReq.mock.calls.length, 2);
    assert.ok(callArguments(makeReq, 1)[0].methodCalls[0][1].create);
  });

  it('on a transport error while disposing (after a good create): returns orphan warning, does NOT throw', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    stubRequests(client, async (req: any) => {
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
    assert.match(result.orphanedOldDraftReason!, /network down/);
  });

  // ---- reply-quote preservation on body edit (#37, redesigned #42) ----
  //
  // The guard decides on the EXISTING (stored) body, so these fixtures use the RAW body shapes
  // Fastmail returns for reply drafts — captured from a live store/fetch round-trip (2026-06-28)
  // and trimmed of the bulk quoted body but BYTE-EXACT in the marker region the guard reads.
  // Only that marker region (the attribution line and the "\n> " / "\n\n> " structure after it)
  // is byte-exact; the quoted lines themselves carry synthetic content, because the guard never
  // reads them. Keep it that way — a fixture must not reproduce a real message's contents.
  // Pinning to Fastmail's re-serialized shape (not our buildReplyBodies output) is the point:
  // an html-derived text fallback comes back as "wrote:\n\n> " (blank line). The coercion of
  // noQuote ("true"/"garbage") lives at the index.ts handler seam and is pinned by coerce.test.ts
  // (coerceBool) + the live harness; updateDraft only ever sees a real boolean, so it is not
  // re-tested here.
  const RAW_HTML_QUOTE = '<p>my reply</p><div><br></div><div>On Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:</div><blockquote type="cite" style="margin:0 0 0 .8ex;border-left:1px solid #ccc;padding-left:1ex">\n  1 new planning application near 1 Example Street\n</blockquote>';
  const RAW_TEXT_QUOTE = 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:\n> 2/2 Example St Sampleton NSW 2000: Change of Use and Fitout of a Studio\n> \n> Contact us if you have questions.';
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

  // An original whose whole body is an embedded image. Quotable — because the image can be
  // carried into the rebuilt quote — but only into an html body; there is no plain-text form
  // of a picture, so a text-only edit still finds nothing to restore.
  const IMAGE_ONLY_ORIGINAL = {
    id: 'orig-image',
    messageId: ['orig-msg@example.com'],
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [],
    htmlBody: [{ partId: 'oh', type: 'text/html' }],
    bodyValues: { oh: { value: '<div><img src="cid:pic-1"></div>' } },
    attachments: [{ partId: '2', blobId: 'blob-pic', type: 'image/png', size: 90, name: 'pic.png', disposition: 'inline', cid: 'pic-1' }],
  };

  // An original whose body points two images at paths only its own sender's origin could
  // resolve. A quote is re-sent from a different message, so those references cannot come
  // with it — the quote ships without them, and the count is what says so.
  const RELATIVE_IMAGE_ORIGINAL = {
    id: 'orig-relative',
    messageId: ['orig-msg@example.com'],
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [],
    htmlBody: [{ partId: 'oh', type: 'text/html' }],
    bodyValues: {
      oh: { value: '<div>ORIGINAL HTML BODY<img src="/logo.png"><img src="//cdn.example.com/a.png"></div>' },
    },
    attachments: [],
  };

  // Dispatch Email/get BY ID — the chosen draft fixture for the draft id, the original fixture
  // for 'orig-1'. A single-fixture mock would make the regenerate test quote the DRAFT as its
  // own original and prove nothing, so id-dispatch is mandatory here. 'orig-missing' → notFound
  // (drives the not-found path). getEmailById issues Email/get + Mailbox/get; we answer only
  // Email/get (its mailbox read is defensive/optional).
  function mockReplyUpdate(c: JmapClient, draft: any = DUAL_REPLY) {
    mock.method(c, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        const id = params.ids?.[0];
        if (id === 'orig-1') return { methodResponses: [['Email/get', { list: [ORIGINAL_FOR_REPLY] }, 'email']] };
        if (id === 'orig-empty') return { methodResponses: [['Email/get', { list: [NONQUOTABLE_ORIGINAL] }, 'email']] };
        if (id === 'orig-image') return { methodResponses: [['Email/get', { list: [IMAGE_ONLY_ORIGINAL] }, 'email']] };
        if (id === 'orig-relative') return { methodResponses: [['Email/get', { list: [RELATIVE_IMAGE_ORIGINAL] }, 'email']] };
        if (id === 'orig-missing') return { methodResponses: [['Email/get', { list: [], notFound: ['orig-missing'] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [draft] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
  }

  // Find the create call by predicate — the regenerate path inserts a second Email/get, so
  // the create is no longer at a fixed index.
  function createdDraft(makeReq: RequestMock) {
    const [request] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][1].create,
      'creating the replacement draft',
    );
    return request.methodCalls[0][1].create.draft;
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
      (err: Error) => {
        assert.match(err.message, /originalEmailId 'orig-missing' could not be fetched/);
        // the #37 not-found originalEmailId reject is now InvalidInputError → InvalidParams (#41)
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
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

  it('keeps the quote of an original whose only content is an embedded image', async () => {
    // Such a message used to count as having nothing to quote, so this keep was refused. The
    // image is now carried into the rebuilt quote, which makes it real content.
    const makeReq = mockReplyUpdate(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>my edited reply</p>', originalEmailId: 'orig-image' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);
    assert.match(draft.bodyValues.html.value, /<img src="cid:ii-[0-9a-f]{32}@inline\.invalid"/);
    // The image itself rides the rebuilt draft under that same identifier.
    assert.equal(draft.attachments.some((a: any) => a.blobId === 'blob-pic' && a.disposition === 'inline'), true);
  });

  // A keep names an original this draft may never have quoted before, so the rebuilt quote can
  // lose an image for the first time here. That loss gets the same sentence a fresh reply gets.
  it('says how many images the rebuilt quote dropped for a reference form it cannot carry', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    const result = await client.updateDraft(
      'draft-1', { htmlBody: '<p>my edited reply</p>', originalEmailId: 'orig-relative' },
    );
    assert.deepEqual(result.notes, [
      '2 image(s) in the quoted message used a reference form this server cannot carry into' +
      ' a quote and were dropped; the rest of the quote was kept.',
    ]);
  });

  it('reports the same loss again on the next edit that rebuilds the same quote', async () => {
    // Each edit rebuilds the quote from the original, so each edit loses them again. Repeating
    // the disclosure is correct: the alternative is an edit that drops images and says nothing.
    mockReplyUpdate(client, DUAL_REPLY);
    const args = { htmlBody: '<p>edited twice</p>', originalEmailId: 'orig-relative' };
    const first = await client.updateDraft('draft-1', args);
    const second = await client.updateDraft('draft-1', args);
    assert.deepEqual(second.notes, first.notes);
  });

  it('says nothing of the sort when the rebuilt quote carries every reference it found', async () => {
    mockReplyUpdate(client, DUAL_REPLY);
    const result = await client.updateDraft(
      'draft-1', { htmlBody: '<p>my edited reply</p>', originalEmailId: 'orig-image' },
    );
    assert.equal(
      (result.notes ?? []).some((n) => /reference form/.test(n)), false,
    );
  });

  it('still refuses a TEXT-only keep of an image-only original — a picture has no plain-text quote', async () => {
    mockReplyUpdate(client, TEXT_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'my edited reply', originalEmailId: 'orig-image' }),
      /has no quotable content.*images.*plain-text body.*noQuote/s,
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
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const { submissionId } = await client.sendDraft('draft-1');
    assert.equal(submissionId, 'sub-1');

    // Verify submission call structure
    const submitCall = callArguments(makeReq, 1)[0];
    assert.equal(submitCall.methodCalls[0][0], 'EmailSubmission/set');
    assert.equal(submitCall.methodCalls[0][1].create.submission.emailId, 'draft-1');
    assert.equal(submitCall.methodCalls[0][1].create.submission.identityId, IDENTITY.id);

    // Verify envelope has all recipients (to + cc)
    const rcptTo = submitCall.methodCalls[0][1].create.submission.envelope.rcptTo;
    assert.equal(rcptTo.length, 2);
    assert.deepEqual(rcptTo[0], { email: 'bob@example.com' });
    assert.deepEqual(rcptTo[1], { email: 'cc@example.com' });
  });

  // sendDraft is the only transmit path, so the envelope must cover every recipient
  // field — including bcc, which appears in the SMTP envelope but not the headers.
  it('includes bcc recipients in the envelope rcptTo', async () => {
    const withBcc = { ...SENDABLE_DRAFT, bcc: [{ email: 'hidden@example.com' }] };
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [withBcc] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    await client.sendDraft('draft-1');

    const rcptTo = callArguments(makeReq, 1)[0].methodCalls[0][1].create.submission.envelope.rcptTo;
    assert.deepEqual(rcptTo, [
      { email: 'bob@example.com' },
      { email: 'cc@example.com' },
      { email: 'hidden@example.com' },
    ]);
  });

  it('reports the embedded images the sent message carried', async () => {
    const withImage = {
      ...SENDABLE_DRAFT,
      attachments: [
        { partId: '3', blobId: 'blob-pic', type: 'image/png', name: 'pic.png', cid: 'pic@x', disposition: 'inline', size: 2048 },
        { partId: '4', blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', cid: null, disposition: 'attachment', size: 9999 },
      ],
    };
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [withImage] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const outcome = await client.sendDraft('draft-1');
    assert.deepEqual(outcome.notes, ['Sent with 1 embedded image(s) (2 KB).']);
    // The receipt needs the part listing, so the pre-send read must ask for it.
    const getParams = callArguments(makeReq)[0].methodCalls[0][1];
    assert.ok(getParams.properties.includes('attachments'));
    for (const p of ['disposition', 'cid', 'name']) assert.ok(getParams.bodyProperties.includes(p));
  });

  it('says nothing about images on a message that carried none', async () => {
    stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });
    const outcome = await client.sendDraft('draft-1');
    assert.equal(outcome.notes, undefined);
  });

  it('rejects non-draft email', async () => {
    const nonDraft = { ...SENDABLE_DRAFT, keywords: { $seen: true } };
    stubRequests(client, async () => ({
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
    stubRequests(client, async () => ({
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

  // sendDraft is the transmit gate for every compose tool, so its identity check is the
  // last line of defence against submitting mail from an unverified address (e.g. a draft
  // created externally, or one whose from was edited after the sending identity was removed).
  it('rejects a draft whose from address matches no sending identity', async () => {
    const foreignFrom = { ...SENDABLE_DRAFT, from: [{ email: 'stranger@other.com' }] };
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [foreignFrom] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /does not match any sending identity/i);
        return true;
      },
    );

    // The reject fires before submission: only the Email/get call was made.
    assert.equal(makeReq.mock.calls.length, 1);
  });

  it('throws when email not found', async () => {
    stubRequests(client, async () => ({
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
    stubRequests(client, async (req: any) => {
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
    stubRequests(client, async () => ({
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
    stubRequests(client, async () => ({
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
    stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [dual] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal((await client.sendDraft('draft-1')).submissionId, 'sub-1');
  });

  it('sends a clean text-only draft (absent partner is undefined, not empty)', async () => {
    const textOnly = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1', type: 'text/plain' }],
      htmlBody: [{ partId: '1', type: 'text/plain' }], // server aliases the one part into both lists
      bodyValues: { '1': { value: 'Real text' } },
    };
    stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [textOnly] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal((await client.sendDraft('draft-1')).submissionId, 'sub-1');
  });

  it('sends an html-only draft with real content as-is (image-only/html-only mail is valid)', async () => {
    const htmlOnly = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '2', type: 'text/html' }], // single html part aliased into both lists
      htmlBody: [{ partId: '2', type: 'text/html' }],
      bodyValues: { '2': { value: '<div><img src="https://x/banner.jpg"></div>' } },
    };
    stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [htmlOnly] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal((await client.sendDraft('draft-1')).submissionId, 'sub-1'); // not rejected — html-only is sendable
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
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails({ query: 'quarterly', limit: 10, excludeDrafts: true });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    // text in the base, $draft as its own keyword condition, AND-wrapped.
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.text === 'quarterly'));
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$draft'));
  });

  it('includes drafts by default (no notKeyword; flat text filter)', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails({ query: 'quarterly', limit: 10 });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.text, 'quarterly');
    assert.equal(filter.notKeyword, undefined);
  });

  // The JMAP filter takes a UTCDate; a bare date reaches the server as invalidArguments
  // unless it is expanded first (#70).
  it('normalises date-only and offset date bounds into the JMAP UTCDate shape', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: [], total: 0 }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));
    await client.searchEmails({ after: '2026-07-20', before: '2026-07-25T09:00:00+02:00', limit: 10 });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.after, '2026-07-20T00:00:00Z');
    assert.equal(filter.before, '2026-07-25T07:00:00Z');
  });

  it('rejects an unparseable date bound before making any request', async () => {
    const makeReq = stubRequests(client, async () => ({ methodResponses: [] }));
    await assert.rejects(
      () => client.searchEmails({ before: 'yesterday' }),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /^before is not a valid date/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 0);
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
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'd-1' } } }, 'createDraft']],
    }));

    await client.createDraft({ subject: 'Hi', to: ['Alice <a@x.example>'] });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ name: 'Alice', email: 'a@x.example' }]);
  });

  it('updateDraft parses "Name <email>" recipients into { name, email }', async () => {
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [EXISTING_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'd-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { to: ['Alice <a@x.example>'] });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.to, [{ name: 'Alice', email: 'a@x.example' }]);
  });

});

// ---------- createDraft replyTo ----------

describe('createDraft replyTo', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('includes replyTo in created email object when provided', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { created: { draft: { id: 'email-draft' } } }, 'createDraft'],
      ],
    }));

    await client.createDraft({
      subject: 'Draft with replyTo',
      replyTo: ['noreply@example.com'],
    });

    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
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

    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWithReplyTo] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-new' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { replyTo: ['new-reply@example.com'] });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'new-reply@example.com' }]);
  });

  it('preserves existing replyTo when not provided in updates', async () => {
    const existingWithReplyTo = {
      ...EXISTING_DRAFT,
      replyTo: [{ email: 'keep-me@example.com' }],
    };

    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWithReplyTo] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-new' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Updated subject only' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.replyTo, [{ email: 'keep-me@example.com' }]);
  });
});

// ---------- wildcard identity ----------

const WILDCARD_IDENTITY = { id: 'id-wild', name: 'Jonathan Godley', email: '*@example.com', mayDelete: true };

describe('sendDraft wildcard identity', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    mock.method(client, 'getIdentities', async () => [WILDCARD_IDENTITY]);
    mock.method(client, 'getMailboxes', async () => [DRAFTS_MAILBOX, SENT_MAILBOX]);
  });

  it('matches wildcard identity when draft has concrete from address', async () => {
    const wildcardDraft = { ...SENDABLE_DRAFT, from: [{ email: 'work@example.com' }] };
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [wildcardDraft] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    const { submissionId } = await client.sendDraft('draft-1');
    assert.equal(submissionId, 'sub-1');

    const submitCall = callArguments(makeReq, 1)[0];
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
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWild] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { from: 'new@example.com' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Jonathan Godley', email: 'new@example.com' }]);
  });

  it('preserves concrete from when updating without changing from', async () => {
    const existingWild = { ...EXISTING_DRAFT, from: [{ email: 'work@example.com' }] };
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [existingWild] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Changed subject only' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Jonathan Godley', email: 'work@example.com' }]);
  });
});

// ---------- version sync ----------

describe('version sync', () => {
  it('package.json, manifest.json, index.ts, and package-lock.json all have the same version', async () => {
    const { readFileSync } = await import('fs');
    const { resolve: r } = await import('path');
    const root = r(import.meta.dirname, '..');

    const pkg = JSON.parse(readFileSync(r(root, 'package.json'), 'utf8'));
    const manifest = JSON.parse(readFileSync(r(root, 'manifest.json'), 'utf8'));
    const indexSrc = readFileSync(r(root, 'src', 'index.ts'), 'utf8');
    const lock = JSON.parse(readFileSync(r(root, 'package-lock.json'), 'utf8'));

    const indexMatch = indexSrc.match(/version:\s*'([^']+)'/);
    assert.ok(indexMatch, 'Could not find version string in index.ts');

    assert.equal(pkg.version, manifest.version, 'package.json and manifest.json versions must match');
    assert.equal(pkg.version, indexMatch[1], 'package.json and index.ts versions must match');
    // The lockfile carries the version in two places (root and the "" package
    // entry); a bump that misses either would ship a mismatched .dxt.
    assert.equal(pkg.version, lock.version, 'package.json and package-lock.json versions must match');
    assert.equal(pkg.version, lock.packages['']?.version, 'package.json and package-lock.json root package versions must match');
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
      uploadUrl: 'https://api.fastmail.com/jmap/upload/{accountId}/',
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

    assert.equal(captured.url, `https://api.fastmail.com/jmap/upload/${ACCOUNT_ID}/`);
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
      uploadUrl: 'https://api.fastmail.com/jmap/upload/{accountId}/',
    }));
    return client;
  }

  it('throws the opt-in error when attachDir is undefined', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ path: 'x.pdf' }], undefined, false),
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
      const parts = await client.uploadAttachments([{ path: 'report.pdf' }], root, false);
      assert.deepEqual(parts, [
        { blobId: 'blob-up', type: 'application/pdf', name: 'report.pdf', disposition: 'attachment' },
      ]);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('builds a 5-key inline part when the composed message displays the cid', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'uploadBlob', async () => ({ blobId: 'blob-img', type: 'image/png', size: 3 }));
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'logo.png'), 'png');
      const parts = await client.uploadAttachments(
        [{ path: 'logo.png', cid: 'logo' }],
        root,
        false,
        { inlineCids: new Set(['logo']) },
      );
      assert.deepEqual(parts, [
        { blobId: 'blob-img', type: 'image/png', name: 'logo.png', disposition: 'inline', cid: 'logo' },
      ]);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  // The hard case: the caller gave the file a Content-ID but the message ships no html body
  // that could display it. Fastmail refuses an inline part in that shape outright, so the
  // file rides as an ordinary attachment — keeping its Content-ID, and never dropped.
  it('keeps a cid but disposes it as an attachment when nothing displays it', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'uploadBlob', async () => ({ blobId: 'blob-img', type: 'image/png', size: 3 }));
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'logo.png'), 'png');
      const parts = await client.uploadAttachments([{ path: 'logo.png', cid: 'logo' }], root, false, {});
      assert.deepEqual(parts, [
        { blobId: 'blob-img', type: 'image/png', name: 'logo.png', disposition: 'attachment', cid: 'logo' },
      ]);
      // And a supplied cid that is simply not among the displayed ones behaves the same.
      const other = await client.uploadAttachments(
        [{ path: 'logo.png', cid: 'logo' }],
        root,
        false,
        { inlineCids: new Set(['something-else']) },
      );
      assert.equal(other[0].disposition, 'attachment');
      assert.equal(other[0].cid, 'logo');
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
        () => client.uploadAttachments([{ path: 'f.bin', contentType: 'not a mime type' }], root, false),
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
      const parts = await client.uploadAttachments([{ path: 'a.txt' }, { path: 'b.txt', contentType: 'text/plain' }], root, false);
      assert.deepEqual(parts.map(p => p.name), ['a.txt', 'b.txt']);
      assert.deepEqual(parts.map(p => p.blobId), ['blob-1', 'blob-2']);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it('infers the upload Content-Type from the file extension, falling back to octet-stream', async (t) => {
    const client = clientWithUpload();
    const seen: string[] = [];
    t.mock.method(client, 'uploadBlob', async (data: Buffer, ct: string) => {
      seen.push(ct);
      return { blobId: 'blob-' + seen.length, type: ct, size: data.length };
    });
    // .ics and .eml are the two that change how the mail renders rather than just its
    // icon: a client decides from this header whether to show a calendar invitation or
    // a forwarded message inline. .eml is also what forward_email's asAttachment writes.
    const names = ['invite.ics', 'original.eml', 'logo.svg', 'notes.md', 'shot.webp', 'old.doc', 'old.xls', 'old.ppt', 'blob.unknownext'];
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      for (const n of names) await fsWriteFile(join(root, n), 'x');
      await client.uploadAttachments(names.map(path => ({ path })), root, false);
      assert.deepEqual(seen, [
        'text/calendar',
        'message/rfc822',
        'image/svg+xml',
        'text/markdown',
        'image/webp',
        'application/msword',
        'application/vnd.ms-excel',
        'application/vnd.ms-powerpoint',
        'application/octet-stream',
      ]);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  // ---- the two in-account sources, behind their own opt-in ----

  // The gate is per SOURCE. FASTMAIL_ATTACH_DIR governs reading a local file off disk;
  // these two reference content the account already holds, so the local-disk opt-in has
  // nothing to say about them and refusing them is FASTMAIL_ALLOW_BLOB_ATTACH's job.
  it('refuses a blobId item when blob attaching is off, even with an attach dir set', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ blobId: 'G1', name: 'r.pdf' }], '/attach/root', false),
      (e: unknown) => e instanceof InvalidInputError
        && /attachments\[0\] attaches by blobId, which is disabled/.test((e as Error).message)
        && /FASTMAIL_ALLOW_BLOB_ATTACH/.test((e as Error).message),
    );
  });

  it('refuses an emailId + attachmentId item when blob attaching is off', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ emailId: 'M1', attachmentId: 'p2', name: 'r.pdf' }], '/attach/root', false),
      (e: unknown) => e instanceof InvalidInputError
        && /attaches by emailId \+ attachmentId, which is disabled/.test((e as Error).message)
        && /FASTMAIL_ALLOW_BLOB_ATTACH/.test((e as Error).message),
    );
  });

  // And the mirror: the blob opt-in does NOT open the local-disk source.
  it('still refuses a path item on the attach-dir gate when only blob attaching is on', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ path: 'x.pdf' }], undefined, true),
      (e: unknown) => e instanceof PathAccessError && /FASTMAIL_ATTACH_DIR/.test((e as Error).message),
    );
  });

  it('references a blobId without uploading anything, inferring the type from the name', async (t) => {
    const client = clientWithUpload();
    let uploads = 0;
    t.mock.method(client, 'uploadBlob', async () => { uploads++; return { blobId: 'nope', type: 'x/y', size: 1 }; });
    const parts = await client.uploadAttachments([{ blobId: 'G-stored', name: 'report.pdf' }], undefined, true);
    assert.deepEqual(parts, [
      { blobId: 'G-stored', type: 'application/pdf', name: 'report.pdf', disposition: 'attachment' },
    ]);
    assert.equal(uploads, 0);
  });

  it('lets a caller contentType override the inferred type on a blobId item, through the MIME vet', async () => {
    const client = clientWithUpload();
    const parts = await client.uploadAttachments(
      [{ blobId: 'G-stored', name: 'report.bin', contentType: 'application/pdf' }], undefined, true,
    );
    assert.equal(parts[0].type, 'application/pdf');
    await assert.rejects(
      () => client.uploadAttachments([{ blobId: 'G1', name: 'r.bin', contentType: 'not a mime type' }], undefined, true),
      /invalid contentType/,
    );
  });

  it("resolves an emailId + attachmentId to the part's blob, name and type", async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'getAttachmentInfo', async () => ({
      blobId: 'G-part', type: 'image/png', name: 'chart.png', size: 12, matchedBy: 'partId',
    }));
    const parts = await client.uploadAttachments([{ emailId: 'M1', attachmentId: '2' }], undefined, true);
    assert.deepEqual(parts, [
      { blobId: 'G-part', type: 'image/png', name: 'chart.png', disposition: 'attachment' },
    ]);
  });

  it('lets the caller override the resolved name and type of a message part', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'getAttachmentInfo', async () => ({
      blobId: 'G-part', type: 'image/png', name: 'chart.png', size: 12, matchedBy: 'blobId',
    }));
    const parts = await client.uploadAttachments(
      [{ emailId: 'M1', attachmentId: 'G-part', name: 'renamed.png', contentType: 'image/webp', cid: 'chart' }],
      undefined,
      true,
      { inlineCids: new Set(['chart']) },
    );
    assert.deepEqual(parts, [
      { blobId: 'G-part', type: 'image/webp', name: 'renamed.png', disposition: 'inline', cid: 'chart' },
    ]);
  });

  // The compose direction refuses a reference that resolved ONLY by position: an entry
  // number names a different file after any change to the listing, and here the wrong file
  // is baked into a draft that send_draft then transmits.
  it('refuses a message part that resolved only through the entry-number fallback', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'getAttachmentInfo', async () => ({
      blobId: 'G-part', type: 'image/png', name: 'chart.png', matchedBy: 'index',
    }));
    await assert.rejects(
      () => client.uploadAttachments([{ emailId: 'M1', attachmentId: '2' }], undefined, true),
      (e: unknown) => e instanceof InvalidInputError
        && /only as an entry number/.test((e as Error).message)
        && /get_email_attachments/.test((e as Error).message),
    );
  });

  // The rejection is on HOW the reference resolved, never on how the string looks.
  // parseInt("2abc") is 2, so a looks-numeric test would call this the index form — while
  // the resolver rejects it outright as unusable, and a partId of "2" is a real part that
  // such a test would wrongly refuse.
  it('decides the entry-number refusal from the resolution, not from a numeric-looking string', async (t) => {
    const client = clientWithUpload();
    // "2abc" never reaches the index branch at all: the resolver refuses it as unusable.
    t.mock.method(client, 'getAttachmentInfo', async () => {
      throw new InvalidInputError('attachmentId "2abc" is not a usable attachment reference.');
    });
    await assert.rejects(
      () => client.uploadAttachments([{ emailId: 'M1', attachmentId: '2abc' }], undefined, true),
      (e: unknown) => e instanceof InvalidInputError && /not a usable attachment reference/.test((e as Error).message),
    );

    // And a digit string that matched a real partId is accepted, though it looks numeric.
    const digits = clientWithUpload();
    t.mock.method(digits, 'getAttachmentInfo', async () => ({
      blobId: 'G-part', type: 'text/plain', name: 'note.txt', matchedBy: 'partId',
    }));
    const parts = await digits.uploadAttachments([{ emailId: 'M1', attachmentId: '2' }], undefined, true);
    assert.equal(parts[0].blobId, 'G-part');
  });

  // A batch rejects on one bad entry, so the wrap has to say WHICH entry was bad and keep
  // the lookup's own advice about what to pass instead — a caller that gets neither is left
  // re-checking every item by hand.
  it('names the attachments index when a part does not resolve, keeping the advice', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'getAttachmentInfo', async () => {
      throw new InvalidInputError(
        'Attachment not found: attachmentId "p404" matches no part of that message. ' +
        'List its parts with get_email_attachments and pass a partId or blobId from it.'
      );
    });
    await assert.rejects(
      () => client.uploadAttachments([{ emailId: 'M1', attachmentId: 'p404' }], undefined, true),
      (e: unknown) => e instanceof InvalidInputError
        && /attachments\[0\]/.test((e as Error).message)
        && /matches no part of that message/.test((e as Error).message)
        && /get_email_attachments/.test((e as Error).message),
    );
  });

  // The wrap is by CLASS. A transport failure is not a caller mistake and must not be
  // relabelled as one — that would send the caller off correcting ids that were fine.
  it('does not relabel a transport failure as bad caller input', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'getAttachmentInfo', async () => {
      throw new Error('socket hang up');
    });
    await assert.rejects(
      () => client.uploadAttachments([{ emailId: 'M1', attachmentId: 'p2' }], undefined, true),
      (e: unknown) => !(e instanceof InvalidInputError) && /socket hang up/.test((e as Error).message),
    );
  });

  // coerceAttachments guarantees exactly one source, but that guarantee lives in another
  // module. If a future source were added there and not here, this dispatch would fall
  // through to the path branch with no path — an attachment built from a file nobody named.
  it('refuses a spec that reaches it naming no source at all', async () => {
    const client = clientWithUpload();
    await assert.rejects(
      () => client.uploadAttachments([{ name: 'orphan.pdf' } as any], '/attach/root', true),
      (e: unknown) => e instanceof InvalidInputError && /names no source/.test((e as Error).message),
    );
  });

  it('keeps parts in spec order across a mixed batch of sources', async (t) => {
    const client = clientWithUpload();
    t.mock.method(client, 'uploadBlob', async () => ({ blobId: 'blob-local', type: 'text/plain', size: 2 }));
    t.mock.method(client, 'getAttachmentInfo', async () => ({
      blobId: 'G-part', type: 'image/png', name: 'chart.png', matchedBy: 'partId',
    }));
    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-att-'));
    try {
      await fsWriteFile(join(root, 'a.txt'), 'aa');
      const parts = await client.uploadAttachments(
        [{ blobId: 'G-stored', name: 'first.pdf' }, { path: 'a.txt' }, { emailId: 'M1', attachmentId: 'p9' }],
        root,
        true,
      );
      assert.deepEqual(parts.map((p) => p.blobId), ['G-stored', 'blob-local', 'G-part']);
      assert.deepEqual(parts.map((p) => p.name), ['first.pdf', 'a.txt', 'chart.png']);
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
        () => client.uploadAttachments([{ path: 'good.txt' }, { path: '../escape.txt' }], root, false),
        /must be within/,
      );
      assert.equal(uploads, 0);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

// ---------- outgoing attachments wired into createDraft ----------

describe('attachments on create', () => {
  it('createDraft places attachment parts in the email object (reply draft branch)', async () => {
    const client = makeClient();
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-42' } } }, 'createDraft']],
    }));
    const part = { blobId: 'b2', type: 'image/png', name: 'p.png', disposition: 'attachment' };
    await client.createDraft({ subject: 'Hello', attachments: [part] });
    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.attachments, [part]);
  });

  it('createDraft accepts an attachment-only draft (no to/subject/body) — attachments count as content', async () => {
    const client = makeClient();
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-43' } } }, 'createDraft']],
    }));
    const part = { blobId: 'b3', type: 'application/pdf', name: 'only.pdf', disposition: 'attachment' };
    const id = await client.createDraft({ attachments: [part] }); // must NOT throw the contentless guard
    assert.equal(id, 'email-43');
    assert.deepEqual(callArguments(makeReq)[0].methodCalls[0][1].create.draft.attachments, [part]);
  });

  it('createDraft still rejects a truly empty draft (no fields, no attachments)', async () => {
    const client = makeClient();
    await assert.rejects(
      () => client.createDraft({}),
      (err: Error) => {
        assert.match(err.message, /At least one of to, subject, textBody, htmlBody, or attachments/);
        // compose-path input reject → InvalidInputError → InvalidParams (#41)
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
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

  // Every other carry test stubs the Email/get RESPONSE, so all of them stay green if
  // someone drops a name from the REQUEST's properties list — the recreate would then
  // silently rebuild the draft without whatever was dropped. This pins the request side.
  it('asks the server for every property the faithful recreate carries', async () => {
    const makeReq = mockUpdate(client, DRAFT_ONE_ATT);
    await client.updateDraft('draft-1', { subject: 'New' });
    const [method, params] = callArguments(makeReq)[0].methodCalls[0];
    assert.equal(method, 'Email/get');
    for (const prop of [
      'attachments', 'inReplyTo', 'references', 'keywords', 'mailboxIds',
      'header:X-Forwarded-Message-Id:asMessageIds',
    ]) {
      assert.ok(params.properties.includes(prop), `Email/get must request '${prop}'`);
    }
    // Attachment metadata rides on bodyProperties, not properties — a carried part needs
    // its name/disposition, and the inline-image reject needs cid.
    for (const prop of ['blobId', 'type', 'name', 'disposition', 'cid']) {
      assert.ok(params.bodyProperties.includes(prop), `Email/get must request bodyProperty '${prop}'`);
    }
  });

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

  it('matches a stored name that carries surrounding whitespace', async () => {
    // BOTH sides of the comparison are trimmed, and the pair is what makes it safe. The
    // coercer trims the caller's ref so the three ways of writing a list agree, which on its
    // own would make ' doc.pdf ' unreachable — a caller could name it exactly and still be
    // told it matched nothing, with no spelling that works. Trimming the stored side too
    // restores it. Delete either half and this test fails.
    const paddedDraft = {
      ...EXISTING_DRAFT,
      attachments: [
        { blobId: 'b-pad', type: 'application/pdf', name: ' doc.pdf ', disposition: 'attachment', cid: null, partId: '3', size: 1 },
      ],
    };
    const makeReq = mockUpdate(client, paddedDraft);
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

describe('updateDraft embedded images (#13)', () => {
  let client: JmapClient;
  beforeEach(() => { client = makeClient(); });

  // The shape of an identifier this server mints for a quoted image. Tests can't inject the
  // generator through updateDraft, so a fresh mint is recognized by shape.
  const MINT_SHAPE = /^ii-[0-9a-f]{32}@inline\.invalid$/;
  // A stored one, standing in for a mint an earlier call made.
  const STORED_MINT = 'ii-0123456789abcdef0123456789abcdef@inline.invalid';
  const SECOND_MINT = 'ii-fedcba9876543210fedcba9876543210@inline.invalid';

  function imagePart(over: any = {}) {
    return {
      partId: 'p9', blobId: 'blob-img', type: 'image/png', name: 'pic.png',
      cid: 'pic@x', disposition: 'inline', size: 2048, ...over,
    };
  }

  // An html-only draft: `parts` populates the JMAP attachments array, `bodyParts` the html
  // body list (which is where a server routes an embedded image on some MIME shapes).
  function htmlDraft(html: string, parts: any[] = [], bodyParts: any[] = []) {
    return {
      ...EXISTING_DRAFT,
      textBody: null,
      htmlBody: [{ partId: 'h', type: 'text/html' }, ...bodyParts],
      bodyValues: { h: { value: html } },
      attachments: parts,
    };
  }

  // Dispatches Email/get by id so the post-edit re-read can return a DIFFERENT draft from the
  // one the edit started with — the whole point of that read is to report on the saved state.
  function mockEdit(c: JmapClient, before: any, saved?: any, opts: { readBackFails?: boolean } = {}) {
    mock.method(c, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        if (params.ids?.[0] === 'draft-2') {
          if (opts.readBackFails) throw new Error('read-back unavailable');
          return { methodResponses: [['Email/get', { list: saved ? [saved] : [] }, 'getEmail']] };
        }
        return { methodResponses: [['Email/get', { list: [before] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
  }

  function createdDraft(makeReq: RequestMock) {
    const [request] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][1].create,
      'creating the replacement draft',
    );
    return request.methodCalls[0][1].create.draft;
  }

  // ---- what the draft carries survives an edit that isn't about it ----

  it('carries a body-list-routed image the JMAP attachments array never listed', async () => {
    const draft = htmlDraft('<p>see <img src="cid:pic@x"></p>', [], [imagePart()]);
    const makeReq = mockEdit(client, draft);
    await client.updateDraft('draft-1', { subject: 'New' });
    assert.deepEqual(createdDraft(makeReq).attachments, [
      { blobId: 'blob-img', type: 'image/png', name: 'pic.png', disposition: 'inline', cid: 'pic@x' },
    ]);
  });

  it('says nothing about images on an edit that cannot have changed what the body displays', async () => {
    const draft = htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]);
    const makeReq = mockEdit(client, draft);
    const result = await client.updateDraft('draft-1', { subject: 'New' });
    assert.equal(result.notes, undefined);
    assert.equal(result.inlineImages, undefined);
    assert.equal(makeReq.mock.calls.length, 3); // no re-read: nothing was attached
  });

  it('resolves a percent-encoded reference to the part that supplies it', async () => {
    const draft = htmlDraft('<p>x</p>', [imagePart()]);
    const makeReq = mockEdit(client, draft, htmlDraft('<p>y</p>', [imagePart()]));
    await client.updateDraft('draft-1', { htmlBody: '<p>y <img src="cid:pic%40x"></p>' });
    assert.match(createdDraft(makeReq).bodyValues.html.value, /cid:pic%40x/);
  });

  // ---- a draft whose stored body already references an image it doesn't carry ----

  const BROKEN = htmlDraft('<p>hi</p><img src="cid:gone@x">');

  it('refuses a body edit while the stored body references an image nothing supplies', async () => {
    mockEdit(client, BROKEN);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>new</p><img src="cid:gone@x">' }),
      /stored body references image identifier\(s\) with no matching attachment.*gone@x/s,
    );
  });

  it('still runs a metadata edit on such a draft, carrying the broken body verbatim', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await client.updateDraft('draft-1', { subject: 'Renamed' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>hi</p><img src="cid:gone@x">');
    assert.equal(draft.textBody, undefined); // body-invariant: no text part invented
  });

  it('still appends an unrelated attachment to such a draft', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await client.updateDraft('draft-1', {
      attachments: [{ blobId: 'b-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' }],
    });
    assert.equal(createdDraft(makeReq).attachments.length, 1);
  });

  it('accepts an attachment that supplies the missing image', async () => {
    const supplied = imagePart({ cid: 'gone@x', blobId: 'b-gone' });
    const makeReq = mockEdit(client, BROKEN, htmlDraft('<p>hi</p><img src="cid:gone@x">', [supplied]));
    const result = await client.updateDraft('draft-1', {
      attachments: [{ blobId: 'b-gone', type: 'image/png', name: 'pic.png', cid: 'gone@x', disposition: 'inline' }],
    });
    assert.equal(createdDraft(makeReq).attachments.length, 1);
    assert.deepEqual(result.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('accepts a body edit that replaces the broken body outright', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await client.updateDraft('draft-1', { htmlBody: '<p>all new</p>' });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, '<p>all new</p>');
  });

  // ---- identifiers the recreate cannot reproduce ----

  const EXOTIC_CID = htmlDraft('<p>hi</p>', [imagePart({ cid: 'has space@x' })]);

  it('refuses a body edit when a stored identifier cannot be re-created faithfully', async () => {
    mockEdit(client, EXOTIC_CID);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>new</p>' }),
      /cannot safely re-create.*Recreate it — with reply_email or forward_email/s,
    );
  });

  it('leaves metadata edits of such a draft alone (the carry reproduces the value verbatim)', async () => {
    const makeReq = mockEdit(client, EXOTIC_CID);
    await client.updateDraft('draft-1', { subject: 'Renamed' });
    assert.equal(createdDraft(makeReq).attachments[0].cid, 'has space@x');
  });

  // ---- body shapes the recreate cannot express ----

  it('refuses a body that interleaves two parts of the same text type', async () => {
    const interleaved = {
      ...EXISTING_DRAFT,
      textBody: null,
      htmlBody: [{ partId: 'h1', type: 'text/html' }, imagePart(), { partId: 'h2', type: 'text/html' }],
      bodyValues: { h1: { value: '<p>above</p>' }, h2: { value: '<p>below</p>' } },
      attachments: [],
    };
    mockEdit(client, interleaved);
    await assert.rejects(
      () => client.updateDraft('draft-1', { subject: 'X' }),
      /interleaves multiple text parts of the same type.*see issue #85/s,
    );
  });

  it('raises a body-shape refusal before an attachment one', async () => {
    const weird = {
      ...EXISTING_DRAFT,
      textBody: [{ partId: '1', type: 'text/calendar' }],
      htmlBody: null,
      bodyValues: { '1': { value: 'BEGIN:VCALENDAR' } },
      attachments: [],
    };
    mockEdit(client, weird);
    await assert.rejects(
      () => client.updateDraft('draft-1', { removeAttachments: ['no-such-blob'] }),
      /cannot carry/,
    );
  });

  // ---- removals ----

  it('refuses a removal that would leave the surviving body pointing at nothing', async () => {
    mockEdit(client, htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]));
    await assert.rejects(
      () => client.updateDraft('draft-1', { removeAttachments: ['blob-img'] }),
      /removeAttachments would remove an image the draft's body still references/,
    );
  });

  it('refuses an attachment wipe that would leave the surviving body pointing at nothing', async () => {
    mockEdit(client, htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]));
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['attachments'] }),
      /would strip image\(s\) the surviving body still references/,
    );
  });

  it('wipes the attachments when the surviving body references none of them', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>no images here</p>', [imagePart()]));
    await client.updateDraft('draft-1', { clearFields: ['attachments'] });
    assert.equal(createdDraft(makeReq).attachments, undefined);
  });

  it('takes a server-managed image off the draft when the rewritten body stops displaying it', async () => {
    const draft = htmlDraft(`<p><img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    const makeReq = mockEdit(client, draft);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>text only now</p>' });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 0, removed: 1 });
    assert.deepEqual(result.notes, ['Removed 1 image(s) that were embedded in the quote.']);
  });

  it("keeps someone else's image as a regular attachment when the body stops displaying it", async () => {
    const draft = htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]);
    const makeReq = mockEdit(client, draft);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>text only now</p>' });
    assert.deepEqual(createdDraft(makeReq).attachments, [
      { blobId: 'blob-img', type: 'image/png', name: 'pic.png', disposition: 'attachment', cid: 'pic@x' },
    ]);
    assert.deepEqual(result.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('counts both parts when two embedded images share one blob', async () => {
    // The same bytes displayed twice under different Content-IDs are two parts, and both
    // leave the draft. Counting by blob would report one and quietly lose the other.
    const shared = [
      imagePart({ partId: 'a', cid: STORED_MINT }),
      imagePart({ partId: 'b', cid: SECOND_MINT }),
    ];
    const draft = htmlDraft(`<p><img src="cid:${STORED_MINT}"><img src="cid:${SECOND_MINT}"></p>`, shared);
    const makeReq = mockEdit(client, draft);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>text only now</p>' });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 0, removed: 2 });
    assert.deepEqual(result.notes, ['Removed 2 image(s) that were embedded in the quote.']);
  });

  it('counts both parts when two blob-sharing images degrade to attachments', async () => {
    const shared = [
      imagePart({ partId: 'a', cid: 'first@x' }),
      imagePart({ partId: 'b', cid: 'second@x' }),
    ];
    const draft = htmlDraft('<p><img src="cid:first@x"><img src="cid:second@x"></p>', shared);
    const makeReq = mockEdit(client, draft);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>text only now</p>' });
    assert.equal(createdDraft(makeReq).attachments.length, 2);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 2, removed: 0 });
    assert.deepEqual(result.notes, ['2 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('reports a part the caller removed once, as a removal it asked for', async () => {
    // The explicit removal is the caller's own action and needs no sentence; what must not
    // happen is the same part ALSO being counted as one the edit took off the body.
    const draft = htmlDraft(`<p><img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    const result = await client.updateDraft.call(
      (mockEdit(client, draft), client),
      'draft-1',
      { htmlBody: '<p>gone</p>', removeAttachments: ['blob-img'] },
    );
    assert.equal(result.notes, undefined);
    assert.equal(result.inlineImages, undefined);
  });

  // ---- collisions ----

  it('refuses two attachments in one call sharing an identifier', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        htmlBody: '<p>y <img src="cid:dup@me"></p>',
        attachments: [
          { blobId: 'b1', type: 'image/png', cid: 'dup@me', disposition: 'inline' },
          { blobId: 'b2', type: 'image/png', cid: 'dup@me', disposition: 'inline' },
        ],
      }),
      /2 attachments items share cid "dup@me"/,
    );
  });

  it('refuses an attachment whose identifier is already used on the draft', async () => {
    mockEdit(client, htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]));
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        attachments: [{ blobId: 'b-other', type: 'image/png', cid: 'pic@x', disposition: 'inline' }],
      }),
      /already used by another attachment on this draft/,
    );
  });

  it('refuses a caller reference to a server-managed identifier', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: `<p><img src="cid:${STORED_MINT}"></p>` }),
      /a server-managed identifier for quoted images/,
    );
  });

  it('refuses a caller reference nothing supplies', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p><img src="cid:mine-logo"></p>' }),
      /references cid "mine-logo" but no attachment supplies it.*add an attachments item with cid: "mine-logo"/s,
    );
  });

  it('offers no attachments repair when this server cannot attach files at all', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => client.updateDraft(
        'draft-1',
        { htmlBody: '<p><img src="cid:mine-logo"></p>' },
        { attachmentsEnabled: false },
      ),
      /sending attachments is disabled on this server/,
    );
  });

  // ---- an added file the edited body displays ----

  // The upload cannot know what body the edit ends up with, so a freshly uploaded part
  // arrives dispositioned as an ordinary attachment. A Content-ID never makes a part inline
  // on its own, so the marking has to happen here, against the body this call assembled.
  it('marks an added file inline when the edited body references its cid', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>x</p>', []));
    await client.updateDraft('draft-1', {
      htmlBody: '<p><img src="cid:mine-logo"></p>',
      attachments: [{ blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'attachment' }],
    });
    assert.deepEqual(createdDraft(makeReq).attachments, [
      { blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'inline' },
    ]);
  });

  it('leaves an added file with an unreferenced cid as an ordinary attachment, and says so', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>x</p>', []));
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>no image here</p>',
      attachments: [{ blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'attachment' }],
    });
    assert.equal(createdDraft(makeReq).attachments[0].disposition, 'attachment');
    assert.equal(createdDraft(makeReq).attachments[0].cid, 'mine-logo');
    assert.deepEqual(result.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  // The other way to end up with a file nothing displays: the reference was written, but
  // the same edit left the draft with no html body to display it from.
  it('says an added file was attached when the edit clears the html body', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    const result = await client.updateDraft('draft-1', {
      textBody: 'plain text only now',
      clearFields: ['htmlBody'],
      attachments: [{ blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'attachment' }],
    });
    assert.deepEqual(result.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('says nothing about an ordinary added file that asked to embed nothing', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    const result = await client.updateDraft('draft-1', {
      subject: 'New subject',
      attachments: [{ blobId: 'b-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' }],
    });
    assert.equal(result.notes, undefined);
  });

  // ---- confirming what the saved draft carries ----

  const NEW_INLINE = { blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'logo@me', disposition: 'inline' };
  const WITH_LOGO = htmlDraft('<p><img src="cid:logo@me"></p>', [imagePart({ blobId: 'b-logo', cid: 'logo@me', name: 'logo.png' })]);

  it('re-reads the saved draft when it attached an embedded image, and reports its size', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>x</p>', []), WITH_LOGO);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p><img src="cid:logo@me"></p>',
      attachments: [NEW_INLINE],
    });
    assert.equal(makeReq.mock.calls.length, 4); // the re-read is the fourth call
    assert.deepEqual(result.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('says so when the saved draft comes back without the image it attached', async () => {
    const result = await client.updateDraft.call(
      (mockEdit(client, htmlDraft('<p>x</p>', []), htmlDraft('<p>x</p>', [])), client),
      'draft-1',
      { htmlBody: '<p><img src="cid:logo@me"></p>', attachments: [NEW_INLINE] },
    );
    assert.ok(result.notes?.some((n) => /were not found on the saved draft/.test(n)));
  });

  it('never fails the edit when the confirming read does', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []), undefined, { readBackFails: true });
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p><img src="cid:logo@me"></p>',
      attachments: [NEW_INLINE],
    });
    assert.equal(result.id, 'draft-2');
    assert.ok(result.notes?.some((n) => /could not re-read it to confirm/.test(n)));
  });
});

describe('updateDraft quote rebuild with embedded images (#13)', () => {
  let client: JmapClient;
  beforeEach(() => { client = makeClient(); });

  const MINT_SHAPE = /^ii-[0-9a-f]{32}@inline\.invalid$/;
  const STORED_MINT = 'ii-0123456789abcdef0123456789abcdef@inline.invalid';
  const SECOND_MINT = 'ii-fedcba9876543210fedcba9876543210@inline.invalid';

  function storedQuotePart(cid: string, blobId = 'blob-one') {
    return { partId: 'p2', blobId, type: 'image/png', name: 'one.png', cid, disposition: 'inline', size: 2048 };
  }

  // A reply draft this server made: a note, an attribution, and a cited quote whose image is
  // supplied by a part already on the draft.
  function replyDraft(quoteImageCids: string[], parts: any[]) {
    const imgs = quoteImageCids.map((c) => `<img src="cid:${c}">`).join('');
    return {
      id: 'draft-1', subject: 'Re: Hello',
      from: [{ email: 'me@example.com' }], to: [{ email: 'bob@example.com' }], cc: [], bcc: [],
      mailboxIds: { 'mb-drafts': true }, keywords: { $draft: true },
      inReplyTo: ['orig-msg@example.com'], references: ['orig-msg@example.com'],
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: {
        t: { value: 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Jon wrote:\n> ORIGINAL' },
        h: { value: `<p>my reply</p><div>On Sun, Jun 28, 2026, at 12:46 AM, Jon wrote:</div><blockquote type="cite">${imgs}<p>ORIGINAL</p></blockquote>` },
      },
      attachments: parts,
    };
  }

  // The message being replied to. `imageCids` are the references its own html makes, each
  // supplied by one of `parts`.
  function originalWith(imageCids: string[], parts: any[]) {
    const imgs = imageCids.map((c) => `<img src="cid:${c}">`).join('');
    return {
      id: 'orig-1',
      messageId: ['orig-msg@example.com'],
      from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
      sentAt: '2026-06-15T03:29:02Z',
      subject: 'Hello',
      textBody: [{ partId: 'ot', type: 'text/plain' }],
      htmlBody: [{ partId: 'oh', type: 'text/html' }],
      bodyValues: { ot: { value: 'ORIGINAL TEXT' }, oh: { value: `<p>ORIGINAL${imgs}</p>` } },
      attachments: parts,
    };
  }

  const ORIG_ONE_IMAGE = originalWith(['one@orig'], [
    { partId: 'op', blobId: 'blob-one', type: 'image/png', name: 'one.png', cid: 'one@orig', disposition: 'inline', size: 2048 },
  ]);
  const ORIG_NO_IMAGES = originalWith([], []);
  const ORIG_TEXT_ONLY = {
    id: 'orig-1', messageId: ['orig-msg@example.com'],
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    sentAt: '2026-06-15T03:29:02Z', subject: 'Hello',
    textBody: [{ partId: 'ot', type: 'text/plain' }], htmlBody: [], bodyValues: { ot: { value: 'ORIGINAL TEXT' } },
    attachments: [],
  };

  // Echoes the created draft back on a re-read of the saved copy: what the server stores is
  // what the create sent, which is what makes the confirming read meaningful here (the
  // freshly minted identifiers aren't knowable in advance).
  function mockKeep(c: JmapClient, draft: any, original: any) {
    mock.method(c, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    let created: any = null;
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        if (params.ids?.[0] === 'orig-1') return { methodResponses: [['Email/get', { list: [original] }, 'email']] };
        if (params.ids?.[0] === 'draft-2') {
          return { methodResponses: [['Email/get', { list: [{ id: 'draft-2', attachments: created?.attachments ?? [] }] }, 'getEmail']] };
        }
        return { methodResponses: [['Email/get', { list: [draft] }, 'getEmail']] };
      }
      if (params.create) {
        created = params.create.draft;
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
  }

  function createdDraft(makeReq: RequestMock) {
    const [request] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][1].create,
      'creating the replacement draft',
    );
    return request.methodCalls[0][1].create.draft;
  }

  it('keeps the identifier a stored quote image already has', async () => {
    const makeReq = mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_ONE_IMAGE);
    await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.html.value, new RegExp(`cid:${STORED_MINT}`));
    assert.deepEqual(draft.attachments, [
      { blobId: 'blob-one', type: 'image/png', name: 'one.png', disposition: 'inline', cid: STORED_MINT },
    ]);
  });

  it('mints a fresh identifier when the draft carries no part for the quoted image', async () => {
    const makeReq = mockKeep(client, replyDraft([], []), ORIG_ONE_IMAGE);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.attachments.length, 1);
    assert.match(draft.attachments[0].cid, MINT_SHAPE);
    assert.equal(draft.attachments[0].disposition, 'inline');
    assert.match(draft.bodyValues.html.value, new RegExp(`cid:${draft.attachments[0].cid.replace(/[.@]/g, '\\$&')}`));
    assert.deepEqual(result.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('mints afresh when the stored part carries a different blob', async () => {
    const makeReq = mockKeep(
      client,
      replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT, 'blob-stale')]),
      ORIG_ONE_IMAGE,
    );
    await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    // The stale part is unreferenced by the rebuilt quote and comes off; the minted one rides
    // its own channel at the end.
    assert.equal(draft.attachments.length, 1);
    assert.match(draft.attachments[0].cid, MINT_SHAPE);
    assert.equal(draft.attachments[0].blobId, 'blob-one');
  });

  it('claims one survivor per reference when two references share a blob', async () => {
    const original = originalWith(['one@orig', 'two@orig'], [
      { partId: 'oa', blobId: 'blob-one', type: 'image/png', name: 'one.png', cid: 'one@orig', disposition: 'inline', size: 2048 },
      { partId: 'ob', blobId: 'blob-one', type: 'image/png', name: 'one.png', cid: 'two@orig', disposition: 'inline', size: 2048 },
    ]);
    const draft = replyDraft([STORED_MINT, SECOND_MINT], [
      storedQuotePart(STORED_MINT),
      { ...storedQuotePart(SECOND_MINT), partId: 'p3' },
    ]);
    const makeReq = mockKeep(client, draft, original);
    await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    const created = createdDraft(makeReq);
    assert.deepEqual(created.attachments.map((p: any) => p.cid), [STORED_MINT, SECOND_MINT]);
    assert.match(created.bodyValues.html.value, new RegExp(`cid:${STORED_MINT}`));
    assert.match(created.bodyValues.html.value, new RegExp(`cid:${SECOND_MINT}`));
  });

  it('takes a stored quote image off the draft when the rebuilt quote no longer shows it', async () => {
    const makeReq = mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_NO_IMAGES);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 0, removed: 1 });
    assert.deepEqual(result.notes, ['Removed 1 image(s) that were embedded in the quote.']);
  });

  it('re-embeds the quote images after an attachment wipe, and says it did both', async () => {
    const makeReq = mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_ONE_IMAGE);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>edited</p>', originalEmailId: 'orig-1', clearFields: ['attachments'],
    });
    const draft = createdDraft(makeReq);
    assert.equal(draft.attachments.length, 1);
    assert.match(draft.attachments[0].cid, MINT_SHAPE); // wiped, then re-embedded afresh
    assert.ok(result.notes?.includes('Cleared 1 attachment(s); the kept quote re-embedded 1 image(s).'));
  });

  it('refuses to remove an individual image the kept quote supplies', async () => {
    mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_ONE_IMAGE);
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        htmlBody: '<p>edited</p>', originalEmailId: 'orig-1', removeAttachments: ['blob-one'],
      }),
      /embedded by the kept quote.*Use noQuote with a replacement body/s,
    );
  });

  it('drops the quote and its images when noQuote replaces the body', async () => {
    const makeReq = mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_ONE_IMAGE);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>bare</p>', noQuote: true });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.deepEqual(result.notes, ['Removed 1 image(s) that were embedded in the quote.']);
  });

  it('counts a quote image once when the caller also names it for removal', async () => {
    // Dropping the quote would take the image off anyway. Because the named removal is
    // applied first, the image is gone exactly once and the result reports the caller's own
    // removal rather than also claiming the edit took it off the body.
    const makeReq = mockKeep(client, replyDraft([STORED_MINT], [storedQuotePart(STORED_MINT)]), ORIG_ONE_IMAGE);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>bare</p>', noQuote: true, removeAttachments: ['blob-one'],
    });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.equal(result.notes, undefined);
    assert.equal(result.inlineImages, undefined);
  });

  it('quotes a text-only original on the keep path', async () => {
    const makeReq = mockKeep(client, replyDraft([], []), ORIG_TEXT_ONLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>edited</p>', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.match(draft.bodyValues.html.value, /ORIGINAL TEXT/);
    assert.match(draft.bodyValues.html.value, /<blockquote type="cite"/);
  });

  it('leaves a text-only keep with no minted parts (no html body would ship them)', async () => {
    const textOnlyReply = {
      ...replyDraft([], []),
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Jon wrote:\n> ORIGINAL' } },
    };
    const makeReq = mockKeep(client, textOnlyReply, ORIG_ONE_IMAGE);
    await client.updateDraft('draft-1', { textBody: 'edited', originalEmailId: 'orig-1' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.attachments, undefined);
    assert.match(draft.bodyValues.text.value, /> ORIGINAL/);
  });
});

describe('updateDraft — forwarded-block guard (#30, Q6 gating)', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const FWD_HEADER_PROP = 'header:X-Forwarded-Message-Id:asMessageIds';
  // Body shapes as buildForwardBodies emits them (the live store/fetch round-trip is
  // pinned separately by the release verification's marker round-trip check).
  const FWD_HTML = '<p>note</p><div><br>----- Original message -----<br>From: Ada Lovelace &lt;ada@example.com&gt;<br>Subject: Hello<br></div><div type="cite"><p>orig body</p></div>';
  const FWD_TEXT = 'note\n\n\n----- Original message -----\nFrom: Ada Lovelace <ada@example.com>\nSubject: Hello\n\norig text';
  const PLAIN_FWD_TEXT = 'no marker here at all';
  const PLAIN_FWD_HTML = '<p>no marker here at all</p>';

  const FORWARD_BASE = {
    id: 'fdraft-1', subject: 'Fwd: Hello',
    from: [{ email: 'me@example.com' }], to: [{ email: 'bob@example.com' }],
    cc: [], bcc: [],
    mailboxIds: { 'mb-drafts': true }, keywords: { $draft: true },
    // No inReplyTo/references — a forward starts a new conversation.
    [FWD_HEADER_PROP]: ['fwd-orig-msg@example.com'],
  };
  // dual: both bodies carry the forwarded block (what forward_email stores).
  const DUAL_FORWARD = { ...FORWARD_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: FWD_TEXT }, h: { value: FWD_HTML } } };
  // text-only forward (of a text-only original): one aliased text part.
  const TEXT_ONLY_FORWARD = { ...FORWARD_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: FWD_TEXT } } };
  // Q2 cell: marker-bearing bodies but NO header (Message-ID-less original, or pasted content).
  const { [FWD_HEADER_PROP]: _drop, ...HEADERLESS_BASE } = FORWARD_BASE as any;
  const HEADERLESS_FORWARD = { ...HEADERLESS_BASE, id: 'fdraft-1',
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: FWD_TEXT }, h: { value: FWD_HTML } } };
  // Q6 floor cell: header present, bodies in NO recognizable shape (foreign client).
  const HEADER_ONLY_FORWARD = { ...FORWARD_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: PLAIN_FWD_TEXT }, h: { value: PLAIN_FWD_HTML } } };
  // Pathological both-marker draft: inReplyTo + reply markers AND forward header + markers.
  const BOTH_MARKER_DRAFT = { ...FORWARD_BASE,
    inReplyTo: ['some-reply-target@example.com'], references: ['some-reply-target@example.com'],
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: {
      t: { value: 'top\n\nOn Sun, Jun 28, 2026, at 12:46 AM, X wrote:\n> quoted\n\n' + FWD_TEXT },
      h: { value: '<p>top</p><blockquote type="cite">quoted</blockquote>' + FWD_HTML },
    } };
  // asAttachment forward draft as forward_email now saves it: header recorded (for
  // send_draft's keyword maintenance), non-marker filler body, .eml attachment.
  const ASATTACH_FORWARD = { ...FORWARD_BASE, id: 'fdraft-1',
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: 'Forwarded message attached.' } },
    attachments: [{ partId: '2', blobId: 'blob-eml', type: 'message/rfc822', size: 999, name: 'Hello.eml', disposition: 'attachment', cid: null }] };
  // The same draft saved BEFORE the header was recorded on asAttachment forwards
  // (or by a client that doesn't set it): still edits freely via no-header + no-marker.
  const LEGACY_ASATTACH_FORWARD = { ...HEADERLESS_BASE, id: 'fdraft-1',
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: 'Forwarded message attached.' } },
    attachments: [{ partId: '2', blobId: 'blob-eml', type: 'message/rfc822', size: 999, name: 'Hello.eml', disposition: 'attachment', cid: null }] };
  // A REPLY draft whose quoted content contains a forwarded-message line ("> "-prefixed):
  // must dispatch to the REPLY variant (the quote-prefix anchor rejects the forward marker).
  const REPLY_QUOTING_FORWARD = { ...HEADERLESS_BASE, id: 'fdraft-1',
    inReplyTo: ['r@example.com'], references: ['r@example.com'],
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: 'my reply\n\nOn Sun, Jun 28, 2026, at 1:00 AM, Y wrote:\n> ----- Original message -----\n> From: someone\n> forwarded stuff' } } };

  // The message the forward draft forwards (id 'orig-1', quotable) and a non-quotable one.
  const ORIGINAL_FOR_FORWARD = {
    id: 'orig-1',
    messageId: ['fwd-orig-msg@example.com'],
    from: [{ name: 'Ada Lovelace', email: 'ada@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [{ partId: 'ot', type: 'text/plain' }],
    htmlBody: [{ partId: 'oh', type: 'text/html' }],
    bodyValues: { ot: { value: 'ORIGINAL TEXT BODY' }, oh: { value: '<p>ORIGINAL HTML BODY</p>' } },
  };
  const NONQUOTABLE_FOR_FORWARD = {
    id: 'orig-empty',
    messageId: ['fwd-orig-msg@example.com'],
    from: [{ name: 'Ada Lovelace', email: 'ada@example.com' }],
    subject: 'Hello',
    textBody: [], htmlBody: [], bodyValues: {},
  };

  function mockForwardUpdate(c: JmapClient, draft: any = DUAL_FORWARD) {
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        const id = params.ids?.[0];
        if (id === 'orig-1') return { methodResponses: [['Email/get', { list: [ORIGINAL_FOR_FORWARD] }, 'email']] };
        if (id === 'orig-empty') return { methodResponses: [['Email/get', { list: [NONQUOTABLE_FOR_FORWARD] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [draft] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'fdraft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
  }

  function createdDraftObj(makeReq: RequestMock) {
    const [request] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][1].create,
      'creating the replacement draft',
    );
    return request.methodCalls[0][1].create.draft;
  }

  it('fetches the forward-marking header with the draft (the guard reads it)', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    await client.updateDraft('fdraft-1', { subject: 'X' });
    const getProps = callArguments(makeReq)[0].methodCalls[0][1].properties;
    assert.ok(getProps.includes(FWD_HEADER_PROP));
  });

  it('rejects a body edit on a forward draft without a flag (normal case: runnable recovery recipe)', async () => {
    mockForwardUpdate(client, DUAL_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>new note</p>' }),
      /would drop the forwarded-message block.*forwardedMessageId via get_email.*bare id/s,
    );
  });

  it('originalEmailId keep: regenerates the block into the written body and carries the header', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>new note</p>', originalEmailId: 'orig-1' });
    const draft = createdDraftObj(makeReq);
    const html = draft.bodyValues[draft.htmlBody[0].partId].value;
    assert.match(html, /^<p>new note<\/p><div><br>----- Original message -----/);
    assert.match(html, /ORIGINAL HTML BODY/);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('originalEmailId keep regenerates BOTH bodies when both are written (custom text side kept)', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>h</p>', textBody: 'my custom t', originalEmailId: 'orig-1' });
    const draft = createdDraftObj(makeReq);
    const text = draft.bodyValues[draft.textBody[0].partId].value;
    assert.match(text, /^my custom t\n\n\n----- Original message -----/);
    assert.match(text, /ORIGINAL TEXT BODY/);
  });

  it('forward keep with a body-less original still succeeds (block always regenerates; wrong-id asymmetry is intentional)', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-empty' });
    const draft = createdDraftObj(makeReq);
    const html = draft.bodyValues[draft.htmlBody[0].partId].value;
    assert.match(html, /----- Original message -----/);
  });

  it('noQuote drops the block AND clears the forward-marking header on the recreate', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>bare note</p>', noQuote: true });
    const draft = createdDraftObj(makeReq);
    const html = draft.bodyValues[draft.htmlBody[0].partId].value;
    assert.equal(html, '<p>bare note</p>');
    assert.equal(draft[FWD_HEADER_PROP], undefined);
  });

  it('rejects originalEmailId + noQuote together (forward wording)', async () => {
    mockForwardUpdate(client, DUAL_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-1', noQuote: true }),
      /either originalEmailId \(keep the forwarded block\) or noQuote/,
    );
  });

  it('metadata-only edit is untouched by the guard and CARRIES the header', async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    const r = await client.updateDraft('fdraft-1', { subject: 'Fwd: renamed' });
    assert.equal(r.id, 'fdraft-2');
    const draft = createdDraftObj(makeReq);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it("plain-text conversion (clearFields:['htmlBody']) keeps the text-side block by construction — no challenge (recompute seam)", async () => {
    const makeReq = mockForwardUpdate(client, DUAL_FORWARD);
    const r = await client.updateDraft('fdraft-1', { clearFields: ['htmlBody'] });
    assert.equal(r.id, 'fdraft-2');
    const draft = createdDraftObj(makeReq);
    const text = draft.bodyValues[draft.textBody[0].partId].value;
    assert.match(text, /----- Original message -----/);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('text-only forward draft: editing textBody challenges; originalEmailId keeps text-only', async () => {
    mockForwardUpdate(client, TEXT_ONLY_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { textBody: 'new note' }),
      /would drop the forwarded-message block/,
    );
    const makeReq = mockForwardUpdate(client, TEXT_ONLY_FORWARD);
    await client.updateDraft('fdraft-1', { textBody: 'new note', originalEmailId: 'orig-1' });
    const draft = createdDraftObj(makeReq);
    const text = draft.bodyValues[draft.textBody[0].partId].value;
    assert.match(text, /^new note\n\n\n----- Original message -----/);
    assert.equal(draft.htmlBody, undefined);
  });

  it('Q2 cell — marker body, NO header: challenge leads with what happened and offers noQuote first', async () => {
    mockForwardUpdate(client, HEADERLESS_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>' }),
      (err: any) => {
        assert.match(err.message, /body matches a forwarded-message marker/);
        assert.match(err.message, /noQuote:true to drop that block/);
        // The un-runnable get_email step must NOT appear (there is no recorded source).
        assert.doesNotMatch(err.message, /forwardedMessageId via get_email/);
        return true;
      },
    );
  });

  it('Q6 floor — header present, unrecognizable body: challenged with the recorded-source wording', async () => {
    mockForwardUpdate(client, HEADER_ONLY_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>' }),
      (err: any) => {
        assert.match(err.message, /marked as a forward/);
        assert.match(err.message, /isn't in a shape this server can regenerate in place/);
        assert.match(err.message, /forwardedMessageId via get_email/);
        return true;
      },
    );
  });

  it('Q6 floor — a noQuote resolution clears the header, so the NEXT edit is unchallenged', async () => {
    const makeReq = mockForwardUpdate(client, HEADER_ONLY_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>rewritten</p>', noQuote: true });
    const draft = createdDraftObj(makeReq);
    assert.equal(draft[FWD_HEADER_PROP], undefined);
    // Simulate the next edit on the recreated (header-less, marker-less) draft: no challenge.
    const NEXT = { ...HEADER_ONLY_FORWARD, [FWD_HEADER_PROP]: undefined };
    mockForwardUpdate(client, NEXT);
    const r = await client.updateDraft('fdraft-1', { htmlBody: '<p>again</p>' });
    assert.equal(r.id, 'fdraft-2');
  });

  it('both-marker draft: REPLY variant wins the dispatch (reply challenge wording)', async () => {
    mockForwardUpdate(client, BOTH_MARKER_DRAFT);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>' }),
      /would drop the quoted original/,
    );
  });

  it('both-marker draft: a reply-variant noQuote ALSO clears the forward header in the same step', async () => {
    const makeReq = mockForwardUpdate(client, BOTH_MARKER_DRAFT);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>bare</p>', noQuote: true });
    const draft = createdDraftObj(makeReq);
    assert.equal(draft[FWD_HEADER_PROP], undefined);
  });

  it('asAttachment draft (header + .eml, filler body): note edits pass with NO challenge, header and .eml both carried', async () => {
    const makeReq = mockForwardUpdate(client, ASATTACH_FORWARD);
    const r = await client.updateDraft('fdraft-1', { textBody: 'updated note about the attached message' });
    assert.equal(r.id, 'fdraft-2');
    const draft = createdDraftObj(makeReq);
    // The .eml is an ordinary carried attachment on the recreate, and the recorded
    // source survives the edit (send_draft resolves it to mark the original).
    assert.equal(draft.attachments[0].type, 'message/rfc822');
    assert.equal(draft.attachments[0].blobId, 'blob-eml');
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('legacy asAttachment draft (no header): note edits still pass with NO challenge', async () => {
    const makeReq = mockForwardUpdate(client, LEGACY_ASATTACH_FORWARD);
    const r = await client.updateDraft('fdraft-1', { textBody: 'updated note about the attached message' });
    assert.equal(r.id, 'fdraft-2');
    const draft = createdDraftObj(makeReq);
    assert.equal(draft.attachments[0].blobId, 'blob-eml');
    assert.equal(draft[FWD_HEADER_PROP], undefined);
  });

  it('noQuote on an asAttachment draft clears the recorded source (deliberate de-forward)', async () => {
    const makeReq = mockForwardUpdate(client, ASATTACH_FORWARD);
    await client.updateDraft('fdraft-1', { textBody: 'bare note', noQuote: true });
    const draft = createdDraftObj(makeReq);
    assert.equal(draft[FWD_HEADER_PROP], undefined);
  });

  it('a body-less noQuote de-forwards without touching the draft\'s parts', async () => {
    // noQuote on an edit that writes no body has one documented effect — clearing the forward
    // marking — and must keep everything else exactly as it was, embedded images included:
    // there is no body change for the parts to be reconciled against.
    const withImage = {
      ...DUAL_FORWARD,
      attachments: [{
        partId: '2', blobId: 'blob-pic', type: 'image/png', name: 'pic.png',
        cid: 'ii-0123456789abcdef0123456789abcdef@inline.invalid', disposition: 'inline', size: 2048,
      }],
    };
    const makeReq = mockForwardUpdate(client, withImage);
    const result = await client.updateDraft('fdraft-1', { subject: 'Renamed', noQuote: true });
    const draft = createdDraftObj(makeReq);
    assert.equal(draft[FWD_HEADER_PROP], undefined);
    assert.equal(draft.bodyValues.html.value, FWD_HTML); // body untouched
    assert.deepEqual(draft.attachments, [{
      blobId: 'blob-pic', type: 'image/png', name: 'pic.png', disposition: 'inline',
      cid: 'ii-0123456789abcdef0123456789abcdef@inline.invalid',
    }]);
    assert.equal(result.notes, undefined);
  });

  it('keeps the guard inert when the attached message is routed into the body list', async () => {
    // Which list a server puts the .eml in is a MIME-shape accident; the carve-out that keeps
    // an attached-message forward editable has to hold either way.
    const bodyRouted = {
      ...FORWARD_BASE, id: 'fdraft-1',
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [
        { partId: 't', type: 'text/plain' },
        { partId: '2', blobId: 'blob-eml', type: 'message/rfc822', size: 999, name: 'Hello.eml', disposition: 'attachment', cid: null },
      ],
      bodyValues: { t: { value: 'Forwarded message attached.' } },
      attachments: [],
    };
    const makeReq = mockForwardUpdate(client, bodyRouted);
    await client.updateDraft('fdraft-1', { textBody: 'updated note about the attached message' });
    const draft = createdDraftObj(makeReq);
    assert.equal(draft.attachments[0].blobId, 'blob-eml');
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('originalEmailId + noQuote together are rejected even on an unengaged (asAttachment) draft', async () => {
    mockForwardUpdate(client, ASATTACH_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { textBody: 'note', originalEmailId: 'orig-1', noQuote: true }),
      /not both/,
    );
  });

  it('a reply draft QUOTING a forward dispatches to the REPLY variant (quote-prefixed dashed line is not a forward marker)', async () => {
    mockForwardUpdate(client, REPLY_QUOTING_FORWARD);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { textBody: 'new text' }),
      /would drop the quoted original/,
    );
  });

  it('keep with a DIFFERENT original re-points the recorded source to that original (no stale recovery pointer)', async () => {
    // The draft records fwd-orig-msg@…; the caller corrects the source to orig-2. The
    // recreate must carry orig-2's Message-ID — a stale carry would make the guard's
    // advertised recovery rebuild the block for the wrong message on a later edit.
    const ORIG2 = { ...ORIGINAL_FOR_FORWARD, id: 'orig-2', messageId: ['corrected-mid@example.com'] };
    const makeReq = stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        const id = params.ids?.[0];
        if (id === 'orig-2') return { methodResponses: [['Email/get', { list: [ORIG2] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [DUAL_FORWARD] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'fdraft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
    await client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-2' });
    const draft = createdDraftObj(makeReq);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['corrected-mid@example.com']);
  });

  it('keep on a header-less (Q2) draft RECORDS the source, upgrading later challenges to the standard recipe', async () => {
    const makeReq = mockForwardUpdate(client, HEADERLESS_FORWARD);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-1' });
    const draft = createdDraftObj(makeReq);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('keep naming an original with NO settable Message-ID keeps the existing carry (stale beats stripped)', async () => {
    const ORIG_NOMID = { ...ORIGINAL_FOR_FORWARD, id: 'orig-nomid', messageId: undefined };
    const makeReq = stubRequests(client, async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        const id = params.ids?.[0];
        if (id === 'orig-nomid') return { methodResponses: [['Email/get', { list: [ORIG_NOMID] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [DUAL_FORWARD] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'fdraft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
    await client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>', originalEmailId: 'orig-nomid' });
    const draft = createdDraftObj(makeReq);
    assert.deepEqual(draft[FWD_HEADER_PROP], ['fwd-orig-msg@example.com']);
  });

  it('inReplyTo + forward markers but NO reply markers → FORWARD variant engages (no first-match crack on isReply alone)', async () => {
    // A draft with an In-Reply-To header whose body carries only the forward shape:
    // replyGuardArmed requires reply MARKERS too, so dispatch must land on the forward
    // variant — a regression that gated on isReply alone would emit the reply wording.
    const INREPLYTO_FORWARD_BODY = { ...FORWARD_BASE,
      inReplyTo: ['some-reply-target@example.com'], references: ['some-reply-target@example.com'],
      textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { t: { value: FWD_TEXT }, h: { value: FWD_HTML } } };
    mockForwardUpdate(client, INREPLYTO_FORWARD_BODY);
    await assert.rejects(
      () => client.updateDraft('fdraft-1', { htmlBody: '<p>x</p>' }),
      /would drop the forwarded-message block/,
    );
  });
});

// ---------- source-instance header (X-Fastmail-MCP-Source-Id) ----------
// The exact stored instance a draft was composed from: stamped by reply_email /
// forward_email, carried by the edit recreate, consumed by send_draft's keyword
// maintenance. These tests pin the stamp, the carry, and the drop/re-point rules.

describe('source-instance header (X-Fastmail-MCP-Source-Id)', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const SRC_PROP = 'header:X-Fastmail-MCP-Source-Id:asText';
  const FWD_PROP = 'header:X-Forwarded-Message-Id:asMessageIds';
  const SRC_FWD_HTML = '<p>note</p><div><br>----- Original message -----<br>From: Ada Lovelace &lt;ada@example.com&gt;<br>Subject: Hello<br></div><div type="cite"><p>orig body</p></div>';
  const SRC_FWD_TEXT = 'note\n\n\n----- Original message -----\nFrom: Ada Lovelace <ada@example.com>\nSubject: Hello\n\norig text';

  const REPLY_QUOTED = {
    id: 'rdraft-1', subject: 'Re: Hello',
    from: [{ email: 'me@example.com' }], to: [{ email: 'bob@example.com' }], cc: [], bcc: [],
    mailboxIds: { 'mb-drafts': true }, keywords: { $draft: true },
    inReplyTo: ['orig-msg@example.com'], references: ['orig-msg@example.com'],
    [SRC_PROP]: 'orig-1',
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: {
      t: { value: 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, X wrote:\n> quoted' },
      h: { value: '<p>my reply</p><blockquote type="cite">quoted</blockquote>' },
    },
  };
  const FORWARD_WITH_SRC = {
    id: 'fdraft-1', subject: 'Fwd: Hello',
    from: [{ email: 'me@example.com' }], to: [{ email: 'bob@example.com' }], cc: [], bcc: [],
    mailboxIds: { 'mb-drafts': true }, keywords: { $draft: true },
    [FWD_PROP]: ['fwd-orig-msg@example.com'],
    [SRC_PROP]: 'stale-src',
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: SRC_FWD_TEXT }, h: { value: SRC_FWD_HTML } },
  };
  const SRC_ORIGINAL = {
    id: 'orig-1',
    messageId: ['fwd-orig-msg@example.com'],
    from: [{ name: 'Ada Lovelace', email: 'ada@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    subject: 'Hello',
    textBody: [{ partId: 'ot', type: 'text/plain' }],
    htmlBody: [{ partId: 'oh', type: 'text/html' }],
    bodyValues: { ot: { value: 'ORIGINAL TEXT BODY' }, oh: { value: '<p>ORIGINAL HTML BODY</p>' } },
  };

  function mockSrcUpdate(c: JmapClient, draft: any) {
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        if (params.ids?.[0] === 'orig-1') return { methodResponses: [['Email/get', { list: [SRC_ORIGINAL] }, 'email']] };
        return { methodResponses: [['Email/get', { list: [draft] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'new-draft' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { destroyed: params.destroy ?? [] }, 'destroyDraft']] };
    });
  }

  it('createDraft stamps the header from sourceEmailId', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-1' } } }, 'createDraft']],
    }));
    await client.createDraft({ subject: 'Re: x', sourceEmailId: 'orig-1' });
    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj[SRC_PROP], 'orig-1');
  });

  it('createDraft treats a non-JMAP-id sourceEmailId as absent (vetted, never fails the create)', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'email-1' } } }, 'createDraft']],
    }));
    await client.createDraft({ subject: 'Re: x', sourceEmailId: 'not a jmap id!' });
    const emailObj = callArguments(makeReq)[0].methodCalls[0][1].create.draft;
    assert.equal(emailObj[SRC_PROP], undefined);
  });

  it('updateDraft carries the header through a metadata-only edit', async () => {
    const makeReq = mockSrcUpdate(client, REPLY_QUOTED);
    await client.updateDraft('rdraft-1', { subject: 'Re: Hello (edited)' });
    const draft = draftFromCall(makeReq);
    assert.equal(draft[SRC_PROP], 'orig-1');
    assert.deepEqual(draft.inReplyTo, ['orig-msg@example.com']);
  });

  it('KEEPS the header on a reply-draft noQuote (In-Reply-To survives, so does the instance pointer)', async () => {
    const makeReq = mockSrcUpdate(client, REPLY_QUOTED);
    await client.updateDraft('rdraft-1', { htmlBody: '<p>fresh body</p>', noQuote: true });
    const draft = draftFromCall(makeReq);
    assert.equal(draft[SRC_PROP], 'orig-1');
    assert.deepEqual(draft.inReplyTo, ['orig-msg@example.com']);
  });

  it('DROPS the header with the forward marking on a forward-draft noQuote (a de-forward)', async () => {
    const makeReq = mockSrcUpdate(client, FORWARD_WITH_SRC);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>fresh body</p>', noQuote: true });
    const draft = draftFromCall(makeReq);
    assert.equal(draft[FWD_PROP], undefined);
    assert.equal(draft[SRC_PROP], undefined);
  });

  it('re-points the header alongside the forward keep (both name the fetched original)', async () => {
    const makeReq = mockSrcUpdate(client, FORWARD_WITH_SRC);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>new note</p>', originalEmailId: 'orig-1' });
    // The keep path adds an Email/get for the original, shifting the create call —
    // find it by shape rather than by index.
    const createCall = makeReq.mock.calls
      .map((c: any) => c.arguments[0].methodCalls[0][1])
      .find((p: any) => p.create?.draft);
    const draft = createCall.create.draft;
    assert.deepEqual(draft[FWD_PROP], ['fwd-orig-msg@example.com']);
    assert.equal(draft[SRC_PROP], 'orig-1'); // no longer the stale 'stale-src'
  });

  it('a draft that never had the header stays without it', async () => {
    const { [SRC_PROP]: _drop, ...rest } = REPLY_QUOTED as any;
    const makeReq = mockSrcUpdate(client, { ...rest, id: 'rdraft-1' });
    await client.updateDraft('rdraft-1', { subject: 'Re: Hello (edited)' });
    assert.equal(draftFromCall(makeReq)[SRC_PROP], undefined);
  });
});
