import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { homedir } from 'os';
import { resolve, join, basename, sep } from 'path';
import { JmapClient, findBlankBodyPart } from './jmap-client.js';
import type { JmapRequest } from './jmap-client.js';
import { FastmailAuth } from './auth.js';
import { InvalidInputError, PathAccessError } from './coerce.js';
import { bodyHash, collectDraftBodyParts, resolveDraftBodyHash } from './body-hash.js';
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

  // The default save target is resolved by EXACT role, the same question send_draft asks
  // before it will transmit. A substring name match would file the draft into a custom
  // folder and report success, and the send would then refuse it with no move that repairs
  // it — a create surface producing a record the send gate can never accept.
  it('refuses to default-save when no mailbox carries the drafts role, even beside a "Draft notes" folder', async () => {
    mock.method(client, 'getMailboxes', async () => [
      { id: 'mb-notes', name: 'Draft notes', role: null },
    ]);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'x' } } }, 'createDraft']],
    }));

    await assert.rejects(() => client.createDraft({ subject: 'X' }), /drafts.*role/i);
    // Refused before the write: no draft is left behind in the wrong folder.
    assert.equal(makeReq.mock.calls.length, 0);
  });

  // Exact but case-INSENSITIVE, so a server spelling the role differently still resolves
  // by role rather than falling to a name match.
  it('resolves the drafts role whatever case the server spells it in', async () => {
    mock.method(client, 'getMailboxes', async () => [
      { id: 'mb-d', name: 'Entwürfe', role: 'Drafts' },
    ]);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { created: { draft: { id: 'x' } } }, 'createDraft']],
    }));

    await client.createDraft({ subject: 'X' });
    assert.equal(callArguments(makeReq)[0].methodCalls[0][1].create.draft.mailboxIds['mb-d'], true);
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

/**
 * The `bodyHash` `get_email` would issue for a draft fixture — the value a caller reads off
 * that response and passes back to its next body edit.
 *
 * Computed through the same two functions the server uses, deliberately: this is the token's
 * only contract (the same stored body always produces the same string), and a test that
 * spelled out the hash by hand would be pinning the digest rather than the guard.
 */
function hashOf(fixture: any): string {
  return bodyHash(collectDraftBodyParts(fixture));
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
    await client.updateDraft('draft-1', { htmlBody: '<p>Write <code>&lt;p&gt;</code> like this</p>', bodyHash: hashOf(RICH_DRAFT) });
    assert.equal(draftFromCall(makeReq).bodyValues.html.value, '<p>Write <code>&lt;p&gt;</code> like this</p>');
  });

  // ---- one-sided guard + text-fallback regeneration on html edit ----
  //
  // Every body edit below carries `bodyHash: hashOf(<the fixture being served>)`, because a
  // body edit without one is refused outright. The tests that leave it off do so on purpose:
  // they are the ones pinning which guard speaks first, and the guards above this comment all
  // beat the hash check.

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
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>', bodyHash: hashOf(RICH_DRAFT) });
    const draft = draftFromCall(makeReq);
    // The old "The text" is replaced by the fallback regenerated from the NEW html.
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW' }, html: { value: '<p>NEW</p>' } });
  });

  it('writes textBody and drops htmlBody when the partner is named in clearFields', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text', clearFields: ['htmlBody'], bodyHash: hashOf(RICH_DRAFT) });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW text' } });
  });

  it('updates both bodies when both are supplied (no throw)', async () => {
    const makeReq = mockUpdate(client, RICH_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text', htmlBody: '<p>NEW</p>', bodyHash: hashOf(RICH_DRAFT) });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft.bodyValues, {
      text: { value: 'NEW text' },
      html: { value: '<p>NEW</p>' },
    });
  });

  it('writes textBody on a text-only draft (no partner, stays text-only, no throw)', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { textBody: 'NEW text', bodyHash: hashOf(EXISTING_DRAFT) });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.htmlBody, undefined);
    assert.deepEqual(draft.bodyValues, { text: { value: 'NEW text' } });
  });

  it('regenerates the text fallback when htmlBody is edited alone on a text-only draft', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>', bodyHash: hashOf(EXISTING_DRAFT) });
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
    await client.updateDraft('draft-1', { htmlBody: '<div><img src="banner.jpg"></div>', bodyHash: hashOf(EXISTING_DRAFT) });
    const draft = draftFromCall(makeReq);
    assert.equal(draft.textBody, undefined); // no derivable text → no fallback part
    assert.deepEqual(draft.bodyValues, { html: { value: '<div><img src="banner.jpg"></div>' } });
  });

  it('rejects an edited htmlBody that has no visible content (no-body)', async () => {
    mockUpdate(client, EXISTING_DRAFT);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p></p>', bodyHash: hashOf(EXISTING_DRAFT) }),
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
      () => client.updateDraft('draft-1', { clearFields: ['textBody'], bodyHash: hashOf(EXISTING_DRAFT) }),
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
    await client.updateDraft('draft-1', { clearFields: ['htmlBody'], bodyHash: hashOf(RICH_DRAFT) });
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

  // ---- body edits: stored as written, proved by a hash (#37/#42's guard replaced) ----
  //
  // What lived here was a quote-preservation guard: the stored body was scanned for a quote
  // marker, an edit that would drop one was refused, and a kept quote was rebuilt from the
  // original message. All of it is gone. This tool now stores the body it is handed byte for
  // byte — a quote survives an edit because the caller handed it back, and vanishes because
  // the caller did not, with no challenge either way.
  //
  // What replaces it is narrower and mechanical: a body edit must carry the `bodyHash` of
  // the read it was written against. That proves the caller SAW the body it is replacing; it
  // never proves the caller kept any of it. So these fixtures keep the raw Fastmail reply
  // shapes, because "a body with a quote in it is stored with the quote in it, unchanged" is
  // exactly the property to pin. Captured from a live store/fetch round-trip (2026-06-28)
  // and trimmed; the quoted lines carry synthetic content, because nothing reads them.
  const RAW_HTML_QUOTE = '<p>my reply</p><div><br></div><div>On Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:</div><blockquote type="cite" style="margin:0 0 0 .8ex;border-left:1px solid #ccc;padding-left:1ex">\n  1 new planning application near 1 Example Street\n</blockquote>';
  const RAW_TEXT_QUOTE = 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:\n> 2/2 Example St Sampleton NSW 2000: Change of Use and Fitout of a Studio\n> \n> Contact us if you have questions.';

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
  // text-only: the ONE text/plain part aliases into BOTH lists (RFC 8621 §4.1.4), which is
  // why the hash dedupes by part identity rather than counting list entries.
  const TEXT_ONLY_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: RAW_TEXT_QUOTE } } };
  // html-only: the ONE text/html part aliases into both lists (a foreign-client shape).
  const HTML_ONLY_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { h: { value: RAW_HTML_QUOTE } } };
  // A draft whose stored html carries a literal `{{signature}}` inside the quoted original —
  // text somebody else wrote, handed back on every edit. The security case for the flag.
  const PLANTED_TOKEN_REPLY = { ...REPLY_BASE,
    textBody: [{ partId: 't', type: 'text/plain' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: {
      t: { value: 'my reply\n\nOn Sun, Jun 28, 2026, at 12:46 AM, Example Alerts wrote:\n> {{signature}}' },
      h: { value: '<p>my reply</p><blockquote type="cite"><p>{{signature}}</p></blockquote>' },
    } };

  const SIGNING_IDENTITY = {
    id: 'id-1', name: 'Test User', email: 'me@example.com', mayDelete: false,
    textSignature: '-- \nTest User', htmlSignature: '<div>Test User</div>',
  };

  /**
   * Serve `fixture` as the draft, and serve the draft the create call ACTUALLY WROTE back as
   * the re-read of the new id.
   *
   * The re-read is what makes the returned hash a statement about stored bytes. A harness
   * that echoed the caller's arguments instead would let an implementation hashing the sent
   * bytes pass every test here, so the saved email is reconstructed from the create call and
   * from nothing else. `savedOverride` replaces it, which is how the provenance test below
   * makes the two differ on purpose.
   */
  function mockBodyEdit(c: JmapClient, fixture: any, savedOverride?: any) {
    mock.method(c, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    let created: any;
    return mock.method(c, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        if (params.ids?.[0] === 'draft-2') {
          return { methodResponses: [['Email/get', { list: [savedOverride ?? created] }, 'getEmail']] };
        }
        return { methodResponses: [['Email/get', { list: [fixture] }, 'getEmail']] };
      }
      if (params.create) {
        created = params.create.draft;
        return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      }
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
  }

  /** The replacement draft the create call wrote, found by predicate (indices vary). */
  function createdDraft(makeReq: RequestMock) {
    const [request] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][1].create,
      'creating the replacement draft',
    );
    return request.methodCalls[0][1].create.draft;
  }

  // -- the body is stored exactly as written --

  it('stores a quote-bearing body back byte for byte, adding and removing nothing', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    const html = RAW_HTML_QUOTE.replace('<p>my reply</p>', '<p>my EDITED reply</p>');
    const result = await client.updateDraft('draft-1', {
      htmlBody: html, textBody: RAW_TEXT_QUOTE, bodyHash: hashOf(DUAL_REPLY),
    });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, html);
    assert.equal(draft.bodyValues.text.value, RAW_TEXT_QUOTE);
    assert.deepEqual(result.notes, undefined);
  });

  // The behaviour the old guard existed to prevent, now allowed on purpose: the caller is
  // handed the body, so dropping the quote is the caller's edit, not a loss to challenge.
  it('drops a quote the caller did not hand back, with no challenge and no note', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>just my new reply</p>', textBody: 'just my new reply', bodyHash: hashOf(DUAL_REPLY),
    });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>just my new reply</p>');
    assert.equal(/blockquote/.test(draft.bodyValues.html.value), false);
    assert.equal(/wrote:/.test(draft.bodyValues.text.value), false);
    assert.deepEqual(result.notes, undefined);
  });

  // The one thing a quote-dropping edit IS told, and it is about the plain-text part rather
  // than the quote: html alone discards the stored text part, because that part is a derived
  // fallback and is re-derived from the new html. Nothing here challenges the drop.
  it('says the stored plain-text part was discarded when html is written alone', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>just my new reply</p>', bodyHash: hashOf(DUAL_REPLY),
    });
    assert.deepEqual(result.notes, [
      'This edit wrote htmlBody alone, so the draft\'s stored plain-text part was discarded and a fresh fallback derived from your html. If that part was hand-written, supply it as textBody alongside htmlBody.',
    ]);
    assert.equal(createdDraft(makeReq).bodyValues.text.value, 'just my new reply');
  });

  it('stores a text-only draft\'s body as written and stays text-only', async () => {
    const makeReq = mockBodyEdit(client, TEXT_ONLY_REPLY);
    await client.updateDraft('draft-1', { textBody: 'bare reply', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.text.value, 'bare reply');
    assert.equal(draft.htmlBody, undefined);
  });

  it('never fetches the message a reply draft answers', async () => {
    // The rebuild path used to read it on every kept edit. Nothing does now, and a tool that
    // stores what it is handed has no reason to: the assertion is that no Email/get in the
    // whole exchange names anything but the draft and its replacement.
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: hashOf(DUAL_REPLY) });
    for (const call of makeReq.mock.calls) {
      const [method, params] = call.arguments[0].methodCalls[0];
      if (method !== 'Email/get') continue;
      assert.ok(['draft-1', 'draft-2'].includes(params.ids?.[0]), params.ids?.[0]);
    }
  });

  // -- the hash: what it refuses --

  it('refuses a body write with no bodyHash, naming the read that issues one', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>' }),
      (e: unknown) => e instanceof InvalidInputError && /needs bodyHash.*get_email/s.test((e as Error).message),
    );
    // Nothing was written: the refusal is not a partial edit.
    assert.equal(makeReq.mock.calls.some((c) => c.arguments[0].methodCalls[0][1].create), false);
  });

  it('refuses a body CLEAR with no bodyHash — a clear replaces the body too', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['htmlBody'] }),
      /needs bodyHash/,
    );
  });

  it('refuses a hash that is not this draft\'s current one, and writes nothing', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: hashOf(TEXT_ONLY_REPLY) }),
      /not this draft's current one.*Nothing was written/s,
    );
    assert.equal(makeReq.mock.calls.some((c) => c.arguments[0].methodCalls[0][1].create), false);
  });

  it('refuses a blank hash the same way as an absent one', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: '   ' }),
      /needs bodyHash/,
    );
  });

  it('accepts a hash with surrounding whitespace (a value copied out of a response)', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    await client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: `  ${hashOf(DUAL_REPLY)}  ` });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, '<p>x</p>');
  });

  it('needs no hash for a metadata-only edit, which is body-invariant', async () => {
    const makeReq = mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', { subject: 'Re: Hello (edited)' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.text.value, RAW_TEXT_QUOTE);
    assert.equal(draft.bodyValues.html.value, RAW_HTML_QUOTE);
    // Neither half of the pair: nothing was promised, so there is nothing to withhold.
    assert.equal(result.bodyHash, undefined);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  it('needs no hash for an attachment-only edit either', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', { removeAttachments: [] });
    assert.equal(result.bodyHash, undefined);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  // The compose side issues none, in any mode. createDraft's whole result is the new id —
  // there is nowhere for a hash to ride, and that is the design: a hash certifies that the
  // caller SAW the stored body, and a compose result can only ever echo the bytes it just
  // sent, which proves nothing about what the server stored. The caller reads the draft.
  it('createDraft hands back an id and nothing else — no compose path issues a hash', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { created: { draft: { id: 'email-42' } } }, 'createDraft']],
    });
    const result = await client.createDraft({ subject: 'Hello', textBody: 'body' });
    assert.equal(result, 'email-42');
    assert.equal(typeof result, 'string');
  });

  // -- the hash: provenance --

  // THE test for the returned hash, and the only one that can tell the two implementations
  // apart. The saved draft is served back with DIFFERENT bytes from the ones the call sent;
  // a hash computed over the sent bytes would match hashOf(sent) and pass every other
  // assertion in this file. The result must carry the re-read's hash.
  it('returns the hash of the RE-READ, never of the bytes the call sent', async () => {
    const sentHtml = '<p>what the caller wrote</p>';
    const saved = { ...REPLY_BASE, id: 'draft-2',
      textBody: [{ partId: 'text', type: 'text/plain' }],
      htmlBody: [{ partId: 'html', type: 'text/html' }],
      bodyValues: { text: { value: 'what the SERVER stored' }, html: { value: '<p>what the SERVER stored</p>' } } };
    mockBodyEdit(client, HTML_ONLY_REPLY, saved);

    const result = await client.updateDraft('draft-1', { htmlBody: sentHtml, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(result.bodyHash, hashOf(saved));
    assert.notEqual(result.bodyHash, hashOf({
      textBody: [{ partId: 'html', type: 'text/html' }],
      htmlBody: [{ partId: 'html', type: 'text/html' }],
      bodyValues: { html: { value: sentHtml } },
    }));
  });

  it('the hash it returns is the one the caller\'s next body edit is accepted with', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const first = await client.updateDraft('draft-1', { htmlBody: '<p>one</p>', bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.ok(first.bodyHash);
    // Re-point the harness at the draft the first edit saved, then edit again with the hash
    // it handed back — the run-of-edits case the return exists for.
    const savedFirst = createdDraft(makeReq);
    mockBodyEdit(client, { ...savedFirst, id: 'draft-1', keywords: { $draft: true }, mailboxIds: { 'mb-drafts': true } });
    const second = await client.updateDraft('draft-1', { htmlBody: '<p>two</p>', bodyHash: first.bodyHash });
    assert.ok(second.bodyHash);
  });

  it('withholds the hash with a reason when the re-read fails, never falling back to the sent bytes', async () => {
    mock.method(client, 'getMailboxes', async () => MAILBOXES_WITH_TRASH);
    mock.method(client, 'makeRequest', async (req: any) => {
      const [method, params] = req.methodCalls[0];
      if (method === 'Email/get') {
        if (params.ids?.[0] === 'draft-2') return { methodResponses: [['Email/get', { list: [] }, 'getEmail']] };
        return { methodResponses: [['Email/get', { list: [HTML_ONLY_REPLY] }, 'getEmail']] };
      }
      if (params.create) return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } } }, 'createDraft']] };
      return { methodResponses: [['Email/set', { updated: { 'draft-1': null } }, 'trashOldDraft']] };
    });
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /re-reading the saved draft.*failed/s);
    // The edit itself still landed: a hash that cannot be issued is a degrade, not a failure.
    assert.equal(result.id, 'draft-2');
  });

  it('withholds the hash when the re-read comes back truncated', async () => {
    const truncated = { ...REPLY_BASE, id: 'draft-2',
      textBody: [{ partId: 'text', type: 'text/plain' }],
      bodyValues: { text: { value: 'partial', isTruncated: true } } };
    mockBodyEdit(client, TEXT_ONLY_REPLY, truncated);
    const result = await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /truncated or as having encoding problems/);
    assert.match(result.bodyHashWithheld!, /the saved draft was re-read/);
  });

  // The saved body is judged by the READ side's rule, so a shape get_email refuses to hash
  // is one edit_draft refuses to hash too. Before, this draft got a hash here and "recreate
  // the draft" from the very next read of the same saved object.
  it('withholds the hash when the saved draft carries a body part no read returns', async () => {
    const mismatched = { ...REPLY_BASE, id: 'draft-2',
      textBody: [{ partId: 'text', type: 'text/plain' }, { partId: 'stray', type: 'text/html' }],
      bodyValues: { text: { value: 'x' }, stray: { value: '<p>hidden</p>' } } };
    mockBodyEdit(client, TEXT_ONLY_REPLY, mismatched);
    const result = await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /a body part no read returns/);
    assert.match(result.bodyHashWithheld!, /Recreate the draft/);
    // The same reason the read side gives for the same saved object, said on this surface.
    const readSide = resolveDraftBodyHash(mismatched, { bodyText: true, bodyHtml: true, stripQuoted: false })!;
    assert.ok('bodyHashWithheld' in readSide);
    assert.ok(result.bodyHashWithheld!.endsWith((readSide as { bodyHashWithheld: string }).bodyHashWithheld));
  });

  // The other half of the same routing: an empty saved part set is a body, not a degraded
  // read. get_email hashes it, so this does too — the old predicate reported it as the
  // server having flagged truncation, which it had not.
  it('issues the hash when the saved draft comes back with no body parts at all', async () => {
    const empty = { ...REPLY_BASE, id: 'draft-2', textBody: [], htmlBody: [], bodyValues: {} };
    mockBodyEdit(client, TEXT_ONLY_REPLY, empty);
    const result = await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    assert.equal(result.bodyHashWithheld, undefined);
    assert.equal(result.bodyHash, hashOf(empty));
  });

  // THE pin on the routing, and it is not implied by the shapes above. resolveDraftBodyHash
  // takes a descriptor of what a read SHOWED, and two of its branches exist only because a
  // read can be field-scoped or quote-stripped. This path controls its own re-read and so
  // claims a whole one; passing anything less would make an edit that has no `fields` at all
  // withhold with a reason about the caller's `fields` selection. This stays meaningful as
  // the read side grows branches, which is exactly when it would otherwise break in silence.
  it('never withholds with a reason about the read\'s scope, whatever the saved draft looks like', async () => {
    const shapes: Record<string, any> = {
      truncated: { textBody: [{ partId: 't', type: 'text/plain' }],
        bodyValues: { t: { value: 'partial', isTruncated: true } } },
      encodingProblem: { textBody: [{ partId: 't', type: 'text/plain' }],
        bodyValues: { t: { value: 'partial', isEncodingProblem: true } } },
      unreturnablePart: { textBody: [{ partId: 't', type: 'text/plain' }, { partId: 's', type: 'text/html' }],
        bodyValues: { t: { value: 'x' }, s: { value: '<p>hidden</p>' } } },
      emptyPartSet: { textBody: [], htmlBody: [], bodyValues: {} },
      htmlOnly: { htmlBody: [{ partId: 'h', type: 'text/html' }], bodyValues: { h: { value: '<p>x</p>' } } },
    };
    for (const [name, shape] of Object.entries(shapes)) {
      mockBodyEdit(client, TEXT_ONLY_REPLY, { ...REPLY_BASE, id: 'draft-2', ...shape });
      const result = await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
      const withheld = result.bodyHashWithheld ?? '';
      assert.equal(/fields:/.test(withheld), false, name);
      assert.equal(/stripQuoted/.test(withheld), false, name);
      assert.equal(/verbose/.test(withheld), false, name);
      assert.equal(/did not return the draft's stored body whole/.test(withheld), false, name);
    }
  });

  // resolveDraftBodyHash answers for drafts and returns nothing for anything else, so a
  // re-read that does not come back as one would drop the promised field with no trace.
  it('withholds with a reason when the saved message does not read back as a draft', async () => {
    const notADraft = { ...REPLY_BASE, id: 'draft-2', keywords: {},
      textBody: [{ partId: 't', type: 'text/plain' }], bodyValues: { t: { value: 'x' } } };
    mockBodyEdit(client, TEXT_ONLY_REPLY, notADraft);
    const result = await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /did not read back as a draft/);
  });

  it('asks the server for the keywords that decide it, so the re-read can be judged at all', async () => {
    const makeReq = mockBodyEdit(client, TEXT_ONLY_REPLY);
    await client.updateDraft('draft-1', { textBody: 'x', bodyHash: hashOf(TEXT_ONLY_REPLY) });
    const [reRead] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][0] === 'Email/get' && req.methodCalls[0][1].ids?.[0] === 'draft-2',
      're-reading the saved draft',
    );
    assert.ok(reRead.methodCalls[0][1].properties.includes('keywords'));
  });

  it('withholds the hash when the governing part is one this server derived', async () => {
    // clearFields:['htmlBody'] with no body supplied leaves the draft's own stored text
    // standing. The caller did not write those bytes, so it has not read what now governs.
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      clearFields: ['htmlBody'], bodyHash: hashOf(DUAL_REPLY),
    });
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /derived from html rather than written by you/);
  });

  it('issues the hash on the same clear when the caller supplies the surviving text', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      clearFields: ['htmlBody'], textBody: 'my plain-text reply', bodyHash: hashOf(DUAL_REPLY),
    });
    assert.ok(result.bodyHash);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  // A body part carrying no stored VALUE is not a degraded read. It is an embedded image the
  // server routed into a body list, which the hash covers with a sentinel of its own — and
  // the read side treats it exactly that way, so counting it as degradation here would make
  // get_email and edit_draft give different answers about the identical saved draft, under a
  // note claiming the server flagged truncation when it flagged nothing.
  it('issues the hash when the saved draft carries a body part with no stored value', async () => {
    const savedWithImagePart = { ...REPLY_BASE, id: 'draft-2',
      textBody: [{ partId: 'text', type: 'text/plain' }],
      htmlBody: [
        { partId: 'html', type: 'text/html' },
        { blobId: 'blob-img', type: 'image/png', cid: 'img@example.com', disposition: 'inline' },
      ],
      bodyValues: { text: { value: '[image]' }, html: { value: '<p>x</p><img src="cid:img@example.com">' } } };
    mockBodyEdit(client, HTML_ONLY_REPLY, savedWithImagePart);

    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>x</p>', bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    assert.equal(result.bodyHashWithheld, undefined);
    assert.equal(result.bodyHash, hashOf(savedWithImagePart));
    // The two sides agree on that draft, which is the property the predicate exists to keep.
    assert.deepEqual(
      resolveDraftBodyHash(savedWithImagePart, { bodyText: true, bodyHtml: true, stripQuoted: false }),
      { bodyHash: result.bodyHash },
    );
  });

  // Clearing htmlBody on a draft that never HAD html leaves the caller's own stored text
  // standing — not a part derived from html — and this call proved it read those bytes.
  it('issues the hash when an html clear leaves the caller\'s own text on a text-only draft', async () => {
    mockBodyEdit(client, TEXT_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      clearFields: ['htmlBody'], bodyHash: hashOf(TEXT_ONLY_REPLY),
    });
    assert.equal(result.bodyHashWithheld, undefined);
    assert.ok(result.bodyHash);
  });

  it('issues the hash on an html-alone edit, whose derived text part is not the governing one', async () => {
    // The governing part is the html when the message ships one. An html-alone edit derives
    // a fresh text fallback, but that part is not what a later edit has to hand back.
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', { htmlBody: '<p>x</p>', bodyHash: hashOf(DUAL_REPLY) });
    assert.ok(result.bodyHash);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  // -- expandSignature --

  it('expands a {{signature}} the caller wrote when the flag is passed', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Thanks.</p>{{signature}}', expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    const html = createdDraft(makeReq).bodyValues.html.value;
    assert.match(html, /Test User/);
    assert.equal(/\{\{signature\}\}/.test(html), false);
    // An expansion rewrote the body, so what landed is not what the caller wrote.
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld!, /expanded \{\{signature\}\}/);
  });

  // A flagged call whose identity has no sign-off REMOVES the token and says why. It is not
  // a refusal: the flag was honoured, there was simply nothing to put there, and leaving the
  // braces in a body the caller declared its own would ship them to the recipient.
  it('removes the token and names the cause when the identity has no signature', async () => {
    mock.method(client, 'getIdentities', async () => [{ ...SIGNING_IDENTITY, textSignature: '', htmlSignature: '' }]);
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Thanks.</p>{{signature}}', expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, '<p>Thanks.</p>');
    assert.ok(result.notes?.some((n) => /the sending identity has no signature configured/.test(n)));
  });

  // Both parts supplied, the token in only one: the other ships unsigned, and a recipient
  // reading that alternative sees no sign-off. Said out loud, because the body that ships is
  // the caller's and nothing here will add the missing one.
  it('says so when one supplied part took the sign-off and the other carried no token', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Thanks.</p>{{signature}}', textBody: 'Thanks.',
      expandSignature: true, bodyHash: hashOf(DUAL_REPLY),
    });
    assert.ok(result.notes?.some((n) => /textBody carries none/.test(n)));
  });

  // The same shape with NOTHING to expand. The other part carried the token, but the identity
  // has no sign-off, so the token was removed and no sign-off landed anywhere — saying it
  // "expanded in htmlBody" beside the note saying it could not would be a receipt for an
  // event that did not happen, and the two notes would contradict each other.
  it('does not claim the other part expanded when nothing expanded there', async () => {
    mock.method(client, 'getIdentities', async () => [{ ...SIGNING_IDENTITY, textSignature: '', htmlSignature: '' }]);
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>Thanks.</p>{{signature}}', textBody: 'Thanks.',
      expandSignature: true, bodyHash: hashOf(DUAL_REPLY),
    });
    assert.ok(result.notes?.some((n) => /has no signature configured/.test(n)));
    assert.equal(result.notes?.some((n) => /expanded in htmlBody/.test(n)), false);
  });

  // -- a {{…}} spelling that expands to nothing --
  // edit_draft stores the body as written, so a mistyped token ships with its braces showing.
  // The compose tool reports these on its receipt; this tool said nothing, on the tool where
  // the caller is likelier to mistype one because it is hand-editing an existing body.

  it('reports a {{…}} spelling this edit introduced, which is not a token and ships as written', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = '<p>Thanks.</p><p>{{sig}}</p>';
    const result = await client.updateDraft('draft-1', { htmlBody: body, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, body);
    const note = result.notes?.find((n) => /\{\{sig\}\}/.test(n));
    assert.ok(note, `expected a note naming {{sig}}, got ${JSON.stringify(result.notes)}`);
    assert.match(note!, /stored as written/);
  });

  // The count in the sentence and the list beside it must count the SAME thing. One typo
  // written twice is one spelling to fix, and a note saying "2 … ("{{sig}}")" sends the
  // caller hunting for a second spelling that is not there.
  it('counts one repeated spelling once, within a single part', async () => {
    mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{sig}}</p><p>{{sig}}</p>', bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    const note = result.notes?.find((n) => /\{\{sig\}\}/.test(n));
    assert.ok(note, `expected a note, got ${JSON.stringify(result.notes)}`);
    assert.match(note!, /carries 1 \{\{…\}\} spelling this edit added/);
    assert.equal(/and \d+ more/.test(note!), false, `implied more spellings than exist: ${note}`);
  });

  // ACROSS the two parts, which is the ordinary path: supplying both bodies means writing the
  // same typo twice, so this is the common case rather than an unusual one.
  it('counts one repeated spelling once, across both written parts', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{sig}}</p>', textBody: '{{sig}}', bodyHash: hashOf(DUAL_REPLY),
    });
    const note = result.notes?.find((n) => /\{\{sig\}\}/.test(n));
    assert.ok(note, `expected a note, got ${JSON.stringify(result.notes)}`);
    assert.match(note!, /carries 1 \{\{…\}\} spelling this edit added/);
    assert.equal(/and \d+ more/.test(note!), false, `implied more spellings than exist: ${note}`);
  });

  // Two DIFFERENT spellings are two things to fix and must still count as two.
  it('counts two distinct spellings as two', async () => {
    mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{sig}}</p><p>{{quotes}}</p>', bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    const note = result.notes?.find((n) => /\{\{sig\}\}/.test(n));
    assert.ok(note, `expected a note, got ${JSON.stringify(result.notes)}`);
    assert.match(note!, /carries 2 \{\{…\}\} spellings this edit added/);
    assert.match(note!, /\{\{quotes\}\}/);
  });

  // The flag expands {{signature}} and nothing else, so a mistyped spelling ships from a
  // flagged edit exactly as it does from an unflagged one and is reported on both.
  it('reports the spelling on a flagged edit too', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{signature}}</p><p>{{sig}}</p>',
      expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    assert.ok(result.notes?.some((n) => /\{\{sig\}\}/.test(n)));
  });

  // COUNT-RISE, the rule that keeps this from nagging. A spelling already in the stored body
  // is text someone else wrote, handed back on every edit; reporting it would fire on every
  // edit of that draft forever and say nothing about what this call did.
  it('says nothing about a spelling the stored body already carried', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p>hi</p><blockquote><p>{{sig}}</p></blockquote>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>hi, edited</p><blockquote><p>{{sig}}</p></blockquote>',
      bodyHash: hashOf(planted),
    });
    assert.equal(result.notes?.some((n) => /\{\{sig\}\}/.test(n)) ?? false, false);
  });

  // ...and a SECOND copy of that same spelling is this edit's doing, so the count rise is
  // reported even though the spelling itself is not new.
  it('reports a second copy of a spelling the stored body carried once', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p>{{sig}}</p>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{sig}}</p><p>{{sig}}</p>', bodyHash: hashOf(planted),
    });
    assert.ok(result.notes?.some((n) => /\{\{sig\}\}/.test(n)));
  });

  it('refuses the flag when the body carries no {{signature}} to expand', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        htmlBody: '<p>Thanks.</p>', expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY),
      }),
      /carries no \{\{signature\}\}/,
    );
    // Refused before anything is written, the same as its sibling below: the caller's body
    // was otherwise perfectly storable, so an implementation that stored it and complained
    // afterwards would leave a draft rewritten by a call that reported failure.
    assert.equal(makeReq.mock.calls.some((c) => c.arguments[0].methodCalls[0][1].create), false);
  });

  // THE PAIR. Same body, same draft, the flag the only difference. A table built only from
  // flagged inputs would read green on an implementation that refused both.
  it('UNFLAGGED: two bare {{signature}} the stored body did not have are stored, with a note naming the flag', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = '<p>{{signature}}</p><p>and again {{signature}}</p>';
    const result = await client.updateDraft('draft-1', { htmlBody: body, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, body);
    assert.deepEqual(result.notes, [
      'htmlBody carries 2 {{signature}} tokens the stored body did not, and this edit stored ' +
      'the body as written. Pass expandSignature: true to expand it.',
    ]);
  });

  it('FLAGGED: the identical body is refused, naming the count and the escape', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = '<p>{{signature}}</p><p>and again {{signature}}</p>';
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: body, expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY) }),
      (e: unknown) => {
        const m = (e as Error).message;
        assert.match(m, /htmlBody carries 2 \{\{signature\}\} tokens/);
        assert.match(m, /\\\{\{signature\}\}/);
        return true;
      },
    );
    // Refused BEFORE anything expands: no create call, so nothing was written and no block
    // was built. An oracle that only asked "was it refused" would miss this.
    assert.equal(makeReq.mock.calls.some((c) => c.arguments[0].methodCalls[0][1].create), false);
  });

  // The escape, under the flag: the backslash is consumed and the bare braces are stored.
  // NO note is emitted for it, and that is what this pins. The escape notes exist to tell an
  // UNFLAGGED caller that its backslash shipped as literal text, which is not what happened
  // here — under the flag the escape did exactly what it is for.
  it('consumes the escape under the flag and stores the bare token, silently', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: String.raw`<p>write \{{signature}} to mean the token</p>{{signature}}`,
      expandSignature: true, bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    const html = createdDraft(makeReq).bodyValues.html.value;
    assert.match(html, /write \{\{signature\}\} to mean the token/);
    assert.match(html, /Test User/); // the unescaped one still expanded
    assert.equal(result.notes, undefined); // the escape note is for UNFLAGGED edits
  });

  // Unflagged, the backslash is not consumed — the body is stored byte for byte, backslash
  // included — so the caller is told, because it almost certainly meant the braces alone.
  it('reports an escape the caller added on an UNFLAGGED edit, where the backslash ships', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = String.raw`<p>write \{{signature}} to mean the token</p>`;
    const result = await client.updateDraft('draft-1', { htmlBody: body, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, body);
    assert.ok(result.notes?.some((n) => /an escaped token spelling the stored body did not/.test(n)));
  });

  // The converse, and the half the description used to state unconditionally. The escape
  // notes are keyed on a COUNT RISE against the stored bytes, exactly as the {{signature}}
  // note is, so an escape the body handed back already carried ships in silence — otherwise
  // the original author's text would be reported back on every edit of that draft forever.
  it('says nothing about an escape the body handed back already carried', async () => {
    const stored = String.raw`<p>they wrote \{{signature}} in the original</p>`;
    const carriesEscape = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: stored } } };
    const makeReq = mockBodyEdit(client, carriesEscape);
    const result = await client.updateDraft('draft-1', { htmlBody: stored, bodyHash: hashOf(carriesEscape) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, stored);
    assert.equal((result.notes ?? []).some((n) => /an escaped token spelling/.test(n)), false);
  });


  // A near-miss is REPORTED here rather than refused, which is the opposite of draft_email.
  // The body may be a foreign one handed back, so a refusal keyed on its text could be
  // planted by the original's author and would recur on every edit of that draft.
  it('reports a near-miss spelling rather than refusing it', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = '<p>{{Signature}} goes here</p>';
    const result = await client.updateDraft('draft-1', { htmlBody: body, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, body);
    assert.ok(result.notes?.some((n) => /\{\{Signature\}\}/.test(n)));
  });

  // The security case for the flag rather than the token: the second token here is one the
  // ORIGINAL's author wrote, sitting inside the quoted history the caller hands back.
  it('stores a planted {{signature}} unchanged and in silence on an unflagged edit', async () => {
    const makeReq = mockBodyEdit(client, PLANTED_TOKEN_REPLY);
    const stored = PLANTED_TOKEN_REPLY.bodyValues.h.value;
    const result = await client.updateDraft('draft-1', {
      htmlBody: stored, textBody: PLANTED_TOKEN_REPLY.bodyValues.t.value,
      bodyHash: hashOf(PLANTED_TOKEN_REPLY),
    });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, stored);
    // No note: the count did not rise, so nothing the caller did is being reported back.
    assert.deepEqual(result.notes, undefined);
  });

  it('stores {{quote}} and {{forward}} as literal text and says so — this tool places no history', async () => {
    mockBodyEdit(client, HTML_ONLY_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>see {{quote}} and {{forward}}</p>', bodyHash: hashOf(HTML_ONLY_REPLY),
    });
    assert.equal((result.notes ?? []).some((n) => /\{\{quote\}\} and \{\{forward\}\} were stored as written/.test(n)), true);
  });

  // The other direction, for both text-keyed notes above. A {{quote}} or a {{Signature}} the
  // STORED body already carried is text the original's author wrote, handed back on every
  // edit: reporting it says nothing about what this call did, cannot be acted on (the caller
  // does not own those words), and is plantable by whoever composed the original.

  it('says nothing about a history token the stored body already carried', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p>hi</p><blockquote><p>see {{quote}}</p></blockquote>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>hi, edited</p><blockquote><p>see {{quote}}</p></blockquote>',
      bodyHash: hashOf(planted),
    });
    assert.equal((result.notes ?? []).some((n) => /\{\{quote\}\}/.test(n)), false);
  });

  it('still reports a history token this edit added beside one the body already had', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p><blockquote>{{quote}}</blockquote></p>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>see {{forward}}</p><blockquote>{{quote}}</blockquote>', bodyHash: hashOf(planted),
    });
    const note = (result.notes ?? []).find((n) => /stored as written: those tokens expand/.test(n));
    assert.ok(note, `expected a history note, got ${JSON.stringify(result.notes)}`);
    // Only the one this edit introduced is named; the carried-over {{quote}} is not.
    assert.match(note!, /\{\{forward\}\} was stored as written/);
  });

  it('says nothing about a near-miss the stored body already carried', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p>hi</p><blockquote><p>{{Signature}}</p></blockquote>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>hi, edited</p><blockquote><p>{{Signature}}</p></blockquote>',
      bodyHash: hashOf(planted),
    });
    assert.equal((result.notes ?? []).some((n) => /\{\{Signature\}\}/.test(n)), false);
  });

  it('still reports a second near-miss this edit added beside one the body already had', async () => {
    const planted = { ...REPLY_BASE,
      textBody: [{ partId: 'h', type: 'text/html' }], htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<p>{{Signature}}</p>' } } };
    mockBodyEdit(client, planted);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{Signature}}</p><p>{{Signature}}</p>', bodyHash: hashOf(planted),
    });
    assert.ok((result.notes ?? []).some((n) => /\{\{Signature\}\}/.test(n)));
  });

  // -- one note per thing, not one per part --
  // The escape, near-miss and history notes name a token and never a part, so the same
  // spelling reaching them from both written bodies is the identical sentence twice, and
  // nothing downstream deduplicates the notes list. Supplying both bodies is what this
  // tool's description tells a caller to do, so this is the instructed path.

  const notesMatching = (result: any, re: RegExp): string[] =>
    (result.notes ?? []).filter((n: string) => re.test(n));

  it('reports an escape written into both parts once', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: String.raw`<p>write \{{signature}} here</p>`,
      textBody: String.raw`write \{{signature}} here`,
      bodyHash: hashOf(DUAL_REPLY),
    });
    assert.equal(notesMatching(result, /an escaped token spelling the stored body did not/).length, 1);
  });

  it('reports a near-miss written into both parts once', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>{{Signature}} goes here</p>',
      textBody: '{{Signature}} goes here',
      bodyHash: hashOf(DUAL_REPLY),
    });
    assert.equal(notesMatching(result, /\{\{Signature\}\}/).length, 1);
  });

  it('reports a history token written into both parts once', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>see {{quote}}</p>',
      textBody: 'see {{quote}}',
      bodyHash: hashOf(DUAL_REPLY),
    });
    assert.equal(notesMatching(result, /stored as written: those tokens expand/).length, 1);
  });

  // Deduplicating across parts must not silence a token that only ONE part introduced: two
  // different history tokens, one per body, are two things to say.
  it('still names a history token that only one of the two parts added', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    const result = await client.updateDraft('draft-1', {
      htmlBody: '<p>see {{quote}}</p>',
      textBody: 'see {{forward}}',
      bodyHash: hashOf(DUAL_REPLY),
    });
    const named = notesMatching(result, /stored as written: those tokens expand/).join(' ');
    assert.match(named, /\{\{quote\}\}/);
    assert.match(named, /\{\{forward\}\}/);
  });

  it('leaves a {{quote}} in the body it stores rather than removing it', async () => {
    const makeReq = mockBodyEdit(client, HTML_ONLY_REPLY);
    const body = '<p>see {{quote}} here</p>';
    await client.updateDraft('draft-1', { htmlBody: body, bodyHash: hashOf(HTML_ONLY_REPLY) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, body);
  });

  // -- the refusal order, over inputs that trip more than one guard --

  it('coupling guard beats a stale hash: the caller is told the SHAPE to fix', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { textBody: 'new text', bodyHash: 'bh1-not-this-draft' }),
      /editing textBody alone won't change what most recipients see/,
    );
  });

  it('coupling guard beats a stale hash on a clear, too', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { clearFields: ['textBody'], bodyHash: 'bh1-not-this-draft' }),
      /textBody can't be cleared on its own while htmlBody is present/,
    );
  });

  it('a stale hash beats the flagged count refusal: re-read first, then worry about tokens', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    mockBodyEdit(client, HTML_ONLY_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        htmlBody: '<p>{{signature}}</p><p>{{signature}}</p>',
        expandSignature: true, bodyHash: 'bh1-not-this-draft',
      }),
      /not this draft's current one/,
    );
  });

  it('assertBodyInputs beats the coupling guard: a malformed body is the first thing to fix', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<![CDATA[<p>Hi</p>]]>', clearFields: ['textBody'] }),
      (e: unknown) => {
        const m = (e as Error).message;
        assert.match(m, /htmlBody contains a CDATA section/);
        assert.doesNotMatch(m, /can't be cleared on its own/);
        return true;
      },
    );
  });

  it('assertBodyInputs beats the hash check as well', async () => {
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', { htmlBody: '<![CDATA[<p>Hi</p>]]>', bodyHash: 'bh1-not-this-draft' }),
      (e: unknown) => {
        const m = (e as Error).message;
        assert.match(m, /htmlBody contains a CDATA section/);
        assert.doesNotMatch(m, /bodyHash/);
        return true;
      },
    );
  });

  it('coupling guard beats the flagged count refusal', async () => {
    mock.method(client, 'getIdentities', async () => [SIGNING_IDENTITY]);
    mockBodyEdit(client, DUAL_REPLY);
    await assert.rejects(
      () => client.updateDraft('draft-1', {
        textBody: '{{signature}} and {{signature}}',
        expandSignature: true, bodyHash: hashOf(DUAL_REPLY),
      }),
      /editing textBody alone won't change what most recipients see/,
    );
  });

  // -- clearFields:['forwardedMessageId'] --

  const FORWARD_DRAFT = { ...REPLY_BASE,
    inReplyTo: null, references: null,
    'header:X-Forwarded-Message-Id:asMessageIds': ['orig-msg@example.com'],
    textBody: [{ partId: 't', type: 'text/plain' }],
    bodyValues: { t: { value: 'FYI\n\n----- Original message -----\nFrom: someone' } } };

  it('de-forwards a draft, dropping the recorded X-Forwarded-Message-Id', async () => {
    const makeReq = mockBodyEdit(client, FORWARD_DRAFT);
    await client.updateDraft('draft-1', { clearFields: ['forwardedMessageId'] });
    const draft = createdDraft(makeReq);
    assert.equal(draft['header:X-Forwarded-Message-Id:asMessageIds'], undefined);
  });

  it('carries the header forward when the clear is not asked for', async () => {
    const makeReq = mockBodyEdit(client, FORWARD_DRAFT);
    await client.updateDraft('draft-1', { subject: 'Fwd: Hello (edited)' });
    const draft = createdDraft(makeReq);
    assert.deepEqual(draft['header:X-Forwarded-Message-Id:asMessageIds'], ['orig-msg@example.com']);
  });

  it('de-forwarding is metadata: it needs no bodyHash and leaves the body alone', async () => {
    const makeReq = mockBodyEdit(client, FORWARD_DRAFT);
    const result = await client.updateDraft('draft-1', { clearFields: ['forwardedMessageId'] });
    assert.equal(createdDraft(makeReq).bodyValues.text.value, FORWARD_DRAFT.bodyValues.t.value);
    assert.equal(result.bodyHash, undefined);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  it('de-forwards alongside a body edit, which still needs its hash', async () => {
    const makeReq = mockBodyEdit(client, FORWARD_DRAFT);
    await client.updateDraft('draft-1', {
      clearFields: ['forwardedMessageId'], textBody: 'just my note', bodyHash: hashOf(FORWARD_DRAFT),
    });
    const draft = createdDraft(makeReq);
    assert.equal(draft['header:X-Forwarded-Message-Id:asMessageIds'], undefined);
    assert.equal(draft.bodyValues.text.value, 'just my note');
  });

  // -- non-reply draft --

  it('treats a plain (non-reply) draft exactly the same — the hash is not about replies', async () => {
    const makeReq = mockBodyEdit(client, EXISTING_DRAFT);
    await client.updateDraft('draft-1', { htmlBody: '<p>NEW</p>', bodyHash: hashOf(EXISTING_DRAFT) });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, '<p>NEW</p>');
  });
});

// ---------- sendDraft ----------

// In the Drafts folder, because send_draft only sends a draft that is filed there. A
// fixture without mailboxIds is an under-stubbed read, not a draft the send should accept.
const SENDABLE_DRAFT = {
  id: 'draft-1',
  from: [{ email: 'me@example.com' }],
  to: [{ email: 'bob@example.com' }],
  cc: [{ email: 'cc@example.com' }],
  bcc: [],
  keywords: { $draft: true },
  mailboxIds: { 'mb-drafts': true },
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

  // ---- a draft is sendable only from the Drafts folder ----

  it('refuses to send a draft that is not in the Drafts folder, and names the repair', async () => {
    const archived = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-archive': true } };
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [archived] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /not in the Drafts folder/i);
        assert.match(err.message, /move_email/);
        return true;
      },
    );
    // Refused before submission: only the Email/get went out.
    assert.equal(makeReq.mock.calls.length, 1);
  });

  // The check needs the property, and nothing else in this method reads it, so a read that
  // stopped asking for it would make every draft look filed nowhere.
  it('asks for mailboxIds on the pre-send read', async () => {
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    await client.sendDraft('draft-1');
    assert.ok(callArguments(makeReq)[0].methodCalls[0][1].properties.includes('mailboxIds'));
  });

  it('sends a draft that is in Drafts alongside another mailbox', async () => {
    const alsoLabelled = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-drafts': true, 'mb-archive': true } };
    stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [alsoLabelled] }, 'getEmail']] };
      }
      return { methodResponses: [['EmailSubmission/set', { created: { submission: { id: 'sub-1' } } }, 'submitDraft']] };
    });

    assert.equal((await client.sendDraft('draft-1')).submissionId, 'sub-1');
  });

  it('refuses when no mailbox carries the drafts role', async () => {
    mock.method(client, 'getMailboxes', async () => [SENT_MAILBOX]);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [SENDABLE_DRAFT] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /drafts/i);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 1);
  });

  // The membership map is parsed from the response and read by a caller-independent id, so
  // it is read the way this file's other mailboxIds reads are: an own-property test, not an
  // index. A drafts-role mailbox whose id collides with an Object.prototype key would
  // otherwise resolve to a truthy function and open the gate for a draft filed anywhere.
  it('does not treat a prototype key as membership of Drafts', async () => {
    mock.method(client, 'getMailboxes', async () => [
      { id: 'constructor', name: 'Drafts', role: 'drafts' },
      SENT_MAILBOX,
    ]);
    const archived = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-archive': true } };
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [archived] }, 'getEmail']],
    }));

    await assert.rejects(() => client.sendDraft('draft-1'), /not in the Drafts folder/i);
    assert.equal(makeReq.mock.calls.length, 1);
  });

  // The locations the refusal names, read out of its parenthetical as a LIST. Object.keys
  // order is not something this code fixes, so a pattern like /Drafts,|: Drafts/ passes or
  // fails on where a name happened to land rather than on whether it was named at all.
  // `mb-archive` has no mailbox in this suite's fixture, so it renders as its own id — which
  // is also the fallback these assertions pin.
  const locationsNamed = (message: string): string[] => {
    const listed = /\(it is in: ([^)]*)\)/.exec(message);
    return listed ? listed[1].split(',').map((s) => s.trim()) : [];
  };

  // A `false` value is not a membership, so it must not be reported as one. Refusing because
  // the draft is not in Drafts and then telling the caller it IS in Drafts is a refusal that
  // argues against itself.
  it('does not name a mailbox whose membership value is false', async () => {
    const notReally = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-drafts': false, 'mb-archive': true } };
    stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [notReally] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /not in the Drafts folder/i);
        assert.deepEqual(locationsNamed(err.message), ['mb-archive']);
        return true;
      },
    );
  });

  // A filing this server cannot read is its own outcome, not "filed somewhere else". Folding
  // it into the ordinary refusal produced a sentence that named no location and then handed
  // back a move_email repair for a draft that may already be in Drafts.
  //
  // The array row is also where `typeof [] === 'object'` is covered. An array read as a map
  // would be walked by Object.keys into its INDICES and the refusal would say "(it is in:
  // 0)"; it now cannot reach the sentence that lists locations at all, so the shape is
  // pinned HERE rather than by a separate test asserting that an index is absent from a
  // message which has no location list to put one in.
  const UNREADABLE_FILINGS: [string, any][] = [
    ['array-shaped', ['mb-archive']],
    ['a bare string', 'mb-archive'],
    ['null', null],
    ['absent', undefined],
  ];

  for (const [label, mailboxIds] of UNREADABLE_FILINGS) {
    it(`refuses with its own message when mailboxIds is ${label}`, async () => {
      const wrongShape = { ...SENDABLE_DRAFT, mailboxIds };
      const makeReq = stubRequests(client, async () => ({
        methodResponses: [['Email/get', { list: [wrongShape] }, 'getEmail']],
      }));

      await assert.rejects(
        () => client.sendDraft('draft-1'),
        (err: Error) => {
          assert.match(err.message, /no readable mailboxIds/i);
          // Not the ordinary refusal: that one asserts where the draft is and how to fix it,
          // and neither is known here.
          assert.equal(/not in the Drafts folder/i.test(err.message), false, err.message);
          assert.equal(/move_email/.test(err.message), false, err.message);
          return true;
        },
      );
      // Refused before submission: only the Email/get went out.
      assert.equal(makeReq.mock.calls.length, 1);
    });
  }

  // The gate demands `=== true`, and the sentence that reports where the draft IS has to
  // apply the same test. Plain truthiness there refuses the draft for not being in Drafts and
  // names Drafts as somewhere it is, in one sentence.
  it('does not name a mailbox whose membership value is truthy but not true', async () => {
    const notReallyEither = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-drafts': 1, 'mb-archive': true } };
    stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [notReallyEither] }, 'getEmail']],
    }));

    await assert.rejects(
      () => client.sendDraft('draft-1'),
      (err: Error) => {
        assert.match(err.message, /not in the Drafts folder/i);
        assert.deepEqual(locationsNamed(err.message), ['mb-archive']);
        return true;
      },
    );
  });

  it('does not accept a custom mailbox whose name merely contains "draft"', async () => {
    mock.method(client, 'getMailboxes', async () => [
      { id: 'mb-notes', name: 'Draft notes', role: null },
      SENT_MAILBOX,
    ]);
    const filedInNotes = { ...SENDABLE_DRAFT, mailboxIds: { 'mb-notes': true } };
    stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [filedInNotes] }, 'getEmail']],
    }));

    await assert.rejects(() => client.sendDraft('draft-1'), /drafts/i);
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

  // A part that declares no content type is displayed by whichever list carries it, so an
  // empty one renders blank exactly as an empty typed part does. Nothing saw it before,
  // because the guard selected a part by matching its type.
  it('rejects a draft whose only body part declares no content type and is blank', async () => {
    const typelessBlank = {
      ...SENDABLE_DRAFT,
      textBody: [{ partId: '1' }],
      htmlBody: [{ partId: '1' }], // the server aliases the one part into both lists
      bodyValues: { '1': { value: '   ' } },
    };
    stubRequests(client, async () => ({
      methodResponses: [['Email/get', { list: [typelessBlank] }, 'getEmail']],
    }));

    await assert.rejects(() => client.sendDraft('draft-1'), /empty htmlBody that would render blank/);
  });

  // The reason the guard reads every part rather than selecting one. These two drafts hold
  // the SAME parts and render the same way; only the list order differs. A selector answered
  // them differently, and the one it let through shipped an empty text/html part that shadows
  // the real body (RFC 2046).
  it('rejects a blank part whatever its position in the body list', async () => {
    for (const htmlBody of [
      [{ partId: '2' }, { partId: '3', type: 'text/html' }],
      [{ partId: '3', type: 'text/html' }, { partId: '2' }],
    ]) {
      const mixed = {
        ...SENDABLE_DRAFT,
        textBody: [{ partId: '1', type: 'text/plain' }],
        htmlBody,
        bodyValues: { '1': { value: 'Real text' }, '2': { value: '<p>Real html</p>' }, '3': { value: '' } },
      };
      stubRequests(client, async () => ({
        methodResponses: [['Email/get', { list: [mixed] }, 'getEmail']],
      }));

      await assert.rejects(
        () => client.sendDraft('draft-1'),
        /empty htmlBody that would render blank/,
        JSON.stringify(htmlBody),
      );
    }
  });
});

describe('findBlankBodyPart', () => {
  const values = (v: Record<string, string>) =>
    Object.fromEntries(Object.entries(v).map(([k, value]) => [k, { value }]));

  it('answers undefined for an ordinary draft with real content in both formats', () => {
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: values({ t: 'text', h: '<p>html</p>' }),
    }), undefined);
  });

  it('names the body field a blank part sits in, typed or typeless', () => {
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: values({ t: 'text', h: '  ' }),
    }), 'htmlBody');
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: values({ t: '', h: '<p>html</p>' }),
    }), 'textBody');
  });

  // htmlBody first, the order the two refusals have always been raised in: an empty
  // text/html part shadows a real text/plain alternative, so it is the worse of the two.
  it('names htmlBody when both formats are blank', () => {
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: values({ t: '', h: '' }),
    }), 'htmlBody');
  });

  // Not blank, and the distinction is the whole reason a draft with one format sends: the
  // draft has no body in that format at all, rather than one that renders to nothing.
  it('does not treat a part with no stored value as blank', () => {
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: values({ t: 'text' }),
    }), undefined);
  });

  // A part a body list does not display as text renders nothing on its own account, so an
  // empty value on one says nothing about what the recipient sees.
  it('ignores a blank part that is not displayed as text', () => {
    assert.equal(findBlankBodyPart({
      textBody: [{ partId: 't', type: 'text/plain' }],
      htmlBody: [{ partId: 'i', type: 'image/png', blobId: 'B' }],
      bodyValues: values({ t: 'text', i: '' }),
    }), undefined);
  });

  it('finds a blank part wherever it sits in the list', () => {
    const parts = [{ partId: 'a', type: 'text/plain' }, { partId: 'b', type: 'text/plain' }];
    const bodyValues = values({ a: 'real', b: '' });
    assert.equal(findBlankBodyPart({ textBody: parts, htmlBody: [], bodyValues }), 'textBody');
    assert.equal(findBlankBodyPart({ textBody: [...parts].reverse(), htmlBody: [], bodyValues }), 'textBody');
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

  // Both fixtures above are nameless drafts, so they stayed green under the old precedence
  // too and don't pin #152 for the wildcard-identity path specifically. This one does.
  it('keeps the draft\'s own name over a wildcard identity on a metadata-only edit (#152)', async () => {
    const namedWild = { ...EXISTING_DRAFT, from: [{ name: 'Work Sender', email: 'work@example.com' }] };
    const makeReq = stubRequests(client, async (req: any) => {
      if (req.methodCalls[0][0] === 'Email/get') {
        return { methodResponses: [['Email/get', { list: [namedWild] }, 'getEmail']] };
      }
      return { methodResponses: [['Email/set', { created: { draft: { id: 'draft-2' } }, destroyed: ['draft-1'] }, 'updateDraft']] };
    });

    await client.updateDraft('draft-1', { subject: 'Changed subject only' });

    const emailObj = callArguments(makeReq, 1)[0].methodCalls[0][1].create.draft;
    assert.deepEqual(emailObj.from, [{ name: 'Work Sender', email: 'work@example.com' }]);
  });
});

// ---------- updateDraft display-name resolution (#152) ----------
//
// Reversed precedence: the name the stored draft already carries against the address being
// written wins over the verified identity's configured name, which is now only a fallback
// for a draft that carries none. See writtenFromName in src/jmap-client.ts.
describe('updateDraft display name resolution', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient(); // getIdentities -> [IDENTITY] = { name: 'Test User', email: 'me@example.com' }
  });

  it('keeps the draft\'s own display name over a differently-named identity (#152)', async () => {
    const named = { ...EXISTING_DRAFT, from: [{ name: 'Custom Sender', email: 'me@example.com' }] };
    const makeReq = mockUpdate(client, named);

    await client.updateDraft('draft-1', { subject: 'New Subject' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Custom Sender', email: 'me@example.com' }]);
  });

  it('falls back to the identity\'s name when the stored draft carries none', async () => {
    const makeReq = mockUpdate(client, EXISTING_DRAFT); // from: [{ email: 'me@example.com' }], no name

    await client.updateDraft('draft-1', { subject: 'New Subject' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Test User', email: 'me@example.com' }]);
  });

  it('falls back to the identity\'s name when the stored name is empty/whitespace-only (#152)', async () => {
    const blankName = { ...EXISTING_DRAFT, from: [{ name: '   ', email: 'me@example.com' }] };
    const makeReq = mockUpdate(client, blankName);

    await client.updateDraft('draft-1', { subject: 'New Subject' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Test User', email: 'me@example.com' }]);
  });

  it('keeps a foreign-address draft\'s own name, with no identity name leaking in', async () => {
    const foreign = { ...EXISTING_DRAFT, from: [{ name: 'Elsewhere Sender', email: 'gone@elsewhere.example' }] };
    const makeReq = mockUpdate(client, foreign);

    await client.updateDraft('draft-1', { subject: 'New Subject' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Elsewhere Sender', email: 'gone@elsewhere.example' }]);
  });

  it('writes the new identity\'s name when the caller switches to a different own address, not the old stored name', async () => {
    const altIdentity = { id: 'id-2', name: 'Alias User', email: 'alias@example.com', mayDelete: true };
    mock.method(client, 'getIdentities', async () => [IDENTITY, altIdentity]);
    const named = { ...EXISTING_DRAFT, from: [{ name: 'Custom Sender', email: 'me@example.com' }] };
    const makeReq = mockUpdate(client, named);

    await client.updateDraft('draft-1', { from: 'alias@example.com' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Alias User', email: 'alias@example.com' }]);
  });

  // The stored-name address match must be case-insensitive: matchesIdentity (used to find
  // signingIdentity) lowercases both sides, so a caller re-passing the SAME mailbox in a
  // different case must not make the stored-name comparison miss and fall through to the
  // identity's name — that fallthrough is exactly the revert #152 exists to stop.
  it('keeps the draft\'s own name when the caller re-passes the same address in a different case (#152)', async () => {
    const mixedCase = { ...EXISTING_DRAFT, from: [{ name: 'Custom Sender', email: 'Me@Example.com' }] };
    const makeReq = mockUpdate(client, mixedCase);

    await client.updateDraft('draft-1', { from: 'me@example.com' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Custom Sender', email: 'me@example.com' }]);
  });

  it('keeps the draft\'s own name when the caller re-passes the identical stored address', async () => {
    const named = { ...EXISTING_DRAFT, from: [{ name: 'Custom Sender', email: 'me@example.com' }] };
    const makeReq = mockUpdate(client, named);

    await client.updateDraft('draft-1', { from: 'me@example.com' });

    assert.deepEqual(draftFromCall(makeReq).from, [{ name: 'Custom Sender', email: 'me@example.com' }]);
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
    // a forwarded message inline. .eml is also what an asAttachment forward writes.
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
      () => client.updateDraft('draft-1', { attachments: [NEW_PART], clearFields: ['textBody'], bodyHash: hashOf(DRAFT_ONE_ATT) }),
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
  let servedDraft: any;
  function mockEdit(c: JmapClient, before: any, saved?: any, opts: { readBackFails?: boolean } = {}) {
    servedDraft = before;
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

  /**
   * An edit of the draft the harness is currently serving, carrying that draft's `bodyHash`.
   *
   * The hash is a lost-update guard and is not what these tests are about, so it is supplied
   * from the fixture rather than written out at every call site — and it is supplied on
   * metadata-only edits too, where updateDraft ignores it, so no call site has to know which
   * kind it is. Passing an explicit `bodyHash` in `updates` overrides it, which is how the
   * few tests that DO care about the guard say so.
   */
  function edit(updates: any) {
    return client.updateDraft('draft-1', { bodyHash: hashOf(servedDraft), ...updates });
  }

  // ---- what the draft carries survives an edit that isn't about it ----

  it('carries a body-list-routed image the JMAP attachments array never listed', async () => {
    const draft = htmlDraft('<p>see <img src="cid:pic@x"></p>', [], [imagePart()]);
    const makeReq = mockEdit(client, draft);
    await edit({ subject: 'New' });
    assert.deepEqual(createdDraft(makeReq).attachments, [
      { blobId: 'blob-img', type: 'image/png', name: 'pic.png', disposition: 'inline', cid: 'pic@x' },
    ]);
  });

  it('says nothing about images on an edit that cannot have changed what the body displays', async () => {
    const draft = htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]);
    const makeReq = mockEdit(client, draft);
    const result = await edit({ subject: 'New' });
    assert.equal(result.notes, undefined);
    assert.equal(result.inlineImages, undefined);
    assert.equal(makeReq.mock.calls.length, 3); // no re-read: nothing was attached
  });

  it('resolves a percent-encoded reference to the part that supplies it', async () => {
    const draft = htmlDraft('<p>x</p>', [imagePart()]);
    const makeReq = mockEdit(client, draft, htmlDraft('<p>y</p>', [imagePart()]));
    await edit({ htmlBody: '<p>y <img src="cid:pic%40x"></p>' });
    assert.match(createdDraft(makeReq).bodyValues.html.value, /cid:pic%40x/);
  });

  // ---- a draft whose stored body already references an image it doesn't carry ----

  const BROKEN = htmlDraft('<p>hi</p><img src="cid:gone@x">');

  it('refuses a body edit while the stored body references an image nothing supplies', async () => {
    mockEdit(client, BROKEN);
    await assert.rejects(
      () => edit({ htmlBody: '<p>new</p><img src="cid:gone@x">' }),
      /stored body references image identifier\(s\) with no matching attachment.*gone@x/s,
    );
  });

  it('still runs a metadata edit on such a draft, carrying the broken body verbatim', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await edit({ subject: 'Renamed' });
    const draft = createdDraft(makeReq);
    assert.equal(draft.bodyValues.html.value, '<p>hi</p><img src="cid:gone@x">');
    assert.equal(draft.textBody, undefined); // body-invariant: no text part invented
  });

  it('still appends an unrelated attachment to such a draft', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await edit({
      attachments: [{ blobId: 'b-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' }],
    });
    assert.equal(createdDraft(makeReq).attachments.length, 1);
  });

  it('accepts an attachment that supplies the missing image', async () => {
    const supplied = imagePart({ cid: 'gone@x', blobId: 'b-gone' });
    const makeReq = mockEdit(client, BROKEN, htmlDraft('<p>hi</p><img src="cid:gone@x">', [supplied]));
    const result = await edit({
      attachments: [{ blobId: 'b-gone', type: 'image/png', name: 'pic.png', cid: 'gone@x', disposition: 'inline' }],
    });
    assert.equal(createdDraft(makeReq).attachments.length, 1);
    assert.deepEqual(result.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('accepts a body edit that replaces the broken body outright', async () => {
    const makeReq = mockEdit(client, BROKEN);
    await edit({ htmlBody: '<p>all new</p>' });
    assert.equal(createdDraft(makeReq).bodyValues.html.value, '<p>all new</p>');
  });

  // ---- identifiers the recreate cannot reproduce ----

  const EXOTIC_CID = htmlDraft('<p>hi</p>', [imagePart({ cid: 'has space@x' })]);

  it('refuses a body edit when a stored identifier cannot be re-created faithfully', async () => {
    mockEdit(client, EXOTIC_CID);
    await assert.rejects(
      () => edit({ htmlBody: '<p>new</p>' }),
      /cannot safely re-create.*Recreate it with draft_email/s,
    );
  });

  it('leaves metadata edits of such a draft alone (the carry reproduces the value verbatim)', async () => {
    const makeReq = mockEdit(client, EXOTIC_CID);
    await edit({ subject: 'Renamed' });
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
      () => edit({ subject: 'X' }),
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
      () => edit({ removeAttachments: ['no-such-blob'] }),
      /cannot carry/,
    );
  });

  // ---- removals ----

  it('refuses a removal that would leave the surviving body pointing at nothing', async () => {
    mockEdit(client, htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]));
    await assert.rejects(
      () => edit({ removeAttachments: ['blob-img'] }),
      /removeAttachments would remove an image the draft's body still references/,
    );
  });

  it('refuses an attachment wipe that would leave the surviving body pointing at nothing', async () => {
    mockEdit(client, htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]));
    await assert.rejects(
      () => edit({ clearFields: ['attachments'] }),
      /would strip image\(s\) the surviving body still references/,
    );
  });

  it('wipes the attachments when the surviving body references none of them', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>no images here</p>', [imagePart()]));
    await edit({ clearFields: ['attachments'] });
    assert.equal(createdDraft(makeReq).attachments, undefined);
  });

  it('takes a server-managed image off the draft when the rewritten body stops displaying it', async () => {
    const draft = htmlDraft(`<p><img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    const makeReq = mockEdit(client, draft);
    const result = await edit({ htmlBody: '<p>text only now</p>' });
    assert.equal(createdDraft(makeReq).attachments, undefined);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 0, removed: 1 });
    assert.deepEqual(result.notes, ['Removed 1 image(s) that were embedded in the quote.']);
  });

  it("keeps someone else's image as a regular attachment when the body stops displaying it", async () => {
    const draft = htmlDraft('<p><img src="cid:pic@x"></p>', [imagePart()]);
    const makeReq = mockEdit(client, draft);
    const result = await edit({ htmlBody: '<p>text only now</p>' });
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
    const result = await edit({ htmlBody: '<p>text only now</p>' });
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
    const result = await edit({ htmlBody: '<p>text only now</p>' });
    assert.equal(createdDraft(makeReq).attachments.length, 2);
    assert.deepEqual(result.inlineImages, { embedded: 0, degraded: 2, removed: 0 });
    assert.deepEqual(result.notes, ['2 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  // A metadata-only edit writes no body and is not asked for a hash — but taking a part out
  // of a body list changes the SET the hash is taken over, so a hash the caller was holding
  // stops matching after an edit that wrote nothing. Nothing is lost: the next body edit is
  // refused and the caller re-reads, which is the guard working. Both descriptions say this;
  // this pins that they are describing the real behaviour.
  it('stales a held hash when removeAttachments takes an image out of a body list', async () => {
    const before = htmlDraft('<p>no image displayed here</p>', [], [imagePart()]);
    const heldHash = hashOf(before);
    const after = htmlDraft('<p>no image displayed here</p>');

    mockEdit(client, before, after);
    const removal = await client.updateDraft('draft-1', { removeAttachments: ['blob-img'] });
    // Metadata-only: no body written, so no hash is issued for the next edit either.
    assert.equal(removal.bodyHash, undefined);

    mockEdit(client, after, after);
    await assert.rejects(
      () => client.updateDraft('draft-2', { htmlBody: '<p>next</p>', bodyHash: heldHash }),
      /not this draft's current one/,
    );
  });

  // The other half of the same claim, and the reason both descriptions say "a part the
  // server routed into a body list" rather than "an attachment". collectDraftBodyParts walks
  // textBody and htmlBody and nothing else, so a part that only ever sat in the attachments
  // array is not in the hash and taking it off cannot stale one. Telling a caller their hash
  // is dead when it is not is the same defect as not telling them when it is.
  it('leaves a held hash valid when removeAttachments takes off a part no body list carries', async () => {
    const attached = {
      partId: 'p1', blobId: 'blob-att', type: 'application/pdf',
      name: 'doc.pdf', disposition: 'attachment', size: 10,
    };
    const before = htmlDraft('<p>unchanged body</p>', [attached]);
    const heldHash = hashOf(before);
    const after = htmlDraft('<p>unchanged body</p>');

    mockEdit(client, before, after);
    await client.updateDraft('draft-1', { removeAttachments: ['blob-att'] });

    mockEdit(client, after, after);
    const next = await client.updateDraft('draft-2', { htmlBody: '<p>next</p>', bodyHash: heldHash });
    assert.ok(next.id, 'the held hash was rejected after a removal the hash does not cover');
  });

  it('reports a part the caller removed once, as a removal it asked for', async () => {
    // The explicit removal is the caller's own action and needs no sentence; what must not
    // happen is the same part ALSO being counted as one the edit took off the body.
    const draft = htmlDraft(`<p><img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    mockEdit(client, draft);
    const result = await edit({ htmlBody: '<p>gone</p>', removeAttachments: ['blob-img'] });
    assert.equal(result.notes, undefined);
    assert.equal(result.inlineImages, undefined);
  });

  // ---- collisions ----

  it('refuses two attachments in one call sharing an identifier', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => edit({
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
      () => edit({
        attachments: [{ blobId: 'b-other', type: 'image/png', cid: 'pic@x', disposition: 'inline' }],
      }),
      /already used by another attachment on this draft/,
    );
  });

  // Narrowed to AUTHORING one. A minted identifier is durable now — it survives an edit for
  // as long as the body keeps referencing it — so the two halves are a matched pair and the
  // pair is what pins the rule: naming a part the draft does not carry is refused, naming one
  // it does is the ordinary read-edit-write shape. Refusing both would refuse the first edit
  // of every image-bearing draft this server made.
  it('refuses a caller reference AUTHORING a server-managed identifier the draft has no part for', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => edit({ htmlBody: `<p><img src="cid:${STORED_MINT}"></p>` }),
      /a server-managed identifier for quoted images, and this draft carries no part under it/,
    );
  });

  it('accepts the same reference handed back on a draft that carries the part', async () => {
    const draft = htmlDraft(`<p>x <img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    const makeReq = mockEdit(client, draft);
    await edit({ htmlBody: `<p>my edited prose <img src="cid:${STORED_MINT}"></p>` });
    const saved = createdDraft(makeReq);
    assert.match(saved.bodyValues.html.value, new RegExp(`cid:${STORED_MINT}`));
    assert.equal(saved.attachments.some((a: any) => a.cid === STORED_MINT), true);
  });

  it('takes the part off when the caller drops that reference, rather than refusing', async () => {
    const draft = htmlDraft(`<p>x <img src="cid:${STORED_MINT}"></p>`, [imagePart({ cid: STORED_MINT })]);
    const makeReq = mockEdit(client, draft);
    const result = await edit({ htmlBody: '<p>my edited prose</p>' });
    assert.equal(createdDraft(makeReq).attachments, undefined); // an empty list is omitted, not written
    assert.equal(result.inlineImages?.removed, 1);
  });

  it('refuses a caller reference nothing supplies', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => edit({ htmlBody: '<p><img src="cid:mine-logo"></p>' }),
      /references cid "mine-logo" but no attachment supplies it.*add an attachments item with cid: "mine-logo"/s,
    );
  });

  it('offers no attachments repair when this server cannot attach files at all', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    await assert.rejects(
      () => client.updateDraft(
        'draft-1',
        { htmlBody: '<p><img src="cid:mine-logo"></p>', bodyHash: hashOf(servedDraft) },
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
    await edit({
      htmlBody: '<p><img src="cid:mine-logo"></p>',
      attachments: [{ blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'attachment' }],
    });
    assert.deepEqual(createdDraft(makeReq).attachments, [
      { blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'inline' },
    ]);
  });

  it('leaves an added file with an unreferenced cid as an ordinary attachment, and says so', async () => {
    const makeReq = mockEdit(client, htmlDraft('<p>x</p>', []));
    const result = await edit({
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
    const result = await edit({
      textBody: 'plain text only now',
      clearFields: ['htmlBody'],
      attachments: [{ blobId: 'b-logo', type: 'image/png', name: 'logo.png', cid: 'mine-logo', disposition: 'attachment' }],
    });
    assert.deepEqual(result.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('says nothing about an ordinary added file that asked to embed nothing', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []));
    const result = await edit({
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
    const result = await edit({
      htmlBody: '<p><img src="cid:logo@me"></p>',
      attachments: [NEW_INLINE],
    });
    // get, create, trash, then TWO reads of the saved draft: the inline-image confirm and
    // the read the returned bodyHash is taken from. They are separate on purpose — the hash
    // must come from the stored parts, not from the bytes this call sent.
    assert.equal(makeReq.mock.calls.length, 5);
    assert.deepEqual(result.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('says so when the saved draft comes back without the image it attached', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []), htmlDraft('<p>x</p>', []));
    const result = await edit({ htmlBody: '<p><img src="cid:logo@me"></p>', attachments: [NEW_INLINE] });
    assert.ok(result.notes?.some((n) => /were not found on the saved draft/.test(n)));
  });

  it('never fails the edit when the confirming read does', async () => {
    mockEdit(client, htmlDraft('<p>x</p>', []), undefined, { readBackFails: true });
    const result = await edit({
      htmlBody: '<p><img src="cid:logo@me"></p>',
      attachments: [NEW_INLINE],
    });
    assert.equal(result.id, 'draft-2');
    assert.ok(result.notes?.some((n) => /could not re-read it to confirm/.test(n)));
  });
});

// ---------- source-instance header (X-Fastmail-MCP-Source-Id) ----------
// The exact stored instance a draft was composed from: stamped by draft_email's reply and
// forward modes, carried by the edit recreate, consumed by send_draft's keyword
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
    [SRC_PROP]: 'fwd-src-1',
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

  it('KEEPS the header when a reply draft has its whole body rewritten (In-Reply-To survives, so does the instance pointer)', async () => {
    const makeReq = mockSrcUpdate(client, REPLY_QUOTED);
    await client.updateDraft('rdraft-1', { htmlBody: '<p>fresh body</p>', bodyHash: hashOf(REPLY_QUOTED) });
    const draft = draftFromCall(makeReq);
    assert.equal(draft[SRC_PROP], 'orig-1');
    assert.deepEqual(draft.inReplyTo, ['orig-msg@example.com']);
  });

  // De-forwarding is now an explicit act — clearFields:['forwardedMessageId'] — rather than
  // something a body rewrite inferred. The pointer still follows the marking it refines: a
  // draft that no longer forwards anything has no instance to name.
  it('DROPS the header with the forward marking when the caller clears forwardedMessageId (a de-forward)', async () => {
    const makeReq = mockSrcUpdate(client, FORWARD_WITH_SRC);
    await client.updateDraft('fdraft-1', { clearFields: ['forwardedMessageId'] });
    const draft = draftFromCall(makeReq);
    assert.equal(draft[FWD_PROP], undefined);
    assert.equal(draft[SRC_PROP], undefined);
  });

  // The pair to the test above: a body rewrite alone is NOT a de-forward. Both markings ride
  // through unchanged, and nothing re-points them — the draft still forwards the same
  // instance however its note is reworded, and the original is never fetched to find out.
  it('carries both markings through a forward-draft body rewrite, and fetches no original', async () => {
    const makeReq = mockSrcUpdate(client, FORWARD_WITH_SRC);
    await client.updateDraft('fdraft-1', { htmlBody: '<p>new note</p>', bodyHash: hashOf(FORWARD_WITH_SRC) });
    const draft = draftFromCall(makeReq);
    assert.deepEqual(draft[FWD_PROP], ['fwd-orig-msg@example.com']);
    assert.equal(draft[SRC_PROP], 'fwd-src-1');
    const fetchedIds = makeReq.mock.calls
      .map((c: any) => c.arguments[0].methodCalls[0])
      .filter(([m]: any) => m === 'Email/get')
      .flatMap(([, p]: any) => p.ids ?? []);
    assert.equal(fetchedIds.includes('orig-1'), false);
  });

  it('a draft that never had the header stays without it', async () => {
    const { [SRC_PROP]: _drop, ...rest } = REPLY_QUOTED as any;
    const makeReq = mockSrcUpdate(client, { ...rest, id: 'rdraft-1' });
    await client.updateDraft('rdraft-1', { subject: 'Re: Hello (edited)' });
    assert.equal(draftFromCall(makeReq)[SRC_PROP], undefined);
  });
});
