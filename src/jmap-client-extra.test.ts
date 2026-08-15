import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { JmapClient, EMAIL_PROPERTIES_COMPACT, EMAIL_PROPERTIES_VERBOSE, EMAIL_BODY_PROPERTIES, buildMailboxInfoMap, attachMailboxInfo, findMailboxExact, resolveMailbox, buildMailboxPathMap, filterMailboxesByParent, computeExclusion, readSourceReferences, fillUrlTemplate } from './jmap-client.js';
import type { JmapRequest } from './jmap-client.js';
import { callArguments } from './testing/mock-calls.js';
import { InvalidInputError } from './coerce.js';
import { validateFastmailUrl } from './url-validation.js';
import { buildExclusionNote, simplifyMailbox } from './response-formatters.js';
import { FastmailAuth } from './auth.js';
import { mkdtemp, writeFile as fsWriteFile, readFile, rm } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';

// ---------- helpers ----------

const ACCOUNT_ID = 'acct-123';
const INBOX_MAILBOX = { id: 'mb-inbox', name: 'Inbox', role: 'inbox', totalEmails: 42, unreadEmails: 5 };
const DRAFTS_MAILBOX = { id: 'mb-drafts', name: 'Drafts', role: 'drafts' };
const TRASH_MAILBOX = { id: 'mb-trash', name: 'Trash', role: 'trash' };
const SENT_MAILBOX = { id: 'mb-sent', name: 'Sent', role: 'sent' };
const ARCHIVE_MAILBOX = { id: 'mb-archive', name: 'Archive', role: 'archive' };
const JUNK_MAILBOX = { id: 'mb-junk', name: 'Spam', role: 'junk' };

// Default fixture set carries both trash and junk roles (so the default Trash/Spam
// exclusion resolves both) plus archive (a move target).
const DEFAULT_MAILBOXES = [INBOX_MAILBOX, DRAFTS_MAILBOX, TRASH_MAILBOX, SENT_MAILBOX, ARCHIVE_MAILBOX, JUNK_MAILBOX];

function makeClient(): JmapClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new JmapClient(auth);

  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {},
  }));

  return client;
}

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

function stubMakeRequest(client: JmapClient, response: any) {
  stubRequests(client, async () => response);
}

function stubMailboxes(client: JmapClient, mailboxes: any[] = DEFAULT_MAILBOXES) {
  mock.method(client, 'getMailboxes', async () => mailboxes);
}

// A request-aware Email/query stub: the visible query reads ids/total from `query`, the
// get reads from `emails`, and (when present) the count query reads from `count`. Because
// getMailboxes is stubbed separately in these tests, makeRequest only ever sees the
// Email/query batch, so this returns the same fixed shape regardless of the methodCalls.
function queryResponse(opts: { ids?: string[]; list?: any[]; total?: number; broaderTotal?: number; position?: number }) {
  const responses: any[] = [
    ['Email/query', { ids: opts.ids ?? [], total: opts.total, position: opts.position }, 'query'],
    ['Email/get', { list: opts.list ?? [] }, 'emails'],
  ];
  if (opts.broaderTotal !== undefined) {
    responses.push(['Email/query', { ids: [], total: opts.broaderTotal }, 'count']);
  }
  return { methodResponses: responses };
}

// ---------- getMailboxes ----------

describe('getMailboxes', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns list of mailboxes on valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
    assert.equal(mailboxes[0].role, 'inbox');
    assert.equal(mailboxes[1].id, 'mb-drafts');
  });

  it('returns empty array when response list is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', {}, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });
});

// ---------- getRecentEmails ----------

describe('getRecentEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns recent emails on valid response', async () => {
    stubMailboxes(client);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1', 'e2'] }, 'query'],
        ['Email/get', { list: [
          { id: 'e1', subject: 'First' },
          { id: 'e2', subject: 'Second' },
        ] }, 'emails'],
      ],
    });

    const result = await client.getRecentEmails(10, 'inbox');
    assert.equal(result.items.length, 2);
    assert.equal(result.items[0].subject, 'First');
  });

  it('throws InvalidInputError when mailbox is not found', async () => {
    stubMailboxes(client, [INBOX_MAILBOX]);

    await assert.rejects(
      () => client.getRecentEmails(10, 'nonexistent'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /not found/);
        return true;
      },
    );
  });

  it('matches mailbox by role', async () => {
    stubMailboxes(client, [{ id: 'mb-custom', name: 'My Inbox', role: 'inbox' }]);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    });

    const result = await client.getRecentEmails(5, 'inbox');
    assert.deepEqual(result.items, []);
  });

  it('falls back to inbox for a blank mailbox (does not throw)', async () => {
    stubMailboxes(client);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: [] }, 'query'],
        ['Email/get', { list: [] }, 'emails'],
      ],
    }));

    await client.getRecentEmails(5, '   ');
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
  });
});

// ---------- getEmails ----------

describe('getEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns emails scoped to an explicit mailbox (no exclusion, no count query)', async () => {
    stubMailboxes(client);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1', subject: 'Filtered' }], total: 1 }),
    );

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });
    assert.equal(result.items.length, 1);

    const batch = callArguments(makeReq)[0].methodCalls;
    const filter = batch[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.inMailboxOtherThan, undefined);
    // Explicit mailbox => no exclusion => no count query and no exclusion metadata.
    assert.equal(batch.length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('default (no mailbox) excludes Trash + Spam via inMailboxOtherThan and runs a count query', async () => {
    stubMailboxes(client);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1', subject: 'All' }], total: 8, broaderTotal: 11 }),
    );

    const result = await client.getEmails({ limit: 10 });
    assert.equal(result.items.length, 1);

    const batch = callArguments(makeReq)[0].methodCalls;
    const filter = batch[0][1].filter;
    assert.deepEqual([...filter.inMailboxOtherThan].sort(), ['mb-junk', 'mb-trash']);
    // Count query present (visible filter minus inMailboxOtherThan) at index 2.
    assert.equal(batch.length, 3);
    assert.equal(batch[2][1].filter.inMailboxOtherThan, undefined);
    // hidden = broaderTotal - visibleTotal = 11 - 8 = 3.
    assert.equal(result.exclusion?.hidden, 3);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Trash', 'Spam']);
    assert.deepEqual(result.exclusion?.unresolvedRoles, []);
  });

  it('includeTrash + includeSpam disable the exclusion entirely', async () => {
    stubMailboxes(client);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 1 }),
    );

    const result = await client.getEmails({ includeTrash: true, includeSpam: true });
    const batch = callArguments(makeReq)[0].methodCalls;
    assert.equal(batch[0][1].filter.inMailboxOtherThan, undefined);
    assert.equal(batch.length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('excludeDrafts adds notKeyword $draft, AND-wrapped with the default exclusion', async () => {
    stubMailboxes(client);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }),
    );

    await client.getEmails({ excludeDrafts: true });
    const batch = callArguments(makeReq)[0].methodCalls;
    const filter = batch[0][1].filter;
    // AND of the base (carrying inMailboxOtherThan) and the $draft keyword condition.
    assert.equal(filter.operator, 'AND');
    const hasExclusion = filter.conditions.some((c: any) => c.inMailboxOtherThan);
    const hasDraft = filter.conditions.some((c: any) => c.notKeyword === '$draft');
    assert.ok(hasExclusion && hasDraft);
    // Count filter keeps the $draft cond but drops inMailboxOtherThan -> differs from visible.
    const countFilter = batch[2][1].filter;
    assert.equal(countFilter.notKeyword, '$draft');
    assert.equal(countFilter.inMailboxOtherThan, undefined);
  });
});

// ---------- paging (#51) ----------

describe('paging with position', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  const emailQuery = (makeReq: ReturnType<typeof stubRequests>) =>
    callArguments(makeReq)[0].methodCalls[0][1];

  it('getEmails sends the requested position to Email/query', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 137, broaderTotal: 137, position: 40 }),
    );

    const result = await client.getEmails({ limit: 20, position: 40 });
    assert.equal(emailQuery(makeReq).position, 40);
    assert.equal(result.position, 40);
    assert.equal(result.total, 137);
  });

  // 0 is the JMAP default, so the paging parameter must not change the request the
  // non-paging callers have always sent.
  it('getEmails sends no position for position 0', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }),
    );

    await client.getEmails({ limit: 20, position: 0 });
    assert.equal('position' in emailQuery(makeReq), false);
  });

  it('getEmails sends no position when none is asked for', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }),
    );

    await client.getEmails({ limit: 20 });
    assert.equal('position' in emailQuery(makeReq), false);
  });

  // The Trash/Spam exclusion lives inside the JMAP filter, so it applies to whichever
  // page is served — and the hidden count still describes the whole match set.
  it('getEmails keeps the Trash/Spam exclusion and its hidden count on a later page', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8, broaderTotal: 11, position: 40 }),
    );

    const result = await client.getEmails({ limit: 20, position: 40 });
    const batch = callArguments(makeReq)[0].methodCalls;
    assert.deepEqual([...batch[0][1].filter.inMailboxOtherThan].sort(), ['mb-junk', 'mb-trash']);
    // The hidden-count query is a count, not a page: it carries no position.
    assert.equal('position' in batch[2][1], false);
    assert.equal(result.exclusion?.hidden, 3);
  });

  // RFC 8620 section 5.5 has the server report the position it actually served, and
  // a request past the end is clamped there rather than erroring.
  it('getEmails reports the server position over the requested one', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 137, broaderTotal: 137, position: 137 }),
    );

    const result = await client.getEmails({ limit: 20, position: 500 });
    assert.deepEqual(result.items, []);
    assert.equal(result.position, 137);
    assert.equal(result.total, 137);
  });

  it('getEmails falls back to the requested position when the response omits it', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 137, broaderTotal: 137 }),
    );

    const result = await client.getEmails({ limit: 20, position: 40 });
    assert.equal(result.position, 40);
  });

  it('searchEmails sends the requested position to Email/query', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 137, broaderTotal: 137, position: 100 }),
    );

    const result = await client.searchEmails({ query: 'invoice', limit: 100, position: 100 });
    assert.equal(emailQuery(makeReq).position, 100);
    assert.equal(result.position, 100);
    assert.equal(result.total, 137);
  });

  // get_recent_emails caps limit at 50, so position is the only way past that cap.
  it('getRecentEmails sends the requested position and reports it back', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 90, position: 50 }),
    );

    const result = await client.getRecentEmails(50, 'inbox', false, 50);
    assert.equal(emailQuery(makeReq).position, 50);
    assert.equal(result.position, 50);
    assert.equal(result.total, 90);
  });

  it('getRecentEmails sends no position by default', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }),
    );

    await client.getRecentEmails(10, 'inbox');
    assert.equal('position' in emailQuery(makeReq), false);
  });
});

// ---------- getEmailById ----------

describe('getEmailById', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns email on valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/get', { list: [{ id: 'e1', subject: 'Found' }] }, 'email'],
      ],
    });

    const email = await client.getEmailById('e1');
    assert.equal(email.id, 'e1');
    assert.equal(email.subject, 'Found');
  });

  it('throws when email is not found (empty list)', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/get', { list: [] }, 'email'],
      ],
    });

    await assert.rejects(
      () => client.getEmailById('missing'),
      (err: Error) => {
        assert.match(err.message, /not found/);
        // a bad id is caller-fixable input → InvalidInputError → InvalidParams (#41)
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('throws when email is in notFound list', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/get', { list: [], notFound: ['gone'] }, 'email'],
      ],
    });

    await assert.rejects(
      () => client.getEmailById('gone'),
      (err: Error) => {
        assert.match(err.message, /not found/);
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });
});

// ---------- moveEmail ----------

describe('moveEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    // moveEmail resolves the destination against getMailboxes(); stub it so 'mb-archive'
    // resolves and makeRequest only sees the Email/set.
    stubMailboxes(client);
  });

  it('sets mailboxIds whole-value in a single request, and writes no keywords', async () => {
    // The membership is REPLACED rather than read and patched id-by-id, which is what the
    // tool promises and what keeps a mailbox added between a read and a write from
    // surviving the move. And a move is a filing change only: writing a keyword here would
    // silently change a message's read or flagged state.
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { updated: { 'e1': null } }, 'moveEmail']],
    }));

    await client.moveEmail('e1', 'mb-archive');

    assert.equal(makeReq.mock.callCount(), 1, 'no pre-read round trip');
    const methodCalls = callArguments(makeReq, 0)[0].methodCalls;
    assert.equal(methodCalls.length, 1);
    assert.equal(methodCalls[0][0], 'Email/set');
    const update = methodCalls[0][1].update.e1;
    assert.deepEqual(update, { mailboxIds: { 'mb-archive': true } });
    assert.deepEqual(Object.keys(update), ['mailboxIds']);
  });

  // Drive a notUpdated failure with a chosen SetError.
  function stubMoveFailure(setError: { type: string; description?: string }) {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e1': setError } }, 'moveEmail'],
      ],
    });
  }

  it('throws InvalidInputError with the JMAP reason when the id is notFound (#22 + #41)', async () => {
    stubMoveFailure({ type: 'notFound' });

    await assert.rejects(
      () => client.moveEmail('e1', 'mb-archive'),
      (err: Error) => {
        // #22: the server's reason is surfaced (not a bare "Failed to move email.")
        assert.match(err.message, /Failed to move email: notFound/);
        // type-only SetError → no dangling " - " from describeSetError
        assert.doesNotMatch(err.message, / - /);
        // #41: notFound is a caller-fixable bad id → InvalidInputError → InvalidParams
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('surfaces type AND description, and stays a plain Error for an operational failure (#22 + #41)', async () => {
    stubMoveFailure({ type: 'forbidden', description: 'mailbox is read-only' });

    await assert.rejects(
      () => client.moveEmail('e1', 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /Failed to move email: forbidden - mailbox is read-only/);
        // a server refusal is not caller-fixable by re-forming args → plain Error → InternalError
        assert.equal(err.name, 'Error');
        return true;
      },
    );
  });

  // Both halves of the JMAP set-error split are pinned by type. Asserting only the
  // caller-fixable half would let a later blanket widening — every type becoming
  // InvalidInputError — pass unnoticed, which would tell callers to re-form arguments
  // in the face of failures no argument can clear.
  const CALLER_FIXABLE_TYPES = [
    'notFound',
    'invalidProperties',
    'invalidArguments',
    'invalidPatch',
    'tooLarge',
    'singleton',
  ];

  for (const type of CALLER_FIXABLE_TYPES) {
    it(`classes a ${type} set error as caller-fixable input`, async () => {
      stubMoveFailure({ type });

      await assert.rejects(
        () => client.moveEmail('e1', 'mb-archive'),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError, `${type} should map to InvalidParams`);
          // the message is the shared "Failed to <action>: <type>" format either way
          assert.equal(err.message, `Failed to move email: ${type}`);
          return true;
        },
      );
    });
  }

  const SERVER_SIDE_TYPES = [
    'forbidden',
    'overQuota',
    'rateLimit',
    'serverFail',
    'serverPartialFail',
    'stateMismatch',
    'accountReadOnly',
    'cannotCalculateChanges',
    'willDestroy',
    // A type this client has no judgement about falls to the same side: an unknown
    // failure is not evidence that the caller's input was wrong.
    'someFutureJmapError',
  ];

  for (const type of SERVER_SIDE_TYPES) {
    it(`keeps a ${type} set error a server-side failure`, async () => {
      stubMoveFailure({ type });

      await assert.rejects(
        () => client.moveEmail('e1', 'mb-archive'),
        (err: Error) => {
          assert.ok(!(err instanceof InvalidInputError), `${type} must not map to InvalidParams`);
          assert.equal(err.message, `Failed to move email: ${type}`);
          return true;
        },
      );
    });
  }
});

// ---------- archiveEmail ----------

describe('archiveEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  it('replaces membership with the archive mailbox alone, and writes no keywords', async () => {
    // A pure move: exactly one destination id, set whole-value, and no keywords key at
    // all. The read state is deliberately untouched — archiving a message does not read it
    // (mark_email_read is the separate call), so a $seen write here would be a silent
    // second effect.
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'archiveEmail']],
    }));

    await client.archiveEmail('e1');

    assert.equal(makeReq.mock.callCount(), 1, 'no pre-read round trip');
    const methodCalls = callArguments(makeReq, 0)[0].methodCalls;
    assert.equal(methodCalls.length, 1);
    assert.equal(methodCalls[0][0], 'Email/set');
    const update = methodCalls[0][1].update.e1;
    assert.deepEqual(update, { mailboxIds: { 'mb-archive': true } });
    assert.deepEqual(Object.keys(update), ['mailboxIds']);
    assert.equal(Object.keys(update.mailboxIds).length, 1);
  });

  it('prefers the archive ROLE over a folder that merely carries the name', async () => {
    // Half of the guard: with both present, the role wins. On its own this case cannot
    // tell an exact-role lookup apart from a role-then-name fallback, because a
    // role-first fallback picks the same mailbox — the case below is the one that
    // separates them, and the two are only meaningful together.
    stubMailboxes(client, [
      INBOX_MAILBOX,
      { id: 'mb-imposter', name: 'archive' },
      { id: 'mb-real-archive', name: 'Old mail', role: 'archive' },
    ]);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'archiveEmail']],
    }));

    await client.archiveEmail('e1');

    const update = callArguments(makeReq, 0)[0].methodCalls[0][1].update.e1;
    assert.deepEqual(update.mailboxIds, { 'mb-real-archive': true });
  });

  it('refuses a folder named "archive" when no mailbox carries the role', async () => {
    // The case the whole role-only rule exists for, and the only one that discriminates:
    // a name fallback would file the mail into a folder ANY caller can create, under an
    // innocuous verb. Nothing may be sent at all here — the assertion on makeRequest is
    // the substantive half, since a wrong destination would still look like a success.
    stubMailboxes(client, [INBOX_MAILBOX, TRASH_MAILBOX, { id: 'mb-imposter', name: 'Archive' }]);
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [['Email/set', { updated: { e1: null } }, 'archiveEmail']],
    }));

    await assert.rejects(
      () => client.archiveEmail('e1'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /archive role/);
        return true;
      },
    );
    assert.equal(makeReq.mock.callCount(), 0, 'a named folder must not receive the mail');
  });

  it('rejects an account with no archive-role mailbox, naming the alternative', async () => {
    stubMailboxes(client, [INBOX_MAILBOX, TRASH_MAILBOX]);

    await assert.rejects(
      () => client.archiveEmail('e1'),
      (err: Error) => {
        // Caller-fixable: they can name any destination on move_email, so the message has
        // to say so and the class has to tell the client to re-form rather than retry.
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /archive role/);
        assert.match(err.message, /move_email/);
        return true;
      },
    );
  });

  it('routes a notUpdated bad id through the shared set-error classifier', async () => {
    stubMakeRequest(client, {
      methodResponses: [['Email/set', { notUpdated: { e1: { type: 'notFound' } } }, 'archiveEmail']],
    });

    await assert.rejects(
      () => client.archiveEmail('e1'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.equal(err.message, 'Failed to archive email: notFound');
        return true;
      },
    );
  });

  it('keeps an operational notUpdated reason a server-side failure', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { e1: { type: 'forbidden', description: 'mailbox is read-only' } } }, 'archiveEmail'],
      ],
    });

    await assert.rejects(
      () => client.archiveEmail('e1'),
      (err: Error) => {
        assert.ok(!(err instanceof InvalidInputError));
        assert.equal(err.message, 'Failed to archive email: forbidden - mailbox is read-only');
        return true;
      },
    );
  });
});

// ---------- deleteEmail ----------

describe('deleteEmail', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('deletes email by moving to trash', async () => {
    stubMailboxes(client);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'moveToTrash'],
      ],
    });

    await client.deleteEmail('e1');
    // No error means success
  });

  it('throws when trash mailbox is not found', async () => {
    stubMailboxes(client, [INBOX_MAILBOX, DRAFTS_MAILBOX]);

    await assert.rejects(
      () => client.deleteEmail('e1'),
      (err: Error) => {
        assert.match(err.message, /Trash/);
        return true;
      },
    );
  });
});

// ---------- markEmailRead ----------

describe('markEmailRead', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('marks email as read', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', true);

    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
  });

  it('marks email as unread', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', false);

    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': null });
  });

  it('throws when update fails', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e1': { type: 'notFound' } } }, 'updateEmail'],
      ],
    });

    await assert.rejects(
      () => client.markEmailRead('e1'),
      (err: Error) => {
        assert.match(err.message, /Failed to mark/);
        return true;
      },
    );
  });
});

// ---------- addKeywords ----------

describe('addKeywords', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('emits an additive keywords/<k>:true patch for each keyword', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'addKeywords'],
      ],
    }));

    await client.addKeywords('e1', ['$answered', '$seen']);

    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$answered': true, 'keywords/$seen': true });
  });

  it('routes a notUpdated entry through throwSingleSetError', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e1': { type: 'notFound' } } }, 'addKeywords'],
      ],
    });

    await assert.rejects(
      () => client.addKeywords('e1', ['$answered']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Failed to add keywords to email/);
        return true;
      },
    );
  });
});

// ---------- bulkMarkRead ----------

describe('bulkMarkRead', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('marks multiple emails as read in one request', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null, 'e3': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2', 'e3'], true);

    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
    assert.deepEqual(update['e2'], { 'keywords/$seen': true });
    assert.deepEqual(update['e3'], { 'keywords/$seen': true });
  });

  it('marks multiple emails as unread', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2'], false);

    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': null });
    assert.deepEqual(update['e2'], { 'keywords/$seen': null });
  });

  it('throws when some emails fail to update, surfacing counts + failing ids + reason (#22)', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e2': { type: 'notFound' } } }, 'bulkUpdate'],
      ],
    });

    await assert.rejects(
      () => client.bulkMarkRead(['e1', 'e2']),
      (err: Error) => {
        // counts: 1 of 2 failed, 1 succeeded; failing id grouped by its reason
        assert.match(err.message, /Failed to mark as read 1 of 2 emails \(1 succeeded\)/);
        assert.match(err.message, /notFound: e2/);
        // every failure is notFound (a bad id) → caller-fixable → InvalidInputError (#41)
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });
});

// ---------- bulk set-error formatting (#22 + #41) ----------

describe('bulk set-error formatting', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  // bulkMove issues only the Email/set (the destination resolves off the stubbed mailbox
  // list), so a single stub drives it to a notUpdated map of our choosing.
  function stubBulkMoveFailure(notUpdated: Record<string, { type: string; description?: string }>) {
    stubMakeRequest(client, { methodResponses: [['Email/set', { notUpdated }, 'bulkMove']] });
  }

  it('groups failing ids by reason, surfacing type AND description (#22)', async () => {
    stubBulkMoveFailure({
      e1: { type: 'invalidArguments', description: 'bad patch' },
      e2: { type: 'invalidArguments', description: 'bad patch' },
    });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /Failed to move 2 of 2 emails \(0 succeeded\)/);
        // both ids share one reason → grouped under it together
        assert.match(err.message, /invalidArguments - bad patch: e1, e2/);
        // a malformed argument is fixed by re-forming the call → InvalidInputError
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('groups a server-side reason the same way, and keeps the batch a plain Error', async () => {
    // The grouping/description formatting is independent of the classification, so it is
    // pinned on both sides of the split rather than only on the caller-fixable one.
    stubBulkMoveFailure({
      e1: { type: 'overQuota', description: 'mailbox full' },
      e2: { type: 'overQuota', description: 'mailbox full' },
    });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /overQuota - mailbox full: e1, e2/);
        assert.equal(err.name, 'Error');
        return true;
      },
    );
  });

  it('throws InvalidInputError only when EVERY failure is caller-fixable (#41)', async () => {
    stubBulkMoveFailure({ e1: { type: 'notFound' }, e2: { type: 'notFound' } });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /notFound: e1, e2/);
        // type-only entries → no dangling " - "
        assert.doesNotMatch(err.message, / - /);
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('stays a plain Error when failures are MIXED (one operational keeps the whole batch InternalError)', async () => {
    stubBulkMoveFailure({ e1: { type: 'notFound' }, e2: { type: 'forbidden' } });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.equal(err.name, 'Error');
        return true;
      },
    );
  });

  it('treats a batch of assorted caller-fixable reasons as caller-fixable', async () => {
    // The batch rule is per-type, not "all the same type": a mix of reasons that are each
    // fixable by re-forming the call is still one re-form away from succeeding.
    stubBulkMoveFailure({ e1: { type: 'invalidProperties' }, e2: { type: 'invalidPatch' } });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  it('leaves an unrecognised reason on the server-side default', async () => {
    // A type this client has made no judgement about must not be presumed to be the
    // caller's fault; the default stays the plain Error.
    stubBulkMoveFailure({ e1: { type: 'someFutureJmapError' }, e2: { type: 'notFound' } });

    await assert.rejects(
      () => client.bulkMove(['e1', 'e2'], 'mb-archive'),
      (err: Error) => {
        assert.equal(err.name, 'Error');
        return true;
      },
    );
  });

  it('marks a truncated id list as partial and idempotent-retryable (never reads as complete)', async () => {
    // 12 ids under one reason exceeds the 10-per-reason cap.
    const ids = Array.from({ length: 12 }, (_, i) => `id${i}`);
    const notUpdated: Record<string, { type: string }> = {};
    ids.forEach(id => { notUpdated[id] = { type: 'notFound' }; });
    // bulkDelete issues only the Email/set (no pre-get), so a single stub suffices.
    stubMakeRequest(client, { methodResponses: [['Email/set', { notUpdated }, 'bulkDelete']] });

    await assert.rejects(
      () => client.bulkDelete(ids),
      (err: Error) => {
        assert.match(err.message, /Failed to delete 12 of 12 emails/);
        assert.match(err.message, /Partial list/);
        assert.match(err.message, /idempotent/);
        // the 11th/12th ids are NOT shown (capped at 10)
        assert.doesNotMatch(err.message, /id11/);
        return true;
      },
    );
  });
});

// ---------- getMethodResult ----------

describe('getMethodResult', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('throws on JMAP error response', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'serverFail', description: 'internal error' }, 'op'],
      ],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /serverFail/);
        assert.match(err.message, /internal error/);
        return true;
      },
    );
  });

  it('throws when index exceeds response length', async () => {
    stubMakeRequest(client, {
      methodResponses: [],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /missing expected method/i);
        return true;
      },
    );
  });

  it('throws on malformed entry (not an array)', async () => {
    stubMakeRequest(client, {
      methodResponses: ['not-a-tuple' as any],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /malformed/i);
        return true;
      },
    );
  });

  it('throws on error without description', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'unknownMethod' }, 'op'],
      ],
    });

    await assert.rejects(
      () => client.getEmailById('e1'),
      (err: Error) => {
        assert.match(err.message, /unknownMethod/);
        return true;
      },
    );
  });
});

// ---------- getListResult ----------

describe('getListResult', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('extracts list from valid response', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', { list: [INBOX_MAILBOX, DRAFTS_MAILBOX] }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.equal(mailboxes.length, 2);
  });

  it('returns empty array when list property is missing', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', { notList: 'something' }, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });

  it('returns empty array when result is null-ish', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Mailbox/get', null, 'mailboxes'],
      ],
    });

    const mailboxes = await client.getMailboxes();
    assert.deepEqual(mailboxes, []);
  });
});

// ---------- getThread ----------

describe('getThread', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('requests compact properties by default (no body data)', async () => {
    let callCount = 0;
    const makeReq = stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail'],
          ],
        };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
        ],
      };
    });

    await client.getThread('e1');

    const emailGetArgs = callArguments(makeReq, 1)[0].methodCalls[1][1];
    assert.ok(emailGetArgs.properties.includes('preview'), 'should request preview');
    assert.ok(emailGetArgs.properties.includes('inReplyTo'), 'should request inReplyTo');
    // textBody *structure* is requested for the bodyTextSize hint (#59), but no body
    // *content*: bodyValues / fetchTextBodyValues are absent, so this stays "no bodies".
    assert.ok(emailGetArgs.properties.includes('textBody'), 'should request textBody structure');
    assert.ok(!emailGetArgs.properties.includes('bodyValues'), 'should NOT request bodyValues');
    assert.ok(!emailGetArgs.properties.includes('htmlBody'), 'should NOT request htmlBody');
    assert.ok(!emailGetArgs.properties.includes('attachments'), 'should NOT request attachments');
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
    assert.equal(emailGetArgs.fetchHTMLBodyValues, undefined);
    assert.equal(emailGetArgs.bodyProperties, undefined);
  });

  it('fetches text body values (and only those) when includeBodies is set', async () => {
    let callCount = 0;
    const makeReq = stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail'],
          ],
        };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
        ],
      };
    });

    await client.getThread('e1', false, true);

    const emailGetArgs = callArguments(makeReq, 1)[0].methodCalls[1][1];
    // The defined VERBOSE superset, not a third ad-hoc property list.
    assert.ok(emailGetArgs.properties.includes('bodyValues'), 'should request bodyValues');
    assert.ok(emailGetArgs.properties.includes('textBody'), 'should request textBody structure');
    assert.equal(emailGetArgs.fetchTextBodyValues, true);
    // The html alternative is the expensive half of a thread read and is never fetched.
    assert.equal(emailGetArgs.fetchHTMLBodyValues, undefined);
    assert.ok(Array.isArray(emailGetArgs.bodyProperties), 'should request body properties');
  });

  it('throws InvalidInputError when the thread is not found (a bad id is caller-fixable input, #41)', async () => {
    let callCount = 0;
    stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        // probe resolves the email's threadId
        return { methodResponses: [['Email/get', { list: [{ threadId: 'thread-gone' }] }, 'checkEmail']] };
      }
      // Thread/get reports the thread missing
      return { methodResponses: [['Thread/get', { list: [], notFound: ['thread-gone'] }, 'getThread']] };
    });

    await assert.rejects(
      () => client.getThread('e1'),
      (err: Error) => {
        assert.match(err.message, /Thread with ID 'thread-gone' not found/);
        // the get_thread index local catch re-raises this BARE so the top-level maps it to InvalidParams
        assert.equal(err.name, 'InvalidInputError');
        return true;
      },
    );
  });

  // A thread containing a normal email plus an in-progress draft reply.
  const threadWithDraftResponse = () => {
    let callCount = 0;
    return stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail'],
          ],
        };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1', 'e2'] }] }, 'getThread'],
          ['Email/get', { list: [
            { id: 'e1', subject: 'Sent message', keywords: { $seen: true } },
            { id: 'e2', subject: 'Draft reply', keywords: { $draft: true } },
          ] }, 'emails'],
        ],
      };
    });
  };

  it('excludes draft messages by default and reports the hidden count', async () => {
    threadWithDraftResponse();
    const { emails, hiddenDraftCount } = await client.getThread('e1');
    assert.equal(emails.length, 1);
    assert.equal(emails[0].id, 'e1');
    assert.equal(hiddenDraftCount, 1);
  });

  it('includes drafts when includeDrafts is true (count 0)', async () => {
    threadWithDraftResponse();
    const { emails, hiddenDraftCount } = await client.getThread('e1', true);
    assert.equal(emails.length, 2);
    assert.deepEqual(emails.map((e: any) => e.id), ['e1', 'e2']);
    assert.equal(hiddenDraftCount, 0);
  });

  // A thread carrying an active draft reply plus a draft that now lives only in Trash —
  // the shape every edit_draft leaves behind, since the replaced copy keeps its $draft
  // keyword. `mailboxes` is what the batch's trailing Mailbox/get returns.
  const threadWithTrashedDraftResponse = (mailboxes: any[]) => {
    let callCount = 0;
    return stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return { methodResponses: [['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail']] };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1', 'e2', 'e3'] }] }, 'getThread'],
          ['Email/get', { list: [
            { id: 'e1', subject: 'Sent message', keywords: { $seen: true }, mailboxIds: { 'mb-inbox': true } },
            { id: 'e2', subject: 'Active draft reply', keywords: { $draft: true }, mailboxIds: { 'mb-drafts': true } },
            { id: 'e3', subject: 'Draft replaced by an edit', keywords: { $draft: true }, mailboxIds: { 'mb-trash': true } },
          ] }, 'emails'],
          ['Mailbox/get', { list: mailboxes }, 'mailboxes'],
        ],
      };
    });
  };

  const THREAD_MAILBOXES = [
    { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
    { id: 'mb-drafts', name: 'Drafts', role: 'drafts' },
    { id: 'mb-trash', name: 'Trash', role: 'trash' },
  ];

  it('does not count a draft that now lives only in Trash', async () => {
    threadWithTrashedDraftResponse(THREAD_MAILBOXES);
    const { emails, hiddenDraftCount } = await client.getThread('e1');
    assert.deepEqual(emails.map((e: any) => e.id), ['e1']); // both drafts stay hidden
    assert.equal(hiddenDraftCount, 1);                      // only the active one is counted
  });

  it('still counts a draft that is in Trash AND another mailbox', async () => {
    let callCount = 0;
    stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return { methodResponses: [['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail']] };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [
            { id: 'e1', subject: 'Draft filed in two places', keywords: { $draft: true }, mailboxIds: { 'mb-trash': true, 'mb-drafts': true } },
          ] }, 'emails'],
          ['Mailbox/get', { list: THREAD_MAILBOXES }, 'mailboxes'],
        ],
      };
    });
    const { hiddenDraftCount } = await client.getThread('e1');
    assert.equal(hiddenDraftCount, 1);
  });

  it('counts every draft when the trash role cannot be resolved (fails toward over-warning)', async () => {
    // Same thread, but no mailbox carries the trash role — a folder merely NAMED "Trash"
    // must not be treated as one.
    threadWithTrashedDraftResponse([
      { id: 'mb-drafts', name: 'Drafts', role: 'drafts' },
      { id: 'mb-trash', name: 'Trash', role: null },
    ]);
    const { hiddenDraftCount } = await client.getThread('e1');
    assert.equal(hiddenDraftCount, 2);
  });

  it('drops a hidden draft together with its body when bodies are fetched', async () => {
    let callCount = 0;
    stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return { methodResponses: [['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail']] };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1', 'e2'] }] }, 'getThread'],
          ['Email/get', { list: [
            {
              id: 'e1', subject: 'Sent message', keywords: { $seen: true },
              textBody: [{ partId: 't1', type: 'text/plain' }],
              bodyValues: { t1: { value: 'the sent text' } },
            },
            {
              id: 'e2', subject: 'Draft reply', keywords: { $draft: true },
              textBody: [{ partId: 't2', type: 'text/plain' }],
              bodyValues: { t2: { value: 'unsent words nobody should read yet' } },
            },
          ] }, 'emails'],
        ],
      };
    });

    const { emails, hiddenDraftCount } = await client.getThread('e1', false, true);

    assert.equal(hiddenDraftCount, 1);
    assert.deepEqual(emails.map((e: any) => e.id), ['e1']);
    assert.ok(
      !JSON.stringify(emails).includes('unsent words nobody should read yet'),
      "a hidden draft's body must not reach the caller",
    );
  });

  it('reports hiddenDraftCount 0 for a thread with no drafts', async () => {
    let callCount = 0;
    stubRequests(client, async () => {
      callCount++;
      if (callCount === 1) {
        return { methodResponses: [['Email/get', { list: [{ threadId: 'thread-1' }] }, 'checkEmail']] };
      }
      return {
        methodResponses: [
          ['Thread/get', { list: [{ id: 'thread-1', emailIds: ['e1'] }] }, 'getThread'],
          ['Email/get', { list: [{ id: 'e1', subject: 'Only message', keywords: { $seen: true } }] }, 'emails'],
        ],
      };
    });
    const { emails, hiddenDraftCount } = await client.getThread('e1');
    assert.equal(emails.length, 1);
    assert.equal(hiddenDraftCount, 0);
  });
});

// ---------- list method property checks ----------

describe('list method property checks', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const standardQueryResponse = {
    methodResponses: [
      ['Email/query', { ids: ['e1'], total: 1 }, 'query'],
      ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
    ],
  };

  // getEmails/searchEmails fetch getMailboxes() separately now, so stub it (mocked, not
  // via makeRequest) — makeRequest then only ever sees the Email/query batch as calls[0].
  function mockAndCall(method: string, callFn: () => Promise<any>) {
    mock.method(client, 'getMailboxes', async () => DEFAULT_MAILBOXES);
    const makeReq = stubRequests(client, async () => standardQueryResponse);
    return { makeReq, result: callFn() };
  }

  it('getEmails always requests compact properties (no bodies)', async () => {
    const { makeReq, result } = mockAndCall('getEmails', () => client.getEmails());
    await result;
    const emailGetArgs = callArguments(makeReq)[0].methodCalls[1][1];
    // textBody structure requested (bodyTextSize hint, #59); body content still not.
    assert.ok(emailGetArgs.properties.includes('textBody'), 'should request textBody structure');
    assert.ok(!emailGetArgs.properties.includes('bodyValues'), 'should NOT request bodyValues');
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('searchEmails always requests compact properties (no bodies)', async () => {
    const { makeReq, result } = mockAndCall('searchEmails', () => client.searchEmails({ query: 'test' }));
    await result;
    const emailGetArgs = callArguments(makeReq)[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('getRecentEmails always requests compact properties (no bodies)', async () => {
    mock.method(client, 'getMailboxes', async () => [INBOX_MAILBOX]);
    const makeReq = stubRequests(client, async () => standardQueryResponse);
    await client.getRecentEmails();
    const emailGetArgs = callArguments(makeReq)[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('searchEmails with structured filters still requests compact properties', async () => {
    const { makeReq, result } = mockAndCall('searchEmails', () => client.searchEmails({ query: 'test', from: 'a@b.example' }));
    await result;
    const emailGetArgs = callArguments(makeReq)[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });
});

// ---------- JMAP property consistency ----------

describe('JMAP property consistency', () => {
  // The property lists are readonly tuples of string literals, so a name a list does
  // NOT contain is not even an allowed argument to that list's own .includes(). The
  // assertions below are about absence and have to keep working — by failing — on the
  // day someone adds one of these names, so they read the list as plain strings.
  // Widening the list, rather than casting each argument, keeps the argument itself
  // checked as a string and leaves nothing to go stale silently.
  const compactProperties: readonly string[] = EMAIL_PROPERTIES_COMPACT;

  it('verbose properties are a superset of compact properties', () => {
    for (const prop of EMAIL_PROPERTIES_COMPACT) {
      assert.ok(
        EMAIL_PROPERTIES_VERBOSE.includes(prop),
        `verbose properties missing compact property: ${prop}`
      );
    }
  });

  it('verbose includes body-content properties that compact does not', () => {
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('htmlBody'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('bodyValues'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('attachments'));
    // Body *content* stays verbose-only; compact must never fetch it.
    assert.ok(!compactProperties.includes('htmlBody'));
    assert.ok(!compactProperties.includes('bodyValues'));
    assert.ok(!compactProperties.includes('attachments'));
  });

  it('verbose fetches both draft provenance headers; compact fetches neither', () => {
    // Get-path reads must surface which message a forward names (forwardedMessageId)
    // and which stored copy send_draft will mark (sourceEmailId); list items show
    // forward-ness via isForwarded instead.
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('header:X-Forwarded-Message-Id:asMessageIds'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('header:X-Fastmail-MCP-Source-Id:asText'));
    assert.ok(!compactProperties.includes('header:X-Forwarded-Message-Id:asMessageIds'));
    assert.ok(!compactProperties.includes('header:X-Fastmail-MCP-Source-Id:asText'));
  });

  it('compact includes textBody structure (part sizes) for the bodyTextSize hint (#59)', () => {
    // textBody is a compact property: it fetches the part *structure* (partId/type/size),
    // not content, so the response stays "no bodies" while exposing the text part size.
    assert.ok(EMAIL_PROPERTIES_COMPACT.includes('textBody'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('textBody'));
  });

  it('body properties include required fields', () => {
    assert.ok(EMAIL_BODY_PROPERTIES.includes('partId'));
    assert.ok(EMAIL_BODY_PROPERTIES.includes('blobId'));
    assert.ok(EMAIL_BODY_PROPERTIES.includes('type'));
    assert.ok(EMAIL_BODY_PROPERTIES.includes('size'));
    assert.ok(EMAIL_BODY_PROPERTIES.includes('name'));
  });
});

// ---------- ascending sort parameter ----------

describe('ascending sort parameter', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  const QUERY_GET_RESPONSE = {
    methodResponses: [
      ['Email/query', { ids: ['e1'] }, 'query'],
      ['Email/get', { list: [{ id: 'e1', subject: 'Test' }] }, 'emails'],
    ],
  };

  describe('getEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.getEmails({ mailbox: 'mb-inbox', limit: 5, ascending: true });

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('searchEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.searchEmails({ query: 'test', limit: 10 });

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.searchEmails({ query: 'test', limit: 10, ascending: true });

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('getRecentEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox');

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = stubRequests(client, async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox', true);

      const sort = callArguments(makeReq)[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });
});

// ---------- #10 mailbox-name resolution ----------

describe('mailbox location (#10)', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('COMPACT property set requests mailboxIds (propagates to every read path)', () => {
    assert.ok(EMAIL_PROPERTIES_COMPACT.includes('mailboxIds'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('mailboxIds')); // superset
  });

  // Names now come from the separately-fetched getMailboxes() list, NOT an in-batch
  // Mailbox/get (searchEmails/getEmails no longer append one).
  const NAME_MAILBOXES = [
    { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
    { id: 'mb-receipts', name: 'Receipts', role: null },
  ];
  const NAME_QUERY_RESPONSE = {
    methodResponses: [
      ['Email/query', { ids: ['e1', 'e2'], total: 2 }, 'query'],
      ['Email/get', { list: [
        { id: 'e1', subject: 'A', mailboxIds: { 'mb-inbox': true, 'mb-receipts': true } },
        { id: 'e2', subject: 'B', mailboxIds: { 'mb-unknown': true } }, // id not in the map → omit
      ] }, 'emails'],
    ],
  };

  it('getEmails attaches resolved names + roles from getMailboxes (multi-membership; unresolved surfaced)', async () => {
    mock.method(client, 'getMailboxes', async () => NAME_MAILBOXES);
    const makeReq = stubRequests(client, async () => NAME_QUERY_RESPONSE);

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });

    // No in-batch Mailbox/get — explicit mailbox => just query + get.
    const calls = callArguments(makeReq)[0].methodCalls;
    assert.equal(calls.length, 2);
    assert.equal(calls.some((c: any) => c[0] === 'Mailbox/get'), false);

    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox', 'Receipts']);
    // roles is an independent set — Receipts has role null, so only 'inbox'.
    assert.deepEqual((result.items[0] as any)._mailboxRoles, ['inbox']);
    // e2's only mailbox id didn't resolve → surfaced as a raw id, never silently dropped (#53).
    assert.equal('_mailboxNames' in (result.items[1] as any), false);
    assert.deepEqual((result.items[1] as any)._unresolvedMailboxIds, ['mb-unknown']);
  });

  it('surfaces an unresolved id (does NOT throw) when an email\'s mailbox id is not in the fetched list (#53)', async () => {
    mock.method(client, 'getMailboxes', async () => [{ id: 'mb-inbox', name: 'Inbox', role: 'inbox' }]);
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/query', { ids: ['e1'], total: 1 }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'A', mailboxIds: { 'mb-other': true } }] }, 'emails'],
      ],
    });

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });
    assert.equal(result.items.length, 1);
    assert.equal('_mailboxNames' in (result.items[0] as any), false);
    assert.deepEqual((result.items[0] as any)._unresolvedMailboxIds, ['mb-other']);
  });

  it('searchEmails attaches names + roles and surfaces an unresolved id from getMailboxes (#53)', async () => {
    mock.method(client, 'getMailboxes', async () => NAME_MAILBOXES);
    const makeReq = stubRequests(client, async () => NAME_QUERY_RESPONSE);
    // includeTrash/includeSpam true => no exclusion/count query, just query + get.
    const result = await client.searchEmails({ query: 'x', limit: 5, includeTrash: true, includeSpam: true });
    assert.equal(callArguments(makeReq)[0].methodCalls.length, 2);
    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox', 'Receipts']);
    assert.deepEqual((result.items[0] as any)._mailboxRoles, ['inbox']);
    // e2's id is not in the map → surfaced, not dropped.
    assert.deepEqual((result.items[1] as any)._unresolvedMailboxIds, ['mb-unknown']);
  });

  it('getEmailById attaches names + roles from its appended Mailbox/get (index 1), requesting role', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/get', { list: [{ id: 'e1', subject: 'A', mailboxIds: { 'mb-trash': true } }] }, 'email'],
        ['Mailbox/get', { list: [{ id: 'mb-trash', name: 'Trash', role: 'trash' }] }, 'mailboxes'],
      ],
    }));

    const email = await client.getEmailById('e1');
    const mailboxGet = callArguments(makeReq)[0].methodCalls[1];
    assert.equal(mailboxGet[0], 'Mailbox/get');
    assert.deepEqual(mailboxGet[1].properties, ['id', 'name', 'role']);
    assert.deepEqual((email as any)._mailboxNames, ['Trash']);
    assert.deepEqual((email as any)._mailboxRoles, ['trash']);
  });

  it('getRecentEmails reuses its existing getMailboxes list (no third methodCall) and attaches names + roles', async () => {
    stubMailboxes(client); // INBOX/DRAFTS/TRASH/SENT
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'A', mailboxIds: { 'mb-inbox': true } }] }, 'emails'],
      ],
    }));

    const result = await client.getRecentEmails(10, 'inbox');
    // Reuses getMailboxes — the request stays a 2-call batch, no appended Mailbox/get.
    assert.equal(callArguments(makeReq)[0].methodCalls.length, 2);
    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox']);
    assert.deepEqual((result.items[0] as any)._mailboxRoles, ['inbox']);
  });

  it('getThread attaches names + roles to retained messages before the draft filter, requesting role', async () => {
    const makeReq = stubRequests(client, async () => ({
      methodResponses: [
        ['Thread/get', { list: [{ id: 't1', emailIds: ['e1', 'e2'] }] }, 'getThread'],
        ['Email/get', { list: [
          { id: 'e1', subject: 'Kept', mailboxIds: { 'mb-inbox': true } },
          { id: 'e2', subject: 'Draft', keywords: { $draft: true }, mailboxIds: { 'mb-drafts': true } },
        ] }, 'emails'],
        ['Mailbox/get', { list: [
          { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
          { id: 'mb-drafts', name: 'Drafts', role: 'drafts' },
        ] }, 'mailboxes'],
      ],
    }));

    const { emails } = await client.getThread('t1'); // drafts excluded by default
    // getThread makes two requests: an email-id probe, then the Thread/Email/Mailbox batch.
    const mailboxGet = callArguments(makeReq, 1)[0].methodCalls[2];
    assert.equal(mailboxGet[0], 'Mailbox/get');
    assert.deepEqual(mailboxGet[1].properties, ['id', 'name', 'role']);
    assert.equal(emails.length, 1);
    assert.equal(emails[0].id, 'e1');
    assert.deepEqual((emails[0] as any)._mailboxNames, ['Inbox']);
    assert.deepEqual((emails[0] as any)._mailboxRoles, ['inbox']);
  });

  describe('buildMailboxInfoMap', () => {
    it('maps id -> {name, role}, keying on the real name (custom labels included)', () => {
      const map = buildMailboxInfoMap([
        { id: 'a', name: 'Inbox', role: 'inbox' },
        { id: 'b', name: 'My Label', role: null }, // role null but still mapped
      ]);
      assert.deepEqual(map.get('a'), { name: 'Inbox', role: 'inbox' });
      assert.deepEqual(map.get('b'), { name: 'My Label', role: null });
    });

    it('lowercases the role (docs promise lowercase regardless of server casing)', () => {
      const map = buildMailboxInfoMap([{ id: 'a', name: 'Spam', role: 'Junk' }]);
      assert.equal(map.get('a')!.role, 'junk');
    });

    it('returns an empty map for [] (the degradation input)', () => {
      assert.equal(buildMailboxInfoMap([]).size, 0);
    });

    it('skips entries lacking an id or a string name', () => {
      const map = buildMailboxInfoMap([{ id: 'a' }, { name: 'x' }, null as any]);
      assert.equal(map.size, 0);
    });
  });

  describe('attachMailboxInfo', () => {
    const map = new Map([
      ['a', { name: 'Inbox', role: 'inbox' }],
      ['b', { name: 'Receipts', role: null }],
    ]);

    it('attaches non-enumerable _mailboxNames + _mailboxRoles (absent from JSON, readable directly)', () => {
      const email: any = { id: 'e', mailboxIds: { a: true, b: true } };
      attachMailboxInfo([email], map);
      assert.deepEqual(email._mailboxNames, ['Inbox', 'Receipts']);
      // Independent sets: Receipts has role null, so only 'inbox'.
      assert.deepEqual(email._mailboxRoles, ['inbox']);
      const json = JSON.stringify(email);
      assert.equal(json.includes('_mailboxNames'), false);
      assert.equal(json.includes('_mailboxRoles'), false);
    });

    it('surfaces an unresolved id in _unresolvedMailboxIds (never silently dropped, never throws) (#53)', () => {
      const email: any = { id: 'e', mailboxIds: { z: true } };
      attachMailboxInfo([email], map);
      assert.equal('_mailboxNames' in email, false);
      assert.deepEqual(email._unresolvedMailboxIds, ['z']);
      assert.equal(JSON.stringify(email).includes('_unresolvedMailboxIds'), false);
    });

    it('partial: resolves the known id and surfaces the unknown one (#53)', () => {
      const email: any = { id: 'e', mailboxIds: { a: true, z: true } };
      attachMailboxInfo([email], map);
      assert.deepEqual(email._mailboxNames, ['Inbox']);
      assert.deepEqual(email._mailboxRoles, ['inbox']);
      assert.deepEqual(email._unresolvedMailboxIds, ['z']);
    });

    it('dropped trailing Mailbox/get (empty map): every id surfaced, no throw (#53)', () => {
      const email: any = { id: 'e', mailboxIds: { a: true, z: true } };
      attachMailboxInfo([email], new Map()); // [] from readListResultIfPresent
      assert.equal('_mailboxNames' in email, false);
      assert.equal('_mailboxRoles' in email, false);
      assert.deepEqual(email._unresolvedMailboxIds, ['a', 'z']);
    });

    it('omits every field when mailboxIds is absent or empty (NOT an unresolved case)', () => {
      const noIds: any = { id: 'e' };
      const emptyIds: any = { id: 'e', mailboxIds: {} };
      attachMailboxInfo([noIds, emptyIds], map);
      assert.equal('_mailboxNames' in noIds, false);
      assert.equal('_unresolvedMailboxIds' in noIds, false);
      assert.equal('_mailboxNames' in emptyIds, false);
      assert.equal('_unresolvedMailboxIds' in emptyIds, false);
    });
  });
});

// ---------- resolveMailbox (exact-only) ----------

describe('resolveMailbox', () => {
  const mbs = [
    { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
    { id: 'mb-archive', name: 'Archive', role: 'archive' },
    { id: 'mb-receipts', name: 'Receipts', role: null },
    { id: 'mb-junkrules', name: 'Junk mail rules', role: null },
  ];

  it('resolves by exact id', () => {
    assert.equal(resolveMailbox(mbs, 'mb-receipts').id, 'mb-receipts');
  });

  it('resolves by role (case-insensitive)', () => {
    assert.equal(resolveMailbox(mbs, 'INBOX').id, 'mb-inbox');
  });

  it('resolves by exact name (case-insensitive)', () => {
    assert.equal(resolveMailbox(mbs, 'receipts').id, 'mb-receipts');
  });

  it('does NOT substring-match (a partial name throws)', () => {
    assert.throws(() => resolveMailbox(mbs, 'arch'), (err: Error) => {
      assert.ok(err instanceof InvalidInputError);
      return true;
    });
  });

  it('throws InvalidInputError with a valid list when not found', () => {
    assert.throws(() => resolveMailbox(mbs, 'nope'), (err: Error) => {
      assert.ok(err instanceof InvalidInputError);
      assert.match(err.message, /not found/);
      assert.match(err.message, /Inbox \(inbox\)/);
      return true;
    });
  });

  it('caps the listed names and points at list_mailboxes', () => {
    const many = Array.from({ length: 40 }, (_, i) => ({ id: `mb-${i}`, name: `Folder ${i}`, role: null }));
    assert.throws(() => resolveMailbox(many, 'nope'), (err: Error) => {
      assert.match(err.message, /and \d+ more — call list_mailboxes/);
      return true;
    });
  });
});

// ---------- computeExclusion (exact role only) ----------

describe('computeExclusion', () => {
  const withRoles = [
    { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
    { id: 'mb-trash', name: 'Trash', role: 'trash' },
    { id: 'mb-junk', name: 'Spam', role: 'junk' },
    { id: 'mb-junkrules', name: 'Junk mail rules', role: null }, // must NOT be mis-hit as junk
  ];

  it('excludes trash + junk by exact role when nothing is included', () => {
    const r = computeExclusion(withRoles, {});
    assert.deepEqual([...r.excludeIds].sort(), ['mb-junk', 'mb-trash']);
    assert.deepEqual(r.excludedRoles, ['Trash', 'Spam']);
    assert.deepEqual(r.unresolvedRoles, []);
  });

  it('does NOT substring-mis-hit a custom "Junk mail rules" mailbox', () => {
    const r = computeExclusion(withRoles, {});
    assert.ok(!r.excludeIds.includes('mb-junkrules'));
  });

  it('includeTrash/includeSpam drop the respective ids', () => {
    assert.deepEqual(computeExclusion(withRoles, { includeTrash: true }).excludedRoles, ['Spam']);
    assert.deepEqual(computeExclusion(withRoles, { includeSpam: true }).excludedRoles, ['Trash']);
  });

  it('an explicit mailbox scope disables exclusion entirely', () => {
    const r = computeExclusion(withRoles, { hasExplicitScope: true });
    assert.deepEqual(r.excludeIds, []);
    assert.deepEqual(r.excludedRoles, []);
    assert.deepEqual(r.unresolvedRoles, []);
  });

  it('a role the caller already excluded is dropped from both the ids and the note roles', () => {
    const r = computeExclusion(withRoles, { callerExcludedIds: ['mb-trash'] });
    // The id would otherwise be sent twice (the caller's copy is in the union already),
    // and the role would make the note prescribe includeTrash:true — which cannot reveal
    // those messages while the caller's own exclusion still hides them.
    assert.deepEqual(r.excludeIds, ['mb-junk']);
    assert.deepEqual(r.excludedRoles, ['Spam']);
    assert.deepEqual(r.unresolvedRoles, []);
  });

  it('caller-excluding both roles leaves nothing for the default exclusion to add', () => {
    const r = computeExclusion(withRoles, { callerExcludedIds: ['mb-trash', 'mb-junk'] });
    assert.deepEqual(r.excludeIds, []);
    assert.deepEqual(r.excludedRoles, []);
    assert.deepEqual(r.unresolvedRoles, []);
  });

  it('a caller exclusion of an unrelated mailbox changes nothing', () => {
    const r = computeExclusion(withRoles, { callerExcludedIds: ['mb-inbox'] });
    assert.deepEqual([...r.excludeIds].sort(), ['mb-junk', 'mb-trash']);
    assert.deepEqual(r.excludedRoles, ['Trash', 'Spam']);
  });

  it('a missing role is flagged unresolved (fail-loud), not silently included', () => {
    const r = computeExclusion([{ id: 'mb-inbox', name: 'Inbox', role: 'inbox' }], {});
    assert.deepEqual(r.excludeIds, []);
    assert.deepEqual(r.unresolvedRoles, ['Trash', 'Spam']);
  });

  it('an empty/degraded mailbox list flags both roles unresolved', () => {
    const r = computeExclusion([], {});
    assert.deepEqual(r.unresolvedRoles, ['Trash', 'Spam']);
  });
});

// ---------- searchEmails exclusion + hidden-count semantics ----------

describe('searchEmails exclusion + count', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client); // DEFAULT_MAILBOXES has both trash + junk roles
  });

  it('explicit mailbox => inMailbox, no exclusion, no count query, no metadata', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 1 }));
    const result = await client.searchEmails({ query: 'x', mailbox: 'inbox' });
    const batch = callArguments(makeReq)[0].methodCalls;
    // text + inMailbox both live in the single base FilterCondition (no keyword conds
    // => no AND-wrap).
    const filter = batch[0][1].filter;
    assert.equal(filter.text, 'x');
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.inMailboxOtherThan, undefined);
    assert.equal(batch.length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('explicit mailbox + includeSpam:true => inMailbox only (Spam NOT OR-d back in)', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ mailbox: 'inbox', includeSpam: true });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.inMailboxOtherThan, undefined);
  });

  it('isUnread:false => hasKeyword $seen; isPinned:false => notKeyword $flagged (polarity kept)', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ isUnread: false, isPinned: false, includeTrash: true, includeSpam: true });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$seen'));
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$flagged'));
  });

  it('isUnread:true + isPinned:true => two separate keyword conditions', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ isUnread: true, isPinned: true, includeTrash: true, includeSpam: true });
    const filter = callArguments(makeReq)[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$seen'));
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$flagged'));
  });

  it('hidden = broaderTotal - visibleTotal', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8, broaderTotal: 11 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 3);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Trash', 'Spam']);
  });

  it('fires hidden>0 even when the visible result set is empty', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 4 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 4);
  });

  it('hidden === 0 when nothing was withheld', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 5, broaderTotal: 5 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 0);
  });

  it('FAIL-CLOSED: an absent count total => hidden:null (degraded), never silence', async () => {
    // Exclusion is active (so a count methodCall is in the request), but the mocked
    // response omits the count entry -> the count read throws -> hidden:null.
    stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, null);
  });

  it('FAIL-CLOSED: a negative hidden (bad broader total) => hidden:null', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8, broaderTotal: 0 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, null);
  });

  it('a missing junk/trash role surfaces in unresolvedRoles (no silent inclusion)', async () => {
    stubMailboxes(client, [{ id: 'mb-inbox', name: 'Inbox', role: 'inbox' }]);
    stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.deepEqual(result.exclusion?.unresolvedRoles, ['Trash', 'Spam']);
    assert.deepEqual(result.exclusion?.excludedRoles, []);
  });

  it('count filter keeps the keyword conds but drops inMailboxOtherThan (differs from visible)', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }));
    await client.searchEmails({ query: 'x', excludeDrafts: true });
    const batch = callArguments(makeReq)[0].methodCalls;
    const countFilter = batch[2][1].filter;
    // $draft cond survives; inMailboxOtherThan is gone -> count filter != visible filter.
    const flatHasDraft = countFilter.notKeyword === '$draft'
      || (countFilter.conditions || []).some((c: any) => c.notKeyword === '$draft');
    assert.ok(flatHasDraft);
    const stringified = JSON.stringify(countFilter);
    assert.ok(!stringified.includes('inMailboxOtherThan'));
  });
});

// ---------- searchEmails multi-mailbox scoping: requiredMailboxes / excludeMailboxes ----------

describe('searchEmails requiredMailboxes + excludeMailboxes', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  const filterOf = (makeReq: ReturnType<typeof stubRequests>) =>
    callArguments(makeReq)[0].methodCalls[0][1].filter;
  const batchOf = (makeReq: ReturnType<typeof stubRequests>) =>
    callArguments(makeReq)[0].methodCalls;

  // JMAP's inMailbox is singular, so "in all of these" is N AND-ed conditions.
  it('two requiredMailboxes become two separate inMailbox conditions', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ requiredMailboxes: ['Archive', 'Sent'] });
    const filter = filterOf(makeReq);
    assert.equal(filter.operator, 'AND');
    const inMailboxes = filter.conditions.filter((c: any) => c.inMailbox).map((c: any) => c.inMailbox);
    assert.deepEqual([...inMailboxes].sort(), ['mb-archive', 'mb-sent']);
    // requiredMailboxes is an explicit scope: default exclusion off, no count query,
    // no exclusion metadata.
    assert.equal(JSON.stringify(filter).includes('inMailboxOtherThan'), false);
    assert.equal(batchOf(makeReq).length, 2);
  });

  it('mailbox and requiredMailboxes fold into one intersection', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ mailbox: 'inbox', requiredMailboxes: ['Archive'] });
    const filter = filterOf(makeReq);
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.inMailbox === 'mb-inbox'));
    assert.ok(filter.conditions.some((c: any) => c.inMailbox === 'mb-archive'));
  });

  // The combine() shortcut collapses an empty base plus a single condition down to that
  // condition alone. The caller's exclusion has to be in `base` before that is decided,
  // or a query with no text/from fields would ship without it — fail-open on the
  // parameter whose whole job is narrowing.
  it('excludeMailboxes survives an otherwise-empty base with a single keyword condition', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }));
    await client.searchEmails({ isUnread: true, excludeMailboxes: ['Archive'] });
    const filter = filterOf(makeReq);
    assert.equal(filter.operator, 'AND');
    const exclusionCond = filter.conditions.find((c: any) => c.inMailboxOtherThan);
    assert.ok(exclusionCond, 'the exclusion must not be dropped by the single-condition shortcut');
    assert.ok(exclusionCond.inMailboxOtherThan.includes('mb-archive'));
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$seen'));
  });

  // The discriminating form of the case above: with the default exclusion OFF, the
  // caller's ids are the ONLY thing in `base`, so a mis-ordering that computed baseEmpty
  // first would collapse the filter to the lone keyword condition and ship no exclusion at
  // all. With the default active, the default ids mask that mistake.
  it('excludeMailboxes survives the empty-base shortcut with the default exclusion off', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({
      isUnread: true, excludeMailboxes: ['Archive'], includeTrash: true, includeSpam: true,
    });
    const filter = filterOf(makeReq);
    assert.equal(filter.operator, 'AND');
    assert.deepEqual(
      filter.conditions.find((c: any) => c.inMailboxOtherThan)?.inMailboxOtherThan,
      ['mb-archive'],
    );
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$seen'));
  });

  // The exclusion ids are assigned UNGATED, so turning the default Trash/Spam exclusion
  // off does not take the caller's own excludes with it.
  it('excludeMailboxes still reaches the filter with includeTrash + includeSpam set', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    const result = await client.searchEmails({
      query: 'x', excludeMailboxes: ['Archive'], includeTrash: true, includeSpam: true,
    });
    assert.deepEqual(filterOf(makeReq).inMailboxOtherThan, ['mb-archive']);
    // No default exclusion => no count query and no disclosure note metadata.
    assert.equal(batchOf(makeReq).length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('excludeMailboxes still reaches the filter under an explicit mailbox scope', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ mailbox: 'inbox', excludeMailboxes: ['Archive'] });
    const filter = filterOf(makeReq);
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.deepEqual(filter.inMailboxOtherThan, ['mb-archive']);
  });

  // One union set, not two conditions: a message in {Trash, an excluded label} is hidden
  // by the union even though neither exclusion alone would hide it.
  it('caller excludes are unioned with the default Trash/Spam ids', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 2, broaderTotal: 5 }));
    const result = await client.searchEmails({ query: 'x', excludeMailboxes: ['Archive'] });
    assert.deepEqual(
      [...filterOf(makeReq).inMailboxOtherThan].sort(),
      ['mb-archive', 'mb-junk', 'mb-trash'],
    );
    // The default exclusion is still active, so its note still fires.
    assert.equal(result.exclusion?.hidden, 3);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Trash', 'Spam']);
  });

  it('excludeMailboxes alone keeps the default exclusion and its note', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 4 }));
    const result = await client.searchEmails({ query: 'x', excludeMailboxes: ['Archive'] });
    const note = buildExclusionNote(result.exclusion);
    assert.ok(note.includes('4 message(s) in Trash/Spam were excluded'));
    assert.ok(note.includes('includeTrash:true / includeSpam:true'));
  });

  // Excluding Trash yourself takes that role out of the default set, so the note stops
  // prescribing includeTrash:true — a flag that could not override the caller's exclusion.
  it('caller-excluding Trash drops it from the note roles but keeps the id in the filter', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 1, broaderTotal: 3 }));
    const result = await client.searchEmails({ query: 'x', excludeMailboxes: ['trash'] });
    assert.deepEqual([...filterOf(makeReq).inMailboxOtherThan].sort(), ['mb-junk', 'mb-trash']);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Spam']);
    const note = buildExclusionNote(result.exclusion);
    assert.ok(note.includes('in Spam were excluded'));
    // BOTH halves of the recovery clause have to drop Trash, not just the flag half. The
    // mailbox override is the one that fails loudest if it survives: mailbox:"trash"
    // against inMailboxOtherThan:["mb-trash"] is a self-contradicting query that can only
    // return nothing.
    assert.ok(!note.includes('includeTrash:true'));
    assert.ok(!note.includes('"trash"'));
    assert.ok(note.includes('(or mailbox:"junk")'));
  });

  // The same derivation, with no caller exclusion involved: excluding only Spam must not
  // prescribe mailbox:"trash" either.
  it('the recovery clause names only the roles actually excluded', async () => {
    stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 1, broaderTotal: 3 }));
    const result = await client.searchEmails({ query: 'x', includeTrash: true });
    assert.deepEqual(result.exclusion?.excludedRoles, ['Spam']);
    const note = buildExclusionNote(result.exclusion);
    assert.ok(note.includes('set includeSpam:true (or mailbox:"junk")'));
    assert.ok(!note.includes('"trash"'));
  });

  it('requiredMailboxes and excludeMailboxes combine into one AND-ed filter', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({
      requiredMailboxes: ['Archive', 'Sent'],
      excludeMailboxes: ['Drafts'],
    });
    const filter = filterOf(makeReq);
    assert.equal(filter.operator, 'AND');
    const base = filter.conditions.find((c: any) => c.inMailboxOtherThan);
    assert.deepEqual(base.inMailboxOtherThan, ['mb-drafts']);
    const inMailboxes = filter.conditions.filter((c: any) => c.inMailbox).map((c: any) => c.inMailbox);
    assert.deepEqual([...inMailboxes].sort(), ['mb-archive', 'mb-sent']);
  });

  // The count query is what makes the hidden count mean "withheld to Trash/Spam". It
  // subtracts the DEFAULT ids only: subtracting the union would make it identical to the
  // visible query (hidden always 0, silence becomes fail-open), and subtracting nothing
  // would count caller-excluded matches the note then tells the caller to recover with
  // includeTrash/includeSpam, which cannot reveal them.
  it('the count query drops the default ids but keeps the caller ids', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 1, broaderTotal: 4 }));
    await client.searchEmails({ query: 'x', excludeMailboxes: ['Archive'] });
    const countFilter = batchOf(makeReq)[2][1].filter;
    assert.deepEqual(countFilter.inMailboxOtherThan, ['mb-archive']);
    assert.equal(countFilter.text, 'x');
  });

  it('the count query carries no inMailboxOtherThan when there are no caller excludes', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 1, broaderTotal: 4 }));
    await client.searchEmails({ query: 'x' });
    const countFilter = batchOf(makeReq)[2][1].filter;
    assert.equal('inMailboxOtherThan' in countFilter, false);
  });

  // doExclude also goes false on a degraded runtime path — neither role resolving — which
  // is not a caller choice at all. The ungated assignment has to cover that too.
  it('caller excludes survive a mailbox list where neither Trash nor Spam resolves', async () => {
    stubMailboxes(client, [
      { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
      { id: 'mb-archive', name: 'Archive', role: 'archive' },
    ]);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    const result = await client.searchEmails({ query: 'x', excludeMailboxes: ['Archive'] });
    assert.deepEqual(filterOf(makeReq).inMailboxOtherThan, ['mb-archive']);
    // The default exclusion was still INTENDED, so the fail-loud note still fires.
    assert.deepEqual(result.exclusion?.unresolvedRoles, ['Trash', 'Spam']);
    assert.deepEqual(result.exclusion?.excludedRoles, []);
    assert.equal(result.exclusion?.hidden, 0);
  });

  it('an entry that resolves to nothing rejects the whole call', async () => {
    stubRequests(client, async () => queryResponse({ ids: [], list: [], total: 0 }));
    await assert.rejects(
      () => client.searchEmails({ query: 'x', requiredMailboxes: ['Archive', 'No Such Folder'] }),
      InvalidInputError,
    );
    await assert.rejects(
      () => client.searchEmails({ query: 'x', excludeMailboxes: ['No Such Folder'] }),
      InvalidInputError,
    );
  });

  // The scope arrays share the label arrays' resolver, so they inherit its single-retry
  // property: EVERY failing entry is named at once, rather than one per round trip.
  it('names every failing entry in one rejection', async () => {
    stubRequests(client, async () => queryResponse({ ids: [], list: [], total: 0 }));
    await assert.rejects(
      () => client.searchEmails({ requiredMailboxes: ['Nope1', 'Nope2'] }),
      (err: unknown) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Nope1/);
        assert.match(err.message, /Nope2/);
        return true;
      },
    );
  });

  // A typo and an ambiguous name need different corrections, so they stay in separate
  // buckets instead of both reading as "not found".
  it('keeps a typo and an ambiguous name in separate buckets', async () => {
    stubMailboxes(client, [
      ...DEFAULT_MAILBOXES,
      { id: 'mb-dup-a', name: 'Receipts', parentId: 'mb-archive' },
      { id: 'mb-dup-b', name: 'Receipts', parentId: 'mb-inbox' },
    ]);
    stubRequests(client, async () => queryResponse({ ids: [], list: [], total: 0 }));
    await assert.rejects(
      () => client.searchEmails({ excludeMailboxes: ['Receipts', 'Nope'] }),
      (err: unknown) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Nope/);
        assert.match(err.message, /Receipts/);
        // The ambiguity is answered with the full paths that disambiguate it.
        assert.match(err.message, /Archive\/Receipts/);
        return true;
      },
    );
  });

  it('accepts every mailbox reference form, like the scalar mailbox param', async () => {
    // A nested mailbox so the root-anchored path form has something to resolve against.
    stubMailboxes(client, [
      ...DEFAULT_MAILBOXES,
      { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
    ]);
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    // id, role, name and root-anchored path in one call.
    await client.searchEmails({
      requiredMailboxes: ['mb-inbox', 'archive', 'Archive/2026'],
      excludeMailboxes: ['Drafts'],
    });
    const filter = filterOf(makeReq);
    const inMailboxes = filter.conditions.filter((c: any) => c.inMailbox).map((c: any) => c.inMailbox);
    assert.deepEqual([...inMailboxes].sort(), ['mb-2026', 'mb-archive', 'mb-inbox']);
    assert.deepEqual(
      filter.conditions.find((c: any) => c.inMailboxOtherThan).inMailboxOtherThan,
      ['mb-drafts'],
    );
  });

  it('two entries naming the same mailbox collapse to one id', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({
      requiredMailboxes: ['mb-archive', 'Archive'],
      excludeMailboxes: ['Drafts', 'drafts'],
      includeTrash: true,
      includeSpam: true,
    });
    const filter = filterOf(makeReq);
    const inMailboxes = filter.conditions.filter((c: any) => c.inMailbox).map((c: any) => c.inMailbox);
    assert.deepEqual(inMailboxes, ['mb-archive']);
    assert.deepEqual(
      filter.conditions.find((c: any) => c.inMailboxOtherThan).inMailboxOtherThan,
      ['mb-drafts'],
    );
  });

  it('empty scope arrays behave exactly like omitting them', async () => {
    const makeReq = stubRequests(client, async () =>
      queryResponse({ ids: [], list: [], total: 1, broaderTotal: 1 }));
    const result = await client.searchEmails({ query: 'x', requiredMailboxes: [], excludeMailboxes: [] });
    assert.deepEqual([...filterOf(makeReq).inMailboxOtherThan].sort(), ['mb-junk', 'mb-trash']);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Trash', 'Spam']);
  });
});

// ---------- getMailboxStats resolution ----------

describe('getMailboxStats resolution', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('resolves a named mailbox and reads stats off the fetched list', async () => {
    mock.method(client, 'getMailboxes', async () => [
      { id: 'mb-inbox', name: 'Inbox', role: 'inbox', totalEmails: 42, unreadEmails: 5, totalThreads: 30, unreadThreads: 3 },
    ]);
    const stats = await client.getMailboxStats('Inbox');
    assert.equal(stats.id, 'mb-inbox');
    assert.equal(stats.totalEmails, 42);
    assert.equal(stats.unreadEmails, 5);
  });

  it('throws InvalidInputError on an unknown mailbox', async () => {
    mock.method(client, 'getMailboxes', async () => [{ id: 'mb-inbox', name: 'Inbox', role: 'inbox' }]);
    await assert.rejects(
      () => client.getMailboxStats('nope'),
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
  });

  it('returns all mailboxes when no argument is given', async () => {
    mock.method(client, 'getMailboxes', async () => DEFAULT_MAILBOXES);
    const stats = await client.getMailboxStats();
    assert.ok(Array.isArray(stats));
    assert.equal(stats.length, DEFAULT_MAILBOXES.length);
  });
});

// ---------- bulkMove resolution ----------

describe('bulkMove resolution', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  it('resolves the destination by name and replaces membership whole-value', async () => {
    const makeReq = stubRequests(client, async () => (
      { methodResponses: [['Email/set', { updated: { e1: null } }, 'bulkMove']] }
    ));
    await client.bulkMove(['e1'], 'Archive');
    assert.equal(makeReq.mock.callCount(), 1, 'no pre-read round trip');
    const update = callArguments(makeReq, 0)[0].methodCalls[0][1].update;
    // Same shape as moveEmail: one whole-value mailboxIds per email, and no keyword.
    assert.deepEqual(update.e1, { mailboxIds: { 'mb-archive': true } });
  });

  it('throws InvalidInputError on an unknown destination', async () => {
    await assert.rejects(
      () => client.bulkMove(['e1'], 'nope'),
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
  });
});

// ---------- label tools resolve mailbox inputs by id/role/name (#50) ----------

// A name-only mailbox (no role) so name resolution is provably distinct from the role
// branch — "Archive" resolves via role, so it does NOT exercise name matching.
const RECEIPTS_MAILBOX = { id: 'mb-receipts', name: 'Receipts' };
const LABEL_MAILBOXES = [...DEFAULT_MAILBOXES, RECEIPTS_MAILBOX];

describe('label mailboxId resolution (#50)', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client, LABEL_MAILBOXES);
  });

  function stubSet(c: JmapClient, callId: string) {
    return stubRequests(c, async () =>
      ({ methodResponses: [['Email/set', { updated: { e1: null } }, callId]] }));
  }

  it('resolves a ROLE to its id and emits the mailboxIds/<id> patch', async () => {
    const makeReq = stubSet(client, 'addLabels');
    await client.addLabels('e1', ['archive']); // role
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-archive'], true);
  });

  it('resolves a NAME (name-only mailbox, no role) to its id', async () => {
    const makeReq = stubSet(client, 'addLabels');
    await client.addLabels('e1', ['Receipts']); // name, resolves via the name branch only
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-receipts'], true);
  });

  it('accepts a raw id (resolves to itself)', async () => {
    const makeReq = stubSet(client, 'addLabels');
    await client.addLabels('e1', ['mb-archive']);
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-archive'], true);
  });

  it('collapses a duplicate id + its own name into one patch key', async () => {
    const makeReq = stubSet(client, 'addLabels');
    // Receipts is name-only, so this exercises id + NAME collapse (not id + role): both
    // 'mb-receipts' and 'Receipts' resolve to the same id and dedupe to one patch key.
    await client.addLabels('e1', ['mb-receipts', 'Receipts']);
    const patch = callArguments(makeReq)[0].methodCalls[0][1].update.e1;
    const labelKeys = Object.keys(patch).filter(k => k.startsWith('mailboxIds/'));
    assert.deepEqual(labelKeys, ['mailboxIds/mb-receipts']);
  });

  it('names ALL unresolved values and applies nothing (all-or-nothing)', async () => {
    const makeReq = stubRequests(client, async () => { throw new Error('should not be called'); });
    await assert.rejects(
      () => client.addLabels('e1', ['nope', 'alsobad']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /'nope'/);
        assert.match(err.message, /'alsobad'/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 0); // rejected before any Email/set
  });

  it('rejects the whole call when one value of a mix is unresolved', async () => {
    const makeReq = stubRequests(client, async () => { throw new Error('should not be called'); });
    await assert.rejects(
      () => client.addLabels('e1', ['archive', 'nope']),
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
    assert.equal(makeReq.mock.calls.length, 0);
  });

  it('rejects a SUBSTRING input (exact-only, no injection-steering)', async () => {
    await assert.rejects(
      () => client.addLabels('e1', ['Arch']), // substring of Archive
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
  });

  it('caps the reflected unresolved list with a "…and N more" tail', async () => {
    const many = Array.from({ length: 32 }, (_, i) => `bad${i}`);
    await assert.rejects(
      () => client.addLabels('e1', many),
      (err: Error) => { assert.match(err.message, /…and 2 more/); return true; },
    );
  });

  it('clamps an over-long unresolved value in the error message', async () => {
    const long = 'x'.repeat(200);
    await assert.rejects(
      () => client.addLabels('e1', [long]),
      (err: Error) => {
        assert.ok(!err.message.includes(long)); // full value not reflected wholesale
        assert.match(err.message, /xxx…/);       // clamped with an ellipsis
        return true;
      },
    );
  });

  it('removeLabels resolves a role and emits the null patch', async () => {
    const makeReq = stubSet(client, 'removeLabels');
    await client.removeLabels('e1', ['archive']);
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-archive'], null);
  });

  it('bulkAddLabels resolves a name/role across the array', async () => {
    const makeReq = stubSet(client, 'bulkAddLabels');
    await client.bulkAddLabels(['e1'], ['Receipts', 'archive']);
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-receipts'], true);
    assert.equal(update.e1['mailboxIds/mb-archive'], true);
  });

  it('bulkAddLabels still rejects a genuinely-unresolvable value', async () => {
    await assert.rejects(
      () => client.bulkAddLabels(['e1'], ['nope']),
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
  });
});

// The label arrays call the matcher directly rather than going through resolveMailbox, so
// the path form has to be asserted here too: a form accepted on move_email but not in a
// mailboxIds entry would be exactly the split vocabulary this server resolves against one
// matcher to avoid.
describe('label mailboxId resolution accepts a path, and reports failures in separate buckets', () => {
  let client: JmapClient;

  // Two "Receipts" folders under different parents, plus one unambiguous nested folder.
  const NESTED_LABEL_MAILBOXES = [
    ...DEFAULT_MAILBOXES,
    { id: 'mb-personal', name: 'Personal' },
    { id: 'mb-work', name: 'Work' },
    { id: 'mb-personal-receipts', name: 'Receipts', parentId: 'mb-personal' },
    { id: 'mb-work-receipts', name: 'Receipts', parentId: 'mb-work' },
  ];

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client, NESTED_LABEL_MAILBOXES);
  });

  it('resolves a full path given as a mailboxIds entry', async () => {
    const makeReq = stubRequests(client, async () =>
      ({ methodResponses: [['Email/set', { updated: { e1: null } }, 'addLabels']] }));
    await client.addLabels('e1', ['Work/Receipts']);
    const update = callArguments(makeReq)[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-work-receipts'], true);
  });

  it('reports an ambiguous entry as ambiguous, never as "not found"', async () => {
    await assert.rejects(
      () => client.addLabels('e1', ['Receipts']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Ambiguous mailbox name\(s\)/);
        assert.match(err.message, /Personal\/Receipts, Work\/Receipts/);
        assert.doesNotMatch(err.message, /not found/);
        return true;
      },
    );
  });

  it('folds an ambiguous entry and a typo into ONE error, each in its own bucket', async () => {
    const makeReq = stubRequests(client, async () => { throw new Error('should not be called'); });
    await assert.rejects(
      () => client.addLabels('e1', ['Receipts', 'nope']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Mailbox\(es\) not found: 'nope'\./);
        assert.match(err.message, /Ambiguous mailbox name\(s\)[^.]*'Receipts' matches Personal\/Receipts, Work\/Receipts/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 0); // still all-or-nothing
  });

  it('says when a bucket list was truncated, so a capped list never reads as complete', async () => {
    // 31 mailboxes sharing one name: the candidate list is capped at 30 and has to say so.
    const many = Array.from({ length: 31 }, (_, i) => [
      { id: `mb-p${i}`, name: `P${i}` },
      { id: `mb-r${i}`, name: 'Receipts', parentId: `mb-p${i}` },
    ]).flat();
    client = makeClient();
    stubMailboxes(client, many);
    await assert.rejects(
      () => client.addLabels('e1', ['Receipts']),
      (err: Error) => {
        assert.match(err.message, /Ambiguous mailbox name\(s\)/);
        assert.match(err.message, /…and 1 more/);
        return true;
      },
    );
  });

  it('reports a name/path collision in its own bucket, never as "not found"', async () => {
    // The label arrays inherit the matcher, so the collision has to reach a caller here too —
    // and in a bucket of its own, because its correction (an id) differs from both a typo's
    // and a duplicated name's.
    client = makeClient();
    stubMailboxes(client, [
      ...DEFAULT_MAILBOXES,
      { id: 'mb-literal', name: 'A/B' },
      { id: 'mb-a', name: 'A' },
      { id: 'mb-b', name: 'B', parentId: 'mb-a' },
    ]);
    const makeReq = stubRequests(client, async () => { throw new Error('should not be called'); });
    await assert.rejects(
      () => client.addLabels('e1', ['A/B', 'nope']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Mailbox\(es\) not found: 'nope'\./);
        assert.match(
          err.message,
          /name a folder AND describe a path to a different mailbox[^.]*'A\/B' matches folder named 'A\/B' \(id: mb-literal\), nested folder A > B \(id: mb-b\)/,
        );
        assert.match(err.message, /retry with an id/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 0); // still all-or-nothing
  });

  it('reports an unwalkable path entry in its own bucket', async () => {
    client = makeClient();
    stubMailboxes(client, [...DEFAULT_MAILBOXES, ...LOOPED_TREE]);
    await assert.rejects(
      () => client.addLabels('e1', ['A/B']),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /never reaches a top-level mailbox/);
        assert.match(err.message, /mb-loop-a/);
        return true;
      },
    );
  });
});

// The single-input not-found message is pinned byte-for-byte, because it is the recovery
// instruction a caller acts on: it has to enumerate every input form the resolver actually
// accepts (id, role, name, path), and its "Valid:" list has to be written in the same
// vocabulary the ambiguity error uses, which is full paths. The label-array message is a
// superset of this one.
describe('resolveMailbox not-found message', () => {
  it('produces the exact single-input message', () => {
    assert.throws(
      () => resolveMailbox(DEFAULT_MAILBOXES, 'nope'),
      (err: Error) => {
        assert.equal(
          err.message,
          "Mailbox 'nope' not found. Use an id, a role (inbox/archive/sent/drafts/trash/junk), a name, or a full path (Parent/Child). " +
          'Valid: Inbox (inbox), Drafts (drafts), Trash (trash), Sent (sent), Archive (archive), Spam (junk)',
        );
        return true;
      },
    );
  });

  it('lists a nested mailbox by full path, so the hint speaks the form the ambiguity error hands back', () => {
    assert.throws(
      () => resolveMailbox(
        [
          { id: 'mb-archive', name: 'Archive', role: 'archive' },
          { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
        ],
        'nope',
      ),
      (err: Error) => {
        assert.match(err.message, /Valid: Archive \(archive\), Archive\/2026$/);
        return true;
      },
    );
  });
});

// ---------- findMailboxExact: the shared matcher both callers inherit ----------

// Every mailbox-taking parameter and every mailboxIds entry goes through this one
// function, so its resolution order and its failure shapes are asserted directly rather
// than only through whichever tool happens to call it.

// Two folders of the same name under different parents: the shape that makes a bare name
// ambiguous and a full path the only unique handle.
const DUPLICATE_NAME_TREE = [
  { id: 'mb-personal', name: 'Personal' },
  { id: 'mb-work', name: 'Work' },
  { id: 'mb-personal-receipts', name: 'Receipts', parentId: 'mb-personal' },
  { id: 'mb-work-receipts', name: 'Receipts', parentId: 'mb-work' },
];

// A parentId loop: no mailbox in it can be given a root-anchored path.
const LOOPED_TREE = [
  { id: 'mb-loop-a', name: 'A', parentId: 'mb-loop-b' },
  { id: 'mb-loop-b', name: 'B', parentId: 'mb-loop-a' },
];

describe('findMailboxExact', () => {
  it('resolves a nested mailbox by its root-anchored path', () => {
    const match = findMailboxExact(DUPLICATE_NAME_TREE, 'Personal/Receipts');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-personal-receipts');
  });

  it('matches path segments case-insensitively', () => {
    const match = findMailboxExact(DUPLICATE_NAME_TREE, 'personal/RECEIPTS');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-personal-receipts');
  });

  it('matches an id case-sensitively, so a case-folded id is not treated as a match', () => {
    const match = findMailboxExact(DEFAULT_MAILBOXES, 'MB-INBOX');
    assert.equal(match, undefined);
  });

  it('lets a unique flat name win when no other mailbox answers to the same text', () => {
    // A top-level folder literally named "A/B", with no A > B nesting. Its own computed path
    // is "A/B" too, so the name branch and the path branch land on the SAME mailbox — nothing
    // to be ambiguous about, and the folder stays reachable by the only name its owner knows
    // it by.
    const tree = [
      { id: 'mb-literal', name: 'A/B' },
      { id: 'mb-a', name: 'A' },
    ];
    const match = findMailboxExact(tree, 'A/B');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-literal');
  });

  it('resolves a flat name containing the separator that matches nothing by path', () => {
    // Here the folder named "A/B" is nested, so its path is "X/A/B" and the input matches
    // by name only. The name branch answers, exactly as it did before the collision check.
    const tree = [
      { id: 'mb-x', name: 'X' },
      { id: 'mb-literal', name: 'A/B', parentId: 'mb-x' },
    ];
    const match = findMailboxExact(tree, 'A/B');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-literal');
  });

  it('resolves a path that matches no flat name', () => {
    const tree = [
      { id: 'mb-a', name: 'A' },
      { id: 'mb-b', name: 'B', parentId: 'mb-a' },
    ];
    const match = findMailboxExact(tree, 'A/B');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-b');
  });

  it('reports a folder name that also reads as a path to a DIFFERENT mailbox as ambiguous', () => {
    // A folder literally named "A/B" alongside a real A > B nesting. Silently applying the
    // flat-name tie-break here files a write into the flat folder while the caller had every
    // reason to mean the nesting, and nothing in the response says a second mailbox answered
    // to the same text.
    const tree = [
      { id: 'mb-literal', name: 'A/B' },
      { id: 'mb-a', name: 'A' },
      { id: 'mb-b', name: 'B', parentId: 'mb-a' },
    ];
    const match = findMailboxExact(tree, 'A/B');
    assert.ok(match && 'ambiguous' in match);
    assert.equal(match.ambiguous, 'A/B');
    assert.equal(match.nameVsPath, true);
    // Both candidates carry the id, which is the only form that separates them, and each says
    // which mailbox it is — two renderings of the path "A/B" would be useless here.
    assert.deepEqual(match.candidates, [
      "folder named 'A/B' (id: mb-literal)",
      'nested folder A > B (id: mb-b)',
    ]);
    // Every id offered resolves when pasted back.
    for (const id of ['mb-literal', 'mb-b']) {
      const again = findMailboxExact(tree, id);
      assert.ok(again && 'mailbox' in again, id);
    }
  });

  it('reports the collision the same way for a deeper nesting', () => {
    const tree = [
      { id: 'mb-literal', name: 'Archive/2026' },
      { id: 'mb-archive', name: 'Archive', role: 'archive' },
      { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
    ];
    const match = findMailboxExact(tree, 'Archive/2026');
    assert.ok(match && 'ambiguous' in match);
    assert.equal(match.nameVsPath, true);
    assert.deepEqual(match.candidates, [
      "folder named 'Archive/2026' (id: mb-literal)",
      'nested folder Archive > 2026 (id: mb-2026)',
    ]);
  });

  it('reports a duplicated flat name as ambiguous, with the candidates as full paths', () => {
    const match = findMailboxExact(DUPLICATE_NAME_TREE, 'Receipts');
    assert.ok(match && 'ambiguous' in match);
    assert.equal(match.ambiguous, 'Receipts');
    assert.deepEqual(match.candidates, ['Personal/Receipts', 'Work/Receipts']);
  });

  it('reports an unwalkable parent chain rather than silently answering "not found"', () => {
    const match = findMailboxExact(LOOPED_TREE, 'A/B');
    assert.ok(match && 'unwalkable' in match);
    assert.equal(match.id, 'mb-loop-a');
  });

  it('reports an unwalkable chain when a parentId names a mailbox the account did not return', () => {
    const match = findMailboxExact([{ id: 'mb-orphan', name: 'Orphan', parentId: 'mb-gone' }], 'Gone/Orphan');
    assert.ok(match && 'unwalkable' in match);
    assert.equal(match.id, 'mb-orphan');
  });

  it('answers a healthy branch even while another branch is unwalkable', () => {
    const match = findMailboxExact([...DUPLICATE_NAME_TREE, ...LOOPED_TREE], 'Work/Receipts');
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-work-receipts');
  });

  it('rejects a path with a leading, trailing or doubled separator as a miss', () => {
    for (const input of ['/Personal/Receipts', 'Personal/Receipts/', 'Personal//Receipts']) {
      assert.equal(findMailboxExact(DUPLICATE_NAME_TREE, input), undefined, input);
    }
  });

  it('returns undefined for a genuine miss', () => {
    assert.equal(findMailboxExact(DUPLICATE_NAME_TREE, 'Nope/Nowhere'), undefined);
  });

  it('offers the ID, never the bare name, for a candidate that has no path', () => {
    // The name is what the caller just had rejected as ambiguous, so handing it back as a
    // candidate would send them round the same failure forever.
    const tree = [
      { id: 'mb-personal', name: 'Personal' },
      { id: 'mb-personal-receipts', name: 'Receipts', parentId: 'mb-personal' },
      { id: 'mb-loose', name: 'Receipts', parentId: 'mb-gone' },
    ];
    const match = findMailboxExact(tree, 'Receipts');
    assert.ok(match && 'ambiguous' in match);
    assert.deepEqual(match.candidates, ['Personal/Receipts', 'mb-loose']);
    // Every candidate offered has to resolve when pasted back.
    for (const candidate of match.candidates) {
      const again = findMailboxExact(tree, candidate);
      assert.ok(again && 'mailbox' in again, candidate);
    }
  });

  it('does not blame an unrelated broken chain for an ordinary path typo', () => {
    // One orphan elsewhere in the account must not turn every path miss into "paths cannot
    // be computed" — in this very tree, Archive/2026 resolves.
    const tree = [
      { id: 'mb-archive', name: 'Archive', role: 'archive' },
      { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
      { id: 'mb-orphan', name: 'Orphan', parentId: 'mb-gone' },
    ];
    assert.equal(findMailboxExact(tree, 'Archive/Recepts'), undefined);
    const good = findMailboxExact(tree, 'Archive/2026');
    assert.ok(good && 'mailbox' in good);
    assert.throws(
      () => resolveMailbox(tree, 'Archive/Recepts'),
      (err: Error) => {
        assert.match(err.message, /not found/);
        assert.doesNotMatch(err.message, /cannot be computed|never reaches/);
        return true;
      },
    );
  });

  it('blames the broken chain when the caller actually named the unpathable mailbox', () => {
    const tree = [
      { id: 'mb-archive', name: 'Archive', role: 'archive' },
      { id: 'mb-orphan', name: 'Orphan', parentId: 'mb-gone' },
    ];
    const match = findMailboxExact(tree, 'Archive/Orphan');
    assert.ok(match && 'unwalkable' in match);
    assert.equal(match.id, 'mb-orphan');
  });

  it('round-trips a path through a name carrying stray whitespace', () => {
    const tree = [
      { id: 'mb-work', name: 'Work' },
      { id: 'mb-q1', name: ' Q1', parentId: 'mb-work' },
    ];
    const { paths } = buildMailboxPathMap(tree);
    assert.equal(paths.get('mb-q1'), 'Work/Q1');
    const match = findMailboxExact(tree, paths.get('mb-q1')!);
    assert.ok(match && 'mailbox' in match);
    assert.equal(match.mailbox.id, 'mb-q1');
    // The same normalisation reaches the flat-name branch, so the name a caller would type
    // matches the padded name the server stores.
    const byName = findMailboxExact(tree, 'Q1');
    assert.ok(byName && 'mailbox' in byName);
    assert.equal(byName.mailbox.id, 'mb-q1');
  });

  it('treats a blank or missing name as unpathable rather than emitting an empty segment', () => {
    // A blank root would otherwise yield "/Kid", which the path parser rejects outright —
    // an emitted path that the resolver will not take back.
    const tree = [
      { id: 'mb-root', name: '   ' },
      { id: 'mb-kid', name: 'Kid', parentId: 'mb-root' },
      { id: 'mb-nameless', parentId: null },
    ];
    const { paths, unpathable } = buildMailboxPathMap(tree);
    assert.equal(paths.size, 0);
    assert.deepEqual(unpathable.sort(), ['mb-kid', 'mb-nameless', 'mb-root']);
    for (const emitted of paths.values()) {
      assert.doesNotMatch(emitted, /(^\/)|(\/\/)|(\/$)/);
    }
  });
});

describe('resolveMailbox failure shapes', () => {
  it('gives an ambiguous name its own message naming the candidate paths', () => {
    assert.throws(
      () => resolveMailbox(DUPLICATE_NAME_TREE, 'Receipts'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /is ambiguous: 2 mailboxes share that name/);
        assert.match(err.message, /Candidates: Personal\/Receipts, Work\/Receipts/);
        // A caller told "not found" would go re-spell a name that was already right.
        assert.doesNotMatch(err.message, /not found/);
        return true;
      },
    );
  });

  it('tells a name/path collision to retry with an id, not with a path', () => {
    // The generic ambiguity advice ("retry with one of their full paths") is the input that
    // just failed here, so this failure gets its own message.
    const tree = [
      { id: 'mb-literal', name: 'A/B' },
      { id: 'mb-a', name: 'A' },
      { id: 'mb-b', name: 'B', parentId: 'mb-a' },
    ];
    assert.throws(
      () => resolveMailbox(tree, 'A/B'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /is ambiguous: it is both the name of one folder and the path to a different mailbox/);
        assert.match(err.message, /Retry with the id of the one you mean/);
        assert.match(err.message, /Candidates: folder named 'A\/B' \(id: mb-literal\), nested folder A > B \(id: mb-b\)/);
        assert.doesNotMatch(err.message, /not found/);
        return true;
      },
    );
  });

  it('gives an unwalkable tree a message distinct from both ambiguity and a miss', () => {
    assert.throws(
      () => resolveMailbox(LOOPED_TREE, 'A/B'),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /never reaches a top-level mailbox/);
        assert.match(err.message, /mb-loop-a/);
        assert.doesNotMatch(err.message, /not found/);
        assert.doesNotMatch(err.message, /ambiguous/);
        return true;
      },
    );
  });
});

// ---------- the path a listing returns is the path a parameter accepts ----------

describe('mailbox path round-trip', () => {
  it('resolves a path taken straight out of the simplified listing back to that mailbox', () => {
    const { paths } = buildMailboxPathMap(DUPLICATE_NAME_TREE);
    for (const mb of DUPLICATE_NAME_TREE) {
      const listed = simplifyMailbox(mb, { path: paths.get(mb.id) });
      assert.ok(listed.path, `no path emitted for ${mb.id}`);
      assert.equal(resolveMailbox(DUPLICATE_NAME_TREE, listed.path).id, mb.id);
    }
  });
});

// ---------- filterMailboxesByParent ----------

describe('filterMailboxesByParent', () => {
  const TREE = [
    { id: 'mb-archive', name: 'Archive', role: 'archive' },
    { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
    { id: 'mb-2025', name: '2025', parentId: 'mb-archive' },
    { id: 'mb-q1', name: 'Q1', parentId: 'mb-2026' },
  ];

  it('returns the whole list when no parent is given', () => {
    assert.equal(filterMailboxesByParent(TREE).length, 4);
    assert.equal(filterMailboxesByParent(TREE, '   ').length, 4);
  });

  it('narrows to DIRECT children only, by any accepted parent form', () => {
    for (const form of ['mb-archive', 'archive', 'Archive']) {
      assert.deepEqual(filterMailboxesByParent(TREE, form).map(mb => mb.id), ['mb-2026', 'mb-2025'], form);
    }
    assert.deepEqual(filterMailboxesByParent(TREE, 'Archive/2026').map(mb => mb.id), ['mb-q1']);
  });

  it('rejects an unknown parent rather than listing everything', () => {
    assert.throws(() => filterMailboxesByParent(TREE, 'nope'), InvalidInputError);
  });
});

// ---------- createMailbox ----------

describe('createMailbox', () => {
  let client: JmapClient;

  const TREE = [
    ARCHIVE_MAILBOX,
    { id: 'mb-2026', name: '2026', parentId: 'mb-archive' },
  ];

  beforeEach(() => {
    client = makeClient();
  });

  function stubCreate(setResult: any = { created: { newMailbox: { id: 'mb-new' } } }) {
    stubMailboxes(client, TREE);
    return stubRequests(client, async () =>
      ({ methodResponses: [['Mailbox/set', setResult, 'createMailbox']] }));
  }

  function createArgs(makeReq: ReturnType<typeof stubRequests>) {
    return callArguments(makeReq)[0].methodCalls[0][1].create.newMailbox;
  }

  it('creates at the top level when no parent is given, and returns the mailbox with its path', async () => {
    const makeReq = stubCreate();
    const result = await client.createMailbox({ name: 'Receipts' });
    assert.deepEqual(createArgs(makeReq), { name: 'Receipts', parentId: null });
    assert.equal(result.mailbox.id, 'mb-new');
    assert.equal(result.mailbox.name, 'Receipts');
    assert.equal(result.path, 'Receipts');
  });

  it('resolves the parent by id, role, name or path', async () => {
    for (const [form, expectedParent, expectedPath] of [
      ['mb-archive', 'mb-archive', 'Archive/Receipts'],
      ['archive', 'mb-archive', 'Archive/Receipts'],
      ['Archive', 'mb-archive', 'Archive/Receipts'],
      ['Archive/2026', 'mb-2026', 'Archive/2026/Receipts'],
    ] as const) {
      client = makeClient();
      const makeReq = stubCreate();
      const result = await client.createMailbox({ name: 'Receipts', parent: form });
      assert.equal(createArgs(makeReq).parentId, expectedParent, form);
      assert.equal(result.path, expectedPath, form);
    }
  });

  it('rejects a name containing the path separator and points at the parent parameter', async () => {
    stubCreate();
    await assert.rejects(
      () => client.createMailbox({ name: 'Archive/Receipts' }),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /must not contain "\/"/);
        assert.match(err.message, /parent parameter/);
        return true;
      },
    );
  });

  it('rejects an empty name', async () => {
    stubCreate();
    await assert.rejects(() => client.createMailbox({ name: '   ' }), InvalidInputError);
  });

  it('rejects an unknown parent', async () => {
    stubCreate();
    await assert.rejects(() => client.createMailbox({ name: 'Receipts', parent: 'nope' }), InvalidInputError);
  });

  it('routes a caller-fixable set error to InvalidInputError, carrying the server reason', async () => {
    stubCreate({ notCreated: { newMailbox: { type: 'invalidProperties', description: 'name already in use' } } });
    await assert.rejects(
      () => client.createMailbox({ name: 'Receipts' }),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /Failed to create mailbox: invalidProperties/);
        assert.match(err.message, /name already in use/);
        return true;
      },
    );
  });

  it('routes an operational set error to a plain Error', async () => {
    stubCreate({ notCreated: { newMailbox: { type: 'forbidden' } } });
    await assert.rejects(
      () => client.createMailbox({ name: 'Receipts' }),
      (err: Error) => {
        assert.ok(!(err instanceof InvalidInputError));
        assert.match(err.message, /Failed to create mailbox: forbidden/);
        return true;
      },
    );
  });

  it('fails loudly when the server reports success but returns no id', async () => {
    stubCreate({ created: { newMailbox: {} } });
    await assert.rejects(() => client.createMailbox({ name: 'Receipts' }), /returned no mailbox id/);
  });

  it('returns the server created object untouched, alongside the merged mailbox', async () => {
    // `created` is what the raw path emits, so it must stay exactly what Mailbox/set sent
    // back (RFC 8620 §5.3: only the properties the server set) with nothing merged in.
    stubCreate({ created: { newMailbox: { id: 'mb-new', sortOrder: 7 } } });
    const result = await client.createMailbox({ name: 'Receipts' });
    assert.deepEqual(result.created, { id: 'mb-new', sortOrder: 7 });
    assert.equal(result.mailbox.name, 'Receipts');
    assert.equal(result.mailbox.parentId, null);
  });

  it('omits path rather than fabricating one when the parent has no computable path', async () => {
    // The parent sits in a loop, so no root-anchored path exists for it. Concatenating the
    // leaf onto its bare name would produce a plausible string that does not resolve.
    client = makeClient();
    stubMailboxes(client, [
      { id: 'mb-loop-a', name: 'Loop', parentId: 'mb-loop-b' },
      { id: 'mb-loop-b', name: 'Other', parentId: 'mb-loop-a' },
    ]);
    stubRequests(client, async () =>
      ({ methodResponses: [['Mailbox/set', { created: { newMailbox: { id: 'mb-new' } } }, 'createMailbox']] }));
    const result = await client.createMailbox({ name: 'Receipts', parent: 'mb-loop-a' });
    assert.equal(result.path, undefined);
    assert.equal(result.mailbox.parentId, 'mb-loop-a');
  });
});

// ---------- buildExclusionNote ----------

describe('buildExclusionNote', () => {
  it('returns empty string when there is no exclusion metadata', () => {
    assert.equal(buildExclusionNote(undefined), '');
  });

  it('no note when hidden === 0 (silence is the trustworthy signal)', () => {
    assert.equal(buildExclusionNote({ hidden: 0, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] }), '');
  });

  it('hidden > 0 produces a count note naming the folders + flags', () => {
    const note = buildExclusionNote({ hidden: 3, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] });
    assert.match(note, /3 message\(s\) in Trash\/Spam were excluded/);
    assert.match(note, /includeTrash:true/);
    assert.match(note, /includeSpam:true/);
  });

  it('hidden === null produces a front-loaded degraded note', () => {
    const note = buildExclusionNote({ hidden: null, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] });
    assert.match(note, /^\n\nRe-run/);
    assert.match(note, /couldn't be confirmed/);
  });

  it('an unresolved role produces a front-loaded fail-loud note', () => {
    const note = buildExclusionNote({ hidden: 0, excludedRoles: [], unresolvedRoles: ['Spam'] });
    assert.match(note, /Re-run to be sure/);
    assert.match(note, /NOT excluded/);
  });
});

// ---------- draft provenance + Message-ID lookup (#60) ----------

describe('readSourceReferences', () => {
  it('reads the reply and forward provenance headers off a raw Email', () => {
    assert.deepEqual(
      readSourceReferences({
        id: 'd1',
        inReplyTo: ['orig@example.com'],
        'header:X-Forwarded-Message-Id:asMessageIds': ['fwd@example.com'],
      }),
      { inReplyTo: ['orig@example.com'], forwardedMessageId: ['fwd@example.com'] },
    );
  });

  it('normalises missing headers to empty arrays', () => {
    assert.deepEqual(readSourceReferences({ id: 'd1' }), { inReplyTo: [], forwardedMessageId: [] });
    assert.deepEqual(readSourceReferences(undefined), { inReplyTo: [], forwardedMessageId: [] });
  });
});

describe('findEmailIdsByMessageId', () => {
  function stubLookup(client: JmapClient, list: any[]) {
    return stubRequests(client, async () => ({
      methodResponses: [
        ['Email/query', { ids: list.map(e => e.id) }, 'query'],
        ['Email/get', { list }, 'emails'],
      ],
    }));
  }

  it('queries the BARE Message-ID and keeps only messages that own it', async () => {
    const client = makeClient();
    const makeReq = stubLookup(client, [
      { id: 'orig-1', messageId: ['orig@example.com'] },   // the message itself
      { id: 'reply-1', messageId: ['reply@example.com'] }, // merely references or quotes it
    ]);

    assert.deepEqual(await client.findEmailIdsByMessageId('orig@example.com'), ['orig-1']);
    const query = (callArguments(makeReq)[0] as any).methodCalls[0][1];
    assert.deepEqual(query.filter, { text: 'orig@example.com' });
    assert.equal(query.filter.inMailboxOtherThan, undefined); // Trash/Spam are NOT excluded
  });

  it('sorts oldest-first so the message that owns the id can never fall off the capped page', async () => {
    const client = makeClient();
    const makeReq = stubLookup(client, [{ id: 'orig-1', messageId: ['orig@example.com'] }]);
    await client.findEmailIdsByMessageId('orig@example.com');
    const query = (callArguments(makeReq)[0] as any).methodCalls[0][1];
    // Every message that references or quotes an id postdates the message that owns it,
    // so oldest-first puts the owner on the first page regardless of the limit.
    assert.deepEqual(query.sort, [{ property: 'receivedAt', isAscending: true }]);
    assert.equal(typeof query.limit, 'number');
  });

  it('strips angle brackets before querying (the bracketed form matches nothing)', async () => {
    const client = makeClient();
    const makeReq = stubLookup(client, [{ id: 'orig-1', messageId: ['orig@example.com'] }]);
    assert.deepEqual(await client.findEmailIdsByMessageId('<orig@example.com>'), ['orig-1']);
    const query = (callArguments(makeReq)[0] as any).methodCalls[0][1];
    assert.deepEqual(query.filter, { text: 'orig@example.com' });
  });

  it('returns every id when more than one stored message carries it (the caller decides)', async () => {
    const client = makeClient();
    stubLookup(client, [
      { id: 'copy-1', messageId: ['dupe@example.com'] },
      { id: 'copy-2', messageId: ['dupe@example.com'] },
    ]);
    assert.deepEqual(await client.findEmailIdsByMessageId('dupe@example.com'), ['copy-1', 'copy-2']);
  });

  it('returns [] when nothing owns the id', async () => {
    const client = makeClient();
    stubLookup(client, [{ id: 'other-1', messageId: ['other@example.com'] }]);
    assert.deepEqual(await client.findEmailIdsByMessageId('orig@example.com'), []);
  });

  it('makes no request for a blank Message-ID', async () => {
    const client = makeClient();
    const makeReq = stubLookup(client, []);
    assert.deepEqual(await client.findEmailIdsByMessageId('  <>  '), []);
    assert.equal(makeReq.mock.calls.length, 0);
  });
});

// ---------- total result count (calculateTotal / QueryResult) ----------

// These pin the QueryResult contract itself: the query always asks the server to count
// the full match set, the reported total rides alongside the fetched page, and an absent
// server total leaves the key OFF rather than inventing a 0 (which would read as "no
// matches"). An explicit mailbox is used throughout so the default Trash/Spam exclusion
// and its extra count query stay out of the picture — those have their own coverage.
describe('query total result count', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  it('getEmails requests calculateTotal on Email/query', async () => {
    const makeReq = stubRequests(client, async () => queryResponse({}));

    await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });

    assert.equal(callArguments(makeReq)[0].methodCalls[0][1].calculateTotal, true);
  });

  it('surfaces the server-reported total alongside the items', async () => {
    stubMakeRequest(client, queryResponse({
      ids: ['e1'],
      list: [{ id: 'e1', subject: 'Only one fetched' }],
      total: 42,
    }));

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 1 });
    assert.equal(result.total, 42);
    assert.equal(result.items.length, 1);
  });

  it('omits total when the server does not report one', async () => {
    stubMakeRequest(client, queryResponse({ ids: ['e1'], list: [{ id: 'e1' }] }));

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 1 });
    assert.equal('total' in result, false);
    assert.equal(result.items.length, 1);
  });
});

// ---------- attachment upload: token-bearing request discipline ----------

describe('uploadBlob request discipline', () => {
  // The upload host and the advertised ceiling both come from the session, so the stub
  // carries a Fastmail-shaped uploadUrl (the post-substitution allowlist check rejects
  // anything else) and a small maxSizeUpload.
  const SESSION_WITH_UPLOAD = {
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: { 'urn:ietf:params:jmap:core': { maxSizeUpload: 1024 } },
    uploadUrl: 'https://api.fastmail.com/jmap/upload/{accountId}/',
    downloadUrl: 'https://www.fastmailusercontent.com/jmap/download/{accountId}/{blobId}/{name}?type={type}',
  };

  function makeUploadClient(): JmapClient {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    const client = new JmapClient(auth);
    mock.method(client, 'getSession', async () => SESSION_WITH_UPLOAD);
    return client;
  }

  it('POSTs to the substituted uploadUrl with Content-Type and refuses to follow redirects', async () => {
    const client = makeUploadClient();
    const fetchMock = mock.method(globalThis, 'fetch', async () => new Response(
      JSON.stringify({ blobId: 'B1', type: 'text/plain', size: 5 }), { status: 200 },
    ));
    try {
      const out = await client.uploadBlob(Buffer.from('hello'), 'text/plain');
      assert.equal(out.blobId, 'B1');
      const [url, opts] = callArguments(fetchMock) as [string, any];
      assert.equal(url, `https://api.fastmail.com/jmap/upload/${ACCOUNT_ID}/`);
      assert.equal(opts.method, 'POST');
      assert.equal(opts.headers['Content-Type'], 'text/plain');
      // The request carries the bearer token, so a 3xx must fail rather than replay it
      // to whatever host the redirect names.
      assert.equal(opts.redirect, 'error');
    } finally {
      fetchMock.mock.restore();
    }
  });

  it('rejects payloads above the advertised maxSizeUpload before any network call', async () => {
    const client = makeUploadClient();
    const fetchMock = mock.method(globalThis, 'fetch', async () => { throw new Error('should not fetch'); });
    try {
      await assert.rejects(
        () => client.uploadBlob(Buffer.alloc(2048), 'application/octet-stream'),
        /upload limit/,
      );
      assert.equal(fetchMock.mock.calls.length, 0);
    } finally {
      fetchMock.mock.restore();
    }
  });
});

// ---------- session/API calls: token-bearing request discipline ----------

// getSession and makeRequest both send the bearer token to a URL that has already been
// checked against the host allowlist (the base URL at construction, session.apiUrl inside
// getSession), so a redirect on either can only point off-allowlist. Both therefore pass
// redirect: 'error'. A grep proves the string is in the file; only reading the options
// off the fetch call proves it reaches fetch, which is what these two cover.
describe('session and API request discipline', () => {
  const AUTH_HEADERS = {
    'Authorization': 'Bearer fake-token',
    'Content-Type': 'application/json',
  };

  it('getSession refuses to follow redirects on the token-bearing session GET', async () => {
    const client = new JmapClient(new FastmailAuth({ apiToken: 'fake-token' }));
    const fetchMock = mock.method(globalThis, 'fetch', async () => new Response(
      JSON.stringify({
        apiUrl: 'https://api.fastmail.com/jmap/api/',
        primaryAccounts: { 'urn:ietf:params:jmap:mail': ACCOUNT_ID },
        accounts: { [ACCOUNT_ID]: {} },
        capabilities: {},
      }),
      { status: 200 },
    ));
    try {
      const session = await client.getSession();
      assert.equal(session.accountId, ACCOUNT_ID);
      const [url, opts] = callArguments(fetchMock) as [string, any];
      assert.equal(url, 'https://api.fastmail.com/jmap/session');
      // Whole option surface, not just the redirect key: the token is in these headers,
      // and a 3xx must fail rather than replay it to whatever host the redirect names.
      assert.deepEqual(opts, {
        method: 'GET',
        headers: AUTH_HEADERS,
        redirect: 'error',
      });
    } finally {
      fetchMock.mock.restore();
    }
  });

  it('makeRequest refuses to follow redirects on the token-bearing API POST', async () => {
    const client = new JmapClient(new FastmailAuth({ apiToken: 'fake-token' }));
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.fastmail.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
    }));
    const fetchMock = mock.method(globalThis, 'fetch', async () => new Response(
      JSON.stringify({ methodResponses: [] }), { status: 200 },
    ));
    const request = { using: ['urn:ietf:params:jmap:core'], methodCalls: [] };
    try {
      const out = await client.makeRequest(request as any);
      assert.deepEqual(out.methodResponses, []);
      const [url, opts] = callArguments(fetchMock) as [string, any];
      assert.equal(url, 'https://api.fastmail.com/jmap/api/');
      assert.deepEqual(opts, {
        method: 'POST',
        headers: AUTH_HEADERS,
        body: JSON.stringify(request),
        redirect: 'error',
      });
    } finally {
      fetchMock.mock.restore();
    }
  });
});

// ---------- downloadAttachmentToFile: exclusive-create discipline ----------

// downloadAttachmentToFile writes with O_EXCL ('wx') so a symlink planted at the target
// during the download is never followed. Overwriting a plain file therefore takes the
// EEXIST -> re-validate -> unlink -> rewrite path, and the REWRITE has to keep 'wx':
// the unlink frees the name, so a default-flag rewrite would re-open exactly the symlink
// window the exclusive first write closed. Only the flag actually passed to the second
// write proves that, and it is invisible from the filesystem afterwards — so this loads a
// private second copy of the client module with 'fs/promises' redirected to a recording
// wrapper, and reads the flags off the recorded calls. The redirect is keyed on the
// importing module's URL, so nothing else in the process sees it.
const FS_RECORDER = 'data:text/javascript,' + encodeURIComponent(`
export * from 'node:fs/promises';
import { writeFile as realWriteFile } from 'node:fs/promises';
export const writeFileCalls = [];
export function writeFile(path, data, options) {
  writeFileCalls.push({ path, options });
  return realWriteFile(path, data, options);
}
`);

describe('downloadAttachmentToFile write flags', () => {
  it('uses an exclusive create on BOTH writes when replacing an existing file', async (t) => {
    // Module hooks that can redirect a builtin arrived in Node 22.15; below that the
    // flags are not observable and the case is reported as skipped rather than passing
    // on nothing.
    const nodeModule = await import('node:module') as any;
    if (typeof nodeModule.registerHooks !== 'function') {
      t.skip('module hooks (registerHooks) are not available on this Node version');
      return;
    }

    const freshUrl = new URL('./jmap-client.ts?fs-recorder=1', import.meta.url).href;
    const hooks = nodeModule.registerHooks({
      resolve(specifier: string, context: any, nextResolve: any) {
        if (
          (specifier === 'fs/promises' || specifier === 'node:fs/promises') &&
          String(context.parentURL).includes('fs-recorder=1')
        ) {
          return { url: FS_RECORDER, shortCircuit: true };
        }
        return nextResolve(specifier, context);
      },
    });

    const root = await mkdtemp(join(tmpdir(), 'fastmail-mcp-toctou-'));
    try {
      const fresh = await import(freshUrl) as any;
      const recorder = await import(FS_RECORDER) as any;
      recorder.writeFileCalls.length = 0;

      const client = new fresh.JmapClient(new FastmailAuth({ apiToken: 'fake-token' }));
      // The network half is not what is under test; the bytes just have to arrive.
      mock.method(client, 'fetchAttachmentBuffer', async () => ({
        buffer: Buffer.from('NEW BYTES'),
        url: 'https://www.fastmailusercontent.com/jmap/download/acct/blob/report.txt',
        name: 'report.txt',
        type: 'text/plain',
        blobId: 'blob-1',
      }));
      const guard = mock.method(fresh.JmapClient, 'safeWritePath');

      // A pre-existing plain file at the target is what forces the EEXIST branch.
      await fsWriteFile(join(root, 'report.txt'), 'OLD BYTES');
      const result = await client.downloadAttachmentToFile('e1', 'a1', 'report.txt', root);

      // Overwriting a plain file still works end to end.
      assert.equal(await readFile(result.savedPath, 'utf8'), 'NEW BYTES');

      const calls = recorder.writeFileCalls;
      assert.equal(calls.length, 2, 'expected the exclusive write plus the post-unlink rewrite');
      assert.deepEqual(calls[0].options, { flag: 'wx' }, 'first write must be an exclusive create');
      assert.deepEqual(calls[1].options, { flag: 'wx' }, 'rewrite after the unlink must ALSO be an exclusive create');

      // Three guard calls: before the fetch, immediately before the write, and again
      // after EEXIST — the last one is what refuses a symlink planted at the target.
      assert.equal(guard.mock.calls.length, 3);
    } finally {
      hooks.deregister();
      await rm(root, { recursive: true, force: true });
    }
  });
});
// ---------- attachment reads: the part listing and the download forms (#13) ----------

// A Fastmail-shaped host is required, not cosmetic: the built URL is re-validated
// against the allowlist after substitution, before it would carry the bearer token, so
// a made-up host would fail every case below for the wrong reason. The path and query
// are kept minimal so the assertions read as URL shape rather than as Fastmail trivia.
const DOWNLOAD_TEMPLATE =
  'https://www.fastmailusercontent.com/{accountId}/{blobId}?type={type}&name={name}';

// A client whose session advertises a download URL template, so the built URL can be
// asserted end to end.
function makeDownloadClient(): JmapClient {
  const client = makeClient();
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: ACCOUNT_ID,
    capabilities: {},
    downloadUrl: DOWNLOAD_TEMPLATE,
  }));
  return client;
}

// The message shape these tests read: a text body, an html body, an embedded image the
// server routed into the body lists only, and one genuinely attached file.
const TEXT_PART = { partId: '1', type: 'text/plain', size: 12 };
const HTML_PART = { partId: '2', type: 'text/html', size: 40 };
const EMBEDDED_IMAGE = {
  partId: '3', type: 'image/png', size: 2048, blobId: 'blob-logo', cid: 'logo@example.com',
};
const ATTACHED_FILE = {
  partId: '4', type: 'application/pdf', size: 900, blobId: 'blob-pdf', name: 'report.pdf',
};

function stubEmail(client: JmapClient, email: any) {
  return stubRequests(client, async () => ({
    methodResponses: [['Email/get', { list: email ? [email] : [] }, 'getEmail']],
  }));
}

function mixedShapeEmail() {
  return {
    id: 'e1',
    attachments: [ATTACHED_FILE],
    textBody: [TEXT_PART, EMBEDDED_IMAGE],
    htmlBody: [HTML_PART, EMBEDDED_IMAGE],
  };
}

describe('getEmailAttachments', () => {
  it('fetches both body lists alongside attachments', async () => {
    const client = makeClient();
    const makeReq = stubEmail(client, mixedShapeEmail());
    await client.getEmailAttachments('e1');
    const params = callArguments(makeReq)[0].methodCalls[0][1];
    assert.deepEqual(params.properties, ['attachments', 'textBody', 'htmlBody']);
    assert.deepEqual(params.bodyProperties, [...EMAIL_BODY_PROPERTIES]);
  });

  it('lists the body-routed image alongside the attached file', async () => {
    const client = makeClient();
    stubEmail(client, mixedShapeEmail());
    const { attachments } = await client.getEmailAttachments('e1');
    assert.deepEqual(attachments.map((a: any) => a.partId), ['4', '3']);
  });

  it('returns raw JMAP part objects, not a simplified shape', async () => {
    const client = makeClient();
    stubEmail(client, mixedShapeEmail());
    const { attachments } = await client.getEmailAttachments('e1');
    assert.equal(attachments[1], EMBEDDED_IMAGE);
    assert.equal(attachments[1].cid, 'logo@example.com');
    assert.equal('isInline' in attachments[1], false);
  });

  it('reports the untouched JMAP array separately, so the withheld set is countable', async () => {
    const client = makeClient();
    stubEmail(client, mixedShapeEmail());
    const { attachments, rawAttachments } = await client.getEmailAttachments('e1');
    assert.deepEqual(rawAttachments, [ATTACHED_FILE]);
    assert.equal(attachments.length - rawAttachments.length, 1);
  });

  it('withholds nothing when every part is already in the JMAP array', async () => {
    const client = makeClient();
    stubEmail(client, { id: 'e1', attachments: [ATTACHED_FILE], textBody: [TEXT_PART] });
    const { attachments, rawAttachments } = await client.getEmailAttachments('e1');
    assert.equal(attachments.length, rawAttachments.length);
  });

  it('returns empty lists for a message that is not found', async () => {
    const client = makeClient();
    stubEmail(client, null);
    const result = await client.getEmailAttachments('missing');
    assert.deepEqual(result, { attachments: [], rawAttachments: [], omittedFromRaw: 0 });
  });
});

describe('downloadAttachment — fetch shape', () => {
  it('fetches both body lists so an embedded image is downloadable at all', async () => {
    const client = makeDownloadClient();
    const makeReq = stubEmail(client, mixedShapeEmail());
    await client.downloadAttachment('e1', '3');
    const params = callArguments(makeReq)[0].methodCalls[0][1];
    assert.deepEqual(params.properties, ['attachments', 'textBody', 'htmlBody']);
  });

  it('requests the full body property set, including disposition and cid', async () => {
    const client = makeDownloadClient();
    const makeReq = stubEmail(client, mixedShapeEmail());
    await client.downloadAttachment('e1', '3');
    const params = callArguments(makeReq)[0].methodCalls[0][1];
    assert.deepEqual(params.bodyProperties, [...EMAIL_BODY_PROPERTIES]);
  });

  it('no longer requests bodyValues, which it never read', async () => {
    const client = makeDownloadClient();
    const makeReq = stubEmail(client, mixedShapeEmail());
    await client.downloadAttachment('e1', '3');
    const params = callArguments(makeReq)[0].methodCalls[0][1];
    assert.equal(params.properties.includes('bodyValues'), false);
  });
});

describe('downloadAttachment — reference forms', () => {
  it('resolves a partId', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    assert.match(await client.downloadAttachment('e1', '4'), /blob-pdf/);
  });

  it('resolves a blobId', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    assert.match(await client.downloadAttachment('e1', 'blob-logo'), /blob-logo/);
  });

  it('resolves a prefixed cid', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    assert.match(await client.downloadAttachment('e1', 'cid:logo@example.com'), /blob-logo/);
  });

  it('accepts the cid: prefix in any case', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    assert.match(await client.downloadAttachment('e1', 'CID:logo@example.com'), /blob-logo/);
  });

  it('resolves an entry number against the full listing, not the JMAP array', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    // Entry 0 is the attached file; entry 1 is the body-routed image, which the JMAP
    // attachments array does not contain at all — so indexing it proves the union is
    // what the entry-number form counts from.
    assert.match(await client.downloadAttachment('e1', '0'), /blob-pdf/);
    assert.match(await client.downloadAttachment('e1', '1'), /blob-logo/);
  });

  it('gives a partId precedence over the entry-number reading of the same digits', async () => {
    // Fastmail partIds are digit strings, so "1" must mean the part with partId "1"
    // whenever one exists — never the second entry.
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [ATTACHED_FILE, { partId: '1', type: 'image/png', size: 5, blobId: 'blob-one' }],
      textBody: [],
    });
    assert.match(await client.downloadAttachment('e1', '1'), /blob-one/);
  });

  it('falls back to the entry number only when no part claims those digits', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'a', type: 'image/png', size: 5, blobId: 'blob-a' },
                    { partId: 'b', type: 'image/png', size: 5, blobId: 'blob-b' }],
    });
    assert.match(await client.downloadAttachment('e1', '1'), /blob-b/);
  });

  it('never lets a bare digit string alias a cid', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p', type: 'image/png', size: 5, blobId: 'blob-x', cid: '7' }],
    });
    // "7" is a Content-ID here, but the cid form requires the prefix, so a bare 7 is an
    // out-of-range entry number, not this part.
    await assert.rejects(() => client.downloadAttachment('e1', '7'), /Attachment not found/);
    assert.match(await client.downloadAttachment('e1', 'cid:7'), /blob-x/);
  });

  it('strips only the first cid: prefix, so a Content-ID beginning with cid: is reachable', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p', type: 'image/png', size: 5, blobId: 'blob-odd', cid: 'cid:x' }],
    });
    assert.match(await client.downloadAttachment('e1', 'cid:cid:x'), /blob-odd/);
  });

  it('matches a cid literally before trying the decoded form', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [
        { partId: 'p1', type: 'image/png', size: 5, blobId: 'blob-literal', cid: '%78' },
        { partId: 'p2', type: 'image/png', size: 5, blobId: 'blob-decoded', cid: 'x' },
      ],
    });
    // The handle round-trips from get_email's verbatim echo, so a literal hit wins and
    // the decode never runs.
    assert.match(await client.downloadAttachment('e1', 'cid:%78'), /blob-literal/);
  });

  it('falls back to the decoded cid only when the literal matches nothing', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p2', type: 'image/png', size: 5, blobId: 'blob-decoded', cid: 'x' }],
    });
    assert.match(await client.downloadAttachment('e1', 'cid:%78'), /blob-decoded/);
  });
});

describe('downloadAttachment — input errors stay actionable, lookups stay generic', () => {
  it('rejects an ambiguous cid instead of guessing at one part', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [
        { partId: 'p1', type: 'image/png', size: 5, blobId: 'b1', cid: 'dupe@host' },
        { partId: 'p2', type: 'image/png', size: 5, blobId: 'b2', cid: 'dupe@host' },
      ],
    });
    await assert.rejects(
      () => client.downloadAttachment('e1', 'cid:dupe@host'),
      (error: unknown) => {
        assert.ok(error instanceof InvalidInputError);
        assert.equal(
          (error as Error).message,
          'attachmentId "cid:dupe@host" matches more than one part. ' +
          "Pass the part's blobId or partId from get_email_attachments instead.",
        );
        return true;
      },
    );
  });

  it('names no blobIds in the ambiguity message', async () => {
    // The tool's no-metadata-leak posture: an input error may say what to pass, never
    // what the mailbox contains.
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [
        { partId: 'p1', type: 'image/png', size: 5, blobId: 'secret-blob-1', cid: 'd@h' },
        { partId: 'p2', type: 'image/png', size: 5, blobId: 'secret-blob-2', cid: 'd@h' },
      ],
    });
    await assert.rejects(
      () => client.downloadAttachment('e1', 'cid:d@h'),
      (error: unknown) => !/secret-blob/.test((error as Error).message),
    );
  });

  it('renders a hostile cid as bounded quoted data', async () => {
    const client = makeDownloadClient();
    const hostile = `${'a'.repeat(200)}‮"`;
    stubEmail(client, {
      id: 'e1',
      attachments: [
        { partId: 'p1', type: 'image/png', size: 5, blobId: 'b1', cid: hostile },
        { partId: 'p2', type: 'image/png', size: 5, blobId: 'b2', cid: hostile },
      ],
    });
    await assert.rejects(
      () => client.downloadAttachment('e1', `cid:${hostile}`),
      (error: unknown) => {
        const message = (error as Error).message;
        assert.ok(message.length < 300);
        assert.equal(message.includes('‮'), false);
        return true;
      },
    );
  });

  it('rejects a bare cid: with no value', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    await assert.rejects(
      () => client.downloadAttachment('e1', 'cid:'),
      (error: unknown) => error instanceof InvalidInputError && /names no Content-ID/.test((error as Error).message),
    );
  });

  it('rejects a number with junk in it instead of silently indexing', async () => {
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    for (const bad of ['3a', '-1', '1.5', ' 1']) {
      await assert.rejects(
        () => client.downloadAttachment('e1', bad),
        (error: unknown) => error instanceof InvalidInputError,
        `expected ${JSON.stringify(bad)} to be rejected as an input error`,
      );
    }
  });

  it('rejects a part that carries no downloadable blob', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p1', type: 'image/png', size: 5, cid: 'no-blob@host' }],
    });
    await assert.rejects(
      () => client.downloadAttachment('e1', 'p1'),
      (error: unknown) => error instanceof InvalidInputError && /no blobId/.test((error as Error).message),
    );
  });

  it('keeps a reference that simply matches nothing generic', async () => {
    // An existence failure must not become an actionable input error: that is how the
    // tool avoids confirming what a mailbox holds.
    const client = makeDownloadClient();
    stubEmail(client, mixedShapeEmail());
    for (const missing of ['nope', 'cid:absent@host', '99']) {
      await assert.rejects(
        () => client.downloadAttachment('e1', missing),
        (error: unknown) => !(error instanceof InvalidInputError) && /Attachment not found/.test((error as Error).message),
        `expected ${JSON.stringify(missing)} to stay a generic not-found`,
      );
    }
  });
});

describe('downloadAttachment — declared filename', () => {
  it('sanitizes the sender-supplied name without appending .eml', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p1', type: 'image/png', size: 5, blobId: 'b1', name: 'lo‮go.png' }],
    });
    const url = await client.downloadAttachment('e1', 'p1');
    assert.match(url, /name=logo\.png$/);
    assert.equal(url.includes('.eml'), false);
  });

  it('falls back to a usable name when the part has none', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p1', type: 'image/png', size: 5, blobId: 'b1' }],
    });
    assert.match(await client.downloadAttachment('e1', 'p1'), /name=attachment$/);
  });

  it('defuses a Windows device name', async () => {
    const client = makeDownloadClient();
    stubEmail(client, {
      id: 'e1',
      attachments: [{ partId: 'p1', type: 'image/png', size: 5, blobId: 'b1', name: 'CON.png' }],
    });
    assert.match(await client.downloadAttachment('e1', 'p1'), /name=CON_\.png$/);
  });
});

// ---------- download URL: the allowlist check that runs AFTER substitution ----------

// Both the download URL template and the values spliced into it come from the session,
// so a session can be allowlisted as a template and off-allowlist once substituted. Here
// the account id lands in the userinfo position and closes the authority early with a
// '/', which moves the effective host to one the allowlist rejects. Checking the template
// at session time sees nothing wrong with this; only re-checking the finished URL does.
const SPLICING_DOWNLOAD_TEMPLATE =
  'https://{accountId}@www.fastmailusercontent.com/{blobId}?type={type}&name={name}';
const SPLICING_ACCOUNT_ID = 'attacker.example.com/';

function makeSplicingDownloadClient(): JmapClient {
  const client = makeClient();
  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: SPLICING_ACCOUNT_ID,
    capabilities: {},
    downloadUrl: SPLICING_DOWNLOAD_TEMPLATE,
  }));
  return client;
}

describe('downloadAttachment — the URL is re-validated after substitution', () => {
  it('finds nothing wrong with the template on its own', () => {
    // The premise the other two cases rest on: a check that only ever saw the template
    // would pass this session through, because the template's host is allowlisted.
    const parsed = validateFastmailUrl(SPLICING_DOWNLOAD_TEMPLATE, 'downloadUrl');
    assert.equal(parsed.hostname, 'www.fastmailusercontent.com');
  });

  it('refuses a substituted URL whose real host is off the allowlist', async () => {
    const client = makeSplicingDownloadClient();
    stubEmail(client, mixedShapeEmail());
    await assert.rejects(
      () => client.downloadAttachment('e1', '4'),
      (error: unknown) => {
        assert.match(
          (error as Error).message,
          /downloadUrl host 'attacker\.example\.com' is not in the Fastmail allowlist/,
        );
        return true;
      },
    );
  });

  it('makes no request at all, so the bearer token never reaches that host', async () => {
    // The whole point of the check is where it lands: before the token-bearing fetch,
    // not on its response.
    const client = makeSplicingDownloadClient();
    stubEmail(client, mixedShapeEmail());
    const fetchMock = mock.method(globalThis, 'fetch', async () => new Response('bytes', { status: 200 }));
    try {
      await assert.rejects(
        () => client.fetchAttachmentBuffer('e1', '4'),
        /not in the Fastmail allowlist/,
      );
      assert.equal(fetchMock.mock.calls.length, 0);
    } finally {
      fetchMock.mock.restore();
    }
  });
});

// ---------- URL templates: values are inserted literally ----------

// Every sequence String.prototype.replace treats specially inside a *string* replacement:
// the whole match, the text before it, the text after it, a numbered group, and an
// escaped dollar. Substituted with a string replacement, each of these would splice part
// of the template back into the URL instead of the value; substituted literally, they
// come out exactly as written. Both the templates and the values arrive from the session,
// so this is what a hostile or compromised server would reach for.
const DOLLAR_SPLICE = "a$&b$`c$'d$1e$$f";

describe('fillUrlTemplate', () => {
  it('inserts a value containing replacement-pattern sequences verbatim', () => {
    assert.equal(
      fillUrlTemplate('https://host.example/{a}/end', { '{a}': DOLLAR_SPLICE }),
      `https://host.example/${DOLLAR_SPLICE}/end`,
    );
  });

  it('fills every named slot, leaving unnamed ones untouched', () => {
    assert.equal(
      fillUrlTemplate('https://host.example/{a}/{b}?x={c}', { '{a}': '1', '{b}': '2' }),
      'https://host.example/1/2?x={c}',
    );
  });

  it('does not re-substitute a placeholder that a value introduced', () => {
    // A value naming a slot must stay data: filling {a} with the literal text "{b}"
    // leaves that text in the URL, and does not consume {b}'s own value on the way.
    assert.equal(
      fillUrlTemplate('https://host.example/{a}/{b}', { '{a}': '{b}', '{b}': 'real' }),
      'https://host.example/{b}/real',
    );
  });
});

describe('download and upload URLs are built with literal substitution', () => {
  // An allowlisted template with all four slots, so one assertion covers the raw
  // substitutions (accountId, blobId) and the percent-encoded ones (type, name).
  const TEMPLATE =
    'https://www.fastmailusercontent.com/{accountId}/{blobId}?type={type}&name={name}';

  function clientWithTemplate(): JmapClient {
    const client = makeClient();
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: ACCOUNT_ID,
      capabilities: {},
      downloadUrl: TEMPLATE,
    }));
    return client;
  }

  it('builds the exact download URL for a blobId, type and name full of dollar sequences', async () => {
    const client = clientWithTemplate();
    stubEmail(client, {
      id: 'e1',
      attachments: [{
        partId: 'p1',
        type: `image/${DOLLAR_SPLICE}`,
        size: 5,
        blobId: DOLLAR_SPLICE,
        name: `${DOLLAR_SPLICE}.png`,
      }],
    });

    // Asserted as exact bytes, not as "it parsed" or "it validated": a spliced URL still
    // parses and still passes the host allowlist, so only the string shows the bug.
    // blobId is inserted raw; type and name are percent-encoded, which turns every '$'
    // into %24 and makes them inert regardless.
    assert.equal(
      await client.downloadAttachment('e1', 'p1'),
      "https://www.fastmailusercontent.com/acct-123/a$&b$`c$'d$1e$$f"
        + "?type=image%2Fa%24%26b%24%60c%24'd%241e%24%24f"
        + "&name=a%24%26b%24%60c%24'd%241e%24%24f.png",
    );
  });

  it('builds the exact upload URL for an accountId full of dollar sequences', async () => {
    const auth = new FastmailAuth({ apiToken: 'fake-token' });
    const client = new JmapClient(auth);
    mock.method(client, 'getSession', async () => ({
      apiUrl: 'https://api.example.com/jmap/api/',
      accountId: DOLLAR_SPLICE,
      capabilities: {},
      uploadUrl: 'https://api.fastmail.com/jmap/upload/{accountId}/',
    }));

    const fetchMock = mock.method(globalThis, 'fetch', async () => new Response(
      JSON.stringify({ blobId: 'B1', type: 'text/plain', size: 5 }), { status: 200 },
    ));
    try {
      await client.uploadBlob(Buffer.from('hello'), 'text/plain');
      const [url] = callArguments(fetchMock) as [string, any];
      assert.equal(url, "https://api.fastmail.com/jmap/upload/a$&b$`c$'d$1e$$f/");
    } finally {
      fetchMock.mock.restore();
    }
  });

  it('still re-validates the finished URL, which literal substitution does not replace', async () => {
    // Literal insertion stops the template from leaking into the URL; it does not stop a
    // value from damaging the URL. The origin check after substitution remains the
    // control that catches that.
    const client = makeSplicingDownloadClient();
    stubEmail(client, mixedShapeEmail());
    await assert.rejects(
      () => client.downloadAttachment('e1', '4'),
      /not in the Fastmail allowlist/,
    );
  });
});
