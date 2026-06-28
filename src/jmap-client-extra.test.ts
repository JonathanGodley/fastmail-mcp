import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { JmapClient, EMAIL_PROPERTIES_COMPACT, EMAIL_PROPERTIES_VERBOSE, EMAIL_BODY_PROPERTIES, buildMailboxNameMap, attachMailboxNames, resolveMailbox, computeExclusion } from './jmap-client.js';
import { InvalidInputError } from './coerce.js';
import { buildExclusionNote } from './response-formatters.js';
import { FastmailAuth } from './auth.js';

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

function stubMakeRequest(client: JmapClient, response: any) {
  mock.method(client, 'makeRequest', async () => response);
}

function stubMailboxes(client: JmapClient, mailboxes: any[] = DEFAULT_MAILBOXES) {
  mock.method(client, 'getMailboxes', async () => mailboxes);
}

// A request-aware Email/query stub: the visible query reads ids/total from `query`, the
// get reads from `emails`, and (when present) the count query reads from `count`. Because
// getMailboxes is stubbed separately in these tests, makeRequest only ever sees the
// Email/query batch, so this returns the same fixed shape regardless of the methodCalls.
function queryResponse(opts: { ids?: string[]; list?: any[]; total?: number; broaderTotal?: number }) {
  const responses: any[] = [
    ['Email/query', { ids: opts.ids ?? [], total: opts.total }, 'query'],
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
});

// ---------- getEmails ----------

describe('getEmails', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('returns emails scoped to an explicit mailbox (no exclusion, no count query)', async () => {
    stubMailboxes(client);
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1', subject: 'Filtered' }], total: 1 }),
    );

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });
    assert.equal(result.items.length, 1);

    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
    const filter = batch[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.inMailboxOtherThan, undefined);
    // Explicit mailbox => no exclusion => no count query and no exclusion metadata.
    assert.equal(batch.length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('default (no mailbox) excludes Trash + Spam via inMailboxOtherThan and runs a count query', async () => {
    stubMailboxes(client);
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1', subject: 'All' }], total: 8, broaderTotal: 11 }),
    );

    const result = await client.getEmails({ limit: 10 });
    assert.equal(result.items.length, 1);

    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
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
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 1 }),
    );

    const result = await client.getEmails({ includeTrash: true, includeSpam: true });
    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
    assert.equal(batch[0][1].filter.inMailboxOtherThan, undefined);
    assert.equal(batch.length, 2);
    assert.equal(result.exclusion, undefined);
  });

  it('excludeDrafts adds notKeyword $draft, AND-wrapped with the default exclusion', async () => {
    stubMailboxes(client);
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }),
    );

    await client.getEmails({ excludeDrafts: true });
    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
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
    // resolves and makeRequest only sees the Email/get + Email/set pair.
    stubMailboxes(client);
  });

  it('moves email successfully', async () => {
    // First call: getEmail to read current mailboxIds
    // Second call: Email/set to move
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ id: 'e1', mailboxIds: { 'mb-inbox': true } }] }, 'getEmail'],
          ],
        };
      }
      return {
        methodResponses: [
          ['Email/set', { updated: { 'e1': null } }, 'moveEmail'],
        ],
      };
    });

    await client.moveEmail('e1', 'mb-archive');
    assert.equal(callCount, 2);
  });

  it('throws when update fails', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
      callCount++;
      if (callCount === 1) {
        return {
          methodResponses: [
            ['Email/get', { list: [{ id: 'e1', mailboxIds: { 'mb-inbox': true } }] }, 'getEmail'],
          ],
        };
      }
      return {
        methodResponses: [
          ['Email/set', { notUpdated: { 'e1': { type: 'notFound' } } }, 'moveEmail'],
        ],
      };
    });

    await assert.rejects(
      () => client.moveEmail('e1', 'mb-archive'),
      (err: Error) => {
        assert.match(err.message, /Failed to move/);
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
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
  });

  it('marks email as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null } }, 'updateEmail'],
      ],
    }));

    await client.markEmailRead('e1', false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
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

// ---------- bulkMarkRead ----------

describe('bulkMarkRead', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('marks multiple emails as read in one request', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null, 'e3': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2', 'e3'], true);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': true });
    assert.deepEqual(update['e2'], { 'keywords/$seen': true });
    assert.deepEqual(update['e3'], { 'keywords/$seen': true });
  });

  it('marks multiple emails as unread', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/set', { updated: { 'e1': null, 'e2': null } }, 'bulkUpdate'],
      ],
    }));

    await client.bulkMarkRead(['e1', 'e2'], false);

    const update = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].update;
    assert.deepEqual(update['e1'], { 'keywords/$seen': null });
    assert.deepEqual(update['e2'], { 'keywords/$seen': null });
  });

  it('throws when some emails fail to update', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Email/set', { notUpdated: { 'e2': { type: 'notFound' } } }, 'bulkUpdate'],
      ],
    });

    await assert.rejects(
      () => client.bulkMarkRead(['e1', 'e2']),
      (err: Error) => {
        assert.match(err.message, /Failed to update/);
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
    const makeReq = mock.method(client, 'makeRequest', async () => {
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

    const emailGetArgs = makeReq.mock.calls[1].arguments[0].methodCalls[1][1];
    assert.ok(emailGetArgs.properties.includes('preview'), 'should request preview');
    assert.ok(emailGetArgs.properties.includes('inReplyTo'), 'should request inReplyTo');
    assert.ok(!emailGetArgs.properties.includes('bodyValues'), 'should NOT request bodyValues');
    assert.ok(!emailGetArgs.properties.includes('textBody'), 'should NOT request textBody');
    assert.ok(!emailGetArgs.properties.includes('htmlBody'), 'should NOT request htmlBody');
    assert.ok(!emailGetArgs.properties.includes('attachments'), 'should NOT request attachments');
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
    assert.equal(emailGetArgs.fetchHTMLBodyValues, undefined);
    assert.equal(emailGetArgs.bodyProperties, undefined);
  });

  it('always requests compact properties (no bodies)', async () => {
    let callCount = 0;
    const makeReq = mock.method(client, 'makeRequest', async () => {
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

    const emailGetArgs = makeReq.mock.calls[1].arguments[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'), 'should NOT request bodyValues');
    assert.ok(!emailGetArgs.properties.includes('textBody'), 'should NOT request textBody');
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  // A thread containing a normal email plus an in-progress draft reply.
  const threadWithDraftResponse = () => {
    let callCount = 0;
    return mock.method(client, 'makeRequest', async () => {
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

  it('reports hiddenDraftCount 0 for a thread with no drafts', async () => {
    let callCount = 0;
    mock.method(client, 'makeRequest', async () => {
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
    const makeReq = mock.method(client, 'makeRequest', async () => standardQueryResponse);
    return { makeReq, result: callFn() };
  }

  it('getEmails always requests compact properties (no bodies)', async () => {
    const { makeReq, result } = mockAndCall('getEmails', () => client.getEmails());
    await result;
    const emailGetArgs = makeReq.mock.calls[0].arguments[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'), 'should NOT request bodyValues');
    assert.ok(!emailGetArgs.properties.includes('textBody'), 'should NOT request textBody');
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('searchEmails always requests compact properties (no bodies)', async () => {
    const { makeReq, result } = mockAndCall('searchEmails', () => client.searchEmails({ query: 'test' }));
    await result;
    const emailGetArgs = makeReq.mock.calls[0].arguments[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('getRecentEmails always requests compact properties (no bodies)', async () => {
    mock.method(client, 'getMailboxes', async () => [INBOX_MAILBOX]);
    const makeReq = mock.method(client, 'makeRequest', async () => standardQueryResponse);
    await client.getRecentEmails();
    const emailGetArgs = makeReq.mock.calls[0].arguments[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });

  it('searchEmails with structured filters still requests compact properties', async () => {
    const { makeReq, result } = mockAndCall('searchEmails', () => client.searchEmails({ query: 'test', from: 'a@b.com' }));
    await result;
    const emailGetArgs = makeReq.mock.calls[0].arguments[0].methodCalls[1][1];
    assert.ok(!emailGetArgs.properties.includes('bodyValues'));
    assert.equal(emailGetArgs.fetchTextBodyValues, undefined);
  });
});

// ---------- JMAP property consistency ----------

describe('JMAP property consistency', () => {
  it('verbose properties are a superset of compact properties', () => {
    for (const prop of EMAIL_PROPERTIES_COMPACT) {
      assert.ok(
        EMAIL_PROPERTIES_VERBOSE.includes(prop),
        `verbose properties missing compact property: ${prop}`
      );
    }
  });

  it('verbose includes body-specific properties that compact does not', () => {
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('textBody'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('htmlBody'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('bodyValues'));
    assert.ok(EMAIL_PROPERTIES_VERBOSE.includes('attachments'));
    assert.ok(!EMAIL_PROPERTIES_COMPACT.includes('textBody'));
    assert.ok(!EMAIL_PROPERTIES_COMPACT.includes('htmlBody'));
    assert.ok(!EMAIL_PROPERTIES_COMPACT.includes('bodyValues'));
    assert.ok(!EMAIL_PROPERTIES_COMPACT.includes('attachments'));
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
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getEmails({ mailbox: 'mb-inbox', limit: 5, ascending: true });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('searchEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.searchEmails({ query: 'test', limit: 10 });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.searchEmails({ query: 'test', limit: 10, ascending: true });

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: true }]);
    });
  });

  describe('getRecentEmails', () => {
    it('defaults to isAscending: false', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox');

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
      assert.deepEqual(sort, [{ property: 'receivedAt', isAscending: false }]);
    });

    it('passes ascending=true as isAscending: true', async () => {
      stubMailboxes(client);
      const makeReq = mock.method(client, 'makeRequest', async () => QUERY_GET_RESPONSE);

      await client.getRecentEmails(10, 'inbox', true);

      const sort = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].sort;
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

  it('getEmails attaches resolved names from getMailboxes (multi-membership; unresolved omitted)', async () => {
    mock.method(client, 'getMailboxes', async () => NAME_MAILBOXES);
    const makeReq = mock.method(client, 'makeRequest', async () => NAME_QUERY_RESPONSE);

    const result = await client.getEmails({ mailbox: 'mb-inbox', limit: 5 });

    // No in-batch Mailbox/get — explicit mailbox => just query + get.
    const calls = makeReq.mock.calls[0].arguments[0].methodCalls;
    assert.equal(calls.length, 2);
    assert.equal(calls.some((c: any) => c[0] === 'Mailbox/get'), false);

    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox', 'Receipts']);
    // e2's only mailbox id didn't resolve → no field at all (don't fabricate).
    assert.equal('_mailboxNames' in (result.items[1] as any), false);
  });

  it('attaches nothing (does NOT throw) when an email\'s mailbox id is not in the fetched list', async () => {
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
  });

  it('searchEmails attaches names from getMailboxes', async () => {
    mock.method(client, 'getMailboxes', async () => NAME_MAILBOXES);
    const makeReq = mock.method(client, 'makeRequest', async () => NAME_QUERY_RESPONSE);
    // includeTrash/includeSpam true => no exclusion/count query, just query + get.
    const result = await client.searchEmails({ query: 'x', limit: 5, includeTrash: true, includeSpam: true });
    assert.equal(makeReq.mock.calls[0].arguments[0].methodCalls.length, 2);
    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox', 'Receipts']);
  });

  it('getEmailById attaches names from its appended Mailbox/get (index 1)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/get', { list: [{ id: 'e1', subject: 'A', mailboxIds: { 'mb-trash': true } }] }, 'email'],
        ['Mailbox/get', { list: [{ id: 'mb-trash', name: 'Trash' }] }, 'mailboxes'],
      ],
    }));

    const email = await client.getEmailById('e1');
    assert.equal(makeReq.mock.calls[0].arguments[0].methodCalls[1][0], 'Mailbox/get');
    assert.deepEqual((email as any)._mailboxNames, ['Trash']);
  });

  it('getRecentEmails reuses its existing getMailboxes list (no third methodCall) and attaches names', async () => {
    stubMailboxes(client); // INBOX/DRAFTS/TRASH/SENT
    const makeReq = mock.method(client, 'makeRequest', async () => ({
      methodResponses: [
        ['Email/query', { ids: ['e1'] }, 'query'],
        ['Email/get', { list: [{ id: 'e1', subject: 'A', mailboxIds: { 'mb-inbox': true } }] }, 'emails'],
      ],
    }));

    const result = await client.getRecentEmails(10, 'inbox');
    // Reuses getMailboxes — the request stays a 2-call batch, no appended Mailbox/get.
    assert.equal(makeReq.mock.calls[0].arguments[0].methodCalls.length, 2);
    assert.deepEqual((result.items[0] as any)._mailboxNames, ['Inbox']);
  });

  it('getThread attaches names to retained messages before the draft filter runs', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['Thread/get', { list: [{ id: 't1', emailIds: ['e1', 'e2'] }] }, 'getThread'],
        ['Email/get', { list: [
          { id: 'e1', subject: 'Kept', mailboxIds: { 'mb-inbox': true } },
          { id: 'e2', subject: 'Draft', keywords: { $draft: true }, mailboxIds: { 'mb-drafts': true } },
        ] }, 'emails'],
        ['Mailbox/get', { list: [
          { id: 'mb-inbox', name: 'Inbox' },
          { id: 'mb-drafts', name: 'Drafts' },
        ] }, 'mailboxes'],
      ],
    });

    const { emails } = await client.getThread('t1'); // drafts excluded by default
    assert.equal(emails.length, 1);
    assert.equal(emails[0].id, 'e1');
    assert.deepEqual((emails[0] as any)._mailboxNames, ['Inbox']);
  });

  describe('buildMailboxNameMap', () => {
    it('maps id -> name, keying on the real name (custom labels included)', () => {
      const map = buildMailboxNameMap([
        { id: 'a', name: 'Inbox', role: 'inbox' },
        { id: 'b', name: 'My Label', role: null }, // role null but still mapped
      ]);
      assert.equal(map.get('a'), 'Inbox');
      assert.equal(map.get('b'), 'My Label');
    });

    it('returns an empty map for [] (the degradation input)', () => {
      assert.equal(buildMailboxNameMap([]).size, 0);
    });

    it('skips entries lacking an id or a string name', () => {
      const map = buildMailboxNameMap([{ id: 'a' }, { name: 'x' }, null as any]);
      assert.equal(map.size, 0);
    });
  });

  describe('attachMailboxNames', () => {
    const map = new Map([['a', 'Inbox'], ['b', 'Receipts']]);

    it('attaches a non-enumerable _mailboxNames (absent from JSON, readable directly)', () => {
      const email: any = { id: 'e', mailboxIds: { a: true, b: true } };
      attachMailboxNames([email], map);
      assert.deepEqual(email._mailboxNames, ['Inbox', 'Receipts']);
      assert.equal(JSON.stringify(email).includes('_mailboxNames'), false);
    });

    it('omits the field when no id resolves', () => {
      const email: any = { id: 'e', mailboxIds: { z: true } };
      attachMailboxNames([email], map);
      assert.equal('_mailboxNames' in email, false);
    });

    it('omits the field when mailboxIds is absent or empty', () => {
      const noIds: any = { id: 'e' };
      const emptyIds: any = { id: 'e', mailboxIds: {} };
      attachMailboxNames([noIds, emptyIds], map);
      assert.equal('_mailboxNames' in noIds, false);
      assert.equal('_mailboxNames' in emptyIds, false);
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

  it('an explicit mailbox disables exclusion entirely', () => {
    const r = computeExclusion(withRoles, { hasExplicitMailbox: true });
    assert.deepEqual(r.excludeIds, []);
    assert.deepEqual(r.excludedRoles, []);
    assert.deepEqual(r.unresolvedRoles, []);
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
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 1 }));
    const result = await client.searchEmails({ query: 'x', mailbox: 'inbox' });
    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
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
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ mailbox: 'inbox', includeSpam: true });
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.inMailbox, 'mb-inbox');
    assert.equal(filter.inMailboxOtherThan, undefined);
  });

  it('isUnread:false => hasKeyword $seen; isPinned:false => notKeyword $flagged (polarity kept)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ isUnread: false, isPinned: false, includeTrash: true, includeSpam: true });
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$seen'));
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$flagged'));
  });

  it('isUnread:true + isPinned:true => two separate keyword conditions', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    await client.searchEmails({ isUnread: true, isPinned: true, includeTrash: true, includeSpam: true });
    const filter = makeReq.mock.calls[0].arguments[0].methodCalls[0][1].filter;
    assert.equal(filter.operator, 'AND');
    assert.ok(filter.conditions.some((c: any) => c.notKeyword === '$seen'));
    assert.ok(filter.conditions.some((c: any) => c.hasKeyword === '$flagged'));
  });

  it('hidden = broaderTotal - visibleTotal', async () => {
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8, broaderTotal: 11 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 3);
    assert.deepEqual(result.exclusion?.excludedRoles, ['Trash', 'Spam']);
  });

  it('fires hidden>0 even when the visible result set is empty', async () => {
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 4 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 4);
  });

  it('hidden === 0 when nothing was withheld', async () => {
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 5, broaderTotal: 5 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, 0);
  });

  it('FAIL-CLOSED: an absent count total => hidden:null (degraded), never silence', async () => {
    // Exclusion is active (so a count methodCall is in the request), but the mocked
    // response omits the count entry -> the count read throws -> hidden:null.
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, null);
  });

  it('FAIL-CLOSED: a negative hidden (bad broader total) => hidden:null', async () => {
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: ['e1'], list: [{ id: 'e1' }], total: 8, broaderTotal: 0 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.equal(result.exclusion?.hidden, null);
  });

  it('a missing junk/trash role surfaces in unresolvedRoles (no silent inclusion)', async () => {
    stubMailboxes(client, [{ id: 'mb-inbox', name: 'Inbox', role: 'inbox' }]);
    mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0 }));
    const result = await client.searchEmails({ query: 'x' });
    assert.deepEqual(result.exclusion?.unresolvedRoles, ['Trash', 'Spam']);
    assert.deepEqual(result.exclusion?.excludedRoles, []);
  });

  it('count filter keeps the keyword conds but drops inMailboxOtherThan (differs from visible)', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () =>
      queryResponse({ ids: [], list: [], total: 0, broaderTotal: 0 }));
    await client.searchEmails({ query: 'x', excludeDrafts: true });
    const batch = makeReq.mock.calls[0].arguments[0].methodCalls;
    const countFilter = batch[2][1].filter;
    // $draft cond survives; inMailboxOtherThan is gone -> count filter != visible filter.
    const flatHasDraft = countFilter.notKeyword === '$draft'
      || (countFilter.conditions || []).some((c: any) => c.notKeyword === '$draft');
    assert.ok(flatHasDraft);
    const stringified = JSON.stringify(countFilter);
    assert.ok(!stringified.includes('inMailboxOtherThan'));
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

  it('resolves the destination by name and targets the resolved id', async () => {
    let call = 0;
    const makeReq = mock.method(client, 'makeRequest', async () => {
      call++;
      if (call === 1) {
        return { methodResponses: [['Email/get', { list: [{ id: 'e1', mailboxIds: { 'mb-inbox': true } }] }, 'getEmails']] };
      }
      return { methodResponses: [['Email/set', { updated: { e1: null } }, 'bulkMove']] };
    });
    await client.bulkMove(['e1'], 'Archive');
    const update = makeReq.mock.calls[1].arguments[0].methodCalls[0][1].update;
    assert.equal(update.e1['mailboxIds/mb-archive'], true);
    assert.equal(update.e1['mailboxIds/mb-inbox'], null);
  });

  it('throws InvalidInputError on an unknown destination', async () => {
    await assert.rejects(
      () => client.bulkMove(['e1'], 'nope'),
      (err: Error) => { assert.ok(err instanceof InvalidInputError); return true; },
    );
  });
});

// ---------- label tools reject non-id mailbox values ----------

describe('label mailboxId validation', () => {
  let client: JmapClient;

  beforeEach(() => {
    client = makeClient();
    stubMailboxes(client);
  });

  it('addLabels rejects a value that is not a real mailbox id', async () => {
    await assert.rejects(
      () => client.addLabels('e1', ['Archive']), // a name, not an id
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /mailbox IDs only/);
        return true;
      },
    );
  });

  it('bulkAddLabels accepts real ids and proceeds', async () => {
    const makeReq = mock.method(client, 'makeRequest', async () =>
      ({ methodResponses: [['Email/set', { updated: { e1: null } }, 'bulkAddLabels']] }));
    await client.bulkAddLabels(['e1'], ['mb-archive']);
    assert.equal(makeReq.mock.calls.length, 1);
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
