import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { listEmailsTool, searchEmailsTool } from './email-list-handler.js';
import type { EmailListClient } from './email-list-handler.js';
import type { QueryResult } from './jmap-client.js';
import { buildExclusionNote, formatEmailQueryResult, formatRawQueryResult } from './response-formatters.js';

const EMAIL = {
  id: 'M1',
  threadId: 'T1',
  mailboxIds: { inbox: true },
  keywords: { $seen: true },
  from: [{ name: 'Ann', email: 'ann@example.com' }],
  to: [{ email: 'bob@example.com' }],
  subject: 'Hello',
  receivedAt: '2027-03-02T09:00:00Z',
  preview: 'Hi there',
};

const EXCLUSION = { hidden: 3, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] };

function page(overrides: Partial<QueryResult> = {}): QueryResult {
  return { items: [EMAIL, { ...EMAIL, id: 'M2' }], total: 10, position: 4, ...overrides };
}

function stubClient(result: QueryResult = page()) {
  const getEmails = mock.fn(async (..._args: Parameters<EmailListClient['getEmails']>) => result);
  const searchEmails = mock.fn(async (..._args: Parameters<EmailListClient['searchEmails']>) => result);
  return { client: { getEmails, searchEmails } satisfies EmailListClient, getEmails, searchEmails };
}

const TOOLS = [
  { name: 'listEmailsTool', run: listEmailsTool },
  { name: 'searchEmailsTool', run: searchEmailsTool },
] as const;

function calls(stub: ReturnType<typeof stubClient>) {
  return stub.getEmails.mock.callCount() + stub.searchEmails.mock.callCount();
}

// Every argument is bad at once; each call must refuse with exactly the error that argument
// alone produces, and fixing it must move the refusal on to the next one in `order`.
async function assertCheckOrder(run: typeof listEmailsTool, order: Array<[string, unknown]>) {
  const stub = stubClient();
  const bad: Record<string, unknown> = Object.fromEntries(order);
  for (const [param, value] of order) {
    let alone = '';
    await run({ [param]: value }, 20, stub.client).catch((err: Error) => { alone = err.message; });
    assert.notEqual(alone, '', `${param} alone was not refused`);
    await assert.rejects(() => run(bad, 20, stub.client), (err: Error) => err.message === alone, param);
    delete bad[param];
  }
  assert.equal(calls(stub), 0);
}

describe('listEmailsTool', () => {
  // MCP `arguments` is optional, so a call can arrive with none.
  it('runs with the defaults when args is undefined', async () => {
    const stub = stubClient();
    await listEmailsTool(undefined, 20, stub.client);
    assert.deepEqual(stub.getEmails.mock.calls[0].arguments, [{
      mailbox: undefined,
      limit: 20,
      position: undefined,
      ascending: false,
      includeTrash: false,
      includeSpam: false,
      excludeDrafts: false,
    }]);
  });

  it('forwards every option, with the flags false by default and the limit as passed', async () => {
    const stub = stubClient();
    await listEmailsTool({ mailbox: 'Inbox', limit: 999 }, 37, stub.client);
    assert.equal(stub.getEmails.mock.callCount(), 1);
    assert.equal(stub.searchEmails.mock.callCount(), 0);
    assert.deepEqual(stub.getEmails.mock.calls[0].arguments, [{
      mailbox: 'Inbox',
      limit: 37,
      position: undefined,
      ascending: false,
      includeTrash: false,
      includeSpam: false,
      excludeDrafts: false,
    }]);
  });

  it('coerces stringified flags and position', async () => {
    const stub = stubClient();
    await listEmailsTool({
      ascending: 'true', includeTrash: 'TRUE', includeSpam: '1', excludeDrafts: 'true', position: '8',
    }, 20, stub.client);
    assert.deepEqual(stub.getEmails.mock.calls[0].arguments, [{
      mailbox: undefined,
      limit: 20,
      position: 8,
      ascending: true,
      includeTrash: true,
      includeSpam: true,
      excludeDrafts: true,
    }]);

    const off = stubClient();
    await listEmailsTool({
      ascending: 'false', includeTrash: 'false', includeSpam: '0', excludeDrafts: 'false',
    }, 20, off.client);
    const opts = off.getEmails.mock.calls[0].arguments[0]!;
    assert.equal(opts.ascending, false);
    assert.equal(opts.includeTrash, false);
    assert.equal(opts.includeSpam, false);
    assert.equal(opts.excludeDrafts, false);
  });

  it('checks its arguments in a fixed order', async () => {
    await assertCheckOrder(listEmailsTool, [
      ['ascending', 'x'], ['raw', 'x'], ['fields', 'nope'], ['position', -1],
      ['includeTrash', 'x'], ['includeSpam', 'x'], ['excludeDrafts', 'x'],
    ]);
  });
});

describe('searchEmailsTool', () => {
  it('runs with the defaults when args is undefined', async () => {
    const stub = stubClient();
    await searchEmailsTool(undefined, 20, stub.client);
    assert.deepEqual(stub.searchEmails.mock.calls[0].arguments, [{
      query: undefined,
      from: undefined,
      to: undefined,
      cc: undefined,
      bcc: undefined,
      subject: undefined,
      hasAttachment: undefined,
      isUnread: undefined,
      isPinned: undefined,
      mailbox: undefined,
      requiredMailboxes: undefined,
      excludeMailboxes: undefined,
      after: undefined,
      before: undefined,
      limit: 20,
      position: undefined,
      ascending: false,
      excludeDrafts: false,
      includeTrash: false,
      includeSpam: false,
    }]);
  });

  it('forwards every option, with the filter flags absent and the others false by default', async () => {
    const stub = stubClient();
    await searchEmailsTool({
      query: 'q', from: 'ann@example.com', to: 'bob@example.com', cc: 'cy@example.com',
      bcc: 'di@example.com', subject: 's', mailbox: 'Inbox',
      after: '2027-01-01', before: '2027-02-01', limit: 999,
    }, 37, stub.client);
    assert.equal(stub.searchEmails.mock.callCount(), 1);
    assert.equal(stub.getEmails.mock.callCount(), 0);
    assert.deepEqual(stub.searchEmails.mock.calls[0].arguments, [{
      query: 'q',
      from: 'ann@example.com',
      to: 'bob@example.com',
      cc: 'cy@example.com',
      bcc: 'di@example.com',
      subject: 's',
      hasAttachment: undefined,
      isUnread: undefined,
      isPinned: undefined,
      mailbox: 'Inbox',
      requiredMailboxes: undefined,
      excludeMailboxes: undefined,
      after: '2027-01-01',
      before: '2027-02-01',
      limit: 37,
      position: undefined,
      ascending: false,
      excludeDrafts: false,
      includeTrash: false,
      includeSpam: false,
    }]);
  });

  it('coerces stringified flags, position and the mailbox arrays', async () => {
    const stub = stubClient();
    await searchEmailsTool({
      hasAttachment: 'true', isUnread: 'false', isPinned: '1', ascending: 'true',
      excludeDrafts: 'true', includeTrash: 'true', includeSpam: 'true', position: '6',
      requiredMailboxes: '["Work","Clients"]', excludeMailboxes: 'Archive',
    }, 20, stub.client);
    assert.deepEqual(stub.searchEmails.mock.calls[0].arguments, [{
      query: undefined,
      from: undefined,
      to: undefined,
      cc: undefined,
      bcc: undefined,
      subject: undefined,
      hasAttachment: true,
      isUnread: false,
      isPinned: true,
      mailbox: undefined,
      requiredMailboxes: ['Work', 'Clients'],
      excludeMailboxes: ['Archive'],
      after: undefined,
      before: undefined,
      limit: 20,
      position: 6,
      ascending: true,
      excludeDrafts: true,
      includeTrash: true,
      includeSpam: true,
    }]);
  });

  it('rejects an uncoercible mailbox array before the query', async () => {
    for (const param of ['requiredMailboxes', 'excludeMailboxes']) {
      const stub = stubClient();
      await assert.rejects(
        () => searchEmailsTool({ [param]: ['Work', 7] }, 20, stub.client),
        (err: Error) => err.name === 'InvalidInputError' && err.message.includes(param),
      );
      assert.equal(calls(stub), 0, param);
    }
  });

  it('checks its arguments in a fixed order', async () => {
    await assertCheckOrder(searchEmailsTool, [
      ['ascending', 'x'], ['raw', 'x'], ['fields', 'nope'], ['position', -1],
      ['requiredMailboxes', [1]], ['excludeMailboxes', [1]], ['hasAttachment', 'x'],
      ['isUnread', 'x'], ['isPinned', 'x'], ['excludeDrafts', 'x'], ['includeTrash', 'x'],
      ['includeSpam', 'x'],
    ]);
  });

  it('names its own argument when refusing a filter flag', async () => {
    for (const flag of ['hasAttachment', 'isUnread', 'isPinned']) {
      const stub = stubClient();
      await assert.rejects(
        () => searchEmailsTool({ [flag]: 'maybe' }, 20, stub.client),
        (err: Error) => err.name === 'InvalidInputError' && err.message.startsWith(`${flag} must be true or false`),
        flag,
      );
      assert.equal(calls(stub), 0, flag);
    }
  });
});

describe('listEmailsTool and searchEmailsTool', () => {
  for (const { name, run } of TOOLS) {
    describe(name, () => {
      it('renders the simplified page, nextPosition included, with no note when nothing is excluded', async () => {
        const result = page();
        const content = await run({}, 20, stubClient(result).client);
        assert.deepEqual(content, [{ type: 'text', text: formatEmailQueryResult(result, { fields: undefined }) }]);
        assert.match(content[0].text, /nextPosition: 6 \(pass position:6/);
      });

      it('renders the raw page for raw:true and the simplified one for raw:"false"', async () => {
        const result = page();
        const raw = await run({ raw: true }, 20, stubClient(result).client);
        assert.deepEqual(raw, [{ type: 'text', text: formatRawQueryResult(result) }]);
        const simplified = await run({ raw: 'false' }, 20, stubClient(result).client);
        assert.equal(simplified[0].text, formatEmailQueryResult(result));
        assert.notEqual(raw[0].text, simplified[0].text);
      });

      it('projects the simplified page onto `fields`', async () => {
        const result = page();
        const content = await run({ fields: ['subject'] }, 20, stubClient(result).client);
        assert.equal(content[0].text, formatEmailQueryResult(result, { fields: new Set(['subject']) }));
        assert.notEqual(content[0].text, formatEmailQueryResult(result));
      });

      it('appends the exclusion note after the JSON on both renderings', async () => {
        const result = page({ exclusion: EXCLUSION });
        const note = buildExclusionNote(EXCLUSION);
        assert.ok(note.length > 0);
        const simplified = await run({}, 20, stubClient(result).client);
        assert.equal(simplified[0].text, formatEmailQueryResult(result) + note);
        const raw = await run({ raw: true }, 20, stubClient(result).client);
        assert.equal(raw[0].text, formatRawQueryResult(result) + note);
      });

      it('rejects fields with raw:true before the query', async () => {
        const stub = stubClient();
        await assert.rejects(
          () => run({ raw: 'true', fields: ['subject'] }, 20, stub.client),
          (err: Error) => err.name === 'InvalidInputError' && /^fields cannot be combined with raw:true/.test(err.message),
        );
        assert.equal(calls(stub), 0);
      });

      it('rejects an unknown field before the query', async () => {
        const stub = stubClient();
        await assert.rejects(
          () => run({ fields: ['nope'] }, 20, stub.client),
          (err: Error) => err.name === 'InvalidInputError' && err.message.includes('nope'),
        );
        assert.equal(calls(stub), 0);
      });

      it('rejects an unusable position before the query, pointing at ascending', async () => {
        for (const position of [-1, '1.5', 'abc', [2], true]) {
          const stub = stubClient();
          await assert.rejects(
            () => run({ position }, 20, stub.client),
            (err: Error) => err.name === 'InvalidInputError' && /^position /.test(err.message),
            JSON.stringify(position),
          );
          assert.equal(calls(stub), 0);
        }
        await assert.rejects(
          () => run({ position: -1 }, 20, stubClient().client),
          /pass ascending:true/,
        );
      });

      it('rejects an uncoercible flag naming it', async () => {
        for (const flag of ['ascending', 'raw', 'includeTrash', 'includeSpam', 'excludeDrafts']) {
          const stub = stubClient();
          await assert.rejects(
            () => run({ [flag]: 'maybe' }, 20, stub.client),
            (err: Error) => err.name === 'InvalidInputError' && err.message.startsWith(`${flag} must be true or false`),
            flag,
          );
          assert.equal(calls(stub), 0, flag);
        }
      });
    });
  }
});
