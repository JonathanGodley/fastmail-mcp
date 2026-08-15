// The mailbox tools' orchestration, exercised through an injected client. The CallTool
// switch has no harness, so this is what keeps the path column, the parent narrowing, the
// raw branch and the no-path note covered by `npm test` rather than only by a live run.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { listMailboxes, createMailbox } from './mailbox-handler.js';
import type { MailboxClient } from './mailbox-handler.js';

const TREE = [
  { id: 'mb-archive', name: 'Archive', role: 'archive', totalEmails: 3 },
  { id: 'mb-2026', name: '2026', parentId: 'mb-archive', totalEmails: 2 },
  { id: 'mb-q1', name: 'Q1', parentId: 'mb-2026', totalEmails: 1 },
];

function stubClient(overrides: Partial<MailboxClient> = {}): MailboxClient {
  return {
    getMailboxes: async () => TREE,
    createMailbox: async () => ({ mailbox: { id: 'mb-new', name: 'Receipts', parentId: null }, created: { id: 'mb-new' }, path: 'Receipts' }),
    ...overrides,
  };
}

function parseFirst(content: Array<{ type: 'text'; text: string }>): any {
  return JSON.parse(content[0].text);
}

describe('list_mailboxes handler', () => {
  it('emits a root-anchored path for every mailbox in the default shape', async () => {
    const listed = parseFirst(await listMailboxes({}, stubClient()));
    assert.deepEqual(listed.map((mb: any) => mb.path), ['Archive', 'Archive/2026', 'Archive/2026/Q1']);
  });

  it('omits path under raw, which is untransformed JMAP', async () => {
    const content = await listMailboxes({ raw: true }, stubClient());
    assert.equal(content.length, 1);
    const listed = parseFirst(content);
    assert.equal(listed.length, 3);
    for (const mb of listed) assert.equal(mb.path, undefined);
  });

  it('honours a stringified raw, like every other flag on this server', async () => {
    const listed = parseFirst(await listMailboxes({ raw: 'false' }, stubClient()));
    assert.equal(listed[0].path, 'Archive');
  });

  it('narrows to the direct children of parent, by any accepted form', async () => {
    for (const form of ['mb-archive', 'archive', 'Archive']) {
      const listed = parseFirst(await listMailboxes({ parent: form }, stubClient()));
      assert.deepEqual(listed.map((mb: any) => mb.id), ['mb-2026'], form);
    }
    const nested = parseFirst(await listMailboxes({ parent: 'Archive/2026' }, stubClient()));
    assert.deepEqual(nested.map((mb: any) => mb.id), ['mb-q1']);
  });

  it('still computes root-anchored paths when the listing is narrowed', async () => {
    const listed = parseFirst(await listMailboxes({ parent: 'Archive/2026' }, stubClient()));
    assert.equal(listed[0].path, 'Archive/2026/Q1');
  });

  it('rejects an unknown parent rather than listing everything', async () => {
    await assert.rejects(() => listMailboxes({ parent: 'nope' }, stubClient()), /not found/);
  });

  it('includes verbose fields only when asked', async () => {
    const client = stubClient({ getMailboxes: async () => [{ ...TREE[0], sortOrder: 4 }] });
    assert.equal(parseFirst(await listMailboxes({}, client))[0].sortOrder, undefined);
    assert.equal(parseFirst(await listMailboxes({ verbose: true }, client))[0].sortOrder, 4);
  });

  it('reports a mailbox with no computable path in a SEPARATE content item, never inside the JSON', async () => {
    const broken = [
      { id: 'mb-ok', name: 'Archive' },
      { id: 'mb-orphan', name: 'Orphan', parentId: 'mb-gone' },
    ];
    const content = await listMailboxes({}, stubClient({ getMailboxes: async () => broken }));
    assert.equal(content.length, 2);
    const listed = parseFirst(content);
    assert.equal(listed[0].path, 'Archive');
    assert.equal(listed[1].path, undefined);
    assert.match(content[1].text, /mb-orphan/);
    // The first item stays parseable JSON with no note spliced into it.
    assert.doesNotMatch(content[0].text, /have no `path`/);
  });

  it('emits no note when every listed mailbox has a path', async () => {
    assert.equal((await listMailboxes({}, stubClient())).length, 1);
  });

  it('emits no note under raw, which promises no path in the first place', async () => {
    const broken = [{ id: 'mb-orphan', name: 'Orphan', parentId: 'mb-gone' }];
    const content = await listMailboxes({ raw: true }, stubClient({ getMailboxes: async () => broken }));
    assert.equal(content.length, 1);
  });
});

describe('create_mailbox handler', () => {
  it('returns the created mailbox with its path in the default shape', async () => {
    const created = parseFirst(await createMailbox({ name: 'Receipts' }, stubClient()));
    assert.equal(created.id, 'mb-new');
    assert.equal(created.path, 'Receipts');
  });

  it('passes name and parent through to the client untouched', async () => {
    const seen: any[] = [];
    const client = stubClient({
      createMailbox: async (input) => {
        seen.push(input);
        return { mailbox: { id: 'mb-new', name: input.name }, created: { id: 'mb-new' }, path: `Archive/${input.name}` };
      },
    });
    await createMailbox({ name: 'Receipts', parent: 'Archive' }, client);
    assert.deepEqual(seen, [{ name: 'Receipts', parent: 'Archive' }]);
  });

  it('returns the untouched JMAP created object under raw', async () => {
    const client = stubClient({
      createMailbox: async () => ({
        mailbox: { id: 'mb-new', name: 'Receipts', parentId: null },
        created: { id: 'mb-new', sortOrder: 7 },
        path: 'Receipts',
      }),
    });
    const raw = parseFirst(await createMailbox({ name: 'Receipts', raw: true }, client));
    assert.deepEqual(raw, { id: 'mb-new', sortOrder: 7 });
  });

  it('reports a missing path in a separate content item instead of dropping the field silently', async () => {
    const client = stubClient({
      createMailbox: async () => ({ mailbox: { id: 'mb-new', name: 'Receipts', parentId: 'mb-loop' }, created: { id: 'mb-new' } }),
    });
    const content = await createMailbox({ name: 'Receipts', parent: 'mb-loop' }, client);
    assert.equal(content.length, 2);
    assert.equal(parseFirst(content).path, undefined);
    assert.match(content[1].text, /mb-new/);
  });
});
