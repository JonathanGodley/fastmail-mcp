import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { runBulkReadTest, type BulkTestClient } from './bulk-test-handler.js';

function email(id: string, read: boolean) {
  return { id, subject: `Message ${id}`, from: [{ email: 'a@example.com' }], keywords: read ? { $seen: true } : {} };
}

function makeClient(failFirst = false): { client: BulkTestClient; calls: Array<{ ids: string[]; read: boolean }> } {
  const calls: Array<{ ids: string[]; read: boolean }> = [];
  const client: BulkTestClient = {
    async bulkMarkRead(ids, read) {
      calls.push({ ids, read });
      if (failFirst && calls.length === 1) throw new Error('Failed to mark as read 1 of 2 emails (1 succeeded)');
    },
  };
  return { client, calls };
}

const noPause = async () => {};

describe('runBulkReadTest', () => {
  it('restores each message to its own prior read state, leaving read messages read', async () => {
    const { client, calls } = makeClient();
    await runBulkReadTest([email('e1', true), email('e2', false), email('e3', true)], false, client, noPause);
    assert.deepEqual(calls, [
      { ids: ['e1', 'e2', 'e3'], read: true },
      { ids: ['e2'], read: false },
    ]);
  });

  it('writes nothing back when every message was already read', async () => {
    const { client, calls } = makeClient();
    await runBulkReadTest([email('e1', true), email('e2', true)], false, client, noPause);
    assert.deepEqual(calls, [{ ids: ['e1', 'e2'], read: true }]);
  });

  it('after a failed first step, restores only the messages that were unread', async () => {
    const { client, calls } = makeClient(true);
    const text = await runBulkReadTest([email('e1', true), email('e2', false)], false, client, noPause);
    assert.deepEqual(calls[1], { ids: ['e2'], read: false });
    assert.match(text, /"status":"FAILED"/);
  });

  it('writes nothing on a dry run, and reports each message’s prior state', async () => {
    const { client, calls } = makeClient();
    const text = await runBulkReadTest([email('e1', true), email('e2', false)], true, client, noPause);
    assert.equal(calls.length, 0);
    assert.match(text, /"wasRead":true/);
    assert.doesNotMatch(text, /undo/i);
  });
});
