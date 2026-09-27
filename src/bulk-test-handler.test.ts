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

  it('leaves a message whose keywords were not reported out of both steps', async () => {
    const { client, calls } = makeClient();
    const unknown = { id: 'e9', subject: 'Message e9', from: [{ email: 'a@example.com' }] };
    const odd = { ...unknown, id: 'e8', keywords: 'garbage' };
    await runBulkReadTest([email('e1', false), unknown, odd], false, client, noPause);
    assert.deepEqual(calls, [
      { ids: ['e1'], read: true },
      { ids: ['e1'], read: false },
    ]);
  });

  it('reports a message with no reported read state as excluded, not as unread', async () => {
    const { client } = makeClient();
    const text = await runBulkReadTest([{ id: 'e1' }, { id: 'e2', keywords: {} }], true, client, noPause);
    const report = JSON.parse(text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1));
    const [e1, e2] = report.testEmails;
    assert.equal('wasRead' in e1, false);
    assert.equal(e1.excluded, true);
    assert.equal(e2.wasRead, false);
    assert.equal('excluded' in e2, false);
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

  it('lists both planned steps on a dry run, the restore naming only the unread messages', async () => {
    const { client } = makeClient();
    const text = await runBulkReadTest([email('e1', true), email('e2', false)], true, client, noPause);
    const report = JSON.parse(text.slice(text.indexOf('{'), text.lastIndexOf('}') + 1));
    const status = 'DRY RUN - Would execute but not actually performed';
    assert.deepEqual(report.operations, [
      { name: 'bulk_mark_read', description: 'Mark 2 emails as read', parameters: { emailIds: ['e1', 'e2'], read: true }, status, executed: false },
      {
        name: 'bulk_mark_read (restore)',
        description: 'Mark the 1 of them that were unread before as unread again',
        parameters: { emailIds: ['e2'], read: false },
        status,
        executed: false,
      },
    ]);
  });

  it('pauses between the two steps and not after the last', async () => {
    const events: string[] = [];
    const client: BulkTestClient = { async bulkMarkRead() { events.push('write'); } };
    const pause = async (ms: number) => { events.push(`pause ${ms}`); };
    await runBulkReadTest([email('e1', true), email('e2', false)], false, client, pause);
    assert.deepEqual(events, ['write', 'pause 500', 'write']);
    events.length = 0;
    await runBulkReadTest([email('e1', true)], false, client, pause);
    assert.deepEqual(events, ['write']);
  });

  it('really waits between the steps when no pause is injected', { timeout: 5000 }, async () => {
    const { client } = makeClient();
    const started = Date.now();
    await runBulkReadTest([email('e1', true), email('e2', false)], false, client);
    assert.ok(Date.now() - started >= 400, `took ${Date.now() - started} ms`);
  });
});
