// The experimental snooze tools (#35), through an injected transport and an injected fetch.
// The shapes the fakes answer with are the ones scripts/probes/snooze-*.probe.mjs measured.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { InvalidInputError } from './coerce.js';
import {
  SNOOZE_CAPABILITY, SNOOZE_TOOLS, WAKE_KEYWORDS, WEB_SESSION_EXPIRED, WebSessionError, WebSessionTransport,
  listSnoozedEmails, parseWakeTime, snoozeEmail, unsnoozeEmail,
} from './snooze.js';
import type { SnoozeTransport } from './snooze.js';

const NOW = new Date('2026-11-02T09:00:00Z');
const MAILBOXES = [
  { id: 'mb-inbox', name: 'Inbox', role: 'inbox' },
  { id: 'mb-snoozed', name: 'Snoozed', role: 'snoozed' },
  { id: 'mb-archive', name: 'Archive', role: 'archive' },
  { id: 'mb-label', name: 'Receipts', role: null },
];

// An in-memory account answering the four methods the tools send, recording every call.
function fakeAccount(emails: Record<string, any>, opts: { refuse?: any } = {}) {
  const calls: any[][] = [];
  const t: SnoozeTransport = {
    accountId: async () => 'acc',
    call: async (methodCalls) => {
      calls.push(methodCalls);
      return methodCalls.map(([method, a, tag]: any) => {
        if (method === 'Mailbox/get') return [method, { list: MAILBOXES }, tag];
        if (method === 'Email/get') {
          const ids = a.ids ?? Object.keys(emails);
          return [method, { list: ids.filter((id: string) => emails[id]).map((id: string) => ({ id, ...emails[id] })) }, tag];
        }
        if (method === 'Email/query') return [method, { ids: Object.keys(emails), total: Object.keys(emails).length }, tag];
        if (method === 'Email/set') {
          const [id, patch] = Object.entries<any>(a.update)[0];
          if (opts.refuse) return [method, { notUpdated: { [id]: opts.refuse } }, tag];
          const e = emails[id];
          for (const [k, v] of Object.entries(patch)) {
            if (k.startsWith('mailboxIds/')) {
              const mb = k.slice('mailboxIds/'.length);
              if (v === null) delete e.mailboxIds[mb]; else e.mailboxIds[mb] = v;
            } else if (k === 'snoozed/until') e.snoozed = { ...e.snoozed, until: v };
            else e[k] = v;
          }
          return [method, { updated: { [id]: null } }, tag];
        }
        throw new Error(`unexpected method ${method}`);
      });
    },
  };
  const sets = () => calls.flat().filter(([m]) => m === 'Email/set').map(([, a]) => a.update);
  return { t, calls, sets };
}

const out = (content: any) => JSON.parse(content[0].text);

describe('parseWakeTime', () => {
  it('takes an offset as written and returns UTC', () => {
    assert.equal(parseWakeTime('2026-11-03T09:00:00+11:00', NOW), '2026-11-02T22:00:00Z');
  });
  it('reads a value with no offset in the configured zone', () => {
    assert.equal(parseWakeTime('2026-11-03T09:00:00', NOW, 'Australia/Sydney'), '2026-11-02T22:00:00Z');
  });
  it('refuses a time that is not in the future', () => {
    assert.throws(() => parseWakeTime('2026-11-02T09:00:00Z', NOW), (e: any) => e instanceof InvalidInputError && /not in the future/.test(e.message));
  });
  it('refuses a missing value', () => {
    assert.throws(() => parseWakeTime(undefined, NOW), (e: any) => e instanceof InvalidInputError && /until is required/.test(e.message));
  });
  it('refuses something that is not a date', () => {
    assert.throws(() => parseWakeTime('tomorrow', NOW), InvalidInputError);
  });
});

describe('snooze_email', () => {
  it('moves an Inbox message to Snoozed with the web client\'s own shape, keeping other labels', async () => {
    const emails = { e1: { mailboxIds: { 'mb-inbox': true, 'mb-label': true }, snoozed: null } };
    const { t, sets } = fakeAccount(emails);
    const result = out(await snoozeEmail({ emailId: 'e1', until: '2026-11-03T09:00:00+11:00' }, t, NOW));
    assert.deepEqual(sets(), [{ e1: {
      'mailboxIds/mb-inbox': null,
      'mailboxIds/mb-snoozed': true,
      snoozed: { until: '2026-11-02T22:00:00Z', setKeywords: WAKE_KEYWORDS },
    } }]);
    assert.deepEqual(emails.e1.mailboxIds, { 'mb-label': true, 'mb-snoozed': true });
    assert.deepEqual(result, { emailId: 'e1', wakeAt: '2026-11-02T22:00:00Z', wakeInto: 'Inbox' });
  });

  it('changes only the wake time of a message already snoozed', async () => {
    const emails = { e1: { mailboxIds: { 'mb-snoozed': true }, snoozed: { until: '2026-12-01T00:00:00Z', setKeywords: WAKE_KEYWORDS } } };
    const { t, sets } = fakeAccount(emails);
    const result = out(await snoozeEmail({ emailId: 'e1', until: '2026-11-05T00:00:00Z' }, t, NOW));
    assert.deepEqual(sets(), [{ e1: { 'snoozed/until': '2026-11-05T00:00:00Z' } }]);
    assert.equal(result.previousWakeAt, '2026-12-01T00:00:00Z');
  });

  it('refuses a message outside the Inbox and writes nothing', async () => {
    const { t, sets } = fakeAccount({ e1: { mailboxIds: { 'mb-archive': true }, snoozed: null } });
    await assert.rejects(snoozeEmail({ emailId: 'e1', until: '2026-11-05T00:00:00Z' }, t, NOW), (e: any) => e instanceof InvalidInputError && /Inbox only/.test(e.message));
    assert.deepEqual(sets(), []);
  });

  it('validates until before touching the account', async () => {
    const { t, calls } = fakeAccount({});
    await assert.rejects(snoozeEmail({ emailId: 'e1', until: '2026-01-01T00:00:00Z' }, t, NOW), InvalidInputError);
    assert.equal(calls.length, 0);
  });

  it('reports an unknown id as not found', async () => {
    const { t } = fakeAccount({});
    await assert.rejects(snoozeEmail({ emailId: 'nope', until: '2026-11-05T00:00:00Z' }, t, NOW), /Email not found/);
  });

  it('surfaces the server\'s refusal', async () => {
    const { t } = fakeAccount({ e1: { mailboxIds: { 'mb-inbox': true }, snoozed: null } }, { refuse: { type: 'invalidProperties', description: 'mailboxIds' } });
    await assert.rejects(snoozeEmail({ emailId: 'e1', until: '2026-11-05T00:00:00Z' }, t, NOW), /refused the update: "?invalidProperties/);
  });
});

describe('unsnooze_email', () => {
  it('returns a snoozed message to the Inbox and clears the snooze', async () => {
    const emails = { e1: { mailboxIds: { 'mb-snoozed': true, 'mb-label': true }, snoozed: { until: '2026-12-01T00:00:00Z' } } };
    const { t, sets } = fakeAccount(emails);
    const result = out(await unsnoozeEmail({ emailId: 'e1' }, t));
    assert.deepEqual(sets(), [{ e1: { 'mailboxIds/mb-snoozed': null, 'mailboxIds/mb-inbox': true, snoozed: null } }]);
    assert.deepEqual(emails.e1.mailboxIds, { 'mb-label': true, 'mb-inbox': true });
    assert.deepEqual(result, { emailId: 'e1', movedTo: 'Inbox', wasWakeAt: '2026-12-01T00:00:00Z' });
  });

  it('refuses a message that is not snoozed and writes nothing', async () => {
    const { t, sets } = fakeAccount({ e1: { mailboxIds: { 'mb-inbox': true }, snoozed: null } });
    await assert.rejects(unsnoozeEmail({ emailId: 'e1' }, t), (e: any) => e instanceof InvalidInputError && /not snoozed/.test(e.message));
    assert.deepEqual(sets(), []);
  });
});

describe('list_snoozed_emails', () => {
  it('queries Snoozed by wake time and names where each message wakes', async () => {
    const emails = {
      a: { subject: 'A', from: [{ email: 'a@example.invalid' }], mailboxIds: { 'mb-snoozed': true }, snoozed: { until: '2026-11-03T00:00:00Z' } },
      b: { subject: 'B', from: [{ email: 'b@example.invalid' }], mailboxIds: { 'mb-snoozed': true }, snoozed: { until: '2026-11-04T00:00:00Z', moveToMailboxId: 'mb-label' } },
    };
    const { t, calls } = fakeAccount(emails);
    const result = out(await listSnoozedEmails({}, t, 50));
    const query = calls.flat().find(([m]) => m === 'Email/query')[1];
    assert.deepEqual(query.filter, { inMailbox: 'mb-snoozed' });
    assert.deepEqual(query.sort, [{ property: 'snoozedUntil', mailboxId: 'mb-snoozed', isAscending: true }]);
    assert.equal(query.limit, 50);
    assert.equal(result.total, 2);
    assert.deepEqual(result.emails.map((e: any) => [e.id, e.wakeAt, e.wakeInto]), [
      ['a', '2026-11-03T00:00:00Z', 'Inbox'],
      ['b', '2026-11-04T00:00:00Z', 'Receipts'],
    ]);
  });
});

describe('WebSessionTransport', () => {
  const COOKIE = '__Http-s_abc=cookie-value-not-a-real-credential';
  const TOKEN = 'token-value-not-a-real-credential'; // allowlist-secret (synthetic placeholder, not a token shape)
  const SESSION = [{ accessToken: TOKEN, apiUrl: 'https://api.fastmail.com/jmap/api/', primaryAccounts: { 'urn:ietf:params:jmap:mail': 'acc' } }];

  function fakeFetch(script: Array<(url: string, init: any) => { status: number; body?: any }>) {
    const seen: Array<{ url: string; init: any }> = [];
    const f = (async (url: string, init: any) => {
      seen.push({ url, init });
      const step = script.shift();
      if (!step) throw new Error(`unexpected fetch ${url}`);
      const { status, body } = step(url, init);
      return { status, ok: status >= 200 && status < 300, json: async () => body } as any;
    }) as unknown as typeof fetch;
    return { f, seen };
  }
  const sessions = () => ({ status: 200, body: SESSION });
  const api = (status = 200) => () => ({ status, body: { methodResponses: [['Core/echo', {}, 'a']] } });

  it('fetches the token with the cookie alone, then sends both on the JMAP call with the snooze capability', async () => {
    const { f, seen } = fakeFetch([sessions, api()]);
    await new WebSessionTransport(COOKIE, f).call([['Core/echo', {}, 'a']]);
    assert.equal(seen[0].url, 'https://api.fastmail.com/auth/sessions');
    assert.equal(seen[0].init.headers.Cookie, COOKIE);
    assert.equal(seen[0].init.headers.Authorization, undefined);
    assert.equal(seen[1].url, SESSION[0].apiUrl);
    assert.equal(seen[1].init.headers.Authorization, `Bearer ${TOKEN}`);
    assert.equal(seen[1].init.headers.Cookie, COOKIE);
    assert.ok(JSON.parse(seen[1].init.body).using.includes(SNOOZE_CAPABILITY));
  });

  it('refetches the token once on a 401 and retries', async () => {
    const { f, seen } = fakeFetch([sessions, api(401), sessions, api()]);
    const responses = await new WebSessionTransport(COOKIE, f).call([['Core/echo', {}, 'a']]);
    assert.equal(seen.length, 4);
    assert.equal(responses[0][0], 'Core/echo');
  });

  it('reports an ended login after a second 401', async () => {
    const { f } = fakeFetch([sessions, api(401), sessions, api(401)]);
    await assert.rejects(new WebSessionTransport(COOKIE, f).call([]), (e: any) => e instanceof WebSessionError && e.message === WEB_SESSION_EXPIRED);
  });

  it('reports an ended login when the cookie gets no session', async () => {
    const { f } = fakeFetch([() => ({ status: 200, body: [] })]);
    await assert.rejects(new WebSessionTransport(COOKIE, f).call([]), (e: any) => e instanceof WebSessionError && !e.message.includes('cookie-value'));
  });

  it('refuses an API URL off Fastmail\'s HTTPS hosts and sends nothing to it', async () => {
    const { f, seen } = fakeFetch([() => ({ status: 200, body: [{ ...SESSION[0], apiUrl: 'https://evil.example/jmap/api/' }] })]);
    await assert.rejects(new WebSessionTransport(COOKIE, f).call([]), /no usable apiUrl/);
    assert.equal(seen.length, 1);
  });
});

describe('snooze tool surface', () => {
  it('marks every snooze tool experimental and names the cookie it runs on', () => {
    assert.deepEqual(SNOOZE_TOOLS.map(t => t.name), ['list_snoozed_emails', 'snooze_email', 'unsnooze_email']);
    for (const tool of SNOOZE_TOOLS) assert.match(tool.description, /EXPERIMENTAL: runs on the Fastmail web login's session cookie \(FASTMAIL_WEB_COOKIE\)/);
  });

  it('documents exactly these tools in README\'s experimental section', () => {
    const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8').replace(/\r/g, '');
    const start = readme.indexOf('\n## Experimental: snooze');
    assert.notEqual(start, -1, 'README has no "## Experimental: snooze" section');
    const end = readme.indexOf('\n## ', start + 1);
    const section = readme.slice(start, end === -1 ? undefined : end);
    const documented = [...section.matchAll(/^- \*\*([a-z_]+)\*\*:/gm)].map(m => m[1]);
    assert.deepEqual(documented, SNOOZE_TOOLS.map(t => t.name));
  });
});
