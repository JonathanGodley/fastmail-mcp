import { describe, it, mock } from 'node:test';
import assert from 'node:assert/strict';
import { getEmailTool } from './email-get-handler.js';
import type { EmailGetClient } from './email-get-handler.js';
import { InvalidInputError, toolJson } from './coerce.js';
import { callArguments } from './testing/mock-calls.js';

const QUOTED_BODY = 'Thanks, that works.\n\n> Can you make Tuesday?\n> It is the only free slot.\n';

function message(overrides: Record<string, any> = {}): any {
  return {
    id: 'M1',
    threadId: 'T1',
    subject: 'Tuesday',
    from: [{ name: 'Ann', email: 'ann@example.com' }],
    to: [{ email: 'bob@example.com' }],
    receivedAt: '2026-03-02T09:00:00Z',
    keywords: { $seen: true },
    textBody: [{ partId: 't1', type: 'text/plain' }],
    htmlBody: [{ partId: 'h1', type: 'text/html' }],
    bodyValues: { t1: { value: QUOTED_BODY }, h1: { value: '<p>Thanks, that works.</p>' } },
    ...overrides,
  };
}

// Text-only, so a default read shows the whole body and can be issued a hash.
const draft = () => message({ keywords: { $draft: true }, htmlBody: [], bodyValues: { t1: { value: QUOTED_BODY } } });

function stubClient(email: any = message()) {
  const getEmailById = mock.fn(async (_id: string) => email);
  return { client: { getEmailById } satisfies EmailGetClient, getEmailById };
}

async function read(args: any, email?: any): Promise<Record<string, any>> {
  const { client } = stubClient(email);
  const content = await getEmailTool(args, client);
  assert.equal(content.length, 1);
  return JSON.parse(content[0].text);
}

async function assertRefusedBeforeFetch(args: any, pattern: RegExp) {
  const { client, getEmailById } = stubClient();
  await assert.rejects(
    () => getEmailTool(args, client),
    (err: Error) => err instanceof InvalidInputError && pattern.test(err.message),
  );
  assert.equal(getEmailById.mock.callCount(), 0, 'the refusal must come before the fetch');
}

describe('getEmailTool', () => {
  it('refuses a missing emailId before the fetch', async () => {
    await assertRefusedBeforeFetch({}, /^emailId is required$/);
  });

  it('fetches the message the caller named', async () => {
    const { client, getEmailById } = stubClient();
    await getEmailTool({ emailId: 'M1' }, client);
    assert.equal(callArguments(getEmailById)[0], 'M1');
  });

  it('returns the simplified message by default: text body, no HTML body', async () => {
    const result = await read({ emailId: 'M1' });
    assert.equal(result.bodyText, QUOTED_BODY);
    assert.equal(result.bodyHtml, undefined);
    assert.equal(result.bodyValues, undefined);
  });

  it('returns the fetched JMAP object unchanged under raw', async () => {
    const email = draft();
    const { client } = stubClient(email);
    const content = await getEmailTool({ emailId: 'M1', raw: true }, client);
    assert.equal(content[0].text, toolJson(email));
  });

  it('adds the HTML body under verbose', async () => {
    const result = await read({ emailId: 'M1', verbose: true });
    assert.equal(result.bodyHtml, '<p>Thanks, that works.</p>');
  });

  it('reads a stringified "false" verbose as false', async () => {
    const result = await read({ emailId: 'M1', verbose: 'false' });
    assert.equal(result.bodyHtml, undefined);
  });

  it('projects the simplified message onto fields', async () => {
    const result = await read({ emailId: 'M1', fields: ['subject'] });
    assert.deepEqual(result, { subject: 'Tuesday' });
  });

  it('turns on the HTML body when fields names bodyHtml, without verbose', async () => {
    const result = await read({ emailId: 'M1', fields: ['bodyHtml'] });
    assert.deepEqual(result, { bodyHtml: '<p>Thanks, that works.</p>' });
  });

  // RFC 8621 puts the text/html part in textBody when there is no text/plain alternative.
  const htmlOnly = () => message({
    textBody: [{ partId: 'h1', type: 'text/html' }],
    bodyValues: { h1: { value: '<p>Thanks, that works.</p>' } },
  });

  it('flags an HTML-only message bodyTextUnavailable when fields names only bodyText (#248)', async () => {
    const result = await read({ emailId: 'M1', fields: ['bodyText'] }, htmlOnly());
    assert.deepEqual(result, { bodyTextUnavailable: true });
  });

  it('flags an HTML-only message bodyTextUnavailable beside its HTML fallback on a default read', async () => {
    const result = await read({ emailId: 'M1' }, htmlOnly());
    assert.equal(result.bodyHtml, '<p>Thanks, that works.</p>');
    assert.equal(result.bodyTextUnavailable, true);
  });

  it('does not flag bodyTextUnavailable on a message with a plain-text body', async () => {
    const result = await read({ emailId: 'M1' });
    assert.equal(result.bodyTextUnavailable, undefined);
  });

  it('refuses fields with raw before the fetch', async () => {
    await assertRefusedBeforeFetch({ emailId: 'M1', raw: true, fields: ['subject'] }, /^fields cannot be combined with raw:true/);
  });

  it('strips quoted history from bodyText under stripQuoted', async () => {
    const result = await read({ emailId: 'M1', stripQuoted: true });
    assert.equal(result.bodyText, 'Thanks, that works.');
    assert.ok(result.quotedBytesStripped > 0, `quotedBytesStripped: ${result.quotedBytesStripped}`);
  });

  it('refuses stripQuoted with raw before the fetch', async () => {
    await assertRefusedBeforeFetch({ emailId: 'M1', raw: true, stripQuoted: true }, /^stripQuoted cannot be combined with raw/);
  });

  it('refuses fields before stripQuoted when both conflict with raw', async () => {
    await assertRefusedBeforeFetch(
      { emailId: 'M1', raw: true, fields: ['subject'], stripQuoted: true },
      /^fields cannot be combined with raw:true/,
    );
  });

  it('reads a stringified "false" raw as false, so stripQuoted is not refused', async () => {
    const result = await read({ emailId: 'M1', raw: 'false', stripQuoted: true });
    assert.equal(result.bodyText, 'Thanks, that works.');
  });

  it('attaches a draft body hash to a default draft read', async () => {
    const result = await read({ emailId: 'M1' }, draft());
    assert.match(result.bodyHash, /^bh1-/);
    assert.equal(result.bodyHashWithheld, undefined);
  });

  it('withholds the draft body hash, with the reason, when the read stripped quoted history', async () => {
    const result = await read({ emailId: 'M1', stripQuoted: true }, draft());
    assert.equal(result.bodyHash, undefined);
    assert.match(result.bodyHashWithheld, /stripped quoted history/);
  });

  it('attaches no body hash to a message that is not a draft', async () => {
    const result = await read({ emailId: 'M1' });
    assert.equal(result.bodyHash, undefined);
    assert.equal(result.bodyHashWithheld, undefined);
  });
});
