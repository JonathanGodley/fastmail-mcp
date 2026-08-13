import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readThread, assertThreadBodiesWithinCap, THREAD_BODY_BYTE_CAP } from './thread-handler.js';
import type { ThreadClient } from './thread-handler.js';
import { InvalidInputError } from './coerce.js';

// A raw-JMAP-shaped thread message with a plain-text body, as Email/get returns it under
// includeBodies (text values fetched, html values deliberately not).
function makeEmail(id: string, text: string, over: any = {}) {
  return {
    id,
    subject: `Message ${id}`,
    from: [{ name: 'Alice Example', email: 'alice@example.com' }],
    receivedAt: '2026-06-15T03:29:02Z',
    textBody: [{ partId: `t-${id}`, type: 'text/plain' }],
    bodyValues: { [`t-${id}`]: { value: text } },
    keywords: { $seen: true },
    ...over,
  };
}

// An HTML-only message: its only body part is text/html, and no html value is fetched.
function makeHtmlOnlyEmail(id: string) {
  return {
    id,
    subject: `Message ${id}`,
    from: [{ email: 'alice@example.com' }],
    receivedAt: '2026-06-15T03:29:02Z',
    textBody: [{ partId: `h-${id}`, type: 'text/html' }],
    htmlBody: [{ partId: `h-${id}`, type: 'text/html' }],
    bodyValues: {},
    keywords: { $seen: true },
  };
}

interface Call { threadId: string; includeDrafts?: boolean; includeBodies?: boolean }

function makeClient(emails: any[], hiddenDraftCount = 0): { client: ThreadClient; calls: Call[] } {
  const calls: Call[] = [];
  const client: ThreadClient = {
    async getThread(threadId, includeDrafts, includeBodies) {
      calls.push({ threadId, includeDrafts, includeBodies });
      return { emails, hiddenDraftCount };
    },
  };
  return { client, calls };
}

// The response is a JSON array plus (sometimes) a trailing note; parse just the JSON.
function parseMessages(text: string): any[] {
  return JSON.parse(text.split('\n\nNote:')[0]);
}

describe('readThread — flag wiring', () => {
  it('asks for no bodies by default', async () => {
    const { client, calls } = makeClient([makeEmail('e1', 'hello')]);
    const text = await readThread({ threadId: 't1' }, client);
    assert.deepEqual(calls, [{ threadId: 't1', includeDrafts: false, includeBodies: false }]);
    assert.equal(parseMessages(text)[0].bodyTextUnavailable, undefined);
  });

  it('passes includeBodies and includeDrafts through, coercing stringified booleans', async () => {
    const { client, calls } = makeClient([makeEmail('e1', 'hello')]);
    await readThread({ threadId: 't1', includeBodies: 'true', includeDrafts: 'true' }, client);
    assert.deepEqual(calls, [{ threadId: 't1', includeDrafts: true, includeBodies: true }]);
  });

  it('returns each message body under includeBodies', async () => {
    const { client } = makeClient([makeEmail('e1', 'first message'), makeEmail('e2', 'second message')]);
    const messages = parseMessages(await readThread({ threadId: 't1', includeBodies: true }, client));
    assert.deepEqual(messages.map((m: any) => m.bodyText), ['first message', 'second message']);
    // No strip requested, so no strip signal is emitted.
    assert.equal(messages[0].quotedBytesStripped, undefined);
  });
});

describe('readThread — rejected flag combinations', () => {
  it('rejects stripQuoted with raw', async () => {
    const { client } = makeClient([makeEmail('e1', 'hello')]);
    await assert.rejects(
      () => readThread({ threadId: 't1', raw: true, includeBodies: true, stripQuoted: true }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /stripQuoted cannot be combined with raw/);
        return true;
      },
    );
  });

  it('rejects stripQuoted without includeBodies rather than letting it do nothing', async () => {
    const { client, calls } = makeClient([makeEmail('e1', 'hello')]);
    await assert.rejects(
      () => readThread({ threadId: 't1', stripQuoted: true }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /nothing to strip without includeBodies/);
        return true;
      },
    );
    assert.equal(calls.length, 0, 'must reject before fetching');
  });
});

describe('readThread — stripQuoted composition', () => {
  const quoted = (newText: string) =>
    `${newText}\n\nOn Mon, Jun 15, 2026, Alice wrote:\n> the whole earlier conversation\n> repeated again`;

  it('returns each message new text only, with a per-message signal', async () => {
    const { client } = makeClient([makeEmail('e1', quoted('First reply')), makeEmail('e2', quoted('Second reply'))]);
    const messages = parseMessages(
      await readThread({ threadId: 't1', includeBodies: true, stripQuoted: true }, client),
    );
    assert.deepEqual(messages.map((m: any) => m.bodyText), ['First reply', 'Second reply']);
    for (const m of messages) assert.ok(m.quotedBytesStripped > 0);
  });

  it('reports 0 for a message with no recognisable quote (body left verbatim)', async () => {
    const { client } = makeClient([makeEmail('e1', 'A short note with no quote.')]);
    const messages = parseMessages(
      await readThread({ threadId: 't1', includeBodies: true, stripQuoted: true }, client),
    );
    assert.equal(messages[0].bodyText, 'A short note with no quote.');
    assert.equal(messages[0].quotedBytesStripped, 0);
  });

  it('flags an HTML-only message instead of silently returning it body-less', async () => {
    const { client } = makeClient([makeEmail('e1', 'plain'), makeHtmlOnlyEmail('e2')]);
    const messages = parseMessages(
      await readThread({ threadId: 't1', includeBodies: true, stripQuoted: true }, client),
    );
    assert.equal(messages[1].bodyTextUnavailable, true);
    assert.equal(messages[1].quotedStripSkipped, 'no non-empty plain-text body to strip');
    assert.equal(messages[0].bodyTextUnavailable, undefined);
  });

  it('flags an HTML-only message under includeBodies even without stripQuoted', async () => {
    const { client } = makeClient([makeHtmlOnlyEmail('e1')]);
    const messages = parseMessages(await readThread({ threadId: 't1', includeBodies: true }, client));
    assert.equal(messages[0].bodyTextUnavailable, true);
    assert.equal(messages[0].quotedStripSkipped, undefined);
  });
});

describe('readThread — total body size cap', () => {
  const bigQuote = 'q'.repeat(60_000);
  const heavy = (id: string, newText: string) =>
    makeEmail(id, `${newText}\n\nOn Mon, Jun 15, 2026, Alice wrote:\n> ${bigQuote}`);

  it('errors instead of truncating, naming the largest messages and the remedy', async () => {
    const { client } = makeClient([heavy('e1', 'one'), heavy('e2', 'two')]);
    await assert.rejects(
      () => readThread({ threadId: 't1', includeBodies: true }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /over the 100000-byte limit/);
        assert.match(err.message, /stripQuoted:true/);
        assert.match(err.message, /Largest: e1 \(\d+ bytes\)/);
        return true;
      },
    );
  });

  it('measures what it is about to return, so stripQuoted brings the same thread under the cap', async () => {
    const { client } = makeClient([heavy('e1', 'one'), heavy('e2', 'two')]);
    const messages = parseMessages(
      await readThread({ threadId: 't1', includeBodies: true, stripQuoted: true }, client),
    );
    assert.deepEqual(messages.map((m: any) => m.bodyText), ['one', 'two']);
  });

  it('points an already-stripped over-cap thread at get_email instead of stripQuoted', async () => {
    const { client } = makeClient([makeEmail('e1', 'x'.repeat(120_000))]);
    await assert.rejects(
      () => readThread({ threadId: 't1', includeBodies: true, stripQuoted: true }, client),
      (err: Error) => {
        assert.match(err.message, /fetch these messages individually with get_email/);
        assert.doesNotMatch(err.message, /Retry with stripQuoted/);
        return true;
      },
    );
  });

  it('applies the cap on the raw path, with a remedy raw can actually run', async () => {
    // "Retry with stripQuoted:true" would be rejected outright on the raw path, so the
    // raw message has to say drop raw / fetch individually instead.
    const { client } = makeClient([makeEmail('e1', 'x'.repeat(120_000))]);
    await assert.rejects(
      () => readThread({ threadId: 't1', includeBodies: true, raw: true }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /over the 100000-byte limit/);
        assert.match(err.message, /drop raw and retry with stripQuoted:true, or fetch the messages individually/);
        assert.doesNotMatch(err.message, /^.*Retry with stripQuoted:true to drop/);
        return true;
      },
    );
  });

  it('does not apply the cap when no bodies were requested', async () => {
    // The preview-sized compact response can never approach the cap; the guard must not
    // fire on metadata.
    const { client } = makeClient([makeEmail('e1', 'x'.repeat(120_000))]);
    const messages = parseMessages(await readThread({ threadId: 't1' }, client));
    assert.equal(messages.length, 1);
  });

  it('assertThreadBodiesWithinCap passes a thread exactly at the cap', () => {
    const mode = { stripQuoted: false, raw: false };
    assert.doesNotThrow(() => assertThreadBodiesWithinCap([{ id: 'e1', bytes: THREAD_BODY_BYTE_CAP }], mode));
    assert.throws(() => assertThreadBodiesWithinCap([{ id: 'e1', bytes: THREAD_BODY_BYTE_CAP + 1 }], mode));
  });
});

describe('readThread — hidden drafts', () => {
  it('keeps the hidden-draft note and returns no draft body', async () => {
    // The client filters drafts out of the fetched set, so a hidden draft's body is
    // dropped with the message: only the count survives.
    const { client } = makeClient([makeEmail('e1', 'visible message')], 1);
    const text = await readThread({ threadId: 't1', includeBodies: true }, client);
    assert.match(text, /Note: 1 draft\(s\) in this thread are hidden/);
    const messages = parseMessages(text);
    assert.equal(messages.length, 1);
    assert.equal(messages[0].bodyText, 'visible message');
  });

  it('omits the note when nothing was hidden', async () => {
    const { client } = makeClient([makeEmail('e1', 'visible message')], 0);
    const text = await readThread({ threadId: 't1', includeBodies: true }, client);
    assert.doesNotMatch(text, /hidden/);
  });

  it('keeps the raw path pure JSON — no note appended', async () => {
    const { client } = makeClient([makeEmail('e1', 'visible message')], 2);
    const text = await readThread({ threadId: 't1', raw: true }, client);
    assert.deepEqual(JSON.parse(text).map((e: any) => e.id), ['e1']);
    assert.doesNotMatch(text, /hidden/);
  });
});
