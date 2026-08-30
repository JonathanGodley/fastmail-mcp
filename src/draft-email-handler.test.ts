import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { composeDraftEmail } from './draft-email-handler.js';
import type { DraftEmailClient } from './draft-email-handler.js';

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------

// A raw-JMAP-shaped original, as getEmailById returns it.
function makeOriginal(over: any = {}) {
  return {
    id: 'o1',
    blobId: 'blob-orig',
    messageId: ['orig-msg@example.com'],
    references: ['root@example.com'],
    subject: 'Project update',
    from: [{ name: 'Jon Godley', email: 'jon@example.com' }],
    to: [{ email: 'me@example.com' }],
    sentAt: '2026-06-15T03:29:02Z',
    textBody: [{ partId: 't', type: 'text/plain' }],
    htmlBody: [{ partId: 'h', type: 'text/html' }],
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>original html</p>' } },
    ...over,
  };
}

const SIGNED_IDENTITY = {
  id: 'id-1',
  name: 'Test User',
  email: 'me@example.com',
  mayDelete: false,
  textSignature: 'Kind regards,\nTest User',
  htmlSignature: '<div>Kind regards,</div><div>Test User</div>',
};

const UNSIGNED_IDENTITY = { id: 'id-1', name: 'Test User', email: 'me@example.com' };

// An original whose html displays an embedded image the message really carries.
const inlinePng = {
  partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'pic.png',
  disposition: 'inline', cid: 'img-1',
};
function withInlineImage(over: any = {}) {
  return makeOriginal({
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:img-1"></p>' } },
    attachments: [inlinePng],
    ...over,
  });
}

/** A spy client: records what each method was called with, serves one canned original. */
function spyClient(original: any = makeOriginal(), over: Partial<DraftEmailClient> = {}) {
  const calls: any = { gets: [] as string[] };
  const client: DraftEmailClient = {
    getEmailById: async (id) => {
      calls.gets.push(id);
      // Serves two jobs: fetching the original, and the post-save read-back. The saved id
      // is what tells them apart, and the read-back must echo what this call attached or
      // every image-bearing draft grows a spurious "not found on the saved draft" note.
      if (id === 'draft-9') {
        return { id, attachments: (calls.draft?.attachments ?? []).map((p: any) => ({ ...p, size: 70 })) };
      }
      return original;
    },
    getIdentities: async () => [SIGNED_IDENTITY],
    uploadAttachments: async (specs, dir, allowBlob, options) => {
      calls.upload = { specs, dir, allowBlob, options };
      return [];
    },
    createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    ...over,
  };
  return { client, calls };
}

const compose = (args: any, client: DraftEmailClient, dir?: string, allowBlob = false) =>
  composeDraftEmail(args, client, dir, allowBlob);

function count(haystack: string, needle: string): number {
  return haystack.split(needle).length - 1;
}

/**
 * The refusal message a call produced, for assertions about what a message must NOT contain.
 * A `assert.rejects` regex can only pin what is present; an absent spelling needs the string.
 * Fails the test if the call resolved rather than refusing.
 */
async function messageFrom(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (err) {
    return (err as Error).message;
  }
  assert.fail('expected the call to be refused, but it resolved');
}

// ---------------------------------------------------------------------------
// Mode, and the parameters that belong to one mode
// ---------------------------------------------------------------------------

describe('draft_email — mode', () => {
  it('refuses a missing mode, naming all three and saying it is not defaulted', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => compose({ textBody: 'hi' }, client),
      /mode is required and must be exactly one of "new", "reply" or "forward".*not defaulted/s,
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses a mode that is not one of the three, rather than guessing', async () => {
    const { client } = spyClient();
    await assert.rejects(() => compose({ mode: 'Reply', textBody: 'hi' }, client), /must be exactly one of/);
    await assert.rejects(() => compose({ mode: 'replyall', textBody: 'hi' }, client), /must be exactly one of/);
  });

  it('requires originalEmailId on reply and forward, and refuses it on new', async () => {
    const { client } = spyClient();
    await assert.rejects(() => compose({ mode: 'reply', textBody: 'hi' }, client), /originalEmailId is required for mode:'reply'/);
    await assert.rejects(() => compose({ mode: 'forward', textBody: 'hi' }, client), /originalEmailId is required for mode:'forward'/);
    await assert.rejects(
      () => compose({ mode: 'new', textBody: 'hi', originalEmailId: 'o1' }, client),
      /originalEmailId applies to mode:'reply' and mode:'forward' only/,
    );
  });

  it('refuses each mode-only parameter in one message shape, so the rule is learnt once', async () => {
    const { client } = spyClient();
    const cases: [string, any][] = [
      ['mailbox', { mode: 'reply', originalEmailId: 'o1', textBody: 'hi', mailbox: 'Drafts' }],
      ['inReplyTo', { mode: 'reply', originalEmailId: 'o1', textBody: 'hi', inReplyTo: ['thread@example.com'] }],
      ['references', { mode: 'reply', originalEmailId: 'o1', textBody: 'hi', references: ['thread@example.com'] }],
      ['asAttachment', { mode: 'new', textBody: 'hi', asAttachment: true }],
      ['includeOriginalAttachments', { mode: 'new', textBody: 'hi', includeOriginalAttachments: false }],
    ];
    for (const [param, args] of cases) {
      await assert.rejects(
        () => compose(args, client),
        new RegExp(`^McpError: MCP error -32602: ${param} applies to mode:'(new|forward)' only \\(this call is mode:'(new|reply)'\\)\\.$`),
        param,
      );
    }
  });

  it('requires `to` on a forward — there is no default recipient', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', htmlBody: '{{forward}}' }, client),
      /to is required for a forward/,
    );
  });

  it('takes the recipient fields and the threading headers as an array OR a bare string', async () => {
    // Both declarations are `['array','string']` and both run through the lenient coercer,
    // so a client that sends one address or one Message-ID as a scalar is not rejected.
    const { client, calls } = spyClient();
    await compose(
      {
        mode: 'new', to: 'sam@example.com', cc: 'cara@example.com',
        inReplyTo: 'thread@example.com', references: 'root@example.com,thread@example.com',
        textBody: 'hi',
      },
      client,
    );
    assert.deepEqual(calls.draft.to, ['sam@example.com']);
    assert.deepEqual(calls.draft.cc, ['cara@example.com']);
    assert.deepEqual(calls.draft.inReplyTo, ['thread@example.com']);
    assert.deepEqual(calls.draft.references, ['root@example.com', 'thread@example.com']);
  });

  it("keeps create_draft's contentless guard on mode:'new', testing bodies with isBlank", async () => {
    const { client } = spyClient();
    await assert.rejects(() => compose({ mode: 'new' }, client), /At least one of to, subject, textBody, htmlBody, or attachments/);
    // A whitespace-only body is contentless too. The old handler tested truthiness here and
    // let this through to a later, vaguer refusal.
    await assert.rejects(() => compose({ mode: 'new', textBody: '   ' }, client), /At least one of to, subject/);
  });
});

// ---------------------------------------------------------------------------
// Token acceptance — every refusal from the scan alone
// ---------------------------------------------------------------------------

describe('draft_email — token refusals, decided before anything is built', () => {
  it('refuses a near-miss spelling rather than guessing what was meant', async () => {
    const { client, calls } = spyClient();
    for (const miss of ['{{Quote}}', '{{{quote}}}', '{quote}}']) {
      await assert.rejects(
        () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: `<p>hi</p>${miss}` }, client),
        /which is not a token.*on mode:'reply' the exact spellings are \{\{signature\}\} and \{\{quote\}\}/s,
        miss,
      );
    }
    assert.equal(calls.draft, undefined);
  });

  it('names only the spellings the mode accepts, so the refusal offers nothing it would reject', async () => {
    // The mode gate refuses the other mode's history token in every spelling, so listing it
    // here would hand the caller a spelling this same call is about to reject.
    const { client } = spyClient();

    const replyMessage = await messageFrom(
      () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{Signature}}' }, client),
    );
    assert.match(replyMessage, /which is not a token/);
    assert.ok(replyMessage.includes('{{quote}}'), replyMessage);
    assert.ok(!replyMessage.includes('{{forward}}'), replyMessage);

    const forwardMessage = await messageFrom(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>hi</p>{{Signature}}' },
        client,
      ),
    );
    assert.ok(forwardMessage.includes('{{forward}}'), forwardMessage);
    assert.ok(!forwardMessage.includes('{{quote}}'), forwardMessage);

    // A new message has no history token at all, so it is offered neither — and the sentence
    // goes singular rather than leaving "the exact spellings are {{signature}}".
    const newMessage = await messageFrom(
      () => compose({ mode: 'new', htmlBody: '<p>hi</p>{{Signature}}' }, client),
    );
    assert.match(newMessage, /on mode:'new' the exact spelling is \{\{signature\}\}, lower case/);
    assert.ok(!newMessage.includes('{{quote}}'), newMessage);
    assert.ok(!newMessage.includes('{{forward}}'), newMessage);

    // The half that was already right stays: the caller's own spelling is named back, and
    // the escape — valid in every mode — is still offered.
    for (const message of [replyMessage, forwardMessage, newMessage]) {
      assert.ok(message.includes('"{{Signature}}"'), message);
      assert.match(message, /escape them: \\\{\{signature\}\} ships the literal token/);
      assert.match(message, /consumed only before a token name/);
    }
  });

  it('treats internal whitespace as part of the token, not as a near-miss', async () => {
    // `\s` inside the braces is trimmed by the one alternation, so `{{ quote }}` IS the
    // token. Pinned here because it is the spelling a caller is likeliest to write, and
    // refusing it would read as a near-miss bug rather than as a rule.
    const { client, calls } = spyClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{ quote }}' }, client);
    assert.match(calls.draft.textBody, /wrote:\n> original text/);
  });

  it('offers each mode exactly its own history token', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{forward}}' }, client),
      /\{\{forward\}\} does not apply to mode:'reply'; use \{\{quote\}\} instead\./,
    );
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>hi</p>{{quote}}' }, client),
      /\{\{quote\}\} does not apply to mode:'forward'; use \{\{forward\}\} instead\./,
    );
    await assert.rejects(
      () => compose({ mode: 'new', htmlBody: '<p>hi</p>{{quote}}' }, client),
      /\{\{quote\}\} does not apply to mode:'new' \(a new message has no history to place\)\./,
    );
  });

  it('refuses {{forward}} alongside asAttachment: the original already rides whole', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], asAttachment: true, htmlBody: '<p>n</p>{{forward}}' },
        client,
      ),
      /\{\{forward\}\} does not apply to an asAttachment forward/,
    );
  });

  it('refuses a forward that places neither {{forward}} nor asAttachment', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>FYI</p>' }, client),
      /A forward must place \{\{forward\}\} in a body part, or pass asAttachment:true/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses a token placed in one supplied part but not the other', async () => {
    const { client } = spyClient();
    await assert.rejects(
      () => compose(
        { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p><div>{{signature}}</div>', textBody: 'hi' },
        client,
      ),
      /\{\{signature\}\} is in htmlBody but not in textBody\./,
    );
    await assert.rejects(
      () => compose(
        { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>', textBody: 'hi\n{{quote}}' },
        client,
      ),
      /\{\{quote\}\} is in textBody but not in htmlBody\./,
    );
  });

  it('refuses the same token twice in one part', async () => {
    const { client, calls } = spyClient();
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'a {{quote}} b {{quote}} c' }, client),
      /\{\{quote\}\} appears 2 times in textBody; a token may be placed once per part.*escape it \(\\\{\{quote\}\}\)/s,
    );
    await assert.rejects(
      () => compose({ mode: 'new', htmlBody: '<div>{{signature}}</div><div>{{signature}}</div>' }, client),
      /\{\{signature\}\} appears 2 times in htmlBody/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('counts unescaped EXACT tokens only, so an escaped twin is not a repeat', async () => {
    const { client, calls } = spyClient();
    await compose({ mode: 'new', textBody: 'a {{signature}} b \\{{signature}} c' }, client);
    assert.equal(calls.draft.textBody, 'a Kind regards,\nTest User b {{signature}} c');
  });

  it("refuses a wrong-mode token in ANY spelling, for the mode rather than the spelling", async () => {
    // The near-miss message lists this mode's own spellings, so on a reply it would answer
    // {{{forward}}} with {{signature}} and {{quote}} and never say the true thing — that a
    // reply has no forwarded block at all. The mode gate has to run first to say it.
    const { client } = spyClient();
    for (const spelling of ['{{forward}}', '{{Forward}}', '{{{forward}}}']) {
      await assert.rejects(
        () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: `<p>hi</p>${spelling}` }, client),
        /\{\{forward\}\} does not apply to mode:'reply'; use \{\{quote\}\} instead\./,
        spelling,
      );
    }
    // A single brace each side is ordinary prose and reaches no gate at all.
    const { client: c2, calls } = spyClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'see {forward} below\n{{quote}}' }, c2);
    assert.match(calls.draft.textBody, /see \{forward\} below/);
  });

  it('runs each gate across the WHOLE call, so no part can jump the order', async () => {
    // textBody is scanned first, so a per-part loop would report its near-miss and never
    // reach htmlBody's wrong-mode token.
    const { client } = spyClient();
    await assert.rejects(
      () => compose(
        { mode: 'reply', originalEmailId: 'o1', textBody: 'hi {{Quote}}', htmlBody: '<p>hi</p>{{forward}}' },
        client,
      ),
      /\{\{forward\}\} does not apply to mode:'reply'/,
    );
    // And the one-part refusal comes before both forward gates.
    await assert.rejects(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], asAttachment: true, htmlBody: '<p>n</p>{{forward}}', textBody: 'n' },
        client,
      ),
      /\{\{forward\}\} is in htmlBody but not in textBody\./,
    );
    await assert.rejects(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>n</p>{{signature}}', textBody: 'n' },
        client,
      ),
      /\{\{signature\}\} is in htmlBody but not in textBody\./,
    );
  });

  it('accepts a token in the only supplied part, and derives the other from it', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p><div>{{signature}}</div>' }, client,
    );
    assert.equal(r.emailId, 'draft-9');
    assert.match(calls.draft.htmlBody, /Kind regards/);
    assert.equal(calls.draft.textBody, undefined);
  });
});

// ---------------------------------------------------------------------------
// THE SECURITY PIN
// ---------------------------------------------------------------------------
//
// Expansion is a SINGLE PASS over the caller's OWN authored body, run before any fetched
// content is joined in. The quote and forwarded blocks are built from an attacker-authored
// original: a token spelling inside a stranger's email lands on the REPLACEMENT side of the
// substitution, where String.prototype.replace never rescans it, and must ship as inert text.
//
// Six pairs — each of the two history blocks against each of the three spellings — asserted
// in BOTH parts. The sharp assertion is the COUNT: the spelling appears exactly once in the
// stored body (the copy carried out of the original) and never twice, never zero times, and
// the signature block's own text never appears at all, because the caller placed no
// {{signature}} anywhere in these calls.

describe('draft_email — a token inside a fetched block stays inert (single-pass expansion)', () => {
  const SPELLINGS = ['{{signature}}', '{{quote}}', '{{forward}}'] as const;

  function poisoned(spelling: string) {
    return makeOriginal({
      bodyValues: {
        t: { value: `before ${spelling} after` },
        h: { value: `<p>before ${spelling} after</p>` },
      },
    });
  }

  for (const spelling of SPELLINGS) {
    it(`carries ${spelling} out of a QUOTED original as literal text, in both parts`, async () => {
      const { client, calls } = spyClient(poisoned(spelling));
      await compose(
        {
          mode: 'reply', originalEmailId: 'o1',
          htmlBody: '<p>my reply</p>{{quote}}',
          textBody: 'my reply\n{{quote}}',
        },
        client,
      );
      for (const part of ['htmlBody', 'textBody'] as const) {
        const body: string = calls.draft[part];
        assert.equal(count(body, spelling), 1, `${part} should carry exactly one inert ${spelling}`);
        assert.doesNotMatch(body, /Kind regards/, `${part} must not have grown a signature`);
      }
      // A re-run would have expanded the carried {{quote}} into a second attribution line.
      assert.equal(count(calls.draft.textBody, 'wrote:'), 1);
    });

    it(`carries ${spelling} out of a FORWARDED original as literal text, in both parts`, async () => {
      const { client, calls } = spyClient(poisoned(spelling));
      await compose(
        {
          mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'],
          htmlBody: '<p>FYI</p>{{forward}}',
          textBody: 'FYI\n{{forward}}',
        },
        client,
      );
      for (const part of ['htmlBody', 'textBody'] as const) {
        const body: string = calls.draft[part];
        assert.equal(count(body, spelling), 1, `${part} should carry exactly one inert ${spelling}`);
        assert.doesNotMatch(body, /Kind regards/, `${part} must not have grown a signature`);
      }
      // A re-run would have expanded the carried {{forward}} into a second header block.
      assert.equal(count(calls.draft.textBody, '----- Original message -----'), 1);
    });
  }
});

// ---------------------------------------------------------------------------
// The image-ordering pin
// ---------------------------------------------------------------------------

describe('draft_email — the authored-image plan reads PRE-expansion, the closure check POST', () => {
  it('embeds a quoted image with no attachment source enabled, and points the body at it', async () => {
    // This one call pins BOTH ends of the order. planAuthoredInlineImages runs on the
    // caller's own html, which carries no cid: reference at all — run after expansion it
    // would see the block's minted identifier as an authored reference with no attachment
    // source enabled and refuse. checkInlineClosure runs on what actually ships — run
    // before expansion it would find the minted cid referenced nowhere and throw.
    const { client, calls } = spyClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{quote}}' }, client,
    );
    assert.equal(r.emailId, 'draft-9');
    const minted = calls.draft.attachments.filter((p: any) => p.blobId === 'blob-png');
    assert.equal(minted.length, 1);
    assert.match(minted[0].cid, /^ii-[0-9a-f]{32}@inline\.invalid$/);
    assert.equal(minted[0].disposition, 'inline');
    // The stored quote points at the MINTED identifier, not the original's own cid:img-1.
    assert.ok(calls.draft.htmlBody.includes(`src="cid:${minted[0].cid}"`), calls.draft.htmlBody);
    assert.doesNotMatch(calls.draft.htmlBody, /cid:img-1/);
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the quoted message (1 KB).',
      'Identity me@example.com has a signature; this body has no {{signature}}, so it was ' +
      'stored as written. To add one on edit_draft, place the token and pass expandSignature: true.',
    ]);
  });

  it('drops a minted image the expanded body does not really reference, and says so', async () => {
    // A token expanded inside an html comment leaves the block's markup in the body with its
    // <img> outside the document. The old path for an unreferenced minted part was
    // checkInlineClosure THROWING, which is the wrong answer for something a caller can
    // cause: it is dropped before assembly and named on the result instead.
    const { client, calls } = spyClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p><!-- {{quote}} -->' }, client,
    );
    assert.equal(r.emailId, 'draft-9');
    assert.equal((calls.draft.attachments ?? []).length, 0);
    assert.ok(
      r.notes!.some((n) => /were dropped: after expansion no body written by this call references them/.test(n)),
      JSON.stringify(r.notes),
    );
  });
});

// ---------------------------------------------------------------------------
// Placement: nothing is added that the caller did not place
// ---------------------------------------------------------------------------

describe('draft_email — nothing is added that the caller did not place', () => {
  it('stores a reply with no {{signature}} unsigned, and warns that the identity has one', async () => {
    const { client, calls } = spyClient();
    const r = await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{quote}}' }, client);
    assert.doesNotMatch(calls.draft.textBody, /Kind regards/);
    assert.ok(r.notes!.some((n) => /has a signature; this body has no \{\{signature\}\}/.test(n)));
  });

  it('says nothing about the signature when the identity has none', async () => {
    const { client } = spyClient(makeOriginal(), { getIdentities: async () => [UNSIGNED_IDENTITY] });
    const r = await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{quote}}' }, client);
    assert.equal(r.notes, undefined);
  });

  it('stores a reply with no {{quote}} without the conversation, and says so', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{signature}}' }, client,
    );
    assert.equal(calls.draft.textBody, 'hi\nKind regards,\nTest User');
    assert.doesNotMatch(calls.draft.textBody, /wrote:/);
    assert.ok(r.notes!.includes(
      'This reply was stored without the original: place {{quote}} in the body to include it.',
    ));
  });

  it('substitutes the block at the token position with nothing added around it', async () => {
    // The spacing at the join is the caller's. A sign-off placed BELOW the history is stored
    // below it, unjudged, and shows up in the receipt's position order.
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: '{{quote}}\n---\n{{signature}}' }, client,
    );
    assert.ok(calls.draft.textBody.endsWith('\n---\nKind regards,\nTest User'), calls.draft.textBody);
    assert.deepEqual(r.tokens!.parts[0].order, ['quote', 'signature']);
  });

  it('threads a reply and defaults the recipient to the original sender', async () => {
    const { client, calls } = spyClient();
    const r = await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{quote}}' }, client);
    assert.equal(r.subject, 'Re: Project update');
    assert.deepEqual(calls.draft.to, ['Jon Godley <jon@example.com>']);
    assert.deepEqual(calls.draft.inReplyTo, ['orig-msg@example.com']);
    assert.deepEqual(calls.draft.references, ['root@example.com', 'orig-msg@example.com']);
    assert.equal(calls.draft.sourceEmailId, 'o1');
  });

  it('records the forwarded source and starts a new conversation', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>FYI</p>{{forward}}' },
      client,
    );
    assert.equal(r.subject, 'Fwd: Project update');
    assert.deepEqual(calls.draft.forwardedMessageId, ['orig-msg@example.com']);
    assert.equal(calls.draft.inReplyTo, undefined);
    assert.equal(calls.draft.sourceEmailId, 'o1');
  });
});

// ---------------------------------------------------------------------------
// A token with nothing behind it
// ---------------------------------------------------------------------------

describe('draft_email — a token with nothing to expand to', () => {
  const NOTHING_QUOTABLE = () => makeOriginal({
    textBody: undefined, htmlBody: undefined, bodyValues: {},
  });

  it('removes it, stores the rest as written, and names the cause per part', async () => {
    const { client, calls } = spyClient(NOTHING_QUOTABLE());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{quote}}' }, client,
    );
    assert.equal(calls.draft.textBody, 'hi\n');
    assert.ok(r.notes!.includes(
      '{{quote}} in textBody was removed: the original has nothing quotable in any format. ' +
      'The rest of that part was stored as written.',
    ));
    assert.deepEqual(r.tokens!.parts[0].removed, [
      { token: 'quote', count: 1, cause: 'nothing-quotable' },
    ]);
    assert.deepEqual(r.tokens!.parts[0].expanded, []);
  });

  it('removes {{signature}} when the identity has none, naming that cause instead', async () => {
    const { client } = spyClient(makeOriginal(), { getIdentities: async () => [UNSIGNED_IDENTITY] });
    const r = await compose({ mode: 'new', textBody: 'hi\n{{signature}}' }, client);
    assert.deepEqual(r.tokens!.parts[0].removed, [
      { token: 'signature', count: 1, cause: 'no-signature' },
    ]);
    assert.ok(r.notes!.some((n) => /the sending identity has no signature configured/.test(n)));
  });

  it('refuses a part that was content before expansion and is empty after it', async () => {
    const { client, calls } = spyClient(NOTHING_QUOTABLE());
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '{{quote}}' }, client),
      /htmlBody is empty after expansion: it was nothing but tokens, and the original has nothing quotable in any format\..*Write prose beside the token/s,
    );
    assert.equal(calls.draft, undefined);
  });
});

// ---------------------------------------------------------------------------
// asAttachment
// ---------------------------------------------------------------------------

describe("draft_email — mode:'forward' with asAttachment", () => {
  it('attaches the original whole and writes the filler note when there is no body', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], asAttachment: true }, client,
    );
    assert.equal(calls.draft.textBody, 'Forwarded message attached.');
    assert.equal(calls.draft.htmlBody, undefined);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-orig', type: 'message/rfc822', name: 'Project update.eml', disposition: 'attachment' },
    ]);
    assert.equal(r.tokens!.fillerBody, true);
  });

  it('leaves a caller-supplied note alone and reports no filler', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], asAttachment: true, textBody: 'see attached' },
      client,
    );
    assert.equal(calls.draft.textBody, 'see attached');
    assert.equal(r.tokens, undefined); // no token written at all, so no receipt
  });
});

// ---------------------------------------------------------------------------
// A forwarded block that ships only in its text form
// ---------------------------------------------------------------------------

describe('draft_email — {{forward}} shipping in the text form over an html original', () => {
  const TEXT_FORM_NOTE =
    '{{forward}} ships in the text form only and the original ships HTML, so this forward ' +
    'loses its formatting and its inline images ride as attachments; put {{forward}} in ' +
    'htmlBody to keep both.';

  it('says the formatting is lost and names the remedy that fixes it', async () => {
    const { client, calls } = spyClient(withInlineImage());
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], textBody: 'FYI\n{{forward}}' },
      client,
    );
    assert.equal(calls.draft.htmlBody, undefined);
    assert.ok(r.notes!.includes(TEXT_FORM_NOTE), JSON.stringify(r.notes));
    // The image sentence's remedy is corrected to match: the shared default sends the caller
    // to `asAttachment: true`, which this tool refuses while {{forward}} is in the body.
    assert.ok(r.notes!.some((n) => n.startsWith('1 media part(s) could not be embedded')
      && n.endsWith('put {{forward}} in htmlBody to embed them, or drop the token and pass '
        + 'asAttachment: true to forward the original whole.')), JSON.stringify(r.notes));
    assert.ok(r.notes!.every((n) => !/re-run with asAttachment: true for full fidelity/.test(n)));
  });

  it('says nothing when the block ships as html', async () => {
    const { client } = spyClient(withInlineImage());
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>FYI</p>{{forward}}' },
      client,
    );
    assert.ok(r.notes!.every((n) => n !== TEXT_FORM_NOTE), JSON.stringify(r.notes));
  });

  it('says nothing when the original has no html to lose', async () => {
    const textOnly = makeOriginal({ htmlBody: undefined, bodyValues: { t: { value: 'plain original' } } });
    const { client } = spyClient(textOnly);
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], textBody: 'FYI\n{{forward}}' },
      client,
    );
    assert.ok((r.notes ?? []).every((n) => n !== TEXT_FORM_NOTE), JSON.stringify(r.notes));
  });
});

// ---------------------------------------------------------------------------
// The identity whose signature expands
// ---------------------------------------------------------------------------

describe('draft_email — {{signature}} expands the FROM identity, not the first or the default', () => {
  // Three identities, deliberately arranged so "first", "default" and "the From one" are all
  // different rows: nothing here can pass by picking a fixed index.
  const IDENTITIES = [
    { id: 'a', email: 'first@example.com', textSignature: 'FIRST SIGN-OFF' },
    { id: 'b', email: 'me@example.com', mayDelete: false, textSignature: 'DEFAULT SIGN-OFF' },
    { id: 'c', email: 'alias@example.com', textSignature: 'ALIAS SIGN-OFF' },
  ];

  it('expands the identity named by `from`', async () => {
    const { client, calls } = spyClient(makeOriginal(), { getIdentities: async () => IDENTITIES });
    await compose(
      { mode: 'new', from: 'alias@example.com', to: ['sam@example.com'], textBody: 'hi\n{{signature}}' },
      client,
    );
    assert.equal(calls.draft.textBody, 'hi\nALIAS SIGN-OFF');
    assert.equal(calls.draft.from, 'alias@example.com');
  });

  it('expands the account default when `from` is omitted, not the first row', async () => {
    const { client, calls } = spyClient(makeOriginal(), { getIdentities: async () => IDENTITIES });
    await compose({ mode: 'new', to: ['sam@example.com'], textBody: 'hi\n{{signature}}' }, client);
    assert.equal(calls.draft.textBody, 'hi\nDEFAULT SIGN-OFF');
  });

  it('names the From identity in the not-placed warning too', async () => {
    const { client } = spyClient(makeOriginal(), { getIdentities: async () => IDENTITIES });
    const r = await compose(
      { mode: 'new', from: 'alias@example.com', to: ['sam@example.com'], textBody: 'hi' }, client,
    );
    assert.ok(r.notes!.some((n) => n.startsWith('Identity alias@example.com has a signature')));
  });
});

// ---------------------------------------------------------------------------
// The receipt
// ---------------------------------------------------------------------------

describe('draft_email — the receipt', () => {
  it('is absent when the call wrote no token at all', async () => {
    const { client } = spyClient(makeOriginal(), { getIdentities: async () => [UNSIGNED_IDENTITY] });
    const r = await compose({ mode: 'new', textBody: 'plain note' }, client);
    assert.equal(r.tokens, undefined);
  });

  it('reports per part: position order, what expanded and how many times', async () => {
    // Position order, not an inferred one: the sign-off is placed BELOW the history here, and
    // the receipt says so without judging it. A token may appear only once per part, so the
    // counts are pinned across a part that carries two DIFFERENT tokens.
    const { client } = spyClient();
    const r = await compose(
      {
        mode: 'reply', originalEmailId: 'o1',
        htmlBody: '<p>hi</p><div>{{quote}}</div><div>{{signature}}</div>',
        textBody: 'hi\n{{quote}}\n{{signature}}',
      },
      client,
    );
    const html = r.tokens!.parts.find((p) => p.part === 'htmlBody')!;
    assert.deepEqual(html.order, ['quote', 'signature']);
    assert.deepEqual(html.expanded, [{ token: 'quote', count: 1 }, { token: 'signature', count: 1 }]);
    assert.deepEqual(html.removed, []);
    assert.deepEqual(r.tokens!.parts.map((p) => p.part), ['htmlBody', 'textBody']);
  });

  it("names a {{…}} spelling it left as written, so a typo does not ship braces silently", async () => {
    const { client, calls } = spyClient();
    const r = await compose({ mode: 'new', textBody: 'hi {{sig}}' }, client);
    assert.equal(calls.draft.textBody, 'hi {{sig}}');
    assert.match(r.tokens!.unexpanded!, /\{\{sig\}\}/);
  });

  it('reports the escape as consumed, with the braces shipped as text', async () => {
    const { client, calls } = spyClient();
    const r = await compose({ mode: 'new', textBody: 'write \\{{signature}} to sign' }, client);
    assert.equal(calls.draft.textBody, 'write {{signature}} to sign');
    // An escape is not an expansion and must not be reported as one.
    assert.equal(r.tokens, undefined);
  });
});
