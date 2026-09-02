import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { composeDraftEmail, sanitizeEmlFilename } from './draft-email-handler.js';
import type { DraftEmailClient } from './draft-email-handler.js';
import { InvalidInputError } from './coerce.js';
import { EMAIL_BODY_PROPERTIES } from './jmap-client.js';
import { normalizeBodies } from './body-format.js';

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

/**
 * The shape of an identifier this server mints for an image it carries out of a fetched
 * message. Production mints from the CSPRNG, so the value cannot be pinned — but the shape
 * can, and every assertion that follows a minted part through to the stored body reads the
 * value back off `createDraft` rather than predicting it.
 */
const MINTED_CID = /^ii-[0-9a-f]{32}@inline\.invalid$/;

/** The note a reply gets when it placed no {{quote}} — the default is now unquoted. */
const NOTE_REPLY_UNQUOTED =
  'This reply was stored without the original: place {{quote}} in the body to include it.';

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

  it("keeps the contentless-draft guard on mode:'new', testing bodies with isBlank", async () => {
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

  // The receipt lists spellings and counts them in the same sentence, so both have to count
  // DISTINCT ones.
  it('counts one spelling written into both bodies once', async () => {
    const { client } = spyClient();
    const r = await compose(
      { mode: 'new', textBody: 'hi {{sig}}', htmlBody: '<p>hi {{sig}}</p>' },
      client,
    );
    assert.match(r.tokens!.unexpanded!, /^"\{\{sig\}\}"$/);
  });

  // With the display cap at three, four occurrences of two spellings must not quote one of
  // them twice and then promise a further spelling that was never in the body.
  it('does not spend the display cap on a repeat, nor promise a spelling that is not there', async () => {
    const { client } = spyClient();
    const r = await compose(
      { mode: 'new', textBody: 'hi {{sig}} {{fwd}}', htmlBody: '<p>hi {{sig}} {{fwd}}</p>' },
      client,
    );
    const listed = r.tokens!.unexpanded!;
    assert.match(listed, /\{\{sig\}\}/);
    assert.match(listed, /\{\{fwd\}\}/);
    assert.equal(/and \d+ more/.test(listed), false, `promised absent spellings: ${listed}`);
    assert.equal(listed.match(/\{\{sig\}\}/g)!.length, 1, `quoted one spelling twice: ${listed}`);
  });

  it('reports the escape as consumed, with the braces shipped as text', async () => {
    const { client, calls } = spyClient();
    const r = await compose({ mode: 'new', textBody: 'write \\{{signature}} to sign' }, client);
    assert.equal(calls.draft.textBody, 'write {{signature}} to sign');
    // An escape is not an expansion and must not be reported as one.
    assert.equal(r.tokens, undefined);
  });
});

// ---------------------------------------------------------------------------
// The handoff to createDraft: what a caller supplies, and what reaches the server
// ---------------------------------------------------------------------------

const UPLOADED: any[] = [
  { blobId: 'up-1', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' },
];

/**
 * The spy above, with two differences that matter to a whole-`notes` assertion.
 *
 * The identity has NO configured sign-off, because a signed identity with no
 * `{{signature}}` in the body adds a warning of its own — correct, and noise in every
 * assertion that is about something else. And `uploadAttachments` returns a caller-supplied
 * result, so an uploaded part can be followed all the way to `createDraft`.
 */
function plainClient(
  original: any = makeOriginal(),
  over: Partial<DraftEmailClient> = {},
  uploadResult: any[] = [],
) {
  const calls: any = { gets: [] as string[] };
  const client: DraftEmailClient = {
    getEmailById: async (id) => {
      calls.gets.push(id);
      if (id !== 'draft-9') return original;
      return {
        id,
        attachments: (calls.draft?.attachments ?? []).map(
          (p: any) => ({ ...p, size: p.size ?? 70 }),
        ),
      };
    },
    getIdentities: async () => [UNSIGNED_IDENTITY],
    uploadAttachments: async (specs, dir, allowBlob, options) => {
      calls.upload = { specs, dir, allowBlob, options };
      return uploadResult;
    },
    createDraft: async (p) => { calls.draft = p; return 'draft-9'; },
    ...over,
  };
  return { client, calls };
}

/** An uploaded part as it comes back when the message really does display it. */
function inlinePart(cid: string, name = 'logo.png') {
  return { blobId: 'blob-' + cid, type: 'image/png', name, disposition: 'inline', cid };
}

describe("draft_email — mode:'new' reaches createDraft with every field it was given", () => {
  it('returns the saved id, the mode, the subject and the recipients', async () => {
    const { client, calls } = plainClient();
    const r = await compose(
      { mode: 'new', to: ['a@b.example'], cc: ['c@d.example'], subject: 'Hi', textBody: 'hello' },
      client,
    );
    assert.deepEqual(r, {
      emailId: 'draft-9', mode: 'new', subject: 'Hi', to: ['a@b.example'], cc: ['c@d.example'],
    });
    assert.equal(calls.draft.subject, 'Hi');
  });

  it('coerces a lone recipient string to an array (lenient-client input)', async () => {
    const { client, calls } = plainClient();
    await compose({ mode: 'new', to: 'a@b.example', subject: 'Hi', textBody: 'hello' }, client);
    assert.deepEqual(calls.draft.to, ['a@b.example']);
  });

  // Pin the WHOLE outgoing object, so a field dropped from the handoff fails the suite
  // rather than silently never reaching the server.
  it('passes every supported field through, and nothing beyond them', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, UPLOADED);
    await compose({
      mode: 'new',
      to: ['a@b.example'],
      cc: ['c@d.example'],
      bcc: ['e@f.example'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: [{ path: 'a.pdf' }],
    }, client, '/attach/root');
    assert.deepEqual(calls.draft, {
      to: ['a@b.example'],
      cc: ['c@d.example'],
      bcc: ['e@f.example'],
      from: 'me@example.com',
      mailbox: 'Drafts',
      subject: 'Hi',
      textBody: 'hello',
      htmlBody: '<p>hello</p>',
      inReplyTo: ['prev@example.com'],
      references: ['root@example.com', 'prev@example.com'],
      replyTo: ['reply@example.com'],
      attachments: UPLOADED,
    });
  });

  it('allows an attachment-only draft: attachments count as content', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, UPLOADED);
    const r = await compose({ mode: 'new', attachments: [{ path: 'a.pdf' }] }, client, '/attach/root');
    assert.equal(r.emailId, 'draft-9');
    assert.deepEqual(calls.draft.attachments, UPLOADED);
  });
});

describe('draft_email — uploading the attachments the caller named', () => {
  it('uploads with the given attachDir and threads the parts into the draft', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, UPLOADED);
    await compose(
      { mode: 'new', to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    // The upload is told which Content-IDs the message displays; an ordinary attachment
    // displays none, so the set is empty and the part is dispositioned as a file.
    assert.deepEqual(calls.upload, {
      specs: [{ path: 'a.pdf' }], dir: '/attach/root', allowBlob: false, options: { inlineCids: new Set() },
    });
    assert.deepEqual(calls.draft.attachments, UPLOADED);
  });

  // The other half of the flag. Every other case passes false, so a handler that dropped the
  // argument and hardcoded "off" would still satisfy them — and the only symptom would be an
  // in-account source refused on a server configured to allow it. Pinned in all three modes,
  // because one shared orchestration is exactly the shape in which a mode could lose it.
  it('passes the blob opt-in through when it is on, in every mode', async () => {
    const spec = { blobId: 'G1', name: 'a.pdf' };
    const forNew = plainClient();
    await compose({ mode: 'new', to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [spec] }, forNew.client, undefined, true);
    assert.deepEqual(forNew.calls.upload, { specs: [spec], dir: undefined, allowBlob: true, options: { inlineCids: new Set() } });

    const forReply = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi', attachments: [{ emailId: 'M1', attachmentId: 'p2' }] }, forReply.client, undefined, true);
    assert.deepEqual(forReply.calls.upload, {
      specs: [{ emailId: 'M1', attachmentId: 'p2' }], dir: undefined, allowBlob: true, options: { inlineCids: new Set() },
    });

    const forForward = plainClient();
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}', attachments: [spec] }, forForward.client, undefined, true);
    assert.deepEqual(forForward.calls.upload, { specs: [spec], dir: undefined, allowBlob: true, options: { inlineCids: new Set() } });
  });

  it('does not call uploadAttachments at all when no attachments are given, in every mode', async () => {
    const forNew = plainClient();
    await compose({ mode: 'new', to: ['a@b.example'], subject: 'Hi', textBody: 'hello' }, forNew.client, '/attach/root');
    assert.equal(forNew.calls.upload, undefined);

    const forReply = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi' }, forReply.client, '/attach/root');
    assert.equal(forReply.calls.upload, undefined);

    const forForward = plainClient();
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, forForward.client, '/attach/root');
    assert.equal(forForward.calls.upload, undefined);
  });
});

// ---------------------------------------------------------------------------
// Malformed caller bodies, refused before anything is fetched, built or stored
// ---------------------------------------------------------------------------

describe('draft_email — a malformed body is refused before the draft is created', () => {
  it('refuses a non-string body, naming the type it got, and creates nothing', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'new', subject: 'Hi', textBody: ['a', 'b'] }, client),
      (e: any) => e instanceof InvalidInputError && /textBody must be a string; received array/.test(e.message),
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses an entirely HTML-escaped htmlBody', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'new', subject: 'Hi', htmlBody: '&lt;p&gt;Hi&lt;/p&gt;' }, client),
      /htmlBody appears to be HTML-escaped/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses a CDATA-wrapped body in either part', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'new', subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>' }, client),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => compose({ mode: 'new', subject: 'Hi', textBody: '<![CDATA[Hi]]>' }, client),
      /textBody is wrapped in a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
  });

  // The history block is what makes this class of malformed body dangerous on a reply or a
  // forward: it supplies the real tags an escaped-markup body lacks and the visible content
  // an empty one lacks, so every downstream check passes and the message ships with the
  // caller's own words missing from the plain-text alternative. The guard therefore runs on
  // the caller's authored body, before any block is built or expanded into it.
  it('refuses a CDATA body on a reply, even though the quote would supply visible content', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '<![CDATA[<p>Just checking in</p>]]>{{quote}}' }, client),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', textBody: '<![CDATA[Just checking in]]>{{quote}}' }, client),
      /textBody is wrapped in a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
    // Refused from the arguments alone: the original is never even fetched.
    assert.deepEqual(calls.gets, []);
  });

  it('refuses a CDATA note on a forward, even though the forwarded block would supply content', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<![CDATA[<p>FYI</p>]]>{{forward}}' }, client),
      /htmlBody contains a CDATA section/,
    );
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: '<![CDATA[FYI]]>{{forward}}' }, client),
      /textBody is wrapped in a CDATA section/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses it on the asAttachment path too, where the note is the whole body', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true, htmlBody: '<![CDATA[<p>FYI</p>]]>' },
        client,
      ),
      /CDATA/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('refuses before uploading, so nothing reaches the blob store, in every mode', async () => {
    const forNew = plainClient();
    await assert.rejects(
      () => compose({ mode: 'new', subject: 'Hi', htmlBody: '<![CDATA[<p>Hi</p>]]>', attachments: [{ path: 'a.pdf' }] }, forNew.client, '/attach/root'),
      /CDATA/,
    );
    assert.equal(forNew.calls.upload, undefined);

    const forReply = plainClient();
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '<![CDATA[<p>hi</p>]]>', attachments: [{ path: 'a.pdf' }] }, forReply.client, '/attach/root'),
      /CDATA/,
    );
    assert.equal(forReply.calls.upload, undefined);
    assert.equal(forReply.calls.draft, undefined);

    const forForward = plainClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<![CDATA[<p>FYI</p>]]>', attachments: [{ path: 'a.pdf' }] }, forForward.client, '/attach/root'),
      /CDATA/,
    );
    assert.equal(forForward.calls.upload, undefined);
    assert.equal(forForward.calls.draft, undefined);
  });

  it('still accepts bodies that only look malformed: escaped markup in real tags, a bare ]]>', async () => {
    const { client, calls } = plainClient();
    await compose(
      {
        mode: 'new',
        subject: 'Hi',
        htmlBody: '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>',
        textBody: 'Close it with ]]> here',
      },
      client,
    );
    assert.equal(calls.draft.htmlBody, '<pre>&lt;p&gt;paragraph&lt;/p&gt;</pre>');
    assert.equal(calls.draft.textBody, 'Close it with ]]> here');
  });

  it('leaves a legitimate reply body alone: escaped markup inside real tags', async () => {
    const { client, calls } = plainClient();
    await compose(
      {
        mode: 'reply',
        originalEmailId: 'o1',
        htmlBody: '<p>Write it as <code>&lt;p&gt;</code>, and end with ]]&gt;</p>{{quote}}',
      },
      client,
    );
    assert.match(calls.draft.htmlBody, /&lt;p&gt;/);
  });
});

// ---------------------------------------------------------------------------
// Threading headers on mode:'new' — coerced, then routed
// ---------------------------------------------------------------------------

// Coercing a helper correctly is only half the guarantee; the other half is that the
// orchestration actually routes the parameter through it. An uncoerced string would reach
// JMAP as a per-character header list instead of one Message-ID.
describe("draft_email — threading headers on mode:'new'", () => {
  async function draftWith(args: object): Promise<any> {
    const { client, calls } = plainClient();
    await compose({ mode: 'new', subject: 'threading probe', ...args }, client);
    return calls.draft;
  }

  it('leaves a real array untouched', async () => {
    const params = await draftWith({ inReplyTo: ['<a@example.com>'], references: ['<a@example.com>'] });
    assert.deepEqual(params.inReplyTo, ['<a@example.com>']);
    assert.deepEqual(params.references, ['<a@example.com>']);
  });

  it('wraps a bare single Message-ID string in an array', async () => {
    const params = await draftWith({ inReplyTo: '<a@example.com>' });
    assert.deepEqual(params.inReplyTo, ['<a@example.com>']);
  });

  it('parses a JSON-stringified array', async () => {
    const params = await draftWith({ references: '["<a@example.com>", "<b@example.com>"]' });
    assert.deepEqual(params.references, ['<a@example.com>', '<b@example.com>']);
  });

  it('splits a comma-separated string', async () => {
    const params = await draftWith({ references: '<a@example.com>, <b@example.com>' });
    assert.deepEqual(params.references, ['<a@example.com>', '<b@example.com>']);
  });

  it('leaves both headers undefined when neither is supplied', async () => {
    const params = await draftWith({});
    assert.equal(params.inReplyTo, undefined);
    assert.equal(params.references, undefined);
  });
});

// ---------------------------------------------------------------------------
// The caller's OWN embedded images (cid: references they authored)
// ---------------------------------------------------------------------------

describe("draft_email — embedding the caller's own images", () => {
  const EMBED_ARGS = {
    mode: 'new',
    to: ['a@b.example'],
    subject: 'Logo',
    htmlBody: '<p>See:</p><img src="cid:logo">',
    attachments: [{ path: 'logo.png', cid: 'logo' }],
  };

  it('tells the upload which Content-ID the body displays', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, [inlinePart('logo')]);
    await compose(EMBED_ARGS, client, '/attach/root');
    assert.deepEqual(calls.upload.options, { inlineCids: new Set(['logo']) });
  });

  it('normalises the supplied spelling before matching it to the reference', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, [inlinePart('logo')]);
    await compose(
      { ...EMBED_ARGS, attachments: [{ path: 'logo.png', cid: '<cid:logo>' }] }, client, '/attach/root',
    );
    assert.deepEqual(calls.upload.options, { inlineCids: new Set(['logo']) });
  });

  it('reports what the saved draft embeds, with the size the server confirmed', async () => {
    let readBacks = 0;
    const { client } = plainClient(makeOriginal(), {
      getEmailById: async (id) => {
        readBacks += 1;
        return { id, attachments: [{ blobId: 'blob-logo', cid: 'logo', size: 214000, disposition: 'inline' }] };
      },
    }, [inlinePart('logo')]);
    const r = await compose(EMBED_ARGS, client, '/attach/root');
    assert.deepEqual(r.notes, ['This draft embeds 1 image(s) (209 KB).']);
    assert.equal(readBacks, 1);
  });

  it('does not read the draft back when nothing was embedded', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, UPLOADED);
    const r = await compose(
      { mode: 'new', to: ['a@b.example'], subject: 'Hi', textBody: 'hello', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.gets, []);
    assert.equal(r.notes, undefined);
  });

  it('says so when the saved draft could not be re-read', async () => {
    const { client } = plainClient(makeOriginal(), {
      getEmailById: async () => { throw new Error('network'); },
    }, [inlinePart('logo')]);
    const r = await compose(EMBED_ARGS, client, '/attach/root');
    assert.equal(r.emailId, 'draft-9');
    assert.ok(r.notes!.some((n) => /could not re-read it to confirm/.test(n)), JSON.stringify(r.notes));
  });

  it('says so when the saved draft is short of an image this call attached', async () => {
    const { client } = plainClient(makeOriginal(), {
      getEmailById: async (id) => ({ id, attachments: [] }),
    }, [inlinePart('logo')]);
    const r = await compose(EMBED_ARGS, client, '/attach/root');
    assert.ok(
      r.notes!.some((n) => /1 embedded image\(s\) this call attached were not found/.test(n)),
      JSON.stringify(r.notes),
    );
  });

  it('refuses a reference no attachment supplies, before anything is uploaded', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'new', to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:missing">', attachments: [{ path: 'a.pdf' }] },
        client, '/attach/root',
      ),
      (e: any) => e instanceof InvalidInputError
        && /references cid "missing" but no attachment supplies it/.test(e.message)
        && /add an attachments item with cid: "missing"/.test(e.message),
    );
    assert.equal(calls.upload, undefined);
    assert.equal(calls.draft, undefined);
  });

  it('drops the add-an-attachment repair when no attachment source is configured', async () => {
    const { client } = plainClient();
    await assert.rejects(
      () => compose({ mode: 'new', to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:missing">' }, client),
      (e: any) => e instanceof InvalidInputError
        && /neither FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is set/.test(e.message)
        && !/add an attachments item with cid/.test(e.message),
    );
  });

  it('refuses a reference to a server-managed identifier, without diagnosing a draft', async () => {
    const { client } = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'new', to: ['a@b.example'], subject: 'Hi', htmlBody: '<img src="cid:ii-' + 'a'.repeat(32) + '@inline.invalid">' },
        client, '/attach/root',
      ),
      // A compose has no draft, so the refusal must not diagnose one: it says the identifier
      // is the server's to assign, and leaves the remedy (attach your own cid) unchanged.
      (e: any) => e instanceof InvalidInputError
        && /server-managed identifier/.test(e.message)
        && !/this draft/.test(e.message)
        && /add an attachments item with a cid of your choosing/.test(e.message),
    );
  });

  it('refuses two attachments sharing one Content-ID, interpolating the real count', async () => {
    const { client, calls } = plainClient();
    await assert.rejects(
      () => compose({
        mode: 'new',
        to: ['a@b.example'],
        subject: 'Hi',
        htmlBody: '<img src="cid:logo">',
        // Two spellings of ONE identifier: the collision is judged on the canonical value.
        attachments: [{ path: 'a.png', cid: 'logo' }, { path: 'b.png', cid: '<logo>' }],
      }, client, '/attach/root'),
      (e: any) => e instanceof InvalidInputError && /2 attachments/.test(e.message) && /"logo"/.test(e.message),
    );
    assert.equal(calls.upload, undefined);
  });

  // The file is the caller's; a message that cannot display it still carries it.
  it('attaches a cid-bearing file to a text-only draft and says it was not embedded', async () => {
    const degraded = [{ blobId: 'blob-logo', type: 'image/png', name: 'logo.png', disposition: 'attachment', cid: 'logo' }];
    const { client, calls } = plainClient(makeOriginal(), {}, degraded);
    const r = await compose(
      { mode: 'new', to: ['a@b.example'], subject: 'Hi', textBody: 'no html here', attachments: [{ path: 'logo.png', cid: 'logo' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.upload.options, { inlineCids: new Set() });
    assert.deepEqual(calls.draft.attachments, degraded);
    assert.deepEqual(r.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
    assert.deepEqual(calls.gets, []);
  });

  // The second route to the same outcome: an html body DOES ship, it just displays
  // something else. The note must not claim there was no html body.
  it('says an unreferenced image was attached even when an html body ships', async () => {
    const degraded = [{ blobId: 'blob-spare', type: 'image/png', name: 'spare.png', disposition: 'attachment', cid: 'spare' }];
    const { client, calls } = plainClient(makeOriginal(), {}, degraded);
    const r = await compose({
      mode: 'new',
      to: ['a@b.example'],
      subject: 'Hi',
      htmlBody: '<p>text, and no image reference at all</p>',
      attachments: [{ path: 'spare.png', cid: 'spare' }],
    }, client, '/attach/root');
    assert.deepEqual(calls.upload.options, { inlineCids: new Set() });
    assert.deepEqual(r.notes, ['1 of your image(s) became regular attachments (nothing in the body displays them).']);
  });

  it('notes reference-shaped prose it could not act on, without refusing the message', async () => {
    const { client } = plainClient(makeOriginal(), {}, [inlinePart('logo')]);
    const r = await compose({
      ...EMBED_ARGS,
      htmlBody: '<p>Reference it as cid:whatever in the css</p><img src="cid:logo">',
    }, client, '/attach/root');
    assert.ok(
      r.notes!.some((n) => /looks like an embedded-image \(cid:\) reference/.test(n)),
      JSON.stringify(r.notes),
    );
  });

  // A lenient client may send the whole array as a JSON string; the checks must see the
  // items, not an opaque string that would make every reference look unsupplied.
  it('sees attachments supplied as a JSON string, in every mode', async () => {
    const asString = '[{"path":"chart.png","cid":"chart"}]';
    const chart = [inlinePart('chart', 'chart.png')];

    const forNew = plainClient(makeOriginal(), {}, chart);
    await compose(
      { mode: 'new', to: ['a@b.example'], subject: 'Hi', htmlBody: '<p>See:</p><img src="cid:chart">', attachments: asString },
      forNew.client, '/attach/root',
    );
    assert.deepEqual(forNew.calls.upload.options, { inlineCids: new Set(['chart']) });

    const forReply = plainClient(makeOriginal(), {}, chart);
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>See:</p><img src="cid:chart">{{quote}}', attachments: asString },
      forReply.client, '/attach/root',
    );
    assert.deepEqual(forReply.calls.upload.options, { inlineCids: new Set(['chart']) });

    const forForward = plainClient(makeOriginal(), {}, chart);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>See:</p><img src="cid:chart">{{forward}}', attachments: asString },
      forForward.client, '/attach/root',
    );
    assert.deepEqual(forForward.calls.upload.options, { inlineCids: new Set(['chart']) });
  });

  it('reports a malformed body before an attachment problem, in every mode', async () => {
    const forNew = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'new', to: ['a@b.example'], subject: 'Hi', htmlBody: '<p>ok</p><img src="cid:missing"><![CDATA[x]]>' },
        forNew.client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );

    const forReply = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p><img src="cid:missing"><![CDATA[x]]>' },
        forReply.client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );

    const forForward = plainClient();
    await assert.rejects(
      () => compose(
        { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>FYI</p><img src="cid:missing"><![CDATA[x]]>' },
        forForward.client, '/attach/root',
      ),
      /htmlBody contains a CDATA section/,
    );
  });

  // The wording differs from a fresh compose on purpose: a reply or forward already carries
  // history below the note, and its images arrive on their own rather than being authored.
  it("refuses a note reference nothing supplies in the note wording, on a reply and a forward", async () => {
    const forReply = plainClient();
    const replyMessage = await messageFrom(() => compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<img src="cid:chart">{{quote}}' }, forReply.client, '/attach/root',
    ));
    assert.match(replyMessage, /your note/);
    assert.match(replyMessage, /Quoted images appear inside the quote automatically/);
    assert.equal(forReply.calls.upload, undefined);
    assert.equal(forReply.calls.draft, undefined);

    const forForward = plainClient();
    const forwardMessage = await messageFrom(() => compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<img src="cid:chart">{{forward}}' },
      forForward.client, '/attach/root',
    ));
    assert.match(forwardMessage, /your note/);
    assert.match(forwardMessage, /Quoted images appear inside the quote automatically/);
    assert.equal(forForward.calls.draft, undefined);
  });

  it("reports the saved draft's authored embeds on a reply and a forward too", async () => {
    const chart = [inlinePart('chart', 'chart.png')];
    const readBack = async (id: string) => (
      id === 'o1'
        ? makeOriginal()
        : { id, attachments: [{ blobId: 'blob-chart', cid: 'chart', size: 2048, disposition: 'inline' }] }
    );

    const forReply = plainClient(makeOriginal(), { getEmailById: readBack }, chart);
    const replyResult = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>Compare with:</p><img src="cid:chart">{{quote}}', attachments: [{ path: 'chart.png', cid: 'chart' }] },
      forReply.client, '/attach/root',
    );
    assert.deepEqual(replyResult.notes, ['This draft embeds 1 image(s) (2 KB).']);

    const forForward = plainClient(makeOriginal(), { getEmailById: readBack }, chart);
    const forwardResult = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>Context:</p><img src="cid:chart">{{forward}}', attachments: [{ path: 'chart.png', cid: 'chart' }] },
      forForward.client, '/attach/root',
    );
    assert.deepEqual(forwardResult.notes, ['This draft embeds 1 image(s) (2 KB).']);
  });

  it('does not re-read the saved draft for an ordinary reply or forward', async () => {
    const forReply = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'hi\n{{quote}}' }, forReply.client);
    assert.deepEqual(forReply.calls.gets, ['o1']);

    const forForward = plainClient();
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, forForward.client);
    assert.deepEqual(forForward.calls.gets, ['o1']);
  });
});

// ---------------------------------------------------------------------------
// mode:'reply' — subject, recipients, threading and the recorded source
// ---------------------------------------------------------------------------

describe("draft_email — mode:'reply' subject, recipients and threading", () => {
  it('prefixes the subject with Re: and does not double-prefix one that has it', async () => {
    const plain = plainClient(makeOriginal({ subject: 'Hello' }));
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, plain.client);
    assert.equal(plain.calls.draft.subject, 'Re: Hello');

    const already = plainClient(makeOriginal({ subject: 'Re: Hello' }));
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, already.client);
    assert.equal(already.calls.draft.subject, 'Re: Hello');
  });

  it('uses a caller subject verbatim, prefixes nothing, and threads the same way', async () => {
    const { client, calls } = plainClient(makeOriginal({ subject: 'Re: Project update' }));
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'x', subject: 'New topic' }, client,
    );
    assert.equal(calls.draft.subject, 'New topic');
    assert.equal(r.subject, 'New topic');
    // The override changes only the subject: the threading chain is built the same way.
    assert.deepEqual(calls.draft.inReplyTo, ['orig-msg@example.com']);
    assert.deepEqual(calls.draft.references, ['root@example.com', 'orig-msg@example.com']);
  });

  it('treats a supplied-but-blank or null subject as omitted, on a reply and a forward', async () => {
    // '​' is a zero-width space: visually empty, so it reads as blank like the rest.
    for (const blank of ['', '   ', '​', null]) {
      const reply = plainClient();
      await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x', subject: blank }, reply.client);
      assert.equal(reply.calls.draft.subject, 'Re: Project update', `reply, subject ${JSON.stringify(blank)}`);

      const forward = plainClient();
      await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'x\n{{forward}}', subject: blank }, forward.client);
      assert.equal(forward.calls.draft.subject, 'Fwd: Project update', `forward, subject ${JSON.stringify(blank)}`);
    }
  });

  it('refuses a non-string subject rather than silently inheriting the default', async () => {
    for (const value of [42, ['a'], { s: 1 }, true]) {
      const reply = plainClient();
      await assert.rejects(
        () => compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x', subject: value }, reply.client),
        /subject must be a string/,
      );
      const forward = plainClient();
      await assert.rejects(
        () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'x\n{{forward}}', subject: value }, forward.client),
        /subject must be a string/,
      );
    }
  });

  it("defaults the recipient to the original sender, keeping the display name", async () => {
    const named = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, named.client);
    assert.deepEqual(named.calls.draft.to, ['Jon Godley <jon@example.com>']);

    const bare = plainClient(makeOriginal({ from: [{ email: 'noname@example.com' }] }));
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, bare.client);
    assert.deepEqual(bare.calls.draft.to, ['noname@example.com']);

    const explicit = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x', to: ['alice@x.example'] }, explicit.client);
    assert.deepEqual(explicit.calls.draft.to, ['alice@x.example']);
  });

  it('allows a body-less reply draft — fill it in with edit_draft before sending', async () => {
    const { client, calls } = plainClient();
    const r = await compose({ mode: 'reply', originalEmailId: 'o1' }, client);
    assert.equal(r.emailId, 'draft-9');
    assert.equal(calls.draft.textBody, undefined);
    assert.equal(calls.draft.htmlBody, undefined);
  });

  it('refuses an original with no Message-ID, and one with no determinable recipient', async () => {
    const noId = plainClient(makeOriginal({ messageId: undefined }));
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, noId.client),
      /does not have a Message-ID/,
    );
    const noRecipient = plainClient(makeOriginal({ from: [] }));
    await assert.rejects(
      () => compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, noRecipient.client),
      /Could not determine reply recipient/,
    );
  });

  it("records the original's JMAP id as sourceEmailId, and omits it when there is none", async () => {
    const withId = plainClient();
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, withId.client);
    assert.equal(withId.calls.draft.sourceEmailId, 'o1');

    const withoutId = plainClient(makeOriginal({ id: undefined }));
    await compose({ mode: 'reply', originalEmailId: 'o1', textBody: 'x' }, withoutId.client);
    assert.equal(withoutId.calls.draft.sourceEmailId, undefined);
  });
});

// ---------------------------------------------------------------------------
// The images a reply quote carries
// ---------------------------------------------------------------------------

describe('draft_email — the images a {{quote}} carries', () => {
  it('reproduces the original html inside a cite blockquote', async () => {
    const { client, calls } = plainClient();
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    assert.match(calls.draft.htmlBody, /<blockquote type="cite"[^>]*>.*original html/s);
    assert.match(calls.draft.htmlBody, /my reply/);
  });

  it('carries nothing and reports nothing when no {{quote}} is placed at all', async () => {
    const { client, calls } = plainClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>' }, client,
    );
    assert.equal('attachments' in calls.draft, false);
    // Block construction runs the collect pass, so an unquoted reply must not gain
    // dropped-image notes by doing work nobody asked for. The unquoted-reply note is the
    // only thing this call has to say.
    assert.deepEqual(r.notes, [NOTE_REPLY_UNQUOTED]);
  });

  it('mints nothing for a text-only reply, and says the quote dropped the images', async () => {
    const { client, calls } = plainClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'my reply\n{{quote}}' }, client,
    );
    assert.equal('attachments' in calls.draft, false);
    assert.match(calls.draft.textBody, /On .*wrote:\n> original text/);
    assert.deepEqual(r.notes, [
      '1 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('quotes an image-only original, which has no text of its own to quote', async () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
    });
    const { client, calls } = plainClient(original);
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    const minted = calls.draft.attachments.filter((p: any) => p.blobId === 'blob-png');
    assert.equal(minted.length, 1);
    assert.ok(calls.draft.htmlBody.includes(`src="cid:${minted[0].cid}"`), calls.draft.htmlBody);
    assert.match(calls.draft.htmlBody, /wrote:/);
  });

  // The phantom-quote row: an image-only original whose sole reference resolves to nothing
  // has no content to quote at all, so the reply must not open an attribution over an empty
  // blockquote. The token is removed and the result says why.
  it('ships no quote at all when an image-only original references a part it does not carry', async () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:gone"></div>' } },
      attachments: [],
    });
    const { client, calls } = plainClient(original);
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    assert.equal(calls.draft.htmlBody, '<p>my reply</p>');
    assert.doesNotMatch(calls.draft.htmlBody, /wrote:/);
    assert.ok(
      r.notes!.some((n) => /1 reference\(s\) matched no part and were skipped\./.test(n)),
      JSON.stringify(r.notes),
    );
    assert.deepEqual(
      r.tokens!.parts[0].removed, [{ token: 'quote', count: 1, cause: 'nothing-quotable' }],
    );
  });

  // The same original, but with an image it CAN carry and a reply that ships no html. The
  // quote has nothing to put in the plain-text alternative, so an attribution there would
  // describe a quote that is not present.
  it('gives a text-only reply to an image-only original no attribution in its text part', async () => {
    const original = withInlineImage({
      textBody: [],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
    });
    const { client, calls } = plainClient(original);
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'my reply\n{{quote}}' }, client,
    );
    assert.doesNotMatch(calls.draft.textBody, /wrote:/);
    assert.equal('attachments' in calls.draft, false);
    assert.deepEqual(
      r.tokens!.parts[0].removed,
      [{ token: 'quote', count: 1, cause: 'nothing-quotable-in-this-form' }],
    );
  });

  it('carries an SVG the original displayed, like any other embedded image', async () => {
    const svg = { partId: '4', blobId: 'blob-svg', type: 'image/svg+xml', size: 70, name: 'logo.svg', disposition: 'inline', cid: 'img-1' };
    const { client, calls } = plainClient(withInlineImage({ attachments: [svg] }));
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    assert.equal(calls.draft.attachments.length, 1);
    assert.equal(calls.draft.attachments[0].blobId, 'blob-svg');
    assert.equal(calls.draft.attachments[0].type, 'image/svg+xml');
    assert.match(calls.draft.attachments[0].cid, MINTED_CID);
  });

  // Two parts under one Content-ID name nothing unambiguously, so neither can be embedded —
  // and neither is a reference that matched nothing either, so the shortfall sentence is the
  // whole of what this call has to say.
  it('carries neither of two parts sharing one Content-ID, and counts both as a shortfall', async () => {
    const twin = { ...inlinePng, partId: '5', blobId: 'blob-png-2', name: 'pic2.png' };
    const { client, calls } = plainClient(withInlineImage({ attachments: [inlinePng, twin] }));
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    assert.equal('attachments' in calls.draft, false);
    assert.doesNotMatch(calls.draft.htmlBody, /<img/);
    assert.deepEqual(r.notes, [
      'Embedded 0 of 2 image part(s) referenced by the quote (0 KB embedded); ' +
      '2 could not be embedded and are not part of this draft.',
    ]);
  });

  it('counts a data:-URI image in the quoted body without carrying it', async () => {
    const original = withInlineImage({
      bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="data:image/png;base64,AAA="></p>' } },
      attachments: [],
    });
    const { client, calls } = plainClient(original);
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>my reply</p>{{quote}}' }, client,
    );
    assert.equal('attachments' in calls.draft, false);
    assert.ok(
      r.notes!.some((n) => /data:/.test(n) || /dropped/.test(n)),
      JSON.stringify(r.notes),
    );
  });

  it("appends the quote's images BEHIND the caller's own files rather than replacing them", async () => {
    const { client, calls } = plainClient(withInlineImage(), {}, UPLOADED);
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{quote}}', attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['up-1', 'blob-png']);
    assert.equal(calls.draft.attachments[1].disposition, 'inline');
  });

  it('re-reads the saved draft when the only images come from the quote', async () => {
    const { client, calls } = plainClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{quote}}' }, client,
    );
    assert.deepEqual(calls.gets, ['o1', 'draft-9']);
    assert.deepEqual(r.notes, ['This draft embeds 1 image(s) from the quoted message (1 KB).']);
  });

  it('says so when an image the quote carries is not on the saved draft', async () => {
    const { client } = plainClient(withInlineImage(), {
      getEmailById: async (id) => (id === 'draft-9' ? { id, attachments: [] } : withInlineImage()),
    });
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{quote}}' }, client,
    );
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the quoted message (1 KB).',
      '1 embedded image(s) this call attached were not found on the saved draft.' +
      ' Open the draft to check how it renders.',
    ]);
  });

  // A quote is re-sent from a different message, so an image referenced by a path relative to
  // the original's own origin cannot come with it. The quote still ships; the loss is said.
  it('says how many images the quote dropped for a reference form it cannot carry', async () => {
    const original = withInlineImage({
      bodyValues: {
        t: { value: 'original text' },
        h: { value: '<p><img src="cid:img-1"><img src="/logo.png"><img src="//cdn.example.com/a.png"></p>' },
      },
    });
    const { client } = plainClient(original);
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>hi</p>{{quote}}' }, client,
    );
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the quoted message (1 KB).',
      '2 image(s) in the quoted message used a reference form this server cannot carry into' +
      ' a quote and were dropped; the rest of the quote was kept.',
    ]);
  });
});

// ---------------------------------------------------------------------------
// mode:'forward' — subject, recipients and the recorded source
// ---------------------------------------------------------------------------

// The drop predicate in this file needs disposition+cid on every fetched part — pin the
// property list so a later "cleanup" of EMAIL_BODY_PROPERTIES cannot silently disarm it.
describe('EMAIL_BODY_PROPERTIES — what a forward needs fetched', () => {
  it('fetches disposition and cid, the inline-drop predicate inputs', () => {
    assert.ok((EMAIL_BODY_PROPERTIES as readonly string[]).includes('disposition'));
    assert.ok((EMAIL_BODY_PROPERTIES as readonly string[]).includes('cid'));
  });
});

describe("draft_email — mode:'forward' subject, recipients and the recorded source", () => {
  it('does not double-prefix a subject that already says it is a forward', async () => {
    for (const s of ['Fwd: Hello', 'fw: Hello', 'FWD: Hello', 'FW: Hello']) {
      const { client, calls } = plainClient(makeOriginal({ subject: s }));
      await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, client);
      assert.equal(calls.draft.subject, s);
    }
  });

  it('uses a caller subject verbatim on a forward', async () => {
    const { client, calls } = plainClient();
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}', subject: 'Custom line' }, client,
    );
    assert.equal(calls.draft.subject, 'Custom line');
    assert.equal(r.subject, 'Custom line');
  });

  it('handles an original with no subject', async () => {
    const { client, calls } = plainClient(makeOriginal({ subject: undefined }));
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, client);
    assert.equal(calls.draft.subject, 'Fwd: ');
  });

  it('requires a recipient — a forward has no default one, unlike a reply', async () => {
    const missing = plainClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', textBody: 'FYI\n{{forward}}' }, missing.client),
      /to is required for a forward; there is no default recipient/,
    );
    const empty = plainClient();
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: [], textBody: 'FYI\n{{forward}}' }, empty.client),
      /to is required for a forward/,
    );
  });

  it('splits a comma-separated to string (lenient coercion)', async () => {
    const { client, calls } = plainClient();
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: 'a@x.example, b@y.example', textBody: 'FYI\n{{forward}}' }, client,
    );
    assert.deepEqual(calls.draft.to, ['a@x.example', 'b@y.example']);
  });

  it("records the original's Message-ID, and omits a malformed one rather than failing the create", async () => {
    const good = plainClient();
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, good.client);
    assert.deepEqual(good.calls.draft.forwardedMessageId, ['orig-msg@example.com']);

    const absent = plainClient(makeOriginal({ messageId: undefined }));
    await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, absent.client);
    assert.equal(absent.calls.draft.forwardedMessageId, undefined);
    // The id still names WHICH stored copy was forwarded; only the header marking is lost.
    assert.equal(absent.calls.draft.sourceEmailId, 'o1');

    // Fastmail rejects or mangles each of these on Email/set, so they are treated as absent.
    for (const value of [
      'has space@example.com',
      'angle<bracket@example.com',
      'angle>bracket@example.com',
      'non-ascii-käse@example.com',
      'ctrlbell@example.com',
      'x'.repeat(999) + '@example.com',
      '',
    ]) {
      const { client, calls } = plainClient(makeOriginal({ messageId: [value] }));
      await compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'FYI\n{{forward}}' }, client);
      assert.equal(
        calls.draft.forwardedMessageId, undefined, `should omit: ${JSON.stringify(value.slice(0, 40))}`,
      );
    }
  });

  it('records both the Message-ID and the source id on an asAttachment forward too', async () => {
    const { client, calls } = plainClient();
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true }, client,
    );
    assert.deepEqual(calls.draft.forwardedMessageId, ['orig-msg@example.com']);
    assert.equal(calls.draft.sourceEmailId, 'o1');
  });

  it('leaves both caller bodies alone on asAttachment, with no forwarded block added', async () => {
    const { client, calls } = plainClient();
    await compose({
      mode: 'forward',
      originalEmailId: 'o1',
      to: ['x@y.example'],
      asAttachment: true,
      textBody: 'see attached',
      htmlBody: '<p>see attached</p>',
    }, client);
    assert.equal(calls.draft.textBody, 'see attached');
    assert.equal(calls.draft.htmlBody, '<p>see attached</p>');
    assert.doesNotMatch(calls.draft.htmlBody, /----- Original message -----/);
  });

  it('refuses an asAttachment forward of an original with no blobId', async () => {
    const { client, calls } = plainClient(makeOriginal({ blobId: undefined }));
    await assert.rejects(
      () => compose({ mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true }, client),
      /no blobId/,
    );
    assert.equal(calls.draft, undefined);
  });

  it('appends caller uploads behind the .eml on an asAttachment forward', async () => {
    const { client, calls } = plainClient(makeOriginal(), {}, UPLOADED);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], asAttachment: true, attachments: [{ path: 'a.pdf' }] },
      client, '/attach/root',
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-orig', 'up-1']);
    assert.equal(calls.draft.attachments[0].type, 'message/rfc822');
  });
});

// ---------------------------------------------------------------------------
// What a {{forward}} carries out of the original
// ---------------------------------------------------------------------------

describe('draft_email — what a {{forward}} carries out of the original', () => {
  const FWD_DOC = {
    partId: '3', blobId: 'blob-doc', type: 'application/pdf', size: 1234, name: 'doc.pdf',
    disposition: 'attachment', cid: null,
  };
  const FWD_PNG = {
    partId: '4', blobId: 'blob-png', type: 'image/png', size: 70, name: 'p.png',
    disposition: 'inline', cid: 'img-1',
  };
  /** An original whose html displays the embedded image, so the block can carry it. */
  const referencing = (over: any = {}) => makeOriginal({
    bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:img-1"></p>' } },
    attachments: [FWD_PNG],
    ...over,
  });
  const HTML_FORWARD = { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], htmlBody: '<p>FYI</p>{{forward}}' };

  it('carries a non-inline original through the whitelist, never its server-set partId/size', async () => {
    const { client, calls } = plainClient(makeOriginal({ attachments: [FWD_DOC] }));
    await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
    ]);
  });

  it('says nothing at all about an ordinary file riding along with the forward', async () => {
    const { client } = plainClient(makeOriginal({ attachments: [FWD_DOC] }));
    const r = await compose(HTML_FORWARD, client);
    assert.equal(r.notes, undefined);
  });

  it('omits attachments entirely when the original carries none', async () => {
    const { client, calls } = plainClient(makeOriginal({ attachments: [] }));
    await compose(HTML_FORWARD, client);
    assert.equal('attachments' in calls.draft, false);
  });

  it('carries a body-referenced image under a minted identifier and points the block at it', async () => {
    const { client, calls } = plainClient(referencing());
    const r = await compose(HTML_FORWARD, client);
    assert.equal(calls.draft.attachments.length, 1);
    const carriedPart = calls.draft.attachments[0];
    assert.equal(carriedPart.blobId, 'blob-png');
    assert.equal(carriedPart.disposition, 'inline');
    assert.match(carriedPart.cid, MINTED_CID);
    assert.ok(calls.draft.htmlBody.includes(`src="cid:${carriedPart.cid}"`), calls.draft.htmlBody);
    assert.doesNotMatch(calls.draft.htmlBody, /cid:img-1/);
    assert.deepEqual(r.notes, ['This draft embeds 1 image(s) from the original (1 KB).']);
  });

  it('re-reads the saved draft when the only images come from the original', async () => {
    const { client, calls } = plainClient(referencing());
    await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.gets, ['o1', 'draft-9']);
  });

  it('says so when a carried image is not on the saved draft', async () => {
    const { client } = plainClient(referencing(), {
      getEmailById: async (id) => (id === 'draft-9' ? { id, attachments: [] } : referencing()),
    });
    const r = await compose(HTML_FORWARD, client);
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the original (1 KB).',
      '1 embedded image(s) this call attached were not found on the saved draft.' +
      ' Open the draft to check how it renders.',
    ]);
  });

  it("pools an inline image the original's body never references, with its Content-ID stripped", async () => {
    const { client, calls } = plainClient(makeOriginal({ attachments: [FWD_DOC, FWD_PNG] }));
    const r = await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-doc', type: 'application/pdf', name: 'doc.pdf', disposition: 'attachment' },
      { blobId: 'blob-png', type: 'image/png', name: 'p.png', disposition: 'attachment' },
    ]);
    assert.deepEqual(r.notes, [
      '1 media part(s) could not be embedded and were attached as regular attachments: "p.png"' +
      ' — re-run with asAttachment: true for full fidelity, then delete this draft.',
    ]);
  });

  // A part carrying NO disposition, listed only under attachments, is the ordinary shape a
  // body-embedded image arrives in — so what makes it body content is the body referencing
  // it, not its metadata.
  it('pools a referenced image the block could not display, even with no disposition on it', async () => {
    const twins = [
      { partId: '7', blobId: 'blob-a', type: 'image/png', size: 20, name: 'a.png', cid: 'a@x' },
      { partId: '8', blobId: 'blob-b', type: 'image/png', size: 30, name: 'b.png', cid: 'a@x' },
    ];
    const original = makeOriginal({
      bodyValues: { t: { value: 'original text' }, h: { value: '<p>hi <img src="cid:a@x"></p>' } },
      attachments: twins,
    });
    const { client } = plainClient(original);
    const r = await compose(HTML_FORWARD, client);
    assert.deepEqual(r.notes, [
      '2 media part(s) could not be embedded and were attached as regular attachments:' +
      ' "a.png", "b.png" — re-run with asAttachment: true for full fidelity, then delete this draft.',
    ]);
  });

  it('never carries a foreign Content-ID of the shape this server mints', async () => {
    const forged = {
      partId: '9', blobId: 'blob-forged', type: 'image/png', name: 'f.png', disposition: 'inline',
      cid: 'ii-0123456789abcdef0123456789abcdef@inline.invalid',
    };
    const { client, calls } = plainClient(makeOriginal({ attachments: [forged] }));
    await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-forged', type: 'image/png', name: 'f.png', disposition: 'attachment' },
    ]);
  });

  it("normalises inline-WITHOUT-cid to disposition 'attachment' and counts it as a file", async () => {
    const inlineNoCid = {
      partId: '5', blobId: 'blob-odd', type: 'application/zip', size: 10, name: 'odd.zip',
      disposition: 'inline', cid: null,
    };
    const { client, calls } = plainClient(makeOriginal({ attachments: [inlineNoCid] }));
    const r = await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-odd', type: 'application/zip', name: 'odd.zip', disposition: 'attachment' },
    ]);
    // Marked inline but with no Content-ID, so nothing could have displayed it: an ordinary
    // file riding along, not a media part the block failed to show.
    assert.equal(r.notes, undefined);
  });

  it('carries the body-referenced image past includeOriginalAttachments:false, and leaves the files', async () => {
    const { client, calls } = plainClient(referencing({ attachments: [FWD_PNG, FWD_DOC] }));
    const r = await compose({ ...HTML_FORWARD, includeOriginalAttachments: false }, client);
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-png']);
    assert.match(calls.draft.attachments[0].cid, MINTED_CID);
    assert.deepEqual(r.notes, [
      'This draft embeds 1 image(s) from the original (1 KB).',
      '1 attachment(s), including 0 image(s), were not included because includeOriginalAttachments is false.' +
      ' Body-embedded images were still carried — they are part of the message body.',
    ]);
  });

  it('omits the carried-anyway sentence when the forward embedded nothing', async () => {
    const { client } = plainClient(makeOriginal({ attachments: [FWD_DOC] }));
    const r = await compose({ ...HTML_FORWARD, includeOriginalAttachments: false }, client);
    assert.deepEqual(r.notes, [
      '1 attachment(s), including 0 image(s), were not included because includeOriginalAttachments is false.',
    ]);
  });

  // The force-carry is bounded to parts the sender declared an image. A body can point an
  // <img> at anything; the bound is what stops a reference dragging an arbitrary file past
  // includeOriginalAttachments:false.
  it('does not force-carry a referenced part the sender did not declare an image', async () => {
    const doc = {
      partId: '7', blobId: 'blob-doc2', type: 'application/pdf', name: 'sneaky.pdf',
      disposition: 'inline', cid: 'img-1',
    };
    const { client, calls } = plainClient(referencing({ attachments: [doc] }));
    await compose({ ...HTML_FORWARD, includeOriginalAttachments: false }, client);
    assert.equal('attachments' in calls.draft, false);
  });

  it('unions in an embedded image the server routed into a body list rather than attachments', async () => {
    const original = referencing({
      attachments: [],
      htmlBody: [{ partId: 'h', type: 'text/html' }, FWD_PNG],
    });
    const { client, calls } = plainClient(original);
    await compose(HTML_FORWARD, client);
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-png']);
  });

  it('reports a reference in the original that matched no part', async () => {
    const { client } = plainClient(referencing({ attachments: [] }));
    const r = await compose(HTML_FORWARD, client);
    assert.deepEqual(r.notes, [
      "1 image reference(s) in the original's body had no matching part; nothing was carried for them.",
    ]);
  });

  it('carries an SVG the original displayed, like any other embedded image', async () => {
    const svg = {
      partId: '6', blobId: 'blob-svg', type: 'image/svg+xml', size: 70, name: 'logo.svg',
      disposition: 'inline', cid: 'img-1',
    };
    const { client, calls } = plainClient(referencing({ attachments: [svg] }));
    await compose(HTML_FORWARD, client);
    assert.equal(calls.draft.attachments.length, 1);
    assert.equal(calls.draft.attachments[0].blobId, 'blob-svg');
    assert.equal(calls.draft.attachments[0].type, 'image/svg+xml');
    assert.match(calls.draft.attachments[0].cid, MINTED_CID);
  });

  it('appends caller uploads BEHIND the originals the forward carried', async () => {
    const { client, calls } = plainClient(makeOriginal({ attachments: [FWD_DOC] }), {}, UPLOADED);
    await compose(
      { ...HTML_FORWARD, attachments: [{ path: 'a.pdf' }] }, client, '/attach/root',
    );
    assert.deepEqual(calls.draft.attachments.map((a: any) => a.blobId), ['blob-doc', 'up-1']);
  });

  it('mints nothing for a {{forward}} that ships in the text form, and pools the image', async () => {
    const { client, calls } = plainClient(referencing());
    const r = await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['x@y.example'], textBody: 'see below\n{{forward}}' },
      client,
    );
    assert.equal(calls.draft.htmlBody, undefined);
    assert.deepEqual(calls.draft.attachments, [
      { blobId: 'blob-png', type: 'image/png', name: 'p.png', disposition: 'attachment' },
    ]);
    assert.ok(
      r.notes!.some((n) => n.startsWith('1 media part(s) could not be embedded')),
      JSON.stringify(r.notes),
    );
  });

  it('embeds the image of an original whose only content IS an embedded image', async () => {
    const original = makeOriginal({
      textBody: [],
      htmlBody: [{ partId: 'h', type: 'text/html' }],
      bodyValues: { h: { value: '<div><img src="cid:img-1"></div>' } },
      attachments: [FWD_PNG],
    });
    const { client, calls } = plainClient(original);
    await compose(HTML_FORWARD, client);
    const carriedPart = calls.draft.attachments[0];
    assert.match(carriedPart.cid, MINTED_CID);
    assert.ok(calls.draft.htmlBody.includes(`src="cid:${carriedPart.cid}"`), calls.draft.htmlBody);
  });
});

// ---------------------------------------------------------------------------
// The .eml filename a forward derives from the original's subject
// ---------------------------------------------------------------------------

// The value is consumed as a SAVE NAME by receiving clients, and it comes verbatim from an
// attacker-controlled subject line.
describe('sanitizeEmlFilename', () => {
  it('keeps an ordinary subject', () => {
    assert.equal(sanitizeEmlFilename('Quarterly report'), 'Quarterly report.eml');
  });
  it('falls back on a blank/absent subject', () => {
    assert.equal(sanitizeEmlFilename(''), 'forwarded-message.eml');
    assert.equal(sanitizeEmlFilename('   '), 'forwarded-message.eml');
    assert.equal(sanitizeEmlFilename(undefined), 'forwarded-message.eml');
  });
  it('neutralizes traversal-shaped subjects (path separators, leading dots, colon)', () => {
    assert.equal(sanitizeEmlFilename('..\\..\\Startup\\x'), '_.._Startup_x.eml');
    assert.equal(sanitizeEmlFilename('../x'), '_x.eml');
    assert.equal(sanitizeEmlFilename('C:autorun'), 'C_autorun.eml');
  });
  it('strips control and format/bidi characters', () => {
    // U+202E (RLO) and U+0085 (NEL) are built from char codes so no raw invisible
    // character lives in this source file.
    const rlo = String.fromCharCode(0x202e);
    const nel = String.fromCharCode(0x85);
    assert.equal(sanitizeEmlFilename(`abc${rlo}lme.gpj`), 'abclme.gpj.eml');
    assert.equal(sanitizeEmlFilename(`ab${nel}cd`), 'abcd.eml');
  });
  it('caps the length before appending .eml', () => {
    const name = sanitizeEmlFilename('x'.repeat(500));
    assert.ok(name.length <= 84);
    assert.ok(name.endsWith('.eml'));
  });
  it('caps by code points — never splits a surrogate pair at the boundary', () => {
    // 79 ASCII chars then an astral char: a UTF-16-unit slice(0,80) would cut the
    // pair in half and leave a lone surrogate (invalid on the wire).
    const name = sanitizeEmlFilename('x'.repeat(79) + '\u{1F600}' + 'tail');
    assert.equal(name, 'x'.repeat(79) + '\u{1F600}.eml');
    assert.ok([...name].every((c) => !(c.length === 1 && c.charCodeAt(0) >= 0xd800 && c.charCodeAt(0) <= 0xdfff)));
  });
});

// ---------------------------------------------------------------------------
// What a quoted or forwarded block may carry INTO THE STORED MESSAGE
// ---------------------------------------------------------------------------
//
// These nine ask one question: can a construct out of a fetched message reach the recipient?
// That question is about the message this call STORES, so every assertion below reads
// `calls.draft` — the body handed to createDraft, after the block was built, after the single
// expansion pass joined it into the caller's own body.
//
// They are deliberately not written against buildQuoteBlocks / buildForwardBlocks, which is
// where the sanitising actually happens and where the same assertion text would compile and
// pass. A block-level pin answers "is the construct absent from the block", which is strictly
// narrower — it says nothing about what the join produces, and the difference is invisible in
// the assertion. The surface has to be the stored one.

describe('draft_email — the quote sanitiser, over the stored body', () => {
  // A reply whose html part carries the token, so the html quote is the one that ships.
  const storedHtmlQuoting = async (originalHtml: string) => {
    const original = makeOriginal({ bodyValues: { t: { value: 'x' }, h: { value: originalHtml } } });
    const { client, calls } = plainClient(original);
    await compose({ mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>r</p>{{quote}}' }, client);
    return calls.draft.htmlBody as string;
  };

  it('stores a full document as its body content only', async () => {
    const out = await storedHtmlQuoting(
      '<!DOCTYPE html><html><head><style>p{color:red}</style></head><body><p>kept</p></body></html>',
    );
    assert.match(out, /<p>kept<\/p>/);
    assert.doesNotMatch(out, /<style>|<head>|DOCTYPE|color:red/i);
  });

  it('stores no script and no event handler, and keeps the formatting', async () => {
    const out = await storedHtmlQuoting(
      '<p onclick="evil()">hi <b>bold</b> <a href="http://x.com">link</a></p><script>steal()</script>',
    );
    assert.doesNotMatch(out, /onclick|script|steal/i);
    assert.match(out, /<b>bold<\/b>/);
    assert.match(out, /<a href="http:\/\/x\.com">link<\/a>/);
  });

  it('stores no style attribute on a kept tag (no global "*" allowance)', async () => {
    const out = await storedHtmlQuoting('<p style="background:url(evil)">x</p>');
    assert.doesNotMatch(out, /style="background/);
    assert.match(out, /<p>x<\/p>/);
  });

  it('stores a real http(s) image but no cid:/data: image (no broken placeholder)', async () => {
    const out = await storedHtmlQuoting(
      '<p><img src="https://cdn/x.png" alt="real"> <img src="cid:logo@x"> <img src="data:image/png;base64,AAAA"></p>',
    );
    assert.match(out, /<img src="https:\/\/cdn\/x\.png" alt="real"/);
    assert.doesNotMatch(out, /cid:/);
    assert.doesNotMatch(out, /data:image/);
  });

  it('stores a no-<body> fragment intact (no regex extraction)', async () => {
    assert.match(await storedHtmlQuoting('<p>hi</p>'), /<p>hi<\/p>/);
  });
});

// A construct that swallows the REST OF THE DOCUMENT is a different question from the ones
// above, and none of them can ask it: every case up there puts `{{quote}}` last, and a
// construct that eats everything after itself eats nothing when there is nothing after it.
// Here the token sits ABOVE the caller's prose, so a swallower surviving the sanitiser has
// the caller's own words in its mouth — the message would ship with the prose invisible in
// a recipient's client, and nothing in the stored body would look wrong.
//
// This is the property the quote design rests on: the sanitiser's ALLOWLIST, not a
// blocklist. None of these six tags is in QUOTE_ALLOWED_TAGS (src/inline-images.ts), and
// these cases exist so that stays true — a future widening that admits one of them turns
// red here rather than in someone's inbox.
//
// The construct comes out of the FETCHED ORIGINAL, never the caller's body. The caller's
// text is checked by assertBodyInputs, which rejects several of these outright; the fetched
// original is deliberately not put through that gate, so the sanitiser is the only thing
// standing between it and the recipient.

describe('draft_email — a swallowing construct out of a fetched original cannot eat the caller\'s prose', () => {
  // Both parts are supplied, and the token leads each of them, so the stored draft carries a
  // text part to assert on as well as an html one. `compose` refuses a token placed in one
  // supplied part and not the other, and it derives no text part when only html is given.
  const storedWithQuoteAbovePros = async (originalHtml: string) => {
    const original = makeOriginal({
      bodyValues: { t: { value: 'original text' }, h: { value: originalHtml } },
    });
    const { client, calls } = plainClient(original);
    await compose(
      {
        mode: 'reply',
        originalEmailId: 'o1',
        htmlBody: '{{quote}}<p>PROSE-KEPT</p>',
        textBody: '{{quote}}\nPROSE-KEPT',
      },
      client,
    );
    return { html: calls.draft.htmlBody as string, text: calls.draft.textBody as string };
  };

  // Each is left UNCLOSED where the construct has a closing form, because that is the shape
  // that swallows: a well-formed pair only eats its own contents.
  const SWALLOWERS: [name: string, originalHtml: string, absent: RegExp][] = [
    ['a CDATA section', '<p>quoted</p><![CDATA[', /CDATA/i],
    ['an unclosed <style>', '<p>quoted</p><style>body{}', /<\/?style/i],
    ['an unclosed comment', '<p>quoted</p><!-- still going', /<!--/],
    ['a <script>', '<p>quoted</p><script>steal()', /<\/?script|steal/i],
    ['a <textarea>', '<p>quoted</p><textarea>', /<\/?textarea/i],
    ['an <xmp>', '<p>quoted</p><xmp>', /<\/?xmp/i],
  ];

  for (const [name, originalHtml, absent] of SWALLOWERS) {
    it(`drops ${name} and keeps the prose below the quote`, async () => {
      const { html, text } = await storedWithQuoteAbovePros(originalHtml);
      // The original's body really did reach the stored draft. Without this the case passes
      // when the quote never lands at all — the construct is absent and the prose is intact
      // because nothing was quoted, which is the vacuous version of this test.
      assert.match(html, /quoted/, `the quote block never reached the stored body: ${html}`);
      // Both halves, because either alone passes for the wrong reason: a construct that
      // vanishes while taking the prose with it satisfies the first, and prose that survives
      // beside a surviving swallower satisfies the second.
      assert.doesNotMatch(html, absent, `swallowing construct survived: ${html}`);
      assert.match(html, /PROSE-KEPT/, `the caller's prose was lost: ${html}`);
      assert.match(text, /PROSE-KEPT/, `the caller's prose was lost from the text part: ${text}`);

      // THE DERIVED TEXT PART, which is where a surviving swallower would do its damage: the
      // html→text conversion is the step that would read one and eat everything after it.
      // The supplied part above cannot answer this — it never passes through the converter.
      //
      // Through `normalizeBodies`, the seam createDraft uses (shapeBodies at
      // jmap-client.ts:2354 calls it, and createDraft calls that), so the conversion mode is
      // production's rather than this test's — `htmlToText` takes a mode argument, and
      // calling it directly here would let the test pick behaviour the shipping path does
      // not use. Html alone, because that is the shape that derives: given both parts
      // normalizeBodies passes them straight through and converts nothing.
      const derived = normalizeBodies({ htmlBody: html });
      assert.match(
        derived.textBody ?? '',
        /PROSE-KEPT/,
        `the derived text part lost the caller's prose: ${JSON.stringify(derived)}`,
      );
    });
  }
});

describe('draft_email — hostile fields out of a forwarded message, over the stored body', () => {
  it('stores every header field escaped in the html block (subject + display names)', async () => {
    const original = makeOriginal({
      subject: '<img src=x onerror=alert(1)>',
      from: [{ name: '"><script>alert(2)</script>', email: 'evil@example.com' }],
    });
    const { client, calls } = plainClient(original);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>n</p>{{forward}}' },
      client,
    );
    assert.doesNotMatch(calls.draft.htmlBody, /<img src=x onerror/);
    assert.doesNotMatch(calls.draft.htmlBody, /<script>/);
    assert.match(calls.draft.htmlBody, /&lt;img src=x onerror=alert\(1\)&gt;/);
  });

  it('stores text header fields with line-splitting whitespace collapsed — incl. U+2028/U+2029/U+0085', async () => {
    // Invisible characters are built from char codes so no raw control char lives in this
    // source file. NEL is NOT in ECMAScript's \s, so the assertion checks ACTUAL collapse.
    const nel = String.fromCharCode(0x85);
    const ls = String.fromCharCode(0x2028);
    const ps = String.fromCharCode(0x2029);
    const original = makeOriginal({
      from: [{ name: 'Fake' + nel + 'To: victim@example.com' + ls + 'X:' + ps + 'Y', email: 'evil@example.com' }],
      subject: 'line1\r\nline2',
    });
    const { client, calls } = plainClient(original);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], textBody: 'n\n{{forward}}' },
      client,
    );
    const lines = (calls.draft.textBody as string).split('\n');
    const fromLine = lines.find((l) => l.startsWith('From: '))!;
    assert.equal(fromLine.includes(nel), false);
    assert.equal(fromLine.includes(ls), false);
    assert.equal(fromLine.includes(ps), false);
    assert.match(fromLine, /^From: Fake To: victim@example\.com X: Y <evil@example\.com>$/);
    assert.equal(lines.find((l) => l.startsWith('Subject: ')), 'Subject: line1 line2');
  });

  it('stores the reproduced original html sanitised (script stripped, http img kept, cid img dropped)', async () => {
    const original = makeOriginal({
      bodyValues: {
        t: { value: 'x' },
        h: { value: '<p onclick="x()">hi</p><script>evil()</script><img src="https://x.com/a.png"><img src="cid:img1">' },
      },
    });
    const { client, calls } = plainClient(original);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>note</p>{{forward}}' },
      client,
    );
    assert.doesNotMatch(calls.draft.htmlBody, /<script|onclick/);
    assert.match(calls.draft.htmlBody, /<img src="https:\/\/x\.com\/a\.png" \/>/);
    assert.doesNotMatch(calls.draft.htmlBody, /cid:img1/);
  });

  it('stores a forwarded REPLY with the nested cite attribute stripped and its content kept', async () => {
    // The forwarded original is itself a reply, so its html carries a quote blockquote of its
    // own. The sanitiser drops the attribute that marks it as quoted history while keeping
    // what it said — the forwarded block this call writes is the only cited region.
    const original = makeOriginal({
      bodyValues: { t: { value: 'x' }, h: { value: '<p>top</p><blockquote type="cite">quoted</blockquote>' } },
    });
    const { client, calls } = plainClient(original);
    await compose(
      { mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'], htmlBody: '<p>n</p>{{forward}}' },
      client,
    );
    assert.match(calls.draft.htmlBody, /<blockquote>quoted<\/blockquote>/);
    assert.doesNotMatch(calls.draft.htmlBody, /<blockquote type="cite"/);
  });
});

// ---------------------------------------------------------------------------
// "Does this message ship html" is NON-BLANK, not merely supplied
// ---------------------------------------------------------------------------
//
// One test decides whether the quoted images are MINTED (`historyHtmlShips`, which the block
// builders take as `htmlShips`) and one decides whether the part carrying them SHIPS
// (`messageShipsHtml`, and buildBodyParts downstream). Both are `!isBlank`. These pins hold
// them together: a regression that moves one and not the other stores a quote whose images
// have gone, with nothing in the output saying so.
//
// A blank html part beside a token-bearing text part cannot reach any of that, because the
// symmetry rule refuses the call first — so the first test here pins the REFUSAL for those
// spellings rather than an outcome no caller can produce. That the blank part is dropped
// downstream anyway is not an exemption from the rule, and the exemption is the plausible
// regression: it would store a message whose two parts say different things.

/** A part did not ship when it is absent or blank — buildBodyParts drops both. */
const isBlankBody = (v: unknown) => v === undefined || (typeof v === 'string' && v.trim() === '');

describe('draft_email — a blank html part ships no html, and mints nothing for one', () => {
  it('refuses a token-bearing text part beside a blank html part, in either spelling', async () => {
    for (const htmlBody of ['', '   ']) {
      const { client, calls } = plainClient(withInlineImage());
      await assert.rejects(
        () => compose(
          { mode: 'reply', originalEmailId: 'o1', htmlBody, textBody: 'r\n{{quote}}' }, client,
        ),
        /\{\{quote\}\} is in textBody but not in htmlBody\./,
        JSON.stringify(htmlBody),
      );
      assert.equal(calls.draft, undefined);
    }
  });

  it('refuses the same shape on a forward, so neither history token gets the exemption', async () => {
    const { client, calls } = plainClient(withInlineImage());
    await assert.rejects(
      () => compose(
        {
          mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'],
          htmlBody: '   ', textBody: 'note\n{{forward}}',
        },
        client,
      ),
      /\{\{forward\}\} is in textBody but not in htmlBody\./,
    );
    assert.equal(calls.draft, undefined);
  });

  it('ships no html part, and no minted image, for a reply that supplies only a text part', async () => {
    // The reachable spelling of "this message ships no html": the part is absent, not blank.
    const { client, calls } = plainClient(withInlineImage());
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'r\n{{quote}}' }, client,
    );
    assert.ok(isBlankBody(calls.draft.htmlBody), String(calls.draft.htmlBody));
    assert.equal('attachments' in calls.draft, false);
    // And the images went with it — reported, not silently absent.
    assert.deepEqual(r.notes, [
      '1 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('ships html — and mints — for a body that is blank-looking MARKUP but not blank text', async () => {
    // The test is `isBlank` (what buildBodyParts drops), not "has visible content". A
    // `<div> </div>` renders as nothing but is a real html part, so it ships, and the quote
    // it carries is the html one.
    const { client, calls } = plainClient(withInlineImage());
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<div> </div>{{quote}}' }, client,
    );
    assert.ok((calls.draft.htmlBody as string).startsWith('<div> </div>'), calls.draft.htmlBody);
    const minted = calls.draft.attachments.filter((p: any) => MINTED_CID.test(p.cid));
    assert.equal(minted.length, 1);
    assert.ok((calls.draft.htmlBody as string).includes(`src="cid:${minted[0].cid}"`), calls.draft.htmlBody);
  });

  it('mints nothing for a forward whose note is text only', async () => {
    const { client, calls } = plainClient(withInlineImage());
    await compose(
      {
        mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'],
        textBody: 'note\n{{forward}}',
      },
      client,
    );
    assert.ok(isBlankBody(calls.draft.htmlBody), String(calls.draft.htmlBody));
    // The forward still carries the original's image as an ordinary attachment — it is a file
    // the caller is forwarding — but nothing was minted for a block that ships no html.
    assert.deepEqual(
      (calls.draft.attachments ?? []).filter((p: any) => MINTED_CID.test(p.cid ?? '')),
      [],
    );
  });
});

// ---------------------------------------------------------------------------
// {{signature}} on the branches that carry no history
// ---------------------------------------------------------------------------

describe('draft_email — {{signature}} does not depend on the history landing', () => {
  it('signs a reply that placed no {{quote}} at all', async () => {
    const { client, calls } = spyClient();
    const r = await compose(
      { mode: 'reply', originalEmailId: 'o1', textBody: 'r\n{{signature}}' }, client,
    );
    assert.equal(calls.draft.textBody, 'r\nKind regards,\nTest User');
    assert.deepEqual(r.notes, [NOTE_REPLY_UNQUOTED]);
  });

  it('signs a reply whose {{quote}} found nothing quotable', async () => {
    // The quote token is removed and reported; the sign-off is a separate expansion and lands.
    const barren = makeOriginal({ textBody: [], htmlBody: [], bodyValues: {} });
    const { client, calls } = spyClient(barren);
    await compose(
      { mode: 'reply', originalEmailId: 'o1', htmlBody: '<p>r</p>{{quote}}{{signature}}' }, client,
    );
    assert.doesNotMatch(calls.draft.htmlBody, /wrote:/);
    assert.match(calls.draft.htmlBody, /Kind regards,/);
  });

  it("signs the caller's own note on an asAttachment forward", async () => {
    const { client, calls } = spyClient();
    await compose(
      {
        mode: 'forward', originalEmailId: 'o1', to: ['sam@example.com'],
        asAttachment: true, textBody: 'FYI\n{{signature}}',
      },
      client,
    );
    assert.equal(calls.draft.textBody, 'FYI\nKind regards,\nTest User');
  });

  it('resolves the identity BEFORE any attachment is uploaded', async () => {
    // The sign-off has to be in the body the upload plan is read against, so the order is
    // load-bearing rather than incidental.
    const order: string[] = [];
    const { client } = spyClient(makeOriginal(), {
      getIdentities: async () => { order.push('identities'); return [SIGNED_IDENTITY]; },
      uploadAttachments: async () => { order.push('upload'); return []; },
    });
    await compose(
      { mode: 'new', to: ['sam@example.com'], textBody: 'hi\n{{signature}}', attachments: [{ path: 'a.pdf' }] },
      client,
      '/tmp/attach',
    );
    assert.deepEqual(order, ['identities', 'upload']);
  });
});
