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
        /which is not a token.*exact spellings are \{\{signature\}\}, \{\{quote\}\} and \{\{forward\}\}/s,
        miss,
      );
    }
    assert.equal(calls.draft, undefined);
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
// The receipt
// ---------------------------------------------------------------------------

describe('draft_email — the receipt', () => {
  it('is absent when the call wrote no token at all', async () => {
    const { client } = spyClient(makeOriginal(), { getIdentities: async () => [UNSIGNED_IDENTITY] });
    const r = await compose({ mode: 'new', textBody: 'plain note' }, client);
    assert.equal(r.tokens, undefined);
  });

  it('reports per part: position order, what expanded and how many times', async () => {
    const { client } = spyClient();
    const r = await compose(
      {
        mode: 'reply', originalEmailId: 'o1',
        htmlBody: '<p>hi</p><div>{{signature}}</div><div>{{signature}}</div><div>{{quote}}</div>',
        textBody: 'hi\n{{signature}}\n{{signature}}\n{{quote}}',
      },
      client,
    );
    const html = r.tokens!.parts.find((p) => p.part === 'htmlBody')!;
    assert.deepEqual(html.order, ['signature', 'signature', 'quote']);
    assert.deepEqual(html.expanded, [{ token: 'signature', count: 2 }, { token: 'quote', count: 1 }]);
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
