import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  DEFAULT_INLINE_CONTEXT, planAuthoredInlineImages, recordQuoteImages,
  reportAuthoredInlineImages,
} from './compose-inline.js';
import { InvalidInputError } from './coerce.js';
import {
  InlineNoteLedger, noteEmbedMissingAfterSave, noteEmbedUnconfirmed, noteUnparsableCidText,
  rejectCidCollisionInCall, rejectDanglingCidRef, rejectNoteCidRef, rejectReservedCidRef,
} from './inline-notes.js';
import type { QuoteImageOutcome } from './reply-quote.js';

// The three exported functions here are shared by every compose path, so they are pinned
// directly rather than through whichever tool happens to call them. The wordings themselves
// belong to inline-notes.ts and are imported rather than transcribed: what these assert is
// that compose-inline routes each outcome to the right one.

// ---------------------------------------------------------------------------
// planAuthoredInlineImages — what the caller's own html may reference
// ---------------------------------------------------------------------------

const plan = (over: Partial<Parameters<typeof planAuthoredInlineImages>[0]> = {}) =>
  planAuthoredInlineImages({
    callerHtml: '', htmlShips: true, attachmentsEnabled: true, surface: 'compose', ...over,
  });

/** An identifier of the shape this server mints when it carries an image out of an original. */
const RESERVED = `ii-${'a'.repeat(32)}@inline.invalid`;

describe('planAuthoredInlineImages — files supplied for one call', () => {
  it('refuses two files sharing an identifier, with the count interpolated', () => {
    // Three items sharing an identifier must not read as two.
    assert.throws(
      () => plan({ specs: [{ cid: 'logo' }, { cid: 'logo' }, { cid: 'logo' }] as any }),
      (err: unknown) => err instanceof InvalidInputError
        && (err as Error).message === rejectCidCollisionInCall(3, 'logo'),
    );
  });

  it('refuses the collision before reading the html, so an unreferenced pair still fails', () => {
    // The refusal is about the call being ambiguous, not about what the body happens to show.
    assert.throws(
      () => plan({ callerHtml: '', specs: [{ cid: 'logo' }, { cid: 'logo' }] as any }),
      InvalidInputError,
    );
  });

  it('displays nothing, and refuses nothing, when the caller wrote no html', () => {
    // A file carrying an identifier still rides along as an ordinary attachment; the degrade
    // is reported after the upload rather than refused here.
    const out = plan({ callerHtml: '   ', specs: [{ cid: 'logo' }] as any });
    assert.deepEqual([...out.inlineCids], []);
    assert.equal(out.unparsableCidText, false);
  });

  it('treats a non-string callerHtml as no html at all', () => {
    const out = plan({ callerHtml: { toString: () => '<img src="cid:logo">' } as any });
    assert.deepEqual([...out.inlineCids], []);
  });
});

describe('planAuthoredInlineImages — references the caller authored', () => {
  it('displays a reference a supplied file answers', () => {
    const out = plan({ callerHtml: '<p><img src="cid:logo"></p>', specs: [{ cid: 'logo' }] as any });
    assert.deepEqual([...out.inlineCids], ['logo']);
  });

  it('displays nothing when the message ships no html, though the reference resolves', () => {
    const out = plan({
      callerHtml: '<p><img src="cid:logo"></p>', htmlShips: false, specs: [{ cid: 'logo' }] as any,
    });
    assert.deepEqual([...out.inlineCids], []);
  });

  it('leaves a supplied file the body never references out of the displayed set', () => {
    const out = plan({
      callerHtml: '<p>no images</p>', specs: [{ cid: 'logo' }, { cid: 'chart' }] as any,
    });
    assert.deepEqual([...out.inlineCids], []);
  });

  it("refuses an identifier of this server's own minted shape, with the compose wording", () => {
    // A compose call has no earlier message to have carried an image out of, so a reference
    // of that shape names something that can never exist.
    assert.throws(
      () => plan({ callerHtml: `<img src="cid:${RESERVED}">` }),
      (err: unknown) => (err as Error).message === rejectReservedCidRef(RESERVED, 'compose'),
    );
  });

  it('refuses a dangling reference with the compose wording on a compose surface', () => {
    assert.throws(
      () => plan({ callerHtml: '<img src="cid:missing">' }),
      (err: unknown) => (err as Error).message
        === rejectDanglingCidRef('missing', { attachmentsEnabled: true }),
    );
  });

  it('refuses a dangling reference with the NOTE wording on a reply or forward', () => {
    // The note wording has to say quoted images arrive on their own, or the caller reads the
    // refusal as an instruction to author references for them.
    const message = (() => {
      try { plan({ callerHtml: '<img src="cid:missing">', surface: 'note' }); } catch (err) {
        return (err as Error).message;
      }
      return assert.fail('expected a refusal');
    })();
    assert.equal(message, rejectNoteCidRef('missing', { attachmentsEnabled: true }));
    assert.notEqual(message, rejectDanglingCidRef('missing', { attachmentsEnabled: true }));
  });

  it('carries attachmentsEnabled into the refusal, so a repair is never one this server cannot do', () => {
    assert.throws(
      () => plan({ callerHtml: '<img src="cid:missing">', attachmentsEnabled: false }),
      (err: unknown) => (err as Error).message
        === rejectDanglingCidRef('missing', { attachmentsEnabled: false }),
    );
  });

  it('notes reference-shaped text it could not act on, rather than refusing it', () => {
    // A CSS url() is not an <img> this server would rewrite. The false-positive class is
    // unbounded, so it is reported and left alone.
    const out = plan({ callerHtml: '<div style="background:url(cid:bg)">x</div>' });
    assert.equal(out.unparsableCidText, true);
    assert.deepEqual([...out.inlineCids], []);
  });

  it('does not double-count a real reference as unparsable text', () => {
    const out = plan({ callerHtml: '<img src="cid:logo">', specs: [{ cid: 'logo' }] as any });
    assert.equal(out.unparsableCidText, false);
  });

  it('assumes attachments are available when no context is supplied', () => {
    assert.equal(DEFAULT_INLINE_CONTEXT.attachmentsEnabled, true);
    assert.equal(DEFAULT_INLINE_CONTEXT.specs, undefined);
  });
});

// ---------------------------------------------------------------------------
// reportAuthoredInlineImages — what the saved draft turned out to carry
// ---------------------------------------------------------------------------

const EMPTY_PLAN = { inlineCids: new Set<string>(), unparsableCidText: false };

const inlineUpload = (cid: string, name = 'logo.png') =>
  ({ blobId: 'b-' + cid, type: 'image/png', name, disposition: 'inline', cid }) as any;

describe('reportAuthoredInlineImages — the confirmation read', () => {
  it('does not re-read the draft when the call embedded nothing', async () => {
    let reads = 0;
    const notes = await reportAuthoredInlineImages({
      uploaded: [{ blobId: 'b', type: 'application/pdf', name: 'a.pdf', disposition: 'attachment' } as any],
      plan: EMPTY_PLAN,
      emailId: 'd1',
      readBack: async () => { reads += 1; return {}; },
    });
    assert.equal(reads, 0);
    assert.ok(notes.every((n) => n !== noteEmbedUnconfirmed()), notes.join(' | '));
  });

  it('reads the draft back for an image the call minted, not only for one it uploaded', async () => {
    // A reply whose only images come from the quote is the ordinary case of this feature;
    // gating the read on uploads alone would make it the one case never checked.
    const seen: string[] = [];
    await reportAuthoredInlineImages({
      plan: EMPTY_PLAN, mintedCids: ['m1'], emailId: 'd1',
      readBack: async (id) => { seen.push(id); return { attachments: [{ cid: 'm1', size: 10 }] }; },
    });
    assert.deepEqual(seen, ['d1']);
  });

  it('says so when the saved draft is short of what this call attached', async () => {
    const notes = await reportAuthoredInlineImages({
      uploaded: [inlineUpload('logo')], plan: EMPTY_PLAN, mintedCids: ['m1'], emailId: 'd1',
      readBack: async () => ({ attachments: [{ cid: 'logo', size: 70 }] }),
    });
    assert.ok(notes.includes(noteEmbedMissingAfterSave(1)), notes.join(' | '));
  });

  it('says the save was not confirmed when the read fails, and still returns notes', async () => {
    // The draft already exists by this point; nothing this step finds may turn it into a
    // failure.
    const notes = await reportAuthoredInlineImages({
      uploaded: [inlineUpload('logo')], plan: EMPTY_PLAN, emailId: 'd1',
      readBack: async () => { throw new Error('network'); },
    });
    assert.ok(notes.includes(noteEmbedUnconfirmed()), notes.join(' | '));
  });

  it('reports a file that carried an identifier but was not displayed as degraded, not dropped', async () => {
    const notes = await reportAuthoredInlineImages({
      uploaded: [{ blobId: 'b', type: 'image/png', name: 'logo.png', disposition: 'attachment', cid: 'logo' } as any],
      plan: EMPTY_PLAN, emailId: 'd1',
      readBack: async () => ({}),
    });
    assert.ok(notes.some((n) => /attachment/i.test(n)), notes.join(' | '));
    assert.ok(!notes.includes(noteEmbedUnconfirmed()), notes.join(' | '));
  });

  it('passes the unparsable-text finding through to a sentence', async () => {
    const notes = await reportAuthoredInlineImages({
      plan: { inlineCids: new Set(), unparsableCidText: true }, emailId: 'd1',
      readBack: async () => ({}),
    });
    assert.ok(notes.includes(noteUnparsableCidText()), notes.join(' | '));
  });

  it('takes the embedded sizes from the read-back, which is the only place they exist', async () => {
    // A file just uploaded has no size on the part object; the server sets it.
    const notes = await reportAuthoredInlineImages({
      uploaded: [inlineUpload('logo')], plan: EMPTY_PLAN, emailId: 'd1',
      readBack: async () => ({ attachments: [{ cid: 'logo', size: 2048 }] }),
    });
    assert.ok(notes.some((n) => n.includes('2 KB')), notes.join(' | '));
  });
});

// ---------------------------------------------------------------------------
// recordQuoteImages — what a quoted or forwarded block did with the originals
// ---------------------------------------------------------------------------

const png = (name: string, size = 100) => ({ cid: name, blobId: 'B-' + name, type: 'image/png', name, size }) as any;

function outcome(over: Partial<QuoteImageOutcome> = {}): QuoteImageOutcome {
  return {
    minted: [], mappings: [], resolvedParts: [], unresolvedRefs: [],
    droppedDataImages: 0, droppedUnsupportedImages: 0, htmlQuoteShips: true, ...over,
  } as QuoteImageOutcome;
}

describe('recordQuoteImages', () => {
  it('records nothing and carries nothing when there is no outcome at all', () => {
    const ledger = new InlineNoteLedger();
    const carry = recordQuoteImages(ledger, undefined, 'reply');
    assert.deepEqual(carry.minted, []);
    assert.deepEqual(carry.mintedCids, []);
    assert.equal(carry.embedded.size, 0);
    assert.equal(carry.resolvedPartCount, undefined);
    assert.deepEqual(ledger.emit({ surface: 'draft' }), []);
  });

  it('copies the minted parts rather than aliasing the outcome that produced them', () => {
    const minted = [{ blobId: 'B1', type: 'image/png', name: 'logo.png', cid: 'm1', disposition: 'inline' }] as any;
    const carry = recordQuoteImages(new InlineNoteLedger(), outcome({ minted }), 'reply');
    assert.deepEqual(carry.minted, minted);
    assert.notEqual(carry.minted[0], minted[0]);
    assert.deepEqual(carry.mintedCids, ['m1']);
  });

  it('reports a resolved part that DID embed through the mappings, keyed by its emitted id', () => {
    const source = png('logo');
    const ledger = new InlineNoteLedger();
    const carry = recordQuoteImages(
      ledger, outcome({ mappings: [{ cid: 'm1', source }] as any, resolvedParts: [source] }), 'reply',
    );
    assert.deepEqual([...carry.embedded], [source]);
    assert.equal(carry.resolvedPartCount, 1);
    assert.ok(ledger.emit({ surface: 'draft' }).length > 0);
  });

  it('reports a reply that ships no html quote as having DROPPED the images', () => {
    // A reply is a new message quoting an original; with no html quote there is no body to
    // display them in, so they are lost and that is said out loud.
    const ledger = new InlineNoteLedger();
    const carry = recordQuoteImages(
      ledger, outcome({ resolvedParts: [png('a'), png('b')], htmlQuoteShips: false }), 'reply',
    );
    assert.equal(carry.resolvedPartCount, undefined);
    assert.ok(ledger.emit({ surface: 'draft' }).some((n) => /dropped/i.test(n)));
  });

  it('reports nothing dropped on a forward, which carries the image as an attachment instead', () => {
    const ledger = new InlineNoteLedger();
    recordQuoteImages(
      ledger, outcome({ resolvedParts: [png('a'), png('b')], htmlQuoteShips: false }), 'forward',
    );
    assert.deepEqual(ledger.emit({ surface: 'draft' }), []);
  });

  it('counts a reference that resolved to nothing, and the image forms it cannot carry', () => {
    const ledger = new InlineNoteLedger();
    recordQuoteImages(
      ledger,
      outcome({ unresolvedRefs: ['gone'], droppedDataImages: 2, droppedUnsupportedImages: 3 }),
      'reply',
    );
    const notes = ledger.emit({ surface: 'draft' }).join(' | ');
    assert.match(notes, /2/);
    assert.match(notes, /3/);
  });
});
