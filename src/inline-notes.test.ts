import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  describePartNames,
  emitInlineNotes,
  formatSize,
  InlineNoteLedger,
  noteAttachmentsExcluded,
  noteDegradedToAttachments,
  noteDraftEmbeds,
  noteDroppedDataImages,
  noteDroppedQuoteImages,
  noteEmbeddedFromOriginal,
  noteEmbeddedFromQuote,
  noteEmbeddedPartially,
  noteForwardPooled,
  noteForwardUnresolvedReferences,
  noteRemovedEmbedded,
  noteSentWithEmbedded,
  noteSkippedReferences,
  noteUnparsableCidText,
  rejectBrokenDraft,
  rejectCidCollisionInCall,
  rejectCidCollisionOnDraft,
  rejectClearAttachmentsDanglingRefs,
  rejectDanglingCidRef,
  rejectNoteCidRef,
  rejectRemovalDanglingRef,
  rejectReservedCidRef,
  rejectUnrecreatableCid,
} from './inline-notes.js';
import type { NoteTally } from './inline-notes.js';

const KB214 = 214 * 1024;
const MB1_4 = Math.round(1.4 * 1024 * 1024);
const MINT = `ii-${'a'.repeat(32)}@inline.invalid`;

const ENABLED = { attachmentsEnabled: true };
const DISABLED = { attachmentsEnabled: false };

describe('formatSize', () => {
  it('reports kilobytes below a megabyte', () => {
    assert.equal(formatSize(KB214), '214 KB');
    assert.equal(formatSize(1024), '1 KB');
  });

  it('reports one decimal place of megabytes at or above a megabyte', () => {
    assert.equal(formatSize(MB1_4), '1.4 MB');
    assert.equal(formatSize(1024 * 1024), '1.0 MB');
    assert.equal(formatSize(12 * 1024 * 1024), '12.0 MB');
  });

  it('never reports a signature logo as "0 MB"', () => {
    // The whole reason for the unit split: these are the real sizes of embedded images.
    assert.equal(formatSize(30 * 1024), '30 KB');
    assert.equal(formatSize(1024 * 1024 - 1), '1024 KB');
  });

  it('never reports something that was carried as "0 KB"', () => {
    assert.equal(formatSize(1), '1 KB');
    assert.equal(formatSize(400), '1 KB');
  });

  it('reports nothing as 0 KB', () => {
    assert.equal(formatSize(0), '0 KB');
  });

  it('treats an absent or nonsensical size as nothing', () => {
    assert.equal(formatSize(-5), '0 KB');
    assert.equal(formatSize(NaN), '0 KB');
    assert.equal(formatSize(Infinity), '0 KB');
    assert.equal(formatSize(undefined as any), '0 KB');
  });
});

describe('describePartNames', () => {
  it('lists the names it has', () => {
    assert.equal(describePartNames(['a.png', 'b.gif']), '"a.png", "b.gif"');
  });

  it('summarizes past the display cap', () => {
    assert.equal(
      describePartNames(['a.png', 'b.png', 'c.png', 'd.png', 'e.png']),
      '"a.png", "b.png", "c.png" …and 2 more',
    );
  });

  it('counts the parts it could not name, not just the names it has', () => {
    assert.equal(describePartNames(['logo.png', null, undefined], 3), '"logo.png" …and 2 more');
  });

  it('renders a hostile name as bounded, quoted data', () => {
    // A filename is sender-controlled, so it can neither close the quoted span it sits in
    // nor smuggle in characters that reorder the sentence around it.
    assert.equal(describePartNames(['a"b']), '"a\'b"');
    assert.equal(describePartNames(['a\nb']), '"ab"');
    assert.equal(describePartNames(['inv‮fdp.exe']), '"invfdp.exe"');
    assert.equal(describePartNames(['z'.repeat(200)]), `"${'z'.repeat(64)}…"`);
  });

  it('renders nothing when no part had a name', () => {
    assert.equal(describePartNames([null, undefined, ''], 3), '');
    assert.equal(describePartNames([]), '');
  });
});

describe('the notes a call emits', () => {
  it('reports a reply that embedded everything the quote referenced', () => {
    assert.equal(
      noteEmbeddedFromQuote(3, KB214),
      'This draft embeds 3 image(s) from the quoted message (214 KB).',
    );
  });

  it('reports a reply that embedded only some of them', () => {
    assert.equal(
      noteEmbeddedPartially(2, 5, KB214),
      'Embedded 2 of 5 image part(s) referenced by the quote (214 KB embedded); ' +
      '3 could not be embedded and are not part of this draft.',
    );
  });

  it('reports references that matched no part at all', () => {
    assert.equal(noteSkippedReferences(3), '3 reference(s) matched no part and were skipped.');
  });

  it('reports a forward that carried the original\'s embedded images', () => {
    assert.equal(
      noteEmbeddedFromOriginal(2, MB1_4),
      'This draft embeds 2 image(s) from the original (1.4 MB).',
    );
  });

  it('reports references in a forwarded original that matched no part', () => {
    assert.equal(
      noteForwardUnresolvedReferences(4),
      "4 image reference(s) in the original's body had no matching part; " +
      'nothing was carried for them.',
    );
  });

  it('reports media a forward attached rather than embedded', () => {
    assert.equal(
      noteForwardPooled(3, ['logo.png']),
      '3 media part(s) could not be embedded and were attached as regular attachments: ' +
      '"logo.png" …and 2 more — re-run with asAttachment: true for full fidelity, ' +
      'then delete this draft.',
    );
  });

  it('reports pooled media even when no part had a name', () => {
    assert.equal(
      noteForwardPooled(1, []),
      '1 media part(s) could not be embedded and were attached as regular attachments — ' +
      're-run with asAttachment: true for full fidelity, then delete this draft.',
    );
  });

  it('takes the remedy from its caller, for a tool whose caller has a better lever', () => {
    // The default remedy is right where the forward's format is inferred and re-running as
    // .eml is the only lever. On a tool where the caller places the block themselves it is
    // not, so the sentence ends where that caller's fix is.
    assert.equal(
      noteForwardPooled(1, ['logo.png'], 'put the token in htmlBody.'),
      '1 media part(s) could not be embedded and were attached as regular attachments: ' +
      '"logo.png" — put the token in htmlBody.',
    );
  });

  it('reports attachments the caller asked to leave behind', () => {
    assert.equal(
      noteAttachmentsExcluded(5, 2, false),
      '5 attachment(s), including 2 image(s), were not included because ' +
      'includeOriginalAttachments is false.',
    );
  });

  it('adds the body-embedded sentence only when such an image really was carried', () => {
    // Unconditional, it would assert carriage on every ordinary forward of a message that
    // has attachments but no embedded images at all.
    assert.equal(
      noteAttachmentsExcluded(5, 2, true),
      '5 attachment(s), including 2 image(s), were not included because ' +
      'includeOriginalAttachments is false. Body-embedded images were still carried — ' +
      'they are part of the message body.',
    );
  });

  it('names the block a removal emptied, in the wording the edit already uses', () => {
    assert.equal(
      noteRemovedEmbedded(2, 'the quote'),
      'Removed 2 image(s) that were embedded in the quote.',
    );
    assert.equal(
      noteRemovedEmbedded(1, 'the forwarded block'),
      'Removed 1 image(s) that were embedded in the forwarded block.',
    );
  });

  it('reports the caller\'s own images that could not be embedded', () => {
    assert.equal(
      noteDegradedToAttachments(3),
      '3 of your image(s) became regular attachments (nothing in the body displays them).',
    );
  });

  it('reports quoted images that could not be carried at all', () => {
    assert.equal(
      noteDroppedQuoteImages(2),
      '2 image(s) from the quoted message were dropped and are not part of this draft.',
    );
  });

  it('reports data: images the quote carried', () => {
    assert.equal(
      noteDroppedDataImages(2),
      '2 data:-URI image(s) in the quoted message were dropped (not supported); ' +
      'the rest of the quote was kept.',
    );
  });

  it('reports a compose or edit that embedded the caller\'s images', () => {
    assert.equal(noteDraftEmbeds(1, KB214), 'This draft embeds 1 image(s) (214 KB).');
  });

  it('reports what a sent message actually carried', () => {
    assert.equal(
      noteSentWithEmbedded(2, MB1_4),
      'Sent with 2 embedded image(s) (1.4 MB).',
    );
  });

  it('reports reference-shaped text it could not act on, as a note and not an error', () => {
    assert.equal(
      noteUnparsableCidText(),
      "htmlBody contains text that looks like an embedded-image (cid:) reference this " +
      "server can't parse as an image; it was left as-is.",
    );
  });
});

describe('the refusals a call raises', () => {
  it('offers the repair when the identifier is one a caller could author', () => {
    assert.equal(
      rejectDanglingCidRef('logo', ENABLED),
      'htmlBody references cid "logo" but no attachment supplies it. ' +
      'Remove the <img> reference, or add an attachments item with cid: "logo".',
    );
  });

  it('echoes an authorable identifier raw, so the repair can be pasted back', () => {
    assert.ok(rejectDanglingCidRef('logo_1.png-x', ENABLED).includes('cid: "logo_1.png-x"'));
  });

  it('withholds the repair for an identifier the vet would then reject', () => {
    // A foreign identifier always contains an "@", so pointing at "add an attachments
    // item" would be a dead end.
    const message = rejectDanglingCidRef('logo@sender.example', ENABLED);
    assert.equal(
      message,
      'htmlBody references cid "logo@sender.example" but no attachment supplies it. ' +
      'Remove the <img> reference, or recreate the draft without it.',
    );
    assert.ok(!message.includes('add an attachments item'));
  });

  it('says why no repair exists when this server cannot attach files at all', () => {
    assert.equal(
      rejectDanglingCidRef('logo', DISABLED),
      'htmlBody references cid "logo" but no attachment supplies it. ' +
      'Remove the <img> reference (sending attachments is disabled on this server — neither ' +
      'FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is set — so no attachments item can supply it).',
    );
  });

  it('renders a hostile dangling value as bounded, quoted data', () => {
    const message = rejectDanglingCidRef(`x‮${'y'.repeat(200)}`, ENABLED);
    assert.ok(message.includes(`"x${'y'.repeat(63)}…"`));
    assert.ok(!message.includes('‮'));
  });

  // Scoped to AUTHORING one: the caller's body names a minted identifier the draft carries
  // no part under. Handing back a reference to a part the draft actually has is the normal
  // read-edit-write shape and never reaches this message, so the wording must not tell the
  // caller to stop referencing them — only to stop inventing them.
  it('refuses a reference authoring one of this server\'s own identifiers', () => {
    assert.equal(
      rejectReservedCidRef(MINT),
      `htmlBody references cid "${MINT}", a server-managed identifier for quoted images, ` +
      'and this draft carries no part under it. A minted identifier survives an edit for as ' +
      'long as your body keeps referencing it, and a reference you drop takes the part with ' +
      'it — but never author a NEW one. To embed your own image, add an attachments item ' +
      'with a cid of your choosing.',
    );
  });

  // The compose surface has no draft, so the edit wording would describe a thing that does
  // not exist. The remedy is the same on both; only the diagnosis moves.
  it('diagnoses the same refusal differently on a compose, where there is no draft', () => {
    const message = rejectReservedCidRef(MINT, 'compose');
    assert.ok(!message.includes('this draft'));
    assert.ok(message.includes('this server assigns them itself — a body never authors one.'));
    assert.ok(message.includes('add an attachments item with a cid of your choosing.'));
  });

  it('refuses a removal the surviving body still references', () => {
    assert.equal(
      rejectRemovalDanglingRef('logo'),
      'removeAttachments would remove an image the draft\'s body still references ' +
      '(cid "logo"). Remove the <img> reference in the same call, or keep the attachment.',
    );
  });

  it('refuses an attachment wipe the surviving body still references', () => {
    assert.equal(
      rejectClearAttachmentsDanglingRefs(),
      'clearFields: ["attachments"] would strip image(s) the surviving body still ' +
      'references. Rewrite or clear that body in the same call, or keep the attachments.',
    );
  });

  it('refuses a Content-ID already in use on the draft', () => {
    assert.equal(
      rejectCidCollisionOnDraft('logo'),
      'cid "logo" is already used by another attachment on this draft; ' +
      'each embedded image needs a distinct cid.',
    );
  });

  it('reports how many items in one call share a Content-ID', () => {
    // Interpolated rather than worded, so three sharers do not read as "Two".
    assert.equal(
      rejectCidCollisionInCall(3, 'logo'),
      '3 attachments items share cid "logo"; each embedded image needs a distinct cid.',
    );
  });

  it('refuses to edit a draft whose identifier it cannot reproduce', () => {
    assert.equal(
      rejectUnrecreatableCid('a\x00b'),
      'This draft has an attachment whose embedded-image identifier (Content-ID "ab") ' +
      'this server cannot safely re-create. This server does not edit drafts in that ' +
      'state. Recreate it with draft_email — mode:\'reply\' or mode:\'forward\' if it is a ' +
      'reply or forward (that preserves the conversation threading; read the draft\'s ' +
      'In-Reply-To or X-Forwarded-Message-Id via get_email, then find the original with ' +
      'search_emails using the bare id, without angle brackets), otherwise mode:\'new\' — ' +
      'then delete this one; or edit it in the mail client that created it.',
    );
  });

  it('routes a reply or forward draft through the modes that keep the threading', () => {
    // Recreating one as a new message splits the conversation.
    for (const message of [
      rejectUnrecreatableCid('x'),
      rejectBrokenDraft(['a@b'], ENABLED),
    ]) {
      assert.ok(message.includes("mode:'reply' or mode:'forward'"));
      assert.ok(message.includes('without angle brackets'));
    }
  });

  it('refuses a body edit on a draft whose stored body already dangles', () => {
    assert.equal(
      rejectBrokenDraft(['logo@sender.example'], ENABLED),
      'This draft\'s stored body references image identifier(s) with no matching ' +
      'attachment ("logo@sender.example"). This server won\'t edit its body in that state ' +
      'unless the edit resolves the missing reference(s) — metadata and attachment edits ' +
      'still work. Recreate it with draft_email — mode:\'reply\' or mode:\'forward\' if it ' +
      'is a reply or forward (read the draft\'s In-Reply-To or X-Forwarded-Message-Id via ' +
      'get_email, then find the original with search_emails using the bare id, without ' +
      'angle brackets), otherwise mode:\'new\' — then delete this one.',
    );
  });

  it('says plainly that a metadata edit still works, so the cheap repair is not foreclosed', () => {
    assert.ok(
      rejectBrokenDraft(['a@b'], ENABLED).includes('metadata and attachment edits still work'),
    );
  });

  it('offers the supply-the-image repair only when every value could be authored', () => {
    assert.ok(
      rejectBrokenDraft(['logo'], ENABLED).endsWith(
        'Or add an attachments item with cid "logo" to supply the missing image.',
      ),
    );
    assert.ok(
      rejectBrokenDraft(['logo', 'icon'], ENABLED).endsWith(
        'Or add attachments items with cid "logo", "icon" to supply the missing images.',
      ),
    );
    assert.ok(!rejectBrokenDraft(['logo', 'a@b'], ENABLED).includes('Or add'));
    assert.ok(!rejectBrokenDraft(['logo'], DISABLED).includes('Or add'));
  });

  it('lists every dangling value it found', () => {
    assert.ok(rejectBrokenDraft(['a@b', 'c@d'], ENABLED).includes('("a@b", "c@d")'));
  });

  it('tells a note author that quoted images arrive on their own', () => {
    assert.equal(
      rejectNoteCidRef('logo', ENABLED),
      'htmlBody (your note) references cid "logo", which no attachments item supplies. ' +
      'Quoted images appear inside the quote automatically; to embed your own image, ' +
      'add an attachments item with cid: "logo" — otherwise remove the reference.',
    );
  });

  it('drops the note author\'s repair clause when attachments are disabled', () => {
    assert.ok(rejectNoteCidRef('logo', DISABLED).includes('neither FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is set'));
  });
});

describe('InlineNoteLedger', () => {
  it('counts each part once, from its final outcome', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'embedded', bytes: KB214 });
    ledger.record({ key: 'B2', outcome: 'embedded', bytes: 1024 });
    const tally = ledger.tally();
    assert.equal(tally.embedded, 2);
    assert.equal(tally.embeddedBytes, KB214 + 1024);
  });

  it('lets a later outcome replace an earlier candidate for the same part', () => {
    // The property the whole mechanism exists for: a part that was going to be embedded
    // and then came off the draft is reported once, as removed.
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'embedded', bytes: KB214 });
    ledger.record({ key: 'B1', outcome: 'removed' });
    const tally = ledger.tally();
    assert.equal(tally.embedded, 0);
    assert.equal(tally.embeddedBytes, 0);
    assert.equal(tally.removed, 1);
  });

  it('never emits a note claiming a part survived when the call ultimately cleared it', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'degraded' });
    ledger.record({ key: 'B1', outcome: 'removed' });
    const notes = ledger.emit({ surface: 'draft', keepNoun: 'the quote' });
    assert.deepEqual(notes, ['Removed 1 image(s) that were embedded in the quote.']);
  });

  it('speaks up about a part it could not carry at all', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'dropped' });
    ledger.record({ key: 'B2', outcome: 'dropped' });
    assert.deepEqual(ledger.emit({ surface: 'reply' }), [
      '2 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('reports a part that was going to be embedded and was then dropped exactly once', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'embedded', bytes: KB214 });
    ledger.record({ key: 'B1', outcome: 'dropped' });
    assert.deepEqual(ledger.emit({ surface: 'reply' }), [
      '1 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('keeps reference-level losses separate from part outcomes', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'B1', outcome: 'embedded', bytes: 1024 });
    ledger.countRefs('unresolvedRefs', 2);
    ledger.countRefs('droppedDataImages');
    ledger.countRefs('droppedUnsupportedImages', 3);
    const tally = ledger.tally();
    assert.equal(tally.embedded, 1);
    assert.equal(tally.unresolvedRefs, 2);
    assert.equal(tally.droppedDataImages, 1);
    assert.equal(tally.droppedUnsupportedImages, 3);
  });

  it('ignores a non-positive reference count', () => {
    const ledger = new InlineNoteLedger();
    ledger.countRefs('unresolvedRefs', 0);
    ledger.countRefs('droppedDataImages', -3);
    assert.equal(ledger.tally().unresolvedRefs, 0);
    assert.equal(ledger.tally().droppedDataImages, 0);
  });

  it('breaks out how many of the excluded attachments were images', () => {
    const ledger = new InlineNoteLedger();
    ledger.record({ key: 'a', outcome: 'notIncluded', isImage: true });
    ledger.record({ key: 'b', outcome: 'notIncluded' });
    const tally = ledger.tally();
    assert.equal(tally.notIncluded, 2);
    assert.equal(tally.notIncludedImages, 1);
  });

  it('copies what it is handed, so a later mutation cannot rewrite history', () => {
    const ledger = new InlineNoteLedger();
    const record = { key: 'B1', outcome: 'embedded' as const, bytes: 1024 };
    ledger.record(record);
    record.bytes = 99999;
    assert.equal(ledger.tally().embeddedBytes, 1024);
  });

  it('says nothing at all when nothing happened', () => {
    assert.deepEqual(new InlineNoteLedger().emit({ surface: 'draft' }), []);
  });
});

describe('emitInlineNotes', () => {
  // Typed as NoteTally in both directions on purpose: a counter added to the tally
  // has to be defaulted here, in one place, and a typo in an override is rejected
  // rather than silently ignored. Without that, a new counter reads as 0 in every
  // case below and the note it drives goes untested.
  const tally = (over: Partial<NoteTally> = {}): NoteTally => ({
    embedded: 0,
    embeddedBytes: 0,
    pooled: 0,
    pooledNames: [],
    degraded: 0,
    notIncluded: 0,
    notIncludedImages: 0,
    dropped: 0,
    removed: 0,
    attached: 0,
    unresolvedRefs: 0,
    droppedDataImages: 0,
    droppedUnsupportedImages: 0,
    ...over,
  });

  it('passes a caller-supplied pooled remedy through to the sentence', () => {
    assert.deepEqual(
      emitInlineNotes(
        tally({ pooled: 1, pooledNames: ['logo.png'] }),
        { surface: 'forward', pooledRemedy: 'put the token in htmlBody.' },
      ),
      ['1 media part(s) could not be embedded and were attached as regular attachments: ' +
       '"logo.png" — put the token in htmlBody.'],
    );
    // Omitted, the shared default stands.
    assert.match(
      emitInlineNotes(tally({ pooled: 1, pooledNames: [] }), { surface: 'forward' })[0],
      /re-run with asAttachment: true for full fidelity, then delete this draft\.$/,
    );
  });

  it('describes an embed in the words of the tool that did it', () => {
    const t = tally({ embedded: 2, embeddedBytes: KB214 });
    assert.deepEqual(emitInlineNotes(t, { surface: 'reply' }), [
      'This draft embeds 2 image(s) from the quoted message (214 KB).',
    ]);
    assert.deepEqual(emitInlineNotes(t, { surface: 'forward' }), [
      'This draft embeds 2 image(s) from the original (214 KB).',
    ]);
    assert.deepEqual(emitInlineNotes(t, { surface: 'draft' }), [
      'This draft embeds 2 image(s) (214 KB).',
    ]);
    assert.deepEqual(emitInlineNotes(t, { surface: 'send' }), [
      'Sent with 2 embedded image(s) (214 KB).',
    ]);
  });

  it('switches a reply to the partial wording when parts were left behind', () => {
    assert.deepEqual(
      emitInlineNotes(tally({ embedded: 2, embeddedBytes: KB214 }), {
        surface: 'reply',
        resolvedPartCount: 5,
      }),
      [
        'Embedded 2 of 5 image part(s) referenced by the quote (214 KB embedded); ' +
        '3 could not be embedded and are not part of this draft.',
      ],
    );
  });

  it('reports the shortfall even when nothing at all could be embedded', () => {
    // The case that most needs saying: a quote referencing three images, none of them
    // usable, would otherwise ship with the images missing and no word about it.
    assert.deepEqual(
      emitInlineNotes(tally({ embedded: 0 }), { surface: 'reply', resolvedPartCount: 3 }),
      [
        'Embedded 0 of 3 image part(s) referenced by the quote (0 KB embedded); ' +
        '3 could not be embedded and are not part of this draft.',
      ],
    );
  });

  it('says nothing about an embed on the other surfaces when nothing embedded', () => {
    for (const surface of ['forward', 'draft', 'send'] as const) {
      assert.deepEqual(emitInlineNotes(tally({ embedded: 0 }), { surface }), []);
    }
  });

  it('reports images it could not carry at all', () => {
    assert.deepEqual(emitInlineNotes(tally({ dropped: 2 }), { surface: 'reply' }), [
      '2 image(s) from the quoted message were dropped and are not part of this draft.',
    ]);
  });

  it('counts references with no part apart from the parts it did resolve', () => {
    // Disjoint sets: summing them would read as one loss counted twice.
    const notes = emitInlineNotes(
      tally({ embedded: 2, embeddedBytes: KB214, unresolvedRefs: 3 }),
      { surface: 'reply', resolvedPartCount: 2 },
    );
    assert.deepEqual(notes, [
      'This draft embeds 2 image(s) from the quoted message (214 KB).',
      '3 reference(s) matched no part and were skipped.',
    ]);
  });

  it('uses the forward\'s own wording for references that matched nothing', () => {
    assert.deepEqual(emitInlineNotes(tally({ unresolvedRefs: 4 }), { surface: 'forward' }), [
      "4 image reference(s) in the original's body had no matching part; " +
      'nothing was carried for them.',
    ]);
  });

  it('emits what was carried, then what was not, then what came off', () => {
    const notes = emitInlineNotes(
      tally({
        embedded: 1,
        embeddedBytes: 1024,
        pooled: 1,
        pooledNames: ['logo.png'],
        notIncluded: 2,
        notIncludedImages: 1,
        removed: 1,
        degraded: 1,
        dropped: 1,
        droppedDataImages: 1,
        droppedUnsupportedImages: 1,
      }),
      { surface: 'forward', keepNoun: 'the forwarded block' },
    );
    assert.deepEqual(notes, [
      'This draft embeds 1 image(s) from the original (1 KB).',
      '1 media part(s) could not be embedded and were attached as regular attachments: ' +
      '"logo.png" — re-run with asAttachment: true for full fidelity, then delete this draft.',
      '2 attachment(s), including 1 image(s), were not included because ' +
      'includeOriginalAttachments is false. Body-embedded images were still carried — ' +
      'they are part of the message body.',
      'Removed 1 image(s) that were embedded in the forwarded block.',
      '1 of your image(s) became regular attachments (nothing in the body displays them).',
      '1 image(s) from the quoted message were dropped and are not part of this draft.',
      '1 data:-URI image(s) in the quoted message were dropped (not supported); ' +
      'the rest of the quote was kept.',
      '1 image(s) in the quoted message used a reference form this server cannot carry ' +
      'into a quote and were dropped; the rest of the quote was kept.',
    ]);
  });

  it('reports an image dropped for the form of its reference apart from a data: URI', () => {
    // Two different losses with two different causes: a data: URI is content this
    // server declines to re-encode, an unsupported reference form named a location
    // the new message has no origin to resolve against. One sentence for both would
    // tell the reader the wrong thing about what to do next.
    assert.deepEqual(
      emitInlineNotes(tally({ droppedUnsupportedImages: 2 }), { surface: 'reply' }),
      [
        '2 image(s) in the quoted message used a reference form this server cannot carry ' +
        'into a quote and were dropped; the rest of the quote was kept.',
      ],
    );
  });

  it('defaults the removal note to the quote when no block was named', () => {
    assert.deepEqual(emitInlineNotes(tally({ removed: 1 }), { surface: 'draft' }), [
      'Removed 1 image(s) that were embedded in the quote.',
    ]);
  });

  it('reports reference-shaped text that was left alone', () => {
    assert.deepEqual(emitInlineNotes(tally(), { surface: 'draft', unparsableCidText: true }), [
      "htmlBody contains text that looks like an embedded-image (cid:) reference this " +
      "server can't parse as an image; it was left as-is.",
    ]);
  });

  it('says nothing about parts that simply rode along as attachments', () => {
    assert.deepEqual(emitInlineNotes(tally({ attached: 4 }), { surface: 'draft' }), []);
  });
});
