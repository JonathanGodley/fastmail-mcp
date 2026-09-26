// The embedded-image (cid:) checks and reporting for draft_email's three modes (#13), in one
// module so the modes refuse the same mismatches with the same words.
import { InvalidInputError } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { isBlank } from './body-format.js';
import {
  buildUnionParts, extractCidRefs, isReservedCid, sanitizeQuoteHtml,
} from './inline-images.js';
import type { CidPart } from './inline-images.js';
import type { QuoteImageOutcome } from './reply-quote.js';
import {
  InlineNoteLedger, noteEmbedMissingAfterSave, noteEmbedUnconfirmed,
  rejectCidCollisionInCall, rejectDanglingCidRef, rejectNoteCidRef, rejectReservedCidRef,
} from './inline-notes.js';
import type { AttachmentPart } from './jmap-client.js';

/**
 * Which refusal wording a dangling reference gets. `note` (a reply or forward) has to say
 * quoted images arrive on their own, or the caller reads the refusal as an instruction to
 * author references for them.
 */
export type AuthoredInlineSurface = 'compose' | 'note';

/**
 * What a pure body builder needs to check the caller's embedded-image references.
 *
 * Passed in rather than read from the tool args, because a lenient client may send the
 * attachments array as a JSON string, and a check reading the raw argument would see no items.
 */
export interface AuthoredInlineContext {
  specs?: AttachmentSpec[];
  attachmentsEnabled: boolean;
}

// Only for a direct call that is not exercising embedded images; production passes the real value.
export const DEFAULT_INLINE_CONTEXT: AuthoredInlineContext = { attachmentsEnabled: true };

export interface AuthoredInlineInput {
  /** The caller's OWN html, before any quote or forwarded block is merged into it. */
  callerHtml?: unknown;
  htmlShips: boolean;
  /** Already coerced (canonical Content-IDs, see coerceAttachments). */
  specs?: AttachmentSpec[];
  /**
   * False when neither FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is enabled, so a
   * refusal must not suggest supplying a file.
   */
  attachmentsEnabled: boolean;
  surface: AuthoredInlineSurface;
}

export interface AuthoredInlinePlan {
  /** Content-IDs the shipping body displays, for uploadAttachments to disposition. */
  inlineCids: Set<string>;
  /**
   * A FINDING rather than its sentence, so every sentence a compose call emits is composed in
   * one place, from the call's final state.
   */
  unparsableCidText: boolean;
}

/**
 * Validate the caller's embedded-image references against the files they supplied, and
 * decide which of those files the message will actually display.
 *
 * Runs BEFORE anything is uploaded, so a rejected call leaves no orphaned blob, and on the
 * caller's own html rather than the merged body, so the server-generated quote and its
 * identifiers are never held against the caller.
 *
 *  - A real `<img>` reference that no file supplies is a REFUSAL.
 *  - Text that merely LOOKS like a reference (a CSS url(), prose about `cid:`, a pasted MIME
 *    fragment) is a NOTE: the false-positive class is unbounded, so refusing over it would
 *    block real mail.
 */
export function planAuthoredInlineImages(input: AuthoredInlineInput): AuthoredInlinePlan {
  const html = typeof input.callerHtml === 'string' ? input.callerHtml : '';
  const specs = input.specs ?? [];

  // Two files sharing one identifier make every reference to it ambiguous.
  const counts = new Map<string, number>();
  for (const spec of specs) {
    if (spec.cid) counts.set(spec.cid, (counts.get(spec.cid) ?? 0) + 1);
  }
  for (const [cid, n] of counts) {
    if (n > 1) throw new InvalidInputError(rejectCidCollisionInCall(n, cid));
  }

  if (isBlank(html)) {
    // A file carrying an identifier still rides along as a regular attachment, never dropped;
    // the degrade is reported after the upload.
    return { inlineCids: new Set(), unparsableCidText: false };
  }

  const availability = { attachmentsEnabled: input.attachmentsEnabled };
  const suppliedCids = new Set(counts.keys());

  // Read by the same pass the quote rewriter uses, so the two cannot disagree.
  const preciseRefs = sanitizeQuoteHtml(html, { mode: 'collect' }).refs;
  for (const ref of preciseRefs) {
    // A compose call has no original to have carried an image out of, so a reference of the
    // server's minted shape names nothing that can ever exist.
    if (isReservedCid(ref)) throw new InvalidInputError(rejectReservedCidRef(ref, 'compose'));
    if (suppliedCids.has(ref)) continue;
    throw new InvalidInputError(
      input.surface === 'note'
        ? rejectNoteCidRef(ref, availability)
        : rejectDanglingCidRef(ref, availability),
    );
  }

  const broadOnly = extractCidRefs(html).filter((ref) => !preciseRefs.includes(ref));

  const inlineCids = new Set(input.htmlShips ? preciseRefs.filter((r) => suppliedCids.has(r)) : []);
  return { inlineCids, unparsableCidText: broadOnly.length > 0 };
}

export interface AuthoredInlineReport {
  /** The parts uploadAttachments returned for this call's specs, in the same order. */
  uploaded?: AttachmentPart[];
  plan: AuthoredInlinePlan;
  /**
   * Content-IDs this call minted for a quote or forwarded block's images. The confirmation
   * read covers them too, or a reply whose only images come from the quote would never be
   * checked.
   */
  mintedCids?: string[];
  emailId: string;
  /** The client's getEmailById, injected so this stays testable without a network. */
  readBack: (id: string) => Promise<any>;
}

/**
 * Say what the saved draft carries, after the fact.
 *
 * The confirmation read runs ONLY when this call embedded something, the one outcome a caller
 * cannot check from an email id. Its failure is a sentence, never an error: the draft is
 * already saved. Sizes come from that read too, since an uploaded part carries none.
 *
 * No equivalent on the send path, deliberately: send_draft submits the draft by reference,
 * so there is nothing new to confirm.
 */
export async function reportAuthoredInlineImages(
  input: AuthoredInlineReport,
): Promise<string[]> {
  const ledger = new InlineNoteLedger();
  const uploaded = input.uploaded ?? [];
  const followUp: string[] = [];

  const embeddedCids = [
    ...uploaded
      .filter((p) => p.disposition === 'inline' && typeof p.cid === 'string')
      .map((p) => p.cid as string),
    ...(input.mintedCids ?? []),
  ];

  const bytesByCid = new Map<string, number>();
  if (embeddedCids.length > 0) {
    try {
      const saved = await input.readBack(input.emailId);
      const savedByCid = new Map<string, any>();
      for (const { part } of buildUnionParts(saved)) {
        if (typeof part?.cid === 'string' && part.cid !== '') savedByCid.set(part.cid, part);
      }
      const missing = embeddedCids.filter((cid) => !savedByCid.has(cid));
      if (missing.length > 0) followUp.push(noteEmbedMissingAfterSave(missing.length));
      for (const [cid, part] of savedByCid) {
        if (typeof part?.size === 'number' && part.size > 0) bytesByCid.set(cid, part.size);
      }
    } catch {
      followUp.push(noteEmbedUnconfirmed());
    }
  }

  uploaded.forEach((part, index) => {
    // Keyed by position: two files can legitimately be the same bytes under two
    // identifiers, and a blob-keyed record would collapse the pair into one.
    const key = `upload:${index}`;
    const cid = typeof part.cid === 'string' ? part.cid : '';
    if (part.disposition === 'inline' && cid) {
      ledger.record({
        key, outcome: 'embedded', bytes: bytesByCid.get(cid) ?? 0, name: part.name, isImage: true,
      });
      return;
    }
    // An identified file that could not be displayed is on the draft as an ordinary
    // attachment, and only this says so.
    ledger.record({ key, outcome: cid ? 'degraded' : 'attached', name: part.name });
  });

  return [
    ...ledger.emit({ surface: 'draft', unparsableCidText: input.plan.unparsableCidText }),
    ...followUp,
  ];
}

// ---------------------------------------------------------------------------
// Images carried out of a quoted or forwarded original
// ---------------------------------------------------------------------------

/** What a compose path must attach so the block it wrote can display the original's images. */
export interface QuoteCarry {
  /**
   * Parts to attach, each under a freshly minted Content-ID. APPENDED to the caller's own
   * attachments, never assigned over them. Empty whenever no html quote ships.
   */
  minted: AttachmentPart[];
  mintedCids: string[];
  /**
   * The original's parts this call embedded, compared by object identity: a forward walks
   * the same parts afterwards, and matching on a Content-ID would re-derive the decision.
   */
  embedded: Set<CidPart>;
  /** See InlineNoteContext.resolvedPartCount. Undefined when no html quote ships. */
  resolvedPartCount?: number;
}

/**
 * Record what a quote or forwarded block did with the original's embedded images, and hand
 * back the parts the draft has to carry for it.
 *
 * An image that could NOT be embedded is dropped by a reply, and recorded here, but carried
 * by a forward as a regular attachment, so nothing is recorded for it.
 *
 * With an html quote, a resolved part that did not embed is reported ONLY through the
 * shortfall denominator, never also as dropped, which would count one lost image twice.
 */
export function recordQuoteImages(
  ledger: InlineNoteLedger,
  outcome: QuoteImageOutcome | undefined,
  surface: 'reply' | 'forward',
): QuoteCarry {
  const embedded = new Set<CidPart>();
  if (!outcome) return { minted: [], mintedCids: [], embedded };

  for (const mapping of outcome.mappings) {
    embedded.add(mapping.source);
    ledger.record({
      // Unique per mapping, whether minted fresh or claimed from one survivor.
      key: `quote:${mapping.cid}`,
      outcome: 'embedded',
      bytes: typeof mapping.source.size === 'number' ? mapping.source.size : 0,
      name: mapping.source.name,
      isImage: true,
    });
  }

  if (surface === 'reply' && !outcome.htmlQuoteShips) {
    outcome.resolvedParts.forEach((part, index) => {
      ledger.record({ key: `quote-drop:${index}`, outcome: 'dropped', name: part.name, isImage: true });
    });
  }

  ledger.countRefs('unresolvedRefs', outcome.unresolvedRefs.length);
  ledger.countRefs('droppedDataImages', outcome.droppedDataImages);
  ledger.countRefs('droppedUnsupportedImages', outcome.droppedUnsupportedImages);

  return {
    minted: outcome.minted.map((part) => ({ ...part })),
    mintedCids: outcome.minted.map((part) => part.cid),
    embedded,
    ...(outcome.htmlQuoteShips && { resolvedPartCount: outcome.resolvedParts.length }),
  };
}
