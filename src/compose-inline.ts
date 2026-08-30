// The embedded-image (cid:) checks and reporting for the compose path — draft_email, in
// all three of its modes (#13).
//
// Its own module rather than inline in the handler because the modes have to agree. Each of
// them lets a caller author `<img src="cid:...">` against a file in the same call, each has
// to refuse the same mismatches with the same words, and each has to say afterwards whether
// the file was actually embedded or merely attached — which is the one outcome the caller
// cannot see from an email id. Extracting them also keeps the checks testable apart from
// the orchestration that calls them.
//
// Everything here is pure except the confirmation read, which takes the read as an injected
// function so the compose paths stay unit-testable with a mock client.
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
 * Which refusal wording a dangling reference gets.
 *
 * `compose` is a message the caller wrote whole: the repair is to add the file or drop the
 * reference. `note` is a reply or forward, where a quote or forwarded block sits below the
 * caller's own text — that wording has to say quoted images arrive on their own, or the
 * caller reads the refusal as an instruction to author references for them.
 */
export type AuthoredInlineSurface = 'compose' | 'note';

/**
 * What a pure body builder needs to check the caller's embedded-image references.
 *
 * Passed in rather than read from the tool args, because a lenient client may send the whole
 * attachments array as a JSON string: a check reading the raw argument would see no items
 * and wave through a message whose images are all missing. `attachmentsEnabled` rides along
 * because a refusal must not suggest supplying a file on a server that cannot attach one —
 * true when EITHER attachment source is enabled, since either one can supply the image.
 */
export interface AuthoredInlineContext {
  specs?: AttachmentSpec[];
  attachmentsEnabled: boolean;
}

// What a builder assumes when a caller supplies no context at all. Every production path
// passes the real value; this default only governs a direct call that is not exercising
// embedded images, where the difference is unobservable.
export const DEFAULT_INLINE_CONTEXT: AuthoredInlineContext = { attachmentsEnabled: true };

export interface AuthoredInlineInput {
  /** The caller's OWN html, before any quote or forwarded block is merged into it. */
  callerHtml?: unknown;
  /** Whether the composed message will ship an html body at all. */
  htmlShips: boolean;
  /** The already-coerced attachment specs (canonical Content-IDs — see coerceAttachments). */
  specs?: AttachmentSpec[];
  /**
   * False when this server can attach nothing at all — neither a local file
   * (FASTMAIL_ATTACH_DIR) nor content already in the account (FASTMAIL_ALLOW_BLOB_ATTACH).
   * It changes what a repair may suggest: with either source open, "add an attachments
   * item" is a repair the caller can actually carry out.
   */
  attachmentsEnabled: boolean;
  surface: AuthoredInlineSurface;
}

export interface AuthoredInlinePlan {
  /** Content-IDs the shipping body displays, for uploadAttachments to disposition. */
  inlineCids: Set<string>;
  /**
   * The body carried reference-shaped text this server could not act on.
   *
   * Carried as a FINDING rather than as its sentence, so every sentence a compose call
   * emits is still composed in one place, from the call's final state.
   */
  unparsableCidText: boolean;
}

/**
 * Validate the caller's embedded-image references against the files they supplied, and
 * decide which of those files the message will actually display.
 *
 * Runs BEFORE anything is uploaded, and before a reply quote or forwarded block is merged
 * in. Both orderings are deliberate. Refusing first means a rejected call leaves no orphaned
 * blob in the mail store, matching how the malformed-body guards already behave. Checking
 * the caller's own html rather than the merged body means the quote — which this server
 * generates and which legitimately carries its own identifiers — is never held against the
 * caller.
 *
 * Two collectors, two very different consequences:
 *
 *  - A real `<img>` reference that no file supplies is a REFUSAL. The message would ship
 *    with a broken image, and the caller can fix it from the message alone.
 *  - Text that merely LOOKS like a reference — a CSS url(), an SVG href, prose about this
 *    feature, a pasted MIME fragment, this server's own refusal quoted back — is a NOTE. The
 *    false-positive class there is unbounded, and refusing to compose mail over a sentence
 *    that mentions `cid:` would be far worse than saying plainly that it was left alone.
 */
export function planAuthoredInlineImages(input: AuthoredInlineInput): AuthoredInlinePlan {
  const html = typeof input.callerHtml === 'string' ? input.callerHtml : '';
  const specs = input.specs ?? [];

  // Two files sharing one identifier make every reference to it ambiguous, so this is
  // refused before either is read off disk. The count is interpolated: three items sharing
  // an identifier must not read as two.
  const counts = new Map<string, number>();
  for (const spec of specs) {
    if (spec.cid) counts.set(spec.cid, (counts.get(spec.cid) ?? 0) + 1);
  }
  for (const [cid, n] of counts) {
    if (n > 1) throw new InvalidInputError(rejectCidCollisionInCall(n, cid));
  }

  if (isBlank(html)) {
    // Nothing authored a reference, so nothing can be displayed. Any file carrying an
    // identifier still rides along as a regular attachment — the caller supplied it, and
    // dropping it is never an option — and the degrade is reported after the upload.
    return { inlineCids: new Set(), unparsableCidText: false };
  }

  const availability = { attachmentsEnabled: input.attachmentsEnabled };
  const suppliedCids = new Set(counts.keys());

  // The precise collector: references this server would really act on, read off the <img>
  // elements by the same pass the quote rewriter uses, so the two cannot disagree.
  const preciseRefs = sanitizeQuoteHtml(html, { mode: 'collect' }).refs;
  for (const ref of preciseRefs) {
    // Identifiers of the server's own shape are refused outright rather than treated as
    // dangling, and on THIS surface the reason is not durability. A minted identifier is
    // durable — it survives an edit for as long as the body keeps referencing it, which is
    // what lets an image-bearing draft be read and handed straight back. What cannot happen
    // is a body claiming one in advance: this server assigns them when it carries an image
    // out of a quoted or forwarded original, and a compose call has no earlier message to
    // have carried anything out of, so a reference of that shape was authored by the caller
    // and names nothing that can ever exist.
    if (isReservedCid(ref)) throw new InvalidInputError(rejectReservedCidRef(ref, 'compose'));
    if (suppliedCids.has(ref)) continue;
    throw new InvalidInputError(
      input.surface === 'note'
        ? rejectNoteCidRef(ref, availability)
        : rejectDanglingCidRef(ref, availability),
    );
  }

  // The broad collector, over the same html: anything reference-shaped the precise pass did
  // not see. A note, never a refusal — see the contract above.
  const broadOnly = extractCidRefs(html).filter((ref) => !preciseRefs.includes(ref));

  // A reference only displays a file if an html body ships at all. When none does, the
  // references are moot and every supplied file degrades to a regular attachment.
  const inlineCids = new Set(input.htmlShips ? preciseRefs.filter((r) => suppliedCids.has(r)) : []);
  return { inlineCids, unparsableCidText: broadOnly.length > 0 };
}

export interface AuthoredInlineReport {
  /** The parts uploadAttachments returned for this call's specs, in the same order. */
  uploaded?: AttachmentPart[];
  plan: AuthoredInlinePlan;
  /**
   * Content-IDs this call minted so a quote or forwarded block could display the ORIGINAL's
   * images. They are embedded content this call put on the draft just as an uploaded file
   * is, so the confirmation read covers them: a reply whose only images come from the quote
   * is the ordinary case of that feature, and gating the read on uploads alone would make it
   * the one case that never gets checked. Their sizes are already known from the original's
   * own parts, so the read is here for the did-it-survive question, not for bytes.
   */
  mintedCids?: string[];
  /** The draft this call saved, for the confirmation read. */
  emailId: string;
  /** The client's getEmailById, injected so this stays testable without a network. */
  readBack: (id: string) => Promise<any>;
}

/**
 * Say what the saved draft carries, after the fact.
 *
 * The confirmation read runs ONLY when this call embedded something. Assembling a message
 * that displays an image is the one outcome whose success the result text would otherwise
 * assert without evidence, and it is also the one a caller cannot check from an email id; an
 * ordinary attachment-only draft gets no extra round trip. It has its own try/catch and its
 * own sentence, because the draft already exists by this point and nothing this step finds
 * may turn a saved draft into a failure.
 *
 * There is no equivalent on the send path, and that is deliberate rather than an oversight:
 * send_draft submits a draft BY REFERENCE without recreating it, so there is nothing new to
 * confirm, and it never re-validates a draft's images — the checks above, at create time,
 * are the whole of this server's say over what a message embeds.
 *
 * Sizes come from that read too. A file this call just uploaded has no size on the part
 * object (a server-set field, deliberately not sent back on the Email), so when the read
 * does not happen the sizes are simply unknown — which is exactly when the unconfirmed
 * sentence is there to explain the gap.
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
    // A file the caller gave an identifier that could not be displayed did not fail and was
    // not dropped — it is on the draft as an ordinary attachment, and only this says so.
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
   * Parts to attach, each under a freshly minted Content-ID.
   *
   * APPENDED to whatever the caller's own attachments produced, never assigned over them:
   * a reply that carries a file AND quotes an original with an embedded image has to ship
   * both. Empty whenever no html quote ships.
   */
  minted: AttachmentPart[];
  /** The Content-IDs of those parts, for the closure check. */
  mintedCids: string[];
  /**
   * The original's parts this call embedded.
   *
   * Compared by object identity, because that is what the resolution actually returned —
   * a forward walks the same parts afterwards to decide what else to carry, and matching
   * on a Content-ID would re-derive a decision that has already been made.
   */
  embedded: Set<CidPart>;
  /**
   * How many distinct parts the quote's references resolved to, when a shortfall is
   * possible. Undefined on a branch that ships no html quote, where nothing embedded and
   * the parts are accounted for individually instead — see the contract on
   * InlineNoteContext.resolvedPartCount.
   */
  resolvedPartCount?: number;
}

/**
 * Record what a quote or forwarded block did with the original's embedded images, and hand
 * back the parts the draft has to carry for it.
 *
 * Shared by reply and forward because the embed side is identical and the counts have to
 * agree; the two differ only in what becomes of an image that could NOT be embedded, and
 * that difference stays with the caller. A reply drops such an image (a reply is a new
 * message quoting an original, and there is no body to display it), so this records the
 * drop. A forward carries it as a regular attachment, so this records nothing and the
 * forward pools the part alongside the original's other files.
 *
 * On a branch that ships an html quote, a resolved part that did not embed is reported ONLY
 * through the shortfall denominator — never also as a dropped part, which would count one
 * lost image twice.
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
      // The emitted Content-ID: minted fresh or claimed from one survivor, so unique per
      // mapping either way.
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
