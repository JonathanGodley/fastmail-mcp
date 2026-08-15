// The wording this server uses when it embeds, carries, demotes, removes or refuses an
// embedded (cid:) image, plus the bookkeeping that decides which of those sentences is
// true at the end of a call (#13).
//
// Everything a caller says about embedded images lives here rather than at the call sites,
// for two reasons. The same sentence is emitted from several tools — compose, reply,
// forward and edit all report an embed — so a copy per tool would drift. And the counts in
// those sentences have to agree with each other: a part that was going to be embedded and
// then got removed must be reported once, as removed, not twice. The ledger below is what
// makes that true by construction.
import { describePart, isAuthorableCid } from './inline-images.js';

// ---------------------------------------------------------------------------
// Sizes
// ---------------------------------------------------------------------------

const KB = 1024;
const MB = 1024 * 1024;

/**
 * A byte count as a person reads it: kilobytes below a megabyte, one decimal place at or
 * above.
 *
 * The unit split is not cosmetic. A signature logo is tens of kilobytes, so reporting
 * everything in megabytes would tell the user their draft embeds "0 MB" of images. The
 * floor of one kilobyte for any non-zero size is the same guard one step down: a tracking
 * pixel is a few hundred bytes, and "0 KB" reads as nothing at all when something really
 * was carried.
 */
export function formatSize(bytes: number): string {
  const n = Number.isFinite(bytes) && bytes > 0 ? bytes : 0;
  if (n >= MB) return `${(n / MB).toFixed(1)} MB`;
  if (n === 0) return '0 KB';
  return `${Math.max(1, Math.round(n / KB))} KB`;
}

// How many filenames a note lists before it summarizes the rest. Enough to recognize what
// was affected, few enough that a message carrying fifty parts does not produce a wall of
// text — and every name is rendered as quoted data, since a filename is sender-controlled.
const MAX_NAMED_PARTS = 3;

/**
 * Render a list of part names as quoted data, summarizing beyond the display cap.
 *
 * `total` is how many parts the sentence is really about, which is not always how many
 * names there are: a part can arrive with no filename at all, and the summary has to
 * account for it rather than quietly shrinking the total the reader was given.
 */
export function describePartNames(
  names: (string | null | undefined)[],
  total = names.length,
): string {
  const usable = names.filter((n): n is string => typeof n === 'string' && n !== '');
  const shownCount = Math.min(usable.length, MAX_NAMED_PARTS);
  const shown = usable.slice(0, shownCount).map((n) => `"${describePart(n)}"`).join(', ');
  const rest = Math.max(0, total - shownCount);
  if (!shown) return '';
  return rest > 0 ? `${shown} …and ${rest} more` : shown;
}

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

/** Every image referenced by the quote was embedded. */
export function noteEmbeddedFromQuote(count: number, bytes: number): string {
  return `This draft embeds ${count} image(s) from the quoted message (${formatSize(bytes)}).`;
}

/**
 * Some of the images the quote references were embedded and some were not.
 *
 * The two numbers count different things and are never added together: `resolvedParts` is
 * how many distinct parts the references landed on, and the separate skipped-references
 * sentence counts references that landed on nothing at all. Keeping them apart is what
 * stops one lost image being read as two.
 */
export function noteEmbeddedPartially(
  embedded: number,
  resolvedParts: number,
  bytes: number,
): string {
  const shortfall = Math.max(0, resolvedParts - embedded);
  return (
    `Embedded ${embedded} of ${resolvedParts} image part(s) referenced by the quote ` +
    `(${formatSize(bytes)} embedded); ${shortfall} could not be embedded and are not part of this draft.`
  );
}

/** References in the quote that matched no part at all. */
export function noteSkippedReferences(count: number): string {
  return `${count} reference(s) matched no part and were skipped.`;
}

/** Every image the original's body displayed was carried into the forward. */
export function noteEmbeddedFromOriginal(count: number, bytes: number): string {
  return `This draft embeds ${count} image(s) from the original (${formatSize(bytes)}).`;
}

/** References in the forwarded original that matched no part. */
export function noteForwardUnresolvedReferences(count: number): string {
  return (
    `${count} image reference(s) in the original's body had no matching part; ` +
    'nothing was carried for them.'
  );
}

/** Media that could not be embedded and rides the forward as a regular attachment. */
export function noteForwardPooled(count: number, names: (string | null | undefined)[]): string {
  const listed = describePartNames(names, count);
  return (
    `${count} media part(s) could not be embedded and were attached as regular attachments` +
    `${listed ? `: ${listed}` : ''} — re-run with asAttachment: true for full fidelity, ` +
    'then delete this draft.'
  );
}

/**
 * Attachments the caller asked not to carry.
 *
 * The counts cover the EXCLUDED set only — never the body-embedded images, which are body
 * content and are carried regardless. The second sentence is emitted only when at least one
 * such image really was carried, because on an ordinary forward of a message with no
 * embedded images it would assert something that did not happen.
 */
export function noteAttachmentsExcluded(
  count: number,
  imageCount: number,
  carriedEmbedded: boolean,
): string {
  const first =
    `${count} attachment(s), including ${imageCount} image(s), were not included ` +
    'because includeOriginalAttachments is false.';
  return carriedEmbedded
    ? `${first} Body-embedded images were still carried — they are part of the message body.`
    : first;
}

/**
 * Images taken off the draft because the body that displayed them is gone.
 *
 * `keepNoun` names the block that was dropped, in the wording the surrounding edit already
 * uses for it ("the quote", "the forwarded block").
 */
export function noteRemovedEmbedded(count: number, keepNoun: string): string {
  return `Removed ${count} image(s) that were embedded in ${keepNoun}.`;
}

/**
 * The caller's images could not be embedded, so they ride as ordinary attachments.
 *
 * The reason is stated in the one form that is true of every route here, because there are
 * several and the note has no way to tell them apart: the message ships no html body at
 * all; it ships one that references some other identifier; or an edit's new body stopped
 * referencing a part the draft already carried. Naming only the first would state a false
 * reason for the others, and the promise made on the attachments parameter is an honest
 * account of what happened to the file, not a guess at why.
 */
export function noteDegradedToAttachments(count: number): string {
  return `${count} of your image(s) became regular attachments (nothing in the body displays them).`;
}

/**
 * Images the quote referenced that could not be carried at all.
 *
 * A reply drops what it cannot embed rather than attaching it, because a reply is a new
 * message that quotes an original — mail clients do not attach a quoted message's images
 * to a reply that has no body to display them. Dropping is therefore the right outcome,
 * but it is never a silent one: the images the reader saw in the original are not in the
 * draft, and only this sentence says so.
 */
export function noteDroppedQuoteImages(count: number): string {
  return `${count} image(s) from the quoted message were dropped and are not part of this draft.`;
}

/** Images the quoted message carried as data URIs, which this server does not support. */
export function noteDroppedDataImages(count: number): string {
  return (
    `${count} data:-URI image(s) in the quoted message were dropped (not supported); ` +
    'the rest of the quote was kept.'
  );
}

/** The caller's own images were embedded in a draft they composed or edited. */
export function noteDraftEmbeds(count: number, bytes: number): string {
  return `This draft embeds ${count} image(s) (${formatSize(bytes)}).`;
}

/** An attachment wipe on an edit that also rebuilt the quote. */
export function noteClearedAndReEmbedded(cleared: number, reEmbedded: number): string {
  return `Cleared ${cleared} attachment(s); the kept quote re-embedded ${reEmbedded} image(s).`;
}

/** What a sent message actually carried, reported back on the send. */
export function noteSentWithEmbedded(count: number, bytes: number): string {
  return `Sent with ${count} embedded image(s) (${formatSize(bytes)}).`;
}

/**
 * The saved draft could not be re-read, so what it actually carries is unconfirmed.
 *
 * Never an error: the edit itself succeeded and the draft exists. This says only that the
 * confirmation step did not run, which matters because embedding is the one outcome the
 * caller cannot check from the result text alone.
 */
export function noteEmbedUnconfirmed(): string {
  return (
    'The draft was saved, but this server could not re-read it to confirm the embedded ' +
    'image(s) are attached. Open the draft to check the images appear.'
  );
}

/** The re-read found the saved draft short of images this call attached to it. */
export function noteEmbedMissingAfterSave(count: number): string {
  return (
    `${count} embedded image(s) this call attached were not found on the saved draft. ` +
    'Open the draft to check how it renders.'
  );
}

/**
 * Text that looks like an embedded-image reference but is not one this server can act on.
 *
 * A note, never an error: the text may be someone writing about this feature, a pasted MIME
 * fragment, or this server's own message quoted back, and refusing to compose mail over any
 * of those would be worse than saying plainly that it was left alone.
 */
export function noteUnparsableCidText(): string {
  return (
    "htmlBody contains text that looks like an embedded-image (cid:) reference this server " +
    "can't parse as an image; it was left as-is."
  );
}

// ---------------------------------------------------------------------------
// Rejects
// ---------------------------------------------------------------------------

/** Whether this server can attach files at all, which changes what a repair can suggest. */
export interface AttachmentAvailability {
  attachmentsEnabled: boolean;
}

// The sentence offered instead of "add an attachments item" when this server cannot attach
// anything. Suggesting a repair the server would then refuse is worse than saying why.
const ATTACHMENTS_DISABLED_CLAUSE =
  'Remove the <img> reference (sending attachments is disabled on this server — ' +
  'FASTMAIL_ATTACH_DIR is not set — so no attachments item can supply it).';

/**
 * The body references an embedded image nothing supplies.
 *
 * The repair clause is conditional in two ways. It is offered only when the value is one a
 * caller could actually author, because pointing at a repair the vet would then reject is
 * a dead end — and a foreign identifier, which always contains an `@`, is exactly that
 * population. And an authorable value is echoed RAW in that clause: it is drawn from a
 * narrow allowlist and bounded in length, so it is safe by construction, and a mangled echo
 * would break the copy-and-paste the clause exists for. Any other value renders through the
 * display treatment, for reading only.
 */
export function rejectDanglingCidRef(value: string, availability: AttachmentAvailability): string {
  const authorable = isAuthorableCid(value);
  const shown = authorable ? value : describePart(value);
  const lead = `htmlBody references cid "${shown}" but no attachment supplies it.`;
  if (!availability.attachmentsEnabled) return `${lead} ${ATTACHMENTS_DISABLED_CLAUSE}`;
  return authorable
    ? `${lead} Remove the <img> reference, or add an attachments item with cid: "${value}".`
    : `${lead} Remove the <img> reference, or recreate the draft without it.`;
}

/** The body references one of this server's own identifiers, which callers must not author. */
export function rejectReservedCidRef(value: string): string {
  return (
    `htmlBody references cid "${describePart(value)}", a server-managed identifier for ` +
    'quoted images. These identifiers are reused across edits but regenerated when the ' +
    'quote is dropped and re-added — never reference them directly. To embed your own ' +
    'image, add an attachments item with a cid of your choosing.'
  );
}

/** A removal would leave the body pointing at an image that is no longer there. */
export function rejectRemovalDanglingRef(value: string): string {
  return (
    `removeAttachments would remove an image the draft's body still references ` +
    `(cid "${describePart(value)}"). Remove the <img> reference in the same call, ` +
    'or keep the attachment.'
  );
}

/**
 * A removal targets an image the kept quote supplies.
 *
 * The recovery names a body write on purpose: dropping the quote on an edit that writes no
 * body rewrites nothing, so suggesting the flag alone would steer straight into a second
 * refusal.
 */
export function rejectRemovalOfQuoteCarriedPart(): string {
  return (
    'That attachment is embedded by the kept quote; the rebuilt quote would re-embed it. ' +
    'Use noQuote with a replacement body to drop the quote and its images instead.'
  );
}

/** Clearing the attachments would leave the body pointing at images that are gone. */
export function rejectClearAttachmentsDanglingRefs(): string {
  return (
    'clearFields: ["attachments"] would strip image(s) the surviving body still references. ' +
    'Rewrite or clear that body in the same call, or keep the attachments.'
  );
}

/**
 * A caller asked for an identifier this server will not author.
 *
 * Names the real item index, matching every other per-item refusal on the attachments
 * parameter, and shows the value as quoted data — this refusal is raised before any
 * redaction layer, so the echo has to be bounded here. The example is spelled the way the
 * parameter takes it, so the fix is a copy away.
 */
export function rejectUnusableCid(index: number, value: unknown): string {
  return (
    `attachments[${index}].cid "${describePart(value)}" isn't usable here. This server ` +
    'accepts a simple token (up to 64 characters) of letters, digits, dot, dash, ' +
    'underscore (e.g. cid: "logo").'
  );
}

/** Two attachments on the draft would share one identifier. */
export function rejectCidCollisionOnDraft(value: string): string {
  return (
    `cid "${describePart(value)}" is already used by another attachment on this draft; ` +
    'each embedded image needs a distinct cid.'
  );
}

/** Two or more items in this call share one identifier. */
export function rejectCidCollisionInCall(count: number, value: string): string {
  return (
    `${count} attachments items share cid "${describePart(value)}"; ` +
    'each embedded image needs a distinct cid.'
  );
}

// The recreate recipe both draft-level refusals end with. It routes a reply or forward
// through the tools that rebuild the threading, because recreating one of those with the
// plain compose tool splits the conversation, and it deliberately stops short of promising
// the recreated draft will re-embed the images — with attachments disabled it cannot, and
// the advice has to stay true either way.
const RECREATE_RECIPE_WITH_THREADING =
  'Recreate it — with reply_email or forward_email if it is a reply or forward ' +
  '(read the draft\'s In-Reply-To or X-Forwarded-Message-Id via get_email, then find the ' +
  'original with search_emails using the bare id, without angle brackets), otherwise ' +
  'create_draft — then delete this one.';

/** A stored identifier this server cannot reproduce faithfully on a recreated draft. */
export function rejectUnrecreatableCid(value: string): string {
  return (
    'This draft has an attachment whose embedded-image identifier ' +
    `(Content-ID "${describePart(value)}") this server cannot safely re-create. ` +
    'This server does not edit drafts in that state. Recreate it — with reply_email or ' +
    'forward_email if it is a reply or forward (that preserves the conversation threading; ' +
    "read the draft's In-Reply-To or X-Forwarded-Message-Id via get_email, then find the " +
    'original with search_emails using the bare id, without angle brackets), otherwise ' +
    'create_draft — then delete this one; or edit it in the mail client that created it.'
  );
}

/**
 * The draft's stored body already references images that are not on it.
 *
 * Deliberately checked after the caller's own edit is merged in, so an edit that replaces
 * the body wholesale and eliminates the dangling references succeeds — the "unless" clause
 * in the sentence says so, which keeps the cheapest repair from being foreclosed. The
 * add-clause appears only when every dangling value is one a caller could author, for the
 * same reason the dangling-reference refusal gates its own repair clause.
 */
export function rejectBrokenDraft(
  values: string[],
  availability: AttachmentAvailability,
): string {
  const listed = values.map((v) => `"${describePart(v)}"`).join(', ');
  const base =
    `This draft's stored body references image identifier(s) with no matching attachment ` +
    `(${listed}). This server won't edit its body in that state unless the edit resolves ` +
    'the missing reference(s) — metadata and attachment edits still work. ' +
    RECREATE_RECIPE_WITH_THREADING;

  const authorable = values.length > 0 && values.every((v) => isAuthorableCid(v));
  if (!authorable || !availability.attachmentsEnabled) return base;

  // One clause covers however many values there are: repeating the whole sentence per
  // identifier would bury the recipe above it.
  return values.length === 1
    ? `${base} Or add an attachments item with cid "${values[0]}" to supply the missing image.`
    : `${base} Or add attachments items with cid ${values.map((v) => `"${v}"`).join(', ')} ` +
      'to supply the missing images.';
}

// The short recreate recipe the body-shape refusals end with. It names the threading tools
// for the same reason the longer recipe does — recreating a reply or forward with the plain
// compose tool splits the conversation — but it needs no header-reading detour: these
// refusals fire on the draft's own body shape, which the caller can see in get_email.
// Unterminated on purpose: one caller ends the sentence, the other appends a pointer first.
const RECREATE_RECIPE =
  'Recreate it (reply_email/forward_email for replies and forwards, create_draft ' +
  'otherwise), then delete this one';

/**
 * The draft's body carries a part the recreate cannot reproduce.
 *
 * Editing rebuilds the message from its parts, so a part this server cannot re-reference
 * would be dropped or mangled by an edit that claims to have preserved it. Refusing is the
 * loud alternative. Both interpolated values come from the message itself, so both render
 * as quoted data.
 */
export function rejectUncarriableBodyPart(
  name: string | null | undefined,
  type: string | null | undefined,
  isMedia: boolean,
): string {
  const noun = isMedia ? 'a media part' : 'a part';
  return (
    `This draft's body contains ${noun} this server cannot carry ` +
    `(part "${describePart(name)}", content type "${describePart(type)}"). ${RECREATE_RECIPE}.`
  );
}

/**
 * The draft's body puts two parts of one text type around something else.
 *
 * The parts are individually carriable — what cannot be preserved is their ORDER around the
 * content between them, which this server's flat rebuild has no way to express. Saying
 * "cannot carry" here would be false, so this refusal has its own wording (see issue #85).
 */
export function rejectInterleavedTextParts(): string {
  return (
    "This draft's body interleaves multiple text parts of the same type (a layout this " +
    `server cannot preserve). ${RECREATE_RECIPE} (see issue #85).`
  );
}

/**
 * A note authored alongside a quote references an image nothing supplies.
 *
 * Distinct from the plain dangling-reference refusal because a quote DOES exist here: the
 * sentence has to say that quoted images arrive on their own, or the caller reads it as an
 * instruction to author references for them.
 */
export function rejectNoteCidRef(value: string, availability: AttachmentAvailability): string {
  const authorable = isAuthorableCid(value);
  const shown = authorable ? value : describePart(value);
  const lead = `htmlBody (your note) references cid "${shown}", which no attachments item supplies.`;
  if (!availability.attachmentsEnabled) {
    return `${lead} Quoted images appear inside the quote automatically. ${ATTACHMENTS_DISABLED_CLAUSE}`;
  }
  return authorable
    ? `${lead} Quoted images appear inside the quote automatically; to embed your own image, ` +
      `add an attachments item with cid: "${value}" — otherwise remove the reference.`
    : `${lead} Quoted images appear inside the quote automatically; to embed your own image, ` +
      'add an attachments item with a cid of your choosing — otherwise remove the reference.';
}

// ---------------------------------------------------------------------------
// The ledger: one outcome per part, counted once, at the end
// ---------------------------------------------------------------------------

/** What ultimately became of one part. Exactly one of these is true per part. */
export type PartOutcome =
  /** Displayed by the message body. */
  | 'embedded'
  /** Carried as a regular attachment because it could not be embedded. */
  | 'pooled'
  /** One of the caller's own images, carried as a regular attachment instead. */
  | 'degraded'
  /** Left behind at the caller's request. */
  | 'notIncluded'
  /** Could not be carried at all. */
  | 'dropped'
  /** Taken off a draft it was already on. */
  | 'removed'
  /** An ordinary attachment. Counted, but nothing is said about it. */
  | 'attached';

export interface PartRecord {
  /** Whatever identifies this part uniquely within the call — a blob id, a part id. */
  key: string;
  outcome: PartOutcome;
  bytes?: number;
  name?: string | null;
  /** Whether the part is an image, for the counts that break images out separately. */
  isImage?: boolean;
}

/** Reference-level losses, which count references rather than parts. */
export type RefCounter = 'unresolvedRefs' | 'droppedDataImages';

export interface NoteTally {
  embedded: number;
  embeddedBytes: number;
  pooled: number;
  pooledNames: (string | null | undefined)[];
  degraded: number;
  notIncluded: number;
  notIncludedImages: number;
  dropped: number;
  removed: number;
  attached: number;
  unresolvedRefs: number;
  droppedDataImages: number;
}

/**
 * Accumulates what a call did to each part, then reports it once at the end.
 *
 * The point is that recording is a CANDIDATE and only the final state is counted. A part
 * can be recorded as embedded early in an assembly and then removed later, and the removal
 * simply replaces the earlier record for that key — so the notes say it was removed, once,
 * instead of claiming it was both embedded and removed. Every count comes from this single
 * terminal read, which is why no note can outlive the disposition it describes.
 */
export class InlineNoteLedger {
  private readonly parts = new Map<string, PartRecord>();
  private readonly refCounts: Record<RefCounter, number> = {
    unresolvedRefs: 0,
    droppedDataImages: 0,
  };

  /** Record what happened to a part. A later record for the same key replaces the earlier. */
  record(record: PartRecord): void {
    this.parts.set(record.key, { ...record });
  }

  /** Count a reference-level loss, which has no part to attach an outcome to. */
  countRefs(counter: RefCounter, n = 1): void {
    if (n > 0) this.refCounts[counter] += n;
  }

  /** The final state, read once. */
  tally(): NoteTally {
    const t: NoteTally = {
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
      unresolvedRefs: this.refCounts.unresolvedRefs,
      droppedDataImages: this.refCounts.droppedDataImages,
    };
    for (const record of this.parts.values()) {
      const bytes = typeof record.bytes === 'number' && record.bytes > 0 ? record.bytes : 0;
      switch (record.outcome) {
        case 'embedded':
          t.embedded++;
          t.embeddedBytes += bytes;
          break;
        case 'pooled':
          t.pooled++;
          t.pooledNames.push(record.name);
          break;
        case 'degraded': t.degraded++; break;
        case 'notIncluded':
          t.notIncluded++;
          if (record.isImage) t.notIncludedImages++;
          break;
        case 'dropped': t.dropped++; break;
        case 'removed': t.removed++; break;
        case 'attached': t.attached++; break;
      }
    }
    return t;
  }

  /** The notes this call should emit, composed once from the final state. */
  emit(context: InlineNoteContext): string[] {
    return emitInlineNotes(this.tally(), context);
  }
}

export interface InlineNoteContext {
  /** Which tool is speaking; decides how an embed is described. */
  surface: 'reply' | 'forward' | 'draft' | 'send';
  /**
   * How many distinct parts the body's references resolved to — the denominator a partial
   * embed reports against. Defaults to the embedded count, i.e. no shortfall.
   */
  resolvedPartCount?: number;
  /** Names the block a removal emptied, in the wording the surrounding edit uses. */
  keepNoun?: string;
  /** False when the caller asked for the original's attachments to be left behind. */
  includeOriginalAttachments?: boolean;
  /** Set when this call wiped the draft's attachments and rebuilt the quote. */
  clearedAttachmentCount?: number;
  reEmbeddedCount?: number;
  /** Set when the body contained reference-shaped text this server could not act on. */
  unparsableCidText?: boolean;
}

/**
 * Turn a final tally into the sentences a caller sees, in a fixed order: what the draft
 * carries, then what it could not carry, then what came off it.
 *
 * Every part is described by exactly one of these sentences. Parts left behind at the
 * caller's request are covered by the exclusion sentence alone, and the pooled sentence is
 * only reached by parts that were actually carried, so the two cannot double-count. What a
 * part's outcome cost the reader is always said: the only outcome that stays silent is an
 * ordinary attachment riding along, which is not a loss and needs no sentence.
 */
export function emitInlineNotes(tally: NoteTally, context: InlineNoteContext): string[] {
  const notes: string[] = [];
  const resolved = context.resolvedPartCount ?? tally.embedded;

  if (context.surface === 'reply') {
    // A shortfall is reported even when NOTHING embedded. A quote referencing three images
    // that all turn out to be unusable is the case where the reader most needs to be told,
    // and gating the sentence on a successful embed would make exactly that case silent.
    const shortfall = resolved > tally.embedded;
    if (tally.embedded > 0 || shortfall) {
      notes.push(
        shortfall
          ? noteEmbeddedPartially(tally.embedded, resolved, tally.embeddedBytes)
          : noteEmbeddedFromQuote(tally.embedded, tally.embeddedBytes),
      );
    }
  } else if (tally.embedded > 0) {
    if (context.surface === 'forward') {
      notes.push(noteEmbeddedFromOriginal(tally.embedded, tally.embeddedBytes));
    } else if (context.surface === 'send') {
      notes.push(noteSentWithEmbedded(tally.embedded, tally.embeddedBytes));
    } else {
      notes.push(noteDraftEmbeds(tally.embedded, tally.embeddedBytes));
    }
  }

  if (tally.unresolvedRefs > 0) {
    notes.push(
      context.surface === 'forward'
        ? noteForwardUnresolvedReferences(tally.unresolvedRefs)
        : noteSkippedReferences(tally.unresolvedRefs),
    );
  }

  if (tally.pooled > 0) notes.push(noteForwardPooled(tally.pooled, tally.pooledNames));

  if (tally.notIncluded > 0) {
    notes.push(
      noteAttachmentsExcluded(tally.notIncluded, tally.notIncludedImages, tally.embedded > 0),
    );
  }

  if (tally.removed > 0) {
    notes.push(noteRemovedEmbedded(tally.removed, context.keepNoun ?? 'the quote'));
  }

  if (tally.degraded > 0) notes.push(noteDegradedToAttachments(tally.degraded));

  if (tally.dropped > 0) notes.push(noteDroppedQuoteImages(tally.dropped));

  if (tally.droppedDataImages > 0) notes.push(noteDroppedDataImages(tally.droppedDataImages));

  if (context.clearedAttachmentCount) {
    notes.push(
      noteClearedAndReEmbedded(context.clearedAttachmentCount, context.reEmbeddedCount ?? 0),
    );
  }

  if (context.unparsableCidText) notes.push(noteUnparsableCidText());

  return notes;
}
