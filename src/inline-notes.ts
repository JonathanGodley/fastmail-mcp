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
import type { BlockUnavailableCause } from './body-tokens.js';

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

/** The remedy the pooled sentence ends on when the caller names no other one. */
export const POOLED_REMEDY_RERUN =
  're-run with asAttachment: true for full fidelity, then delete this draft.';

/**
 * Media that could not be embedded and rides the forward as a regular attachment.
 *
 * The REMEDY is a parameter because it is not the same sentence on every tool. On a tool
 * where the forward's format is inferred, re-running as .eml is the only lever the caller
 * has. On one where the caller places the block themselves, the fix is usually to place it
 * in the html part instead — and telling that caller to "re-run with asAttachment: true"
 * sends them at a call the token gate would refuse.
 */
export function noteForwardPooled(
  count: number, names: (string | null | undefined)[], remedy: string = POOLED_REMEDY_RERUN,
): string {
  const listed = describePartNames(names, count);
  return (
    `${count} media part(s) could not be embedded and were attached as regular attachments` +
    `${listed ? `: ${listed}` : ''} — ${remedy}`
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

/*
 * WORDING CONSTRAINT for the two sentences below, and any that join them.
 *
 * Both render on a forward as well as a reply, and both say "the quoted message" there —
 * which reads correctly for a forwarded block, and is what the sibling drop sentences
 * already say. Do NOT give one member of the family a caller-supplied noun while the others
 * keep the fixed wording: the inconsistency between two sentences that can appear in the
 * SAME result is worse than the slight imprecision either one carries alone. If the noun is
 * ever parameterised, parameterise the whole family in one change.
 */

/** Images the quoted message carried as data URIs, which this server does not support. */
export function noteDroppedDataImages(count: number): string {
  return (
    `${count} data:-URI image(s) in the quoted message were dropped (not supported); ` +
    'the rest of the quote was kept.'
  );
}

/**
 * Images the quote could not carry because of HOW they were referenced — a relative or
 * scheme-less path, a protocol-relative URL, an exotic scheme. Distinct from a data: URI
 * (which is content this server declines to re-encode) and from an unmatched embedded-image
 * reference (which named a part that was not there): here the reference form itself is one
 * a quote cannot carry, since it resolves against an origin the new message does not have.
 */
export function noteDroppedUnsupportedImages(count: number): string {
  return (
    `${count} image(s) in the quoted message used a reference form this server cannot carry ` +
    'into a quote and were dropped; the rest of the quote was kept.'
  );
}

/** The caller's own images were embedded in a draft they composed or edited. */
export function noteDraftEmbeds(count: number, bytes: number): string {
  return `This draft embeds ${count} image(s) (${formatSize(bytes)}).`;
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

/**
 * Whether this server can attach anything at all — from local disk or from the account's own
 * content — which changes what a repair can suggest.
 */
export interface AttachmentAvailability {
  attachmentsEnabled: boolean;
}

// The sentence offered instead of "add an attachments item" when this server cannot attach
// anything. Suggesting a repair the server would then refuse is worse than saying why.
const ATTACHMENTS_DISABLED_CLAUSE =
  'Remove the <img> reference (sending attachments is disabled on this server — neither ' +
  'FASTMAIL_ATTACH_DIR nor FASTMAIL_ALLOW_BLOB_ATTACH is set — so no attachments item can supply it).';

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

/**
 * The body authors a NEW reference to one of this server's own identifiers.
 *
 * Scoped to a reference naming no part the draft already carries: a minted identifier is
 * durable now — it survives an edit for as long as the body keeps referencing it — so an
 * image-bearing draft read back and handed straight back is full of these references, and
 * refusing them all would refuse the first edit of every such draft. What stays refused is
 * AUTHORING one, which names an image the draft does not have.
 */
export function rejectReservedCidRef(value: string, surface: 'edit' | 'compose' = 'edit'): string {
  // The DIAGNOSIS differs by surface; the remedy does not. On an edit there is a draft to
  // check the reference against, and what is wrong is that it carries no part under that
  // identifier. On a compose there is no draft at all, so the same clause would describe a
  // thing that does not exist — what is wrong there is simply that the body authored an
  // identifier this server assigns itself.
  const diagnosis = surface === 'edit'
    ? 'and this draft carries no part under it'
    : 'and this server assigns them itself — a body never authors one';
  return (
    `htmlBody references cid "${describePart(value)}", a server-managed identifier for ` +
    `quoted images, ${diagnosis}. A minted identifier survives ` +
    'an edit for as long as your body keeps referencing it, and a reference you drop takes ' +
    'the part with it — but never author a NEW one. To embed your own image, add an ' +
    'attachments item with a cid of your choosing.'
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

// ---------------------------------------------------------------------------
// Body tokens, and the read a body edit has to prove
// ---------------------------------------------------------------------------
//
// These sentences serve BOTH compose and edit, which is why they live here rather than in
// either handler. The two tools treat the same body very differently — draft_email refuses
// what would ship wrong, because the body is wholly the caller's; edit_draft NOTES it,
// because the body may be a foreign one handed back and a refusal keyed on its text could
// be planted by the original's author and would then recur on every edit — so the split
// between a refusal and a note below is deliberate and is not a wording choice.

/** Why a block had nothing to put at a token's position, as one clause of a sentence. */
export const CAUSE_SENTENCE: Record<BlockUnavailableCause, string> = {
  'no-signature': 'the sending identity has no signature configured',
  'no-text-form': 'the signature has no plain-text form (it is images only)',
  'nothing-quotable': 'the original has nothing quotable in any format',
  'nothing-quotable-in-this-form': 'the original has nothing quotable in this part\'s format',
};

/** A token that was placed and had nothing to expand to. Per part, never merged. */
export function noteTokenEmpty(token: string, part: string, cause: BlockUnavailableCause): string {
  return (
    `{{${token}}} in ${part} was removed: ${CAUSE_SENTENCE[cause]}. ` +
    'The rest of that part was stored as written.'
  );
}

/**
 * A body edit arrived with no proof that the caller read the body it replaces.
 *
 * The sentence explains WHY rather than just naming the parameter, because the reason is
 * the whole of the rule: this tool stores what it is handed and preserves nothing, so the
 * hash is the only thing standing between a stale read and a silent overwrite.
 */
export function rejectMissingBodyHash(): string {
  return (
    'This edit writes the draft\'s body, so it needs bodyHash — proof you have read the body ' +
    'you are replacing. edit_draft stores the body you supply byte for byte and keeps nothing ' +
    'of what is there, so an edit written from a stale read overwrites whatever changed in ' +
    'between. Read the draft with get_email (fields: ["bodyText", "bodyHtml", "bodyHash"]) and ' +
    'pass the bodyHash it returns. Metadata-only edits need none.'
  );
}

/** The hash was current when it was issued and is not current now. */
export function rejectStaleBodyHash(): string {
  return (
    'The bodyHash you passed is not this draft\'s current one, so its body changed after the ' +
    'read that issued it (the web UI, another client, or an edit you have already made). ' +
    'Nothing was written. Read the draft again with get_email, re-apply your changes to the ' +
    'body it returns, and pass the bodyHash from that read.'
  );
}

/** expandSignature was passed with nothing for it to expand. */
export function rejectExpandSignatureWithoutToken(wroteAnyBody: boolean): string {
  return wroteAnyBody
    ? 'expandSignature: true was passed but the body you supplied carries no {{signature}}, so ' +
      'there is nothing to expand. Place {{signature}} where the sign-off goes — above any ' +
      'quoted or forwarded history — or drop the flag and the body is stored as written.'
    : 'expandSignature: true was passed but this edit writes no body, so there is nothing to ' +
      'expand it in. Supply textBody or htmlBody carrying {{signature}}, or drop the flag.';
}

/**
 * A flagged edit's written part carries more than one `{{signature}}`.
 *
 * Refused rather than noted, and that is the one text-keyed refusal on this tool: passing
 * the flag is the caller claiming the written part as its own, so the compose-style refusal
 * applies. The escape is named because the extra token is usually one the caller did not
 * write — a literal `{{signature}}` inside the quoted original it handed back.
 */
export function rejectRepeatedSignatureToken(part: string, count: number): string {
  return (
    `${part} carries ${count} {{signature}} tokens, and expandSignature would expand every one ` +
    'of them. Keep the one where the sign-off goes and write the others as \\{{signature}} to ' +
    'store them as text. A body you read back from a reply can carry one inside the quoted ' +
    'original; that is the one to escape, and the escape is needed in each expandSignature call.'
  );
}

/** The written part carries a `{{signature}}` the stored part did not, and no flag came with it. */
export function noteSignatureTokenStored(part: string, count: number): string {
  const plural = count === 1 ? '' : 's';
  return (
    `${part} carries ${count} {{signature}} token${plural} the stored body did not, and this ` +
    'edit stored the body as written. Pass expandSignature: true to expand it.'
  );
}

/** `{{quote}}` / `{{forward}}` reached the edit path, where neither expands. */
export function noteHistoryTokenStored(tokens: string[]): string {
  const named = tokens.map((t) => `{{${t}}}`).join(' and ');
  const verb = tokens.length > 1 ? 'were' : 'was';
  return (
    `${named} ${verb} stored as written: those tokens expand on draft_email only, where the ` +
    'block is built from the message being replied to or forwarded. edit_draft stores whatever ' +
    'history you hand back to it.'
  );
}

/** A spelling close enough to a token to be worth naming, stored as the caller wrote it. */
export function noteNearMissToken(text: string, token: string): string {
  return (
    `The body you supplied carries "${describePart(text)}", which is not a token and was stored ` +
    `as written. The token is exactly {{${token}}}.`
  );
}

/** An escaped spelling this edit is about to ship with its backslash intact. */
export function noteEscapedTokenShips(text: string): string {
  return (
    `The body you supplied carries "${describePart(text)}", an escaped token spelling the stored ` +
    'body did not. An edit without expandSignature stores the body byte for byte, so the ' +
    'backslash ships with it; the escape is only needed in an expandSignature call.'
  );
}

/** One supplied part took the sign-off and the other carried no token to take it. */
export function noteSignatureExpandedInOnePart(withToken: string, without: string): string {
  return (
    `{{signature}} expanded in ${withToken}; ${without} carries none, so it was stored without a ` +
    'sign-off and a recipient reading that alternative sees none.'
  );
}

/** An html-alone edit drops whatever plain-text part the draft was storing. */
export function noteDiscardedTextPart(): string {
  return (
    'This edit wrote htmlBody alone, so the draft\'s stored plain-text part was discarded and a ' +
    'fresh fallback derived from your html. If that part was hand-written, supply it as textBody ' +
    'alongside htmlBody.'
  );
}

// Why this edit returns no bodyHash. Each names the caller's way out, because the
// alternative — omitting the field and saying nothing — is the silent-drop failure: a
// caller that got a hash from the last edit and none from this one has no way to tell a
// withheld hash from a forgotten one.
//
// The way out is a re-read for every case EXCEPT a degraded stored body, where a re-read
// issues no hash either (get_email withholds on the same flags) and the remedy is to
// recreate the draft. A note that sends the caller to a read that cannot answer is a
// non-terminating remedy, which is no remedy at all.
//
// A hash is returned ONLY over a re-read of the stored parts after the write, never over
// the bytes the call sent: a store may normalise what it stores, and a hash over the sent
// bytes would refuse the next edit as stale on a draft nobody had touched.

export const NOTE_BODY_HASH_AFTER_EXPANSION =
  'this edit expanded {{signature}}, so the stored body carries a block you have not read. ' +
  'Read the draft with get_email to get a bodyHash for the next body edit.';

export const NOTE_BODY_HASH_DERIVED_PART =
  'this edit left the draft\'s plain-text part standing as its body, and that part is derived ' +
  'from html rather than written by you. Read the draft with get_email to get a bodyHash for ' +
  'the next body edit.';

export const NOTE_BODY_HASH_DEGRADED =
  'the saved draft was re-read to compute one, but the server flagged its stored body as ' +
  'truncated or as having encoding problems, so no read can prove you saw it whole. ' +
  'Recreate the draft rather than editing its body again.';

export function noteBodyHashUnreadable(reason: string): string {
  return (
    `re-reading the saved draft to compute one failed (${reason}). Read the draft with ` +
    'get_email to get a bodyHash.'
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
export type RefCounter = 'unresolvedRefs' | 'droppedDataImages' | 'droppedUnsupportedImages';

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
  droppedUnsupportedImages: number;
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
    droppedUnsupportedImages: 0,
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
      droppedUnsupportedImages: this.refCounts.droppedUnsupportedImages,
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
   *
   * CALLER CONTRACT: never pass a count that includes a part this same call also recorded
   * as removed, dropped or pooled. The shortfall this produces already speaks for every
   * resolved part that did not embed, so a part counted here AND given its own outcome is
   * reported twice — one lost image read as two. Leave it undefined on a branch that
   * accounts for those parts individually.
   */
  resolvedPartCount?: number;
  /** Names the block a removal emptied, in the wording the surrounding edit uses. */
  keepNoun?: string;
  /** False when the caller asked for the original's attachments to be left behind. */
  includeOriginalAttachments?: boolean;
  /** Set when the body contained reference-shaped text this server could not act on. */
  unparsableCidText?: boolean;
  /**
   * How to end the pooled sentence, for a tool whose caller has a better lever than
   * re-running as .eml. Defaults to POOLED_REMEDY_RERUN. See noteForwardPooled.
   */
  pooledRemedy?: string;
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

  if (tally.pooled > 0) {
    notes.push(noteForwardPooled(tally.pooled, tally.pooledNames, context.pooledRemedy));
  }

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

  if (tally.droppedUnsupportedImages > 0) {
    notes.push(noteDroppedUnsupportedImages(tally.droppedUnsupportedImages));
  }

  if (context.unparsableCidText) notes.push(noteUnparsableCidText());

  return notes;
}
