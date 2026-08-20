import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { isBlank, assertBodyInputs } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import { applySignature, buildForwardBodies, emptyQuoteImages } from './reply-quote.js';
import type { QuoteImageOutcome } from './reply-quote.js';
import { resolveSignature } from './identity.js';
import type { ResolvedSignature } from './identity.js';
import {
  DEFAULT_INLINE_CONTEXT, planAuthoredInlineImages, recordQuoteImages,
  reportAuthoredInlineImages,
} from './compose-inline.js';
import type { AuthoredInlineContext, AuthoredInlinePlan } from './compose-inline.js';
import { buildUnionParts, checkInlineClosure, isImageType } from './inline-images.js';
import type { CidPart } from './inline-images.js';
import { InlineNoteLedger } from './inline-notes.js';
import type { AttachmentPart, UploadAttachmentsOptions } from './jmap-client.js';

// Parameters passed to createDraft for a forward (matches its input shape).
export interface ForwardParams {
  to: string[];
  cc?: string[];
  bcc?: string[];
  from?: string;
  subject: string;
  textBody?: string;
  htmlBody?: string;
  replyTo?: string[];
  // The forwarded original's Message-ID, recorded as X-Forwarded-Message-Id on the
  // draft (no In-Reply-To/References — a forward starts a new conversation). Recorded
  // on BOTH forward shapes, including asAttachment; omitted only when the original
  // has no settable Message-ID.
  forwardedMessageId?: string[];
  // The JMAP id of the exact stored instance being forwarded, recorded on the draft
  // as X-Fastmail-MCP-Source-Id so send_draft can mark precisely that copy forwarded
  // — the Message-ID above names the message, which can exist as several stored
  // copies. Recorded on both forward shapes.
  sourceEmailId?: string;
  attachments?: AttachmentPart[];
}

// Fastmail validates header:…:asMessageIds values on Email/set (probed live
// 2026-07-05): embedded CR/LF and non-ASCII are REJECTED (failing the whole create),
// and embedded angle brackets round-trip MANGLED (split into two ids). The value
// comes verbatim from the forwarded — attacker-controlled — message, so pre-vet it
// and treat a malformed id as absent: the forward still works, and the edit guard's
// body markers still gate; only the recorded-source affordance is lost. Exported for
// edit_draft's forward keep path, which re-points the recorded source the same way.
export function isSettableMessageId(id: unknown): id is string {
  return (
    typeof id === 'string' &&
    id.length > 0 &&
    id.length <= 998 && // RFC 5322 line limit; anything longer is garbage, not an id
    /^[\x21-\x7e]+$/.test(id) && // printable ASCII only — no spaces, controls, or non-ASCII
    !id.includes('<') &&
    !id.includes('>')
  );
}

// Filename for the attached .eml, derived from the ORIGINAL's subject (the name
// describes the attached message, not the new wrapper — a caller subject override
// does not affect it). The value is consumed as a SAVE NAME by receiving clients,
// so strip control chars (\p{Cc} covers C0 and C1 incl. U+0085), Unicode
// format/bidi controls (\p{Cf}, e.g. U+202E right-to-left override), path
// separators and the Windows drive/ADS colon, and leading dots; cap the length.
// Windows reserved device names (CON, NUL, …) deliberately survive as e.g.
// "CON.eml" — a save-time nuisance the receiving client handles, same
// receiver-sanitizes posture as carried attachment names (docs/security-model.md).
export function sanitizeEmlFilename(subject: string | null | undefined): string {
  const stripped = (subject ?? '')
    .replace(/[\p{Cc}\p{Cf}]/gu, '')
    .replace(/[/\\:]/g, '_')
    .replace(/^\.+/, '')
    .trim();
  // Cap by CODE POINTS, not UTF-16 units — a unit slice can split a surrogate pair
  // and leave a lone surrogate, invalid on the wire.
  const cleaned = [...stripped].slice(0, 80).join('').trim();
  return `${cleaned || 'forwarded-message'}.eml`;
}

// The carry whitelist for re-referencing an existing part on a new Email — the same
// field set edit_draft's carry uses: caller-meaningful fields only, never the
// server-set partId/size.
//
// The original's Content-ID is deliberately NOT carried. Every part that reaches this
// function rides the forward as a regular attachment, so nothing in the new message
// references it: the forwarded block's own image references are rewritten to the
// identifiers this server mints for the parts it embeds, and a part that could not be
// embedded is not referenced at all. Carrying the identifier anyway would be wrong twice
// over — it would claim an identity no body uses, and an identifier of this server's own
// shape arriving on someone else's message would then be classified as server-managed and
// removed from a later edit under an explanation that does not apply to it. Disposition is
// normalized here for the same reason: some clients mark ordinary attachments inline, and
// an inline marking with no body behind it misdescribes what the draft carries.
function carryPart(a: any): AttachmentPart {
  const part: AttachmentPart = { blobId: a.blobId, type: a.type };
  if (a.name != null) part.name = a.name;
  if (a.disposition != null) part.disposition = a.disposition === 'inline' ? 'attachment' : a.disposition;
  return part;
}

/** What became of the original's parts, beyond the ones the forwarded block embeds. */
export interface ForwardCarry {
  // Media the forwarded block could not display, riding along as regular attachments
  // instead. Reported loudly: the reader sees a file where the original showed an image,
  // and asAttachment is the lossless alternative.
  pooled: CidPart[];
  // The original's ordinary files, carried as they were. Counted, but unremarkable.
  attached: CidPart[];
  // Parts left behind because includeOriginalAttachments is false. Never includes the
  // images the block embeds — those are body content and are carried regardless.
  notIncluded: CidPart[];
}

export interface BuiltForward {
  asAttachment: boolean;
  forwardParams: ForwardParams;
  // What the caller's own note asks this server to embed, and whether the forward will
  // actually be able to display it (#13).
  inlinePlan: AuthoredInlinePlan;
  // What the forwarded block did with the images the original's body displayed (#13).
  quoteImages: QuoteImageOutcome;
  // What became of every other part of the original.
  carry: ForwardCarry;
}

// Whether a part is CONTENT of the original's body rather than a file attached to it.
// Two signals, either of which is enough: the server routed the part into a body list
// (RFC 8621 §4.1.4, which is where an embedded image lands for some MIME shapes), or the
// part declares itself inline and carries a Content-ID for a body to reference. A part
// matching neither is an ordinary attachment. The distinction decides which sentence a
// carried part gets: body content that could not be displayed is a visible loss, an
// ordinary file riding along is not.
function isBodyMedia(entry: { part: any; inBodyList: boolean }): boolean {
  return entry.inBodyList || (entry.part?.disposition === 'inline' && !!entry.part?.cid);
}


// Assemble the forward parameters from the caller's args and the already-fetched
// original email. Pure (no I/O) so the forward_email orchestration is unit-testable:
// coerce inputs, require an explicit recipient, build the Fwd: subject, record the
// original's Message-ID, assemble the forwarded-message bodies, and map the carried
// attachments (all-or-none; per-part subsetting goes through the draft +
// edit_draft's removeAttachments). Caller uploads are appended later by
// composeForward (the I/O boundary). Throws McpError on invalid input.
//
// `inline` carries the caller's already-coerced attachments for the embedded-image checks
// (see AuthoredInlineContext for why they are passed in rather than read from args).
//
// `mint` is injected only so tests are deterministic; production leaves it unset and the
// forwarded-block builder draws from the CSPRNG.
//
// `signature` is the sending identity's sign-off, already resolved by composeForward (the
// lookup is I/O, and this function is pure). Absent unless the caller asked for one.
export function buildForwardParams(
  args: any,
  originalEmail: any,
  inline: AuthoredInlineContext = DEFAULT_INLINE_CONTEXT,
  mint?: () => string,
  signature?: ResolvedSignature,
): BuiltForward {
  const a = args ?? {};
  // Validate the caller's note FIRST — before the forwarded-message block is assembled
  // below, which would otherwise mask a malformed note the same way a reply quote does
  // (see the equivalent guard in buildReplyParams).
  assertBodyInputs(a);
  const { from, textBody, htmlBody } = a;
  const { to, cc, bcc, replyTo } = coerceRecipients(a);
  const includeOriginalAttachments = coerceBool(a.includeOriginalAttachments) ?? true;
  const asAttachment = coerceBool(a.asAttachment) ?? false;

  if (!to || to.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, 'to is required for a forward; there is no default recipient');
  }

  // Subject: caller override used verbatim when non-blank; a supplied-but-blank value
  // falls to the default and a non-string is rejected (shared with reply_email's identical
  // parameter, so the two can't drift — #68). Default is "Fwd: <original>" without
  // double-prefixing an existing Fwd:/Fw:/FW:.
  const rawSubject = coerceSubjectOverride(a.subject, 'Omit it to use "Fwd: <original subject>".');
  let subject: string;
  if (rawSubject !== undefined) {
    subject = rawSubject;
  } else {
    const orig = originalEmail?.subject || '';
    subject = /^fwd?:/i.test(orig.trim()) ? orig : `Fwd: ${orig}`;
  }

  // Checked against the caller's OWN note, BEFORE the forwarded-message block is assembled
  // below — the block is this server's own markup and legitimately carries identifiers the
  // caller never wrote. An image the caller supplies can only be displayed if their note
  // ships as html; with no html note there is nothing to reference it, whatever format the
  // forwarded block itself ends up in.
  const inlinePlan = planAuthoredInlineImages({
    callerHtml: htmlBody,
    htmlShips: !isBlank(htmlBody),
    specs: inline.specs,
    attachmentsEnabled: inline.attachmentsEnabled,
    surface: 'note',
  });

  const params: ForwardParams = { to, cc, bcc, from, subject, replyTo };

  // Record the source on the draft — on BOTH forward shapes. The header is what
  // send_draft resolves to mark the original forwarded+read on transmit (#60), so an
  // asAttachment forward needs it as much as an inline one; the attached .eml carries
  // the content but is not machine-resolvable as provenance. The edit guard does not
  // arm on the bare header when the draft carries a message/rfc822 attachment (the
  // forwarded content lives there, untouched by body edits), so recording it here
  // leaves note edits on an asAttachment draft unchallenged.
  const originalMessageId = originalEmail?.messageId?.[0];
  if (isSettableMessageId(originalMessageId)) {
    params.forwardedMessageId = [originalMessageId];
  }
  // The caller named this exact instance; record it too (createDraft vets the value).
  // Recorded even when the original has no settable Message-ID (the JMAP id is real
  // either way, and the draft self-documents its source) — though send_draft keys its
  // marking off the provenance HEADERS (they classify reply-vs-forward), so a draft
  // with no recorded Message-ID still sends unmarked; the id only refines WHICH copy.
  if (typeof originalEmail?.id === 'string' && originalEmail.id !== '') {
    params.sourceEmailId = originalEmail.id;
  }

  // The part set is the GATED union of the original's `attachments` and the media parts
  // the server routed into its body lists — `attachments` alone is not a complete listing
  // (RFC 8621 §4.1.4), and a message whose only images are embedded can list nothing there
  // at all. A part without a blobId can't be re-referenced on a new Email (JMAP carry
  // works by blobId), so it is skipped without a count: there is no alternative to point
  // the caller at, and RFC 8621 attachments are blob-backed, so the case is theoretical.
  const sourceParts = buildUnionParts(originalEmail).filter((u) => u.part?.blobId);
  const emptyCarry = (): ForwardCarry => ({ pooled: [], attached: [], notIncluded: [] });

  if (asAttachment) {
    // Lossless form: the Email's own blobId is the raw RFC 5322 message; attach it
    // whole (verified live 2026-07-05: byte-identical round-trip). The body is just
    // the caller's note — no forwarded-message block — and when no note is given, a
    // short filler that matches NO forward marker (the guard's marker path stays
    // inert; the header it does carry is neutralized by the .eml carve-out) and the
    // draft ships with a readable body.
    if (isBlank(textBody) && isBlank(htmlBody)) {
      params.textBody = 'Forwarded message attached.';
    } else {
      if (textBody !== undefined) params.textBody = textBody;
      if (htmlBody !== undefined) params.htmlBody = htmlBody;
    }
    // Signed here rather than in a builder, because this shape has no builder: the note IS
    // the body. The filler above is signed too — a signature under one line of filler is
    // still the sign-off the caller asked for, and leaving it off would make the flag mean
    // something different on this one forward shape.
    const signedNote = applySignature(
      { textBody: params.textBody, htmlBody: params.htmlBody },
      signature,
    );
    if (signedNote.textBody !== undefined) params.textBody = signedNote.textBody;
    if (signedNote.htmlBody !== undefined) params.htmlBody = signedNote.htmlBody;
    if (!originalEmail?.blobId) {
      throw new McpError(ErrorCode.InternalError, 'Original email has no blobId; cannot attach it as .eml');
    }
    params.attachments = [{
      blobId: originalEmail.blobId,
      type: 'message/rfc822',
      name: sanitizeEmlFilename(originalEmail?.subject),
      disposition: 'attachment',
    }];
    return {
      asAttachment,
      forwardParams: params,
      inlinePlan,
      quoteImages: emptyQuoteImages(),
      carry: emptyCarry(),
    };
  }

  // Inline forward: reproduce the original under the forwarded-message block, resolving the
  // images its body displayed against the original's own parts.
  const bodies = buildForwardBodies({
    original: originalEmail,
    textBody,
    htmlBody,
    quoteImages: { sourceParts: sourceParts.map((u) => u.part), ...(mint && { mint }) },
    // Appended inside the builder, between the caller's note and the forwarded-message
    // block — see the signature section of src/reply-quote.ts.
    ...(signature && { signature }),
  });
  params.textBody = bodies.textBody;
  params.htmlBody = bodies.htmlBody;
  const quoteImages = bodies.quoteImages ?? emptyQuoteImages();

  // An image the forwarded block displays is BODY CONTENT, so it is carried whatever
  // includeOriginalAttachments says: leaving it behind would forward a body with a hole in
  // it, and the flag is about the original's FILES. Everything else goes to the pool — the
  // regular-attachment carry set the flag really governs — including an image the block
  // referenced but could not display, which is a part with nothing to reference it.
  const embedded = new Set(quoteImages.mappings.map((m) => m.source));
  // A part the body REFERENCED is body content by definition, whatever its metadata says.
  // The reference is the stronger signal and the one to pool on: a part routed only into
  // `attachments`, with its disposition omitted, is the usual shape for an embedded image,
  // so judging it by disposition alone would file a part the reader sees as a missing
  // picture as an ordinary file — and the pooled sentence is the only one that would have
  // mentioned it.
  const referenced = new Set(quoteImages.resolvedParts);
  const carry = emptyCarry();
  const carried: AttachmentPart[] = [];
  for (const entry of sourceParts) {
    if (embedded.has(entry.part)) continue;
    if (!includeOriginalAttachments) {
      carry.notIncluded.push(entry.part);
      continue;
    }
    carried.push(carryPart(entry.part));
    if (referenced.has(entry.part) || isBodyMedia(entry)) carry.pooled.push(entry.part);
    else carry.attached.push(entry.part);
  }

  // The minted parts ride last and unconditionally. Kept out of `carried` on purpose: a
  // part this call created to make the block display is not one of the original's files,
  // and merging the two would make the flag's promise unreadable.
  const attachments = [...carried, ...quoteImages.minted];
  if (attachments.length > 0) params.attachments = attachments;

  return { asAttachment, forwardParams: params, inlinePlan, quoteImages, carry };
}

// The minimal client surface composeForward needs; JmapClient satisfies it
// structurally. Declared here (like ReplyClient) so the orchestration stays
// unit-testable with a mock and free of a hard dependency on the concrete client.
export interface ForwardClient {
  getEmailById(id: string): Promise<any>;
  /** The sending identities, for the signature lookup (#33). Called only when one is asked for. */
  getIdentities(): Promise<any[]>;
  uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options?: UploadAttachmentsOptions,
  ): Promise<AttachmentPart[]>;
  createDraft(params: ForwardParams): Promise<string>;
}

export interface ComposeForwardResult {
  subject: string;
  emailId: string;
  /** What the draft ended up embedding, or could not (#13). Absent when there is nothing to say. */
  notes?: string[];
}

// Orchestrate a forward end to end: fetch the original, assemble the (pure) forward
// params, upload any NEW attachments and append them behind the carried/.eml parts,
// then save the forward as a draft. This tool only ever drafts — send_draft is the
// single tool that transmits mail, and it also does the thread-state maintenance
// (marking the original forwarded + read), resolved from the draft's recorded
// X-Forwarded-Message-Id header (#60; see docs/conventions.md "Draft provenance"),
// recorded on both inline and asAttachment forwards.
// Mirrors composeReply so the index handler stays a thin wrapper. attachDir and
// allowBlobAttach are passed in (resolved by the caller) so this reads no environment
// itself; between them they say which attachment SOURCES this server will accept.
export async function composeForward(
  args: any,
  client: ForwardClient,
  attachDir: string | undefined,
  allowBlobAttach: boolean,
): Promise<ComposeForwardResult> {
  const originalEmailId = args?.originalEmailId;
  if (!originalEmailId) {
    throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
  }

  const originalEmail = await client.getEmailById(originalEmailId);

  // The caller's note is validated before their attachments are even coerced, so a malformed
  // note is reported as a note problem rather than as whatever the attachments list happens
  // to be wrong about too. The builder re-runs this check as its own first step; running it
  // here as well is an ordering belt, not a second authority.
  assertBodyInputs(args ?? {});
  const specs = coerceAttachments(args?.attachments);

  // Off by default, and the lookup is skipped when it is off — see composeReply for why the
  // sign-off is resolved fresh on every call rather than remembered.
  const appendSignature = coerceBool(args?.appendSignature) === true;
  const signature = appendSignature
    ? resolveSignature(await client.getIdentities(), args?.from)
    : undefined;

  const { forwardParams, inlinePlan, quoteImages, carry } = buildForwardParams(
    args,
    originalEmail,
    { specs, attachmentsEnabled: !!attachDir || allowBlobAttach },
    undefined,
    signature,
  );

  // Upload NEW attachments (if any) after the pure builder and append them behind
  // the carried/.eml parts.
  let uploaded: AttachmentPart[] | undefined;
  if (specs?.length) {
    uploaded = await client.uploadAttachments(specs, attachDir, allowBlobAttach, { inlineCids: inlinePlan.inlineCids });
    forwardParams.attachments = [...(forwardParams.attachments ?? []), ...uploaded];
  }

  // Checked BEFORE the draft is saved, for the reason composeReply gives: an assembled
  // message that references an image it does not carry is this server's own error.
  checkInlineClosure({
    htmlBodies: [forwardParams.htmlBody],
    finalPartCids: (forwardParams.attachments ?? []).map((part) => part.cid),
    attachedMintedCids: quoteImages.minted.map((part) => part.cid),
  });

  const ledger = new InlineNoteLedger();
  recordQuoteImages(ledger, quoteImages, 'forward');
  // Keyed by identity within this call: the union has already deduped the parts, and a
  // forward can legitimately carry two files that differ only in their bytes.
  carry.pooled.forEach((part, i) => {
    ledger.record({
      key: `pool:${i}`, outcome: 'pooled', name: part.name, isImage: isImageType(part.type),
    });
  });
  carry.attached.forEach((part, i) => {
    ledger.record({ key: `carry:${i}`, outcome: 'attached', name: part.name });
  });
  carry.notIncluded.forEach((part, i) => {
    ledger.record({
      key: `excluded:${i}`, outcome: 'notIncluded', name: part.name, isImage: isImageType(part.type),
    });
  });

  const emailId = await client.createDraft(forwardParams);
  const notes = [
    ...ledger.emit({ surface: 'forward' }),
    ...await reportAuthoredInlineImages({
      // Caller uploads only (the forwarded block's own parts are on the ledger above), but
      // their identifiers go in so the confirmation read covers them too.
      uploaded,
      mintedCids: quoteImages.minted.map((part) => part.cid),
      plan: inlinePlan,
      emailId,
      readBack: (id) => client.getEmailById(id),
    }),
  ];
  return {
    subject: forwardParams.subject,
    emailId,
    ...(notes.length > 0 && { notes }),
  };
}
