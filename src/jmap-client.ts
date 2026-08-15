import { FastmailAuth } from './auth.js';
import { validateFastmailUrl } from './url-validation.js';
import { parseAddress, requireNonEmpty, validateClearFields, coerceUtcDate, PathAccessError, InvalidInputError } from './coerce.js';
import { normalizeBodies, htmlHasVisibleContent, buildBodyParts, isBlank, assertBodyInputs } from './body-format.js';
import { buildReplyBodies, hasQuoteMarker, hasTextQuoteMarker, buildForwardBodies, hasForwardMarker, hasTextForwardMarker, readQuotableHtml } from './reply-quote.js';
import { isSettableMessageId } from './forward-handler.js';
import {
  buildUnionParts, cidKey, describePart, sanitizeDownloadFilename,
  buildCidMap, checkInlineClosure, isRecreatableCid, isReservedCid, reconcileInlineParts,
  sanitizeQuoteHtml,
} from './inline-images.js';
import type { CidMapping, CidPart, MintedInlinePart, UnionPart } from './inline-images.js';
import {
  InlineNoteLedger, emitInlineNotes,
  noteEmbedMissingAfterSave, noteEmbedUnconfirmed, noteSentWithEmbedded,
  rejectBrokenDraft, rejectCidCollisionInCall, rejectCidCollisionOnDraft,
  rejectClearAttachmentsDanglingRefs, rejectDanglingCidRef, rejectInterleavedTextParts,
  rejectRemovalDanglingRef, rejectRemovalOfQuoteCarriedPart, rejectReservedCidRef,
  rejectUncarriableBodyPart, rejectUnrecreatableCid,
} from './inline-notes.js';
import type { AttachmentAvailability } from './inline-notes.js';
// unlink is a security control, not a convenience: the exclusive-create download
// path removes the file it just refused to trust before rewriting it.
import { writeFile, mkdir, realpath, stat, lstat, open, unlink } from 'fs/promises';
import type { FileHandle } from 'fs/promises';
import { dirname, resolve, normalize, sep, basename, join } from 'path';
import { homedir } from 'os';

// A JMAP Email "attachment" body part referencing an uploaded blob. blobId is the
// stable handle the server assigns; type is the stored MIME type. New uploads carry
// only blobId/type/name/disposition; a carried (re-referenced) part may also pass
// through cid/disposition from the existing draft.
export interface AttachmentPart {
  blobId: string;
  type: string;
  name?: string;
  disposition?: string;
  cid?: string;
}

/** How a compose path wants freshly uploaded files dispositioned. */
export interface UploadAttachmentsOptions {
  /**
   * The Content-IDs the message being composed actually displays.
   *
   * A file whose Content-ID is in this set is marked `inline` — the message body shows it.
   * Anything else, including every file when this is absent or empty, is a regular
   * attachment. The caller owns this decision because only the caller knows what html will
   * ship; see uploadAttachments for why guessing is not an option.
   */
  inlineCids?: ReadonlySet<string>;
}

// True if `child` is `parent` itself or nested beneath it. Case-fold is a parameter:
// the read guard compares case-insensitively on Win32 (NTFS folds case, so a
// case-sensitive startsWith is bypassable), while the write guard keeps its existing
// byte-exact compare so download behaviour is unchanged.
function isPathContained(child: string, parent: string, caseInsensitive: boolean): boolean {
  let c = child, p = parent;
  if (caseInsensitive) { c = c.toLowerCase(); p = p.toLowerCase(); }
  return c === p || c.startsWith(p + sep);
}

// Shared lexical pre-check: resolve `inputPath` against `allowedDir` (relative incl.
// a bare filename lands inside the root in one step; absolute must already be within),
// reject null bytes, and verify lexical containment. Returns the resolved absolute
// path. Throws PathAccessError on null bytes or escape. The canonical (realpath)
// re-verification is the caller's job — this is only the cheap first gate.
function lexicalContainedPath(inputPath: string, allowedDir: string, caseInsensitive: boolean): string {
  const resolved = resolve(allowedDir, normalize(inputPath));
  if (resolved.includes('\0')) {
    throw new PathAccessError('path contains null bytes');
  }
  if (!isPathContained(resolved, allowedDir, caseInsensitive)) {
    throw new PathAccessError(`path must be within ${allowedDir}. Received: ${inputPath}`);
  }
  return resolved;
}

// Reject Windows path forms that can dodge the lexical containment compare. Applied
// to the raw input on every platform: these shapes are never a legitimate attachment
// path, and resolving them first would mask the escape. (Device namespaces and UNC
// roots jump outside the drive-relative root; a drive-relative `C:foo` resolves
// against the drive's own CWD; a `:` past the drive names an NTFS alternate data
// stream; a `~` segment can be an 8.3 short name aliasing a long name past the compare.)
function rejectWindowsPathEscapes(input: string): void {
  if (/^\\\\[?.]\\/.test(input)) {
    throw new PathAccessError('path uses a Windows device namespace (\\\\?\\ or \\\\.\\), which is not allowed.');
  }
  if (/^(\\\\|\/\/)/.test(input)) {
    throw new PathAccessError('path is a UNC network path, which is not allowed.');
  }
  if (/^[A-Za-z]:(?![\\/])/.test(input)) {
    throw new PathAccessError('path is drive-relative (e.g. C:foo); use an absolute path or a name under the attach directory.');
  }
  // Strip a leading drive letter's own colon before scanning for a stream colon.
  if (input.replace(/^[A-Za-z]:/, '').includes(':')) {
    throw new PathAccessError("path contains a ':' (NTFS alternate data stream), which is not allowed.");
  }
  // 8.3 short-name segment (e.g. PROGRA~1) can alias a long name past the compare. Match
  // the 8.3 form specifically — a tilde followed by a digit — so legitimate filenames that
  // merely contain a tilde (e.g. report~final.pdf) are NOT rejected.
  if (/~\d/.test(input)) {
    throw new PathAccessError("path contains an 8.3 short-name segment (e.g. PROGRA~1), which is not allowed; use the full name.");
  }
}

// Extension -> MIME map for the Content-Type we POST when the caller omits one. Only
// sets the upload header; the recipient sees whatever type the server echoes back.
// A missing extension is not an error - it falls back to application/octet-stream - but
// that fallback is not equally harmless for every type: a receiving client decides from
// this header whether to render a calendar invitation or a forwarded message inline, so
// text/calendar and message/rfc822 change how the mail presents, not just its icon.
// message/rfc822 also covers forward_email's asAttachment output, which is a .eml.
const EXT_CONTENT_TYPES: Record<string, string> = {
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  zip: 'application/zip',
  txt: 'text/plain',
  md: 'text/markdown',
  csv: 'text/csv',
  json: 'application/json',
  html: 'text/html',
  ics: 'text/calendar',
  eml: 'message/rfc822',
};

function guessContentType(path: string): string {
  const ext = basename(path).split('.').pop()?.toLowerCase() ?? '';
  return EXT_CONTENT_TYPES[ext] ?? 'application/octet-stream';
}

// RFC 2045 token grammar for type/subtype, with a length cap. A positive grammar
// (reject anything outside the token set) closes header injection via the
// Content-Type we POST — defense in depth alongside undici's own header validation.
// MIME parameters (e.g. "; charset=utf-8") are intentionally not accepted: only the
// type/subtype is needed to upload a blob.
const MIME_TYPE_PATTERN = /^[A-Za-z0-9!#$%&'*+.^_`|~-]+\/[A-Za-z0-9!#$%&'*+.^_`|~-]+$/;

function validateContentType(value: string, index: number): string {
  const v = value.trim();
  if (v.length > 255 || !MIME_TYPE_PATTERN.test(v)) {
    throw new PathAccessError(`attachments[${index}] has an invalid contentType '${value}'. Use a MIME type like application/pdf.`);
  }
  return v;
}

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, any>;
  downloadUrl?: string;
  uploadUrl?: string;
  /**
   * Per-capability primary account ids, keyed by capability URN. `accountId` above is
   * the mail account; other capabilities (contacts in particular) can be a DIFFERENT
   * account on the same session, so callers that touch a non-mail capability must read
   * their own account id from here rather than reusing `accountId`.
   */
  primaryAccounts?: Record<string, string>;
}

/**
 * An attachment resolved to the fields needed to reference or fetch its blob. `type` and
 * `name` are defaulted at resolution time, so consumers never re-apply a fallback, and
 * `name` is additionally the SANITIZED form of the sender-supplied filename — safe to
 * hand onward as a save name. The raw declared name is not carried: nothing downstream
 * of resolution has a use for it that the sanitized form does not serve. `blobId` is
 * non-optional because resolution rejects a part that carries none.
 */
export interface AttachmentInfo {
  blobId: string;
  type: string;
  name: string;
  size?: number;
}

export interface JmapRequest {
  using: string[];
  methodCalls: [string, any, string][];
}

export interface JmapResponse {
  methodResponses: Array<[string, any, string]>;
  sessionState: string;
}

export interface QueryResult<T = any> {
  items: T[];
  total?: number;
  // The 0-based index the returned items start at within the full result set (RFC 8620
  // section 5.5). The server reports the position it actually served — a request past
  // the end is clamped — and we fall back to the position we asked for if a response
  // omits it. Read by the formatters to state where a page starts and to compute
  // `nextPosition`; like `exclusion`, it describes the QUERY, not a message, so it is
  // never serialized into the JSON body on either the simplified or the raw path.
  position?: number;
  // Out-of-band metadata for the default Trash/Spam exclusion. Populated by
  // searchEmails/getEmails when an exclusion was active; read by the handlers to
  // emit a trailing note. NEVER serialized into the JSON body or the raw path
  // (same discipline as getThread's hiddenDraftCount). `hidden:null` = the hidden
  // count could not be computed (degraded — emit the fail-closed note).
  exclusion?: {
    hidden: number | null;
    excludedRoles: string[];   // roles actually excluded (note fires iff hidden>0)
    unresolvedRoles: string[]; // roles intended-but-NOT-excluded (fail-loud note)
  };
}

// Result of computeExclusion: the mailbox ids to exclude via inMailboxOtherThan,
// the display labels of roles actually excluded, and the labels of roles we meant
// to exclude but couldn't resolve (fail-loud, never silently included).
export interface ExclusionResult {
  excludeIds: string[];
  excludedRoles: string[];
  unresolvedRoles: string[];
}

// Shared Email/get property lists — keep in sync per CLAUDE.md rules.
// COMPACT: used by the list/search tools and getThread (metadata + preview, no bodies).
// `textBody` here fetches only the body-part *structure* (partId/type/size), NOT the
// body content — no `bodyValues`/fetchTextBodyValues — so the response stays "no bodies"
// while giving the formatter the text part size for the `bodyTextSize` hint (#59).
export const EMAIL_PROPERTIES_COMPACT = [
  'id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'receivedAt',
  'preview', 'keywords', 'threadId', 'messageId', 'references', 'inReplyTo',
  'hasAttachment', 'header:List-Unsubscribe:asURLs', 'blobId', 'size', 'mailboxIds',
  'textBody',
] as const;

// VERBOSE: superset with body properties — used by verbose mode and getEmailById.
// `textBody` is already in COMPACT (structure); VERBOSE adds the content (`bodyValues`)
// and the html/attachment parts. `sentAt` is a get-path superset addition (allowed by the
// property-consistency rule) for reply-quote attribution (when the original was written).
// The two provenance headers this server records on drafts are deliberately
// VERBOSE-tier, not COMPACT: list items already show forward-ness via the isForwarded
// keyword flag, and the header VALUES are needed only when operating on a specific
// draft — the edit guard's recovery (X-Forwarded-Message-Id) and inspecting which
// stored copy send_draft will mark (X-Fastmail-MCP-Source-Id, see SOURCE_ID_HEADER
// below) — both get_email contexts.
// Corollary, not a gap: getThread uses EMAIL_PROPERTIES_COMPACT by default, so an
// ordinary thread read (including raw:true) doesn't surface forwardedMessageId or
// sourceEmailId — consistent with this decision. getThread's includeBodies mode (#74)
// switches to this same VERBOSE superset rather than inventing a third property list,
// so a thread read with bodies is a get-path read and does carry them.
export const EMAIL_PROPERTIES_VERBOSE = [
  ...EMAIL_PROPERTIES_COMPACT,
  'htmlBody', 'attachments', 'bodyValues', 'sentAt',
  'header:X-Forwarded-Message-Id:asMessageIds',
  'header:X-Fastmail-MCP-Source-Id:asText',
] as const;

// The provenance headers a draft carries about the message it was composed from:
// In-Reply-To (set by reply_email and every other client's reply) and
// X-Forwarded-Message-Id (set by forward_email and Fastmail's own clients). Both are
// arrays of BARE Message-IDs — JMAP MessageIds carry no angle brackets.
export interface SourceReferences {
  inReplyTo: string[];
  forwardedMessageId: string[];
  // The JMAP id of the exact stored instance the draft was composed from (the
  // X-Fastmail-MCP-Source-Id header). A Message-ID names a MESSAGE, not a stored
  // instance — an account can hold several copies of one message (e.g. a Sent copy
  // plus a self-delivered copy filed elsewhere), and only the compose call knew which
  // one the caller meant. Absent on drafts made by other clients or before this
  // header existed; send_draft then falls back to the Message-ID lookup.
  sourceEmailId?: string;
}

// The JMAP header form used to SET and GET the recorded source instance. Private
// bookkeeping in the Exchange/Thunderbird/Fastmail-client tradition (Fastmail's own
// clients stamp e.g. X-PersonalityId on drafts). It is NOT stripped on send:
// EmailSubmission transmits the stored bytes verbatim (probed live 2026-08-14 — even
// Fastmail's mobile app ships its private headers), and the value is an opaque,
// account-scoped id that discloses strictly less than the In-Reply-To header next to
// it. See docs/security-model.md for the recorded decision.
export const SOURCE_ID_HEADER = 'header:X-Fastmail-MCP-Source-Id:asText';

// Pre-vet a value before recording it as the source instance. RFC 8620 ids are
// URL-safe (A-Za-z0-9, hyphen, underscore); anything else is not a JMAP id and is
// treated as absent rather than risking a header-set rejection failing the create.
function isSettableSourceId(id: unknown): id is string {
  return typeof id === 'string' && /^[A-Za-z0-9_-]{1,255}$/.test(id);
}

// Read the provenance headers off a raw Email, tolerating a server that returns none.
export function readSourceReferences(email: any): SourceReferences {
  const forwarded = email?.['header:X-Forwarded-Message-Id:asMessageIds'];
  const sourceId = email?.[SOURCE_ID_HEADER];
  return {
    inReplyTo: Array.isArray(email?.inReplyTo) ? email.inReplyTo : [],
    forwardedMessageId: Array.isArray(forwarded) ? forwarded : [],
    ...(typeof sourceId === 'string' && sourceId.trim() !== '' && { sourceEmailId: sourceId.trim() }),
  };
}

// What sendDraft returns: the submission, plus the provenance of the message just sent —
// read off the same pre-send Email/get that validates the draft, so the caller's
// thread-state maintenance costs no extra round trip.
export interface SendDraftOutcome {
  submissionId: string;
  sourceReferences: SourceReferences;
  // What the transmitted message actually carried as embedded images (#13). A receipt, read
  // off the same pre-send fetch: the caller sees what went out, not what was intended.
  notes?: string[];
}

// Recall cap for the Message-ID lookup. The full-text step can match messages that merely
// quote or reference the id; the oldest-first sort guarantees the owner sits on the first
// page (see findEmailIdsByMessageId), so the cap only bounds how many extra mentions are
// fetched and filtered out.
const MESSAGE_ID_LOOKUP_LIMIT = 50;

// `disposition`/`cid` are load-bearing for forward_email's inline-image handling:
// without them every attachments[] part reads as disposition-less and the
// true-inline (cid) drop predicate can never fire (#30).
//
// `type` is load-bearing the same way for the part listing (#13): buildUnionParts
// classifies a body-list part by its media type, and a part with no type is read as
// body text. Dropping `type` from a property set used on an attachment-fetching path
// would therefore empty the body-embedded half of the listing SILENTLY — every
// embedded image simply absent, with no error. Do not trim this list per-call.
export const EMAIL_BODY_PROPERTIES = ['partId', 'blobId', 'type', 'size', 'name', 'disposition', 'cid'] as const;

// What get_email_attachments needs to answer honestly in both of its modes.
// `attachments` is the full part listing (the union of the JMAP attachments array and
// the media parts routed into the body lists) and is the basis every download index
// counts from. `rawAttachments` is the untouched JMAP array, which the tool's `raw`
// mode returns — a SET escape (#13).
//
// `omittedFromRaw` is what that mode withholds, counted by MEMBERSHIP rather than by
// subtracting lengths. The two agree on every real message, but a length difference is
// not the withheld set: two attachments sharing a blobId, or a null entry in the array,
// make the arithmetic under-report and silently suppress the disclosure the count
// exists to produce.
export interface EmailAttachmentsResult {
  attachments: any[];
  rawAttachments: any[];
  omittedFromRaw: number;
}

/**
 * Resolve a download_attachment `attachmentId` against a message's part listing.
 *
 * Four input forms, in this fixed order: partId, blobId, `cid:<value>`, then a plain
 * entry number. The order is load-bearing rather than arbitrary — Fastmail's partIds
 * ARE digit strings, so resolving digits as an index first would make "1" ambiguous
 * between "the part whose partId is 1" and "the second entry". A part always wins.
 *
 * Malformed INPUT throws InvalidInputError, which download_attachment lets through to
 * the caller; a well-formed reference that simply matches nothing returns undefined
 * and becomes the tool's generic not-found, which deliberately reveals no mailbox
 * metadata. That input-vs-existence split is documented in docs/conventions.md.
 */
function resolveAttachmentRef(parts: any[], attachmentId: string): any {
  const byPartId = parts.find((p: any) => p?.partId === attachmentId);
  if (byPartId) return byPartId;

  // First match, deliberately NOT rejected when several parts share the blobId. Blobs
  // are content-addressed, so sharing one means the parts ARE the same bytes — the
  // common case being one image both attached and embedded. The download is identical
  // whichever is picked; only the declared filename differs. That is why the
  // reject-on-ambiguity rule below belongs to the cid form, where two parts sharing a
  // Content-ID are genuinely different content.
  const byBlobId = parts.find((p: any) => p?.blobId === attachmentId);
  if (byBlobId) return byBlobId;

  if (/^cid:/i.test(attachmentId)) {
    // First occurrence only: `cid:cid:x` names the Content-ID "cid:x". A Content-ID may
    // legitimately begin with "cid:", and stripping repeatedly would make it unreachable.
    const literal = attachmentId.slice('cid:'.length);
    if (!literal) {
      throw new InvalidInputError(
        'attachmentId "cid:" names no Content-ID. Pass cid:<value> using the cid from ' +
        'get_email, or the part\'s blobId or partId from get_email_attachments.'
      );
    }
    // LITERAL first, decoded only as a fallback. The handle round-trips from get_email's
    // verbatim `cid` echo, so decoding first would break pasting back a cid that really
    // does contain a percent escape. Ambiguity WITHIN a stage is rejected, never guessed.
    const ambiguous = () => new InvalidInputError(
      `attachmentId "cid:${describePart(literal)}" matches more than one part. ` +
      'Pass the part\'s blobId or partId from get_email_attachments instead.'
    );
    const literalMatches = parts.filter((p: any) => p?.cid === literal);
    if (literalMatches.length > 1) throw ambiguous();
    if (literalMatches.length === 1) return literalMatches[0];

    const decoded = cidKey(attachmentId);
    if (decoded !== literal) {
      const decodedMatches = parts.filter((p: any) => p?.cid === decoded);
      if (decodedMatches.length > 1) throw ambiguous();
      if (decodedMatches.length === 1) return decodedMatches[0];
    }
    return undefined;
  }

  if (/^\d+$/.test(attachmentId)) {
    return parts[Number(attachmentId)];
  }

  // Anything parseInt would have swallowed as a number ("3a", "-1", "1.5") used to
  // index the list silently, so a typo downloaded the wrong file. It is now an input
  // error rather than a wrong answer.
  if (!Number.isNaN(Number.parseInt(attachmentId, 10))) {
    throw new InvalidInputError(
      `attachmentId "${describePart(attachmentId)}" is not a usable attachment reference. ` +
      'Pass a partId or blobId from get_email_attachments, cid:<value> for an embedded ' +
      'image, or a plain entry number (0, 1, 2, ...) counting from the start of that listing.'
    );
  }

  return undefined;
}

// ---------------------------------------------------------------------------
// The draft-edit part model: body shape, removals, carry (#13)
// ---------------------------------------------------------------------------

// The identity a part is deduped by across the three lists — the same rule the part union
// uses, so a part counted once there is counted once here. RFC 8621 §4.1.4 puts one
// displayed part into BOTH body lists, and a single-format draft aliases its one text part
// into both, so counting raw array entries would double every part in the message.
function draftPartKey(part: any, fallback: number): string {
  if (typeof part?.partId === 'string' && part.partId) return `p:${part.partId}`;
  if (typeof part?.blobId === 'string' && part.blobId) return `b:${part.blobId}`;
  return `i:${fallback}`;
}

// Content type without its parameters, lowercased. For CLASSIFYING only — the value stored
// or sent for a part is always the server's own string (RFC 2045 §5.1).
function classifyPartType(type: unknown): string {
  if (typeof type !== 'string') return '';
  const semicolon = type.indexOf(';');
  return (semicolon === -1 ? type : type.slice(0, semicolon)).trim().toLowerCase();
}

const TEXT_BODY_TYPES = new Set(['text/plain', 'text/html']);

// The non-text body parts the recreate can reproduce. The media three are what RFC 8621
// §4.1.4 routes into a body list in the first place; message/rfc822 is there because
// forward_email writes exactly that part itself (asAttachment), and a server that lists it
// as body content must not make its own drafts uneditable. Everything else — a calendar
// part, a signature part, an arbitrary application/* — has no flat-property spelling, so
// carrying it would change what the message IS.
function isCarriableBodyType(type: string): boolean {
  return /^(?:image|audio|video)\//.test(type) || type === 'message/rfc822';
}

export interface DraftBodyShape {
  // A part the recreate cannot reproduce, with whether it is media (which decides the
  // wording — "a media part" reads wrong for a calendar attachment).
  uncarriablePart?: { part: any; isMedia: boolean };
  // Set when two DISTINCT parts of one text type sit in the body lists: the Apple Mail
  // text-image-text layout, whose ordering a flat rebuild cannot express (issue #85).
  interleavedTextType?: string;
}

/**
 * Decide whether a draft's body is a shape the immutable-email recreate can rebuild.
 *
 * Deduping FIRST is load-bearing, not an optimization: a single-format draft lists its one
 * text part under both textBody and htmlBody, so a raw count would see two text/plain parts
 * on an ordinary plain-text draft and refuse to edit it. A typeless part is left alone —
 * the body reader treats one as body text, and inventing a refusal for it would reject
 * drafts that have always worked.
 */
export function classifyDraftBodyShape(email: any): DraftBodyShape {
  const lists = [email?.textBody, email?.htmlBody].filter(Array.isArray) as any[][];
  const seen = new Set<string>();
  const textTypeCounts = new Map<string, number>();
  const shape: DraftBodyShape = {};
  let index = 0;

  for (const list of lists) {
    for (const part of list) {
      if (!part) continue;
      const key = draftPartKey(part, index++);
      if (seen.has(key)) continue;
      seen.add(key);

      const type = classifyPartType(part.type);
      if (!type || TEXT_BODY_TYPES.has(type)) {
        if (!type) continue;
        const count = (textTypeCounts.get(type) ?? 0) + 1;
        textTypeCounts.set(type, count);
        if (count > 1 && !shape.interleavedTextType) shape.interleavedTextType = type;
        continue;
      }

      const carriable = isCarriableBodyType(type)
        && typeof part.blobId === 'string' && part.blobId !== '';
      if (!carriable && !shape.uncarriablePart) {
        shape.uncarriablePart = { part, isMedia: /^(?:image|audio|video)\//.test(type) };
      }
    }
  }

  return shape;
}

export interface AttachmentRemovalPlan {
  /** Parts this call takes off the draft, in stored order. */
  removed: any[];
  /** Parts that stay, in stored order. */
  survivors: any[];
  /**
   * The refusal a bad ref earned, held rather than thrown.
   *
   * Resolution runs EARLY (the rebuilt quote has to know which parts survive it) but the
   * body-shape guards must keep raising first, so this is raised at the point in the order
   * the removal loop always occupied. Resolution itself performs no I/O and throws nothing.
   */
  error?: PathAccessError;
}

/**
 * Match each removeAttachments ref against the draft's parts, or clear them all.
 *
 * A ref names a blobId (every part with that blob comes off — a blob shared by two parts is
 * the same bytes) or, failing that, a unique non-null name. A ref that matches nothing, or a
 * name that matches several, is an error rather than a silent no-op: a removal the caller
 * believes happened is the confidently-wrong result this codebase refuses to produce.
 */
export function resolveAttachmentRemovals(
  storedParts: any[],
  refs: string[] | undefined,
  clearAll: boolean,
): AttachmentRemovalPlan {
  if (clearAll) return { removed: storedParts.slice(), survivors: [] };

  let survivors = storedParts.slice();
  const removed: any[] = [];
  for (const ref of refs ?? []) {
    const byBlob = survivors.filter((p) => p.blobId !== ref);
    if (byBlob.length < survivors.length) {
      removed.push(...survivors.filter((p) => p.blobId === ref));
      survivors = byBlob;
      continue;
    }
    const nameMatches = survivors.filter((p) => p.name != null && p.name === ref);
    if (nameMatches.length === 1) {
      removed.push(nameMatches[0]);
      survivors = survivors.filter((p) => p !== nameMatches[0]);
      continue;
    }
    if (nameMatches.length > 1) {
      return {
        removed,
        survivors,
        error: new PathAccessError(
          `removeAttachments ref '${ref}' matches ${nameMatches.length} attachments by name; pass the blobId instead (one of: ${survivors.map((p) => p.blobId).join(', ')}).`,
        ),
      };
    }
    return {
      removed,
      survivors,
      error: new PathAccessError(
        `removeAttachments ref '${ref}' matched no attachment on this draft. Carried blobIds: ${storedParts.map((p) => p.blobId).join(', ') || '(none)'}.`,
      ),
    };
  }
  return { removed, survivors };
}

// Re-reference a stored part on the recreated draft. Whitelist exactly these fields — a
// blob-backed part is blobId XOR partId, and `size` is server-set, so sending partId/size
// would be rejected by a strict server. `demote` is what an embedded image becomes when the
// body that displayed it is gone but the bytes are not this server's to discard: an ordinary
// attachment. Its Content-ID rides along — a Content-ID never makes a part inline on its
// own (probed live 2026-08-14) — so nothing is lost if a later edit displays it again.
function carriedPartFrom(part: any, demote = false): AttachmentPart {
  return {
    blobId: part.blobId,
    type: part.type,
    ...(part.name != null && { name: part.name }),
    ...(demote ? { disposition: 'attachment' } : part.disposition != null && { disposition: part.disposition }),
    ...(part.cid != null && { cid: part.cid }),
  };
}

// The embedded-image references an html body makes, as comparison keys. The collecting pass
// reports them without deciding anything, so this reads the same references the quote
// rewriter would act on.
function htmlCidRefs(html: string | null | undefined): string[] {
  if (!html) return [];
  return sanitizeQuoteHtml(html, { mode: 'collect' }).refs;
}

function partCid(part: any): string {
  return typeof part?.cid === 'string' ? part.cid : '';
}

function partBytes(part: any): number {
  return typeof part?.size === 'number' && part.size > 0 ? part.size : 0;
}

function isImagePart(part: any): boolean {
  return classifyPartType(part?.type).startsWith('image/');
}

// A compact fingerprint of the draft an edit replaced, echoed back so a caller that
// edited from a stale copy sees immediately what it overwrote (#65). Body sizes are the
// character lengths of the stored values, not the bodies themselves: the old draft
// survives in Trash, so its full content is one get_email away and echoing it back
// would be bulk with no extra signal. Bcc is left out for the same reason.
export interface ReplacedDraftInfo {
  id: string;
  subject?: string;
  to?: string[];
  cc?: string[];
  textBodySize?: number;
  htmlBodySize?: number;
}

// updateDraft's result. Exactly ONE of trashedOldDraftId / orphanedOldDraftId is always
// set: the old copy either reached Trash or is still sitting where it was. It is never
// destroyed, and its fate is never left unstated.
export interface UpdateDraftResult {
  id: string;
  replacedDraft: ReplacedDraftInfo;
  trashedOldDraftId?: string;
  orphanedOldDraftId?: string;
  // Why the Trash move didn't happen. Set whenever orphanedOldDraftId is.
  orphanedOldDraftReason?: string;
  // What the edit did to the draft's embedded images (#13). Present only when it did
  // something: an edit that neither embeds, demotes nor removes one omits both fields.
  inlineImages?: { embedded: number; degraded: number; removed: number };
  // The sentences describing those outcomes, composed once from the final state. Rendered
  // by the handler — a count nobody prints is not a disclosure.
  notes?: string[];
}

export interface MailboxInfo {
  name: string;
  // The stable JMAP role (lowercased), or null for a custom folder/label. Only
  // the Mailbox object carries a role; the Email object never does (RFC 8621).
  role: string | null;
}

// Build an id -> {name, role} lookup from a Mailbox/get list. We require a real
// string `name` (so a malformed/missing name leaves the id unresolvable — surfaced
// later, not silently dropped) but key on the id, so custom labels — which have
// role:null — still resolve their name; default mailboxes carry names like
// "Trash"/"Archive" AND a role. Role is lowercased here so callers get the
// docs-promised lowercase form regardless of server casing. (#10, #49)
export function buildMailboxInfoMap(mailboxes: any[]): Map<string, MailboxInfo> {
  const map = new Map<string, MailboxInfo>();
  for (const mb of mailboxes || []) {
    if (mb && mb.id && typeof mb.name === 'string') {
      const role = typeof mb.role === 'string' && mb.role ? mb.role.toLowerCase() : null;
      map.set(mb.id, { name: mb.name, role });
    }
  }
  return map;
}

// Attach the resolved mailbox/label location onto each raw email as NON-enumerable
// properties (`_mailboxNames`, `_mailboxRoles`, `_unresolvedMailboxIds`) so
// JSON.stringify — the raw:true paths — omits all three while simplifyEmail can
// still read them. Raw output therefore stays pure JMAP (opaque mailboxIds), and
// the simplified output carries friendly names + stable roles.
//
// NEVER-SILENT, NON-THROWING resolution (#53). For each of the email's mailboxIds:
//   - resolves to a name  -> name goes to `_mailboxNames`; its role (if any) to `_mailboxRoles`.
//   - does NOT resolve    -> the raw id goes to `_unresolvedMailboxIds` (the fallback +
//                            the explicit "this couldn't be named" indicator).
// So a promised location is never silently dropped: it appears either as a friendly
// name or as a raw id flagged in `unresolvedMailboxIds`. This function does NOT throw
// on an unresolved id, and it never fabricates a name.
//
// WHY not throw, and WHY not silently omit (do not "fix" this back to either):
//   - An unresolved id is rare and benign — a just-created custom folder, a TOCTOU
//     mailbox-creation race on the separately-fetched list/search mailbox list, or a
//     mailbox with a malformed/missing `name`. Trash/Spam (and all role mailboxes)
//     ALWAYS resolve, so the location that motivated #49/#53 is never the unresolved one.
//   - Strict-throw was rejected as disproportionate: it would fail a whole
//     list_emails/search_emails page over one rare/benign id, against the codebase's
//     documented best-effort resolver posture.
//   - Silently omitting the id (the prior behaviour) WAS the #53 bug — a promised
//     field vanished with no trace.
// A genuine Mailbox/get `error` response is a different thing and still throws via the
// callers' existing catches (a real failure stays loud); that is not this path.
//
// `_mailboxRoles` and `_unresolvedMailboxIds` are attached only when non-empty (so the
// formatter's addIf omits them); `_mailboxNames` keeps its existing attach-when-non-empty
// behaviour. `roles` and `mailboxes` are INDEPENDENT sets, not parallel arrays — a custom
// folder contributes a name but no role, so their lengths can differ. (#10, #49, #53)
export function attachMailboxInfo(emails: any[], map: Map<string, MailboxInfo>): void {
  for (const email of emails || []) {
    if (!email || !email.mailboxIds) continue;
    const ids = Object.keys(email.mailboxIds);
    if (ids.length === 0) continue;
    const names: string[] = [];
    const roles: string[] = [];
    const unresolved: string[] = [];
    for (const id of ids) {
      const info = map.get(id);
      if (info) {
        names.push(info.name);
        if (info.role) roles.push(info.role);
      } else {
        unresolved.push(id);
      }
    }
    if (names.length > 0) {
      Object.defineProperty(email, '_mailboxNames', { value: names, enumerable: false, configurable: true });
    }
    if (roles.length > 0) {
      Object.defineProperty(email, '_mailboxRoles', { value: roles, enumerable: false, configurable: true });
    }
    if (unresolved.length > 0) {
      Object.defineProperty(email, '_unresolvedMailboxIds', { value: unresolved, enumerable: false, configurable: true });
    }
  }
}

// Cap the mailbox names listed in a not-found error so a large account doesn't
// produce a huge message; the list_mailboxes pointer keeps a truncated list actionable.
const MAILBOX_LIST_CAP = 30;

// The shared "Valid: …" tail listing known mailboxes (name + role), capped so a large
// account can't produce a huge message. Reused by both the single-input and the multi-input
// (label array) not-found messages so they truncate identically and read consistently.
function mailboxListHint(mailboxes: any[]): string {
  const entries = (mailboxes || [])
    .filter(mb => mb && typeof mb.name === 'string')
    .map(mb => (mb.role ? `${mb.name} (${mb.role})` : mb.name));
  const shown = entries.slice(0, MAILBOX_LIST_CAP);
  let list = shown.join(', ');
  if (entries.length > shown.length) {
    list += `, …and ${entries.length - shown.length} more — call list_mailboxes for the full list`;
  }
  return `Use a name, or a role (inbox/archive/sent/drafts/trash/junk). Valid: ${list}`;
}

function formatMailboxNotFound(input: string, mailboxes: any[]): string {
  return `Mailbox '${input}' not found. ${mailboxListHint(mailboxes)}`;
}

// Multi-input not-found variant for the label arrays: name EVERY unresolved value in one
// message (so an agent fixes all typos in a single retry), then the same shared hint so it
// reads as a natural superset of the single-input message. The reflected values are the
// caller's own input (redacted at the top-level catch), but bound the volume anyway: cap
// the count with a "…and N more" tail and clamp each value's length so a huge/long
// mailboxIds array can't be reflected wholesale.
const UNRESOLVED_VALUE_MAXLEN = 80;
function formatMailboxesNotFound(unresolved: string[], mailboxes: any[]): string {
  const clamp = (v: string) => (v.length > UNRESOLVED_VALUE_MAXLEN ? `${v.slice(0, UNRESOLVED_VALUE_MAXLEN)}…` : v);
  const shown = unresolved.slice(0, MAILBOX_LIST_CAP).map(v => `'${clamp(v)}'`);
  let listed = shown.join(', ');
  if (unresolved.length > shown.length) {
    listed += `, …and ${unresolved.length - shown.length} more`;
  }
  return `Mailbox(es) not found: ${listed}. ${mailboxListHint(mailboxes)}`;
}

// Exact-only mailbox match: exact id -> role (case-insensitive) -> exact name
// (case-insensitive); returns undefined on a miss. NO substring matching (substring is an
// injection-steering primitive on write paths and can mis-resolve). Edge: a custom mailbox
// literally named after a role (e.g. "Archive") is shadowed by the role branch — acceptable,
// and the reason the docs use "Receipts" not "Archive" as the name example. Input is
// trimmed/stringified HERE so every caller (the single-mailbox params and the label arrays)
// normalises identically. Exported pure for unit testing.
export function findMailboxExact(mailboxes: any[], input: string): any {
  const list = mailboxes || [];
  const raw = String(input).trim();
  const byId = list.find(mb => mb && mb.id === raw);
  if (byId) return byId;
  const lower = raw.toLowerCase();
  const byRole = list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === lower);
  if (byRole) return byRole;
  const byName = list.find(mb => mb && typeof mb.name === 'string' && mb.name.toLowerCase() === lower);
  if (byName) return byName;
  return undefined;
}

// Throwing wrapper over findMailboxExact for the single-mailbox callers: resolve one input
// or throw InvalidInputError with the (capped) valid list. Exported pure for unit testing.
export function resolveMailbox(mailboxes: any[], input: string): any {
  const found = findMailboxExact(mailboxes, input);
  if (found) return found;
  throw new InvalidInputError(formatMailboxNotFound(String(input).trim(), mailboxes || []));
}

// Compute the default Trash/Spam exclusion. Resolves trash/junk by EXACT role only
// (case-insensitive) — NEVER findMailboxByRoleOrName, whose substring name fallback
// could mis-hit a custom mailbox (e.g. "Junk mail rules") and silently hide real mail.
// When an explicit mailbox is set, exclusion is off (the explicit scope wins). When we
// intend to exclude a role we can't resolve (role absent, OR an empty/degraded mailbox
// list), DO NOT silently include it: flag it in unresolvedRoles so the handler emits a
// fail-loud "not excluded" note — never run a default search/list with zero exclusion
// ids and zero disclosure. Exported pure for unit testing.
export function computeExclusion(
  mailboxes: any[],
  opts: { includeTrash?: boolean; includeSpam?: boolean; hasExplicitMailbox?: boolean },
): ExclusionResult {
  const excludeIds: string[] = [];
  const excludedRoles: string[] = [];
  const unresolvedRoles: string[] = [];
  if (opts.hasExplicitMailbox) {
    return { excludeIds, excludedRoles, unresolvedRoles };
  }
  const list = mailboxes || [];
  const findRole = (role: string) => list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === role);
  if (!opts.includeTrash) {
    const tb = findRole('trash');
    if (tb) { excludeIds.push(tb.id); excludedRoles.push('Trash'); }
    else unresolvedRoles.push('Trash');
  }
  if (!opts.includeSpam) {
    const jb = findRole('junk');
    if (jb) { excludeIds.push(jb.id); excludedRoles.push('Spam'); }
    else unresolvedRoles.push('Spam');
  }
  return { excludeIds, excludedRoles, unresolvedRoles };
}

/** Match an email address against an identity, supporting wildcard identities (e.g. *@example.com). */
function matchesIdentity(identityEmail: string, address: string): boolean {
  const identity = identityEmail.toLowerCase();
  const addr = address.toLowerCase();
  if (identity === addr) return true;
  if (identity.startsWith('*@')) {
    const domain = identity.slice(1); // "@example.com"
    // A wildcard identity is only honoured for a single well-formed addr-spec. Without
    // this, a composite value like "a@evil.com,b@example.com" (or one carrying CR/LF or
    // a quoted local part) satisfies the endsWith test and lands unparsed in the
    // outgoing `from`/`mailFrom`, turning the "verified identity" check into a pass.
    // Note the pattern admits a BARE addr-spec only — a "Name <a@b.com>" form is
    // rejected on purpose, because the display name is supplied separately and is never
    // part of the value matched here. Do not widen it to accept angle-addr shapes.
    if (!/^[^\s@,;"]+@[^\s@,;"]+$/.test(addr)) return false;
    return addr.endsWith(domain);
  }
  return false;
}

export class JmapClient {
  private auth: FastmailAuth;
  private session: JmapSession | null = null;

  constructor(auth: FastmailAuth) {
    this.auth = auth;
  }

  /**
   * Extract the result from a JMAP method response, throwing on method-level errors.
   */
  protected getMethodResult(response: JmapResponse, index: number): any {
    if (!response.methodResponses || index >= response.methodResponses.length) {
      throw new Error(
        `JMAP response missing expected method at index ${index} (got ${response.methodResponses?.length ?? 0} responses)`
      );
    }
    const entry = response.methodResponses[index];
    if (!Array.isArray(entry) || entry.length < 2) {
      throw new Error(`JMAP response entry at index ${index} is malformed`);
    }
    const [tag, result] = entry;
    if (tag === 'error') {
      // This is a top-level method-`error` entry (RFC 8620 §3.6.1) — a DIFFERENT
      // shape from the per-id SetError that describeSetError() formats. Same {type,
      // description} fields, distinct error class, so this copy is intentionally
      // separate rather than routed through that helper.
      throw new Error(`JMAP error: ${result.type}${result.description ? ' - ' + result.description : ''}`);
    }
    return result;
  }

  /**
   * Format a JMAP SetError (RFC 8620 §5.3) for a human-readable message. The single
   * chokepoint every throwing notCreated/notUpdated set-error site routes through, so
   * the "type - description" format lives in one place and can't drift.
   *
   * SetError.type is required by the spec (so no missing-type guard is needed — a
   * guard there would silently reintroduce a reasonless message); description is
   * optional. We concatenate ONLY the server-authored type/description — we never
   * attach our own copy of the message body. (A server may put a snippet in its own
   * description; that is the server's text, identical to what the notCreated path has
   * always shipped — so the promise is "we add no content," not "no content can ever
   * appear.") Caller-supplied failing ids are added by the bulk callers, not here.
   */
  protected describeSetError(entry: { type: string; description?: string }): string {
    return `${entry.type}${entry.description ? ' - ' + entry.description : ''}`;
  }

  /**
   * Throw the correctly-classified error for a single-id Email/set failure, surfacing
   * the server's reason (#22) and discriminating the MCP code by SetError type (#41):
   * a `notFound` is a bad id the caller can fix → InvalidInputError (InvalidParams);
   * any other type (`forbidden`, `serverFail`, …) is an operational failure the caller
   * can't fix by re-forming args → plain Error (InternalError). `action` is the verb
   * phrase, e.g. "move email", yielding "Failed to move email: notFound - …".
   */
  protected throwSingleSetError(entry: { type: string; description?: string }, action: string): never {
    const message = `Failed to ${action}: ${this.describeSetError(entry)}`;
    if (entry.type === 'notFound') {
      throw new InvalidInputError(message);
    }
    throw new Error(message);
  }

  /**
   * Throw the correctly-classified error for a bulk Email/set partial failure. Reports
   * success/fail counts plus the caller's own failing ids grouped by reason (#22), so an
   * agent can retry exactly the failures. Iterates ONLY the caller's input ids
   * (Object.keys(notUpdated)) — never echoes a server-originated id or message content.
   * Caps the ids-per-reason and reason count; on truncation it says the list is PARTIAL
   * (so a truncated list never reads as complete) and points at re-running the full input
   * set, which is safe because these mutators are idempotent. Classified per #41: if every
   * failure is `notFound` the whole batch is caller-fixable bad ids → InvalidInputError
   * (InvalidParams); any other type means at least one operational failure → plain Error
   * (InternalError).
   */
  protected throwBulkSetError(
    notUpdated: Record<string, { type: string; description?: string }>,
    total: number,
    action: string,
  ): never {
    const MAX_REASONS = 5;
    const MAX_IDS_PER_REASON = 10;

    const failedIds = Object.keys(notUpdated);
    const failCount = failedIds.length;
    const successCount = total - failCount;

    // Group failing ids by their server-stated reason.
    const byReason = new Map<string, string[]>();
    for (const id of failedIds) {
      const reason = this.describeSetError(notUpdated[id]);
      const ids = byReason.get(reason);
      if (ids) ids.push(id);
      else byReason.set(reason, [id]);
    }

    let truncated = false;
    const reasonEntries = [...byReason.entries()];
    if (reasonEntries.length > MAX_REASONS) truncated = true;
    const groups = reasonEntries.slice(0, MAX_REASONS).map(([reason, ids]) => {
      if (ids.length > MAX_IDS_PER_REASON) truncated = true;
      return `${reason}: ${ids.slice(0, MAX_IDS_PER_REASON).join(', ')}`;
    });

    let message = `Failed to ${action} ${failCount} of ${total} emails (${successCount} succeeded). ${groups.join('; ')}`;
    if (truncated) {
      message += '. (Partial list — not every failure is shown. These operations are idempotent, so re-run with the full input set to retry every failure safely.)';
    }

    if (failedIds.every(id => notUpdated[id].type === 'notFound')) {
      throw new InvalidInputError(message);
    }
    throw new Error(message);
  }

  /**
   * Extract the .list array from a JMAP method response, with null safety.
   */
  protected getListResult(response: JmapResponse, index: number): any[] {
    const result = this.getMethodResult(response, index);
    return result?.list || [];
  }

  /**
   * Like getListResult, but returns [] when the method at `index` is absent
   * instead of throwing. Used for an appended Mailbox/get (the trailing #10
   * mailbox-name resolver): a server that drops the trailing method, or an
   * older 2-entry test stub, degrades to "no names resolved" rather than an error.
   */
  protected readListResultIfPresent(response: JmapResponse, index: number): any[] {
    if (!response.methodResponses || index >= response.methodResponses.length) return [];
    return this.getListResult(response, index);
  }

  /**
   * Build a QueryResult from a query + get pair.
   * queryIndex is the /query response; listIndex is the /get response.
   */
  protected getQueryResult(response: JmapResponse, queryIndex: number, listIndex: number): QueryResult {
    const queryResult = this.getMethodResult(response, queryIndex);
    const items = this.getListResult(response, listIndex);
    const total = queryResult?.total;
    const result: QueryResult = total != null ? { items, total } : { items };
    // The position the server actually served these items from. Kept only when it is a
    // number: paging arithmetic on a garbled value would be worse than falling back to
    // the position we requested (which the paging callers supply).
    if (typeof queryResult?.position === 'number') result.position = queryResult.position;
    return result;
  }

  async getSession(): Promise<JmapSession> {
    if (this.session) {
      return this.session;
    }

    const response = await fetch(this.auth.getSessionUrl(), {
      method: 'GET',
      headers: this.auth.getAuthHeaders(),
      // Never follow a redirect on a token-bearing request: the session URL is built from
      // the allowlist-validated base URL, so a 3xx could only point off-allowlist — the
      // bearer token would be replayed to an unvalidated host.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Failed to get session: ${response.statusText}`);
    }

    const sessionData = await response.json() as any;

    // Validate every URL the server hands us before we send the bearer token to it.
    // The downloadUrl/uploadUrl are URL templates with {accountId}/{blobId}/etc.
    // placeholders, so we strip those for parsing and validate origin only.
    const allowUnsafe = this.auth.getAllowUnsafe();
    const stripTemplate = (url: string) => url.replace(/\{[^}]+\}/g, 'x');
    if (typeof sessionData.apiUrl !== 'string') {
      throw new Error('Invalid session response: apiUrl missing');
    }
    validateFastmailUrl(sessionData.apiUrl, 'session.apiUrl', allowUnsafe);
    // Reject a present-but-non-string download/upload URL rather than storing it
    // unvalidated: a `typeof === 'string'` guard around validation alone lets a
    // non-string value skip the check and still be stored, so validate and store
    // must not diverge.
    if (sessionData.downloadUrl !== undefined) {
      if (typeof sessionData.downloadUrl !== 'string') {
        throw new Error('Invalid session response: downloadUrl is not a string');
      }
      validateFastmailUrl(stripTemplate(sessionData.downloadUrl), 'session.downloadUrl', allowUnsafe);
    }
    if (sessionData.uploadUrl !== undefined) {
      if (typeof sessionData.uploadUrl !== 'string') {
        throw new Error('Invalid session response: uploadUrl is not a string');
      }
      validateFastmailUrl(stripTemplate(sessionData.uploadUrl), 'session.uploadUrl', allowUnsafe);
    }

    this.session = {
      apiUrl: sessionData.apiUrl,
      accountId: sessionData.primaryAccounts?.['urn:ietf:params:jmap:mail']
        || sessionData.primaryAccounts?.['urn:ietf:params:jmap:core']
        || Object.keys(sessionData.accounts)[0],
      capabilities: sessionData.capabilities,
      downloadUrl: sessionData.downloadUrl,
      uploadUrl: sessionData.uploadUrl,
      // Carried whole so a non-mail capability (contacts) can resolve its own account
      // id; `accountId` above only answers for mail.
      primaryAccounts: sessionData.primaryAccounts
    };

    return this.session;
  }

  async makeRequest(request: JmapRequest): Promise<JmapResponse> {
    const session = await this.getSession();
    
    const response = await fetch(session.apiUrl, {
      method: 'POST',
      headers: this.auth.getAuthHeaders(),
      body: JSON.stringify(request),
      // Never follow a redirect on a token-bearing request: session.apiUrl was
      // allowlist-validated in getSession, so a 3xx could only point off-allowlist — the
      // bearer token would be replayed to an unvalidated host.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`JMAP request failed: ${response.statusText}`);
    }

    const data = await response.json();
    if (!data || !Array.isArray(data.methodResponses)) {
      throw new Error('Invalid JMAP response: missing or malformed methodResponses');
    }
    return data as JmapResponse;
  }

  // Resolve a fixed role with a SUBSTRING name fallback. The substring fallback is an
  // injection-steering / mis-resolution hazard on any exclusion/delete/move target, so
  // this is kept ONLY for the compose path (drafts/sent save target), where it resolves
  // a benign save destination. Default-exclusion uses computeExclusion (exact role),
  // delete/move/the #12 sweep use resolveMailbox / resolveMailboxId (exact only).
  protected findMailboxByRoleOrName(mailboxes: any[], role: string, nameFallback?: string): any | undefined {
    return mailboxes.find(mb => mb.role === role) ||
           (nameFallback ? mailboxes.find(mb => mb.name.toLowerCase().includes(nameFallback)) : undefined);
  }

  // Resolve trash/junk for the default Trash/Spam exclusion by EXACT role only
  // (case-insensitive) — used by both searchEmails and getEmails. Fixed-role lookup
  // with a private helper to share the resolved id between the visible filter and the
  // hidden-count query.
  private findByExactRole(mailboxes: any[], role: string): any | undefined {
    const target = role.toLowerCase();
    return (mailboxes || []).find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === target);
  }

  // Resolve an optional mailbox input to an id. undefined/blank -> undefined (no filter).
  // Else resolve EXACTLY against a passed-in list (shared, no double-fetch) or one
  // getMailboxes(); throws InvalidInputError on no match. Used by every swept tool
  // (reads + writes) — safe to share now that matching is exact.
  private async resolveMailboxId(input?: string, mailboxes?: any[]): Promise<string | undefined> {
    if (input === undefined || input === null || String(input).trim() === '') return undefined;
    const list = mailboxes ?? await this.getMailboxes();
    return resolveMailbox(list, input).id;
  }

  async getMailboxes(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/get', { accountId: session.accountId }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0);
  }

  // Options-object signature (a positional add of the three scope bools would be
  // fragile). When no explicit `mailbox` is set, applies the same default Trash/Spam
  // exclusion + hidden-count as searchEmails. Its only structural filter is the
  // excludeDrafts keyword — isUnread/isPinned are NOT exposed on list_emails.
  async getEmails(opts: {
    mailbox?: string;
    limit?: number;
    position?: number;
    ascending?: boolean;
    includeTrash?: boolean;
    includeSpam?: boolean;
    excludeDrafts?: boolean;
  } = {}): Promise<QueryResult> {
    const mailboxes = await this.getMailboxes();
    const resolvedMailboxId = await this.resolveMailboxId(opts.mailbox, mailboxes);

    const base: any = {};
    if (resolvedMailboxId) base.inMailbox = resolvedMailboxId;

    const conds: any[] = [];
    if (opts.excludeDrafts) conds.push({ notKeyword: '$draft' });

    const hasExplicitMailbox = !!resolvedMailboxId;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: opts.includeTrash,
      includeSpam: opts.includeSpam,
      hasExplicitMailbox,
    });
    const exclusionIntended = !hasExplicitMailbox && (!opts.includeTrash || !opts.includeSpam);

    return this.runFilteredQuery({
      base,
      conds,
      exclusion,
      exclusionIntended,
      limit: opts.limit ?? 20,
      ascending: opts.ascending ?? false,
      mailboxes,
      position: opts.position,
    });
  }

  async getEmailById(id: string): Promise<any> {
    const session = await this.getSession();

    // No maxBodyValueBytes: verified live (2026-06-24) that Fastmail does NOT truncate body
    // values by default (a 5 MB body returned whole, isTruncated=false), so reply_email gets
    // the complete original to quote. (An explicit maxBodyValueBytes:0 is REJECTED by Fastmail
    // with invalidArguments, so we must not send it.) The reply-quote module still appends an
    // elision marker if any bodyValue ever reports isTruncated, as a defensive net.
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [id],
          properties: [...EMAIL_PROPERTIES_VERBOSE],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'email'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name', 'role'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notFound && result.notFound.includes(id)) {
      throw new InvalidInputError(`Email with ID '${id}' not found`);
    }

    const email = result.list?.[0];
    if (!email) {
      throw new InvalidInputError(`Email with ID '${id}' not found or not accessible`);
    }

    attachMailboxInfo([email], buildMailboxInfoMap(this.readListResultIfPresent(response, 1)));
    return email;
  }

  async getIdentities(): Promise<any[]> {
    const session = await this.getSession();

    // No `properties` filter: RFC 8620 section 5.1 defines an omitted/null `properties`
    // as "return every property of the object", so this already fetches the full RFC 8621
    // section 6 Identity — including textSignature/htmlSignature, which list_identities
    // surfaces (#33). Naming properties explicitly would be the narrowing change, not the
    // widening one: it would cap raw/verbose output at whatever list we hard-coded and
    // drop any Fastmail-specific property those modes return today.
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Identity/get', {
          accountId: session.accountId
        }, 'identities']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0);
  }

  async getDefaultIdentity(): Promise<any> {
    const identities = await this.getIdentities();
    
    // Find the default identity (usually the one that can't be deleted)
    return identities.find((id: any) => id.mayDelete === false) || identities[0];
  }

  async createDraft(email: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    mailbox?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
    forwardedMessageId?: string[];
    sourceEmailId?: string;
    attachments?: AttachmentPart[];
  }): Promise<string> {
    const session = await this.getSession();

    // Validate at least one meaningful field is present (zero-width/whitespace-only
    // bodies count as absent). Attachments count as content too: an attachment-only
    // draft is a valid artifact (and is consistent with edit_draft, which preserves a
    // body-less draft that carries attachments).
    if (!email.to?.length && !email.subject && isBlank(email.textBody) && isBlank(email.htmlBody) && !email.attachments?.length) {
      throw new InvalidInputError('At least one of to, subject, textBody, htmlBody, or attachments must be provided');
    }

    // Get all identities to resolve from address
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (email.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, email.from!));
      if (!selectedIdentity) {
        throw new InvalidInputError('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
    }

    const fromEmail = email.from || selectedIdentity.email;

    // Resolve the save target. Fetch the mailbox list unconditionally now (a name/role
    // needs it, and an explicit id is validated against it too) and share it. An unknown
    // mailbox throws InvalidInputError; otherwise default to the Drafts mailbox.
    const mailboxes = await this.getMailboxes();
    let draftMailboxId: string;
    if (email.mailbox) {
      draftMailboxId = resolveMailbox(mailboxes, email.mailbox).id;
    } else {
      const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
      if (!draftsMailbox) {
        throw new Error('Could not find Drafts mailbox');
      }
      draftMailboxId = draftsMailbox.id;
    }

    const mailboxIds: Record<string, boolean> = {};
    mailboxIds[draftMailboxId] = true;

    const emailObject: any = {
      mailboxIds,
      keywords: { $draft: true },
      from: [{ name: selectedIdentity.name, email: fromEmail }],
    };

    if (email.to?.length) emailObject.to = email.to.map(parseAddress);
    if (email.cc?.length) emailObject.cc = email.cc.map(parseAddress);
    if (email.bcc?.length) emailObject.bcc = email.bcc.map(parseAddress);
    if (email.subject) emailObject.subject = email.subject;
    if (email.inReplyTo?.length) emailObject.inReplyTo = email.inReplyTo;
    if (email.references?.length) emailObject.references = email.references;
    if (email.replyTo?.length) emailObject.replyTo = email.replyTo.map(parseAddress);
    // The forwarded original's Message-ID (forward_email). A header SET, unlike every
    // other header use in this file (which are GETs) — Fastmail accepts it and
    // round-trips it through store/fetch and the edit recreate (probed live
    // 2026-07-05). The value is pre-vetted by the forward handler; Fastmail itself
    // rejects CRLF/non-ASCII, so no injection is possible here.
    if (email.forwardedMessageId?.length) emailObject['header:X-Forwarded-Message-Id:asMessageIds'] = email.forwardedMessageId;
    // The exact stored instance this draft was composed from (reply_email /
    // forward_email pass the fetched original's own id). Vetted here — the single
    // seam that sets the header — so a malformed value degrades to absent (send_draft
    // falls back to the Message-ID lookup) instead of failing the whole create.
    if (isSettableSourceId(email.sourceEmailId)) emailObject[SOURCE_ID_HEADER] = email.sourceEmailId;
    if (email.attachments?.length) emailObject.attachments = email.attachments;
    // Generate the body parts (auto text/plain fallback for html-only input where
    // derivable; ships html-only otherwise; no-body html is rejected by shapeBodies). A
    // draft with neither body is allowed — shapeBodies returns empty shaping in that case.
    Object.assign(emailObject, this.shapeBodies(email.textBody, email.htmlBody));

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject }
        }, 'createDraft']
      ]
    };

    const response = await this.makeRequest(request);

    const result = this.getMethodResult(response, 0);

    // Propagate server-provided error details from notCreated
    if (result.notCreated?.draft) {
      throw new Error(`Failed to create draft: ${this.describeSetError(result.notCreated.draft)}`);
    }

    // Throw if created ID is missing instead of returning silently
    const emailId = result.created?.draft?.id;
    if (!emailId) {
      throw new Error('Draft creation returned no email ID');
    }

    return emailId;
  }

  // Extract a stored body value by MIME type, keyed into bodyValues by partId.
  //
  // Server behaviour (verified live against Fastmail, 2026-06-23):
  //  - The server does NOT auto-generate the missing text/html partner in either
  //    direction; the client owns keeping the pair in sync.
  //  - A single-format draft has its ONE part aliased into BOTH the textBody and
  //    htmlBody lists (e.g. a text-only draft lists the text/plain part under htmlBody
  //    too, with type "text/plain"). So we select by the part's actual MIME type — not
  //    mere presence in a list — otherwise we'd read the text value into the html slot
  //    and synthesise a phantom text/html part on recreate.
  //  - JMAP body properties are immutable (RFC 8621 §4.1), which is why updateDraft
  //    rebuilds and re-sends the bodies via a recreate rather than patching.
  // Takes the first part of the given type (drafts here carry at most one per type). If a
  // value were ever elided from bodyValues, that format reads as undefined rather than a
  // partial body (callers fetch full values, so this won't occur in practice).
  private bodyValueForType(parts: any[] | undefined, mimeType: string, bodyValues: Record<string, any>): string | undefined {
    const part = parts?.find((p: any) => p.type === mimeType && p.partId != null && bodyValues[p.partId]);
    return part ? bodyValues[part.partId].value : undefined;
  }

  // Generate the JMAP body-part shaping for the authoring path (createDraft) to splat
  // into the email object. normalizeBodies derives the text/plain fallback from html
  // when none was supplied; we degrade gracefully — ship html-only when the html has
  // visible media but no derivable text — and reject ONLY a genuinely no-body message
  // (html present that renders to nothing AND has no image). A message with neither body
  // returns empty shaping (a body-less draft is
  // allowed; the no-body reject only fires when an html body was actually provided).
  private shapeBodies(textBody?: string, htmlBody?: string) {
    const normalized = normalizeBodies({ textBody, htmlBody });
    if (normalized.htmlOnly && !htmlHasVisibleContent(htmlBody!)) {
      throw new InvalidInputError('This message has no readable body; add text or visible content.');
    }
    return buildBodyParts(normalized);
  }

  async updateDraft(emailId: string, updates: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    replyTo?: string[];
    clearFields?: string[];
    attachments?: AttachmentPart[];
    removeAttachments?: string[];
    // Reply-quote preservation on body edit (#37, redesigned #42). If a reply draft already
    // carries the quoted original (detected on its EXISTING body, which this server generated)
    // and the edit would touch the body in a way that could drop the quote, the caller must
    // say what to do: originalEmailId = regenerate and keep the quote from that (caller-named)
    // message; noQuote = deliberately drop it. Absent both, the edit is rejected (no silent loss).
    originalEmailId?: string;
    noQuote?: boolean;
  }, options: {
    // Whether this server can attach files at all (FASTMAIL_ATTACH_DIR is set). It changes
    // only which repair a refusal offers: pointing at "add an attachments item" when the
    // server would then refuse to read one is a dead end. Defaults to the attachment-capable
    // wording, which is right for every caller that never had the gate to begin with.
    attachmentsEnabled?: boolean;
  } = {}): Promise<UpdateDraftResult> {
    // The caller-supplied-body guard for edit_draft. It lives here, unlike the other four
    // compose paths (which guard in their handlers), because updateDraft is edit_draft's
    // only caller — so `updates` IS the caller's own input, before this method regenerates
    // a reply quote or forwarded block from it, and this is where the rest of the body
    // rules already live. createDraft can't host it: it is shared with the reply and
    // forward paths, which reach it with an already-merged body.
    assertBodyInputs(updates);

    const session = await this.getSession();

    // Fetch the existing email
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'textBody', 'htmlBody', 'bodyValues', 'mailboxIds', 'keywords', 'inReplyTo', 'references', 'attachments', 'header:X-Forwarded-Message-Id:asMessageIds', SOURCE_ID_HEADER],
          // The full part properties: the faithful recreate carries a part's metadata, and
          // `cid`/`disposition`/`type` are what tell it which parts the body displays.
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const existingEmail = this.getListResult(getResponse, 0)[0];
    if (!existingEmail) {
      throw new InvalidInputError(`Email with ID '${emailId}' not found`);
    }

    // Verify it's a draft
    if (!existingEmail.keywords?.$draft) {
      throw new InvalidInputError('Cannot edit a non-draft email');
    }

    const availability: AttachmentAvailability = {
      attachmentsEnabled: options.attachmentsEnabled !== false,
    };

    // The part set this edit works from is the UNION of the JMAP attachments array and the
    // media parts the server routed into the body lists. The same embedded image lands in
    // one list or the other depending on the message's MIME shape (RFC 8621 §4.1.4), so
    // reading `attachments` alone would make a body-routed image invisible to the carry —
    // it would simply vanish from the recreated draft — and would let the forward guard
    // re-arm on a draft whose attached .eml the server listed as body content.
    const storedParts: any[] = buildUnionParts(existingEmail).map((u: UnionPart) => u.part);

    // Faithful-recreate guard. The recreate below rebuilds the message from flat convenience
    // props (textBody/htmlBody/attachments). That reproduces every blob-backed part along
    // with its embedded-image linkage — Fastmail assembles the multipart/related shape from
    // a flat create (probed live 2026-08-14) — but it has no spelling for a part it cannot
    // re-reference, nor for a body that alternates between two parts of the same text type.
    // Those two shapes are refused loudly rather than mangled (#13, #85).
    const bodyShape = classifyDraftBodyShape(existingEmail);
    if (bodyShape.uncarriablePart) {
      const { part, isMedia } = bodyShape.uncarriablePart;
      throw new InvalidInputError(rejectUncarriableBodyPart(part?.name, part?.type, isMedia));
    }
    if (bodyShape.interleavedTextType) {
      throw new InvalidInputError(rejectInterleavedTextParts());
    }

    // Resolve identity
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (updates.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, updates.from!));
      if (!selectedIdentity) {
        throw new InvalidInputError('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      // Use existing from, or fall back to default identity
      const existingFrom = existingEmail.from?.[0]?.email;
      if (existingFrom) {
        selectedIdentity = identities.find(id => matchesIdentity(id.email, existingFrom))
          || identities.find(id => id.mayDelete === false) || identities[0];
      } else {
        selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
      }
    }

    // Extract existing body values by MIME type (see bodyValueForType for the
    // MIME-match-not-list-presence rationale; bodies are immutable so we recreate).
    const bodyValues = existingEmail.bodyValues || {};
    const existingTextValue = this.bodyValueForType(existingEmail.textBody, 'text/plain', bodyValues);
    const existingHtmlValue = this.bodyValueForType(existingEmail.htmlBody, 'text/html', bodyValues);

    // Strict empty-reject + explicit clearFields. A provided-but-empty value is a
    // loud error (it's almost always an accidental clobber); deliberately blanking
    // a field is done by naming it in clearFields. Every field is clearable EXCEPT
    // `from` (identity-resolved; a draft always has a sender, matching the Fastmail UI).
    const CLEARABLE = new Set(['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments']); // NOT 'from'
    const SETTABLE = ['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'from'] as const;
    const provided = new Set<string>(SETTABLE.filter(f => (updates as any)[f] !== undefined));
    // `attachments` isn't a string SETTABLE field, so add it to `provided` explicitly
    // when an attachment add/remove was requested. This is the seam that makes
    // validateClearFields throw "can't set and clear attachments" if the caller also
    // passes clearFields:['attachments'] (clear-then-append in one call is ambiguous).
    if (updates.attachments?.length || updates.removeAttachments?.length) provided.add('attachments');
    validateClearFields(updates.clearFields, CLEARABLE, provided);
    const clear = new Set(updates.clearFields ?? []);

    // The `!clear.has(f)` guard is belt-and-suspenders: validateClearFields already
    // throws when a field is BOTH provided and cleared, so a cleared field can't also
    // reach these checks — the skip just keeps each check self-evidently correct.
    const clearHint = 'omit to leave it unchanged, or list it in clearFields to clear it';
    if (updates.subject  !== undefined && !clear.has('subject'))  requireNonEmpty(updates.subject,  'subject',  clearHint);
    if (updates.textBody !== undefined && !clear.has('textBody')) requireNonEmpty(updates.textBody, 'textBody', clearHint);
    if (updates.htmlBody !== undefined && !clear.has('htmlBody')) requireNonEmpty(updates.htmlBody, 'htmlBody', clearHint);
    if (updates.from     !== undefined) requireNonEmpty(updates.from, 'from'); // not clearable; no hint about clearFields
    for (const f of ['to', 'cc', 'bcc', 'replyTo'] as const) {
      if (updates[f] !== undefined && !clear.has(f) && updates[f]!.length === 0) {
        throw new InvalidInputError(`${f} cannot be empty; ${clearHint}`);
      }
    }
    // Body requireNonEmpty calls above are GUARDS ONLY — their trimmed return is
    // discarded so stored bodies keep their exact (untrimmed) value below.

    // ---- Quoted-original / forwarded-block preservation guard (#37, redesigned #42,
    // extended to forward drafts #30) ----
    // A reply draft keeps the quoted original in its body (buildReplyBodies appends a cited
    // <blockquote> to html and a "> "-quoted block to text); a forward draft keeps the
    // forwarded-message block (buildForwardBodies). An edit that rewrites or clears that
    // body would silently drop it. We decide on the EXISTING (stored) body — which THIS
    // server generated, so its shape is reliable — never on the caller's NEW body
    // (untrusted: it can't tell a real quote from any quote-shaped content, and prose can
    // false-positive). When the draft is protected and the edit touches the body in a way
    // that isn't preserving by construction, force the caller to choose: regenerate+keep
    // from a caller-named originalEmailId, or deliberately noQuote. Supersedes the fork.8
    // new-body-scan (bypassable + html-only); see docs/email-bodies.md, #42, #30.
    const isReply = !!existingEmail.inReplyTo?.length;
    const oldHtmlQuoted = hasQuoteMarker(existingHtmlValue);
    const oldTextQuoted = hasTextQuoteMarker(existingTextValue);

    // Forward gating: the X-Forwarded-Message-Id header — set by this server's
    // forward_email AND by Fastmail's own clients — is a challenge FLOOR (protects a
    // draft whose block shape isn't recognizable: foreign header-setting clients, or a
    // marker lost to re-serialization), while the body markers ALSO gate on their own
    // (protects forwards of a Message-ID-less original, where no header could be set).
    // Markers additionally decide the challenge wording (recognizable = regenerable).
    // The bare header does NOT arm when the draft carries a message/rfc822 attachment:
    // that is an asAttachment forward (which records the header for send_draft's
    // keyword maintenance), whose forwarded content lives in the attached .eml — a
    // body edit can't drop it, and the recreate carries both the attachment and the
    // header. Residual: a foreign draft with BOTH an unrecognizable inline block and
    // an .eml attachment loses the block's challenge, but the .eml still preserves
    // the forwarded content in full, so nothing is unrecoverable.
    // Read off the part UNION, not the attachments array: which list a server routes the
    // .eml into is a MIME-shape accident, and the carve-out has to hold either way.
    const forwardHeader: string[] = existingEmail['header:X-Forwarded-Message-Id:asMessageIds'] || [];
    const emlAttached = storedParts.some((p: any) => classifyPartType(p?.type) === 'message/rfc822');
    const isForward = forwardHeader.length > 0 && !emlAttached;
    const oldHtmlForwarded = hasForwardMarker(existingHtmlValue);
    const oldTextForwarded = hasTextForwardMarker(existingTextValue);
    const forwardMarkers = oldHtmlForwarded || oldTextForwarded;

    // Mutually exclusive dispatch: compute both, run at most ONE variant — the keep path
    // MUTATES updates.htmlBody/textBody and falls through, so two sequential blocks would
    // rebuild the body twice. Both engaged → reply wins: on a pathological both-marker
    // draft an originalEmailId keep regenerates the REPLY shape — loud, no silent loss,
    // but a tie-break, not a claim of correctness. (A forward OF a reply carries
    // "wrote:\n> " in its reproduced text but has no inReplyTo, so isReply stays false
    // and it correctly dispatches to the forward variant; asAttachment forwards record
    // the header for send_draft's keyword maintenance but keep a deliberately
    // non-marker filler body, and the .eml carve-out above keeps the guard inert on
    // them — their .eml is an ordinary carried attachment the recreate preserves,
    // and the recreate carries the header too.)
    const replyGuardArmed = isReply && (oldHtmlQuoted || oldTextQuoted);
    const forwardGuardArmed = isForward || forwardMarkers;
    const guardVariant: 'reply' | 'forward' | null =
      replyGuardArmed ? 'reply' : (forwardGuardArmed ? 'forward' : null);

    // Pre-merge signals: what the edit does to each body. existingHtmlValue/existingTextValue
    // were fetched above; these are the same inputs the coupling guards below use.
    const wroteHtml = updates.htmlBody !== undefined;
    const wroteText = updates.textBody !== undefined;
    const clearedHtml = clear.has('htmlBody');
    const clearedText = clear.has('textBody');
    const touchesBody = wroteHtml || wroteText || clearedHtml || clearedText;

    // The caller's OWN html, snapshotted before the keep path can replace it with a rebuilt
    // body. The checks that must not see this server's own output — an authored reference to
    // a server-managed identifier is an error, while the rebuilt quote is full of them by
    // design — read this, never updates.htmlBody.
    const callerWrittenHtml = updates.htmlBody;

    // Which parts this call takes off the draft, resolved HERE because the quote rebuild
    // below has to know what survives it: a surviving part supplies its own embedded image,
    // and a removed one has to be re-attached under a fresh identifier. Resolution is pure
    // and holds its refusal (see AttachmentRemovalPlan.error) so the body-shape guards keep
    // raising first; the removal is APPLIED at the point in the order it always was.
    const removalPlan = resolveAttachmentRemovals(
      storedParts,
      updates.removeAttachments,
      clear.has('attachments'),
    );

    // Embedded-image references the STORED body already makes with nothing to supply them.
    // A draft in that state is broken however it got that way, and the refusal it earns is
    // scoped to edits that write its body — see the check below.
    const storedPartCids = new Set(storedParts.map(partCid).filter((c) => c !== ''));
    const storedDanglingRefs = htmlCidRefs(existingHtmlValue).filter((r) => !storedPartCids.has(r));

    // The quote survives WITHOUT inspecting any new content in exactly two shapes:
    //  - a metadata-only edit (no body written or cleared) leaves both bodies untouched;
    //  - a plain-text conversion (clearFields:['htmlBody'] alone) keeps the old text, but only
    //    counts as quote-preserving when that surviving text actually carries the variant's
    //    own text marker ("> " quote for replies, the dashed forwarded-message line for
    //    forwards — always true for drafts this server made). Checking the OTHER variant's
    //    marker here would misfire a plain-text conversion of a forward draft. If the
    //    surviving text has no marker, this is NOT a clean carve-out and the edit correctly
    //    falls through to the guard below.
    const oldTextMarked = guardVariant === 'forward' ? oldTextForwarded : oldTextQuoted;
    const quoteKeptByConstruction =
         !touchesBody
      || (clearedHtml && !wroteHtml && !wroteText && !clearedText && oldTextMarked);

    // Text-side edits while a non-empty html survives are owned by the two coupling guards
    // further down (textBody-alone; clearFields:['textBody']-while-html), which emit the
    // correct remedy and (for guard ii) need the post-merge mergedHtmlRaw, so they can't move
    // up here. Exclude those cases so this pre-merge guard doesn't pre-empt them. On a text-
    // only draft existingHtmlValue is blank, so this is false → a text edit there correctly
    // falls through to the guard (exactly the #42 case this guard exists to catch).
    const coupledTextEdit =
      !wroteHtml && !clearedHtml && !isBlank(existingHtmlValue) && (wroteText || clearedText);

    // Set when a noQuote drop resolves the challenge: a deliberate de-quote is also a
    // de-forward, so the recreate below skips the X-Forwarded-Message-Id carry
    // REGARDLESS of which variant handled it — otherwise a reply-variant noQuote on a
    // pathological both-marker draft would strand the header and the NEXT body edit
    // would be re-challenged by the header floor. One step must fully de-arm.
    let dropForwardHeader = false;
    // The embedded images a rebuilt quote will display. Filled only on the keep path: the
    // pure builders see the ORIGINAL but not this draft's surviving parts, so the decision
    // of which surviving part supplies which reference is made here and the finished map is
    // handed to them. They rewrite; they never mint.
    let mintedParts: MintedInlinePart[] = [];
    let quoteImageMappings: CidMapping[] = [];
    let rebuiltQuote = false;
    // How the surrounding refusal or note names the block being kept or dropped.
    let keepNoun = 'the quote';
    // The recorded source the recreate carries. A forward-variant keep RE-POINTS this
    // to the fetched original's own Message-ID (when settable): the caller may name a
    // DIFFERENT original than the recorded one (the blessed correcting-a-wrong-original
    // case), and carrying the stale id would leave the guard's advertised recovery
    // pointer rebuilding the block for the wrong message on a later edit.
    let carriedForwardHeader: string[] = forwardHeader;
    // The recorded source INSTANCE (X-Fastmail-MCP-Source-Id, the exact JMAP id
    // send_draft marks) rides the same carry: re-pointed with the forward header on a
    // keep, dropped with it on a de-forwarding noQuote. On a REPLY draft a noQuote
    // drops only the quote text — In-Reply-To survives, the draft is still a reply to
    // that same instance — so the instance pointer survives with it.
    const storedSourceId = existingEmail[SOURCE_ID_HEADER];
    let carriedSourceId: string | undefined =
      typeof storedSourceId === 'string' && storedSourceId.trim() !== '' ? storedSourceId.trim() : undefined;
    if (guardVariant && touchesBody && !quoteKeptByConstruction && !coupledTextEdit) {
      keepNoun = guardVariant === 'reply' ? 'the quote' : 'the forwarded block';
      if (updates.originalEmailId && updates.noQuote === true) {
        throw new InvalidInputError(`Pass either originalEmailId (keep ${keepNoun}) or noQuote (discard it), not both.`);
      } else if (updates.originalEmailId) {
        // Regenerate from the caller-named original — never re-resolved from the draft's
        // In-Reply-To / X-Forwarded-Message-Id (attacker-controllable), so there's no spoof
        // surface. The id is trusted, not validated against the draft's headers (that check
        // would false-reject legitimate cases, e.g. correcting a wrong original).
        // getEmailById throws on not-found; rethrow with a message naming the param so the
        // caller can fix it (it surfaces via index's error wrap like the other guards).
        let original: any;
        try {
          original = await this.getEmailById(updates.originalEmailId);
        } catch {
          const relation = guardVariant === 'reply' ? 'replies to' : 'forwards';
          throw new InvalidInputError(`originalEmailId '${updates.originalEmailId}' could not be fetched (no such message, or not accessible). Pass the id of the message this draft ${relation}.`);
        }
        // Regenerate the quote/block into EVERY body the edit is writing — both, when the
        // caller supplies both (a new html + a custom text alternative), so neither side
        // silently loses it on the keep path. The builders emit into exactly the formats
        // passed. A clear-only edit writes neither body, so there's nowhere to regenerate
        // into — reject loudly rather than silently no-op the keep intent. This pre-empts
        // the downstream no-body reject for a clear-the-last-body edit (the caller sees the
        // regenerate message, not the no-body one); both are loud and lose no data.
        if (wroteHtml || wroteText) {
          // Resolve the rebuilt quote's embedded images BEFORE building it, but ONLY when
          // this edit writes an html body: an embedded image needs an html body to display
          // it (the mail server refuses an inline part without one), so a text-only keep
          // mints nothing and quotes exactly as it always did. Each reference the original's
          // html makes is matched to one of the original's parts; a part already on this
          // draft that survived the removals above and carries an identifier of this
          // server's own shape supplies it under that SAME identifier, so an ordinary edit
          // does not renumber images a client has already rendered. Anything unmatched is
          // carried under a fresh identifier and attached as its own assembly step.
          let quoteCidMap: Map<string, string> | undefined;
          if (wroteHtml) {
            const resolvedQuoteImages = buildCidMap({
              refs: htmlCidRefs(readQuotableHtml(original)),
              sourceParts: buildUnionParts(original).map((u: UnionPart) => u.part),
              // Every surviving part is offered; buildCidMap keeps only the ones carrying an
              // identifier of this server's own shape.
              survivors: removalPlan.survivors as CidPart[],
            });
            quoteCidMap = resolvedQuoteImages.cidMap;
            mintedParts = resolvedQuoteImages.minted;
            quoteImageMappings = resolvedQuoteImages.mappings;
          }
          rebuiltQuote = true;
          const rebuilt = guardVariant === 'reply'
            ? buildReplyBodies({
                original,
                ...(wroteHtml && { htmlBody: updates.htmlBody }),
                ...(wroteText && { textBody: updates.textBody }),
                quoteOriginal: true,
                ...(quoteCidMap && { cidMap: quoteCidMap }),
              })
            : buildForwardBodies({
                original,
                ...(wroteHtml && { htmlBody: updates.htmlBody }),
                ...(wroteText && { textBody: updates.textBody }),
                ...(quoteCidMap && { cidMap: quoteCidMap }),
              });
          if (wroteHtml) updates.htmlBody = rebuilt.htmlBody;
          if (wroteText) updates.textBody = rebuilt.textBody;
          if (guardVariant === 'reply') {
            // Loud-fail a self-inconsistent keep request: the caller asked to KEEP via
            // originalEmailId, but the named message has no quotable content (attachment-only /
            // calendar-only / cid-image-only), so buildReplyBodies passed the body through
            // unquoted. This is reachable only by naming the WRONG/empty original — a draft naming
            // its own original can't hit it (a quote exists only if that original was quotable, and
            // JMAP message content is immutable). It loses no caller input (the new body is kept);
            // it just turns a confusing quote-less result into an actionable error instead of a
            // silent one. The `||` accepts the edit if ANY written format kept a marker, so a
            // partially-quotable original still keeps.
            const restored = (wroteHtml && hasQuoteMarker(updates.htmlBody))
              || (wroteText && hasTextQuoteMarker(updates.textBody));
            if (!restored) {
              throw new InvalidInputError(`originalEmailId '${updates.originalEmailId}' has no quotable content (e.g. an attachment-only or calendar-only message), so the quote can't be restored. Check the id, or use noQuote to drop the quote deliberately.`);
            }
          }
          // Forward variant: no restored-check — buildForwardBodies always emits at least
          // the header block into every written body, so the check would be tautological.
          // Flip side (intentional asymmetry, inherent to forwarding): a WRONG
          // originalEmailId here produces a valid-looking block for the wrong message with
          // no loud fail — any original yields a block — unlike the reply path's restored
          // catch. No caller data is lost either way.
          if (guardVariant === 'forward') {
            // Re-point the recorded source at the original the block was just rebuilt
            // from (the same pre-vet forward_email applies; a malformed/absent id keeps
            // the existing carry — stale beats stripped, since the header is also the
            // guard's challenge floor). This also records the source on a marker-only
            // draft that had no header, upgrading its later challenges to the runnable
            // standard recipe.
            const repointedId = original?.messageId?.[0];
            if (isSettableMessageId(repointedId)) carriedForwardHeader = [repointedId];
            // Keep the instance pointer consistent with the Message-ID it just
            // re-pointed: both now name the fetched original. (The reply variant
            // deliberately re-points neither — its In-Reply-To stays as stored, and
            // the instance pointer stays consistent with THAT.)
            if (isSettableSourceId(original?.id)) carriedSourceId = original.id;
          }
        } else {
          const regenNoun = guardVariant === 'reply' ? 'a quote' : 'the forwarded block';
          throw new InvalidInputError(`originalEmailId can't regenerate ${regenNoun} on a body you're not writing — edit the body (htmlBody or textBody) to keep ${keepNoun}, or use noQuote to drop it.`);
        }
      } else if (updates.noQuote === true) {
        // Proceed: the quote/block is dropped on explicit request.
        dropForwardHeader = true;
      } else if (guardVariant === 'reply') {
        // Error names ONLY the data-preserving keep path; noQuote is deliberately omitted so
        // the model is never nudged toward discarding the quote (it stays in the schema).
        throw new InvalidInputError("Editing this reply draft's body would drop the quoted original. Pass originalEmailId (the message it replies to) to keep the quote. If you only have the draft, resolve the original from its In-Reply-To Message-ID via search_emails first.");
      } else if (isForward && forwardMarkers) {
        // The normal forward case: recorded source + recognizable block. Recipe mirrors the
        // reply guard's (resolve the id, then keep). The recovery deliberately says plain
        // get_email — forwardedMessageId is always returned by it, not gated behind a
        // verbose knob. Search the BARE id: Fastmail's full-text lookup matches the
        // Message-ID without its angle brackets (verified live 2026-07-05).
        throw new InvalidInputError("Editing this forward draft's body would drop the forwarded-message block. Pass originalEmailId (the message it forwards) to keep the block. If you only have the draft, read its forwardedMessageId via get_email, then find that message with search_emails (search the bare id, without angle brackets).");
      } else if (forwardMarkers) {
        // Marker only, no recorded source (a forward of a Message-ID-less original, or
        // pasted forwarded-looking content — an accepted false-positive cost; see
        // docs/email-bodies.md). Leads with what happened, readable by a caller that never
        // forwarded anything, and offers noQuote first: there may be nothing to rebuild.
        throw new InvalidInputError("This draft's body matches a forwarded-message marker. Pass noQuote:true to drop that block and proceed with your edit, or originalEmailId (the message it forwards) to rebuild the block.");
      } else {
        // Header floor: the source IS recorded but the block isn't in a shape this server
        // can regenerate in place (a foreign client's forward, or our marker lost to
        // re-serialization). The standard recipe IS runnable; noQuote also clears the
        // forward marking so later edits aren't re-challenged.
        throw new InvalidInputError("This draft is marked as a forward (it records the forwarded message's id), but its forwarded block isn't in a shape this server can regenerate in place. Pass originalEmailId to rebuild the block from that message (read the draft's forwardedMessageId via get_email, then find it with search_emails using the bare id, without angle brackets), or noQuote:true to drop the block and the forward marking (later edits won't be re-challenged).");
      }
    }
    // Honor an explicit noQuote even when the guard never armed — an asAttachment
    // forward (guard inert via the .eml carve-out) still carries a recorded source,
    // and noQuote's documented meaning includes clearing the forward marking; leaving
    // the header would have send_draft mark the original against the caller's stated
    // intent. This also covers metadata-only edits (documented on the noQuote param).
    // Redundant (same assignment) when the armed guard already handled it. The
    // keep-vs-drop exclusivity is enforced here too: the armed guard's identical
    // check is unreachable on an unengaged draft, and silently obeying half of a
    // contradictory request would be worse than either behavior.
    if (updates.noQuote === true && updates.originalEmailId) {
      throw new InvalidInputError('Pass either originalEmailId (keep the quoted or forwarded content) or noQuote (discard it), not both.');
    }
    if (updates.noQuote === true) dropForwardHeader = true;

    // Merge non-body fields: updates override existing; clearFields force the empty value.
    const mergedSubject = clear.has('subject') ? '' : (updates.subject !== undefined ? updates.subject : (existingEmail.subject || ''));
    const mergedTo      = clear.has('to')      ? [] : (updates.to      !== undefined ? updates.to.map(parseAddress)      : (existingEmail.to || []));
    const mergedCc      = clear.has('cc')      ? [] : (updates.cc      !== undefined ? updates.cc.map(parseAddress)      : (existingEmail.cc || []));
    const mergedBcc     = clear.has('bcc')     ? [] : (updates.bcc     !== undefined ? updates.bcc.map(parseAddress)     : (existingEmail.bcc || []));
    const mergedReplyTo = clear.has('replyTo') ? [] : (updates.replyTo !== undefined ? updates.replyTo.map(parseAddress) : (existingEmail.replyTo || null));

    // ---- Body pipeline: one-sided guard + text-fallback generation ----
    // The text part is a DERIVED fallback when html is present. So:
    //  - editing htmlBody alone REGENERATES the text fallback from the new html (no throw);
    //  - editing textBody alone (while a non-empty html survives) is rejected — it won't
    //    change what most recipients render (the html), and the fallback is auto-managed;
    //  - a metadata-only edit (no body written) stays body-invariant (both bodies kept).
    // wroteText/wroteHtml are computed once at the reply-quote guard above (same values; the
    // originalEmailId path only ever replaces an already-written body with another, so their
    // truth doesn't change). Reuse them here.
    const wroteAnyBody = wroteText || wroteHtml;

    // Raw merge: a written body drops the unwritten partner (single-format intent);
    // a no-body edit preserves both; clearFields force the body absent.
    const mergedTextRaw = clear.has('textBody') ? undefined
      : (updates.textBody !== undefined ? updates.textBody
      : (wroteAnyBody ? undefined : existingTextValue));
    const mergedHtmlRaw = clear.has('htmlBody') ? undefined
      : (updates.htmlBody !== undefined ? updates.htmlBody
      : (wroteAnyBody ? undefined : existingHtmlValue));

    // Guard: editing textBody alone while a non-empty htmlBody survives (checked against
    // the EXISTING html, since the raw merge has already dropped the unwritten partner).
    if (wroteText && !wroteHtml && !clear.has('htmlBody') && !isBlank(existingHtmlValue)) {
      throw new InvalidInputError('editing textBody alone won\'t change what most recipients see (they render htmlBody). To change the message, edit htmlBody (the text fallback regenerates automatically); to save a custom plain-text alternative, supply htmlBody alongside it; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }

    // Guard: clearFields:['textBody'] while htmlBody survives — the text fallback is
    // managed automatically (regenerated from html, or html-only if none is derivable), so
    // clearing it on its own is rejected. Evaluated against the MERGED html and BEFORE the
    // fallback step runs (else that step would silently refill it). Allowed when html is also cleared.
    if (clear.has('textBody') && !clear.has('htmlBody') && !isBlank(mergedHtmlRaw)) {
      throw new InvalidInputError('textBody can\'t be cleared on its own while htmlBody is present — the text fallback is managed automatically (regenerated from htmlBody, or html-only if none can be derived). Omit textBody from clearFields; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }

    // Generate the text fallback, but ONLY when a body was actually written — a
    // metadata-only edit must stay body-invariant (it must NOT inject a text part into an
    // html-only draft). When html was (re)written without text, regenerate the text fallback
    // from the new html; ship html-only if none is derivable but the html has visible
    // content; reject a genuinely no-body result.
    let textBodyValue = mergedTextRaw;
    let htmlBodyValue = mergedHtmlRaw;
    if (wroteAnyBody) {
      const normalized = normalizeBodies({ textBody: mergedTextRaw, htmlBody: mergedHtmlRaw });
      textBodyValue = normalized.textBody;
      htmlBodyValue = normalized.htmlBody;
      if (normalized.htmlOnly && !htmlHasVisibleContent(mergedHtmlRaw!)) {
        throw new InvalidInputError('This message has no readable body; add text or visible content.');
      }
    }

    // Reject a body-less RESULT, but only when this edit actually touched the body —
    // a written body that came out empty, or a cleared body. An attachment-only (or
    // any metadata-only) edit must NOT trip this: it stays body-invariant and may run
    // against a draft that legitimately has no body yet. Gating on
    // `wroteAnyBody || clearedAnyBody` (not `wroteAnyBody` alone) keeps the throw firing
    // when the last body is cleared — incl. alongside an attachments change — so a
    // caller can't silently strip a draft down to no body. A draft keeps >=1 body.
    // Distinct from the clear-text-while-html guard (which only fires when merged html IS
    // present), so the two can't both match.
    const clearedAnyBody = clear.has('textBody') || clear.has('htmlBody');
    if ((wroteAnyBody || clearedAnyBody) && isBlank(textBodyValue) && isBlank(htmlBodyValue)) {
      throw new InvalidInputError('a draft needs a body; supply textBody or htmlBody (this edit would leave it with neither).');
    }

    // ---- Part assembly: apply the removals resolved earlier, work out what the surviving
    // parts still do for the body that actually ships, then append (#13) ----

    // The removal refusal is raised HERE, at the position the removal loop always occupied,
    // so a body-shape error keeps its precedence over an attachment one. Resolution happened
    // before the quote rebuild because the rebuild needs to know what survives it.
    if (removalPlan.error) throw removalPlan.error;

    const bodyTouching = wroteAnyBody || clearedAnyBody;
    const finalHtmlRefs = htmlCidRefs(htmlBodyValue);

    // A stored Content-ID this server cannot reproduce faithfully is a refusal rather than a
    // silent mangle. Evaluated AFTER the removals, so a call that takes the offending part
    // off the draft is not blocked by it, and only on an edit that touches the body: the
    // carry copies such a value through verbatim (measured round-tripping exactly against
    // Fastmail, 2026-08-14), so metadata and attachment edits are already safe and refusing
    // them would cost the caller their only in-place repair.
    if (bodyTouching) {
      for (const part of removalPlan.survivors) {
        const cid = partCid(part);
        if (cid && !isRecreatableCid(cid)) throw new InvalidInputError(rejectUnrecreatableCid(cid));
      }
    }

    // What becomes of each surviving part now the shipping body is known: still displayed →
    // unchanged; no longer displayed and carrying one of this server's own identifiers →
    // taken off (it was only ever there to show an image the body no longer shows); no
    // longer displayed but someone else's → demoted to an ordinary attachment rather than
    // dropped, because those bytes are not this server's to discard. Runs only when this
    // call wrote or cleared a body — a metadata or attachment edit is body-invariant by
    // design, so nothing about the parts' relationship to the body can have changed.
    const reconciled = bodyTouching
      ? reconcileInlineParts({
          storedParts: removalPlan.survivors as CidPart[],
          referencedCids: finalHtmlRefs,
          htmlShips: !isBlank(htmlBodyValue),
        })
      : null;

    const ledger = new InlineNoteLedger();
    const carriedParts: AttachmentPart[] = [];
    if (reconciled) {
      reconciled.parts.forEach(({ part, action }, index) => {
        // Keyed by POSITION, not by blob. Two parts can share one blobId — the same bytes
        // displayed twice under different Content-IDs is a real shape (and one Fastmail
        // stores as two distinct parts), and the ledger replaces by key, so a blob-keyed
        // record would collapse the pair into one and report half of what came off the
        // draft. The disclosure has to count parts, because parts are what the reader loses.
        const key = `part:${index}`;
        if (action === 'removed') {
          ledger.record({ key, outcome: 'removed', name: (part as any).name });
          return;
        }
        if (action === 'degraded') {
          ledger.record({ key, outcome: 'degraded', name: (part as any).name });
        }
        carriedParts.push(carriedPartFrom(part, action === 'degraded'));
      });
    } else {
      for (const part of removalPlan.survivors) carriedParts.push(carriedPartFrom(part));
    }

    // Assembly order: what survived, then the caller's own additions, then the parts minted
    // for the rebuilt quote. Minted parts ride their own channel to the very end so they
    // stay distinguishable from anything carried or supplied.
    let finalAttachments: AttachmentPart[] = carriedParts.slice();
    if (updates.attachments?.length) {
      // A file the caller supplied with a Content-ID is marked inline exactly when the
      // shipping body references it, which is only knowable here: the upload happens before
      // this call and cannot see the body the edit ends up with (the caller's html, or the
      // draft's own when the edit leaves the body alone). A Content-ID never makes a part
      // inline on its own, so without this an image the caller asked to embed would arrive
      // as an ordinary attachment. The reverse case — a supplied Content-ID nothing
      // references — stays an attachment and is reported as such rather than dropped.
      const displays = (p: AttachmentPart) =>
        !isBlank(htmlBodyValue) && partCid(p) !== '' && finalHtmlRefs.includes(partCid(p));
      finalAttachments = finalAttachments.concat(
        updates.attachments.map((p, index) => {
          if (displays(p)) return { ...p, disposition: 'inline' };
          // A file supplied WITH a Content-ID that the edited body does not display is the
          // caller asking to embed something and not getting it. It is kept — those bytes
          // are theirs — but the outcome is recorded so the result says what became of it,
          // exactly as the compose tools do. Keyed by position rather than by identifier,
          // so a file the body DOES display (recorded under its identifier further down)
          // can never be collapsed into the same record. A file supplied with no
          // Content-ID asked for nothing and needs no sentence.
          if (partCid(p) !== '') {
            ledger.record({ key: `supplied:${index}`, outcome: 'degraded', name: p.name });
          }
          return p;
        }),
      );
    }
    if (mintedParts.length) finalAttachments = finalAttachments.concat(mintedParts);

    // ---- Checks over the assembled state, in the order a caller can act on ----

    const finalPartCids = new Set(
      finalAttachments.map((p) => (typeof p.cid === 'string' ? p.cid : '')).filter((c) => c !== ''),
    );
    const danglingRefs = finalHtmlRefs.filter((r) => !finalPartCids.has(r));

    // Removing an image the kept quote supplies cannot be honoured at all: the rebuild would
    // put it straight back under a fresh identifier, and a removal that silently does nothing
    // is worse than a refusal. Pruning individual quote images is not a thing this server can
    // do — dropping the quote with a replacement body is the way out, which the message says.
    // Scoped to a NAMED removal. Wiping the attachments wholesale alongside a kept quote is
    // coherent and supported: the stored quote parts go, the rebuild re-embeds under fresh
    // identifiers, and the result says both happened.
    if (rebuiltQuote && !clear.has('attachments') && removalPlan.removed.length > 0) {
      const reEmbeddedBlobs = new Set(mintedParts.map((p) => p.blobId));
      if (removalPlan.removed.some((p) => reEmbeddedBlobs.has(p.blobId))) {
        throw new InvalidInputError(rejectRemovalOfQuoteCarriedPart());
      }
    }

    // A reference left pointing at nothing BY THIS CALL'S REMOVAL, named as such. This is a
    // diff, not a body scan: a reference that was already broken before the call is not
    // attributed here, so removing part C never reports a pre-existing break on part B.
    const removedCids = new Set(removalPlan.removed.map(partCid).filter((c) => c !== ''));
    const removalCaused = danglingRefs.filter((r) => removedCids.has(r));
    if (removalCaused.length > 0) {
      throw new InvalidInputError(
        clear.has('attachments')
          ? rejectClearAttachmentsDanglingRefs()
          : rejectRemovalDanglingRef(removalCaused[0]),
      );
    }

    if (bodyTouching) {
      // The draft's own stored body already referenced images nothing supplies. The draft's
      // state dominates: this wording wins over anything the caller's html got wrong in the
      // same call, because recreating the draft is the repair either way. Checked POST-MERGE
      // on purpose — an edit that replaces the body and eliminates the references passes, and
      // so does one that adds the attachment supplying them.
      const stillBroken = danglingRefs.filter((r) => storedDanglingRefs.includes(r));
      if (stillBroken.length > 0) throw new InvalidInputError(rejectBrokenDraft(stillBroken, availability));

      // Scoped to the caller's OWN html: the rebuilt quote references server-managed
      // identifiers by construction, so scanning the merged body here would refuse every
      // keep-edit of a draft that carries quoted images.
      for (const ref of htmlCidRefs(callerWrittenHtml)) {
        if (isReservedCid(ref)) throw new InvalidInputError(rejectReservedCidRef(ref));
      }

      if (danglingRefs.length > 0) {
        throw new InvalidInputError(rejectDanglingCidRef(danglingRefs[0], availability));
      }
    }

    // Two parts sharing one identifier make every reference to it ambiguous, so a colliding
    // addition is refused before it can be uploaded into the draft. The in-call wording
    // interpolates the real count — three items sharing an identifier must not read "two".
    const appendedCids = (updates.attachments ?? [])
      .map((p) => (typeof p.cid === 'string' ? p.cid : ''))
      .filter((c) => c !== '');
    const appendedCidCounts = new Map<string, number>();
    for (const cid of appendedCids) appendedCidCounts.set(cid, (appendedCidCounts.get(cid) ?? 0) + 1);
    for (const [cid, n] of appendedCidCounts) {
      if (n > 1) throw new InvalidInputError(rejectCidCollisionInCall(n, cid));
    }
    const carriedCids = new Set(carriedParts.map((p) => p.cid ?? '').filter((c) => c !== ''));
    for (const cid of appendedCidCounts.keys()) {
      if (carriedCids.has(cid)) throw new InvalidInputError(rejectCidCollisionOnDraft(cid));
    }

    // Self-check, not a caller-facing rule: every reference in a body this call produced
    // resolves to a part the message carries, and every part it minted is referenced by one
    // of those bodies. Nothing above can be true and this false, so a failure means the
    // assembly itself is wrong.
    checkInlineClosure({
      htmlBodies: bodyTouching ? [htmlBodyValue] : [],
      finalPartCids: finalAttachments.map((p) => p.cid),
      attachedMintedCids: mintedParts.map((p) => p.cid),
      skip: !bodyTouching && mintedParts.length === 0,
    });

    // What the shipping body displays, counted once per part whatever supplied it: a part
    // already on the draft, one the caller just added, or one minted for the rebuilt quote.
    // The identifier IS the key, so a part that is both reused and re-referenced is one
    // record, never two.
    //
    // Reported only when this call could have CHANGED what the body displays — it wrote or
    // cleared a body, or it supplied parts. An edit that touches neither is body-invariant by
    // design, and announcing what such a draft has always embedded would be noise on every
    // subject change.
    //
    // The basis is REFERENCE MEMBERSHIP, not the part's stored disposition: a body displays
    // an image by resolving its Content-ID, and a part carrying the referenced identifier is
    // displayed whatever disposition it arrived with (the assembly above marks it inline for
    // exactly that reason). The compose paths read the disposition instead, which is the same
    // answer there because every part is one this call just uploaded and dispositioned. They
    // diverge only for a part already ON the draft as a plain attachment that the edited body
    // starts referencing — a shape compose cannot reach, and one this basis gets right.
    const reportsEmbeds = bodyTouching || !!updates.attachments?.length;
    const bytesByCid = new Map<string, number>();
    for (const part of storedParts) {
      if (partCid(part)) bytesByCid.set(partCid(part), partBytes(part));
    }
    for (const mapping of quoteImageMappings) bytesByCid.set(mapping.cid, partBytes(mapping.source));
    const embeddedNow = reportsEmbeds
      ? finalAttachments.filter(
          (p) => typeof p.cid === 'string' && p.cid !== '' && finalHtmlRefs.includes(p.cid),
        )
      : [];

    const emailObject: any = {
      mailboxIds: existingEmail.mailboxIds,
      // Preserve all existing keywords (e.g. $flagged, custom labels), not just $draft.
      keywords: { ...(existingEmail.keywords || {}), $draft: true },
      from: [{ name: selectedIdentity.name, email: updates.from || existingEmail.from?.[0]?.email || selectedIdentity.email }],
      to: mergedTo,
      cc: mergedCc,
      bcc: mergedBcc,
      subject: mergedSubject,
      ...(mergedReplyTo?.length && { replyTo: mergedReplyTo }),
      // Threading: carry inReplyTo/references as JMAP structured properties so the
      // In-Reply-To/References headers regenerate (fixes silent threading loss on reply
      // drafts this client creates via reply_email send=false).
      ...(existingEmail.inReplyTo && { inReplyTo: existingEmail.inReplyTo }),
      ...(existingEmail.references && { references: existingEmail.references }),
      // Carry the forward marking (the recorded source AND the guard's challenge floor)
      // unless a noQuote drop deliberately cleared it; a forward-variant keep may have
      // re-pointed or newly recorded it above. A foreign stored value that fails
      // Fastmail's header-SET validation fails the CREATE loudly with the old draft
      // intact (create-first order below) — recoverable in one step via noQuote, never
      // a silent drop.
      ...(carriedForwardHeader.length > 0 && !dropForwardHeader && { 'header:X-Forwarded-Message-Id:asMessageIds': carriedForwardHeader }),
      // The exact-instance pointer follows the provenance it refines: dropped when a
      // noQuote de-forwards a FORWARD draft (the forward header above goes with it),
      // kept on a reply draft (noQuote drops the quote text, but In-Reply-To — and so
      // the reply itself, and the instance it replies to — survives).
      ...(carriedSourceId !== undefined && !(dropForwardHeader && !isReply) && { [SOURCE_ID_HEADER]: carriedSourceId }),
      ...(finalAttachments.length && { attachments: finalAttachments }),
    };

    Object.assign(emailObject, buildBodyParts({ textBody: textBodyValue, htmlBody: htmlBodyValue }));

    // What this edit is about to replace, captured BEFORE the recreate so the caller can
    // detect an overwrite it didn't intend (#65). See ReplacedDraftInfo for why this is a
    // fingerprint rather than the previous content.
    const addressList = (addrs: any): string[] =>
      (addrs || []).map((a: any) => a?.email).filter((e: any): e is string => typeof e === 'string' && e !== '');
    const replacedTo = addressList(existingEmail.to);
    const replacedCc = addressList(existingEmail.cc);
    const replacedDraft: ReplacedDraftInfo = {
      id: emailId,
      ...(existingEmail.subject && { subject: existingEmail.subject }),
      ...(replacedTo.length && { to: replacedTo }),
      ...(replacedCc.length && { cc: replacedCc }),
      ...(existingTextValue !== undefined && { textBodySize: existingTextValue.length }),
      ...(existingHtmlValue !== undefined && { htmlBodySize: existingHtmlValue.length }),
    };

    // Create-then-dispose (NOT a single combined create+destroy call). JMAP content is
    // immutable — verified live 2026-06-24 that Fastmail SILENTLY NO-OPS an in-place
    // subject/body update (returns success but changes nothing), so a recreate is
    // mandatory. RFC 8620 §6.3 guarantees blob lifetime within a call but says NOTHING
    // about create/destroy atomicity: a server MAY apply the destroy even when the create
    // lands in notCreated, which would vanish the draft. So we create FIRST, confirm it
    // succeeded, and only THEN dispose of the old draft. Worst case is a harmless duplicate
    // (recoverable), never a vanished draft (unrecoverable).
    // WARNING: do NOT "optimize" this back into one Email/set call; that reintroduces the
    // data-loss window.
    const createRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject },
        }, 'createDraft']
      ]
    };

    const createResponse = await this.makeRequest(createRequest);
    const createResult = this.getMethodResult(createResponse, 0);

    if (createResult.notCreated?.draft) {
      // Create failed → old draft is untouched (no destroy was issued). This is the
      // data-loss-prevention path: surface the error, leave the draft as-is.
      throw new Error(`Failed to create updated draft: ${this.describeSetError(createResult.notCreated.draft)}`);
    }

    const newEmailId = createResult.created?.draft?.id;
    if (!newEmailId) {
      throw new Error('Draft update returned no email ID');
    }

    // The new draft is valid → the edit has SUCCEEDED. Now dispose of the old copy by
    // moving it to TRASH — never Email/set destroy (#65). An edit made from a stale copy
    // of the draft (it was changed in the web UI in between, say) silently overwrites
    // whatever the caller didn't know about, and a destroy made that unrecoverable: the
    // old draft was gone from the account with only a support-side backup restore left.
    // A Trash move makes the same mistake a one-step undo, and whatever Trash retention
    // the account is set to bounds the clutter. Only mailboxIds is patched; keywords
    // (including $draft) are left alone, matching delete_email — a trashed draft is still
    // a draft, so restoring it is a move back to Drafts, and Trash retention applies by
    // mailbox, not by keyword. getThread compensates: a $draft whose only mailbox is Trash
    // is not counted as an active draft, so repeated edits can't inflate its hidden-draft
    // warning.
    //
    // Any failure here — no Trash mailbox, a structured notUpdated, or a thrown
    // transport/method error — leaves a duplicate holding the OLD pre-edit content.
    // Report it as an orphan warning with its reason, but do NOT throw: the edit itself
    // succeeded, and throwing would tell the caller it failed. There is deliberately NO
    // destroy fallback — destroying is the exact unrecoverable act this path exists to
    // avoid, and a visible duplicate the caller can delete is strictly cheaper.
    let trashedOldDraftId: string | undefined;
    let orphanedOldDraftId: string | undefined;
    let orphanedOldDraftReason: string | undefined;
    try {
      // EXACT role only (case-insensitive), the same rule delete_email uses: a custom
      // folder merely NAMED like a trash folder must never become the destination.
      // (getMailboxes is uncached, so this costs one extra round trip per edit —
      // accepted for a write path that already makes several.)
      const mailboxes = await this.getMailboxes();
      const trashMailbox = this.findByExactRole(mailboxes, 'trash');
      if (!trashMailbox) {
        orphanedOldDraftId = emailId;
        orphanedOldDraftReason = 'this account has no mailbox with the trash role';
      } else {
        const trashResponse = await this.makeRequest({
          using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
          methodCalls: [
            ['Email/set', {
              accountId: session.accountId,
              update: { [emailId]: { mailboxIds: { [trashMailbox.id]: true } } },
            }, 'trashOldDraft']
          ]
        });
        // Confirm the move POSITIVELY: RFC 8620 §5.3 puts every requested id in exactly
        // one of updated/notUpdated, so absence from both means the server told us
        // nothing — and inferring success from "not in notUpdated" would report
        // "recoverable in Trash" for a draft still sitting in Drafts. `updated` maps
        // id -> object|null, so the null case makes a truthiness test wrong; test for
        // the key.
        const trashResult = this.getMethodResult(trashResponse, 0);
        if (trashResult.notUpdated?.[emailId]) {
          orphanedOldDraftId = emailId;
          orphanedOldDraftReason = this.describeSetError(trashResult.notUpdated[emailId]);
        } else if (Object.prototype.hasOwnProperty.call(trashResult.updated ?? {}, emailId)) {
          trashedOldDraftId = emailId;
        } else {
          orphanedOldDraftId = emailId;
          orphanedOldDraftReason = 'the server did not report the move as applied';
        }
      }
    } catch (err) {
      orphanedOldDraftId = emailId;
      orphanedOldDraftReason = err instanceof Error ? err.message : String(err);
    }

    // Confirm what the saved draft carries, but ONLY when this call attached or minted an
    // embedded image. The flat create is what assembles the multipart shape that makes an
    // image display, so an embed is the one outcome the result text would otherwise assert
    // without evidence — while an ordinary edit gets no extra round trip. Its own try/catch
    // and its own sentence: the edit has already succeeded here, so nothing this step finds
    // may fail it. The confirmed sizes replace the ones known at assembly time, which is
    // where a freshly uploaded part's size comes from at all.
    const followUpNotes: string[] = [];
    const attachedInlineCids = embeddedNow
      .map((p) => p.cid as string)
      .filter((cid) => mintedParts.some((m) => m.cid === cid) || appendedCidCounts.has(cid));
    if (attachedInlineCids.length > 0) {
      try {
        const savedParts = await this.readDraftParts(newEmailId);
        const savedByCid = new Map<string, any>();
        for (const part of savedParts) {
          if (partCid(part)) savedByCid.set(partCid(part), part);
        }
        const missing = attachedInlineCids.filter((cid) => !savedByCid.has(cid));
        if (missing.length > 0) followUpNotes.push(noteEmbedMissingAfterSave(missing.length));
        for (const [cid, part] of savedByCid) {
          if (partBytes(part) > 0) bytesByCid.set(cid, partBytes(part));
        }
      } catch {
        followUpNotes.push(noteEmbedUnconfirmed());
      }
    }

    for (const part of embeddedNow) {
      const cid = part.cid as string;
      ledger.record({
        key: `cid:${cid}`,
        outcome: 'embedded',
        bytes: bytesByCid.get(cid) ?? 0,
        name: part.name,
        isImage: true,
      });
    }

    const tally = ledger.tally();
    const notes = [
      ...emitInlineNotes(tally, {
        surface: 'draft',
        keepNoun,
        ...(clear.has('attachments') && rebuiltQuote && {
          clearedAttachmentCount: removalPlan.removed.length,
          reEmbeddedCount: quoteImageMappings.length,
        }),
      }),
      ...followUpNotes,
    ];
    const touchedInlineImages = tally.embedded > 0 || tally.degraded > 0 || tally.removed > 0;

    return {
      id: newEmailId,
      replacedDraft,
      ...(trashedOldDraftId && { trashedOldDraftId }),
      ...(orphanedOldDraftId && { orphanedOldDraftId, orphanedOldDraftReason }),
      ...(touchedInlineImages && {
        inlineImages: { embedded: tally.embedded, degraded: tally.degraded, removed: tally.removed },
      }),
      ...(notes.length > 0 && { notes }),
    };
  }

  /**
   * Re-read a draft's part listing, for confirming what a just-saved draft actually carries.
   *
   * Deliberately a listing read and nothing else: no body values, no keywords, no guard.
   * Callers use it after a successful write, where the only question left is whether the
   * parts landed.
   */
  private async readDraftParts(emailId: string): Promise<any[]> {
    const session = await this.getSession();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'attachments', 'textBody', 'htmlBody'],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
        }, 'getEmail']
      ]
    });
    const email = this.getListResult(response, 0)[0];
    if (!email) throw new Error(`Email with ID '${emailId}' not found`);
    return buildUnionParts(email).map((u: UnionPart) => u.part);
  }

  async sendDraft(emailId: string): Promise<SendDraftOutcome> {
    const session = await this.getSession();

    // Fetch the existing email to verify it's a draft. The two provenance headers ride
    // along on this existing read: they say which message the draft replies to or
    // forwards, and they can't change during submission (the draft is submitted by
    // reference, not recreated), so reading them here rather than after the send costs
    // nothing and keeps a failed read indistinguishable from the draft-not-found path
    // this method already has.
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: [
            'id', 'from', 'to', 'cc', 'bcc', 'replyTo', 'keywords', 'textBody', 'htmlBody', 'bodyValues',
            'attachments',
            'inReplyTo', 'header:X-Forwarded-Message-Id:asMessageIds', SOURCE_ID_HEADER,
          ],
          // `attachments` above and disposition/cid/name here are for the RECEIPT ONLY —
          // the sentence below that reports what the sent message carried. They are never a
          // send-time vet: this method submits the stored draft by reference, exactly as it
          // is, and adds no check that could refuse a message the caller already approved.
          // Anything that would refuse a draft belongs on the edit path, which can offer a
          // repair; a refusal here would only strand a finished message.
          bodyProperties: ['partId', 'blobId', 'type', 'size', 'disposition', 'cid', 'name'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];
    if (!email) {
      throw new InvalidInputError(`Email with ID '${emailId}' not found`);
    }

    if (!email.keywords?.$draft) {
      throw new InvalidInputError('Cannot send a non-draft email');
    }

    // Reject an empty body part before an irreversible send. sendDraft submits the draft by
    // reference WITHOUT recreating it, so unlike edit_draft it can't truthy-gate the body away.
    // An empty/whitespace body part renders blank; an empty text/html part even shadows a real
    // text/plain (RFC 2046: clients render the richest alternative), so the recipient sees
    // nothing. We never emit such a part, but an externally-created draft (Fastmail web UI, etc.)
    // can carry one — refuse to ship it. Reject, not silent sanitize: recreating to strip the part
    // would change the email id and rewrite the message, against this codebase's loud-reject
    // philosophy. The hardened edit_draft is the fix path.
    const textVal = this.bodyValueForType(email.textBody, 'text/plain', email.bodyValues || {});
    const htmlVal = this.bodyValueForType(email.htmlBody, 'text/html', email.bodyValues || {});
    if (htmlVal !== undefined && htmlVal.trim() === '') {
      throw new InvalidInputError('This draft has an empty htmlBody that would render blank to recipients. Edit the draft to supply or clear htmlBody before sending.');
    }
    if (textVal !== undefined && textVal.trim() === '') {
      throw new InvalidInputError('This draft has an empty textBody that would render blank for plain-text recipients. Edit the draft to supply or clear textBody before sending.');
    }

    // Collect all recipients for the envelope
    const allRecipients: { email: string }[] = [
      ...(email.to || []),
      ...(email.cc || []),
      ...(email.bcc || []),
    ];

    if (allRecipients.length === 0) {
      throw new InvalidInputError('Draft has no recipients. Edit the draft to add a to/cc/bcc recipient before sending.');
    }

    // Determine identity from the email's from field
    const fromEmail = email.from?.[0]?.email;
    if (!fromEmail) {
      throw new InvalidInputError('Draft has no from address. Edit the draft to set a from address before sending.');
    }

    const identities = await this.getIdentities();
    const selectedIdentity = identities.find(id => matchesIdentity(id.email, fromEmail));
    if (!selectedIdentity) {
      throw new InvalidInputError('From address on draft does not match any sending identity. Edit the draft to set a from address matching one of your verified identities before sending.');
    }

    // Find the Sent mailbox
    const mailboxes = await this.getMailboxes();
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox');
    }

    const sentMailboxIds: Record<string, boolean> = {};
    sentMailboxIds[sentMailbox.id] = true;

    // Submit the draft
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['EmailSubmission/set', {
          accountId: session.accountId,
          create: {
            submission: {
              emailId,
              identityId: selectedIdentity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: allRecipients.map(addr => ({ email: addr.email })),
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              'keywords/$draft': null,
              'keywords/$seen': true,
            }
          }
        }, 'submitDraft']
      ]
    };

    const response = await this.makeRequest(request);
    const submissionResult = this.getMethodResult(response, 0);
    if (submissionResult.notCreated?.submission) {
      throw new Error(`Failed to submit draft: ${this.describeSetError(submissionResult.notCreated.submission)}`);
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Draft submission returned no submission ID');
    }

    // The receipt: what the message that just went out actually displayed inline. Computed
    // after the submission because it reports on a completed send and must never influence
    // one — an image counted wrong here changes a sentence, never whether the mail ships.
    // An embedded image is one the body displays, which the server signals either by routing
    // the part into a body list or by marking it inline; a message with neither says nothing.
    const embedded = buildUnionParts(email).filter(
      (u: UnionPart) => isImagePart(u.part) && (u.inBodyList || u.part?.disposition === 'inline'),
    );
    const embeddedBytes = embedded.reduce((sum, u) => sum + partBytes(u.part), 0);

    return {
      submissionId,
      sourceReferences: readSourceReferences(email),
      ...(embedded.length > 0 && { notes: [noteSentWithEmbedded(embedded.length, embeddedBytes)] }),
    };
  }

  /**
   * Resolve an RFC 5322 Message-ID to the JMAP id(s) of the message(s) that CARRY it as
   * their own Message-ID. Header values name messages by Message-ID, but every write path
   * needs a JMAP id, so this is the bridge.
   *
   * Two steps, because neither alone is right. The full-text query is the recall step:
   * Fastmail matches a message by its BARE Message-ID (the <bracketed> form finds nothing,
   * probed live 2026-07-05 — see docs/email-bodies.md), but it is a text search, so it also
   * returns every message that merely mentions the id — replies whose In-Reply-To/References
   * carry it, quoted bodies, and the sent copy that referenced it in the first place. The
   * exact `messageId` comparison is the precision step that keeps only the message the id
   * actually belongs to. (RFC 8621 section 4.4.1's `header` FilterCondition also works on
   * Fastmail and would do this server-side, but it is unproven against other JMAP servers,
   * and the two-step form needs no such assumption.)
   *
   * The OLDEST-FIRST sort is load-bearing, not cosmetic. The recall step is capped, and in a
   * busy thread the id can be mentioned by more messages than the cap — every later reply
   * carries it in References, every quoted body repeats it. A message can only reference an
   * id that already exists, so the message that OWNS the id predates all of them: sorting
   * oldest-first puts the owner on the first page whatever the cap is. Without the sort the
   * owner can fall off the page and the caller is told, wrongly, that no message carries it.
   *
   * No Trash/Spam exclusion: this answers "which message is this", not "what should a search
   * show", and a message being in Trash does not make it the wrong one.
   */
  async findEmailIdsByMessageId(messageId: string): Promise<string[]> {
    const bare = String(messageId ?? '').trim().replace(/^<+/, '').replace(/>+$/, '').trim();
    if (!bare) return [];

    const session = await this.getSession();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: { text: bare },
          sort: [{ property: 'receivedAt', isAscending: true }],
          limit: MESSAGE_ID_LOOKUP_LIMIT,
        }, 'query'],
        ['Email/get', {
          accountId: session.accountId,
          '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
          properties: ['id', 'messageId'],
        }, 'emails'],
      ],
    });

    return this.getListResult(response, 1)
      .filter((e: any) => Array.isArray(e?.messageId) && e.messageId.includes(bare))
      .map((e: any) => e.id);
  }

  /**
   * Read the Message-ID list of one stored message by its JMAP id; null when no such
   * message exists (destroyed, or a foreign/hand-set pointer). send_draft's keyword
   * maintenance uses this to validate the draft's recorded source INSTANCE before
   * marking it — the instance must still exist and still carry the Message-ID the
   * draft's provenance header names, otherwise the caller falls back to the
   * Message-ID lookup above rather than marking whatever the stale pointer hits.
   */
  async getEmailMessageId(emailId: string): Promise<string[] | null> {
    const session = await this.getSession();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'messageId'],
        }, 'getSourceInstance'],
      ],
    });
    const email = this.getListResult(response, 0)[0];
    if (!email) return null;
    return Array.isArray(email.messageId) ? email.messageId : [];
  }

  async getRecentEmails(limit: number = 10, mailbox: string = 'inbox', ascending: boolean = false, position: number = 0): Promise<QueryResult> {
    const session = await this.getSession();

    // Resolve the target mailbox EXACTLY (id/role/name) — replaces the old substring
    // match, so this stays consistent with the #12 sweep and carries no substring
    // injection-steering primitive. A blank/whitespace mailbox falls back to the inbox
    // default (matching the resolveMailboxId blank handling the swept tools use), rather
    // than throwing. Throws InvalidInputError on a non-blank unknown mailbox.
    const target = mailbox && mailbox.trim() ? mailbox : 'inbox';
    const mailboxes = await this.getMailboxes();
    const targetMailbox = resolveMailbox(mailboxes, target);

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
    };

    emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];

    const query: any = {
      accountId: session.accountId,
      filter: { inMailbox: targetMailbox.id },
      sort: [{ property: 'receivedAt', isAscending: ascending }],
      limit: Math.min(limit, 50),
      calculateTotal: true
    };
    // Paging offset (#51), sent only when non-zero — 0 is the JMAP default, so
    // position:0 and an omitted position are the same request. This tool caps `limit`
    // at 50, so `position` is how a caller reads past the cap.
    if (position > 0) query.position = position;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', query, 'query'],
        ['Email/get', emailGetParams, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getQueryResult(response, 0, 1);
    if (typeof result.position !== 'number') result.position = position;
    // Reuse the mailbox list already fetched above to resolve names + roles — no extra methodCall.
    attachMailboxInfo(result.items, buildMailboxInfoMap(mailboxes));
    return result;
  }

  async markEmailRead(emailId: string, read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = read
      ? { 'keywords/$seen': true }
      : { 'keywords/$seen': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'updateEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], `mark email as ${read ? 'read' : 'unread'}`);
    }
  }

  // Additively set one or more keyword flags on an email (e.g. $answered, $seen) without
  // clobbering the others: a per-keyword `keywords/${k}: true` patch. Mirrors markEmailRead
  // (a single Email/set routing a notUpdated entry through throwSingleSetError). Used by the
  // reply path for best-effort thread-state maintenance (#52/#54).
  async addKeywords(emailId: string, keywords: string[]): Promise<void> {
    const session = await this.getSession();

    const patch: Record<string, any> = {};
    keywords.forEach(k => {
      patch[`keywords/${k}`] = true;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'addKeywords']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], 'add keywords to email');
    }
  }

  async pinEmail(emailId: string, pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = pinned
      ? { 'keywords/$flagged': true }
      : { 'keywords/$flagged': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'pinEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], `${pinned ? 'pin' : 'unpin'} email`);
    }
  }

  async deleteEmail(emailId: string): Promise<void> {
    const session = await this.getSession();

    // Find the trash mailbox by EXACT role only (case-insensitive). NOT the substring
    // findMailboxByRoleOrName: a custom "Trash bin rules" mailbox (no trash role) must
    // never be the delete destination, and computeExclusion's exact-role Trash would
    // then never count mail mis-filed there.
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findByExactRole(mailboxes, 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: {
              mailboxIds: trashMailboxIds
            }
          }
        }, 'moveToTrash']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], 'delete email');
    }
  }

  async moveEmail(emailId: string, target: string): Promise<void> {
    const session = await this.getSession();

    // Resolve the destination EXACTLY (id/role/name) — a new capability (moveEmail
    // previously took a raw id with no resolution). Exact-only, so it carries no
    // substring mis-resolution; deliberate move-to-any-mailbox stays open by design
    // (a move-target restriction is tracked separately, fork #43). Throws
    // InvalidInputError on an unknown destination.
    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // Fetch current mailboxIds to build a proper JMAP patch
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['mailboxIds']
        }, 'getEmail']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];

    // Build patch: remove from all current mailboxes, add to target
    const patch: Record<string, boolean | null> = {};
    if (email?.mailboxIds) {
      for (const mbId of Object.keys(email.mailboxIds)) {
        patch[`mailboxIds/${mbId}`] = null;
      }
    }
    patch[`mailboxIds/${targetMailboxId}`] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'moveEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], 'move email');
    }
  }

  // Resolve each label mailbox input to a real mailbox id by id/role/name (exact —
  // findMailboxExact, no substring), so the label arrays accept the same id/role/name a
  // caller learned works on every other mailbox-taking tool (fork #50). Collects EVERY
  // unresolved input so the caller fixes all typos in one retry. All-or-nothing: if any
  // input can't be resolved, throw InvalidInputError and apply no labels (avoids a
  // half-applied mutation the caller must reconcile). Duplicate inputs (an id and its own
  // name) collapse downstream since the Email/set patch keys by id. A real id absent from
  // the live list is still rejected — accepted residual, see docs/security-model.md.
  private async resolveLabelMailboxIds(inputs: string[]): Promise<string[]> {
    const mailboxes = await this.getMailboxes();
    const resolved: string[] = [];
    const unresolved: string[] = [];
    for (const input of inputs) {
      const mb = findMailboxExact(mailboxes, input);
      if (mb) resolved.push(mb.id);
      else unresolved.push(String(input).trim());
    }
    if (unresolved.length > 0) {
      throw new InvalidInputError(formatMailboxesNotFound(unresolved, mailboxes || []));
    }
    return resolved;
  }

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const resolvedIds = await this.resolveLabelMailboxIds(mailboxIds);

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    resolvedIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'addLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], 'add labels to email');
    }
  }

  async removeLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const resolvedIds = await this.resolveLabelMailboxIds(mailboxIds);

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    resolvedIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'removeLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      this.throwSingleSetError(result.notUpdated[emailId], 'remove labels from email');
    }
  }

  async bulkAddLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const resolvedIds = await this.resolveLabelMailboxIds(mailboxIds);

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    resolvedIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkAddLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, 'add labels to');
    }
  }

  async bulkRemoveLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const resolvedIds = await this.resolveLabelMailboxIds(mailboxIds);

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    resolvedIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkRemoveLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, 'remove labels from');
    }
  }

  async getEmailAttachments(emailId: string): Promise<EmailAttachmentsResult> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          // The body lists are fetched alongside `attachments` because an embedded
          // image is routed into them rather than into `attachments` on some MIME
          // shapes (RFC 8621 §4.1.4) — without them this tool reports nothing for a
          // message that visibly shows a picture (#13).
          properties: ['attachments', 'textBody', 'htmlBody'],
          // Pinned explicitly (was previously the RFC 8621 §4.4 server default, which
          // already includes disposition/cid) so this tool's ability to distinguish
          // inline parts never rides on an implicit default.
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
        }, 'getAttachments']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];
    const attachments = buildUnionParts(email).map((u) => u.part);
    const rawAttachments = email?.attachments || [];
    // buildUnionParts yields the server's own part objects, so identity is an exact
    // membership test for "this part is not in the JMAP attachments array".
    const inRaw = new Set<any>(rawAttachments);
    return {
      attachments,
      rawAttachments,
      omittedFromRaw: attachments.filter((part) => !inRaw.has(part)).length,
    };
  }

  /**
   * Resolve an attachment reference on an email to its blob metadata. The reference may
   * be a partId, a blobId, `cid:<value>` for an embedded image, or a plain entry number
   * counting from the start of the union listing — see resolveAttachmentRef for the
   * fixed precedence between those forms.
   *
   * This is the SINGLE attachment-resolution path in this client: every consumer (URL
   * build, byte fetch, save-to-file) resolves through here exactly once and then passes
   * the resolved info around, so a message's metadata and its bytes can never come from
   * two different reads of the same email.
   *
   * The error split is deliberate. A malformed REFERENCE (and a part carrying no blob)
   * throws InvalidInputError, which download_attachment lets through so the caller can
   * correct what it passed. A well-formed reference that simply matches nothing, and a
   * missing email, stay a plain `Error` that maps to a generic InternalError — that is
   * how a lookup avoids confirming what a mailbox holds.
   */
  async getAttachmentInfo(emailId: string, attachmentId: string): Promise<AttachmentInfo> {
    const session = await this.getSession();

    // Both body lists are fetched as well as `attachments`: an embedded image is
    // routed into them on some MIME shapes, and without them it is not downloadable
    // at all (#13). `bodyValues` used to be requested here and was never read — the
    // request never set fetchTextBodyValues/fetchHTMLBodyValues, so it returned
    // nothing; it is dropped rather than carried as an inert property.
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['attachments', 'textBody', 'htmlBody'],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
        }, 'getEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];

    if (!email) {
      throw new Error('Email not found');
    }

    const parts = buildUnionParts(email).map((u) => u.part);
    const attachment = resolveAttachmentRef(parts, attachmentId);

    if (!attachment) {
      throw new Error('Attachment not found.');
    }

    // Rejected here rather than at URL-build time because resolution is single: this one
    // check covers every consumer (metadata read, URL build, byte fetch, save-to-file),
    // and it is what makes AttachmentInfo's non-optional `blobId` honest. A part with no
    // blob is listable but not fetchable, so naming that is more useful than a later
    // failure on an undefined substituted into the URL.
    if (!attachment.blobId) {
      throw new InvalidInputError(
        'That part has no downloadable content: it carries no blobId. ' +
        'Pick a part from get_email_attachments that has one.'
      );
    }

    return {
      blobId: attachment.blobId,
      type: attachment.type || 'application/octet-stream',
      // The declared filename is sender-supplied, so it is sanitized at resolution, not
      // at URL build: a receiving client may use it as a save name, and every consumer
      // of AttachmentInfo reads this same field. The sanitizer subsumes the old
      // `|| 'attachment'` fallback (it returns that for a missing or unusable name).
      name: sanitizeDownloadFilename(attachment.name),
      size: attachment.size,
    };
  }

  /**
   * Build the blob download URL for an ALREADY-resolved attachment. Taking the resolved
   * info rather than an (emailId, attachmentId) pair is what keeps resolution single:
   * callers that also need the bytes or the metadata reuse the one resolution instead of
   * issuing a second Email/get whose answer could differ.
   */
  private async downloadUrlFor(info: AttachmentInfo): Promise<string> {
    const session = await this.getSession();

    const downloadUrl = session.downloadUrl;
    if (!downloadUrl) {
      throw new Error('Download capability not available in session');
    }

    const url = downloadUrl
      .replace('{accountId}', session.accountId)
      .replace('{blobId}', info.blobId)
      .replace('{type}', encodeURIComponent(info.type))
      .replace('{name}', encodeURIComponent(info.name));

    // Re-validate after substitution, before the URL is used to send the bearer token:
    // the session-time check saw the template, and placeholder values could rewrite the
    // origin. Belt-and-suspenders over that origin check.
    validateFastmailUrl(url, 'downloadUrl', this.auth.getAllowUnsafe());

    return url;
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    return this.downloadUrlFor(await this.getAttachmentInfo(emailId, attachmentId));
  }

  static readonly DEFAULT_DOWNLOADS_DIR = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  static validateSavePath(savePath: string, downloadDir?: string): string {
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    // Resolve relative paths against the allowed download directory rather than
    // the process cwd (which is unpredictable for an MCP server launched by a
    // client). Absolute paths are taken as-is; either way the containment check
    // is the security boundary. So a bare filename lands safely in the configured
    // dir in one step, and an absolute path inside that dir writes exactly there.
    // Write side keeps a byte-exact (case-sensitive) compare so download behaviour
    // is unchanged; the read guard opts into case-insensitive containment on Win32.
    // Throws PathAccessError so the index layer maps it to InvalidParams uniformly.
    return lexicalContainedPath(savePath, allowedDir, false);
  }

  /**
   * Symlink-safe canonicalization of a save path. Walks up to the longest
   * existing ancestor, realpaths it, and verifies it lives under the canonical
   * allowed directory. Refuses to overwrite an existing symlink at the target.
   *
   * Returns the canonical path that is safe to write to. Throws on escape.
   */
  static async safeWritePath(savePath: string, downloadDir?: string): Promise<string> {
    // Lexical pre-check first (cheap and gives nice errors)
    const lexical = JmapClient.validateSavePath(savePath, downloadDir);
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;

    // Ensure allowed dir exists so realpath can resolve it.
    await mkdir(allowedDir, { recursive: true });
    const canonicalAllowed = await realpath(allowedDir);

    // Walk up from the target until we find an existing ancestor.
    let ancestor = dirname(lexical);
    const missingSegments: string[] = [];
    while (true) {
      try {
        await stat(ancestor);
        break;
      } catch (e: any) {
        if (e.code !== 'ENOENT') throw e;
        missingSegments.unshift(basename(ancestor));
        const parent = dirname(ancestor);
        if (parent === ancestor) {
          throw new PathAccessError(`Could not find an existing ancestor for path: ${lexical}`);
        }
        ancestor = parent;
      }
    }

    // Canonicalize the existing ancestor — this is what catches symlink escapes.
    const canonicalAncestor = await realpath(ancestor);
    if (canonicalAncestor !== canonicalAllowed && !canonicalAncestor.startsWith(canonicalAllowed + sep)) {
      throw new PathAccessError(
        `path resolves to '${canonicalAncestor}' which is outside the allowed directory '${canonicalAllowed}'. ` +
        `Refusing to follow symlink escape.`,
      );
    }

    // Reconstruct the safe canonical path under the canonical ancestor.
    const safePath = join(canonicalAncestor, ...missingSegments, basename(lexical));

    // If a symlink already exists at the target, refuse — writing through it
    // would still escape the allowed directory.
    try {
      const lst = await lstat(safePath);
      if (lst.isSymbolicLink()) {
        throw new PathAccessError(`Refusing to overwrite an existing symlink at the target: ${safePath}`);
      }
    } catch (e: any) {
      if (e.code !== 'ENOENT') throw e;
    }

    return safePath;
  }

  // Per-file / aggregate fail-fast guards. NOT authoritative — Fastmail's own ceiling
  // governs; these just bound the in-memory read and reject obviously-too-large inputs
  // before we upload. The per-file cap also bounds the fd read in uploadAttachments.
  static readonly MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
  static readonly MAX_TOTAL_ATTACHMENT_BYTES = 45 * 1024 * 1024;

  /**
   * Read-shaped, handle-based path confinement for the attachment-send capability.
   * Distinct from the write-shaped safeWritePath (which mkdir -p's the root and walks
   * MISSING segments) — a read creates nothing and must validate the OPEN file:
   *
   *  - attachDir undefined → throw the opt-in error BEFORE any fs syscall (the hard
   *    gate; an exfiltration capability stays disabled until the operator sets the var);
   *  - reject the Windows escape shapes and lexically contain the path (case-insensitively
   *    on Win32) against the resolved root;
   *  - open(path,'r') ONCE, then fstat the handle (require a regular file) and realpath
   *    the FULL target (not an ancestor), re-verifying canonical containment. The caller
   *    reads from the returned handle, so the bytes uploaded are the bytes of the file we
   *    validated — TOCTOU is narrowed, not eliminated (see docs/security-model.md).
   *
   * Returns the open handle and its size; the CALLER must close the handle.
   */
  static async safeReadPath(inputPath: string, attachDir: string | undefined): Promise<{ handle: FileHandle; size: number }> {
    if (!attachDir) {
      throw new PathAccessError(
        'Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR to the directory attachable files live in, then restart the server to enable it.'
      );
    }

    rejectWindowsPathEscapes(inputPath);

    const allowedDir = resolve(normalize(attachDir));
    const caseInsensitive = process.platform === 'win32';
    const lexical = lexicalContainedPath(inputPath, allowedDir, caseInsensitive);

    // The attach root itself must exist — a missing root is a config error, reported
    // distinctly from the opt-in gate above (not a raw realpath ENOENT).
    let canonicalAllowed: string;
    try {
      canonicalAllowed = await realpath(allowedDir);
    } catch (e: any) {
      if (e.code === 'ENOENT') {
        throw new PathAccessError(`FASTMAIL_ATTACH_DIR (${allowedDir}) does not exist. Create it or fix the path, then restart.`);
      }
      throw e;
    }

    let handle: FileHandle;
    try {
      handle = await open(lexical, 'r');
    } catch (e: any) {
      if (e.code === 'ENOENT') {
        throw new PathAccessError(`File not found: ${inputPath} (resolved under ${allowedDir}).`);
      }
      if (e.code === 'EISDIR') {
        throw new PathAccessError(`Not a regular file: ${inputPath}.`);
      }
      throw e;
    }

    try {
      const st = await handle.stat();
      if (!st.isFile()) {
        throw new PathAccessError(`Not a regular file: ${inputPath}.`);
      }
      // Re-verify against the canonical full target — this catches a symlinked leaf or
      // an intermediate-dir symlink that escapes the root.
      const canonicalTarget = await realpath(lexical);
      if (!isPathContained(canonicalTarget, canonicalAllowed, caseInsensitive)) {
        throw new PathAccessError(
          `path resolves to '${canonicalTarget}' which is outside the allowed directory '${canonicalAllowed}'. Refusing to follow symlink escape.`
        );
      }
      return { handle, size: st.size };
    } catch (e) {
      await handle.close().catch(() => {});
      throw e;
    }
  }

  /**
   * Upload a single blob and return its server-assigned blobId. POSTs the raw bytes to
   * the {accountId}-substituted session uploadUrl with ONLY Authorization + the given
   * Content-Type — deliberately NOT spreading getAuthHeaders() (which hardcodes
   * application/json) and NOT JSON.stringifying the body (unlike every other call here).
   * The server-returned `type` is authoritative for the stored blob (it echoes the
   * Content-Type we sent — a best-effort hint, not content sniffing).
   */
  async uploadBlob(data: Buffer, contentType: string): Promise<{ blobId: string; type: string; size: number }> {
    const session = await this.getSession();
    if (!session.uploadUrl) {
      throw new Error('Upload capability not available in session');
    }

    // Reject an over-limit payload before any network call — the server advertises its
    // own ceiling in the core capability, so a doomed upload need not be sent at all.
    const maxSize = session.capabilities?.['urn:ietf:params:jmap:core']?.maxSizeUpload;
    if (typeof maxSize === 'number' && data.length > maxSize) {
      throw new Error(`Attachment is ${data.length} bytes; the server's upload limit is ${maxSize} bytes`);
    }

    const url = session.uploadUrl.replace('{accountId}', session.accountId);
    // Re-validate after substitution, mirroring the download path: the session-time check
    // saw the template, and the substituted value could rewrite the origin the bearer
    // token is about to be sent to.
    validateFastmailUrl(url, 'uploadUrl', this.auth.getAllowUnsafe());

    // POST the raw bytes — no JSON.stringify, unlike every other call here. The copy
    // constructor `new Uint8Array(data)` yields a concrete Uint8Array<ArrayBuffer> (not the
    // ArrayBufferLike-backed view a Buffer/`.subarray` carries), which IS assignable to
    // fetch's BodyInit — so this stays fully type-checked, no `any` escape hatch.
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': this.auth.getAuthHeaders()['Authorization'],
        'Content-Type': contentType,
      },
      body: new Uint8Array(data),
      // Same rationale as the download path: never follow a redirect on a token-bearing
      // request — the token would be replayed to an unvalidated host.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Blob upload failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json() as any;
    if (!result || typeof result.blobId !== 'string') {
      throw new Error('Blob upload returned no blobId');
    }
    // End-to-end integrity check: the stored blob size the server reports back must equal
    // the byte count we sent. A mismatch means a truncated/corrupted upload (proxy or
    // transport), so fail loudly rather than attach a corrupt file. (Only checked when the
    // server returns a numeric size.)
    if (typeof result.size === 'number' && result.size !== data.length) {
      throw new Error(`Blob upload size mismatch: sent ${data.length} bytes, server stored ${result.size}.`);
    }
    return { blobId: result.blobId, type: result.type || contentType, size: result.size ?? data.length };
  }

  /**
   * Confine, read, and upload each file spec, returning the JMAP attachment parts to
   * splat into an Email. Re-checks the opt-in gate (so a caller that skipped safeReadPath
   * can't bypass it), validates the caller contentType against the MIME token grammar,
   * rejects an over-cap file via fstat.size BEFORE reading, then does a bounded read from
   * the confined handle.
   *
   * Each part is a FRESH literal built here, never the carriedAttachments shape — that one
   * passes through a server-set `size` a strict server rejects. A file with no Content-ID
   * is a 4-key part; a file the caller gave one is a 5-key part carrying its `cid`.
   *
   * The DISPOSITION of a Content-ID-bearing part is the caller's decision expressed through
   * `inlineCids`: a part is marked `inline` only when the message being composed actually
   * displays that identifier. This is not a preference. Fastmail refuses a create that marks
   * any part inline when the message ships no html body at all — it fails the whole call
   * with invalidProperties:["htmlBody"] — so a file the body cannot display has to ride as a
   * regular attachment or nothing is saved. It keeps its Content-ID either way: the caller
   * asked for that identifier, a Content-ID never flips a part inline on its own, and
   * keeping it is what lets a later edit add the html that displays the file.
   */
  async uploadAttachments(
    specs: { path: string; name?: string; contentType?: string; cid?: string }[],
    attachDir: string | undefined,
    options: UploadAttachmentsOptions = {},
  ): Promise<AttachmentPart[]> {
    if (!attachDir) {
      throw new PathAccessError(
        'Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR to the directory attachable files live in, then restart the server to enable it.'
      );
    }

    // Two passes so a confinement/size failure orphans NO blobs. Pass 1 validates and
    // opens every file (path confinement + per-file/total size caps) before a single
    // upload; a bad path or oversize file anywhere in the batch rejects with zero blobs
    // uploaded. Pass 2 reads + uploads the already-validated handles. (A network failure
    // mid-upload can still orphan a blob — unavoidable without server-side transactions;
    // Fastmail garbage-collects unreferenced blobs.)
    const inlineCids = options.inlineCids;
    const opened: { handle: FileHandle; size: number; contentType: string; name: string; cid?: string }[] = [];
    try {
      let totalBytes = 0;
      for (let i = 0; i < specs.length; i++) {
        const spec = specs[i];
        // contentType grammar is validated before any fs work (no handle to leak yet).
        const contentType = spec.contentType
          ? validateContentType(spec.contentType, i)
          : guessContentType(spec.path);
        const { handle, size } = await JmapClient.safeReadPath(spec.path, attachDir);
        // Push BEFORE the size checks so the finally closes this handle even if a cap throws.
        opened.push({ handle, size, contentType, name: spec.name ?? basename(spec.path), cid: spec.cid });
        if (size > JmapClient.MAX_ATTACHMENT_BYTES) {
          throw new PathAccessError(
            `attachments[${i}] (${basename(spec.path)}) is ${size} bytes, over the ${JmapClient.MAX_ATTACHMENT_BYTES}-byte per-file guard. Fastmail's own limit ultimately governs.`
          );
        }
        totalBytes += size;
        if (totalBytes > JmapClient.MAX_TOTAL_ATTACHMENT_BYTES) {
          throw new PathAccessError(
            `attachments total exceeds the ${JmapClient.MAX_TOTAL_ATTACHMENT_BYTES}-byte fail-fast guard. Fastmail's own limit ultimately governs.`
          );
        }
      }

      const parts: AttachmentPart[] = [];
      for (const o of opened) {
        // Bounded read of exactly `size` bytes from the validated handle (never read the
        // whole handle then check .length — that buffers a hostile oversize file first).
        const buffer = Buffer.alloc(o.size);
        const { bytesRead } = await o.handle.read(buffer, 0, o.size, 0);
        const data = bytesRead === o.size ? buffer : buffer.subarray(0, bytesRead);

        const uploaded = await this.uploadBlob(data, o.contentType);
        parts.push({
          blobId: uploaded.blobId,
          type: uploaded.type,
          name: o.name,
          disposition: o.cid && inlineCids?.has(o.cid) ? 'inline' : 'attachment',
          ...(o.cid && { cid: o.cid }),
        });
      }
      return parts;
    } finally {
      for (const o of opened) await o.handle.close().catch(() => {});
    }
  }

  /**
   * Fetch an attachment's bytes into memory, alongside the metadata they were resolved
   * from. One resolution feeds both the URL and the returned name/type, so the bytes and
   * the metadata describing them always come from the same read of the email.
   */
  async fetchAttachmentBuffer(emailId: string, attachmentId: string): Promise<{ buffer: Buffer; url: string } & AttachmentInfo> {
    const info = await this.getAttachmentInfo(emailId, attachmentId);
    const url = await this.downloadUrlFor(info);

    const response = await fetch(url, {
      headers: { 'Authorization': this.auth.getAuthHeaders()['Authorization'] },
      // Never follow a redirect on a token-bearing request: an allowlisted host that
      // 3xx-redirects cross-origin would otherwise source the attachment body from an
      // unvalidated host, with the bearer token replayed to it.
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    return { buffer: Buffer.from(await response.arrayBuffer()), url, ...info };
  }

  async downloadAttachmentToFile(emailId: string, attachmentId: string, savePath: string, downloadDir?: string): Promise<{ url: string; bytesWritten: number; savedPath: string }> {
    // Check + create the parent directory before the (slow) network fetch, so the window
    // in which a co-resident process could swap a checked directory for a symlink is the
    // re-check below rather than the whole download.
    await JmapClient.safeWritePath(savePath, downloadDir);
    const { buffer, url } = await this.fetchAttachmentBuffer(emailId, attachmentId);

    // Re-validate immediately before writing — the target may have been swapped for a
    // symlink during the fetch — then write with O_EXCL so a symlink planted at the path
    // is never followed. Overwriting a pre-existing regular file stays supported: on
    // EEXIST we re-run the symlink-safe check (which refuses a symlink) and replace the
    // plain file. The rewrite ALSO uses 'wx' — a default-flag rewrite would reopen the
    // exact symlink window the unlink just created.
    const safePath = await JmapClient.safeWritePath(savePath, downloadDir);
    await mkdir(dirname(safePath), { recursive: true });
    try {
      await writeFile(safePath, buffer, { flag: 'wx' });
    } catch (e: any) {
      if (e?.code !== 'EEXIST') throw e;
      await JmapClient.safeWritePath(savePath, downloadDir); // refuses a symlink at the target
      await unlink(safePath);
      await writeFile(safePath, buffer, { flag: 'wx' });
    }

    return { url, bytesWritten: buffer.length, savedPath: safePath };
  }

  // Shared engine for searchEmails + getEmails: assemble the filter from a flat base
  // FilterCondition plus a list of single-keyword condition objects, inject the default
  // Trash/Spam exclusion, run the visible query + a hidden-count query, and populate
  // QueryResult.exclusion. Both callers fetch `mailboxes` themselves (to resolve their
  // `mailbox` param + compute the exclusion) and pass it in, so names attach with no
  // extra round-trip and there is no in-batch Mailbox/get.
  private async runFilteredQuery(opts: {
    base: any;
    conds: any[];
    exclusion: ExclusionResult;
    exclusionIntended: boolean;
    limit: number;
    ascending: boolean;
    mailboxes: any[];
    position?: number;
  }): Promise<QueryResult> {
    const session = await this.getSession();
    const { base, conds, exclusion, exclusionIntended, limit, ascending, mailboxes } = opts;
    const position = opts.position ?? 0;

    const doExclude = exclusion.excludeIds.length > 0;
    // Inject the exclusion into `base` BEFORE computing baseEmpty — otherwise an
    // exclusion-only query (no text/from fields) would see base as {} and take the
    // conds[0]-alone branch, silently dropping the folder exclusion (fail-open).
    if (doExclude) base.inMailboxOtherThan = exclusion.excludeIds;

    // Combine the base FilterCondition with N single-keyword conditions. Each keyword
    // is its own condition object because a single JMAP FilterCondition allows only one
    // hasKeyword/notKeyword. baseEmpty alone + one cond -> the lone cond; else AND-wrap.
    const combine = (b: any, c: any[], bEmpty: boolean) =>
      c.length === 0 ? b
      : (bEmpty && c.length === 1) ? c[0]
      : { operator: 'AND', conditions: [...(bEmpty ? [] : [b]), ...c] };

    const baseEmpty = Object.keys(base).length === 0;
    const visibleFilter = combine(base, conds, baseEmpty);

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
      properties: [...EMAIL_PROPERTIES_COMPACT],
    };

    const visibleQuery: any = {
      accountId: session.accountId,
      filter: visibleFilter,
      sort: [{ property: 'receivedAt', isAscending: ascending }],
      limit,
      calculateTotal: true,
    };
    // Paging offset (#51). Sent only when non-zero: 0 is the JMAP default, so an
    // omitted `position` and position:0 are the same request on the wire. The
    // exclusion lives inside `filter`, so every page is filtered server-side and the
    // hidden-count query below (which is a count, not a page) stays unaffected.
    if (position > 0) visibleQuery.position = position;

    const methodCalls: [string, any, string][] = [
      ['Email/query', visibleQuery, 'query'],
      ['Email/get', emailGetParams, 'emails'],
    ];

    if (doExclude) {
      // Hidden-count query = the visible filter with ONLY inMailboxOtherThan removed, so
      // hidden = broaderTotal - visibleTotal = matches withheld to Trash/Spam (the
      // complement, which never overcounts a message cross-filed in {Trash, a visible
      // mailbox} — that message is in both totals). Reconstruct from a COPY of base
      // minus the key, then re-run the identical combine: a naive top-level delete on
      // the assembled filter would no-op when the key sits inside conditions[0] under
      // the AND-wrap (count == visible -> note never fires, fail-open). Issued in the
      // SAME makeRequest at a higher index: one atomic snapshot (no two-query race) and
      // one fewer round-trip; the visible indices 0/1 are unchanged.
      const countBase = { ...base };
      delete countBase.inMailboxOtherThan;
      const countBaseEmpty = Object.keys(countBase).length === 0;
      const countFilter = combine(countBase, conds, countBaseEmpty);
      methodCalls.push(['Email/query', {
        accountId: session.accountId,
        filter: countFilter,
        limit: 0,
        calculateTotal: true,
      }, 'count']);
    }

    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls,
    });
    const result = this.getQueryResult(response, 0, 1);
    if (typeof result.position !== 'number') result.position = position;
    attachMailboxInfo(result.items, buildMailboxInfoMap(mailboxes));

    // Populate exclusion metadata whenever an exclusion was INTENDED — even if
    // excludeIds came back empty (a role couldn't be resolved), so the handler still
    // fires the fail-loud "not excluded" note rather than silently running unfiltered.
    if (exclusionIntended) {
      let hidden: number | null = 0;
      if (doExclude) {
        // FAIL-CLOSED: the published "no note => nothing hidden" contract is only safe
        // if a missing/garbled count fails loud. calculateTotal is server-discretionary
        // and the count methodCall can error. If either total is non-numeric, or the
        // count method errored, or hidden computes negative (a wrong total:0 on the
        // broader query), set hidden=null (degraded note) — never clamp to 0.
        const visibleTotal = result.total;
        let broaderTotal: number | undefined;
        try { broaderTotal = this.getMethodResult(response, 2)?.total; } catch { broaderTotal = undefined; }
        if (typeof visibleTotal !== 'number' || typeof broaderTotal !== 'number') {
          hidden = null;
        } else {
          const h = broaderTotal - visibleTotal;
          hidden = h < 0 ? null : h;
        }
      }
      result.exclusion = {
        hidden,
        excludedRoles: exclusion.excludedRoles,
        unresolvedRoles: exclusion.unresolvedRoles,
      };
    }

    return result;
  }

  async searchEmails(filters: {
    query?: string;
    from?: string;
    to?: string;
    cc?: string;
    bcc?: string;
    subject?: string;
    hasAttachment?: boolean;
    isUnread?: boolean;
    isPinned?: boolean;
    mailbox?: string;
    after?: string;
    before?: string;
    limit?: number;
    position?: number;
    ascending?: boolean;
    excludeDrafts?: boolean;
    includeTrash?: boolean;
    includeSpam?: boolean;
  }): Promise<QueryResult> {
    // Normalise the date bounds before any network work: JMAP wants a UTCDate and
    // rejects a bare YYYY-MM-DD with an `invalidArguments` that names no argument, so
    // an unusable value fails here with one that does (#70).
    const after = coerceUtcDate(filters.after, 'after');
    const before = coerceUtcDate(filters.before, 'before');

    const mailboxes = await this.getMailboxes();
    const resolvedMailboxId = await this.resolveMailboxId(filters.mailbox, mailboxes);

    const base: any = {};
    if (filters.query) base.text = filters.query;
    if (filters.from) base.from = filters.from;
    if (filters.to) base.to = filters.to;
    if (filters.cc) base.cc = filters.cc;
    if (filters.bcc) base.bcc = filters.bcc;
    if (filters.subject) base.subject = filters.subject;
    if (filters.hasAttachment !== undefined) base.hasAttachment = filters.hasAttachment;
    if (after) base.after = after;
    if (before) base.before = before;
    if (resolvedMailboxId) base.inMailbox = resolvedMailboxId;

    // Each keyword is its own condition (mixed polarities can't share one FilterCondition).
    const conds: any[] = [];
    if (filters.isUnread === true) conds.push({ notKeyword: '$seen' });
    else if (filters.isUnread === false) conds.push({ hasKeyword: '$seen' });
    if (filters.isPinned === true) conds.push({ hasKeyword: '$flagged' });
    else if (filters.isPinned === false) conds.push({ notKeyword: '$flagged' });
    if (filters.excludeDrafts) conds.push({ notKeyword: '$draft' });

    const hasExplicitMailbox = !!resolvedMailboxId;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: filters.includeTrash,
      includeSpam: filters.includeSpam,
      hasExplicitMailbox,
    });
    const exclusionIntended = !hasExplicitMailbox && (!filters.includeTrash || !filters.includeSpam);

    return this.runFilteredQuery({
      base,
      conds,
      exclusion,
      exclusionIntended,
      limit: Math.min(filters.limit || 20, 100),
      ascending: filters.ascending ?? false,
      mailboxes,
      position: filters.position,
    });
  }

  async getThread(
    threadId: string,
    includeDrafts: boolean = false,
    includeBodies: boolean = false,
  ): Promise<{ emails: any[]; hiddenDraftCount: number }> {
    const session = await this.getSession();

    // First, check if threadId is actually an email ID and resolve the thread
    let actualThreadId = threadId;

    // Try to get the email first to see if we need to resolve thread ID
    try {
      const emailRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/get', {
            accountId: session.accountId,
            ids: [threadId],
            properties: ['threadId']
          }, 'checkEmail']
        ]
      };

      const emailResponse = await this.makeRequest(emailRequest);
      const email = this.getListResult(emailResponse, 0)[0];

      if (email && email.threadId) {
        actualThreadId = email.threadId;
      }
    } catch (error) {
      // If email lookup fails, assume threadId is correct
    }

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'getThread', name: 'Thread/get', path: '/list/*/emailIds' },
    };

    // Default: the compact list set — metadata + preview, body *structure* only.
    // includeBodies (#74) fetches the text body CONTENT for every message in one call, so
    // a conversation can be read without one get_email per message. It reuses the defined
    // VERBOSE superset rather than a third ad-hoc property list, keeping the two-list
    // discipline documented above (and so raw:true returns a complete JMAP email). The
    // HTML values are deliberately NOT fetched: a thread multiplies body size by the
    // message count and the html alternative is the expensive half.
    if (includeBodies) {
      emailGetParams.properties = [...EMAIL_PROPERTIES_VERBOSE];
      emailGetParams.bodyProperties = [...EMAIL_BODY_PROPERTIES];
      emailGetParams.fetchTextBodyValues = true;
    } else {
      emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];
    }

    // Use Thread/get with the resolved thread ID
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Thread/get', {
          accountId: session.accountId,
          ids: [actualThreadId]
        }, 'getThread'],
        ['Email/get', emailGetParams, 'emails'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name', 'role'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const threadResult = this.getMethodResult(response, 0);

    // Check if thread was found
    if (threadResult.notFound && threadResult.notFound.includes(actualThreadId)) {
      throw new InvalidInputError(`Thread with ID '${actualThreadId}' not found`);
    }

    // Resolve mailbox names onto the FULL list before filtering, so the draft
    // filter below doesn't skip the attach for retained messages.
    const emails = this.getListResult(response, 1);
    const threadMailboxes = this.readListResultIfPresent(response, 2);
    attachMailboxInfo(emails, buildMailboxInfoMap(threadMailboxes));

    // Drafts (e.g. an in-progress reply) are noise when reading a conversation,
    // so exclude them by default. Identify by the $draft keyword (survives a
    // draft moved out of the Drafts mailbox); opt back in via includeDrafts.
    // Return the hidden count (no extra query — derived from the already-fetched
    // thread) so the handler can ANNOUNCE that drafts were hidden without surfacing
    // them (the duplicate-draft trap: an agent reading a thread to reply must not
    // miss that a draft reply already exists). Assumes the full thread is returned
    // (Thread/get -> Email/get, no limit); threads are small so this holds.
    // The filter runs on the whole fetched email object, so under includeBodies a
    // hidden draft's BODY is dropped with it — the count is all that survives.
    if (includeDrafts) {
      return { emails, hiddenDraftCount: 0 };
    }
    const filtered = emails.filter((e: any) => !e.keywords?.$draft);
    // A draft sitting ONLY in Trash is not an active draft, so it must not inflate the
    // count: the note exists to warn that a draft reply already exists, and every
    // edit_draft leaves its replaced copy in Trash (as does deleting a draft), so
    // counting those would make a thread with one real draft warn about three. Trash is
    // resolved by EXACT role, from the mailbox list this batch already fetched. If that
    // role can't be resolved, every draft is counted as before — fail toward
    // over-warning, never toward missing a real draft. (A trashed draft IS still
    // returned under includeDrafts:true, which is an explicit ask for everything;
    // trashed non-drafts show there too.)
    const trashMailboxId = this.findByExactRole(threadMailboxes, 'trash')?.id;
    const isTrashedDraft = (e: any): boolean => {
      if (trashMailboxId == null) return false;
      const ids = Object.entries(e.mailboxIds || {}).filter(([, v]) => v).map(([id]) => id);
      return ids.length > 0 && ids.every(id => id === trashMailboxId);
    };
    const hiddenDraftCount = emails.filter((e: any) => e.keywords?.$draft && !isTrashedDraft(e)).length;
    return { emails: filtered, hiddenDraftCount };
  }

  async getMailboxStats(mailbox?: string): Promise<any> {
    // Fetch the full mailbox list once and read counts off it (getMailboxes returns all
    // fields, including the stat fields — it must NOT be narrowed). Resolving a specific
    // mailbox by id/role/name shares this list rather than issuing a second Mailbox/get.
    const mailboxes = await this.getMailboxes();
    const toStats = (mb: any) => ({
      id: mb.id,
      name: mb.name,
      role: mb.role,
      totalEmails: mb.totalEmails || 0,
      unreadEmails: mb.unreadEmails || 0,
      totalThreads: mb.totalThreads || 0,
      unreadThreads: mb.unreadThreads || 0,
    });

    if (mailbox !== undefined && String(mailbox).trim() !== '') {
      // Exact resolution (id/role/name); throws InvalidInputError on unknown. A real id
      // present but absent from the fetched list (a hidden/role-less mailbox) now throws
      // rather than returning stats — accepted residual (see docs/security-model.md).
      const mb = resolveMailbox(mailboxes, mailbox);
      return toStats(mb);
    }
    // Stats for all mailboxes.
    return mailboxes.map(toStats);
  }

  async getAccountSummary(): Promise<any> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const identities = await this.getIdentities();

    // Calculate totals
    const totals = mailboxes.reduce((acc, mb) => ({
      totalEmails: acc.totalEmails + (mb.totalEmails || 0),
      unreadEmails: acc.unreadEmails + (mb.unreadEmails || 0),
      totalThreads: acc.totalThreads + (mb.totalThreads || 0),
      unreadThreads: acc.unreadThreads + (mb.unreadThreads || 0)
    }), { totalEmails: 0, unreadEmails: 0, totalThreads: 0, unreadThreads: 0 });

    return {
      accountId: session.accountId,
      mailboxCount: mailboxes.length,
      identityCount: identities.length,
      ...totals,
      mailboxes: mailboxes.map(mb => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0
      }))
    };
  }

  async bulkMarkRead(emailIds: string[], read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = read
        ? { 'keywords/$seen': true }
        : { 'keywords/$seen': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkUpdate']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, `mark as ${read ? 'read' : 'unread'}`);
    }
  }

  async bulkPinEmails(emailIds: string[], pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = pinned
        ? { 'keywords/$flagged': true }
        : { 'keywords/$flagged': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkFlag']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, `${pinned ? 'pin' : 'unpin'}`);
    }
  }

  async bulkMove(emailIds: string[], target: string): Promise<void> {
    const session = await this.getSession();

    // Resolve the destination EXACTLY (id/role/name) — see moveEmail for the rationale
    // (new capability, exact-only, deliberate move-to-any stays open per fork #43).
    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // Fetch current mailboxIds for all emails to build proper JMAP patches
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: emailIds,
          properties: ['id', 'mailboxIds']
        }, 'getEmails']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const emails: any[] = this.getListResult(getResponse, 0);
    const mailboxMap: Record<string, Record<string, boolean>> = {};
    emails.forEach((e: any) => { mailboxMap[e.id] = e.mailboxIds || {}; });

    // Build patch per email: remove all current mailboxes, add target
    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      const patch: Record<string, boolean | null> = {};
      for (const mbId of Object.keys(mailboxMap[id] || {})) {
        patch[`mailboxIds/${mbId}`] = null;
      }
      patch[`mailboxIds/${targetMailboxId}`] = true;
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkMove']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, 'move');
    }
  }

  async bulkDelete(emailIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Find the trash mailbox by EXACT role only (case-insensitive) — see deleteEmail.
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findByExactRole(mailboxes, 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = { mailboxIds: trashMailboxIds };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkDelete']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      this.throwBulkSetError(result.notUpdated, emailIds.length, 'delete');
    }
  }
}