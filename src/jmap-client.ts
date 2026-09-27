import { FastmailAuth } from './auth.js';
import { validateFastmailUrl } from './url-validation.js';
import { parseAddress, requireNonEmpty, validateClearFields, coerceUtcDate, describeUntrusted, echoPath, PathAccessError, InvalidInputError } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { normalizeBodies, htmlHasVisibleContent, buildBodyParts, isBlank, assertBodyInputs } from './body-format.js';
import { rejectSignatureEmbeddedImage, signatureBlock, signatureCidRefs } from './reply-quote.js';
import { defaultIdentity, identityFor, signatureOf } from './identity.js';
import { expandBodyTokens, scanBodyTokens } from './body-tokens.js';
import type { BodyBlocks, BodyTokenScan } from './body-tokens.js';
import {
  bodyHash, classifyPartType, collectDraftBodyParts, draftInterleavedTextType, draftPartKey,
  isTextBodyType, resolveDraftBodyHash,
} from './body-hash.js';
import {
  buildUnionParts, cidKey, describePart, sanitizeDownloadFilename,
  checkInlineClosure, isRecreatableCid, isReservedCid, reconcileInlineParts,
  sanitizeQuoteHtml,
} from './inline-images.js';
import type { CidPart, UnionPart } from './inline-images.js';
import {
  InlineNoteLedger, describePartNames, emitInlineNotes,
  noteBodyHashUnreadable, noteDiscardedTextPart, noteEmbedMissingAfterSave, noteEmbedUnconfirmed,
  noteEscapedTokenShips, noteHistoryTokenStored, noteNearMissToken, noteSentWithEmbedded,
  noteSignatureExpandedInOnePart, noteSignatureTokenStored, noteTokenEmpty,
  noteUnexpandedSpelling,
  rejectBrokenDraft, rejectCidCollisionInCall, rejectCidCollisionOnDraft,
  rejectClearAttachmentsDanglingRefs, rejectDanglingCidRef, rejectExpandSignatureWithoutToken,
  rejectInterleavedTextParts, rejectMissingBodyHash,
  rejectRemovalDanglingRef, rejectRepeatedSignatureToken, rejectReservedCidRef,
  rejectStaleBodyHash, rejectUncarriableBodyPart, rejectUnrecreatableCid,
  NOTE_BODY_HASH_AFTER_EXPANSION, NOTE_BODY_HASH_DERIVED_PART, noteBodyHashAfterReRead,
} from './inline-notes.js';
import type { AttachmentAvailability } from './inline-notes.js';
import { matchSubjectPrefix, noteEditSubjectPrefix } from './subject-prefix.js';
import { buildIdCollapseNote } from './id-collapse-note.js';
// unlink is a security control, not a convenience: the exclusive-create download
// path removes the file it just refused to trust before rewriting it.
import { writeFile, mkdir, realpath, stat, lstat, open, unlink } from 'fs/promises';
import type { FileHandle } from 'fs/promises';
import { dirname, resolve, normalize, sep, basename, join } from 'path';
import { homedir } from 'os';

// A JMAP Email "attachment" body part referencing an uploaded blob. A carried
// (re-referenced) part may also pass through cid/disposition from the existing draft.
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
   * A file whose Content-ID is in this set is marked `inline`; anything else is a regular
   * attachment. Only the caller knows what html will ship; see uploadAttachments.
   */
  inlineCids?: ReadonlySet<string>;
}

// True if `child` is `parent` itself or nested beneath it. The read guard folds case on
// Win32 (NTFS folds case, so a case-sensitive compare is bypassable); the write guard
// stays byte-exact.
function isPathContained(child: string, parent: string, caseInsensitive: boolean): boolean {
  let c = child, p = parent;
  if (caseInsensitive) { c = c.toLowerCase(); p = p.toLowerCase(); }
  return c === p || c.startsWith(p + sep);
}

// Shared lexical pre-check: resolve `inputPath` against `allowedDir`, reject null bytes,
// and verify lexical containment. Only the cheap first gate: the canonical (realpath)
// re-verification is the caller's job.
function lexicalContainedPath(inputPath: string, allowedDir: string, caseInsensitive: boolean): string {
  const resolved = resolve(allowedDir, normalize(inputPath));
  if (resolved.includes('\0')) {
    throw new PathAccessError('path contains null bytes');
  }
  if (!isPathContained(resolved, allowedDir, caseInsensitive)) {
    // Nothing upstream rejects a line separator in a path or bounds one, so it goes through
    // `echoPath` inside `"…"` (docs/conventions.md, untrusted values in prose).
    throw new PathAccessError(`path must be within "${echoPath(allowedDir)}". Received: "${echoPath(inputPath)}"`);
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
  // Tilde followed by a digit specifically, so `report~final.pdf` is NOT rejected.
  if (/~\d/.test(input)) {
    throw new PathAccessError("path contains an 8.3 short-name segment (e.g. PROGRA~1), which is not allowed; use the full name.");
  }
}

// Extension -> MIME map for the Content-Type we POST when the caller omits one. A missing
// entry falls back to application/octet-stream, which is not harmless for every type: a
// receiving client decides from text/calendar and message/rfc822 whether to render an
// invitation or a forwarded message (draft_email's asAttachment .eml) inline.
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

// RFC 2045 token grammar for type/subtype. A positive grammar closes header injection via
// the Content-Type we POST. MIME parameters ("; charset=utf-8") are intentionally not
// accepted: only the type/subtype is needed to upload a blob.
const MIME_TYPE_PATTERN = /^[A-Za-z0-9!#$%&'*+.^_`|~-]+\/[A-Za-z0-9!#$%&'*+.^_`|~-]+$/;

function validateContentType(value: string, index: number): string {
  const v = value.trim();
  if (v.length > 255 || !MIME_TYPE_PATTERN.test(v)) {
    throw new PathAccessError(`attachments[${index}] has an invalid contentType "${describeUntrusted(value)}". Use a MIME type like application/pdf.`);
  }
  return v;
}

// The opt-in gate for attaching content that is ALREADY in the account. InvalidInputError,
// not PathAccessError: no filesystem path is involved on these sources.
function assertBlobAttachEnabled(allowBlobAttach: boolean, index: number, source: string): void {
  if (allowBlobAttach) return;
  throw new InvalidInputError(
    `attachments[${index}] attaches by ${source}, which is disabled. ` +
    'Set FASTMAIL_ALLOW_BLOB_ATTACH=true to allow attaching content already in the account, then restart the server to enable it.'
  );
}

/**
 * Substitute `{placeholder}` slots in a JMAP session URL template (RFC 8620 §1.6.2 —
 * downloadUrl, uploadUrl) with values inserted LITERALLY.
 *
 * A string replacement would interpret `$&`, `` $` ``, `$1` and friends inside a value and
 * splice template text into the URL (a blobId of `` a$`b `` injects the template's prefix),
 * so the replacement is a FUNCTION. A SINGLE pass keeps a value that reads `{accountId}` from
 * becoming a slot. An unfilled slot is left visible rather than emptied.
 *
 * Callers still re-run `validateFastmailUrl` on the result.
 */
export function fillUrlTemplate(template: string, values: Record<string, string>): string {
  return template.replace(/\{[^{}]*\}/g, (slot) => (
    Object.prototype.hasOwnProperty.call(values, slot) ? values[slot] : slot
  ));
}

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, any>;
  downloadUrl?: string;
  uploadUrl?: string;
  /**
   * Per-capability primary account ids. `accountId` above is the mail account; a non-mail
   * capability (contacts in particular) can be a DIFFERENT account, so read it from here.
   */
  primaryAccounts?: Record<string, string>;
}

/**
 * An attachment resolved to the fields needed to reference or fetch its blob. `type` and
 * `name` are defaulted at resolution time, and `name` is the SANITIZED filename, safe as a
 * save name; the raw declared name is deliberately not carried.
 */
export interface AttachmentInfo {
  blobId: string;
  type: string;
  name: string;
  size?: number;
  /**
   * WHICH form of the caller's `attachmentId` matched (see resolveAttachmentRef).
   * Non-optional so a consumer that cares cannot forget to ask.
   */
  matchedBy: AttachmentRefMatch;
}

/** How resolveAttachmentRef matched a reference; `'index'` is the positional fallback. */
export type AttachmentRefMatch = 'partId' | 'blobId' | 'cid' | 'index';

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
  // The 0-based index the server actually served (RFC 8620 section 5.5), falling back to
  // the one asked for. Describes the QUERY, not a message, so it is never serialized into
  // the JSON body on either path.
  position?: number;
  // Out-of-band metadata for the default Trash/Spam exclusion, for the handlers' trailing
  // note. NEVER serialized into the JSON body or the raw path. `hidden:null` = the count
  // could not be computed (emit the fail-closed note).
  exclusion?: {
    hidden: number | null;
    excludedRoles: string[];   // roles actually excluded (note fires iff hidden>0)
    unresolvedRoles: string[]; // roles intended-but-NOT-excluded (fail-loud note)
  };
}

// `unresolvedRoles` are roles meant to be excluded but not resolved: fail-loud, never
// silently included.
export interface ExclusionResult {
  excludeIds: string[];
  excludedRoles: string[];
  unresolvedRoles: string[];
}

// Shared Email/get property lists; keep in sync per CONTRIBUTING.md (JMAP properties).
// COMPACT: list/search tools and getThread. `textBody` fetches only the part STRUCTURE,
// not content, for the `bodyTextSize` hint (#59).
export const EMAIL_PROPERTIES_COMPACT = [
  'id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'receivedAt',
  'preview', 'keywords', 'threadId', 'messageId', 'references', 'inReplyTo',
  'hasAttachment', 'header:List-Unsubscribe:asURLs', 'blobId', 'size', 'mailboxIds',
  'textBody',
] as const;

// VERBOSE: superset with body properties, for verbose mode, getEmailById and getThread's
// includeBodies mode (#74). `sentAt` is for reply-quote attribution. The two provenance
// headers are deliberately VERBOSE-tier, not COMPACT: their values matter only when
// operating on one draft, so an ordinary thread read does not surface them.
export const EMAIL_PROPERTIES_VERBOSE = [
  ...EMAIL_PROPERTIES_COMPACT,
  'htmlBody', 'attachments', 'bodyValues', 'sentAt',
  'header:X-Forwarded-Message-Id:asMessageIds',
  'header:X-Fastmail-MCP-Source-Id:asText',
] as const;

// The provenance headers a draft carries about the message it was composed from:
// In-Reply-To (set by draft_email's reply mode and every other client's reply) and
// X-Forwarded-Message-Id (set by its forward mode and Fastmail's own clients). Both are
// arrays of BARE Message-IDs — JMAP MessageIds carry no angle brackets.
export interface SourceReferences {
  inReplyTo: string[];
  forwardedMessageId: string[];
  // The JMAP id of the exact stored instance the draft was composed from. A Message-ID
  // names a MESSAGE, and an account can hold several copies of one. Absent on other
  // clients' drafts; send_draft then falls back to the Message-ID lookup.
  sourceEmailId?: string;
}

// The JMAP header form used to SET and GET the recorded source instance. It is NOT
// stripped on send (EmailSubmission transmits the stored bytes verbatim); the decision is
// recorded in docs/security-model.md.
export const SOURCE_ID_HEADER = 'header:X-Fastmail-MCP-Source-Id:asText';

// Anything that is not an RFC 8620 id is treated as absent rather than risking a
// header-set rejection failing the create.
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

// What sendDraft returns, read off the same pre-send Email/get that validates the draft.
export interface SendDraftOutcome {
  submissionId: string;
  sourceReferences: SourceReferences;
  // What the transmitted message actually carried as embedded images (#13).
  notes?: string[];
}

// Recall cap for the Message-ID lookup. The oldest-first sort puts the owner on the first
// page (see findEmailIdsByMessageId), so this only bounds the extra mentions fetched.
const MESSAGE_ID_LOOKUP_LIMIT = 50;

// Do not trim this list per-call. `disposition`/`cid` drive draft_email's forward-mode
// true-inline drop (#30). `type` drives buildUnionParts (#13): a part with no type reads
// as body text, so dropping it empties the embedded-image half of the listing SILENTLY.
export const EMAIL_BODY_PROPERTIES = ['partId', 'blobId', 'type', 'size', 'name', 'disposition', 'cid'] as const;

// `attachments` is the full part listing (JMAP attachments plus media parts routed into
// the body lists), the basis every download index counts from. `rawAttachments` is the
// untouched JMAP array for `raw` mode (#13). `omittedFromRaw` is counted by MEMBERSHIP,
// not by subtracting lengths, which under-reports on a shared blobId or a null entry.
export interface EmailAttachmentsResult {
  attachments: any[];
  rawAttachments: any[];
  omittedFromRaw: number;
}

/**
 * Resolve a download_attachment `attachmentId` against a message's part listing.
 *
 * Four input forms, in this fixed order: partId, blobId, `cid:<value>`, then a plain
 * entry number. Fastmail's partIds ARE digit strings, so a part always wins over an index.
 *
 * Malformed INPUT throws; a well-formed reference that matches nothing returns undefined.
 *
 * WHICH form matched is returned because uploadAttachments refuses the positional form (a
 * shifted listing would bake the wrong file into a draft). Decide from what the resolver
 * DID, never from how the string looks: a partId of "2" is a real part.
 */
function resolveAttachmentRef(parts: any[], attachmentId: string): { part: any; matchedBy: AttachmentRefMatch } | undefined {
  const byPartId = parts.find((p: any) => p?.partId === attachmentId);
  if (byPartId) return { part: byPartId, matchedBy: 'partId' };

  // First match, deliberately NOT rejected when several parts share the blobId: blobs are
  // content-addressed, so the bytes are identical. Ambiguity is rejected only for cid,
  // where two parts sharing a Content-ID are different content.
  const byBlobId = parts.find((p: any) => p?.blobId === attachmentId);
  if (byBlobId) return { part: byBlobId, matchedBy: 'blobId' };

  if (/^cid:/i.test(attachmentId)) {
    // First occurrence only: `cid:cid:x` names the Content-ID "cid:x".
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
      `attachmentId "cid:${describeUntrusted(literal)}" matches more than one part. ` +
      'Pass the part\'s blobId or partId from get_email_attachments instead.'
    );
    const literalMatches = parts.filter((p: any) => p?.cid === literal);
    if (literalMatches.length > 1) throw ambiguous();
    if (literalMatches.length === 1) return { part: literalMatches[0], matchedBy: 'cid' };

    const decoded = cidKey(attachmentId);
    if (decoded !== literal) {
      const decodedMatches = parts.filter((p: any) => p?.cid === decoded);
      if (decodedMatches.length > 1) throw ambiguous();
      if (decodedMatches.length === 1) return { part: decodedMatches[0], matchedBy: 'cid' };
    }
    return undefined;
  }

  if (/^\d+$/.test(attachmentId)) {
    const part = parts[Number(attachmentId)];
    return part ? { part, matchedBy: 'index' } : undefined;
  }

  // Anything parseInt would swallow as a number ("3a", "-1") is an input error, never an
  // index into the list.
  if (!Number.isNaN(Number.parseInt(attachmentId, 10))) {
    throw new InvalidInputError(
      `attachmentId "${describeUntrusted(attachmentId)}" is not a usable attachment reference. ` +
      'Pass a partId or blobId from get_email_attachments, cid:<value> for an embedded ' +
      'image, or a plain entry number (0, 1, 2, ...) counting from the start of that listing.'
    );
  }

  return undefined;
}

// ---------------------------------------------------------------------------
// The draft-edit part model: body shape, removals, carry (#13)
// ---------------------------------------------------------------------------

// The part-identity rule (`draftPartKey`) and the part classifiers live in body-hash.ts, so
// the body-shape check and the body hash agree on what counts as one part.

// The non-text body parts the recreate can reproduce: the media three RFC 8621 §4.1.4
// routes into a body list, plus message/rfc822, which draft_email writes itself (an
// asAttachment forward). Anything else has no flat-property spelling.
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
 * The interleaved half is `draftInterleavedTextType` in body-hash.ts, shared with the read
 * that issues a `bodyHash`.
 *
 * Deduping FIRST is load-bearing: a single-format draft lists its one text part under both
 * textBody and htmlBody. A typeless part is left alone, since the body reader treats it as
 * body text.
 */
export function classifyDraftBodyShape(email: any): DraftBodyShape {
  const lists = [email?.textBody, email?.htmlBody].filter(Array.isArray) as any[][];
  const seen = new Set<string>();
  const shape: DraftBodyShape = {};
  let index = 0;

  const interleaved = draftInterleavedTextType(email);
  if (interleaved) shape.interleavedTextType = interleaved;

  for (const list of lists) {
    for (const part of list) {
      if (!part) continue;
      const key = draftPartKey(part, index++);
      if (seen.has(key)) continue;
      seen.add(key);

      const type = classifyPartType(part.type);
      if (!type || isTextBodyType(type)) continue;

      const carriable = isCarriableBodyType(type)
        && typeof part.blobId === 'string' && part.blobId !== '';
      if (!carriable && !shape.uncarriablePart) {
        shape.uncarriablePart = { part, isMedia: /^(?:image|audio|video)\//.test(type) };
      }
    }
  }

  return shape;
}

/**
 * The blank body part a send would ship, named by the body field a caller would fix.
 *
 * Reads the DEDUPLICATED part set whole rather than selecting one part per format, so part
 * order cannot hide a blank one. "Displays" is the same rule the read and the hash use.
 *
 * A part with no stored value is not blank (the draft has no body in that format); blank
 * means a present value that renders to nothing.
 *
 * htmlBody is answered first: an empty text/html part SHADOWS a real text/plain
 * alternative (RFC 2046, clients render the richest one).
 */
export function findBlankBodyPart(email: any): 'htmlBody' | 'textBody' | undefined {
  const blank = collectDraftBodyParts(email).filter(
    (p) => p.value !== undefined && p.value.trim() === '',
  );
  if (blank.some((p) => p.showsInHtml)) return 'htmlBody';
  if (blank.some((p) => p.showsInText)) return 'textBody';
  return undefined;
}

export interface AttachmentRemovalPlan {
  /** Parts this call takes off the draft, in stored order. */
  removed: any[];
  /** Parts that stay, in stored order. */
  survivors: any[];
  /**
   * The refusal a bad ref earned, held rather than thrown: resolution runs early, but the
   * body-shape guards and the hash check must still be the first refusal a caller hears.
   */
  error?: PathAccessError;
}

/**
 * Match each removeAttachments ref against the draft's parts, or clear them all.
 *
 * A ref names a blobId (every part with that blob comes off) or, failing that, a unique
 * non-null name. A ref that matches nothing, or a name that matches several, is an error
 * rather than a silent no-op.
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
    // The stored side is trimmed too, because coerceStringArray trims the ref and a MIME
    // filename may legally carry surrounding spaces.
    const nameMatches = survivors.filter((p) => p.name != null && String(p.name).trim() === ref);
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
          `removeAttachments ref "${describeUntrusted(ref)}" matches ${nameMatches.length} attachments by name; pass the blobId instead (one of: ${joinCapped(survivors.map((p) => p.blobId))}).`,
        ),
      };
    }
    return {
      removed,
      survivors,
      error: new PathAccessError(
        `removeAttachments ref "${describeUntrusted(ref)}" matched no attachment on this draft. Carried blobIds: ${joinCapped(storedParts.map((p) => p.blobId)) || '(none)'}.`,
      ),
    };
  }
  return { removed, survivors };
}

// Re-reference a stored part on the recreated draft. Whitelist exactly these fields: a
// blob-backed part is blobId XOR partId, and `size` is server-set. `demote` turns an
// embedded image whose displaying body is gone into an ordinary attachment; its Content-ID
// rides along, since a Content-ID never makes a part inline on its own.
function carriedPartFrom(part: any, demote = false): AttachmentPart {
  return {
    blobId: part.blobId,
    type: part.type,
    ...(part.name != null && { name: part.name }),
    ...(demote ? { disposition: 'attachment' } : part.disposition != null && { disposition: part.disposition }),
    ...(part.cid != null && { cid: part.cid }),
  };
}

// The embedded-image references an html body makes, as comparison keys: the same
// references the quote rewriter would act on.
function htmlCidRefs(html: string | null | undefined): string[] {
  if (!html) return [];
  return sanitizeQuoteHtml(html, { mode: 'collect' }).refs;
}

// ---- Wildcard identities are not addresses (#160) ----
//
// A WILDCARD identity's `email` is the pattern `*@domain`, not an address. `matchesIdentity`
// (src/identity.ts) honours it for any concrete address in the domain, and that is correct.
// What the pattern must never be is a `from` VALUE: it is a valid addr-spec, so nothing
// downstream rejects it, and a recipient would see a literal asterisk as the sender.
//
// It reaches that position three ways, each with its own refusal:
//
//   PASSED BY THE CALLER as `from` (rejectWildcardFromValue), checked before any identity
//     match, because matchesIdentity's equality test would read it as verified.
//   SELECTED as the identity's own `email` when no `from` was supplied
//     (rejectWildcardIdentityFrom).
//   ALREADY STORED on a draft: refused by the send path.
//
// WHAT A DRAFT ALREADY STORES IS RE-WRITTEN UNCHANGED on edit, pattern included: refusing it
// would block the very edit that fixes it. The send path is where a stored pattern is caught.
//
// `selectIdentity` in src/identity.ts is deliberately unchanged: a wildcard identity is still
// the correct selection, and supplies the signature.
function isWildcardIdentityEmail(email: unknown): boolean {
  return typeof email === 'string' && email.startsWith('*@');
}

/** The refusal both draft-WRITE paths raise, so they read as one rule. */
function rejectWildcardIdentityFrom(identityEmail: string): string {
  return `The default sending identity is the wildcard pattern "${describeUntrusted(identityEmail)}", not an address. ` +
    'Pass from with a concrete address in that domain.';
}

/**
 * The refusal both draft-WRITE paths raise for a `from` the CALLER passed whose address half
 * is the pattern itself. A separate sentence from `rejectWildcardIdentityFrom`, whose fix
 * ("pass from") would read as nonsense to a caller who did. Echoes the address HALF only, so
 * the display name does not blur which half to fix.
 */
function rejectWildcardFromValue(fromAddress: string): string {
  return `The from address "${describeUntrusted(fromAddress)}" is a wildcard identity's pattern, not an address. ` +
    'A from needs a concrete address in that domain; the wildcard identity still verifies it and still supplies its signature.';
}

const REJECT_UNVERIFIED_FROM =
  'From address is not verified for sending. Choose one of your verified identities.';

/**
 * createDraft's refusal of a caller `from` address half, in its order (the pattern before the
 * identity match), for a compose handler that has to raise it earlier. Undefined when valid.
 */
export function rejectFromAddress(identities: any[], fromAddress: string): string | undefined {
  if (isWildcardIdentityEmail(fromAddress)) return rejectWildcardFromValue(fromAddress);
  return identityFor(identities, fromAddress)
    ? undefined
    : REJECT_UNVERIFIED_FROM;
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

// A compact fingerprint of the draft an edit replaced, so a caller that edited from a
// stale copy sees what it overwrote (#65). Sizes, not bodies, and no Bcc: the old draft
// survives in Trash.
export interface ReplacedDraftInfo {
  id: string;
  subject?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  replyTo?: string[];
  textBodySize?: number;
  htmlBodySize?: number;
}

// updateDraft's result. Exactly ONE of trashedOldDraftId / orphanedOldDraftId is always
// set: the old copy's fate is never left unstated.
export interface UpdateDraftResult {
  id: string;
  replacedDraft: ReplacedDraftInfo;
  trashedOldDraftId?: string;
  orphanedOldDraftId?: string;
  // Why the Trash move didn't happen. Set whenever orphanedOldDraftId is.
  orphanedOldDraftReason?: string;
  // What the edit did to the draft's embedded images (#13); present only when it did something.
  inlineImages?: { embedded: number; degraded: number; removed: number };
  // The hash the caller's NEXT edit has to pass, from a RE-READ of the saved draft, never
  // from the bytes sent. Exactly one of these two is present on an edit that wrote or cleared
  // a body; neither on a metadata-only edit, where the caller's hash is still current.
  bodyHash?: string;
  bodyHashWithheld?: string;
  notes?: string[];
}

export interface MailboxInfo {
  name: string;
  // The stable JMAP role (lowercased), or null for a custom folder/label.
  role: string | null;
}

// Build an id -> {name, role} lookup from a Mailbox/get list (#10, #49). A mailbox with no
// string `name` is left out, so its id surfaces later as unresolved.
//
// Deliberately NO root-anchored path: that would widen the cheap ['id','name','role']
// projection on every list/search page. Paths belong to the mailbox surface.
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
// properties (`_mailboxNames`, `_mailboxRoles`, `_unresolvedMailboxIds`), so the raw:true
// paths' JSON.stringify omits them while simplifyEmail can still read them.
//
// NEVER-SILENT, NON-THROWING resolution (#53). A resolved id contributes its name (and
// role, if any); an unresolved id goes to `_unresolvedMailboxIds`. A promised location is
// never silently dropped, and a name is never fabricated.
//
// WHY not throw, and WHY not silently omit (do not "fix" this back to either):
//   - An unresolved id is rare and benign: a just-created folder, a race with the
//     separately-fetched mailbox list, or a mailbox with no `name`. Role mailboxes always
//     resolve. Throwing would fail a whole list/search page over one such id.
//   - Silently omitting the id WAS the #53 bug: a promised field vanished with no trace.
// A genuine Mailbox/get `error` response still throws via the callers' catches; that is
// not this path.
//
// Each property is attached only when non-empty. `roles` and `mailboxes` are INDEPENDENT
// sets, not parallel arrays: a custom folder contributes a name but no role.
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

// ---------- archive ----------

/** What archiving decided to do to one message, before the write is attempted. */
export type ArchiveBranch = 'movedToArchive' | 'removedFromInbox' | 'notInInbox' | 'refused';

/** What archiving actually did to one message. The two extra buckets are write outcomes. */
export type ArchiveAction = ArchiveBranch | 'notFound' | 'failed';

/**
 * What remove_labels / bulk_remove_labels report back.
 *
 * `rescued` names the messages left with no mailbox and filed in Archive to keep them from
 * being destroyed: a change of filing the caller did not ask for, so it must be reported.
 * `unchangedCount` is the messages none of the named labels was on, which the write skips,
 * so a call that changed nothing does not read like one that changed everything.
 */
export interface LabelRemovalResult {
  rescued: string[];
  unchangedCount: number;
}

/** Internal shape of the shared removal write; `distinctCount` is post-de-duplication. */
interface LabelRemovalOutcome extends LabelRemovalResult {
  notUpdated: Record<string, any>;
  distinctCount: number;
  /**
   * Ids submitted in the write's own `update` map, acknowledged in `updated`, and absent
   * from the final `notUpdated`: the success count for `throwBulkSetError`. Never
   * `total - failCount`, which would count a skipped no-op as a success.
   */
  updatedCount: number;
}

export interface ArchiveEmailResult {
  id: string;
  action: ArchiveAction;
  /**
   * Where the message is filed. PROJECTED from the pre-write read for movedToArchive and
   * removedFromInbox; OBSERVED for notInInbox and refused. For failed, the filing BEFORE the
   * attempt, or neither field when the filing could not be read. Absent for notFound.
   *
   * INDEPENDENT sets, not parallel arrays: read "did it reach Archive" off
   * roles.includes('archive'), never off the branch name (Inbox+Archive takes
   * removedFromInbox yet ends in Archive).
   */
  mailboxes?: string[];
  roles?: string[];
  /** Ids that could not be resolved to a name. Never omitted silently (#53). */
  unresolvedMailboxIds?: string[];
  reason?: {
    role?: string;
    setErrorType?: string;
    description?: string;
    /**
     * The write was dispatched and the server acknowledged this id in NEITHER its `updated`
     * nor its `notUpdated` map, so nothing confirmed the outcome either way.
     *
     * A STRUCTURAL marker, not inferred from `description`: the renderer hedges its headline
     * on it, and keying off wording would break silently on a rewording or be spoofed by a
     * server description containing the phrase.
     */
    outcomeUnknown?: boolean;
  };
}

export interface ArchiveResult {
  results: ArchiveEmailResult[];
  /** One entry per action, always all six. Sums to the number of DISTINCT ids passed in. */
  counts: Record<ArchiveAction, number>;
}

/**
 * The roles where Fastmail's client offers no Archive action at all, in the fixed order
 * used to pick one when a message sits in two of them.
 *
 * Exactly the set measured to omit Archive IN EACH ROLE'S OWN VIEW; applying it from any
 * view is a deliberate, small extrapolation (docs/fastmail-action-availability.md). Extend
 * it only by measuring a view, never from a role's name. Deliberately not a general
 * "system mailbox" predicate.
 */
export const ARCHIVE_REFUSING_ROLES = ['trash', 'junk', 'drafts', 'scheduled', 'sent', 'snoozed'] as const;

/**
 * How many caller-supplied email ids any one error message names before it summarises the
 * rest. ONE cap for every such list (#134).
 */
const EMAIL_ID_LIST_CAP = 10;

/**
 * Whether a value returned inside a JMAP method response can be read as an id-keyed map.
 *
 * An array or string would answer hasOwnProperty for index-shaped ids ("0", "1"), and
 * fabricate an answer about the write.
 */
function isPlainResponseMap(value: any): boolean {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

/**
 * An `Email/set` `update` map keyed by CALLER-SUPPLIED email ids.
 *
 * Object.create(null), not `{}`: on an ordinary object `updates['__proto__'] = patch` hits
 * the prototype SETTER, the entry never reaches the request, and the success text reports a
 * change nothing made. `__proto__` is a legal JMAP id. Read it back with
 * Object.prototype.hasOwnProperty.call, or `constructor` reads as a response.
 */
function newUpdateMap(): Record<string, any> {
  return Object.create(null);
}

/**
 * Decide what archiving does to one message, from its current membership alone.
 *
 * The order of the tests is the measured rule: Inbox membership FIRST, because the client
 * offers Archive on an Inbox+Trash message.
 *
 * Nothing is added when the surviving memberships are all refusing roles (Inbox+Trash keeps
 * Trash and gains nothing), as the client does. A "rescue" into Archive belongs to #43.
 */
export function decideArchiveBranch(
  currentIds: string[],
  inboxId: string,
  roleById: Map<string, string | null>,
): { branch: ArchiveBranch; keptIds: string[]; refusingRole?: string } {
  if (currentIds.includes(inboxId)) {
    const keptIds = currentIds.filter(id => id !== inboxId);
    return keptIds.length > 0
      ? { branch: 'removedFromInbox', keptIds }
      : { branch: 'movedToArchive', keptIds: [] };
  }

  const presentRoles = new Set(currentIds.map(id => roleById.get(id)).filter((r): r is string => !!r));
  const refusingRole = ARCHIVE_REFUSING_ROLES.find(role => presentRoles.has(role));
  return refusingRole
    ? { branch: 'refused', keptIds: [], refusingRole }
    : { branch: 'notInInbox', keptIds: [] };
}

/**
 * Resolve a set of mailbox ids to names/roles for an archive report, following
 * attachMailboxInfo's never-silent rule (resolved -> name and role, unresolved -> the raw
 * id in unresolvedMailboxIds) but as ORDINARY ENUMERABLE fields.
 *
 * Not attachMailboxInfo's non-enumerable carrier: this result IS serialized, so those
 * fields would silently vanish (#53).
 */
function describeMailboxIds(
  ids: string[],
  map: Map<string, MailboxInfo>,
): { mailboxes: string[]; roles?: string[]; unresolvedMailboxIds?: string[] } {
  const mailboxes: string[] = [];
  const roles: string[] = [];
  const unresolvedMailboxIds: string[] = [];
  for (const id of ids) {
    const info = map.get(id);
    if (info) {
      mailboxes.push(info.name);
      if (info.role) roles.push(info.role);
    } else {
      unresolvedMailboxIds.push(id);
    }
  }
  return {
    mailboxes,
    ...(roles.length > 0 ? { roles } : {}),
    ...(unresolvedMailboxIds.length > 0 ? { unresolvedMailboxIds } : {}),
  };
}

// Cap the mailbox names listed in a not-found error so a large account doesn't
// produce a huge message; the list_mailboxes pointer keeps a truncated list actionable.
const MAILBOX_LIST_CAP = 30;

// The bound on how far a parent chain is followed, so a corrupt chain terminates. Hitting it
// is REPORTED as unwalkable, never truncated into "not found". A legitimate chain deeper
// than this would be reported too, which is the safe direction to be wrong in.
const MAILBOX_PATH_SEPARATOR = '/';
const MAILBOX_PARENT_CHAIN_CAP = 100;

// The ONE normalisation a path segment gets, applied on BOTH sides of every comparison, or
// a path the server emits fails to resolve when pasted back. '' means "no path", never an
// empty segment.
function normalizeMailboxSegment(name: unknown): string {
  return typeof name === 'string' ? name.trim() : '';
}

// Root-anchored path for every mailbox in the list, keyed by id ("Archive/2026/Receipts").
//
// A mailbox whose chain never reaches the root within the cap, or passes a blank name, gets
// NO path and is reported in `unpathable`: a partial string would look root-anchored and
// could match a different mailbox's real path.
//
// Shared by list_mailboxes' `path` column and the resolver, so output and input describe
// the same tree. Deliberately NOT memoised: it is one pass over tens of mailboxes, and a
// cache would depend on no caller mutating its list.
export function buildMailboxPathMap(mailboxes: any[]): { paths: Map<string, string>; unpathable: string[] } {
  const list = (mailboxes || []).filter(mb => mb && typeof mb.id === 'string');
  const byId = new Map<string, any>(list.map(mb => [mb.id as string, mb]));
  const paths = new Map<string, string>();
  const unpathable: string[] = [];
  for (const mb of list) {
    const segments: string[] = [];
    let cursor: any = mb;
    let rooted = false;
    for (let depth = 0; depth < MAILBOX_PARENT_CHAIN_CAP; depth++) {
      const segment = normalizeMailboxSegment(cursor.name);
      if (segment === '') break;
      segments.unshift(segment);
      if (!cursor.parentId) { rooted = true; break; }
      const parent = byId.get(cursor.parentId);
      if (!parent) break;
      cursor = parent;
    }
    if (rooted) paths.set(mb.id, segments.join(MAILBOX_PATH_SEPARATOR));
    else unpathable.push(mb.id);
  }
  return { paths, unpathable };
}

// The handle a not-found/ambiguity message offers for one mailbox: its path, else its id.
// Never its bare name, which is what the caller just had rejected as ambiguous.
function mailboxLabel(mb: any, paths: Map<string, string>): string {
  return paths.get(mb?.id) ?? String(mb?.id);
}

// The shared "Valid: …" tail listing known mailboxes as PATHS (the vocabulary the ambiguity
// errors use), capped. Each VALUE is sanitised here, never the finished sentence the call
// sites append (#131). A path cut by the cap stops being pasteable; accepted, since the
// tail points at list_mailboxes.
function mailboxListHint(mailboxes: any[]): string {
  const { paths } = buildMailboxPathMap(mailboxes || []);
  const entries = (mailboxes || [])
    .filter(mb => mb && typeof mb.name === 'string')
    .map(mb => {
      const label = describeUntrusted(mailboxLabel(mb, paths));
      return mb.role ? `${label} (${describeUntrusted(mb.role)})` : label;
    });
  const shown = entries.slice(0, MAILBOX_LIST_CAP);
  let list = shown.join(', ');
  if (entries.length > shown.length) {
    list += `, …and ${entries.length - shown.length} more — call list_mailboxes for the full list`;
  }
  return `Use an id, a role (inbox/archive/sent/drafts/trash/junk), a name, or a full path (Parent/Child). Valid: ${list}`;
}

function formatMailboxNotFound(input: string, mailboxes: any[]): string {
  return `Mailbox "${describeUntrusted(input)}" not found. ${mailboxListHint(mailboxes)}`;
}

// A name shared by several mailboxes is NOT a typo, so it never renders as "not found".
// `candidates` are RAW paths, so each is sanitised here (contrast formatMailboxNameVsPath).
function formatMailboxAmbiguous(input: string, candidates: string[]): string {
  return `Mailbox "${describeUntrusted(input)}" is ambiguous: ${candidates.length} mailboxes share that name. ` +
    `Retry with one of their full paths, or with an id. Candidates: ${joinCapped(candidates.map(describeUntrusted))}`;
}

// The other ambiguity: one reference is BOTH a folder name containing the separator and a
// path to a different mailbox. Only an id separates them, so "retry with a full path" would
// be wrong advice. `candidates` are already-composed descriptions whose values were
// sanitised where built; re-describing them would truncate the whole sentence.
function formatMailboxNameVsPath(input: string, candidates: string[]): string {
  return `Mailbox "${describeUntrusted(input)}" is ambiguous: it is both the name of one folder and the path to a different ` +
    `mailbox, and a path cannot tell those apart. Retry with the id of the one you mean. ` +
    `Candidates: ${joinCapped(candidates)}`;
}

// Both sides of a name/path collision render as the SAME path string, so each is described
// for what it is and carries the id, the one form that picks either.
function describeFlatCandidate(mb: any): string {
  return `folder named "${describeUntrusted(normalizeMailboxSegment(mb?.name))}" (id: ${describeUntrusted(mb?.id)})`;
}

function describeNestedCandidate(mb: any, paths: Map<string, string>): string {
  const path = paths.get(mb?.id);
  // Sanitised per SEGMENT, so a long path is not cut off mid-way.
  const nesting = path
    ? path.split(MAILBOX_PATH_SEPARATOR).map(describeUntrusted).join(' > ')
    : describeUntrusted(mb?.id);
  return `nested folder ${nesting} (id: ${describeUntrusted(mb?.id)})`;
}

// Distinct from "not found" and "ambiguous": the tree itself is unwalkable.
function formatMailboxUnwalkable(input: string, id: string): string {
  return `Mailbox "${describeUntrusted(input)}" could not be resolved as a path: mailbox "${describeUntrusted(id)}" has a parent chain that never reaches ` +
    `a top-level mailbox (a loop, or a parent missing from the mailbox list), so full paths cannot be computed. ` +
    `Refer to the mailbox by id, role, or name instead.`;
}

// The multi-input messages for the label arrays name EVERY failed value in one message, in
// separate buckets because each calls for a different correction.

// Join a capped list and SAY when it was capped: a truncated list with no tail reads as a
// complete one.
export function joinCapped(items: string[], separator = ', '): string {
  const shown = items.slice(0, MAILBOX_LIST_CAP);
  const listed = shown.join(separator);
  return items.length > shown.length ? `${listed}${separator}…and ${items.length - shown.length} more` : listed;
}

function formatMailboxesNotFound(unresolved: string[], mailboxes: any[]): string {
  const listed = joinCapped(unresolved.map(v => `"${describeUntrusted(v)}"`));
  return `Mailbox(es) not found: ${listed}. ${mailboxListHint(mailboxes)}`;
}

export interface MailboxResolutionFailures {
  notFound: string[];
  ambiguous: Array<{ input: string; candidates: string[] }>;
  // A folder name AND a different mailbox's path. Separate from `ambiguous` because only an
  // id fixes it.
  nameVsPath: Array<{ input: string; candidates: string[] }>;
  unwalkable: Array<{ input: string; id: string }>;
}

function formatMailboxesNotResolved(failures: MailboxResolutionFailures, mailboxes: any[]): string {
  const parts: string[] = [];
  if (failures.notFound.length > 0) {
    const listed = joinCapped(failures.notFound.map(v => `"${describeUntrusted(v)}"`));
    parts.push(`Mailbox(es) not found: ${listed}.`);
  }
  if (failures.ambiguous.length > 0) {
    // Raw paths, so each is described; nameVsPath below carries composed descriptions.
    const listed = joinCapped(
      failures.ambiguous.map(a => `"${describeUntrusted(a.input)}" matches ${joinCapped(a.candidates.map(describeUntrusted))}`),
      '; ',
    );
    parts.push(`Ambiguous mailbox name(s) — retry with a full path or an id: ${listed}.`);
  }
  if (failures.nameVsPath.length > 0) {
    const listed = joinCapped(
      failures.nameVsPath.map(a => `"${describeUntrusted(a.input)}" matches ${joinCapped(a.candidates)}`),
      '; ',
    );
    parts.push(
      `Mailbox reference(s) that name a folder AND describe a path to a different mailbox — ` +
      `retry with an id, which is the only form that tells them apart: ${listed}.`,
    );
  }
  if (failures.unwalkable.length > 0) {
    const listed = joinCapped(
      failures.unwalkable.map(u => `"${describeUntrusted(u.input)}" (blocked by mailbox "${describeUntrusted(u.id)}")`),
    );
    parts.push(
      `Mailbox path(s) unresolvable because a parent chain never reaches a top-level mailbox: ${listed}. ` +
      'Refer to those by id, role, or name.',
    );
  }
  return `${parts.join(' ')} ${mailboxListHint(mailboxes)}`;
}

/**
 * The mailboxes named for a label write that Fastmail does not treat as labels (#133),
 * rendered name-and-role for the refusal below.
 *
 * MEASURED from the client's pickers: "Labels" offers only the Inbox and user labels, so a
 * role mailbox is a FOLDER, with the Inbox (in both pickers) the sole exception. The test is
 * the ROLE, never a name list, so it stays correct as Fastmail adds roles.
 *
 * Deliberately NOT ARCHIVE_REFUSING_ROLES, which answers a different measured question.
 */
function findNonLabelMailboxes(resolvedIds: string[], mailboxes: any[]): string[] {
  const byId = new Map<string, any>();
  for (const mailbox of mailboxes || []) {
    if (mailbox && typeof mailbox.id === 'string') byId.set(mailbox.id, mailbox);
  }
  const named: string[] = [];
  const seen = new Set<string>();
  for (const id of resolvedIds) {
    // One mailbox named twice (id and name) is one offender.
    if (seen.has(id)) continue;
    seen.add(id);
    const mailbox = byId.get(id);
    const role = mailbox && typeof mailbox.role === 'string' ? mailbox.role.trim().toLowerCase() : '';
    if (!role || role === 'inbox') continue;
    const name = mailbox && typeof mailbox.name === 'string' && mailbox.name.trim() !== ''
      ? mailbox.name
      : id;
    named.push(`"${describeUntrusted(name)}" (${describeUntrusted(role)})`);
  }
  return named;
}

/**
 * The label tools' namespace check, one message for add and remove. Runs AFTER mailbox
 * resolution (a raw-argument check would miss every form but a role) and BEFORE any
 * Email/set. All-or-nothing: serving the rest would report a subset as success.
 */
function assertLabelNamespace(resolvedIds: string[], mailboxes: any[]): void {
  const offenders = findNonLabelMailboxes(resolvedIds, mailboxes);
  if (offenders.length === 0) return;
  const one = offenders.length === 1;
  throw new InvalidInputError(
    `${joinCapped(offenders)} ${one ? 'is a folder' : 'are folders'} in Fastmail's model, not ` +
    `${one ? 'a label' : 'labels'}, so ${one ? 'it' : 'they'} cannot be added to or removed from a ` +
    'message as a label. Fastmail\'s label picker offers only the Inbox and your own labels; every ' +
    'other role mailbox (Archive, Trash, Spam, Drafts, Sent, Snoozed, Scheduled) is offered under ' +
    '"Move to" instead. Nothing was changed. Use move_email (or bulk_move) to file a message ' +
    `${one ? 'there' : 'in one of them'} instead. The Inbox is the one exception, because it is both: ` +
    'removing the inbox label (which is what archiving a message is) and adding it back are served here.'
  );
}

/**
 * The outcome of resolving ONE mailbox reference.
 *
 * RETURNED rather than thrown: `resolveMailbox` throws on the spot, while the label arrays
 * aggregate every failure into one error.
 *
 *   { mailbox }              resolved.
 *   { ambiguous, candidates} a flat name matched several mailboxes; `candidates` are paths.
 *   { ambiguous, candidates, nameVsPath } one folder's name AND a different mailbox's path;
 *                            only an id separates them, so `candidates` are descriptions.
 *   { unwalkable, id }       a parent chain never reached the root. Not "not found": a
 *                            corrupt tree is not a caller typo.
 *   undefined                a genuine miss.
 */
export type MailboxMatch =
  | { mailbox: any }
  | { ambiguous: string; candidates: string[]; nameVsPath?: true }
  | { unwalkable: true; id: string };

// Exact-only mailbox match, in resolution order:
//   1. exact id       case-SENSITIVE: a JMAP id is an opaque token.
//   2. role           case-insensitive.
//   3. exact flat name case-insensitive.
//   4. `/`-joined path root-anchored, every segment case-insensitive.
//
// A flat name matching exactly one mailbox WINS over the path reading, so a name containing
// "/" stays reachable, UNLESS the path resolves to a DIFFERENT mailbox: then it is ambiguous,
// since a wrong write destination is worth a retry to avoid.
//
// NO substring matching (an injection-steering primitive on write paths). A custom mailbox
// named after a role is shadowed by the role branch, which is accepted.
//
// Input is normalised HERE, and the path walk lives here, so the single-mailbox params and
// the label arrays accept one vocabulary.
export function findMailboxExact(mailboxes: any[], input: string): MailboxMatch | undefined {
  const list = mailboxes || [];
  const raw = String(input).trim();
  const byId = list.find(mb => mb && mb.id === raw);
  if (byId) return { mailbox: byId };
  const lower = raw.toLowerCase();
  const byRole = list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === lower);
  if (byRole) return { mailbox: byRole };
  const byName = list.filter(mb => mb && normalizeMailboxSegment(mb.name).toLowerCase() === lower);
  if (byName.length > 1) {
    const { paths } = buildMailboxPathMap(list);
    return { ambiguous: raw, candidates: byName.map(mb => mailboxLabel(mb, paths)) };
  }

  // A leading, trailing or doubled separator leaves an empty segment: not a path, rather than
  // a guess at which slash was meant. A unique flat name can still win below.
  const rawSegments = raw.includes(MAILBOX_PATH_SEPARATOR)
    ? raw.split(MAILBOX_PATH_SEPARATOR).map(normalizeMailboxSegment)
    : undefined;
  const segments = rawSegments && !rawSegments.some(s => s === '') ? rawSegments : undefined;

  // BEFORE the flat-name tie-break, which depends on what the path resolves to.
  let paths = new Map<string, string>();
  let unpathable: string[] = [];
  let matches: any[] = [];
  if (segments) {
    const target = segments.join(MAILBOX_PATH_SEPARATOR).toLowerCase();
    ({ paths, unpathable } = buildMailboxPathMap(list));
    matches = list.filter(mb => {
      const p = mb && paths.get(mb.id);
      return typeof p === 'string' && p.toLowerCase() === target;
    });
  }

  if (byName.length === 1) {
    const flat = byName[0];
    const collidingWith = matches.filter(mb => mb.id !== flat.id);
    if (collidingWith.length === 0) return { mailbox: flat };
    return {
      ambiguous: raw,
      nameVsPath: true,
      candidates: [describeFlatCandidate(flat), ...collidingWith.map(mb => describeNestedCandidate(mb, paths))],
    };
  }

  if (!segments) return undefined;
  if (matches.length === 1) return { mailbox: matches[0] };
  if (matches.length > 1) return { ambiguous: raw, candidates: matches.map(mb => mailboxLabel(mb, paths)) };

  // The path matched nothing. A broken chain elsewhere is blamed only when it plausibly IS
  // the cause (no mailbox has a path, or a named segment is unpathable); otherwise a typo
  // would be told paths cannot be computed.
  if (unpathable.length > 0) {
    if (paths.size === 0) return { unwalkable: true, id: unpathable[0] };
    const named = new Set(segments.map(s => s.toLowerCase()));
    const blamed = unpathable.find(id => {
      const mb = list.find(m => m && m.id === id);
      return !!mb && named.has(normalizeMailboxSegment(mb.name).toLowerCase());
    });
    if (blamed) return { unwalkable: true, id: blamed };
  }
  return undefined;
}

// Throwing wrapper over findMailboxExact for the single-mailbox callers. Each failure shape
// gets its OWN message.
export function resolveMailbox(mailboxes: any[], input: string): any {
  const match = findMailboxExact(mailboxes, input);
  if (match && 'mailbox' in match) return match.mailbox;
  const raw = String(input).trim();
  if (match && 'ambiguous' in match) {
    throw new InvalidInputError(
      match.nameVsPath
        ? formatMailboxNameVsPath(raw, match.candidates)
        : formatMailboxAmbiguous(raw, match.candidates),
    );
  }
  if (match && 'unwalkable' in match) {
    throw new InvalidInputError(formatMailboxUnwalkable(raw, match.id));
  }
  throw new InvalidInputError(formatMailboxNotFound(raw, mailboxes || []));
}

// Narrow a mailbox list to the DIRECT children of one parent. A blank parent means no
// filter. A pure operation rather than a getMailboxes option, because list_mailboxes needs
// the WHOLE tree for paths before narrowing.
export function filterMailboxesByParent(mailboxes: any[], parent?: string): any[] {
  const list = mailboxes || [];
  if (parent === undefined || parent === null || String(parent).trim() === '') return list;
  const parentId = resolveMailbox(list, parent).id;
  return list.filter(mb => mb && mb.parentId === parentId);
}

// A mailbox name is a LEAF name: a "/" would be permanently ambiguous against the path form,
// and is far more likely a caller expressing nesting inline.
export function assertLeafMailboxName(name: string): void {
  if (name.includes(MAILBOX_PATH_SEPARATOR)) {
    throw new InvalidInputError(
      `Mailbox name must not contain "${MAILBOX_PATH_SEPARATOR}": "${describeUntrusted(name)}". ` +
      'Pass the leaf name and nest it with the parent parameter (e.g. name: "2026", parent: "Archive").',
    );
  }
}

// Compute the default Trash/Spam exclusion. Resolves trash/junk by EXACT role only, NEVER a
// name substring, which could hide real mail in a custom "Junk mail rules". An explicit scope
// turns it off. A role that cannot be resolved goes to unresolvedRoles for the fail-loud note,
// never silently included.
//
// ONE of the two places that decide whether the default exclusion runs; the other is each
// runFilteredQuery caller's `exclusionIntended` (docs/conventions.md). Teach both about any
// new explicit scope.
//
// A role the caller already excluded (`callerExcludedIds`) is dropped from BOTH arrays, or the
// note would prescribe includeTrash/includeSpam, which cannot reveal those messages.
export function computeExclusion(
  mailboxes: any[],
  opts: {
    includeTrash?: boolean;
    includeSpam?: boolean;
    hasExplicitScope?: boolean;
    callerExcludedIds?: string[];
  },
): ExclusionResult {
  const excludeIds: string[] = [];
  const excludedRoles: string[] = [];
  const unresolvedRoles: string[] = [];
  if (opts.hasExplicitScope) {
    return { excludeIds, excludedRoles, unresolvedRoles };
  }
  const list = mailboxes || [];
  const alreadyExcluded = new Set(opts.callerExcludedIds || []);
  const findRole = (role: string) => list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === role);
  const add = (role: string, label: string) => {
    const mb = findRole(role);
    if (!mb) { unresolvedRoles.push(label); return; }
    if (alreadyExcluded.has(mb.id)) return;
    excludeIds.push(mb.id);
    excludedRoles.push(label);
  };
  if (!opts.includeTrash) add('trash', 'Trash');
  if (!opts.includeSpam) add('junk', 'Spam');
  return { excludeIds, excludedRoles, unresolvedRoles };
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
      // A method-level `error` (RFC 8620 §3.6.1), a different class from the per-id SetError
      // describeSetError() formats, so intentionally not routed through it. Its fields are
      // still server-authored and rendered untrusted (#134).
      throw new Error(`JMAP error: ${describeUntrusted(result.type)}${result.description ? ' - ' + describeUntrusted(result.description) : ''}`);
    }
    return result;
  }

  /**
   * Format a JMAP SetError (RFC 8620 §5.3). The single chokepoint every throwing
   * notCreated/notUpdated site routes through. We add no content of our own (a server's
   * description may carry a snippet; that is its text); failing ids are added by the bulk
   * callers.
   *
   * Each field is rendered untrusted SEPARATELY, so the " - " is still our punctuation
   * (#134). Sanitising here covers the returned edit_draft orphan reason too, which never
   * meets the CallTool catch. The cap can cost a long description its tail; accepted, since
   * `type` is what a caller recovers from and is capped separately.
   */
  protected describeSetError(entry: { type: string; description?: string }): string {
    const type = describeUntrusted(entry.type);
    return `${type}${entry.description ? ' - ' + describeUntrusted(entry.description) : ''}`;
  }

  /**
   * The JMAP set-error types a caller can resolve by re-forming the call, and therefore
   * the ones that surface as `InvalidParams` rather than `InternalError`.
   *
   * The dividing question is "what should the caller do next": a type belongs here only when
   * re-forming the request is the route to success. An unknown type stays out. Drawn from RFC
   * 8620 §5.3, plus `invalidArguments` (§3.6.2), which real servers return per record.
   *
   * Deliberately OUT: `overQuota` (the fix is freeing capacity, not a new call), `rateLimit`
   * (a pause then retry), `stateMismatch` (re-read then retry), `forbidden`, the server
   * failures, and `willDestroy` (this client never batches update and destroy for one id).
   */
  private static readonly CALLER_FIXABLE_SET_ERROR_TYPES: ReadonlySet<string> = new Set([
    'notFound',
    'invalidProperties',
    'invalidArguments',
    'invalidPatch',
    'tooLarge',
    'singleton',
  ]);

  private static isCallerFixableSetError(type: string): boolean {
    return JmapClient.CALLER_FIXABLE_SET_ERROR_TYPES.has(type);
  }

  /**
   * Throw the correctly-classified error for a single-id Email/set failure (#22, #41):
   * InvalidInputError for a CALLER_FIXABLE_SET_ERROR_TYPES type, plain Error otherwise.
   * `action` is the verb phrase, e.g. "move email".
   */
  protected throwSingleSetError(entry: { type: string; description?: string }, action: string): never {
    const message = `Failed to ${action}: ${this.describeSetError(entry)}`;
    if (JmapClient.isCallerFixableSetError(entry.type)) {
      throw new InvalidInputError(message);
    }
    throw new Error(message);
  }

  /**
   * Throw the correctly-classified error for a bulk Email/set partial failure: counts plus
   * the failing ids grouped by reason (#22), so an agent can retry exactly the failures. An
   * id in `notUpdated` may be one nobody submitted, so ids go through `nameEmailIds`.
   * Truncation says the list is PARTIAL and points at re-running the full input (these
   * mutators are idempotent). InvalidInputError only when EVERY failure is caller-fixable
   * (#41).
   *
   * `successCount` is the CALLER's (see `countAcknowledged`), never `total - failCount`,
   * which counts an unsubmitted id as done. `effectiveTotal` floors the denominator, since a
   * server can name an id `total` never counted. The "no reported outcome" clause is 0 for
   * every caller here (all close `notUpdated` via `withUnaccountedFailures`) and stays for a
   * subclass that does not: an unaccounted id must not vanish from the sentence.
   */
  protected throwBulkSetError(
    notUpdated: Record<string, { type: string; description?: string }>,
    total: number,
    successCount: number,
    action: string,
    // A side effect the SUCCEEDING messages incurred, which a partial failure must not hide.
    trailingNote?: string,
  ): never {
    const MAX_REASONS = 5;

    const failedIds = Object.keys(notUpdated);
    const failCount = failedIds.length;
    const effectiveTotal = Math.max(total, failCount + successCount);
    const unaccountedCount = effectiveTotal - failCount - successCount;

    // Keyed on the RAW type+description, not describeSetError's truncated text, which would
    // merge two different errors into one false shared cause. NUL keeps the key unambiguous.
    type SetErrorEntry = { type: string; description?: string };
    const byReason = new Map<string, { entry: SetErrorEntry; ids: string[] }>();
    for (const id of failedIds) {
      const entry = notUpdated[id];
      const key = `${entry.type}\u0000${entry.description ?? ''}`;
      const group = byReason.get(key);
      if (group) group.ids.push(id);
      else byReason.set(key, { entry, ids: [id] });
    }

    let truncated = false;
    const reasonEntries = [...byReason.values()];
    if (reasonEntries.length > MAX_REASONS) truncated = true;
    const groups = reasonEntries.slice(0, MAX_REASONS).map(({ entry, ids }) => {
      if (ids.length > EMAIL_ID_LIST_CAP) truncated = true;
      return `${this.describeSetError(entry)}: ${JmapClient.nameEmailIds(ids)}`;
    });

    const successPhrase = unaccountedCount > 0
      ? `${successCount} succeeded, ${unaccountedCount} with no reported outcome`
      : `${successCount} succeeded`;
    let message = `Failed to ${action} ${failCount} of ${effectiveTotal} emails (${successPhrase}). ${groups.join('; ')}.`;
    if (truncated) {
      message += ' (Partial list — not every failure is shown. These operations are idempotent, so re-run with the full input set to retry every failure safely.)';
    }
    if (trailingNote) message += ` ${trailingNote}`;

    if (failedIds.every(id => JmapClient.isCallerFixableSetError(notUpdated[id].type))) {
      throw new InvalidInputError(message);
    }
    throw new Error(message);
  }

  /**
   * Count of `submittedIds` (a bulk write's own `Object.keys(update)`) acknowledged in
   * `updated` and NOT also in `notUpdated`: the success count for `throwBulkSetError`. A
   * non-compliant server can list one id in both maps, which would otherwise double-count it.
   */
  private static countAcknowledged(
    submittedIds: string[],
    updated: any,
    notUpdated: Record<string, any>,
  ): number {
    const acknowledged = isPlainResponseMap(updated) ? updated : {};
    return submittedIds.filter(id =>
      Object.prototype.hasOwnProperty.call(acknowledged, id) &&
      !Object.prototype.hasOwnProperty.call(notUpdated, id)
    ).length;
  }

  /**
   * A fresh Email/set `notUpdated` map: the server's own map, plus a synthesized
   * `outcomeUnknown` entry for every id in `submittedIds` the server acknowledged in
   * NEITHER `updated` nor its own `notUpdated`: an unconfirmed id is a failure.
   *
   * A FRESH newUpdateMap(), never `{}` or a spread (see newUpdateMap). Membership is
   * `hasOwnProperty`, never truthiness: RFC 8620 §5.3 lets a success come back as null.
   *
   * Must run BEFORE countAcknowledged, so an id this adds is excluded from the success
   * count it computes.
   */
  private static withUnaccountedFailures(
    submittedIds: string[],
    updated: any,
    serverNotUpdated: any,
  ): Record<string, any> {
    const notUpdated: Record<string, any> = newUpdateMap();
    if (isPlainResponseMap(serverNotUpdated)) Object.assign(notUpdated, serverNotUpdated);
    const acknowledged = isPlainResponseMap(updated) ? updated : {};
    for (const id of submittedIds) {
      if (Object.prototype.hasOwnProperty.call(notUpdated, id)) continue;
      if (Object.prototype.hasOwnProperty.call(acknowledged, id)) continue;
      notUpdated[id] = { type: 'outcomeUnknown' };
    }
    return notUpdated;
  }

  /**
   * Extract the .list array from a JMAP method response, with null safety.
   */
  protected getListResult(response: JmapResponse, index: number): any[] {
    const result = this.getMethodResult(response, index);
    return result?.list || [];
  }

  /**
   * Like getListResult, but returns [] instead of throwing, for an auxiliary method appended
   * to a batch whose primary work must not be turned into a failure. Covers both an absent
   * entry and an `error` ENTRY (RFC 8620 section 3.6.1), the likelier real failure.
   *
   * It cannot say what went wrong, so every caller must state the degradation itself (e.g.
   * `unresolvedMailboxIds`). A caller that swallows the [] silently is the bug.
   */
  protected readListResultIfPresent(response: JmapResponse, index: number): any[] {
    if (!response.methodResponses || index >= response.methodResponses.length) return [];
    const entry = response.methodResponses[index];
    if (!Array.isArray(entry) || entry.length < 2 || entry[0] === 'error') return [];
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
    // Kept only when numeric; the paging callers otherwise fall back to the requested one.
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

    // Validate every URL the server hands us before we send the bearer token to it. The
    // templates' {placeholders} are stripped for parsing.
    const allowUnsafe = this.auth.getAllowUnsafe();
    const stripTemplate = (url: string) => url.replace(/\{[^}]+\}/g, 'x');
    if (typeof sessionData.apiUrl !== 'string') {
      throw new Error('Invalid session response: apiUrl missing');
    }
    validateFastmailUrl(sessionData.apiUrl, 'session.apiUrl', allowUnsafe);
    // Reject a present-but-non-string URL: validate and store must not diverge.
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
      // As in getSession: a redirect would replay the token to an unvalidated host.
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

  // Find a mailbox by EXACT role (case-insensitive). A USABLE id is part of the match:
  // every caller reads `.id` straight away, and a missing one becomes the literal
  // "undefined" in a silently corrupt write.
  /**
   * edit_draft refuses a draft filed in Trash. A superseded copy keeps `$draft` there, and the
   * recreate carries the old copy's mailboxIds, so an edit would write the replacement into
   * Trash too. A draft filed anywhere else (draft_email's `mailbox`) stays editable. A map
   * this cannot read, or an account with no trash-role mailbox, is not refused here.
   */
  private refuseTrashedDraftEdit(filing: any, mailboxes: any[]): void {
    const trash = this.findByExactRole(mailboxes, 'trash');
    if (!trash || !isPlainResponseMap(filing)) return;
    if (!(Object.prototype.hasOwnProperty.call(filing, trash.id) && filing[trash.id] === true)) return;
    const filedIn = Object.keys(filing)
      .filter(id => filing[id] === true)
      .map(id => {
        const mailbox = mailboxes.find(mb => mb?.id === id);
        const name = describeUntrusted(mailbox?.name);
        if (mailbox && name.trim() !== '') return `"${name}"`;
        return `${mailbox ? 'unnamed' : 'unknown'} mailbox (id: "${describeUntrusted(id)}")`;
      });
    throw new InvalidInputError(
      `This draft is in Trash, so it will not be edited (it is in: ${joinCapped(filedIn)}). ` +
      'Move it back to Drafts with move_email and edit it again.',
    );
  }

  private findByExactRole(mailboxes: any[], role: string): any | undefined {
    const target = role.toLowerCase();
    return (mailboxes || []).find(mb =>
      mb && typeof mb.role === 'string' && mb.role.toLowerCase() === target
      && typeof mb.id === 'string' && !!mb.id
    );
  }

  // A per-id set-error out of a JMAP `notUpdated` map, or undefined when the server did not
  // list that id. Three things a bare `notUpdated[id]` gets wrong:
  //
  // 1. hasOwnProperty: the id is CALLER-supplied, and "constructor" would index a prototype
  //    function and fabricate a failure.
  // 2. KEY PRESENCE is the refusal, not truthiness: a null value is still a refusal, hence
  //    the `?? {}`.
  // 3. isPlainResponseMap: an array-shaped map would answer for the id "0".
  private setErrorFor(notUpdated: any, id: string): any | undefined {
    if (!isPlainResponseMap(notUpdated)) return undefined;
    return Object.prototype.hasOwnProperty.call(notUpdated, id) ? (notUpdated[id] ?? {}) : undefined;
  }

  // Resolve an optional mailbox input to an id; blank means no filter. Pass `mailboxes` to
  // avoid a second fetch.
  private async resolveMailboxId(input?: string, mailboxes?: any[]): Promise<string | undefined> {
    if (input === undefined || input === null || String(input).trim() === '') return undefined;
    const list = mailboxes ?? await this.getMailboxes();
    return resolveMailbox(list, input).id;
  }

  // Fetch the account's mailboxes, whole and unprojected. Narrowing to one parent is
  // `filterMailboxesByParent`.
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

  // Create a mailbox. Returns:
  //   - `created`: the server's created object, UNTOUCHED, for the raw path. RFC 8620 §5.3
  //     means it often holds just the id.
  //   - `mailbox`: that object merged over the create arguments.
  //   - `path`: OMITTED when the parent has no computable path, rather than a plausible
  //     string that does not resolve.
  async createMailbox(input: { name: string; parent?: string }): Promise<{ mailbox: any; created: any; path?: string }> {
    const name = requireNonEmpty(input?.name, 'name', 'pass the leaf name of the mailbox to create');
    assertLeafMailboxName(name);

    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const { paths } = buildMailboxPathMap(mailboxes);

    let parentId: string | null = null;
    // undefined = "the parent has no computable path", unlike the top-level case.
    let parentPath: string | undefined;
    const parentInput = input.parent;
    if (parentInput !== undefined && parentInput !== null && String(parentInput).trim() !== '') {
      const parent = resolveMailbox(mailboxes, parentInput);
      parentId = parent.id;
      parentPath = paths.get(parent.id);
    }

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/set', {
          accountId: session.accountId,
          create: {
            newMailbox: { name, parentId },
          },
        }, 'createMailbox']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notCreated && result.notCreated.newMailbox) {
      this.throwSingleSetError(result.notCreated.newMailbox, 'create mailbox');
    }

    const created = result.created?.newMailbox;
    if (!created?.id) {
      throw new Error('Mailbox creation reported success but the server returned no mailbox id');
    }
    const mailbox = { name, parentId, ...created };
    const leaf = normalizeMailboxSegment(mailbox.name) || name;
    const path = parentId === null
      ? leaf
      : (parentPath === undefined ? undefined : `${parentPath}${MAILBOX_PATH_SEPARATOR}${leaf}`);
    return { mailbox, created, path };
  }

  // With no explicit `mailbox`, applies the same default Trash/Spam exclusion and hidden
  // count as searchEmails.
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

    // searchEmails has its OWN copy of this line; change both when what counts as an
    // explicit scope changes (see computeExclusion).
    const hasExplicitScope = !!resolvedMailboxId;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: opts.includeTrash,
      includeSpam: opts.includeSpam,
      hasExplicitScope,
    });
    const exclusionIntended = !hasExplicitScope && (!opts.includeTrash || !opts.includeSpam);

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

    // No maxBodyValueBytes: Fastmail does not truncate by default, and REJECTS an explicit
    // maxBodyValueBytes:0 with invalidArguments.
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
      throw new InvalidInputError(`Email with ID "${describeUntrusted(id)}" not found`);
    }

    const email = result.list?.[0];
    if (!email) {
      throw new InvalidInputError(`Email with ID "${describeUntrusted(id)}" not found or not accessible`);
    }

    attachMailboxInfo([email], buildMailboxInfoMap(this.readListResultIfPresent(response, 1)));
    return email;
  }

  async getIdentities(): Promise<any[]> {
    const session = await this.getSession();

    // No `properties` filter: omitted means every property (RFC 8620 section 5.1), which
    // is what raw/verbose output relies on (#33). Naming them would narrow it.
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
    
    return defaultIdentity(identities);
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

    // At least one meaningful field; blank bodies count as absent, attachments as content.
    if (!email.to?.length && !email.subject && isBlank(email.textBody) && isBlank(email.htmlBody) && !email.attachments?.length) {
      throw new InvalidInputError('At least one of to, subject, textBody, htmlBody, or attachments must be provided');
    }

    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    // `from` takes "Name <address>" (#161) and is split BEFORE anything looks at it:
    // matchesIdentity takes a bare addr-spec only (do not widen it). The name half is never
    // validated, because nothing on the platform reads it.
    const parsedFrom = email.from ? parseAddress(email.from) : undefined;

    // Before the identity match; see isWildcardIdentityEmail.
    if (parsedFrom && isWildcardIdentityEmail(parsedFrom.email)) {
      throw new InvalidInputError(rejectWildcardFromValue(parsedFrom.email));
    }

    let selectedIdentity;
    if (email.from) {
      selectedIdentity = identityFor(identities, parsedFrom!.email);
      if (!selectedIdentity) {
        throw new InvalidInputError(REJECT_UNVERIFIED_FROM);
      }
    } else {
      selectedIdentity = defaultIdentity(identities);
    }

    // With no `from`, the identity's OWN `email` is written, which for a wildcard is the
    // pattern (#160).
    if (!email.from && isWildcardIdentityEmail(selectedIdentity.email)) {
      throw new InvalidInputError(rejectWildcardIdentityFrom(selectedIdentity.email));
    }

    const fromEmail = parsedFrom?.email || selectedIdentity.email;
    // The caller's own display name if given, else the verifying identity's.
    const fromName = parsedFrom?.name ?? selectedIdentity.name;

    const mailboxes = await this.getMailboxes();
    let draftMailboxId: string;
    if (email.mailbox) {
      draftMailboxId = resolveMailbox(mailboxes, email.mailbox).id;
    } else {
      // EXACT role, the same question sendDraft's gate asks, or this would produce a draft
      // the send can never accept. Never a name substring.
      const draftsMailbox = this.findByExactRole(mailboxes, 'drafts');
      if (!draftsMailbox) {
        throw new Error(
          'Could not find a Drafts mailbox (no mailbox in this account carries the "drafts" role). ' +
          'Pass the mailbox parameter to save this draft somewhere explicitly — but note that ' +
          'send_draft only sends a draft that is in the Drafts folder.',
        );
      }
      draftMailboxId = draftsMailbox.id;
    }

    const mailboxIds: Record<string, boolean> = {};
    mailboxIds[draftMailboxId] = true;

    const emailObject: any = {
      mailboxIds,
      keywords: { $draft: true },
      from: [{ name: fromName, email: fromEmail }],
    };

    if (email.to?.length) emailObject.to = email.to.map(parseAddress);
    if (email.cc?.length) emailObject.cc = email.cc.map(parseAddress);
    if (email.bcc?.length) emailObject.bcc = email.bcc.map(parseAddress);
    if (email.subject) emailObject.subject = email.subject;
    if (email.inReplyTo?.length) emailObject.inReplyTo = email.inReplyTo;
    if (email.references?.length) emailObject.references = email.references;
    if (email.replyTo?.length) emailObject.replyTo = email.replyTo.map(parseAddress);
    // A header SET, round-tripped by Fastmail. Pre-vetted by the compose handler, and
    // Fastmail rejects CRLF/non-ASCII.
    if (email.forwardedMessageId?.length) emailObject['header:X-Forwarded-Message-Id:asMessageIds'] = email.forwardedMessageId;
    // Vetted at this single seam, so a malformed value degrades to absent rather than
    // failing the create.
    if (isSettableSourceId(email.sourceEmailId)) emailObject[SOURCE_ID_HEADER] = email.sourceEmailId;
    if (email.attachments?.length) emailObject.attachments = email.attachments;
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

    if (result.notCreated?.draft) {
      this.throwSingleSetError(result.notCreated.draft, 'create draft');
    }

    const emailId = result.created?.draft?.id;
    if (!emailId) {
      throw new Error('Draft creation returned no email ID');
    }

    return emailId;
  }

  // Extract a stored body value by MIME type. A single-format draft has its ONE part
  // aliased into BOTH textBody and htmlBody (docs/email-bodies.md), so select by the part's
  // actual type, not its list.
  //
  // THE MATCH IS EXACT ON PURPOSE, unlike the read's rule for a typeless part
  // (draftTextBodyType): a typeless part would answer both formats, and the recreate would
  // write the text value into the html slot too, silently, even on a metadata-only edit
  // (#179).
  private bodyValueForType(parts: any[] | undefined, mimeType: string, bodyValues: Record<string, any>): string | undefined {
    const part = parts?.find((p: any) => p.type === mimeType && p.partId != null && bodyValues[p.partId]);
    return part ? bodyValues[part.partId].value : undefined;
  }

  // The JMAP body-part shaping for createDraft. Rejects ONLY html that renders to nothing
  // and has no image; a draft with neither body gets empty shaping.
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
    // Expand `{{signature}}` in the bodies this call writes. Absent means false; see the
    // token step below for why the trigger is a flag, not the token's presence.
    expandSignature?: boolean;
    // The `bodyHash` a `get_email` of this draft returned. Required by every edit that
    // writes or clears a body.
    bodyHash?: string;
  }, options: {
    // Whether this server can attach anything at all. Changes only which repair a refusal
    // offers.
    attachmentsEnabled?: boolean;
  } = {}): Promise<UpdateDraftResult> {
    // Guarded here, unlike the other compose paths, because `updates` IS edit_draft's raw
    // input; createDraft is shared with paths that pass an already-merged body.
    assertBodyInputs(updates);

    const session = await this.getSession();

    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'textBody', 'htmlBody', 'bodyValues', 'mailboxIds', 'keywords', 'inReplyTo', 'references', 'attachments', 'header:X-Forwarded-Message-Id:asMessageIds', SOURCE_ID_HEADER],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const existingEmail = this.getListResult(getResponse, 0)[0];
    if (!existingEmail) {
      throw new InvalidInputError(`Email with ID "${describeUntrusted(emailId)}" not found`);
    }

    if (!existingEmail.keywords?.$draft) {
      throw new InvalidInputError('Cannot edit a non-draft email');
    }
    this.refuseTrashedDraftEdit(existingEmail.mailboxIds, await this.getMailboxes());

    const availability: AttachmentAvailability = {
      attachmentsEnabled: options.attachmentsEnabled !== false,
    };

    // The UNION of `attachments` and the media parts routed into the body lists (RFC 8621
    // §4.1.4): reading `attachments` alone would silently drop a body-routed image.
    const storedParts: any[] = buildUnionParts(existingEmail).map((u: UnionPart) => u.part);

    // The recreate rebuilds from flat props. Fastmail assembles the multipart/related
    // embedding from a flat create (live-probed; scripts/probes/foreign-draft-roundtrip.mjs),
    // so embedded images survive it. What flat props cannot spell (a part it cannot
    // re-reference, an interleaved body) is refused loudly rather than mangled (#13, #85).
    const bodyShape = classifyDraftBodyShape(existingEmail);
    if (bodyShape.uncarriablePart) {
      const { part, isMedia } = bodyShape.uncarriablePart;
      throw new InvalidInputError(rejectUncarriableBodyPart(part?.name, part?.type, isMedia));
    }
    if (bodyShape.interleavedTextType) {
      throw new InvalidInputError(rejectInterleavedTextParts());
    }

    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    // Split as createDraft does (#161). `updates.from` stays raw for the non-empty check.
    const parsedUpdateFrom = updates.from ? parseAddress(updates.from) : undefined;

    if (parsedUpdateFrom && isWildcardIdentityEmail(parsedUpdateFrom.email)) {
      throw new InvalidInputError(rejectWildcardFromValue(parsedUpdateFrom.email));
    }

    let selectedIdentity;
    if (updates.from) {
      selectedIdentity = identityFor(identities, parsedUpdateFrom!.email);
      if (!selectedIdentity) {
        throw new InvalidInputError(REJECT_UNVERIFIED_FROM);
      }
    } else {
      const existingFrom = existingEmail.from?.[0]?.email;
      if (existingFrom) {
        selectedIdentity = identityFor(identities, existingFrom) ?? defaultIdentity(identities);
      } else {
        selectedIdentity = defaultIdentity(identities);
      }
    }

    const bodyValues = existingEmail.bodyValues || {};
    const existingTextValue = this.bodyValueForType(existingEmail.textBody, 'text/plain', bodyValues);
    const existingHtmlValue = this.bodyValueForType(existingEmail.htmlBody, 'text/html', bodyValues);

    // A provided-but-empty value is a loud error (almost always an accidental clobber);
    // blanking is done through clearFields. `from` is not clearable: a draft always has a
    // sender. `forwardedMessageId` is clearable but not settable.
    const CLEARABLE = new Set(['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments', 'forwardedMessageId']); // NOT 'from'
    const SETTABLE = ['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'from'] as const;
    const provided = new Set<string>(SETTABLE.filter(f => (updates as any)[f] !== undefined));
    // So validateClearFields refuses clear-then-append of attachments in one call.
    if (updates.attachments?.length || updates.removeAttachments?.length) provided.add('attachments');
    validateClearFields(updates.clearFields, CLEARABLE, provided);
    const clear = new Set(updates.clearFields ?? []);

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
    // The body checks above are GUARDS ONLY: their trimmed return is discarded so stored
    // bodies keep their exact value.

    // ---- What this edit does to each body ----
    // Order below: the coupling guards, the hash, the token refusals, then the merge.
    const wroteHtml = updates.htmlBody !== undefined;
    const wroteText = updates.textBody !== undefined;
    const clearedHtml = clear.has('htmlBody');
    const clearedText = clear.has('textBody');
    const wroteAnyBody = wroteText || wroteHtml;
    const clearedAnyBody = clearedText || clearedHtml;
    const touchesBody = wroteAnyBody || clearedAnyBody;

    // The caller's OWN html, before the `{{signature}}` expansion. Checks that must not see
    // this server's text (an expanded signature may carry a reserved identifier) read this,
    // never workingHtml.
    const callerWrittenHtml = updates.htmlBody;

    // Replaced only by the expansion below. Never write back into `updates`: an in-process
    // caller can reuse that object across calls.
    let workingHtml = updates.htmlBody;
    let workingText = updates.textBody;

    // ---- The draft's provenance, carried or cleared ----
    // In-Reply-To survives every edit. The forward marking is cleared only through
    // clearFields, on a body edit and a metadata-only edit alike.
    const isReply = !!existingEmail.inReplyTo?.length;
    const carriedForwardHeader: string[] = existingEmail['header:X-Forwarded-Message-Id:asMessageIds'] || [];
    const dropForwardHeader = clear.has('forwardedMessageId');
    // An asAttachment forward is a forward for the marking but not for the noun below. Read
    // off the part UNION: which list holds the .eml is a MIME-shape accident.
    const emlAttached = storedParts.some((p: any) => classifyPartType(p?.type) === 'message/rfc822');
    // The source instance rides with the forward marking: dropped when a FORWARD draft is
    // de-forwarded, kept on a reply draft. Vetted as the create vets it (isSettableSourceId).
    const storedSourceId = existingEmail[SOURCE_ID_HEADER];
    const trimmedSourceId = typeof storedSourceId === 'string' ? storedSourceId.trim() : undefined;
    const carriedSourceId: string | undefined =
      isSettableSourceId(trimmedSourceId) ? trimmedSourceId : undefined;

    // How the notes name the carried block, from the HEADERS ALONE, never the body's markup.
    const keepNoun = !isReply && carriedForwardHeader.length > 0 && !emlAttached
      ? 'the forwarded block'
      : 'the quote';

    // Holds its refusal (see AttachmentRemovalPlan.error).
    const removalPlan = resolveAttachmentRemovals(
      storedParts,
      updates.removeAttachments,
      clear.has('attachments'),
    );

    // References the STORED body already makes with nothing to supply them; see the check
    // further down.
    const storedPartCids = new Set(storedParts.map(partCid).filter((c) => c !== ''));
    const storedDanglingRefs = htmlCidRefs(existingHtmlValue).filter((r) => !storedPartCids.has(r));

    // ---- The sending identity ----
    // The address this edit writes into `from` (the emailObject build below must use the
    // same value). The sign-off resolves against THAT, not `selectedIdentity`, which may
    // have fallen back to the account default while a stored address is written.
    //
    // Only the THIRD arm writes the identity's own `email`, so only it is refused for a
    // wildcard (#160); a stored `from` is re-written unchanged (see isWildcardIdentityEmail).
    if (!updates.from && !existingEmail.from?.[0]?.email && isWildcardIdentityEmail(selectedIdentity.email)) {
      throw new InvalidInputError(rejectWildcardIdentityFrom(selectedIdentity.email));
    }
    const writtenFromAddress: string | undefined =
      parsedUpdateFrom?.email || existingEmail.from?.[0]?.email || selectedIdentity.email;
    const signingIdentity = writtenFromAddress
      ? identityFor(identities, writtenFromAddress)
      : undefined;
    // The display name written alongside that address: the caller's own in THIS edit's
    // `from` (#161), else the name the stored draft carries against that address, else the
    // signing identity's (#152). Only passed fields change, so a stored name must not be
    // reverted by an edit that never touched `from`. `selectedIdentity.name` is deliberately
    // not a fallback: it would put the default's name before a foreign address.
    // Case-folded rather than matchesIdentity, whose wildcard branch must not apply here.
    const storedFromEmailLower = existingEmail.from?.[0]?.email?.toLowerCase();
    const storedFromName = storedFromEmailLower !== undefined && storedFromEmailLower === writtenFromAddress?.toLowerCase()
      ? existingEmail.from?.[0]?.name
      : undefined;
    // A blank stored name counts as no name, so it cannot beat a real identity name.
    const storedFromNameIfPresent: string | undefined =
      storedFromName && storedFromName.trim() !== '' ? storedFromName : undefined;
    // parseAddress gives no `name` for a blank name half, so it falls through to the stored
    // arm. The documented limit (#161): a stored display name can be REPLACED, never REMOVED.
    const writtenFromName: string | null =
      parsedUpdateFrom?.name ?? storedFromNameIfPresent ?? signingIdentity?.name ?? null;
    const editSignature = signatureOf(signingIdentity);

    // ---- The merge rule, written once and applied twice ----
    // The guards judge the pre-expansion body and the real merge the post-expansion one;
    // one helper keeps the two from drifting. A written body drops the unwritten partner; a
    // no-body edit preserves both; clearFields force the body absent.
    const mergeHtml = (html: string | undefined): string | undefined =>
      clearedHtml ? undefined : (html !== undefined ? html : (wroteAnyBody ? undefined : existingHtmlValue));
    const mergeText = (text: string | undefined): string | undefined =>
      clearedText ? undefined : (text !== undefined ? text : (wroteAnyBody ? undefined : existingTextValue));

    // ---- Body-shape coupling guards ----
    // The text part is a DERIVED fallback when html is present (docs/email-bodies.md).
    //
    // These run AHEAD OF THE HASH CHECK: they name the shape to fix, where a hash failure
    // would send the caller to re-read for a request that is refused anyway.
    if (wroteText && !wroteHtml && !clearedHtml && !isBlank(existingHtmlValue)) {
      throw new InvalidInputError('editing textBody alone won\'t change what most recipients see (they render htmlBody). To change the message, edit htmlBody (the text fallback regenerates automatically); to save a custom plain-text alternative, supply htmlBody alongside it; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }
    // PRE-expansion html on purpose; an html that expands to nothing is refused downstream.
    if (clearedText && !clearedHtml && !isBlank(mergeHtml(callerWrittenHtml))) {
      throw new InvalidInputError('textBody can\'t be cleared on its own while htmlBody is present — the text fallback is managed automatically (regenerated from htmlBody, or html-only if none can be derived). Omit textBody from clearFields; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }

    // ---- Proof that the caller read the body it is replacing ----
    // `bodyHash` is a LOST-UPDATE GUARD and nothing else: it proves the caller saw the body,
    // never that it kept any of it. Recomputed over the same part set `get_email` hashes. The
    // window to the write is accepted; there is no `ifInState` below.
    //
    // Gated on `touchesBody`, not on body-invariance, which the metadata path cannot promise
    // (see the hand-back below).
    if (touchesBody) {
      if (typeof updates.bodyHash !== 'string' || updates.bodyHash.trim() === '') {
        throw new InvalidInputError(rejectMissingBodyHash());
      }
      if (updates.bodyHash.trim() !== bodyHash(collectDraftBodyParts(existingEmail))) {
        throw new InvalidInputError(rejectStaleBodyHash());
      }
    }

    // ---- Body tokens ----
    // The ONE thing this tool does to a body it is handed, and it is opt-in.
    //
    // WHY THE TRIGGER IS A FLAG AND NOT THE TOKEN'S PRESENCE: part of a handed-back body was
    // authored by the original's sender, who could plant any in-band trigger. A flag cannot
    // be planted, so a stored `{{signature}}` is stable under every unflagged edit.
    const expandSignature = updates.expandSignature === true;
    const writtenParts: { part: 'textBody' | 'htmlBody'; authored: string; stored?: string }[] = [];
    if (wroteText) writtenParts.push({ part: 'textBody', authored: updates.textBody!, ...(existingTextValue !== undefined && { stored: existingTextValue }) });
    if (wroteHtml) writtenParts.push({ part: 'htmlBody', authored: updates.htmlBody!, ...(existingHtmlValue !== undefined && { stored: existingHtmlValue }) });
    const scans = new Map<'textBody' | 'htmlBody', BodyTokenScan>(
      writtenParts.map((p) => [p.part, scanBodyTokens(p.authored)] as const),
    );
    const tokenNotes: string[] = [];

    if (expandSignature) {
      // The only text-keyed refusal on this tool: the flag claims the written part as the
      // caller's own. After the hash check, so a stale caller is told to re-read first.
      const placed = writtenParts.reduce((n, p) => n + scans.get(p.part)!.counts.signature, 0);
      if (placed === 0) throw new InvalidInputError(rejectExpandSignatureWithoutToken(wroteAnyBody));
      for (const p of writtenParts) {
        const n = scans.get(p.part)!.counts.signature;
        if (n > 1) throw new InvalidInputError(rejectRepeatedSignatureToken(p.part, n));
      }

      // The sign-off's text form depends on whether the MESSAGE ships an html part.
      const messageShipsHtml = !isBlank(mergeHtml(callerWrittenHtml));
      // Which parts a sign-off ACTUALLY landed in (an identity with no signature lands
      // nothing), read after the loop because each part's note is about the other.
      const signatureLanded = new Map<'textBody' | 'htmlBody', boolean>();
      for (const p of writtenParts) {
        const blocks: BodyBlocks = {
          signature: signingIdentity
            ? signatureBlock(editSignature, p.part, messageShipsHtml)
            : { available: false, cause: 'no-identity' },
          // Neither history token expands or is REMOVED here: it is stored text.
          quote: { available: 'as-written' },
          forward: { available: 'as-written' },
        };
        // THE single pass, over the caller's OWN string. Runs on every written part, because
        // a `\{{signature}}` escape in a flagged call must still be resolved.
        const expansion = expandBodyTokens(p.authored, blocks);
        if (p.part === 'htmlBody') workingHtml = expansion.text;
        else workingText = expansion.text;
        signatureLanded.set(
          p.part,
          expansion.tokens.some((site) => site.name === 'signature' && site.expanded),
        );
        for (const site of expansion.tokens) {
          if (site.name === 'signature' && !site.expanded && site.cause) {
            tokenNotes.push(noteTokenEmpty('signature', p.part, site.cause));
          }
        }
      }

      // The uneven-sign-off note is owed only when a sign-off ACTUALLY landed.
      for (const p of writtenParts) {
        if (scans.get(p.part)!.counts.signature !== 0 || writtenParts.length <= 1) continue;
        const other = writtenParts.find((q) => q.part !== p.part)!;
        if (!signatureLanded.get(other.part)) continue;
        tokenNotes.push(noteSignatureExpandedInOnePart(other.part, p.part));
      }
    }

    // NOTES, never refusals: the body may be a foreign one handed back, so a refusal keyed on
    // its text could be planted by the original's author and recur on every edit.
    //
    // EVERY TEXT-KEYED NOTE BELOW IS GATED ON A RISE AGAINST THE STORED SCAN (an exact count
    // over the stored bytes), so only what the caller just added is reported. A note added
    // here without the gate is the defect.
    //
    // Every DISTINCT unknown `{{…}}` this call added, as one sentence. DISTINCT is the unit
    // for list and count alike (`describePartNames`' invariant).
    const addedSpellings: string[] = [];

    // A note whose signature takes a part name is per part (counter inside the loop); the
    // escape, near-miss and history notes take only the text, so they dedupe ACROSS the edit.
    // The count-rise budgets stay per part, because the stored bytes are per part.

    const reportedEscapes = new Set<string>();
    const namedMisses = new Set<string>();
    const notedHistory = new Set<string>();

    for (const p of writtenParts) {
      const scan = scans.get(p.part)!;
      const storedScan = scanBodyTokens(p.stored ?? '');

      // Flagged and unflagged alike: an unknown spelling ships either way.
      const storedSpellings = new Map<string, number>();
      for (const s of storedScan.otherSpellings) {
        storedSpellings.set(s.text, (storedSpellings.get(s.text) ?? 0) + 1);
      }
      for (const s of scan.otherSpellings) {
        const budget = storedSpellings.get(s.text) ?? 0;
        if (budget > 0) { storedSpellings.set(s.text, budget - 1); continue; }
        if (!addedSpellings.includes(s.text)) addedSpellings.push(s.text);
      }

      if (!expandSignature) {
        const added = scan.counts.signature - storedScan.counts.signature;
        if (added > 0) tokenNotes.push(noteSignatureTokenStored(p.part, added));
        const storedEscapes = new Map<string, number>();
        for (const e of storedScan.escapes) storedEscapes.set(e.text, (storedEscapes.get(e.text) ?? 0) + 1);
        for (const e of scan.escapes) {
          const budget = storedEscapes.get(e.text) ?? 0;
          if (budget > 0) { storedEscapes.set(e.text, budget - 1); continue; }
          if (reportedEscapes.has(e.text)) continue;
          reportedEscapes.add(e.text);
          tokenNotes.push(noteEscapedTokenShips(e.text));
        }
      }
      // Count rise per token, the signature note's gate: a `{{quote}}` sitting in the quoted
      // history is the original author's text and rides along on every edit.
      const history = (['quote', 'forward'] as const)
        .filter((t) => scan.counts[t] - storedScan.counts[t] > 0)
        .filter((t) => !notedHistory.has(t));
      for (const t of history) notedHistory.add(t);
      if (history.length > 0) tokenNotes.push(noteHistoryTokenStored([...history]));

      // Per-text budget, the escape note's gate rather than the count one, because a
      // near-miss is reported BY ITS LITERAL TEXT and two different spellings must not
      // cancel each other out.
      const storedMisses = new Map<string, number>();
      for (const m of storedScan.nearMisses) storedMisses.set(m.text, (storedMisses.get(m.text) ?? 0) + 1);
      for (const miss of scan.nearMisses) {
        const budget = storedMisses.get(miss.text) ?? 0;
        if (budget > 0) { storedMisses.set(miss.text, budget - 1); continue; }
        if (namedMisses.has(miss.text)) continue;
        namedMisses.add(miss.text);
        tokenNotes.push(noteNearMissToken(miss.text, miss.name));
      }
    }

    if (addedSpellings.length > 0) {
      tokenNotes.push(
        noteUnexpandedSpelling(describePartNames(addedSpellings), addedSpellings.length),
      );
    }

    // An html-alone edit drops a stored text part that may have been hand-written; say so.
    if (wroteHtml && !wroteText && !clearedText && !isBlank(existingTextValue)) {
      tokenNotes.push(noteDiscardedTextPart());
    }

    // A reply or forward prefix edited ONTO a draft that cannot thread (#188), with the
    // matcher compose uses.
    //
    // THE TRIGGER IS `updates.subject`, never the merged value, or every later body edit of
    // a prefixed draft would repeat the warning.
    //
    // ALL THREE stored markers count as earning the prefix: References alone threads, and a
    // forward carries only the forwarded id.
    const storedThreadMarkers =
      (existingEmail.inReplyTo?.length ?? 0) > 0
      || (existingEmail.references?.length ?? 0) > 0
      || (existingEmail['header:X-Forwarded-Message-Id:asMessageIds']?.length ?? 0) > 0;
    const prefixTyped = storedThreadMarkers ? undefined : matchSubjectPrefix(updates.subject);

    const mergedSubject = clear.has('subject') ? '' : (updates.subject !== undefined ? updates.subject : (existingEmail.subject || ''));
    const mergedTo      = clear.has('to')      ? [] : (updates.to      !== undefined ? updates.to.map(parseAddress)      : (existingEmail.to || []));
    const mergedCc      = clear.has('cc')      ? [] : (updates.cc      !== undefined ? updates.cc.map(parseAddress)      : (existingEmail.cc || []));
    const mergedBcc     = clear.has('bcc')     ? [] : (updates.bcc     !== undefined ? updates.bcc.map(parseAddress)     : (existingEmail.bcc || []));
    const mergedReplyTo = clear.has('replyTo') ? [] : (updates.replyTo !== undefined ? updates.replyTo.map(parseAddress) : (existingEmail.replyTo || null));

    // ---- The bodies that ship ----
    const mergedTextRaw = mergeText(workingText);
    const mergedHtmlRaw = mergeHtml(workingHtml);

    // ONLY when a body was written: a metadata-only edit must NOT inject a text part into an
    // html-only draft.
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

    // Reject a body-less RESULT only when this edit touched the body: a metadata-only edit
    // may run against a draft with no body yet. `touchesBody`, not `wroteAnyBody`, so
    // clearing the last body is caught.
    if (touchesBody && isBlank(textBodyValue) && isBlank(htmlBodyValue)) {
      throw new InvalidInputError('a draft needs a body; supply textBody or htmlBody (this edit would leave it with neither).');
    }

    // ---- Part assembly: apply the removals resolved earlier, work out what the surviving
    // parts still do for the body that actually ships, then append (#13) ----

    // Raised HERE so a body-shape error keeps precedence over an attachment one.
    if (removalPlan.error) throw removalPlan.error;

    const finalHtmlRefs = htmlCidRefs(htmlBodyValue);

    // A stored Content-ID this server cannot reproduce is refused rather than mangled. AFTER
    // the removals, so removing the offending part is not blocked, and only on a body edit:
    // the carry copies such a value verbatim, so other edits are safe and are the repair.
    if (touchesBody) {
      for (const part of removalPlan.survivors) {
        const cid = partCid(part);
        if (cid && !isRecreatableCid(cid)) throw new InvalidInputError(rejectUnrecreatableCid(cid));
      }
    }

    // What becomes of each surviving part now the shipping body is known (see
    // reconcileInlineParts). A part no longer displayed that is someone else's is demoted,
    // never dropped: those bytes are not this server's to discard. Body edits only.
    const reconciled = touchesBody
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
        // Keyed by POSITION, not by blob: two parts can share a blobId, and the ledger
        // replaces by key.
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

    // What survived, then the caller's own additions. This method mints no parts.
    let finalAttachments: AttachmentPart[] = carriedParts.slice();
    if (updates.attachments?.length) {
      // Marked inline exactly when the shipping body references it, which is only knowable
      // here (the upload cannot see the body the edit ends up with).
      const displays = (p: AttachmentPart) =>
        !isBlank(htmlBodyValue) && partCid(p) !== '' && finalHtmlRefs.includes(partCid(p));
      finalAttachments = finalAttachments.concat(
        updates.attachments.map((p, index) => {
          if (displays(p)) return { ...p, disposition: 'inline' };
          // A Content-ID the body does not display: kept, and reported. Keyed by position so
          // it cannot collide with a displayed file's identifier-keyed record.
          if (partCid(p) !== '') {
            ledger.record({ key: `supplied:${index}`, outcome: 'degraded', name: p.name });
          }
          return p;
        }),
      );
    }

    // ---- Checks over the assembled state, in the order a caller can act on ----

    const finalPartCids = new Set(
      finalAttachments.map((p) => (typeof p.cid === 'string' ? p.cid : '')).filter((c) => c !== ''),
    );
    const danglingRefs = finalHtmlRefs.filter((r) => !finalPartCids.has(r));

    // A reference left dangling BY THIS CALL'S REMOVAL: a diff, so a pre-existing break is
    // not attributed to it.
    const removedCids = new Set(removalPlan.removed.map(partCid).filter((c) => c !== ''));
    const removalCaused = danglingRefs.filter((r) => removedCids.has(r));
    if (removalCaused.length > 0) {
      throw new InvalidInputError(
        clear.has('attachments')
          ? rejectClearAttachmentsDanglingRefs()
          : rejectRemovalDanglingRef(removalCaused[0]),
      );
    }

    if (touchesBody) {
      // The stored body was already broken; this wording wins, since recreating the draft is
      // the repair either way. POST-MERGE, so an edit that fixes the references passes.
      const stillBroken = danglingRefs.filter((r) => storedDanglingRefs.includes(r));
      if (stillBroken.length > 0) throw new InvalidInputError(rejectBrokenDraft(stillBroken, availability));

      // A reserved identifier naming NO part this draft carries: the caller minted it, and no
      // later call can supply it. One that resolves is a handed-back embed and is fine.
      for (const ref of htmlCidRefs(callerWrittenHtml)) {
        if (isReservedCid(ref) && !finalPartCids.has(ref)) {
          throw new InvalidInputError(rejectReservedCidRef(ref));
        }
      }

      if (danglingRefs.length > 0) {
        // One the caller did not write came in with an expanded sign-off.
        const fromSignature = expandSignature
          && signatureCidRefs(editSignature).includes(danglingRefs[0])
          && !htmlCidRefs(callerWrittenHtml).includes(danglingRefs[0]);
        throw new InvalidInputError(fromSignature
          ? rejectSignatureEmbeddedImage(danglingRefs[0])
          : rejectDanglingCidRef(danglingRefs[0], availability));
      }
    }

    // Two parts sharing one identifier make every reference to it ambiguous.
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

    // Self-check, not a caller-facing rule. This method mints nothing, hence the empty list.
    checkInlineClosure({
      htmlBodies: touchesBody ? [htmlBodyValue] : [],
      finalPartCids: finalAttachments.map((p) => p.cid),
      attachedMintedCids: [],
      skip: !touchesBody,
    });

    // What the shipping body displays, keyed by identifier. Reported only when this call
    // could have CHANGED it, or every subject edit would repeat it.
    //
    // The basis is REFERENCE MEMBERSHIP, not stored disposition: a stored plain attachment
    // the edited body starts referencing is displayed. Compose reads the disposition, which
    // agrees there because it dispositions every part itself.
    const reportsEmbeds = touchesBody || !!updates.attachments?.length;
    const bytesByCid = new Map<string, number>();
    for (const part of storedParts) {
      if (partCid(part)) bytesByCid.set(partCid(part), partBytes(part));
    }
    const embeddedNow = reportsEmbeds
      ? finalAttachments.filter(
          (p) => typeof p.cid === 'string' && p.cid !== '' && finalHtmlRefs.includes(p.cid),
        )
      : [];

    const emailObject: any = {
      mailboxIds: existingEmail.mailboxIds,
      keywords: { ...(existingEmail.keywords || {}), $draft: true },
      from: [{ name: writtenFromName, email: writtenFromAddress }],
      to: mergedTo,
      cc: mergedCc,
      bcc: mergedBcc,
      subject: mergedSubject,
      ...(mergedReplyTo?.length && { replyTo: mergedReplyTo }),
      ...(existingEmail.inReplyTo && { inReplyTo: existingEmail.inReplyTo }),
      ...(existingEmail.references && { references: existingEmail.references }),
      // A foreign value failing Fastmail's header-SET validation fails the CREATE loudly with
      // the old draft intact; clearing the marking is the repair.
      ...(carriedForwardHeader.length > 0 && !dropForwardHeader && { 'header:X-Forwarded-Message-Id:asMessageIds': carriedForwardHeader }),
      ...(carriedSourceId !== undefined && !(dropForwardHeader && !isReply) && { [SOURCE_ID_HEADER]: carriedSourceId }),
      ...(finalAttachments.length && { attachments: finalAttachments }),
    };

    Object.assign(emailObject, buildBodyParts({ textBody: textBodyValue, htmlBody: htmlBodyValue }));

    const addressList = (addrs: any): string[] =>
      (addrs || []).map((a: any) => a?.email).filter((e: any): e is string => typeof e === 'string' && e !== '');
    const replacedTo = addressList(existingEmail.to);
    const replacedCc = addressList(existingEmail.cc);
    const replacedBcc = addressList(existingEmail.bcc);
    const replacedReplyTo = addressList(existingEmail.replyTo);
    const replacedDraft: ReplacedDraftInfo = {
      id: emailId,
      ...(existingEmail.subject && { subject: existingEmail.subject }),
      ...(replacedTo.length && { to: replacedTo }),
      ...(replacedCc.length && { cc: replacedCc }),
      ...(replacedBcc.length && { bcc: replacedBcc }),
      ...(replacedReplyTo.length && { replyTo: replacedReplyTo }),
      ...(existingTextValue !== undefined && { textBodySize: existingTextValue.length }),
      ...(existingHtmlValue !== undefined && { htmlBodySize: existingHtmlValue.length }),
    };

    // Create-then-dispose, NOT one combined Email/set. Fastmail silently no-ops an in-place
    // body update, so a recreate is mandatory, and RFC 8620 §6.3 does not make create and
    // destroy atomic: a server MAY destroy even when the create fails. Worst case here is a
    // duplicate, never a vanished draft. Do NOT "optimize" this into one call.
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
      this.throwSingleSetError(createResult.notCreated.draft, 'create updated draft');
    }

    const newEmailId = createResult.created?.draft?.id;
    if (!newEmailId) {
      throw new Error('Draft update returned no email ID');
    }

    // The edit has SUCCEEDED. The old copy moves to TRASH, never destroy (#65), so an edit
    // from a stale copy is a one-step undo. Only mailboxIds is patched; `$draft` stays, as on
    // delete_email, and getThread does not count a Trash-only draft as active.
    //
    // Any failure here leaves a duplicate with the OLD content: reported as an orphan, NOT
    // thrown, since the edit succeeded. Deliberately NO destroy fallback.
    let trashedOldDraftId: string | undefined;
    let orphanedOldDraftId: string | undefined;
    let orphanedOldDraftReason: string | undefined;
    try {
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
        // Confirm the move POSITIVELY, by key in `updated` (its values may be null): absence
        // from both maps means the server said nothing.
        const trashResult = this.getMethodResult(trashResponse, 0);
        const trashSetError = this.setErrorFor(trashResult.notUpdated, emailId);
        if (trashSetError) {
          orphanedOldDraftId = emailId;
          orphanedOldDraftReason = this.describeSetError(trashSetError);
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

    // Confirm what the saved draft carries, ONLY when this call attached an embedded image,
    // the one outcome the result would otherwise assert without evidence. Nothing here may
    // fail the edit, which has already succeeded.
    const followUpNotes: string[] = [];
    const attachedInlineCids = embeddedNow
      .map((p) => p.cid as string)
      .filter((cid) => appendedCidCounts.has(cid));
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

    // ---- The hash the caller's NEXT edit will need ----
    // Issued only where the caller already knows, byte for byte, what the saved draft's
    // governing part says (the html when one ships, else the text):
    //
    //  - A metadata-only edit issues none. It COPIES the bodies, so a held hash usually stays
    //    good, but that is a property of the DRAFT: a typeless body part is not carried (see
    //    `draftInterleavedTextType` in body-hash.ts), and re-filing an image out of a body list
    //    changes the hashed part set.
    //  - A flagged edit expanded `{{signature}}` into the body: withheld.
    //  - A governing part this server DERIVED during this call (e.g. a text fallback from new
    //    html) was never shown to the caller: withheld. The test is SEEN-NESS, not authorship:
    //    clearing htmlBody on a draft that never had one leaves stored text the caller saw, so
    //    that edit issues a hash.
    //
    // PROVENANCE is the whole point: a returned hash comes from a RE-READ of the saved draft,
    // never from the bytes sent, which would assert what the server stored on the strength of
    // what was asked. A failed re-read has NO fallback: no hash, and the reason.
    const governingSupplied = !isBlank(htmlBodyValue)
      ? wroteHtml
      : wroteText || (isBlank(existingHtmlValue) && textBodyValue === existingTextValue);
    let issuedBodyHash: string | undefined;
    let bodyHashWithheld: string | undefined;
    if (touchesBody) {
      if (expandSignature) {
        bodyHashWithheld = NOTE_BODY_HASH_AFTER_EXPANSION;
      } else if (!governingSupplied) {
        bodyHashWithheld = NOTE_BODY_HASH_DERIVED_PART;
      } else {
        try {
          const saved = await this.readDraftBody(newEmailId);
          // WHETHER THE SAVED BODY CAN BE HASHED IS THE READ SIDE'S RULE, ASKED RATHER THAN
          // RE-STATED, so this hash and a later `get_email`'s agree by construction.
          //
          // THE DESCRIPTOR IS A CLAIM about what the response showed: `readDraftBody` fetches
          // both lists whole, so anything less would withhold with a false read-scope reason.
          const outcome = resolveDraftBodyHash(saved, {
            bodyText: true,
            bodyHtml: true,
            stripQuoted: false,
          });
          if (!outcome) {
            // Undefined only for a non-draft: the re-read did not come back as what we wrote.
            // Raised so the read-failure note says so, rather than the field vanishing.
            throw new Error(`the saved draft "${describeUntrusted(newEmailId)}" did not read back as a draft`);
          }
          if ('bodyHash' in outcome) {
            issuedBodyHash = outcome.bodyHash;
          } else {
            bodyHashWithheld = noteBodyHashAfterReRead(outcome.bodyHashWithheld);
          }
        } catch (err) {
          bodyHashWithheld = noteBodyHashUnreadable(err instanceof Error ? err.message : String(err));
        }
      }
    }

    const tally = ledger.tally();
    const notes = [
      ...emitInlineNotes(tally, { surface: 'draft', keepNoun }),
      ...tokenNotes,
      ...followUpNotes,
      ...(prefixTyped ? [noteEditSubjectPrefix(prefixTyped)] : []),
    ];
    const touchedInlineImages = tally.embedded > 0 || tally.degraded > 0 || tally.removed > 0;

    return {
      id: newEmailId,
      replacedDraft,
      ...(trashedOldDraftId && { trashedOldDraftId }),
      ...(orphanedOldDraftId && { orphanedOldDraftId, orphanedOldDraftReason }),
      ...(issuedBodyHash !== undefined && { bodyHash: issuedBodyHash }),
      ...(bodyHashWithheld !== undefined && { bodyHashWithheld }),
      ...(touchedInlineImages && {
        inlineImages: { embedded: tally.embedded, degraded: tally.degraded, removed: tally.removed },
      }),
      ...(notes.length > 0 && { notes }),
    };
  }

  /**
   * Re-read a just-saved draft's body parts WITH values, for the hash the result hands back.
   * `keywords` is required: without it `resolveDraftBodyHash` reads a non-draft and the hash
   * silently goes missing.
   */
  private async readDraftBody(emailId: string): Promise<any> {
    const session = await this.getSession();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'keywords', 'textBody', 'htmlBody', 'bodyValues'],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    });
    const email = this.getListResult(response, 0)[0];
    if (!email) throw new Error(`the saved draft "${describeUntrusted(emailId)}" could not be read back`);
    return email;
  }

  /**
   * Re-read a draft's part listing (no values), to confirm what a just-saved draft carries.
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
    if (!email) throw new Error(`Email with ID "${describeUntrusted(emailId)}" not found`);
    return buildUnionParts(email).map((u: UnionPart) => u.part);
  }

  async sendDraft(emailId: string): Promise<SendDraftOutcome> {
    const session = await this.getSession();

    // The provenance headers ride on this read: the draft is submitted by reference, so
    // they cannot change during submission.
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: [
            'id', 'from', 'to', 'cc', 'bcc', 'replyTo', 'keywords', 'mailboxIds',
            'textBody', 'htmlBody', 'bodyValues',
            'attachments',
            'inReplyTo', 'header:X-Forwarded-Message-Id:asMessageIds', SOURCE_ID_HEADER,
          ],
          // `attachments` and disposition/cid/name are for the RECEIPT ONLY; they vet nothing.
          bodyProperties: ['partId', 'blobId', 'type', 'size', 'disposition', 'cid', 'name'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];
    if (!email) {
      throw new InvalidInputError(`Email with ID "${describeUntrusted(emailId)}" not found`);
    }

    if (!email.keywords?.$draft) {
      throw new InvalidInputError('Cannot send a non-draft email');
    }

    // Reject an empty body part before an irreversible send: another client's draft can carry
    // one. Reject rather than sanitise, since stripping it means a recreate; edit_draft is the
    // fix path.
    const blankPart = findBlankBodyPart(email);
    if (blankPart === 'htmlBody') {
      throw new InvalidInputError('This draft has an empty htmlBody that would render blank to recipients. Edit the draft to supply or clear htmlBody before sending.');
    }
    if (blankPart === 'textBody') {
      throw new InvalidInputError('This draft has an empty textBody that would render blank for plain-text recipients. Edit the draft to supply or clear textBody before sending.');
    }

    const allRecipients: { email: string }[] = [
      ...(email.to || []),
      ...(email.cc || []),
      ...(email.bcc || []),
    ];

    if (allRecipients.length === 0) {
      throw new InvalidInputError('Draft has no recipients. Edit the draft to add a to/cc/bcc recipient before sending.');
    }

    const fromEmail = email.from?.[0]?.email;
    if (!fromEmail) {
      throw new InvalidInputError('Draft has no from address. Edit the draft to set a from address before sending.');
    }

    // A stored wildcard PATTERN never goes on the wire (#160). The ONLY check a stored
    // pattern meets (see isWildcardIdentityEmail); the identity check below would pass it.
    if (isWildcardIdentityEmail(fromEmail)) {
      throw new InvalidInputError(
        `This draft's from address is the wildcard pattern "${describeUntrusted(fromEmail)}", not an address. ` +
        'Edit the draft with an explicit from before sending.',
      );
    }

    const identities = await this.getIdentities();
    const selectedIdentity = identityFor(identities, fromEmail);
    if (!selectedIdentity) {
      throw new InvalidInputError('From address on draft does not match any sending identity. Edit the draft to set a from address matching one of your verified identities before sending.');
    }

    // One mailbox fetch serves both the Drafts gate below and the Sent target further down.
    const mailboxes = await this.getMailboxes();

    // A DRAFT IS SENDABLE ONLY FROM THE DRAFTS FOLDER: filed anywhere else, someone has made
    // it something other than outbound mail. MEMBERSHIP, NOT EXCLUSIVITY: a label beside
    // Drafts is not a move.
    //
    // EXACT role, never a name substring, which in front of an irreversible send would
    // PERMIT a "Draft notes" folder. Both arms refuse: no drafts-role mailbox is not a permit.
    const draftsMailbox = this.findByExactRole(mailboxes, 'drafts');
    if (!draftsMailbox) {
      throw new Error(
        'Could not find a Drafts mailbox (no mailbox in this account carries the "drafts" role), ' +
        'so this draft cannot be confirmed to be in Drafts. send_draft only sends a draft that is ' +
        'in the Drafts folder.',
      );
    }
    // Read like this file's other mailboxIds reads (see setErrorFor): hasOwnProperty, or
    // "constructor" OPENS the gate; the VALUE must be `true`, not merely present; and
    // isPlainResponseMap. Nothing on Fastmail produces these shapes, but this check stands in
    // front of the only irreversible action here.
    //
    // AN UNREADABLE MAP GETS ITS OWN REFUSAL rather than `{}`, which would hand back a
    // move_email repair for a draft that may already be in Drafts.
    const filing = email.mailboxIds;
    if (!isPlainResponseMap(filing)) {
      throw new Error(
        'The server returned this draft with no readable mailboxIds, so this server cannot ' +
        'tell whether it is in the Drafts folder and will not send it. Nothing was sent. ' +
        'This is a fault in the response rather than in the call, and no tool here can ' +
        'repair it: report it rather than retrying.',
      );
    }
    const inMailbox = (id: string) =>
      Object.prototype.hasOwnProperty.call(filing, id) && filing[id] === true;

    if (!inMailbox(draftsMailbox.id)) {
      const filedIn = Object.keys(filing)
        .filter(id => filing[id] === true)
        .map(id => {
          const mailbox = mailboxes.find(mb => mb?.id === id);
          const name = describeUntrusted(mailbox?.name);
          if (mailbox && name.trim() !== '') return `"${name}"`;
          return `${mailbox ? 'unnamed' : 'unknown'} mailbox (id: "${describeUntrusted(id)}")`;
        });
      throw new InvalidInputError(
        'This draft is not in the Drafts folder, so it will not be sent' +
        (filedIn.length > 0 ? ` (it is in: ${joinCapped(filedIn)})` : '') +
        '. Move it back to Drafts with move_email and send it again.',
      );
    }

    // EXACT role, as for Drafts: a name substring would file the sent copy into
    // "Presentations". Refused BEFORE the submission, so nothing is transmitted.
    const sentMailbox = this.findByExactRole(mailboxes, 'sent');
    if (!sentMailbox) {
      throw new Error(
        'Could not find a Sent mailbox (no mailbox in this account carries the "sent" role), ' +
        'so there is nowhere to file the sent copy of this message. Nothing was sent.',
      );
    }

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
          // THE FILING IS A PATCH, NOT A REPLACEMENT: the send trades Drafts for Sent and
          // keeps every other label. Never add a whole-value `mailboxIds` beside these keys:
          // it would silently win and drop the labels.
          //
          // Removal BEFORE addition, so if the two keys ever name one mailbox the message
          // stays filed somewhere.
          onSuccessUpdateEmail: {
            '#submission': {
              [`mailboxIds/${draftsMailbox.id}`]: null,
              [`mailboxIds/${sentMailbox.id}`]: true,
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
      this.throwSingleSetError(submissionResult.notCreated.submission, 'submit draft');
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Draft submission returned no submission ID');
    }

    // The receipt, after the submission so it can never influence the send. Embedded means
    // routed into a body list or marked inline.
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
   * Two steps: a full-text query on the BARE id for recall (docs/email-bodies.md), which also
   * matches every message that merely mentions it, then an exact `messageId` comparison for
   * precision. The `header` FilterCondition would do it server-side but is unproven beyond
   * Fastmail.
   *
   * The OLDEST-FIRST sort is load-bearing: recall is capped, and the owner predates every
   * message that references it, so it is always on the first page.
   *
   * No Trash/Spam exclusion: this answers "which message is this".
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
   * message exists. Validates a draft's recorded source INSTANCE before send_draft marks it,
   * so a stale pointer falls back to the Message-ID lookup.
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

  async markEmailRead(emailId: string, read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = newUpdateMap();
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

    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, `mark email as ${read ? 'read' : 'unread'}`);
    }
  }

  // Additively set keyword flags without clobbering the others, for the reply path's
  // thread-state maintenance (#52/#54).
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

    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, 'add keywords to email');
    }
  }

  async pinEmail(emailId: string, pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = newUpdateMap();
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

    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, `${pinned ? 'pin' : 'unpin'} email`);
    }
  }

  async deleteEmail(emailId: string): Promise<void> {
    const session = await this.getSession();

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
    
    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, 'delete email');
    }
  }

  async moveEmail(emailId: string, target: string): Promise<void> {
    const session = await this.getSession();

    // Move-to-any-mailbox stays open by design; a target restriction is #43.
    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // WHOLE-VALUE, matching the promise ("replaces all mailbox membership"); a read-then-patch
    // races a mailbox added in between. NO keyword is written.
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: { mailboxIds: { [targetMailboxId]: true } }
          }
        }, 'moveEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, 'move email');
    }
  }

  /**
   * Archive one or more emails the way the Fastmail client does: REMOVE the Inbox
   * membership and leave everything else alone, adding the Archive mailbox only when
   * removing the Inbox would otherwise leave the message filed nowhere.
   *
   * NOT a whole-value membership replace: an Inbox + label message keeps the label and gets
   * no Archive (measured, docs/fastmail-action-availability.md).
   *
   * Archive is resolved by EXACT ROLE ONLY, with deliberately no destination parameter: a
   * mailbox NAMED "archive" can be created by text a model merely read.
   *
   * NO keyword is written, but read state can still change: $seen is reported only when
   * EVERY per-mailbox record carries it, so dropping an unread Inbox record can flip it.
   *
   * Never throws per message; every id lands in one of the six buckets. Throws only
   * account-wide. Before any write, saying "Nothing was archived": the read failed or was
   * incomplete; no `inbox` role (a plain Error, since no other call can do this); an
   * Inbox-only message with no `archive` role (InvalidInputError naming move_email). After
   * dispatch, an Email/set that failed outright, which leaves every id's outcome unknown.
   */
  async archiveEmails(emailIds: string[]): Promise<ArchiveResult> {
    const session = await this.getSession();

    // De-duplicated so one id cannot occupy two buckets.
    const ids = [...new Set(emailIds)];

    // One batch for both reads; getMailboxes() would add a round trip.
    const readRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/get', {
          accountId: session.accountId,
          properties: ['id', 'name', 'role'],
        }, 'mailboxes'],
        ['Email/get', {
          accountId: session.accountId,
          ids,
          properties: ['id', 'mailboxIds'],
        }, 'emails'],
      ],
    };

    // Wrapped so EVERY pre-write abort, transport and method errors included, carries the
    // "Nothing was archived." tell the tool description tells callers to key off. Appended,
    // so the server's diagnosis survives; the InvalidInputError arm keeps a future
    // caller-fixable read error in its class.
    let readResponse: JmapResponse;
    let mailboxes: any[];
    let emailsResult: any;
    try {
      readResponse = await this.makeRequest(readRequest);
      // Not readListResultIfPresent: neither read is optional. Array.isArray because
      // getListResult hands a non-array `list` straight back, which would die outside this
      // block without the tell; [] routes it to the empty-list guard.
      const rawMailboxes = this.getListResult(readResponse, 0);
      mailboxes = Array.isArray(rawMailboxes) ? rawMailboxes : [];
      emailsResult = this.getMethodResult(readResponse, 1);
    } catch (error: any) {
      const detail = error instanceof Error ? error.message : String(error);
      const suffixed = `${detail} Nothing was archived.`;
      if (error instanceof InvalidInputError) throw new InvalidInputError(suffixed);
      throw new Error(suffixed);
    }
    if (mailboxes.length === 0) {
      throw new Error(
        'Could not read this account\'s mailboxes, so there is no way to tell which messages are in the Inbox. Nothing was archived.'
      );
    }

    const inboxMailbox = this.findByExactRole(mailboxes, 'inbox');
    // Not defensive: without it the tool silently becomes a no-op reporting success, or
    // writes "mailboxIds/undefined", and `any | undefined` hides that from the compiler.
    if (!inboxMailbox) {
      throw new Error(
        'This account has no mailbox with the inbox role, so there is no Inbox membership to remove. Nothing was archived.'
      );
    }
    const inboxId: string = inboxMailbox.id;
    const archiveMailbox = this.findByExactRole(mailboxes, 'archive');

    const emails: any[] = Array.isArray(emailsResult?.list) ? emailsResult.list : [];
    // Array.isArray, not `|| []`: a non-array would throw a bare TypeError. Its ids are then
    // reported as unaccounted below, so nothing is swallowed.
    const rawNotFound = emailsResult?.notFound;
    const notFoundIds: string[] = Array.isArray(rawNotFound) ? rawNotFound : [];
    const byId = new Map<string, any>();
    for (const email of emails) {
      if (email && typeof email.id === 'string') byId.set(email.id, email);
    }
    const notFound = new Set<string>(notFoundIds);
    // Fail closed on an id the server accounted for in neither list.
    const unaccounted = ids.filter(id => !byId.has(id) && !notFound.has(id));
    if (unaccounted.length > 0) {
      // Caller-supplied ids: sanitised and capped (docs/conventions.md, untrusted values in prose).
      const shown = unaccounted.slice(0, EMAIL_ID_LIST_CAP).map(describeUntrusted).join(', ');
      const more = unaccounted.length > EMAIL_ID_LIST_CAP ? `, …and ${unaccounted.length - EMAIL_ID_LIST_CAP} more` : '';
      throw new Error(
        `The server returned neither a result nor a not-found entry for ${unaccounted.length} of the ${ids.length} requested email(s): ${shown}${more}. Nothing was archived.`
      );
    }

    const roleById = new Map<string, string | null>();
    for (const mb of mailboxes) {
      if (mb && typeof mb.id === 'string') {
        roleById.set(mb.id, typeof mb.role === 'string' && mb.role ? mb.role.toLowerCase() : null);
      }
    }

    // Decide every branch BEFORE writing, so the missing-Archive guard can ask whether any
    // message needs Archive.
    const decisions = new Map<string, { branch: ArchiveBranch; keptIds: string[]; refusingRole?: string; currentIds: string[] }>();
    // Unreadable filing, per id, with the sentence saying WHY: a per-message `failed`, not a
    // throw, and each way in is a different fact so each gets its own sentence.
    const unreadableFiling = new Map<string, string>();
    for (const id of ids) {
      const email = byId.get(id);
      if (!email) continue;
      // Never default to {}: an empty membership decides `notInInbox` and reports a false
      // "already archived". An array does the same through its index keys.
      const filing = email.mailboxIds;
      if (!isPlainResponseMap(filing)) {
        unreadableFiling.set(id, 'The server returned this message with no readable mailboxIds object, so its current filing is unknown. Nothing was written for it.');
        continue;
      }
      // Truthy values only: every writing branch RE-ASSERTS what it read, so a
      // `{"mb-label": false}` would be written back as a membership it never had.
      const currentIds = Object.keys(filing).filter(mbId => filing[mbId]);
      // An empty membership is a server fault, refused like an absent one. The guard lives
      // here, not in decideArchiveBranch, which is total by design.
      if (currentIds.length === 0) {
        const hadEntries = Object.keys(filing).length > 0;
        unreadableFiling.set(id, hadEntries
          ? 'The server returned a mailboxIds for this message in which no entry is set to true, so it is filed in no mailbox — which is not a state a message can be in. Its current filing is unknown and nothing was written for it.'
          : 'The server returned an empty mailboxIds for this message, which is not a filing a message can have, so its current filing is unknown. Nothing was written for it.');
        continue;
      }
      decisions.set(id, { ...decideArchiveBranch(currentIds, inboxId, roleById), currentIds });
    }

    // Raised only when a message reaches the Inbox-only branch, or a batch this tool can
    // serve would be rejected. Resolving Archive INTO keptIds leaves one destination set for
    // both the write and the report. It rejects the WHOLE batch: an account-wide condition,
    // not a per-message one.
    //
    // DO NOT delete it as a can't-happen path: without it the patch key becomes
    // "mailboxIds/undefined", a silently corrupt write.
    if ([...decisions.values()].some(d => d.branch === 'movedToArchive')) {
      if (!archiveMailbox) {
        throw new InvalidInputError(
          'This account has no mailbox with the archive role, and at least one of these messages is filed ONLY in the Inbox, so removing it from the Inbox would leave it filed nowhere. Nothing was archived. ' +
          'Use move_email with a destination of your choice instead.'
        );
      }
      for (const decision of decisions.values()) {
        if (decision.branch === 'movedToArchive') decision.keptIds = [archiveMailbox.id];
      }
    }

    const update: Record<string, any> = newUpdateMap();
    for (const [id, decision] of decisions) {
      // Remove Inbox, then RE-ASSERT a non-empty destination set. Not redundancy, and not
      // about concurrency: Cyrus's emptiness guard counts mailboxes the message was expunged
      // from, so a lone null can destroy the message. Why this patches where move_email
      // writes whole-value: docs/conventions.md, "The membership-subtracting tools reverse
      // that convention, deliberately".
      if (decision.branch !== 'removedFromInbox' && decision.branch !== 'movedToArchive') continue;
      const patch: Record<string, any> = { [`mailboxIds/${inboxId}`]: null };
      for (const keptId of decision.keptIds) patch[`mailboxIds/${keptId}`] = true;
      update[id] = patch;
    }

    let notUpdated: Record<string, any> = {};
    let updated: Record<string, any> = {};
    const wrote = Object.keys(update).length > 0;
    if (wrote) {
      const writeResponse = await this.makeRequest({
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/set', { accountId: session.accountId, update }, 'archiveEmails'],
        ],
      });
      const setResult = this.getMethodResult(writeResponse, 0);
      // isPlainResponseMap, not `|| {}`: for a STRING, hasOwnProperty('0') is true and would
      // fabricate a set-error (see setErrorFor, which this path cannot call).
      notUpdated = isPlainResponseMap(setResult?.notUpdated) ? setResult.notUpdated : {};
      // An id in neither map has no known outcome. RFC 8620 §5.3 forbids that and Cyrus
      // complies; this makes a server that stops complying a reported failure, not a
      // false success.
      updated = isPlainResponseMap(setResult?.updated) ? setResult.updated : {};
    }

    const infoMap = buildMailboxInfoMap(mailboxes);
    const results: ArchiveEmailResult[] = [];
    for (const id of ids) {
      const unreadable = unreadableFiling.get(id);
      if (unreadable) {
        results.push({ id, action: 'failed', reason: { description: unreadable } });
        continue;
      }

      const decision = decisions.get(id);
      if (!decision) {
        // Email/get put it in notFound; there is no filing to report.
        results.push({ id, action: 'notFound' });
        continue;
      }

      // Only an id this call wrote can carry a set-error. hasOwnProperty on both maps
      // ("constructor"). The KEY's presence is the refusal: a falsy value is still one.
      const wasWritten = Object.prototype.hasOwnProperty.call(update, id);
      const wasRefused = wasWritten && Object.prototype.hasOwnProperty.call(notUpdated, id);
      const setError = wasRefused ? (notUpdated[id] || {}) : undefined;
      if (setError) {
        // A write-time `notFound` is the read's notFound a round trip later, so same bucket.
        // An empty-after-patch is NOT this: it is `invalidProperties` on "mailboxIds"
        // (Cyrus jmap_mail.c:13511), and lands in `failed`.
        const action: ArchiveAction = setError.type === 'notFound' ? 'notFound' : 'failed';
        if (action === 'notFound') {
          results.push({ id, action, reason: { setErrorType: setError.type, description: setError.description } });
        } else {
          results.push({
            id,
            action,
            ...describeMailboxIds(decision.currentIds, infoMap),
            reason: { setErrorType: setError.type, description: setError.description },
          });
        }
        continue;
      }

      if (decision.branch === 'removedFromInbox' || decision.branch === 'movedToArchive') {
        // hasOwnProperty: RFC 8620 §5.3 allows a null value for a successful update.
        if (!Object.prototype.hasOwnProperty.call(updated, id)) {
          results.push({
            id,
            action: 'failed',
            ...describeMailboxIds(decision.currentIds, infoMap),
            reason: {
              outcomeUnknown: true,
              description: 'The server acknowledged this id in neither the updated nor the notUpdated map, so the outcome of the write is unknown. The filing shown is what was read BEFORE the write.',
            },
          });
          continue;
        }
        // PROJECTED from the pre-write read, not read back.
        results.push({ id, action: decision.branch, ...describeMailboxIds(decision.keptIds, infoMap) });
        continue;
      }

      // OBSERVED filing, unchanged — nothing was written for these.
      results.push({
        id,
        action: decision.branch,
        ...describeMailboxIds(decision.currentIds, infoMap),
        ...(decision.refusingRole ? { reason: { role: decision.refusingRole } } : {}),
      });
    }

    const counts: Record<ArchiveAction, number> = {
      movedToArchive: 0, removedFromInbox: 0, notInInbox: 0, refused: 0, notFound: 0, failed: 0,
    };
    for (const result of results) counts[result.action]++;

    return { results, counts };
  }

  // Resolve an ARRAY of mailbox inputs exactly (findMailboxExact), for the label arrays and
  // search_emails' scope arrays (#50, #27, #26).
  //
  // Collects EVERY failure, each in its own bucket, so one retry fixes them all and an
  // ambiguity is never reported as a typo.
  //
  // All-or-nothing: a dropped scope entry would silently widen a search, a dropped label
  // half-apply a write. Duplicates are NOT collapsed here (the label patch keys by id; the
  // scope arrays de-duplicate at their call site). A real id absent from the live list is
  // rejected, an accepted residual (docs/security-model.md).
  private async resolveMailboxIdList(inputs: string[], mailboxList?: any[]): Promise<string[]> {
    const mailboxes = mailboxList ?? await this.getMailboxes();
    const resolved: string[] = [];
    const failures: MailboxResolutionFailures = { notFound: [], ambiguous: [], nameVsPath: [], unwalkable: [] };
    for (const input of inputs) {
      const match = findMailboxExact(mailboxes, input);
      const raw = String(input).trim();
      if (match && 'mailbox' in match) resolved.push(match.mailbox.id);
      else if (match && 'ambiguous' in match) {
        const bucket = match.nameVsPath ? failures.nameVsPath : failures.ambiguous;
        bucket.push({ input: raw, candidates: match.candidates });
      }
      else if (match && 'unwalkable' in match) failures.unwalkable.push({ input: raw, id: match.id });
      else failures.notFound.push(raw);
    }
    const failed = failures.notFound.length + failures.ambiguous.length
      + failures.nameVsPath.length + failures.unwalkable.length;
    if (failed > 0) {
      if (failures.notFound.length === failed) {
        throw new InvalidInputError(formatMailboxesNotFound(failures.notFound, mailboxes || []));
      }
      throw new InvalidInputError(formatMailboxesNotResolved(failures, mailboxes || []));
    }
    return resolved;
  }

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const resolvedIds = await this.resolveMailboxIdList(mailboxIds, mailboxes);
    assertLabelNamespace(resolvedIds, mailboxes);

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

    const setError = this.setErrorFor(result.notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, 'add labels to email');
    }
  }

  /** Caller-supplied email ids rendered into prose, capped and made safe. */
  private static nameEmailIds(list: string[]): string {
    const shown = list.slice(0, EMAIL_ID_LIST_CAP)
      .map(describeUntrusted).join(', ');
    return list.length > EMAIL_ID_LIST_CAP
      ? `${shown}, …and ${list.length - EMAIL_ID_LIST_CAP} more`
      : shown;
  }

  /**
   * Shared write for remove_labels and bulk_remove_labels.
   *
   * A bare removal of a message's LAST mailbox can DESTROY it: Cyrus's emptiness guard
   * counts tombstoned memberships (scripts/probes/label-emptiness.probe.mjs), and every
   * message archive_email has touched carries one.
   *
   * So, like Fastmail's own client, an emptying removal rehomes into the archive-role
   * mailbox in the same patch, the same rescue archiveEmails applies. Surviving memberships
   * are RE-ASSERTED on every patch (see archiveEmails); do not reduce this to a bare null.
   *
   * Returns notUpdated so each caller keeps its own error shape.
   */
  private async applyLabelRemoval(
    emailIds: string[],
    mailboxIds: string[],
    callId: string
  ): Promise<LabelRemovalOutcome> {
    const session = await this.getSession();
    const ids = [...new Set(emailIds)];

    // Every failure of THESE READS carries the "Nothing was changed." tell, as in
    // archiveEmails. Scoped to the reads: the resolver below throws its own input error.
    let mailboxes: any[];
    let readResult: any;
    try {
      const [mailboxList, readResponse] = await Promise.all([
        this.getMailboxes(),
        this.makeRequest({
          using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
          methodCalls: [
            ['Email/get', {
              accountId: session.accountId,
              ids,
              properties: ['id', 'mailboxIds'],
            }, 'currentFiling'],
          ],
        }),
      ]);
      // Array.isArray: see the same read in archiveEmails.
      mailboxes = Array.isArray(mailboxList) ? mailboxList : [];
      readResult = this.getMethodResult(readResponse, 0);
    } catch (error: any) {
      const detail = error instanceof Error ? error.message : String(error);
      const suffixed = `${detail} Nothing was changed.`;
      if (error instanceof InvalidInputError) throw new InvalidInputError(suffixed);
      throw new Error(suffixed);
    }

    // Before the resolver, which would blame the caller for a server-side problem.
    if (mailboxes.length === 0) {
      throw new Error(
        'Could not read this account\'s mailboxes, so there is no way to tell which labels these messages carry. Nothing was changed.'
      );
    }

    const resolvedIds = await this.resolveMailboxIdList(mailboxIds, mailboxes);
    // Also why no removal can take Archive away or collide with the rescue: a role mailbox
    // never enters `removing`.
    assertLabelNamespace(resolvedIds, mailboxes);
    const removing = new Set(resolvedIds);

    // A non-array `list` or `notFound` must NOT degrade to empty: unknown filing is exactly
    // where a removal destroys a message.
    if (!Array.isArray(readResult?.list) || !Array.isArray(readResult?.notFound)) {
      throw new Error(
        'The server\'s response to the filing read was malformed (list and notFound must both be arrays), so no message\'s current filing is known. Nothing was changed.'
      );
    }
    const byId = new Map<string, any>();
    for (const email of readResult.list) {
      if (email && typeof email.id === 'string') byId.set(email.id, email);
    }
    const notFound = new Set<string>(readResult.notFound);

    const archiveMailbox = this.findByExactRole(mailboxes, 'archive');
    const update: Record<string, any> = newUpdateMap();
    // Unaccounted or unreadable filing: fails closed, never reaches the Email/set.
    const unknownFiling: string[] = [];
    // Collected so the error can name the ids. Both abort the WHOLE batch, unlike
    // archiveEmails, only because these tools have no per-message report; if they gain one,
    // move these there.
    const needArchiveRole: string[] = [];
    const rescued: string[] = [];
    const unchanged: string[] = [];

    for (const id of ids) {
      const email = byId.get(id);
      if (!email) {
        if (!notFound.has(id)) {
          // Said nothing is not "does not exist": do NOT fall through to a removal.
          unknownFiling.push(id);
          continue;
        }
        // Omitted from the write; its set-error is synthesized below.
        continue;
      }

      const filing = email.mailboxIds;
      if (!isPlainResponseMap(filing)) {
        unknownFiling.push(id);
        continue;
      }
      // Truthy values only, so a {id: false} entry is never re-asserted as a real membership.
      const currentIds = Object.keys(filing).filter(mbId => filing[mbId]);
      if (currentIds.length === 0) {
        unknownFiling.push(id);
        continue;
      }

      const survivors = currentIds.filter(mbId => !removing.has(mbId));
      // None of the named labels is here: skip, since a re-assert could resurrect a
      // membership a concurrent client just removed. Recorded, so a no-op call does not
      // render like a full relabel.
      if (survivors.length === currentIds.length) {
        unchanged.push(id);
        continue;
      }

      const keptIds = [...survivors];
      if (keptIds.length === 0) {
        if (!archiveMailbox) {
          needArchiveRole.push(id);
          continue;
        }
        keptIds.push(archiveMailbox.id);
        rescued.push(id);
      }

      const patch: Record<string, any> = {};
      for (const mbId of resolvedIds) patch[`mailboxIds/${mbId}`] = null;
      for (const keptId of keptIds) patch[`mailboxIds/${keptId}`] = true;
      update[id] = patch;
    }

    if (unknownFiling.length > 0) {
      throw new Error(
        `The server did not report a readable current filing for ${unknownFiling.length} of the ${ids.length} requested email(s): ${JmapClient.nameEmailIds(unknownFiling)}. ` +
        'Removing a label without knowing what else holds the message risks destroying it, so nothing was changed. Re-read them with get_email and try again.'
      );
    }
    if (needArchiveRole.length > 0) {
      throw new InvalidInputError(
        `This account has no mailbox with the archive role, and removing these labels would leave ${needArchiveRole.length} message(s) filed nowhere, which would destroy them: ${JmapClient.nameEmailIds(needArchiveRole)}. ` +
        'Nothing was changed, for any message in this call. Use move_email (or bulk_move) to file them somewhere else instead.'
      );
    }

    // No Email/set for an empty update, which would advance the account state for nothing.
    //
    // newUpdateMap(), not {}: an id of `__proto__` is assigned into it below.
    const notUpdated: Record<string, any> = newUpdateMap();
    let wroteEmailSet = false;
    let updatedCount = 0;
    if (Object.keys(update).length > 0) {
      wroteEmailSet = true;
      const response = await this.makeRequest({
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/set', { accountId: session.accountId, update }, callId],
        ],
      });
      const result = this.getMethodResult(response, 0);
      // BEFORE countAcknowledged, so an id it adds is excluded from that count.
      Object.assign(notUpdated, JmapClient.withUnaccountedFailures(Object.keys(update), result?.updated, result?.notUpdated));
      updatedCount = JmapClient.countAcknowledged(Object.keys(update), result?.updated, notUpdated);
    }

    // What the server would have reported, written after its own map so a real set-error
    // wins. Keyed off the ids the loop ACTUALLY skipped: an id in both `list` and `notFound`
    // was written.
    for (const id of ids) {
      if (!byId.has(id) && notFound.has(id) &&
          !Object.prototype.hasOwnProperty.call(notUpdated, id)) {
        notUpdated[id] = { type: 'notFound' };
      }
    }

    // A rescue is only a fact once the write carrying it succeeded.
    const confirmedRescued = wroteEmailSet
      ? rescued.filter(id => !Object.prototype.hasOwnProperty.call(notUpdated, id))
      : [];

    return {
      notUpdated,
      rescued: confirmedRescued,
      unchangedCount: unchanged.length,
      distinctCount: ids.length,
      updatedCount,
    };
  }

  async removeLabels(emailId: string, mailboxIds: string[]): Promise<LabelRemovalResult> {
    const { notUpdated, rescued, unchangedCount } =
      await this.applyLabelRemoval([emailId], mailboxIds, 'removeLabels');

    const setError = this.setErrorFor(notUpdated, emailId);
    if (setError) {
      this.throwSingleSetError(setError, 'remove labels from email');
    }
    return { rescued, unchangedCount };
  }

  async bulkAddLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const resolvedIds = await this.resolveMailboxIdList(mailboxIds, mailboxes);
    assertLabelNamespace(resolvedIds, mailboxes);

    const patch: Record<string, any> = {};
    resolvedIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const updates: Record<string, any> = newUpdateMap();
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

    // Not emailIds.length: duplicate ids collapse in the map.
    const submittedIds = Object.keys(updates);
    const notUpdated = JmapClient.withUnaccountedFailures(submittedIds, result.updated, result.notUpdated);
    if (Object.keys(notUpdated).length > 0) {
      const successCount = JmapClient.countAcknowledged(submittedIds, result.updated, notUpdated);
      this.throwBulkSetError(notUpdated, submittedIds.length, successCount, 'add labels to', buildIdCollapseNote(emailIds));
    }
  }

  async bulkRemoveLabels(emailIds: string[], mailboxIds: string[]): Promise<LabelRemovalResult> {
    const { notUpdated, rescued, unchangedCount, distinctCount, updatedCount } =
      await this.applyLabelRemoval(emailIds, mailboxIds, 'bulkRemoveLabels');

    if (Object.keys(notUpdated).length > 0) {
      const notes: string[] = [];
      if (rescued.length > 0) {
        // Rides on the error, or a message relocated to Archive is never mentioned.
        notes.push(`Of the messages that did succeed, ${rescued.length} had no mailbox left and ${rescued.length === 1 ? 'was' : 'were'} filed in Archive: ${JmapClient.nameEmailIds(rescued)}.`);
      }
      if (unchangedCount > 0) {
        // Same wording as formatLabelRemoval's success path (response-formatters.ts).
        notes.push(`${unchangedCount} of the messages named did not carry any of these labels and ${unchangedCount === 1 ? 'was' : 'were'} left untouched.`);
      }
      const collapseNote = buildIdCollapseNote(emailIds);
      if (collapseNote) notes.push(collapseNote);
      this.throwBulkSetError(
        notUpdated,
        // Unchanged ids were never submitted, so they are out of this total (notFound ids,
        // synthesized as failures, stay in). Deliberately unlike formatLabelRemoval's
        // success-path `total`, which counts every message named.
        distinctCount - unchangedCount,
        updatedCount,
        'remove labels from',
        notes.length > 0 ? notes.join(' ') : undefined,
      );
    }
    return { rescued, unchangedCount };
  }

  async getEmailAttachments(emailId: string): Promise<EmailAttachmentsResult> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          // The body lists too: an embedded image can be routed there instead of into
          // `attachments` (RFC 8621 §4.1.4, #13).
          properties: ['attachments', 'textBody', 'htmlBody'],
          // Pinned, so telling inline parts apart never rides on a server default.
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
   * Resolve an attachment reference (precedence: resolveAttachmentRef) to its blob
   * metadata.
   *
   * The SINGLE attachment-resolution path: every consumer resolves here once and passes
   * the result on, so metadata and bytes never come from two different reads.
   *
   * Every failure here is caller-fixable, so InvalidInputError; callers add context but
   * never reclassify. A transport failure stays a plain Error.
   */
  async getAttachmentInfo(emailId: string, attachmentId: string): Promise<AttachmentInfo> {
    const session = await this.getSession();

    // Body lists too, as in getEmailAttachments (#13).
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
      throw new InvalidInputError(
        'Email not found: that emailId matches no message. ' +
        'Pass an id from list_emails, search_emails or get_thread.'
      );
    }

    const parts = buildUnionParts(email).map((u) => u.part);
    const resolved = resolveAttachmentRef(parts, attachmentId);

    if (!resolved) {
      throw new InvalidInputError(
        `Attachment not found: attachmentId "${describeUntrusted(attachmentId)}" matches no part of that message. ` +
        'List its parts with get_email_attachments and pass a partId or blobId from it.'
      );
    }
    const attachment = resolved.part;

    // Here, not at URL build, so it covers every consumer and keeps `blobId` non-optional.
    if (!attachment.blobId) {
      throw new InvalidInputError(
        'That part has no downloadable content: it carries no blobId. ' +
        'Pick a part from get_email_attachments that has one.'
      );
    }

    return {
      blobId: attachment.blobId,
      type: attachment.type || 'application/octet-stream',
      // Sender-supplied and possibly used as a save name: sanitized once, at resolution.
      name: sanitizeDownloadFilename(attachment.name),
      size: attachment.size,
      matchedBy: resolved.matchedBy,
    };
  }

  /**
   * Build the blob download URL for an ALREADY-resolved attachment, which keeps
   * resolution single (see getAttachmentInfo).
   */
  private async downloadUrlFor(info: AttachmentInfo): Promise<string> {
    const session = await this.getSession();

    const downloadUrl = session.downloadUrl;
    if (!downloadUrl) {
      throw new Error('Download capability not available in session');
    }

    const url = fillUrlTemplate(downloadUrl, {
      '{accountId}': session.accountId,
      '{blobId}': info.blobId,
      '{type}': encodeURIComponent(info.type),
      '{name}': encodeURIComponent(info.name),
    });

    // Re-validate after substitution: the session-time check saw only the template, and a
    // placeholder value could rewrite the origin the bearer token goes to.
    validateFastmailUrl(url, 'downloadUrl', this.auth.getAllowUnsafe());

    return url;
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    return this.downloadUrlFor(await this.getAttachmentInfo(emailId, attachmentId));
  }

  static readonly DEFAULT_DOWNLOADS_DIR = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  static validateSavePath(savePath: string, downloadDir?: string): string {
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    // Relative paths resolve against the allowed dir, not the unpredictable cwd; the
    // containment check is the boundary. Case-sensitive on the write side, unlike the read
    // guard (docs/security-model.md).
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
          throw new PathAccessError(`Could not find an existing ancestor for path: "${echoPath(lexical)}"`);
        }
        ancestor = parent;
      }
    }

    // Canonicalize the existing ancestor — this is what catches symlink escapes.
    const canonicalAncestor = await realpath(ancestor);
    if (canonicalAncestor !== canonicalAllowed && !canonicalAncestor.startsWith(canonicalAllowed + sep)) {
      throw new PathAccessError(
        `path resolves to "${echoPath(canonicalAncestor)}" which is outside the allowed directory "${echoPath(canonicalAllowed)}". ` +
        `Refusing to follow symlink escape.`,
      );
    }

    const safePath = join(canonicalAncestor, ...missingSegments, basename(lexical));

    // If a symlink already exists at the target, refuse — writing through it
    // would still escape the allowed directory.
    try {
      const lst = await lstat(safePath);
      if (lst.isSymbolicLink()) {
        throw new PathAccessError(`Refusing to overwrite an existing symlink at the target: "${echoPath(safePath)}"`);
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
        // Operator-set, but echoed by the same rule as the caller's path beside it.
        throw new PathAccessError(`FASTMAIL_ATTACH_DIR ("${echoPath(allowedDir)}") does not exist. Create it or fix the path, then restart.`);
      }
      throw e;
    }

    let handle: FileHandle;
    try {
      handle = await open(lexical, 'r');
    } catch (e: any) {
      if (e.code === 'ENOENT') {
        throw new PathAccessError(`File not found: "${echoPath(inputPath)}" (resolved under "${echoPath(allowedDir)}").`);
      }
      // Not a duplicate of isFile() below, which is the real guard (Windows open() succeeds
      // on a directory; FIFOs and devices raise no EISDIR). This only translates the POSIX
      // errno so a raw error does not escape.
      if (e.code === 'EISDIR') {
        throw new PathAccessError(`Not a regular file: "${echoPath(inputPath)}".`);
      }
      throw e;
    }

    try {
      const st = await handle.stat();
      if (!st.isFile()) {
        throw new PathAccessError(`Not a regular file: "${echoPath(inputPath)}".`);
      }
      // Re-verify against the canonical full target — this catches a symlinked leaf or
      // an intermediate-dir symlink that escapes the root.
      const canonicalTarget = await realpath(lexical);
      if (!isPathContained(canonicalTarget, canonicalAllowed, caseInsensitive)) {
        throw new PathAccessError(
          `path resolves to "${echoPath(canonicalTarget)}" which is outside the allowed directory "${echoPath(canonicalAllowed)}". Refusing to follow symlink escape.`
        );
      }
      return { handle, size: st.size };
    } catch (e) {
      await handle.close().catch(() => {});
      throw e;
    }
  }

  /**
   * Upload a single blob and return its server-assigned blobId. Deliberately NOT
   * getAuthHeaders() (it hardcodes application/json) and NOT a JSON body. The returned
   * `type` echoes the Content-Type sent; the server does not sniff content.
   */
  async uploadBlob(data: Buffer, contentType: string): Promise<{ blobId: string; type: string; size: number }> {
    const session = await this.getSession();
    if (!session.uploadUrl) {
      throw new Error('Upload capability not available in session');
    }

    const maxSize = session.capabilities?.['urn:ietf:params:jmap:core']?.maxSizeUpload;
    if (typeof maxSize === 'number' && data.length > maxSize) {
      // Caller-fixable, so not InternalError, which would invite a doomed retry.
      throw new InvalidInputError(`Attachment is ${data.length} bytes; the server's upload limit is ${maxSize} bytes`);
    }

    const url = fillUrlTemplate(session.uploadUrl, { '{accountId}': session.accountId });
    // Re-validate after substitution, as in downloadUrlFor.
    validateFastmailUrl(url, 'uploadUrl', this.auth.getAllowUnsafe());

    // `new Uint8Array(data)` copies into a Uint8Array<ArrayBuffer>, which fetch's BodyInit
    // accepts where a Buffer view is not, so no `any` is needed.
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': this.auth.getAuthHeaders()['Authorization'],
        'Content-Type': contentType,
      },
      body: new Uint8Array(data),
      // Never follow a redirect with the token (see fetchAttachmentBuffer).
      redirect: 'error',
    });

    if (!response.ok) {
      throw new Error(`Blob upload failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json() as any;
    if (!result || typeof result.blobId !== 'string') {
      throw new Error('Blob upload returned no blobId');
    }
    // A size mismatch is a truncated upload: fail rather than attach a corrupt file.
    if (typeof result.size === 'number' && result.size !== data.length) {
      throw new Error(`Blob upload size mismatch: sent ${data.length} bytes, server stored ${result.size}.`);
    }
    return { blobId: result.blobId, type: result.type || contentType, size: result.size ?? data.length };
  }

  /**
   * Turn each attachment spec into the JMAP attachment part to splat into an Email.
   *
   * THREE SOURCES, TWO GATES, per SOURCE not per call: `path` reads local disk and stays
   * behind `attachDir` (FASTMAIL_ATTACH_DIR, re-checked here as well as in safeReadPath);
   * `blobId` and `emailId`+`attachmentId` reference content already in the account and are
   * gated on `allowBlobAttach` (FASTMAIL_ALLOW_BLOB_ATTACH).
   *
   * The size caps apply to local reads ONLY; a referenced blob is never read client-side
   * (docs/security-model.md).
   *
   * Each part is a FRESH literal, never the carriedAttachments shape, whose server-set
   * `size` a strict server rejects.
   *
   * `inline` only when `inlineCids` says the body displays that Content-ID: Fastmail fails
   * the whole create with invalidProperties:["htmlBody"] on an inline part with no html
   * body. The Content-ID is kept either way, so a later edit can add the html.
   */
  async uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options: UploadAttachmentsOptions = {},
  ): Promise<AttachmentPart[]> {
    // Two passes, so a validation failure anywhere orphans NO blobs (a mid-upload network
    // failure still can; Fastmail garbage-collects them). Entries stay in SPEC ORDER: the
    // compose paths map parts back onto specs by position.
    const inlineCids = options.inlineCids;
    type PreparedFile = { kind: 'file'; handle: FileHandle; size: number; contentType: string; name: string; file: string; cid?: string };
    type PreparedRef = { kind: 'ref'; blobId: string; type: string; name: string; cid?: string };
    const prepared: (PreparedFile | PreparedRef)[] = [];
    try {
      let totalBytes = 0;
      for (let i = 0; i < specs.length; i++) {
        const spec = specs[i];
        const callerType = spec.contentType ? validateContentType(spec.contentType, i) : undefined;

        if (spec.blobId !== undefined) {
          assertBlobAttachEnabled(allowBlobAttach, i, 'blobId');
          prepared.push({
            kind: 'ref',
            blobId: spec.blobId,
            // A blob has no metadata to ask for. Not probed either: an unknown blobId fails
            // the Email/set loudly.
            type: callerType ?? guessContentType(spec.name as string),
            name: spec.name as string,
            cid: spec.cid,
          });
          continue;
        }

        if (spec.emailId !== undefined) {
          assertBlobAttachEnabled(allowBlobAttach, i, 'emailId + attachmentId');
          // Adds the item index; wrapped by CLASS so a transport failure is not relabelled.
          let info: AttachmentInfo;
          try {
            info = await this.getAttachmentInfo(spec.emailId, spec.attachmentId as string);
          } catch (e) {
            if (e instanceof InvalidInputError) {
              throw new InvalidInputError(
                `attachments[${i}] names emailId "${describeUntrusted(spec.emailId)}" and attachmentId ` +
                `"${describeUntrusted(spec.attachmentId as string)}". ${e.message}`
              );
            }
            throw e;
          }
          // Refused ONLY on the way out, where a shifted position mails the wrong file.
          // Decided from what the resolver did, never from whether the string looks numeric.
          if (info.matchedBy === 'index') {
            throw new InvalidInputError(
              `attachments[${i}] resolved attachmentId "${describeUntrusted(spec.attachmentId as string)}" only as an entry number in the part listing. ` +
              'An entry number moves whenever the listing does, and this attaches the file to mail you may then send. ' +
              'Pass the partId or blobId from get_email_attachments instead.'
            );
          }
          prepared.push({
            kind: 'ref',
            blobId: info.blobId,
            type: callerType ?? info.type,
            name: spec.name ?? info.name,
            cid: spec.cid,
          });
          continue;
        }

        // Not trusted to coerceAttachments: a future source falling through here would
        // attach the attach root itself, a file the caller never named.
        if (typeof spec.path !== 'string') {
          throw new InvalidInputError(
            `attachments[${i}] names no source this server can attach. ` +
            "Give exactly one of: 'path', 'blobId', or 'emailId' + 'attachmentId'."
          );
        }

        if (!attachDir) {
          throw new PathAccessError(
            'Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR to the directory attachable files live in, then restart the server to enable it.'
          );
        }
        const path = spec.path;
        const contentType = callerType ?? guessContentType(path);
        const { handle, size } = await JmapClient.safeReadPath(path, attachDir);
        // Push BEFORE the size checks so the finally closes this handle even if a cap throws.
        prepared.push({ kind: 'file', handle, size, contentType, name: spec.name ?? basename(path), file: basename(path), cid: spec.cid });
        if (size > JmapClient.MAX_ATTACHMENT_BYTES) {
          throw new PathAccessError(
            // basename is still caller text, so it is echoed like the rest.
            `attachments[${i}] ("${echoPath(basename(path))}") is ${size} bytes, over the ${JmapClient.MAX_ATTACHMENT_BYTES}-byte per-file guard. Fastmail's own limit ultimately governs.`
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
      for (const o of prepared) {
        const disposition = o.cid && inlineCids?.has(o.cid) ? 'inline' : 'attachment';
        if (o.kind === 'ref') {
          parts.push({
            blobId: o.blobId,
            type: o.type,
            name: o.name,
            disposition,
            ...(o.cid && { cid: o.cid }),
          });
          continue;
        }
        // Bounded read, never read-then-check, which would buffer an oversize file first.
        // A read may return fewer bytes than asked, so loop; a file that ends early (it
        // shrank after the stat) is refused rather than uploaded truncated.
        const buffer = Buffer.alloc(o.size);
        let filled = 0;
        while (filled < o.size) {
          const { bytesRead } = await o.handle.read(buffer, filled, o.size - filled, filled);
          if (bytesRead === 0) {
            throw new PathAccessError(
              `The attachment "${echoPath(o.file)}" could be read for only ${filled} of its ${o.size} bytes; ` +
              'it changed while being read. Nothing was uploaded for it. Try again once the file is complete.'
            );
          }
          filled += bytesRead;
        }

        const uploaded = await this.uploadBlob(buffer, o.contentType);
        parts.push({
          blobId: uploaded.blobId,
          type: uploaded.type,
          name: o.name,
          disposition,
          ...(o.cid && { cid: o.cid }),
        });
      }
      return parts;
    } finally {
      for (const o of prepared) if (o.kind === 'file') await o.handle.close().catch(() => {});
    }
  }

  /**
   * Fetch an attachment's bytes into memory, with the metadata from the same resolution.
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
    // Checked before the slow fetch, so a bad path fails fast.
    await JmapClient.safeWritePath(savePath, downloadDir);
    const { buffer, url } = await this.fetchAttachmentBuffer(emailId, attachmentId);

    // Re-validated after the fetch (a symlink may have been swapped in), then O_EXCL. On
    // EEXIST the rewrite ALSO uses 'wx': a default flag would reopen the symlink window the
    // unlink just created.
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

  // Shared engine for searchEmails + getEmails: the filter, the default Trash/Spam
  // exclusion, the visible query plus a hidden-count query, and QueryResult.exclusion.
  private async runFilteredQuery(opts: {
    base: any;
    conds: any[];
    exclusion: ExclusionResult;
    exclusionIntended: boolean;
    // Kept apart from exclusion.excludeIds all the way down: includeTrash/includeSpam
    // cannot recover these, and the hidden count depends on telling them apart.
    callerExcludeIds?: string[];
    limit: number;
    ascending: boolean;
    mailboxes: any[];
    position?: number;
  }): Promise<QueryResult> {
    const session = await this.getSession();
    const { base, conds, exclusion, exclusionIntended, limit, ascending, mailboxes } = opts;
    const position = opts.position ?? 0;
    const callerExcludeIds = opts.callerExcludeIds ?? [];

    // "The DEFAULT exclusion is active": drives the hidden count and the note, NOTHING
    // else. Gating the filter on it would drop the caller's excludes (fail-open).
    const doExclude = exclusion.excludeIds.length > 0;
    // ONE union array, assigned UNGATED (docs/conventions.md, "Scoping a query").
    // Into `base` BEFORE baseEmpty, or an exclusion-only query drops the exclusion.
    const allExcludeIds = [...exclusion.excludeIds, ...callerExcludeIds];
    if (allExcludeIds.length > 0) base.inMailboxOtherThan = allExcludeIds;

    // One condition per keyword: a FilterCondition allows only one hasKeyword/notKeyword.
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
    // Paging offset (#51); the hidden count is not a page and takes none.
    if (position > 0) visibleQuery.position = position;

    const methodCalls: [string, any, string][] = [
      ['Email/query', visibleQuery, 'query'],
      ['Email/get', emailGetParams, 'emails'],
    ];

    if (doExclude) {
      // hidden = broaderTotal - visibleTotal, the broader query dropping ONLY the DEFAULT
      // ids (the caller's excludes stay, since includeTrash/includeSpam cannot reveal
      // those). Rebuilt from a COPY of base through the same combine: deleting the key from
      // the assembled filter no-ops under the AND-wrap (fail-open). Same makeRequest, so
      // one atomic snapshot.
      const countBase = { ...base };
      if (callerExcludeIds.length > 0) countBase.inMailboxOtherThan = callerExcludeIds;
      else delete countBase.inMailboxOtherThan;
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

    // Whenever INTENDED, even with no role resolved, so the "not excluded" note fires.
    if (exclusionIntended) {
      let hidden: number | null = 0;
      if (doExclude) {
        // FAIL-CLOSED: "no note => nothing hidden" is a contract, so a missing, errored or
        // negative count is null (degraded note), never clamped to 0.
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
    requiredMailboxes?: string[];
    excludeMailboxes?: string[];
    after?: string;
    before?: string;
    limit?: number;
    position?: number;
    ascending?: boolean;
    excludeDrafts?: boolean;
    includeTrash?: boolean;
    includeSpam?: boolean;
  }): Promise<QueryResult> {
    // Before any network work, so a bad value fails naming its argument (#70).
    const after = coerceUtcDate(filters.after, 'after');
    const before = coerceUtcDate(filters.before, 'before');

    const mailboxes = await this.getMailboxes();
    const resolvedMailboxId = await this.resolveMailboxId(filters.mailbox, mailboxes);
    // Scope arrays (#26) share the label resolver; duplicates collapse here, since nothing
    // downstream collapses them.
    const resolveAll = async (inputs?: string[]) =>
      inputs && inputs.length > 0
        ? [...new Set(await this.resolveMailboxIdList(inputs, mailboxes))]
        : [];
    const requiredMailboxIds = await resolveAll(filters.requiredMailboxes);
    const callerExcludeIds = await resolveAll(filters.excludeMailboxes);

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
    // inMailbox is SINGULAR (RFC 8621 §4.4.1): N AND-ed conditions, not an array.
    for (const id of requiredMailboxIds) conds.push({ inMailbox: id });

    // Caller excludes are NOT an explicit scope. getEmails carries its own copy of this
    // expression; keep the two in step (docs/conventions.md, "Scoping a query").
    const hasExplicitScope = !!resolvedMailboxId || requiredMailboxIds.length > 0;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: filters.includeTrash,
      includeSpam: filters.includeSpam,
      hasExplicitScope,
      callerExcludedIds: callerExcludeIds,
    });
    const exclusionIntended = !hasExplicitScope && (!filters.includeTrash || !filters.includeSpam);

    return this.runFilteredQuery({
      base,
      conds,
      exclusion,
      exclusionIntended,
      callerExcludeIds,
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

    // threadId may be an email id; resolve it to its thread.
    let actualThreadId = threadId;

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

    // includeBodies (#74) reuses VERBOSE rather than a third property list. HTML values are
    // deliberately NOT fetched: a thread multiplies body size by the message count.
    if (includeBodies) {
      emailGetParams.properties = [...EMAIL_PROPERTIES_VERBOSE];
      emailGetParams.bodyProperties = [...EMAIL_BODY_PROPERTIES];
      emailGetParams.fetchTextBodyValues = true;
    } else {
      emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];
    }

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

    if (threadResult.notFound && threadResult.notFound.includes(actualThreadId)) {
      throw new InvalidInputError(`Thread with ID "${describeUntrusted(actualThreadId)}" not found`);
    }

    // Onto the FULL list, before the draft filter.
    const emails = this.getListResult(response, 1);
    const threadMailboxes = this.readListResultIfPresent(response, 2);
    attachMailboxInfo(emails, buildMailboxInfoMap(threadMailboxes));

    // Drafts hidden by default, by the $draft keyword (it survives a move out of Drafts),
    // but COUNTED so the handler can announce them: an agent replying must not miss that
    // a draft reply already exists.
    if (includeDrafts) {
      return { emails, hiddenDraftCount: 0 };
    }
    const filtered = emails.filter((e: any) => !e.keywords?.$draft);
    // A draft ONLY in Trash is not counted: every edit_draft leaves one there. With no
    // resolvable trash role every draft counts, failing toward over-warning.
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
    // Reads the stat fields off getMailboxes, which must therefore NOT be narrowed.
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
      // A real id absent from the fetched list throws: accepted residual
      // (docs/security-model.md).
      const mb = resolveMailbox(mailboxes, mailbox);
      return toStats(mb);
    }
    return mailboxes.map(toStats);
  }

  async getAccountSummary(): Promise<any> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const identities = await this.getIdentities();

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

    const updates: Record<string, any> = newUpdateMap();
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
    
    // Same accounting as bulkAddLabels.
    const submittedIds = Object.keys(updates);
    const notUpdated = JmapClient.withUnaccountedFailures(submittedIds, result.updated, result.notUpdated);
    if (Object.keys(notUpdated).length > 0) {
      const successCount = JmapClient.countAcknowledged(submittedIds, result.updated, notUpdated);
      this.throwBulkSetError(notUpdated, submittedIds.length, successCount, `mark as ${read ? 'read' : 'unread'}`, buildIdCollapseNote(emailIds));
    }
  }

  async bulkPinEmails(emailIds: string[], pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = newUpdateMap();
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

    // Same accounting as bulkAddLabels.
    const submittedIds = Object.keys(updates);
    const notUpdated = JmapClient.withUnaccountedFailures(submittedIds, result.updated, result.notUpdated);
    if (Object.keys(notUpdated).length > 0) {
      const successCount = JmapClient.countAcknowledged(submittedIds, result.updated, notUpdated);
      this.throwBulkSetError(notUpdated, submittedIds.length, successCount, `${pinned ? 'pin' : 'unpin'}`, buildIdCollapseNote(emailIds));
    }
  }

  async bulkMove(emailIds: string[], target: string): Promise<void> {
    const session = await this.getSession();

    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // Whole-value, as in moveEmail.
    const updates: Record<string, any> = newUpdateMap();
    emailIds.forEach(id => {
      updates[id] = { mailboxIds: { [targetMailboxId]: true } };
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

    // Same accounting as bulkAddLabels.
    const submittedIds = Object.keys(updates);
    const notUpdated = JmapClient.withUnaccountedFailures(submittedIds, result.updated, result.notUpdated);
    if (Object.keys(notUpdated).length > 0) {
      const successCount = JmapClient.countAcknowledged(submittedIds, result.updated, notUpdated);
      this.throwBulkSetError(notUpdated, submittedIds.length, successCount, 'move', buildIdCollapseNote(emailIds));
    }
  }

  async bulkDelete(emailIds: string[]): Promise<void> {
    const session = await this.getSession();

    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findByExactRole(mailboxes, 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const updates: Record<string, any> = newUpdateMap();
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

    // Same accounting as bulkAddLabels.
    const submittedIds = Object.keys(updates);
    const notUpdated = JmapClient.withUnaccountedFailures(submittedIds, result.updated, result.notUpdated);
    if (Object.keys(notUpdated).length > 0) {
      const successCount = JmapClient.countAcknowledged(submittedIds, result.updated, notUpdated);
      this.throwBulkSetError(notUpdated, submittedIds.length, successCount, 'delete', buildIdCollapseNote(emailIds));
    }
  }
}