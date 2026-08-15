import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { isBlank, assertBodyInputs } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import { buildForwardBodies } from './reply-quote.js';
import type { AttachmentPart } from './jmap-client.js';

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
function carryPart(a: any): AttachmentPart {
  const part: AttachmentPart = { blobId: a.blobId, type: a.type };
  if (a.name != null) part.name = a.name;
  if (a.disposition != null) part.disposition = a.disposition;
  if (a.cid != null) part.cid = a.cid;
  return part;
}

export interface BuiltForward {
  asAttachment: boolean;
  forwardParams: ForwardParams;
  // Count of the original's true-inline (cid-referenced) image parts NOT carried by
  // an inline forward — the loud runtime degrade for the inline-image gap. Computed
  // from the SOURCE message regardless of includeOriginalAttachments (the body's
  // cid <img>s are stripped by the sanitizer either way). Absent on asAttachment
  // forwards: the .eml embeds those parts losslessly, nothing is dropped.
  droppedInlineImages?: number;
}

// Assemble the forward parameters from the caller's args and the already-fetched
// original email. Pure (no I/O) so the forward_email orchestration is unit-testable:
// coerce inputs, require an explicit recipient, build the Fwd: subject, record the
// original's Message-ID, assemble the forwarded-message bodies, and map the carried
// attachments (all-or-none; per-part subsetting goes through the draft +
// edit_draft's removeAttachments). Caller uploads are appended later by
// composeForward (the I/O boundary). Throws McpError on invalid input.
export function buildForwardParams(args: any, originalEmail: any): BuiltForward {
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

  // A part without a blobId can't be re-referenced on a new Email at all (JMAP carry
  // works by blobId), so it is skipped without a count — unlike the inline drop below,
  // there is no alternative to point the caller at, and RFC 8621 attachments are
  // blob-backed, so the case is essentially theoretical.
  const sourceAttachments: any[] = Array.isArray(originalEmail?.attachments)
    ? originalEmail.attachments.filter((p: any) => p && p.blobId)
    : [];
  const isTrueInline = (p: any) => p.disposition === 'inline' && p.cid;

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
    if (!originalEmail?.blobId) {
      throw new McpError(ErrorCode.InternalError, 'Original email has no blobId; cannot attach it as .eml');
    }
    params.attachments = [{
      blobId: originalEmail.blobId,
      type: 'message/rfc822',
      name: sanitizeEmlFilename(originalEmail?.subject),
      disposition: 'attachment',
    }];
    return { asAttachment, forwardParams: params };
  }

  // Inline forward: reproduce the original under the forwarded-message block.
  const bodies = buildForwardBodies({ original: originalEmail, textBody, htmlBody });
  params.textBody = bodies.textBody;
  params.htmlBody = bodies.htmlBody;

  // True-inline (cid) image parts are NOT carried — their <img> references are
  // stripped from the forwarded html by the sanitizer, so carrying the bytes would
  // attach orphans. Counted so the degrade is loud; asAttachment is the lossless
  // alternative. Parts marked inline WITHOUT a cid are content (some clients mark
  // regular attachments inline); carry them normalized to disposition:'attachment',
  // since nothing can reference a part with no Content-ID and an inline marking no
  // body backs would misdescribe what the draft carries.
  const droppedInlineImages = sourceAttachments.filter(isTrueInline).length;
  if (includeOriginalAttachments) {
    params.attachments = sourceAttachments
      .filter((p) => !isTrueInline(p))
      .map((p) => {
        const part = carryPart(p);
        if (part.disposition === 'inline') part.disposition = 'attachment';
        return part;
      });
  }

  return { asAttachment, forwardParams: params, droppedInlineImages };
}

// The minimal client surface composeForward needs; JmapClient satisfies it
// structurally. Declared here (like ReplyClient) so the orchestration stays
// unit-testable with a mock and free of a hard dependency on the concrete client.
export interface ForwardClient {
  getEmailById(id: string): Promise<any>;
  uploadAttachments(specs: AttachmentSpec[], attachDir: string | undefined): Promise<AttachmentPart[]>;
  createDraft(params: ForwardParams): Promise<string>;
}

export interface ComposeForwardResult {
  subject: string;
  emailId: string;
  droppedInlineImages?: number; // see BuiltForward
}

// Orchestrate a forward end to end: fetch the original, assemble the (pure) forward
// params, upload any NEW attachments and append them behind the carried/.eml parts,
// then save the forward as a draft. This tool only ever drafts — send_draft is the
// single tool that transmits mail, and it also does the thread-state maintenance
// (marking the original forwarded + read), resolved from the draft's recorded
// X-Forwarded-Message-Id header (#60; see docs/conventions.md "Draft provenance"),
// recorded on both inline and asAttachment forwards.
// Mirrors composeReply so the index handler stays a thin wrapper. attachDir is passed
// in (resolved by the caller) so this reads no environment itself.
export async function composeForward(
  args: any,
  client: ForwardClient,
  attachDir: string | undefined,
): Promise<ComposeForwardResult> {
  const originalEmailId = args?.originalEmailId;
  if (!originalEmailId) {
    throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
  }

  const originalEmail = await client.getEmailById(originalEmailId);
  const { forwardParams, droppedInlineImages } = buildForwardParams(args, originalEmail);

  // Upload NEW attachments (if any) after the pure builder and append them behind
  // the carried/.eml parts.
  const specs = coerceAttachments(args?.attachments);
  if (specs?.length) {
    const uploaded = await client.uploadAttachments(specs, attachDir);
    forwardParams.attachments = [...(forwardParams.attachments ?? []), ...uploaded];
  }

  const emailId = await client.createDraft(forwardParams);
  return { subject: forwardParams.subject, emailId, droppedInlineImages };
}
