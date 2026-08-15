import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs, isBlank } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import { buildReplyBodies, emptyQuoteImages } from './reply-quote.js';
import type { QuoteImageOutcome } from './reply-quote.js';
import { formatAddress } from './email-formatter.js';
import {
  DEFAULT_INLINE_CONTEXT, planAuthoredInlineImages, recordQuoteImages,
  reportAuthoredInlineImages,
} from './compose-inline.js';
import type { AuthoredInlineContext, AuthoredInlinePlan } from './compose-inline.js';
import { buildUnionParts, checkInlineClosure } from './inline-images.js';
import { InlineNoteLedger } from './inline-notes.js';
import type { AttachmentPart, UploadAttachmentsOptions } from './jmap-client.js';

// Parameters passed to createDraft for a reply (matches its input shape).
export interface ReplyParams {
  to: string[];
  cc?: string[];
  bcc?: string[];
  from?: string;
  subject: string;
  textBody?: string;
  htmlBody?: string;
  inReplyTo: string[];
  references: string[];
  replyTo?: string[];
  // The JMAP id of the exact stored instance being replied to, recorded on the draft
  // as X-Fastmail-MCP-Source-Id so send_draft can mark precisely that copy answered —
  // In-Reply-To alone names the message, which can exist as several stored copies.
  sourceEmailId?: string;
  // Set by the reply_email handler AFTER buildReplyParams (which stays pure / no I/O).
  attachments?: AttachmentPart[];
}

// Assemble the reply parameters from the caller's args and the already-fetched original
// email. Pure (no I/O), so the reply_email handler's logic is unit-testable without
// spinning the server: coerce inputs, default quoteOriginal to true, build the threading
// headers and the subject (the caller's override, else "Re: <original>"), default the
// recipient to the original sender, and append the attributed quote. The caller-supplied
// bodies are returned as-is when quoteOriginal is false; a body-less reply draft is
// allowed (it can be filled in via edit_draft before send_draft transmits it).
// createDraft adds the auto text/plain fallback downstream for an html-only reply.
// Throws McpError on invalid input.
//
// `inline` carries the caller's already-coerced attachments for the embedded-image checks
// (see AuthoredInlineContext for why they are passed in rather than read from args). The
// specs are used for validation only — the uploaded parts are threaded onto the result by
// the orchestration, which owns the I/O.
//
// `mint` is injected only so tests are deterministic; production leaves it unset and the
// quote builder draws from the CSPRNG.
export function buildReplyParams(
  args: any,
  originalEmail: any,
  inline: AuthoredInlineContext = DEFAULT_INLINE_CONTEXT,
  mint?: () => string,
): {
  quoteOriginal: boolean;
  replyParams: ReplyParams;
  inlinePlan: AuthoredInlinePlan;
  /** What the quote did with the original's embedded images (#13). */
  quoteImages: QuoteImageOutcome;
} {
  const a = args ?? {};
  // Validate the caller's own bodies FIRST — before the quote is appended below. Once the
  // quoted original is merged in, a malformed new message is masked by it: the quote
  // supplies the real tags an escaped-HTML body lacks, and it supplies the visible content
  // that would otherwise trip the no-readable-body reject, so a CDATA-wrapped reply ships
  // with its new message silently dropped from the plain-text part (#78).
  assertBodyInputs(a);
  const subjectOverride = coerceSubjectOverride(a.subject, 'Omit it to inherit "Re: <original subject>".');
  const { from, textBody, htmlBody } = a;
  const { to: toArray, cc, bcc, replyTo } = coerceRecipients(a);
  const quoteOriginal = coerceBool(a.quoteOriginal) ?? true;

  const originalMessageId = originalEmail?.messageId?.[0];
  if (!originalMessageId) {
    throw new McpError(ErrorCode.InternalError, 'Original email does not have a Message-ID; cannot thread reply');
  }
  const inReplyTo = [originalMessageId];
  const references = [...(originalEmail.references || []), originalMessageId];

  // The caller's override wins verbatim; otherwise inherit the original's subject with a
  // "Re: " prefix, without double-prefixing one that already has it. The threading headers
  // above are built the same way either way, so an overridden subject still carries a
  // correct In-Reply-To/References chain to the recipient (#68).
  let subject = subjectOverride ?? (originalEmail.subject || '');
  if (subjectOverride === undefined && !/^Re:/i.test(subject)) {
    subject = `Re: ${subject}`;
  }

  // Default the recipient to the original sender, preserving the display name via the
  // shared formatAddress ("Name <email>"). The result round-trips through parseAddress
  // downstream (split on the last <>), and this default array bypasses coerceStringArray
  // so a comma in the name is never re-split into a bogus second recipient (#31).
  const to = (toArray && toArray.length > 0)
    ? toArray
    : (Array.isArray(originalEmail.from) ? originalEmail.from.filter((a: any) => a?.email).map(formatAddress) : []);
  if (to.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, 'Could not determine reply recipient. Please provide "to" explicitly.');
  }

  // Checked against the caller's OWN note, BEFORE the quote is merged in below. The quote is
  // this server's own markup and carries its own identifiers; scanning the merged body would
  // hold those against the caller. An html note ships whenever the caller supplied one, so
  // that is also what decides whether their images can be displayed at all.
  const inlinePlan = planAuthoredInlineImages({
    callerHtml: htmlBody,
    htmlShips: !isBlank(htmlBody),
    specs: inline.specs,
    attachmentsEnabled: inline.attachmentsEnabled,
    surface: 'note',
  });

  // The quote resolves the original's embedded images against the original's OWN parts (the
  // gated union — an embedded image lands in a body list rather than in `attachments` for
  // some MIME shapes, so `attachments` alone would miss it). Carrying them is the default:
  // the images are part of the quoted body, and a quote that silently loses them shows the
  // recipient a conversation the sender never had. quoteOriginal: false drops the whole
  // quote, images included; there is no setting for quote text without its images.
  const quoted = buildReplyBodies({
    original: originalEmail,
    textBody,
    htmlBody,
    quoteOriginal,
    quoteImages: {
      sourceParts: buildUnionParts(originalEmail).map((u) => u.part),
      ...(mint && { mint }),
    },
  });

  return {
    quoteOriginal,
    inlinePlan,
    quoteImages: quoted.quoteImages ?? emptyQuoteImages(),
    replyParams: {
      to,
      cc,
      bcc,
      from,
      subject,
      textBody: quoted.textBody,
      htmlBody: quoted.htmlBody,
      inReplyTo,
      references,
      replyTo,
      // The caller named this exact instance; record it (createDraft vets the value).
      ...(typeof originalEmail?.id === 'string' && originalEmail.id !== '' && { sourceEmailId: originalEmail.id }),
    },
  };
}

// The minimal client surface composeReply needs; JmapClient satisfies it structurally.
// Declared here (rather than importing JmapClient) so the orchestration stays unit-
// testable with a mock and free of a hard dependency on the concrete client.
export interface ReplyClient {
  getEmailById(id: string): Promise<any>;
  uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options?: UploadAttachmentsOptions,
  ): Promise<AttachmentPart[]>;
  createDraft(params: ReplyParams): Promise<string>;
}

export interface ComposeReplyResult {
  subject: string;
  emailId: string;
  /** What the draft ended up embedding, or could not (#13). Absent when there is nothing to say. */
  notes?: string[];
}

// Orchestrate a reply end to end: fetch the original, assemble the (pure) reply params,
// upload any attachments, then save the reply as a draft. This tool only ever drafts —
// send_draft is the single tool that transmits mail, and it also does the thread-state
// maintenance (marking the original answered + read), targeting the exact instance
// recorded on the draft, with the In-Reply-To Message-ID as the fallback (#60; see
// docs/conventions.md "Draft provenance"). Extracted from
// the index tool handler so the attachment-threading seam (the one piece that touches
// I/O via the injected client) is unit-testable without the MCP server or a live
// account — the handler is a thin wrapper over this. attachDir and allowBlobAttach are
// passed in (resolved by the caller) so this function reads no environment itself; between
// them they say which attachment SOURCES this server will accept.
export async function composeReply(
  args: any,
  client: ReplyClient,
  attachDir: string | undefined,
  allowBlobAttach: boolean,
): Promise<ComposeReplyResult> {
  const originalEmailId = args?.originalEmailId;
  if (!originalEmailId) {
    throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
  }

  // Fetch the original, then assemble the reply (threading headers, Re: subject, recipient
  // defaulting, the attributed quote, body validation) via the pure, unit-tested builder.
  const originalEmail = await client.getEmailById(originalEmailId);

  // The caller's bodies are validated before their attachments are even coerced, so a
  // malformed body is reported as a body problem rather than as whatever the attachments
  // list happens to be wrong about too. The builder re-runs this check as its own first
  // step; running it here as well is an ordering belt, not a second authority.
  assertBodyInputs(args ?? {});
  const specs = coerceAttachments(args?.attachments);

  const { replyParams, inlinePlan, quoteImages } = buildReplyParams(args, originalEmail, {
    specs,
    attachmentsEnabled: !!attachDir || allowBlobAttach,
  });

  // Upload attachments (if any) after the pure builder, then thread the parts into
  // the draft.
  const uploaded = specs?.length
    ? await client.uploadAttachments(specs, attachDir, allowBlobAttach, { inlineCids: inlinePlan.inlineCids })
    : undefined;

  // The images the quote carries are APPENDED to the caller's own files, never assigned over
  // them: a reply that attaches a file and quotes a message with an embedded image ships
  // both. They ride even when this server cannot read files off disk at all — a carried
  // image re-references a blob the account already holds, so it needs no attachments
  // directory. The field stays unset when there is nothing to carry, so an ordinary reply's
  // parameters are unchanged.
  const ledger = new InlineNoteLedger();
  const carry = recordQuoteImages(ledger, quoteImages, 'reply');
  const attachments = [...(uploaded ?? []), ...carry.minted];
  if (attachments.length > 0) replyParams.attachments = attachments;

  // Checked BEFORE the draft is saved: a message that references an image it does not carry,
  // or carries one nothing references, is this server's own assembly error, and it is worth
  // more to refuse than to leave the caller a broken draft.
  checkInlineClosure({
    htmlBodies: [replyParams.htmlBody],
    finalPartCids: attachments.map((part) => part.cid),
    attachedMintedCids: carry.mintedCids,
  });

  const emailId = await client.createDraft(replyParams);
  const notes = [
    ...ledger.emit({ surface: 'reply', ...(carry.resolvedPartCount !== undefined && { resolvedPartCount: carry.resolvedPartCount }) }),
    ...await reportAuthoredInlineImages({
      // The caller's own uploads only: the quote's minted parts are reported by the ledger
      // above, and passing them here as well would count each embed twice. Their identifiers
      // still go in, so the confirmation read covers what the quote carries.
      uploaded,
      mintedCids: carry.mintedCids,
      plan: inlinePlan,
      emailId,
      readBack: (id) => client.getEmailById(id),
    }),
  ];
  return { subject: replyParams.subject, emailId, ...(notes.length > 0 && { notes }) };
}
