import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import { buildReplyBodies } from './reply-quote.js';
import { formatAddress } from './email-formatter.js';
import type { AttachmentPart } from './jmap-client.js';

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
export function buildReplyParams(
  args: any,
  originalEmail: any,
): { quoteOriginal: boolean; replyParams: ReplyParams } {
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

  const quoted = buildReplyBodies({ original: originalEmail, textBody, htmlBody, quoteOriginal });

  return {
    quoteOriginal,
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
  uploadAttachments(specs: AttachmentSpec[], attachDir: string | undefined): Promise<AttachmentPart[]>;
  createDraft(params: ReplyParams): Promise<string>;
}

export interface ComposeReplyResult {
  subject: string;
  emailId: string;
}

// Orchestrate a reply end to end: fetch the original, assemble the (pure) reply params,
// upload any attachments, then save the reply as a draft. This tool only ever drafts —
// send_draft is the single tool that transmits mail, and it also does the thread-state
// maintenance (marking the original answered + read), targeting the exact instance
// recorded on the draft, with the In-Reply-To Message-ID as the fallback (#60; see
// docs/conventions.md "Draft provenance"). Extracted from
// the index tool handler so the attachment-threading seam (the one piece that touches
// I/O via the injected client) is unit-testable without the MCP server or a live
// account — the handler is a thin wrapper over this. attachDir is passed in (resolved
// by the caller) so this function reads no environment itself.
export async function composeReply(
  args: any,
  client: ReplyClient,
  attachDir: string | undefined,
): Promise<ComposeReplyResult> {
  const originalEmailId = args?.originalEmailId;
  if (!originalEmailId) {
    throw new McpError(ErrorCode.InvalidParams, 'originalEmailId is required');
  }

  // Fetch the original, then assemble the reply (threading headers, Re: subject, recipient
  // defaulting, the attributed quote, body validation) via the pure, unit-tested builder.
  const originalEmail = await client.getEmailById(originalEmailId);
  const { replyParams } = buildReplyParams(args, originalEmail);

  // Upload attachments (if any) after the pure builder, then thread the parts into
  // the draft.
  const specs = coerceAttachments(args?.attachments);
  if (specs?.length) {
    replyParams.attachments = await client.uploadAttachments(specs, attachDir);
  }

  const emailId = await client.createDraft(replyParams);
  return { subject: replyParams.subject, emailId };
}
