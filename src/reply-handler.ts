import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { isBlank } from './body-format.js';
import { buildReplyBodies } from './reply-quote.js';
import type { AttachmentPart } from './jmap-client.js';

// Parameters passed to createDraft/sendEmail for a reply (matches their input shapes).
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
  // Set by the reply_email handler AFTER buildReplyParams (which stays pure / no I/O);
  // threaded into both the send and save-as-draft branches.
  attachments?: AttachmentPart[];
}

// Assemble the reply parameters from the caller's args and the already-fetched original
// email. Pure (no I/O), so the reply_email handler's logic is unit-testable without
// spinning the server: coerce inputs, default quoteOriginal to true, reject a body-less
// send (trim/zero-width-aware), build the threading headers and the "Re:" subject, default
// the recipient to the original sender, and append the attributed quote. The caller-supplied
// bodies are returned as-is when quoteOriginal is false. createDraft/sendEmail add the auto
// text/plain fallback downstream for an html-only reply. Throws McpError on invalid input.
export function buildReplyParams(
  args: any,
  originalEmail: any,
): { shouldSend: boolean; quoteOriginal: boolean; replyParams: ReplyParams } {
  const a = args ?? {};
  const { from, textBody, htmlBody, send } = a;
  const { to: toArray, cc, bcc, replyTo } = coerceRecipients(a);
  const shouldSend = coerceBool(send) ?? false;
  const quoteOriginal = coerceBool(a.quoteOriginal) ?? true;

  // Trim/zero-width-aware so a whitespace-only htmlBody can't slip through and produce a
  // "   " + quote reply; a body-less reply flows to the same no-body handling elsewhere.
  if (shouldSend && isBlank(textBody) && isBlank(htmlBody)) {
    throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
  }

  const originalMessageId = originalEmail?.messageId?.[0];
  if (!originalMessageId) {
    throw new McpError(ErrorCode.InternalError, 'Original email does not have a Message-ID; cannot thread reply');
  }
  const inReplyTo = [originalMessageId];
  const references = [...(originalEmail.references || []), originalMessageId];

  let subject = originalEmail.subject || '';
  if (!/^Re:/i.test(subject)) {
    subject = `Re: ${subject}`;
  }

  const to = (toArray && toArray.length > 0)
    ? toArray
    : (Array.isArray(originalEmail.from) ? originalEmail.from.map((addr: any) => addr.email).filter(Boolean) : []);
  if (to.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, 'Could not determine reply recipient. Please provide "to" explicitly.');
  }

  const quoted = buildReplyBodies({ original: originalEmail, textBody, htmlBody, quoteOriginal });

  return {
    shouldSend,
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
  sendEmail(params: ReplyParams): Promise<string>;
}

export interface ComposeReplyResult {
  sent: boolean;
  subject: string;
  emailId?: string;      // set on the draft (send=false) branch
  submissionId?: string; // set on the send branch
}

// Orchestrate a reply end to end: fetch the original, assemble the (pure) reply params,
// upload any attachments and thread them into whichever branch runs, then create the
// draft or send. Extracted from the index tool handler so the attachment-threading seam
// (the one piece that touches I/O via the injected client) is unit-testable without the
// MCP server or a live account — the handler is now a thin wrapper over this. attachDir
// is passed in (resolved by the caller) so this function reads no environment itself.
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
  const { shouldSend, replyParams } = buildReplyParams(args, originalEmail);

  // Upload attachments (if any) after the pure builder, then thread the parts into
  // whichever branch runs (send or save-as-draft).
  const specs = coerceAttachments(args?.attachments);
  if (specs?.length) {
    replyParams.attachments = await client.uploadAttachments(specs, attachDir);
  }

  if (!shouldSend) {
    const emailId = await client.createDraft(replyParams);
    return { sent: false, subject: replyParams.subject, emailId };
  }
  const submissionId = await client.sendEmail(replyParams);
  return { sent: true, subject: replyParams.subject, submissionId };
}
