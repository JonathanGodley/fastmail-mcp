import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceAttachments } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs } from './body-format.js';
import type { AttachmentPart } from './jmap-client.js';

// Parameters passed to createDraft for a freshly composed message (matches its input
// shape; everything is optional because an attachment-only or subject-only draft is
// legitimate).
export interface DraftParams {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  from?: string;
  mailbox?: string;
  subject?: string;
  textBody?: string;
  htmlBody?: string;
  inReplyTo?: string[];
  references?: string[];
  replyTo?: string[];
  attachments?: AttachmentPart[];
}

// Sending additionally requires a recipient and a subject.
export interface SendParams extends DraftParams {
  to: string[];
  subject: string;
}

// The minimal client surface these orchestrations need; JmapClient satisfies it
// structurally. Declared here (rather than importing JmapClient) so the handlers stay
// unit-testable with a mock, matching ReplyClient / ForwardClient.
export interface ComposeClient {
  uploadAttachments(specs: AttachmentSpec[], attachDir: string | undefined): Promise<AttachmentPart[]>;
  createDraft(params: DraftParams): Promise<string>;
  sendEmail(params: SendParams): Promise<string>;
}

export interface ComposeDraftResult {
  emailId: string;
  subject?: string;
  to?: string[];
  cc?: string[];
}

// Orchestrate send_email: validate the caller's bodies and required fields, coerce the
// recipient lists, upload any attachments, then transmit. Extracted from the index tool
// handler so the validation and the attachment seam are unit-testable with a mock client
// (the handler is now a thin submissionId-to-text wrapper). attachDir is passed in so
// this function reads no environment itself.
//
// The body guard belongs HERE rather than in JmapClient.sendEmail: that method is shared
// with reply_email / forward_email, where it receives the MERGED body (the caller's text
// plus the quoted or forwarded original). Validating there would reject legitimate mail
// whose quoted original happens to contain, say, a CDATA token.
export async function composeSend(
  args: any,
  client: ComposeClient,
  attachDir: string | undefined,
): Promise<string> {
  const a = args ?? {};
  assertBodyInputs(a);

  const { from, mailbox, subject, textBody, htmlBody, inReplyTo, references } = a;
  const { to, cc, bcc, replyTo } = coerceRecipients(a);

  if (!to || to.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, 'to field is required and must be a non-empty array');
  }
  if (!subject) {
    throw new McpError(ErrorCode.InvalidParams, 'subject is required');
  }
  if (!textBody && !htmlBody) {
    throw new McpError(ErrorCode.InvalidParams, 'Either textBody or htmlBody is required');
  }

  const specs = coerceAttachments(a.attachments);
  const attachments = specs?.length ? await client.uploadAttachments(specs, attachDir) : undefined;

  return client.sendEmail({
    to,
    cc,
    bcc,
    from,
    mailbox,
    subject,
    textBody,
    htmlBody,
    inReplyTo,
    references,
    replyTo,
    attachments,
  });
}

// Orchestrate create_draft. Same shape as composeSend; returns the fields the handler
// needs for its summary line so the handler stays a thin wrapper.
export async function composeDraft(
  args: any,
  client: ComposeClient,
  attachDir: string | undefined,
): Promise<ComposeDraftResult> {
  const a = args ?? {};
  assertBodyInputs(a);

  const { from, mailbox, subject, textBody, htmlBody, inReplyTo, references } = a;
  const { to, cc, bcc, replyTo } = coerceRecipients(a);

  // Coerce attachments BEFORE the contentless guard so an attachment-only draft
  // (a legitimate "stash this file" artifact, consistent with edit_draft accepting
  // a body-less attachment edit) counts as content. A lenient client may send a
  // JSON-string array, so test the coerced specs, not the raw arg.
  const specs = coerceAttachments(a.attachments);

  if (!to?.length && !subject && !textBody && !htmlBody && !specs?.length) {
    throw new McpError(ErrorCode.InvalidParams, 'At least one of to, subject, textBody, htmlBody, or attachments must be provided');
  }

  const attachments = specs?.length ? await client.uploadAttachments(specs, attachDir) : undefined;

  const emailId = await client.createDraft({
    to,
    cc,
    bcc,
    from,
    mailbox,
    subject,
    textBody,
    htmlBody,
    inReplyTo,
    references,
    replyTo,
    attachments,
  });

  return { emailId, subject, to, cc };
}
