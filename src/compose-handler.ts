import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceAttachments } from './coerce.js';
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

// The minimal client surface this orchestration needs; JmapClient satisfies it
// structurally. Declared here (rather than importing JmapClient) so the handler stays
// unit-testable with a mock, matching ReplyClient / ForwardClient.
export interface ComposeClient {
  uploadAttachments(specs: AttachmentSpec[], attachDir: string | undefined): Promise<AttachmentPart[]>;
  createDraft(params: DraftParams): Promise<string>;
}

export interface ComposeDraftResult {
  emailId: string;
  subject?: string;
  to?: string[];
  cc?: string[];
}

// Orchestrate create_draft: validate the caller's bodies, coerce the recipient lists,
// upload any attachments, then create the draft. Extracted from the index tool handler so
// the validation and the attachment seam are unit-testable with a mock client (the
// handler is a thin result-to-text wrapper). attachDir is passed in so this function
// reads no environment itself. Transmission is a separate step: send_draft is the only
// tool that submits mail.
//
// The body guard belongs HERE rather than in JmapClient.createDraft: that method is
// shared with reply_email / forward_email, where it receives the MERGED body (the
// caller's text plus the quoted or forwarded original). Validating there would reject
// legitimate mail whose quoted original happens to contain, say, a CDATA token.
export async function composeDraft(
  args: any,
  client: ComposeClient,
  attachDir: string | undefined,
): Promise<ComposeDraftResult> {
  const a = args ?? {};
  assertBodyInputs(a);

  const { from, mailbox, subject, textBody, htmlBody } = a;
  const { to, cc, bcc, replyTo } = coerceRecipients(a);

  // The threading headers are string[] to createDraft, which maps over them. A lenient
  // client can send either as a bare or JSON-encoded string, which would reach JMAP as a
  // per-character header list (or crash the map). coerceStringArray, matching how the
  // recipient lists above are handled. (#54)
  const inReplyTo = coerceStringArray(a.inReplyTo);
  const references = coerceStringArray(a.references);

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
