import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceAttachments, coerceBool } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs, isBlank } from './body-format.js';
import { planAuthoredInlineImages, reportAuthoredInlineImages } from './compose-inline.js';
import { applySignature, noteSignatureNotAppended, signatureSkipReason } from './reply-quote.js';
import { selectIdentity, signatureOf } from './identity.js';
import type { AttachmentPart, UploadAttachmentsOptions } from './jmap-client.js';

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
  /** The sending identities, for the signature lookup (#33). Called only when one is asked for. */
  getIdentities(): Promise<any[]>;
  uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options?: UploadAttachmentsOptions,
  ): Promise<AttachmentPart[]>;
  createDraft(params: DraftParams): Promise<string>;
  // Only for the post-save confirmation read; this tool needs no fetch to compose.
  getEmailById(id: string): Promise<any>;
}

export interface ComposeDraftResult {
  emailId: string;
  subject?: string;
  to?: string[];
  cc?: string[];
  /** What the draft ended up embedding, or could not (#13). Absent when there is nothing to say. */
  notes?: string[];
}

// Orchestrate create_draft: validate the caller's bodies, coerce the recipient lists,
// upload any attachments, then create the draft. Extracted from the index tool handler so
// the validation and the attachment seam are unit-testable with a mock client (the
// handler is a thin result-to-text wrapper). attachDir and allowBlobAttach are passed in
// so this function reads no environment itself; between them they say which attachment
// SOURCES this server will accept (a local file, and content already in the account). Transmission is a separate step: send_draft is the only
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
  allowBlobAttach: boolean,
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

  // Embedded-image checks run here, before a single byte is read off disk, so a refused
  // call leaves nothing behind in the blob store — the same ordering the malformed-body
  // guard above already has.
  const plan = planAuthoredInlineImages({
    callerHtml: htmlBody,
    htmlShips: !isBlank(htmlBody),
    specs,
    attachmentsEnabled: !!attachDir || allowBlobAttach,
    surface: 'compose',
  });

  // Off by default, and the lookup is skipped when it is off (see composeReply).
  //
  // Ordered BEFORE the upload below, for the reason the embedded-image plan is: a lookup
  // that fails after the upload would orphan freshly written blobs in the account with no
  // draft referencing them. It still runs AFTER the contentless guard above, which is what
  // keeps a signature from turning an empty call into a saveable draft here — note that
  // guard is this function's own, and createDraft's downstream "html body has no visible
  // content" reject DOES see the signature, so create_draft(htmlBody:'<p></p>',
  // appendSignature:true) saves as a signature-only draft rather than being refused. It is
  // also resolved after the embedded-image plan, which reads the caller's OWN html — the
  // signature is this server's markup and is not held against them.
  const appendSignature = coerceBool(a.appendSignature) === true;
  const identity = appendSignature ? selectIdentity(await client.getIdentities(), from) : undefined;
  const signature = appendSignature ? signatureOf(identity) : undefined;
  // Read before the append, because the reason for an empty append is only legible against
  // the bodies as the caller supplied them.
  const signatureSkip = appendSignature
    ? signatureSkipReason({ textBody, htmlBody }, signature)
    : undefined;

  const attachments = specs?.length
    ? await client.uploadAttachments(specs, attachDir, allowBlobAttach, { inlineCids: plan.inlineCids })
    : undefined;

  const signed = applySignature({ textBody, htmlBody }, signature);

  const emailId = await client.createDraft({
    to,
    cc,
    bcc,
    from,
    mailbox,
    subject,
    textBody: signed.textBody,
    htmlBody: signed.htmlBody,
    inReplyTo,
    references,
    replyTo,
    attachments,
  });

  const notes = [
    ...await reportAuthoredInlineImages({
      uploaded: attachments,
      plan,
      emailId,
      readBack: (id) => client.getEmailById(id),
    }),
    ...(signatureSkip ? [noteSignatureNotAppended(signatureSkip, identity?.email ?? from)] : []),
  ];

  return { emailId, subject, to, cc, ...(notes.length > 0 && { notes }) };
}
