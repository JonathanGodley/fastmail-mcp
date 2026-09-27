import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceAttachments, coerceBool } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs } from './body-format.js';
import type { AttachmentPart, UpdateDraftResult, UploadAttachmentsOptions } from './jmap-client.js';

/** The client surface editDraft needs; JmapClient satisfies it structurally. */
export interface EditDraftClient {
  uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options?: UploadAttachmentsOptions,
  ): Promise<AttachmentPart[]>;
  updateDraft(
    emailId: string,
    updates: {
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
      expandSignature?: boolean;
      bodyHash?: string;
    },
    options?: { attachmentsEnabled?: boolean },
  ): Promise<UpdateDraftResult>;
}

/**
 * Orchestrate edit_draft: coerce the caller's fields, upload/resolve any attachments, then
 * apply the edit.
 *
 * attachDir and allowBlobAttach are resolved by the caller, so this reads no environment.
 * Together they say which attachment sources are accepted, and so whether a refusal may
 * suggest supplying a missing embedded image through `attachments` at all.
 *
 * Unlike the compose paths, NO inline decision is made here: an edit's shipping body is only
 * settled inside updateDraft (the caller's html merged with the draft's own), so that is
 * where a supplied Content-ID is matched against the body and marked.
 */
export async function editDraft(
  args: any,
  client: EditDraftClient,
  attachDir: string | undefined,
  allowBlobAttach: boolean,
): Promise<UpdateDraftResult> {
  const a = args ?? {};
  const { emailId, from, subject, textBody, htmlBody, bodyHash } = a;
  const { to, cc, bcc, replyTo } = coerceRecipients(a);
  const clearFields = coerceStringArray(a.clearFields);
  const removeAttachments = coerceStringArray(a.removeAttachments);
  // `=== true` fails closed: a non-bool like "garbage" reads as false, which stores the
  // body exactly as written rather than rewriting it unasked.
  //
  // READ OFF `args?.` RATHER THAN THE `a` ALIAS, and leave it that way. The lenient-boolean
  // guard in tool-schema.test.ts matches `!!expandSignature` and `!!args?.expandSignature`
  // but not `!!a.expandSignature`, so tidying this back to the alias would put any future
  // bare-`!!` read of it outside the only check that looks for one.
  const expandSignature = coerceBool(args?.expandSignature) === true;
  if (!emailId) {
    throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
  }

  // Ordering belt only: updateDraft runs the same check and is authoritative. Here it
  // refuses a malformed body before any attachment is uploaded. It stays ABOVE the
  // attachment coercion, as in draft_email, so both tools report the same first error on
  // identical input.
  assertBodyInputs(a);

  const specs = coerceAttachments(a.attachments);
  const attachments = specs?.length
    ? await client.uploadAttachments(specs, attachDir, allowBlobAttach)
    : undefined;

  return client.updateDraft(emailId, {
    to,
    cc,
    bcc,
    from,
    subject,
    textBody,
    htmlBody,
    replyTo,
    clearFields,
    attachments,
    removeAttachments,
    expandSignature,
    // Passed through UNVALIDATED on purpose: updateDraft checks it behind the body-shape
    // guards, and a presence check here would jump that refusal order.
    bodyHash,
  }, {
    attachmentsEnabled: !!attachDir || allowBlobAttach,
  });
}
