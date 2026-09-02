import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceAttachments, coerceBool } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs } from './body-format.js';
import type { AttachmentPart, UpdateDraftResult, UploadAttachmentsOptions } from './jmap-client.js';

/**
 * The minimal client surface editDraft needs; JmapClient satisfies it structurally.
 * Declared here (rather than importing JmapClient) so the orchestration stays unit-testable
 * with a mock, matching ComposeClient / ReplyClient / ForwardClient.
 */
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
 * Extracted from the index tool handler so the attachment seam — which decides whether a
 * capability gate refuses the call — is exercised by unit tests with a mock client rather
 * than only by running the server against a live account. The handler above it is a thin
 * result-to-text wrapper.
 *
 * attachDir and allowBlobAttach are passed in (resolved by the caller) so this function
 * reads no environment itself; between them they say which attachment SOURCES this server
 * will accept, and their combination is also what decides whether a refusal may suggest
 * supplying a missing embedded image through `attachments` at all.
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
  // Coerce to a real boolean (lenient clients send "true"/"false"). The explicit
  // `=== true` is the fail-closed direction: a non-bool like "garbage" coerces to
  // undefined and therefore reads as FALSE, which stores the body exactly as written.
  // Guessing the other way would have this server rewrite a body nobody asked it to
  // touch.
  //
  // READ OFF `args?.` RATHER THAN THE `a` ALIAS, and leave it that way. The lenient-boolean
  // guard in tool-schema.test.ts matches `!!expandSignature` and `!!args?.expandSignature`
  // but not `!!a.expandSignature`, so tidying this back to the alias would put any future
  // bare-`!!` read of it outside the only check that looks for one.
  const expandSignature = coerceBool(args?.expandSignature) === true;
  if (!emailId) {
    throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
  }

  // Ordering belt only: updateDraft runs the same check and remains the authoritative
  // seam for this tool. Running it here too means a malformed body is refused before
  // any attachment is read off disk and pushed to the blob store, matching the compose
  // path (the guard is a pure, idempotent input check). It stays ABOVE the attachment
  // coercion, as draft_email has it: a body defect is the caller's first problem to fix,
  // and reporting an attachments key ahead of it would make this tool disagree with
  // draft_email on identical input.
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
    // Passed through UNVALIDATED on purpose: updateDraft is where the hash is required
    // and checked, so it keeps its place in that method's refusal order — behind the
    // body-shape coupling guards, which name the shape the caller has to fix before a
    // stale-read complaint is any use to it. A presence check here would jump the queue.
    bodyHash,
  }, {
    attachmentsEnabled: !!attachDir || allowBlobAttach,
  });
}
