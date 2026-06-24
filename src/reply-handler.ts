import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceBool } from './coerce.js';
import { isBlank } from './body-format.js';
import { buildReplyBodies } from './reply-quote.js';

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
  const shouldSend = coerceBool(send) ?? true;
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
