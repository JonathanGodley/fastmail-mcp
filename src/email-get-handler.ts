import { attachDraftBodyHash } from './body-hash.js';
import { coerceBool, InvalidInputError, toolJson } from './coerce.js';
import { simplifyEmail } from './email-formatter.js';
import { parseEmailFields, projectEmail, wantsHtmlBody } from './field-projection.js';
import { assertStripQuotedNotRaw } from './quote-strip.js';
import type { JmapClient } from './jmap-client.js';
import type { ToolContent } from './contacts-handler.js';

/** The client surface get_email needs. `JmapClient` satisfies it structurally. */
export interface EmailGetClient {
  getEmailById: JmapClient['getEmailById'];
}

export async function getEmailTool(args: any, client: EmailGetClient): Promise<ToolContent> {
  const { emailId, stripQuoted } = args ?? {};
  if (!emailId) {
    throw new InvalidInputError('emailId is required');
  }
  // Same coercion as listEmailsTool. Here `!!` on raw:"false" would also make
  // assertStripQuotedNotRaw reject a legitimate stripQuoted read.
  const raw = coerceBool(args.raw, 'raw') ?? false;
  const verbose = coerceBool(args.verbose, 'verbose') ?? false;
  // Validated before the fetch so a typo'd field name costs no round trip.
  // Selecting bodyHtml implies verbose's includeHtml: without that, projecting a
  // field the simplifier never emitted would return {} — the trap the parameter
  // exists to avoid (#69).
  const fields = parseEmailFields(args.fields, { raw });
  // Rejected before the fetch: raw is unmodified JMAP, so honouring stripQuoted
  // there would be impossible and ignoring it would be silent (#73).
  const strip = coerceBool(stripQuoted, 'stripQuoted') ?? false;
  assertStripQuotedNotRaw(strip, raw);
  const email = await client.getEmailById(emailId);
  const simplified = simplifyEmail(email, { includeHtml: verbose || wantsHtmlBody(fields), stripQuoted: strip });
  // Set before projection so it rides along with a projected bodyText, which would
  // otherwise come back as {} on an HTML-only message (#248).
  if (typeof simplified.bodyText !== 'string') simplified.bodyTextUnavailable = true;
  // The draft body hash, attached HERE rather than inside simplifyEmail: whether one
  // can be issued honestly depends on what THIS READ returns, not on the message.
  // The rules are in attachDraftBodyHash; this is the only tool that calls it.
  attachDraftBodyHash(email, simplified, { raw, fields, stripQuoted: strip });
  return [{ type: 'text', text: raw ? toolJson(email) : toolJson(projectEmail(simplified, fields)) }];
}
