import { coerceBool, coercePosition, coerceStringArrayStrict } from './coerce.js';
import { parseEmailFields } from './field-projection.js';
import { buildExclusionNote, formatEmailQueryResult, formatRawQueryResult } from './response-formatters.js';
import type { JmapClient } from './jmap-client.js';
import type { ToolContent } from './contacts-handler.js';

/**
 * The slice of the JMAP client list_emails and search_emails need, so each handler can be
 * exercised with a stub. `JmapClient` satisfies it structurally.
 */
export interface EmailListClient {
  getEmails: JmapClient['getEmails'];
  searchEmails: JmapClient['searchEmails'];
}

// In both handlers: every flag goes through coerceBool, not !!, because a lenient client's
// stringified "false" is truthy (#54). `fields` and `position` are validated before the
// query, so a bad value costs no round trip. `limit` arrives clamped: the clamp stays in the
// CallTool case, where tool-schema.test.ts checks it against the advertised cap. The
// exclusion note rides after the JSON on both raw and simplified, so the JSON block stays
// parseable.

export async function listEmailsTool(args: any, limit: number, client: EmailListClient): Promise<ToolContent> {
  const { mailbox } = args ?? {};
  const ascending = coerceBool(args?.ascending, 'ascending') ?? false;
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const fields = parseEmailFields(args?.fields, { raw });
  const position = coercePosition(args?.position, { ascendingHint: true });
  const result = await client.getEmails({
    mailbox,
    limit,
    position,
    ascending,
    includeTrash: coerceBool(args?.includeTrash, 'includeTrash') ?? false,
    includeSpam: coerceBool(args?.includeSpam, 'includeSpam') ?? false,
    excludeDrafts: coerceBool(args?.excludeDrafts, 'excludeDrafts') ?? false,
  });
  const body = raw ? formatRawQueryResult(result) : formatEmailQueryResult(result, { fields });
  return [{ type: 'text', text: body + buildExclusionNote(result.exclusion) }];
}

export async function searchEmailsTool(args: any, limit: number, client: EmailListClient): Promise<ToolContent> {
  const { query, from, to, cc, bcc, subject, hasAttachment, isUnread, isPinned, mailbox, after, before } = args ?? {};
  const ascending = coerceBool(args?.ascending, 'ascending') ?? false;
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const fields = parseEmailFields(args?.fields, { raw });
  const position = coercePosition(args?.position, { ascendingHint: true });
  // STRICT: dropping an uncoercible scope array would silently widen the query.
  const requiredMailboxes = coerceStringArrayStrict(args?.requiredMailboxes, 'requiredMailboxes');
  const excludeMailboxes = coerceStringArrayStrict(args?.excludeMailboxes, 'excludeMailboxes');
  const result = await client.searchEmails({
    query, from, to, cc, bcc, subject,
    hasAttachment: coerceBool(hasAttachment, 'hasAttachment'),
    isUnread: coerceBool(isUnread, 'isUnread'),
    isPinned: coerceBool(isPinned, 'isPinned'),
    mailbox, requiredMailboxes, excludeMailboxes,
    after, before, limit, position,
    ascending,
    excludeDrafts: coerceBool(args?.excludeDrafts, 'excludeDrafts') ?? false,
    includeTrash: coerceBool(args?.includeTrash, 'includeTrash') ?? false,
    includeSpam: coerceBool(args?.includeSpam, 'includeSpam') ?? false,
  });
  const body = raw ? formatRawQueryResult(result) : formatEmailQueryResult(result, { fields });
  return [{ type: 'text', text: body + buildExclusionNote(result.exclusion) }];
}
