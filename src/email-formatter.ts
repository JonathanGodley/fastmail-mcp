export interface SimplifiedEmail {
  id: string;
  subject: string;
  from: string;
  date?: string;
  threadId?: string;
  messageId?: string[];
  references?: string[];
  to?: string[];
  cc?: string[];
  bcc?: string[];
  isReply?: boolean;
  isRead?: boolean;
  isFlagged?: boolean;
  isDraft?: boolean;
  bodyText?: string;
  bodyHtml?: string;
  attachments?: Array<{
    name?: string;
    contentType: string;
    size: number;
    blobId: string;
  }>;
  _extra?: Record<string, unknown>;
}

// Fields consumed by simplifyEmail — anything not in this set goes to _extra
const KNOWN_FIELDS = new Set([
  'id', 'threadId', 'messageId', 'references', 'subject', 'from', 'to', 'cc', 'bcc',
  'receivedAt', 'inReplyTo', 'keywords',
  'textBody', 'htmlBody', 'bodyValues', 'attachments',
]);

// Fields to silently drop (redundant with simplified fields)
const DROP_FIELDS = new Set(['hasAttachment']);

export function formatAddress(addr: { name?: string; email: string }): string {
  if (!addr) return 'unknown';
  return addr.name ? `${addr.name} <${addr.email}>` : addr.email;
}

function extractBody(parts: any[] | undefined | null, bodyValues: Record<string, any> | undefined | null): string | null {
  if (!parts?.length || !bodyValues) return null;

  const chunks: string[] = [];
  for (const part of parts) {
    const bv = bodyValues[part.partId];
    if (!bv?.value) continue;

    let text = bv.value;
    if (bv.isTruncated) text += '\n[body truncated]';
    if (bv.isEncodingProblem) text += '\n[encoding issues detected]';
    chunks.push(text);
  }

  return chunks.length > 0 ? chunks.join('\n') : null;
}

function addIf<T>(obj: Record<string, any>, key: string, value: T): void {
  if (value === null || value === undefined) return;
  if (Array.isArray(value) && value.length === 0) return;
  if (typeof value === 'boolean' && value === false) return;
  obj[key] = value;
}

export function simplifyEmail(raw: any): SimplifiedEmail {
  const extra: Record<string, unknown> = {};
  for (const key of Object.keys(raw)) {
    if (!KNOWN_FIELDS.has(key) && !DROP_FIELDS.has(key)) {
      extra[key] = raw[key];
    }
  }

  const result: Record<string, any> = {
    id: raw.id,
    subject: raw.subject ?? '(no subject)',
    from: raw.from?.length ? formatAddress(raw.from[0]) : 'unknown',
  };

  addIf(result, 'date', raw.receivedAt);
  addIf(result, 'threadId', raw.threadId);
  addIf(result, 'messageId', raw.messageId);
  addIf(result, 'references', raw.references);
  addIf(result, 'to', (raw.to ?? []).map(formatAddress));
  addIf(result, 'cc', (raw.cc ?? []).map(formatAddress));
  addIf(result, 'bcc', (raw.bcc ?? []).map(formatAddress));
  addIf(result, 'isReply', !!(raw.inReplyTo?.length));
  addIf(result, 'isRead', !!(raw.keywords?.$seen));
  addIf(result, 'isFlagged', !!(raw.keywords?.$flagged));
  addIf(result, 'isDraft', !!(raw.keywords?.$draft));
  addIf(result, 'bodyText', extractBody(raw.textBody, raw.bodyValues));
  addIf(result, 'bodyHtml', extractBody(raw.htmlBody, raw.bodyValues));

  const attachments = (raw.attachments ?? []).map((a: any) => {
    const att: Record<string, any> = {
      contentType: a.type ?? 'application/octet-stream',
      size: a.size ?? 0,
      blobId: a.blobId,
    };
    if (a.name) att.name = a.name;
    return att;
  });
  addIf(result, 'attachments', attachments);

  if (Object.keys(extra).length > 0) {
    result._extra = extra;
  }

  return result as SimplifiedEmail;
}
