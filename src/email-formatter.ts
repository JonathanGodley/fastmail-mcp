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
  replyTo?: string[];
  inReplyTo?: string[];
  isReply?: boolean;
  isRead?: boolean;
  isFlagged?: boolean;
  isDraft?: boolean;
  preview?: string;
  listUnsubscribe?: string[];
  hasAttachment?: boolean;
  bodyText?: string;
  bodyHtml?: string;
  bodyHtmlSize?: number;
  blobId?: string;
  size?: number;
  keywords?: Record<string, boolean>;
  attachments?: Array<{
    partId?: string;
    name?: string;
    contentType: string;
    size: number;
    blobId: string;
  }>;
}

export interface SimplifyOptions {
  includeHtml?: boolean;
}

const STANDARD_KEYWORDS = new Set(['$seen', '$flagged', '$draft', '$answered', '$forwarded']);

export function formatAddress(addr: { name?: string; email: string }): string {
  if (!addr) return 'unknown';
  return addr.name ? `${addr.name} <${addr.email}>` : addr.email;
}

function extractBody(
  parts: any[] | undefined | null,
  bodyValues: Record<string, any> | undefined | null,
  preferType: 'text/plain' | 'text/html'
): string | null {
  if (!parts?.length || !bodyValues) return null;

  const chunks: string[] = [];
  for (const part of parts) {
    // Skip parts that don't match the preferred type (defensive: allow parts with no type)
    if (part.type && part.type !== preferType) continue;
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
  obj[key] = value;
}

// Only these specific flags are noise when false — omit them to save tokens.
// Other booleans (including unknown future ones) pass through as-is.
const DROP_WHEN_FALSE = new Set(['isReply', 'isFlagged', 'isDraft', 'hasAttachment']);

function addFlag(obj: Record<string, any>, key: string, value: boolean): void {
  if (!value && DROP_WHEN_FALSE.has(key)) return;
  obj[key] = value;
}

export function simplifyEmail(raw: any, options?: SimplifyOptions): SimplifiedEmail {
  const result: Record<string, any> = {
    id: raw.id,
    subject: raw.subject || '(no subject)',
    from: raw.from?.length ? formatAddress(raw.from[0]) : 'unknown',
  };

  addIf(result, 'date', raw.receivedAt);
  addIf(result, 'threadId', raw.threadId);
  addIf(result, 'messageId', raw.messageId);
  addIf(result, 'references', raw.references);
  addIf(result, 'to', (raw.to ?? []).map(formatAddress));
  addIf(result, 'cc', (raw.cc ?? []).map(formatAddress));
  addIf(result, 'bcc', (raw.bcc ?? []).map(formatAddress));
  addIf(result, 'replyTo', (raw.replyTo ?? []).map(formatAddress));
  addIf(result, 'inReplyTo', raw.inReplyTo);
  addFlag(result, 'isReply', !!(raw.inReplyTo?.length));
  addFlag(result, 'isRead', !!(raw.keywords?.$seen));
  addFlag(result, 'isFlagged', !!(raw.keywords?.$flagged));
  addFlag(result, 'isDraft', !!(raw.keywords?.$draft));
  addIf(result, 'preview', raw.preview);
  addIf(result, 'listUnsubscribe', raw['header:List-Unsubscribe:asURLs']);
  // hasAttachment is redundant when attachments array is present
  if (!raw.attachments) {
    addFlag(result, 'hasAttachment', !!raw.hasAttachment);
  }
  addIf(result, 'blobId', raw.blobId);
  addIf(result, 'size', raw.size);

  // Surface non-standard keywords (anything not $seen/$flagged/$draft/$answered/$forwarded)
  if (raw.keywords) {
    const nonStandard: Record<string, boolean> = {};
    for (const [key, value] of Object.entries(raw.keywords)) {
      if (!STANDARD_KEYWORDS.has(key) && value) {
        nonStandard[key] = true;
      }
    }
    if (Object.keys(nonStandard).length > 0) {
      result.keywords = nonStandard;
    }
  }

  const bodyText = extractBody(raw.textBody, raw.bodyValues, 'text/plain');
  const bodyHtml = extractBody(raw.htmlBody, raw.bodyValues, 'text/html');

  addIf(result, 'bodyText', bodyText);
  if (options?.includeHtml) {
    addIf(result, 'bodyHtml', bodyHtml);
  } else if (!bodyText && bodyHtml) {
    // HTML-only email — include HTML as fallback since there's no plain text
    addIf(result, 'bodyHtml', bodyHtml);
  } else if (bodyHtml) {
    // Include size so agent knows HTML exists and can request it
    addIf(result, 'bodyHtmlSize', bodyHtml.length);
  }

  const attachments = (raw.attachments ?? []).map((a: any) => {
    const att: Record<string, any> = {
      contentType: a.type ?? 'application/octet-stream',
      size: a.size ?? 0,
      blobId: a.blobId,
    };
    if (a.partId) att.partId = a.partId;
    if (a.name) att.name = a.name;
    return att;
  });
  addIf(result, 'attachments', attachments);

  return result as SimplifiedEmail;
}
