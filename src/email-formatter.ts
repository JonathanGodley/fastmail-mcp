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
  isAnswered?: boolean;
  isForwarded?: boolean;
  mailboxes?: string[];
  roles?: string[];
  unresolvedMailboxIds?: string[];
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
  // IANA timezone name (e.g. 'America/New_York') to render `date` in. Takes
  // precedence over the module default set by setDefaultTimezone(); falls back
  // to the host zone when neither is set.
  timezone?: string;
}

// Single source of truth for the canonical message-state keywords (RFC 3501/5788/
// 8621): each maps a JMAP keyword to the simplified boolean it promotes to. Two
// concerns are derived from this one map so they can never drift (the bug behind
// #49: $answered/$forwarded were listed as "standard" — excluded from passthrough —
// but had no promotion line, so they vanished from output entirely):
//   1. Promotion — every entry produces its `is*` flag (see the loop in simplifyEmail).
//   2. Passthrough exclusion — STANDARD_KEYWORDS (the keys) is what the non-standard
//      `keywords` map omits, so a promoted flag is never also echoed in `keywords`.
// Adding a keyword here both promotes it AND excludes it from passthrough in one step.
export const KEYWORD_FLAGS: Record<string, string> = {
  $seen: 'isRead',
  $flagged: 'isFlagged',
  $draft: 'isDraft',
  $answered: 'isAnswered',
  $forwarded: 'isForwarded',
};

const STANDARD_KEYWORDS = new Set(Object.keys(KEYWORD_FLAGS));

// Static deployment config: the timezone all emails render their `date` in,
// resolved once at startup from FASTMAIL_TIMEZONE (see setDefaultTimezone). An
// explicit options.timezone overrides it per call; both unset means host zone.
let defaultTimezone: string | undefined;

export function setDefaultTimezone(tz?: string): void {
  defaultTimezone = tz && tz.trim() ? tz.trim() : undefined;
}

// Format a UTC instant in an explicit zone (or host zone when undefined).
// Throws RangeError for an invalid IANA name — callers handle the fallback.
function renderLocalIso(date: Date, zone: string | undefined): string {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: zone,
    hour12: false,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(date);

  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? '';
  const year = get('year');
  const month = get('month');
  const day = get('day');
  // Intl can emit '24' for midnight in some engines; normalize to '00'.
  let hour = get('hour');
  if (hour === '24') hour = '00';
  const minute = get('minute');
  const second = get('second');

  // Second formatter purely to read the offset for this instant/zone.
  const offsetPart = new Intl.DateTimeFormat('en-US', {
    timeZone: zone,
    timeZoneName: 'longOffset',
  })
    .formatToParts(date)
    .find((p) => p.type === 'timeZoneName')?.value ?? '';
  // 'GMT+10:00' / 'GMT-05:30' → '+10:00' / '-05:30'; bare 'GMT' (UTC) → '+00:00'.
  const stripped = offsetPart.replace('GMT', '');
  const offset = stripped === '' ? '+00:00' : stripped;

  return `${year}-${month}-${day}T${hour}:${minute}:${second}${offset}`;
}

// Render a UTC ISO instant as local ISO-8601 with numeric offset, e.g.
// "2026-03-02T08:00:00+10:00". timeZone is an IANA name; invalid/empty falls
// back to the host zone, then to the original UTC string. Never throws —
// simplifyEmail is a per-email hot path.
export function toLocalIso(utcIso: string, timeZone?: string): string {
  const zone = timeZone || defaultTimezone || undefined;
  const date = new Date(utcIso);
  try {
    return renderLocalIso(date, zone);
  } catch {
    // Invalid IANA name throws RangeError. Retry with the host zone (bypassing
    // the configured default); if even that fails, hand back the UTC string.
    if (zone) {
      try {
        return renderLocalIso(date, undefined);
      } catch {
        return utcIso;
      }
    }
    return utcIso;
  }
}

export function formatAddress(addr: { name?: string; email: string }): string {
  if (!addr) return 'unknown';
  return addr.name ? `${addr.name} <${addr.email}>` : addr.email;
}

// Format a UTC instant as a reply-attribution date in LOCAL time, e.g.
// "Mon, Jun 15, 2026, at 1:29 PM". Returns '' when the instant is absent or unparseable
// (so the caller omits the date rather than emit the literal "Invalid Date"). Reuses the
// same zone resolution as toLocalIso (timezone || defaultTimezone || host). Node ICU
// inserts a narrow no-break space (U+202F) before AM/PM, so we normalize every Unicode
// space separator to a plain ASCII space to match the captured Fastmail format. Never
// throws (a bad IANA zone falls back to the host zone, then to '').
export function formatReplyDate(utcIso: string | null | undefined, timezone?: string): string {
  if (!utcIso) return '';
  const date = new Date(utcIso);
  if (Number.isNaN(date.getTime())) return '';
  const zone = timezone || defaultTimezone || undefined;
  const render = (z: string | undefined): string => {
    const datePart = new Intl.DateTimeFormat('en-US', {
      timeZone: z, weekday: 'short', month: 'short', day: 'numeric', year: 'numeric',
    }).format(date);
    // hour:'numeric' + hour12 renders midnight as "12:00 AM" (not "0:00"/"24:00").
    const timePart = new Intl.DateTimeFormat('en-US', {
      timeZone: z, hour: 'numeric', minute: '2-digit', hour12: true,
    }).format(date);
    return `${datePart}, at ${timePart}`.replace(/\p{Zs}/gu, ' ');
  };
  try {
    return render(zone);
  } catch {
    try { return render(undefined); } catch { return ''; }
  }
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
// NOTE: isRead is deliberately NOT here — isRead:false is the only reliable "unread"
// signal (JMAP has no $unseen; absence of $seen is ambiguous), so it's always shown.
const DROP_WHEN_FALSE = new Set(['isReply', 'isFlagged', 'isDraft', 'isAnswered', 'isForwarded', 'hasAttachment']);

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

  addIf(result, 'date', raw.receivedAt ? toLocalIso(raw.receivedAt, options?.timezone) : undefined);
  addIf(result, 'threadId', raw.threadId);
  addIf(result, 'messageId', raw.messageId);
  addIf(result, 'references', raw.references);
  addIf(result, 'to', (raw.to ?? []).map(formatAddress));
  addIf(result, 'cc', (raw.cc ?? []).map(formatAddress));
  addIf(result, 'bcc', (raw.bcc ?? []).map(formatAddress));
  addIf(result, 'replyTo', (raw.replyTo ?? []).map(formatAddress));
  addIf(result, 'inReplyTo', raw.inReplyTo);
  addFlag(result, 'isReply', !!(raw.inReplyTo?.length));
  // Promote the canonical message-state keywords to `is*` flags from the single
  // KEYWORD_FLAGS map (subsumes the old hand-written isRead/isFlagged/isDraft lines
  // and adds isAnswered/isForwarded for free). DROP_WHEN_FALSE then decides which
  // show only when true (isRead always shows; the rest are dropped when false).
  for (const [kw, flag] of Object.entries(KEYWORD_FLAGS)) {
    addFlag(result, flag, !!(raw.keywords?.[kw]));
  }
  // Two AXES of message metadata, intentionally separate (do not dedupe them):
  //   - KEYWORD axis (`is*`): what the message IS / what's been done to it.
  //   - LOCATION axis (`mailboxes`/`roles`): where it's filed. Resolved in the client
  //     layer and attached as non-enumerable `_mailboxNames`/`_mailboxRoles`.
  // They normally agree (isDraft + roles:["drafts"]) but can legitimately diverge — a
  // draft moved to Trash is isDraft:true with roles:["trash"], and both are correct.
  // `mailboxes` (display names, user-renamable) and `roles` (stable JMAP roles) describe
  // the SAME set but are NOT parallel arrays — a custom folder appears in `mailboxes`
  // with no `roles` entry, so they can differ in length. Any id that couldn't be
  // resolved to a name surfaces in `unresolvedMailboxIds` (never silently dropped). (#10, #49, #53)
  addIf(result, 'mailboxes', raw._mailboxNames);
  addIf(result, 'roles', raw._mailboxRoles);
  addIf(result, 'unresolvedMailboxIds', raw._unresolvedMailboxIds);
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
