import { stripQuotedText } from './quote-strip.js';
import { buildUnionParts } from './inline-images.js';

// Reason string for `quotedStripSkipped`. Fixed wording so a client can match on it.
// "non-empty" is load-bearing: extractBody cannot tell a missing text/plain part from an
// empty one, so the wording covers both.
export const NO_PLAIN_TEXT_BODY = 'no non-empty plain-text body to strip';

export interface SimplifiedEmail {
  id: string;
  subject: string;
  from: string;
  date?: string;
  sentAt?: string;
  threadId?: string;
  messageId?: string[];
  references?: string[];
  to?: string[];
  cc?: string[];
  bcc?: string[];
  replyTo?: string[];
  inReplyTo?: string[];
  isReply?: boolean;
  // The forwarded original's Message-ID (from X-Forwarded-Message-Id). VERBOSE-tier:
  // list items show forward-ness via isForwarded instead.
  forwardedMessageId?: string[];
  // The JMAP id of the exact stored copy a reply/forward draft was composed from (from
  // X-Fastmail-MCP-Source-Id): the copy send_draft will mark. VERBOSE-tier; absent on
  // drafts made by other clients.
  sourceEmailId?: string;
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
  bodyTextSize?: number;
  // Quote-stripping signals (#73), emitted ONLY when the caller asked for stripQuoted.
  // Exactly one of the two appears, and neither is ever omitted for being "empty":
  //   quotedBytesStripped: 0 -> the body is verbatim, no marker matched.
  //   quotedBytesStripped: N -> N UTF-8 bytes of quoted history were removed.
  //   quotedStripSkipped     -> stripping did not run; the string says why.
  quotedBytesStripped?: number;
  quotedStripSkipped?: string;
  // Set by get_thread's includeBodies path for an HTML-only message (thread reads never
  // carry HTML).
  bodyTextUnavailable?: true;
  // The lost-update token edit_draft requires before it will write or clear this draft's
  // body. On a get_email read of a draft, exactly one of bodyHash / bodyHashWithheld is
  // present.
  //
  // Set by get_email after simplification, not by simplifyEmail, because whether a hash
  // can be issued depends on what the read asked for: a truncated part, a projected-away
  // body field or a stripQuoted body would make the hash describe something other than
  // the stored draft.
  bodyHash?: string;
  // Why no hash came back, naming the read that would issue one. Without it a draft read
  // with no hash is indistinguishable from a non-draft.
  bodyHashWithheld?: string;
  blobId?: string;
  size?: number;
  keywords?: Record<string, boolean>;
  // The message's parts: the JMAP `attachments` array plus the media parts the server
  // routed into the body lists (see buildUnionParts). Present only on the read paths
  // that fetch `attachments` — list/search results carry `hasAttachment` instead.
  attachments?: Array<{
    partId?: string;
    name?: string;
    contentType: string;
    size: number;
    blobId: string;
    // The server routed the part into a body list, or the sender marked it inline.
    // SENDER-DECLARED, like `name` and `contentType`: never a sign the part is safe to
    // skip (#13).
    isInline?: true;
    // The part's Content-ID, verbatim: what a `cid:` reference in the HTML body points
    // at. Omitted when the part has none.
    cid?: string;
  }>;
}

export interface SimplifyOptions {
  includeHtml?: boolean;
  // false: an HTML-only message returns no bodyHtml (only bodyHtmlSize). get_thread sets it,
  // because its body cap and its description count plain text only.
  htmlFallback?: boolean;
  // Remove recognised quoted correspondence from `bodyText` and report how much went
  // (#73). Opt-in per call; the default output is always verbatim. Applies to the plain
  // text body ONLY — a `bodyHtml` returned alongside it (verbose) is untouched.
  stripQuoted?: boolean;
  // IANA timezone name (e.g. 'America/New_York') to render `date` and `sentAt` in. Takes
  // precedence over the module default set by setDefaultTimezone(); falls back
  // to the host zone when neither is set.
  timezone?: string;
}

// The canonical message-state keywords (RFC 3501/5788/8621), each mapped to the `is*`
// flag it promotes to. Both promotion and the passthrough exclusion (STANDARD_KEYWORDS)
// derive from this one map so they cannot drift (#49).
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

// The same stored value, for paths that INTERPRET a local date (list_calendar_events'
// date-only window). Read from here, never re-derived from the environment, so a calendar
// day and an email's printed `date` cannot disagree. `undefined` means the host zone.
export function getDefaultTimezone(): string | undefined {
  return defaultTimezone;
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
  // A local-mean-time offset carries seconds, which ISO 8601 and new Date() reject; UTC
  // renders the same instant.
  if (!/^[+-]\d\d:\d\d$/.test(offset)) return renderLocalIso(date, 'UTC');

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
// "Mon, Jun 15, 2026, at 1:29 PM". Returns '' when the instant is absent or unparseable,
// so the caller omits the date rather than print "Invalid Date". Node ICU inserts a
// narrow no-break space (U+202F) before AM/PM, so every Unicode space separator becomes
// an ASCII space to match the captured Fastmail format. Never throws.
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
  addIf(result, 'sentAt', raw.sentAt ? toLocalIso(raw.sentAt, options?.timezone) : undefined);
  addIf(result, 'threadId', raw.threadId);
  addIf(result, 'messageId', raw.messageId);
  addIf(result, 'references', raw.references);
  addIf(result, 'to', (raw.to ?? []).map(formatAddress));
  addIf(result, 'cc', (raw.cc ?? []).map(formatAddress));
  addIf(result, 'bcc', (raw.bcc ?? []).map(formatAddress));
  addIf(result, 'replyTo', (raw.replyTo ?? []).map(formatAddress));
  addIf(result, 'inReplyTo', raw.inReplyTo);
  addFlag(result, 'isReply', !!(raw.inReplyTo?.length));
  addIf(result, 'forwardedMessageId', raw['header:X-Forwarded-Message-Id:asMessageIds']);
  // asText header values can carry folding whitespace; trim, and treat blank as absent
  // (mirrors readSourceReferences in jmap-client.ts).
  const rawSourceId = raw['header:X-Fastmail-MCP-Source-Id:asText'];
  addIf(result, 'sourceEmailId', typeof rawSourceId === 'string' && rawSourceId.trim() !== '' ? rawSourceId.trim() : undefined);
  for (const [kw, flag] of Object.entries(KEYWORD_FLAGS)) {
    addFlag(result, flag, !!(raw.keywords?.[kw]));
  }
  // Two AXES of message metadata, intentionally separate (do not dedupe them): the
  // keyword axis (`is*`, what the message is) and the location axis (`mailboxes`/`roles`,
  // where it is filed, attached by the client layer as non-enumerable `_mailbox*`). They
  // can legitimately diverge: a draft moved to Trash is isDraft:true with roles:["trash"].
  // `mailboxes` and `roles` are NOT parallel arrays: a custom folder has no role.
  // (#10, #49, #53)
  addIf(result, 'mailboxes', raw._mailboxNames);
  addIf(result, 'roles', raw._mailboxRoles);
  addIf(result, 'unresolvedMailboxIds', raw._unresolvedMailboxIds);
  addIf(result, 'preview', raw.preview);
  addIf(result, 'listUnsubscribe', raw['header:List-Unsubscribe:asURLs']);
  // hasAttachment is suppressed whenever the read fetched `attachments` at all (an empty
  // array counts): the listing is authoritative. The two are NOT interchangeable, since
  // hasAttachment is a server content-or-decoration heuristic (docs/conventions.md): a
  // message whose only image is embedded in its body can report false yet list a part.
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
  } else if (!bodyText && bodyHtml && options?.htmlFallback !== false) {
    // HTML-only email — include HTML as fallback since there's no plain text
    addIf(result, 'bodyHtml', bodyHtml);
  } else if (bodyHtml) {
    // Include size so agent knows HTML exists and can request it
    addIf(result, 'bodyHtmlSize', bodyHtml.length);
  }

  // bodyTextSize: compact (list/search) reads fetch the textBody part structure but no
  // bodyValues, so the total text/* size tells a ~256-char `preview` apart from a large
  // message (#59). An upper bound: quoted history is included. Emitted only when no body
  // content is present.
  if (!bodyText && !bodyHtml && Array.isArray(raw.textBody)) {
    const textBytes = raw.textBody.reduce(
      (sum: number, p: any) =>
        typeof p?.type === 'string' && p.type.startsWith('text/') && typeof p.size === 'number'
          ? sum + p.size
          : sum,
      0
    );
    if (textBytes > 0) addIf(result, 'bodyTextSize', textBytes);
  }

  // Quote stripping (#73) runs on the plain-text body only: HTML quoting has no reliable
  // text-level boundary, so any `bodyHtml` above is left verbatim and not counted.
  if (options?.stripQuoted) {
    if (typeof bodyText === 'string') {
      const { text, quotedBytesStripped } = stripQuotedText(bodyText);
      result.bodyText = text;
      result.quotedBytesStripped = quotedBytesStripped;
    } else {
      result.quotedStripSkipped = NO_PLAIN_TEXT_BODY;
    }
  }

  // The part listing is the UNION of `attachments` and the media parts the server routed
  // into the body lists, where some MIME shapes put an embedded image (#13).
  //
  // Computed BESIDE the raw email, never onto it: `raw: true` paths hand `raw` back
  // untouched.
  const attachments = buildUnionParts(raw).map(({ part, inBodyList }) => {
    const att: Record<string, any> = {
      contentType: part.type ?? 'application/octet-stream',
      size: part.size ?? 0,
      blobId: part.blobId,
    };
    if (part.partId) att.partId = part.partId;
    if (part.name) att.name = part.name;
    // Two independent grounds, per RFC 8621 §4.1.4: the routing itself, and the
    // sender's own Content-Disposition. Neither is available on the other's shapes,
    // so both are needed for the flag to mean the same thing on every message.
    if (inBodyList || String(part.disposition ?? '').trim().toLowerCase() === 'inline') {
      att.isInline = true;
    }
    // Verbatim and untruncated: the value has to compare equal to the `cid:` reference
    // in the body and to download_attachment's `cid:` handle, so it is data in a JSON
    // field, not prose. (Prose that echoes a cid renders it through describePart.)
    if (part.cid) att.cid = part.cid;
    return att;
  });
  addIf(result, 'attachments', attachments);

  return result as SimplifiedEmail;
}
