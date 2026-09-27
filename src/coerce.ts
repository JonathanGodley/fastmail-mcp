import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { describePart, DESCRIBE_PART_MAX, isAuthorableCid, stripCidSpelling } from './inline-images.js';
import { rejectUnusableCid } from './inline-notes.js';

// Tagged error for filesystem-path access decisions (path confinement and the
// attachment opt-in gate). Thrown by the path guards and attachment upload in
// jmap-client.ts, which deliberately stays free of MCP SDK types — the index
// boundary maps every PathAccessError to McpError(InvalidParams). instanceof is
// the discriminator, so the message text carries no routing burden.
export class PathAccessError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PathAccessError';
  }
}

// Tagged error for caller-supplied input that is well-formed JSON but semantically
// invalid (e.g. a `mailbox` that resolves to nothing). Thrown from jmap-client.ts, and
// mapped at the index boundary to McpError(InvalidParams) like PathAccessError. That
// boundary redacts every branch with no per-class exemption, so "no unredacted error text
// reaches tool output" stays a grep; do not add one here. Redaction is not what makes the
// reflected-input oracle acceptable; see docs/security-model.md.
export class InvalidInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InvalidInputError';
  }
}

// Some MCP clients (e.g. Claude Cowork as of 2026-04-08, issue #54) stringify
// structured params before dispatch. These helpers coerce such values back to
// their expected shapes so the handlers work against both strict and lenient clients.

// Credential-shaped substrings scrubbed from anything reflected back to the caller.
// Deliberately narrow: provider error messages are what the caller recovers from.
const BEARER_PATTERN = /Bearer\s+\S+/gi;
// The CalDAV path authenticates with HTTP Basic, so a reflected header or a
// tsdav error carries the base64 credential blob — redact that shape too.
const BASIC_PATTERN = /Basic\s+[A-Za-z0-9+/=]+/gi;
// Fastmail token shape. The charset is `[\w-]`, not `[A-Za-z0-9-]`: under the
// narrower class an underscore inside the token ends the match early, and with
// fewer than 20 characters before it the `{20,}` quantifier fails outright, so
// the whole token would pass through in clear.
const FASTMAIL_TOKEN_PATTERN = /fmu\d+-[\w-]{20,}/g;

// Exact secret values registered at startup, for credentials the patterns above cannot
// see (a CalDAV password, a self-hosted token with neither prefix). Never logged.
const KNOWN_SECRETS = new Set<string>();

// Values under 8 characters are ignored: an over-broad match would mangle legitimate
// output for no security gain.
export function registerSecret(value: string | undefined): void {
  if (typeof value === 'string' && value.length >= 8) {
    KNOWN_SECRETS.add(value);
  }
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function redactBearerTokens(input: string): string {
  let out = input
    .replace(BEARER_PATTERN, 'Bearer [REDACTED]')
    .replace(BASIC_PATTERN, 'Basic [REDACTED]')
    .replace(FASTMAIL_TOKEN_PATTERN, 'fmu[REDACTED]');
  for (const secret of KNOWN_SECRETS) {
    out = out.replace(new RegExp(escapeRegExp(secret), 'g'), '[REDACTED]');
  }
  return out;
}

/**
 * Render an untrusted value into prose: REDACT it, then neutralise and truncate it.
 *
 * Any value not written by this server (a caller-supplied id, a mailbox name, a
 * server-authored set-error description) that is interpolated into a message a caller reads
 * back goes through this. Only the value, never the server's own sentence around it. The
 * model is in docs/conventions.md.
 *
 * The order is the landmine (#131). `describePart` strips line breaks and bidi overrides,
 * swaps `"` for `'`, and truncates at 64 code points. `redactBearerTokens` is
 * length-sensitive (FASTMAIL_TOKEN_PATTERN needs 20+ characters, a registered secret is an
 * exact match), so truncating first lets a token prefix out verbatim. Backwards still passes
 * every line-forging test.
 *
 * No second parameter on purpose: callers pass this to `.map` bare, and `map` would hand it
 * the index as a bound. A wider bound goes through `describeUntrustedAt`.
 *
 * A caller that quotes the value uses `"…"`: the swap protects that span only, and inside
 * `'…'` the value's own `'` closes it (#190). A bare render is judged on the whole sentence,
 * so a new `'…'` span in any sentence that renders a bare value reopens this. The drift guard
 * in coerce.test.ts catches a single-quoted `${describeUntrusted(…)}`; the whole-sentence
 * half is a reading at the sentence you are editing.
 *
 * Not for a structured result item; see `redactedJson`.
 */
export function describeUntrusted(value: unknown): string {
  return describeUntrustedAt(value, DESCRIBE_PART_MAX);
}

/**
 * `describeUntrusted` at a bound named by the caller, for a value the 64-code-point default
 * renders useless — the same two steps in the same order, at a width that value can survive.
 *
 * The bound moves the TRUNCATION and nothing else. Redaction still runs over the whole value
 * first, so a wider echo cannot let a credential through. Pass a named constant carrying its
 * reason, as `PATH_ECHO_LIMIT` does.
 */
export function describeUntrustedAt(value: unknown, max: number): string {
  const source = typeof value === 'string' ? value : value == null ? '' : String(value);
  return describePart(redactBearerTokens(source), max);
}

/**
 * Serialise a value to JSON with every string inside it redacted.
 *
 * The ONLY safe way to redact a structured result item. BEARER_PATTERN's `\S+` runs to the
 * next whitespace, which over a finished document is the string's closing quote and comma,
 * so redacting the serialised text eats delimiters and the item stops parsing (a mailbox
 * named "Bearer Bonds" is enough). Per value there is no trailing delimiter to swallow.
 * Prose calls redactBearerTokens directly; anything JSON.stringify touches comes here.
 *
 * Compact like toolJson below: no indent argument.
 */
export function redactedJson(value: any): string {
  return JSON.stringify(value, (_key, v) => (typeof v === 'string' ? redactBearerTokens(v) : v));
}

/**
 * Serialise a tool result payload. THE one seam every JSON result item goes through, across
 * every handler and formatter, so how this server serialises is decided once (#40).
 *
 * Compact, with no option to indent: every payload is read by a machine, and indentation was
 * ~17% of a 25-message list page's bytes. That includes JSON embedded in a prose frame (a list
 * summary line, the bulk-operations diagnostic). See docs/conventions.md, result serialisation.
 *
 * Use redactedJson above instead where the values may carry credentials.
 */
export function toolJson(value: unknown): string {
  return JSON.stringify(value);
}

// Every branch TRIMS its elements, so the three spellings of a list agree; a padded value
// otherwise reaches the server and comes back as a not-found. coerceStringArrayStrict
// relies on this for its trim.
//
// One non-obvious case: a removeAttachments ref is matched against a MIME filename, which
// may legally carry surrounding spaces. resolveAttachmentRemovals trims both sides; if that
// ever stops being true, this trim starts hiding such an attachment.
const trimAll = (values: unknown[]): string[] => values.map(v => String(v).trim());

export function coerceStringArray(value: unknown): string[] | undefined {
  if (value === undefined || value === null) return undefined;
  if (Array.isArray(value)) return trimAll(value);
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return [];
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return trimAll(parsed);
    } catch { /* fall through to comma-split */ }
  }
  return trimmed.split(',').map(s => s.trim()).filter(Boolean);
}

// coerceStringArray for a parameter that must FAIL CLOSED: a present value that cannot be
// coerced is rejected instead of coming back as `undefined` and read as "not supplied".
// Use it where the value NARROWS what the call touches, since a dropped scoping argument
// silently widens the query (docs/conventions.md, lenient input coercion).
//
// Strict per ELEMENT too: a non-string entry is rejected by index rather than passed through
// `String()`, which would send `[null]` to the matcher as the text "null". A top-level `null`
// is still absent: a lenient client emits it for every key it has nothing to say about.
export function coerceStringArrayStrict(value: unknown, paramName: string): string[] | undefined {
  if (value === undefined || value === null) return undefined;

  // Unwrap a JSON-string array HERE, so the elements are type-checked before `String()`
  // erases what they were.
  let candidate: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) candidate = parsed;
      } catch { /* not JSON: leave it for the comma-split path */ }
    }
  }

  if (Array.isArray(candidate)) {
    candidate.forEach((entry, i) => {
      if (typeof entry !== 'string') {
        const kind = entry === null ? 'null' : Array.isArray(entry) ? 'array' : typeof entry;
        throw new InvalidInputError(`${paramName}[${i}] must be a string; received ${kind}.`);
      }
      // coerceStringArray drops blanks only on the comma-split branch, where "a,,b" is a
      // separator artefact; an array's `''` would otherwise reach the lookup as a value.
      if (entry.trim() === '') {
        throw new InvalidInputError(`${paramName}[${i}] must be a non-empty string.`);
      }
    });
  }

  const coerced = coerceStringArray(candidate);
  if (coerced === undefined) {
    throw new InvalidInputError(
      `${paramName} must be an array of strings (or a comma-separated string); received ${typeof value}.`,
    );
  }
  return coerced;
}

// Coerce the four recipient fields into string[] | undefined, so the JMAP client's
// .map(parseAddress) never receives a bare string (#54).
//
// STRICT, because a dropped recipient reads as a legitimate outcome on both tools: on a
// draft_email reply an omitted `to` means reply-all (carrying the original's Bcc), and on
// edit_draft it means "leave unchanged" while the edit reports success. `''` and `[]` still
// coerce to the empty list, which the bcc description documents as "treated as omitted".
export function coerceRecipients(args: { to?: unknown; cc?: unknown; bcc?: unknown; replyTo?: unknown }): {
  to?: string[]; cc?: string[]; bcc?: string[]; replyTo?: string[];
} {
  return {
    to: coerceStringArrayStrict(args.to, 'to'),
    cc: coerceStringArrayStrict(args.cc, 'cc'),
    bcc: coerceStringArrayStrict(args.bcc, 'bcc'),
    replyTo: coerceStringArrayStrict(args.replyTo, 'replyTo'),
  };
}

// Hard-reject any argument key the tool didn't declare in its inputSchema, so a
// misspelled/hallucinated param (e.g. `mailbox` vs `mailboxId`) fails loudly
// instead of being silently dropped and the tool running with defaults (#11).
// KEY-strictness only — value coercion is handled separately and is untouched.
// `additionalProperties: true` on a tool's schema opts that tool out (none today).
export function assertKnownParams(
  toolName: string,
  args: Record<string, unknown> | null | undefined,
  allowedKeys: Set<string>,
  additionalProperties: boolean,
): void {
  if (additionalProperties) return;
  if (args === null || args === undefined) return;
  const unknown = Object.keys(args).filter(k => !allowedKeys.has(k));
  if (unknown.length === 0) return;
  throw new McpError(
    ErrorCode.InvalidParams,
    `Unknown parameter(s): ${unknown.join(', ')}. Valid: ${[...allowedKeys].join(', ')}`,
  );
}

// A boolean parameter: true, "true" in any case, 1 or "1" read as true; false, "false", 0 or
// "0" as false; null/undefined as absent. Anything else is REFUSED naming the parameter: read
// as absent it would silently drop a filter (`isUnread:"yes"`) or keep the default the caller
// was trying to change (`includeTrash:"1"`).
export function coerceBool(value: unknown, paramName: string): boolean | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value === 'boolean') return value;
  if (value === 1) return true;
  if (value === 0) return false;
  if (typeof value === 'string') {
    const word = value.trim().toLowerCase();
    if (word === 'true' || word === '1') return true;
    if (word === 'false' || word === '0') return false;
  }
  const received = typeof value === 'string'
    ? `"${describeUntrusted(value)}"`
    : typeof value === 'number' ? String(value)
    : Array.isArray(value) ? 'an array' : `a ${typeof value}`;
  throw new InvalidInputError(
    `${paramName} must be true or false ("true"/"false" in any case, or 1/0, are also accepted); received ${received}.`,
  );
}

// JMAP filter conditions take a UTCDate (RFC 8620 §1.4): an RFC 3339 date-time whose
// offset is literally `Z`, e.g. `2026-07-20T00:00:00Z`. A bare `2026-07-20` is valid
// ISO 8601 but the server rejects it with an opaque `invalidArguments` that names no
// argument (#70), so normalise here instead of passing the caller's string through.
//
// The two accepted shapes are matched explicitly and everything else is rejected, NOT
// coerced leniently: `new Date()`'s fallback parser reads `2026/07/20` as HOST-LOCAL
// midnight and rolls `2026-2-31` into the next month, silently moving the search window.
const DATE_ONLY_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
const DATE_TIME_PATTERN = /^(\d{4}-\d{2}-\d{2})T.+$/;

// Longest caller value echoed back in a rejection message. Enough to recognise the bad
// value, short enough that a pasted blob doesn't become the error.
const DATE_ECHO_LIMIT = 60;

/**
 * The ONE way this server quotes caller-supplied text back inside an error message. Four
 * rules, and they travel together:
 *
 *   TRIM, so the value quoted is the value the coercion actually judged.
 *   SCRUB control characters and U+2028/U+2029, which would forge extra lines.
 *   NEUTRALISE `"` into `'`, so a value cannot close the `"…"` span a caller renders it in
 *     and have the rest read as the server's next sentence. That protects `"…"` only, so
 *     EVERY CALLER THAT QUOTES USES `"…"` (#190). A caller that renders the value bare is
 *     judged on the WHOLE sentence: adding a `'…'` span to a sentence that carries a bare
 *     value reopens this. `describePart` applies the same rule.
 *   BOUND it, with a VISIBLE truncation marker, so a reader can tell a cut value from a
 *     short one.
 */
export function echoCallerText(value: unknown, limit: number = DATE_ECHO_LIMIT): string {
  const text = typeof value === 'string' ? value : String(value);
  const clean = text.replace(/[\u0000-\u001F\u007F\u2028\u2029]/g, ' ').replace(/"/g, "'").trim();
  return clean.length > limit ? `${clean.slice(0, limit)}…` : clean;
}

// Normalise a caller-supplied date/datetime into the JMAP UTCDate shape, reading anything
// without a zone in `zone`, the configured zone (`undefined` = the host's), exactly as the
// calendar window does:
//
//   2026-07-20                -> midnight at the start of the 20th in `zone`
//   2026-07-20T14:30:00       -> 14:30 in `zone`
//   2026-07-20T14:30:00Z      -> 2026-07-20T14:30:00Z
//   2026-07-20T14:30:00+01:00 -> 2026-07-20T13:30:00Z   (offset applied)
//
// Anything else is REJECTED, naming the parameter. An empty string is rejected too rather
// than treated as "no filter": silently dropping a date bound widens the search.
export function coerceUtcDate(value: unknown, paramName: string, zone: string | undefined): string | undefined {
  return resolveWindowBound(value, paramName, zone, 0, acceptedDateFormats(describeTimezone(zone)));
}

// What SHAPE a caller's date argument is. The distinction the calendar cares about is the
// third one: a datetime carrying no zone designator names a wall clock, not an instant, so
// something has to decide which zone reads it.
type DateValueKind = 'date' | 'local-datetime' | 'zoned-datetime';

// Every accepted datetime ends in `Z` or a numeric offset; anything else is a wall clock.
const ZONE_DESIGNATOR_PATTERN = /(?:Z|[+-]\d{2}:?\d{2})$/i;
// The wall-clock datetimes the calendar resolves itself, captured so the components can be
// placed in a zone. Stricter than DATE_TIME_PATTERN's `T.+` so an unreadable hour is
// rejected, not guessed. SHAPE ONLY: the ranges are checked in `isWallClockInRange`.
const LOCAL_DATETIME_PATTERN = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?$/;

/**
 * The shared validation behind every date argument: type, emptiness, shape, and a real
 * calendar day, so `coerceUtcDate` and the calendar-window pair below reject identically and
 * diverge only on what an accepted value resolves to.
 *
 * It does NOT check the TIME components: `coerceUtcDate` gets that from `new Date()`, and the
 * calendar pair must do it in `isWallClockInRange`, or `Date.UTC` rolls `99:99:99` forward.
 */
function classifyDateValue(
  value: unknown,
  paramName: string,
  formats: string,
): { trimmed: string; kind: DateValueKind } {
  if (typeof value !== 'string') {
    throw new InvalidInputError(
      `${paramName} must be a date string, not ${Array.isArray(value) ? 'an array' : `a ${typeof value}`}. ${formats}`,
    );
  }
  const trimmed = value.trim();
  if (!trimmed) {
    throw new InvalidInputError(
      `${paramName} cannot be empty; omit it to search without that date bound. ${formats}`,
    );
  }

  const dateOnly = DATE_ONLY_PATTERN.test(trimmed);
  const datePart = dateOnly ? trimmed : DATE_TIME_PATTERN.exec(trimmed)?.[1];
  if (!datePart) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${formats}`,
    );
  }

  // A nonexistent day parses rather than failing (2026-02-31 becomes 2026-03-03). Probe the
  // date part alone, since an offset legitimately moves the whole value's UTC date.
  const dayProbe = new Date(`${datePart}T00:00:00Z`);
  if (Number.isNaN(dayProbe.getTime()) || !dayProbe.toISOString().startsWith(datePart)) {
    throw new InvalidInputError(
      `${paramName} is not a real calendar date: "${echoDate(trimmed)}". ${formats}`,
    );
  }

  const kind: DateValueKind = dateOnly
    ? 'date'
    : ZONE_DESIGNATOR_PATTERN.test(trimmed)
      ? 'zoned-datetime'
      : 'local-datetime';
  return { trimmed, kind };
}

/** A caller's date value, quoted back in a rejection under the one shared echo policy. */
function echoDate(value: string): string {
  return echoCallerText(value, DATE_ECHO_LIMIT);
}

// ===========================================================================
// Calendar window bounds: a DAY is a local day, not a UTC day.
// ===========================================================================
//
// `list_calendar_events`' startDate/endDate resolve differently from every email search
// bound on purpose; the reasoning is in docs/conventions.md, "A calendar window's DAY is a local day".
//
//   startDate: 2026-08-12   ->  local midnight on the 12th
//   endDate:   2026-08-12   ->  local midnight on the 13th   (the whole of the 12th)
//   either:    2026-08-12T09:00:00      ->  09:00 in the CONFIGURED zone, not the host's
//   either:    2026-08-12T09:00:00Z     ->  exactly what it says; zone rules do not apply
//   either:    2026-08-12T09:00:00+10:00 -> exactly what it says
//
// A UTC-day reading silently drops a +10:00 user's morning appointments from "the 12th".
// The end is EXCLUSIVE (RFC 4791 section 9.9), so the next midnight is "through the end of
// that day" and a same-date start and end is not a zero-length window.
//
// `zone` is passed in rather than read from module state so a test can exercise a zone other
// than the host's; host-zone-only tests pass under either reading. `undefined` means the host.

function acceptedWindowFormats(zoneLabel: string): string {
  return `Accepted: a date such as 2026-08-12 (read as a whole day in ${zoneLabel}), or a full datetime such as ` +
    '2026-08-12T14:30:00Z or 2026-08-12T14:30:00+10:00 (taken exactly as written). Also accepted: no ' +
    'seconds (2026-08-12T14:30Z), fractional seconds, dropped (2026-08-12T14:30:00.5Z), a lowercase z ' +
    '(2026-08-12T14:30:00z) and an offset without its colon (2026-08-12T14:30:00+1000). A datetime with no ' +
    `Z and no offset is read as ${zoneLabel} local time.`;
}

/** The host's own IANA zone name, for the two places a configured zone is not usable. */
function hostTimezone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}

/** Whether ICU can actually resolve an IANA name, as opposed to it merely being a string. */
export function isUsableTimezone(zone: string): boolean {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: zone });
    return true;
  } catch {
    return false;
  }
}

// The bound for an echoed zone name, shared by every zone rejection.
export const ZONE_ECHO_LIMIT = 40;

// The bound for a filesystem path echoed back by a path-confinement refusal. Far wider than
// `describePart`'s 64: such a refusal names the resolved path AND the allowed directory, and
// two paths sharing a long ancestor truncate to the same prefix at 64, which cannot be acted on.
export const PATH_ECHO_LIMIT = 200;

/**
 * The path spelling of `echoCallerText`: redact, then neutralise and bound at
 * `PATH_ECHO_LIMIT`, the same two steps in the same order `describeUntrusted` runs. Callers
 * quote it with `"…"`.
 */
export function echoPath(value: unknown): string {
  const source = typeof value === 'string' ? value : value == null ? '' : String(value);
  return echoCallerText(redactBearerTokens(source), PATH_ECHO_LIMIT);
}

const zoneCanonicalizationCache = new Map<string, string>();
export const ZONE_CANONICALIZATION_CACHE_LIMIT = 512;

export function zoneCanonicalizationCacheSize(): number {
  return zoneCanonicalizationCache.size;
}

export function zoneCanonicalizationCacheHas(zone: string): boolean {
  return zoneCanonicalizationCache.has(zone);
}

/**
 * ICU's canonical spelling for a zone name — `Intl.DateTimeFormat`'s own name for whatever the
 * string resolves to, or the string unchanged when ICU cannot resolve it. The one seam every
 * zone comparison and every written zone routes through (write side #157, read-side
 * `zoneNamesEqual` #139), so an alias canonicalises identically on write and on read.
 *
 * Cached because `zoneNamesEqual` runs per event on every list read. The keys are untrusted:
 * every stored TZID reaches here, and an invitation's sender chooses it. So only names ICU
 * resolves are kept, and at most `ZONE_CANONICALIZATION_CACHE_LIMIT` of them, oldest dropped
 * first, since ICU matches case-insensitively and each case variant is a distinct key.
 */
export function canonicalZoneName(zone: string): string {
  const cached = zoneCanonicalizationCache.get(zone);
  if (cached !== undefined) return cached;
  if (!isUsableTimezone(zone)) return zone;
  const resolved = new Intl.DateTimeFormat('en-US', { timeZone: zone }).resolvedOptions().timeZone;
  if (zoneCanonicalizationCache.size >= ZONE_CANONICALIZATION_CACHE_LIMIT) {
    zoneCanonicalizationCache.delete(zoneCanonicalizationCache.keys().next().value!);
  }
  zoneCanonicalizationCache.set(zone, resolved);
  return resolved;
}

/**
 * The IANA name actually used for a configured zone: `zone` itself, canonicalised through
 * `canonicalZoneName`, when it is set and ICU can resolve it; otherwise the host's own zone.
 * The one rule for "which zone is really in force".
 *
 * Canonicalised here too because the configured default is written straight into a TZID when
 * create_calendar_event gets no `timeZone`; a raw `australia/sydney` would later compare
 * unequal to the canonical spelling on a read-modify-write.
 */
export function resolveUsableTimezone(zone: string | undefined): string {
  if (zone && isUsableTimezone(zone)) return canonicalZoneName(zone);
  return hostTimezone();
}

// The ways a zone candidate fails the rule shared by validateCallerTimezone and
// resolveConfiguredTimezone (#157); a reason rather than a boolean so each site frames its own
// sentence.
type ZoneRejectionReason = 'offset-shaped' | 'unresolvable' | 'shorthand';

// Quoted by both shorthand rejections, so the warning reads the same on either path.
const SHORTHAND_ZONE_WARNING =
  '"EST" resolves to a fixed-offset zone with no daylight saving, not US Eastern, and other ' +
  'abbreviations and aliases ("NZ", "PST", "GMT", "Zulu"...) are just as ambiguous';

/**
 * Classify a zone candidate, or return `null` when it is acceptable. The candidate must be
 * PRE-canonical with any leading `/` stripped (RFC 5545 §3.2.19): a canonical name would let
 * `NZ` through as `Pacific/Auckland`.
 *
 * The order is the contract:
 *   1. Offset shapes (`+10:00`, `GMT+10`, ...) first, because ICU resolves several of them as
 *      though they were zone names.
 *   2. ICU resolvability next, so `Blah` is not told that a slash is all it needed.
 *   3. Then the slash rule (#157): a name needs a region-qualifying slash, with "UTC" the one
 *      exception. It rejects every abbreviation ICU resolves ("EST" is a fixed-offset Panama
 *      zone, not US Eastern), with no safe-list. "US/Pacific" is unaffected.
 */
function zoneRejectionReason(zoneCandidate: string): ZoneRejectionReason | null {
  if (/^[+-]/.test(zoneCandidate) || /^(GMT|UTC|UT)[+-]/i.test(zoneCandidate) || /^\d/.test(zoneCandidate)) {
    return 'offset-shaped';
  }
  if (!isUsableTimezone(zoneCandidate)) {
    return 'unresolvable';
  }
  if (!zoneCandidate.includes('/') && zoneCandidate.toUpperCase() !== 'UTC') {
    return 'shorthand';
  }
  return null;
}

/**
 * Validate a caller-supplied `timeZone` argument (#157) and return the canonical IANA name to
 * write. `isUsableTimezone` alone is the WRONG gate: ICU resolves `"+10:00"` and similar, which
 * would bake an offset into a TZID. The rule and its order are `zoneRejectionReason`'s.
 *
 * Fails closed on `null` and blank rather than reading them as "write floating": the read side
 * emits `timeZone: null` for a floating event (#139), so an echoed `null` would decide that by
 * accident. Nothing here writes a zone-less time (docs/conventions.md).
 *
 * A Windows zone id ("AUS Eastern Standard Time") rejects: ICU cannot resolve it, so there is
 * nothing to canonicalise against. A known round-trip limit, stated in both tools' descriptions.
 */
export function validateCallerTimezone(value: unknown): string {
  if (value === null || (typeof value === 'string' && value.trim().length === 0)) {
    throw new InvalidInputError(
      'timeZone cannot be null, empty, or whitespace-only. A floating time (no zone at all) cannot ' +
      'be written through this parameter. Omit timeZone instead: on create that writes the ' +
      "account's configured zone, and on update it leaves whatever the event already has unchanged."
    );
  }
  if (typeof value !== 'string') {
    throw new InvalidInputError(`timeZone must be an IANA zone name string (e.g. "Australia/Sydney"), not ${typeof value}.`);
  }
  const trimmed = value.trim();
  // RFC 5545 §3.2.19 permits a leading '/' on a TZID, and the read side (#139) emits it
  // verbatim, so strip it or a read-modify-write echo is rejected. ICU throws on the slash, so
  // this runs before every check. A vendor-prefixed form still fails, the safe direction.
  const zoneCandidate = trimmed.startsWith('/') ? trimmed.slice(1) : trimmed;
  const rejection = zoneRejectionReason(zoneCandidate);
  if (rejection === 'offset-shaped') {
    throw new InvalidInputError(
      `timeZone must be an IANA zone NAME such as "Australia/Sydney" or "America/New_York", not a ` +
      `fixed UTC offset ("${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}"). A calendar time is never ` +
      'written with an offset — pass the zone the wall clock is actually in and this server works ' +
      'out the offset itself.'
    );
  }
  if (rejection === 'unresolvable') {
    throw new InvalidInputError(
      `timeZone "${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}" is not a time zone this server can ` +
      'resolve. Pass a standard IANA name such as "Australia/Sydney" or "America/New_York".'
    );
  }
  if (rejection === 'shorthand') {
    throw new InvalidInputError(
      `timeZone "${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}" is a zone abbreviation or alias, not a ` +
      'full IANA zone name. Pass a name that contains a region, such as "Australia/Sydney" or ' +
      `"America/New_York" — "UTC" is the one accepted exception. ${SHORTHAND_ZONE_WARNING}.`
    );
  }
  // Do not return the caller's spelling: ICU resolves names case-insensitively and through
  // links, but Cyrus looks a TZID up by exact string, so only the canonical name is one the
  // calendar server can read back.
  return canonicalZoneName(zoneCandidate);
}

/**
 * Resolve the zone the server actually runs with, from FASTMAIL_TIMEZONE or the host zone, held
 * to `zoneRejectionReason`'s rule (#157). Called once, from `runServer()` in index.ts.
 *
 * The two sources are handled ASYMMETRICALLY on purpose:
 *
 * - A set-and-invalid FASTMAIL_TIMEZONE THROWS, and `runServer()` refuses to start: someone
 *   chose that value, and falling back silently is the wrong-day failure the rule prevents.
 * - A rejected host zone (used only when nothing is configured) falls back to `'UTC'` with a
 *   `warning` the caller MUST print, since date-only window bounds then read as UTC days.
 *   Refusing to start over a setting nobody touched would punish the operator. Close to
 *   unreachable, since ICU reports a canonical name; it guards a misconfigured TZ variable.
 */
export function resolveConfiguredTimezone(
  configuredValue: string | undefined,
  hostZone: string = hostTimezone(),
): { zone: string; warning?: string } {
  if (configuredValue !== undefined) {
    const trimmed = configuredValue.trim();
    const zoneCandidate = trimmed.startsWith('/') ? trimmed.slice(1) : trimmed;
    const rejection = zoneRejectionReason(zoneCandidate);
    if (rejection === 'offset-shaped') {
      throw new InvalidInputError(
        `FASTMAIL_TIMEZONE is set to "${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}", a fixed UTC offset, not a zone ` +
        'name — a calendar time is never written with an offset. This server refuses to start on an unusable ' +
        'configured time zone rather than falling back to it silently, because that silence is exactly how a ' +
        'wrong day happens with nothing said. Set FASTMAIL_TIMEZONE to a full IANA zone name such as ' +
        '"Australia/Sydney" or "America/New_York" (or "UTC"), or unset it to use this server\'s own zone.'
      );
    }
    if (rejection === 'unresolvable') {
      throw new InvalidInputError(
        `FASTMAIL_TIMEZONE is set to "${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}", which is not a time zone this ` +
        'server can resolve. This server refuses to start on an unusable configured time zone rather than ' +
        'falling back to it silently, because that silence is exactly how a wrong day happens with nothing ' +
        'said. Set FASTMAIL_TIMEZONE to a full IANA zone name such as "Australia/Sydney" or "America/New_York" ' +
        '(or "UTC"), or unset it to use this server\'s own zone.'
      );
    }
    if (rejection === 'shorthand') {
      throw new InvalidInputError(
        `FASTMAIL_TIMEZONE is set to "${echoCallerText(trimmed, ZONE_ECHO_LIMIT)}", a zone abbreviation or alias, ` +
        `not a full IANA zone name. ${SHORTHAND_ZONE_WARNING}. This server refuses to start on an unusable ` +
        'configured time zone rather than falling back to it silently, because that silence is exactly how a ' +
        'wrong day happens with nothing said. Set FASTMAIL_TIMEZONE to a name that contains a region, such as ' +
        '"Australia/Sydney" or "America/New_York" (or "UTC"), or unset it to use this server\'s own zone.'
      );
    }
    return { zone: canonicalZoneName(zoneCandidate) };
  }
  const rejection = zoneRejectionReason(hostZone);
  if (rejection) {
    return {
      zone: 'UTC',
      warning:
        `This server's own time zone ("${echoCallerText(hostZone, ZONE_ECHO_LIMIT)}") is not a full IANA zone ` +
        'name this server will use unqualified (it has no region-qualifying slash and is not "UTC"), so it ' +
        'falls back to UTC rather than writing an ambiguous zone into every calendar read. This also means a ' +
        'bare date-only window bound (list_calendar_events) is now read as a UTC day, not this machine\'s ' +
        'local day. Set FASTMAIL_TIMEZONE to a full IANA zone name (one containing a slash, e.g. ' +
        '"Australia/Sydney") to fix this.'
    };
  }
  return { zone: canonicalZoneName(hostZone) };
}

/**
 * The IANA name to show a caller, resolving `undefined` to whatever the host zone is.
 *
 * Names the zone and nothing else: this string lands in every date rejection, so where the
 * zone came from belongs in the tool description instead. A name ICU cannot resolve is flagged
 * as such rather than printed as though dates were read in it.
 */
export function describeTimezone(zone: string | undefined): string {
  if (!zone) return hostTimezone();
  if (isUsableTimezone(zone)) return zone;
  const echoed = echoCallerText(zone, ZONE_ECHO_LIMIT);
  return `"${echoed}" (the configured time zone, which is not a time zone this server can resolve)`;
}

/**
 * The UTC offset an IANA zone is at, at one instant, in milliseconds.
 *
 * Read by formatting the instant in the zone and treating the printed wall clock as UTC; the
 * difference IS the offset. No tz database ships here, so ICU is the only source.
 *
 * `undefined` means the host zone. A name ICU cannot resolve THROWS rather than falling back
 * to the host offset, which would turn a caller bug into a silently wrong time: gate an
 * untrusted name with `isUsableTimezone` first.
 *
 * The formatter is cached per zone because `src/vtimezone.ts` (#166) makes hundreds of these
 * calls per `generateVTimezone`.
 */
// Keyed on the canonical name (`undefined` for the host zone), and only for a zone that
// resolves, so stored TZID spellings cannot grow it past the set of real zones.
const zoneOffsetFormatterCache = new Map<string | undefined, Intl.DateTimeFormat>();

export function zoneOffsetFormatterCacheSize(): number {
  return zoneOffsetFormatterCache.size;
}

function zoneOffsetFormatterFor(zone: string | undefined): Intl.DateTimeFormat | null {
  const key = zone === undefined ? undefined : canonicalZoneName(zone);
  const cached = zoneOffsetFormatterCache.get(key);
  if (cached !== undefined) return cached;
  let formatter: Intl.DateTimeFormat;
  try {
    formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: zone,
      hour12: false,
      // ERA IS REQUESTED BECAUSE THE YEAR IS READ BACK: without it `Intl` prints the
      // era-relative year (proleptic year 0 formats as "1"), and a bound near year 0 lands
      // on the wrong day. The mapping below puts a BC year back on the proleptic scale.
      era: 'short',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    });
  } catch {
    return null;
  }
  zoneOffsetFormatterCache.set(key, formatter);
  return formatter;
}

export function zoneOffsetMsAt(utcMsInput: number, zone: string | undefined): number {
  // Floored to a whole second: `formatToParts` reads whole seconds, so subtracting an
  // unfloored input would leak the sub-second remainder into the returned offset.
  const utcMs = Math.floor(utcMsInput / 1000) * 1000;
  const formatter = zoneOffsetFormatterFor(zone);
  if (!formatter) {
    throw new Error('zoneOffsetMsAt was given a time zone name ICU cannot resolve; check it with isUsableTimezone first.');
  }
  const parts = formatter.formatToParts(new Date(utcMs));
  const get = (type: string) => Number(parts.find(p => p.type === type)?.value);
  // ISO 8601 / proleptic Gregorian has a year 0; the BC/AD scale does not. 1 BC IS year 0,
  // 2 BC is year -1, so the mapping is `1 - n`.
  const era = parts.find(p => p.type === 'era')?.value ?? '';
  const rawYear = get('year');
  const year = /^b/i.test(era) ? 1 - rawYear : rawYear;
  // Intl can render midnight as hour 24 in some engines; the same normalisation toLocalIso
  // carries, for the same reason.
  const asIfUtc = utcMsFromComponents(year, get('month'), get('day'), get('hour') % 24, get('minute'), get('second'));
  return asIfUtc - utcMs;
}

// A whole Gregorian cycle: 400 years is exactly 146097 days, leap rules included.
export const GREGORIAN_CYCLE_YEARS = 400;
const GREGORIAN_CYCLE_MS = 146097 * 24 * 60 * 60 * 1000;

/**
 * `Date.UTC` without its legacy two-digit-year mapping.
 *
 * `Date.UTC(26, 7, 12)` is the year 1926, not the year 26. Shifting by one whole Gregorian
 * cycle steps over the mapping and back without disturbing leap days.
 */
export function utcMsFromComponents(y: number, mo: number, d: number, h: number, mi: number, s: number): number {
  if (y >= 0 && y <= 99) {
    return Date.UTC(y + GREGORIAN_CYCLE_YEARS, mo - 1, d, h, mi, s) - GREGORIAN_CYCLE_MS;
  }
  return Date.UTC(y, mo - 1, d, h, mi, s);
}

// Far enough either side of a wall clock to bracket the instant it names: no IANA zone has
// ever been more than 14 hours from UTC, so the answer is always within a day of the naive
// reading. Wide enough to see a nearby transition, narrow enough that it can only see one.
const OFFSET_SAMPLE_SPAN_MS = 24 * 60 * 60 * 1000;

/**
 * The UTC instant a wall clock names in a zone.
 *
 * The offset has to be sampled at the instant being solved for, so the offsets a day either
 * side are read first and a candidate is CHECKED against the offset in force where it lands.
 * Both awkward cases follow RFC 5545's and Temporal's `compatible` disambiguation:
 *
 *   REPEATED (clocks go back): the EARLIER of the two instants.
 *   SKIPPED (clocks go forward): FORWARD BY THE LENGTH OF THE GAP. Resolving it backward
 *     makes a window's exclusive end drop the last hour of the day in a zone whose transition
 *     is at midnight (America/Santiago, America/Havana).
 *
 * Neither case is refused: a day with a transition in it still has an answer.
 */
function wallClockToUtcMs(y: number, mo: number, d: number, h: number, mi: number, s: number, zone: string | undefined): number {
  const naive = utcMsFromComponents(y, mo, d, h, mi, s);
  const before = zoneOffsetMsAt(naive - OFFSET_SAMPLE_SPAN_MS, zone);
  const after = zoneOffsetMsAt(naive + OFFSET_SAMPLE_SPAN_MS, zone);
  if (before === after) return naive - before;

  // Checked first, so a repeated hour resolves to the earlier instant.
  const early = naive - before;
  if (zoneOffsetMsAt(early, zone) === before) return early;
  const late = naive - after;
  if (zoneOffsetMsAt(late, zone) === after) return late;
  // Skipped: `early` is the wall clock shifted forward by the length of the gap.
  return early;
}

function toUtcIso(ms: number): string {
  return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** Resolve one window bound, adding `dayOffset` whole local days to a date-only value. */
function resolveWindowBound(
  value: unknown,
  paramName: string,
  zone: string | undefined,
  dayOffset: 0 | 1,
  formats: string = acceptedWindowFormats(describeTimezone(zone)),
): string | undefined {
  if (value === undefined || value === null) return undefined;
  const { trimmed, kind } = classifyDateValue(value, paramName, formats);

  if (kind === 'zoned-datetime') {
    const parsed = new Date(trimmed);
    if (Number.isNaN(parsed.getTime())) {
      throw new InvalidInputError(
        `${paramName} is not a valid date: "${echoDate(trimmed)}". ${formats}`,
      );
    }
    return toUtcIso(parsed.getTime());
  }

  if (kind === 'date') {
    const [y, mo, d] = trimmed.split('-').map(Number);
    // Advanced in LOCAL days, not by adding 24 hours: a DST day is 23 or 25 hours long.
    return toUtcIso(wallClockToUtcMs(y, mo, d + dayOffset, 0, 0, 0, zone));
  }

  const m = LOCAL_DATETIME_PATTERN.exec(trimmed);
  if (!m || !isWallClockInRange(Number(m[4]), Number(m[5]), Number(m[6] ?? 0))) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${formats}`,
    );
  }
  // `dayOffset` does NOT apply: a wall-clock datetime names a time of day, not a day.
  return toUtcIso(wallClockToUtcMs(Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] ?? 0), zone));
}

/**
 * Whether a wall clock's components name a time that exists.
 *
 * `Date.UTC` ROLLS an out-of-range component instead of rejecting it, so without this the
 * calendar pair would accept what `coerceUtcDate` and `create_calendar_event` reject.
 *
 * `24:00:00` is deliberately allowed, because `new Date()` accepts it as the end of the day
 * and the UTC coercion takes it; it rolls into the following midnight, which is what it names.
 */
function isWallClockInRange(h: number, mi: number, s: number): boolean {
  if (h === 24) return mi === 0 && s === 0;
  return h <= 23 && mi <= 59 && s <= 59;
}

/**
 * The UTC instant a value parsed out of an iCalendar payload names, in milliseconds, for
 * ORDERING two of them against each other. `NaN` when there is nothing to order on.
 *
 * `formatICalDate` drops the TZID, so a zoned event arrives bare while a UTC one keeps its
 * `Z`; comparing those as STRINGS misorders them on any non-UTC account. A bare value is
 * therefore read through the configured zone, as the window coercions read it.
 *
 * Best-effort, not a validation: it orders server data, so an unreadable value returns NaN
 * rather than throwing, and an out-of-range component is left to roll.
 */
export function resolveCalendarInstantMs(value: string | undefined, zone: string | undefined): number {
  if (typeof value !== 'string') return NaN;
  const trimmed = value.trim();
  if (!trimmed) return NaN;
  if (ZONE_DESIGNATOR_PATTERN.test(trimmed)) return Date.parse(trimmed);
  if (DATE_ONLY_PATTERN.test(trimmed)) {
    const [y, mo, d] = trimmed.split('-').map(Number);
    // Local midnight, the instant a date-only window bound resolves to, so both share a scale.
    return wallClockToUtcMs(y, mo, d, 0, 0, 0, zone);
  }
  const m = LOCAL_DATETIME_PATTERN.exec(trimmed);
  if (!m) return NaN;
  return wallClockToUtcMs(Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] ?? 0), zone);
}

/** The INCLUSIVE start of a calendar window: a date-only value is local midnight that day. */
export function coerceCalendarWindowStart(value: unknown, paramName: string, zone?: string): string | undefined {
  return resolveWindowBound(value, paramName, zone, 0);
}

/** The EXCLUSIVE end of a calendar window: a date-only value is local midnight the NEXT day. */
export function coerceCalendarWindowEnd(value: unknown, paramName: string, zone?: string): string | undefined {
  return resolveWindowBound(value, paramName, zone, 1);
}

/**
 * The instant local midnight at the START OF TODAY resolves to, for a caller that named no
 * window at all.
 *
 * Resolved through the same helpers as a date-only `startDate`, so the default window starts
 * on the same day as passing today's date would; do not re-derive either step here.
 */
export function startOfLocalDayUtcIso(nowMs: number, zone?: string): string {
  const local = new Date(nowMs + zoneOffsetMsAt(nowMs, zone));
  return toUtcIso(wallClockToUtcMs(
    local.getUTCFullYear(), local.getUTCMonth() + 1, local.getUTCDate(), 0, 0, 0, zone,
  ));
}

// The pagination offset shared by the list/search tools (JMAP `position`, RFC 8620
// section 5.5). A string "40" is accepted, but a shape that would page somewhere unmeant
// is REJECTED rather than repaired, since a wrong offset silently skips messages: a
// NEGATIVE value (JMAP reads it from the END; `ascending` is the way to do that), a
// fraction, or a non-safe integer. Blank shapes mean "start at the first result".
const POSITION_HINT =
  'Pass a whole number of results to skip (0 or greater), e.g. position:20 for the second page of a limit:20 listing.';
const POSITION_ECHO_LIMIT = 40;

export function coercePosition(value: unknown, paramName = 'position'): number | undefined {
  if (value === undefined || value === null) return undefined;

  let n: number;
  if (typeof value === 'number') {
    n = value;
  } else if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    n = Number(trimmed);
  } else {
    throw new InvalidInputError(
      `${paramName} must be a number, not ${Array.isArray(value) ? 'an array' : `a ${typeof value}`}. ${POSITION_HINT}`,
    );
  }

  if (!Number.isSafeInteger(n)) {
    throw new InvalidInputError(
      `${paramName} is not a whole number: "${echoPosition(value)}". ${POSITION_HINT}`,
    );
  }
  if (n < 0) {
    throw new InvalidInputError(
      `${paramName} cannot be negative: "${echoPosition(value)}". It is an offset from the START of the results; to read from the oldest end pass ascending:true. ${POSITION_HINT}`,
    );
  }
  return n;
}

function echoPosition(value: unknown): string {
  const text = String(value);
  return text.length > POSITION_ECHO_LIMIT ? `${text.slice(0, POSITION_ECHO_LIMIT)}...` : text;
}

// Clamp a caller-supplied limit into [1, max], dropping any fractional part (a JMAP limit
// is an UnsignedInt). The `|| fallback` guards NaN, which JMAP serializes as
// `"limit": null`, an unbounded query, and a value that truncates to 0. Never throws, unlike
// the other coercers: a caller who fat-fingers a limit wants results, not an error.
export function clampLimit(value: unknown, fallback: number, max: number): number {
  return Math.min(Math.max(Math.trunc(Number(value)) || fallback, 1), max);
}

function acceptedDateFormats(zoneLabel: string): string {
  return `Accepted: a date such as 2026-07-20 (read as midnight at the start of that day in ${zoneLabel}), ` +
    'or a full datetime such as 2026-07-20T14:30:00Z or 2026-07-20T14:30:00+01:00 (taken exactly as written). ' +
    `A datetime with no Z and no offset, such as 2026-07-20T14:30:00, is read as ${zoneLabel} local time.`;
}

// Loud-reject a settable string field that was provided but is blank or null. Call it only
// for fields that are present, so omitting a field stays distinct from blanking it.
export function requireNonEmpty(value: unknown, fieldName: string, hint = 'omit the field to leave it unchanged'): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new InvalidInputError(`${fieldName} cannot be empty; ${hint}`);
  }
  return value.trim();
}

// Validate a clearFields list: every entry must be in the allowed set, and no
// entry may also appear as a settable param (can't both set and clear a field).
// No-op when clearFields is empty/undefined.
export function validateClearFields(clearFields: string[] | undefined, allowed: ReadonlySet<string>, provided: ReadonlySet<string>): void {
  if (!clearFields || clearFields.length === 0) return;
  for (const field of clearFields) {
    if (!allowed.has(field)) {
      throw new InvalidInputError(`Cannot clear "${field}"; clearable fields are: ${[...allowed].join(', ')}`);
    }
    if (provided.has(field)) {
      throw new InvalidInputError(`cannot both set and clear ${field}; pass it as a value or in clearFields, not both`);
    }
  }
}

// Parse an RFC 5322 "Display Name <email>" recipient string into a JMAP EmailAddress.
// A pragmatic parse, not the full RFC grammar. Input is assumed non-empty.
export function parseAddress(input: string): { name?: string; email: string } {
  const trimmed = String(input).trim();
  const open = trimmed.lastIndexOf('<');
  const close = trimmed.lastIndexOf('>');
  if (open !== -1 && close > open) {
    const email = trimmed.slice(open + 1, close).trim();
    let name = trimmed.slice(0, open).trim();
    if (name.length >= 2 && name.startsWith('"') && name.endsWith('"')) {
      name = name.slice(1, -1).trim();
    }
    return name ? { name, email } : { email };
  }
  return { email: trimmed };
}

// One thing-to-attach spec as it arrives from a tool call (before path confinement,
// gating and upload/resolution, which happen in jmap-client.ts).
//
// THREE SOURCES, exactly one per item, each gated separately (see uploadAttachments):
//
//   path                    a local file, read off disk and uploaded (FASTMAIL_ATTACH_DIR)
//   blobId                  content already in the account's blob store
//   emailId + attachmentId  a part of an existing message, resolved to its blob
//
// Every source field is optional HERE; coerceAttachments, the only producer, guarantees
// exactly one is set.
export interface AttachmentSpec {
  path?: string;
  blobId?: string;
  emailId?: string;
  attachmentId?: string;
  name?: string;
  contentType?: string;
  // The Content-ID an html body references this file by. Stored in CANONICAL form (see
  // stripCidSpelling), so everything downstream compares one value per identifier.
  cid?: string;
}

// The keys that NAME the bytes. Exactly one source per item; `emailId` and `attachmentId`
// are one source spelled in two keys, so they are required together.
const ATTACHMENT_SOURCE_KEYS = ['path', 'blobId', 'emailId', 'attachmentId'] as const;
// The keys that DESCRIBE whatever the source named. Valid on every source.
const ATTACHMENT_COMMON_KEYS = ['name', 'contentType', 'cid'] as const;
const ATTACHMENT_KEYS = new Set<string>([...ATTACHMENT_SOURCE_KEYS, ...ATTACHMENT_COMMON_KEYS]);

const ATTACHMENT_ITEM_SHAPE = '{ path | blobId | emailId+attachmentId, name?, contentType?, cid? }';

const ATTACHMENT_SOURCE_RULE =
  "Give exactly one source per item: 'path' (a local file), 'blobId' (content already in " +
  "the account), or 'emailId' + 'attachmentId' together (a part of an existing message).";

// `null` counts as ABSENT: a lenient client emits null for every declared key it has
// nothing to say about, and would otherwise always name four sources.
function namedAttachmentSourceKeys(obj: Record<string, unknown>): string[] {
  return ATTACHMENT_SOURCE_KEYS.filter((k) => obj[k] !== undefined && obj[k] !== null);
}

// Returns the trimmed value, so a stray space does not read as "not found".
function requireAttachmentString(obj: Record<string, unknown>, key: string, index: number, hint: string): string {
  const value = obj[key];
  if (typeof value !== 'string' || value.trim() === '') {
    throw new McpError(ErrorCode.InvalidParams, `attachments[${index}] is missing a non-empty '${key}'; ${hint}`);
  }
  return value.trim();
}

// Coerce the `attachments` tool param into AttachmentSpec[] | undefined. Per element it
// REJECTS, naming the index, rather than silently dropping: assertKnownParams is top-level
// only, so this is the sole guard on the item shape. A bare string is rejected rather than
// guessed as a path. The exactly-one-source rule is also what refuses a key from a source
// the item did not choose (`{ blobId, attachmentId }`); there is no second allowlist.
export function coerceAttachments(value: unknown): AttachmentSpec[] | undefined {
  if (value === undefined || value === null) return undefined;

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    try {
      arr = JSON.parse(trimmed);
    } catch {
      throw new McpError(ErrorCode.InvalidParams, `attachments must be an array of ${ATTACHMENT_ITEM_SHAPE} objects.`);
    }
  }

  if (!Array.isArray(arr)) {
    throw new McpError(ErrorCode.InvalidParams, `attachments must be an array of ${ATTACHMENT_ITEM_SHAPE} objects.`);
  }

  const specs: AttachmentSpec[] = [];
  for (let i = 0; i < arr.length; i++) {
    let item: unknown = arr[i];
    if (typeof item === 'string') {
      const t = item.trim();
      if (t.startsWith('{') && t.endsWith('}')) {
        try {
          item = JSON.parse(t);
        } catch {
          throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] is a string that isn't valid JSON; pass an object with a path.`);
        }
      } else {
        throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object naming a source, not a bare string. ${ATTACHMENT_SOURCE_RULE}`);
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object shaped ${ATTACHMENT_ITEM_SHAPE}.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter(k => !ATTACHMENT_KEYS.has(k));
    if (unknownKeys.length > 0) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...ATTACHMENT_KEYS].join(', ')}`);
    }

    // Pick the source BEFORE validating anything else, so a caller that named two sources
    // is told that rather than that the first one is malformed.
    const named = namedAttachmentSourceKeys(obj);
    const namesMessagePart = named.includes('emailId') || named.includes('attachmentId');
    // Count SOURCES, not keys: emailId+attachmentId is one source in two keys.
    const distinctSources =
      (named.includes('path') ? 1 : 0) + (named.includes('blobId') ? 1 : 0) + (namesMessagePart ? 1 : 0);
    if (named.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] names no source. ${ATTACHMENT_SOURCE_RULE}`);
    }
    if (distinctSources > 1) {
      throw new McpError(
        ErrorCode.InvalidParams,
        `attachments[${i}] names more than one source (${named.join(', ')}). ${ATTACHMENT_SOURCE_RULE}`,
      );
    }

    const spec: AttachmentSpec = {};
    if (namesMessagePart) {
      spec.emailId = requireAttachmentString(obj, 'emailId', i, "a message part is named by 'emailId' AND 'attachmentId' together.");
      spec.attachmentId = requireAttachmentString(obj, 'attachmentId', i, "a message part is named by 'emailId' AND 'attachmentId' together.");
    } else if (named[0] === 'blobId') {
      spec.blobId = requireAttachmentString(obj, 'blobId', i, 'give the blobId of content already in the account.');
      // A blob carries no filename to default to, and inventing one would put it on
      // outgoing mail.
      if (typeof obj.name !== 'string' || obj.name.trim() === '') {
        throw new McpError(
          ErrorCode.InvalidParams,
          `attachments[${i}] gives a blobId but no 'name'. A stored blob carries no filename, so name the file recipients will see.`,
        );
      }
    } else {
      spec.path = requireAttachmentString(obj, 'path', i, 'give the file to attach.');
    }

    if (obj.name !== undefined) {
      if (typeof obj.name !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].name must be a string.`);
      // A BLANK name reads as absent, so the source's documented default applies. Done here
      // because consumers use `??`, and `'' ?? info.name` is `''`.
      const trimmed = obj.name.trim();
      if (trimmed) spec.name = trimmed;
    }
    if (obj.contentType !== undefined) {
      if (typeof obj.contentType !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].contentType must be a string.`);
      spec.contentType = obj.contentType;
    }
    if (obj.cid !== undefined) {
      // Normalize before validating, then keep the NORMALIZED value: `cid:logo` and `<logo>`
      // are one identifier, and collision detection has to see them as one.
      const raw = typeof obj.cid === 'string' ? obj.cid : '';
      const canonical = stripCidSpelling(raw);
      if (!isAuthorableCid(canonical)) {
        throw new McpError(ErrorCode.InvalidParams, rejectUnusableCid(i, obj.cid));
      }
      spec.cid = canonical;
    }
    specs.push(spec);
  }
  return specs;
}

// One calendar-event attendee as it arrives from a tool call, before
// validateAttendeeEmail vets the address and the iCal ATTENDEE line is built.
export interface ParticipantSpec {
  email: string;
  name?: string;
}

const PARTICIPANT_KEYS = new Set(['email', 'name']);

const PARTICIPANT_ITEM_SHAPE = '{ email, name? }';

// Coerce the `participants` tool param into ParticipantSpec[] | undefined, the same way
// coerceAttachments does, and the sole guard on the item shape (the SDK does not enforce
// inputSchema). A comma-joined string is NOT split: an item is an object.
//
// A BARE STRING element is read as the address, unlike coerceAttachments: the address is a
// participant's only required key, as for `to`/`cc`/`bcc`. The address is NOT vetted here;
// validateAttendeeEmail (src/caldav-client.ts) owns those rules.
//
// Objects are built fresh from the validated keys, so a key added to ParticipantSpec later
// has to be handled here explicitly.
export function coerceParticipants(value: unknown): ParticipantSpec[] | undefined {
  if (value === undefined || value === null) return undefined;

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    // A blank string reads as "not supplied", NOT the empty list: on update_calendar_event
    // an empty array removes every attendee.
    if (!trimmed) return undefined;
    try {
      arr = JSON.parse(trimmed);
    } catch {
      throw new InvalidInputError(`participants must be an array of ${PARTICIPANT_ITEM_SHAPE} objects.`);
    }
  }

  if (!Array.isArray(arr)) {
    throw new InvalidInputError(`participants must be an array of ${PARTICIPANT_ITEM_SHAPE} objects.`);
  }

  const specs: ParticipantSpec[] = [];
  for (let i = 0; i < arr.length; i++) {
    let item: unknown = arr[i];
    if (typeof item === 'string') {
      const t = item.trim();
      if (t.startsWith('{') && t.endsWith('}')) {
        try {
          item = JSON.parse(t);
        } catch {
          throw new InvalidInputError(`participants[${i}] is a string that isn't valid JSON; pass an email address or an object with an email.`);
        }
      } else {
        if (!t) {
          throw new InvalidInputError(`participants[${i}] is an empty string; pass an email address or an object with an email.`);
        }
        specs.push({ email: t });
        continue;
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new InvalidInputError(`participants[${i}] must be an email address or an object with an email.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter(k => !PARTICIPANT_KEYS.has(k));
    if (unknownKeys.length > 0) {
      throw new InvalidInputError(`participants[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...PARTICIPANT_KEYS].join(', ')}`);
    }
    if (typeof obj.email !== 'string' || obj.email.trim() === '') {
      throw new InvalidInputError(`participants[${i}] is missing a non-empty 'email'.`);
    }
    const spec: ParticipantSpec = { email: obj.email.trim() };
    if (obj.name !== undefined) {
      if (typeof obj.name !== 'string') {
        throw new InvalidInputError(`participants[${i}].name must be a string.`);
      }
      spec.name = obj.name;
    }
    specs.push(spec);
  }
  return specs;
}

// ---------- contact write inputs ----------

/** One `emails` entry of a contact write, after coercion. */
export interface ContactEmailSpec {
  address: string;
  label?: string;
}

/** One `phones` entry of a contact write, after coercion. */
export interface ContactPhoneSpec {
  number: string;
  label?: string;
}

/** One `addresses` entry of a contact write, after coercion. */
export interface ContactAddressSpec {
  full: string;
  label?: string;
}

/** The `name` parameter of a contact write, after coercion (a bare string becomes `full`). */
export interface ContactNameSpec {
  given?: string;
  surname?: string;
  full?: string;
}

const CONTACT_NAME_KEYS = new Set(['given', 'surname', 'full']);
const CONTACT_NAME_SHAPE = '{ given?, surname?, full? }';

/**
 * Coerce one of the contact entry-array parameters (`emails`, `phones`, `addresses`) into a
 * validated array of fresh objects. These values are copied into a `ContactCard/set` patch,
 * and the SDK does not enforce `inputSchema`, so: unknown keys are rejected, every known key
 * is TYPE-CHECKED (an allowlist alone passes `{label: []}`), and the output is a FRESH LITERAL,
 * never a spread, so a new key is a conscious edit here.
 *
 * A blank string reads as "not supplied", never the empty array, which these parameters
 * reject. `allowBareString` is off for `addresses`, which has no single scalar reading.
 *
 * Duplicates are REJECTED: on `emails`/`phones` a repeat cannot be matched against the stored
 * card twice, and would surface as an unknown addition.
 */
function coerceContactEntries<T extends Record<string, any>>(
  value: unknown,
  paramName: string,
  keyField: 'address' | 'number' | 'full',
  allowBareString: boolean,
): T[] | undefined {
  if (value === undefined || value === null) return undefined;

  const keys = new Set([keyField, 'label']);
  const itemShape = `{ ${keyField}, label? }`;
  const bareNote = allowBareString ? ` (or a bare ${keyField} string)` : '';

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    try {
      arr = JSON.parse(trimmed);
    } catch {
      throw new InvalidInputError(`${paramName} must be an array of ${itemShape} objects${bareNote}.`);
    }
  }

  if (!Array.isArray(arr)) {
    throw new InvalidInputError(`${paramName} must be an array of ${itemShape} objects${bareNote}.`);
  }

  const specs: T[] = [];
  const seen = new Map<string, number>();
  for (let i = 0; i < arr.length; i++) {
    let item: unknown = arr[i];
    if (typeof item === 'string') {
      const t = item.trim();
      if (t.startsWith('{') && t.endsWith('}')) {
        try {
          item = JSON.parse(t);
        } catch {
          throw new InvalidInputError(
            `${paramName}[${i}] is a string that isn't valid JSON; pass an object shaped ${itemShape}${bareNote}.`,
          );
        }
      } else if (allowBareString) {
        if (!t) {
          throw new InvalidInputError(`${paramName}[${i}] is an empty string; pass a ${keyField} or an object shaped ${itemShape}.`);
        }
        item = { [keyField]: t };
      } else {
        throw new InvalidInputError(
          `${paramName}[${i}] must be an object shaped ${itemShape}, not a bare string.`,
        );
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new InvalidInputError(`${paramName}[${i}] must be an object shaped ${itemShape}${bareNote}.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter((k) => !keys.has(k));
    if (unknownKeys.length > 0) {
      throw new InvalidInputError(
        `${paramName}[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...keys].join(', ')}`,
      );
    }
    if (typeof obj[keyField] !== 'string') {
      throw new InvalidInputError(
        `${paramName}[${i}].${keyField} must be a string, not ${
          Array.isArray(obj[keyField]) ? 'an array' : `a ${typeof obj[keyField]}`
        }.`,
      );
    }
    const primary = (obj[keyField] as string).trim();
    if (!primary) {
      throw new InvalidInputError(`${paramName}[${i}] is missing a non-empty '${keyField}'.`);
    }
    const firstAt = seen.get(primary);
    if (firstAt !== undefined) {
      throw new InvalidInputError(
        `${paramName}[${i}] repeats the ${keyField} already given at ${paramName}[${firstAt}]: "${primary}". ` +
          `List each ${keyField} once.`,
      );
    }
    seen.set(primary, i);

    const spec: Record<string, any> = { [keyField]: primary };
    if (obj.label !== undefined) {
      if (typeof obj.label !== 'string') {
        throw new InvalidInputError(
          `${paramName}[${i}].label must be a string, not ${
            Array.isArray(obj.label) ? 'an array' : `a ${typeof obj.label}`
          }.`,
        );
      }
      // A blank label would write something that reads as absent. Deliberately NOT a way to
      // remove a label: this tool has no clearing mechanism for one.
      if (obj.label.trim() === '') {
        throw new InvalidInputError(
          `${paramName}[${i}].label cannot be empty; omit it to leave the entry's label unchanged.`,
        );
      }
      spec.label = obj.label;
    }
    specs.push(spec as T);
  }
  return specs;
}

export function coerceContactEmails(value: unknown): ContactEmailSpec[] | undefined {
  return coerceContactEntries<ContactEmailSpec>(value, 'emails', 'address', true);
}

export function coerceContactPhones(value: unknown): ContactPhoneSpec[] | undefined {
  return coerceContactEntries<ContactPhoneSpec>(value, 'phones', 'number', true);
}

export function coerceContactAddresses(value: unknown): ContactAddressSpec[] | undefined {
  return coerceContactEntries<ContactAddressSpec>(value, 'addresses', 'full', false);
}

/**
 * Coerce the `name` parameter of a contact write; a bare string is the full name. Same
 * discipline as coerceContactEntries. A blank value is rejected rather than read as "clear the
 * name": `name` is not clearable, and dropping it would look like a removal that never happened.
 */
export function coerceContactName(value: unknown): ContactNameSpec | undefined {
  if (value === undefined || value === null) return undefined;

  let raw: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
      try {
        raw = JSON.parse(trimmed);
      } catch {
        throw new InvalidInputError(`name must be a full-name string or an object shaped ${CONTACT_NAME_SHAPE}.`);
      }
    } else {
      if (!trimmed) {
        throw new InvalidInputError('name cannot be empty; omit it to leave the stored name unchanged.');
      }
      return { full: trimmed };
    }
  }

  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    throw new InvalidInputError(`name must be a full-name string or an object shaped ${CONTACT_NAME_SHAPE}.`);
  }
  const obj = raw as Record<string, unknown>;
  const unknownKeys = Object.keys(obj).filter((k) => !CONTACT_NAME_KEYS.has(k));
  if (unknownKeys.length > 0) {
    throw new InvalidInputError(
      `name has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...CONTACT_NAME_KEYS].join(', ')}`,
    );
  }

  const spec: ContactNameSpec = {};
  for (const key of CONTACT_NAME_KEYS) {
    const supplied = obj[key];
    if (supplied === undefined) continue;
    if (typeof supplied !== 'string') {
      throw new InvalidInputError(
        `name.${key} must be a string, not ${Array.isArray(supplied) ? 'an array' : `a ${typeof supplied}`}.`,
      );
    }
    const trimmed = supplied.trim();
    if (!trimmed) {
      throw new InvalidInputError(`name.${key} cannot be empty; omit it to leave that part of the name unchanged.`);
    }
    (spec as Record<string, string>)[key] = trimmed;
  }
  if (Object.keys(spec).length === 0) {
    throw new InvalidInputError(`name must set at least one of ${[...CONTACT_NAME_KEYS].join(', ')}.`);
  }
  return spec;
}
