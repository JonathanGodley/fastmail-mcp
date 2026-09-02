import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { describePart, isAuthorableCid, stripCidSpelling } from './inline-images.js';
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
// invalid (e.g. a `mailbox` that resolves to nothing, or a label `mailboxIds`
// element that isn't a real id). Thrown from jmap-client.ts (which stays free of
// MCP SDK types); the index boundary maps every InvalidInputError to
// McpError(InvalidParams), mirroring PathAccessError. instanceof is the
// discriminator. Like every other branch of that catch — the McpError rethrow, the
// PathAccessError mapping and the generic InternalError wrap — the message goes
// through redactBearerTokens. Redaction there is unconditional and has no exemptions,
// which is what lets the audit ("no unredacted error text reaches tool output") be a
// grep anyone can run instead of a claim resting on a per-error-class exemption list.
// These messages in particular reflect caller input and mailbox names, so a
// token-shaped echo is a real shape to scrub (it is NOT what makes the reflected-input
// oracle acceptable; see docs/security-model.md).
export class InvalidInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InvalidInputError';
  }
}

// Some MCP clients (e.g. Claude Cowork as of 2026-04-08, issue #54) stringify
// structured params before dispatch. These helpers coerce such values back to
// their expected shapes so the handlers work against both strict and lenient clients.

// Defense-in-depth: scrub credential-shaped substrings from any string that
// might be reflected back to the MCP caller (e.g. a JMAP error message). This
// is intentionally narrow — provider error messages are useful for the LLM to
// recover from, so we don't want to over-sanitize.
const BEARER_PATTERN = /Bearer\s+\S+/gi;
// The CalDAV path authenticates with HTTP Basic, so a reflected header or a
// tsdav error carries the base64 credential blob — redact that shape too.
const BASIC_PATTERN = /Basic\s+[A-Za-z0-9+/=]+/gi;
// Fastmail token shape. The charset is `[\w-]`, not `[A-Za-z0-9-]`: under the
// narrower class an underscore inside the token ends the match early, and with
// fewer than 20 characters before it the `{20,}` quantifier fails outright, so
// the whole token would pass through in clear.
const FASTMAIL_TOKEN_PATTERN = /fmu\d+-[\w-]{20,}/g;

// Exact secret values registered at startup (API token, CalDAV password, and
// self-hosted tokens carrying neither a `Bearer` prefix nor the `fmu` shape).
// Value-based redaction catches the credentials the patterns above cannot see.
// Populated by registerSecret(); never logged.
const KNOWN_SECRETS = new Set<string>();

// Register a literal secret value so redactBearerTokens scrubs any exact
// occurrence of it. Values under 8 characters are ignored — an over-broad
// match would mangle legitimate output for no security gain.
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
 * The criterion this exists to make checkable: **any untrusted value interpolated into a
 * message a caller reads back — thrown or returned — goes through this, and nothing else.**
 * "Untrusted" is not "attacker-authored"; it is "not written by this server": a
 * caller-supplied id, a mailbox name, a Content-ID, a server-authored set-error
 * description. The server's own sentence around the value is never passed through here —
 * only the value.
 *
 * Two hazards, and they need the two steps in this order:
 *
 * 1. LINE FORGING (the reason this matters in practice). A value carrying CR, LF or
 *    U+2028 splits one message into what reads as several, and the forged lines read to
 *    an agent as further sentences from the server. `describePart` strips those, collapses
 *    space runs, drops bidi overrides, and turns a double quote into a single one so the
 *    value cannot close the quoted span it is rendered inside. It also caps the length, so
 *    one hostile id cannot become the whole error message.
 *
 * 2. CREDENTIAL ECHO (narrow, but a leak rather than a style point). `redactBearerTokens`
 *    is length-sensitive at both ends — FASTMAIL_TOKEN_PATTERN needs 20+ characters after
 *    the `fmu<n>-` prefix, and a registered secret is matched as an exact string — so
 *    running `describePart` FIRST, which truncates at 64 code points, hands the redactor a
 *    string the secret no longer fits in and the surviving prefix goes out verbatim.
 *
 * Hence the order: redact the full value, THEN neutralise and truncate it. Getting it
 * backwards still reads correctly and still passes every line-forging test, which is
 * exactly why the two steps live behind one name instead of at each call site (#131).
 *
 * Not applicable to a structured result item — see `redactedJson` for why redacting a
 * finished JSON document eats its delimiters.
 */
export function describeUntrusted(value: unknown): string {
  const source = typeof value === 'string' ? value : value == null ? '' : String(value);
  return describePart(redactBearerTokens(source));
}

/**
 * Serialise a value to JSON with every string inside it redacted.
 *
 * This is the ONLY safe way to redact a structured result item, and the reason is
 * BEARER_PATTERN: `/Bearer\s+\S+/gi` runs `\S+` to the next whitespace, which over an
 * already-serialised document is the string's own closing quote and the comma after it.
 * Redacting the finished document therefore eats JSON delimiters and the item stops
 * parsing — triggered by nothing more exotic than a mailbox named "Bearer Bonds", and
 * also by a server description that ends in a real token, which loses the caller the
 * whole report exactly when a credential was present.
 *
 * Redacting per value has neither problem: a value has no trailing delimiter for `\S+` to
 * swallow, JSON.stringify escapes whatever the replacer hands back, and a genuine
 * `Bearer <token>` inside a value is still redacted. Prose has no delimiters to protect,
 * so it calls redactBearerTokens directly; anything JSON.stringify touches comes here.
 *
 * The redacting counterpart to toolJson below, and compact for the same reason: it takes no
 * indent argument, so this path cannot pretty-print even by accident.
 */
export function redactedJson(value: any): string {
  return JSON.stringify(value, (_key, v) => (typeof v === 'string' ? redactBearerTokens(v) : v));
}

/**
 * Serialise a tool result payload. THE one seam every JSON result item goes through, across
 * every handler and formatter, so how this server serialises is decided once (#40).
 *
 * Compact, with no option to indent. Every payload here is read by a machine — an MCP client
 * parses it, or a model reads it as data — and neither needs indentation, while the reader
 * pays for every byte of it. On a 25-message list page the indentation alone was ~17% of the
 * response by bytes, carrying no information at all. The saving scales with the number of
 * JSON tokens, so it is largest exactly where the payload is largest: the list and search
 * seams in response-formatters.ts.
 *
 * This applies to a payload embedded in a prose frame too (a list result's summary line, the
 * bulk-operations diagnostic). The prose stays prose; the JSON inside it is still a payload
 * the caller parses, so it is serialised the same way as a payload standing alone. Whitespace
 * is not the thing that makes a result legible — its structure and field names are, and
 * compacting changes neither.
 *
 * Use redactedJson above instead where the values may carry credentials.
 */
export function toolJson(value: unknown): string {
  return JSON.stringify(value);
}

// Every branch TRIMS its elements, so the three ways of expressing the same list agree.
// Without it the branches disagree: the comma-split branch has always trimmed, so
// "e1, e2" arrives clean while ["e1", " e2"] arrives padded, and a padded value then
// reaches the server and comes back as a not-found — a whitespace problem wearing a
// lookup error's clothes. coerceStringArrayStrict relies on this too: it rejects a blank
// element itself, then delegates here for the trim rather than repeating it.
//
// The call sites are email ids, mailbox references, addresses, message-ids, field names, and
// edit_draft's removeAttachments. Whitespace is not meaningful in any of them, with ONE case
// worth naming because it is not obvious: a removeAttachments ref is matched against an
// attachment's own name, and a MIME filename may legally carry surrounding spaces. That
// comparison trims both sides (resolveAttachmentRemovals), so such an attachment stays
// reachable by name; if that ever stops being true, this trim starts hiding it.
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

// coerceStringArray for a parameter that must FAIL CLOSED: a value that is present but
// cannot be coerced (a number, an object, a boolean) is rejected instead of coming back
// as `undefined` and being read as "not supplied". The plain coercer's silent-undefined
// is right where dropping the value only means "field unchanged", and wrong where the
// value NARROWS what the call touches — a dropped scoping argument silently widens the
// query, which is the failure the argument was passed to prevent.
//
// Strict per ELEMENT as well, following coerceAttachments' discipline: a non-string entry
// is rejected BY INDEX rather than passed through `String()`. Without that, `[null]` and
// `[{}]` reach the mailbox matcher as the literal text "null" and "[object Object]" and
// come back as `Mailbox 'null' not found. Use an id, a role...` — a type error wearing a
// typo's error message, which sends the caller re-spelling a value that was never text.
// (The lenient whole-value forms stay: a real array, a JSON-string array, and a
// comma-separated string are all accepted, and only their elements are checked.)
//
// `null` is treated as absent at the TOP level, not as an error, matching
// coerceStringArray: a lenient client that fills every declared key emits `null` for the
// ones it has nothing to say about, and that is a statement of absence rather than an
// unusable value. Inside the array it is an unusable value, and rejects.
export function coerceStringArrayStrict(value: unknown, paramName: string): string[] | undefined {
  if (value === undefined || value === null) return undefined;

  // Unwrap a JSON-string array HERE rather than leaving it to coerceStringArray, so the
  // elements can be type-checked before `.map(String)` erases what they were. A string
  // that is not a JSON array falls through untouched and is comma-split as before.
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
      // An empty or whitespace-only element is rejected for the same reason a non-string one
      // is, and it is the SAME failure by a different door: `['']` is a string, so the type
      // check above passes it, and coerceStringArray's `.filter(Boolean)` only runs on the
      // comma-split branch — so it reaches the downstream lookup as a real value and comes
      // back as "not found" or "unknown mailbox", a type error wearing a lookup error's
      // clothes. The comma-split branch keeps dropping blanks, because there the blank is a
      // separator artefact ("a,,b") rather than something a caller wrote down.
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

// Coerce the four recipient list fields from whatever shape a (possibly lenient)
// client sent into string[] | undefined, so the JMAP client's .map(parseAddress)
// calls never receive a bare string (issue #54). Pass the raw tool args; reads
// only to/cc/bcc/replyTo and returns the coerced quartet.
export function coerceRecipients(args: { to?: unknown; cc?: unknown; bcc?: unknown; replyTo?: unknown }): {
  to?: string[]; cc?: string[]; bcc?: string[]; replyTo?: string[];
} {
  return {
    to: coerceStringArray(args.to),
    cc: coerceStringArray(args.cc),
    bcc: coerceStringArray(args.bcc),
    replyTo: coerceStringArray(args.replyTo),
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

export function coerceBool(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') return value;
  if (value === 'true') return true;
  if (value === 'false') return false;
  return undefined;
}

// JMAP filter conditions take a UTCDate (RFC 8620 §1.4): an RFC 3339 date-time whose
// offset is literally `Z`, e.g. `2026-07-20T00:00:00Z`. A bare `2026-07-20` is valid
// ISO 8601 but the server rejects it with an opaque `invalidArguments` that names no
// argument (#70), so normalise here instead of passing the caller's string through.
//
// The two accepted shapes are matched explicitly and everything else is rejected — this
// is the one place the codebase does NOT coerce leniently, because `new Date()`'s legacy
// fallback parser guesses in ways that would silently move the search window: it reads
// `2026/07/20` and `20 July 2026` as HOST-LOCAL midnight (not the documented UTC
// midnight), and rolls an impossible day like `2026-2-31` over into the next month
// instead of failing. A rejection the caller can read and fix beats a window that is
// quietly off by the host's UTC offset.
const DATE_ONLY_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
const DATE_TIME_PATTERN = /^(\d{4}-\d{2}-\d{2})T.+$/;

// Longest caller value echoed back in a rejection message. Enough to recognise the bad
// value, short enough that a pasted blob doesn't become the error.
const DATE_ECHO_LIMIT = 60;

/**
 * The ONE way this server quotes caller-supplied text back inside an error message.
 *
 * An error message is read by an agent, so any caller text inside it is untrusted content in
 * a trusted-looking channel. Three rules, and they travel together:
 *
 *   TRIM, so the value quoted is the value the coercion actually looked at. Echoing
 *     `"  2027-03-10  "` back at someone whose padding was stripped before validation quotes
 *     a string that is not what was judged.
 *   SCRUB the control characters — and U+2028/U+2029 with them — that would otherwise forge
 *     extra lines in the message. A raw ESC in an echoed argument reaches a terminal intact.
 *   BOUND it, with a VISIBLE truncation marker, so a pasted blob does not become the error
 *     and a reader can tell a cut value from a short one.
 *
 * It lives here, and every echo site calls it, because the three used to disagree: the date
 * coercions echoed raw control characters, the calendar window scrubbed them, and the zone
 * describer scrubbed but cut silently at 40 characters with no marker. Same class of value,
 * same message family, three policies.
 */
export function echoCallerText(value: unknown, limit: number = DATE_ECHO_LIMIT): string {
  const text = typeof value === 'string' ? value : String(value);
  const clean = text.replace(/[\u0000-\u001F\u007F\u2028\u2029]/g, ' ').trim();
  return clean.length > limit ? `${clean.slice(0, limit)}…` : clean;
}

// Normalise a caller-supplied date/datetime into the JMAP UTCDate shape.
//
//   2026-07-20                -> 2026-07-20T00:00:00Z   (midnight UTC on that date)
//   2026-07-20T14:30:00Z      -> 2026-07-20T14:30:00Z
//   2026-07-20T14:30:00+01:00 -> 2026-07-20T13:30:00Z   (offset applied)
//   2026-07-20T14:30:00       -> the same instant in UTC (no zone = host local time)
//
// Anything else is REJECTED with an InvalidInputError naming the parameter and the
// accepted shapes, so the caller never has to guess which argument JMAP disliked. That
// includes an unpadded or slash-separated date, free text, a reduced-precision `2026` /
// `2026-07`, and a day that doesn't exist in its month. An empty/whitespace-only string
// is rejected too rather than treated as "no filter": silently dropping a date bound
// widens the search and reads as "the filter did nothing". Milliseconds are trimmed so
// the emitted value is the canonical seconds-precision form.
export function coerceUtcDate(value: unknown, paramName: string): string | undefined {
  if (value === undefined || value === null) return undefined;
  const { trimmed, kind } = classifyDateValue(value, paramName, acceptedDateFormats());

  // A date-only value is expanded explicitly rather than left to Date's parse so the
  // intent (midnight UTC, never host-local) is visible in the code, not a spec detail.
  const parsed = new Date(kind === 'date' ? `${trimmed}T00:00:00Z` : trimmed);
  if (Number.isNaN(parsed.getTime())) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${acceptedDateFormats()}`,
    );
  }
  return parsed.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

// What SHAPE a caller's date argument is. The distinction the calendar cares about is the
// third one: a datetime carrying no zone designator names a wall clock, not an instant, so
// something has to decide which zone reads it.
type DateValueKind = 'date' | 'local-datetime' | 'zoned-datetime';

// Every accepted datetime ends in `Z` or a numeric offset; anything else is a wall clock.
const ZONE_DESIGNATOR_PATTERN = /(?:Z|[+-]\d{2}:?\d{2})$/i;
// The wall-clock datetimes the calendar resolves itself, captured so the components can be
// placed in a zone. Deliberately stricter than DATE_TIME_PATTERN's `T.+`: a zone-less value
// this does not match is one whose hour and minute cannot be read out, and guessing at it is
// exactly what the strict-parsing rule above exists to prevent.
//
// SHAPE ONLY — the RANGES are checked separately, in `isWallClockInRange`. Matching here
// is not acceptance: `2026-08-12T99:99:99` has a readable hour, minute and second, and every
// one of them is out of range. See that function for why the check is not folded into this
// pattern.
const LOCAL_DATETIME_PATTERN = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?$/;

/**
 * The shared validation behind every date argument: type, emptiness, shape, and a real
 * calendar day. Split out so `coerceUtcDate` (email search bounds, resolved in UTC) and the
 * calendar-window pair below (resolved in the user's zone) reject an identical set of bad
 * values with identical wording, and only DIVERGE where they mean to — on what an accepted
 * value resolves to. `formats` is the accepted-shapes sentence, which differs between them
 * because the two describe a date-only value differently.
 *
 * It does NOT check the TIME components, because it does not read them: `coerceUtcDate` gets
 * that from `new Date()` refusing an out-of-range hour, and the calendar pair has to do it
 * itself, in `isWallClockInRange`. The two therefore agree on `2026-08-12T25:00:00` only for
 * as long as both halves stay — and they did not, once: the calendar pair read the components
 * out with a shape-only pattern and handed them to `Date.UTC`, which rolled `99:99:99` into a
 * window three days wide of the one asked for.
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

  // A day that doesn't exist in its month parses rather than failing (2026-02-31 becomes
  // 2026-03-03), which would silently shift the search window off the dates the caller
  // asked for. Probe the calendar date on its own — probing the whole value wouldn't
  // work, since an offset legitimately moves the UTC date.
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
  // `datePart` is deliberately NOT returned: it exists only for the real-calendar-date probe
  // above, and no caller ever read it.
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
// bound, and the divergence is the whole point rather than an inconsistency to tidy away.
//
//   startDate: 2026-08-12   ->  local midnight on the 12th
//   endDate:   2026-08-12   ->  local midnight on the 13th   (the whole of the 12th)
//   either:    2026-08-12T09:00:00      ->  09:00 local, deliberately
//   either:    2026-08-12T09:00:00Z     ->  exactly what it says; zone rules do not apply
//   either:    2026-08-12T09:00:00+10:00 -> exactly what it says
//
// WHY THIS IS NOT THE UTC RULE. "What is on the 12th?" is a question about the asker's own
// day. Reading it as a UTC day answered a +10:00 user with the 12th 10:00 through the 13th
// 10:00 — so a 08:00 appointment on the 12th, whose own title said "Wednesday 12 Aug 2026",
// fell outside the window and a day with three appointments in it came back holding one.
// Silently, because two events missing look exactly like a quiet morning. Every hour of
// UTC offset is an hour of somebody's day answered from the wrong date, and the further a
// user is from Greenwich the more of their day it is.
//
// A ZONE-LESS DATETIME follows the same rule, for the same reason and one more: it already
// resolved in the host's zone, by accident of how `new Date()` parses a value with no
// designator. That made the machine the server happens to run on part of the answer, with
// nothing in the schema admitting it. Now the zone is the CONFIGURED one and it is stated.
//
// AN EXPLICIT Z OR OFFSET IS NEVER TOUCHED. A caller that named an instant named an
// instant; re-reading it in another zone would be the same class of silent shift.
//
// The email search bounds keep the UTC rule (`coerceUtcDate` above): `before`/`after` name
// an instant to compare a message's `receivedAt` against, and JMAP's UTCDate is that
// instant. A calendar window names DAYS on somebody's wall. Same accepted spellings, same
// rejections, different question — see docs/conventions.md.
//
// The end is EXCLUSIVE, so a date-only end resolves to local midnight of the FOLLOWING day:
// CalDAV's <C:time-range> end is exclusive (RFC 4791 section 9.9), which makes the next
// midnight precisely "through the end of that day", where 23:59:59 would drop the last
// second. Without that, `startDate: 2026-08-12, endDate: 2026-08-12` would be a zero-length
// window returning nothing, which reads as an empty day rather than as a mistake.
//
// `zone` is passed in rather than read from module state so the resolution is injectable —
// a test that only ever exercises the host zone passes under either behaviour, which is how
// the UTC-day reading survived the whole suite. `undefined` means the host zone.

function acceptedWindowFormats(zoneLabel: string): string {
  return `Accepted: a date such as 2026-08-12 (read as a whole day in ${zoneLabel}), or a full datetime such as ` +
    '2026-08-12T14:30:00Z or 2026-08-12T14:30:00+10:00 (taken exactly as written). A datetime with no ' +
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

// A misconfigured zone name is echoed back to the caller, so it is bounded and stripped of
// the control characters that would forge extra lines in an error message. Exported so
// validateCallerTimezone's own rejections (#157) use the identical bound rather than a second
// number that could quietly drift from this one.
export const ZONE_ECHO_LIMIT = 40;

const zoneCanonicalizationCache = new Map<string, string>();

/**
 * ICU's canonical spelling for a zone name — `Intl.DateTimeFormat`'s own name for whatever the
 * string resolves to — or the string unchanged when ICU cannot resolve it at all. This is the
 * one canonicalization seam every zone comparison and every zone actually written routes
 * through: `validateCallerTimezone` (the create/update write path, #157) and
 * `resolveUsableTimezone` below both return through here, and so does the read-side comparison
 * in caldav-client.ts's `zoneNamesEqual` (#139). An alias spelling ('NZ', 'US/Pacific', a
 * lowercase 'australia/sydney') therefore canonicalises identically wherever it is checked —
 * without a single seam, the same string could compare equal to itself on write but not on
 * read, or vice versa.
 *
 * Cached by exact input string: `zoneNamesEqual` runs once per event on every calendar list
 * read, and constructing an `Intl.DateTimeFormat` per call is not free.
 *
 * The cache is deliberately unbounded. Its keys are TZID strings out of the account holder's own
 * calendar plus the configured and caller-supplied zone names, so the distinct set is the handful
 * of zones that account actually uses — there is no path by which an untrusted party feeds it
 * unbounded distinct strings. An eviction policy here would cost more than it could ever save.
 */
export function canonicalZoneName(zone: string): string {
  const cached = zoneCanonicalizationCache.get(zone);
  if (cached !== undefined) return cached;
  const resolved = isUsableTimezone(zone)
    ? new Intl.DateTimeFormat('en-US', { timeZone: zone }).resolvedOptions().timeZone
    : zone;
  zoneCanonicalizationCache.set(zone, resolved);
  return resolved;
}

/**
 * The IANA name actually used for a configured zone: `zone` itself, canonicalised through
 * `canonicalZoneName`, when it is set and ICU can resolve it — otherwise the host's own zone.
 * This is the one place that decides which zone wins — `describeTimezone` and every calendar
 * read path call through here rather than re-deriving the fallback, so there is exactly one
 * rule for "which zone is really in force."
 *
 * Canonicalising here (not just on the caller-supplied `timeZone` argument) matters because the
 * configured default is interpolated straight into a written TZID on `create_calendar_event`
 * when no `timeZone` is passed: without this, `FASTMAIL_TIMEZONE=australia/sydney` would write
 * `TZID=australia/sydney` while the same zone arriving as a caller's `timeZone` argument writes
 * the canonical `Australia/Sydney` — two different spellings for the same zone depending on
 * which path set it, which then falsely compare unequal on a later read-modify-write.
 */
export function resolveUsableTimezone(zone: string | undefined): string {
  if (zone && isUsableTimezone(zone)) return canonicalZoneName(zone);
  return hostTimezone();
}

// The three ways a zone candidate can fail the rule shared by validateCallerTimezone (the
// caller-supplied `timeZone` parameter, #157) and resolveConfiguredTimezone (FASTMAIL_TIMEZONE
// and the host zone it falls back to, #157 amendment below). Kept as a closed set of reasons
// rather than a bare boolean so each call site can compose its own framing sentence around
// WHICH way the candidate failed, without re-deriving that classification itself.
type ZoneRejectionReason = 'offset-shaped' | 'unresolvable' | 'shorthand';

// The specific abbreviation that decided the shorthand rule, and why it is dangerous rather
// than merely nonstandard: "EST" resolves through ICU to a real, fixed-offset zone (Panama's,
// with no daylight saving) that is NOT US Eastern. Both places that reject a shorthand zone name
// quote this same sentence, so the warning cannot read differently depending on which path a
// caller happened to hit.
const SHORTHAND_ZONE_WARNING =
  '"EST" resolves to a fixed-offset zone with no daylight saving, not US Eastern, and other ' +
  'abbreviations and aliases ("NZ", "PST", "GMT", "Zulu"...) are just as ambiguous';

/**
 * Classify a zone candidate against the rule validateCallerTimezone and
 * resolveConfiguredTimezone both enforce, or return `null` when it is acceptable. The candidate
 * must already be pre-canonical and have any leading `/` stripped (RFC 5545 §3.2.19) — see
 * validateCallerTimezone's own comment for why: checking a CANONICAL name would let a shorthand
 * like `NZ` (which canonicalises to `Pacific/Auckland`, a slash-bearing name) straight through.
 *
 * Order is part of the contract, not an implementation detail:
 *   1. Offset-shaped strings (`+10:00`, `GMT+10`, ...) are rejected before the ICU check even
 *      runs — ICU resolves several offset spellings as though they were legitimate zone names,
 *      so checking resolvability first would wrongly accept them. See the offset-shape comment
 *      inline below for the exact denylist.
 *   2. ICU resolvability is checked next, so a string ICU cannot resolve AT ALL (`Blah`) gets
 *      "not a time zone this server can resolve" — not the shorthand message, which would wrongly
 *      imply that adding a slash is all that string needed.
 *   3. Only then the slash rule (#157 amendment, decided by "EST" silently resolving to a
 *      fixed-offset Panama zone rather than US Eastern): a zone name must contain a
 *      region-qualifying slash, with the literal name "UTC" (compared case-insensitively) as the
 *      rule's one exception. This rejects every abbreviation and alias that ICU nonetheless
 *      resolves — "EST", "NZ", "PST", "GMT", "Zulu" and similar — even the harmless ones; there
 *      is no safe-list, the caller writes the region name instead ("Pacific/Auckland", not
 *      "NZ"). A slash-bearing alias like "US/Pacific" is unaffected and keeps resolving exactly
 *      as it does today.
 */
function zoneRejectionReason(zoneCandidate: string): ZoneRejectionReason | null {
  // Offset-shape denylist — see point 1 above for why this runs before isUsableTimezone.
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
 * Validate a caller-supplied `timeZone` argument (create_calendar_event / update_calendar_event,
 * fork issue #157) and return the canonical IANA name to write.
 *
 * `isUsableTimezone` alone is the WRONG gate for this. Probed on this host it returns `true`
 * for `"+10:00"`, `"+1000"`, `"+10"` and `"Etc/GMT-10"` — ICU resolves several offset spellings
 * as though they were legitimate zone names. Waving those through here would let a caller write
 * `DTSTART;TZID=+10:00:...` — an offset baked into a TZID, which is the one shape this whole
 * design exists to prevent: unresolvable to this server as a NAME, and mis-indexed for its own
 * time-range queries. So the offset shape is rejected BEFORE the ICU check runs at all.
 *
 * The gate is a DENYLIST of offset spellings, not an allowlist of IANA shapes: `Etc/GMT-10`,
 * `UTC`, `Japan` and `EST5EDT` are all legitimate zone SHAPES an allowlist would reject — though
 * `EST5EDT` is separately rejected below, by the slash rule rather than the offset denylist.
 *
 * ALSO REJECTED (#157 amendment): a zone name that ICU resolves but that has no region-qualifying
 * slash and is not literally "UTC" — an abbreviation or alias such as "EST", "NZ", "PST", "GMT"
 * or "Zulu". "EST" resolving to a fixed-offset Panama zone with no daylight saving, silently and
 * with no way to tell from the input, is the case that decided this: ICU accepting a spelling
 * says nothing about whether the caller meant the zone ICU picked. There is no safe-list of
 * harmless abbreviations (a caller writing "NZ" almost certainly does mean Pacific/Auckland) —
 * the caller writes the region name instead. See `zoneRejectionReason` for the exact rule and
 * why it runs after the ICU check.
 *
 * Fails closed on `null`, empty, or whitespace-only rather than treating any of them as "write
 * floating". The read side emits `timeZone: null` for a genuinely floating event (see
 * `CalendarEvent` in caldav-client.ts, #139), so a read-modify-write caller can echo `null`
 * straight back — silently reinterpreting that as "make this floating" would decide an open
 * design question by accident rather than ask. Omit `timeZone` instead: on `create` an omitted
 * `timeZone` writes the account's configured zone (never floating, see docs/conventions.md); on
 * `update` it leaves whatever the event already has unchanged. This narrowing is deliberate, not
 * incidental to the null rejection: there is no benefit to this server ever writing a zone-less
 * calendar time, so nothing here exists to make that possible (see docs/conventions.md).
 *
 * Not an IANA name and not accepted: a non-IANA name such as a Windows zone id ("AUS Eastern
 * Standard Time") still rejects here even though it is a real zone identifier somewhere — ICU
 * cannot resolve it, so there is nothing to canonicalise it against. That is a genuine
 * round-trip limit, not an oversight; both tools' descriptions say `timeZone` takes an IANA name
 * for this reason.
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
  // RFC 5545 §3.2.19 permits a leading '/' on a TZID (a zone registered by its own creator
  // rather than plain IANA form); normalizeZoneForComparison in caldav-client.ts already strips
  // it for comparison, and the read half (#139) emits it verbatim when a stored event carries
  // it. Strip it here too so a caller echoing a `timeZone` this server just handed back
  // (read-modify-write) is not rejected for a spelling this server itself produced. This has to
  // run before both the offset-shape check and the canonicalisation call below: ICU throws
  // outright on a leading slash (`new Intl.DateTimeFormat('en-US', { timeZone:
  // '/Australia/Sydney' })` throws), so an unstripped value never reaches isUsableTimezone as
  // true in the first place. A vendor-prefixed form ('/vendor.example/.../Zone/Name') still
  // fails after stripping only the one leading slash — that is the safe direction, a real
  // rejection rather than a guess at which registry the rest of the string names.
  const zoneCandidate = trimmed.startsWith('/') ? trimmed.slice(1) : trimmed;
  // Runs the shared rule (offset-shape, then ICU resolvability, then the slash rule — see
  // zoneRejectionReason's own comment for why that order matters) and composes this call site's
  // own framing around whichever way it failed.
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
  // ICU resolves a zone name case-insensitively and through backward-compatibility links
  // ("australia/sydney", "NZ", "Zulu" all resolve to a real zone) but the CalDAV server behind
  // this account (Cyrus) looks a TZID up with an exact-string match against its own tzdata,
  // which is far stricter than ICU. Writing the caller's raw spelling would let an ICU-blessed
  // string reach the calendar as a TZID Cyrus itself cannot resolve — the same failure class the
  // offset-shape check above exists to prevent, one dimension over: a string ICU accepts that
  // the calendar server cannot read back as a name. `canonicalZoneName` returns ICU's own
  // canonical spelling for whatever was resolved (guaranteed not to fall back here, since
  // `isUsableTimezone` above already proved `zoneCandidate` resolves), so returning THAT is what
  // actually lands on a name Cyrus recognises, and it is also what the create/update response
  // reports as written. Do not "simplify" this back to returning the caller's input — the
  // canonical name is deliberately not an echo of what was typed.
  return canonicalZoneName(zoneCandidate);
}

/**
 * Resolve the zone the server actually runs with, from FASTMAIL_TIMEZONE (or the host zone when
 * it is unset) — held to the same rule `validateCallerTimezone` enforces on the caller-supplied
 * `timeZone` parameter (#157 amendment). Called once, from `runServer()` in index.ts, before
 * `setDefaultTimezone` — see that call site for why it is not called at module load instead.
 *
 * The two sources are handled ASYMMETRICALLY on purpose, because only one of them was actually
 * chosen by anyone:
 *
 * - FASTMAIL_TIMEZONE, when set, was written by whoever configured this server. A shorthand or
 *   unresolvable value there is exactly the silent-wrong-day failure the rule exists to prevent
 *   — today it silently falls back to the host zone with nothing said, and a caller has no way
 *   to learn that "EST" was never US Eastern. So a set-and-invalid value THROWS here, and
 *   `runServer()` turns that into a refusal to start: printing the exact value and the rule to
 *   stderr and exiting is strictly better than starting up and quietly working out of the wrong
 *   zone for every calendar read from then on.
 * - The host's own zone (`hostZone`, defaulted to `hostTimezone()` so a test can pass a fixed
 *   value instead of depending on the machine it runs on) is used only when FASTMAIL_TIMEZONE is
 *   unset — nobody configured it, so refusing to start over it would punish the operator for a
 *   setting they never touched. A rejected host zone instead falls back to `zone: 'UTC'` (the
 *   rule's own exception, and unambiguous by construction — the rejected name is, by definition,
 *   ambiguous) with a `warning` string the caller MUST print loudly rather than swallow: the
 *   server still starts, but a bare date-only window bound (list_calendar_events) now reads as a
 *   UTC day rather than this machine's local day, which is exactly the kind of change that must
 *   not happen silently.
 *
 * In practice the host-zone branch is close to unreachable: `Intl.DateTimeFormat().
 * resolvedOptions().timeZone` returns ICU's own canonical name for whatever the OS/TZ env
 * reports, and a canonical IANA zone name is either slash-bearing or is `UTC` itself — this is a
 * guard against a misconfigured HOST (e.g. a TZ environment variable ICU cannot canonicalise),
 * not an expected path.
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
 * Names the zone and nothing else in the ordinary case: this string lands in the
 * accepted-shapes sentence on every date rejection, so where the zone came from belongs in
 * the tool description (said once, where a model reads it) rather than repeated inside every
 * error.
 *
 * The exception is a zone name ICU cannot resolve. `zoneOffsetMsAt` deliberately falls back
 * to the HOST zone there rather than throwing, so naming the configured value alone would
 * print a zone the dates were not read in — the disclosure and the behaviour disagreeing, on
 * the one call where the caller is trying to work out why their days look wrong. So both are
 * named: what actually resolved, and the configured value that did not.
 *
 * Since `resolveConfiguredTimezone` (#157) validates `FASTMAIL_TIMEZONE` at startup and
 * refuses to start on a value that would land here, every production caller of this function
 * now always passes an already-usable zone, so this branch is not reachable through a
 * misconfigured `FASTMAIL_TIMEZONE` any more. It stays: `describeTimezone` is a general
 * utility, not something that gets to assume its argument was pre-validated by any one
 * caller, and covering the branch directly is cheaper than proving every future call site
 * always will be.
 */
export function describeTimezone(zone: string | undefined): string {
  if (!zone) return hostTimezone();
  if (isUsableTimezone(zone)) return zone;
  // Through the shared echo, so a cut zone name shows that it was cut. Slicing silently at
  // 40 characters printed a name the caller could neither recognise nor correct.
  const echoed = echoCallerText(zone, ZONE_ECHO_LIMIT);
  return `${resolveUsableTimezone(zone)} (the configured time zone "${echoed}" is not a time zone this server can resolve, ` +
    "so this server's own zone was used)";
}

/**
 * The UTC offset an IANA zone is at, at one instant, in milliseconds.
 *
 * Read by formatting the instant in the zone and treating the wall-clock components it
 * prints as if they were UTC: the difference between that and the real instant IS the
 * offset. No timezone database ships with this server, so ICU (through `Intl`) is the only
 * thing here that knows when a zone changes offset.
 *
 * An unusable IANA name falls back to the host zone rather than throwing, matching
 * `toLocalIso`'s posture on the same kind of bad zone string. Every current caller passes
 * `getDefaultTimezone()` down through `coerceCalendarWindowStart`/`End`, and since #157
 * `resolveConfiguredTimezone` validates that value at server startup — refusing to start
 * rather than let an unusable configured (or host) zone reach request time at all — so this
 * fallback is not reachable in production today. It stays because this is a low-level helper
 * with no way to enforce that every future caller pre-validates its `zone` argument the same
 * way, and because it is covered directly by its own unit tests.
 */
function zoneOffsetMsAt(utcMs: number, zone: string | undefined): number {
  let parts;
  try {
    parts = new Intl.DateTimeFormat('en-US', {
      timeZone: zone,
      hour12: false,
      // ERA IS REQUESTED BECAUSE THE YEAR IS READ BACK, and without it `Intl` prints the
      // ERA-RELATIVE year: proleptic year 0 formats as "1", year -1 as "2". That number then
      // went straight back into `utcMsFromComponents` as though it were the proleptic year,
      // so the offset came out a whole year wrong and a window bound near the start of the
      // era resolved to the wrong DAY with nothing said — `startDate: "0000-12-31"` answered
      // with 0000-01-01. Requesting the era makes the two eras distinguishable; the negation
      // below puts a BC year back on the proleptic scale ICU's own year is not on.
      era: 'short',
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
    }).formatToParts(new Date(utcMs));
  } catch {
    if (zone === undefined) return 0;
    return zoneOffsetMsAt(utcMs, undefined);
  }
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

// A whole Gregorian cycle: 400 years is exactly 146097 days, leap rules included. Exported
// because caldav-client.ts's date-only rollover steps over the same legacy two-digit-year
// mapping by the same trick, and two copies of this number could drift apart.
export const GREGORIAN_CYCLE_YEARS = 400;
const GREGORIAN_CYCLE_MS = 146097 * 24 * 60 * 60 * 1000;

/**
 * `Date.UTC` without its legacy two-digit-year mapping.
 *
 * `Date.UTC(26, 7, 12)` is the year 1926, not the year 26 — the same rule that makes
 * `new Date(99, 0)` a 1999 date. Every year the calendar window handles arrives as an
 * already-validated four-digit string, so `0026-08-12` would otherwise have been answered
 * with a window in 1926: a different window, silently, where `coerceUtcDate` on the same
 * value correctly returns the year 26. Shifting by one whole Gregorian cycle steps over the
 * mapping and back without disturbing the arithmetic, so a leap day still lands on the day
 * the proleptic Gregorian calendar puts it.
 */
function utcMsFromComponents(y: number, mo: number, d: number, h: number, mi: number, s: number): number {
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
 * The offset has to be sampled at the instant the wall clock names, and that instant is what
 * is being solved for — so the offsets in force a day either side are read first, and a
 * candidate answer is CHECKED against the offset actually in force where it lands. Without
 * that check a wall clock within a day of a DST transition lands an hour out.
 *
 * The two awkward cases are the reason the check exists rather than a second blind pass:
 *
 *   REPEATED (clocks go back) — the wall clock names two instants. The EARLIER one is
 *     returned, matching RFC 5545's and Temporal's `compatible` disambiguation.
 *   SKIPPED (clocks go forward) — the wall clock names none, and the answer is resolved
 *     FORWARD BY THE LENGTH OF THE GAP, again matching `compatible`. That equals the
 *     transition instant only for a clock sitting at the very start of the gap:
 *     `America/New_York 2026-03-08T02:00:00` gives 07:00:00Z (the transition), but
 *     `…T02:29:59` gives 07:29:59Z, half an hour past it. Resolving it BACKWARD is what a
 *     blind second pass did, and for the exclusive END of a window that quietly dropped the
 *     last hour of the requested day: in a zone whose transition is at midnight
 *     (America/Santiago, America/Havana) a single-day window ran local 00:00 to 23:00 and an
 *     event at 23:30 was never searched for. Neither case is refused — a caller asking about
 *     a day is entitled to an answer on the day a transition happens.
 */
function wallClockToUtcMs(y: number, mo: number, d: number, h: number, mi: number, s: number, zone: string | undefined): number {
  const naive = utcMsFromComponents(y, mo, d, h, mi, s);
  const before = zoneOffsetMsAt(naive - OFFSET_SAMPLE_SPAN_MS, zone);
  const after = zoneOffsetMsAt(naive + OFFSET_SAMPLE_SPAN_MS, zone);
  // No transition in range: one offset answers it.
  if (before === after) return naive - before;

  // `early` uses the pre-transition offset, so where both readings are valid (a repeated
  // hour) it is the earlier instant of the two.
  const early = naive - before;
  if (zoneOffsetMsAt(early, zone) === before) return early;
  const late = naive - after;
  if (zoneOffsetMsAt(late, zone) === after) return late;
  // Neither reading is valid: the wall clock was skipped. `early` is that wall clock shifted
  // FORWARD by the length of the gap — which lands ON the transition instant only when the
  // clock sat at the very start of the gap, and past it by however far into the gap it sat.
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
): string | undefined {
  if (value === undefined || value === null) return undefined;
  const { trimmed, kind } = classifyDateValue(value, paramName, acceptedWindowFormats(describeTimezone(zone)));

  if (kind === 'zoned-datetime') {
    // Named an instant; hand back that instant. This is the one branch the zone never
    // touches, and it is why an offset-carrying value is the way to ask for something the
    // local-day rule cannot express.
    const parsed = new Date(trimmed);
    if (Number.isNaN(parsed.getTime())) {
      throw new InvalidInputError(
        `${paramName} is not a valid date: "${echoDate(trimmed)}". ${acceptedWindowFormats(describeTimezone(zone))}`,
      );
    }
    return toUtcIso(parsed.getTime());
  }

  if (kind === 'date') {
    const [y, mo, d] = trimmed.split('-').map(Number);
    // The day is advanced in LOCAL days, not by adding 24 hours to the resolved instant: a
    // day that a DST transition makes 23 or 25 hours long would otherwise land the exclusive
    // end an hour inside or past the day the caller named. Date.UTC rolls month, year and
    // leap-day ends for us, and the date has already been proved real.
    return toUtcIso(wallClockToUtcMs(y, mo, d + dayOffset, 0, 0, 0, zone));
  }

  const m = LOCAL_DATETIME_PATTERN.exec(trimmed);
  if (!m || !isWallClockInRange(Number(m[4]), Number(m[5]), Number(m[6] ?? 0))) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${acceptedWindowFormats(describeTimezone(zone))}`,
    );
  }
  // A wall-clock datetime names a time of day, not a day, so `dayOffset` deliberately does
  // NOT apply to it — `endDate: 2026-08-12T17:00:00` is the exclusive end at five in the
  // afternoon, exactly as the Z-designated form is.
  return toUtcIso(wallClockToUtcMs(Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] ?? 0), zone));
}

/**
 * Whether a wall clock's components name a time that exists.
 *
 * The shape check alone is not acceptance, and this is where the calendar pair would
 * otherwise diverge from `coerceUtcDate` in the direction that matters. `Date.UTC` ROLLS an
 * out-of-range component instead of rejecting it, so `2026-08-12T99:99:99` resolved to a
 * window starting three and a half days later, silently, while `coerceUtcDate` and
 * `create_calendar_event` both reject the same value — the family accepted on a read what it
 * refused on a write, and only the clamp path says anything about a window it moved.
 *
 * `24:00:00` is deliberately allowed: the ECMAScript Date Time String Format accepts it as
 * the end of the day, so `new Date('2026-08-12T24:00:00')` is valid and the UTC coercion
 * takes it. Rejecting it here would be the same divergence pointing the other way. Date.UTC
 * rolls it into the following midnight, which is what it names.
 */
function isWallClockInRange(h: number, mi: number, s: number): boolean {
  if (h === 24) return mi === 0 && s === 0;
  return h <= 23 && mi <= 59 && s <= 59;
}

/**
 * The UTC instant a value parsed out of an iCalendar payload names, in milliseconds, for
 * ORDERING two of them against each other. `NaN` when there is nothing to order on.
 *
 * Lives beside the window coercions because it resolves a zone-less value the same way they
 * do — through the configured zone — and for the same reason. `formatICalDate` drops the
 * TZID parameter, so an event stored as `DTSTART;TZID=Australia/Sydney:20260325T083000`
 * reaches a caller as the bare `2026-03-25T08:30:00` while a UTC-stored one keeps its `Z`.
 * Comparing those two as STRINGS puts them in the wrong order whenever the account is not on
 * UTC (on a +10:00 account the bare one is 9.5 hours the earlier of the two), which is how a
 * genuinely earlier event was dropped by a `limit` that kept a later one.
 *
 * This is a best-effort reading, not a validation: it is ordering server data, not accepting
 * caller input, so an unreadable value returns NaN for the caller to place rather than
 * throwing, and an out-of-range component is left to roll rather than rejected.
 */
export function resolveCalendarInstantMs(value: string | undefined, zone: string | undefined): number {
  if (typeof value !== 'string') return NaN;
  const trimmed = value.trim();
  if (!trimmed) return NaN;
  // Carries its own zone: it names an instant, and no local reading applies.
  if (ZONE_DESIGNATOR_PATTERN.test(trimmed)) return Date.parse(trimmed);
  if (DATE_ONLY_PATTERN.test(trimmed)) {
    const [y, mo, d] = trimmed.split('-').map(Number);
    // An all-day value is placed at local midnight, the same instant the window's own
    // date-only bound resolves to, so the two are on one scale.
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
 * It sits here, beside the window coercions, because it has to agree with them: a bounds-free
 * window starts on the same day a caller would get by passing today's date as a DATE-ONLY
 * `startDate`, and the only way to guarantee that is to resolve it through the same
 * wall-clock-to-instant path (`zoneOffsetMsAt` to read which local day `nowMs` falls in, then
 * `wallClockToUtcMs` to put that day's midnight back on the UTC scale). Re-deriving either
 * step here would give the default window its own DST and offset behaviour, which is exactly
 * the drift the shared helpers exist to prevent.
 *
 * `nowMs` is passed in rather than read from the clock so the caller — and its tests — decide
 * what "today" is. Year-0000 and era handling are irrelevant here in a way they are not for a
 * caller-named bound: this value is always the present.
 */
export function startOfLocalDayUtcIso(nowMs: number, zone?: string): string {
  // The wall clock in `zone` at `nowMs`, read by shifting the instant by the offset in force
  // and taking the UTC components of the result — the same trick `zoneOffsetMsAt` uses in
  // reverse, and the reason the offset is sampled AT `nowMs` rather than assumed.
  const local = new Date(nowMs + zoneOffsetMsAt(nowMs, zone));
  return toUtcIso(wallClockToUtcMs(
    local.getUTCFullYear(), local.getUTCMonth() + 1, local.getUTCDate(), 0, 0, 0, zone,
  ));
}

// The pagination offset shared by the list/search tools: a 0-based index into the
// full result set (the JMAP `position` argument, RFC 8620 section 5.5). Values are
// coerced leniently like `limit` — a stringified "40" from a client that stringifies
// numbers is accepted — but the three shapes that would page somewhere the caller did
// not mean are REJECTED rather than repaired, because a wrong offset silently skips or
// repeats messages instead of erroring:
//
//   - A NEGATIVE value. JMAP reads a negative position as an offset from the END of
//     the results, so `-1` would quietly return the last page. That is a footgun with
//     no upside here: reading from the other end is what `ascending` is for.
//   - A FRACTION. Rounding 1.5 is a guess about which message the caller meant to
//     start at.
//   - Anything non-numeric, or an integer too large to be exact.
//
// An omitted value, `null`, an empty string, and 0 all mean the same thing: start at
// the first result. There is nothing to guess there, so the blank shapes coerce rather
// than reject.
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

// Clamp a caller-supplied limit into [1, max], tolerating string and NaN input.
// A lenient client may send "20", and a bare Number("abc") yields NaN — which
// JMAP serializes as `"limit": null`, i.e. a query with no bound at all. The
// `|| fallback` is therefore a guard against an unbounded result set, not a
// nicety. Unlike the other coercers this never throws: a limit is a convenience
// knob, and a caller who fat-fingers it wants results, not an error.
export function clampLimit(value: unknown, fallback: number, max: number): number {
  return Math.min(Math.max(Number(value) || fallback, 1), max);
}

function acceptedDateFormats(): string {
  return 'Accepted: a date such as 2026-07-20 (treated as 00:00:00 UTC on that date), or a full datetime such as 2026-07-20T14:30:00Z or 2026-07-20T14:30:00+01:00.';
}

// Loud-reject a settable string field that was provided but is empty,
// whitespace-only, or null. Callers invoke this only for fields that were
// actually present (i.e. !== undefined at the call site), so silently omitting
// a field stays distinct from explicitly blanking it. Returns the trimmed value.
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

// Parse an RFC 5322 "Display Name <email>" recipient string into a JMAP
// EmailAddress object. Bare addresses pass through as { email }, and a blank
// display name is omitted. This is a pragmatic parse, not the full RFC grammar.
// Callers map it over already-trimmed, non-empty arrays (coerceStringArray
// filters blanks), so input is assumed non-empty.
export function parseAddress(input: string): { name?: string; email: string } {
  const trimmed = String(input).trim();
  const open = trimmed.lastIndexOf('<');
  const close = trimmed.lastIndexOf('>');
  if (open !== -1 && close > open) {
    const email = trimmed.slice(open + 1, close).trim();
    let name = trimmed.slice(0, open).trim();
    // Strip one pair of surrounding double-quotes from a quoted display name.
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
// THREE SOURCES, exactly one per item. They are three ways to name the bytes, not three
// optional extras, and each is gated separately (see uploadAttachments):
//
//   path                    a local file, read off disk and uploaded (FASTMAIL_ATTACH_DIR)
//   blobId                  content already in the account's blob store
//   emailId + attachmentId  a part of an existing message, resolved to its blob
//
// Every source field is optional HERE because the shape is a union that TypeScript cannot
// express as one interface without making every consumer narrow it; coerceAttachments is
// the gate that guarantees exactly one source is set, and it is the only producer of these
// objects.
export interface AttachmentSpec {
  path?: string;
  blobId?: string;
  emailId?: string;
  attachmentId?: string;
  name?: string;
  contentType?: string;
  // The Content-ID an html body references this file by, to display it inside the message
  // rather than hang it off the end. Stored here in CANONICAL form: coerceAttachments
  // normalizes the two spellings a caller realistically copies (see stripCidSpelling)
  // before validating, so everything downstream compares one value per identifier.
  cid?: string;
}

// The keys that NAME the bytes. Exactly one source per item; `emailId` and `attachmentId`
// are one source spelled in two keys, so they are required together.
const ATTACHMENT_SOURCE_KEYS = ['path', 'blobId', 'emailId', 'attachmentId'] as const;
// The keys that DESCRIBE whatever the source named. Valid on every source.
const ATTACHMENT_COMMON_KEYS = ['name', 'contentType', 'cid'] as const;
const ATTACHMENT_KEYS = new Set<string>([...ATTACHMENT_SOURCE_KEYS, ...ATTACHMENT_COMMON_KEYS]);

// The item shape named in every whole-parameter refusal below, kept in one place so the
// copies cannot drift from the schema.
const ATTACHMENT_ITEM_SHAPE = '{ path | blobId | emailId+attachmentId, name?, contentType?, cid? }';

// The sentence every source-selection refusal ends with, so the three sources are always
// spelled out the same way wherever the caller lands.
const ATTACHMENT_SOURCE_RULE =
  "Give exactly one source per item: 'path' (a local file), 'blobId' (content already in " +
  "the account), or 'emailId' + 'attachmentId' together (a part of an existing message).";

// Which source an item names. `null` counts as ABSENT, not as a value: a lenient client
// that fills every declared key emits null for the ones it has nothing to say about, and
// reading those as "this item names four sources" would reject every call from such a
// client.
function namedAttachmentSourceKeys(obj: Record<string, unknown>): string[] {
  return ATTACHMENT_SOURCE_KEYS.filter((k) => obj[k] !== undefined && obj[k] !== null);
}

// Read one required string key off an item, rejecting a missing/blank/non-string value by
// index. Returns the trimmed value: an accidental leading/trailing space would otherwise
// reach the filesystem (path) or the server (an id) and read as "not found".
function requireAttachmentString(obj: Record<string, unknown>, key: string, index: number, hint: string): string {
  const value = obj[key];
  if (typeof value !== 'string' || value.trim() === '') {
    throw new McpError(ErrorCode.InvalidParams, `attachments[${index}] is missing a non-empty '${key}'; ${hint}`);
  }
  return value.trim();
}

// Coerce the `attachments` tool param into AttachmentSpec[] | undefined. Accepts a
// real array, or a JSON-string array from lenient clients (mirroring
// coerceStringArray). Per element it REJECTS — never silently drops — a non-object,
// an item naming no source or more than one, an `emailId`/`attachmentId` half-pair, or an
// unexpected per-item key, naming the index so the caller can fix it (assertKnownParams is
// top-level only and won't catch nested keys, so this is the sole guard for the item
// shape). A bare string element is rejected rather than guessed as a path (too magic); a
// JSON-object string is parsed.
//
// A key belonging to a source the item did not choose is a REJECTION, not a silent ignore:
// `{ blobId, attachmentId }` reads as two different intentions, and picking one of them
// would attach bytes the caller did not ask for. Because the only source-specific keys ARE
// the source keys, the exactly-one-source rule is what enforces that — there is no second
// per-source allowlist pass to drift from it.
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
    // (or none) is told that rather than being told the first source's value is malformed.
    const named = namedAttachmentSourceKeys(obj);
    const namesMessagePart = named.includes('emailId') || named.includes('attachmentId');
    // Count SOURCES, not keys: emailId+attachmentId is one source spelled in two keys, so
    // counting keys would wave through `{ blobId, attachmentId }` — the exact mix this rule
    // exists to refuse.
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
      // Both halves or neither. One alone is a half-written reference, and guessing the
      // other half is not possible — an emailId names a message, not a part of one.
      spec.emailId = requireAttachmentString(obj, 'emailId', i, "a message part is named by 'emailId' AND 'attachmentId' together.");
      spec.attachmentId = requireAttachmentString(obj, 'attachmentId', i, "a message part is named by 'emailId' AND 'attachmentId' together.");
    } else if (named[0] === 'blobId') {
      spec.blobId = requireAttachmentString(obj, 'blobId', i, 'give the blobId of content already in the account.');
      // A blob is bytes and nothing else: unlike a file it has no basename to fall back on
      // and unlike a message part it carries no declared filename, so there is nothing to
      // derive a name from. Defaulting one would put an invented filename on outgoing mail,
      // so this rejects instead — the fail-loud posture the rest of this coercer takes.
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
      // A BLANK name reads as absent, so every source falls back to the default it
      // documents: the file's basename, or the message part's own name. Kept here rather
      // than at each consumer because `??` cannot see the difference — `'' ?? info.name`
      // is `''`, which would put a nameless attachment on outgoing mail while the schema
      // promised a default. The trim also stops a stray space from becoming the filename
      // recipients see. The blobId branch above rejects a blank name outright instead,
      // and the two agree: a blob has no default to fall back to, so there "absent" is
      // not a usable state.
      const trimmed = obj.name.trim();
      if (trimmed) spec.name = trimmed;
    }
    if (obj.contentType !== undefined) {
      if (typeof obj.contentType !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].contentType must be a string.`);
      spec.contentType = obj.contentType;
    }
    if (obj.cid !== undefined) {
      // Normalize before validating, then keep the NORMALIZED value: `cid:logo` copied out
      // of an html reference and `<logo>` copied out of a header are the same identifier,
      // and every later comparison — collision detection above all — has to see them as
      // one. Validating the raw value instead would bounce both spellings; keeping the raw
      // value would let two spellings of one identifier become two parts sharing a
      // Content-ID, which makes every reference to it ambiguous.
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

// The item shape named in every whole-parameter refusal below, kept in one place so the
// copies cannot drift from the schema.
const PARTICIPANT_ITEM_SHAPE = '{ email, name? }';

// Coerce the `participants` tool param into ParticipantSpec[] | undefined, the same way
// coerceAttachments handles its own array-of-objects param. Accepts a real array or a
// JSON-string array; a comma-joined string is NOT split, because an item here is an
// object, not a scalar, so there is no unambiguous reading of one.
//
// Per element it REJECTS — never silently drops — a non-object, a missing/blank `email`,
// an unexpected key, or a key of the wrong type, naming the index so the caller can fix
// it. assertKnownParams is top-level only and won't see nested keys, and the MCP SDK does
// not enforce inputSchema, so this is the sole guard on the item shape: without it a
// non-string `email` or an object `name` would reach the ATTENDEE serializer.
//
// A BARE STRING element is accepted and read as the address: `["a@example.com"]` means
// the same as `[{ email: "a@example.com" }]`. That differs from coerceAttachments, which
// refuses a bare string, and the difference is deliberate — an attachment spec has four
// keys and a lone string could plausibly be a path or a display name, whereas a
// participant's only required key is the address. It also matches the recipient lists
// (`to`/`cc`/`bcc`), which have always taken bare address strings. The address itself is
// NOT vetted here: validateAttendeeEmail (src/caldav-client.ts) owns the address rules
// for both the create and update paths, and a second copy of them would drift.
//
// The returned objects are built fresh from the validated keys rather than passed through,
// so a key added to ParticipantSpec later has to be handled here explicitly instead of
// arriving by accident.
export function coerceParticipants(value: unknown): ParticipantSpec[] | undefined {
  if (value === undefined || value === null) return undefined;

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    // A blank string reads as "not supplied", not as the empty list. On
    // update_calendar_event an empty array REMOVES every attendee, so resolving an
    // ambiguous blank in that direction would destroy data on a guess; resolving it as
    // omitted leaves the event alone. To clear attendees, pass an actual empty array.
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
    // Trim the address for the same reason coerceAttachments trims a path: a stray
    // leading/trailing space would otherwise fail validateAttendeeEmail's no-whitespace
    // rule and read as "invalid address" for an address that is fine.
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
 * validated array of fresh objects, following the same three-part discipline as
 * coerceAttachments/coerceParticipants — and for the same reason, sharpened here by what the
 * write does with the result: the MCP SDK does not enforce `inputSchema`, and these values
 * are copied into a `ContactCard/set` patch, so an unvalidated object would be written onto a
 * real card verbatim.
 *
 *   1. UNKNOWN KEYS are rejected, naming the index.
 *   2. Every known key is TYPE-CHECKED, naming the index. A key allowlist alone would still
 *      let `{address: {…}}` or `{label: []}` through to the card.
 *   3. The value passed onward is a FRESH LITERAL built from the validated keys, never a
 *      spread of the caller's object. That is what makes adding a key later a conscious edit
 *      here rather than a silent widening of what reaches the server.
 *
 * A JSON-string array is accepted (lenient clients stringify structured params); a blank
 * string reads as "not supplied", never as the empty array — the empty array is a rejected
 * shape on these parameters, and resolving a blank one into it would turn a client quirk into
 * a rejection the caller cannot explain.
 *
 * `allowBareString` is on for `emails` and `phones`, whose only required key is the value
 * itself, so `["a@b.example"]` means `[{address: "a@b.example"}]` — matching how the recipient lists
 * and `participants` already read a bare string. It is off for `addresses`, where an entry has
 * no single obvious scalar reading.
 *
 * Duplicate values are REJECTED naming both positions: on `emails`/`phones` a repeat cannot be
 * matched against the stored card twice, so it would silently surface as an unknown addition,
 * and on any of the three it is a caller mistake with no useful reading.
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
      // A blank label is rejected like every other blank here, rather than written. On a
      // stored card an empty `label` is what "no label" already looks like, so accepting one
      // would write a property that reads as absent — a change with no visible effect, which
      // is worse than an error. It is deliberately NOT repurposed as a way to remove a label:
      // that would be a new clearing mechanism, and removing a label is not something this
      // tool can currently express.
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
 * Coerce the `name` parameter of a contact write. A bare string is the common case and reads
 * as the full name; the structured form names the parts. Same three-part discipline as the
 * entry arrays: unknown keys rejected, every key type-checked, and a fresh literal returned.
 *
 * A blank value is rejected rather than read as "clear the name": `name` is not clearable
 * (see the tool description), so silently dropping a blank one would leave the caller thinking
 * a name had been removed when nothing happened.
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
