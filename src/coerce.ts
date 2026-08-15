import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';
import { isAuthorableCid, stripCidSpelling } from './inline-images.js';
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

export function coerceStringArray(value: unknown): string[] | undefined {
  if (value === undefined || value === null) return undefined;
  if (Array.isArray(value)) return value.map(String);
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return [];
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.map(String);
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
  if (typeof value !== 'string') {
    throw new InvalidInputError(
      `${paramName} must be a date string, not ${Array.isArray(value) ? 'an array' : `a ${typeof value}`}. ${acceptedDateFormats()}`,
    );
  }
  const trimmed = value.trim();
  if (!trimmed) {
    throw new InvalidInputError(
      `${paramName} cannot be empty; omit it to search without that date bound. ${acceptedDateFormats()}`,
    );
  }

  const dateOnly = DATE_ONLY_PATTERN.test(trimmed);
  const datePart = dateOnly ? trimmed : DATE_TIME_PATTERN.exec(trimmed)?.[1];
  if (!datePart) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${acceptedDateFormats()}`,
    );
  }

  // A day that doesn't exist in its month parses rather than failing (2026-02-31 becomes
  // 2026-03-03), which would silently shift the search window off the dates the caller
  // asked for. Probe the calendar date on its own — probing the whole value wouldn't
  // work, since an offset legitimately moves the UTC date.
  const dayProbe = new Date(`${datePart}T00:00:00Z`);
  if (Number.isNaN(dayProbe.getTime()) || !dayProbe.toISOString().startsWith(datePart)) {
    throw new InvalidInputError(
      `${paramName} is not a real calendar date: "${echoDate(trimmed)}". ${acceptedDateFormats()}`,
    );
  }

  // A date-only value is expanded explicitly rather than left to Date's parse so the
  // intent (midnight UTC, never host-local) is visible in the code, not a spec detail.
  const parsed = new Date(dateOnly ? `${trimmed}T00:00:00Z` : trimmed);
  if (Number.isNaN(parsed.getTime())) {
    throw new InvalidInputError(
      `${paramName} is not a valid date: "${echoDate(trimmed)}". ${acceptedDateFormats()}`,
    );
  }
  return parsed.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function echoDate(value: string): string {
  return value.length > DATE_ECHO_LIMIT ? `${value.slice(0, DATE_ECHO_LIMIT)}...` : value;
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

// One file-to-attach spec as it arrives from a tool call (before path confinement
// and upload, which happen in jmap-client.ts).
export interface AttachmentSpec {
  path: string;
  name?: string;
  contentType?: string;
  // The Content-ID an html body references this file by, to display it inside the message
  // rather than hang it off the end. Stored here in CANONICAL form: coerceAttachments
  // normalizes the two spellings a caller realistically copies (see stripCidSpelling)
  // before validating, so everything downstream compares one value per identifier.
  cid?: string;
}

const ATTACHMENT_KEYS = new Set(['path', 'name', 'contentType', 'cid']);

// The item shape named in every whole-parameter refusal below, kept in one place so the
// three copies cannot drift from the schema.
const ATTACHMENT_ITEM_SHAPE = '{ path, name?, contentType?, cid? }';

// Coerce the `attachments` tool param into AttachmentSpec[] | undefined. Accepts a
// real array, or a JSON-string array from lenient clients (mirroring
// coerceStringArray). Per element it REJECTS — never silently drops — a non-object,
// a spec missing `path`, or an unexpected per-item key, naming the index so the
// caller can fix it (assertKnownParams is top-level only and won't catch nested
// keys, so this is the sole guard for the item shape). A bare string element is
// rejected rather than guessed as a path (too magic); a JSON-object string is parsed.
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
        throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object with a path, not a bare string.`);
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object with a path.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter(k => !ATTACHMENT_KEYS.has(k));
    if (unknownKeys.length > 0) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...ATTACHMENT_KEYS].join(', ')}`);
    }
    if (typeof obj.path !== 'string' || obj.path.trim() === '') {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] is missing a non-empty 'path'.`);
    }
    // Trim the path (consistent with coerceStringArray's lenient coercion): an accidental
    // leading/trailing space would otherwise reach the filesystem and read as "file not found".
    const spec: AttachmentSpec = { path: obj.path.trim() };
    if (obj.name !== undefined) {
      if (typeof obj.name !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].name must be a string.`);
      spec.name = obj.name;
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
