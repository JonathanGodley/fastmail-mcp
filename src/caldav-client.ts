import { DAVClient, DAVCalendar, DAVCalendarObject, DAVResponse, davRequest, urlEquals } from 'tsdav';
// Caller-fixable input must throw coerce.ts's tagged InvalidInputError, which the CallTool
// boundary maps to InvalidParams; a plain Error surfaces as InternalError. See
// docs/conventions.md.
import { InvalidInputError, describeUntrustedAt, etcGmtOffsetNote, etcGmtUtcOffset, requireNonEmpty, validateClearFields, coerceCalendarWindowStart, coerceCalendarWindowEnd, startOfLocalDayUtcIso, describeTimezone, resolveCalendarInstantMs, echoCallerText, ZONE_ECHO_LIMIT, resolveUsableTimezone, isUsableTimezone, validateCallerTimezone, canonicalZoneName, GREGORIAN_CYCLE_YEARS } from './coerce.js';
import { trimEnd } from './trim-end.js';
import { foldICalLine } from './ical-fold.js';
// A calendar window interprets local dates in the same zone the rest of the server displays,
// so it reads that stored value rather than re-deriving one from the environment.
import { getDefaultTimezone } from './email-formatter.js';
import { generateVTimezone } from './vtimezone.js';

export interface CalDAVConfig {
  username: string;
  password: string;
  serverUrl?: string;
  // Display name for the ORGANIZER this client emits. Resolved by the caller from the
  // environment (the server does that through its shared multi-name lookup, so a DXT
  // user_config spelling reaches it); left unset it falls back to the username.
  displayName?: string;
  // The clock "today" is read from for a default window; injectable so that window is
  // unit-testable. Read ONCE per call so the two ends of one window cannot straddle a midnight.
  now?: () => number;
}

export interface CalendarInfo {
  id: string;
  displayName: string;
  url: string;
  description?: string;
  color?: string;
}

export interface Participant {
  email: string;
  name?: string;
  role?: string;       // REQ-PARTICIPANT, OPT-PARTICIPANT, CHAIR
  status?: string;     // PARTSTAT: ACCEPTED, DECLINED, TENTATIVE, NEEDS-ACTION
  cutype?: string;     // CUTYPE: INDIVIDUAL, ROOM, RESOURCE, GROUP, UNKNOWN
  rsvp?: boolean;      // RSVP: TRUE/FALSE
}

export interface CalendarEvent {
  id: string;
  // The resource's CalDAV URL, and a second series-wide delete handle: `findCalendarObjectByUID`
  // accepts it interchangeably with `id`. Kept on the row anyway: a UID is unique per
  // collection, not per account, so the url is what tells two same-UID records apart, and for a
  // resource with no VEVENT the url IS the id (parseCalendarObjects' minimal fallback).
  url: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: Participant;
  participants?: Participant[];
  // ---- recurrence (#64) ----
  // Set for a series master AND for an expanded occurrence.
  isRecurring?: boolean;
  // From RECURRENCE-ID: its presence says `start`/`end` are the in-window occurrence rather
  // than the series' original DTSTART.
  recurrenceId?: string;
  // The raw RRULE value. Not exclusive with `recurrenceId`: RFC 5545 §3.8.5.3 lets an override
  // block carry its own rule, and then this is that block's rule, not the series'.
  recurrenceRule?: string;
  // Every RDATE line's values joined into one comma-separated list. It proves `start` is not
  // the only date this entry has, which the window guard in `getCalendarEvents` needs (#162).
  //
  // Values only, deliberately: TZID, VALUE=DATE and VALUE=PERIOD are dropped, so these
  // designator-less values do NOT follow the `timeZone` rule and are only evidence that other
  // dates exist. No tool acts on an individual RDATE, so parameters would claim a precision the
  // field cannot back up.
  //
  // Normally absent on the listing path: Cyrus strips RDATE (and RRULE) from an expanded block
  // (scripts/probes/calendar-expand.probe.mjs, calendar-rdate-expand.probe.mjs).
  recurrenceDates?: string;
  // ---- time zone (#139) ----
  // The IANA name DTSTART was written in, when it differs from the zone this server would
  // assume. NEVER an offset: `start` stays a bare wall clock so DST is worked out by the
  // reader's own zone database.
  //
  // Omitted for rows already in the configured zone. `null` means `start` is genuinely
  // FLOATING (RFC 5545 §3.3.5). A `Z` instant and an all-day value are omitted too, not
  // `null`, since `null` would assert "floating".
  timeZone?: string | null;
  // The IANA name `end` was written in, ONLY when it differs from `start`'s (legal per RFC 5545
  // §3.8.5.3 and #140). Omitted when `end` is absent or computed from a DURATION, which inherits
  // start's zone. `null` means one end is floating and the other is not.
  endTimeZone?: string | null;
  // `busy` or `free` (#194), derived from TRANSP by readTransparency. `string`, not the closed
  // `Transparency` union, because the account can hold tokens this server never writes.
  // Always present from `get_calendar_event`; listed rows carry it only when not busy.
  transparency?: string;
}

// `total` is how many matched before `limit` trimmed the page: `limit` is a hard cap with no
// paging, and a capped page read as the whole answer looks like an empty calendar (#100).
export interface CalendarEventQueryResult {
  events: CalendarEvent[];
  total: number;
  // Set only when the window queried was narrower than the one the caller described. Structure,
  // not prose: the formatter owns the wording, as `QueryResult.exclusion` does for email.
  windowClamp?: CalendarWindowClamp;
  brokenCollections?: string[];
}

/**
 * The entries in the calendar home's own listing that failed to list (#136). Absent, never
 * empty, when every entry answered; `buildBrokenCollectionNote` owns the wording.
 *
 * A broken entry keeps only its href, so these are PATHS: there is no name, and no way to say
 * whether it was a calendar at all. No message built from it may claim a calendar was lost.
 */
export type BrokenCollections = string[];

/** What `getCalendars` returns: the listing, plus any collection that failed to list (#136). */
export interface CalendarListResult {
  calendars: CalendarInfo[];
  brokenCollections?: BrokenCollections;
}

/**
 * What `getCalendarEventById` returns. The event is nested rather than carrying the disclosure
 * as a field of its own because the event object IS the tool's JSON body — a `brokenCollections`
 * key inside it would read as a property of the event.
 */
export interface CalendarEventResult {
  event: CalendarEvent;
  /**
   * The OTHER records this id named, when it named more than one (#101). Absent, never empty,
   * when the id is unambiguous. `event` is the first copy in `CalendarObjectLookup`'s order.
   *
   * On the result, not on `CalendarEvent`: that is the shared row type the list path also
   * serialises, and the list path resolves no id.
   */
  otherCopies?: CalendarEventCopy[];
  /**
   * Whether the id was a resource url rather than a shared UID. Meaningful only beside
   * `otherCopies`: an addressed id's writes act on THIS record instead of being refused, so the
   * note must not tell the caller to pick a url. See `CalendarObjectLookup`.
   */
  addressedByUrl?: boolean;
  /** Set when the writes would refuse the addressed id (`CalendarObjectLookup.collision`). */
  addressCollision?: { addressedUid: string | undefined };
  brokenCollections?: BrokenCollections;
}

export interface DeleteCalendarEventResult {
  /** The deleted record's own UID, as update reports it, whichever form of id was passed. */
  eventId: string;
  /** The resource actually deleted: a UID can spell another record's url. */
  url: string;
  brokenCollections?: BrokenCollections;
}

/** Why, and to what, a calendar window ended up narrower than the one the caller described. */
export interface CalendarWindowClamp {
  // The bound the caller OMITTED, which had to be invented CALENDAR_OPEN_WINDOW_DAYS away
  // from the one they gave. `'both'` is the caller who named neither: the window then starts
  // at local midnight today and runs the same invented span forward. Absent when both bounds
  // were named.
  invented?: 'startDate' | 'endDate' | 'both';
  // Caller-named bounds whose resolved instant ran outside the four-digit-year range every
  // consumer of these values can express, and so were pulled back to its edge. `edge` names
  // WHICH edge, because the disclosure is an opposite statement at each end and a window can
  // saturate at both at once — knowing only the top end, the note told a caller whose bound
  // was pulled UP to year 0000 that it had "resolved past the last date this server can
  // express", the reverse of what happened.
  saturated?: Array<{ bound: 'startDate' | 'endDate'; edge: 'earliest' | 'latest' }>;
  // The window actually queried. `end` is exclusive.
  start: string;
  end: string;
}

/**
 * Falls back when the configured name is unset, blank, or an unresolved DXT placeholder like
 * "${user_config.fastmail_caldav_display_name}", which would otherwise land in generated iCal.
 * The server's own lookup rejects placeholders too; this guard covers a directly-constructed
 * client.
 */
export function resolveDisplayName(raw: string | undefined, fallback: string): string {
  const trimmed = raw?.trim();
  if (!trimmed || /\$\{[^}]+\}/.test(trimmed)) return fallback;
  return trimmed;
}

/**
 * ICALENDAR STRUCTURE IS DECIDED ON WHOLE CONTENT LINES, NEVER WITH A `/m`-ANCHORED REGEX.
 *
 * RFC 5545 §3.1 knows one line break, CRLF (a bare LF is tolerated because real servers emit
 * it). JavaScript's `/m` anchors also match after U+2028, U+2029 and a bare CR, all of which
 * are legal unescaped inside a TEXT value, so a `/m` regex lets anyone who writes a SUMMARY or
 * DESCRIPTION forge structure: split a VEVENT so its dates vanish, fabricate a second event,
 * or make a resource report ANOTHER resource's UID, which `findCalendarObjectByUID` would then
 * hand to an irreversible delete. Do not "simplify" any of this back to a `/m` regex.
 *
 * Unfolding happens LATER than the component split (see `parseICalValue`), so a marker at the
 * head of a continuation line is text, not structure.
 */
interface ICalContentLine {
  /** The line's text, with no trailing line break. */
  text: string;
  /** Offset of the line's first character in the original payload. */
  start: number;
  /** Offset one past the line's last character, BEFORE its line break. */
  end: number;
}

/**
 * Split a payload into content lines on CRLF or LF and NOTHING else, keeping each line's
 * character offsets so a component can be sliced back out of the original string verbatim
 * (`normalizeMasterVEventFirst` does a literal substring replace on what it is handed).
 */
function icalContentLines(data: string): ICalContentLine[] {
  const out: ICalContentLine[] = [];
  let lineStart = 0;
  let i = 0;
  while (i < data.length) {
    const ch = data[i];
    if (ch === '\n') {
      out.push({ text: data.slice(lineStart, i), start: lineStart, end: i });
      i += 1;
      lineStart = i;
    } else if (ch === '\r' && data[i + 1] === '\n') {
      out.push({ text: data.slice(lineStart, i), start: lineStart, end: i });
      i += 2;
      lineStart = i;
    } else {
      i += 1;
    }
  }
  out.push({ text: data.slice(lineStart), start: lineStart, end: data.length });
  return out;
}

/**
 * Whether a line is a FOLDED CONTINUATION of the line above it (RFC 5545 §3.1). Structural
 * scans skip these: libical folds at a fixed octet count, so a fold placing `UID:` at the head
 * of a continuation is deterministic to construct.
 */
function isFoldedContinuation(line: string): boolean {
  return line.startsWith(' ') || line.startsWith('\t');
}

/**
 * Whether an iCalendar block contains the named property as a whole logical line. It fronts
 * the recurrence refusal on update/delete and the ORGANIZER/ATTENDEE patch routing.
 *
 * CASE-INSENSITIVE (RFC 5545 §3.1): libical upper-cases names, but a third party can PUT
 * `rrule:`, and missing a real rule lets a delete destroy a series it should refuse.
 *
 * UNFOLDS FIRST rather than skipping continuations as the structural scans do: skipping reads
 * a name split across a fold (`RRU\r\n LE:`) as absent, which fails open. A continuation is
 * appended to the line above, so it still cannot begin a logical line.
 */
export function hasICalProperty(block: string, key: string): boolean {
  const test = new RegExp(`^${key}[;:]`, 'i');
  return unfoldedICalLines(block).some(line => test.test(line));
}

/**
 * A block's LOGICAL lines, continuations unfolded. Text only, with no offsets: anything that
 * edits the payload uses `icalContentLines` + `structuralLine` instead.
 */
function unfoldedICalLines(block: string): string[] {
  const out: string[] = [];
  for (const { text } of icalContentLines(block)) {
    if (isFoldedContinuation(text) && out.length > 0) out[out.length - 1] += text.slice(1);
    else out.push(text);
  }
  return out;
}

/**
 * A line's text for STRUCTURAL comparison, or null if the line is a folded continuation.
 * Trailing whitespace (and a stray CR from `\r\r\n`) is trimmed; leading whitespace is not,
 * because it is what makes a line a continuation.
 */
function structuralLine(text: string): string | null {
  if (isFoldedContinuation(text)) return null;
  return trimEnd(text, (ch) => ch === '\r' || ch === '\t' || ch === ' ');
}

/** Every VEVENT block in a payload, as verbatim substrings of it. */
function extractVEventBlocks(data: string): string[] {
  const lines = icalContentLines(data);
  const blocks: string[] = [];
  let openIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const text = structuralLine(lines[i].text);
    if (text === null) continue;
    if (text === 'BEGIN:VEVENT') {
      if (openIdx === -1) openIdx = i;
    } else if (text === 'END:VEVENT' && openIdx !== -1) {
      blocks.push(data.slice(lines[openIdx].start, lines[i].end));
      openIdx = -1;
    }
  }
  return blocks;
}

export function extractVEvent(data: string): string | null {
  return extractVEventBlocks(data)[0] ?? null;
}

/**
 * Whether a STORED calendar resource holds a repeating series; what update/delete refuse on.
 * Read off the resource, never a listing row, whose RRULE expansion has stripped.
 *
 * Fails CLOSED, since it fronts an irreversible write. Any of four markers is enough: an
 * RRULE; an RDATE (RFC 5545 §3.8.5.2, a series with no rule at all, #162); more than one
 * VEVENT block (one resource is one UID, so a second block is an override); or any
 * RECURRENCE-ID (a series whose master was removed). The last two match the read path's
 * `blockCountProvesSeries`, so the two halves agree on what a series is.
 *
 * The scan is VEVENT-WIDE, not position-aware, so a marker inside a VALARM counts. The reads
 * are position-aware (`ownPropertyLines`) and ignore it, so such an event reads as one-off
 * while update/delete refuse it as repeating. The split is deliberate: the payload is
 * malformed, and refusing an irreversible write is the fail-closed direction.
 */
export function isRecurringSeriesResource(icalData: string | null | undefined): boolean {
  const blocks = extractVEventBlocks(icalData || '');
  if (blocks.length === 0) return false;
  if (blocks.length > 1) return true;
  return blocks.some((block) =>
    hasICalProperty(block, 'RRULE') || hasICalProperty(block, 'RDATE') || hasRecurrenceId(block));
}

/**
 * The event's SUMMARY for an error message, unescaped, falling back to the caller's id. Read off
 * the MASTER block so an override's re-titled occurrence cannot rename the series.
 */
function calendarObjectTitle(icalData: string | null | undefined, eventId: string): string {
  const blocks = extractVEventBlocks(icalData || '');
  // The master is the block with no RECURRENCE-ID; RFC 5545 does not fix component order, and
  // an all-overrides resource has no master at all, so fall through to the first block.
  const vevent = blocks.find((b) => !hasICalProperty(b, 'RECURRENCE-ID')) ?? blocks[0];
  const summary = vevent ? parseICalValue(vevent, 'SUMMARY') : undefined;
  return summary ? unescapeICalText(summary) : eventId;
}

/**
 * The one refusal update and delete raise on a repeating event. It refuses outright because
 * `create_calendar_event` cannot make a repeating event, so this server must not destroy or
 * rewrite one (CONTRIBUTING.md, "A destroy must not remove what the server cannot recreate").
 * Per-occurrence editing is #146, design in #109. There is deliberately NO override parameter,
 * and the message says so, or an LLM caller spends turns hunting for one.
 */
export function recurringSeriesRefusal(
  action: 'update' | 'delete',
  title: string,
): InvalidInputError {
  const consequence = action === 'delete'
    ? 'Deleting it would remove every occurrence, past and future, and the server would mail a cancellation to every attendee.'
    : 'Changing it would move every occurrence, and where single occurrences have already been edited on their own there is no agreed answer for what should happen to them.';
  // The unescaped title is attacker-choosable text (docs/conventions.md, untrusted values in
  // prose), so it goes through the shared echo inside its double quotes.
  return new InvalidInputError(
    `"${echoCallerText(title)}" is a repeating event, and this server will not ${action} it. `
    + `${consequence} `
    + 'There is no parameter, flag or confirmation that overrides this, so do not look for one: '
    + 'this server cannot CREATE a repeating event (create_calendar_event writes single events only), '
    + 'so it will not destroy or rewrite one it has no way to put back. '
    + 'Use the Fastmail web interface to change or delete a repeating event, including a single occurrence of one. '
    + 'get_calendar_event still works on this id and is read-only.',
  );
}

/**
 * The first colon outside quotes: the parameter/value boundary. A quoted parameter value can
 * hold colons (DELEGATED-FROM="mailto:boss@example.com").
 */
export function findValueBoundary(line: string): number {
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      inQuote = !inQuote;
    } else if (ch === ':' && !inQuote) {
      return i;
    }
  }
  return -1;
}

/**
 * The TZID parameter's value AS STORED (quotes included), searched quote-aware and only
 * before the value boundary, so a `;TZID=` inside another parameter's quoted value or inside
 * the value itself never matches.
 *
 * Takes a whole property line or the segment before its colon, property name included; a
 * string with no unquoted colon is read as all parameters. Do not pass a bare
 * `TZID=Europe/Paris` (no property name): segment 0 would then match.
 *
 * Returns `undefined`, never `''`, for an empty `TZID=`, so callers' no-TZID fallback fires.
 * A repeated TZID (malformed per RFC 5545 §3.2): the first wins. Case-sensitive on `TZID`;
 * RFC 5545 §3.1 conformance is #57/#111.
 */
export function extractTzidParam(line: string): string | undefined {
  const boundary = findValueBoundary(line);
  const params = boundary === -1 ? line : line.slice(0, boundary);

  let inQuote = false;
  let segStart = 0;
  const segments: string[] = [];
  for (let i = 0; i < params.length; i++) {
    const ch = params[i];
    if (ch === '"') {
      inQuote = !inQuote;
    } else if (ch === ';' && !inQuote) {
      segments.push(params.slice(segStart, i));
      segStart = i + 1;
    }
  }
  segments.push(params.slice(segStart));

  for (const segment of segments) {
    if (segment.startsWith('TZID=')) {
      const value = segment.slice('TZID='.length);
      return value === '' ? undefined : value;
    }
  }
  return undefined;
}

/**
 * Which content lines start a property of the block itself: not a folded continuation, not a
 * BEGIN:/END: marker, and not inside a nested component. A VALARM's DESCRIPTION, ATTENDEE,
 * DURATION or UID is the alarm's; read as the event's, an update writes the alarm's recipient
 * back as a real ATTENDEE. The block may open with its own BEGIN: line or be bare properties.
 */
function ownPropertyLines(lines: string[]): boolean[] {
  let depth = 0;
  let base: number | undefined;
  return lines.map((text) => {
    // Upper-cased: RFC 5545 §3.1 names are case-insensitive, as hasICalProperty reads them.
    const marker = structuralLine(text)?.toUpperCase();
    if (marker === undefined) return false;
    if (base === undefined && marker !== '') base = marker.startsWith('BEGIN:') ? 1 : 0;
    if (marker.startsWith('BEGIN:')) { depth++; return false; }
    if (marker.startsWith('END:')) { depth--; return false; }
    return depth <= (base ?? 0);
  });
}

/**
 * The first matching property's value in a VEVENT block, unfolded. Whole content lines only
 * (see the line-model comment above): this read decides which record a destroy resolves to.
 *
 * CASE-SENSITIVE on the property name, deliberately, unlike `hasICalProperty`. A wholly
 * lower-cased payload yields no blocks at all and is invisible, which is fail-closed. A
 * mixed-case payload is not, and only `extractVTimezoneBlocks` guards its one such shape; the
 * rest is the RFC conformance audit (#57, #111).
 */
export function parseICalValue(vevent: string, key: string): string | undefined {
  const lines = icalContentLines(vevent).map(l => l.text);
  const own = ownPropertyLines(lines);
  const test = new RegExp(`^${key}[;:]`);

  for (let i = 0; i < lines.length; i++) {
    if (!own[i]) continue;
    const line = lines[i].replace(/\r$/, '');
    if (!test.test(line)) continue;

    let fullLine = line;
    for (let j = i + 1; j < lines.length; j++) {
      if (!isFoldedContinuation(lines[j])) break;
      fullLine += lines[j].substring(1);
    }
    // A lone `\r` survives `icalContentLines`, and the strip above ran on the first physical
    // line only, so strip one from the last continuation too.
    fullLine = fullLine.replace(/\r$/, '');

    const colonIdx = findValueBoundary(fullLine);
    if (colonIdx === -1) return undefined;
    // Not trimmed: whitespace inside a value is significant. Callers needing an exact match
    // (a UID equality check) trim at their own call site.
    return fullLine.substring(colonIdx + 1);
  }

  return undefined;
}

/**
 * Every occurrence of a property as a full unfolded raw line (ATTENDEE, EXDATE and others
 * repeat). Same line model as parseICalValue.
 */
export function parseAllICalProperties(vevent: string, key: string): string[] {
  const lines = icalContentLines(vevent).map(l => l.text);
  const own = ownPropertyLines(lines);
  const regex = new RegExp(`^${key}[;:]`);
  const results: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    if (!own[i]) continue;
    const line = lines[i].replace(/\r$/, '');
    if (!regex.test(line)) continue;

    let fullLine = line;
    for (let j = i + 1; j < lines.length; j++) {
      const next = lines[j];
      if (isFoldedContinuation(next)) {
        fullLine += next.substring(1);
        i = j; // skip continuation lines in outer loop
      } else {
        break;
      }
    }
    results.push(fullLine);
  }

  return results;
}

/**
 * Parse a raw ATTENDEE or ORGANIZER line into a Participant.
 */
export function parseAttendee(rawLine: string): Participant {
  const boundaryIdx = findValueBoundary(rawLine);
  const paramPart = boundaryIdx >= 0 ? rawLine.substring(0, boundaryIdx) : rawLine;
  const valuePart = boundaryIdx >= 0 ? rawLine.substring(boundaryIdx + 1) : '';

  const email = valuePart.replace(/^mailto:/i, '');

  const params: string[] = [];
  let current = '';
  let inQuote = false;
  const firstSemi = paramPart.indexOf(';');
  const paramStr = firstSemi >= 0 ? paramPart.substring(firstSemi + 1) : '';

  for (let i = 0; i < paramStr.length; i++) {
    const ch = paramStr[i];
    if (ch === '"') {
      inQuote = !inQuote;
      current += ch;
    } else if (ch === ';' && !inQuote) {
      if (current) params.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  if (current) params.push(current);

  const result: Participant = { email };

  for (const param of params) {
    const eqIdx = param.indexOf('=');
    if (eqIdx === -1) continue;
    const pName = param.substring(0, eqIdx).toUpperCase();
    let pValue = param.substring(eqIdx + 1);
    if (pValue.startsWith('"') && pValue.endsWith('"')) {
      pValue = pValue.slice(1, -1);
    }

    switch (pName) {
      case 'CN':
        if (pValue) result.name = pValue;
        break;
      case 'PARTSTAT':
        result.status = pValue;
        break;
      case 'ROLE':
        result.role = pValue;
        break;
      case 'CUTYPE':
        result.cutype = pValue;
        break;
      case 'RSVP':
        result.rsvp = pValue.toUpperCase() === 'TRUE';
        break;
    }
  }

  return result;
}

/**
 * Format an iCalendar date/datetime string to ISO 8601.
 * Input formats: 20260320T083000, 20260320T083000Z, 20260324
 * Output: 2026-03-20T08:30:00, 2026-03-20T08:30:00Z, 2026-03-24
 */
export function formatICalDate(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const cleaned = raw.replace(/\r/g, '');

  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }

  const dtMatch = cleaned.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
  if (dtMatch) {
    const [, y, m, d, hh, mm, ss, z] = dtMatch;
    return `${y}-${m}-${d}T${hh}:${mm}:${ss}${z}`;
  }

  return cleaned;
}

/**
 * Convert an ISO 8601 datetime string to iCalendar UTC format.
 * Handles timezone offsets by converting to UTC via Date.
 * Preserves floating times (no offset, no Z) as-is.
 * e.g. "2026-04-07T18:45:00+10:00" → "20260407T084500Z"
 */
export function toICalUTC(isoString: string): string {
  // A plain Error on purpose: a broken internal contract, not caller-fixable input.
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoString)) {
    throw new Error('date-only input must be handled by caller, not passed to toICalUTC');
  }
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(isoString)) {
    return isoString.replace(/[-:]/g, '');
  }
  const d = new Date(isoString);
  // Caller-supplied and unscreened upstream, so it goes through the shared echo.
  if (isNaN(d.getTime())) throw new InvalidInputError(`Invalid date: "${echoCallerText(isoString)}"`);
  return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

export function detectLineEnding(data: string): string {
  return data.includes('\r\n') ? '\r\n' : '\n';
}

/**
 * Replace or insert an iCal property within the first VEVENT block.
 * Operates on lines only within the first BEGIN:VEVENT/END:VEVENT pair,
 * skipping nested sub-components (VALARM etc.).
 * @param newLine Pre-folded replacement line, or null to remove the property.
 */
export function replaceICalProperty(icalData: string, key: string, newLine: string | null): string {
  if (!icalData) throw new Error('replaceICalProperty: empty input');

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // structuralLine, not `.trim()`: a trimmed compare reads a FOLDED continuation
  // (` BEGIN:VEVENT`) as a component marker.
  const veventStart = lines.findIndex(l => structuralLine(l) === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('replaceICalProperty: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = structuralLine(lines[i]);
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) depth++;
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) {
        veventEnd = i;
        break;
      }
    }
  }
  if (veventEnd === -1) throw new Error('replaceICalProperty: END:VEVENT not found');

  const propRegex = new RegExp(`^${key}[;:]`);
  let foundIdx = -1;
  let foundEndIdx = -1;
  let nestDepth = 0;

  for (let i = veventStart + 1; i < veventEnd; i++) {
    const trimmed = structuralLine(lines[i]);
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) { nestDepth++; continue; }
    if (trimmed.startsWith('END:')) { nestDepth--; continue; }
    if (nestDepth > 0) continue;

    if (propRegex.test(lines[i])) {
      foundIdx = i;
      foundEndIdx = i + 1;
      while (foundEndIdx < veventEnd && (lines[foundEndIdx].startsWith(' ') || lines[foundEndIdx].startsWith('\t'))) {
        foundEndIdx++;
      }
      break;
    }
  }

  if (foundIdx >= 0) {
    const newLines = newLine !== null ? newLine.split(/\r?\n/) : [];
    lines.splice(foundIdx, foundEndIdx - foundIdx, ...newLines);
  } else if (newLine !== null) {
    // Insert before the first sub-component (e.g. VALARM) when present —
    // RFC 5545 ABNF is `eventprop *alarmc`, so properties must precede alarms.
    let insertAt = veventEnd;
    for (let i = veventStart + 1; i < veventEnd; i++) {
      // structuralLine: a trimmed compare would read a folded ` BEGIN:...` as a sub-component
      // and splice the new property into the middle of the one above.
      if (structuralLine(lines[i])?.startsWith('BEGIN:')) { insertAt = i; break; }
    }
    const newLines = newLine.split(/\r?\n/);
    lines.splice(insertAt, 0, ...newLines);
  }

  return lines.join(lineEnding);
}

/**
 * Remove ALL occurrences of a property within the first VEVENT block.
 * Skips nested sub-components.
 */
export function removeAllICalProperties(icalData: string, key: string): string {
  if (!icalData) throw new Error('removeAllICalProperties: empty input');

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // structuralLine, not `.trim()`: a trimmed compare reads a FOLDED continuation
  // (` BEGIN:VEVENT`) as a component marker.
  const veventStart = lines.findIndex(l => structuralLine(l) === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('removeAllICalProperties: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = structuralLine(lines[i]);
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) depth++;
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) {
        veventEnd = i;
        break;
      }
    }
  }
  if (veventEnd === -1) throw new Error('removeAllICalProperties: END:VEVENT not found');

  const propRegex = new RegExp(`^${key}[;:]`);
  const toRemove: Array<[number, number]> = [];
  let nestDepth = 0;

  for (let i = veventStart + 1; i < veventEnd; i++) {
    const trimmed = structuralLine(lines[i]);
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) { nestDepth++; continue; }
    if (trimmed.startsWith('END:')) { nestDepth--; continue; }
    if (nestDepth > 0) continue;

    if (propRegex.test(lines[i])) {
      const startIdx = i;
      let endIdx = i + 1;
      while (endIdx < veventEnd && (lines[endIdx].startsWith(' ') || lines[endIdx].startsWith('\t'))) {
        endIdx++;
      }
      toRemove.push([startIdx, endIdx - startIdx]);
      i = endIdx - 1; // skip past continuation lines
    }
  }

  // Remove in reverse to preserve indices
  for (let r = toRemove.length - 1; r >= 0; r--) {
    lines.splice(toRemove[r][0], toRemove[r][1]);
  }

  return lines.join(lineEnding);
}

/**
 * Insert a property line into the first VEVENT block, before any sub-components
 * (VALARM etc.) per RFC 5545 ABNF (eventprop before alarmc).
 * Falls back to before END:VEVENT if no sub-components exist.
 */
export function insertBeforeEndVEvent(icalData: string, newLine: string): string {
  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // structuralLine, not `.trim()`: a trimmed compare reads a FOLDED continuation
  // (` BEGIN:VEVENT`) as a component marker.
  const veventStart = lines.findIndex(l => structuralLine(l) === 'BEGIN:VEVENT');
  if (veventStart === -1) throw new Error('insertBeforeEndVEvent: BEGIN:VEVENT not found');

  let veventEnd = -1;
  let firstSubComponent = -1;
  let depth = 0;
  for (let i = veventStart; i < lines.length; i++) {
    const trimmed = structuralLine(lines[i]);
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) {
      depth++;
      // Track first nested sub-component (depth 2 = inside VEVENT)
      if (depth === 2 && firstSubComponent === -1) {
        firstSubComponent = i;
      }
    }
    if (trimmed.startsWith('END:')) {
      depth--;
      if (depth === 0) { veventEnd = i; break; }
    }
  }
  if (veventEnd === -1) throw new Error('insertBeforeEndVEvent: END:VEVENT not found');

  const insertIdx = firstSubComponent !== -1 ? firstSubComponent : veventEnd;
  const newLines = newLine.split(/\r?\n/);
  lines.splice(insertIdx, 0, ...newLines);
  return lines.join(lineEnding);
}

/**
 * Remove orphaned VTIMEZONE blocks whose TZID has no remaining references
 * in the file (outside VTIMEZONE blocks themselves).
 */
export function removeOrphanedVTimezones(icalData: string): string {
  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  const tzBlocks = extractVTimezoneBlocks(lines);
  if (tzBlocks.length === 0) return icalData;

  const excludedLines = new Set<number>();
  for (const block of tzBlocks) {
    for (let i = block.start; i <= block.end; i++) excludedLines.add(i);
  }
  const nonTzLines = lines.filter((_, i) => !excludedLines.has(i));
  // Unfold before scanning so a reference split across a folded line isn't missed, but stay in
  // LINES: a TZID parameter is a property of one line, and the parser below reads one at a time.
  const unfoldedNonTzLines = nonTzLines.join('\n').replace(/\n[ \t]/g, '').split('\n');

  // A reference is decided by the same parser that reads a TZID everywhere else (#187). A
  // substring search would count a `;TZID=` inside a quoted parameter or a DESCRIPTION, and
  // match `Europe/Pari` as a prefix of `Europe/Paris`.
  const referenced = new Set<string>();
  for (const line of unfoldedNonTzLines) {
    const tzid = extractTzidParam(line);
    if (tzid !== undefined) referenced.add(tzid.replace(/^"|"$/g, ''));
  }

  const orphaned = tzBlocks.filter(tz => tz.tzid !== '' && !referenced.has(tz.tzid));

  for (let i = orphaned.length - 1; i >= 0; i--) {
    lines.splice(orphaned[i].start, orphaned[i].end - orphaned[i].start + 1);
  }

  return lines.join(lineEnding);
}

/**
 * Remove exception VEVENT blocks whose RECURRENCE-ID matches one of the orphaned dates.
 * Operates on the full iCal string. Never touches the master VEVENT (no RECURRENCE-ID).
 *
 * NO PRODUCTION CALLER, kept deliberately: it is the primitive a series-aware update needs
 * back (#146), and it is pure and tested.
 */
export function removeExceptionVEvents(icalData: string, orphanedRecurrenceIds: Date[]): string {
  if (orphanedRecurrenceIds.length === 0) return icalData;

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  const veventBlocks: Array<{ start: number; end: number; recurrenceId?: string }> = [];
  for (let i = 0; i < lines.length; i++) {
    if (structuralLine(lines[i]) === 'BEGIN:VEVENT') {
      const blockStart = i;
      for (let j = i + 1; j < lines.length; j++) {
        if (structuralLine(lines[j]) === 'END:VEVENT') {
          const veventText = lines.slice(blockStart, j + 1).join('\n');
          // Trimmed here because this feeds formatICalDate, which anchors its pattern.
          const recId = parseICalValue(veventText, 'RECURRENCE-ID')?.trim();
          veventBlocks.push({ start: blockStart, end: j, recurrenceId: recId });
          i = j;
          break;
        }
      }
    }
  }

  // Compared as ISO strings in a fixed UTC frame, so floating and UTC values are not
  // reinterpreted in the process's local zone.
  const orphanedDateStrings = orphanedRecurrenceIds.map(d => {
    return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
  });
  const toRemove = veventBlocks.filter(block => {
    if (!block.recurrenceId) return false; // master VEVENT — never remove
    const recIdFormatted = formatICalDate(block.recurrenceId);
    if (!recIdFormatted) return false;
    const recDate = parseICalDateAsUTC(recIdFormatted);
    const recDateStr = recDate.toISOString().replace(/\.\d{3}Z$/, 'Z');
    return orphanedDateStrings.includes(recDateStr);
  });

  for (let i = toRemove.length - 1; i >= 0; i--) {
    lines.splice(toRemove[i].start, toRemove[i].end - toRemove[i].start + 1);
  }

  return lines.join(lineEnding);
}

/** RFC 5545 §3.3.6: [+/-]P[nW | nDTnHnMnS]. The one pattern every DURATION parse in this file matches against. */
const ICAL_DURATION_RE = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/;

interface ParsedICalDuration {
  sign: 1 | -1;
  weeks: number;
  days: number;
  hours: number;
  minutes: number;
  seconds: number;
}

/**
 * Parse a DURATION value into its components, or undefined if malformed. Shared by
 * `parseICalDuration` and `resolveDurationSpanEndMs` so the two agree on validity.
 */
function parseICalDurationComponents(duration: string): ParsedICalDuration | undefined {
  const m = duration.match(ICAL_DURATION_RE);
  if (!m) return undefined;

  const [, sign, weeks, days, hours, minutes, seconds] = m;

  // At least one component must be present (reject bare "P")
  if (!weeks && !days && !hours && !minutes && !seconds) return undefined;
  // If T is present in input, at least one time component must exist (reject "P1DT")
  if (duration.includes('T') && !hours && !minutes && !seconds) return undefined;

  return {
    sign: sign === '-' ? -1 : 1,
    weeks: parseInt(weeks || '0', 10),
    days: parseInt(days || '0', 10),
    hours: parseInt(hours || '0', 10),
    minutes: parseInt(minutes || '0', 10),
    seconds: parseInt(seconds || '0', 10),
  };
}

/**
 * The end a DURATION implies from `start`, in start's format, or undefined if malformed.
 * A plain millisecond add, which is wrong across a DST transition in the event's zone (#196);
 * `resolveDurationSpanEndMs` does the RFC 5545 §3.3.6 nominal-day/exact-time split.
 */
export function parseICalDuration(duration: string, start: string): string | undefined {
  const parsed = parseICalDurationComponents(duration);
  if (!parsed) return undefined;
  const { sign, weeks, days, hours, minutes, seconds } = parsed;

  const ms = (weeks * 7 * 86400000) + (days * 86400000) + (hours * 3600000) + (minutes * 60000) + (seconds * 1000);

  const startDate = new Date(start);
  if (isNaN(startDate.getTime())) return undefined;

  const endMs = startDate.getTime() + sign * ms;
  const endDate = new Date(endMs);

  if (/^\d{4}-\d{2}-\d{2}$/.test(start)) {
    return endDate.toISOString().slice(0, 10);
  }

  // `new Date()` reads a floating time as process-local, so do the arithmetic in UTC by hand.
  const isFloating = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(start);
  if (isFloating) {
    const [datePart, timePart] = start.split('T');
    const [y, mo, d] = datePart.split('-').map(Number);
    const [h, mi, s] = timePart.split(':').map(Number);
    const utcStart = Date.UTC(y, mo - 1, d, h, mi, s);
    const utcEnd = utcStart + sign * ms;
    const e = new Date(utcEnd);
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${e.getUTCFullYear()}-${pad(e.getUTCMonth() + 1)}-${pad(e.getUTCDate())}T${pad(e.getUTCHours())}:${pad(e.getUTCMinutes())}:${pad(e.getUTCSeconds())}`;
  }

  return endDate.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** Every VEVENT block in an iCalendar payload, in the order the server serialised them. */
function extractAllVEvents(data: string): string[] {
  return extractVEventBlocks(data);
}

/**
 * Parse EVERY event a single CalDAV resource represents.
 *
 *   EXPANDED (a time-range query sent with `expand`): one VEVENT per in-window occurrence,
 *     several in ONE blob (scripts/probes/calendar-expand.probe.mjs); every block is an event.
 *   UNEXPANDED: the series MASTER plus optional overrides; only the master is emitted.
 *   NO VEVENT: a minimal event.
 *
 * WHICH SHAPE IT IS, IS TOLD, NEVER INFERRED FROM CONTENT. Cyrus's expansion (`expand_cb`,
 * imap/http_caldav.c) emits the series' first instance with no RRULE and no RECURRENCE-ID, so
 * a "no master means expanded" sniff takes that instance for a master and drops every sibling.
 *
 * The master is the block WITHOUT a RECURRENCE-ID, not the first block: RFC 5545 does not fix
 * component order. `normalizeMasterVEventFirst` applies the same rule on the write path; the
 * two stay separate implementations until the read/write divergence in #102 is settled.
 */
export function parseCalendarObjects(
  obj: DAVCalendarObject,
  options?: { includeParticipants?: boolean; expanded?: boolean; configuredZone?: string; blocks?: string[]; includeDefaultTransparency?: boolean },
): CalendarEvent[] {
  // `blocks` lets the listing path, which counts blocks first (CALENDAR_MAX_OCCURRENCES_PER_SERIES),
  // pass its extraction through: one walk of a large payload, one definition of a block.
  const blocks = options?.blocks ?? extractAllVEvents(obj.data || '');
  if (blocks.length === 0) {
    return [{
      id: obj.url || '',
      url: obj.url || '',
      title: 'Untitled',
    }];
  }

  if (options?.expanded) {
    const events = blocks.map(b => parseVEvent(obj, b, options));
    if (blockCountProvesSeries(blocks)) {
      // Marks the first instance too, which Cyrus leaves with no RRULE and no RECURRENCE-ID.
      for (const event of events) event.isRecurring = true;
    }
    return events;
  }

  const master = blocks.find(b => !hasRecurrenceId(b));

  // No master: every block is a detached override naming a distinct instance.
  if (!master) return blocks.map(b => parseVEvent(obj, b, options));

  return [parseVEvent(obj, master, options)];
}

function hasRecurrenceId(block: string): boolean {
  return hasICalProperty(block, 'RECURRENCE-ID');
}

/**
 * Whether an EXPANDED blob's block list proves the resource is a repeating series. A `false`
 * means "ask", not "one-off": a lone unmarked block may be a series' only in-window instance,
 * which `settleAmbiguousRecurrence` resolves.
 */
function blockCountProvesSeries(blocks: string[]): boolean {
  return blocks.length > 1 || blocks.some(hasRecurrenceId);
}

/**
 * Parse ONE event out of a resource: for an unexpanded resource, the series master.
 */
export function parseCalendarObject(obj: DAVCalendarObject, options?: { includeParticipants?: boolean; configuredZone?: string; includeDefaultTransparency?: boolean }): CalendarEvent {
  return parseCalendarObjects(obj, options)[0];
}

function parseVEvent(
  obj: DAVCalendarObject,
  vevent: string,
  options?: { includeParticipants?: boolean; configuredZone?: string; includeDefaultTransparency?: boolean },
): CalendarEvent {
  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  // Trimmed here because both feed formatICalDate, which anchors its pattern.
  const rawStart = parseICalValue(vevent, 'DTSTART')?.trim();
  let rawEnd = parseICalValue(vevent, 'DTEND')?.trim();
  const location = parseICalValue(vevent, 'LOCATION');
  // Trimmed because findCalendarObjectByUID matches it by exact equality.
  const uid = parseICalValue(vevent, 'UID')?.trim() || obj.url || '';

  // Injectable so a test can pin the zone rather than depend on the host's.
  const configuredZone = options?.configuredZone ?? resolveUsableTimezone(undefined);

  if (!rawEnd && rawStart) {
    // Trimmed here because this feeds parseICalDuration, which anchors its pattern.
    const rawDuration = parseICalValue(vevent, 'DURATION')?.trim();
    if (rawDuration) {
      const startIso = formatICalDate(rawStart);
      if (startIso) {
        const computedEnd = parseICalDuration(rawDuration, startIso);
        if (computedEnd) {
          const event: CalendarEvent = {
            id: uid,
            url: obj.url || '',
            title: unescapeICalText(title),
            description: description ? unescapeICalText(description) : undefined,
            start: formatICalDate(rawStart),
            end: computedEnd,
            location: location ? unescapeICalText(location) : undefined,
          };
          // Start's zone only, never attachZoneFields: see its doc comment.
          attachStartZone(event, vevent, configuredZone);
          attachTransparency(event, vevent, options?.includeDefaultTransparency);
          addRecurrenceToEvent(event, vevent);
          if (options?.includeParticipants) {
            addParticipantsToEvent(event, vevent);
          }
          return event;
        }
      }
    }
  }

  const event: CalendarEvent = {
    id: uid,
    url: obj.url || '',
    title: unescapeICalText(title),
    description: description ? unescapeICalText(description) : undefined,
    start: formatICalDate(rawStart),
    end: formatICalDate(rawEnd),
    location: location ? unescapeICalText(location) : undefined,
  };

  attachZoneFields(event, vevent, configuredZone);
  attachTransparency(event, vevent, options?.includeDefaultTransparency);
  addRecurrenceToEvent(event, vevent);
  if (options?.includeParticipants) {
    addParticipantsToEvent(event, vevent);
  }

  return event;
}

/**
 * Put `transparency` on an event, or decide not to (#194). BOTH return paths of `parseVEvent`
 * must call this: they build separate literals, and the Fastmail client writes both end
 * shapes (docs/fastmail-action-availability.md).
 */
function attachTransparency(event: CalendarEvent, vevent: string, includeDefault?: boolean): void {
  const transparency = readTransparency(vevent);
  if (includeDefault || transparency !== 'busy') {
    event.transparency = transparency;
  }
}

// A DTSTART/DTEND property's zone, for `timeZone`/`endTimeZone` (#139). `absent` is a missing
// property, which `describeDateProperty` has no line to classify.
type ZoneDescriptor =
  | { kind: 'tzid'; name: string }
  | { kind: 'floating' }
  | { kind: 'none' }
  | { kind: 'absent' };

/**
 * Classify a DTSTART/DTEND property's zone from its raw line(s). Built on
 * `describeDateProperty` so the read path and the write path's consistency check agree on what
 * `zoned` is. Its `date` and `utc` frames both become `none`: neither carries a zone name.
 */
function classifyZoneFromLines(rawLines: string[]): ZoneDescriptor {
  if (rawLines.length === 0) return { kind: 'absent' };
  const d = describeDateProperty(rawLines[0]);
  if (d.frame === 'zoned') return { kind: 'tzid', name: d.tzid! };
  if (d.frame === 'floating') return { kind: 'floating' };
  return { kind: 'none' };
}

// Comparison-only; the stored spelling is what gets emitted. Strips the one leading '/' RFC 5545
// §3.2.19 allows, as libical does; a vendor-prefixed TZID still compares unequal, the safe
// direction. Then canonicalises through `canonicalZoneName` (#157), so an alias such as 'NZ'
// equals the 'Pacific/Auckland' this server writes and a round trip is not falsely rejected
// as a two-zone event (#139). A name ICU cannot resolve is compared as the stripped string.
function normalizeZoneForComparison(name: string): string {
  const trimmed = name.trim();
  const stripped = trimmed.startsWith('/') ? trimmed.slice(1) : trimmed;
  return canonicalZoneName(stripped);
}

// Normalised as above, then case-insensitive, and no looser: an extra emitted field is the
// safe direction, calling two different zones the same is not.
function zoneNamesEqual(a: string, b: string): boolean {
  return normalizeZoneForComparison(a).toLowerCase() === normalizeZoneForComparison(b).toLowerCase();
}

function zoneDescriptorsEqual(a: ZoneDescriptor, b: ZoneDescriptor): boolean {
  if (a.kind !== b.kind) return false;
  if (a.kind === 'tzid' && b.kind === 'tzid') return zoneNamesEqual(a.name, b.name);
  return true;
}

function attachStartZone(event: CalendarEvent, vevent: string, configuredZone: string): ZoneDescriptor {
  const startDesc = classifyZoneFromLines(parseAllICalProperties(vevent, 'DTSTART'));

  if (startDesc.kind === 'tzid') {
    // Trimmed so downstream isUsableTimezone can resolve it; not slash-normalized, since the
    // stored spelling is what named the zone.
    if (!zoneNamesEqual(startDesc.name, configuredZone)) event.timeZone = startDesc.name.trim();
  } else if (startDesc.kind === 'floating') {
    event.timeZone = null;
  }
  return startDesc;
}

function attachEndZone(event: CalendarEvent, startDesc: ZoneDescriptor, vevent: string, configuredZone: string): void {
  const endDesc = classifyZoneFromLines(parseAllICalProperties(vevent, 'DTEND'));
  if (endDesc.kind === 'absent' || endDesc.kind === 'none') return;

  const compareDesc: ZoneDescriptor = startDesc.kind === 'absent'
    ? { kind: 'tzid', name: configuredZone }
    : startDesc;
  if (zoneDescriptorsEqual(endDesc, compareDesc)) return;

  // Trimmed as for `timeZone` above.
  event.endTimeZone = endDesc.kind === 'tzid' ? endDesc.name.trim() : null;
}

/**
 * Set `timeZone`, and where applicable `endTimeZone`, on a parsed event from the VEVENT's raw
 * DTSTART/DTEND lines (#139): the one rule for the emit matrix documented on `CalendarEvent`
 * and in docs/conventions.md. `endTimeZone` is relative to start, or to the configured zone
 * when start is absent.
 *
 * The DURATION branch in `parseVEvent` calls `attachStartZone` alone, deliberately: an empty
 * `DTEND;TZID=Europe/Paris:` takes that branch but still reads as `zoned`, and would leak a zone
 * that never computed `end`. A DURATION-computed end shares start's frame by construction.
 */
function attachZoneFields(event: CalendarEvent, vevent: string, configuredZone: string): void {
  const startDesc = attachStartZone(event, vevent, configuredZone);
  attachEndZone(event, startDesc, vevent, configuredZone);
}

/**
 * Attach the recurrence markers that say what KIND of date `start` is (#64). Read from the
 * block, not the resource, since that is the distinction reported. EVERY RDATE line is read
 * (#162): RFC 5545 §3.8.5.2 lets the property repeat.
 */
function addRecurrenceToEvent(event: CalendarEvent, vevent: string): void {
  const rrule = parseICalValue(vevent, 'RRULE');
  // Trimmed here because this feeds formatICalDate, which anchors its pattern.
  const recurrenceId = parseICalValue(vevent, 'RECURRENCE-ID')?.trim();
  if (recurrenceId) {
    event.isRecurring = true;
    event.recurrenceId = formatICalDate(recurrenceId);
  }
  if (rrule) {
    event.isRecurring = true;
    event.recurrenceRule = rrule;
  }
  // Parameters dropped deliberately; see `CalendarEvent.recurrenceDates`.
  const rdateValues = parseAllICalProperties(vevent, 'RDATE')
    .map(line => {
      const colonIdx = findValueBoundary(line);
      return colonIdx === -1 ? '' : line.substring(colonIdx + 1).trim();
    })
    .filter(value => value.length > 0);
  if (rdateValues.length > 0) {
    event.isRecurring = true;
    event.recurrenceDates = rdateValues.join(',');
  }
}

function addParticipantsToEvent(event: CalendarEvent, vevent: string): void {
  const attendeeLines = parseAllICalProperties(vevent, 'ATTENDEE');
  if (attendeeLines.length > 0) {
    event.participants = attendeeLines.map(parseAttendee);
  }
  const organizerLines = parseAllICalProperties(vevent, 'ORGANIZER');
  if (organizerLines.length > 0) {
    event.organizer = parseAttendee(organizerLines[0]);
  }
}

/**
 * Unescape an iCalendar text value (RFC 5545 §3.3.11). One left-to-right pass: chained
 * .replace() calls would turn the escaped backslash in "\\n" into "\<newline>".
 */
export function unescapeICalText(value: string): string {
  return value.replace(/\\(\\|;|,|[nN])/g, (_, ch) => {
    if (ch === 'n' || ch === 'N') return '\n';
    if (ch === ',') return ',';
    if (ch === ';') return ';';
    return '\\';
  });
}

/**
 * Escape a text value for use in an iCalendar property (RFC 5545 §3.3.11).
 */
export function escapeICalText(value: string): string {
  return value
    // A bare CR would otherwise pass through and act as a line terminator downstream.
    .replace(/\r\n?/g, '\n')
    // HTAB is legal in iCal TEXT; LF is escaped below.
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
    .replace(/\\/g, '\\\\')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,')
    .replace(/\n/g, '\\n');
}

/**
 * Validate and serialize a date/datetime value for use in DTSTART/DTEND.
 * Accepts only:
 *   - YYYY-MM-DD                       (date-only)
 *   - YYYY-MM-DDTHH:MM:SS              (floating local)
 *   - YYYY-MM-DDTHH:MM:SSZ             (UTC)
 *   - YYYY-MM-DDTHH:MM:SS+HH:MM        (with offset, normalized to UTC; +HHMM also parses)
 * Returns the ICS form (`YYYYMMDD`, or a datetime with `Z` for instants).
 *
 * The ONLY thing that turns a caller-supplied start/end into an iCal value;
 * formatDateTimeProperty never parses the caller's string itself. The shapes are anchored so
 * nothing reaches `new Date()`'s legacy parser, which reads `2026/04/18` as HOST-LOCAL
 * midnight, and an impossible day (`2026-02-31`) is refused rather than rolled into the next
 * month. Same two traps as coerceUtcDate in src/coerce.ts.
 */
export function validateAndFormatICalDate(value: string, fieldName: string): string {
  if (typeof value !== 'string') {
    throw new InvalidInputError(`${fieldName} must be a string`);
  }
  if (/[\x00-\x1F\x7F]/.test(value)) {
    throw new InvalidInputError(`${fieldName} contains control characters`);
  }
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    assertRealCalendarDate(trimmed, trimmed, fieldName);
    return trimmed.replace(/-/g, '');
  }
  // The pattern admits any 2-4 digit offset; V8 then parses only +/-HH:MM and +/-HHMM.
  const dtMatch = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})(Z|[+-]\d{2}:?\d{0,2})?$/.exec(trimmed);
  if (!dtMatch) {
    // The only rejection here quoting an unconstrained value (U+2028 passes the control guard),
    // so it echoes; the others quote values already matched to an anchored shape.
    throw new InvalidInputError(`${fieldName} must be ISO-8601 date or datetime (got: "${echoCallerText(trimmed)}")`);
  }
  const [, datePart, timePart, tz] = dtMatch;
  // The date part alone: an offset legitimately moves the UTC date.
  assertRealCalendarDate(datePart, trimmed, fieldName);
  // RFC 5545 §3.3.12 hours run 00-23. V8 reads T24:00:00 as next-day midnight, and the floating
  // form would be written verbatim; a leap second (60) has no instant to normalise to.
  const [hh, mm, ss] = timePart.split(':').map(Number);
  if (hh > 23 || mm > 59 || ss > 59) {
    throw new InvalidInputError(`${fieldName} has a time out of range; hours run 00-23 and minutes and seconds 00-59 (got: ${trimmed.slice(0, 60)})`);
  }
  const isoForParse = `${datePart}T${timePart}${tz || ''}`;
  const d = new Date(isoForParse);
  if (Number.isNaN(d.getTime())) {
    throw new InvalidInputError(`${fieldName} is not a valid datetime (got: ${trimmed.slice(0, 60)})`);
  }
  if (!tz) {
    return `${datePart.replace(/-/g, '')}T${timePart.replace(/:/g, '')}`;
  }
  const utc = d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  return utc;
}

/**
 * Reject a YYYY-MM-DD that names a day its month does not have; `new Date` would silently
 * roll `2026-02-31` to 3 March. A rolled-over date no longer round-trips through toISOString.
 *
 * @param echo the caller's whole value, quoted in the message.
 */
function assertRealCalendarDate(datePart: string, echo: string, fieldName: string): void {
  const probe = new Date(`${datePart}T00:00:00Z`);
  if (Number.isNaN(probe.getTime()) || !probe.toISOString().startsWith(datePart)) {
    throw new InvalidInputError(`${fieldName} is not a real calendar date (got: ${echo.slice(0, 60)})`);
  }
}

/**
 * FREE/BUSY TRANSPARENCY (#194). Callers say `busy`/`free`; iCalendar stores RFC 5545
 * §3.8.2.7's `OPAQUE`/`TRANSPARENT`. The iCal spellings stay in this file and are not accepted
 * as parameter aliases. The read direction is not total; see `readTransparency`.
 */
export const TRANSPARENCY_VALUES = ['busy', 'free'] as const;
export type Transparency = typeof TRANSPARENCY_VALUES[number];

const ICAL_TRANSP: Record<Transparency, string> = { busy: 'OPAQUE', free: 'TRANSPARENT' };
const TRANSPARENCY_BY_ICAL_TRANSP: Record<string, Transparency> = { OPAQUE: 'busy', TRANSPARENT: 'free' };

/**
 * The one write-side mapping, shared by the all-day default and an explicit `transparency`.
 */
function transpLine(transparency: Transparency): string {
  return `TRANSP:${ICAL_TRANSP[transparency]}`;
}

/**
 * A caller's `transparency` argument, resolved to one of the two values this server accepts.
 *
 * WHITELIST, NOT AN ESCAPE: the line is built by concatenation, so only one of two literals
 * this file owns ever reaches the payload, never the caller's text.
 *
 * The trim and case-fold are a backstop for a non-validating client and MUST NOT be
 * advertised: the schema's closed enum stops a validating client first. `OPAQUE`/`TRANSPARENT`
 * are refused too. The refusal echoes an unvalidated value (#190).
 */
export function normalizeTransparency(value: unknown, fieldName = 'transparency'): Transparency {
  const text = typeof value === 'string' ? value.trim().toLowerCase() : undefined;
  const match = TRANSPARENCY_VALUES.find(v => v === text);
  if (!match) {
    throw new InvalidInputError(
      `${fieldName} must be "busy" or "free" (got: "${echoCallerText(value)}"). ` +
      'busy blocks this event\'s time in your free/busy; free leaves you showing as available.',
    );
  }
  return match;
}

/**
 * What an event's `TRANSP` property says, for the `transparency` response field.
 *
 * An absent or empty `TRANSP` means `busy` (RFC 5545 §3.8.2.7 default). A token neither
 * spelling covers is reported verbatim, trimmed: folding it into `busy` would silently drop it.
 *
 * CASE-INSENSITIVE and unfolded, like `hasICalProperty` and unlike `parseICalValue`: a miss
 * here fails OPEN, reporting busy for a free event.
 *
 * FLAT, where `replaceICalProperty` tracks `nestDepth`, so a `TRANSP` inside a VALARM is read
 * here but not cleared by `clearFields`. Left: it needs an already-malformed record.
 */
function readTransparency(vevent: string): string {
  const name = /^TRANSP[;:]/i;
  const line = unfoldedICalLines(vevent).find(l => name.test(l));
  const boundary = line === undefined ? -1 : findValueBoundary(line);
  // §3.1 case-insensitivity applies to the value as well as the property name.
  const raw = boundary === -1 ? undefined : line!.slice(boundary + 1).trim();
  if (!raw) return 'busy';
  return TRANSPARENCY_BY_ICAL_TRANSP[raw.toUpperCase()] ?? raw;
}

/**
 * Validate an email address for use in ATTENDEE lines, against iCal property injection.
 * Rejections are caller-fixable; the configured username goes through
 * validateOrganizerUsername instead, which rethrows as a plain Error.
 */
export function validateAttendeeEmail(email: string): void {
  if (!email || typeof email !== 'string') {
    throw new InvalidInputError('Participant email is required');
  }
  // Both refusals echo: neither shape check screens line separators out of the quoted value.
  if (!/^[^@]+@[^@]+$/.test(email)) {
    throw new InvalidInputError(`Invalid participant email: "${echoCallerText(email)}"`);
  }
  // A CRITERION, not a whitelist: RFC 5322 specials (a route, display name, second address or
  // parameter delimiter), any whitespace, and every Unicode category C code point.
  if (/[()<>[\]:;\\,"]|\s|\p{C}/u.test(email)) {
    throw new InvalidInputError(`Invalid participant email (contains illegal characters): "${echoCallerText(email)}"`);
  }
}

/**
 * The participant addr-spec check, applied to the configured CalDAV username (embedded in the
 * ORGANIZER line). Rethrown as a plain Error: server configuration is not caller-fixable.
 */
function validateOrganizerUsername(username: string): void {
  try {
    validateAttendeeEmail(username);
  } catch (e) {
    // `detail` is already echoed at its throw; do not echo the finished sentence
    // (docs/conventions.md, sanitise the value, never the sentence).
    const detail = e instanceof Error ? e.message : String(e);
    throw new Error(`The configured CalDAV username is not usable as an ORGANIZER address. ${detail}`);
  }
}

/**
 * Quote a CN parameter value per RFC 5545 §3.2: DQUOTE quoting, not backslash escaping.
 * RFC 5545 cannot escape a DQUOTE inside one, so it becomes a single quote, as Python
 * icalendar and Outlook do (RFC 6868 caret encoding is poorly adopted).
 */
export function quoteParamValue(value: string): string {
  let cleaned = value.replace(/[\r\n]+/g, ' ');
  // Controls (HTAB is legal here), DEL, C1, and the bidi overrides and isolates, which let a
  // name display as a different address. LRM/RLM are deliberately kept: they cannot reorder
  // surrounding text and occur legitimately in Arabic and Hebrew names.
  cleaned = cleaned.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\u202A-\u202E\u2066-\u2069]/g, '');
  cleaned = cleaned.replace(/"/g, "'");
  if (/[,;:]/.test(cleaned) || value.includes('"')) {
    return `"${cleaned}"`;
  }
  return cleaned;
}

// Where a freshly written TZID came from (#157), so a frame-mismatch error does not name a zone
// the caller never wrote:
//   - 'caller' - the caller passed `timeZone` on this call.
//   - 'stored' - inherited from the event's existing TZID.
//   - 'default' - nothing to inherit; `create` filled in the configured zone.
export type TzidSource = 'caller' | 'stored' | 'default';

interface FormattedDateProperty {
  line: string;
  /** Set only when a TZID was actually written onto `line`. */
  tzidSource?: TzidSource;
}

/**
 * Format a start/end input value into the correct iCal property line.
 * Handles four cases:
 * 1. Date-only (2026-04-01) → DTXXX;VALUE=DATE:20260401
 * 2. Floating time (2026-03-20T09:30:00), `callerZone` given → DTXXX;TZID=<callerZone>:...
 * 3. Floating time, no `callerZone` → preserve original TZID, else `defaultZone`, else floating
 * 4. UTC/offset (2026-03-20T09:30:00Z) → DTXXX:20260320T093000Z
 *
 * The form is decided from validateAndFormatICalDate's serialized value, never by re-reading
 * the input, so the two cannot disagree about what the caller wrote.
 *
 * `create_calendar_event` passes the configured zone as `defaultZone`; `update_calendar_event`
 * passes it too, except on an event whose stored start is floating, which stays floating
 * (docs/conventions.md).
 *
 * Exported so `tzidSource` is unit-testable: consumers only distinguish 'default' (#157, #102).
 */
export function formatDateTimeProperty(
  propName: string,
  value: string,
  originalVevent: string | null,
  lineEnding: string,
  callerZone?: string,
  defaultZone?: string
): FormattedDateProperty {
  const serialized = validateAndFormatICalDate(value, propName);

  if (/^\d{8}$/.test(serialized)) {
    return { line: foldICalLine(`${propName};VALUE=DATE:${serialized}`, lineEnding) };
  }

  if (!serialized.endsWith('Z')) {
    // Interpolated unescaped, sound only because `validateCallerTimezone` already returned
    // ICU's canonical spelling, which cannot contain `:`, `;`, `"`, CR or LF. That does not
    // extend to the stored TZID below.
    if (callerZone) {
      return { line: foldICalLine(`${propName};TZID=${callerZone}:${serialized}`, lineEnding), tzidSource: 'caller' };
    }
    if (originalVevent) {
      const rawLines = parseAllICalProperties(originalVevent, propName);
      if (rawLines.length > 0) {
        const storedTzid = extractTzidParam(rawLines[0]);
        if (storedTzid !== undefined) {
          return { line: foldICalLine(`${propName};TZID=${storedTzid}:${serialized}`, lineEnding), tzidSource: 'stored' };
        }
      }
      // A DURATION-based event has no DTEND TZID to inherit, so take DTSTART's.
      if (propName === 'DTEND') {
        const startLines = parseAllICalProperties(originalVevent, 'DTSTART');
        if (startLines.length > 0) {
          const storedTzid = extractTzidParam(startLines[0]);
          if (storedTzid !== undefined) {
            return { line: foldICalLine(`${propName};TZID=${storedTzid}:${serialized}`, lineEnding), tzidSource: 'stored' };
          }
        }
      }
    }
    if (defaultZone) {
      return { line: foldICalLine(`${propName};TZID=${defaultZone}:${serialized}`, lineEnding), tzidSource: 'default' };
    }
    return { line: foldICalLine(`${propName}:${serialized}`, lineEnding) };
  }

  return { line: foldICalLine(`${propName}:${serialized}`, lineEnding) };
}

function isDateOnlyProperty(rawLine: string): boolean {
  return /;VALUE=DATE[;:]/.test(rawLine) || /;VALUE=DATE$/.test(rawLine);
}

/**
 * The time frame a DTSTART/DTEND value is expressed in. RFC 5545 §3.3.4/§3.3.5
 * give a date/time property one of these forms, and two properties are only
 * comparable — to each other, or to a wall clock — when they share one:
 *   - date     — VALUE=DATE, an all-day value with no time at all
 *   - floating — a date-time with no zone: "whatever the local clock says",
 *                a different instant for every reader
 *   - utc      — a date-time pinned to an instant (trailing Z; a caller-supplied
 *                offset is normalized to Z before it reaches here)
 *   - zoned    — a date-time carried by a TZID parameter
 */
type DateFrame = 'date' | 'floating' | 'utc' | 'zoned';

interface DatePropertyFrame {
  frame: DateFrame;
  /** TZID parameter value, unquoted. Set only when frame === 'zoned'. */
  tzid?: string;
  /** Where a `zoned` frame's TZID came from (#157); passed in, since a line cannot say. */
  tzidSource?: TzidSource;
  /**
   * Serialized iCal value (20260320 / 20260320T093000 / 20260320T093000Z).
   * Fixed-width within a frame, so lexical order is chronological order.
   */
  value: string;
  /** Human-readable rendering, for error messages. */
  display: string;
}

/**
 * Classify a DTSTART/DTEND property line into its time frame. Runs on the line that will be
 * WRITTEN, not the caller's input, so a floating value that inherited a stored TZID classifies
 * as `zoned`. The stored and the freshly formatted side go through this one classifier so the
 * comparison means something.
 *
 * @param displayOverride the caller's own input, echoed in errors instead of our rendering.
 * @param tzidSourceOverride see `DatePropertyFrame.tzidSource`; omitted for a stored line.
 */
function describeDateProperty(rawLine: string, displayOverride?: string, tzidSourceOverride?: TzidSource): DatePropertyFrame {
  // Unfold first: a long TZID can push the line past the 75-octet fold width.
  const line = rawLine.replace(/\r?\n[ \t]/g, '');
  const colonIdx = findValueBoundary(line);
  const params = colonIdx === -1 ? line : line.slice(0, colonIdx);
  const value = (colonIdx === -1 ? '' : line.slice(colonIdx + 1)).trim();
  const display = displayOverride ?? formatICalDate(value) ?? value;

  // The 8-digit shape is checked alongside VALUE=DATE because a third-party
  // client can write a bare `DTSTART:20260401`; it is still an all-day value.
  if (isDateOnlyProperty(line) || /^\d{8}$/.test(value)) {
    return { frame: 'date', value, display };
  }
  // Before the TZID test: a malformed `;TZID=X:...Z` line names a UTC instant, and the read
  // reports it as one, so every consumer here must too.
  if (/Z$/.test(value)) {
    return { frame: 'utc', value, display };
  }
  const storedTzid = extractTzidParam(params);
  if (storedTzid !== undefined) {
    return { frame: 'zoned', tzid: storedTzid.replace(/^"|"$/g, ''), tzidSource: tzidSourceOverride, value, display };
  }
  return { frame: 'floating', value, display };
}

function describeFrame(d: DatePropertyFrame): string {
  switch (d.frame) {
    case 'date': return 'a date-only (all-day) value';
    case 'floating': return 'a date-time with no time zone';
    case 'utc': return 'a UTC date-time';
    case 'zoned':
      // A `default` TZID (#157) was filled in by this server; do not word it as the caller's.
      if (d.tzidSource === 'default') {
        return `a date-time in the account's configured time zone (${echoCallerText(d.tzid!, ZONE_ECHO_LIMIT)}${etcGmtOffsetNote(d.tzid!)}), applied because you named none`;
      }
      return `a date-time in time zone ${echoCallerText(d.tzid!, ZONE_ECHO_LIMIT)}${etcGmtOffsetNote(d.tzid!)}`;
  }
}

/**
 * Validate the DTSTART/DTEND pair that is about to be written: they must be in
 * the same time frame (RFC 5545 §3.6.1 value-type agreement) and in the right
 * order (RFC 5545 §3.8.2.2, "DTEND MUST be later than DTSTART").
 *
 * Ordering is judged only after the frames agree: a floating/UTC pair has no single duration.
 *
 * Every DTSTART/DTEND value these refusals render goes through `echoCallerText` inside DOUBLE
 * quotes (#190): a side the caller left alone is read verbatim from the stored VEVENT, so an
 * invitation wrote it. `describeFrame`'s TZID is echoed BARE, which is inert only because these
 * sentences single-quote nothing; adding a `'...'` span to any of them reopens it.
 * `suggestion` needs no echo: `nextDay` constrains it to `^\d{4}-\d{2}-\d{2}$`.
 */
function validateDateConsistency(start: DatePropertyFrame, end: DatePropertyFrame): void {
  if (start.frame !== end.frame) {
    throw new InvalidInputError(
      `DTSTART and DTEND must use the same date/time form per RFC 5545 §3.6.1 — ` +
      `start "${echoCallerText(start.display)}" is ${describeFrame(start)} ` +
      `but end "${echoCallerText(end.display)}" is ${describeFrame(end)}. ` +
      `Pass start and end in the same form: both date-only (2026-03-20), both with a zone designator ` +
      `(2026-03-20T09:30:00Z or 2026-03-20T09:30:00+10:00), or both without one (2026-03-20T09:30:00).`
    );
  }

  // Two different zones (a flight) are legal; order them on INSTANTS, not text (#140). An
  // unresolvable vendor TZID (`AUS Eastern Standard Time`) has no instant, so the check stands
  // down rather than reject a record the account already holds. `isUsableTimezone` must run
  // first: `zoneOffsetMsAt` throws on a name it cannot resolve.
  let ordered: boolean;
  if (start.frame === 'zoned' && start.tzid && end.tzid && !zoneNamesEqual(start.tzid, end.tzid)) {
    if (!isUsableTimezone(start.tzid) || !isUsableTimezone(end.tzid)) return;
    const startMs = resolveCalendarInstantMs(formatICalDate(start.value), start.tzid);
    const endMs = resolveCalendarInstantMs(formatICalDate(end.value), end.tzid);
    // An unplaceable value stands down the same way.
    if (Number.isNaN(startMs) || Number.isNaN(endMs)) return;
    ordered = startMs < endMs;
  } else {
    ordered = start.value < end.value;
  }

  if (ordered) return;

  if (start.frame === 'date') {
    const startDate = formatICalDate(start.value) ?? start.value;
    const suggestion = nextDay(startDate);
    throw new InvalidInputError(
      `DTEND is exclusive per RFC 5545 — for a one-day event on "${echoCallerText(startDate)}", ` +
      (suggestion === undefined ? 'pass an end one day later.' : `pass end: "${suggestion}"`)
    );
  }
  throw new InvalidInputError(
    `DTEND must be later than DTSTART per RFC 5545 §3.8.2.2 — start "${echoCallerText(start.display)}" ` +
    `is not before end "${echoCallerText(end.display)}". Pass an end later than the start.`
  );
}

/**
 * The TZID spelling of every usable zoned frame, as WRITTEN, not canonicalised: a reader looks
 * a block up by the spelling on the wire, so `US/Pacific` and `America/Los_Angeles` each need
 * their own block (#166). An unresolvable TZID contributes nothing, as in
 * `validateDateConsistency`.
 */
function referencedZoneTzids(frames: DatePropertyFrame[]): Set<string> {
  const tzids = new Set<string>();
  for (const frame of frames) {
    if (frame.frame === 'zoned' && frame.tzid && isUsableTimezone(frame.tzid)) tzids.add(frame.tzid);
  }
  return tzids;
}

/**
 * The UTC instant of each usable zoned frame, feeding ONE span shared by every zone the VEVENT
 * references (#166): a cross-zone event needs both zones' blocks to cover the same range.
 * An unresolvable value throws rather than being skipped: a span missing an endpoint is wrong.
 * An unresolvable TZID is skipped, not refused: it gets no block (docs/conventions.md, "The
 * VTIMEZONE residual", item 2), and no generated block needs its value.
 */
function collectZoneInstants(labeled: Array<{ label: string; frame: DatePropertyFrame }>): number[] {
  const instants: number[] = [];
  for (const { label, frame } of labeled) {
    if (frame.frame !== 'zoned' || !frame.tzid || !isUsableTimezone(frame.tzid)) continue;
    const ms = resolveCalendarInstantMs(formatICalDate(frame.value), frame.tzid);
    if (Number.isNaN(ms)) {
      throw new InvalidInputError(`Cannot resolve ${label} to an instant for VTIMEZONE generation.`);
    }
    instants.push(ms);
  }
  return instants;
}

/**
 * Every VTIMEZONE block in `lines`, with its TZID and inclusive line-index span. The one scan
 * every caller needing VTIMEZONE boundaries uses, so a malformed resource is refused the same
 * way on every path.
 *
 * Boundaries come from tracking nesting depth, not from scanning to the next END:VTIMEZONE.
 * Refused as too broken to edit: a VTIMEZONE not directly under the VCALENDAR; anything inside
 * one other than STANDARD/DAYLIGHT as direct children (RFC 5545 §3.6.5 allows no deeper
 * component); a BEGIN:/END: hidden behind a fold; and an unterminated block, which takes
 * precedence over "malformed". Component names compare case-insensitively in this scan only
 * (#57, #111). A bare `BEGIN:`/`END:` is ignored.
 */
function extractVTimezoneBlocks(lines: string[]): Array<{ tzid: string; start: number; end: number }> {
  // The depth scan reads PHYSICAL lines (callers splice its physical indices), so a marker split
  // across a fold would be invisible to it. Any logical line that unfolds to a marker is refused
  // up front, even a legal fold no producer seen here emits: fails closed with no
  // logical-to-physical mapping. A property whose text merely CONTAINS a marker is unaffected.
  for (let i = 0; i < lines.length; i++) {
    if (isFoldedContinuation(lines[i])) continue; // only ever reached as part of the group below
    let j = i + 1;
    while (j < lines.length && isFoldedContinuation(lines[j])) j++;
    if (j === i + 1) continue; // this logical line was never folded
    const logical = lines[i] + lines.slice(i + 1, j).map(l => l.slice(1)).join('');
    if (/^(BEGIN|END):/i.test(logical)) {
      // Worded generically: a resource with no VTIMEZONE at all can trip this.
      throw new InvalidInputError(
        'Stored calendar resource has a component boundary hidden behind a folded line.'
      );
    }
  }

  const blocks: Array<{ tzid: string; start: number; end: number }> = [];
  const stack: string[] = [];
  // The VTIMEZONE being tracked. A second BEGIN:VTIMEZONE inside it marks it malformed rather
  // than becoming a second candidate.
  let candidateStart = -1;
  let candidateOpenDepth = -1;
  let malformed = false;

  for (let i = 0; i < lines.length; i++) {
    const structural = structuralLine(lines[i]);
    if (structural === null) continue;
    const match = /^(BEGIN|END):(.+)$/i.exec(structural);
    if (!match) continue;
    const name = match[2].toUpperCase();

    if (match[1].toUpperCase() === 'BEGIN') {
      const openDepth = stack.length;
      if (candidateStart === -1 && name === 'VTIMEZONE') {
        candidateStart = i;
        candidateOpenDepth = openDepth;
        malformed = openDepth !== 1;
      } else if (candidateStart !== -1) {
        const isDirectChild = openDepth === candidateOpenDepth + 1;
        if (!isDirectChild || !['STANDARD', 'DAYLIGHT'].includes(name)) {
          malformed = true;
        }
      }
      stack.push(name);
    } else {
      // A mismatched END: closes nothing.
      if (stack[stack.length - 1] === name) stack.pop();
      if (candidateStart !== -1 && stack.length === candidateOpenDepth) {
        if (malformed) {
          throw new InvalidInputError('Stored calendar resource has a malformed VTIMEZONE block.');
        }
        const tzid = (parseICalValue(lines.slice(candidateStart, i + 1).join('\n'), 'TZID') || '').trim();
        blocks.push({ tzid, start: candidateStart, end: i });
        candidateStart = -1;
      }
    }
  }
  if (candidateStart !== -1) {
    throw new InvalidInputError('Stored calendar resource has an unterminated VTIMEZONE block.');
  }
  return blocks;
}

/**
 * Remove any existing VTIMEZONE block(s) spelled exactly `tzid`, so a stale one never sits
 * beside the replacement `regenerateVTimezones` inserts. Exact, not by zone identity: a
 * same-zone block under another spelling ('/America/New_York') may still be referenced by a
 * TZID nothing regenerates, and `removeOrphanedVTimezones` drops it once it is not.
 */
function stripVTimezoneBlockFor(icalData: string, tzid: string): string {
  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);
  const toRemove = extractVTimezoneBlocks(lines).filter(b => b.tzid === tzid);
  for (let i = toRemove.length - 1; i >= 0; i--) {
    lines.splice(toRemove[i].start, toRemove[i].end - toRemove[i].start + 1);
  }
  return lines.join(lineEnding);
}

/**
 * Insert a generated VTIMEZONE `block` right before the first VEVENT, where
 * `createCalendarEvent` places one.
 */
function insertVTimezoneBlock(icalData: string, block: string, lineEnding: string): string {
  const lines = icalData.split(/\r?\n/);
  const veventIdx = lines.findIndex(l => structuralLine(l) === 'BEGIN:VEVENT');
  lines.splice(veventIdx === -1 ? lines.length : veventIdx, 0, ...block.split(/\r?\n/));
  return lines.join(lineEnding);
}

/**
 * The end instant of a DURATION from a zoned `startIso`, for the VTIMEZONE span only. RFC 5545
 * §3.3.6 makes weeks/days nominal (same wall clock N days later) and hours/minutes/seconds exact
 * elapsed time, so the day shift is resolved to an instant FIRST and the time part added as
 * milliseconds: across a spring-forward, `PT6H` from 23:00 ends at 06:00, not 05:00.
 */
function resolveDurationSpanEndMs(durationValue: string, startIso: string, tzid: string): number | undefined {
  const parsed = parseICalDurationComponents(durationValue);
  if (!parsed) return undefined;
  const { sign, weeks, days, hours, minutes, seconds } = parsed;

  const nominalDays = sign * ((weeks * 7) + days);

  const [datePart, timePart] = startIso.split('T');
  const [y, mo, d] = datePart.split('-').map(Number);
  // `Date.UTC` maps a two-digit year to 19xx, so build a Gregorian cycle away and shift back,
  // as `nextDateOnly` does.
  const shifted = new Date(Date.UTC(y + GREGORIAN_CYCLE_YEARS, mo - 1, d));
  shifted.setUTCDate(shifted.getUTCDate() + nominalDays);
  const year = shifted.getUTCFullYear() - GREGORIAN_CYCLE_YEARS;
  const pad = (n: number) => String(n).padStart(2, '0');
  const nominalIso = `${String(year).padStart(4, '0')}-${pad(shifted.getUTCMonth() + 1)}-${pad(shifted.getUTCDate())}T${timePart}`;

  const nominalMs = resolveCalendarInstantMs(nominalIso, tzid);
  if (Number.isNaN(nominalMs)) return NaN;

  const exactMs = sign * ((hours * 3600000) + (minutes * 60000) + (seconds * 1000));
  return nominalMs + exactMs;
}

/**
 * Recompute the VTIMEZONE block(s) the master VEVENT's current DTSTART/DTEND need, after
 * `updateCalendarEvent` has patched them (#166). Must run after every other patch, so it reads
 * the FINAL start/end, and before `removeOrphanedVTimezones`.
 */
export function regenerateVTimezones(icalData: string, lineEnding: string): string {
  const vevent = extractVEvent(icalData);
  if (!vevent) return icalData;

  // A plain Error: reaching here means the isRecurringSeriesResource refusal failed, a server
  // bug. A series-aware span is designed under #146.
  if (hasICalProperty(vevent, 'RRULE') || hasICalProperty(vevent, 'RDATE')) {
    throw new Error(
      'Cannot compute a VTIMEZONE span for a recurring VEVENT (RRULE/RDATE present) — a single ' +
      'occurrence\'s own DTSTART/DTEND is the wrong span for a series.'
    );
  }

  const startLine = parseAllICalProperties(vevent, 'DTSTART')[0];
  const endLine = parseAllICalProperties(vevent, 'DTEND')[0];
  const startFrame = startLine ? describeDateProperty(startLine) : undefined;
  const endFrame = endLine ? describeDateProperty(endLine) : undefined;
  const frames = [startFrame, endFrame].filter((f): f is DatePropertyFrame => f !== undefined);

  const zoneTzids = referencedZoneTzids(frames);
  if (zoneTzids.size === 0) return icalData;

  const labeled: Array<{ label: string; frame: DatePropertyFrame }> = [];
  if (startFrame) labeled.push({ label: 'DTSTART', frame: startFrame });
  if (endFrame) labeled.push({ label: 'DTEND', frame: endFrame });
  const instants = collectZoneInstants(labeled);

  // A DURATION end shares DTSTART's zone, so it extends the same span (#166).
  if (!endLine && startFrame && startFrame.frame === 'zoned' && startFrame.tzid && isUsableTimezone(startFrame.tzid)) {
    const durationLine = parseAllICalProperties(vevent, 'DURATION')[0];
    if (durationLine) {
      const colonIdx = findValueBoundary(durationLine);
      const durationValue = colonIdx === -1 ? '' : durationLine.slice(colonIdx + 1).trim();
      const startIso = formatICalDate(startFrame.value);
      const endMs = startIso ? resolveDurationSpanEndMs(durationValue, startIso, startFrame.tzid) : undefined;
      if (endMs !== undefined) {
        if (Number.isNaN(endMs)) {
          throw new InvalidInputError('Cannot resolve DURATION to an instant for VTIMEZONE generation.');
        }
        instants.push(endMs);
      }
    }
  }

  const spanMinMs = Math.min(...instants);
  const spanMaxMs = Math.max(...instants);

  let result = icalData;
  for (const tzid of zoneTzids) {
    result = stripVTimezoneBlockFor(result, tzid);
  }
  for (const tzid of zoneTzids) {
    result = insertVTimezoneBlock(result, generateVTimezone(tzid, spanMinMs, spanMaxMs, lineEnding), lineEnding);
  }
  return result;
}

// Returns undefined unless the RESULT is a plain `YYYY-MM-DD`. The input may be any stored
// text, and valid arithmetic can still leave four-digit years: 9999-12-31 renders as
// `+010000-01-01T...`, and near the maximum date `toISOString` throws. So the check runs after
// the increment, on the WHOLE ISO string before slicing, where `^` is what rejects the `+`.
function nextDay(dateStr: string): string | undefined {
  const d = new Date(dateStr);
  d.setUTCDate(d.getUTCDate() + 1);
  if (Number.isNaN(d.getTime())) return undefined;
  const iso = d.toISOString();
  return /^\d{4}-\d{2}-\d{2}T/.test(iso) ? iso.slice(0, 10) : undefined;
}

/**
 * Parse an ISO-ish date/datetime string in a fixed UTC frame, so a naive datetime is not read
 * in the process's local zone. Only for comparing two values that both pass through here, as
 * `removeExceptionVEvents` does; never for placing an event against a window, which must use
 * `resolveCalendarInstantMs` (#162).
 */
export function parseICalDateAsUTC(iso: string): Date {
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return new Date(iso + 'T00:00:00Z');
  if (/Z$|[+-]\d{2}:?\d{2}$/.test(iso)) return new Date(iso);
  return new Date(iso + 'Z');
}

const DATE_ONLY_EVENT_VALUE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Which zone one of an event's values is written in: its own TZID where ICU can resolve it,
 * the configured zone otherwise (absent, floating, or an unresolvable name such as "AUS Eastern
 * Standard Time", on which `zoneOffsetMsAt` would throw and fail the whole read).
 *
 * NORMALISED BEFORE IT IS TESTED: the stored spelling can carry RFC 5545 §3.2.19's leading
 * slash ('/America/New_York'), which `isUsableTimezone` rejects, and the exact window filter
 * would then drop the event. Shared by the window filter and `sortEventsByStart` so they agree.
 */
function zoneForValue<Z extends string | undefined>(
  name: string | null | undefined,
  configuredZone: Z,
): string | Z {
  if (!name) return configuredZone;
  const normalized = normalizeZoneForComparison(name);
  return isUsableTimezone(normalized) ? normalized : configuredZone;
}

/**
 * The date-only value one calendar day after a date-only one. `Date.UTC` maps a two-digit year
 * to 19xx, so the arithmetic runs a whole 400-year Gregorian cycle away, where leap rules repeat.
 */
function nextDateOnly(dateOnly: string): string {
  const [y, mo, d] = dateOnly.split('-').map(Number);
  const shifted = new Date(Date.UTC(y + GREGORIAN_CYCLE_YEARS, mo - 1, d + 1));
  const year = shifted.getUTCFullYear() - GREGORIAN_CYCLE_YEARS;
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${String(year).padStart(4, '0')}-${pad(shifted.getUTCMonth() + 1)}-${pad(shifted.getUTCDate())}`;
}

/**
 * Whether a parsed event falls inside the requested window: EXACT, and authoritative for what
 * the caller is shown (#162). The margin for what the server withholds widens the REQUESTED
 * range in `getCalendarEvents`; a filter cannot keep what the server never sent.
 *
 * Every value is resolved by what it is:
 *   - a `Z` or offset value is the instant it names;
 *   - a date-only value is local midnight in its zone, and a date-only `end` is ALREADY
 *     exclusive (measured, docs/fastmail-action-availability.md);
 *   - a wall clock resolves via `zoneForValue`, never as UTC;
 *   - an event with nothing readable to judge is KEPT: a missing event is what #64 prevents.
 *
 * A master the server declined to expand still carries its ORIGINAL DTSTART; the recurrence
 * guard in `getCalendarEvents` keeps it, not this function.
 *
 * `zone` is REQUIRED so no call site falls through to the host zone.
 */
export function eventIntersectsWindow(
  event: Pick<CalendarEvent, 'start' | 'end' | 'timeZone' | 'endTimeZone'>,
  windowStartMs: number,
  windowEndMs: number,
  zone: string,
): boolean {
  const anchor = event.start || event.end;
  if (!anchor) return true;

  const startZone = zoneForValue(event.timeZone, zone);
  // `endTimeZone` is emitted ONLY when end's zone differs from start's, so `undefined` here
  // means "same as start" and must inherit start's zone rather than fall back to configured.
  const endZone = event.endTimeZone === undefined ? startZone : zoneForValue(event.endTimeZone, zone);
  const anchorZone = event.start ? startZone : endZone;

  const startMs = resolveCalendarInstantMs(anchor, anchorZone);
  if (Number.isNaN(startMs)) return true;

  const parsedEnd = event.end ? resolveCalendarInstantMs(event.end, endZone) : NaN;
  let endMs: number;
  if (!Number.isNaN(parsedEnd)) {
    endMs = Math.max(parsedEnd, startMs);
  } else if (event.start && DATE_ONLY_EVENT_VALUE.test(event.start.trim())) {
    // An all-day event with no end covers its whole day, reached in LOCAL days so a DST day
    // still ends at midnight. The NaN arm is only '9999-12-31', whose next day does not resolve;
    // a flat 24 hours there keeps the event at least as wide as its day.
    const nextMidnight = resolveCalendarInstantMs(nextDateOnly(event.start.trim()), anchorZone);
    endMs = Number.isNaN(nextMidnight) ? startMs + 24 * 60 * 60 * 1000 : nextMidnight;
  } else {
    // A missing or unreadable end makes the event a point in time, not an unbounded one.
    endMs = startMs;
  }

  // A zero-width instant (an event with no duration) has no interval to overlap, so it counts
  // as inside when the window contains it. `windowEnd` is exclusive, matching CalDAV.
  if (startMs === endMs) return startMs >= windowStartMs && startMs < windowEndMs;
  return startMs < windowEndMs && endMs > windowStartMs;
}

/**
 * Order events by the INSTANT their start names, ascending, in place. Not a string sort:
 * starts arrive as bare wall clocks, `Z` instants and bare dates, which interleave wrongly as
 * text, and `limit` then cuts a genuinely earlier event. Zones come from `zoneForValue`, shared
 * with the window filter (#139, #162).
 *
 * An unreadable or absent start sorts FIRST: last is where `limit` truncates.
 */
export function sortEventsByStart(events: CalendarEvent[], zone: string | undefined): void {
  const instants = new Map<CalendarEvent, number>();
  for (const event of events) {
    instants.set(event, resolveCalendarInstantMs(event.start, zoneForValue(event.timeZone, zone)));
  }
  events.sort((a, b) => {
    const aMs = instants.get(a)!;
    const bMs = instants.get(b)!;
    if (Number.isNaN(aMs) || Number.isNaN(bMs)) {
      if (Number.isNaN(aMs) && Number.isNaN(bMs)) return 0;
      return Number.isNaN(aMs) ? -1 : 1;
    }
    return aMs - bMs;
  });
}

/**
 * How far a HALF-OPEN calendar window is allowed to run past the bound the caller gave.
 *
 * The window is also the range the SERVER EXPANDS OVER, one VEVENT per occurrence with no cap
 * in Cyrus, and `limit` bounds none of that work, so the missing half of a one-sided window
 * cannot be open-ended. A month is what a calendar client shows; Cyrus's JMAP
 * `CalendarEvent/query` and Microsoft Graph's `calendarView` both require an upper bound.
 *
 * 31 so the same date next month is always inside. Fixed 24-hour days; see `shiftIsoDays`.
 * Applies only to an INVENTED bound, and is disclosed through `windowClamp`.
 */
export const CALENDAR_OPEN_WINDOW_DAYS = 31;

/**
 * How many in-window occurrences ONE CalDAV resource may expand to before this server declines
 * to materialise it.
 *
 * The caller chooses the span but an invitation sender chooses the DENSITY: one
 * `FREQ=MINUTELY` series fills any window. 5000 passes dense-but-real cases:
 *
 *   10 years of a DAILY series            3,653   passes
 *   a month at every 10 minutes           4,464   passes
 *   a month at every 5 minutes            8,928   trips
 *   a month of FREQ=MINUTELY             44,640   trips
 *
 * It counts blocks returned for the widened request, so it can slightly exceed the in-window
 * count, accepted to avoid a per-block parse. A tripped resource fails the call with an
 * InvalidInputError naming the series.
 *
 * RESIDUAL: this bounds what this server parses and shows, never what Cyrus generates or
 * transfers. Cyrus's `expand_cb` has no limit and `CALDAV:max-instances` has no handler, and
 * tsdav's `fetchCalendarObjects` buffers the whole multistatus with no paging.
 */
export const CALENDAR_MAX_OCCURRENCES_PER_SERIES = 5000;

// The four-digit-year range. Past it `toISOString` emits `+010000-...`, which tsdav rejects
// with a plain Error, surfacing a caller-fixable argument as InternalError.
const LATEST_REPRESENTABLE_INSTANT = '9999-12-31T23:59:59Z';
const EARLIEST_REPRESENTABLE_INSTANT = '0000-01-01T00:00:00Z';

/**
 * Pull an already-resolved instant back inside the representable range. Applies to a
 * caller-named bound too: `endDate: "9999-12-31"` on a UTC-5 account resolves past year 9999.
 * Saturates rather than rejects, since `9999-12-31` is a fair question; a saturated caller
 * bound is named in the window clamp.
 */
function saturateInstant(iso: string): string {
  if (/^\d{4}-/.test(iso)) return iso;
  return iso.startsWith('-') ? EARLIEST_REPRESENTABLE_INSTANT : LATEST_REPRESENTABLE_INSTANT;
}

/**
 * Which end a value `saturateInstant` moved was pulled to. Only meaningful for a value that did
 * move.
 */
function saturationEdge(saturatedValue: string): 'earliest' | 'latest' {
  return saturatedValue === EARLIEST_REPRESENTABLE_INSTANT ? 'earliest' : 'latest';
}

/**
 * Shift an ISO-8601 UTC instant by whole days, saturating at the representable range.
 *
 * 24-hour days deliberately, NOT `coerceCalendarWindowEnd`'s local days: that advances a bound
 * the caller named, while this invents one nobody named, so a DST hour is not a wrong answer
 * and the note names the resulting instant.
 */
function shiftIsoDays(iso: string, days: number): string {
  return shiftIsoMs(iso, days * 24 * 60 * 60 * 1000);
}

/**
 * Shift an ISO-8601 UTC instant by milliseconds, saturating at the representable range: the
 * request margin below pushes `endDate: "9999-12-31"` past it.
 */
function shiftIsoMs(iso: string, ms: number): string {
  const shifted = new Date(Date.parse(iso) + ms)
    .toISOString()
    .replace(/\.\d{3}Z$/, 'Z');
  return saturateInstant(shifted);
}

// The widest UTC offset any IANA zone has used (+14:00), and so how far the range SENT TO THE
// SERVER runs past the caller's window at each edge (#162). Cyrus matches an all-day value on
// its UTC day and resolves a floating time as UTC, so a narrow window loses both server-side
// (scripts/probes/calendar-window-frames.probe.mjs). `eventIntersectsWindow` then judges
// exactly.
const MAX_UTC_OFFSET_MS = 14 * 60 * 60 * 1000;

// Calendar display names are server/user data of unbounded length, so a listing of them is
// capped the way every other echoed list in this server is.
export const CALENDAR_NAME_LIST_CAP = 20;

// A calendar URL in the not-found error is offered as a `calendarId` to paste back, so it must
// survive intact: Fastmail's prefix alone is 47 characters before the address. Names keep the
// 60-char default, since they are read, not pasted.
export const CALENDAR_URL_ECHO_LIMIT = 200;

/**
 * The bound on ONE broken collection's path wherever it is echoed (#136). Not a share of
 * CALENDAR_URL_ECHO_LIMIT: that is a handle that must survive, this is a path the caller can
 * only recognise, not act on.
 */
export const BROKEN_COLLECTION_PATH_ECHO_LIMIT = 160;

/**
 * How many broken collections one message names before the rest are counted. A home listing
 * that failed wholesale would otherwise put every path it holds into the note.
 */
export const BROKEN_COLLECTION_LIST_CAP = 5;

/**
 * How many copies the ambiguity messages name before the rest are counted (#101). Higher than
 * `BROKEN_COLLECTION_LIST_CAP`: a copy that is not named cannot be chosen.
 */
export const AMBIGUOUS_COPY_LIST_CAP = 12;

/**
 * The bound on ONE copy's calendar and resource URL in the ambiguity report (#101). A resource
 * url is a collection url plus a UID-derived name (120 characters allowed here), and it is the
 * handle the caller must pass back, unrecoverable elsewhere if truncated.
 */
export const AMBIGUOUS_COPY_URL_ECHO_LIMIT = 320;

/**
 * The bound on a stored UID echoed as an id to pass back. Nothing caps a UID's length, and
 * Exchange-generated ones run past 100 hex characters, so 320 keeps every real UID whole while
 * a hostile one still cannot fill the reply.
 */
export const CALENDAR_UID_ECHO_LIMIT = 320;

/**
 * The bound on tsdav's own text inside the login refusal (#182).
 *
 * Wider than `describeUntrusted`'s 64 default because the status code sits at the END of
 * `Invalid credentials: PROPFIND <url> returned 401 Unauthorized`.
 */
export const LOGIN_FAILURE_ECHO_LIMIT = 200;

/**
 * Is this entry of the calendar-home listing BROKEN — i.e. did the server fail to describe it?
 *
 * Covers both forms in tsdav's output: a failed RESPONSE (non-2xx status, no props) and a
 * failed PROPSTAT (2xx element, but tsdav keeps props from 2xx propstats only, so
 * `resourcetype` is gone).
 *
 * A MISSING resourcetype is the test, NOT an empty one: `<resourcetype/>` parses to `{}`, an
 * ordinary resource. A non-number `status` is unconfirmed, as in assertDavOk.
 */
export function isBrokenCalendarHomeEntry(entry: DAVResponse): boolean {
  const status = entry?.status;
  if (typeof status !== 'number' || status < 200 || status >= 300) return true;
  return (entry?.props as Record<string, unknown> | undefined)?.resourcetype === undefined;
}

/** An href resolved against the calendar home, or undefined when it cannot be resolved at all. */
function resolveCollectionUrl(url: string, base: string): URL | undefined {
  try {
    return new URL(url, base);
  } catch {
    return undefined;
  }
}

/**
 * The comparable form of a collection path: its SEGMENTS, each decoded on its own. Decoded
 * because the two sides may spell the address's `@` differently (`%40`); per segment because a
 * decoded `%2F` would otherwise become a separator, making `/dav/.../probe%2Fwobbly/` read as a
 * child of `/dav/.../probe/`.
 */
function collectionPathSegments(pathname: string): string[] {
  return pathname
    .split('/')
    .filter(segment => segment.length > 0)
    .map(segment => {
      try {
        return decodeURIComponent(segment);
      } catch {
        // A stray '%' is not an escape; compare the raw text rather than losing the entry.
        return segment;
      }
    });
}

/**
 * Is `entry` a path strictly BELOW `container`? Segment arrays, never a string prefix, which
 * would put `/dav/.../personal-archive/x.ics` under `/dav/.../personal` (#136, #137).
 */
function isPathStrictlyInside(container: string[], entry: string[]): boolean {
  if (entry.length <= container.length) return false;
  return container.every((segment, i) => entry[i] === segment);
}

/**
 * The base a relative CalDAV url is resolved against for the confinement test below. NEVER a
 * request target: it puts both sides on one origin, and makes a bare `meeting.ics` resolve
 * outside every calendar. `.invalid` is RFC 2606 reserved.
 */
const CALDAV_URL_MATCH_BASE = 'https://caldav.invalid/';

/**
 * The calendars a caller-supplied URL-shaped event id sits inside, and the resource url to ask
 * each of them for (#137).
 *
 * THE CALLER'S STRING IS NEVER ITSELF A REQUEST TARGET. Every request carries the app password,
 * so a url-shaped id is MATCHED against discovered collections (origin, then path segments) and
 * only a match produces a fetch against that known collection; anything else is not-found.
 *
 * A list because nothing forbids two discovered collections from nesting.
 */
function resolveEventUrlTargets(
  eventId: string,
  calendars: DAVCalendar[],
): Array<{ calendar: DAVCalendar; objectUrl: string }> {
  const candidate = resolveCollectionUrl(eventId, CALDAV_URL_MATCH_BASE);
  if (candidate === undefined) return [];
  const candidateSegments = collectionPathSegments(candidate.pathname);
  const targets: Array<{ calendar: DAVCalendar; objectUrl: string }> = [];
  for (const calendar of calendars) {
    const collectionUrl = typeof calendar.url === 'string' ? calendar.url.trim() : '';
    if (collectionUrl.length === 0) continue;
    const collection = resolveCollectionUrl(collectionUrl, CALDAV_URL_MATCH_BASE);
    if (collection === undefined) continue;
    if (collection.origin !== candidate.origin) continue;
    if (!isPathStrictlyInside(collectionPathSegments(collection.pathname), candidateSegments)) continue;
    targets.push({ calendar, objectUrl: candidate.href });
  }
  return targets;
}

/**
 * An href as compared to decide whether a resource was ADDRESSED. The fragment and query are
 * dropped from both sides: tsdav never sends a fragment, and a server may answer a query with
 * the resource's own href, so either would stop the caller's spelling matching what came back.
 */
function addressComparisonKey(url: string): string {
  const parsed = resolveCollectionUrl(url, CALDAV_URL_MATCH_BASE);
  if (parsed === undefined) return url;
  parsed.hash = '';
  parsed.search = '';
  return parsed.href;
}

/**
 * Taken from tsdav's own signature so an upgrade that changes the shape surfaces here.
 */
type CalendarQueryFilters = NonNullable<Parameters<DAVClient['fetchCalendarObjects']>[0]['filters']>;

/**
 * Which hrefs a calendar-object fetch will request. Passed by EVERY `fetchCalendarObjects`
 * call in this file, so that no read reaches a record another read reports as absent (#191).
 *
 * Replaces tsdav's default `url.includes('.ics')`, which judges KIND by NAME and made `.ICS` or
 * extensionless resources unreachable. The VEVENT comp-filter keeps other resources out of a
 * calendar-query. The url-form multiget sends none, so there `isResolvedCalendarObject`'s
 * VEVENT test is the only guard. The collection's own url must be excluded HERE: tsdav's
 * calendar branch does not.
 */
function calendarResourceUrlFilter(collectionUrl: string | undefined): (url: string) => boolean {
  return (url: string) => Boolean(url) && !urlEquals(url, collectionUrl);
}

/**
 * The CalDAV `calendar-query` filter that asks one collection for the resources whose VEVENT
 * carries exactly this UID (#137).
 *
 * MEASURED against the live account (scripts/probes/calendar-uid-query.probe.mjs): Cyrus
 * honours the CardDAV `match-type` attribute (RFC 6352 §10.5.1) on CalDAV, and with no
 * `collation` RFC 4791's default `i;ascii-casemap` makes `equals` case-INSENSITIVE. So the
 * caller's exact-equality check on the PARSED UID is LOAD-BEARING: a loose match would invent
 * an ambiguity and freeze the writes (#101).
 *
 * xml-js escapes only `<`, `>` and `&` in `_text`, so a UID cannot inject markup; characters
 * XML cannot carry are refused before this, see XML_UNSENDABLE_CHARS.
 */
function uidEqualsFilter(uid: string): CalendarQueryFilters {
  return [{
    'comp-filter': {
      _attributes: { name: 'VCALENDAR' },
      'comp-filter': {
        _attributes: { name: 'VEVENT' },
        'prop-filter': {
          _attributes: { name: 'UID' },
          'text-match': { _attributes: { 'match-type': 'equals' }, _text: uid },
        },
      },
    },
  }] as unknown as CalendarQueryFilters;
}

/**
 * One copy of an event that a lookup RESOLVED, and the label of the calendar it was found in.
 */
interface CalendarObjectMatch {
  object: DAVCalendarObject;
  calendarLabel: string;
}

/**
 * What a lookup hands back: every copy the id resolved to, whether the caller ADDRESSED one of
 * them, and what discovery could not search.
 *
 * `matches` IS ORDERED, and the order is contract: the ADDRESSED copy first, then UID matches,
 * then anything else the url form resolved. `get_calendar_event` returns `matches[0]`.
 *
 * `addressed` means the caller's string was the resource url of `matches[0]`, not merely a UID.
 * It DECIDES AMBIGUITY: an invitation sender can mint a decoy whose UID is another event's url,
 * but addressing cannot be imitated, so an addressed record is the one the reads answer with.
 * At most one copy can be addressed.
 *
 * `collision` is set when another match's UID is the string: the listing shows that record's id
 * as this url, so the writes refuse rather than act on the addressed one. `addressedUid` is the
 * addressed record's UID only where that UID reaches it alone; otherwise no id does.
 */
interface CalendarObjectLookup {
  matches: CalendarObjectMatch[];
  addressed: boolean;
  collision?: { addressedUid: string | undefined };
  brokenCollections: BrokenCollections;
}

/**
 * Every character XML 1.0's `Char` excludes that SURVIVES TO THE WIRE: C0 except tab, LF and
 * CR, plus the noncharacters U+FFFE and U+FFFF. xml-js copies them raw, the server answers 400,
 * and a plain Error would read as InternalError for a value only the caller can fix, so the
 * argument is refused instead.
 *
 * DELIBERATELY NOT HERE, because each is sendable: tab, CR and LF (legal `Char`: trimmed off the
 * ends of either form, inside a UID they come back not-found, and the url form's URL parser
 * strips them); U+FDD0-U+FDEF (admitted by XML 1.0); lone surrogates (MEASURED:
 * `TextEncoder` repairs them to U+FFFD).
 */
const XML_UNSENDABLE_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/u;

/**
 * Did this throw from an addressed multiget mean "there is no resource at that href"?
 *
 * tsdav collapses every failure into `Collection query failed: <status> <statusText>. ...`, so
 * the status is the ONLY signal: 404 and 410 name an absent resource, anything else rethrows.
 * A non-Error throw is not stringified into a match. See the call site for why a whole
 * collection answering 404 is still safe.
 */
function isAddressedResourceMissing(err: unknown): boolean {
  const message = err instanceof Error ? err.message : '';
  return /^Collection query failed: 4(?:04|10)\b/.test(message);
}

/** The calendar name a message shows, with the url as the handle that always exists. */
function calendarLabel(calendar: DAVCalendar): string {
  return unwrapDisplayName(calendar.displayName)
    ?? (typeof calendar.url === 'string' ? calendar.url.trim() : '');
}

/** One undecided resource: the href to ASK about, and the rows that resolving it would label. */
interface UndecidedResource {
  requestHref: string;
  rows: CalendarEvent[];
}

/**
 * Settle `isRecurring` for the listing rows an EXPANDED read cannot decide (#155).
 *
 * A series' first expanded instance with no sibling in the window is indistinguishable from a
 * one-off (see `parseCalendarObjects`); the STORED resource keeps its rule. One request per
 * calendar, covering nearly every row, since one-offs look the same.
 *
 * AN INCOMPLETE ANSWER FAILS THE WHOLE LISTING: an absent `isRecurring` claims the event does
 * not repeat. See docs/conventions.md.
 */
async function settleAmbiguousRecurrence(
  client: DAVClient,
  calendar: DAVCalendar,
  undecided: Map<string, UndecidedResource>,
): Promise<void> {
  if (undecided.size === 0) return;

  const responses = await client.calendarMultiGet({
    url: calendar.url,
    props: { 'd:getetag': {}, 'c:calendar-data': {} },
    objectUrls: [...undecided.values()].map(entry => entry.requestHref),
    depth: '1',
  });

  const answered = new Set<string>();
  for (const res of responses) {
    const href = res.href;
    if (!href) continue;
    const url = resolveResponseHref(href, calendar.url);
    const entry = undecided.get(url);
    if (!entry) continue;
    const ical = readCalendarData(res);
    if (ical === undefined) continue;
    // No readable VEVENT (e.g. lower-cased keywords) has not answered: `isRecurringSeriesResource`
    // would return false, a positive "does not repeat".
    if (extractVEventBlocks(ical).length === 0) continue;
    answered.add(url);
    if (isRecurringSeriesResource(ical)) for (const row of entry.rows) row.isRecurring = true;
  }

  if (answered.size < undecided.size) {
    // The label is server-authored, not caller-authored, and still goes through the echo.
    throw new Error(
      `Calendar "${echoCallerText(calendarLabel(calendar))}": the follow-up read that settles whether an ` +
      `event repeats did not answer for ${undecided.size - answered.size} of the ${undecided.size} ` +
      'resources it asked about, so this listing cannot say which of its rows repeat. It fails rather ' +
      'than report those rows as one-off events on no evidence. Retry the call; a failure that persists ' +
      'is the calendar server not answering a multiget it accepted.',
    );
  }
}

/**
 * A resource url in one comparable form, resolved against the collection since a DAV server may
 * write a bare path. An unresolvable href is returned as it came: its counterpart was built from
 * the same base.
 */
function resolveResponseHref(href: string, collectionUrl: string | undefined): string {
  try {
    return new URL(href, collectionUrl).href;
  } catch {
    return href;
  }
}

/**
 * A resource url as a multiget `<D:href>`: path and query, no origin, as tsdav writes its own.
 */
function toRequestHref(url: string, collectionUrl: string | undefined): string {
  try {
    const resolved = new URL(url, collectionUrl);
    return resolved.pathname + resolved.search;
  } catch {
    return url;
  }
}

/** The iCalendar payload of a `calendar-data` prop, which tsdav hands back CDATA-wrapped or bare. */
function readCalendarData(res: DAVResponse): string | undefined {
  const data = (res.props as { calendarData?: unknown } | undefined)?.calendarData;
  if (typeof data === 'string') return data;
  const cdata = (data as { _cdata?: unknown } | undefined)?._cdata;
  return typeof cdata === 'string' ? cdata : undefined;
}

/**
 * The resolved copies as a caller sees them (#101). `object.url` is safe here because
 * `isResolvedCalendarObject` admitted only matches that carry one.
 */
/** A resource's own UID, trimmed as the lookup compares it; undefined when it has none. */
function ownUid(obj: DAVCalendarObject): string | undefined {
  return parseICalValue(extractVEvent(obj.data || '') ?? '', 'UID')?.trim();
}

function matchesToCopies(matches: CalendarObjectMatch[]): CalendarEventCopy[] {
  return matches.map(m => ({ calendar: m.calendarLabel, url: m.object.url }));
}

/**
 * Did this resource come back COMPLETE enough to be acted on?
 *
 * One rule for every tool and both resolution paths: a url (the write's address), a parseable
 * VEVENT (what the recurrence refusal reads; without it a series passes as "not recurring"),
 * and an etag (tsdav's `cleanupFalsy` drops an empty `If-Match`, so the delete would be
 * unconditional). Measured present on every live match
 * (`scripts/probes/calendar-uid-query.probe.mjs`, step 2). The VEVENT requirement is also what
 * keeps non-event resources out of the url-form path (see `calendarResourceUrlFilter`), so it
 * must not be relaxed for a read-only caller.
 */
function isResolvedCalendarObject(obj: DAVCalendarObject): boolean {
  if (typeof obj.url !== 'string' || obj.url.trim().length === 0) return false;
  if (typeof obj.etag !== 'string' || obj.etag.trim().length === 0) return false;
  if (typeof obj.data !== 'string' || obj.data.length === 0) return false;
  return extractVEventBlocks(obj.data).length > 0;
}

/**
 * The collections inside `homeUrl`'s own PROPFIND answer that failed to list (#136).
 *
 * An entry qualifies when its href is strictly UNDER the home (origin and path) and
 * `isBrokenCalendarHomeEntry` says so. The positional test keeps a REQUEST-level failure, whose
 * pseudo-entry carries the request URL or no href, out of this list: that is a whole-discovery
 * failure (#100), reported by `discoverCalendars`. The origin test stops a foreign absolute href
 * being echoed as the caller's.
 */
export function findBrokenCalendarHomeCollections(entries: DAVResponse[], homeUrl: string): BrokenCollections {
  const home = resolveCollectionUrl(homeUrl, homeUrl);
  // Silence is the safe direction: a missed disclosure, never a healthy collection reported.
  if (home === undefined) return [];
  const homeSegments = collectionPathSegments(home.pathname);
  const broken: string[] = [];
  const seen = new Set<string>();
  for (const entry of entries) {
    const href = typeof entry?.href === 'string' ? entry.href.trim() : '';
    if (href.length === 0) continue;
    const resolved = resolveCollectionUrl(href, homeUrl);
    if (resolved === undefined) continue;
    if (resolved.origin !== home.origin) continue;
    if (!isPathStrictlyInside(homeSegments, collectionPathSegments(resolved.pathname))) continue;
    if (!isBrokenCalendarHomeEntry(entry)) continue;
    if (seen.has(resolved.href)) continue;
    seen.add(resolved.href);
    broken.push(resolved.href);
  }
  return broken;
}

/**
 * The result field, ABSENT when nothing broke, as the window clamp is.
 */
function asBrokenCollectionsField(broken: BrokenCollections): BrokenCollections | undefined {
  return broken.length > 0 ? broken : undefined;
}

/** What discovery hands back internally: the library's list, plus what failed beside it. */
interface DiscoveredCalendars {
  calendars: DAVCalendar[];
  brokenCollections: BrokenCollections;
}

/**
 * The bytes of one DAV answer, kept so the detection can re-read exactly what tsdav parsed.
 * No `statusText`: nothing reads it, and the `Response` constructor throws on a malformed one.
 */
interface CapturedDavAnswer {
  body: string;
  status: number;
  contentType: string | null;
}

/**
 * Re-parse a captured calendar-home answer and report its broken collections (#136).
 *
 * TSDAV'S OWN parse, via a `fetch` override that replays the captured bytes, so the detection
 * sees exactly what `fetchCalendars` saw. A parse that fails reports nothing broken rather than
 * throwing: most likely it was never a multistatus, which `discoverCalendars` already reports.
 */
async function brokenCollectionsInAnswer(captured: CapturedDavAnswer, homeUrl: string): Promise<BrokenCollections> {
  let entries: DAVResponse[];
  try {
    entries = await davRequest({
      url: homeUrl,
      // Nothing is sent; `convertIncoming: false` stops the builder serialising an absent body.
      init: { method: 'PROPFIND', body: undefined },
      convertIncoming: false,
      fetch: async () => new Response(captured.body, {
        status: captured.status,
        headers: captured.contentType === null ? {} : { 'content-type': captured.contentType },
      }),
    });
  } catch {
    return [];
  }
  return findBrokenCalendarHomeCollections(entries, homeUrl);
}

/**
 * The part of the broken-collection disclosure TRUE OF EVERY COUNT (#136), exported so the tool
 * descriptions quote a string that always prints; the subject before it varies with the count.
 */
export const BROKEN_COLLECTION_PHRASE = 'in the calendar list failed to list';

/** The count-dependent pieces of the disclosure, built once for both places that write it. */
export interface BrokenCollectionSummary {
  /** "a collection in the calendar list failed to list" / "3 collections in the calendar …". */
  subject: string;
  /** The capped, echoed path list: `"a", "b", …and 2 more`. */
  paths: string;
  /** The sentence bounding what the disclosure may claim, agreeing in number with `subject`. */
  disclaimer: string;
  /** Whether more than one collection is named — every pronoun beside this follows it. */
  plural: boolean;
}

/**
 * The shared half of the broken-collection disclosure: the subject, the capped path list, and
 * the sentence bounding what may be claimed (#136).
 *
 * ONE BUILDER for both surfaces, the returned note (response-formatters.ts) and a THROWN clause,
 * so the count-dependent wording cannot drift. Neither may upgrade "collection" to "calendar"
 * (see `BrokenCollections`). Each path is echoed as a value, never the finished sentence
 * (docs/conventions.md).
 */
export function summariseBrokenCollections(broken: BrokenCollections): BrokenCollectionSummary {
  const shown = broken
    .slice(0, BROKEN_COLLECTION_LIST_CAP)
    .map(p => `"${echoCallerText(p, BROKEN_COLLECTION_PATH_ECHO_LIMIT)}"`)
    .join(', ');
  const more = broken.length > BROKEN_COLLECTION_LIST_CAP
    ? `, …and ${broken.length - BROKEN_COLLECTION_LIST_CAP} more`
    : '';
  const plural = broken.length !== 1;
  return {
    subject: plural
      ? `${broken.length} collections ${BROKEN_COLLECTION_PHRASE}`
      : `a collection ${BROKEN_COLLECTION_PHRASE}`,
    paths: `${shown}${more}`,
    disclaimer: plural
      ? 'The failure destroyed their names and their types, so it cannot be said whether they were calendars.'
      : 'The failure destroyed its name and its type, so it cannot be said whether it was a calendar.',
    plural,
  };
}

/**
 * The clause naming broken collections inside a THROWN message (#136), worded as an aside: the
 * caller's own problem in front of it may well be the real one.
 */
function describeBrokenCollections(broken: BrokenCollections | undefined): string {
  if (!broken || broken.length === 0) return '';
  const s = summariseBrokenCollections(broken);
  return ` Separately, ${s.subject}, so nothing in ${s.plural ? 'them' : 'it'} could be read: `
    + `${s.paths}. ${s.disclaimer}`;
}

/**
 * The display name of Fastmail's hidden task collection, which `list_calendars` and the
 * not-found error both hide.
 */
const HIDDEN_TASK_CALENDAR_NAME = 'DEFAULT_TASK_CALENDAR_NAME';

/**
 * The one place a DAV `displayName` is turned into a name: tsdav types it `string`, but hands
 * through whatever xml-js produced (`fetchCalendars`, tsdav 2.3.1). MEASURED shapes:
 *
 *   <displayname>Work</displayname>                 "Work"          a plain string
 *   <displayname>2026</displayname>                 2026            a NUMBER
 *   <displayname>true</displayname>                 true            a BOOLEAN
 *   <displayname><![CDATA[2026]]></displayname>     "2026"          a STRING, not a number
 *   <displayname/>  or  <displayname></displayname> {}              an EMPTY OBJECT
 *   <D:displayname xml:lang="en"/>                  {_attributes:…} an OBJECT
 *   <displayname>A</displayname> twice              ['A','B']       an ARRAY
 *   (property absent)                               undefined
 *
 * `{_cdata}` and `{_text}` are DEFENSIVE only: tsdav flattens both today, so do not build a
 * fixture believing them reachable.
 *
 * tsdav's `nativeType` coerces numeric or boolean TEXT (not CDATA), lossily: `1e3` arrives as
 * 1000. The invariant kept is that the returned name resolves when passed back as `calendarId`,
 * since both sides come through here. `String({})` is a truthy "[object Object]", which is why
 * this exists. An ARRAY (duplicate elements) is deliberately undefined, not "A,B". No fallback
 * lives here: each call site degrades differently.
 */
export function unwrapDisplayName(raw: unknown): string | undefined {
  const scalar =
    typeof raw === 'object' && raw !== null
      ? (raw as { _cdata?: unknown; _text?: unknown })._cdata ??
        (raw as { _text?: unknown })._text
      : raw;
  if (typeof scalar === 'string') {
    const trimmed = scalar.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  }
  // A parser-typed name is still the user's name for the calendar.
  if (typeof scalar === 'number' && Number.isFinite(scalar)) return String(scalar);
  if (typeof scalar === 'boolean') return String(scalar);
  return undefined;
}

/** The calendars a caller is able to name, which is the only set worth listing back at them. */
function selectableCalendars(calendars: DAVCalendar[]): DAVCalendar[] {
  return calendars.filter(c => unwrapDisplayName(c.displayName) !== HIDDEN_TASK_CALENDAR_NAME);
}

/**
 * The one "no calendar matched that id" error, raised by BOTH the read and the write path.
 *
 * Names are matched CASE-SENSITIVELY, so the available names are listed for the caller to fix
 * the spelling. The listing is filtered HERE so neither path can advertise the hidden task
 * collection.
 *
 * `broken` (#136): a collection that failed to list has no name and no calendar entry, so any
 * id aimed at it lands here; naming the broken paths stops the miss reading as proof the
 * calendar does not exist.
 */
function calendarNotFoundError(
  calendarId: unknown,
  available: DAVCalendar[],
  broken?: BrokenCollections,
): InvalidInputError {
  // A nameless calendar is listed by its URL, a `calendarId` that works, at the URL bound so it
  // survives being pasted back.
  const entries = selectableCalendars(available)
    .map(c => {
      const name = unwrapDisplayName(c.displayName);
      if (name !== undefined) return { text: name, limit: undefined };
      const url = typeof c.url === 'string' ? c.url.trim() : '';
      return url.length > 0 ? { text: url, limit: CALENDAR_URL_ECHO_LIMIT } : undefined;
    })
    .filter((e): e is { text: string; limit: number | undefined } => e !== undefined);
  const shown = entries
    .slice(0, CALENDAR_NAME_LIST_CAP)
    .map(e => `"${echoCallerText(e.text, e.limit)}"`)
    .join(', ');
  const more = entries.length > CALENDAR_NAME_LIST_CAP ? `, …and ${entries.length - CALENDAR_NAME_LIST_CAP} more` : '';
  const listing = entries.length > 0 ? ` Available calendars: ${shown}${more}.` : '';
  // Quoted so an empty or blank value is visible; at the URL bound because the rejected value is
  // often a mistyped URL, which 60 characters would not distinguish.
  return new InvalidInputError(
    `Calendar not found: "${echoCallerText(calendarId, CALENDAR_URL_ECHO_LIMIT)}". calendarId takes either a calendar's URL ` +
    '(its `id` from list_calendars) or its display name, and the name is matched CASE-SENSITIVELY; ' +
    'surrounding whitespace is ignored on both sides, and list_calendars reports the trimmed name.' +
    `${listing}${describeBrokenCollections(broken)}`,
  );
}

/**
 * The refusal BOTH calendar paths raise when a `calendarId` NAMES more than one calendar
 * (#173), written once so a read and a write state a single rule.
 *
 * A display name is not unique per account (a shared calendar is named by its owner), and
 * neither a union nor a first pick says which calendar was meant.
 *
 * THE WAY OUT IS THE URL, which cannot be made ambiguous: the resolver tries the url form FIRST,
 * so a decoy calendar whose NAME spells another's url cannot make the remedy circular. Same
 * rule as `ambiguousEventIdError` (#101); see docs/conventions.md.
 */
function ambiguousCalendarNameError(
  calendarId: unknown,
  matches: DAVCalendar[],
  broken?: BrokenCollections,
): InvalidInputError {
  // The same caps and bounds as `calendarNotFoundError`. Deliberately NOT
  // `AMBIGUOUS_COPY_URL_ECHO_LIMIT`: these are COLLECTION urls, which `list_calendars` always
  // hands back whole.
  const shown = matches
    .slice(0, CALENDAR_NAME_LIST_CAP)
    .map(c => `"${echoCallerText(unwrapDisplayName(c.displayName), undefined)}" ("${echoCallerText(c.url, CALENDAR_URL_ECHO_LIMIT)}")`)
    .join(', ');
  const more = matches.length > CALENDAR_NAME_LIST_CAP
    ? `, …and ${matches.length - CALENDAR_NAME_LIST_CAP} more`
    : '';
  // Double-quoted because the echo's neutralisation protects only double-quoted spans.
  //
  // THE BROKEN-COLLECTION CLAUSE IS NOT OPTIONAL (#136): an unsearched collection may hold one
  // more calendar of that name, so the account-wide count would read as complete.
  return new InvalidInputError(
    `The calendarId "${echoCallerText(calendarId, CALENDAR_URL_ECHO_LIMIT)}" names ${matches.length} calendars ` +
    'in this account, and this server will not guess which one you mean. ' +
    `The calendars are: ${shown}${more}. ` +
    "Pass that calendar's URL (its `id` from list_calendars) as calendarId instead — a calendar URL " +
    'ADDRESSES exactly one collection, whatever else carries that display name, and this parameter ' +
    `accepts a URL wherever it accepts a name.${describeBrokenCollections(broken)}`,
  );
}

/**
 * Resolve a `calendarId` to EXACTLY ONE calendar, or refuse: the one rule for the read and
 * write paths (#173). The read path's read-everything branch is deliberately outside it.
 *
 * FAIL-CLOSED ON AN EMPTY VALUE. The name arm cannot match a blank, but the URL comparison is
 * raw, so it is GUARDED: a calendar with an empty url would otherwise be addressed by an empty
 * `calendarId`. The read path's presence test must also stop an empty value reaching here.
 */
function resolveCalendarTarget(
  calendarId: unknown,
  selectable: DAVCalendar[],
  broken?: BrokenCollections,
): DAVCalendar {
  const requested = typeof calendarId === 'string' ? calendarId.trim() : calendarId;
  // BOTH SIDES through the same normaliser: tsdav delivers a calendar called "2026" as a number.
  // Accepted knowingly: a caller's `{_cdata: 'Work'}` now resolves too (lenient coercion,
  // docs/conventions.md).
  const requestedName = unwrapDisplayName(requested);

  // ADDRESSED BEATS NAMED, in a SEPARATE PASS: one url-or-name predicate let a decoy calendar
  // whose NAME spells another's url, listed first, receive the write. Guarded against an empty
  // url, per the fail-closed note above.
  const addressed = selectable.find(
    c => typeof c.url === 'string' && c.url.length > 0 && c.url === requested,
  );
  if (addressed) return addressed;

  const named = requestedName === undefined
    ? []
    : selectable.filter(c => unwrapDisplayName(c.displayName) === requestedName);
  // A miss must throw: an empty read would answer "you are free" because of a typo.
  if (named.length === 0) throw calendarNotFoundError(calendarId, selectable, broken);
  if (named.length > 1) throw ambiguousCalendarNameError(calendarId, named, broken);
  return named[0]!;
}

/**
 * The one "no event matched that id" error, raised by get/update/delete alike.
 *
 * Carries the broken-collection clause (#136): an unsearched collection may hold the record.
 *
 * `eventId` is echoed so it cannot forge a " Separately, a collection ..." clause of its own:
 * the echo, not the quotes, keeps it inside the one quote pair. At the URL bound because the
 * id is often a resource url.
 */
function eventNotFoundError(eventId: string, broken?: BrokenCollections): InvalidInputError {
  return new InvalidInputError(
    `Calendar event not found: "${echoCallerText(eventId, CALENDAR_URL_ECHO_LIMIT)}"${describeBrokenCollections(broken)}`,
  );
}

/**
 * One copy of an event as a CALLER sees it. The url is the point: it is the only thing that
 * picks out one copy of a duplicated id.
 */
export interface CalendarEventCopy {
  /** The calendar's display name, falling back to its url where the collection has none. */
  calendar: string;
  /** This copy's own resource url — the handle that names it and no other. */
  url: string;
}

/**
 * The copies an id resolved to, for the write refusal and the read disclosure alike. Every
 * value is untrusted, so each is echoed inside DOUBLE quotes (#190).
 */
export function describeEventCopies(copies: CalendarEventCopy[]): string {
  const listed = copies
    .slice(0, AMBIGUOUS_COPY_LIST_CAP)
    // BOTH fields at the url bound: `calendar` may be a nameless collection's url, and at 60
    // characters every such calendar would render identically.
    .map(c => `"${echoCallerText(c.calendar, AMBIGUOUS_COPY_URL_ECHO_LIMIT)}" ("${echoCallerText(c.url, AMBIGUOUS_COPY_URL_ECHO_LIMIT)}")`)
    .join(', ');
  const more = copies.length > AMBIGUOUS_COPY_LIST_CAP
    ? `, …and ${copies.length - AMBIGUOUS_COPY_LIST_CAP} more`
    : '';
  return `${listed}${more}`;
}

/**
 * The refusal `update_calendar_event` and `delete_calendar_event` raise when an id names more
 * than one record (#101), written once so the two read as a single rule.
 *
 * A UID is unique per COLLECTION, not per account; acting on the first copy found would patch or
 * destroy a record the caller did not choose.
 *
 * RAISED BEFORE THE REPEATING-SERIES REFUSAL, so the caller learns a second record exists.
 *
 * ACCEPTED: an invitation sender can mint a duplicate UID to freeze writes, which is why the
 * url, which ADDRESSES one record, is the offered way out; where another record's UID spells
 * that url, `addressCollisionError` names the addressed record's UID instead. Word it
 * "addresses", not "names": the copies listed beside it may all name the url-shaped id.
 */
export function ambiguousEventIdError(
  eventId: string,
  action: 'update' | 'delete',
  copies: CalendarEventCopy[],
  broken?: BrokenCollections,
): InvalidInputError {
  return new InvalidInputError(
    // Bounded as eventNotFoundError bounds the same value.
    `The event id "${echoCallerText(eventId, CALENDAR_URL_ECHO_LIMIT)}" names ${copies.length} records `
    + `in this account, and this server will not ${action} one of them without being told which. `
    + `The copies are: ${describeEventCopies(copies)}. `
    + 'Pass the `url` of the copy you mean as eventId instead — a resource url ADDRESSES exactly '
    + 'one record, and this tool accepts it wherever it accepts an id. '
    + 'get_calendar_event still works on this id: it returns the first copy and lists the others.'
    // The count is account-wide; an unsearched collection may hold another copy (#136).
    + describeBrokenCollections(broken),
  );
}

/**
 * The refusal update and delete raise on `CalendarObjectLookup.collision`. The url cannot be the
 * way out here, so each record is named with the handle that reaches it alone, if any.
 */
export function addressCollisionError(
  eventId: string,
  action: 'update' | 'delete',
  copies: CalendarEventCopy[],
  addressedUid: string | undefined,
  broken?: BrokenCollections,
): InvalidInputError {
  const [addressed, ...others] = copies;
  const reach = addressedUid
    ? `pass its own UID "${echoCallerText(addressedUid, CALENDAR_UID_ECHO_LIMIT)}" as eventId to act on it`
    : 'no event id reaches that record alone through this server; change it in the Fastmail web interface';
  return new InvalidInputError(
    `The event id "${echoCallerText(eventId, CALENDAR_URL_ECHO_LIMIT)}" is the url of one record and the UID of `
    + `another, and this server will not ${action} either without being told which. `
    + `The record at that url is ${describeEventCopies([addressed])}; ${reach}. `
    + `Other records it names: ${describeEventCopies(others)}; pass the url of the one you mean as eventId.`
    + describeBrokenCollections(broken),
  );
}

/**
 * Reorder VEVENT blocks so the master (no RECURRENCE-ID) comes first. RFC 5545/4791 do not fix
 * component order, and every in-place patch helper targets the first VEVENT.
 */
export function normalizeMasterVEventFirst(icalData: string): string {
  const vevents = extractAllVEvents(icalData);
  if (vevents.length < 2) return icalData;
  const first = vevents[0];
  if (!first || !hasRecurrenceId(first)) return icalData;
  const master = vevents.find(v => !hasRecurrenceId(v));
  if (!master) return icalData;
  // Swap the two blocks. Function replacements avoid `$`-pattern expansion.
  const SENTINEL = '\u0000MASTER-VEVENT\u0000';
  let out = icalData.replace(master, () => SENTINEL);
  out = out.replace(first, () => master);
  out = out.replace(SENTINEL, () => first);
  return out;
}

/**
 * Assert a tsdav write actually succeeded: tsdav returns raw Responses without throwing on
 * 4xx/5xx. Fails on any status outside 2xx, on a missing numeric status (even `{ ok: true }`),
 * and on an empty array: an outcome nothing confirmed is not a success.
 */
function assertDavOk(resp: unknown, action: string): void {
  const responses = Array.isArray(resp) ? resp : [resp];
  if (responses.length === 0) {
    throw new Error(`Failed to ${action}: server returned no status`);
  }
  for (const r of responses) {
    const status = (r as any)?.status;
    const ok = (r as any)?.ok;
    if (typeof status !== 'number') {
      throw new Error(`Failed to ${action}: server returned no status`);
    }
    if (status < 200 || status >= 300) {
      // Server-authored, so echoed; the default bound, since nobody pastes a reason phrase back.
      const reason = (r as any)?.statusText;
      const phrase = typeof reason === 'string' && reason.trim().length > 0
        ? ` ${echoCallerText(reason)}`
        : '';
      throw new Error(`Failed to ${action}: server returned ${status}${phrase}`);
    }
    if (ok === false) {
      throw new Error(`Failed to ${action}: server rejected the request`);
    }
  }
}

// ---- write-side time zone result (#157) ----
//
// What create/update actually put on the wire for `start`/`end`, computed from the WRITTEN
// line, never the caller's input, so an inherited or defaulted zone is reported truthfully.
export interface CalendarZoneWriteInfo {
  /**
   * 'zoned'    — a TZID was written (from `timeZone`, inherited, or create's default).
   * 'utc'      — the value carries Z; the caller passed (or the stored line already named) a
   *              fixed instant.
   * 'floating' — no TZID, no Z: nothing to inherit and no default applied (only reachable on
   *              update — create always has a default zone to fall back to).
   * 'allday'   — date-only; there is no time component and so no zone to report.
   */
  kind: 'zoned' | 'utc' | 'floating' | 'allday';
  /** The IANA zone name written. Set only when kind === 'zoned'. */
  zone?: string;
}

export interface CreateCalendarEventResult {
  eventId: string;
  start: CalendarZoneWriteInfo;
  end: CalendarZoneWriteInfo;
  /** Collections that failed to list at discovery, so were no possible target here (#136). */
  brokenCollections?: BrokenCollections;
}

export interface UpdateCalendarEventResult {
  eventId: string;
  /** Present only when this call actually wrote (touched) that side. */
  start?: CalendarZoneWriteInfo;
  end?: CalendarZoneWriteInfo;
  /** Collections that failed to list, so were not searched for another copy (#136). */
  brokenCollections?: BrokenCollections;
}

function classifyWrittenLine(formatted: FormattedDateProperty): CalendarZoneWriteInfo {
  const d = describeDateProperty(formatted.line);
  switch (d.frame) {
    case 'date': return { kind: 'allday' };
    case 'utc': return { kind: 'utc' };
    case 'floating': return { kind: 'floating' };
    case 'zoned': return { kind: 'zoned', zone: d.tzid! };
  }
}

function describeCalendarZoneWrite(info: CalendarZoneWriteInfo): string {
  switch (info.kind) {
    case 'zoned': return `zone ${info.zone}${etcGmtOffsetNote(info.zone ?? '')}`;
    case 'utc': return 'UTC';
    case 'floating': return 'floating (no zone)';
    case 'allday': return 'all-day (no time component)';
  }
}

/** Refuse a present, non-string text field by its type; escapeICalText throws a TypeError on one. */
function assertTextType(name: string, value: unknown): void {
  if (value != null && typeof value !== 'string') {
    throw new InvalidInputError(`${name} must be a string; received ${Array.isArray(value) ? 'array' : typeof value}.`);
  }
}

/**
 * The trailing note a calendar read carries when an event's `timeZone`/`endTimeZone` is a signed
 * Etc/GMT name, whose sign is the inverse of its offset (`etcGmtOffsetNote`). A note rather than
 * a field, so both fields stay the zone name a caller can pass back.
 */
export function buildEtcGmtZoneNote(events: CalendarEvent[]): string {
  const zones = new Set<string>();
  for (const e of events) {
    for (const zone of [e.timeZone, e.endTimeZone]) {
      if (zone && etcGmtUtcOffset(zone)) zones.add(zone);
    }
  }
  if (zones.size === 0) return '';
  const offsets = [...zones].map(z => `${z} is UTC${etcGmtUtcOffset(z)}`);
  return `\n\nNote: ${offsets.join(' and ')}; an Etc/GMT name carries the POSIX sign, the inverse of the offset.`;
}

// The sentence create_calendar_event appends, from the written result only (#157).
export function describeCreateCalendarEventResult(result: CreateCalendarEventResult): string {
  const startDesc = describeCalendarZoneWrite(result.start);
  const endDesc = describeCalendarZoneWrite(result.end);
  return startDesc === endDesc
    ? ` Written in ${startDesc}.`
    : ` Start written in ${startDesc}, end written in ${endDesc}.`;
}

// As above, omitting a side the update did not touch.
export function describeUpdateCalendarEventResult(result: UpdateCalendarEventResult): string {
  const parts: string[] = [];
  if (result.start) parts.push(`start ${describeCalendarZoneWrite(result.start)}`);
  if (result.end) parts.push(`end ${describeCalendarZoneWrite(result.end)}`);
  if (parts.length === 0) return '';
  return ` (${parts.join(', ')})`;
}

/**
 * `timeZone` only qualifies a designator-less value; with a `Z`/offset or a date-only value it
 * is a contradiction, and is rejected rather than ignored (docs/conventions.md). Runs before
 * anything is written.
 *
 * Validates under `DTSTART`/`DTEND`, as the formatting call downstream does, so a malformed date
 * gets one rejection sentence whether or not `timeZone` was passed.
 */
function rejectTimezoneConflict(value: string, label: 'start' | 'end', callerZone: string): void {
  const propName = label === 'start' ? 'DTSTART' : 'DTEND';
  const serialized = validateAndFormatICalDate(value, propName);
  // Echoed although validated: validation tested `value.trim()`, and `.trim()` strips U+2028,
  // so leading line separators survive into `value` itself.
  if (/^\d{8}$/.test(serialized)) {
    throw new InvalidInputError(
      `timeZone cannot be combined with a date-only ${label} ("${echoCallerText(value)}") — an all-day value has ` +
      `no time zone. Drop timeZone, or pass ${label} with a time component for it to qualify.`
    );
  }
  if (serialized.endsWith('Z')) {
    throw new InvalidInputError(
      `timeZone cannot be combined with a ${label} that already carries Z or a UTC offset ("${echoCallerText(value)}") ` +
      `— that value already names a fixed instant of its own. Drop timeZone, or pass ${label} as a ` +
      `bare wall-clock value (no Z, no offset) for timeZone to qualify.`
    );
  }
}

/**
 * On update, `timeZone` with only ONE of `start`/`end` can strand the other side in a different
 * stored zone, a two-zone event (#140) nobody asked for that ordering will not catch. Fires only
 * on a stored, differently-named `zoned` value; floating or `Z` already trips the frame check.
 */
function rejectStrandedZoneMismatch(originalVevent: string, updatedSide: 'start' | 'end', callerZone: string): void {
  const strandedProp = updatedSide === 'start' ? 'DTEND' : 'DTSTART';
  const strandedLabel = updatedSide === 'start' ? 'end' : 'start';
  const strandedLines = parseAllICalProperties(originalVevent, strandedProp);
  if (strandedLines.length === 0) return;
  const desc = describeDateProperty(strandedLines[0]);
  if (desc.frame === 'zoned' && desc.tzid && !zoneNamesEqual(desc.tzid, callerZone)) {
    // Quoted differently on purpose: `callerZone` is ICU's canonical spelling (server text), the
    // stored tzid is invitation-authored, so it is echoed inside DOUBLE quotes (#190).
    throw new InvalidInputError(
      `timeZone would rewrite ${updatedSide} into '${callerZone}'${etcGmtOffsetNote(callerZone)} while the stored ${strandedLabel} stays ` +
      `in "${echoCallerText(desc.tzid, ZONE_ECHO_LIMIT)}"${etcGmtOffsetNote(desc.tzid)} untouched — silently producing a two-zone event. ` +
      `Pass BOTH start and end alongside timeZone (re-send the ${strandedLabel} you are not otherwise ` +
      `moving, unchanged, to keep its wall clock), or omit timeZone.`
    );
  }
}

export class CalDAVCalendarClient {
  private config: CalDAVConfig;
  private client: DAVClient | null = null;
  private calendars: DAVCalendar[] | null = null;

  constructor(config: CalDAVConfig) {
    this.config = config;
  }

  private async getClient(): Promise<DAVClient> {
    if (this.client) return this.client;

    const client = new DAVClient({
      serverUrl: this.config.serverUrl || 'https://caldav.fastmail.com',
      credentials: {
        username: this.config.username,
        password: this.config.password,
      },
      authMethod: 'Basic',
      defaultAccountType: 'caldav',
      // Following a redirect would replay the Basic credential at whatever host it names.
      // tsdav merges this into every underlying fetch. See docs/security-model.md.
      fetchOptions: { redirect: 'error' },
    });

    // Assigned to `this.client` only after login() resolves, so a failed login is retried
    // next call instead of a dead client being cached (#143).
    try {
      await client.login();
    } catch (e) {
      const detail = e instanceof Error ? e.message : String(e);
      // A plain Error: configuration, not caller input. `detail` is remote-authored, so it takes
      // `describeUntrustedAt`, which also redacts an echoed credential (#182).
      throw new Error(
        `CalDAV login failed: ${describeUntrustedAt(detail, LOGIN_FAILURE_ECHO_LIMIT)}. Check the configured CalDAV app password ` +
        `(a separate credential from the Fastmail JMAP API token).`,
      );
    }

    this.client = client;
    return this.client;
  }

  /**
   * Discover the account's calendars, failing loudly instead of returning an empty list.
   *
   * tsdav's `fetchCalendars` never checks `response.ok`, so a failed PROPFIND returns `[]`,
   * which read as "no events" (#100). Only a non-empty result is cached.
   *
   * An empty-BUT-SUCCESSFUL discovery is a failure too; do not "tidy" it into `[]`. A Fastmail
   * account always has a calendar, and an empty answer to an availability question reads as
   * free time. Revisit if that account-shape assumption stops holding.
   */
  private async discoverCalendars(): Promise<DiscoveredCalendars> {
    const client = await this.getClient();
    if (this.calendars && this.calendars.length > 0) {
      // Nothing is cached while a collection is broken. Residual, stated on the tool surface: a
      // calendar that breaks after a healthy discovery stays listed for the process's life.
      return { calendars: this.calendars, brokenCollections: [] };
    }

    // THE SAME ANSWER TSDAV PARSED, NEVER A SECOND REQUEST (#136), which could disagree. The
    // wrapper delegates to the client's own fetch, so `redirect: 'error'` still applies. The
    // FIRST request is the home PROPFIND; if a future tsdav reorders, nothing sits under the home
    // and nothing is flagged, the safe direction.
    const baseFetch = client.fetchOverride ?? globalThis.fetch;
    let capturedHomeListing: CapturedDavAnswer | undefined;
    const capturingFetch: typeof globalThis.fetch = async (input, init) => {
      const response = await baseFetch(input, init);
      if (capturedHomeListing === undefined) {
        // CLONED, because tsdav reads the body itself and a body can only be read once.
        capturedHomeListing = {
          body: await response.clone().text(),
          status: response.status,
          contentType: response.headers.get('content-type'),
        };
      }
      return response;
    };

    const calendars = await client.fetchCalendars({ fetch: capturingFetch });
    // An empty home URL needs no case: `findBrokenCalendarHomeCollections` reports nothing for it.
    const homeUrl = (client as any).account?.homeUrl;
    const brokenCollections = typeof homeUrl === 'string' && capturedHomeListing !== undefined
      ? await brokenCollectionsInAnswer(capturedHomeListing, homeUrl)
      : [];

    if (calendars.length > 0) {
      // NEVER CACHED WHILE A COLLECTION IS BROKEN, or a transient failure would become permanent
      // and its note would vanish on the next call.
      if (brokenCollections.length === 0) this.calendars = calendars;
      return { calendars, brokenCollections };
    }

    // Empty: re-ask with a status-checked PROPFIND, since the status is what fetchCalendars
    // discarded. Only on the already-broken path.
    await this.assertCalendarHomeReachable();
    throw new Error(
      'Calendar discovery returned no calendars. The CalDAV server answered successfully but listed ' +
      'no calendar collections, so no calendar could be read — this is reported as an error rather ' +
      'than an empty result, because an empty result would read as "there are no events".',
    );
  }

  /** Status-check the calendar-home PROPFIND that discovery runs, and throw on any non-2xx. */
  private async assertCalendarHomeReachable(): Promise<void> {
    const client = await this.getClient();
    const homeUrl = (client as any).account?.homeUrl;
    if (!homeUrl) {
      throw new Error('Calendar discovery failed: the CalDAV account has no calendar home URL.');
    }
    const responses = await client.propfind({
      url: homeUrl,
      props: { 'd:resourcetype': {} },
      depth: '1',
    });
    assertDavOk(responses, 'discover calendars');
  }

  async getCalendars(): Promise<CalendarListResult> {
    const { calendars, brokenCollections } = await this.discoverCalendars();

    const listed = selectableCalendars(calendars)
      .map(c => ({
        id: c.url || '',
        displayName: unwrapDisplayName(c.displayName) ?? 'Unnamed',
        url: c.url || '',
        description: c.description || undefined,
        // tsdav passes `calendarColor` through raw, so `<calendar-color/>` arrives as `{}`. Not
        // `unwrapDisplayName`: its number/boolean coercions would make a bad colour plausible.
        color: typeof (c as any).calendarColor === 'string' && (c as any).calendarColor.trim().length > 0
          ? (c as any).calendarColor.trim()
          : undefined,
      }));
    return { calendars: listed, brokenCollections: asBrokenCollectionsField(brokenCollections) };
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50, startDate?: string, endDate?: string): Promise<CalendarEventQueryResult> {
    const client = await this.getClient();
    const { calendars, brokenCollections } = await this.discoverCalendars();

    let targetCalendars = selectableCalendars(calendars);
    // PRESENCE, not truthiness: a narrowing argument fails CLOSED, so `''` must be refused, never
    // read as "every calendar" (docs/conventions.md).
    if (calendarId !== undefined && calendarId !== null) {
      targetCalendars = [resolveCalendarTarget(calendarId, targetCalendars, brokenCollections)];
    }

    // Normalised ONCE for both the server's time range and the local re-filter. A DATE IS A
    // LOCAL DAY in the configured zone, the same one email dates render in (docs/conventions.md).
    const zone = getDefaultTimezone();
    // `zone` stays raw for the window coercions, which hand it to `zoneOffsetMsAt` unresolved;
    // safe only because an unusable configured zone stops the server at startup (#157).
    // `configuredZone` is the one value the filter and the sort both read.
    const configuredZone = resolveUsableTimezone(zone);
    const rawStart = coerceCalendarWindowStart(startDate, 'startDate', zone);
    const rawEnd = coerceCalendarWindowEnd(endDate, 'endDate', zone);

    // Caller-named bounds are saturated too; see `saturateInstant`.
    const start = rawStart === undefined ? undefined : saturateInstant(rawStart);
    const end = rawEnd === undefined ? undefined : saturateInstant(rawEnd);
    const saturated: NonNullable<CalendarWindowClamp['saturated']> = [];
    if (rawStart !== undefined && start !== rawStart) {
      saturated.push({ bound: 'startDate', edge: saturationEdge(start!) });
    }
    if (rawEnd !== undefined && end !== rawEnd) {
      saturated.push({ bound: 'endDate', edge: saturationEdge(end!) });
    }

    const fetchOptions: any = {};
    let windowClamp: CalendarWindowClamp | undefined;
    // THE WINDOW THE CALLER ASKED FOR (#162). Everything caller-facing derives from these, never
    // from the widened `fetchOptions.timeRange`.
    let trueWindowStart: string | undefined;
    let trueWindowEnd: string | undefined;
    // UNCONDITIONAL: every listing is windowed, since tsdav drops `expand` without a time range
    // and the window bounds what the server materialises (#142).
    {
      let windowStart = start;
      let windowEnd = end;
      let invented: 'startDate' | 'endDate' | 'both' | undefined;
      if (!windowStart && !windowEnd) {
        // The local-day rule a date-only `startDate` gets; the clock is read once.
        windowStart = startOfLocalDayUtcIso((this.config.now ?? Date.now)(), zone);
        windowEnd = shiftIsoDays(windowStart, CALENDAR_OPEN_WINDOW_DAYS);
        invented = 'both';
      } else if (!windowStart) {
        windowStart = shiftIsoDays(windowEnd!, -CALENDAR_OPEN_WINDOW_DAYS);
        invented = 'startDate';
      } else if (!windowEnd) {
        windowEnd = shiftIsoDays(windowStart, CALENDAR_OPEN_WINDOW_DAYS);
        invented = 'endDate';
      }

      // AFTER the invented half is filled in: saturation can make a ONE-SIDED window
      // zero-length (`startDate: "9999-12-31T23:59:59Z"`).
      if (Date.parse(windowStart!) >= Date.parse(windowEnd!)) {
        // Checked here because tsdav's plain Error would read as InternalError
        // (docs/conventions.md). The message quotes what the caller TYPED beside what it
        // resolved to, and names the zone, so a local-day reading is recognisable as one.
        const resolved = windowStart === windowEnd
          ? `both resolve to the same instant, ${windowStart}, which is a zero-length window`
          : `which resolve to the range ${windowStart} .. ${windowEnd}`;
        const quote = (v: unknown) => (v === undefined || v === null ? '(omitted)' : `"${echoCallerText(v)}"`);
        // An invented bound needs its own advice: "put startDate before endDate" is unfollowable
        // for a caller who passed one or neither.
        const advice = invented === 'both'
          ? 'Neither startDate nor endDate was given, so the window was taken as a month from today — but today ' +
            'sits at the edge of the range this server can express, so that month had nowhere to go. Pass ' +
            'startDate and endDate.'
          : invented
          ? `${invented} was not given, so it was filled in ${CALENDAR_OPEN_WINDOW_DAYS} days away — but the bound ` +
            'you did give sits at the edge of the range this server can express, so the invented one had nowhere ' +
            'to go. Pass both startDate and endDate.'
          : `Dates are read as whole days in ${describeTimezone(zone)}, and a date-only endDate covers the whole of ` +
            'that day — so a single-day window is startDate and endDate on the SAME date, written as dates rather ' +
            'than as the same instant twice.';
        throw new InvalidInputError(
          `startDate must be before endDate (got startDate ${quote(startDate)}, endDate ${quote(endDate)}, ` +
          `${resolved}). ${advice}`,
        );
      }

      // Against the pre-widening bounds: the margin below does not change the caller's window.
      if (invented || saturated.length > 0) {
        windowClamp = { invented, saturated: saturated.length > 0 ? saturated : undefined, start: windowStart!, end: windowEnd! };
      }
      trueWindowStart = windowStart;
      trueWindowEnd = windowEnd;

      // WIDENED on both edges; see MAX_UTC_OFFSET_MS. A saturated widening is deliberately NOT
      // disclosed in `saturated[]`: that reports caller-named bounds, and the exact filter drops
      // whatever the lost hours would have reached anyway.
      fetchOptions.timeRange = {
        start: shiftIsoMs(windowStart!, -MAX_UTC_OFFSET_MS),
        end: shiftIsoMs(windowEnd!, MAX_UTC_OFFSET_MS),
      };
      // Reports the in-window occurrence rather than the original DTSTART (#64). tsdav forwards
      // <C:expand> only alongside a timeRange.
      fetchOptions.expand = true;
    }

    // From the TRUE window, never `fetchOptions.timeRange`.
    const windowStartMs = trueWindowStart === undefined ? NaN : Date.parse(trueWindowStart);
    const windowEndMs = trueWindowEnd === undefined ? NaN : Date.parse(trueWindowEnd);

    const allEvents: CalendarEvent[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({
        calendar: cal,
        ...fetchOptions,
        urlFilter: calendarResourceUrlFilter(cal.url),
      });
      // Settled in one request per calendar; see `settleAmbiguousRecurrence` (#155).
      const undecided = new Map<string, UndecidedResource>();
      for (const obj of objects) {
        // ONE structural extraction per resource, counted here and handed to the parser.
        const blocks = extractVEventBlocks(obj.data || '');
        if (blocks.length > CALENDAR_MAX_OCCURRENCES_PER_SERIES) {
          // THE CALL FAILS on the first such resource, rather than burying the omission under a
          // complete-looking listing. Title and id come from the first block, without a parse.
          const title = parseICalValue(blocks[0], 'SUMMARY') || 'Untitled';
          // Not trimmed: no exact match here, and the echo trims it.
          const id = parseICalValue(blocks[0], 'UID') || obj.url || '';
          const calendar = unwrapDisplayName(cal.displayName) ?? cal.url ?? '';
          // Invitation-authored values, echoed (#141). The guarantee is no extra LINES, not no
          // attacker prose inside the quotes.
          throw new InvalidInputError(
            `Refused: repeating event "${echoCallerText(title)}" (id ${echoCallerText(id)}, ` +
            `calendar ${echoCallerText(calendar)}) expands to ${blocks.length} occurrences in the ` +
            `range searched for this window, more than the ${CALENDAR_MAX_OCCURRENCES_PER_SERIES} ` +
            'this server will materialise for one series. This is a deliberate limit; narrow the ' +
            'window to list around it. If you have a genuine use for a series this dense, open an ' +
            'issue at https://github.com/JonathanGodley/fastmail-mcp/issues.',
          );
        }
        // `expanded` is passed rather than sniffed; see parseCalendarObjects.
        const kept: CalendarEvent[] = [];
        for (const event of parseCalendarObjects(obj, { expanded: !!fetchOptions.expand, configuredZone, blocks })) {
          // A block still carrying RRULE or RDATE (#162) is never dropped on its dates: a master
          // the server declined to expand shows its ORIGINAL DTSTART, and judging that would turn
          // a wrongly-dated row into a missing one.
          const provablyOutside = !event.recurrenceRule
            && !event.recurrenceDates
            && !eventIntersectsWindow(event, windowStartMs, windowEndMs, configuredZone);
          if (provablyOutside) continue;
          allEvents.push(event);
          kept.push(event);
        }
        // Exactly the set `blockCountProvesSeries` cannot decide, among rows still kept.
        if (blocks.length === 1 && kept.length > 0 && !kept.some(e => e.isRecurring) && obj.url) {
          undecided.set(resolveResponseHref(obj.url, cal.url), {
            requestHref: toRequestHref(obj.url, cal.url),
            rows: kept,
          });
        }
      }
      await settleAmbiguousRecurrence(client, cal, undecided);
      // No early exit on `limit`: the slice is only a genuine top-N across calendars once every
      // calendar is read (#100).
    }

    sortEventsByStart(allEvents, configuredZone);

    return {
      events: allEvents.slice(0, limit),
      total: allEvents.length,
      windowClamp,
      // Collections that failed at discovery (#136). A calendar that listed and then failed on
      // its event read still fails the whole call.
      brokenCollections: asBrokenCollectionsField(brokenCollections),
    };
  }

  /**
   * Find every stored copy of the event this id names, by UID or by URL (#137).
   *
   * ONE TARGETED QUERY PER CALENDAR (`uidEqualsFilter`), then an exact-equality check. There is
   * deliberately NO FALLBACK to a full scan when the server refuses the filter: the caller sees
   * a failed call, not one that searched something other than it says.
   *
   * BOTH FORMS OF THE ID ARE TRIED AND UNIONED, never branched on the string's shape: a UID can
   * be url-shaped, and dispatching would make that event unfindable by its own id.
   */
  private async findCalendarObjectByUID(eventId: string): Promise<CalendarObjectLookup> {
    // Here, not in the handler, whose guard tests presence only; all three tools inherit it.
    if (eventId != null && typeof eventId !== 'string') {
      throw new InvalidInputError(
        `eventId must be a string; received ${Array.isArray(eventId) ? 'array' : typeof eventId}. `
        + 'Pass an event id or url from list_calendar_events.',
      );
    }
    const wanted = requireNonEmpty(eventId, 'eventId', 'pass an event id or url from list_calendar_events');
    if (XML_UNSENDABLE_CHARS.test(wanted)) {
      throw new InvalidInputError(
        'eventId contains a character that no XML request can carry, which no calendar id or '
        + 'url holds either. Pass an event id or url from list_calendar_events.',
      );
    }
    const client = await this.getClient();
    // Via discovery, so a discovery failure throws rather than reading as a wrong id (#100).
    const { calendars, brokenCollections } = await this.discoverCalendars();

    // SELECTABLE only: this server must not destroy a record no read tool would show.
    const selectable = selectableCalendars(calendars);

    const matches: CalendarObjectMatch[] = [];
    const seenUrls = new Set<string>();
    const collect = (calendar: DAVCalendar, obj: DAVCalendarObject) => {
      if (!isResolvedCalendarObject(obj)) return;
      // Deduped on resource url: an id can resolve the same resource by UID and by url.
      const url = obj.url;
      if (seenUrls.has(url)) return;
      seenUrls.add(url);
      matches.push({ object: obj, calendarLabel: calendarLabel(calendar) });
    };

    const urlTargets = resolveEventUrlTargets(wanted, selectable);
    const addressedHrefs = new Set(urlTargets.map(t => addressComparisonKey(t.objectUrl)));

    // No early exit: a UID is unique per collection, not per account (#101).
    const uidHolders = async (uid: string) => {
      const holders: Array<{ calendar: DAVCalendar; obj: DAVCalendarObject }> = [];
      for (const calendar of selectable) {
        const objects = await client.fetchCalendarObjects({
          calendar,
          filters: uidEqualsFilter(uid),
          urlFilter: calendarResourceUrlFilter(calendar.url),
        });
        for (const obj of objects) {
          // LOAD-BEARING: the server matches case-insensitively; see `uidEqualsFilter`.
          if (ownUid(obj) === uid) holders.push({ calendar, obj });
        }
      }
      return holders;
    };
    for (const { calendar, obj } of await uidHolders(wanted)) collect(calendar, obj);

    // The URL form; see resolveEventUrlTargets.
    for (const { calendar, objectUrl } of urlTargets) {
      let objects: DAVCalendarObject[];
      try {
        objects = await client.fetchCalendarObjects({
          calendar,
          objectUrls: [objectUrl],
          urlFilter: calendarResourceUrlFilter(calendar.url),
        });
      } catch (err) {
        // NOT-FOUND ONLY IS SWALLOWED. Any other failure is about the collection, and swallowing
        // it would drop a copy from the ambiguity count so a write that should be refused goes
        // through. A collection-level 404 is safe to swallow: this collection answered the UID
        // query just above, so what is left is a race meaning the same thing.
        if (!isAddressedResourceMissing(err)) throw err;
        continue;
      }
      for (const obj of objects) collect(calendar, obj);
    }

    // The addressed copy LEADS (see CalendarObjectLookup); every other order is untouched.
    // `> 0` only skips a no-op, so it is indistinguishable from `>= 0` by any test.
    const addressedIndex = matches.findIndex(m => addressedHrefs.has(addressComparisonKey(m.object.url)));
    if (addressedIndex > 0) matches.unshift(...matches.splice(addressedIndex, 1));

    let collision: CalendarObjectLookup['collision'];
    if (addressedIndex !== -1 && matches.slice(1).some(m => ownUid(m.object) === wanted)) {
      // Offer the addressed record's UID only where it reaches that record alone: not absent,
      // not this same string, and not held by any other resolved record.
      const uid = ownUid(matches[0].object);
      const addressedKey = addressComparisonKey(matches[0].object.url);
      const reachesAlone = uid !== undefined && uid !== '' && uid !== wanted
        && !(await uidHolders(uid)).some(h => isResolvedCalendarObject(h.obj)
          && addressComparisonKey(h.obj.url) !== addressedKey);
      collision = { addressedUid: reachesAlone ? uid : undefined };
    }

    return { matches, addressed: addressedIndex !== -1, collision, brokenCollections };
  }

  async getCalendarEventById(eventId: string): Promise<CalendarEventResult> {
    const { matches, addressed, collision, brokenCollections } = await this.findCalendarObjectByUID(eventId);
    const obj = matches[0]?.object;
    if (!obj) {
      throw eventNotFoundError(eventId, brokenCollections);
    }
    return {
      // Always states free/busy (#194): on a single event, absence cannot mean "busy".
      event: parseCalendarObject(obj, { includeParticipants: true, includeDefaultTransparency: true, configuredZone: resolveUsableTimezone(getDefaultTimezone()) }),
      // ANSWERS AND DISCLOSES where the writes refuse (#101): a read harms no copy, and this is
      // the tool that hands over each copy's url. Disclosed even when a copy was addressed.
      otherCopies: matches.length > 1 ? matchesToCopies(matches.slice(1)) : undefined,
      addressedByUrl: addressed,
      addressCollision: collision,
      brokenCollections: asBrokenCollectionsField(brokenCollections),
    };
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    /**
     * IANA zone name for a designator-less `start`/`end` (#157). Omitted writes the configured
     * zone, never floating (docs/conventions.md). `null` and blank are rejected, not read as
     * "floating".
     */
    timeZone?: string | null;
    /** Free/busy (#194); overrides the all-day default. See the TRANSP block below. */
    transparency?: string;
  }): Promise<CreateCalendarEventResult> {
    // Before discovery, and by update's rules.
    assertTextType('title', event.title);
    const title = requireNonEmpty(event.title, 'title', 'pass the event title');
    assertTextType('description', event.description);
    assertTextType('location', event.location);

    const client = await this.getClient();
    const { calendars, brokenCollections } = await this.discoverCalendars();

    // The same resolver as the read path (#173). A broken collection is never in `selectable`,
    // so the write cannot land in one (#136).
    const selectable = selectableCalendars(calendars);
    const targetCal = resolveCalendarTarget(event.calendarId, selectable, brokenCollections);

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

    // The conflict rules need the start/end values, so they run here rather than in
    // validateCallerTimezone (#157).
    let callerZone: string | undefined;
    if (event.timeZone !== undefined) {
      callerZone = validateCallerTimezone(event.timeZone);
      rejectTimezoneConflict(event.start, 'start', callerZone);
      rejectTimezoneConflict(event.end, 'end', callerZone);
    }

    const callerTransparency = event.transparency !== undefined
      ? normalizeTransparency(event.transparency)
      : undefined;

    // Create's default for a designator-less value, so it is never written floating (#157).
    const configuredZone = resolveUsableTimezone(getDefaultTimezone());

    const startFormatted = formatDateTimeProperty('DTSTART', event.start, null, '\r\n', callerZone, configuredZone);
    const endFormatted = formatDateTimeProperty('DTEND', event.end, null, '\r\n', callerZone, configuredZone);
    const startLine = startFormatted.line;
    const endLine = endFormatted.line;

    // Classified from the serialized lines, as update does, so both reject the same pairs.
    const startFrame = describeDateProperty(startLine, event.start, startFormatted.tzidSource);
    const endFrame = describeDateProperty(endLine, event.end, endFormatted.tzidSource);
    validateDateConsistency(startFrame, endFrame);

    // RFC 5545 §3.6.5 requires a VTIMEZONE for every TZID a component uses. A `text/calendar`
    // PUT — the only body type this server sends — gets none from Cyrus, so this generates one
    // instead (see vtimezone.ts). For what a CalDAV PUT does and does not reach on the platform
    // side, and why, see docs/conventions.md's "The VTIMEZONE residual" (#166).
    const createZoneTzids = referencedZoneTzids([startFrame, endFrame]);
    const createInstants = collectZoneInstants([
      { label: 'start', frame: startFrame },
      { label: 'end', frame: endFrame },
    ]);
    const vtimezoneBlocks = createInstants.length === 0 ? [] : Array.from(
      createZoneTzids,
      tzid => generateVTimezone(tzid, Math.min(...createInstants), Math.max(...createInstants), '\r\n'),
    );

    const icalLines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      ...vtimezoneBlocks.flatMap(block => block.split('\r\n')),
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `LAST-MODIFIED:${now}`,
      startLine,
      endLine,
      foldICalLine(`SUMMARY:${escapeICalText(title)}`),
    ];

    // ALL-DAY IS WRITTEN FREE, TIMED IS LEFT BUSY: the Fastmail client's defaults (#195,
    // docs/fastmail-action-availability.md). An absent TRANSP means OPAQUE (RFC 5545 §3.8.2.7),
    // so the timed path deliberately writes nothing. Choosing a default is create's alone;
    // update never writes TRANSP unasked. An explicit `transparency` wins (#194).
    if (callerTransparency !== undefined) {
      icalLines.push(transpLine(callerTransparency));
    } else if (startFrame.frame === 'date') {
      icalLines.push(transpLine('free'));
    }

    if (event.description) {
      icalLines.push(foldICalLine(`DESCRIPTION:${escapeICalText(event.description)}`));
    }
    if (event.location) {
      icalLines.push(foldICalLine(`LOCATION:${escapeICalText(event.location)}`));
    }

    if (event.participants && event.participants.length > 0) {
      for (const p of event.participants) {
        validateAttendeeEmail(p.email);
      }

      // ORGANIZER is required when ATTENDEEs are present.
      const caldavUsername = this.config.username;
      validateOrganizerUsername(caldavUsername);
      const displayName = resolveDisplayName(this.config.displayName, caldavUsername);
      const cnPart = `;CN=${quoteParamValue(displayName)}`;
      icalLines.push(foldICalLine(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));

      // No RSVP=TRUE by default (RFC 5545 §3.2.17 defaults to FALSE).
      for (const p of event.participants) {
        const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
        icalLines.push(foldICalLine(`ATTENDEE${cnParam}:mailto:${p.email}`));
      }
    }

    icalLines.push('END:VEVENT');
    icalLines.push('END:VCALENDAR');

    // Trailing CRLF per RFC 5545 §3.1.
    const ical = icalLines.join('\r\n') + '\r\n';

    const createResp = await client.createCalendarObject({
      calendar: targetCal,
      filename: `${uid}.ics`,
      iCalString: ical,
    });
    assertDavOk(createResp, 'create calendar event');

    return {
      eventId: uid,
      start: classifyWrittenLine(startFormatted),
      end: classifyWrittenLine(endFormatted),
      brokenCollections: asBrokenCollectionsField(brokenCollections),
    };
  }

  async updateCalendarEvent(eventId: string, fields: {
    title?: string;
    description?: string;
    start?: string;
    end?: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    clearFields?: string[];
    // Explicit zone for a designator-less start/end (#157). Omitted: the stored TZID, else
    // floating on a floating event, else the configured zone.
    timeZone?: string | null;
    /**
     * Free/busy (#194). Omitted leaves the stored value alone; exclusive with
     * `clearFields: ['transparency']`.
     */
    transparency?: string;
  }): Promise<UpdateCalendarEventResult> {
    const client = await this.getClient();
    // PROCEEDS ON THE COPY IT FOUND, naming the collection it could not check (#136). Refusing
    // every write while one collection is unhealthy is NOT the safe default here: it guards only
    // a duplicate UID whose other copy sits in the broken collection, and this path writes once,
    // to the one resource it resolved.
    const { matches, addressed, collision, brokenCollections } = await this.findCalendarObjectByUID(eventId);
    const obj = matches[0]?.object;
    if (!obj) {
      throw eventNotFoundError(eventId, brokenCollections);
    }

    // Before the repeating-series refusal and all argument validation; see ambiguousEventIdError
    // (#101). `addressed` keeps the url escape hatch open.
    if (!addressed && matches.length > 1) {
      throw ambiguousEventIdError(eventId, 'update', matchesToCopies(matches), brokenCollections);
    }
    if (collision) {
      throw addressCollisionError(eventId, 'update', matchesToCopies(matches), collision.addressedUid, brokenCollections);
    }

    // UNREACHABLE DEFENCE (#137): `isResolvedCalendarObject` already requires a VEVENT. A plain
    // Error, since no argument could fix it.
    if (!obj.data || extractVEventBlocks(obj.data).length === 0) {
      throw new Error('Cannot update event: no iCal data found');
    }

    // Before ANY argument validation: nothing in the arguments can make this legal. Keyed on
    // the resolved RESOURCE, so the url form of the id reaches it too (#146).
    if (isRecurringSeriesResource(obj.data)) {
      throw recurringSeriesRefusal('update', calendarObjectTitle(obj.data, eventId));
    }

    const datePattern = /^\d{4}-\d{2}-\d{2}$/;
    const dateTimePattern = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;
    // Trimmed as validateAndFormatICalDate trims, or a padded value it accepts is refused here.
    const isoShaped = (value: string) => {
      const trimmed = String(value).trim();
      return datePattern.test(trimmed) || dateTimePattern.test(trimmed);
    };
    // Unvalidated by definition, so echoed.
    if (fields.start !== undefined && !isoShaped(fields.start)) {
      throw new InvalidInputError(`Invalid start date format: "${echoCallerText(fields.start)}". Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }
    if (fields.end !== undefined && !isoShaped(fields.end)) {
      throw new InvalidInputError(`Invalid end date format: "${echoCallerText(fields.end)}". Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }

    // `transparency` is clearable (#194) though an enum: removing TRANSP is the only way back to
    // the shape the Fastmail client writes for a busy event, a change to the record, not the state.
    const CLEARABLE_FIELDS = new Set(['description', 'location', 'transparency']);
    const providedStringFields = new Set<string>();
    if (fields.description !== undefined) providedStringFields.add('description');
    if (fields.location !== undefined) providedStringFields.add('location');
    if (fields.transparency !== undefined) providedStringFields.add('transparency');
    validateClearFields(fields.clearFields, CLEARABLE_FIELDS, providedStringFields);

    const callerTransparency = fields.transparency !== undefined
      ? normalizeTransparency(fields.transparency)
      : undefined;

    const lineEnding = detectLineEnding(obj.data);
    const fold = (line: string) => foldICalLine(line, lineEnding);

    const normalizedData = normalizeMasterVEventFirst(obj.data);

    const originalVevent = extractVEvent(normalizedData);
    if (!originalVevent) {
      throw new Error('Cannot update event: no VEVENT block found');
    }

    // Trimmed: returned as `eventId`, which a caller may feed to an exact-match lookup.
    const existingUid = parseICalValue(originalVevent, 'UID')?.trim() || eventId;
    let data = normalizedData;

    // timeZone validation (#157), before any patching.
    let callerZone: string | undefined;
    if (fields.timeZone !== undefined) {
      callerZone = validateCallerTimezone(fields.timeZone);
      if (fields.start === undefined && fields.end === undefined) {
        throw new InvalidInputError(
          `timeZone was supplied ('${callerZone}') but neither start nor end was. timeZone only ` +
          "qualifies a start/end value being written in this same call — it cannot be applied to a " +
          "stored value on its own. Re-send start and/or end (even unchanged) alongside timeZone, " +
          "or drop timeZone."
        );
      }
      if (fields.start !== undefined) rejectTimezoneConflict(fields.start, 'start', callerZone);
      if (fields.end !== undefined) rejectTimezoneConflict(fields.end, 'end', callerZone);
      if (fields.start !== undefined && fields.end === undefined) {
        rejectStrandedZoneMismatch(originalVevent, 'start', callerZone);
      }
      if (fields.end !== undefined && fields.start === undefined) {
        rejectStrandedZoneMismatch(originalVevent, 'end', callerZone);
      }
    }

    let newStartFormatted: FormattedDateProperty | null = null;
    let newEndFormatted: FormattedDateProperty | null = null;
    let newStartLine: string | null = null;
    let newEndLine: string | null = null;
    let timeChanged = false;

    if (fields.title !== undefined) {
      assertTextType('title', fields.title);
      const title = requireNonEmpty(fields.title, 'title');
      data = replaceICalProperty(data, 'SUMMARY', fold(`SUMMARY:${escapeICalText(title)}`));
    }

    if (fields.description !== undefined) {
      const description = requireNonEmpty(fields.description, 'description');
      data = replaceICalProperty(data, 'DESCRIPTION', fold(`DESCRIPTION:${escapeICalText(description)}`));
    }

    // Floating only on a floating event: keeping it floating keeps the event's own frame.
    const storedStartLine = parseAllICalProperties(originalVevent, 'DTSTART')[0];
    const defaultZone = storedStartLine && describeDateProperty(storedStartLine).frame === 'floating'
      ? undefined
      : resolveUsableTimezone(getDefaultTimezone());

    if (fields.start !== undefined) {
      newStartFormatted = formatDateTimeProperty('DTSTART', fields.start, originalVevent, lineEnding, callerZone, defaultZone);
      newStartLine = newStartFormatted.line;
      data = replaceICalProperty(data, 'DTSTART', newStartLine);
      timeChanged = true;
    }

    if (fields.end !== undefined) {
      newEndFormatted = formatDateTimeProperty('DTEND', fields.end, originalVevent, lineEnding, callerZone, defaultZone);
      newEndLine = newEndFormatted.line;
      data = replaceICalProperty(data, 'DTEND', newEndLine);
      // DTEND and DURATION are mutually exclusive (RFC 5545 §3.6.1).
      data = removeAllICalProperties(data, 'DURATION');
      timeChanged = true;
    }

    // AN UPDATE CHANGES FREE/BUSY ONLY WHEN ASKED (#195, #194): an absent TRANSP already says
    // busy, so there is no gap to fill. No fold or escape: the line is a short owned literal.
    if (callerTransparency !== undefined) {
      data = replaceICalProperty(data, 'TRANSP', transpLine(callerTransparency));
    }

    // ONE OCCURRENCE EACH: RFC 5545 allows these at most once per VEVENT.
    // `removeAllICalProperties` is the tool if that changes.
    if (fields.clearFields && fields.clearFields.length > 0) {
      const KEY_BY_FIELD: Record<string, string> = { description: 'DESCRIPTION', location: 'LOCATION', transparency: 'TRANSP' };
      for (const field of fields.clearFields) {
        data = replaceICalProperty(data, KEY_BY_FIELD[field], null);
      }
    }

    // Judged on the pair that will be WRITTEN, the stored line standing in for an untouched side,
    // which is where single-sided frame flips come from. Skipped when neither side changed, so
    // a title edit is never blocked by an inconsistency already stored.
    if (fields.start !== undefined || fields.end !== undefined) {
      const startLine = newStartLine ?? parseAllICalProperties(originalVevent, 'DTSTART')[0];
      const endLine = newEndLine ?? parseAllICalProperties(originalVevent, 'DTEND')[0];
      if (startLine && endLine) {
        validateDateConsistency(
          describeDateProperty(startLine, newStartLine ? fields.start : undefined, newStartFormatted?.tzidSource),
          describeDateProperty(endLine, newEndLine ? fields.end : undefined, newEndFormatted?.tzidSource)
        );
      } else if (newStartLine && !newEndLine) {
        // A kept DURATION stands in for DTEND: a DATE DTSTART takes only a dur-day or dur-week
        // one (RFC 5545 §3.6.1).
        const duration = parseICalValue(originalVevent, 'DURATION')?.trim();
        const start = describeDateProperty(newStartLine, fields.start);
        if (duration && start.frame === 'date' && duration.includes('T')) {
          throw new InvalidInputError(
            `A date-only DTSTART takes only a whole-day DURATION per RFC 5545 §3.6.1 — start "${echoCallerText(start.display)}" `
            + `is ${describeFrame(start)} but the stored DURATION "${echoCallerText(duration)}" has a time part. `
            + 'Pass start with a time, or pass end as well, which replaces the DURATION.',
          );
        }
      }
    }

    if (fields.location !== undefined) {
      const location = requireNonEmpty(fields.location, 'location');
      data = replaceICalProperty(data, 'LOCATION', fold(`LOCATION:${escapeICalText(location)}`));
    }

    if (fields.participants !== undefined) {
      for (const p of fields.participants) {
        validateAttendeeEmail(p.email);
      }
      data = removeAllICalProperties(data, 'ATTENDEE');
      // An ORGANIZER with no ATTENDEEs is a malformed scheduling VEVENT (RFC 5545 §3.8.4.3).
      if (fields.participants.length === 0) {
        data = removeAllICalProperties(data, 'ORGANIZER');
      }
      if (fields.participants.length > 0) {
        const attendeeLines = fields.participants.map(p => {
          const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
          return fold(`ATTENDEE${cnParam}:mailto:${p.email}`);
        }).join(lineEnding);
        data = insertBeforeEndVEvent(data, attendeeLines);
      }
      // RFC 5545 §3.8.4.1.
      if (fields.participants.length > 0 && !hasICalProperty(extractVEvent(data) || '', 'ORGANIZER')) {
        const caldavUsername = this.config.username;
        validateOrganizerUsername(caldavUsername);
        const displayName = resolveDisplayName(this.config.displayName, caldavUsername);
        const cnPart = `;CN=${quoteParamValue(displayName)}`;
        data = replaceICalProperty(data, 'ORGANIZER', fold(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));
      }
    }

    const hasAttendees = hasICalProperty(originalVevent, 'ATTENDEE');
    const schedulingSignificant = fields.start !== undefined || fields.end !== undefined ||
      fields.participants !== undefined || fields.location !== undefined;

    if (hasAttendees && schedulingSignificant) {
      const existingSeq = parseInt(parseICalValue(originalVevent, 'SEQUENCE') || '0', 10) || 0;
      data = replaceICalProperty(data, 'SEQUENCE', `SEQUENCE:${existingSeq + 1}`);
    }

    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    data = replaceICalProperty(data, 'DTSTAMP', `DTSTAMP:${now}`);
    data = replaceICalProperty(data, 'LAST-MODIFIED', `LAST-MODIFIED:${now}`);

    // LAST, after every patch; regeneration before the orphan sweep.
    if (timeChanged) {
      data = regenerateVTimezones(data, lineEnding);
      data = removeOrphanedVTimezones(data);
    }

    obj.data = data;
    const updateResp = await client.updateCalendarObject({ calendarObject: obj });
    assertDavOk(updateResp, 'update calendar event');

    return {
      eventId: existingUid,
      start: newStartFormatted ? classifyWrittenLine(newStartFormatted) : undefined,
      end: newEndFormatted ? classifyWrittenLine(newEndFormatted) : undefined,
      brokenCollections: asBrokenCollectionsField(brokenCollections),
    };
  }

  async deleteCalendarEvent(eventId: string): Promise<DeleteCalendarEventResult> {
    const client = await this.getClient();
    // Proceeds on the copy it found, as update does (#136).
    const { matches, addressed, collision, brokenCollections } = await this.findCalendarObjectByUID(eventId);
    const obj = matches[0]?.object;
    if (!obj) {
      throw eventNotFoundError(eventId, brokenCollections);
    }

    // Same rule and order as update's; see ambiguousEventIdError (#101).
    if (!addressed && matches.length > 1) {
      throw ambiguousEventIdError(eventId, 'delete', matchesToCopies(matches), brokenCollections);
    }
    if (collision) {
      throw addressCollisionError(eventId, 'delete', matchesToCopies(matches), collision.addressedUid, brokenCollections);
    }

    // Keyed on the resolved RESOURCE, not the argument's shape, so passing a row's url cannot
    // bypass it (#146).
    if (isRecurringSeriesResource(obj.data)) {
      throw recurringSeriesRefusal('delete', calendarObjectTitle(obj.data, eventId));
    }

    const deleteResp = await client.deleteCalendarObject({ calendarObject: obj });
    assertDavOk(deleteResp, 'delete calendar event');
    const uid = ownUid(obj) || eventId;
    return { eventId: uid, url: obj.url, brokenCollections: asBrokenCollectionsField(brokenCollections) };
  }
}
