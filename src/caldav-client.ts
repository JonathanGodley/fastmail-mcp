import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
// requireNonEmpty/validateClearFields come from coerce.ts rather than being
// defined here, so their rejections throw the tagged InvalidInputError and the
// CallTool boundary maps them to InvalidParams. A plain Error would surface as
// InternalError ("server bug"), which is wrong for caller-fixable input and
// would tell the caller a bare retry might work. See docs/conventions.md.
import { InvalidInputError, requireNonEmpty, validateClearFields, coerceCalendarWindowStart, coerceCalendarWindowEnd, describeTimezone, resolveCalendarInstantMs, echoCallerText } from './coerce.js';
// The deployment's configured timezone, read from the ONE place it is stored — the value
// `setDefaultTimezone` holds and every email `date` renders in. A calendar window has to
// INTERPRET a local date rather than display one, but it must interpret it as the same zone
// the rest of the server displays, so it reads that value instead of re-deriving its own
// from the environment.
import { getDefaultTimezone } from './email-formatter.js';

export interface CalDAVConfig {
  username: string;
  password: string;
  serverUrl?: string;
  // Display name for the ORGANIZER this client emits. Resolved by the caller from the
  // environment (the server does that through its shared multi-name lookup, so a DXT
  // user_config spelling reaches it); left unset it falls back to the username.
  displayName?: string;
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
  // The resource's CalDAV URL — AND A SECOND, EQUALLY SERIES-WIDE DELETE HANDLE, because
  // `findCalendarObjectByUID` accepts it interchangeably with `id`. `delete_calendar_event`
  // given a row's `url` destroys the whole series exactly as its `id` does.
  //
  // KEPT rather than dropped from the row, and the choice is worth stating because "stop
  // emitting it" is the obvious alternative and breaking changes are cheap here. Two reasons
  // it stays: a UID is unique per collection, not per account, so where two calendars hold
  // the same UID the url is the only thing that tells the two records apart; and for a
  // resource carrying no VEVENT at all the url IS the id (see parseCalendarObjects' minimal
  // fallback), so removing the field would not remove the second handle, only hide it. It is
  // documented on the tool surface instead, in the same sentence that settles what `id` names.
  url: string;
  title: string;
  description?: string;
  start?: string;
  end?: string;
  location?: string;
  organizer?: Participant;
  participants?: Participant[];
  // ---- recurrence (#64) ----
  // Whether this entry belongs to a repeating series at all. Set for a series master AND
  // for a single expanded occurrence, so a caller can tell "this repeats" without having
  // to reason about which of the two fields below is present.
  isRecurring?: boolean;
  // The occurrence this entry IS, from RECURRENCE-ID. Its presence is the unambiguous
  // signal that `start`/`end` are the in-window occurrence rather than the series' original
  // DTSTART — which is exactly what #64 asked for, because a recurring event reported at
  // its first occurrence years earlier is indistinguishable from a one-off on that date.
  recurrenceId?: string;
  // The raw RRULE value ("FREQ=WEEKLY;BYDAY=MO"). Usually the mark of a series MASTER, which
  // shows its ORIGINAL DTSTART, as against an occurrence, which carries a `recurrenceId` and
  // shows the real date — reading the two together is how a caller knows which it is holding.
  //
  // The two are NOT mutually exclusive, though, and an earlier version of this comment said
  // they were. RFC 5545 §3.8.5.3 lets an override block carry its own recurrence rule, so a
  // block can have both; where it does, `recurrenceId` names the instance this entry IS and
  // `recurrenceRule` is that block's own rule rather than the series'. The tool description
  // states it the same way.
  recurrenceRule?: string;
}

// What `getCalendarEvents` returns: the page, plus how many events matched before `limit`
// trimmed it. The count is stated rather than left implicit because `limit` is a hard cap
// with no paging — a caller that reads a capped page as the whole answer concludes the
// calendar is empty past that point (#100). Expansion makes this sharper, not softer: a
// fortnightly event across a three-month window is now seven entries where it was one, so
// the cap is reached far more often than it used to be.
export interface CalendarEventQueryResult {
  events: CalendarEvent[];
  total: number;
  // Set only when the window actually queried was NOT the window the caller described. A
  // caller handed a narrower window than it asked for must be told, or "nothing after that
  // date" reads as an empty calendar. Absent means the window was honoured exactly.
  //
  // STRUCTURE, NOT PROSE — deliberately, and it used to be prose. The email listings already
  // established the shape for a disclosure of this kind: the client returns structured
  // metadata (`QueryResult.exclusion`), a formatter beside `buildExclusionNote` owns the
  // wording and the blank-line separator, and the handler concatenates. Building the finished
  // sentence down here instead put the wording somewhere no formatter test can reach it and
  // gave the separator convention a second home.
  windowClamp?: CalendarWindowClamp;
}

/** Why, and to what, a calendar window ended up narrower than the one the caller described. */
export interface CalendarWindowClamp {
  // The bound the caller OMITTED, which had to be invented CALENDAR_OPEN_WINDOW_DAYS away
  // from the one they gave. Absent when both bounds were named.
  invented?: 'startDate' | 'endDate';
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
 * Extract the VEVENT block from iCalendar data.
 * This avoids matching properties from VTIMEZONE or other components.
 */
/**
 * Resolve the ORGANIZER display name from the configured value, falling back when
 * it is unset, blank, or an unresolved DXT config placeholder like
 * "${user_config.fastmail_caldav_display_name}" — without that check the literal
 * placeholder would be embedded into generated iCal.
 *
 * The server resolves the value from the environment before constructing the client,
 * and its lookup rejects placeholders too. This stays the client's own guard so a
 * directly-constructed client (tests, embedders) cannot emit a placeholder CN either.
 */
export function resolveDisplayName(raw: string | undefined, fallback: string): string {
  const trimmed = raw?.trim();
  if (!trimmed || /\$\{[^}]+\}/.test(trimmed)) return fallback;
  return trimmed;
}

/**
 * ICALENDAR STRUCTURE IS DECIDED ON WHOLE CONTENT LINES, NEVER WITH A `/m`-ANCHORED REGEX.
 *
 * This is the single most load-bearing rule in this parser, and the reason the helpers below
 * exist at all rather than each site testing `/^BEGIN:VEVENT/m` for itself.
 *
 * RFC 5545 §3.1 knows exactly one line break: CRLF (and this server tolerates a bare LF,
 * because real servers emit it). JavaScript's `^`/`$` under `/m` additionally anchor after
 * U+2028 LINE SEPARATOR, U+2029 PARAGRAPH SEPARATOR and a BARE CR. None of those three is a
 * line break to iCalendar, and none of them has to be escaped inside a TEXT value — so all
 * three survive verbatim into a SUMMARY or DESCRIPTION written by whoever authored the event.
 *
 * With `/m` anchors that turns any TEXT value into a structure-editing primitive, and the
 * damage is not limited to fabricating a row:
 *
 *   - `DESCRIPTION:hi<U+2028>END:VEVENT<U+2028>BEGIN:VEVENT<U+2028>…` cuts ONE component in two. The
 *     real DTSTART/DTEND land in the discarded tail, so the event comes back with no dates at
 *     all — and `eventIntersectsWindow` keeps a dateless event, so it displays inside a window
 *     it was never shown to be in.
 *   - On the expanded path the same payload yields two events, one wholly attacker-authored
 *     with an attacker-chosen start, and `blockCountProvesSeries` then marks the REAL event
 *     recurring because "two blocks prove a series".
 *   - Worst: a property read takes the FIRST match in the block, so a SUMMARY of
 *     `Lunch<U+2028>UID:board-meeting@victim.example<U+2028>DTSTART:…` makes the attacker's own
 *     resource report ANOTHER resource's UID. `delete_calendar_event`/`update_calendar_event`
 *     resolve an id through `findCalendarObjectByUID`, which returns the first object in the
 *     account whose UID matches — so an agent told "the id names the series, act on it"
 *     irreversibly destroys a record the caller never named, and mails its attendees a
 *     cancellation.
 *
 * Splitting on RFC line breaks and matching WHOLE lines makes U+2028, U+2029 and a bare CR
 * inert characters inside a value, which is what the RFC says they are. Do not "simplify" any
 * of this back to a `/m` regex.
 *
 * Unfolding stays per RFC and stays LATER than the component split (see `parseICalValue`): a
 * continuation line is one beginning with a space or a tab, so a marker at the head of a
 * continuation is text, not structure.
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
 * Whether a line is a FOLDED CONTINUATION of the line above it (RFC 5545 §3.1).
 *
 * Every structural scan skips these, so a `BEGIN:VEVENT` or a `UID:` at the head of a
 * continuation is read as the text it is. This is the fold-injection half of the same rule
 * the line model above closes for U+2028: libical folds at a fixed octet count, so padding a
 * description to put a chosen fold in a chosen place is deterministic to construct.
 */
function isFoldedContinuation(line: string): boolean {
  return line.startsWith(' ') || line.startsWith('\t');
}

/**
 * Whether an iCalendar block contains the named property as a whole content line.
 *
 * The one place a "does this component have an RRULE / a RECURRENCE-ID / an ATTENDEE?"
 * question is answered. Every such test used to be its own `/^KEY[;:]/m` literal, which is
 * exactly the shape described above — and the RRULE gate now decides whether
 * `update_calendar_event` / `delete_calendar_event` refuse the call outright, while the
 * ORGANIZER/ATTENDEE gates steer an in-place patch, so a forged one mis-routes a write.
 *
 * CASE-INSENSITIVE, because RFC 5545 §3.1 says property names are. libical re-serialises them
 * in upper case, so `rrule:FREQ=WEEKLY` does not arrive from Fastmail's own client — but it is
 * legal iCalendar, a third party can PUT it, and this test now fronts an irreversible write:
 * missing a real rule lets `delete_calendar_event` destroy a series it was meant to refuse.
 *
 * UNFOLDS FIRST (RFC 5545 §3.1), rather than skipping continuation lines as the structural
 * scans do. The two disagree only on a property NAME split across a fold — `RRU\r\n LE:…` —
 * which skipping reads as no property at all, i.e. fail-OPEN in front of that same destroy.
 * Unfolding closes it without re-opening the injection hole the skip existed for, because a
 * continuation is APPENDED to the line above and so can never begin a logical line: a
 * `DESCRIPTION:x\r\n UID:evil` still yields one line starting `DESCRIPTION:`, not a `UID:`.
 *
 * Exported for its own test, and still worth having one: a caller-level test of the RRULE gate
 * pins the refusal's own reads rather than this presence test, and the two can disagree.
 */
export function hasICalProperty(block: string, key: string): boolean {
  const test = new RegExp(`^${key}[;:]`, 'i');
  return unfoldedICalLines(block).some(line => test.test(line));
}

/**
 * A block's LOGICAL lines: every folded continuation (RFC 5545 §3.1 — a line beginning with a
 * space or tab) joined onto the line above with its single leading whitespace character
 * removed. Text only, with no offsets, so it cannot be used to slice a component back out —
 * `icalContentLines` + `structuralLine` stay the model for anything that edits the payload.
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
 *
 * Trailing whitespace (and a stray CR from a `\r\r\n` payload) is trimmed because it is not
 * meaningful and real servers emit it; LEADING whitespace deliberately is not, because leading
 * whitespace is precisely what makes a line a continuation rather than a marker.
 */
function structuralLine(text: string): string | null {
  if (isFoldedContinuation(text)) return null;
  return text.replace(/[\r\t ]+$/, '');
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
 * Whether a STORED calendar resource holds a repeating series. Read off the RESOURCE, never
 * off a listing row: `list_calendar_events` expands a series into per-occurrence rows whose
 * RRULE the server has stripped, so a row is not evidence either way — the payload behind the
 * id is. This is what `update_calendar_event` and `delete_calendar_event` refuse on.
 *
 * Three markers, ANY of which is enough, because this gates a refusal in front of an
 * irreversible write and so fails CLOSED — the cost of calling a one-off event a series is one
 * edit the caller does in the web client, the cost of the reverse is a destroyed series:
 *
 *  - a VEVENT carrying an RRULE — the ordinary series master;
 *  - more than one VEVENT block in the resource — one CalDAV resource is one UID, so a second
 *    block can only be an overridden occurrence of a recurrence;
 *  - any block carrying a RECURRENCE-ID — that block IS an overridden occurrence, and a
 *    resource made entirely of them is a series whose master has been removed.
 *
 * The last two are the same test the read path uses to set `isRecurring`
 * (`blockCountProvesSeries`), deliberately, so the two halves of the server cannot disagree
 * about what a series is. Note what is NOT required: the RRULE and the overrides do not have
 * to agree, and a malformed resource carrying an override with no rule still refuses.
 *
 * Line-model reads (hasICalProperty / extractVEventBlocks), never a `/m`-anchored regex over
 * the raw payload. Calendar content is attacker-authored here (anyone who can send an
 * invitation writes a DESCRIPTION), and U+2028, U+2029 and a bare CR all satisfy JavaScript's
 * `/m` anchors while being legal, unescaped characters inside an iCalendar TEXT value — so a
 * regex would let a forged text field turn a one-off event into a refusal, or, read the other
 * way, be defeated by folding a real rule across a continuation line. See docs/conventions.md.
 */
export function isRecurringSeriesResource(icalData: string | null | undefined): boolean {
  const blocks = extractVEventBlocks(icalData || '');
  if (blocks.length === 0) return false;
  if (blocks.length > 1) return true;
  return blocks.some((block) => hasICalProperty(block, 'RRULE') || hasRecurrenceId(block));
}

/**
 * The one refusal `update_calendar_event` and `delete_calendar_event` raise on a repeating
 * event, so the two read as a single rule rather than two similar ones.
 *
 * WHY THIS REFUSES INSTEAD OF ASKING FOR CONFIRMATION. `create_calendar_event` has no RRULE
 * parameter — this server cannot make a repeating event at all — so under the project rule
 * that a destroy must not remove what this server cannot recreate, it must not destroy or
 * rewrite one either. Concretely: a delete removes the whole resource, every occurrence past
 * and future, and the server then mails a cancellation to every attendee; an update patches
 * the master and moves every occurrence, and where the series carries RECURRENCE-ID overrides
 * RFC 5545 does not settle whether those follow the master or stay put, so there is no correct
 * answer to implement. Recurrence expansion (#64) made this reachable in a way it was not
 * before — a series used to appear once at its original start date, and now returns a
 * plausible per-occurrence row carrying the series id, so "delete Thursday's meeting" finds
 * something to call.
 *
 * Restoring per-occurrence and whole-series editing (and the RRULE parameter on create that
 * would make the two surfaces match) is tracked as issue #146; #109 holds the deeper design
 * discussion. Until that lands there is deliberately NO override parameter, and the message
 * says so outright, because an LLM caller that is merely told "no" will otherwise spend turns
 * hunting for the flag that makes it work.
 */
/**
 * The event's SUMMARY for an error message, falling back to the id the caller passed when the
 * resource has none. Unescaped so the caller reads the title they see in the calendar, and
 * read off the MASTER block (the first one) so an override's re-titled occurrence cannot name
 * the series something the caller would not recognise.
 */
function calendarObjectTitle(icalData: string | null | undefined, eventId: string): string {
  const blocks = extractVEventBlocks(icalData || '');
  // The master is the block with no RECURRENCE-ID; RFC 5545 does not fix component order, and
  // an all-overrides resource has no master at all, so fall through to the first block.
  const vevent = blocks.find((b) => !hasICalProperty(b, 'RECURRENCE-ID')) ?? blocks[0];
  const summary = vevent ? parseICalValue(vevent, 'SUMMARY') : undefined;
  return summary ? unescapeICalText(summary) : eventId;
}

export function recurringSeriesRefusal(
  action: 'update' | 'delete',
  title: string,
): InvalidInputError {
  const consequence = action === 'delete'
    ? 'Deleting it would remove every occurrence, past and future, and the server would mail a cancellation to every attendee.'
    : 'Changing it would move every occurrence, and where single occurrences have already been edited on their own there is no agreed answer for what should happen to them.';
  return new InvalidInputError(
    `"${title}" is a repeating event, and this server will not ${action} it. `
    + `${consequence} `
    + 'There is no parameter, flag or confirmation that overrides this, so do not look for one: '
    + 'this server cannot CREATE a repeating event (create_calendar_event writes single events only), '
    + 'so it will not destroy or rewrite one it has no way to put back. '
    + 'Use the Fastmail web interface to change or delete a repeating event, including a single occurrence of one. '
    + 'get_calendar_event still works on this id and is read-only.',
  );
}

/**
 * Find the index of the first colon that separates iCal property parameters
 * from the property value. Colons inside quoted parameter values (e.g.
 * DELEGATED-FROM="mailto:boss@example.com") are skipped.
 * Also correctly handles properties like DESCRIPTION;ALTREP="http://...":text
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
 * Parse an iCalendar property value from within a VEVENT block.
 * Handles simple (KEY:value), parameterized (KEY;TZID=...:value),
 * and VALUE=DATE (KEY;VALUE=DATE:20260319) forms.
 * Also handles line folding (continuation lines starting with space/tab).
 *
 * CASE-SENSITIVE on the property name, unlike `hasICalProperty`, and left that way
 * deliberately. RFC 5545 §3.1 says names are case-insensitive, so a lower-cased `uid:` or
 * `dtstart:` here reads as absent — but every consequence of that is fail-CLOSED, because the
 * structural scan above it is case-sensitive too: `extractVEventBlocks` matches the literal
 * `BEGIN:VEVENT`, so a payload whose keywords are lower-cased yields no blocks at all,
 * `findCalendarObjectByUID` skips the object before it ever reads a value, and the event is
 * simply invisible to every tool rather than editable or destroyable through a mis-read.
 * `hasICalProperty` is the exception because it is the one read that gates a destroy while the
 * surrounding payload IS well-formed — a normal resource with one lower-cased `rrule:` line —
 * so there, missing the property fails open. Making every read case-insensitive is a wider
 * change than that gate needs, and belongs with the RFC conformance audit (#57, #111).
 */
export function parseICalValue(vevent: string, key: string): string | undefined {
  // Whole content lines, split on RFC 5545 line breaks only — never a `/m` regex over the
  // blob. This is the read that a forged UID/DTSTART inside a SUMMARY used to hijack, and it
  // is the one whose answer decides which stored record a destroy resolves to. See the
  // line-model comment above.
  const lines = icalContentLines(vevent).map(l => l.text);
  const test = new RegExp(`^${key}[;:]`);

  for (let i = 0; i < lines.length; i++) {
    if (isFoldedContinuation(lines[i])) continue;
    const line = lines[i].replace(/\r$/, '');
    if (!test.test(line)) continue;

    // Unfold: every following continuation line contributes its text minus the one leading
    // space or tab that marked it as a continuation.
    let fullLine = line;
    for (let j = i + 1; j < lines.length; j++) {
      if (!isFoldedContinuation(lines[j])) break;
      fullLine += lines[j].substring(1);
    }

    // Use quote-aware colon detection for the parameter/value boundary
    const colonIdx = findValueBoundary(fullLine);
    if (colonIdx === -1) return undefined;
    return fullLine.substring(colonIdx + 1).trim();
  }

  return undefined;
}

/**
 * Return all occurrences of a property key as full unfolded raw lines.
 * Needed because ATTENDEE/EXDATE etc. can appear multiple times.
 */
export function parseAllICalProperties(vevent: string, key: string): string[] {
  // Same line model as parseICalValue — RFC line breaks only, folded continuations skipped.
  const lines = icalContentLines(vevent).map(l => l.text);
  const regex = new RegExp(`^${key}[;:]`);
  const results: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    if (isFoldedContinuation(lines[i])) continue;
    const line = lines[i].replace(/\r$/, '');
    if (!regex.test(line)) continue;

    // Unfold continuation lines
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
 * Uses quote-aware scanning for parameter/value boundary detection.
 */
export function parseAttendee(rawLine: string): Participant {
  // Find the parameter/value boundary (first colon outside quotes)
  const boundaryIdx = findValueBoundary(rawLine);
  const paramPart = boundaryIdx >= 0 ? rawLine.substring(0, boundaryIdx) : rawLine;
  const valuePart = boundaryIdx >= 0 ? rawLine.substring(boundaryIdx + 1) : '';

  // Extract email from cal-address value
  const email = valuePart.replace(/^mailto:/i, '');

  // Split parameters on semicolons, respecting quotes
  const params: string[] = [];
  let current = '';
  let inQuote = false;
  // Skip the property name (ATTENDEE or ORGANIZER) — start after first ;
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

  // Extract known parameters
  const result: Participant = { email };

  for (const param of params) {
    const eqIdx = param.indexOf('=');
    if (eqIdx === -1) continue;
    const pName = param.substring(0, eqIdx).toUpperCase();
    let pValue = param.substring(eqIdx + 1);
    // Strip surrounding quotes
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

  // All-day date: 20260324 (8 digits)
  if (/^\d{8}$/.test(cleaned)) {
    return `${cleaned.slice(0, 4)}-${cleaned.slice(4, 6)}-${cleaned.slice(6, 8)}`;
  }

  // DateTime: 20260320T083000 or 20260320T083000Z
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
  // Guard: date-only input must be handled by caller, not passed here. This one stays a
  // plain Error on purpose — it reports a broken internal contract between this function
  // and the code calling it, not anything the tool caller supplied or can correct.
  if (/^\d{4}-\d{2}-\d{2}$/.test(isoString)) {
    throw new Error('date-only input must be handled by caller, not passed to toICalUTC');
  }
  // Floating time (no offset, no Z) — preserve as local iCal datetime
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(isoString)) {
    return isoString.replace(/[-:]/g, '');
  }
  const d = new Date(isoString);
  // The value reaching here is the start/end the tool caller passed (via
  // formatDateTimeProperty), so an unparseable one is caller-fixable input.
  if (isNaN(d.getTime())) throw new InvalidInputError(`Invalid date: ${isoString}`);
  return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

/**
 * Fold an iCalendar content line at 75 octets per RFC 5545 §3.1.
 * @param lineEnding Line ending to use for fold breaks (default '\r\n')
 */
export function foldICalLine(line: string, lineEnding: string = '\r\n'): string {
  const parts: string[] = [];
  while (Buffer.byteLength(line, 'utf8') > 75) {
    // Find the largest character count that fits in 75 bytes
    let cut = 75;
    while (cut > 0 && Buffer.byteLength(line.slice(0, cut), 'utf8') > 75) {
      cut--;
    }
    // Don't split a surrogate pair (characters outside BMP like emoji)
    if (cut > 0 && cut < line.length) {
      const code = line.charCodeAt(cut);
      if (code >= 0xDC00 && code <= 0xDFFF) cut--;
    }
    parts.push(line.slice(0, cut));
    line = ' ' + line.slice(cut);
  }
  parts.push(line);
  return parts.join(lineEnding);
}

/**
 * Detect line ending style from iCal data.
 */
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

  // structuralLine, not `.trim()`: a trimmed compare treats the head of a FOLDED
  // continuation (` BEGIN:VEVENT`) as a component marker, which is the write-path twin of
  // the read-path forgery the line model above closes.
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
    // A folded continuation is text, never a marker: it belongs to the property above it,
    // and the property scan below has already consumed (or will skip) that property whole.
    if (trimmed === null) continue;
    if (trimmed.startsWith('BEGIN:')) { nestDepth++; continue; }
    if (trimmed.startsWith('END:')) { nestDepth--; continue; }
    if (nestDepth > 0) continue;

    if (propRegex.test(lines[i])) {
      foundIdx = i;
      // Find end of this property (including continuation lines)
      foundEndIdx = i + 1;
      while (foundEndIdx < veventEnd && (lines[foundEndIdx].startsWith(' ') || lines[foundEndIdx].startsWith('\t'))) {
        foundEndIdx++;
      }
      break;
    }
  }

  if (foundIdx >= 0) {
    // Replace or remove existing property
    const newLines = newLine !== null ? newLine.split(/\r?\n/) : [];
    lines.splice(foundIdx, foundEndIdx - foundIdx, ...newLines);
  } else if (newLine !== null) {
    // Insert before the first sub-component (e.g. VALARM) when present —
    // RFC 5545 ABNF is `eventprop *alarmc`, so properties must precede alarms.
    let insertAt = veventEnd;
    for (let i = veventStart + 1; i < veventEnd; i++) {
      // structuralLine, like every other scan in this function. A trimmed compare read the head
      // of a FOLDED continuation (` BEGIN:phase two of the agenda`) as a sub-component and
      // spliced the new property INTO the middle of the property above it — the tail of that
      // property was cut loose and swallowed by the inserted line, damaging two records in one
      // write with nothing reported.
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

  // structuralLine, not `.trim()`: a trimmed compare treats the head of a FOLDED
  // continuation (` BEGIN:VEVENT`) as a component marker, which is the write-path twin of
  // the read-path forgery the line model above closes.
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
  // Collect indices to remove (in reverse order to avoid index shifting)
  const toRemove: Array<[number, number]> = [];
  let nestDepth = 0;

  for (let i = veventStart + 1; i < veventEnd; i++) {
    const trimmed = structuralLine(lines[i]);
    // A folded continuation is text, never a marker: it belongs to the property above it,
    // and the property scan below has already consumed (or will skip) that property whole.
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

  // structuralLine, not `.trim()`: a trimmed compare treats the head of a FOLDED
  // continuation (` BEGIN:VEVENT`) as a component marker, which is the write-path twin of
  // the read-path forgery the line model above closes.
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

  // Insert before first sub-component (VALARM etc.) or before END:VEVENT
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

  // Find all VTIMEZONE blocks and their TZIDs
  const tzBlocks: Array<{ tzid: string; start: number; end: number }> = [];
  for (let i = 0; i < lines.length; i++) {
    if (structuralLine(lines[i]) === 'BEGIN:VTIMEZONE') {
      const blockStart = i;
      let blockEnd = -1;
      for (let j = i + 1; j < lines.length; j++) {
        if (structuralLine(lines[j]) === 'END:VTIMEZONE') {
          blockEnd = j;
          break;
        }
      }
      if (blockEnd === -1) { i = lines.length; break; }
      // Use parseICalValue for proper unfolding support
      const tzBlock = lines.slice(blockStart, blockEnd + 1).join('\n');
      const tzid = parseICalValue(tzBlock, 'TZID') || '';
      tzBlocks.push({ tzid, start: blockStart, end: blockEnd });
      i = blockEnd;
    }
  }

  if (tzBlocks.length === 0) return icalData;

  // Build content outside VTIMEZONE blocks for reference scanning
  const nonTzLines: string[] = [];
  let inTz = false;
  for (let i = 0; i < lines.length; i++) {
    if (structuralLine(lines[i]) === 'BEGIN:VTIMEZONE') { inTz = true; continue; }
    if (structuralLine(lines[i]) === 'END:VTIMEZONE') { inTz = false; continue; }
    if (!inTz) nonTzLines.push(lines[i]);
  }
  // Unfold before scanning so a reference split across a folded line isn't
  // missed, and check both bare and quoted parameter forms.
  const nonTzContent = nonTzLines.join('\n').replace(/\n[ \t]/g, '');

  // Check each VTIMEZONE for references
  const orphaned = tzBlocks.filter(tz => {
    if (!tz.tzid) return false;
    return !nonTzContent.includes(`;TZID=${tz.tzid}`) &&
           !nonTzContent.includes(`;TZID="${tz.tzid}"`);
  });

  // Remove orphaned blocks in reverse order
  for (let i = orphaned.length - 1; i >= 0; i--) {
    lines.splice(orphaned[i].start, orphaned[i].end - orphaned[i].start + 1);
  }

  return lines.join(lineEnding);
}

/**
 * Remove exception VEVENT blocks whose RECURRENCE-ID matches one of the orphaned dates.
 * Operates on the full iCal string. Never touches the master VEVENT (no RECURRENCE-ID).
 *
 * NO PRODUCTION CALLER at present, and kept deliberately rather than deleted:
 * `update_calendar_event` refuses a repeating event outright now (see recurringSeriesRefusal),
 * which made the orphan-pruning branch that used this unreachable. This is the primitive that
 * a series-aware update needs back — see #146, which names it — and it is pure and covered by
 * its own tests, so keeping it costs nothing and re-deriving it would.
 */
export function removeExceptionVEvents(icalData: string, orphanedRecurrenceIds: Date[]): string {
  if (orphanedRecurrenceIds.length === 0) return icalData;

  const lineEnding = detectLineEnding(icalData);
  const lines = icalData.split(/\r?\n/);

  // Find all VEVENT blocks
  const veventBlocks: Array<{ start: number; end: number; recurrenceId?: string }> = [];
  for (let i = 0; i < lines.length; i++) {
    if (structuralLine(lines[i]) === 'BEGIN:VEVENT') {
      const blockStart = i;
      for (let j = i + 1; j < lines.length; j++) {
        if (structuralLine(lines[j]) === 'END:VEVENT') {
          // Extract RECURRENCE-ID using parseICalValue for consistency
          // with the orphan detection code path (handles unfolding)
          const veventText = lines.slice(blockStart, j + 1).join('\n');
          const recId = parseICalValue(veventText, 'RECURRENCE-ID');
          veventBlocks.push({ start: blockStart, end: j, recurrenceId: recId });
          i = j;
          break;
        }
      }
    }
  }

  // Only remove exception VEVENTs (those with RECURRENCE-ID) that are orphaned.
  // Compare on ISO date strings (YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS) to avoid
  // timezone interpretation issues (floating vs UTC) with millisecond comparison.
  const orphanedDateStrings = orphanedRecurrenceIds.map(d => {
    // Normalize to ISO date string for comparison
    return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
  });
  const toRemove = veventBlocks.filter(block => {
    if (!block.recurrenceId) return false; // master VEVENT — never remove
    const recIdFormatted = formatICalDate(block.recurrenceId);
    if (!recIdFormatted) return false;
    // Compare in a fixed UTC frame — naive datetimes must not be interpreted
    // in the process's local timezone (must match orphan-detection's frame).
    const recDate = parseICalDateAsUTC(recIdFormatted);
    const recDateStr = recDate.toISOString().replace(/\.\d{3}Z$/, 'Z');
    return orphanedDateStrings.includes(recDateStr);
  });

  // Remove in reverse order
  for (let i = toRemove.length - 1; i >= 0; i--) {
    lines.splice(toRemove[i].start, toRemove[i].end - toRemove[i].start + 1);
  }

  return lines.join(lineEnding);
}

/**
 * Parse an iCalendar DURATION value and compute end datetime.
 * RFC 5545 §3.3.6: [+/-]P[nW | nDTnHnMnS]
 * Returns ISO 8601 end datetime, or undefined for malformed input.
 */
export function parseICalDuration(duration: string, start: string): string | undefined {
  const m = duration.match(/^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
  if (!m) return undefined;

  const [, sign, weeks, days, hours, minutes, seconds] = m;

  // At least one component must be present (reject bare "P")
  if (!weeks && !days && !hours && !minutes && !seconds) return undefined;
  // If T is present in input, at least one time component must exist (reject "P1DT")
  if (duration.includes('T') && !hours && !minutes && !seconds) return undefined;

  const ms =
    (parseInt(weeks || '0', 10) * 7 * 86400000) +
    (parseInt(days || '0', 10) * 86400000) +
    (parseInt(hours || '0', 10) * 3600000) +
    (parseInt(minutes || '0', 10) * 60000) +
    (parseInt(seconds || '0', 10) * 1000);

  const startDate = new Date(start);
  if (isNaN(startDate.getTime())) return undefined;

  const endMs = sign === '-' ? startDate.getTime() - ms : startDate.getTime() + ms;
  const endDate = new Date(endMs);

  // Return in same format as input start
  if (/^\d{4}-\d{2}-\d{2}$/.test(start)) {
    // Date-only: return date-only
    return endDate.toISOString().slice(0, 10);
  }

  // Floating time (no Z, no offset): return floating to match start format.
  // new Date() interprets floating as local, so we add the duration in ms
  // and format back as floating by doing manual arithmetic instead of toISOString().
  const isFloating = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(start);
  if (isFloating) {
    // Parse start components directly to avoid local-time interpretation
    const [datePart, timePart] = start.split('T');
    const [y, mo, d] = datePart.split('-').map(Number);
    const [h, mi, s] = timePart.split(':').map(Number);
    const utcStart = Date.UTC(y, mo - 1, d, h, mi, s);
    const utcEnd = sign === '-' ? utcStart - ms : utcStart + ms;
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
 * One resource is not one event, and which events it holds depends on how it was asked
 * for — so the shape of the payload has to be decided before anything is read out of it:
 *
 *   EXPANDED (a time-range query sent with `expand`): the server applies the recurrence
 *     itself and returns one VEVENT per in-window occurrence, RRULE stripped. Several
 *     occurrences arrive inside ONE calendar-data blob — measured against Fastmail at 3, 6
 *     and 7 VEVENTs in a single blob (scripts/probes/calendar-expand.probe.mjs) — so
 *     reading only the first would silently drop six events out of seven. Every block
 *     becomes its own CalendarEvent.
 *
 *   UNEXPANDED (no time range, or a plain query): the resource holds the series MASTER,
 *     which carries the RRULE, optionally followed by exception blocks that override
 *     individual instances. Only the master is emitted, which is the long-standing
 *     one-event-per-resource behaviour.
 *
 *   NO VEVENT: the minimal-event fallback, unchanged.
 *
 * WHICH SHAPE IT IS, IS TOLD — NEVER INFERRED FROM CONTENT. `expanded` is set by the caller
 * that decided to send `expand`, because the payload cannot be read backwards to recover
 * that decision. The tempting discriminator — "no block lacks a RECURRENCE-ID, therefore
 * the server expanded this" — is FALSE against the server Fastmail runs. Cyrus's expansion
 * (`expand_cb`, imap/http_caldav.c) sets a RECURRENCE-ID only on instances AFTER the
 * series' first: the first instance is emitted with its RRULE stripped and NO
 * RECURRENCE-ID. So any window containing a series' original DTSTART returns
 * [first-instance, occurrence, occurrence, …], the content sniff picks the first block as
 * a "master", and every sibling is discarded with nothing said — a yearly series over a
 * five-year window reported ONE event, and a 2020-2021 window lost 27 of 102 occurrences.
 *
 * The master is selected by looking for the block WITHOUT a RECURRENCE-ID rather than by
 * taking the first block. RFC 5545 does not fix component order, so a resource authored by
 * another client can serialise an override ahead of its master — and the previous
 * first-match read reported that override as though it were the event. `normalizeMasterVEventFirst`
 * encodes the same rule for the write path by reordering the payload; this decides it on
 * the block list instead, because the read path has no reason to re-serialise anything.
 * The two are deliberately left as separate implementations: unifying them means making the
 * write path consume a parsed model it does not have, which is the read/write divergence
 * tracked in #102 rather than something to settle inside a parser.
 */
export function parseCalendarObjects(
  obj: DAVCalendarObject,
  options?: { includeParticipants?: boolean; expanded?: boolean },
): CalendarEvent[] {
  const blocks = extractAllVEvents(obj.data || '');
  if (blocks.length === 0) {
    // No VEVENT found — return minimal event
    return [{
      id: obj.url || '',
      url: obj.url || '',
      title: 'Untitled',
    }];
  }

  if (options?.expanded) {
    // Every block is an in-window occurrence the server generated. Nothing is selected and
    // nothing is dropped: the server already decided what falls in the window.
    const events = blocks.map(b => parseVEvent(obj, b, options));
    if (blockCountProvesSeries(blocks)) {
      // A resource that yielded MORE THAN ONE in-window instance is a repeating series by
      // construction, so its first instance is recurring too even though the server stripped
      // the RRULE and left it without a RECURRENCE-ID. Without this the row that survives at
      // the head of an expanded series reports as a one-off, which is the same lie the
      // dropped-siblings bug told, one row smaller.
      for (const event of events) event.isRecurring = true;
    }
    return events;
  }

  const master = blocks.find(b => !hasRecurrenceId(b));

  // No master at all means every block is a detached override; emitting each of them is
  // right, since each names a distinct instance.
  if (!master) return blocks.map(b => parseVEvent(obj, b, options));

  return [parseVEvent(obj, master, options)];
}

function hasRecurrenceId(block: string): boolean {
  return hasICalProperty(block, 'RECURRENCE-ID');
}

/**
 * Whether an EXPANDED blob's block list proves the resource is a repeating series.
 *
 * Two blocks or more can only come from a recurrence, and a RECURRENCE-ID on any block says
 * the same thing. What is NOT decidable here is the converse: a single block carrying
 * neither marker is a one-off event AND the sole in-window instance of a series that starts
 * inside the window, because Cyrus emits both identically (see above). That residual
 * ambiguity is deliberately left unresolved rather than guessed at in either direction —
 * claiming `isRecurring` on every expanded row would mark genuine one-off events as
 * repeating, and the alternative is a second, unexpanded fetch. That cost is one fetch per
 * AMBIGUOUS resource (a single block carrying neither marker), not per resource, so state it
 * that way rather than inflating it: the reason to leave this alone is that the residue is
 * disclosed and settleable in one call, not that closing it would be expensive. It is
 * documented on the tool instead: `get_calendar_event` on that id returns the master and
 * settles it in one call.
 */
function blockCountProvesSeries(blocks: string[]): boolean {
  return blocks.length > 1 || blocks.some(hasRecurrenceId);
}

/**
 * Parse ONE event out of a resource.
 *
 * Retained as the single-event entry point that most of this file and its callers use.
 * It returns the first event `parseCalendarObjects` produces, which for an unexpanded
 * resource is the series master — the right answer for `get_calendar_event`, which fetches
 * without a time range and so has no occurrence to report.
 */
export function parseCalendarObject(obj: DAVCalendarObject, options?: { includeParticipants?: boolean }): CalendarEvent {
  return parseCalendarObjects(obj, options)[0];
}

function parseVEvent(obj: DAVCalendarObject, vevent: string, options?: { includeParticipants?: boolean }): CalendarEvent {
  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  const rawStart = parseICalValue(vevent, 'DTSTART');
  let rawEnd = parseICalValue(vevent, 'DTEND');
  const location = parseICalValue(vevent, 'LOCATION');
  const uid = parseICalValue(vevent, 'UID') || obj.url || '';

  // DURATION parsing: compute end from start + duration if DTEND absent
  if (!rawEnd && rawStart) {
    const rawDuration = parseICalValue(vevent, 'DURATION');
    if (rawDuration) {
      const startIso = formatICalDate(rawStart);
      if (startIso) {
        const computedEnd = parseICalDuration(rawDuration, startIso);
        if (computedEnd) {
          // computedEnd is already ISO format, return it directly
          const event: CalendarEvent = {
            id: uid,
            url: obj.url || '',
            title: unescapeICalText(title),
            description: description ? unescapeICalText(description) : undefined,
            start: formatICalDate(rawStart),
            end: computedEnd,
            location: location ? unescapeICalText(location) : undefined,
          };
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

  addRecurrenceToEvent(event, vevent);
  if (options?.includeParticipants) {
    addParticipantsToEvent(event, vevent);
  }

  return event;
}

/**
 * Attach the recurrence markers that say what KIND of date `start` is (#64).
 *
 * Read from the block being parsed rather than from the resource, because that is the
 * distinction being reported: an expanded occurrence has had its RRULE stripped by the
 * server and carries a RECURRENCE-ID, a master carries the rule and no RECURRENCE-ID.
 * Both set `isRecurring`; nothing is set for an ordinary one-off event, per the
 * omit-empty-fields convention.
 */
function addRecurrenceToEvent(event: CalendarEvent, vevent: string): void {
  const rrule = parseICalValue(vevent, 'RRULE');
  const recurrenceId = parseICalValue(vevent, 'RECURRENCE-ID');
  if (recurrenceId) {
    event.isRecurring = true;
    event.recurrenceId = formatICalDate(recurrenceId);
  }
  if (rrule) {
    event.isRecurring = true;
    event.recurrenceRule = rrule;
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
 * Unescape an iCalendar text value (RFC 5545 §3.3.11).
 * Reverses escaping of newlines, semicolons, commas, and backslashes.
 *
 * Done in a single left-to-right pass so each escape is decoded exactly once.
 * Chained .replace() calls re-scan the whole string and corrupt an escaped
 * backslash that precedes an escapable char: e.g. "\\n" (an escaped backslash
 * followed by a literal "n") would have its second "\n" turned into a newline,
 * yielding "\<newline>" instead of the correct "\n".
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
 * Backslashes, newlines, commas, and semicolons must be escaped.
 */
export function escapeICalText(value: string): string {
  return value
    // Normalize CRLF and BARE CR to LF first — a lone \r would otherwise pass
    // through untouched and act as a line terminator for downstream parsers,
    // reopening the property-injection class the date paths are guarded against.
    .replace(/\r\n?/g, '\n')
    // Strip remaining control characters (HTAB is legal in iCal TEXT; LF is
    // escaped below).
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
 *   - YYYY-MM-DDTHH:MM:SS+HH:MM        (with offset, normalized to UTC)
 * Rejects any control characters or unexpected content. Returns the ICS-safe
 * serialized form (no `-` or `:`, with `Z` suffix for instants, or `YYYYMMDD`
 * for date-only). Throws on invalid input.
 *
 * This is the ONLY thing that turns a caller-supplied start/end into an iCal value:
 * formatDateTimeProperty wraps the result in the property line and never parses the
 * caller's string itself. That matters because the two jobs have opposite instincts —
 * a serializer reaches for `new Date()` to normalize, and `new Date()`'s legacy
 * fallback parser accepts `2026/04/18` and `April 18 2026` and reads them as
 * HOST-LOCAL midnight, which puts the event on a different day for a caller in a
 * different zone. The anchored shapes above are matched explicitly so nothing reaches
 * that parser, and a real-calendar-date probe rejects an impossible day (`2026-02-31`)
 * that Date would otherwise roll silently into the next month. Same reasoning, and the
 * same two traps, as coerceUtcDate in src/coerce.ts.
 */
export function validateAndFormatICalDate(value: string, fieldName: string): string {
  // Every rejection below names the caller's own field and value, so all of them are
  // caller-fixable input errors (InvalidParams), never server faults.
  if (typeof value !== 'string') {
    throw new InvalidInputError(`${fieldName} must be a string`);
  }
  if (/[\x00-\x1F\x7F]/.test(value)) {
    throw new InvalidInputError(`${fieldName} contains control characters`);
  }
  const trimmed = value.trim();
  // Date-only: 2026-04-18
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    assertRealCalendarDate(trimmed, trimmed, fieldName);
    return trimmed.replace(/-/g, '');
  }
  // Datetime forms: floating, UTC (Z), or with offset (+/-HH:MM, +/-HHMM, +/-HH)
  const dtMatch = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2}:\d{2})(Z|[+-]\d{2}:?\d{0,2})?$/.exec(trimmed);
  if (!dtMatch) {
    throw new InvalidInputError(`${fieldName} must be ISO-8601 date or datetime (got: ${trimmed.slice(0, 60)})`);
  }
  const [, datePart, timePart, tz] = dtMatch;
  // Probe the calendar date on its own rather than the whole value: an offset
  // legitimately moves the UTC date, so the round-trip check below only holds for the
  // bare date part.
  assertRealCalendarDate(datePart, trimmed, fieldName);
  const isoForParse = `${datePart}T${timePart}${tz || ''}`;
  const d = new Date(isoForParse);
  if (Number.isNaN(d.getTime())) {
    throw new InvalidInputError(`${fieldName} is not a valid datetime (got: ${trimmed.slice(0, 60)})`);
  }
  if (!tz) {
    // Floating: emit as-is without zone designator
    return `${datePart.replace(/-/g, '')}T${timePart.replace(/:/g, '')}`;
  }
  // UTC or offset: normalize to UTC instant
  const utc = d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
  return utc;
}

/**
 * Reject a YYYY-MM-DD that names a day its month does not have.
 *
 * `new Date('2026-02-31T00:00:00Z')` parses happily and lands on 3 March, so without
 * this an event asked for on a nonexistent day would be created a few days later and
 * reported as created — a silent move, not an error the caller can see. Round-tripping
 * through toISOString is the check: a rolled-over date no longer starts with the string
 * it was built from.
 *
 * @param echo the caller's whole value, so the message quotes what they wrote rather
 *   than the date fragment this probe happens to look at.
 */
function assertRealCalendarDate(datePart: string, echo: string, fieldName: string): void {
  const probe = new Date(`${datePart}T00:00:00Z`);
  if (Number.isNaN(probe.getTime()) || !probe.toISOString().startsWith(datePart)) {
    throw new InvalidInputError(`${fieldName} is not a real calendar date (got: ${echo.slice(0, 60)})`);
  }
}

/**
 * Validate an email address for use in ATTENDEE lines.
 * Prevents iCal property injection via malicious email values.
 */
export function validateAttendeeEmail(email: string): void {
  // Participant addresses come straight from the tool call, so every rejection here is
  // caller-fixable input (InvalidParams). The one place this function is applied to a
  // value the caller did NOT supply — the configured CalDAV username, embedded in the
  // ORGANIZER line — goes through validateOrganizerUsername below, which restores the
  // plain-Error class for that case.
  if (!email || typeof email !== 'string') {
    throw new InvalidInputError('Participant email is required');
  }
  if (!/^[^@]+@[^@]+$/.test(email)) {
    throw new InvalidInputError(`Invalid participant email: ${email}`);
  }
  if (/[\r\n:;"\\]|\s/.test(email)) {
    throw new InvalidInputError(`Invalid participant email (contains illegal characters): ${email}`);
  }
}

/**
 * Apply the same strict addr-spec check to the configured CalDAV username, which is
 * embedded verbatim in the ORGANIZER line whenever attendees are present.
 *
 * The address rules are identical to a participant's, but the failure is not: this value
 * is server configuration, so re-forming the tool call cannot fix it. Rethrowing as a
 * plain Error keeps it in the InternalError class instead of telling the caller their
 * arguments were wrong. The message is passed through unchanged.
 */
function validateOrganizerUsername(username: string): void {
  try {
    validateAttendeeEmail(username);
  } catch (e) {
    // The shared validator words its message for a participant address. Say whose
    // address this actually is, or an operator with a bad CalDAV username spends the
    // failure hunting a participant who is fine.
    const detail = e instanceof Error ? e.message : String(e);
    throw new Error(`The configured CalDAV username is not usable as an ORGANIZER address. ${detail}`);
  }
}

/**
 * Quote a CN parameter value per RFC 5545 §3.2.
 * Uses DQUOTE quoting (NOT escapeICalText backslash escaping).
 * Literal DQUOTEs in the value are replaced with single quotes since
 * RFC 5545 has no escape mechanism for DQUOTE inside quoted parameter values.
 * RFC 6868 caret encoding (^') exists but is poorly adopted;
 * single-quote replacement matches Python icalendar/Outlook behavior.
 */
export function quoteParamValue(value: string): string {
  // Strip newlines to prevent iCal property injection via CN values
  let cleaned = value.replace(/[\r\n]+/g, ' ');
  // Then strip the rest of the control range, mirroring escapeICalText's strip
  // for TEXT values. Parameter values reach here from model-supplied participant
  // names, so the remaining C0 controls (HTAB excepted — it is WSP, legal in a
  // quoted param value), DEL and the C1 range would otherwise be emitted raw
  // into the ORGANIZER/ATTENDEE lines. The bidi OVERRIDE and ISOLATE characters
  // go with them: they cannot terminate a line, but they reorder the rendered
  // text of every downstream client, so a name can be made to display as a
  // different address than the one it sits beside.
  // U+200E/200F (LRM/RLM) are deliberately NOT stripped. Unlike the overrides
  // they carry no nesting scope and cannot reorder text around themselves, and
  // they occur legitimately in Arabic and Hebrew display names - stripping them
  // would corrupt real participant names to close a far weaker vector.
  cleaned = cleaned.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\u202A-\u202E\u2066-\u2069]/g, '');
  // Replace literal double quotes with single quotes
  cleaned = cleaned.replace(/"/g, "'");
  // Quote if contains comma, semicolon, colon, or if the original had double quotes
  if (/[,;:]/.test(cleaned) || value.includes('"')) {
    return `"${cleaned}"`;
  }
  return cleaned;
}

/**
 * Format a start/end input value into the correct iCal property line.
 * Handles three cases:
 * 1. Date-only (2026-04-01) → DTXXX;VALUE=DATE:20260401
 * 2. Floating time (2026-03-20T09:30:00) → preserve original TZID
 * 3. UTC/offset (2026-03-20T09:30:00Z) → DTXXX:20260320T093000Z
 *
 * Validation and serialization of the caller's string are delegated wholesale to
 * validateAndFormatICalDate; this function only decides which property FORM the
 * result belongs in, and it decides that from the serialized value rather than by
 * re-reading the input. The serialized form is unambiguous — eight digits is an
 * all-day date, a trailing Z is an instant, anything else is a floating wall-clock
 * time — so the two functions cannot disagree about what the caller wrote, which is
 * exactly what went wrong while each of them parsed the input separately.
 */
function formatDateTimeProperty(
  propName: string,
  value: string,
  originalVevent: string | null,
  lineEnding: string
): string {
  const serialized = validateAndFormatICalDate(value, propName);

  // Date-only
  if (/^\d{8}$/.test(serialized)) {
    return foldICalLine(`${propName};VALUE=DATE:${serialized}`, lineEnding);
  }

  // Floating time (no offset, no Z)
  if (!serialized.endsWith('Z')) {
    // Try to preserve original TZID
    if (originalVevent) {
      const rawLines = parseAllICalProperties(originalVevent, propName);
      if (rawLines.length > 0) {
        const tzMatch = rawLines[0].match(/;TZID=("[^"]*"|[^;:]+)/);
        if (tzMatch) {
          return foldICalLine(`${propName};TZID=${tzMatch[1]}:${serialized}`, lineEnding);
        }
      }
      // If propName is DTEND and no TZID found (DURATION-based), fall back to DTSTART's TZID
      if (propName === 'DTEND') {
        const startLines = parseAllICalProperties(originalVevent, 'DTSTART');
        if (startLines.length > 0) {
          const tzMatch = startLines[0].match(/;TZID=("[^"]*"|[^;:]+)/);
          if (tzMatch) {
            return foldICalLine(`${propName};TZID=${tzMatch[1]}:${serialized}`, lineEnding);
          }
        }
      }
    }
    // No TZID to preserve — emit as floating
    return foldICalLine(`${propName}:${serialized}`, lineEnding);
  }

  // UTC or offset — already normalized to a UTC instant
  return foldICalLine(`${propName}:${serialized}`, lineEnding);
}

/**
 * Check if a raw iCal property line represents a date-only value (VALUE=DATE).
 */
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
  /**
   * Serialized iCal value (20260320 / 20260320T093000 / 20260320T093000Z).
   * Fixed-width within a frame, so lexical order is chronological order.
   */
  value: string;
  /** Human-readable rendering, for error messages. */
  display: string;
}

/**
 * Classify a serialized DTSTART/DTEND property line into its time frame.
 *
 * This deliberately runs on the line that will actually be WRITTEN rather than
 * on the caller's raw input, and that is what keeps `formatDateTimeProperty`'s
 * "a floating input preserves the stored TZID" behaviour intact. A floating
 * value aimed at a TZID-bearing event has already been rewritten to carry that
 * TZID by the time it arrives here, so it classifies as `zoned` and agrees with
 * its zoned partner — exactly as before. A floating value aimed at a UTC (or
 * floating) event has no TZID to inherit, stays floating, and is then correctly
 * seen as a different frame from a UTC partner instead of silently converting
 * one half of the event to a wall-clock time.
 *
 * One classifier serves both sides on purpose: an untouched property is read
 * from the stored VEVENT and a changed one from the freshly formatted line, and
 * they have to be judged by identical rules for the comparison to mean anything.
 *
 * @param displayOverride the caller's own input, when the line was built from
 *   it — error messages should echo what the caller wrote, not our rendering.
 */
function describeDateProperty(rawLine: string, displayOverride?: string): DatePropertyFrame {
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
  const tzMatch = params.match(/;TZID=("[^"]*"|[^;:]+)/);
  if (tzMatch) {
    return { frame: 'zoned', tzid: tzMatch[1].replace(/^"|"$/g, ''), value, display };
  }
  if (/Z$/.test(value)) {
    return { frame: 'utc', value, display };
  }
  return { frame: 'floating', value, display };
}

function describeFrame(d: DatePropertyFrame): string {
  switch (d.frame) {
    case 'date': return 'a date-only (all-day) value';
    case 'floating': return 'a date-time with no time zone';
    case 'utc': return 'a UTC date-time';
    case 'zoned': return `a date-time in time zone ${d.tzid}`;
  }
}

/**
 * Validate the DTSTART/DTEND pair that is about to be written: they must be in
 * the same time frame (RFC 5545 §3.6.1 value-type agreement) and in the right
 * order (RFC 5545 §3.8.2.2, "DTEND MUST be later than DTSTART").
 *
 * The two checks are one check in sequence, not two independent ones: values in
 * different frames are not comparable at all, so ordering can only be judged
 * after the frames agree. That is also why the frame check exists — a mixed
 * pair such as `DTSTART:20260320T093000` (floating) beside
 * `DTEND:20260320T093000Z` (UTC) has no single duration; it renders as a
 * different length for every reader, and in some zones ends before it starts.
 */
function validateDateConsistency(start: DatePropertyFrame, end: DatePropertyFrame): void {
  if (start.frame !== end.frame) {
    throw new InvalidInputError(
      `DTSTART and DTEND must use the same date/time form per RFC 5545 §3.6.1 — ` +
      `start '${start.display}' is ${describeFrame(start)} but end '${end.display}' is ${describeFrame(end)}. ` +
      `Pass start and end in the same form: both date-only (2026-03-20), both with a zone designator ` +
      `(2026-03-20T09:30:00Z or 2026-03-20T09:30:00+10:00), or both without one (2026-03-20T09:30:00).`
    );
  }

  // Two TZID-bearing values in DIFFERENT zones are a legal RFC 5545 shape — a
  // flight that departs in one zone and lands in another — so the frames agree
  // and the event is written. Ordering is skipped there and only there, and as
  // of the window work that is a CHOICE, not an inability: both zone names are
  // present on the write path and ICU can resolve them (`zoneOffsetMsAt` does
  // exactly that for the read window). Checking the order here would newly
  // reject input this tool accepts today, which is a behaviour change rather
  // than a fix, so it is left alone and tracked in #140.
  if (start.frame === 'zoned' && start.tzid !== end.tzid) return;

  if (start.value < end.value) return;

  if (start.frame === 'date') {
    const startDate = formatICalDate(start.value) ?? start.value;
    throw new InvalidInputError(
      `DTEND is exclusive per RFC 5545 — for a one-day event on ${startDate}, ` +
      `pass end: '${nextDay(startDate)}'`
    );
  }
  throw new InvalidInputError(
    `DTEND must be later than DTSTART per RFC 5545 §3.8.2.2 — start '${start.display}' ` +
    `is not before end '${end.display}'. Pass an end later than the start.`
  );
}

function nextDay(dateStr: string): string {
  const d = new Date(dateStr);
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

/**
 * Parse an ISO-ish date/datetime string in a fixed UTC frame.
 *
 * A naive datetime ("2026-03-20T09:30:00") is read as UTC rather than in the process's local
 * timezone, which is what `new Date(...)` would do. The FIXED frame is the point: its callers
 * are `eventIntersectsWindow`'s comparisons, which judge an event's start and end against a
 * window, and a value whose zone the parser already dropped (see docs/conventions.md on
 * zone-less reads) would otherwise be placed differently depending on where the server runs —
 * so the same query would include or exclude the same event by deployment.
 *
 * It was originally written for orphaned-exception detection, which needed to agree with
 * rrule's naive-as-UTC convention. That code is gone — recurring events are refused on the
 * write path now (see recurringSeriesRefusal) — but the window comparisons need the same
 * fixed frame for their own reason, so the function stays.
 */
export function parseICalDateAsUTC(iso: string): Date {
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return new Date(iso + 'T00:00:00Z');
  if (/Z$|[+-]\d{2}:?\d{2}$/.test(iso)) return new Date(iso);
  return new Date(iso + 'Z');
}

// The widest UTC offset any IANA zone has ever used (+14:00 for Kiritimati; the negative
// extreme is smaller, so one bound covers both directions).
//
// It is needed because `formatICalDate` DROPS the TZID parameter: a value that arrived as
// `DTSTART;TZID=Australia/Sydney:20270305T083000` becomes the bare `2027-03-05T08:30:00`,
// which is indistinguishable from a floating time and, read as UTC, can be up to fourteen
// hours away from the instant it names.
//
// WHAT IS AND IS NOT MISSING. ICU can resolve an IANA name — `resolveCalendarInstantMs` does
// exactly that, and the sort above relies on it. What is gone by this point is the NAME: the
// parser discarded it, so there is nothing left to resolve and this filter has only the
// configured zone to guess with. Tightening the slack to that guess would narrow the filter
// on a guess about somebody else's event, so the widest-offset bound stays. Carrying the TZID
// through to the parsed event is the fix that would make a tighter bound honest, and it is a
// response-shape change tracked in #139 rather than something to fold in here.
const MAX_UTC_OFFSET_MS = 14 * 60 * 60 * 1000;

/**
 * Whether a parsed event could fall inside the requested window.
 *
 * This is a RESIDUE filter, not a reimplementation of RFC 4791 time-range matching. The
 * server does the authoritative, timezone-correct matching; this exists because the server
 * matches per OCCURRENCE but returns whole RESOURCES, so without expansion a series whose
 * occurrence lands in the window comes back showing a master DTSTART years earlier (#64 —
 * a ten-day window in March 2027 returning an event dated August 2020).
 *
 * IT DOES NOT CLOSE THE "SERVER DECLINED TO EXPAND" GAP, and an earlier version of this
 * comment claimed it did. On an expanded query a surviving master carries its RRULE and its
 * original DTSTART, and judging that date drops the row entirely — turning a wrongly-dated
 * event into a MISSING one, which is the direction this filter exists not to fail in. The
 * caller keeps such a block regardless of its dates (see the `recurrenceRule` guard in
 * `getCalendarEvents`); this function is not the thing that saves it.
 *
 * Because it is a safety net rather than the authority, it is built to DROP ONLY WHAT
 * PROVABLY CANNOT INTERSECT. Any value with no zone designator gets MAX_UTC_OFFSET_MS of
 * slack on both sides, so a zoned event near a window edge is kept even though its real
 * instant is unknown here. That deliberately leaves a little genuine residue — a zoned
 * event just outside the window survives this filter — and that is the right way to be
 * wrong: the alternative is discarding events the server correctly matched, which is the
 * "you are free when you are not" failure this whole issue is about.
 *
 * THE RESIDUE IS NOT ONE-SIDED AND IS NOT BOUNDED AT A DAY. The slack is granted on BOTH
 * edges, to ANY designator-less value, so the row that survives can sit outside EITHER end;
 * which end you actually see follows the sign of the account's UTC offset, because that
 * decides which neighbouring day the server itself matched. Measured: a Sydney account sees
 * the all-day event on the day BEFORE the window, a New York account sees the one on the day
 * AFTER, and both keep a designator-less TIMED value on the neighbouring day too. And 14
 * hours of slack applied to an all-day `end` that is already the following midnight reaches
 * nearly two calendar days, not one.
 */
export function eventIntersectsWindow(
  event: Pick<CalendarEvent, 'start' | 'end'>,
  windowStartMs: number,
  windowEndMs: number,
): boolean {
  // Nothing to judge: keep it rather than guess. A promised event vanishing with no trace
  // is worse than one extra row the caller can see and dismiss.
  const anchor = event.start || event.end;
  if (!anchor) return true;

  const startMs = parseICalDateAsUTC(anchor).getTime();
  if (Number.isNaN(startMs)) return true;

  const parsedEnd = event.end ? parseICalDateAsUTC(event.end).getTime() : NaN;
  // A missing or unreadable end makes the event a point in time, not an unbounded one.
  const endMs = Number.isNaN(parsedEnd) ? startMs : Math.max(parsedEnd, startMs);

  const zoneless = (v?: string) => !!v && !/Z$|[+-]\d{2}:?\d{2}$/.test(v);
  const slack = zoneless(event.start) || zoneless(event.end) ? MAX_UTC_OFFSET_MS : 0;

  const lo = startMs - slack;
  const hi = endMs + slack;

  // A zero-width instant (a zone-designated event with no duration) has no interval to
  // overlap, so it counts as inside when the window contains it. `windowEnd` is exclusive,
  // matching CalDAV.
  if (lo === hi) return lo >= windowStartMs && lo < windowEndMs;
  return lo < windowEndMs && hi > windowStartMs;
}

/**
 * Order events by the INSTANT their start names, ascending, in place.
 *
 * Not a string sort, and the difference is not cosmetic. `formatICalDate` drops the TZID
 * parameter, so one event's start reaches here as the bare `2026-03-25T08:30:00` (stored
 * `TZID=Australia/Sydney`, which is how Fastmail writes an event a user created) while the
 * next keeps its `Z` (how an external invitation usually arrives), and an all-day event is a
 * bare date. Compared as text those spellings interleave by their digits: on a +10:00 account
 * the zone-less 08:30 is 9.5 hours EARLIER than the 08:00Z beside it and sorted after it.
 * `limit` slices this list, and expansion means the cap is reached far more often, so the
 * wrong order silently cuts a genuinely earlier event and keeps a later one — under a summary
 * line promising the earliest N.
 *
 * Each start is resolved through the same machinery the window bounds use, in the same
 * configured zone, so a zone-less value is placed where the caller's own clock puts it. That
 * is a best-effort reading of a value whose real zone was already lost, not a claim to have
 * recovered it: an event stored in a THIRD zone still sorts by the local reading of its wall
 * clock. It is right for the common case and never worse than comparing the spellings.
 *
 * An unreadable or absent start sorts FIRST. It cannot be placed, and placing it last would
 * put it where `limit` truncates — the one position that turns "cannot order this" into
 * "dropped this".
 */
export function sortEventsByStart(events: CalendarEvent[], zone: string | undefined): void {
  const instants = new Map<CalendarEvent, number>();
  for (const event of events) instants.set(event, resolveCalendarInstantMs(event.start, zone));
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
 * `startDate` alone, or `endDate` alone, is a natural call — the schema requires neither —
 * and the missing half has to be filled in with something. It used to be filled with
 * 1970-01-01 / 2099-12-31, which was harmless while the window was only a filter. It stopped
 * being harmless when the same window became the range the SERVER EXPANDS OVER: `expand`
 * makes Fastmail materialise one VEVENT per occurrence, so `startDate: <today>` on its own
 * asked it to generate every occurrence of every recurring event for the next 73 years —
 * roughly 26,600 VEVENTs for one daily series, and calendar content here is authored by
 * anyone who can send an invitation, including one carrying FREQ=MINUTELY. Cyrus caps the
 * iteration nowhere, and nothing on this side does either: the whole response is buffered,
 * every block copied by a global regex, parsed, filtered and sorted, and only THEN sliced to
 * `limit` — so `limit` is not a bound on any of the work.
 *
 * A year is the span chosen: it answers the question a one-sided window is actually asking
 * ("what is coming up", "what led up to this date") without inventing a range nobody named,
 * and it holds an ordinary daily series to a few hundred rows. 366 rather than 365 so a leap
 * year still covers a full year from any starting day. It is a bound on the INVENTED half
 * only — a caller that names both bounds gets exactly the window it named, because that
 * scope is its own decision.
 *
 * The clamp is never silent: `windowClamp` names the window actually queried and how to ask
 * for more, so "nothing after that date" can't be read as an empty calendar.
 */
export const CALENDAR_OPEN_WINDOW_DAYS = 366;

// The ends of the four-digit-year range every consumer of these bounds can express. Past
// them `toISOString` emits the expanded form (`+010000-12-30T…`), which tsdav rejects with a
// plain Error — surfacing a caller-fixable argument as InternalError ("server-side, a bare
// retry might work"), the exact misclassification the backwards-range check exists to avoid.
const LATEST_REPRESENTABLE_INSTANT = '9999-12-31T23:59:59Z';
const EARLIEST_REPRESENTABLE_INSTANT = '0000-01-01T00:00:00Z';

/**
 * Pull an already-resolved instant back inside the representable range.
 *
 * Applies to a CALLER-NAMED bound as well as to the invented half, which is the half it used
 * to cover alone. A caller bound is not immune: the local-day rule resolves it through a zone,
 * and an offset is enough to push it over on its own — `endDate: "9999-12-31"` on a UTC-5
 * account resolves to `+010000-01-01T05:00:00Z`, and a date at the other end runs off the
 * bottom the same way. Every one of those reached tsdav's `^\d{4}` check as a plain Error, so
 * an argument the caller could have fixed was reported as a server fault.
 *
 * Saturating rather than rejecting, for the same reason the invented half saturates:
 * `9999-12-31` is a perfectly good question and the last representable instant answers it.
 * Never silently, though — a saturated caller bound is named in the window clamp, because
 * unlike the invented half it is a bound the caller DID choose and now is not getting.
 */
function saturateInstant(iso: string): string {
  if (/^\d{4}-/.test(iso)) return iso;
  return iso.startsWith('-') ? EARLIEST_REPRESENTABLE_INSTANT : LATEST_REPRESENTABLE_INSTANT;
}

/**
 * Which end of the representable range a value `saturateInstant` moved was pulled to.
 *
 * Only meaningful for a value that DID move — the caller establishes that by comparing the
 * saturated value against the raw one, which is the same test that decides whether to disclose.
 */
function saturationEdge(saturatedValue: string): 'earliest' | 'latest' {
  return saturatedValue === EARLIEST_REPRESENTABLE_INSTANT ? 'earliest' : 'latest';
}

/**
 * Shift an ISO-8601 UTC instant by whole days, keeping the seconds-precision form, and
 * SATURATING at the ends of the representable range rather than running past them.
 *
 * Saturating rather than rejecting because this only ever fills in a bound the caller did NOT
 * give: `startDate: 9999-12-30` is a perfectly good question, and the invented other half
 * landing on the last representable instant answers it. The clamp note states the range
 * actually searched either way, so the saturation is disclosed like any other clamp.
 *
 * The 24-hour day here is deliberate, and deliberately NOT the local-day arithmetic
 * `coerceCalendarWindowEnd` uses. That one advances a bound the CALLER named, where landing an
 * hour inside or past their day is a wrong answer to a question they asked. This one invents a
 * span nobody named, chosen for being roughly a year; a DST hour either side of an arbitrary
 * 366-day bound is not a wrong answer to anything, and the note names the resulting instant.
 *
 * The example that reaches the saturation is `startDate: "9999-12-30"` ON ITS OWN, where the
 * invented END has nowhere to go. `endDate` alone runs the other way and cannot saturate at
 * that end, so the two directions are not interchangeable in an example.
 */
function shiftIsoDays(iso: string, days: number): string {
  const shifted = new Date(Date.parse(iso) + days * 24 * 60 * 60 * 1000)
    .toISOString()
    .replace(/\.\d{3}Z$/, 'Z');
  return saturateInstant(shifted);
}

// The window arguments quoted back in a rejection go through the SHARED echo in coerce.ts:
// scrubbed of the control characters that would forge extra lines, trimmed so the value shown
// is the value the coercion actually judged, and cut with a visible marker. It is the same
// helper the date rejections themselves use, so one message family cannot end up with two
// policies (#141).

// Calendar display names are server/user data of unbounded length, so a listing of them is
// capped the way every other echoed list in this server is.
const CALENDAR_NAME_LIST_CAP = 20;

/**
 * The display name Fastmail gives the hidden task collection, which `list_calendars` filters
 * out. Named once so every place that hides it hides the same thing — including the
 * not-found error, which must not advertise a calendar the caller cannot see.
 */
const HIDDEN_TASK_CALENDAR_NAME = 'DEFAULT_TASK_CALENDAR_NAME';

/** The calendars a caller is able to name, which is the only set worth listing back at them. */
function selectableCalendars(calendars: DAVCalendar[]): DAVCalendar[] {
  return calendars.filter(c => c.displayName !== HIDDEN_TASK_CALENDAR_NAME);
}

/**
 * The one "no calendar matched that id" error, raised by BOTH the read and the write path.
 *
 * `calendarId` accepts a CalDAV URL or a display name, and the name is matched exactly — so
 * "work" for a calendar called "Work" misses. That made a shared, self-correcting message
 * worth more than a bare echo: the available names are listed, so a caller that guessed the
 * spelling can fix the call without a second `list_calendars` round-trip.
 *
 * The listing is filtered HERE rather than at each call site, because the two call sites
 * passed different lists: the read path had already dropped the hidden task collection and
 * the write path had not, so a mistyped id on a create answered with a calendar name that
 * `list_calendars` never shows and no call can obtain. A shared message is only one rule if
 * it is given one input.
 */
function calendarNotFoundError(calendarId: unknown, available: DAVCalendar[]): InvalidInputError {
  const names = selectableCalendars(available).map(c => String(c.displayName ?? '')).filter(Boolean);
  const shown = names.slice(0, CALENDAR_NAME_LIST_CAP).map(n => `"${echoCallerText(n)}"`).join(', ');
  const more = names.length > CALENDAR_NAME_LIST_CAP ? `, …and ${names.length - CALENDAR_NAME_LIST_CAP} more` : '';
  const listing = names.length > 0 ? ` Available calendars: ${shown}${more}.` : '';
  // QUOTED. The two values this path exists to reject — an empty string and a whitespace-only
  // one — render as nothing at all unquoted, so the message read "Calendar not found: ." with
  // no sign of what had been rejected, while the available names two clauses later were
  // quoted. Same value class, same sentence, so the same treatment.
  return new InvalidInputError(
    `Calendar not found: "${echoCallerText(calendarId)}". calendarId takes either a calendar's URL ` +
    '(its `id` from list_calendars) or its display name, and the name is matched EXACTLY — case and ' +
    `spacing included.${listing}`,
  );
}

/**
 * Reorder VEVENT blocks so the master (no RECURRENCE-ID) comes first.
 * RFC 5545/4791 do not guarantee component ordering — a resource authored by
 * a third-party client may list an overridden instance before the master.
 * All in-place patch helpers target the first VEVENT, so without this
 * normalization an exception-first payload would have its exception patched
 * (and the recurring-event guard skipped) instead of the master.
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
 * Assert a tsdav write (create/update/delete calendar object) actually succeeded.
 * tsdav returns the raw Response(s) without throwing on 4xx/5xx, so without this
 * a server-side rejection would be reported to the caller as success. Accepts a
 * single Response or an array; treats a missing status as success (older tsdav
 * shapes) but fails loudly on any status outside 2xx.
 */
function assertDavOk(resp: unknown, action: string): void {
  const responses = Array.isArray(resp) ? resp : [resp];
  for (const r of responses) {
    const status = (r as any)?.status;
    const ok = (r as any)?.ok;
    if (typeof status === 'number' && (status < 200 || status >= 300)) {
      throw new Error(`Failed to ${action}: server returned ${status}${(r as any)?.statusText ? ' ' + (r as any).statusText : ''}`);
    }
    if (ok === false) {
      throw new Error(`Failed to ${action}: server rejected the request`);
    }
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
      // Every request this client makes carries the CalDAV Basic credential, and
      // the destination is a configured host. A redirect is therefore never
      // legitimate: following one would replay the credential at whatever host the
      // response names. The JMAP side applies the same rule to its own fetches.
      // tsdav merges this into the init of each underlying fetch, so it covers
      // every method rather than the ones called explicitly.
      fetchOptions: { redirect: 'error' },
    });

    // Built into a local and only assigned to `this.client` after login() resolves
    // (#143). The previous code assigned `this.client` before the await, so a
    // rejected login still left an unauthenticated client cached on this long-lived
    // instance: every later call took the `if (this.client)` fast path above, handed
    // out the dead client, and failed downstream inside tsdav with a bare "no account
    // for fetchCalendars" instead of the real auth error. Keeping the client local
    // until login succeeds makes the invariant structural — `this.client` only ever
    // holds a logged-in client — so a failed login here means the next call retries
    // login instead of reusing a dead one.
    try {
      await client.login();
    } catch (e) {
      const detail = e instanceof Error ? e.message : String(e);
      // A login rejection is a credentials/config problem, not a caller argument
      // problem, so this stays a plain Error (InternalError) rather than
      // InvalidInputError — same classification as validateOrganizerUsername above.
      // tsdav's own message doesn't say which credential is wrong, and the CalDAV
      // app password is a separate credential from the Fastmail JMAP API token used
      // elsewhere in this server, so name it explicitly.
      throw new Error(
        `CalDAV login failed: ${detail}. Check the configured CalDAV app password ` +
        `(a separate credential from the Fastmail JMAP API token).`,
      );
    }

    this.client = client;
    return this.client;
  }

  /**
   * Discover the account's calendars, failing loudly instead of returning an empty list.
   *
   * tsdav's `fetchCalendars` never checks `response.ok`. On a non-2xx PROPFIND its parser
   * produces an error pseudo-response, which is then filtered out on `props.resourcetype`,
   * so the call returns `[]` and does NOT throw. Every read here used to take that at face
   * value, and the consequences all pointed the same way (#100): `list_calendar_events`
   * reported success with no events, which answers "am I free?" with "yes" on a server
   * failure; `getCalendarEventById` blamed the caller's event id for a discovery that never
   * happened; and because `[]` is truthy, `if (!this.calendars)` CACHED the empty result on
   * this long-lived client, repeating it until the process restarted.
   *
   * So: cache only a non-empty result, and treat empty as a failure to be explained.
   *
   * An empty-BUT-SUCCESSFUL discovery is also treated as a failure, which is the surprising
   * half and should not be "tidied up" into returning `[]`. Two reasons. The dangerous
   * direction here is under-reporting — an availability question answered with nothing reads
   * as free time that may not exist — and this is an account-shape assumption that stays
   * checkable: a Fastmail account always has at least one calendar collection, so an empty
   * successful discovery means something went wrong upstream of the status code far more
   * often than it means the account genuinely has no calendars. If that ever stops being
   * true of the accounts this server targets, this is the line to revisit.
   */
  private async discoverCalendars(): Promise<DAVCalendar[]> {
    const client = await this.getClient();
    if (this.calendars && this.calendars.length > 0) return this.calendars;

    const calendars = await client.fetchCalendars();
    if (calendars.length > 0) {
      this.calendars = calendars;
      return calendars;
    }

    // Empty. Re-ask with a status-checked PROPFIND, because the status is the one thing
    // fetchCalendars threw away — this is the read-path counterpart to assertDavOk, which
    // exists for the same reason on the write paths. The extra round-trip only ever happens
    // on the already-broken path.
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

  async getCalendars(): Promise<CalendarInfo[]> {
    const calendars = await this.discoverCalendars();

    return selectableCalendars(calendars)
      .map(c => ({
        id: c.url || '',
        displayName: String(c.displayName || 'Unnamed'),
        url: c.url || '',
        description: c.description || undefined,
        color: (c as any).calendarColor || undefined,
      }));
  }

  async getCalendarEvents(calendarId?: string, limit: number = 50, startDate?: string, endDate?: string): Promise<CalendarEventQueryResult> {
    const client = await this.getClient();
    const calendars = await this.discoverCalendars();

    let targetCalendars = selectableCalendars(calendars);
    // Trimmed and tested for PRESENCE, not for truthiness. `calendarId` narrows what the call
    // touches, so it fails CLOSED: an empty or whitespace-only value is a value that matched
    // no calendar, never "read every calendar". Under the old truthiness test `''` skipped the
    // filter entirely and quietly widened the query to the whole account while `'   '` was
    // correctly rejected — two spellings of the same mistake answered from two different
    // calendars. See docs/conventions.md on arguments that narrow a call.
    const requested = typeof calendarId === 'string' ? calendarId.trim() : calendarId;
    if (requested !== undefined && requested !== null) {
      const matched = targetCalendars.filter(
        c => c.url === requested || c.displayName === requested
      );
      // A calendarId that matches nothing used to leave the list empty, so the loop below
      // never ran and the tool answered "Showing 0 of 0 results." — an availability question
      // answered "you are free" because of a typo. Matching is exact, so "work" for "Work"
      // is a plausible first-try miss. The write path has always thrown here; both raise the
      // same error through one helper so a caller sees one rule, not two.
      if (matched.length === 0) throw calendarNotFoundError(calendarId, targetCalendars);
      targetCalendars = matched;
    }

    // The window is normalised ONCE and then used for two different things — the server's
    // time-range filter and the local re-filter below — so the two cannot disagree about
    // which days were asked for.
    //
    // A DATE IS A LOCAL DAY. "What is on the 12th?" asks about the caller's own day, so both
    // bounds resolve in the configured zone; only a value carrying `Z` or an offset is taken
    // as the instant it names. Read as UTC days instead, a +10:00 account was answered with
    // the 12th 10:00 to the 13th 10:00 and lost that morning's appointments with nothing
    // said. The zone is the SAME stored value every email timestamp renders in, so the day a
    // calendar query covers and the day an email is dated cannot drift apart.
    const zone = getDefaultTimezone();
    const rawStart = coerceCalendarWindowStart(startDate, 'startDate', zone);
    const rawEnd = coerceCalendarWindowEnd(endDate, 'endDate', zone);

    // A CALLER-NAMED bound is saturated too, not just the invented half. It resolves through
    // a zone, and an offset alone is enough to push it over the four-digit-year range every
    // consumer of these values can express — `endDate: "9999-12-31"` on a UTC-5 account
    // resolves to `+010000-01-01T05:00:00Z`. tsdav's `^\d{4}` check then threw a plain Error,
    // reporting a caller-fixable argument as a server fault.
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
    if (start || end) {
      let windowStart = start;
      let windowEnd = end;
      let invented: 'startDate' | 'endDate' | undefined;
      if (!windowStart) {
        windowStart = shiftIsoDays(windowEnd!, -CALENDAR_OPEN_WINDOW_DAYS);
        invented = 'startDate';
      } else if (!windowEnd) {
        windowEnd = shiftIsoDays(windowStart, CALENDAR_OPEN_WINDOW_DAYS);
        invented = 'endDate';
      }

      // AFTER the invented half is filled in, not as its alternative. The old `else if` made
      // this check the both-bounds case only, and the comment below said so — but saturation
      // means a ONE-SIDED window can come out zero-length too: `startDate:
      // "9999-12-31T23:59:59Z"` leaves the invented +366 days nowhere to go, so both ends land
      // on the same instant. Left to tsdav that was a plain Error (InternalError) or, worse, a
      // silently empty answer under a note claiming a 366-day span had been searched.
      if (Date.parse(windowStart!) >= Date.parse(windowEnd!)) {
        // Checked here rather than left to tsdav, which throws a plain Error for a backwards
        // range. That reaches the tool boundary as InternalError ("server-side, a bare retry
        // might work"), which is false and unactionable for what is plainly a caller-fixable
        // pair of arguments. See docs/conventions.md on error classification.
        //
        // The message echoes what the CALLER typed and states what it resolved to underneath.
        // Printing only the post-coercion values reported `endDate 2026-08-11T00:00:00Z` back
        // at someone who passed `2026-08-10`, so the one line meant to explain the whole-day
        // rule instead quoted a value they could not find in their own call. An omitted bound
        // is quoted as such rather than as the string "undefined".
        //
        // The ZONE is named alongside the resolved range, because otherwise the resolved
        // range is the confusing part: a caller who passed two dates and gets back two
        // UTC instants offset from midnight has no way to tell a correct local-day reading
        // from a bug.
        //
        // Equality gets its own sentence. A window whose two bounds are the same instant is
        // not backwards, it is EMPTY, and printing "the range X .. X" as though one end were
        // before the other reads as a rounding error rather than as the answer — the shape a
        // caller lands on by passing the same instant twice, meaning a single day.
        const resolved = windowStart === windowEnd
          ? `both resolve to the same instant, ${windowStart}, which is a zero-length window`
          : `which resolve to the range ${windowStart} .. ${windowEnd}`;
        const quote = (v: unknown) => (v === undefined || v === null ? '(omitted)' : `"${echoCallerText(v)}"`);
        // The one-sided arm needs its own sentence: the caller cannot "put startDate before
        // endDate" when they only passed one, so the ordinary advice would be unfollowable.
        const advice = invented
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

      if (invented || saturated.length > 0) {
        windowClamp = { invented, saturated: saturated.length > 0 ? saturated : undefined, start: windowStart!, end: windowEnd! };
      }
      fetchOptions.timeRange = { start: windowStart, end: windowEnd };
      // Expansion is what makes a recurring event report the occurrence that actually falls
      // in the window instead of the series' original DTSTART (#64). tsdav only forwards
      // <C:expand> when a timeRange accompanies it, which is why this sits inside the same
      // branch rather than being set unconditionally.
      fetchOptions.expand = true;
    }

    const windowStartMs = fetchOptions.timeRange ? Date.parse(fetchOptions.timeRange.start) : NaN;
    const windowEndMs = fetchOptions.timeRange ? Date.parse(fetchOptions.timeRange.end) : NaN;

    const allEvents: CalendarEvent[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal, ...fetchOptions });
      for (const obj of objects) {
        // Every VEVENT in the blob, not the first: with `expand` a single resource carries
        // one block per in-window occurrence, so a first-match read drops all but one.
        // `expanded` is passed rather than sniffed — see parseCalendarObjects.
        for (const event of parseCalendarObjects(obj, { expanded: !!fetchOptions.expand })) {
          // A block that STILL CARRIES ITS RRULE is never dropped here, whatever its dates
          // say. This branch runs on an expanded query, so a surviving master means the server
          // declined to expand that resource — and the master then shows the series' ORIGINAL
          // DTSTART, which for a long-running weekly event is years before the window. Judged
          // on that date it fails the intersection test and the row disappears: the filter
          // whose whole purpose is "drop only what provably cannot intersect" would be turning
          // a wrongly-dated row into a missing one, on exactly the resource the server told us
          // repeats. A recurrence rule is proof that DTSTART is not the only date this event
          // has, so it is not a date the filter is entitled to judge.
          const provablyOutside = fetchOptions.timeRange
            && !event.recurrenceRule
            && !eventIntersectsWindow(event, windowStartMs, windowEndMs);
          if (provablyOutside) continue;
          allEvents.push(event);
        }
      }
      // NOTE: no early exit on `limit`. Breaking out of this loop once enough events had
      // been gathered meant later calendars were never queried at all, so with several
      // calendars the "earliest N" were the earliest N *of whichever calendar happened to be
      // read first* — silently, and fatally for any cross-calendar availability check (#100).
      // Every target calendar is read, and only then is the combined set sorted and trimmed,
      // which is what makes the slice a genuine top-N.
    }

    sortEventsByStart(allEvents, zone);

    return { events: allEvents.slice(0, limit), total: allEvents.length, windowClamp };
  }

  /**
   * Find the raw DAVCalendarObject by UID or URL.
   * Needed for update/delete which require the original object with url/etag.
   */
  private async findCalendarObjectByUID(eventId: string): Promise<DAVCalendarObject | null> {
    const client = await this.getClient();
    // Routed through discovery so a failed lookup can only mean "no object matched". A
    // discovery failure throws here instead, which is what keeps getCalendarEventById from
    // telling the caller their event id is wrong about a call that never reached a
    // calendar (#100).
    const calendars = await this.discoverCalendars();

    // The SELECTABLE calendars, matching the read path. Searching the unfiltered list let
    // get/update/delete reach a collection `list_calendars` never shows and
    // `list_calendar_events` answers "Calendar not found" for — a record this server would
    // destroy but would not let you look at.
    for (const cal of selectableCalendars(calendars)) {
      const objects = await client.fetchCalendarObjects({ calendar: cal });
      for (const obj of objects) {
        const vevent = extractVEvent(obj.data || '');
        if (!vevent) continue;
        const uid = parseICalValue(vevent, 'UID');
        if (uid === eventId || obj.url === eventId) {
          return obj;
        }
      }
    }

    return null;
  }

  async getCalendarEventById(eventId: string): Promise<CalendarEvent> {
    const obj = await this.findCalendarObjectByUID(eventId);
    // Throw rather than return null so the MCP tool surfaces a real not-found
    // error — matches updateCalendarEvent/deleteCalendarEvent below. A null
    // here used to reach callers as a successful "null" tool response.
    // InvalidInputError, not a plain Error: a wrong event id is caller-fixable,
    // so it must reach the boundary as InvalidParams ("re-form the call") rather
    // than InternalError ("server-side, a bare retry might work").
    if (!obj) {
      throw new InvalidInputError(`Calendar event not found: ${eventId}`);
    }
    return parseCalendarObject(obj, { includeParticipants: true });
  }

  async createCalendarEvent(event: {
    calendarId: string;
    title: string;
    description?: string;
    start: string;
    end: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
  }): Promise<string> {
    const client = await this.getClient();
    const calendars = await this.discoverCalendars();

    // Selected exactly as the read path selects: the same filtered list and the same trim.
    // The two used to share only their ERROR MESSAGE, which made the divergence harder to
    // see rather than easier — an event could be written into (and later deleted from) a
    // collection no read tool would show, and `" Work "` failed here while succeeding on a
    // list. `calendarId` narrows what the call touches, so an empty or whitespace-only value
    // fails closed here for the same reason it does there.
    const requested = typeof event.calendarId === 'string' ? event.calendarId.trim() : event.calendarId;
    const selectable = selectableCalendars(calendars);
    const targetCal = selectable.find(
      c => c.url === requested || c.displayName === requested
    );
    if (!targetCal) {
      // A calendarId that matches no calendar is caller-fixable: they re-issue the call
      // with an id or name from list_calendars. Shared with the read path so both state
      // the same rule — see calendarNotFoundError.
      throw calendarNotFoundError(event.calendarId, selectable);
    }

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

    // Format start/end with all-day event support
    const startLine = formatDateTimeProperty('DTSTART', event.start, null, '\r\n');
    const endLine = formatDateTimeProperty('DTEND', event.end, null, '\r\n');

    // Time frame + ordering consistency. Classifying the serialized lines rather
    // than the raw inputs is the same thing the update path does, so create and
    // update reject an identical set of bad pairs.
    validateDateConsistency(
      describeDateProperty(startLine, event.start),
      describeDateProperty(endLine, event.end)
    );

    const icalLines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//fastmail-mcp//CalDAV//EN',
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      `LAST-MODIFIED:${now}`,
      startLine,
      endLine,
      foldICalLine(`SUMMARY:${escapeICalText(event.title)}`),
    ];

    if (event.description) {
      icalLines.push(foldICalLine(`DESCRIPTION:${escapeICalText(event.description)}`));
    }
    if (event.location) {
      icalLines.push(foldICalLine(`LOCATION:${escapeICalText(event.location)}`));
    }

    // Participant support
    if (event.participants && event.participants.length > 0) {
      // Validate all emails first
      for (const p of event.participants) {
        validateAttendeeEmail(p.email);
      }

      // ORGANIZER required when ATTENDEEs present. Validate the username as a
      // strict addr-spec (rejects ; , : CR LF etc.) so it can't corrupt or inject
      // into the ORGANIZER line when embedded below.
      const caldavUsername = this.config.username;
      validateOrganizerUsername(caldavUsername);
      const displayName = resolveDisplayName(this.config.displayName, caldavUsername);
      const cnPart = `;CN=${quoteParamValue(displayName)}`;
      icalLines.push(foldICalLine(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));

      // ATTENDEE lines — do NOT emit RSVP=TRUE by default (RFC 5545 §3.2.17 defaults to FALSE)
      for (const p of event.participants) {
        const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
        icalLines.push(foldICalLine(`ATTENDEE${cnParam}:mailto:${p.email}`));
      }
    }

    icalLines.push('END:VEVENT');
    icalLines.push('END:VCALENDAR');

    // Trailing CRLF per RFC 5545 §3.1
    const ical = icalLines.join('\r\n') + '\r\n';

    const createResp = await client.createCalendarObject({
      calendar: targetCal,
      filename: `${uid}.ics`,
      iCalString: ical,
    });
    assertDavOk(createResp, 'create calendar event');

    return uid;
  }

  async updateCalendarEvent(eventId: string, fields: {
    title?: string;
    description?: string;
    start?: string;
    end?: string;
    location?: string;
    participants?: Array<{ email: string; name?: string }>;
    clearFields?: string[];
  }): Promise<string> {
    const client = await this.getClient();
    const obj = await this.findCalendarObjectByUID(eventId);
    if (!obj) {
      // Caller-fixable bad id, same as getCalendarEventById: InvalidParams, not the
      // InternalError a plain Error maps to.
      throw new InvalidInputError(`Calendar event not found: ${eventId}`);
    }

    // Stays a plain Error: the event was found, but the server handed back an object with
    // no usable iCal payload. Nothing in the caller's arguments can change that.
    // Structural, not a substring test: `includes('BEGIN:VEVENT')` is true of a payload whose
    // only occurrence of that text sits inside a DESCRIPTION, which then fell through to a
    // different error message about the same condition.
    if (!obj.data || extractVEventBlocks(obj.data).length === 0) {
      throw new Error('Cannot update event: no iCal data found');
    }

    // Raised before ANY argument validation, because nothing in the arguments can make this
    // call legal — reporting a date-format problem first would imply fixing it would help.
    // Keyed on the resolved RESOURCE rather than the argument, so the `url` form of the id
    // (which findCalendarObjectByUID accepts, and which every serialised row carries) reaches
    // the same refusal as the UID form. See recurringSeriesRefusal for the reasoning (#146).
    if (isRecurringSeriesResource(obj.data)) {
      throw recurringSeriesRefusal('update', calendarObjectTitle(obj.data, eventId));
    }

    // Validate date inputs early before any processing
    const datePattern = /^\d{4}-\d{2}-\d{2}$/;
    const dateTimePattern = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;
    if (fields.start !== undefined && !datePattern.test(fields.start) && !dateTimePattern.test(fields.start)) {
      throw new InvalidInputError(`Invalid start date format: ${fields.start}. Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }
    if (fields.end !== undefined && !datePattern.test(fields.end) && !dateTimePattern.test(fields.end)) {
      throw new InvalidInputError(`Invalid end date format: ${fields.end}. Expected ISO 8601 (e.g. 2026-04-07T14:00:00Z or 2026-04-07)`);
    }

    // Validate clearFields: only the optional, string-settable, not-otherwise-
    // clearable fields may be cleared, and a field can't be both set and cleared.
    const CLEARABLE_FIELDS = new Set(['description', 'location']);
    const providedStringFields = new Set<string>();
    if (fields.description !== undefined) providedStringFields.add('description');
    if (fields.location !== undefined) providedStringFields.add('location');
    validateClearFields(fields.clearFields, CLEARABLE_FIELDS, providedStringFields);

    const lineEnding = detectLineEnding(obj.data);
    const fold = (line: string) => foldICalLine(line, lineEnding);

    // All patch helpers target the FIRST VEVENT — make sure that's the master,
    // not an overridden instance (component order is not guaranteed by RFC).
    const normalizedData = normalizeMasterVEventFirst(obj.data);

    // Capture original VEVENT before any patching for reads
    const originalVevent = extractVEvent(normalizedData);
    // Also a plain Error: the stored object is malformed, which is not a caller input fault.
    if (!originalVevent) {
      throw new Error('Cannot update event: no VEVENT block found');
    }

    const existingUid = parseICalValue(originalVevent, 'UID') || eventId;
    let data = normalizedData;

    // --- Patch fields ---
    let newStartLine: string | null = null;
    let newEndLine: string | null = null;
    let timeChanged = false;

    if (fields.title !== undefined) {
      const title = requireNonEmpty(fields.title, 'title');
      data = replaceICalProperty(data, 'SUMMARY', fold(`SUMMARY:${escapeICalText(title)}`));
    }

    if (fields.description !== undefined) {
      const description = requireNonEmpty(fields.description, 'description');
      data = replaceICalProperty(data, 'DESCRIPTION', fold(`DESCRIPTION:${escapeICalText(description)}`));
    }

    if (fields.start !== undefined) {
      newStartLine = formatDateTimeProperty('DTSTART', fields.start, originalVevent, lineEnding);
      data = replaceICalProperty(data, 'DTSTART', newStartLine);
      timeChanged = true;
    }

    if (fields.end !== undefined) {
      newEndLine = formatDateTimeProperty('DTEND', fields.end, originalVevent, lineEnding);
      data = replaceICalProperty(data, 'DTEND', newEndLine);
      // Remove DURATION — DTEND and DURATION are mutually exclusive (RFC 5545 §3.6.1)
      data = removeAllICalProperties(data, 'DURATION');
      timeChanged = true;
    }

    // Time frame + ordering consistency, judged on the pair that will actually
    // be written: the freshly formatted line for a side the caller supplied,
    // and the STORED line for a side they left alone. Comparing only the
    // caller's own values would miss the single-sided update entirely, which is
    // where both a frame flip (a floating start landing beside a UTC end) and a
    // backwards DTEND come from. The check is skipped when neither side was
    // touched, so a title-only edit is never blocked by an inconsistency that
    // was already in the stored event.
    if (fields.start !== undefined || fields.end !== undefined) {
      const startLine = newStartLine ?? parseAllICalProperties(originalVevent, 'DTSTART')[0];
      const endLine = newEndLine ?? parseAllICalProperties(originalVevent, 'DTEND')[0];
      // A DURATION-based event has no stored DTEND — nothing to compare against.
      if (startLine && endLine) {
        validateDateConsistency(
          describeDateProperty(startLine, newStartLine ? fields.start : undefined),
          describeDateProperty(endLine, newEndLine ? fields.end : undefined)
        );
      }
    }

    if (fields.location !== undefined) {
      const location = requireNonEmpty(fields.location, 'location');
      data = replaceICalProperty(data, 'LOCATION', fold(`LOCATION:${escapeICalText(location)}`));
    }

    // Clear requested fields by removing the property line entirely.
    if (fields.clearFields && fields.clearFields.length > 0) {
      const KEY_BY_FIELD: Record<string, string> = { description: 'DESCRIPTION', location: 'LOCATION' };
      for (const field of fields.clearFields) {
        data = replaceICalProperty(data, KEY_BY_FIELD[field], null);
      }
    }

    if (fields.participants !== undefined) {
      // Validate emails
      for (const p of fields.participants) {
        validateAttendeeEmail(p.email);
      }
      // Remove all existing ATTENDEE lines
      data = removeAllICalProperties(data, 'ATTENDEE');
      // Clearing all participants must also strip ORGANIZER — an ORGANIZER with
      // no ATTENDEEs is a malformed scheduling VEVENT (RFC 5545 §3.8.4.3). On the
      // length>0 path below the ORGANIZER is re-added, so this is gated to ===0.
      if (fields.participants.length === 0) {
        data = removeAllICalProperties(data, 'ORGANIZER');
      }
      // Build and insert all ATTENDEE lines in one pass
      if (fields.participants.length > 0) {
        const attendeeLines = fields.participants.map(p => {
          const cnParam = p.name ? `;CN=${quoteParamValue(p.name)}` : '';
          return fold(`ATTENDEE${cnParam}:mailto:${p.email}`);
        }).join(lineEnding);
        data = insertBeforeEndVEvent(data, attendeeLines);
      }
      // Add ORGANIZER if absent and participants are being added (RFC 5545 §3.8.4.1)
      if (fields.participants.length > 0 && !hasICalProperty(extractVEvent(data) || '', 'ORGANIZER')) {
        const caldavUsername = this.config.username;
        // Same strict addr-spec check the create path applies. A bare
        // .includes('@') admits ; , : CR LF, which would corrupt or inject into
        // the ORGANIZER line built below — the two paths emit the identical line
        // from the identical value, so they validate it identically.
        validateOrganizerUsername(caldavUsername);
        const displayName = resolveDisplayName(this.config.displayName, caldavUsername);
        // Always a CN: resolveDisplayName falls back to the username, which the check
        // above has just proved is a usable address, so it can never be empty.
        const cnPart = `;CN=${quoteParamValue(displayName)}`;
        data = replaceICalProperty(data, 'ORGANIZER', fold(`ORGANIZER${cnPart}:mailto:${caldavUsername}`));
      }
    }

    // --- SEQUENCE increment ---
    const hasAttendees = hasICalProperty(originalVevent, 'ATTENDEE');
    const schedulingSignificant = fields.start !== undefined || fields.end !== undefined ||
      fields.participants !== undefined || fields.location !== undefined;

    if (hasAttendees && schedulingSignificant) {
      const existingSeq = parseInt(parseICalValue(originalVevent, 'SEQUENCE') || '0', 10) || 0;
      data = replaceICalProperty(data, 'SEQUENCE', `SEQUENCE:${existingSeq + 1}`);
    }

    // --- Update DTSTAMP and LAST-MODIFIED ---
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    data = replaceICalProperty(data, 'DTSTAMP', `DTSTAMP:${now}`);
    data = replaceICalProperty(data, 'LAST-MODIFIED', `LAST-MODIFIED:${now}`);

    // --- Orphaned VTIMEZONE cleanup (LAST — after all modifications) ---
    if (timeChanged) {
      data = removeOrphanedVTimezones(data);
    }

    obj.data = data;
    const updateResp = await client.updateCalendarObject({ calendarObject: obj });
    assertDavOk(updateResp, 'update calendar event');

    return existingUid;
  }

  async deleteCalendarEvent(eventId: string): Promise<void> {
    const client = await this.getClient();
    const obj = await this.findCalendarObjectByUID(eventId);
    if (!obj) {
      // Caller-fixable bad id, matching getCalendarEventById/updateCalendarEvent.
      throw new InvalidInputError(`Calendar event not found: ${eventId}`);
    }

    // Raised AFTER the lookup and BEFORE the delete, so it covers both ways an id resolves:
    // findCalendarObjectByUID matches a UID or a `url` interchangeably, and `url` is on every
    // row this server serialises, so a guard keyed on the argument's shape would be bypassed
    // by passing the row's url. Reading the resolved RESOURCE cannot be. See
    // recurringSeriesRefusal for why this refuses rather than confirms (#146).
    if (isRecurringSeriesResource(obj.data)) {
      throw recurringSeriesRefusal('delete', calendarObjectTitle(obj.data, eventId));
    }

    const deleteResp = await client.deleteCalendarObject({ calendarObject: obj });
    assertDavOk(deleteResp, 'delete calendar event');
  }
}
