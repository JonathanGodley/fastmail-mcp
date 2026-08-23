import { DAVClient, DAVCalendar, DAVCalendarObject } from 'tsdav';
// requireNonEmpty/validateClearFields come from coerce.ts rather than being
// defined here, so their rejections throw the tagged InvalidInputError and the
// CallTool boundary maps them to InvalidParams. A plain Error would surface as
// InternalError ("server bug"), which is wrong for caller-fixable input and
// would tell the caller a bare retry might work. See docs/conventions.md.
import { InvalidInputError, requireNonEmpty, validateClearFields, coerceCalendarWindowStart, coerceCalendarWindowEnd, startOfLocalDayUtcIso, describeTimezone, resolveCalendarInstantMs, echoCallerText, ZONE_ECHO_LIMIT, resolveUsableTimezone, isUsableTimezone, validateCallerTimezone, canonicalZoneName, GREGORIAN_CYCLE_YEARS } from './coerce.js';
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
  // The clock "today" is read from, for the window a caller that named no bounds gets.
  // Injectable so that behaviour is unit-testable at all: a default window computed from the
  // real clock can only be asserted against a value the test recomputes the same way, which
  // tests nothing. Defaults to `Date.now`, and is read ONCE per call so the two ends of one
  // window cannot straddle a midnight.
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
  // The raw RDATE values of a block that lists its occurrences individually instead of (or as
  // well as) stating a rule — "20260612T090000Z,20260619T090000Z", every RDATE line in the
  // block joined into one comma-separated list, which is well-formed because an RDATE value is
  // already such a list.
  //
  // It exists for the same reason `recurrenceRule` does: it is proof that `start` is NOT the
  // only date this entry has. The window filter is not entitled to judge such a block on its
  // DTSTART alone, and a series that recurs only by RDATE carries no rule at all — so without
  // this field the guard in `getCalendarEvents` could not see it and dropped the row (#162).
  //
  // THE VALUES ARE CARRIED AND THE PARAMETERS ARE NOT — TZID, VALUE=DATE and VALUE=PERIOD are
  // all dropped, and that is a real limit rather than a detail. `RDATE;TZID=America/New_York:
  // 20270302T090000` arrives here as a bare `20270302T090000`, which under the rule that
  // governs `start` would read as FLOATING, and two RDATE lines in different zones join into
  // one list with nothing left to tell them apart. So the designator-less values in this field
  // do NOT follow the `timeZone` rule; they are only evidence that other dates exist. Both tool
  // descriptions say so.
  //
  // The shape stays values-only deliberately. A caller cannot act on an individual RDATE
  // through this server at all — there are no per-occurrence tools, and update/delete refuse
  // the whole series — so the field's entire job is to prove `start` is not the only date this
  // entry has. Carrying parameters would invite a precision it cannot back up. Read a real
  // occurrence date out of `list_calendar_events` with a window, which the server expands.
  //
  // Expected to be absent on the ordinary listing path: Cyrus strips RDATE from an expanded
  // block, the way it strips RRULE. Both halves are MEASURED — calendar-expand.probe.mjs
  // settles RRULE, and calendar-rdate-expand.probe.mjs settles RDATE, observing an RDATE-only
  // series come back from `<C:expand>` as one VEVENT per occurrence with no RDATE line on any
  // of them, in either serialisation. It shows up for certain on `get_calendar_event`, which
  // returns the unexpanded master, and on any block the server declined to expand.
  recurrenceDates?: string;
  // ---- time zone (#139) ----
  // The IANA name DTSTART was written in, when it is anything other than the zone this
  // server would otherwise assume. NEVER an offset: an offset is only valid at the one
  // instant it was computed for, and `start` above stays a bare local wall clock precisely
  // so DST is worked out by the READER's own zone database rather than baked in here.
  //
  // OMITTED — not set to the configured zone's own name — for the overwhelming majority of
  // rows, which are already in that zone; a caller reads "absent" as "the zone I asked
  // about". `null` means something different and narrower: `start` carries a genuinely
  // FLOATING value (RFC 5545 §3.3.5 — no TZID and no `Z`, a different instant for every
  // reader), and there is no zone name to give it. A `Z`-designated instant is also
  // omitted, because it already names its own instant; so is an all-day (date-only) value,
  // which has no zone at all by definition — emitting `null` for either would assert
  // "floating", which is a different, wrong fact.
  timeZone?: string | null;
  // The IANA name `end` was written in, but ONLY when it differs from `start`'s — RFC 5545
  // §3.8.5.3 and the write path (`validateDateConsistency`, fork issue #140) both allow a
  // start and an end in two different named zones, which is legal for the ordinary reason a
  // flight has one: it departs in one zone and lands in another. Omitted whenever end's zone
  // matches start's (the common case), whenever `end` is absent, and whenever `end` was
  // computed from a DURATION rather than a stored DTEND (it inherits start's zone by
  // construction, so there is nothing to disagree about). `null` means end is floating while
  // start is not, or the reverse.
  endTimeZone?: string | null;
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
  // Series left OUT of `events` and `total` because one resource expanded to more than
  // CALENDAR_MAX_OCCURRENCES_PER_SERIES blocks in the RANGE REQUESTED for this window, which
  // is the widened one rather than the caller's own (see CALENDAR_MAX_OCCURRENCES_PER_SERIES
  // for what that does to the threshold). ABSENT when nothing was
  // omitted, so silence means every series was materialised — the same never-silently-degrade
  // rule `windowClamp` follows, and the same STRUCTURE-not-prose shape: the wording lives in
  // `buildCalendarDensityNote` where a formatter test can reach it.
  denseSeries?: DenseCalendarSeries[];
}

/** One repeating event too dense to materialise, described well enough to find and narrow to. */
export interface DenseCalendarSeries {
  title: string;
  id: string;
  // Carried but deliberately NOT rendered in the note. It is here for the same reason
  // `CalendarEvent.url` is kept on an ordinary row (see the comment there): a UID is unique
  // per collection, not per account, so where two calendars hold the same UID the url is the
  // only thing that tells the two records apart — and a series omitted from the rows is
  // exactly the case where no row carries it. It stays out of the note text because title, id
  // and calendar name already identify the series for someone narrowing the window, and the
  // note is long enough without a path in it.
  url: string;
  calendar: string;
  occurrences: number;
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
 * The one place a "does this component have an RRULE / an RDATE / a RECURRENCE-ID / an
 * ATTENDEE?" question is answered. Every such test used to be its own `/^KEY[;:]/m` literal,
 * which is exactly the shape described above — and the recurrence gate now decides whether
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
 * Exported for its own test, and still worth having one: a caller-level test of the recurrence
 * gate pins the refusal's own reads rather than this presence test, and the two can disagree.
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
 * Four markers, ANY of which is enough, because this gates a refusal in front of an
 * irreversible write and so fails CLOSED — the cost of calling a one-off event a series is one
 * edit the caller does in the web client, the cost of the reverse is a destroyed series:
 *
 *  - a VEVENT carrying an RRULE — the ordinary series master;
 *  - a VEVENT carrying an RDATE — a series that LISTS its occurrences rather than stating a
 *    rule (RFC 5545 §3.8.5.2). It carries no rule at all, so an RRULE-only test read a
 *    single-block master of one as a one-off and let `delete_calendar_event` destroy the whole
 *    series and mail every attendee a cancellation (#162). `create_calendar_event` has no RDATE
 *    parameter either, so the same "cannot recreate it, must not destroy it" rule applies;
 *  - more than one VEVENT block in the resource — one CalDAV resource is one UID, so a second
 *    block can only be an overridden occurrence of a recurrence;
 *  - any block carrying a RECURRENCE-ID — that block IS an overridden occurrence, and a
 *    resource made entirely of them is a series whose master has been removed.
 *
 * The last two are the same test the read path uses to set `isRecurring`
 * (`blockCountProvesSeries`), deliberately, so the two halves of the server cannot disagree
 * about what a series is — and the RDATE marker is here for that same reason: the read path
 * sets `isRecurring` and emits `recurrenceDates` on exactly such a block, and the tool
 * descriptions promise that update and delete act on every occurrence of it. Note what is NOT
 * required: the RRULE and the overrides do not have to agree, and a malformed resource
 * carrying an override with no rule still refuses.
 *
 * The marker scan is VEVENT-WIDE and NOT position-aware: it asks whether the block's content
 * lines contain the property anywhere, so a line nested inside a subcomponent of the VEVENT
 * (a VALARM, say) counts as present — `hasRecurrenceId` included, being the same call. None of
 * RRULE, RDATE or RECURRENCE-ID is a DEFINED VALARM property (RFC 5545 §3.6.6 admits any
 * `iana-prop`, so such a line still parses), so a payload placing one there is malformed, and
 * the resulting refusal is the fail-closed direction this whole test already prefers. The
 * identical reach applied to RRULE long before RDATE joined it; do not narrow any of them into
 * a position-aware read without settling what the write path should then do with a malformed
 * resource.
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
  return blocks.some((block) =>
    hasICalProperty(block, 'RRULE') || hasICalProperty(block, 'RDATE') || hasRecurrenceId(block));
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
  options?: { includeParticipants?: boolean; expanded?: boolean; configuredZone?: string; blocks?: string[] },
): CalendarEvent[] {
  // `blocks` is an optimisation with a correctness point behind it. The listing path has to
  // COUNT this resource's blocks before deciding whether to parse it at all (see
  // CALENDAR_MAX_OCCURRENCES_PER_SERIES), and extracting them twice would mean two walks of a
  // payload that is large in exactly the case the count is guarding against — and, worse, two
  // places that could disagree about what a VEVENT block is. Passing the list through keeps
  // one structural extraction per resource.
  const blocks = options?.blocks ?? extractAllVEvents(obj.data || '');
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
export function parseCalendarObject(obj: DAVCalendarObject, options?: { includeParticipants?: boolean; configuredZone?: string }): CalendarEvent {
  return parseCalendarObjects(obj, options)[0];
}

function parseVEvent(
  obj: DAVCalendarObject,
  vevent: string,
  options?: { includeParticipants?: boolean; configuredZone?: string },
): CalendarEvent {
  const title = parseICalValue(vevent, 'SUMMARY') || 'Untitled';
  const description = parseICalValue(vevent, 'DESCRIPTION');
  const rawStart = parseICalValue(vevent, 'DTSTART');
  let rawEnd = parseICalValue(vevent, 'DTEND');
  const location = parseICalValue(vevent, 'LOCATION');
  const uid = parseICalValue(vevent, 'UID') || obj.url || '';

  // The zone this account is configured for, resolved to a name ICU can actually use. Passed
  // in rather than read from module state (see resolveUsableTimezone's own comment in
  // coerce.ts on why), so a test that pins a zone exercises the omit-when-same branch
  // deterministically rather than passing only when it happens to match the host.
  const configuredZone = options?.configuredZone ?? resolveUsableTimezone(undefined);

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
          // Start's own zone only — deliberately not the shared attachZoneFields, which would
          // also classify whatever raw DTEND line the source text happens to still contain.
          // An empty `DTEND;TZID=Europe/Paris:` value takes this branch too (parseICalValue
          // reads it as falsy) but still carries a TZID parameter, so describeDateProperty
          // would read it as `zoned` even though that value was never used to compute `end`.
          // A DURATION-computed end shares start's frame by construction — parseICalDuration
          // returns the new instant in the same spelling it read `start` in (see its own
          // "Return in same format as input start" comment) — so endTimeZone has nothing
          // independent to report here, and is left unset. See attachZoneFields' own doc
          // comment for the full reasoning.
          attachStartZone(event, vevent, configuredZone);
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
  addRecurrenceToEvent(event, vevent);
  if (options?.includeParticipants) {
    addParticipantsToEvent(event, vevent);
  }

  return event;
}

// Where a DTSTART/DTEND property's zone comes from, for the `timeZone`/`endTimeZone`
// response fields (#139). Four outcomes rather than `describeDateProperty`'s three,
// because `describeDateProperty` needs a raw line to classify and a missing property
// never produces one — `absent` is that fourth case.
type ZoneDescriptor =
  | { kind: 'tzid'; name: string }
  | { kind: 'floating' }
  | { kind: 'none' }
  | { kind: 'absent' };

/**
 * Classify a DTSTART/DTEND property's zone from its raw line(s).
 *
 * Built on `describeDateProperty` rather than re-deriving the TZID extraction, so the read
 * path and the write path's DTSTART/DTEND consistency check (`describeDateProperty`,
 * `validateDateConsistency`) agree about what a `zoned` value even is — including sharing
 * its `;TZID=("[^"]*"|[^;:]+)` extraction and unquoting, the same shape `formatDateTimeProperty`
 * uses to preserve a TZID on a floating rewrite.
 *
 * `describeDateProperty`'s `date` (all-day) and `utc` (Z-suffixed) frames collapse to one
 * `none` outcome here: both are already self-describing and neither carries a zone name, so
 * from this field's point of view they are the same fact — "nothing to say".
 */
function classifyZoneFromLines(rawLines: string[]): ZoneDescriptor {
  if (rawLines.length === 0) return { kind: 'absent' };
  const d = describeDateProperty(rawLines[0]);
  if (d.frame === 'zoned') return { kind: 'tzid', name: d.tzid! };
  if (d.frame === 'floating') return { kind: 'floating' };
  return { kind: 'none' };
}

// RFC 5545 §3.2.19 lets a TZID carry a leading '/', naming a zone registered by its creator
// rather than the plain IANA form — the simple case is '/Zone/Name', the same IANA-style name
// with one slash in front, and other clients emit it. libical strips exactly this one
// character before comparing, so stripping it here matches the platform rather than inventing
// a rule. This is comparison-only: `attachZoneFields` emits the stored spelling verbatim, slash
// and all, so the field always reflects what the calendar actually stored.
//
// A vendor-prefixed TZID ('/vendor.example/20050126_1/Australia/Sydney') still compares
// unequal after stripping one slash — that is the safe direction, an extra emitted field
// rather than a false claim that two differently-registered names are the same zone.
//
// After the strip, each side is canonicalised through `canonicalZoneName` (coerce.ts) when ICU
// can resolve it — the same seam `validateCallerTimezone` and `resolveUsableTimezone` route
// through to decide what actually lands on the wire (#157). Without this, a link/alias spelling
// compared unequal to the canonical name it resolves to: a caller echoing back a stored 'NZ'
// TZID as timeZone 'NZ' read as a DIFFERENT zone from the 'Pacific/Auckland' this server itself
// writes for it, producing a false "stranded two-zone event" rejection on an ordinary
// read-modify-write round trip (#139). A name ICU cannot resolve at all (a Windows zone id, a
// vendor-prefixed TZID) falls back to the trimmed/stripped string unchanged — still today's
// conservative string comparison, and still a real rejection when the two sides genuinely
// differ.
function normalizeZoneForComparison(name: string): string {
  const trimmed = name.trim();
  const stripped = trimmed.startsWith('/') ? trimmed.slice(1) : trimmed;
  return canonicalZoneName(stripped);
}

// Zone names compare trimmed, slash-normalized and ICU-alias-canonicalised (see above), then
// case-insensitively ("Australia/Sydney" == "australia/sydney" == "/Australia/Sydney" ==
// "NZ" == "Pacific/Auckland") and no other way: emitting an extra field when two spellings of
// the same zone differ only in these ways is the safe direction, claiming two genuinely
// different zones are the same is not.
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
    // Trimmed because a padded TZID that DIFFERS from the configured zone must still emit a
    // name isUsableTimezone (and everything downstream, e.g. sortEventsByStart) can resolve —
    // isUsableTimezone("Australia/Sydney ") is false. NOT slash-normalized: the stored
    // spelling, leading '/' and all, is what actually named the zone.
    if (!zoneNamesEqual(startDesc.name, configuredZone)) event.timeZone = startDesc.name.trim();
  } else if (startDesc.kind === 'floating') {
    event.timeZone = null;
  }
  // 'none' (Z-instant or all-day) and 'absent' both omit — see the field comment.
  return startDesc;
}

function attachEndZone(event: CalendarEvent, startDesc: ZoneDescriptor, vevent: string, configuredZone: string): void {
  const endDesc = classifyZoneFromLines(parseAllICalProperties(vevent, 'DTEND'));
  if (endDesc.kind === 'absent' || endDesc.kind === 'none') return;

  const compareDesc: ZoneDescriptor = startDesc.kind === 'absent'
    ? { kind: 'tzid', name: configuredZone }
    : startDesc;
  if (zoneDescriptorsEqual(endDesc, compareDesc)) return;

  // Trimmed for the same reason as `timeZone` above; stored spelling otherwise preserved.
  event.endTimeZone = endDesc.kind === 'tzid' ? endDesc.name.trim() : null;
}

/**
 * Set `timeZone`, and where applicable `endTimeZone`, on a parsed event from the VEVENT's raw
 * DTSTART/DTEND lines (#139). This is the one rule for the emit matrix documented on
 * `CalendarEvent` and in docs/conventions.md — read the descriptor comments there before
 * changing the branches above, they are not independent choices.
 *
 * `timeZone` describes start alone. `endTimeZone` describes end relative to start (the
 * flight-lands-elsewhere case `validateDateConsistency` permits on write, #140) — EXCEPT
 * when start is absent entirely, where end is compared against the configured zone instead,
 * which is exactly the rule `timeZone` itself uses.
 *
 * The DURATION branch in `parseVEvent` does NOT call this — it calls `attachStartZone` alone
 * and never touches `endTimeZone` at all, deliberately. Reusing `attachEndZone` there would
 * classify whatever raw DTEND line happens to be sitting in the source text even though its
 * VALUE was never used to compute `end` — an empty `DTEND;TZID=Europe/Paris:` still carries a
 * TZID parameter, so `describeDateProperty` reads it as `zoned` and a genuinely unused zone
 * would leak into the response. A DURATION-computed end shares start's frame by construction
 * (`parseICalDuration` returns the new instant in the same spelling it read `start` in — see
 * its own "Return in same format as input start" comment), so there is nothing independent
 * for `endTimeZone` to report regardless of what text a DTEND line happens to still contain.
 */
function attachZoneFields(event: CalendarEvent, vevent: string, configuredZone: string): void {
  const startDesc = attachStartZone(event, vevent, configuredZone);
  attachEndZone(event, startDesc, vevent, configuredZone);
}

/**
 * Attach the recurrence markers that say what KIND of date `start` is (#64).
 *
 * Read from the block being parsed rather than from the resource, because that is the
 * distinction being reported: an expanded occurrence has had its RRULE stripped by the
 * server and carries a RECURRENCE-ID, a master carries the rule and no RECURRENCE-ID.
 * Both set `isRecurring`; nothing is set for an ordinary one-off event, per the
 * omit-empty-fields convention.
 *
 * RDATE is read the same way and for the same reason as RRULE (#162): a block may list its
 * occurrences individually rather than state a rule, and such a block repeats just as much as
 * a ruled one does. EVERY RDATE line is read, not the first — RFC 5545 §3.8.5.2 allows the
 * property to appear any number of times, and a first-match read would hide the rest.
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
  // Parameters are deliberately dropped and the values kept verbatim — see the field's own
  // comment on `CalendarEvent`. `findValueBoundary` is what splits the two, quote-aware, so a
  // colon inside a quoted parameter value cannot be mistaken for the boundary.
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

// Where a freshly-WRITTEN property's TZID came from (#157) — used only for the
// designator-less (floating-input) branch of formatDateTimeProperty, since that is the only
// branch that can attach a TZID at all. Threaded onto `describeDateProperty`'s result as
// `tzidSource` so a frame-mismatch error can say "your account's configured zone, applied
// because you named none" instead of naming a zone the caller never wrote:
//   - 'caller' — the caller passed `timeZone` on this call.
//   - 'stored' — inherited from the event's own existing TZID (the long-standing behaviour).
//   - 'default' — no caller zone and nothing to inherit; `create` filled in the configured zone.
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
 * Validation and serialization of the caller's string are delegated wholesale to
 * validateAndFormatICalDate; this function only decides which property FORM the
 * result belongs in, and it decides that from the serialized value rather than by
 * re-reading the input. The serialized form is unambiguous — eight digits is an
 * all-day date, a trailing Z is an instant, anything else is a floating wall-clock
 * time — so the two functions cannot disagree about what the caller wrote, which is
 * exactly what went wrong while each of them parsed the input separately.
 *
 * `callerZone`/`defaultZone` only ever matter for case 2/3 — a designator-less value has no
 * zone of its own to override, which is exactly what qualifies it for one. Precedence for that
 * case (#157): the caller's own `timeZone` wins first, then the stored TZID this function has
 * always inherited, then `defaultZone`, then plain floating. `create_calendar_event` passes the
 * account's configured zone as `defaultZone`; `update_calendar_event` passes none at all, which
 * is what keeps its designator-less behaviour exactly as it was before `timeZone` existed —
 * inherit-then-floating, never defaulted. See docs/conventions.md for why that split is
 * deliberate rather than an inconsistency to "fix".
 */
function formatDateTimeProperty(
  propName: string,
  value: string,
  originalVevent: string | null,
  lineEnding: string,
  callerZone?: string,
  defaultZone?: string
): FormattedDateProperty {
  const serialized = validateAndFormatICalDate(value, propName);

  // Date-only
  if (/^\d{8}$/.test(serialized)) {
    return { line: foldICalLine(`${propName};VALUE=DATE:${serialized}`, lineEnding) };
  }

  // Floating time (no offset, no Z)
  if (!serialized.endsWith('Z')) {
    // The caller's own `timeZone` wins over everything this function would otherwise infer.
    // Interpolated with no escaping — sound only because `callerZone` reaches here already
    // validated by `validateCallerTimezone`, which returns ICU's canonical spelling for a name
    // `isUsableTimezone` proved resolvable. No zone name ICU resolves can contain `:`, `;`, `"`,
    // CR or LF, so there is no character here `foldICalLine` could mis-fold or that would forge
    // a second property line. This guard is load-bearing for this call site specifically: it
    // does not extend to `tzMatch[1]` below, which is read back from whatever a PREVIOUS write
    // (by this server or another CalDAV client) already stored.
    if (callerZone) {
      return { line: foldICalLine(`${propName};TZID=${callerZone}:${serialized}`, lineEnding), tzidSource: 'caller' };
    }
    // Try to preserve original TZID
    if (originalVevent) {
      const rawLines = parseAllICalProperties(originalVevent, propName);
      if (rawLines.length > 0) {
        const tzMatch = rawLines[0].match(/;TZID=("[^"]*"|[^;:]+)/);
        if (tzMatch) {
          return { line: foldICalLine(`${propName};TZID=${tzMatch[1]}:${serialized}`, lineEnding), tzidSource: 'stored' };
        }
      }
      // If propName is DTEND and no TZID found (DURATION-based), fall back to DTSTART's TZID
      if (propName === 'DTEND') {
        const startLines = parseAllICalProperties(originalVevent, 'DTSTART');
        if (startLines.length > 0) {
          const tzMatch = startLines[0].match(/;TZID=("[^"]*"|[^;:]+)/);
          if (tzMatch) {
            return { line: foldICalLine(`${propName};TZID=${tzMatch[1]}:${serialized}`, lineEnding), tzidSource: 'stored' };
          }
        }
      }
    }
    // Nothing to inherit — fall back to the caller-independent default (create's configured
    // zone; absent on update, which is what leaves update's no-timeZone behaviour untouched).
    if (defaultZone) {
      return { line: foldICalLine(`${propName};TZID=${defaultZone}:${serialized}`, lineEnding), tzidSource: 'default' };
    }
    // No TZID to preserve or default — emit as floating
    return { line: foldICalLine(`${propName}:${serialized}`, lineEnding) };
  }

  // UTC or offset — already normalized to a UTC instant
  return { line: foldICalLine(`${propName}:${serialized}`, lineEnding) };
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
   * Where a `zoned` frame's TZID came from (#157) — passed in by the caller who built the
   * line, since a raw property line alone cannot say whether its TZID was written because the
   * caller asked for it, inherited from what was already stored, or defaulted by `create` when
   * neither applied. Only ever set alongside `frame === 'zoned'`; see `describeFrame`.
   */
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
 * @param tzidSourceOverride where the line's TZID came from (#157), when the caller of this
 *   function knows — see `DatePropertyFrame.tzidSource`. Omitted for a line read straight from
 *   storage, which is always `stored` in spirit but has no wording that depends on saying so.
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
  const tzMatch = params.match(/;TZID=("[^"]*"|[^;:]+)/);
  if (tzMatch) {
    return { frame: 'zoned', tzid: tzMatch[1].replace(/^"|"$/g, ''), tzidSource: tzidSourceOverride, value, display };
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
    case 'zoned':
      // A `default` TZID (#157) is one this server filled in on `create` because the caller
      // named no zone at all — naming it the same way as a caller-chosen or inherited zone
      // would tell someone who wrote `2026-04-07T14:00:00Z` and no `timeZone` that they typed
      // a zone they never touched. `caller`/`stored`/undefined all read as an ordinary named
      // zone, which is the correct reading for each of those.
      if (d.tzidSource === 'default') {
        return `a date-time in the account's configured time zone (${echoCallerText(d.tzid!, ZONE_ECHO_LIMIT)}), applied because you named none`;
      }
      return `a date-time in time zone ${echoCallerText(d.tzid!, ZONE_ECHO_LIMIT)}`;
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
  if (start.frame === 'zoned' && start.tzid && end.tzid && !zoneNamesEqual(start.tzid, end.tzid)) return;

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
 * timezone, which is what `new Date(...)` would do. The FIXED frame is the point: the two
 * readings it compares have to be placed on one scale, and a frame that varied with the
 * deployment would make the same comparison answer differently by machine.
 *
 * THE WINDOW FILTER USED THIS AND NO LONGER DOES (#162). Placing an event against a window is
 * not a comparison of two like values, and reading a naive value as UTC there is precisely
 * what put a floating time in the wrong day; that path now resolves each value in the zone it
 * belongs to, through `resolveCalendarInstantMs`.
 *
 * SO NOTHING LIVE CALLS THIS. Its only caller in the source is `removeExceptionVEvents`, which
 * matches a RECURRENCE-ID against a set of orphaned ones — and that function has no caller of
 * its own outside its unit tests, because the write path refuses a recurring event outright
 * (see `recurringSeriesRefusal`), so no tool can reach it. Both are kept for their tests and
 * because the fixed frame is what would make that two-sided comparison meaningful if the
 * orphan-removal path is ever wired back up: both sides go through here, so the frame cancels
 * out and only its fixedness matters.
 */
export function parseICalDateAsUTC(iso: string): Date {
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return new Date(iso + 'T00:00:00Z');
  if (/Z$|[+-]\d{2}:?\d{2}$/.test(iso)) return new Date(iso);
  return new Date(iso + 'Z');
}

const DATE_ONLY_EVENT_VALUE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Which zone one of an event's values is really written in: its own TZID where ICU can
 * resolve that name, and the configured zone otherwise.
 *
 * The four cases, and why each lands where it does:
 *   ABSENT — `attachZoneFields` omits the field for a value already in the configured zone
 *     (and for a `Z`-designated or date-only one, which the resolver handles by shape), so
 *     absent means "the zone the caller asked about". Configured.
 *   `null` — a genuinely FLOATING value (RFC 5545 §3.3.5). No zone exists to place it in, so
 *     the caller's own clock is the least-wrong reading. Configured.
 *   A RESOLVABLE TZID — the correct reading rather than a best-effort one. Its own zone.
 *   AN UNRESOLVABLE NAME — a Windows zone name such as "AUS Eastern Standard Time" passed
 *     through verbatim rather than rejected. Configured, and the `isUsableTimezone` check is
 *     what makes that happen: `zoneOffsetMsAt` silently falls back to the HOST zone for a name
 *     it cannot resolve, which would place that one event in the deployment's zone instead of
 *     the account's.
 *
 * THE NAME IS NORMALISED BEFORE IT IS TESTED, and that is load-bearing rather than tidying.
 * `attachStartZone` deliberately emits the STORED spelling, so a TZID written in the RFC 5545
 * §3.2.19 global form ('/America/New_York') reaches here with its leading slash, and
 * `isUsableTimezone` says no to it. Testing the raw string would send a genuine New York event
 * to the configured zone and the exact filter would then DROP it — a missing event, which is
 * the direction #64 exists to prevent, arriving only for calendars whose zone happens to be
 * spelled that way. `normalizeZoneForComparison` strips the one leading slash and canonicalises
 * through ICU, so the global form resolves and an alias spelling ('NZ') resolves to the same
 * zone the comparison seam already treats it as. A VENDOR-PREFIXED name
 * ('/vendor.example/Australia/Sydney') still fails the test after the strip and still falls back
 * to configured, which is correct: nothing here can say which registry that name came from.
 *
 * Same rule, and the same reasoning, as `sortEventsByStart` uses to order the same values —
 * literally the same function, so the two cannot drift apart again.
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
 * The date-only value one calendar day after a date-only one.
 *
 * `Date.UTC` maps a two-digit year to 19xx, so the arithmetic steps over a whole 400-year
 * Gregorian cycle and back: leap rules repeat exactly across that cycle, so the rollover of
 * month, year and leap day is unaffected while the legacy mapping is not reachable.
 */
function nextDateOnly(dateOnly: string): string {
  const [y, mo, d] = dateOnly.split('-').map(Number);
  const shifted = new Date(Date.UTC(y + GREGORIAN_CYCLE_YEARS, mo - 1, d + 1));
  const year = shifted.getUTCFullYear() - GREGORIAN_CYCLE_YEARS;
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${String(year).padStart(4, '0')}-${pad(shifted.getUTCMonth() + 1)}-${pad(shifted.getUTCDate())}`;
}

/**
 * Whether a parsed event falls inside the requested window. EXACT, and authoritative for what
 * the caller is shown (#162).
 *
 * It used to be a residue filter that granted fourteen hours of slack on both edges to any
 * value with no zone designator, on the reasoning that such a value's real instant was
 * unknowable here and keeping an extra row beat dropping a real one. The keeping half of that
 * was right and stays; the fourteen hours were in the wrong place. A FILTER CANNOT KEEP WHAT
 * THE SERVER NEVER SENT, and the server withholds two kinds of event from a narrow window: it
 * matches a date-only value on its UTC day, and it resolves a floating time as UTC. Widening
 * here did nothing about either, while leaving visible residue at both edges. The margin now
 * widens the range REQUESTED (see `getCalendarEvents`), which is the only place it works, and
 * this function judges what comes back against the window the caller actually asked for.
 *
 * Exact does not mean eager to drop. Every value is resolved by WHAT IT IS, never guessed at
 * and never discarded for being hard to read:
 *   - a `Z` or offset value is the instant it names;
 *   - a date-only value is local midnight in the zone the value belongs to, and a date-only
 *     `end` is ALREADY exclusive in iCalendar (measured, docs/fastmail-action-availability.md),
 *     so a DTSTART..DTEND date span is the full multi-day local span with no day added;
 *   - a wall clock resolves in its own TZID where that resolves, and in the CONFIGURED zone
 *     otherwise — see `zoneForValue`, and note that a naive value must never be read as UTC
 *     here, which on a +10 account is what put a floating 20:00 outside the local evening;
 *   - an event with nothing readable to judge is KEPT. A promised event vanishing with no
 *     trace is worse than one extra row the caller can see and dismiss, and a missing event is
 *     the failure #64 exists to prevent.
 *
 * IT STILL DOES NOT CLOSE THE "SERVER DECLINED TO EXPAND" GAP. On an expanded query a
 * surviving master carries its recurrence carrier and its ORIGINAL DTSTART, and judging that
 * date drops the row entirely — turning a wrongly-dated event into a MISSING one. The caller
 * keeps such a block regardless of its dates (see the recurrence guard in
 * `getCalendarEvents`); this function is not the thing that saves it.
 *
 * `zone` is REQUIRED rather than defaulted so that no call site can silently fall through to
 * the host zone: `zoneOffsetMsAt` treats an undefined zone as the deployment's own, which
 * would make the same query include or exclude the same event depending on where this server
 * runs.
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
    // An all-day event with no end covers its whole day. Collapsing it to a zero-width
    // instant at local midnight is what dropped it from every window that started later that
    // morning. The next day is reached in LOCAL days, not by adding 24 hours, so a day a DST
    // transition makes 23 or 25 hours long still ends where the calendar says it does.
    //
    // The NaN arm is reachable only at the very end of the representable range: the day after
    // '9999-12-31' is '10000-01-01', which is not a four-digit-year date and so does not
    // resolve. Falling back to a flat 24 hours there is wrong about a DST transition and right
    // about the thing that matters — an all-day event is never NARROWER than its own day, and
    // collapsing it to local midnight is exactly the drop this branch exists to prevent.
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
 * Each start is resolved through the same machinery the window bounds use — but no longer
 * always in the CONFIGURED zone (#139). An event whose own start carries a `timeZone` ICU
 * can resolve is now sorted in ITS zone, the correct reading rather than a best-effort one;
 * the configured zone is only the fallback, used for a genuinely floating value (RFC 5545
 * §3.3.5 — no zone exists to sort it in, so the caller's own clock is the least-wrong guess)
 * and for a `timeZone` this server was handed but ICU cannot resolve (a Windows zone name
 * such as "AUS Eastern Standard Time" passed through verbatim rather than rejected). That
 * fallback matters beyond correctness: `zoneOffsetMsAt` silently falls back to the HOST zone
 * for an unresolvable name, so sorting by an unresolvable `timeZone` directly would place that
 * one event in the host zone instead of the account's configured one.
 *
 * That choice is `zoneForValue`, CALLED rather than restated. The sort and the window filter
 * have to agree about which zone an event is in — a read that filtered an event in one zone
 * and ordered it in another would put it at the wrong place in a list `limit` then truncates —
 * and two copies of the rule are two things to keep in step. One copy is what let a
 * slash-prefixed TZID be normalised on one side and not the other (#162).
 *
 * `zone` may be undefined here, unlike in the filter, because a sort still has to produce an
 * order when no zone was configured at all; `resolveCalendarInstantMs` takes it from there.
 *
 * An unreadable or absent start sorts FIRST. It cannot be placed, and placing it last would
 * put it where `limit` truncates — the one position that turns "cannot order this" into
 * "dropped this".
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
 * A MONTH is the span chosen. It answers the question a one-sided or bounds-free window is
 * actually asking — "what is coming up", "what led up to this date" — and it matches what a
 * calendar client puts on screen at a time, so the invented span is the one a caller reading
 * the answer already has in mind. The expanding APIs nearest to this one do not invent a span
 * at all: Cyrus's own JMAP `CalendarEvent/query` REJECTS an `expandRecurrences` query with no
 * upper bound, and Microsoft Graph's `calendarView` requires both bounds. Inventing a month is
 * already the generous reading; inventing a year was generous twice over, and it held an
 * ordinary daily series to a few hundred rows rather than a few dozen.
 *
 * 31 rather than 30 so the same date next month is always inside the span, from any starting
 * day in any month. The days are fixed 24-hour days, not local days, for the reason
 * `shiftIsoDays` gives: the span is invented rather than asked for, so a DST hour either side
 * of it is not a wrong answer to anything, and the note states the resulting instant.
 *
 * It bounds an INVENTED bound only — a caller that names both bounds gets exactly the window
 * it named, because that span is its own decision.
 *
 * The clamp is never silent: `windowClamp` names the window actually queried and how to ask
 * for more, so "nothing after that date" can't be read as an empty calendar.
 */
export const CALENDAR_OPEN_WINDOW_DAYS = 31;

/**
 * How many in-window occurrences ONE CalDAV resource may expand to before this server declines
 * to materialise it.
 *
 * The window bound above covers one of the two ways to ask for an unbounded expansion; this
 * covers the other. A caller naming both bounds gets exactly the span it named, because the
 * span is its own decision — but the caller chooses the span and an ATTACKER chooses the
 * DENSITY. Calendar content here is authored by anyone who can send an invitation, and one
 * `FREQ=MINUTELY` series fills any window at all.
 *
 * 5000 is set where it passes the dense-but-real cases and trips the ones nothing legitimate
 * produces:
 *
 *   10 years of a DAILY series            3,653   passes
 *   a month at every 10 minutes           4,464   passes
 *   a month at every 5 minutes            8,928   trips
 *   a month of FREQ=MINUTELY             44,640   trips
 *
 * THE COUNT IS OVER THE RANGE REQUESTED, NOT THE CALLER'S WINDOW. The blocks counted are the
 * ones the server returned for the widened request (MAX_UTC_OFFSET_MS at each edge, #162), so
 * the threshold is applied to a set up to 28 hours wider than the window the caller asked
 * about — a series with slightly under 5000 genuine in-window occurrences can therefore be
 * omitted. For a 31-day window the widened range is 772 hours against the caller's 744, so an
 * evenly spaced series trips the threshold from about 4819 in-window occurrences up
 * (5000 x 744/772); the shorter the window, the wider that gap gets. Narrowing to the
 * caller's window first would need the per-block parse this cap exists to avoid, so the number
 * is described accurately in the note rather than made exact.
 *
 * IT IS A PARSE-AND-SHOW THRESHOLD, NOT A RESPONSE SIZE. The response stays bounded by `limit`
 * (default 50, hard cap 500) exactly as before; this decides whether one resource's blocks are
 * parsed and offered at all.
 *
 * A tripped resource is left OUT of the rows entirely rather than truncated to its first N.
 * Truncating would fill the whole `limit` page with one series and push every real event off
 * it, which is the outcome the cap exists to prevent — one hostile invitation cannot blank a
 * listing. The omission is disclosed by name, count and calendar in a trailing note, so it is
 * never silent, and the call still answers.
 *
 * THE RESIDUAL, stated because it is not fixed: this bounds what this server PARSES and SHOWS,
 * never what Cyrus generates or transfers. Cyrus's `expand_cb` returns 1 unconditionally and
 * its `CALDAV:max-instances` property has no handler, so there is no server-side result limit
 * to ask for; tsdav's `fetchCalendarObjects` has no limit or paging option and buffers the
 * whole multistatus before this code sees any of it. The fetch is the platform's; the parse
 * and the page are the work on this side, and that is what this bounds.
 */
export const CALENDAR_MAX_OCCURRENCES_PER_SERIES = 5000;

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
 * span nobody named, chosen for being roughly a month; a DST hour either side of an arbitrary
 * 31-day bound is not a wrong answer to anything, and the note names the resulting instant.
 *
 * The example that reaches the saturation is `startDate: "9999-12-30"` ON ITS OWN, where the
 * invented END has nowhere to go. `endDate` alone runs the other way and cannot saturate at
 * that end, so the two directions are not interchangeable in an example.
 */
function shiftIsoDays(iso: string, days: number): string {
  return shiftIsoMs(iso, days * 24 * 60 * 60 * 1000);
}

/**
 * Shift an ISO-8601 UTC instant by milliseconds, keeping the seconds-precision form and
 * SATURATING at the ends of the representable range.
 *
 * Saturation is not optional here. `endDate: "9999-12-31"` plus the request margin below runs
 * off the four-digit year, `toISOString` answers with the expanded `+010000-…` form, and
 * tsdav's `^\d{4}` check throws a plain Error — an InternalError ("server-side, a bare retry
 * might work") raised over an argument the caller could have fixed.
 */
function shiftIsoMs(iso: string, ms: number): string {
  const shifted = new Date(Date.parse(iso) + ms)
    .toISOString()
    .replace(/\.\d{3}Z$/, 'Z');
  return saturateInstant(shifted);
}

// The widest UTC offset any IANA zone has ever used (+14:00 for Kiritimati; the negative
// extreme is smaller, so one bound covers both directions), and therefore how far the range
// SENT TO THE SERVER runs past the window the caller asked for, at each edge (#162).
//
// The margin lives here, on the request, because that is the only place it does any work. A
// client-side filter cannot keep an event the server never sent, and the server withholds two
// kinds of event from a narrow window — both measured against the live server by
// scripts/probes/calendar-window-frames.probe.mjs, neither fixable downstream:
//
//   AN ALL-DAY EVENT. Cyrus matches a date-only value on its UTC day, and `<C:expand>` emits
//     zero VEVENTs for a window that touches that UTC day without containing it. "The morning
//     of the 12th" on a UTC+10 account therefore returned no occurrence at all.
//   A FLOATING TIME. It is resolved as UTC server-side, so a floating 20:00 sits outside the
//     same account's local-evening window.
//
// Fourteen hours covers the worst case of either, in either direction. What comes back is then
// judged EXACTLY, by `eventIntersectsWindow`, against the window the caller actually asked
// for — so widening the request costs the caller nothing but the rows the filter then trims.
const MAX_UTC_OFFSET_MS = 14 * 60 * 60 * 1000;

// The window arguments quoted back in a rejection go through the SHARED echo in coerce.ts:
// scrubbed of the control characters that would forge extra lines, trimmed so the value shown
// is the value the coercion actually judged, and cut with a visible marker. It is the same
// helper the date rejections themselves use, so one message family cannot end up with two
// policies (#141).

// Calendar display names are server/user data of unbounded length, so a listing of them is
// capped the way every other echoed list in this server is.
const CALENDAR_NAME_LIST_CAP = 20;

// A calendar URL offered in the not-found error has to arrive USABLE, because it is offered
// as a `calendarId` to paste straight back — and `echoCallerText`'s default limit of 60 cuts
// every real one, with nothing in the truncated value to say it was cut. Fastmail's own
// prefix, `https://caldav.fastmail.com/dav/calendars/user/`, is 47 characters before the
// account's email address and the collection id even gets a look in. 200 clears a long
// address with room to spare while still bounding what one entry can add to the message; the
// list as a whole is bounded separately by CALENDAR_NAME_LIST_CAP. Display NAMES keep the
// 60-char default: a name is offered to be recognised, not pasted, and a caller who sees it
// truncated can still read it off `list_calendars`.
//
// Deliberately HERE rather than in coerce.ts beside DATE_ECHO_LIMIT and ZONE_ECHO_LIMIT: this
// one has a single consumer, and it belongs next to CALENDAR_NAME_LIST_CAP, which bounds the
// other half of the same message.
const CALENDAR_URL_ECHO_LIMIT = 200;

/**
 * The display name Fastmail gives the hidden task collection, which `list_calendars` filters
 * out. Named once so every place that hides it hides the same thing — including the
 * not-found error, which must not advertise a calendar the caller cannot see.
 */
const HIDDEN_TASK_CALENDAR_NAME = 'DEFAULT_TASK_CALENDAR_NAME';

/**
 * The one place a DAV `displayName` is turned into a name, because tsdav types it as `string`
 * and it is not one.
 *
 * tsdav 2.3.1 reads the property as `rs.props?.displayname?._cdata ?? rs.props?.displayname`
 * (`fetchCalendars`, dist/tsdav.cjs:1004) and hands the result straight through — unlike its
 * address-book path (:744), which guards with `typeof === 'string'`. What reaches us is
 * therefore whatever xml-js produced under tsdav's own parser options (:274-300). The
 * OBSERVED shapes, measured against xml-js with those options rather than assumed:
 *
 *   <displayname>Work</displayname>                 "Work"          a plain string
 *   <displayname>2026</displayname>                 2026            a NUMBER
 *   <displayname>true</displayname>                 true            a BOOLEAN
 *   <displayname><![CDATA[2026]]></displayname>     "2026"          a STRING, not a number
 *       (a string by the time it reaches us: the raw `{_cdata:…}` shape is what the
 *        DEFENSIVE table below covers, because tsdav flattens it before we see it)
 *   <displayname/>  or  <displayname></displayname> {}              an EMPTY OBJECT
 *   <D:displayname xml:lang="en"/>                  {_attributes:…} an OBJECT
 *   <displayname>A</displayname> twice              ['A','B']       an ARRAY
 *   (property absent)                               undefined
 *
 * And the DEFENSIVE ones, which tsdav's own read means cannot arrive today — it flattens
 * `_cdata` before we see the value, and its `textFn` replaces an element with its text so
 * `_text` never survives either. They are handled in case tsdav changes that read, and are
 * NOT evidence about what a live server sends. Do not build a fixture on them believing it
 * reachable; that mistake is what made the first version of this fix claim a bug that could
 * not happen.
 *
 *   <displayname><![CDATA[Work]]></displayname>     {_cdata:"Work"} tsdav flattens this today
 *   (a parser leaving compact text in place)        {_text:"Work"}  never produced today
 *
 * The number and boolean come from tsdav's `nativeType` (:104-114), which coerces any element
 * text that looks numeric or reads "true"/"false". It runs on TEXT ONLY, which is why the
 * CDATA row above differs from the plain row directly two lines up: xml-js routes character
 * data past the text callback, so `<![CDATA[2026]]>` survives as the string "2026" while the
 * same four characters written as text arrive as the number 2026. Both name the same
 * calendar, and this function reports both as "2026". THAT COERCION IS LOSSY BEFORE WE SEE IT: a
 * calendar named `1e3` arrives as the number 1000 and can only be listed as "1000", because
 * the original text is gone by then. What this function preserves is not the spelling but the
 * INVARIANT that matters — the name it returns is one the caller can pass straight back as a
 * `calendarId` and have it resolve, because both sides of that comparison come through here.
 * The empty object is xml-js compact mode: tsdav's `textFn` only fires when there IS text, so
 * an empty element keeps its bare compact form. Whitespace-only text is the same case,
 * because `trim: true` leaves nothing behind.
 *
 * The object cases are why this exists. `String({})` is `"[object Object]"` — a TRUTHY string,
 * so every `|| 'Unnamed'` fallback downstream is dead on exactly the input it was written for,
 * and the marker reached the caller as a calendar name.
 *
 * THE ARRAY IS A DELIBERATE DEGRADE, not an oversight: duplicate `<displayname>` elements are
 * a malformed collection, and the old code rendered them joined ("A,B") as though that were
 * the calendar's name. It is not one, so this returns undefined and each call site falls back
 * to something real (the URL, or "Unnamed").
 *
 * No URL fallback lives here. What an absent name should degrade to differs per call site (a
 * literal "Unnamed" in a listing, the collection URL in a density note and in an error
 * message's name list), so the helper answers only "is there a name, and what is it".
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
  // A name the parser typed for us is still the name the user gave the calendar, so it is
  // rendered rather than discarded — a calendar called "2026" keeps its name instead of
  // becoming "Unnamed". NaN/Infinity cannot come from `nativeType`, which rejects both.
  if (typeof scalar === 'number' && Number.isFinite(scalar)) return String(scalar);
  if (typeof scalar === 'boolean') return String(scalar);
  return undefined;
}

/** The calendars a caller is able to name, which is the only set worth listing back at them. */
function selectableCalendars(calendars: DAVCalendar[]): DAVCalendar[] {
  // Unwrapped before comparing, for WHITESPACE SYMMETRY and for consistency with every other
  // read of this field — not because a typed name was leaking through. It was not: the shapes
  // tsdav can hand back ({}, {_attributes}, a number, a boolean, an array) equal the hidden
  // name under a raw comparison exactly as rarely as they do under this one, which is never.
  // Nor is a padded exact name the exception: xml-js parses with `trim: true`, which trims
  // plain text and CDATA alike, so `'  DEFAULT_TASK_CALENDAR_NAME  '` cannot reach this code
  // either. There is NO reachable input where the raw comparison and this one disagree. The
  // unwrap is symmetry insurance — every other read of this field goes through the helper, and
  // a filter that compared raw would be the one place to re-check if that ever stopped being
  // true.
  return calendars.filter(c => unwrapDisplayName(c.displayName) !== HIDDEN_TASK_CALENDAR_NAME);
}

/**
 * The one "no calendar matched that id" error, raised by BOTH the read and the write path.
 *
 * `calendarId` accepts a CalDAV URL or a display name, and the name is matched
 * CASE-SENSITIVELY — surrounding whitespace is ignored on both sides, and `list_calendars`
 * reports the trimmed name, so what it hands back is a value this parameter resolves. "work"
 * for a calendar called "Work" still misses, which is what made a shared, self-correcting
 * message worth more than a bare echo: the available names are listed, so a caller that
 * guessed the spelling can fix the call without a second `list_calendars` round-trip.
 *
 * The listing is filtered HERE rather than at each call site, because the two call sites
 * passed different lists: the read path had already dropped the hidden task collection and
 * the write path had not, so a mistyped id on a create answered with a calendar name that
 * `list_calendars` never shows and no call can obtain. A shared message is only one rule if
 * it is given one input.
 */
function calendarNotFoundError(calendarId: unknown, available: DAVCalendar[]): InvalidInputError {
  // A nameless calendar is listed by its URL rather than dropped. Filtering it out left the
  // caller a message that silently under-reported what they could name — and the URL is not a
  // consolation prize here, it is a `calendarId` that works, which is exactly what this
  // message exists to hand back. Never silently drop a promised field; see CLAUDE.md.
  //
  // Which is why the two carry DIFFERENT echo limits rather than sharing one. A URL that is
  // offered as a working handle and then truncated is worse than one omitted: the caller
  // pastes it back and gets this same error again, with nothing saying the value was cut.
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
  // QUOTED. The two values this path exists to reject — an empty string and a whitespace-only
  // one — render as nothing at all unquoted, so the message read "Calendar not found: ." with
  // no sign of what had been rejected, while the available names two clauses later were
  // quoted. Same value class, same sentence, so the same treatment.
  //
  // And echoed at the URL bound, not the 60-char default, because the value being REJECTED is
  // itself often a URL — a caller who mistyped one is exactly who reaches here. 60 characters
  // leaves nothing past the 47-character fixed prefix, so two different wrong URLs render
  // byte-identical and the echo stops telling the caller which one they sent. Same bound the
  // offered URLs get.
  return new InvalidInputError(
    `Calendar not found: "${echoCallerText(calendarId, CALENDAR_URL_ECHO_LIMIT)}". calendarId takes either a calendar's URL ` +
    '(its `id` from list_calendars) or its display name, and the name is matched CASE-SENSITIVELY; ' +
    'surrounding whitespace is ignored on both sides, and list_calendars reports the trimmed name.' +
    `${listing}`,
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

// ---- write-side time zone result (#157) ----
//
// What createCalendarEvent/updateCalendarEvent actually put on the wire for `start`/`end`,
// computed from the WRITTEN property line — never from the caller's input — so a stored
// inherit or a create default is reported exactly as truthfully as a caller-supplied
// `timeZone`. describe{Create,Update}CalendarEventResult below turn this into the response
// sentence index.ts's handlers append; they live here rather than in index.ts (where the
// handler that calls them lives) because index.ts's CallTool switch has no test harness and
// the module itself runs `server.connect()` as a load-time side effect, which makes it unsafe
// to `import` from a unit test — CLAUDE.md's "Handler logic must be unit-testable" pattern
// (composeReply/reply-handler.ts) is to extract into a safely-importable module instead. This
// one is co-located with the types it formats rather than a third file, since it has no
// dependency of its own beyond them.
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
}

export interface UpdateCalendarEventResult {
  eventId: string;
  /** Present only when this call actually wrote (touched) that side. */
  start?: CalendarZoneWriteInfo;
  end?: CalendarZoneWriteInfo;
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
    case 'zoned': return `zone ${info.zone}`;
    case 'utc': return 'UTC';
    case 'floating': return 'floating (no zone)';
    case 'allday': return 'all-day (no time component)';
  }
}

// Given only the structured write result (never the caller's input — a caller-named zone, an
// inherited stored TZID, and create's configured-zone default all arrive at createCalendarEvent
// as the same "no designator" input, so only the WRITTEN line can say which one actually
// happened), returns the sentence index.ts's create_calendar_event handler appends to its
// response text. Pure and exported for direct unit testing (#157).
export function describeCreateCalendarEventResult(result: CreateCalendarEventResult): string {
  const startDesc = describeCalendarZoneWrite(result.start);
  const endDesc = describeCalendarZoneWrite(result.end);
  return startDesc === endDesc
    ? ` Written in ${startDesc}.`
    : ` Start written in ${startDesc}, end written in ${endDesc}.`;
}

// Same idea as describeCreateCalendarEventResult, but update only ever touches the sides the
// caller actually supplied (start/end are each optional on the result), so an untouched side
// is omitted rather than described.
export function describeUpdateCalendarEventResult(result: UpdateCalendarEventResult): string {
  const parts: string[] = [];
  if (result.start) parts.push(`start ${describeCalendarZoneWrite(result.start)}`);
  if (result.end) parts.push(`end ${describeCalendarZoneWrite(result.end)}`);
  if (parts.length === 0) return '';
  return ` (${parts.join(', ')})`;
}

/**
 * `timeZone` only qualifies a designator-less value — one that carries neither its own
 * `Z`/offset nor a date-only marker. A value that already names its own instant, or an all-day
 * value that has no time component at all, makes `timeZone` a contradiction rather than a
 * qualifier, and this repo rejects a contradicting argument instead of silently ignoring one of
 * the two (docs/conventions.md, fail-closed narrowing arguments). Runs BEFORE
 * formatDateTimeProperty ever sees `callerZone`, so a rejected call never reaches the point of
 * writing anything.
 *
 * Validates through the same field name (`DTSTART`/`DTEND`) the real formatting call downstream
 * uses, not the human-readable `label` — otherwise the same malformed date string produces two
 * different rejection sentences depending only on whether `timeZone` happened to be passed
 * alongside it.
 */
function rejectTimezoneConflict(value: string, label: 'start' | 'end', callerZone: string): void {
  const propName = label === 'start' ? 'DTSTART' : 'DTEND';
  const serialized = validateAndFormatICalDate(value, propName);
  if (/^\d{8}$/.test(serialized)) {
    throw new InvalidInputError(
      `timeZone cannot be combined with a date-only ${label} ('${value}') — an all-day value has ` +
      `no time zone. Drop timeZone, or pass ${label} with a time component for it to qualify.`
    );
  }
  if (serialized.endsWith('Z')) {
    throw new InvalidInputError(
      `timeZone cannot be combined with a ${label} that already carries Z or a UTC offset ('${value}') ` +
      `— that value already names a fixed instant of its own. Drop timeZone, or pass ${label} as a ` +
      `bare wall-clock value (no Z, no offset) for timeZone to qualify.`
    );
  }
}

/**
 * On `update_calendar_event`, `timeZone` combined with only ONE of `start`/`end` can silently
 * strand the untouched side in a different, still-stored zone — manufacturing a two-zone event
 * (the flight-lands-elsewhere shape #140 legitimises) that nobody asked for, and doing it past
 * `validateDateConsistency`'s ordering check rather than through it: two `zoned` values in
 * different TZIDs is the one case that check deliberately stands down on, so this is the one
 * case it will not catch. Only a STORED, DIFFERENTLY-NAMED `zoned` value is a problem — a
 * stored floating or `Z` value already trips the ordinary frame mismatch inside
 * `validateDateConsistency`, and a stored TZID matching `callerZone` produces no discrepancy —
 * so this only ever fires for the one shape that check cannot see.
 */
function rejectStrandedZoneMismatch(originalVevent: string, updatedSide: 'start' | 'end', callerZone: string): void {
  const strandedProp = updatedSide === 'start' ? 'DTEND' : 'DTSTART';
  const strandedLabel = updatedSide === 'start' ? 'end' : 'start';
  const strandedLines = parseAllICalProperties(originalVevent, strandedProp);
  if (strandedLines.length === 0) return;
  const desc = describeDateProperty(strandedLines[0]);
  if (desc.frame === 'zoned' && desc.tzid && !zoneNamesEqual(desc.tzid, callerZone)) {
    throw new InvalidInputError(
      `timeZone would rewrite ${updatedSide} into '${callerZone}' while the stored ${strandedLabel} stays ` +
      `in '${echoCallerText(desc.tzid, ZONE_ECHO_LIMIT)}' untouched — silently producing a two-zone event. ` +
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
        displayName: unwrapDisplayName(c.displayName) ?? 'Unnamed',
        url: c.url || '',
        description: c.description || undefined,
        // Guarded for the same reason `displayName` is, and it is the same defect: tsdav
        // guards `description` with `typeof === 'string'` but passes `calendarColor` through
        // raw (dist/tsdav.cjs:1003), so an empty `<calendar-color/>` parses to `{}` — truthy,
        // so `|| undefined` never fired and an object went out as a colour. Not routed
        // through `unwrapDisplayName`: this is a colour, and the number/boolean coercions
        // that make sense for a name would turn a malformed colour into a plausible one.
        // Trimmed on the way OUT as well as in the guard, so the two agree the way the name
        // path already does. No reachable input differs; consistency only.
        color: typeof (c as any).calendarColor === 'string' && (c as any).calendarColor.trim().length > 0
          ? (c as any).calendarColor.trim()
          : undefined,
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
      // BOTH SIDES through the same normaliser, so the comparison cannot drift. tsdav types a
      // calendar's name away from string (a calendar called "2026" arrives as the number
      // 2026), and normalising only the stored side turned a `calendarId` of 2026 — which
      // matched by raw equality before — into "Calendar not found". Note this is computed
      // separately from `requested` rather than replacing it: `requested` is what the
      // fail-closed presence test above reads, and an empty string must stay a value that
      // matches nothing, never one that widens the call to every calendar.
      // Accepted knowingly: because the CALLER's value comes through here too, a property
      // object such as `{_cdata: 'Work'}` now resolves where it was rejected before. The
      // inputSchema types `calendarId` as a string so no normal MCP call can produce that
      // shape, and lenient coercion of what a caller sends is this repo's documented posture
      // (docs/conventions.md) — so this is left as a widening, not guarded against.
      const requestedName = unwrapDisplayName(requested);
      const matched = targetCalendars.filter(
        c => c.url === requested ||
          (requestedName !== undefined && unwrapDisplayName(c.displayName) === requestedName)
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
    // The resolved (ICU-usable) form of `zone`, for the `timeZone`/`endTimeZone` fields
    // (#139) — those compare a stored TZID against the zone actually in force. Since #157,
    // `getDefaultTimezone()` is already guaranteed usable (an unusable configured or host
    // zone now stops the server at startup instead), so this call is a defensive no-op in
    // production rather than doing real work — kept because `resolveUsableTimezone` is the
    // one shared seam every other zone-resolving call site here already goes through, and
    // this stays consistent with them rather than being a special case that assumes its
    // input differently from the rest. `zone` above stays the raw configured value for
    // `coerceCalendarWindowStart`/`End`, which do not resolve it either — they hand it
    // through to the same `zoneOffsetMsAt`, whose own catch tolerates an unresolvable zone.
    // `sortEventsByStart` is handed `configuredZone`, the same value the window filter gets.
    // The raw value would order the same events the same way — canonicalising a usable zone
    // changes its spelling and not its offset, and for an unresolvable one `zoneOffsetMsAt`
    // falls back to the host zone `resolveUsableTimezone` falls back to — but passing one
    // resolved value removes the dependency on those two fallbacks coinciding rather than
    // documenting it.
    const configuredZone = resolveUsableTimezone(zone);
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
    // THE WINDOW THE CALLER ASKED FOR, kept separate from the widened range sent to the server
    // (#162). Everything caller-facing — the re-filter below and the clamp note — is derived
    // from these, never from `fetchOptions.timeRange`. Reading the request range back out of
    // `fetchOptions` is what would silently reinstate the fourteen hours of residue the
    // exact filter exists to remove.
    let trueWindowStart: string | undefined;
    let trueWindowEnd: string | undefined;
    // UNCONDITIONAL: there is no such thing as an unwindowed listing any more. A call naming
    // neither bound used to be sent with no time range at all, which meant no `expand` either
    // (tsdav drops it without one) — so the one call most likely to be asked "what is on?" was
    // the one call that answered with series masters at their original DTSTART instead of the
    // occurrences that actually fall on those days. Bounding it is also what makes the
    // expansion safe to turn on there: the window IS the range the server materialises over,
    // and an absent one is the unbounded case, not the empty case (#142).
    {
      let windowStart = start;
      let windowEnd = end;
      let invented: 'startDate' | 'endDate' | 'both' | undefined;
      if (!windowStart && !windowEnd) {
        // The same local-day rule a date-only `startDate` gets, so "no bounds" and "today's
        // date as startDate" cannot disagree about which day today is. The clock is read once
        // here, not per bound, so the two ends cannot straddle a midnight.
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

      // AFTER the invented half is filled in, not as its alternative. The old `else if` made
      // this check the both-bounds case only, and the comment below said so — but saturation
      // means a ONE-SIDED window can come out zero-length too: `startDate:
      // "9999-12-31T23:59:59Z"` leaves the invented month nowhere to go, so both ends land
      // on the same instant. Left to tsdav that was a plain Error (InternalError) or, worse, a
      // silently empty answer under a note claiming a month-long span had been searched.
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
        // An INVENTED bound needs its own sentence: the caller cannot "put startDate before
        // endDate" when they passed one or neither, so the ordinary advice would be
        // unfollowable. The no-bounds arm is separate again, because there is no bound of
        // theirs to blame — the default window starts at today and there is nowhere for a
        // month to go only when today itself sits at the end of the representable range.
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

      // The clamp reports the window the CALLER asked for, and it is checked against the
      // pre-widening bounds for the same reason: the caller has to be told when the window
      // they described was not the window searched, and the margin below does not change that
      // window — it changes what is asked of the server so the window can be honoured.
      if (invented || saturated.length > 0) {
        windowClamp = { invented, saturated: saturated.length > 0 ? saturated : undefined, start: windowStart!, end: windowEnd! };
      }
      trueWindowStart = windowStart;
      trueWindowEnd = windowEnd;

      // WIDENED BY THE REQUEST MARGIN ON BOTH EDGES — see MAX_UTC_OFFSET_MS for what the
      // server otherwise withholds. The widened bounds are saturated because 14 hours past
      // `9999-12-31` runs off the four-digit year and tsdav throws a plain Error over it.
      //
      // That saturation is deliberately NOT added to `saturated[]`. The disclosure array
      // reports what happened to bounds the CALLER named, and the caller's window is
      // untouched by the widening: the row the extra hours would have reached is one the
      // exact filter drops anyway, so nothing the caller asked for is lost by the shortfall.
      fetchOptions.timeRange = {
        start: shiftIsoMs(windowStart!, -MAX_UTC_OFFSET_MS),
        end: shiftIsoMs(windowEnd!, MAX_UTC_OFFSET_MS),
      };
      // Expansion is what makes a recurring event report the occurrence that actually falls
      // in the window instead of the series' original DTSTART (#64). tsdav only forwards
      // <C:expand> when a timeRange accompanies it, which is why this sits inside the same
      // branch rather than being set unconditionally. It expands over the range REQUESTED, so
      // widening that range is also what makes the server materialise the occurrences a
      // narrow window would otherwise see none of.
      fetchOptions.expand = true;
    }

    // Derived from the TRUE window, never from `fetchOptions.timeRange`. Both are always set
    // now that the window branch is unconditional; the undefined arms stay because the
    // declarations above are the only thing that says so and a type is not a guarantee.
    const windowStartMs = trueWindowStart === undefined ? NaN : Date.parse(trueWindowStart);
    const windowEndMs = trueWindowEnd === undefined ? NaN : Date.parse(trueWindowEnd);

    const allEvents: CalendarEvent[] = [];
    const denseSeries: DenseCalendarSeries[] = [];
    for (const cal of targetCalendars) {
      const objects = await client.fetchCalendarObjects({ calendar: cal, ...fetchOptions });
      for (const obj of objects) {
        // ONE structural extraction per resource, on whole content lines — never a `/m` regex
        // or a substring count, which a DESCRIPTION containing the text "BEGIN:VEVENT" defeats
        // (see docs/conventions.md). The list is counted here and then handed to the parser
        // rather than re-derived by it.
        const blocks = extractVEventBlocks(obj.data || '');
        if (blocks.length > CALENDAR_MAX_OCCURRENCES_PER_SERIES) {
          // NOT PARSED, and not truncated to its first N either: the first N of a hostile
          // series would fill the whole `limit` page and push every real event off it. The
          // title and id come from the first block, which is the same pair a row would have
          // carried and is two property reads rather than a parse of the whole payload.
          denseSeries.push({
            title: parseICalValue(blocks[0], 'SUMMARY') || 'Untitled',
            id: parseICalValue(blocks[0], 'UID') || obj.url || '',
            url: obj.url || '',
            // UNWRAPPED through the shared helper, the same one `list_calendars` and the
            // not-found error use. A DAV displayName arrives as a property object whenever the
            // element is empty or attribute-only, and stringifying one produced a visible
            // "[object Object]" in the disclosure — truthy, so the url fallback beside it never
            // ran. The helper answers "is there a name"; the url is the handle that always
            // exists when there is not.
            calendar: unwrapDisplayName(cal.displayName) ?? cal.url ?? '',
            occurrences: blocks.length,
          });
          continue;
        }
        // Every VEVENT in the blob, not the first: with `expand` a single resource carries
        // one block per in-window occurrence, so a first-match read drops all but one.
        // `expanded` is passed rather than sniffed — see parseCalendarObjects.
        //
        // `!!fetchOptions.expand` is always true now that the window branch above is
        // unconditional, and the defensive read stays for the same reason the
        // `trueWindowStart === undefined` arms above do: the only thing making it true is that
        // branch, and a reader who changes the branch should get the old behaviour here rather
        // than a hardcoded `true` that has quietly become a lie.
        for (const event of parseCalendarObjects(obj, { expanded: !!fetchOptions.expand, configuredZone, blocks })) {
          // A block that STILL CARRIES A RECURRENCE CARRIER is never dropped here, whatever
          // its dates say. This branch runs on an expanded query, so a surviving master means
          // the server declined to expand that resource — and the master then shows the
          // series' ORIGINAL DTSTART, which for a long-running weekly event is years before
          // the window. Judged on that date it fails the intersection test and the row
          // disappears: the filter would be turning a wrongly-dated row into a missing one, on
          // exactly the resource the server told us repeats. A recurrence carrier is proof
          // that DTSTART is not the only date this event has, so it is not a date the filter
          // is entitled to judge.
          //
          // BOTH CARRIERS COUNT, not just RRULE (#162). A series may list its occurrences as
          // RDATEs instead of stating a rule; such a master carries no RRULE at all, so an
          // RRULE-only guard read it as an ordinary one-off and dropped it on its original
          // DTSTART — the missing-event direction this guard exists to prevent.
          //
          // THE "WAS A WINDOW ASKED FOR" LEG IS GONE, because there is no longer a call
          // without one: a caller naming no bounds is given today plus a month (#142), so
          // `fetchOptions.timeRange` is set on every path through this method and testing it
          // here only asserted that. Removing it keeps the guard's remaining conditions about
          // what they are about — whether THIS block is one the window is entitled to judge.
          const provablyOutside = !event.recurrenceRule
            && !event.recurrenceDates
            && !eventIntersectsWindow(event, windowStartMs, windowEndMs, configuredZone);
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

    sortEventsByStart(allEvents, configuredZone);

    // `total` counts what was materialised, so an omitted series is NOT in it: the total is
    // "how many events matched, of which `limit` trimmed the rest", and folding in occurrences
    // no row represents would make a caller raise `limit` chasing rows that do not exist. The
    // note is what discloses the omission, and it is absent when there was nothing to omit.
    return {
      events: allEvents.slice(0, limit),
      total: allEvents.length,
      windowClamp,
      denseSeries: denseSeries.length > 0 ? denseSeries : undefined,
    };
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
    return parseCalendarObject(obj, { includeParticipants: true, configuredZone: resolveUsableTimezone(getDefaultTimezone()) });
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
     * IANA zone name for a designator-less `start`/`end` (fork issue #157). Omitted means the
     * account's configured zone is written — never floating; see docs/conventions.md for why
     * `create` defaults where `update` does not. `null` and an empty/whitespace string are
     * rejected (validateCallerTimezone) rather than read as "write floating", and so is a
     * value that already carries `Z`/an offset or is date-only (rejectTimezoneConflict).
     */
    timeZone?: string | null;
  }): Promise<CreateCalendarEventResult> {
    const client = await this.getClient();
    const calendars = await this.discoverCalendars();

    // RESOLVED like the read path, but not identically, and the difference is worth naming:
    // the read path `filter`s and queries EVERY calendar a name matches, while this `find`s
    // and writes to the FIRST. A display name is unique per account in practice, so the two
    // agree on every real input — but where they would not, a read unions the matches and a
    // create silently picks one. What an ambiguous write should do (refuse, or say which it
    // chose) is open; see issue #173. Do not "align" these by changing either behaviour here.
    //
    // What the two DO share, and what this comment used to be about: the same filtered
    // calendar list, the same trim, the same fail-closed treatment of an empty value, and the
    // same not-found error. They once shared only the error message, so an event could be
    // written into — and later deleted from — a collection no read tool would show, and
    // `" Work "` failed here while succeeding on a list.
    const requested = typeof event.calendarId === 'string' ? event.calendarId.trim() : event.calendarId;
    // Both sides normalised, for the reason spelled out on the read path: tsdav can type a
    // calendar's name away from string, so normalising only the stored side would reject a
    // numeric or boolean `calendarId` that used to match by raw equality.
    const requestedName = unwrapDisplayName(requested);
    const selectable = selectableCalendars(calendars);
    const targetCal = selectable.find(
      c => c.url === requested ||
        (requestedName !== undefined && unwrapDisplayName(c.displayName) === requestedName)
    );
    if (!targetCal) {
      // A calendarId that matches no calendar is caller-fixable: they re-issue the call
      // with an id or name from list_calendars. Shared with the read path so both state
      // the same rule — see calendarNotFoundError.
      throw calendarNotFoundError(event.calendarId, selectable);
    }

    const uid = `${Date.now()}-${Math.random().toString(36).slice(2)}@fastmail-mcp`;
    const now = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

    // --- timeZone validation (#157) ---
    // Offset-shape/resolvability rejection and the null/empty/whitespace fail-closed live in
    // validateCallerTimezone; the two conflict rules below (timeZone combined with a
    // Z/offset-designated or date-only value) need the start/end VALUES, so they run here.
    let callerZone: string | undefined;
    if (event.timeZone !== undefined) {
      callerZone = validateCallerTimezone(event.timeZone);
      rejectTimezoneConflict(event.start, 'start', callerZone);
      rejectTimezoneConflict(event.end, 'end', callerZone);
    }

    // The account's configured zone, resolved to a name ICU can actually use — `create`'s
    // default for a designator-less value with no caller `timeZone` (#157). Never floating:
    // a bare `2026-04-07T14:00:00` used to be written verbatim (a different instant for every
    // reader); this is the deliberate behaviour change docs/conventions.md documents.
    const configuredZone = resolveUsableTimezone(getDefaultTimezone());

    // Format start/end with all-day event support
    const startFormatted = formatDateTimeProperty('DTSTART', event.start, null, '\r\n', callerZone, configuredZone);
    const endFormatted = formatDateTimeProperty('DTEND', event.end, null, '\r\n', callerZone, configuredZone);
    const startLine = startFormatted.line;
    const endLine = endFormatted.line;

    // Time frame + ordering consistency. Classifying the serialized lines rather
    // than the raw inputs is the same thing the update path does, so create and
    // update reject an identical set of bad pairs. tzidSource is threaded through so a
    // frame-mismatch error names a DEFAULTED zone as defaulted, not as something the caller
    // wrote.
    validateDateConsistency(
      describeDateProperty(startLine, event.start, startFormatted.tzidSource),
      describeDateProperty(endLine, event.end, endFormatted.tzidSource)
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

    return {
      eventId: uid,
      start: classifyWrittenLine(startFormatted),
      end: classifyWrittenLine(endFormatted),
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
    // Explicit zone for a designator-less start/end (#157). Unlike create,
    // update NEVER defaults this when omitted — omitting it preserves
    // whatever the event already has (inherited stored TZID, or floating).
    // See validateCallerTimezone and the module-level reject* helpers above
    // the class for the full set of rejection rules this triggers.
    timeZone?: string | null;
  }): Promise<UpdateCalendarEventResult> {
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

    // --- timeZone validation (#157) ---
    // Runs before any patching so a rejection never leaves the object half-modified.
    let callerZone: string | undefined;
    if (fields.timeZone !== undefined) {
      callerZone = validateCallerTimezone(fields.timeZone);
      // timeZone with neither side supplied has nothing to qualify.
      // Still reachable — re-send start/end unchanged alongside timeZone to re-zone them.
      if (fields.start === undefined && fields.end === undefined) {
        throw new InvalidInputError(
          `timeZone was supplied ('${callerZone}') but neither start nor end was. timeZone only ` +
          "qualifies a start/end value being written in this same call — it cannot be applied to a " +
          "stored value on its own. Re-send start and/or end (even unchanged) alongside timeZone, " +
          "or drop timeZone."
        );
      }
      // timeZone can't be combined with a value that already names its own instant
      // (Z/offset) or has no time component (all-day).
      if (fields.start !== undefined) rejectTimezoneConflict(fields.start, 'start', callerZone);
      if (fields.end !== undefined) rejectTimezoneConflict(fields.end, 'end', callerZone);
      // A single-sided update that would re-zone one side while leaving the other
      // stranded in a DIFFERENT stored zone would silently produce a two-zone event.
      if (fields.start !== undefined && fields.end === undefined) {
        rejectStrandedZoneMismatch(originalVevent, 'start', callerZone);
      }
      if (fields.end !== undefined && fields.start === undefined) {
        rejectStrandedZoneMismatch(originalVevent, 'end', callerZone);
      }
    }

    // --- Patch fields ---
    let newStartFormatted: FormattedDateProperty | null = null;
    let newEndFormatted: FormattedDateProperty | null = null;
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
      // No defaultZone: update never defaults an omitted zone, only create does.
      newStartFormatted = formatDateTimeProperty('DTSTART', fields.start, originalVevent, lineEnding, callerZone);
      newStartLine = newStartFormatted.line;
      data = replaceICalProperty(data, 'DTSTART', newStartLine);
      timeChanged = true;
    }

    if (fields.end !== undefined) {
      newEndFormatted = formatDateTimeProperty('DTEND', fields.end, originalVevent, lineEnding, callerZone);
      newEndLine = newEndFormatted.line;
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
          describeDateProperty(startLine, newStartLine ? fields.start : undefined, newStartFormatted?.tzidSource),
          describeDateProperty(endLine, newEndLine ? fields.end : undefined, newEndFormatted?.tzidSource)
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

    return {
      eventId: existingUid,
      start: newStartFormatted ? classifyWrittenLine(newStartFormatted) : undefined,
      end: newEndFormatted ? classifyWrittenLine(newEndFormatted) : undefined,
    };
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
