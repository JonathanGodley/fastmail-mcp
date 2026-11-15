# Calendar reference

Fastmail does not currently expose calendar access via JMAP API tokens — the `urn:ietf:params:jmap:calendars` scope is not available because the JMAP Calendars specification is still an IETF Internet-Draft ([draft-ietf-jmap-calendars](https://datatracker.ietf.org/doc/draft-ietf-jmap-calendars/)). Fastmail has stated they will add JMAP calendar support once the spec becomes an RFC, but there is no public timeline.

However, Fastmail fully supports **CalDAV** for calendar access via `caldav.fastmail.com`. All calendar tools use CalDAV directly.

## Calendar invitations send mail

Naming attendees on an event is not a local write. When the event is created the server emails each attendee a meeting invitation from this account, and deleting that event later emails them a cancellation. Updating attendees does the same for whoever the change affects. None of this goes through a compose tool, so it happens with no `send_draft` call and produces no draft to inspect first.

This is the one hole in the draft-first posture, and it matters most if you are relying on tool names for permissions: gating `send_draft` gates every message you *compose*, and gates none of this. If a caller can reach `create_calendar_event`, it can cause mail to be sent to an address of its choosing.

Omit `participants` on create to write an event without notifying anyone. On update, omitting it keeps the existing attendees, and a change that affects them still mails them. The behaviour is the server's, not this server's — attendee scheduling is what CalDAV does — so there is no flag here that turns it off. An operator who needs the guarantee that nothing leaves the account has to deny `create_calendar_event`, `update_calendar_event` **and `delete_calendar_event`** alongside `send_draft`. Delete belongs on that list for a reason that is easy to miss: it takes no `participants` of its own, so it reads like a purely local removal, but deleting an event someone else is already on mails all of them. Deleting one occurrence of a repeating event (`occurrence`) never sends mail, because a series with attendees or an organizer is refused.

## What an `eventId` resolves to

`get_calendar_event`, `update_calendar_event` and `delete_calendar_event` all resolve an id through one lookup, so these rules hold identically on the read and on both writes ([#137](https://github.com/JonathanGodley/fastmail-mcp/issues/137)).

- **Two forms, and both are always tried.** An `eventId` may be the event's `id` (its `UID`) or the `url` from a `list_calendar_events` row; the lookup asks every calendar for that UID **and** resolves the url form, then unions what comes back.
- **The UID is asked of the server**, one CalDAV `calendar-query` per calendar. If the server refuses that query, **the call fails**; there is no fallback to a full scan.
- **A url is resolved against this account's own calendars, and nothing else.** A url pointing anywhere else is never requested and comes back as not-found. See [Security model](../security-model.md).
- **An empty or whitespace-only `eventId` is rejected** rather than searched for.
- **A record the server describes incompletely does not resolve.** A resource with no payload, nothing parseable in it, or no entity tag to write against reads as not-found, so this server never changes or destroys a record it could not read whole.
- **What a resource is NAMED does not affect whether it is visible** ([#191](https://github.com/JonathanGodley/fastmail-mcp/issues/191)): a `.ICS`, extensionless or bare-UUID resource is listed, read, updated and deleted like a `.ics` one.

## Start/end agreement

An event's `start` and `end` are only meaningful together, so both `create_calendar_event` and `update_calendar_event` check the pair that will actually be written and reject it if it doesn't hold up. Two rules:

1. **Same form.** Both must land in the same one of: date-only (`2026-03-20`), a zone-designated time (`2026-03-20T09:30:00Z` or `2026-03-20T09:30:00+10:00`, both stored as UTC), or a time carried by a named `TZID` (the account's configured zone by default, an explicit `timeZone`, or an inherited stored zone — see [Writing calendar times](#writing-calendar-times)). A floating time with **no** zone at all can only arise on `update_calendar_event`, on an event that is already floating, where a designator-less value with no stored `TZID` to inherit stays floating; `create_calendar_event` always resolves a designator-less value to a zone, so this form is unreachable there. A mixed pair has no single duration — `DTSTART:20260320T093000` (no zone) beside `DTEND:20260320T093000Z` (UTC) is zero-length in UTC and ends ten hours before it starts in UTC+10, and renders differently for every attendee.
2. **`end` after `start`.** Equal values are rejected too; RFC 5545 §3.8.2.2 requires DTEND to be strictly later. For all-day events, remember DTEND is exclusive.

On `update_calendar_event` the comparison is against **the value it will sit beside** — the other one you passed, or the stored one you left alone. So changing only `start` is fine when it stays in the same form and still lands before the stored `end`, but **moving an event to a different day, or converting one side to UTC, means passing both `start` and `end`.** A floating time on an event that carries a `TZID` inherits that `TZID`, so it agrees with its partner and is accepted. An event stored with a `DURATION` instead of an end keeps it when you change only `start`, and a date-only `start` is then refused beside a `DURATION` with a time part (`PT1H`), which RFC 5545 §3.6.1 forbids; pass `end` as well, which replaces the `DURATION`.

Two `TZID`-bearing values in *different* zones — a flight that departs in one zone and lands in another — are a legal shape and are accepted, and **their order is judged on instants, not on their wall clocks** ([#140](https://github.com/JonathanGodley/fastmail-mcp/issues/140)). **This shape only ever arises on `update_calendar_event`**, editing a record that already holds it: `create_calendar_event` takes no per-value zone and writes a designator-less pair in the configured zone, so both of its ends are always in one zone. The text of a cross-zone pair says nothing about its order in either direction: Rome 10:00 to New York 08:00 the same day is a real four-hour flight and is written, while New York 07:00 to Rome 08:00 ends four hours before it starts and is rejected with the same "end after start" message as any other backwards pair. The one case that still stands down is a `TZID` this server cannot resolve at all — a vendor name such as `AUS Eastern Standard Time`, which real records do carry — because there is then no instant to compare and refusing would reject a record your account already holds. Both rejections are `InvalidParams` and name both values and both forms.

## Accepted date spellings

`start` and `end` are read as strict ISO-8601: `2026-04-07`, `2026-04-07T14:00:00`, `2026-04-07T14:00:00Z`, or `2026-04-07T14:00:00+10:00`, whose offset may drop its colon (`+1000`); surrounding whitespace is trimmed. The time must be a real time of day, hours `00`-`23` and minutes and seconds `00`-`59`, so `24:00:00` is rejected here: RFC 5545 has no hour 24, so pass `00:00:00` on the next day instead. Anything else, such as `2026/04/07` or `April 7 2026`, is rejected as `InvalidParams` rather than guessed at, as is a day that does not exist in its month (`2026-02-31`, in either the date-only or the datetime form). A time outside years `0001`-`9999` is rejected too, whether written (`0000-06-01`) or reached by converting an offset to UTC (`9999-12-31T23:00:00-05:00`): an iCalendar year has four digits, and the Gregorian calendar has no year 0. `create_calendar_event` and `update_calendar_event` apply the identical rule.

## Calendar text size limits

Every text field that ends up in the event (`title`, `description`, `location`, and each participant's `name` and `email`) is sent to Fastmail and stored in the event, so both `create_calendar_event` and `update_calendar_event` bound the size of what one call writes, up front:

| Bound | Limit |
| --- | --- |
| Any single field (`title`, `description`, `location`, participant `name`, participant `email`) | 64 KB |
| Number of `participants` | 500 |
| All of that text combined, in one call | 256 KB |

Oversized input is **rejected, never truncated**. The rejection is `InvalidParams` and names the offending field, its actual size and the limit, so one retry with a shorter value fixes it.

## Calendar window dates

`list_calendar_events`' `startDate` and `endDate` accept a date (`2027-03-01`) or a datetime (`2027-03-01T09:00:00`, `2027-03-01T09:00:00Z`, `2027-03-01T09:00:00+10:00`), with surrounding whitespace trimmed; an offset may drop its colon (`+1000`). They are a little looser than `create_calendar_event`'s `start`/`end` (see [Accepted date spellings](#accepted-date-spellings)): the seconds may be left out (`2027-03-01T09:00Z`), fractional seconds are accepted and dropped (`2027-03-01T09:00:00.5Z`), and the `Z` may be lowercase. `2027/03/01`, `"March 1 2027"` and a day its month does not have (`2027-02-30`) are all rejected naming the parameter, rather than guessed at or silently rolled into the next month — either would move the window off the days you asked about. **What a date-only value MEANS differs from the write tools**, and it is read in your timezone rather than UTC; both are covered below.

**A date is a LOCAL day.** `startDate: "2026-08-12"` means that calendar day where you are, not the UTC day. The zone is `FASTMAIL_TIMEZONE` (or `USER_CONFIG_FASTMAIL_TIMEZONE`), falling back to the zone this server runs in: the same single setting every email's `date` is rendered in (see [Configuration](configuration.md#configuration) for what a valid value is). Where the host zone is itself not a usable IANA name, dates fall back to UTC, with a one-time startup warning naming the rejected host zone. A datetime with **no** `Z` and **no** offset (`2026-08-12T09:00:00`) is local time too. A value carrying `Z` or a numeric offset means exactly the instant it names and is never re-read.

The window runs from `startDate` **inclusive** to `endDate` **exclusive**, and there is one deliberate asymmetry between them:

- A date-only **`startDate`** is local midnight at the start of that day.
- A date-only **`endDate`** covers the **whole of that day** — `2027-03-10` runs through to local midnight at the start of the 11th. So a single-day query is `startDate` and `endDate` on the *same* date. **This is deliberately the opposite of `create_calendar_event`'s date-only `end`**, which is exclusive at the *start* of that day (RFC 5545 DTEND, so a one-day event on April 1 takes `end: "2026-04-02"`).
- A **datetime** `endDate`, with a designator or without, is taken literally as the exclusive end — no day is added, because a wall clock names a time of day rather than a day.
- A day that a daylight-saving change makes 23 or 25 hours long is still covered end to end; the bound advances by a whole *local* day, not by a fixed 24 hours. Where the change happens *at midnight* (Santiago, Havana) the day's own first or last hour is one the clocks skipped — a skipped wall clock resolves **forward by the length of the gap**, so the window still covers every instant of the day that exists (that lands on the moment the clocks jumped to only for a wall clock at the very start of the gap, and past it otherwise). A wall clock the change makes happen twice resolves to the earlier of the two.
- The time of day has to exist: `2026-08-12T25:00:00`, `12:75:00` or `99:99:99` is rejected rather than rolled into another day, as `create_calendar_event` rejects them on a write. (`24:00:00` is accepted here as the end of the day, which is what ISO-8601 says it means, and read as midnight starting the next day. The write tools reject it, because RFC 5545 has no hour 24.)

A `startDate` that is not before the resulting `endDate` is rejected as `InvalidParams`, quoting the values you passed, the range they resolved to, and the zone they were read in. Two bounds that resolve to the *same* instant are rejected as a zero-length window and told so in those words, rather than being shown a range whose two ends print identically.

**Omitting one bound clamps the window to 31 days.** `startDate` alone runs a month forward from it, `endDate` alone a month back, and the response says which range it searched in a trailing `Note:`. A window whose bounds you name yourself is never clamped, however long it is - the span is yours to choose. A bound that resolves past the ends of the four-digit-year range — the invented half, or one you named, since the local-day rule resolves your value through a zone and an offset alone can push `9999-12-31` over — saturates at that edge rather than running past it. So `startDate: "9999-12-30"` on its own is answered rather than erroring on an unrepresentable date, and a saturated bound **you** named is disclosed in the same trailing `Note:`. (`endDate` alone runs the other way and cannot saturate at that end.) Where saturation collapses a one-sided window to zero length, the call is rejected as `InvalidParams` rather than answered with nothing.

## Calendar times: a name, never an offset

This server never puts a UTC offset in a `start`/`end`, and never asks you to compute one yourself. A calendar time comes back as a bare local wall clock (`2026-04-20T10:00:00` — no `Z`, no offset) plus, when it matters, a separate `timeZone` naming the IANA zone that wall clock is in — never as `2026-04-20T10:00:00+12:00`.

`timeZone` describes `start`:

| `start` is stored as | `timeZone` |
| --- | --- |
| A named zone (`TZID=...`) **matching** the account's configured zone | *(omitted)* |
| A named zone **differing** from the configured zone | the zone's IANA name, e.g. `"Pacific/Auckland"` |
| Floating — no `TZID` and no `Z` (RFC 5545 §3.3.5: a different instant for every reader) | `null` |
| `Z`-designated (a UTC instant), or date-only (all-day) | *(omitted)* |

`timeZone` is **omitted**, not set to the configured zone's own name, for the overwhelming majority of rows — most events are already in the account's configured zone (see [Calendar window dates](#calendar-window-dates)), so read "no `timeZone`" as "the configured zone". A `Z`-designated or all-day value is also omitted rather than set to `null`: both already fully describe themselves, and `null` there would wrongly assert "floating".

`endTimeZone` describes `end`, but only **relative to `start`**: omitted whenever `end`'s zone matches `start`'s (the ordinary case), and otherwise set by the same rule as the table above. When `start` is absent entirely — no `DTSTART` at all — `end` is compared against the account's configured zone instead, the same fallback `timeZone` itself uses. A start and end in two different named zones, such as a flight's, is a legal shape (see [Start/end agreement](#startend-agreement)). `endTimeZone` is also omitted when `end` is absent, or was computed from a stored `DURATION` rather than a real `DTEND`: a computed end is in `start`'s zone or is a `Z` value, which describes itself, so there is nothing to disagree about.

A `DURATION`-computed `end` is the wall clock in `start`'s zone once the duration has run, per RFC 5545 §3.3.6: days and weeks keep the time of day, hours, minutes and seconds are elapsed time. So `PT6H` from 23:00 on the night clocks go forward ends at 06:00, not 05:00. When the end falls in the second pass through a repeated hour of `start`'s zone, no wall clock there names it (a repeated clock reads as its first occurrence), so it is returned as a UTC `Z` value. A floating `start`, or one whose zone this server cannot resolve, is plain clock arithmetic. An end that would fall outside years 0001-9999 is omitted.

Where the underlying value comes from: `list_calendar_events` always expands (there is no unwindowed call - see [Calendar window dates](#calendar-window-dates)), and expansion normalises most in-window *occurrences* to a UTC instant server-side, so those rows carry `Z` and no `timeZone`; an all-day row, and any row expansion did not touch, keeps whatever form the stored property had, judged by the same table. `get_calendar_event` is the call that returns the stored property exactly as written, because it fetches the resource itself rather than a window over it.

Non-IANA names pass straight through rather than being rejected or resolved. A calendar written by a non-Fastmail client can carry a Windows zone name (`AUS Eastern Standard Time`) in its `TZID`; this server reports it verbatim in `timeZone`/`endTimeZone` and sorts that event by the account's configured zone instead of that name (below), leaving interpreting the name itself to you.

**Sorting.** `list_calendar_events`' result order compares instants, not the text of the dates, since a calendar mixes floating wall clocks, UTC instants and named zones. An event whose own `timeZone` names a zone this server can resolve sorts in **that** zone; the account's configured zone is only the fallback, used for a genuinely floating `start` (there is no zone to sort it in, so your own clock is the least-wrong guess) and for a `timeZone` this server cannot resolve (a Windows name, passed through rather than rejected — see above).

## Writing calendar times

`create_calendar_event` and `update_calendar_event` never write an offset either, and neither asks you to compute one: a `start`/`end` with no `Z` and no date-only marker is a bare wall clock, and `timeZone` names the IANA zone (e.g. `"Australia/Sydney"`) it's written in. `timeZone` only **qualifies** a designator-less `start`/`end` — the one shape with no zone of its own already. It cannot combine with a value that carries `Z`/an offset, or with a date-only value, since both already name themselves; either combination is rejected as `InvalidParams`, naming which argument to drop.

**A `TZID` is never written alone**: every timed write carries a `VTIMEZONE` for each zone it references, generated from Node's own ICU data, within bounds the tool descriptions state ([#166](https://github.com/JonathanGodley/fastmail-mcp/issues/166)); see [docs/conventions.md](../conventions.md).

**`timeZone` must contain a region-qualifying slash** (`Continent/City`), with exactly one exception: `"UTC"` itself, matched case-insensitively. A bare abbreviation or alias - `"EST"`, `"NZ"`, `"PST"`, `"MST"`, `"GMT"`, `"Zulu"` and others like them - is rejected even though each one resolves to a real zone, because the resolution is ambiguous: `"EST"` names a fixed-offset zone with no daylight saving, **not** US Eastern, and the others are exactly as easy to misread. Write `"Pacific/Auckland"` rather than `"NZ"`, `"America/New_York"` rather than `"EST"`. There is no safe-list of "harmless" abbreviations - the rule is the same for all of them.

**An `Etc/GMT` name is accepted, and its sign is inverted.** `"Etc/GMT+10"` and `"Etc/GMT-5"` are real IANA zones, so `timeZone` and `FASTMAIL_TIMEZONE` take them, but they follow the POSIX convention: `Etc/GMT+10` is **UTC-10:00** and `Etc/GMT-10` is **UTC+10:00**. Wherever this server shows one - a write confirmation (`Written in zone "Etc/GMT+10" (UTC-10:00; the Etc/GMT sign is inverted)`), a start/end refusal, a date message naming the configured zone, the configured zone in the tool descriptions - the real offset is shown beside it. On a read, `timeZone`/`endTimeZone` stay the bare zone name, so you can pass it back, and the response ends with a `Note:` line giving each such zone's real offset.

`timeZone` accepts a leading `/` on a TZID (RFC 5545 §3.2.19 - a zone registered by its own creator rather than plain IANA form): it's accepted and stripped for you, then checked against the slash rule above. What actually gets written is `timeZone`'s **canonical IANA spelling**, which can differ in case, or via a link, from what you passed - `"us/pacific"` is written as `"America/Los_Angeles"`, and `"AUSTRALIA/SYDNEY"` as `"Australia/Sydney"`. A non-IANA name a read can hand back - a Windows zone id such as `"AUS Eastern Standard Time"`, from an external invite - is rejected by `timeZone`; there is nothing to canonicalise it against. Omitting `timeZone` on `update_calendar_event` preserves the stored zone exactly as it is, which is the way to round-trip an event carrying one of those.

**Create and update treat an omitted `timeZone` differently, on purpose:**

- **`create_calendar_event`**: an omitted `timeZone` writes the account's configured zone (see [Calendar window dates](#calendar-window-dates)), so a designator-less `start`/`end` always resolves to a concrete zone.
- **`update_calendar_event`**: an omitted `timeZone` never overrides what the event already says. A designator-less value inherits the event's existing `TZID` when it has one; on an event whose stored start is floating it stays floating, which keeps that event's own form; on any other event (UTC or all-day) it is read in the configured zone, as on create, and never written floating. If the stored `DTSTART` or `DTEND` it would inherit from is malformed (a quote in its parameters that does not open or close a value, whitespace beside a parameter name such as `TZID =Europe/Paris`, a line with no value, or a control character in its `TZID`), the zone cannot be read reliably, so the update is refused as `InvalidParams` rather than written, and the error names that property. Pass **both** `start` and `end` with `timeZone`, both with `Z`/an offset, or both date-only: none of those inherits anything. An unquoted `TZID` containing a comma (`Eastern Time (US & Canada), Bogota`) and a quoted Outlook-style one are not malformed in this sense and inherit as usual.

**`timeZone` is rejected outright** (never silently ignored) in these cases, each naming the argument to drop:

- `null`, an empty string, or a whitespace-only string — neither tool can write a genuinely floating event through this parameter. Note that a read returns `timeZone: null` for a floating event, so a read-modify-write call that echoes that straight back will hit this rejection; pass no `timeZone` at all to leave an already-floating value floating (`update_calendar_event` only - `create_calendar_event` has no existing value to leave alone, so an omitted `timeZone` there always writes the configured zone, per above).
- On `update_calendar_event`, `timeZone` with **neither** `start` nor `end` supplied — it has nothing to qualify. Still reachable: re-send `start` and/or `end` unchanged alongside `timeZone` to re-zone them.
- On `update_calendar_event`, `timeZone` with **only one** of `start`/`end` supplied, when the side you left alone is already stored in a *different* named zone. Re-zoning just one side would silently strand the other, producing a two-zone event nobody asked for. Pass **both** `start` and `end` (re-sending the untouched one unchanged) to change the zone, or omit `timeZone`.

**The response states what was actually written**: both tools report the zone (or `UTC`/all-day/floating, as applicable) that ended up on `start` and `end`.

## Busy and free: an all-day event is free, a timed event busy

The caller-facing values are **`busy`** and **`free`** — what the `transparency` parameter takes on `create_calendar_event` and `update_calendar_event`, and what the `transparency` field reports on a read. (iCalendar stores this as `TRANSP:OPAQUE`/`TRANSP:TRANSPARENT`; those spellings are not accepted as parameter values.)

**On create the default follows the event’s frame, matching the Fastmail client**: an all-day event is created **free**, a timed event **busy** ([#195](https://github.com/JonathanGodley/fastmail-mcp/issues/195), measurements in [docs/fastmail-action-availability.md](../fastmail-action-availability.md)). Pass `transparency` to override that on either frame ([#194](https://github.com/JonathanGodley/fastmail-mcp/issues/194)) — `"busy"` for an all-day event that should block the whole day, `"free"` for a timed one you stay available during.

**`update_calendar_event` changes free/busy only when you ask it to.** `transparency` sets or replaces the value; `clearFields: ["transparency"]` removes the property, which leaves the event reading as busy without saying so. Omit both and the stored value survives everything else the call does — **including a change of frame, in either direction**. An all-day event you convert to a meeting stays marked free, and a timed event you convert to all-day keeps whatever it had, so pass `transparency` alongside the new dates when the converted event should read differently.

## One `eventId` can name more than one record

A `UID` is unique inside **one calendar**, not across the account, so two calendars can hold the same id. The three tools answer that differently, and the split is deliberate ([#101](https://github.com/JonathanGodley/fastmail-mcp/issues/101)).

- **`update_calendar_event` and `delete_calendar_event` refuse it.** The error names every copy - the calendar it sits in and its own resource `url` - and nothing is written or destroyed.
- **`get_calendar_event` answers, and discloses.** It returns the **first** copy in the lookup's stated order - the copy you addressed if you passed a url, otherwise the first found in discovery order. It adds an `otherCopies` array naming each other copy's calendar and `url`, and carries a trailing `Note:` line saying how many records the id names and what the two write tools will do with it.

**The way through is the `url`.** Every one of these tools accepts a resource `url` wherever it accepts an id, and a url **addresses** exactly one record: it is matched against this account's discovered collections and resolves to a single resource. Where an id is ambiguous, that is what tells the copies apart and what gets a write through.

**The one exception: a url that is also another record's UID.** Both forms are always tried, so a record whose UID literally spells another record's `url` puts two records in one result. `get_calendar_event` returns the record the url actually resolved to - the one you addressed - first, and lists the other in `otherCopies`. The write tools refuse, and `get_calendar_event`'s trailing `Note:` says so: `list_calendar_events` shows the other record's id as that very url, so a caller passing it back may mean either. The refusal names both records' urls and quotes the addressed record's own UID. Pass that UID to act on the record at the url, or the other record's url to act on the other. Where that UID would not reach the record alone - it has none, its UID is that same url, its UID is another resource's url, or another record shares it - the refusal offers no id and says so: change that record in the Fastmail web interface. A url-shaped id with no resource at that address addresses nothing and resolves by UID like any other id.

**The accepted consequence.** Whoever sends this account an invitation chooses the `UID` it arrives under, so a stranger who knows an event's id can freeze its writes by minting a duplicate in a calendar you share with them.

The ambiguity refusal is raised **before** the repeating-event refusal below.

## Repeating events: only one occurrence can be deleted

`update_calendar_event` **refuses** any `eventId` that resolves to a repeating event, and `delete_calendar_event` refuses to delete the whole series. There is no flag, parameter or confirmation that overrides either, and the error message says so. `get_calendar_event` is read-only and still works on the same id.

**What counts as repeating** is read off the stored record, not off the row you passed, and any one of four markers is enough: an `RRULE`; an `RDATE`, which is how a series **lists** its occurrences instead of stating a rule ([#162](https://github.com/JonathanGodley/fastmail-mcp/issues/162)); more than one `VEVENT` in the record, since one CalDAV resource is one UID; or a `RECURRENCE-ID` on any block, which makes that block an edited occurrence. The test fails **closed**.

`create_calendar_event` cannot make a repeating event, so this server will not rewrite or destroy a series it could not put back.

### Deleting one occurrence

`delete_calendar_event` with `occurrence` deletes **one** occurrence ([#146](https://github.com/JonathanGodley/fastmail-mcp/issues/146)). Pass the series id (or the row's `url`) as `eventId`, and as `occurrence` the row's `recurrenceId`, or its `start` when the row has no `recurrenceId` (the series' first occurrence has none). The value takes the spellings `create_calendar_event`'s `start` takes: a date-only value for an all-day series and only for one, a `Z` or offset instant, or a time with no `Z`/offset, read in the series' own zone.

The series gains one `EXDATE`, in the form the Fastmail client writes when you delete one occurrence there ([measured](../fastmail-action-availability.md#editing-one-occurrence-of-a-series)): the master's own zone and wall clock for a zoned series, `VALUE=DATE` for an all-day one, a `Z` instant for a UTC one. Nothing else in the record changes except its `DTSTAMP` and `LAST-MODIFIED`: other exclusions, alarms, time zone definitions, edited occurrences and properties this server does not know all stay as they were.

The calendar server confirms the occurrence twice, by expanding the series around it: before the write, that it exists, and after, that it is gone. This server never expands the rule itself. If the write lands but the occurrence is still listed, or the read after the write fails, the call fails with an error quoting the `EXDATE` written. The write is conditional on the record being unchanged since it was read.

The response carries the whole record as it was before the write, raw, as `previousICalendar`. That keeps visible what the write took away, and the excluded start is enough to make the occurrence again as a single event with `create_calendar_event`. It does not make the delete undoable: no tool here removes an exclusion, so the occurrence cannot be put back *into* the series.

Refused, with nothing deleted:

- `occurrence` on an event that does not repeat. It never falls back to deleting the whole event.
- A series with attendees or an organizer on any block, an alarm's attendee included. Whether a series write may mail attendees is undecided, so this path sends no mail at all.
- A series that is not one master carrying one `RRULE` (plus edited occurrences), or that has an `RDATE` or an `EXRULE`, a start with no time zone, a time zone this server cannot resolve, or an edited occurrence whose `RECURRENCE-ID` cannot be read or carries `RANGE`.
- An occurrence that has been edited on its own. Its content is not the series', and what to do with its block is undecided.
- A value that is not an occurrence the calendar server lists for the series, which includes one already deleted; a value away from the series' time of day; a date-only value against a timed series, or a timed one against an all-day series; and a time in the second pass through an hour the clocks repeat.
- In a series with a fixed number of occurrences or an end date: its first occurrence, and any occurrence once the first is gone. A surviving first occurrence is how the server proves such a series is not left empty.

What to do instead, for everything refused: use the Fastmail web interface, whose occurrence picker offers "This event only" and "All occurrences" ([measured](../fastmail-action-availability.md#editing-a-whole-series)). Whole-series delete with a recovery artifact, a recurrence-rule parameter on create and series-aware update are still open on [#146](https://github.com/JonathanGodley/fastmail-mcp/issues/146).

Single (non-repeating) events are unaffected.

## Calendar known limitations

- **Recurring events, on read**: `list_calendar_events` always expands recurrences, while `get_calendar_event` returns the series master at its original date. How rows are windowed, and the cases that put a row outside the window or keep one out of it, are in the `list_calendar_events` and `get_calendar_event` entries in [tools.md](tools.md#calendar-tools).
- **Recurring events, on write**: only deleting one occurrence is supported — see [Repeating events: only one occurrence can be deleted](#repeating-events-only-one-occurrence-can-be-deleted) above.
- **Attendee parameters**: RSVP, ROLE, CUTYPE and other attendee parameters are parsed on read but not settable on create/update — only `email` and `name` are accepted.
