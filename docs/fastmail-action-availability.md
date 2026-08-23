# What the Fastmail client offers and writes

Fastmail's client is the only written-down statement of Fastmail's own semantics. JMAP permits far
more than the client does, Cyrus (the server Fastmail runs) implements almost none of the policy,
and Fastmail's own MCP publishes tool *descriptions* rather than behaviour. So when this server
needs to know what an action means — not what the protocol allows — the answer comes from measuring
the client.

This file records those measurements. Extend it by measuring a screen, never by inferring from a
role's name or from what the protocol permits.

## How these were measured

Two independent methods, and a row is only trusted where they agree:

1. **The client's own action lists** — the toolbar plus the **More** menu, read off screenshots of a
   real account, one screen at a time.
2. **Live JMAP diffs** — perform the action in the client, then read the message's `mailboxIds`
   before and after, plus the per-mailbox `totalEmails` counts.

Everything below was measured on **one account, in labels mode**, in August 2026. Folders mode was
not measured and is called out where it would differ.

### Conversation grouping changes the unit, not the rule

Every tool in this server acts on a **message**. The client's unit of action is a *per-user display
setting*: with conversation grouping on the toolbar acts on the whole conversation, with it off the
toolbar acts on a single message. Neither is the client's behaviour - each account is configured one
way or the other, so there is no default mode to treat as canonical and no way to infer which one a
caller is looking at. Both were measured on the same three-message thread, on 16 August 2026.

**Grouped.** Every message filed in `{Inbox, "Pending Orders"}`, archived with one click on the
thread. All three came back filed in `{"Pending Orders"}` alone: none kept an Inbox membership, and
**none gained Archive**.

**Individual.** The same thread restored to `{Inbox, "Pending Orders"}`, with a single message
archived. That message came back `{"Pending Orders"}`; the other two were left in
`{Inbox, "Pending Orders"}`, untouched.

Three things follow.

1. **The rule is the same in both modes**, and it is the rule in the effect table below: remove the
   Inbox membership, keep everything else, add Archive only when removing the Inbox would leave the
   message filed nowhere.
2. **A grouped Archive is not a move to Archive.** "Move every message in the thread to the Archive
   folder" is refuted outright - the label survived and Archive was never added.
3. **Grouping decides the fan-out, and only the fan-out.** Grouped, the rule is applied to every
   message in the conversation. Individual, it is applied to the one message and its siblings are
   left alone.

**What that means for this server.** `archive_email` acts on the ids it is given, which reproduces
the ungrouped behaviour exactly, siblings included. Against a grouped account it *under-archives*:
called with one message id it clears that message and leaves the rest of the conversation in the
Inbox, so the thread goes on showing in the Inbox list.

Since the setting is per-user and undetectable here - the same blind spot as labels-versus-folders
mode (fork #122) - a caller cannot resolve this by knowing the mode, only by knowing the intent. If
the intent is to archive a conversation, pass every message id in it; `get_thread` lists them.

**Residual.** Every message in the grouped fixture had identical filing, so it does not separate
"apply the per-message rule to each member independently" from "remove the Inbox from every member,
and add Archive only if the *thread* would otherwise be filed nowhere". The two differ only for a
thread whose members are filed differently from one another, say one message in Inbox plus a label
and another in the Inbox only. The individual-mode result makes the per-message form the better
supported reading, and it is the one implemented.

## Availability: where Archive is offered

Seven of the account's nine role mailboxes were measured. `archive` and `memos` are marked
**unmeasured** rather than left blank, because a blank cell reads as "checked, absent" to the next
person.

| Message state / view | Actions offered | Archive? |
| --- | --- | --- |
| Inbox (`inbox`) | Archive · Delete · Snooze · Labels · Move to · More{Mark unread, Pin, Notify, Mute, Forward as attachment, Report spam, Report phishing} | **yes** |
| Inbox + Trash, viewed from the Inbox (see the caveat below) | Archive · Delete · Snooze · Labels · Move to · More | **yes** |
| Sent (`sent`) | Delete · Snooze · Labels · Move to · More{Mark unread, Pin, Notify, Mute, Forward as attachment} | no |
| Drafts (`drafts`) | Discard Draft · Snooze · Labels · Move to · More{Pin, Notify, Mute} | no |
| Scheduled (`scheduled`) | Cancel send · Labels · More{Mark unread, Pin, Notify, Mute} | no |
| Snoozed (`snoozed`), viewed from inside a label | Remove label · Delete · Snooze · Labels · Move to · More{Mark unread, Pin, Notify, Mute, Forward as attachment} | no |
| Spam (`junk`) | Delete permanently · Not spam · Move to · More | no |
| Trash (`trash`) | Delete permanently · Undelete · Move to · More | no |
| A label / All mail (`role: null`), message carries no other role | as Inbox, plus Remove label | yes, and it no-ops if the message is not in the Inbox |
| A label / All mail, message ALSO in `snoozed` | Remove label · Delete · Snooze · Labels · Move to · More | no (this is the `snoozed` row above, measured from inside a label) |
| A label / All mail, message also in one of the other five refusing roles | **unmeasured** | unmeasured |
| `archive`, `memos` | **unmeasured** | unmeasured |

The toolbar is therefore neither purely per-view nor purely per-message-role: a snoozed message
hides Archive even when viewed from inside a label.

That `snoozed` row is the ONLY refusing role measured from inside a label view. Whether a label
view also hides Archive for a message that is in Sent, Drafts, Scheduled, Spam or Trash is
unmeasured, and is left that way on purpose: generalising from the one measured role to the other
five would be inferring a view from a role's name, which the extension rule at the end of this
file forbids. This matters beyond bookkeeping, because `src/jmap-client.ts` cites this table as
the evidence that its refusal set is exactly the set that was measured.

**Role mailboxes are not uniformly exclusive.** Two real messages on this account are filed in both
`snoozed` and `sent`, so "which refusal applies" is a live case rather than a hypothetical. Inbox
and Trash are the two that do behave exclusively.

When a message is in two refusing roles, this server picks the first in the fixed order
`trash, junk, drafts, scheduled, sent, snoozed`, so the `snoozed`+`sent` message above is refused
as **sent**. That ordering is a tiebreak for a deterministic message, not a claim about which
state the client considers primary - nothing here measures that. Note the consequence: the
message gets the Sent refusal, which names `move_email`, rather than the Snoozed one, which says
this server cannot unsnooze. The more actionable sentence wins, and it is the less relevant one.

### Caveat on the Inbox + Trash row

That row was produced artificially, by cross-filing a message into Inbox and Trash with
`add_labels` and opening it from the Inbox. Archive is offered, and Trash is not shown as a chip -
only `Inbox ×`.

But **the client cannot produce that state**. Fastmail's own Delete is a whole-value replace: the
same message, deleted from the client, came back as `roles: ["trash"]` with the Inbox membership
gone. Trash is a folder in Fastmail's model, not a label. Corroboration: all 37 messages then in the
Inbox carried `roles: ["inbox"]` and nothing else.

So the toolbar shown for Inbox + Trash is the Inbox view's *default*, not a considered answer, and
it cannot settle whether the Inbox test should come first. This server puts the Inbox test first
anyway — that is **our** decision, justified by being what the client rendered when confronted with
the state and by being the least surprising reading of "archive this", not by parity. That this
server can manufacture the state at all is tracked separately (fork #124; the label tools joining
the role-mailbox guard is #133, the general destination guard is #43).

## Availability: what the two message-action pickers offer

The toolbar's **Labels** and **Move to** entries each open a picker. Both were opened on the **same
single message**, on 22 August 2026, on the same account and in the same labels mode as everything
above. The rows below are what each open picker contained, read off the screen. **Nothing in either
picker was clicked**, so this section measures what is offered and not what choosing it does.

| Entry | In **Labels** | In **Move to** |
| --- | --- | --- |
| `Inbox` | yes, with a checkbox, ticked for a message in the Inbox | yes |
| `Archive`, `Drafts`, `Sent`, `Spam`, `Trash` | **no** | yes |
| `Snoozed`, `Scheduled` | **no** | yes, but rendered **greyed out** |
| The account's user labels | yes, each with an unticked checkbox | yes |
| Create affordance | "Create label…" | "Create…" |

**Move to** lists the eight role folders first, in the order `Inbox, Snoozed, Archive, Drafts,
Scheduled, Sent, Spam, Trash`, then the user labels, then the create affordance. Each role folder
carries its own glyph and the user labels carry a tag glyph, so the client draws the two kinds as
different kinds of thing. Only `Snoozed` and `Scheduled` are greyed; the other six role folders are
rendered live.

**The finding: a role mailbox is not a label, and Inbox is the sole exception.** `Inbox` is the only
mailbox that appears in *both* pickers, so it is the one mailbox belonging to both namespaces. "Move
to" is the folder namespace, "Labels" is the label namespace, and every other role mailbox sits in
the folder namespace alone - a message cannot be given Archive, Trash, Spam, Drafts, Sent, Snoozed
or Scheduled the way it is given a label, because the client never offers it. This is the general
form of the observation in the Inbox + Trash caveat above, which reached the same conclusion for
Trash alone by watching a client Delete replace the whole `mailboxIds` value.

**Corroboration, by the second method.** The greying matches the live JMAP probe recorded in fork
issue #43, which moved a message into each destination in turn and found `scheduled` and `snoozed` -
and only those two - rejected by the server. That probe sorted the destinations into four groups:
server-protected (`scheduled`, `snoozed`), accepted-but-corrupting (`drafts`, `sent`),
destructive-or-spam (`trash`, `junk`), and normal (`archive`, `inbox`, a user label). The client
greys exactly the server-protected pair and offers the other six live. Two independent methods
agreeing is the standard this file sets at the top, and this row meets it.

**What the greying does NOT settle.** That greyed means "not a valid manual destination" is a
*reading* of the pixels, corroborated by #43 but not measured: neither greyed entry was clicked, so
whether the client swallows the click, shows an error, or acts anyway is **unmeasured**. Two further
things are unmeasured and are called out rather than left blank: folders mode, where the label
namespace does not exist at all; and the conversation-grouping axis, since both pickers were opened
on a single message. The grouping section above measured Archive, not the pickers, so it does not
settle whether a grouped account applies a picker choice to every message in the conversation.

## Effect: what Archive does when it is offered

| Message is filed in | Archive does | Evidence |
| --- | --- | --- |
| Inbox + a label | Inbox removed, label kept, **Archive not added** | a newsletter went `{Inbox, Gmail}` → `{Gmail}`, keywords unchanged |
| Inbox only | **Moves to Archive** | Inbox 43→42, Inbox-only 31→30, Archive 15843→15844 |
| A label only, no Inbox | **Nothing**; the UI says "already archived" | re-archiving changed no mailbox and no count |

Account-wide corroboration for the first row: `Email/query` for "in Archive AND in at least one
other mailbox" returns **0 of 15,843**. Fastmail never creates an Archive-plus-label message.

**Mode dependence.** Fastmail's own MCP describes the operation as "in folders mode the emails are
moved to the Archive folder; in labels mode the Inbox label is removed and any user-applied labels
are preserved" (`docs/official-mcp-surface.md:60`). The rows above are labels mode. In folders mode
every message has a single membership, so the Inbox-plus-others case cannot arise and the
Inbox-only case is already correct; the no-op row is the one that would differ. This server does not
detect the mode (fork #122).

## What the server does NOT settle

Grepping Cyrus's `imap/jmap_mail.c` for `archive` returns exactly one hit — `list-archive` at line
628, an RFC 2369 header name — and no `\Archive` hit at all. **Cyrus has no archive concept.**
Archiving is entirely client policy expressed as `Email/set` on `mailboxIds`, which is why the
client is the authority here and reading the server source cannot substitute for measuring it.

## What is NOT known about Fastmail's own API

`docs/official-mcp-surface.md` was produced by a script that calls only `tools/list` and
deliberately makes no `tools/call`, so it records Fastmail's tool *description* and never their
behaviour. Their `archive_email` documents no refusal and is annotated `idempotentHint=true`, but
absence of a documented refusal is not evidence of no refusal — do not cite it as a contrast with
the client. Settling it needs a real call against their endpoint.

## Authoring: what the client writes for a calendar event

Same charter, a different surface. Everything above measures the message-action screens; this
section measures what the client *writes*, because the client's own stored bytes are the reference
for what this server's write path should author. Nothing here is inferred from what iCalendar
permits — RFC 5545 allows floating time, UTC and offsets, and the client uses none of them.

**Method, and the instrument.** Six reference events were authored in the Fastmail **mobile** app on
22 August 2026, one per shape below, and the stored iCalendar was fetched back over CalDAV the same
day. Four further events were authored in the **web** client on 23 August 2026 and fetched back the
same way, covering the recurrence and DST shapes the first pass left open. The mobile and web
clients share their authoring logic, so these are recorded as the client's shapes rather than either
app's; that sharing is the operator's statement, and the second pass is the first *measured* support
for it — on the four shapes it covers the web client wrote the same model as the mobile one (a zone
name plus a wall clock, `VALUE=DATE` for all-day). That is agreement on four shapes, not a
measurement of every shape, and it is agreement on the *model*, not on every spelling: the two
clients wrote the same three-day all-day event with different end properties (see the storage
paragraph below). This is a third method alongside the two at the top of
the file — not a reading of pixels and not a `mailboxIds` diff, but the resource's bytes as the
server stored them. Bytes need no second
method to corroborate them, which is why one pass settles these rows.

| Event kind | What the client wrote |
| --- | --- |
| Timed, all defaults | `DTSTART;TZID=Australia/Sydney:20260822T090000` + `DURATION:PT1H`, with an embedded `VTIMEZONE` for the zone |
| Timed, zone chosen in the picker | `DTSTART;TZID=Asia/Hong_Kong:20260822T090000` + `DTEND;TZID=Asia/Hong_Kong:20260822T100000`, with an embedded `VTIMEZONE` carrying `TZID:Asia/Hong_Kong` |
| All-day, single day | `DTSTART;VALUE=DATE:20260822` + `DURATION:P1D`, plus `TRANSP:TRANSPARENT` |
| All-day, three days | `DTSTART;VALUE=DATE:20260822` + `DTEND;VALUE=DATE:20260825` — an **exclusive** end |
| Weekly timed series | `DTSTART;TZID=Australia/Sydney:20260822T090000` + `RRULE:FREQ=WEEKLY;COUNT=4` |
| Yearly all-day series | `DTSTART;VALUE=DATE:20260822` + `RRULE:FREQ=YEARLY;COUNT=3` + `DURATION:P1D` |
| Weekly timed series ended by a **date** ("Last occurs on") | `DTSTART;TZID=Australia/Sydney:20260826T093000` + `DURATION:PT1H` + `RRULE:FREQ=WEEKLY;UNTIL=20260923T135959Z`, with an embedded `VTIMEZONE` carrying `TZUNTIL:20260923T145959Z`. The `DTSTART` is the **post-edit** value; it was authored at `090000` and moved by the whole-series edit below |
| Monthly timed series, "every month on the 3rd Tuesday" | `DTSTART;TZID=Australia/Sydney:20260915T090000` + `DURATION:PT1H` + `RRULE:FREQ=MONTHLY;BYDAY=3TU;COUNT=3`, with an embedded `VTIMEZONE` carrying `TZUNTIL:20261116T230000Z` |
| All-day, three days, spanning the Sydney DST change | `DTSTART;VALUE=DATE:20261003` + `DURATION:P3D` + `TRANSP:TRANSPARENT` — no zone, no `VTIMEZONE` |
| All-day **daily series** across the same DST change | `DTSTART;VALUE=DATE:20261003` + `DURATION:P1D` + `RRULE:FREQ=DAILY;COUNT=3` + `TRANSP:TRANSPARENT` |

**`UNTIL` is UTC, and it is the last second of the chosen local day.** The picker was given a date,
"Last occurs on Wed, 23 Sep 2026", and the client wrote `UNTIL=20260923T135959Z` — 23:59:59 on the
23rd in `Australia/Sydney`, converted to UTC. So among the values that *schedule* the series,
`UNTIL` is the one the client writes as a `Z`, and the bound it means is a whole local day rather
than the series' own clock time. (The resource's housekeeping timestamps — `CREATED`, `DTSTAMP`,
`LAST-MODIFIED`, and `TZUNTIL` and `LAST-MODIFIED` inside the `VTIMEZONE` — are UTC as well, but
none of them schedules anything, so they are not what this claim is about.) Note
also what is **absent**: the weekly rule carries no `BYDAY`, so the weekday is taken from `DTSTART`
and a reader must not expect the rule to restate it. The monthly "3rd Tuesday" rule does carry
`BYDAY=3TU`, because there the weekday is not derivable from `DTSTART` alone. Reading the series
back, the client's own popup rendered the `UNTIL` as a count plus a last date — "It occurs 5 times,
starting on Wed, Aug 26, 2026 and last occurring on Wed, Sep 23, 2026" — so a count in the UI is not
evidence of a `COUNT` on the wire.

**A DST boundary leaves no trace in an all-day value.** Both October fixtures run across the Sydney
transition (DST starts 02:00 on Sun 4 Oct 2026) and neither records it: the single event is
`DTSTART;VALUE=DATE:20261003` + `DURATION:P3D`, the series is the same start + `DURATION:P1D` +
`RRULE:FREQ=DAILY;COUNT=3`, and neither carries a zone or an embedded `VTIMEZONE` at all. A
multi-day all-day event spanning a transition is a plain run of dates, which is what makes the
date-only reading in the window filter safe across one. (The noise these four carry differs from
the first pass's — see the client-noise paragraph below.)

**The client never writes a floating or absolute time.** Every timed value in all six is an IANA
zone *name* plus a local wall clock. Not one `Z` form, not one numeric offset, not one bare
`DTSTART:20260822T090000`. The claim is about the values that *schedule* an event — `DTSTART`,
`DTEND`, `RECURRENCE-ID`, `EXDATE`, and a recurrence rule's own bound — and it survives
the second pass, where the only `Z` among them is the `UNTIL`. Read it as scoped to those values,
not to the resource: the housekeeping timestamps (`CREATED`, `DTSTAMP`, `LAST-MODIFIED`, `TZUNTIL`)
are UTC throughout, and always were. That is the measured ratification of the write model this server ships
(#139, #157): a zone name plus wall clock in both directions, and an omitted zone writing the
configured zone rather than leaving the value floating. The client would author the same bytes.

**All-day means `VALUE=DATE`, not a midnight-to-midnight timed span.** Both all-day shapes carry a
date-only `DTSTART`, and the multi-day one ends with a date-only `DTEND` one day past the last day
it covers. So an all-day event is a run of local days with no zone attached, and this is the
measurement `list_calendar_events`' exact window filter rests on (#162): a date-only value is the
configured zone's local day, a date-only `DTEND` is already exclusive so a `DTSTART..DTEND` span is
read as the full multi-day local span with no day added, and an all-day value is never converted to
an instant. This server's create path already serialises date-only input the same way.

**Storage serialisation varies by path, within one account.** Among the six, some resources carry
`PRODID:-//Fastmail/2020.5/EN` and others `PRODID:-//CyrusIMAP.org/Cyrus …//EN`, and the end of an
event is spelled sometimes as `DURATION` and sometimes as `DTEND`. Both end-shapes are real on the
wire from the same account and from the same client, so neither can be treated as the canonical
one; a reader must handle both, and this server's parser does. The second pass sharpens it: the
three-day all-day event authored on 23 August came back as `DURATION:P3D` where the identically
shaped one from 22 August had a date-only `DTEND`, so the two spellings are not even split by event
kind — the same kind of event, in one account, produced both.

### Editing one occurrence of a series

The weekly series was edited twice in the client: the third occurrence deleted, the second moved
half an hour later. Both edits landed in **one resource under one UID**, as two VEVENT blocks.

- The deleted occurrence became `EXDATE;TZID=Australia/Sydney:20260905T090000` on the master — an
  exclusion date, **not** a `STATUS:CANCELLED` override block.
- The moved occurrence became a sibling override VEVENT:
  `RECURRENCE-ID;TZID=Australia/Sydney:20260829T090000` with
  `DTSTART;TZID=Australia/Sydney:20260829T093000`.
- `SEQUENCE` was bumped to 1 on both blocks.

Note that the `RECURRENCE-ID` carries the same zone-name-plus-wall-clock form as everything else,
so identifying an occurrence uses the same model as scheduling one.

### Editing a whole series

The weekly `UNTIL` series was edited from its **first** occurrence, choosing **All occurrences**,
with the title changed and the start moved from 9:00 to 9:30. The client rewrote the **master VEVENT
in place**, and wrote nothing else:

Only the **post-edit** bytes were fetched, so the two-sided claims below are marked with what they
rest on. Read an unmarked one as measured.

- `DTSTART` is `20260826T093000` and the `SUMMARY` is the new title, in the one existing block.
  The `090000` and the old title it moved *from* are the operator's own authoring and edit actions
  — 9:00 was entered in the picker and 9:30 typed over it, and the title was retyped — not a
  pre-edit fetch. The table row above records the same post-edit value.
- `SEQUENCE` is **1**, and `DTSTAMP` and `LAST-MODIFIED` are later than the resource's `CREATED`.
  That it was 0 before is *inferred* from the three sibling resources authored in the same pass and
  not edited, which all carry `SEQUENCE:0`.
- **One resource, one UID, after the edit** — the edit did not fork the series into a second
  resource. This is not a before-and-after comparison of the UID: only the post-edit resource was
  read, so what is measured is that a single resource under a single UID is what the series is.
  It is bounded, though, by how that resource was found: a title-substring sweep across **all five
  calendar collections in the account** over 2026-07-24 to 2026-12-21 returned exactly four
  resources for the pass's four events. A fork would have had to land outside that window or under
  a title not containing the substring to escape it.
- `RRULE` carries `UNTIL=20260923T135959Z`, exactly as authored, so moving the series did not move
  its end bound with it. That the rule is otherwise unchanged is *inferred* from the client's own
  popup, which read the same before and after the edit — "5 times … last occurring on Wed, Sep 23".
- **No override VEVENT, no `RECURRENCE-ID`, no `EXDATE`.** The resource still holds exactly one
  VEVENT. It also carries **no `DESCRIPTION`** at all; the three un-edited siblings each carry an
  empty `DESCRIPTION:`, so the rewrite most likely dropped one — *inferred* from the siblings, not
  from a pre-edit fetch. Noise either way, but it suggests the block is rewritten rather than
  patched.

So the two edit modes have nothing in common on the wire: editing one occurrence adds a sibling
override block beside the master, editing the whole series mutates the master and leaves the
resource single-block. Nothing structural distinguishes a whole-series edit from an event that was
never edited — only `SEQUENCE` and the timestamps record that anything happened.

**The occurrence picker offers exactly two choices.** Editing an occurrence of a series pops "This
event only" and "All occurrences", and nothing else. There is no "this and future occurrences", so
the client never authors the split that option implies elsewhere (capping the old master with an
`UNTIL` and starting a fresh series), and a resource in that shape did not come from this client.

**Client noise a parser must tolerate — in both directions.** All ten resources carry
`X-JMAP-USEDEFAULTALERTS;VALUE=BOOLEAN:TRUE`, so that one is simply what the client writes. The
other two are not. The 22 August six carry `VALARM` blocks; **not one of the 23 August four does**.
The empty `DESCRIPTION:` line is on the 22 August six and on three of the 23 August four — every
one except the whole-series edit, which carries no `DESCRIPTION` at all. Every one of the ten is an
ordinary event nobody configured specially, so a parser must survive each of those two properties
being present *and* being absent, and must not read an absence as meaning the event came from
somewhere other than this client. What made the two passes differ on the `VALARM` was not
identified — an account-level default alarm setting is the obvious candidate and was not checked.
All four of the second pass additionally carry `STATUS:CONFIRMED` and the Cyrus `PRODID`. Read that
as a fact about those four rather than a rule about the client: the first pass's six are mixed
between the Cyrus and the Fastmail `PRODID`, per the storage-serialisation paragraph above.

### How this server's writes render in the client

The section above reads the client's bytes. This reads the client's *pixels*, for bytes this server
wrote — the same comparison run backwards, and the only way to find out whether the client agrees
with what the create path says it produced.

**Method, and the instrument.** On 23 August 2026 the four events below were authored through this
server's own create path (`create_calendar_event`, driven against the built `dist/` over the MCP
harness, no participants) into a collection minted by
`scripts/probes/server-authored-events.probe.mjs` with `MKCALENDAR`, and each was opened in the
Fastmail **web** client the same day. The account's configured zone was `Australia/Sydney`, on AEST
at the time. This is a fourth method in this file: the pixels of the client's event popup, read
against bytes this server wrote rather than bytes the client wrote. Also measured in passing: the
`MKCALENDAR`'d collection appeared in the client's calendar list under its display name with no
further step.

| What this server wrote | What the client showed |
| --- | --- |
| `DTSTART;TZID=Australia/Sydney:20260826T100000` + `DTEND;TZID=Australia/Sydney:20260826T110000` (`timeZone` omitted, so the configured zone was written), `PRODID:-//fastmail-mcp//CalDAV//EN`, no `VTIMEZONE` | `10:00 AM – 11:00 AM AEST (1 hour)`, with no zone annotation |
| `DTSTART;TZID=Asia/Hong_Kong:20260826T100000` + `DTEND;TZID=Asia/Hong_Kong:20260826T110000` (`timeZone: "Asia/Hong_Kong"`), no `VTIMEZONE` | `12:00 PM – 1:00 PM AEST (1 hour)` with a second line `10:00 AM – 11:00 AM HKST` |
| `DTSTART;VALUE=DATE:20260827` + `DTEND;VALUE=DATE:20260828` | `Thursday, August 27, 2026` |
| `DTSTART;VALUE=DATE:20260828` + `DTEND;VALUE=DATE:20260831` (an **exclusive** end) | `Friday, August 28, 2026 – Sunday, August 30, 2026 (3 days)` |

Note that all four end with `DTEND`. This server never writes the `DURATION` form, though the client
writes it on some resources and this server's parser reads both — see "Storage serialisation varies
by path" above.

**A bare `TZID` with no embedded `VTIMEZONE` renders exactly as the client's own events do.** Every
timed event the client authors carries an embedded `VTIMEZONE` for its zone (the two timed rows in
the section above both do); this server writes none, and the popup for the explicitly-zoned event is
identical in format to the client's own zone-picker reference event authored on 22 August, whose
popup — read on 23 August in the same web client, since the 22 August section records bytes only —
gave `11:00 AM – 12:00 PM AEST` over `9:00 AM – 10:00 AM HKST`. So the absence has no visible effect
in the Fastmail client, which resolves the zone name itself. **What this does not measure is
interoperability**: whether a `VTIMEZONE`-less resource resolves the same way in some *other* CalDAV
client was not tested here, and is tracked as #166.

Also worth recording for the read side: #162 changed only the window filter and the refusals, not the
create serialiser, which is unchanged since #157 — so that work produced no newly authored bytes to
view here.

One bound on this whole subsection: a client popup is not a byte-level check. The bytes in the left
column were verified by the probe's CalDAV `REPORT` fetch-back of the stored resource, and only the
right column is pixels.

**Unmeasured.** The 23 August pass closed four of the gaps the first one left: `UNTIL`, a `BYDAY`
expansion, all-day events spanning a DST boundary, and a whole-series edit are all measured above.
What is still not authored, and so still not known:

- **`BYMONTHDAY`** — the monthly picker's other option, "on the 15th". Only the "3rd Tuesday" branch
  was authored, so nothing here says how a day-of-month rule is written.
- **The weekly picker's multi-day form** ("on Saturday & Sunday"). A `BYDAY` list is the obvious
  guess and a guess is not a measurement; the single-weekday case wrote no `BYDAY` at all, which is
  reason enough not to assume the multi-day case by extension.
- **A "this and future occurrences" split** — not merely unmeasured but unavailable: the client's
  occurrence picker has no such option (above), so it cannot be authored from this client at all.
- **A timed series crossing a DST boundary.** Both DST fixtures here are date-only. Whether a
  weekly 9:00 series holds its wall clock or its offset across a transition is the case that
  matters most for a zone-name-plus-wall-clock reader, and it has not been measured.

These are left explicit rather than blank.

## Extending this file

Add a row by **measuring the view**, never by inferring from a role's name. A role that has not been
looked at gets an explicit `unmeasured`, not a blank and not a guess.

Fork issue #121 is the standing audit that fills in the rest of the availability axis; #125 covers
the sibling axis this file only touches for Archive — what each client action *does* to membership
and keywords, and how the client marks the result up.
