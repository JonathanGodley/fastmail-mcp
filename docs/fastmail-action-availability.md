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
day. The mobile and web clients share their authoring logic, so these are recorded as the client's
shapes rather than the mobile app's; that sharing is the operator's statement and is not measured
here. This is a third method alongside the two at the top of the file — not a reading of pixels and
not a `mailboxIds` diff, but the resource's bytes as the server stored them. Bytes need no second
method to corroborate them, which is why one pass settles these rows.

| Event kind | What the client wrote |
| --- | --- |
| Timed, all defaults | `DTSTART;TZID=Australia/Sydney:20260822T090000` + `DURATION:PT1H`, with an embedded `VTIMEZONE` for the zone |
| Timed, zone chosen in the picker | `DTSTART;TZID=Asia/Hong_Kong:20260822T090000` + `DTEND;TZID=Asia/Hong_Kong:20260822T100000`, with an embedded `VTIMEZONE` carrying `TZID:Asia/Hong_Kong` |
| All-day, single day | `DTSTART;VALUE=DATE:20260822` + `DURATION:P1D`, plus `TRANSP:TRANSPARENT` |
| All-day, three days | `DTSTART;VALUE=DATE:20260822` + `DTEND;VALUE=DATE:20260825` — an **exclusive** end |
| Weekly timed series | `DTSTART;TZID=Australia/Sydney:20260822T090000` + `RRULE:FREQ=WEEKLY;COUNT=4` |
| Yearly all-day series | `DTSTART;VALUE=DATE:20260822` + `RRULE:FREQ=YEARLY;COUNT=3` + `DURATION:P1D` |

**The client never writes a floating or absolute time.** Every timed value in all six is an IANA
zone *name* plus a local wall clock. Not one `Z` form, not one numeric offset, not one bare
`DTSTART:20260822T090000`. That is the measured ratification of the write model this server ships
(#139, #157): a zone name plus wall clock in both directions, and an omitted zone writing the
configured zone rather than leaving the value floating. The client would author the same bytes.

**All-day means `VALUE=DATE`, not a midnight-to-midnight timed span.** Both all-day shapes carry a
date-only `DTSTART`, and the multi-day one ends with a date-only `DTEND` one day past the last day
it covers. So an all-day event is a run of local days with no zone attached, and the exact
re-filter in the #162 window redesign has to treat it as a local-day span rather than convert it.
This server's create path already serialises date-only input the same way.

**Storage serialisation varies by path, within one account.** Among the six, some resources carry
`PRODID:-//Fastmail/2020.5/EN` and others `PRODID:-//CyrusIMAP.org/Cyrus …//EN`, and the end of an
event is spelled sometimes as `DURATION` and sometimes as `DTEND`. Both end-shapes are real on the
wire from the same account and from the same client, so neither can be treated as the canonical
one; a reader must handle both, and this server's parser does.

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

**Client noise a parser must tolerate.** The six carry `VALARM` blocks,
`X-JMAP-USEDEFAULTALERTS;VALUE=BOOLEAN:TRUE`, and empty `DESCRIPTION:` lines. None of it is
optional to survive: it is what the client writes by default, so it is present on ordinary events
nobody configured specially.

**Unmeasured.** These six cover single events and two simple `COUNT`-bounded rules. `UNTIL`,
`BYDAY`/`BYMONTHDAY` expansions, all-day events spanning a DST boundary, and what the client writes
when a *whole series* is edited rather than one occurrence were not authored and are left explicit
rather than blank.

## Extending this file

Add a row by **measuring the view**, never by inferring from a role's name. A role that has not been
looked at gets an explicit `unmeasured`, not a blank and not a guess.

Fork issue #121 is the standing audit that fills in the rest of the availability axis; #125 covers
the sibling axis this file only touches for Archive — what each client action *does* to membership
and keywords, and how the client marks the result up.
