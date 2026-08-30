# Live probes

On-demand live verification scripts for externally-observable behavior that unit
tests cannot prove (Fastmail's blob store, MIME assembly, keyword writes). They
run the built server (`dist/index.js`) over real JMAP against the configured
account, via `scripts/mcp-harness.mjs`.

These are **not** durable regression coverage - the unit suite is. A probe run
proves the real external path once, on demand (typically before a release or
after touching an area a probe covers). See CLAUDE.md "Testing".

## Running

1. `npm run build` (the server runs from `dist/`, not `src/`).
2. Run through the token launcher, which injects `FASTMAIL_API_TOKEN` from the
   local MCP client config into the child environment without printing it:

   ```
   python scripts/probes/run-probe.py inline-read.smoke.mjs
   ```

   The launcher also injects the CalDAV username/password when the config
   carries them, because the calendar probes authenticate with a separate app
   password rather than the JMAP token. Those are optional: a config without
   them still runs every JMAP probe, and a calendar probe reports the missing
   credential itself rather than failing obscurely.

Each probe prints one PASS/FAIL line per check and exits non-zero on any
failure.

## Safety rules

- Probes create their own fixtures in the configured account and remove every
  artifact before exiting, including on failure: mail fixtures move to Trash,
  and a probe that needs a container (a label folder, a calendar collection)
  deletes the container, which takes everything inside it in one request.
- One probe (`inline-quotecarry.smoke.mjs`) performs a single send-to-self to
  verify the transmit receipt; the sent and received copies are swept to Trash.
- Never print or persist token values; scripts reference env var names only.
- One probe is a deliberate exception to the first rule.
  `server-authored-events.probe.mjs` **leaves its fixtures in the account between its
  two phases**, because a human has to open the Fastmail client and look at them, which
  cannot happen inside a script run. Its `create` phase writes four events into a
  temporary collection and exits with them in place; its `cleanup` phase deletes those
  collections whole, taking every event inside them in one request. Nothing else sweeps
  them, so a `create` that is never followed by a `cleanup` leaves a stray calendar
  behind. Run the two phases as a pair. Each `create` mints its own collection rather
  than reusing the last one's, so several unswept runs leave several calendars - one
  `cleanup` removes them all, since it deletes every minted collection under the name.

### Creating a calendar event with a participant sends real mail

`create_calendar_event` with a non-empty `participants` is **not** a local
write. Adding an attendee makes the server's scheduling layer send that address
an iTIP meeting invitation, from the account under test, the moment the event is
created - and a `delete_calendar_event` afterwards sends the matching
cancellation. Neither goes through a compose tool, so "the probe never called
`send_draft`" does not mean the probe sent nothing.

That makes it the one live check here that reaches outside the account, and the
consequence is a real invitation in a real person's calendar if the address
belongs to one. Two ways to stay inside the account:

- Omit `participants` entirely when the point of the check is the CalDAV write
  path itself (creation, properties, parsing, deletion). It exercises everything
  except the scheduling hop.
- If the attendee handling is what you are verifying, address it into a domain
  that cannot receive mail. RFC 2606 reserves `example.com`, which publishes a
  null MX, so nothing is delivered and no stranger is contacted.

Note what the second option costs, because it is not free: the invitation is
still *sent*, the null MX merely refuses it, and the refusal arrives back in the
account as an "Undelivered Mail Returned to Sender" bounce - one for the
invitation and one for the cancellation. Those bounces are real messages that
outlive the probe; nothing sweeps them, because they arrive as ordinary inbound
mail rather than as an artifact the probe created and can track. Expect them,
and clear them by hand if you care.

## Inventory

| Probe | Covers |
| --- | --- |
| `inline-read.smoke.mjs` | Read surfacing: `isInline`/`cid` in `get_email`, raw purity, `get_email_attachments`, download by `cid:` (including an `@`-bearing cid), compact lists unchanged |
| `inline-author.smoke.mjs` | Authoring: cid embed + note, lenient cid spellings, text-only degrade, dangling-ref and bad-cid rejects, `[image]` text derivation. Needs `FASTMAIL_ATTACH_DIR` (the probe sets it to the OS temp dir) |
| `inline-quotecarry.smoke.mjs` | Reply/forward quote carry: minted `ii-...@inline.invalid` cids, keep-rebuild reuse, drop/degrade/exclusion notes, asAttachment untouched, `send_draft` transmit receipt |
| `foreign-draft-roundtrip.mjs` | Edit round-trip of a foreign-shape draft (`alternative[text, related[html, inline image]]`, `@`-bearing Content-ID): metadata edit, body-keep edit, ref-dropping edit |
| `probe-exact-instance.mjs` | Exact-instance thread-state marking on duplicated messages (see its header; needs `FASTMAIL_PROBE_TEST_ADDR`) |
| `archive-parity.smoke.mjs` | Archive semantics end to end: the `mailboxIds/` patch form is accepted, an Inbox+label message keeps its label without gaining Archive, an Inbox-only message reaches Archive, a message already out of the Inbox is untouched, a refusing role writes nothing, and a mixed batch sends both patch shapes in one `Email/set`. Creates and destroys its own label folder |
| `calendar-expand.probe.mjs` | The platform fact behind #64: what Fastmail's CalDAV server returns when a `timeRange` query is sent with `expand: true`. Measures that expanded occurrences arrive with `RRULE` stripped and `RECURRENCE-ID` set at the real in-window date, and — the part that decides the fix — that several occurrences arrive as multiple VEVENTs inside ONE `calendar-data` blob, so a first-match parser drops all but one. A second pass then finds a real recurring series and builds a window from **its own DTSTART**, settling the fact the first pass cannot see: the series' FIRST instance comes back with no `RECURRENCE-ID` at all, so "the block without a RECURRENCE-ID is the master" is a false read of an expanded payload and discards every sibling. Read-only; creates nothing. Raw CalDAV via tsdav, not the built server, so it measures the platform rather than our parsing |
| `calendar-window.probe.mjs` | The other half of #64, end to end: whether the shipped `list_calendar_events` answers a window question with dates that are actually in that window. Asserts every returned `start` falls inside the requested range, that a recurring entry arrives marked as an expanded occurrence rather than as the series master it used to be reported as, and that the response carries a total so a trimmed list is visible (#100). Then follows one series back to its own start date and asserts the window reports **every** occurrence in it — including the bare first instance, marked recurring rather than as a one-off — that a date-only single-day window covers the caller's **local** day (asserted by equivalence against the same day written as instants, so it holds whatever the account has on it, and printed alongside what the old UTC-day reading would have returned), and that a window given only one bound is bounded rather than run open-ended, with the range it actually searched stated in a trailing `Note:`. Then the same for a call given **no** bounds at all, which is now the next month from today rather than an unwindowed listing (#142): the note must say neither bound was given rather than blaming one the caller never passed, and must name the span and the range searched. Takes an optional third argument naming the single date to check, defaulting to the date the wrong-day window was reported on. Spawns the BUILT server through the MCP harness, so unlike `calendar-expand.probe.mjs` it measures our parsing and filtering rather than the platform. Read-only; creates nothing |
| `calendar-window-frames.probe.mjs` | Which TIME FRAME Fastmail's CalDAV server resolves each kind of value in when it matches a `time-range`, and what that costs a caller (#162). Settles two missing-event bugs the issue derived from the Cyrus source but never observed. Writes three synthetic fixtures — an all-day `VALUE=DATE`, a genuinely floating timed value, and a `TZID` control that discriminates the windows — into a calendar it creates by MKCALENDAR, then queries five windows in the exact shape the shipped tool sends (`time-range` + `expand`, UTC instants). Measures: a date-only value is matched on its UTC DAY, so a sub-day local-morning window returns no occurrence for it; a floating value is resolved as UTC, so a `+10` caller's evening window misses a floating 20:00 event that a 20:00Z window returns; no `CALDAV:calendar-timezone` exists on the collection or the calendar home, which is why the UTC fallback happens; and expansion Z-stamps floating and `TZID` values but leaves a date-only value a bare DATE. Also separates the filter from the expansion — on the boundary-touching window the filter MATCHES the resource while `expand` emits nothing — which corrects the mechanism the issue had recorded. Raw CalDAV over bare `fetch`, not the built server and not tsdav; raw PUT is mandatory because this server's own create path always writes a `TZID`. Deletes its temporary collection in a `finally` |
| `client-authored-events.probe.mjs` | What Fastmail's own clients write on the wire when a user authors a calendar event — the raw stored iCalendar, fetched back untouched, for extending `docs/fastmail-action-availability.md` by measurement (#165: UNTIL and BYDAY recurrence forms, a DST-spanning all-day event, whole-series edits). The operator authors events in a Fastmail client with distinctive titles; the probe discovers the calendar home (PROPFIND), enumerates collections, calendar-query REPORTs a relative window (30 days back to 120 ahead) and dumps every event whose SUMMARY contains one of the title substrings given as arguments (default: the 22 Aug 2026 reference set). Output redacts the account name and every email-shaped string, and unwraps the CDATA sections the server wraps calendar-data and display names in, so the printed bytes are the stored bytes. Raw CalDAV over bare `fetch`, not the built server — it measures what the client wrote, with none of our parsing in the way. Read-only; creates and deletes nothing |
| `calendar-rdate-expand.probe.mjs` | What Fastmail's CalDAV server does with `RDATE` on the two paths a windowed read uses (#165). Settles the half `docs/conventions.md` had recorded as derived from the Cyrus source: `<C:expand>` DOES strip `RDATE`, exactly as `calendar-expand.probe.mjs` measured it stripping `RRULE` — an `RDATE`-only series comes back as one VEVENT per occurrence, no `RDATE` line on any of them, `RECURRENCE-ID` on every block after the first. The same run measures something the repo had not recorded and had assumed the other way: **the `time-range` filter does not walk `RDATE`s at all.** A window covering an `RDATE` occurrence but not the series `DTSTART` matches the resource not at all, with or without expand, while an `RRULE` control at the identical instant is returned by the same request; two further windows pin the indexed span to `DTSTART..DTSTART+DURATION`. Both `RDATE` serialisations (one comma-joined line, one property per line) are written as separate fixtures and every check runs per form, so a form-specific result reports as a divergence. Raw CalDAV over bare `fetch`, not the built server and not tsdav; raw PUT is mandatory because `create_calendar_event` has no `RDATE` parameter. Every fixture goes into a collection it creates by MKCALENDAR and **there is no fallback** — if MKCALENDAR fails the probe stops rather than writing into a real calendar. Deletes its temporary collection in a `finally` and PROPFINDs to confirm it is gone, naming the collection for a manual delete if either step fails |
| `server-authored-events.probe.mjs` | The inverse of `client-authored-events.probe.mjs` (#164): what THIS server's create path puts in front of a human, checked against the Fastmail client's own rendering. Two phases with a person in between. `create` mints a temporary collection by MKCALENDAR, waits for it to become addressable in `list_calendars`, writes four reference shapes through `create_calendar_event` (timed with `timeZone` omitted, so the configured zone is written; the same wall clock explicitly in `Asia/Hong_Kong`; an all-day single day; an all-day three-day band with its exclusive DTEND), prints each response's statement of what it wrote, fetches every stored resource back raw over CalDAV (calendar-query REPORT, no expand) and prints it, then lists what to look for on screen and on which dates. The fixture dates are computed from today (the next Wednesday at least two days out, then the following days), so a run in any month lands them in the coming week. `cleanup` deletes that collection whole. Takes the calendar display name as an optional second argument (default `MCP probe calendar`), which must match across both phases. If MKCALENDAR fails the probe stops rather than writing into a calendar it did not mint. **Both phases match on the display name AND the `mcp-164-` path segment `create` mints**, so a same-named collection anywhere else is refused rather than written into or deleted - a display name is a CLI argument and could name a real calendar. `create` refuses whenever such a collection exists at all, even alongside one of its own, because the name would not say which of the two a write reaches; the writes themselves are addressed by the minted collection's URL. **Every run mints a fresh collection and never reuses an earlier one** - so the run's resource count means something, and one run's cleanup cannot take another run's fixtures; an earlier minted collection is reported and left alone. `cleanup` deletes every minted collection under the name, so one run of it sweeps every earlier run's leftovers too. No participants, so nothing is mailed. Output redacts the account name and every email-shaped string |
| `label-emptiness.probe.mjs` | The emptiness-guard premise behind #132: a membership patch that would leave a message filed nowhere is REJECTED for a message that has never moved, but ACCEPTED — expunging the message — for one carrying a tombstone from an earlier move. Raw JMAP, not the built server, so it measures the platform rather than the guard `remove_labels` now applies on top of it. Creates and destroys its own mailboxes |

`jmaplib.mjs` is a minimal raw-JMAP helper (session, Email/set, blob upload,
tiny PNG generator) used to build fixtures outside the server under test.
`probelib.mjs` holds the shared check harness and response parsing.
