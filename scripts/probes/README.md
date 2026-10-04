# Live probes

On-demand live verification scripts for externally-observable behavior that unit
tests cannot prove (Fastmail's blob store, MIME assembly, keyword writes). They
run the built server (`dist/index.js`) over real JMAP against the configured
account, via `scripts/mcp-harness.mjs`. The token comes from the environment, not
`~/.fastmail-mcp/.env` (see `inheritHome` in `scripts/mcp-harness.d.mts`).

These are **not** durable regression coverage - the unit suite is. A probe run
proves the real external path once, on demand (typically before a release or
after touching an area a probe covers). See CONTRIBUTING.md "Building and testing".

**Probes run against a live account.** They create and remove real fixtures in
it, and some send real mail to the account's own address. Read a probe's row
below before running it.

## Running

1. `npm run build` (the server runs from `dist/`, not `src/`).
2. Set `FASTMAIL_API_TOKEN` in the environment. The calendar probes also need
   `FASTMAIL_CALDAV_USERNAME` and `FASTMAIL_CALDAV_PASSWORD` (an app password),
   and some read `FASTMAIL_TIMEZONE`; a calendar probe run without them reports
   the missing credential itself rather than failing obscurely.
3. Run the probe with `node`:

   ```
   node scripts/probes/inline-read.smoke.mjs
   ```

If your credentials live in `~/.fastmail-mcp/.env` or Claude Code's MCP server
config rather than your shell, `run-probe.py` is a convenience launcher that
reads them into the child environment without printing them. Each variable comes
from the `.env` file if it sets it, otherwise from the config, otherwise from
your shell:

```
python scripts/probes/run-probe.py inline-read.smoke.mjs
```

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
| `inline-quotecarry.smoke.mjs` | Reply/forward quote carry: minted `ii-...@inline.invalid` cids and their durability across an edit that hands the quote back, drop/degrade/exclusion notes, asAttachment untouched, `send_draft` transmit receipt. **Sends** - the one probe here that transmits (a single send-to-self, swept afterwards) |
| `body-tokens.smoke.mjs` | **Drafts only: this probe never sends.** The `{{signature}}` round trip against a real identity (the configured sign-off lands once, at write, and a body read back carries no token to expand again), the escape through the derived text part, the four `draft_email` token refusals, and `edit_draft`'s `bodyHash` gate. Recipients are addressed into `invalid.example`, which publishes no MX. Ran in full against a live account on 2026-09-02, every check passing |
| `foreign-draft-roundtrip.mjs` | Edit round-trip of a foreign-shape draft (`alternative[text, related[html, inline image]]`, `@`-bearing Content-ID): metadata edit, body-keep edit, ref-dropping edit |
| `probe-exact-instance.mjs` | Exact-instance thread-state marking on duplicated messages (see its header; needs `FASTMAIL_PROBE_TEST_ADDR`) |
| `archive-parity.smoke.mjs` | Archive semantics end to end: the `mailboxIds/` patch form is accepted, an Inbox+label message keeps its label without gaining Archive, an Inbox-only message reaches Archive, a message already out of the Inbox is untouched, a refusing role writes nothing, and a mixed batch sends both patch shapes in one `Email/set`. Creates and destroys its own label folder |
| `calendar-expand.probe.mjs` | The platform fact behind #64: what Fastmail's CalDAV server returns for a `timeRange` query with `expand: true`. Measures that expanded occurrences arrive with `RRULE` stripped and `RECURRENCE-ID` at the real in-window date, several of them as multiple VEVENTs in ONE `calendar-data` blob (so a first-match parser drops all but one), and that a series' FIRST instance carries no `RECURRENCE-ID` at all (so "the block without one is the master" is a false read). Read-only. Raw CalDAV via tsdav, not the built server |
| `calendar-window.probe.mjs` | The other half of #64, end to end through the BUILT server: whether `list_calendar_events` answers a window question with dates actually in that window. Asserts every returned `start` is in range, a recurring entry arrives as an expanded occurrence, one series' window reports every occurrence (the bare first instance included, marked recurring), the response carries a total (#100), a date-only single day is the caller's **local** day, a one-bound window is bounded, and a no-bound one gets the next month from today (#142), each stating the range it searched in a trailing `Note:`. Optional third argument: the single date to check. Read-only |
| `calendar-window-frames.probe.mjs` | Which time frame Fastmail's CalDAV server resolves each kind of value in when it matches a `time-range`, and what that costs a caller (#162). Measures: a date-only value is matched on its UTC DAY; a floating value is resolved as UTC; no `CALDAV:calendar-timezone` exists on the collection or the calendar home, which is why; expansion Z-stamps floating and `TZID` values but leaves a date-only value a bare DATE; and on a boundary-touching window the filter MATCHES while `expand` emits nothing. Raw PUT over bare `fetch` (this server's own create path always writes a `TZID`), into a collection it creates by MKCALENDAR, deleted in a `finally`; if MKCALENDAR is unavailable it falls back to an existing calendar and deletes each resource it PUT |
| `client-authored-events.probe.mjs` | What Fastmail's own clients write on the wire when a user authors a calendar event, fetched back untouched, for extending `docs/fastmail-action-availability.md` by measurement (#165). The operator authors events with distinctive titles; the probe dumps every event from 30 days back to 120 ahead whose SUMMARY contains one of the title substrings given as arguments (default: the 22 Aug 2026 reference set). Output redacts the account name and every email-shaped string. Raw CalDAV over bare `fetch`. Read-only |
| `calendar-rdate-expand.probe.mjs` | What Fastmail's CalDAV server does with `RDATE` on the two paths a windowed read uses (#165). `<C:expand>` strips `RDATE` as it strips `RRULE`. **The `time-range` filter does not walk `RDATE`s at all**: a window covering an `RDATE` occurrence but not `DTSTART` misses the resource, with or without expand, while an `RRULE` control at the same instant is returned, and the indexed span is `DTSTART..DTSTART+DURATION`. Both `RDATE` serialisations are checked separately. Raw PUT over bare `fetch` into a collection it creates by MKCALENDAR, with **no fallback** to a real calendar; deleted in a `finally` and confirmed gone |
| `server-authored-events.probe.mjs` | The inverse of `client-authored-events.probe.mjs` (#164): what THIS server's create path puts in front of a human, checked against the Fastmail client's own rendering. Two phases with a person in between (see Safety rules). `create` writes four reference shapes through `create_calendar_event` into a freshly minted collection, with **no fallback** if MKCALENDAR fails (timed in the configured zone; the same wall clock in `Asia/Hong_Kong`; an all-day day; a three-day all-day band), prints what each response says it wrote and the raw stored bytes, and lists what to look for on which dates, computed from today. `cleanup` deletes every minted collection under the name. Optional second argument: the display name (default `MCP probe calendar`), the same in both phases. **Both phases also match the `mcp-164-` path segment**, so a same-named real calendar is refused, never written or deleted. No participants, so nothing is mailed. Output redacts the account name and every email-shaped string |
| `calendar-uid-query.probe.mjs` | The premise behind #137's targeted lookup, which has no fallback: whether Fastmail's CalDAV server accepts a `calendar-query` on `UID` with `text-match match-type="equals"` and HONORS `equals` rather than CalDAV's default `contains`. An exact-UID query must return the event; a strict substring and an impossible UID must return nothing, since a looser match would manufacture duplicates in an ambiguity count two destructive tools rest on. Also measures, without gating, whether the default collation's case-insensitivity reaches `equals`: on this account it does. Raw CalDAV via tsdav. Read-only and needs no fixture; output is counts and PASS/FAIL only, so a run can be quoted into a public issue or commit |
| `calendar-tzdist.probe.mjs` | Whether this account can be handed a ready-made `VTIMEZONE` by RFC 7808 timezone data distribution (#166). Cyrus implements tzdist behind a per-deployment switch, so only a probe settles it. Three conditions: the service answers; a named zone returns one parseable `VTIMEZONE` carrying the `TZID` asked for; `start`/`end` truncation is honoured, with the `TZUNTIL` the stored shape carries. The base URL is discovered, never guessed, and every 404 names which tier sent it (the routes are in the header). Raw HTTP over bare `fetch`. Read-only; output is safe to quote publicly. **Result on this account, 17 Sep 2026: all three conditions FAIL, so the service is absent**, answered by the backend rather than the edge proxy (measurement in `docs/fastmail-action-availability.md`). Any `VTIMEZONE` this server embeds has to come from somewhere other than the platform |
| `calendar-event-json-put.probe.mjs` | Whether a second route gets the platform to supply the `VTIMEZONE` (#166): a CalDAV PUT of an `application/event+json` JSCalendar body, which Cyrus registers under `WITH_JMAP` and converts with the required timezones added (source refs in the header). PUTs one minimal event in `Australia/Sydney` across the 2026-10-04 spring-forward and checks four conditions: the PUT is accepted; the stored resource holds exactly one `VTIMEZONE` with the event's `TZID`; it carries a `TZUNTIL`; `DTSTART` keeps its zone and wall time, and the end names the intended instant. Raw PUT over bare `fetch` into a collection it creates by MKCALENDAR (**no fallback**), deleted in a `finally` and confirmed gone. No participants. **Result on this account, 25 Sep 2026: the PUT is refused `403` with `CALDAV:supported-calendar-data`; conditions 2-4 are not reached.** That precondition has one emitter on the PUT path, so this deployment does not accept `application/event+json`, and the platform will not supply a `VTIMEZONE` by this route either. Cleanup reported `DELETE 204`, `PROPFIND 404` |
| `label-emptiness.probe.mjs` | The emptiness-guard premise behind #132: a membership patch that would leave a message filed nowhere is REJECTED for a message that has never moved, but ACCEPTED — expunging the message — for one carrying a tombstone from an earlier move. Raw JMAP, not the built server, so it measures the platform rather than the guard `remove_labels` now applies on top of it. Creates and destroys its own mailboxes |
| `contacts-paging.probe.mjs` | The premises behind paging `list_contacts`/`search_contacts` (#94): `ContactCard/query` accepts `sort: [name/given, name/surname, uid]` ascending; `position` is honoured, so consecutive pages neither overlap nor skip and their union equals one large query in the same order; cards with no given name sit together at the start; and the order is byte-wise over given name, surname and uid, so it is case-sensitive (counted as adjacent pairs out of case-insensitive order, and the reverse). Raw JMAP, not the built server. **Read-only**: `ContactCard/query` and `ContactCard/get` (`name` and `uid` only, compared in memory); creates, updates and deletes nothing. Output is PASS/FAIL, counts, booleans and positions only - no name, address, uid or id |
| `calendar-vtimezone.probe.mjs` | End-to-end proof that #166's generated `VTIMEZONE` lands on the wire: creates one timed event in `Australia/Sydney` across the October 2026 spring-forward through the BUILT server, fetches the stored bytes raw over CalDAV, and checks for exactly one `VTIMEZONE` with `TZID:Australia/Sydney`, a `TZUNTIL` equal to the event's `DTEND` in UTC, and exactly one `STANDARD` and one `DAYLIGHT` observance whose `TZOFFSETTO`s match Intl's offsets at `DTSTART` and `DTEND` respectively. The arithmetic is proved in `src/vtimezone.test.ts`; this probe is about the wire. Temporary MKCALENDAR collection, deleted in a `finally`. No participants, so nothing is mailed. Output is PASS/FAIL and counts/offsets only |

`jmaplib.mjs` is a minimal raw-JMAP helper (session, Email/set, blob upload,
tiny PNG generator) used to build fixtures outside the server under test.
`probelib.mjs` holds the shared check harness and response parsing.
