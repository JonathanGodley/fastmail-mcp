# Cross-cutting conventions and gotchas

Developer-facing conventions and non-obvious traps that span multiple tools or are
properties of the toolchain. Per-tool behaviour rationale lives in the relevant GitHub
issue; this file is the shared stuff a developer reading the code needs to know.

## Where to check what the server actually does

Fastmail runs [Cyrus IMAP](https://github.com/cyrusimap/cyrus-imapd) and employs its
maintainers, and the JMAP implementation lives in that tree. So for any question of the
form "what will the server do with this", Cyrus is not a proxy for the answer, it is the
answer — and it outranks the RFC wherever the two disagree, because the RFC says what
should happen and Cyrus says what will. Several decisions recorded in this file and in the
issue tracker were reached by reading it after reasoning from the spec produced a wrong
answer.

The files worth knowing:

| path | what it settles |
|---|---|
| `imap/jscontact.c` | JSContact (RFC 9553) ↔ vCard conversion. Validates nothing on write; silently drops unrecognised `contexts` values on read. A card-level `kind` is **always** emitted — the object is seeded `individual` before any vCard property is read (`jscontact.c:1982`), so a card with no `KIND` line still returns `"individual"` — and the value is lowercased on the way out (`jscontact.c:1169`), with no whitelist, so an arbitrary kind passes through verbatim. `src/contact-card.ts` carries the consequences. |
| `imap/jmap_contact.c` | `ContactCard/set`, address books. |
| `imap/jmap_mail.c` | `Email/get`, `Email/query`, `Email/set`, and what the filter conditions actually mean. |
| `imap/jmap_calendar.c`, `imap/ical_support.c`, `imap/caldav_db.c` | CalDAV and JSCalendar, including the validation `caldav_put` does and does not perform. |
| `imap/jmap_api.c` | Method-response and error framing (RFC 8620 §3.4). |

**Read it; do not copy from it.** `cyrus-imapd` is CMU 4-clause BSD, whose clause 4
requires every redistribution in any form to carry the Carnegie Mellon acknowledgment —
satisfiable, but a deliberate act, not something to do by accident. Its iCalendar parsing
delegates to [libical](https://github.com/libical/libical), which is MPL-2.0 / LGPL-2.1
dual: MPL is file-level copyleft, so a copied file would stay MPL inside this MIT package.
Neither licence restricts reading it to learn what the server does.

## Lenient input coercion

MCP clients (especially LLMs) send sloppy parameter shapes: a comma-joined string where
an array is expected, a single bare id, a stringified boolean, `"to": ""`. The server
coerces rather than rejects, so a reasonable-looking call does not crash. This spans
most tools, so the helpers are centralised in `src/coerce.ts`:

- `coerceStringArray` — array / comma-string / single value to `string[]` (or
  `undefined`). `""` coerces to `[]`. **Every branch trims its elements**, so the three ways
  of writing one list agree: without that, `"e1, e2"` arrives clean while `["e1", " e2"]`
  arrives padded, and the padded value reaches the server and comes back as a not-found — a
  whitespace problem wearing a lookup error's clothes. It is safe at nearly every call site,
  because this coercer otherwise takes only email ids, mailbox references, addresses,
  message-ids and field names, and surrounding whitespace is meaningful in none of them. The
  one call site where it *could* be meaningful is `edit_draft`'s `removeAttachments`, which
  matches against a MIME filename that may legally carry surrounding spaces; that stays safe
  because `resolveAttachmentRemovals` trims the **stored** name too, so both sides of the
  comparison are trimmed and a padded stored name is still reachable.
- `coerceStringArrayStrict` — the same coercion for a parameter that must **fail closed**:
  a value that is present but uncoercible (a number, an object, a boolean) is an
  `InvalidInputError` naming the parameter, instead of the plain coercer's `undefined`.
  `null`/`undefined` still read as absent. The dividing line is what a dropped value
  costs: on a *content* field it means "leave this unchanged", which is harmless; on a
  field that **narrows what the call touches** it silently widens the operation the
  argument was passed to restrict, and the caller reads the wider answer as complete.
  `search_emails`' `requiredMailboxes` / `excludeMailboxes` are the first users.

  `archive_email`'s `emailIds` is the third, and it widens the rule rather than fitting it:
  the id list does not *narrow* anything, it IS the operand. The cost of a silently-coerced
  element is the same shape though — `[null]` would reach `Email/get` as the literal id
  `"null"` and come back in the `notFound` bucket, so a type error would arrive wearing a
  not-found error's clothes and falsify the tool's promise that `notFound` means the server
  did not know the id. So the real dividing line is **whether a dropped or mangled element
  can be mistaken for a legitimate outcome**, and a narrowing argument is the commonest case
  of it rather than the whole of it.

  `list_calendar_events`' `calendarId` is a scalar version of the same rule, and it is worth
  reading as one: it names the single calendar to read, so a value that is present but
  matches nothing must be an error rather than a wider read. `''` used to slip through a
  bare `if (calendarId)` truthiness test and quietly query **every** calendar in the account,
  while `'   '` was correctly rejected — the same mistake answered from two different
  calendars. It is now trimmed and tested for PRESENCE, so both spellings raise the shared
  not-found error. A narrowing argument's failure mode is always this shape: the caller reads
  a wider answer as though it were the narrow one it asked for.

  Strictness is per element, and covers the **empty string** as well as the wrong type. That
  is not pedantry: `['']` passes a `typeof entry !== 'string'` check, and the plain coercer's
  `.filter(Boolean)` runs only on the comma-split branch, so without an explicit check a
  blank element reaches the lookup as a real value and returns the same misleading
  "not found". Blanks are still dropped on the comma-split branch, where they are a separator
  artefact (`"a,,b"`) rather than something a caller wrote down.
- `coerceRecipients` — fans `coerceStringArray` over `to` / `cc` / `bcc` / `replyTo` so
  no recipient field can reach `.map(parseAddress)` as a bare string (the original
  `cc:""` / `bcc:""` crash class).
- `coerceParticipants` — the `participants` array on the calendar write tools to
  `{ email, name? }[]` (or `undefined`). Accepts a real array or a JSON-string array; a
  bare string entry is read as the address, matching the recipient lists. Every other
  shape is a loud reject naming the offending index (unknown per-item key,
  missing/non-string `email`, non-string `name`) — the MCP SDK does not enforce
  `inputSchema`, so this is the only guard on the item shape. A blank string for the
  whole parameter reads as *omitted*, never as the empty list, because an empty list
  removes every attendee on update.
- `coerceBool` — stringified / actual boolean to `boolean` (or `undefined`).
- `coerceUtcDate` — a date or datetime to the JMAP `UTCDate` shape (`2026-07-20T00:00:00Z`)
  for the `search_emails` `after` / `before` filters. `YYYY-MM-DD` expands to midnight UTC
  and `YYYY-MM-DDThh:mm:ss` (with `Z`, an offset, or no zone) is converted; **every other
  shape is a loud reject** naming the parameter, because the mail server's own rejection
  (`invalidArguments`) names no argument at all. This is the deliberate exception to the
  lenient-coercion rule above: `new Date()`'s fallback parser accepts `2026/07/20` and
  `20 July 2026` but reads them as *host-local* midnight, and rolls an impossible day
  (`2026-2-31`) into the next month, so accepting them would silently shift the search
  window instead of failing. A guess the caller can't see is worse than an error they can.
- `coerceContactEmails` / `coerceContactPhones` / `coerceContactAddresses` /
  `coerceContactName` — the contact write inputs, following the same three-part discipline
  as `coerceParticipants` (unknown key rejected by index, every key type-checked by index,
  a fresh literal built from the validated keys). `emails`/`phones` accept a bare value
  string as well as the object form; `addresses` do not, having no single obvious scalar
  reading. A repeated value is rejected naming both positions, because the merge matches
  entries by that value and a repeat could only surface as a phantom addition.
- `requireNonEmpty` / `validateClearFields` — the loud-reject + `clearFields` machinery
  shared by `update_calendar_event`, `edit_draft` and `update_contact`.

The calendar write path carries the same exception for the same reason.
`validateAndFormatICalDate` (`src/caldav-client.ts`) is the single validator for
`start`/`end`: it matches the accepted ISO shapes explicitly so nothing reaches
`new Date()`'s fallback parser, and probes the calendar date so `2026-02-31` cannot roll
into March. `formatDateTimeProperty` calls it and only chooses the property form
(`;VALUE=DATE:`, `TZID`, or UTC) from the serialized result — it never parses the caller's
string itself. Both parsing it separately is exactly how the two drifted apart: for a
while the live path accepted `2026/04/18` and read it as *server-local* midnight, while
the validator that would have rejected it sat exported with no callers.

Leniency has a limit: a value is coerced when the intent is unambiguous, and rejected when
guessing would change the message. `textBody` / `htmlBody` are the reject side — a
non-string body, an entirely HTML-escaped `htmlBody`, and a CDATA-wrapped body are refused
by `assertBodyInputs` (`src/body-format.ts`) rather than repaired, because unescaping or
unwrapping would guess at what the caller meant to send. See `docs/email-bodies.md`.

### Coercing is only half of it: the schema has to declare the lenient shape

A handler that runs `coerceBool` on a parameter the schema declares as `type: 'boolean'`
has written unreachable code. A client that validates arguments against the advertised
`inputSchema` rejects `"true"` before dispatch, so the coercion never sees it, and the
leniency exists only for clients that skip validation. Both halves are required, and they
live next to each other:

- **Schema**: `type: ['boolean', 'string']`, with the description wrapped in
  `lenientBool()` (`src/index.ts`), which appends the note that `"true"`/`"false"` are
  accepted. The prose earns its place on top of the widened type: `["boolean","string"]`
  says a string is accepted but not *which* strings, and `coerceBool` recognises only
  those two spellings. Anything else returns `undefined` and falls to the parameter's
  default rather than erroring, so a caller guessing `"1"` or `"yes"` would silently get
  the default.
- **Handler**: `coerceBool(...) ?? <default>`, never `!!`. Under `!!` the string `"false"`
  is truthy, which inverts the flag. That was a live bug on `raw` and `verbose` across
  nearly every read tool: `raw: "false"` returned untransformed JMAP to a caller that had
  explicitly asked for the simplified shape, and on `get_email` it also made
  `assertStripQuotedNotRaw` reject a legitimate `stripQuoted` read.

`src/tool-schema.test.ts` enforces both halves against the source: it fails on any tool
parameter declared `type: 'boolean'`, and on any bare `!!` read of a boolean parameter in a
handler file. It reads `src/index.ts` as text rather than spawning the built server and
calling `tools/list`, because `npm test` runs `tsx` over `src/` and never builds first, so
a guard reading a stale `dist/` would miss exactly the newly added tool it exists to catch.
`tsc` does not rewrite string literals, so the source and the shipped schema cannot
disagree here.

The same reasoning applies to the array coercions. `participants` on the calendar write
tools now declares `type: ['array', 'string']` alongside `coerceParticipants`, but the rest
are still narrow (`type: 'array'` on `clearFields`, `emailIds`, `mailboxIds`, `attachments`
and the recipient lists) even though `coerceStringArray` / `coerceAttachments` run behind
them. That is unfinished work rather than a decision, and no guard covers it — tracked as
fork issue #98.

### Verifying coercion

The normal MCP tool harness validates the declared `inputSchema` before the call
reaches the handler, so it will reject the malformed inputs these coercions are meant to
accept. That now holds only for the parameters whose schema has *not* been widened: a
compliant client can send a stringified boolean, because every boolean declares
`['boolean', 'string']`. For the rest you must drive a raw JSON-RPC request against the
built server (`dist/index.js`) with `FASTMAIL_API_TOKEN` set, bypassing the
schema-validating harness. `scripts/mcp-harness.mjs` is that client; its `list()` (also
`node scripts/mcp-harness.mjs --list`) dumps the advertised schemas and needs no
credentials. (See the `verify-lenient-client-coercion` note in project memory.)

## Strict parameter keys (the complement to lenient values)

Coercion is the *value* half of input handling; the *key* half is strict. The CallTool
handler runs `assertKnownParams` (`src/coerce.ts`) first, before touching credentials,
and hard-rejects any argument key a tool did not declare in its `inputSchema.properties`
with an `InvalidParams` error that lists the valid keys. The two halves encode one
principle: **recover a clear intent, but refuse to guess at an unclear one.** A
stringified `"true"` or a comma-joined list is unambiguous, so coerce it; a misspelled
`mailbox` for `mailboxId` or a hallucinated `folder` is not, so reject it. A silently
dropped key is worse than a coerced value — the tool runs with defaults and returns
confident wrong results (the original `list_emails {mailbox:'drafts'}` listed *every*
mailbox).

- The allowed-key set is derived from the live `TOOLS` catalog (`TOOL_SCHEMAS` in
  `src/index.ts`), so it never drifts from what clients see via `ListTools`.
- A per-tool `additionalProperties: true` opts that tool out (none set today;
  future-proofing).
- Like coercion, this is unreachable through the normal harness (a compliant client
  cannot send an undeclared key), so verify it with the same raw-JSON-RPC harness.

## Writing a record the tool can only partly see

`update_contact` is the worked example, but the shape generalises to any write whose tool
surface is narrower than the stored record.

**The platform facts first**, live-read off 30 real cards, because the JMAP shape is not
what the property names suggest:

- An `emails`/`phones` entry map is keyed by an **opaque server id**, never by a label.
  Older cards carry 40-character sha1-shaped keys; cards written by the current Fastmail
  interface carry short 6-character ones.
- **Two different properties carry a label, and both are live in one account.** `contexts`
  is a SET (`{"private": true}`) and is what recent cards use — they carry no `label` key
  at all. `label` is a scalar string found on older imported cards, and on every one of
  those observed its value was `""`. So an empty `label` means *no label*, and a
  `contexts` set naming more than one context names no single label either.
- `pref: 1` sits on nearly every entry, and nothing in the simplified read shape shows it.
- A card may be `"kind": "group"`, holding a `members` map of uids and no entries at all.
- A `name` is `{full?, components?: [{kind, value}]}`, and real cards exist with
  `components` and **no** `full`.

**The consequence:** a JMAP PatchObject replaces a top-level property outright, so writing
`emails` from the tool's flat array would silently destroy `contexts` and `pref` on every
entry — fields the caller was never shown and so cannot know to resend. The pattern that
fixes it, in `src/contact-card.ts` (pure) plus `ContactsCalendarClient.updateContact`:

1. **Read the whole record first** — no `properties` filter. An existence probe answers the
   wrong question.
2. **The caller's array decides which entries exist; the stored record decides what each
   surviving entry still carries.** Match by the entry's identifying value, keep the stored
   map key, and write over only the supplied properties.
3. **Refuse the ambiguous edit rather than guessing.** A call that both drops a stored entry
   and adds an unknown one reads equally as "correct this entry" and "delete it, add
   another", and those produce different records. The rejection echoes the dropped entries
   **whole**, so the lossless retry is cheaper than reaching for the override.
4. **Echo the pre-write record, always raw.** `previousCard` (update) and `deletedCard`
   (delete) are the untransformed card whatever `verbose`/`raw` say, because their purpose
   is to keep visible whatever the write took away, and a simplified echo would drop
   exactly the fields the merge exists to protect.

**An echo makes a bad write legible; it does not make it undoable.** Say that plainly
wherever the echo is documented. This server writes a name, emails, phones, addresses and a
note, so `create_contact` rebuilds that much of a deleted card and no more — photos, titles,
organizations, nicknames, URLs, anniversaries, group membership, the `uid` and the per-entry
`contexts`/`pref` are all unreachable from the tool surface. A description promising a
"recreate" would be claiming a fidelity the write path cannot deliver, which is worse than
saying the delete is irreversible. The general rule: an echo's documented promise is bounded
by what the tool can actually write back, not by what the echo contains.

**Where that bound turns into a refusal.** An echo that cannot rebuild a *field* is a
documented limit; a destroy aimed at a record the create surface cannot produce **at all** is
refused outright, because there the echo is worth nothing. That is why `delete_contact` rejects
a contact GROUP: `create_contact` has no `kind` and no `members` parameter, so a group destroyed
here is gone for good, `deletedCard` included - `update_contact` already refused the same card
kind, and both raise it through one shared message so they read as a single rule. The
granularity is the point, and it is the narrow reading that stays correct as the create surface
grows: refuse when the KIND of record is unmakeable, not when a record merely carries fields the
create tool cannot set. Nearly every real card has titles, organizations or photos this server
cannot write, and refusing to delete those would break the tool. The general rule lives in
`CLAUDE.md` ("A destroy must not remove what this server cannot recreate") because it governs
delete paths not yet written.

**The override is scoped to the field that was actually ambiguous**, not to the call.
`allowEntryReplace` is checked per entry array, after that array's own merge has run — so a
call editing `emails` ambiguously and `phones` cleanly whole-replaces `emails` only, and
`phones` still merges. A call-wide flag would quietly strip `contexts`/`pref` off an array
the caller never had a problem with, which is the exact loss the merge exists to prevent.

**A merge that writes a value back unchanged must write nothing.** Resolving a label from
two competing properties (§ the platform facts above) means a round-trip can *look* like an
edit: read an entry whose label came from `contexts`, resend it verbatim, and a naive merge
stamps a scalar `label` on a card that never had one. The merge therefore compares the
supplied label against the *resolved* one and only writes on a genuine difference. The
residual is documented rather than hidden: a changed label lands in `label` while the stored
`contexts` set stays as it was, so a label here can be added or changed but never removed,
and this server and the Fastmail apps can end up disagreeing on one entry. Writing
`contexts` is not implemented.

**A flag is an intent marker; the echo is the safety control.** `update_contact`'s
`allowEntryReplace` and the confirmation parameter `delete_contact` deliberately does *not*
have are the same lesson from both sides: a model retries a rejected call with a flag as
readily as it retries with corrected arguments, so a flag records that a lossy write was
meant — it cannot prevent one. What makes a wrong write survivable is having the previous
state in the response. Do not "strengthen" such a flag into a confirmation handshake, and
do not drop an echo as redundant.

**Where an empty array means "clear", and where it does not.** These diverge across the
server on purpose, and the difference is which mistake is cheaper:

| Tool | `[]` on an entry array | Why |
| --- | --- | --- |
| `update_contact` | REJECTED, naming `clearFields` | `emails: []` is indistinguishable from a mapping bug that produced no entries, and the two outcomes differ by a whole field of the card. |
| `create_contact` | REJECTED, naming the omission | Same read, minus the clear: a card being created has nothing to clear, so the only honest advice is to leave the field out. Accepting `[]` on create while rejecting it on update would make the pair inconsistent for no gain. |
| `update_calendar_event` | `participants: []` REMOVES every attendee | Long-standing published behaviour, stated in the tool description; changing it would break callers that rely on it. |
| `edit_draft` | no bare-array clear at all — `clearFields` only | Same reasoning as `update_contact`. |

`clearFields` is the shape all three agree on: it can only be written deliberately. When
adding a new mutator, follow `update_contact`/`edit_draft` — reject the empty array and
point at `clearFields`.

## Mailbox resolution is uniform (id / role / name / path, exact)

Every mailbox-taking parameter resolves through one exact matcher (`findMailboxExact` in
`src/jmap-client.ts`), in this order:

1. **exact id**, case-**sensitive** - a JMAP id is an opaque server token, and folding its
   case could make two distinct ids collide;
2. **role**, case-insensitive;
3. **exact flat name**, case-insensitive;
4. **root-anchored path** (`Archive/2026/Receipts`), `/`-separated with no leading or
   trailing slash, every segment matched case-insensitively.

No substring matching at any step. A flat name matching exactly one mailbox **wins over**
reading the same text as a path, so a mailbox whose own name contains a `/` stays reachable
by that name. That tie-break is scoped to the case where nothing else answers to the same
text: where a flat `A/B` folder and a real `A > B` nesting **both** exist, the reference is
reported as ambiguous instead, because applying the tie-break there sent a write into the
flat folder with nothing in the response saying a second mailbox had also matched - and a
wrong destination is worth a retry to avoid. A flat name and a path landing on the *same*
mailbox (a top-level folder named `A/B`, no such nesting) is not a collision and resolves.

The walk lives in `findMailboxExact` rather than in its throwing wrapper, and that placement
is the point: both callers inherit it. The label tools' `mailboxIds` arrays (`add_labels` /
`remove_labels` / `bulk_add_labels` / `bulk_remove_labels`) call the matcher directly, and
`search_emails`' `requiredMailboxes` / `excludeMailboxes` entries go through `resolveMailbox`,
so a form accepted on `move_email` but not in one of those arrays would rebuild the split
vocabulary that #50 closed and that made this fork decline upstream's separate
`get_mailbox_by_name` tool. `resolveMailbox` is the throwing single-input wrapper over the
same core, and `list_mailboxes` publishes the same paths it accepts, so a path read out of a
listing goes straight back into any mailbox parameter.

**Four failure shapes, not one.** The matcher RETURNS a discriminated result and never
throws, because only its caller knows what to do with a failure: the single-input wrapper
throws on the spot, while the array resolver collects failures across the whole array. The
shapes are resolved, ambiguous (a flat name matched several mailboxes - candidates are given
as full paths, the form that disambiguates them), a name/path collision (the same text named
one folder and the path to a *different* mailbox), unwalkable (some mailbox's `parentId` chain
never reaches a top-level mailbox, so no path can be computed), and a genuine miss. Upstream's
path lookup silently truncates its bounded parent walk into "not found"; reporting the
unwalkable tree instead is what lets a caller stop hunting for a typo that was never there.

The collision is a *separate* shape from the plain ambiguity for the same reason a typo is:
the correction differs. A duplicated name is fixed by picking one of the candidate paths, but
a path is precisely the text that failed in a collision, so the candidates there are
descriptions - which one is the folder, which one is the nesting - each carrying the **id**,
the only form that separates them. Rendering both as their computed path would hand the
caller the same string twice and no way to choose.

**One array resolver, shared** (`resolveMailboxIdList` in `src/jmap-client.ts`): the label
arrays and `search_emails`' `requiredMailboxes`/`excludeMailboxes` all go through it. It
stays **all-or-nothing** - if any entry fails, nothing happens: no labels are applied, and no
query runs against a scope the caller only half-named - and it names *every* failing entry in
one error, keeping the shapes in separate buckets. That is what preserves the single-retry
property: an ambiguous name and a typo need different corrections (pick a path vs re-spell),
so folding an ambiguity into "not found" would send a caller to re-spell a name that was
already right, costing a retry the split buckets save. When every failure is a plain miss the
message keeps its original single-bucket wording verbatim.

Sharing it rather than giving the read path its own first-failure loop is deliberate: the
single-retry property is a property of resolving a **list**, not of what the list is then
used for, and a read is cheaper to retry than a write, not harder - so there is no version of
"reads can afford to be worse here" that survives being written down. The caller passes in
the mailbox list when it already holds one (`searchEmails` fetches it to resolve `mailbox`
and compute the exclusion), so the sharing costs no extra round trip.

The `path` a mailbox is listed under is computed by walking its `parentId` chain up to a
top-level mailbox, bounded so a corrupt chain terminates. A mailbox whose chain does not
reach a root is listed **without** `path` and its id is named in a trailing note - the
partial walk's shorter string is never emitted, because it would look root-anchored, be
wrong, and could match a different mailbox's real path. Paths are a property of the mailbox
surface only: the per-email `mailboxes`/`roles` enrichment deliberately does not gain them,
because its two `Mailbox/get` calls project `['id','name','role']` to keep a listing cheap
and a path would mean widening that projection on every read page.

**Two destinations are deliberately NOT caller-resolvable at all**: `delete_email`/`bulk_delete`
find Trash, and `archive_email` finds Archive, by **exact role only** (`findByExactRole`), with no
name fallback and no destination parameter. These are fixed defaults with no way to aim them, and a
folder can be created with any name a caller likes, "Trash" or "archive" included. Resolving the
name would make the destination of a destructive default something text the model merely read could
aim; a role is assigned by the server and cannot be minted that way. A caller who wants a different
destination has one: `move_email`, where naming it is the whole point. `archive_email` resolves the
**Inbox** the same way, and needs to, since a caller-created folder named "Inbox" must not become
the membership it strips.

`move_email` and `bulk_move` set `mailboxIds` **whole-value** in a single `Email/set`
(RFC 8620 §5.3) rather than reading the current membership and patching each id away. It states
the promised contract directly ("replaces all mailbox membership") and has no read/write window in
which a newly-added mailbox survives the move. Neither writes a keyword: a move changes where a
message is filed and nothing about its read or flagged state. Marking read is `mark_email_read`.
This is a deliberate divergence from upstream PR `MadLlama25/fastmail-mcp#67`, whose
`archive_email` writes `$seen` as part of the move: folding two effects into one verb means a
caller who wanted only the filing cannot get it, while a caller who wanted both can still make the
second call.

### The membership-subtracting tools reverse that convention, deliberately

`archive_email`, `remove_labels` and `bulk_remove_labels` do **not** use the whole-value form. Each
reads the message's current membership first and emits a `mailboxIds/<id>` patch: `null` for the
mailboxes it is taking away, `true` re-asserted for every mailbox the message keeps. The reason is
their shared contract, which is the opposite of a move's: subtracting one membership must **never**
drop the others, because Fastmail's own Archive and Remove-label actions do not.

What differs is the trigger, not the resolution. `archive_email` adds Archive when the Inbox was
the only filing; `remove_labels`/`bulk_remove_labels` add it, per message, when the named labels
were the only filing. The evidence is the same in both cases and it is narrow: removing a message's
last **user label** in the Fastmail client leaves it in Archive, does not delete it and does not
refuse. What the rescue should do for a message in Trash, Spam or the other role folders used to be
an open extrapolation, because the client offers no remove-label action there. Fork issue #133
settled it by measuring the client's two message-action pickers against each other: the Labels
picker offers the Inbox and the account's own labels and nothing else, while Archive, Trash, Spam,
Drafts, Sent, Snoozed and Scheduled appear only under "Move to". So a role mailbox is a **folder**
in Fastmail's model, with the Inbox as the sole exception, and the label tools reject one before
reading any message's filing - the extrapolation is not needed, because the case cannot arise. The
test is the **role**, never a name list, so a user label someone called "Archive" is still a label
and a role Fastmail adds later is a folder from the day it appears.

Three conditions are refused rather than written, all raised before the write, so a batch
containing one unservable message changes nothing at all:

- some mailbox named for adding or removing is a folder rather than a label, per the namespace rule
  above. `add_labels`, `bulk_add_labels`, `remove_labels` and `bulk_remove_labels` share one
  message for this so the two sides read as a single rule, and it names `move_email`/`bulk_move` as
  the tool that does what the caller meant;
- the account has no archive-role mailbox, so there is no fallback to reach for;
- the server did not report a readable current filing for some message in the batch, which is the
  state a removal destroys a message from.

The namespace rule made a fourth refusal unreachable and it was deleted rather than kept as a
backstop: the removal emptying a message *and* taking Archive away with it cannot happen once
Archive can never be named as a label. The same goes for the patch's null-then-true ordering, which
existed so a rescue could win a key collision with a removal; removed ids and kept ids can no longer
overlap.

One of the three is a per-message condition aborting a whole batch, which is the opposite of the split
`archive_email` draws. The difference is the result shape: the label tools return no per-message
report, so "all of it" and "none of it" are the only honest answers available, and serving the
servable subset would leave the caller a bare success line and no way to learn what was skipped.

The two forms lose different races, and that is what decides it. Whole-value strips a mailbox added
between our read and our write; the patch form resurrects one removed in that window. For a move,
losing the add-race is acceptable — the tool promised to replace everything anyway. For archiving,
losing the add-race is the failure that breaks the feature, while the patch form's worst case is a
concurrently-removed label coming back, which is visible and recoverable.

**The re-assert is not redundancy, and it needs no concurrency to matter.** Cyrus's patch-path
emptiness guard (`imap/jmap_mail.c:13499-13516`) counts a message's mailboxes through
`_email_mailboxes()`, whose callback inserts a key for every mailbox the message has a
*conversations record* in, before testing whether that record is `"added"` (`:822-823`). A mailbox
the message was expunged from days ago therefore still leaves a `{"removed": …}` tombstone in that
map. The `mailboxIds` getter filters on `"added"` (`:7415-7421`); the guard does not. So a bare
`{"mailboxIds/<inbox>": null}` can satisfy a guard that believes the message will still be filed
somewhere, while the executor — which *does* filter tombstones out (`:11413-11417`, and again under
the write lock at `:12916-12933`) — expunges the last live record and destroys the message.

Those tombstones are guaranteed to exist, not merely possible: Fastmail's Undelete reconstructs a
message's previous folders from exactly these expunged records (`Backup/restoreMail`,
`imap/jmap_backup.c:1583-1640`), so every message the account has ever moved carries them for the
length of the restore window. Do not simplify the re-assert back to a lone `null` on the grounds
that a single-user stdio server has no concurrency; concurrency was never what made this reachable.

`ifInState` was considered as a tighter alternative and declined: `Email` state is account-wide, so
one unrelated message arriving mid-call would abort the whole batch.

## Scoping a query: intersection, exclusion, and what `inMailboxOtherThan` really means

The read tools scope by mailbox through two JMAP filter keys, and both have a shape that is
easy to assume wrongly. This is a property of the protocol, so it applies to anything built
on `runFilteredQuery` later, not just to `search_emails` where the parameters live today.

- **`inMailbox` is SINGULAR** (RFC 8621 §4.4.1). "The message must be in all of these" is
  therefore N separate `{ inMailbox }` conditions AND-ed together, never one array-valued
  key. `search_emails`' `requiredMailboxes` builds exactly that, and the scalar `mailbox`
  contributes one more `inMailbox` to the same intersection — which is why the two are
  documented as one thing at different arities rather than as alternatives.
- **`inMailboxOtherThan` is SOLELY-IN, not "not in".** It withholds a message only when
  **every** mailbox that message is filed in is in the excluded set. A message in Inbox plus
  an excluded label is still returned. JMAP offers no strict "not in this mailbox" filter at
  all, so absolute exclusion is only reachable by filtering the returned results caller-side.
  This is stated in the `excludeMailboxes` description *first*, before what the parameter
  does, because the name implies the stricter reading — and it applies equally to the
  default Trash/Spam exclusion, which uses the same key, so a message cross-filed in Trash
  and a normal folder was never hidden and is not in the withheld count.
- **One excluded set, not two conditions.** The default Trash/Spam ids and any
  caller-supplied ids are **unioned** into a single `inMailboxOtherThan`. That is
  deliberately stronger than applying them as two separate exclusions would be: a message
  filed in {Trash, an excluded label} is hidden by the union, though neither exclusion alone
  would hide it. It is also the only shape the protocol offers, since the key takes one set.

**Two flags that look alike and are not.** `runFilteredQuery`'s `doExclude` means *the
default Trash/Spam exclusion is active* — it drives the hidden-count query and the
disclosure note, nothing else. The excluded ids are assigned to the filter **ungated**,
because `doExclude` goes false on paths that say nothing about the caller's own excludes: an
explicit scope, `includeTrash`+`includeSpam`, and the degraded runtime case where neither
role resolves. Gating the assignment on it would drop a caller's excludes on all three — a
fail-open on the parameter that exists to narrow.

**The hidden count subtracts the DEFAULT ids only.** The count query re-runs the visible
filter with the default ids removed and the caller's ids kept, so `hidden` still means
"withheld to Trash/Spam". Subtracting the union instead would make the count identical to the
visible query, so `hidden` would always be 0 and the published "no note means nothing was
withheld" contract would silently become fail-open; subtracting nothing would count
caller-excluded matches while the note prescribes `includeTrash`/`includeSpam`, which cannot
reveal them. For the same reason, a Trash- or Spam-role mailbox the caller excluded itself is
dropped from `excludedRoles` upstream in `computeExclusion` — the note must not prescribe a
recovery flag that the caller's own exclusion overrides.

**Both halves of the note's recovery clause are derived, never written as a constant.** The
`includeTrash`/`includeSpam` flags AND the `mailbox:"trash"/"junk"` override both come from
the surviving `excludedRoles`, because a hard-coded pair goes wrong two ways: it names Trash
when only Spam was excluded, and it names a role the caller excluded itself — where
`mailbox:"trash"` against `inMailboxOtherThan:["mb-trash"]` is a query that contradicts
itself and can only return nothing. A prescription the caller cannot follow is worse than no
prescription.

**Accepted residual: a role-less mailbox NAMED like a role.** On an account with no
`trash`-role mailbox but a folder literally named "Trash", a caller passing
`excludeMailboxes: ["Trash"]` gets the fail-loud "the Trash folder couldn't be found, so it
was NOT excluded" note even though that folder *was* excluded — by the caller's own request.
The note is about the **role**-resolved default exclusion, which genuinely excluded nothing.
Left as-is deliberately: the only way to suppress it is to match the caller's excluded ids
against the role's *name*, which is exactly the substring/name inference the exact-role rule
above exists to forbid (a "Junk mail rules" folder must never be read as the junk role), and
here it would be inference used to **silence** a fail-loud signal — the worse direction to be
wrong in. The cost of keeping it is one redundant "re-run to be sure" in a rare account
shape; the cost of suppressing it is a missing warning on a fuzzy match.

**What disables the default exclusion is decided in two places, and they must agree**:
`computeExclusion`'s `hasExplicitScope` input, and the `exclusionIntended` expression each
caller of `runFilteredQuery` computes for itself (`getEmails` over `opts.*`, `searchEmails`
over `filters.*`). The shared helper does not compute it, so "they both go through
`runFilteredQuery`" does not keep them in step — anything that changes what counts as an
explicit scope has two edit sites. It is still two and not three after `get_recent_emails`
joined the default exclusion (#29): `getRecentEmails` delegates to `getEmails` rather than
assembling its own batch, precisely so it inherits that expression instead of copying it. Naming folders to look **in** (`mailbox`,
`requiredMailboxes`) is an explicit scope and turns the default off; naming folders to
exclude is not, because narrowing by exclusion says nothing about wanting Trash/Spam back.

## Update maps are keyed by caller-supplied ids

Every `Email/set` `update` map built by assignment goes through `newUpdateMap()`
(`Object.create(null)`), never a plain `{}`. An email id of `__proto__` is legal — underscores
are in the base64url alphabet JMAP ids are drawn from — and `updates['__proto__'] = patch` on
an ordinary object invokes the prototype **setter** instead of creating an own key. The entry
never reaches the request, the server is never asked about that id, it cannot appear in
`notUpdated`, and every tool here infers success from the absence of a set-error: so the
message is reported as done while nothing touched it.

Reading those maps back is the mirror image and needs `Object.prototype.hasOwnProperty.call`,
not a bare index. `notUpdated` is parsed from the response, so an id of `constructor` or
`toString` indexes through to a function on the prototype, which is truthy — a **fabricated**
failure for a write the server performed. `setErrorFor` is the shared read; use it rather than
indexing. Two further readings it centralises: **key presence, not the value's truthiness**, is
what counts as a refusal (a server listing an id with a null value has still said it did not
update that record), and the map's **shape is checked** before it is read, because
`typeof [] === 'object'` and `hasOwnProperty.call('oops', '0')` is true — an array or a bare
string would otherwise answer for an id that looks like an index.

An object literal with a computed key (`update: { [emailId]: patch }`) is already safe: that
form is `CreateDataProperty`, not the setter. It is the assignment form that is not.

## Error classification: `InvalidParams` vs `InternalError`

The same recover-clear-intent / refuse-to-guess principle extends to *which* MCP error
code a failure surfaces. MCP clients read `error.code` as a distinct structured field, and
the two codes drive different recovery: `InvalidParams` (-32602) says **"the input is
wrong — re-form it; don't blind-retry as-is,"** while `InternalError` (-32603) says
**"server-side; a bare retry might succeed."** So the dividing line is recoverability:

- **Caller-fixable → `InvalidParams`.** A failure the caller can resolve by re-forming the
  call's arguments, by editing the object (e.g. the draft) the call operates on, OR by
  **replacing that object** — recreating it and deleting the old one — when the object
  itself is unusable rather than merely wrong. That last route is what the body-shape
  refusals on `edit_draft` offer: a draft whose body carries a part the recreate cannot
  reproduce is not something any argument can fix, but recreating the draft is entirely
  within the caller's reach, so it is still their failure to resolve and not a server
  fault. This
  covers bad/empty fields, a not-found id (`get_email`/`get_thread`, `originalEmailId`, a
  draft-mutation target), the body-coupling rejects, an unverified `from`, the
  `send_draft` draft-state guards (no recipients / no from / from not matching an
  identity), and a server-side SetError whose type the caller can act on (see **The JMAP
  set-error split** below). These throw the tagged
  `InvalidInputError` (`src/coerce.ts`), which the top-level CallTool catch maps to
  `InvalidParams` (after `redactBearerTokens`).
- **Operational / server → `InternalError`.** A failure the caller cannot fix by changing
  input: zero sending identities, a missing system mailbox (Drafts/Sent/Trash), a
  transport error, a set-error naming a server/account/state condition, or a
  post-condition like "returned no ID." These stay a plain `Error`.

A **missing Archive mailbox is the one that switches sides**: `archive_email`,
`remove_labels` and `bulk_remove_labels` throw
`InvalidInputError` (`InvalidParams`), not the plain `Error` its Trash counterpart throws. The
split is recoverability, and it really does differ here. "Archive" is a filing convention, so the
caller substitutes any folder they like via `move_email` and gets the same outcome, which is a
re-formed call and so the definition of caller-fixable. A missing Drafts/Sent/Trash has no
substitute: nothing else *is* the trash, and no argument to any tool produces a delete without it.
The error message therefore names `move_email`, because the classification is only honest if the
route it implies exists.

Two conditions on that, both consequences of archiving no longer being a move:

- It is raised **only when a message would otherwise be left filed nowhere** — the Inbox-only
  branch on `archive_email`, the last-label removal on the two label tools. A message that keeps
  other filing never touches Archive at all, so an unconditional guard would reject a batch the
  tool can serve perfectly.
- A missing **Inbox** role sits on the *other* side and throws a plain `Error`. There is no
  substitute call for "remove this from the Inbox" — `move_email` would replace the whole
  membership, which is the thing archiving exists not to do. The guard is load-bearing rather than
  defensive: the entire rule keys on the inbox id, `findByExactRole` is typed `any | undefined` so
  the compiler will not catch its absence, and both failure modes are silent (every message falls
  to the no-Inbox branch and the tool becomes a universal no-op reporting success, or the patch key
  becomes the literal `mailboxIds/undefined`).

### `archive_email` reports per-message failures with no error code at all

`archive_email` takes an array and **never throws on a partial failure**: an unknown id comes back
as a `notFound` entry inside a successful response, not as `InvalidParams`. That is a real change
to an observable contract — the single-id version threw `InvalidInputError` for exactly this case —
and it is accepted deliberately, because a batch tool that threw on one bad id would discard the
outcome of every other id in the call.

The cost is that a caller branching on `error.code` to detect a bad id stops seeing one on this
tool, and has to read the `notFound` bucket instead. The residual is recorded here rather than
fixed locally because it is not specific to archiving: the same question applies to every tool the
array-parameter family will cover, including whether a batch containing *only* failures should
still be success-shaped. Fork issue #120 is where that gets settled.

This rule is **tool-family-agnostic.** Because the calendar tools share the same
`requireNonEmpty` / `validateClearFields` helpers from `src/coerce.ts`, their input
rejects (`create_calendar_event` / `update_calendar_event`) are `InvalidParams` too — the
classification is a property of the shared helpers, not of email specifically. The
calendar rejects raised in `src/caldav-client.ts` itself follow the same rule and throw
`InvalidInputError` directly: the start/end frame and ordering checks below, the date
format and control-character rejects, the participant-address rejects, a `calendarId` or
`eventId` that resolves to nothing, and the repeating-event refusal that
`update_calendar_event` / `delete_calendar_event` raise (see below). `src/contacts-calendar.ts` follows it too — its
input rejects (a contact with neither name nor address, an update naming no field, an
empty entry array, an ambiguous entry edit, an update aimed at a contact group) and its
not-found rejects are `InvalidInputError`, and its create/update/delete set-errors route
through the same `throwSingleSetError` classifier the mail writes use. That classifier puts
a `forbidden` on the operational side, so a contacts token issued read-only surfaces as an
`InternalError` whose message names the read-only scope — the refusal is the only place a
caller can learn that, since the session reports an identical capability either way.

One nearby throw stays a plain `Error` **on purpose**, and looks like an exception to the
rule until you see what it carries. The ORGANIZER address check validates the operator's
configured CalDAV username through the participant validator — a configuration fault no
argument can fix, so it is re-thrown in the plain class with its own wording.

**The repeating-event refusal is `InvalidInputError` even though no argument fixes it**, and
that is not a contradiction. `InvalidParams` means "re-form the call", not "add a parameter":
the caller resolves this by not making the call — pointing a different id at it, or doing the
edit in the web client — which is a caller-side change. `InternalError` would say a bare retry
might work, and it never will: the event repeats, and it will still repeat next time.

`assertICalTextLimits` (`src/ical-limits.ts`) follows the same rule from the other end of
the call chain: it runs in the `create_calendar_event` / `update_calendar_event` handlers,
before anything measures or serializes the values, and throws `InvalidInputError` naming
the field, its size and the limit. The classification matters more than usual there. The
whole point of the bound is that the work it refuses is expensive (see **Bounding a
quadratic serializer** below), so an `InternalError` would not merely be inaccurate, it
would invite the bare retry that repeats the cost.

**`download_attachment` follows the rule with no exception.** A bad `emailId`/`attachmentId`
is caller-fixable input there exactly as it is on `get_email`/`get_thread`, whichever way it
is bad: an unusable reference *form* (a bare `cid:` with no value, a `cid:` value matching
several parts, a number with junk in it, a part carrying no blob) and a well-formed reference
that simply *matches nothing* both throw `InvalidInputError` and reach the caller naming what
to pass instead. Messages quote back only the caller's own input, through `describePart`, so
a hostile value cannot overrun the sentence.

Naming the failure confirms that a message exists and that a reference does or does not
resolve, and that is fine here: `get_email_attachments` enumerates a message's parts on
request and `get_email` confirms the message exists, to the same caller holding the same
token. There is nothing for the download to withhold that the tool beside it does not answer.
So its local `catch` maps `PathAccessError` — a filesystem-confinement decision, a different
control — and re-throws everything else to the top-level catch, which applies the
`InvalidParams` mapping *with* `redactBearerTokens`, and turns a transport or JMAP failure
into an `InternalError` carrying the server's own reason.

`getAttachmentInfo` is shared by that read and by an `attachments` item naming
`emailId` + `attachmentId`, and one class serves both. `uploadAttachments` wraps what it
catches, but only to add context the lookup cannot know — the attachments index, since a
batch rejects on one bad entry. It wraps **by class**, so a transport failure is not relabelled
as a caller mistake.

**Redaction at that boundary is unconditional.** Every branch of the CallTool catch runs
its message through `redactBearerTokens`: the `McpError` rethrow (redacted in place, since
rebuilding would double the SDK's `MCP error <code>: ` prefix), the `PathAccessError` and
`InvalidInputError` mappings, and the generic `InternalError` wrap. There is no exemption
list, which is the point: the audit "no unredacted error text reaches tool output" stays a
grep anyone can run instead of a claim that has to be re-argued per error class. Only
`.message` is scrubbed; `McpError.data` is deliberately left alone, because nothing here
populates it and it is arbitrary JSON a string-shaped scrubber cannot walk.

Server text that reaches the caller on a *successful* result gets no help from that catch,
so it redacts at its own render site. `orphanedOldDraftReason` is the one such field today
(`formatEditDraftResult` in `src/response-formatters.ts`, holding a server or exception
message when an `edit_draft` replacement could not be moved to Trash). It is redacted where
it is rendered rather than where it is assigned, because the render site is single and the
assignment sites are not.

The JMAP set-error reason itself is surfaced (not just the code): every throwing
`Email/set` failure routes its `SetError` through `describeSetError` in
`src/jmap-client.ts` so the server's `type`/`description` reaches the caller, and bulk
mutators additionally report success/fail counts and the caller's failing ids grouped by
reason. The helper concatenates only server-authored text — we add no message body of our
own.

### The JMAP set-error split

When the server refuses a write, `throwSingleSetError` decides the MCP code from the
`SetError.type` alone (`CALLER_FIXABLE_SET_ERROR_TYPES` in `src/jmap-client.ts`; the same
list drives the bulk classifier). The question it answers is not *whose fault is it* but
*what should the caller do next* — so a type is caller-fixable only when re-forming the
request is the route to success:

| Type | Code | Why |
| --- | --- | --- |
| `notFound` | `InvalidParams` | The id names nothing; correcting it is the only route. |
| `invalidProperties` | `InvalidParams` | A property value the server rejected — and every one of them came from the call's arguments. |
| `invalidArguments` | `InvalidParams` | A malformed argument. Method-level in RFC 8620 §3.6.2, but servers return it per-record too, and it means the same thing there. |
| `invalidPatch` | `InvalidParams` | The patch does not apply. We build it from the caller's fields. |
| `tooLarge` | `InvalidParams` | Over the per-object size limit; send something smaller. Matches how `uploadBlob` already classifies an over-limit attachment before it is sent. |
| `singleton` | `InvalidParams` | The target's type forbids the operation. A retry can never succeed; naming a different target can. |
| `forbidden` | `InternalError` | A permission refusal. No argument grants permission. |
| `overQuota` | `InternalError` | The account is at capacity. Neither code fits perfectly — a bare retry will not help either — but the fix is freeing capacity, not re-forming the call. |
| `rateLimit` | `InternalError` | Retrying after a pause is exactly the right recovery, which is what `InternalError` signals. |
| `serverFail` / `serverPartialFail` | `InternalError` | Server-side by definition. |
| `stateMismatch` | `InternalError` | The record moved underneath the request; recovery is re-read then retry. |
| `accountReadOnly` | `InternalError` | The whole account refuses writes. |
| `cannotCalculateChanges` | `InternalError` | A server-side state-window limitation. |
| `willDestroy` | `InternalError` | The server ignored an update because the same call also destroyed the record. This client never batches an update and a destroy for one id, so it cannot come from the caller's arguments. |
| anything else | `InternalError` | An unrecognised type is not evidence that the caller's input was wrong, so the default stays server-side. |

A bulk mutator applies the same table per entry and is caller-fixable only when **every**
failure in the batch is: one operational failure means no single re-form clears the call.

Every throwing set-error site routes through the classifier — `Email/set` `notUpdated`,
`ContactCard/set` `notCreated`/`notUpdated`/`notDestroyed`, and the three draft-lifecycle
`notCreated` throws (`createDraft`, the `edit_draft` recreate, and the `EmailSubmission/set`
in `send_draft`). Those last three built their message from `describeSetError` directly for
a while and stayed a plain `Error` whatever the type, which meant `create_contact` failing
on `invalidProperties` reported a caller-fixable error while `create_draft` failing the same
way reported a server bug. Route new set-error sites through `throwSingleSetError`; reaching
for `describeSetError` alone is how that inconsistency reappears.

The messages are identical on both sides of the split — only the code differs — so a client
that reads `error.message` sees no change.

## Bounding a quadratic serializer

`foldICalLine` (`src/caldav-client.ts`) folds an iCalendar content line to 75 octets per
RFC 5545 §3.1 by repeatedly re-slicing the remainder of the line, allocating a fresh copy
of the tail each time. Its cost therefore grows with the square of the field length:
roughly 135ms to fold a 200KB value, and out of memory somewhere near 800KB. Every
caller-supplied calendar text field reaches it, on both the create and the update path
(SUMMARY, DESCRIPTION, LOCATION, and each ORGANIZER/ATTENDEE line, whose `CN=` parameter
carries a participant name). One oversized value was enough to stall or kill the process
for every other request sharing it.

The guard lives in `src/ical-limits.ts`, ahead of the handlers' own checks, and takes
three bounds rather than one: a per-field cap (64KB), a participant-count cap (500), and
a cap on the combined text of the whole call (256KB). The third is not redundant. A
per-field cap on its own is defeated by many fields each sitting just under it, which is
trivial to arrange through the participants array. Sizes are measured in UTF-8 bytes, not
JS characters, because the fold is defined on octets.

It **rejects and never truncates.** Trimming an over-long description would produce an
event that reads as successfully created while quietly missing content, and the caller
would have no signal at all. The rejection names the field, its actual size and the
limit, which is everything needed to fix it in one retry.

The bounds are properties of that serializer, not of iCalendar or of Fastmail. If the
folding is ever made linear, they are the thing to revisit; until then, removing them
re-opens a denial of service reachable from ordinary tool input.

## Surfacing computed fields without leaking into `raw: true`

When the client layer resolves derived data to attach to a raw JMAP object so a
downstream simplifier can read it — but `raw: true` must stay pure JMAP — attach it as a
**non-enumerable** property:

```js
Object.defineProperty(email, '_mailboxNames', { value: names, enumerable: false, configurable: true });
```

`JSON.stringify` (every `raw: true` path) omits non-enumerable properties, while
`simplifyEmail` reads `raw._mailboxNames` directly. This is how the `mailboxes`, `roles`,
and `unresolvedMailboxIds` fields reach simplified output — mailbox ids are resolved to
names + stable roles in `src/jmap-client.ts` (`buildMailboxInfoMap` / `attachMailboxInfo`,
which attach `_mailboxNames`, `_mailboxRoles`, and `_unresolvedMailboxIds` non-enumerably)
and ride along the email object through every read path — with **zero signature changes**
to `simplifyEmail` / `formatEmailQueryResult`, keeping the formatter pure and testable.
`attachMailboxInfo` is **never-silent but non-throwing**: an id that can't be resolved to a
name is surfaced as a raw id in `_unresolvedMailboxIds` rather than dropped (the #53 fix);
see its in-code comment for why it neither throws nor omits. The convention for the next
computed field: resolve in the client layer, attach non-enumerably under a
leading-underscore name, and read it in the simplifier via `addIf` (so it is omitted when
absent). Do not thread it through function signatures, and do not make it enumerable — it
would leak into raw output.

**The same "beside, never onto" rule covers values that are DERIVED rather than
resolved.** The attachment listing is the union of a message's JMAP `attachments` array
and the media parts the server routed into `textBody`/`htmlBody` (`buildUnionParts` in
`src/inline-images.ts`), and the `isInline` flag is derived from that routing. None of it
is written back to the raw email: the union is computed inside the simplifier from
properties that are already on the object, and the part objects it yields are the
server's own, never copies. A `raw: true` response therefore stays exactly what the
server sent — the same guarantee the non-enumerable route gives, reached without needing
a carrier property at all. Reach for the non-enumerable attachment only when the value
cannot be derived where it is read (a mailbox name needs a second JMAP call; a union of
lists already in hand does not).

**A flag that only a simplified response can honour is REJECTED with `raw`, not ignored.**
`stripQuoted` (#73) rewrites the simplified `bodyText`; there is no honest way to apply it
to a pure-JMAP payload, and silently dropping it would leave a caller believing a response
was stripped when it was not. So `assertStripQuotedNotRaw` (`src/quote-strip.ts`) rejects
the pair from both `get_email` and `get_thread`, naming the two ways out. Same reasoning as
the unknown-parameter guard: a parameter that quietly does nothing produces confident wrong
answers. Contrast `includeBodies` (#74), which is a *fetch*-level knob — it changes the
`Email/get` property set, so `raw` faithfully returns the richer JMAP object and the two
compose fine.

**A third kind: `raw` as a SET escape.** `get_email_attachments` already emits raw JMAP
part objects, so its `raw` has no shape to escape. What it escapes is the *set*: it
returns the JMAP `attachments` array alone, dropping the parts the server routed into the
body lists. Neither rule above fits — there is nothing to reject (the flag is honourable)
and nothing to enrich (the entries are already raw). The rule that does apply is
never-silent: a bare array is indistinguishable from a complete listing, so the withheld
count is stated. It ships as a SECOND content item rather than being appended to the JSON
string, because the JSON staying parseable is the entire point of the flag — the same
reasoning that keeps the Trash/Spam note outside the JSON block on the list tools. The
tool's description says so too, and adds that download entry numbers always count from
the full listing, so a `raw` listing is never an index basis.

## Output width is the caller's to narrow (`fields`)

`verbose` and `raw` both ask for *more*. `fields` is the third axis and the only one that
asks for **less**: an explicit allowlist of simplified field names, honoured exactly.
A named preset (`compact: true`, dropping the threading plumbing and preview) was
considered and **declined**: it would be a second mechanism to document and a second thing
to re-decide every time the shape gains a field, and it guesses at what the caller wanted,
where an allowlist is told. Callers here are overwhelmingly models that know exactly which
fields they are about to read, so being told costs a few tokens and removes the guess.
It exists because response size was otherwise a property of the mailbox rather than of the
call — a 66-message sweep measured 84KB (~47% RFC threading plumbing, ~23% preview,
18% the fields the caller wanted) and an editor-inflated draft body pushed `get_email`
past the same wall (#79, #69). `limit` is not a substitute: per-message size varies by
more than an order of magnitude, so no limit value is both safe and useful.

Four properties define the convention, and a fifth tool adding `fields` should keep all of
them:

- **One seam, post-simplification.** `parseEmailFields` / `projectEmail`
  (`src/field-projection.ts`) run on the *simplified* object, so `raw: true` stays pure
  JMAP and untouched. The list/search tools inherit it through the single
  `formatEmailQueryResult` seam, which is why `list_emails`, `search_emails` and
  `get_recent_emails` cannot drift apart. `get_thread` applies the same pair in
  `readThread`, projecting each message *before* the thread-body byte cap is measured —
  the cap guards what is actually returned, so projecting `bodyText` away also lifts it.
- **One vocabulary, checked by the compiler.** The valid names are the keys of
  `SimplifiedEmail`, held in a `Record<keyof SimplifiedEmail, true>` map — a field added to
  the shape without a line there is a compile error, so a new field can never become
  permanently unselectable. It is deliberately NOT per-tool, since a per-tool allowlist
  would be a second thing to keep in sync. The consequence is that a name can be valid on a
  tool that can never populate it: the list/search tools fetch `EMAIL_PROPERTIES_COMPACT`,
  so every simplified field derived from the `EMAIL_PROPERTIES_VERBOSE` additions (today
  `bodyText`, `bodyHtml`, `bodyHtmlSize`, `attachments`, `forwardedMessageId`) comes back
  absent there. That is documented in the README rather than rejected — it is not a typo,
  and rejecting it would make the vocabulary tool-dependent. State the rule (a field needing
  the full-message fetch is never populated on a list result) rather than only the list, so
  a field added to the verbose tier later inherits the same status without a doc rewrite.
- **Every ambiguity rejects, because every quiet fallback returns the FULL response** —
  the exact refusal the parameter exists to avoid. Unknown name, empty array, and
  `fields` together with `raw: true` are all `InvalidInputError`s naming the fix. Silently
  letting `raw` win would hand the largest response this server produces to the caller who
  asked for the smallest, and a silently-ignored typo would read as "that field is always
  empty" (the same reasoning as the strict-parameter-keys rule above).
- **Subtractive only, with one companion rule.** Projection cannot invent a field or change
  a value's meaning, so caller-directed omission is not the "silently dropped promised
  field" failure — the caller asked. The exception is a field that is not independent but
  the *description of another field's value*: `unresolvedMailboxIds` is the degradation
  half of `mailboxes`/`roles`, and the bodyText signals (`quotedBytesStripped`,
  `quotedStripSkipped`, `bodyTextUnavailable`) say what happened to the `bodyText` being
  returned. Projecting the carrier field emits its companions even unnamed (attached by
  key presence, not truthiness — `quotedBytesStripped: 0` is an answer). Without that, a
  short `mailboxes` array would look complete when it isn't (the #53 bug through a new
  door), and a stripped body would read as verbatim.

Out-of-band signals are not fields and are never projected away: the result-count summary
and the Trash/Spam exclusion note describe the *query*, not a message.

Projection is an **output** transform. It deliberately does not narrow what is fetched from
the server — the JMAP property sets are held identical across the read methods on purpose
(see the JMAP-property-consistency rule in `CLAUDE.md`), and making the fetch shape depend
on a caller's projection would trade that invariant for a saving the caller never sees.

## Result serialisation

Every JSON payload this server returns is serialised **compact**, through one of two seams in
`src/coerce.ts`:

- `toolJson(value)` — the default, used by every handler and formatter.
- `redactedJson(value)` — the same thing with every string value passed through
  `redactBearerTokens` first, for payloads whose values could carry a credential (server error
  descriptions, mailbox names). See its own comment for why redaction has to run per value
  rather than over the finished document.

**Neither takes an indent argument**, so the seams cannot pretty-print even by accident, and
TypeScript rejects a second argument to either. The way back to indented output is a handler
reaching past them for its own `JSON.stringify`, so the drift guard in
`src/tool-schema.test.ts` makes two separate assertions about every shipped file under `src/`
(recursing, with `src/testing/` excluded by name):

- **No three-argument call.** It counts arguments by walking brackets, after blanking comments
  and string, template and regex literals, so it catches any indented call across any number of
  lines whatever sits in the replacer slot. Matching only the literal `null, 2` spelling would
  miss `JSON.stringify(x, someReplacer, 2)` - which is precisely the shape `redactedJson` had
  before its indent parameter was removed, and so the most plausible regression there is.
- **No bare call outside the listed exemptions.** Every `JSON.stringify` written literally in a
  shipped file (with a plain dot or an optional one) must either go through a seam or be one of
  the enumerated non-payload uses, with its exact count pinned; referencing `JSON.stringify`
  without calling it is reported too, since an alias would be invoked out of the scan's sight.
  This half is not about bytes either. What it buys is that serialisation stays in **one place**:
  because neither seam takes an indent, keeping every payload on them makes "no payload can be
  pretty-printed" a property of two function signatures rather than of a text scan that has to
  keep finding every call site forever.

**What redaction does and does not cover, since it is easy to state this too strongly.**
`toolJson` is a bare `JSON.stringify` with no replacer and it is nearly every seam call site, so
routing a payload through it redacts nothing - the drift guard buys the single seam, not
redaction. `redactedJson` is the only serialiser that redacts, and it has one call site (the
bulk-operations result). Success payloads on every other path are unredacted, and were before
the compaction too. The **error** path is covered independently of both seams: every error reply
is redacted centrally in `index.ts`'s CallTool catch, which is where a bearer token in a server
error description is caught.

The reason is that indentation is bytes the caller pays for and nothing parses. Every payload
here is machine-read — an MCP client parses it, or a model reads it as data — and neither
needs the whitespace. It also scales with the number of JSON *tokens* rather than with the
content, so it costs most on exactly the payloads that are already the largest. Measured live
against one real account in August 2026, as a point-in-time reading rather than a rate:
**17.3%** of a 25-message `list_emails` page, **24.9%** of the same page under `raw: true`, **28.5%** of a `list_mailboxes` result (many small flat objects, so the
most delimiters per byte of content), and 6.5% of a single `get_email` (dominated by one long
body string, which carries no delimiters to indent). That is pure whitespace in every case:
the change removes no field and alters no value (#40).

**A payload inside a prose frame is still a payload.** A list result is a summary line, a
newline, then the JSON array; the bulk-operations diagnostic wraps its JSON in a heading and a
follow-up instruction. In both, the prose stays prose and only the JSON is compacted. Treating
the framed ones as a readability exception would be a carve-out the next reader has to
remember, and whitespace is not what makes a payload legible: its structure and field names
are, and compacting changes neither.

The rule is about **payloads**, not about every `JSON.stringify` in the tree - which is why the
in-`src/` exceptions are enumerated in the test rather than left to each reader's judgement.
Two things outside `src/` stay indented on purpose, because their output is not a tool result:
`scripts/dump-official-surface.mjs` writes a checked-in JSON document, where indentation is
what keeps the diffs readable, and `scripts/mcp-harness.mjs` prints to a console for a human
running it by hand.

## Query-level signals: the result count and `position` paging

`total` and `nextPosition` are properties of the *query*, not of a message, and that
placement decides everything else about them. They live in the summary line built by
`formatQuerySummary` (`src/response-formatters.ts`), which both the simplified and the
raw path render, so the two cannot drift — and, like the Trash/Spam exclusion note,
they are never projected away by `fields` (see the out-of-band rule above). The raw
path is not a pure JMAP passthrough here: it already carried a summary line, so the
signals belong on it too; a `raw` caller additionally has the JMAP response's own
`total`/`position` if it re-queries.

- **`total` is always stated.** The old summary printed `20 results.` when a page
  happened to fill, which reads identically whether 20 is the whole match set or just
  the first page. A caller reading a capped page as the whole answer concludes "nothing
  else matched" when plenty did (#51) — the worst failure shape for a sweep, since it
  is a false negative with no signal. When the server declines to compute a total
  (`calculateTotal` is discretionary, RFC 8620 §5.5), the summary says the count was not
  returned rather than substituting the page size.
- **`nextPosition` appears only while more results remain.** Its absence is the
  published "the listing is complete" signal, so there is no `hasMore: false` to
  interpret — the same discipline as the exclusion note, where silence means nothing was
  withheld. The arithmetic is `position actually served + items actually returned`,
  never the requested `limit`, so a short final page ends the listing instead of
  advertising a page that does not exist.
- **One seam.** `position` is threaded through `runFilteredQuery` into `Email/query`, so
  `list_emails`, `search_emails` and `get_recent_emails` inherit it together (the last of
  those through `getEmails`, which it delegates to). The filters (including the default Trash/Spam
  exclusion, which lives inside the JMAP `filter` as `inMailboxOtherThan`) are applied
  server-side to every page, so paging cannot change what matches, and the hidden-count
  query stays a count — it carries no `position`. That count is over the whole filtered
  set, not the page, so the withheld-count note repeats identically on every page and
  must never be summed across pages (the tool descriptions and README say so).
- **`position: 0` and an omitted `position` are the same request.** 0 is the JMAP
  default, so the parameter is only put on the wire when it is non-zero.
- **A position past the end is not an error.** JMAP clamps it and returns an empty page
  with the real total, which is self-describing ("0 of 137 results from position 500")
  and idempotent, so the caller sees it overshot and re-asks. Rejecting it would add a
  failure mode for a case that already explains itself.
- **The coercion is the loud-reject side of the lenient-value rule** (`coercePosition`
  in `src/coerce.ts`): a stringified `"40"` is accepted, but a negative value, a
  fraction, or garbage is refused. Negative is the load-bearing one — JMAP reads a
  negative position as an offset from the *end* of the results, so accepting `-1` would
  quietly serve the last page. Reading from the other end is what `ascending` is for.

**`nextPosition` is gated on the calling tool accepting `position`.** `formatQuerySummary`
takes a `paged` flag, and only the three email tools set it. The contacts listings render
through the same summary (they get the always-stated total, which is an improvement
everywhere) but never the `nextPosition` clause: they declare no `position` parameter, so
a caller following that instruction would have the call rejected outright by the
unknown-parameter guard — an instruction the caller cannot act on is worse than none.
This is carried by *which renderer the handler picks*, not by a flag at every call site:
`formatRawEmailQueryResult` (paged) versus `formatQueryResult` (not), because a forgotten
flag would silently drop a promised signal while a wrong function name is visible in the
handler. Paginating the contacts and calendar protocol paths is tracked on
[#51](https://github.com/JonathanGodley/fastmail-mcp/issues/51).

**The CalDAV calendar listing joins the same discipline, over a different protocol.**
`list_calendar_events` does not go through JMAP at all, so nothing hands it a server-computed
`total` — it counts what it gathered. `getCalendarEvents` therefore returns
`{ events, total }`, where `total` is the number of events that matched *after* the window
re-filter and *before* `limit` trimmed the list, and the handler renders it through the same
`formatQuerySummary` as everything else (unpaged, so no `nextPosition`, since the tool takes
no `position`). The count is load-bearing here rather than cosmetic: recurrence expansion
turns one fortnightly series across a quarter into seven rows, so the cap is reached far
sooner than it was when a series counted once, and a caller reading a capped page as the
whole answer is exactly the false negative this section exists to prevent (#64, #100).

Its sibling failure is disclosed the same way, by refusing to answer at all. Calendar
*discovery* used to return an empty list on a server failure, which the listing reported as
a successful empty result — "you are free" for a question about availability. That now
raises. Where the email tools express a degraded read as an explicit note, the calendar read
path expresses it as an error, because there is no partial answer to annotate: with no
calendars there is nothing that could have matched.

## `hasAttachment` is a server heuristic — passed through by design

`hasAttachment` in simplified output is the server's value, untouched. This is a
deliberate decision (researched for #59, 2026-07-03), not an oversight; do not "fix" it
by deriving our own boolean from the `attachments` list.

**What the server actually does.** RFC 8621 leaves `hasAttachment` to server discretion
(the SHOULD is disposition-based; a MAY allows arbitrary heuristics). Fastmail's
implementation — upstream Cyrus, which Fastmail maintains and runs in production with
only a handful of site patches — ignores the disposition rule and answers the semantic
question "is this part content or decoration?" (`jmap_email_hasattachment` in Cyrus
`imap/jmap_mail_query.c`):

- image parts: `Content-Disposition: attachment` → true; otherwise **true iff both pixel
  dimensions ≥ 256** (signature logos excluded by size, pasted screenshots included);
  **unknown dimensions → true** (falls back to the false positive, the safe direction);
- any part with a filename → true; PDFs, `message/*`, `text/rfc822`, `text/calendar` → true;
- PGP/S-MIME signature parts and unnamed `application/octet-stream` → false.

**Why passthrough beats a deterministic derivation.** A spec-anchored derivation
(`attachments.length > 0`, which deterministically includes every cid inline image) was
considered and declined: it would flag every corporate reply carrying an `image001.png`
signature logo as `hasAttachment: true` — amplifying exactly the decoration noise the
Cyrus heuristic is built to filter, and burying the signal (the #59 incident was an agent
wasting attention on two signature logos). No vendor does better: MS Graph excludes
inline-only by design (a documented complaint generator), Gmail and Thunderbird have
their own long-standing inline-image gray zones. There is no industry-standard answer;
Fastmail's dimension heuristic is the most thoughtful of the lot, and passthrough also
keeps the field consistent with the server-side `hasAttachment` **search filter** (RFC
8621 §4.4.1 defines the filter as a direct comparison against this property) and with
what Fastmail's own UI shows.

**The residual accepted:** the value can change if Fastmail evolves the heuristic, and a
sub-256px image that IS content (a small but meaningful figure) reads as false. Both are
inherent to any heuristic; the mitigations are `bodyTextSize` (an agent can see there is
body to read regardless) and `get_email`/`get_email_attachments` for ground truth.

**Ground truth now includes embedded images, and the divergence is expected.** The part
listing is the union of the JMAP `attachments` array and the media parts routed into the
body lists, so a message reporting `hasAttachment: false` can list an inline logo. That
is the two fields answering different questions, not a bug — do NOT "reconcile" them by
deriving the flag from the listing, which is precisely the declined derivation above.
The consequence worth stating in consumer docs (and stated in the tool descriptions) is
that the **search filter** inherits the heuristic: RFC 8621 §4.4.1 defines
`hasAttachment` as a direct comparison against this property, so `hasAttachment: true`
filters embedded-image-only mail OUT. There is no server-side filter for embedded
images; narrow with other filters, then read parts with `get_email`.

## Two cid consumers, two staging orders — by provenance

A `cid:` value reaches this server from two directions, and the two are deliberately
staged differently. Do not "unify" them.

- **A reference in HTML** (`<img src="cid:...">`) is a URL, so per RFC 2392 its value is
  percent-encoded. It must be DECODED FIRST and the decoded key compared against part
  Content-IDs. `cidKey` (`src/inline-images.ts`) is that key function; the collecting
  sanitizer pass reports references already in that decoded form, so the compose and edit
  paths that match body references to parts inherit the staging rather than re-deriving
  it. `cidKey`'s other production caller is the download path below, as its fallback.
- **`download_attachment`'s `cid:<value>` parameter** is a handle that round-tripped from
  `get_email`'s output, where the `cid` is echoed VERBATIM. It is therefore compared
  LITERALLY first, and only falls back to the same `cidKey` decode when the literal
  matched nothing. Decoding first would break pasting back a Content-ID that genuinely
  contains a percent escape: `cid:%78` would silently resolve to the part whose id is
  `x`. The two consumers share the helper; what differs is which comparison runs first.

Both sides share one rule: only the REFERENCE is ever decoded. A part's own `cid` is a
Content-ID, not a URL, and is compared literally on both paths — decoding it too would
let a reference spelled `x` resolve to a part whose id really is `%78`. Within the `cid`
form, ambiguity is rejected rather than resolved by falling through to the next stage:
two parts sharing a Content-ID are genuinely different content, so picking one would be
a guess. That rule is specific to `cid` — the `blobId` form takes the first match on
purpose, because blobs are content-addressed and parts sharing one ARE the same bytes.

## Authoring a cid: what a caller may supply, and where it is decided

An `attachments` item may carry a `cid`, which is what makes the file display inside the
body instead of hanging off the end (#13). That value is a third provenance, and it is
neither of the two above: it is not a URL to decode, and it is not a handle that
round-tripped from this server's own output. It is caller-chosen text that ends up in a
MIME header, so it is the one direction that gets a vet rather than a comparison rule.

- **Vetted at coercion, in `coerceAttachments`.** A spelling copied out of HTML
  (`cid:logo`) or out of a header (`<logo>`) is pre-stripped to the same canonical value,
  then that value must be a simple token: letters, digits, dot, dash, underscore, up to
  64 characters. Everything else is refused by index, naming the item. The narrow shape is
  not decoration — a value carrying CR or LF stores as a genuinely injected MIME header
  and is invisible in the JMAP read-back (see `security-model.md`).
- **Canonical from that point on.** Every later comparison — collision between two items,
  a body reference resolving to a part, the reserved-shape check — uses the post-strip
  value, so `cid:logo`, `<logo>` and `logo` are one identifier and cannot collide with
  each other by spelling.
- **Matched against the body at compose time, not at send time.** The compose tools decide
  which supplied files the message displays and mark exactly those `inline`; `send_draft`
  submits a draft by reference and never re-validates its images.
- **Never a silent demotion.** A supplied file whose identifier nothing displays — no html
  body ships, or the body references something else — is still attached as an ordinary
  file, and the result says so. A body reference with no matching item is the reverse case
  and is refused outright, because that one ships visibly broken.
- **The refusal is honest about the opt-in.** When NEITHER attachment source is enabled —
  no `FASTMAIL_ATTACH_DIR` for a local file and no `FASTMAIL_ALLOW_BLOB_ATTACH` for content
  already in the account — there is no way to supply the missing file at all, so the wording
  drops the "add it to attachments" repair rather than pointing at a parameter that cannot
  work. Either gate being open restores the repair, because either source can supply the
  image. The combined flag is threaded into the refusal builders for exactly that reason.

## Index tightening: `download_attachment`'s entry-number form

`attachmentId` originally accepted anything `parseInt` would swallow as an array index.
That was leniency in the wrong place, and the decision has been REVERSED: the form is now
`/^\d+$/`, and a value like `3a`, `-1` or `1.5` is rejected instead of indexing.

**The decline history matters, because the reasoning changed rather than the taste.**
Lenient value coercion (the section at the top of this file) exists so a client that
stringifies `20` into `"20"` still works — the coercion recovers the caller's evident
intent. `parseInt("3a")` does not recover an intent; it invents one, and the result is a
silently wrong file, downloaded successfully, with no error to notice. That is the same
failure mode the unknown-parameter guard exists to prevent.

Two further reasons specific to this parameter, both new since the original decision:

- The reference space now has four forms sharing one string field. A permissive numeric
  reading makes a mistyped `partId` or `cid` land on an unrelated entry rather than fail.
- Digit strings are ambiguous here on purpose: Fastmail partIds ARE digit strings, so a
  digit resolves as a partId FIRST and only falls through to the entry-number form when
  no part claims it. Precedence like that is only safe when the fallback form is exact.

Entry numbers stay supported (they are the cheapest reference for a one-shot download)
but they are positional. Adding embedded images to the listing did NOT re-base the
existing ones: the union emits the JMAP `attachments` array first, in server order, and
appends body-routed parts after it, so indices into the old listing still mean what they
did. They remain unstable in general — any change to what the message or the server
reports moves them — which is why the tool description says to prefer a
`partId`/`blobId`/`cid` for any reference that will be reused.

**The form is READ-ONLY: it is refused when the reference feeds outgoing mail.** The same
`attachmentId` grammar is now accepted in a second place — an `attachments` item naming a
part of an existing message — and there the positional form is rejected. The asymmetry is
the consequence, not the direction of travel: a wrong download is a wrong file on disk and
costs one retry, while a wrong attach is baked into a draft that `send_draft` then
transmits, with nothing on the sent message to show it was wrong. So the resolver reports
WHICH form matched (`partId` / `blobId` / `cid` / `index`) and the compose path refuses
`index`.

That refusal is defined by **how the resolver matched, never by how the string looks**.
A looks-numeric test would fail in both directions at once: `parseInt` accepts `"2abc"`
(which the resolver refuses outright as unusable), and a real `partId` of `"2"` is a
perfectly valid exact reference that such a test would wrongly refuse. This is the same
trap the `parseInt` reversal above closed, met from the other side.

## Draft provenance: how a draft names the message it came from

A draft composed from an existing message records which message that was — in headers,
so the record rides the draft itself and survives everything a draft survives (a new
session, an edit, another client). Four surfaces read that record (`reply_email`,
`forward_email`, `send_draft`, and `edit_draft`'s quote guard), so the model belongs here
rather than in any one of them. The record has two tiers: **which message** (by
Message-ID, interoperable, set by other clients too) and **which stored copy of it**
(by JMAP id, this server's own header).

- **The two kind headers.** `In-Reply-To` marks a reply (set by `reply_email` and by every
  other client's reply); `X-Forwarded-Message-Id` marks a forward (set by `forward_email`,
  and by Fastmail's own clients — see `docs/email-bodies.md` for the forward-threading
  rationale and the value-validation rules). Both are JMAP `MessageIds`, i.e. **bare** ids
  with no angle brackets.
- **Reply wins when a draft carries both.** Dispatch is mutually exclusive and
  reply-first, in `edit_draft`'s guard variant and in `send_draft`'s keyword maintenance
  alike. This server never writes both; a foreign draft that does is treated as a reply.
- **Absent provenance is a real state, not a failure.** An ordinary compose records
  neither header; anything keyed on provenance is inert on such a draft — `send_draft`
  marks nothing. Both forward shapes DO record the header, including `asAttachment`
  (originally it did not, on the theory that the attached `.eml` was its recorded
  source — reversed under draft-first, #32/#66, because the `.eml` is not
  machine-resolvable as provenance and `send_draft` is the only transmit path left, so
  a header-less asAttachment forward's original could never be marked forwarded).
  `edit_draft`'s guard does not arm on the bare header when the draft carries a
  `message/rfc822` attachment: the forwarded content lives in the `.eml`, which body
  edits can't drop and the recreate preserves alongside the header.
- **The exact-instance header: `X-Fastmail-MCP-Source-Id`.** A Message-ID names a
  *message*; an account can hold several stored copies of one message (a duplicate
  delivery, or a self-addressed copy the user filed into another folder), and only the
  compose call knows which copy the caller actually had in hand. So `reply_email` and
  `forward_email` (both shapes, including `asAttachment`) also record the JMAP id of the
  fetched original in this header. That id is what lets `send_draft` mark exactly the
  copy the caller composed from — matching Fastmail's own client, which marks the
  instance replied to and leaves other copies of the same Message-ID untouched
  (observed live, 2026-08-14: replying to the Archive copy of a self-addressed message
  set `$answered` on that copy only, never the Sent twin). The recorded id is surfaced
  by `get_email` (and `get_thread` with `includeBodies`) as `sourceEmailId`, so which
  copy will be marked is inspectable before the send.
- **The exact id is validated before use; the Message-ID lookup is the fallback, not a
  peer.** `send_draft` checks that the recorded instance still exists and still carries
  the Message-ID the kind header names. A destroyed instance, a mismatch (a stale or
  hand-set pointer), or a failed read falls through to the lookup below rather than
  marking on faith. The value is also vetted at the single seam that writes it
  (`createDraft`): only an RFC 8620 URL-safe id shape (`[A-Za-z0-9_-]{1,255}`) is
  stamped, and anything else degrades to absent — the fallback covers a draft without
  the header, so a malformed value is never worth an error.
- **`edit_draft` carries the exact id like the kind headers.** The immutable-email
  recreate copies it verbatim; the forward keep path (`originalEmailId`) re-points it at
  the newly-named original alongside the Message-ID re-point; `noQuote` on a forward
  draft drops it together with `X-Forwarded-Message-Id` (the draft stops being a
  forward), while a reply draft keeps it under `noQuote` just as it keeps `In-Reply-To`.
- **Message-ID to JMAP id is a lookup, and it is two steps.** Every write path needs a
  JMAP id, so the header value has to be resolved: a full-text `Email/query` on the
  **bare** id is the recall step (the bracketed form matches nothing — the platform fact
  and its probe are recorded in `docs/email-bodies.md`), and keeping only the messages
  whose own `messageId` equals the target is the precision step. The recall step is a text
  search, so it also returns every message that merely *mentions* the id: thread siblings
  whose `References` carry it, quoted bodies, the sent copy itself. The query sorts
  **oldest-first**, which is load-bearing rather than cosmetic — a message can only
  reference an id that already exists, so the message that owns it predates every mention
  and lands on the first page whatever the result cap is.
- **The lookup can legitimately fail to identify one message**, and callers must not
  guess: no match, or more than one (a duplicate delivery, or a self-addressed message
  held in both Sent and Inbox), means the resolution is unknown. `send_draft` reports the
  skip rather than marking an arbitrary candidate — this is exactly the case the
  exact-instance header exists to avoid, and it remains reachable on drafts that lack
  the header (older drafts, foreign clients, a pointer that failed validation).
  `edit_draft`'s guard never resolves at all and instead requires the caller to pass
  `originalEmailId`, so a quote is never rebuilt from a message the caller didn't name.

Platform facts behind the design (live-probed 2026-08-14 against Fastmail):

- **Fastmail's own reply marks the exact instance**, not every copy of the Message-ID —
  the behaviour the exact-instance header replicates.
- **A Fastmail-UI reply draft records nothing richer than `In-Reply-To`/`References`**,
  plus a private `X-PersonalityId` (an internal sending-identity id). There is no
  platform-provided exact-instance record to reuse, so this server writes its own.
- **A Fastmail-UI draft edit keeps its own convention headers but DROPS truly foreign
  ones** (their edit also recreates the message under a new id).
  `x-forwarded-message-id` came through an edit intact (Fastmail recognizes it as its
  own convention), but `X-Fastmail-MCP-Source-Id` did not (probed live 2026-08-15: a
  tool-made reply draft edited in the UI came back with `In-Reply-To` preserved and the
  source-id header gone). So a UI-edited reply/forward draft loses its exact-instance
  record, and `send_draft` degrades to the Message-ID fallback — the designed path for
  record-less drafts, and the reason that fallback stays load-bearing rather than
  vestigial. A forward's edit guard is unaffected: it keys on `x-forwarded-message-id`,
  which survives.
- **EmailSubmission transmits stored headers verbatim.** The delivered copy of a
  self-forward still carried `x-forwarded-message-id`, and a reply sent from Fastmail's
  mobile app arrived still carrying `X-PersonalityId`. Header stripping at send is a
  per-client habit (their webmail strips private headers; their app does not), not a
  platform guarantee — so `X-Fastmail-MCP-Source-Id` IS transmitted to recipients.
  Why that is accepted is recorded in `docs/security-model.md`.

## Calendar DTSTART/DTEND: four time frames, and why they must agree

An iCalendar date/time property (RFC 5545 §3.3.4 / §3.3.5) carries its value in one of
four **frames**, and the frame — not the digits — decides what the value means:

| frame | shape | means |
| --- | --- | --- |
| date | `DTSTART;VALUE=DATE:20260320` | an all-day value, no time at all |
| floating | `DTSTART:20260320T093000` | whatever the local clock says; a different instant for every reader |
| UTC | `DTSTART:20260320T093000Z` | one fixed instant (a caller-supplied `+HH:MM` offset is normalised to this) |
| zoned | `DTSTART;TZID=Europe/Rome:20260320T093000` | one instant, resolved through a named zone |

Two properties are comparable to each other only inside one frame. That is the whole
reason `validateDateConsistency` in `src/caldav-client.ts` checks frame agreement *before*
ordering: a floating `DTSTART` beside a UTC `DTEND` has no single duration to order, and
writing the pair produces an event that is zero-length in UTC, backwards in UTC+10, and a
different length for every attendee. The ordering rule itself is RFC 5545 §3.8.2.2 —
`DTEND` strictly later than `DTSTART`, with `DTEND` exclusive for all-day events.

Three properties of the implementation are load-bearing and easy to undo by accident:

- **Classification runs on the serialized line, not on the caller's raw input.**
  `formatDateTimeProperty` rewrites a floating input to carry the stored event's `TZID`
  when there is one (the long-standing preserve-the-timezone behaviour), so by the time
  `describeDateProperty` sees it, it is already `zoned` and agrees with its `zoned`
  partner. Classify the raw input instead and that legitimate case starts failing, while
  the floating-against-UTC case it exists to catch starts passing.
- **The comparison is against the value that will sit beside it, not just the caller's
  own arguments.** On an update, an untouched side is read from the *stored* VEVENT. Both
  defects this check closes are single-sided updates, so comparing only what the caller
  passed sees nothing. The visible consequence: moving an event to another day, or
  converting one side to UTC, requires passing both `start` and `end`.
- **Two `zoned` values in different zones are accepted, and only their ordering is
  skipped.** A flight departing Rome and landing New York is a legal VEVENT whose wall
  clocks read backwards, so the check declines to guess rather than reject valid travel
  events — which does mean a genuinely backwards cross-zone pair is written. That is now a
  choice rather than a limit: both zone names are present here and ICU resolves them
  (`zoneOffsetMsAt`, which the read window depends on). Adding the check would newly reject
  input the tool accepts today, so it is tracked as
  [#140](https://github.com/JonathanGodley/fastmail-mcp/issues/140) rather than folded in.

The frame check is deliberately *not* applied when the caller touches neither `start` nor
`end`: it exists to stop us writing a broken pair, not to hold a title edit hostage to an
inconsistency some other client left in the event.

### The read path carries the zone name now; its window filter still only widens, on purpose

The four frames survive intact on the *write* path, where the property line is built and
inspected whole. Reading USED TO be lossy: `formatICalDate` takes only the property's
**value**, so `DTSTART;TZID=Australia/Sydney:20270305T083000` and a floating
`DTSTART:20270305T083000` both become the bare `2027-03-05T08:30:00`, and by the time a
`CalendarEvent` existed the frame had collapsed to indistinguishable digits — the **name**
was gone and nothing downstream could get it back.

That is fixed ([#139](https://github.com/JonathanGodley/fastmail-mcp/issues/139)). `start`/
`end` themselves are UNCHANGED — still a bare local wall clock, a `Z` instant, or a date-only
value, because this server never puts an offset in a calendar value on read any more than it
accepts one on write. What changed is that `parseVEvent` now ALSO reads the raw DTSTART/DTEND
**lines** (`parseAllICalProperties`, not just `parseICalValue`) and classifies each one with
`describeDateProperty` — the same classifier `validateDateConsistency` uses on write, so read
and write agree about what a `zoned` value is — and carries the result forward as `timeZone`
(describing `start`) and `endTimeZone` (describing `end`, relative to `start`) on the parsed
`CalendarEvent`. Never an offset: the whole reason for a NAME here is the same reason
`validateDateConsistency`'s zoned/zoned exemption exists on write — an offset is only valid
at one instant, and delegating DST to the reader's own zone database is the entire point.

**The emit rule.** `timeZone` is set from `start`'s classification: a `tzid` differing from
the account's configured zone emits the name; a `tzid` MATCHING it omits, because the
overwhelming majority of rows are already in the configured zone and silence there is the
token-cheap default (a caller reads "absent" as "the configured zone"); `floating` emits
`null`, because "no zone, a different instant per reader" is the whole fact and there is
nothing else to say; `date` (all-day) and `utc` (`Z`-suffixed) both OMIT rather than emit
`null` — both frames already fully describe themselves, and `null` there would assert
"floating", a different and wrong fact. `endTimeZone` applies the identical rule to `end`,
but compared against `start`'s own classification (falling back to the configured zone only
when `start` itself is absent) rather than against the configured zone directly — RFC 5545
§3.8.5.3 and `validateDateConsistency`'s zoned/zoned exemption both permit a start and an end
in two different named zones, the flight-lands-elsewhere case (#140), so `endTimeZone` is how
that legal shape is disclosed on read. A `DURATION`-computed `end` has no raw `DTEND` line to
classify at all, so it reads back `absent` and `endTimeZone` omits — a computed end shares
`start`'s zone by construction, so there is nothing to disagree about.

**One: the returned `start`/`end` is still a bare wall clock, but a caller no longer has to
guess whether that means "floating" or "the configured zone".** Both read tools' descriptions
state the emit rule once, naming the configured zone, rather than leaving a model to infer it
from a disclosure written for the old, lossy behaviour.

**Two: results are still ORDERED by a resolved instant, but now in the event's OWN zone when
it resolves, not only the configured one.** `sortEventsByStart` reads `event.timeZone`; when
it is a `tzid` `isUsableTimezone` can resolve, that zone is used — a correct reading, not the
old best-effort one. The configured zone remains the fallback, for two cases only: a
genuinely floating `start` (there is no zone to sort it in, so the caller's own clock is the
least-wrong guess) and a `timeZone` this server was HANDED but cannot resolve (a Windows zone
name such as `AUS Eastern Standard Time`, passed through verbatim rather than rejected — see
below). That guard is not cosmetic: `zoneOffsetMsAt` silently falls back to the HOST zone for
an unresolvable name, so sorting by an unresolvable `timeZone` directly would place that one
event in the *host's* zone rather than the account's configured one — a regression on the
zone-blind behaviour this replaces, which is exactly why `isUsableTimezone` gates it.

**Non-IANA names pass through verbatim, on purpose.** A calendar written by a non-Fastmail
client can carry a Windows zone name in its `TZID` rather than an IANA one. This server does
not reject it or try to resolve it: `timeZone`/`endTimeZone` reports the stored name exactly
as written, because the offset-shape alternative could not represent it at all (ICU will not
resolve a Windows name, so there is no offset to compute), and a name-carrying shape at least
lets a caller that understands Windows zone names use it. `resolveUsableTimezone` in
`src/coerce.ts` is the one place that decides which zone is actually IN FORCE for reads —
`describeTimezone` and every calendar read path call through it rather than re-deriving the
host-zone fallback separately.

**Three: the window re-filter can still only widen, and that is now a choice, not a gap.**
`eventIntersectsWindow` (the client-side re-filter behind `list_calendar_events`) is built to
**drop only what provably cannot intersect**, rather than to decide membership, and it is
unchanged by #139: `MAX_UTC_OFFSET_MS` still grants ±14 hours of slack to any
zone-designator-less value, on both edges, so a zone-carrying event near a window edge is
kept even though its true instant is unknown here. Before #139 that width was FORCED — the
parser had already discarded the name, so there was nothing narrower to check a value
against. The name is no longer missing: it sits on the parsed event as
`timeZone`/`endTimeZone`. The slack stays wide anyway, deliberately, because narrowing it to
a per-event zone would trade one visible extra row today for an invisible missing one the
moment a name fails to resolve, is spelled unusually, or belongs to a value
`eventIntersectsWindow` is not even handed — its signature takes only `start`/`end`, not the
zone fields, so tightening it is a real, separately-tested follow-up, not a one-line change
folded in here, tracked as #162. The residue this leaves is accepted deliberately for a second
reason too: the
CalDAV server is the authority on time-range matching (RFC 4791 §9.9, occurrence-based), and
this filter exists only to catch what that matching cannot express — the server matches per
*occurrence* but returns whole *resources*, so an unexpanded series arrives showing a master
`DTSTART` that may be years outside the window (#64). Erring the other way would discard
events the server correctly matched, which turns a busy day into a free one; an extra row a
caller can see and dismiss is the cheaper error.

**The residue that actually shows up is an ALL-DAY event, not a zoned one.** Measured live:
`{startDate: "2026-08-01", endDate: "2026-08-31"}` on a UTC+10 account returns
`{"title":"Newcastle","start":"2026-07-31","end":"2026-08-01"}` as its first row. The window
begins at local midnight (`2026-07-31T14:00:00Z`), Cyrus matches a `VALUE=DATE` event on a
UTC day, and the ±14h slack — granted to *any* designator-less value, which includes a bare
date — keeps it.

**That single measurement does NOT fix the DIRECTION, and reading it as a rule that it does is
wrong for half the world's accounts.** The slack is granted on BOTH edges, so the surviving row
can sit outside EITHER end; which end you see follows the SIGN of the account's UTC offset,
because that is what decides which neighbouring UTC day the server itself matched. Measured
both ways:

| account | window | day before | day after |
| --- | --- | --- | --- |
| Australia/Sydney (+10) | `2026-07-31T14:00Z .. 2026-08-31T14:00Z` | server returns it, filter keeps it | server does not return it |
| America/New_York (-4) | `2026-08-01T04:00Z .. 2026-09-01T04:00Z` | server does not return it | server returns it, filter keeps it |

Nor is the reach one day: 14 hours of slack applied to an all-day `end` that is already the
following midnight lands nearly two calendar days out east of UTC+11. "Bounded at one day" is
not a property of this filter; "only ever ADDS rows" is. So it is a reporting defect rather
than data loss, and `list_calendar_events`' description warns to check each `start` against the
window. Narrowing it would mean resolving date-only values in the configured zone instead of
granting them slack, which is a behaviour change to the filter and is not made here.

**One thing the filter is NOT allowed to judge: a block still carrying its `RRULE`.** The
windowed path runs this filter over blocks the server was asked to EXPAND. If a server ever
declines to expand one, the master arrives with its rule intact and its ORIGINAL `DTSTART` —
years before the window for a long-running weekly event — and judging that date deletes the
row, turning a wrongly-dated event into a MISSING one, which is the exact direction the filter
exists not to fail in. A recurrence rule is proof that `DTSTART` is not the only date the event
has, so `getCalendarEvents` keeps such a block whatever its dates say. Unreachable against
Fastmail today (measured in `calendar-expand.probe.mjs`), and it stays as resilience rather
than being removed, because the claim that this filter "closes the gap if a server declines to
expand" was false in the dangerous direction until the guard existed.

The window's own bounds are a separate concern from an event's stored zone entirely: they
never carry a TZID at all, and are normalised once by `coerceCalendarWindowStart` /
`coerceCalendarWindowEnd` and used for both the server's `time-range` and this filter, so the
two cannot disagree about which days were asked for.

**That is a statement about the bounds agreeing with EACH OTHER, and it is not the whole
story — the bounds had their own zone bug (#138).** Read the paragraph above as "the bounds
cannot drift apart from each other", never as "the bounds are fine". They were resolved as UTC days, so on a UTC+10 account
`startDate: 2026-08-12` searched 12 Aug 10:00 to 13 Aug 10:00 local and a day holding three
appointments answered with one. Two missing events look exactly like a quiet morning, which is
the same silent-under-report failure the rest of this section is about, arriving through the
argument rather than the payload.

### Writing a zone: `timeZone`, and why create and update disagree on the default

The read half above restored the NAME on the way out; `timeZone` on `create_calendar_event`
and `update_calendar_event` ([#157](https://github.com/JonathanGodley/fastmail-mcp/issues/157))
is its write-side counterpart, and it keeps to the same rule: this server never puts an offset
in a calendar value and never asks a caller to compute one. `timeZone` only ever qualifies a
**designator-less** `start`/`end` — a value already carrying `Z` or a numeric offset names its
own instant and `timeZone` would be redundant at best and contradictory at worst, so it is
rejected rather than silently ignored (below).

**Precedence, in the order `formatDateTimeProperty` actually applies it:**

1. `callerZone` — the validated, **canonicalised** `timeZone` argument, when the caller passed
   one. `validateCallerTimezone` returns ICU's canonical spelling for whatever the argument
   resolved to, not an echo of what was typed - see the round-trip note below.
2. the STORED `TZID`, read off the existing property line — **update only**, and only when
   `timeZone` was not supplied. This is the pre-#157 preserve-the-timezone behaviour, kept
   byte-for-byte: an update that touches only one side of an already-zoned event keeps the
   other side's zone without the caller having to re-state it.
3. `defaultZone` — the account's configured zone (`getDefaultTimezone()`, `resolveUsableTimezone`
   gated) — **create only**. Also canonicalised: `resolveUsableTimezone` returns the SAME
   `canonicalZoneName` spelling `validateCallerTimezone` does, so the identical operator-configured
   string ends up as the identical written TZID regardless of which of the two paths supplied it.
4. floating — no `TZID` at all. This is now unreachable on create (step 3 always supplies a
   zone) and is exactly the pre-#157 behaviour on update.

**The write is canonicalised; the read is not, and that is a real round-trip asymmetry.** A read
emits a stored `TZID` verbatim (see "Restoring the NAME" above) - a stored `US/Pacific` reads
back as `timeZone: "US/Pacific"`, not `"America/Los_Angeles"`. Echo that `"US/Pacific"` straight
back as a `timeZone` argument and this server writes `"America/Los_Angeles"`: the same zone, a
different spelling than what was read. A caller comparing the two spellings as strings would
wrongly conclude the zone changed; comparing them as zones (`zoneNamesEqual`, below) is what
actually agrees they didn't. A link name reached through a **slash-qualified** alias
(`US/Pacific` → `America/Los_Angeles`) or a case difference (`australia/sydney` →
`Australia/Sydney`) are the shapes where this is visible on a value the write side still
accepts; an already-canonical spelling round-trips unchanged because canonicalising it is a
no-op.

**A stored BARE alias does not round-trip at all, on purpose ([#157](https://github.com/JonathanGodley/fastmail-mcp/issues/157) amendment).**
A stored `NZ` TZID still reads back as `timeZone: "NZ"` verbatim - the read side is unchanged -
but echoing `"NZ"` straight back as a `timeZone` write argument is now **rejected**, not silently
canonicalised, because a bare abbreviation or alias fails the slash rule below regardless of
where it came from. The rejection names the fix: pass the slash-qualified spelling
(`"Pacific/Auckland"`) instead. This is the deliberate cost of closing the ambiguity `EST` and
its relatives create - a caller who only ever received a canonical name from a read never hits
it, and one who is round-tripping a foreign client's bare-alias TZID gets an actionable message
rather than a silently-different zone.

**Create defaulting to the configured zone is a deliberate behaviour change, not a bug fix.**
Before #157, a designator-less `create_calendar_event` call wrote a bare floating value — "a
different instant for every reader" — silently, because nothing else in the call could name a
zone. That is a worse default than the configured zone for the overwhelming majority of events
a caller creates for themselves, and it is now what happens unless `timeZone` says otherwise.
**Update never defaults**, on purpose: unlike create, an update's untouched side may already
carry a real, meaningful `TZID` — quietly overwriting it with the configured zone the moment a
caller edits the *other* side would be a silent, unrequested rewrite of data the caller never
asked to touch. So omitting `timeZone` on update reproduces exactly what happened before #157
(step 2 or step 4), and reaching the configured zone requires naming it.

**`timeZone` provenance (`tzidSource`) exists so a rejection never misattributes a zone the
caller didn't choose.** `describeDateProperty` threads `'caller' | 'stored' | 'default'`
alongside `frame: 'zoned'` (`DatePropertyFrame.tzidSource`, set only where `formatDateTimeProperty`
itself set the `TZID`, never re-derived from the line). Without it, a `create` call that omitted
`timeZone` and got the configured zone by default would see its own ordering error worded as if
it had explicitly asked for that zone — a caller reading "you named X" when they named nothing
would go looking for a `timeZone` argument that was never in their request. `validateDateConsistency`'s
error text says "applied because you named none" for a `'default'` source instead.

**Rejections (fail-closed on a narrowing argument — the [same posture](#lenient-input-coercion)
`coerceStringArrayStrict` applies elsewhere, applied to this argument specifically):**

- **`timeZone` combined with a `Z`/offset-designated `start` or `end`** (`rejectTimezoneConflict`)
  — the value already names an instant; a second zone claim on top of it is a caller error, not
  something to silently prefer one of.
- **`timeZone` combined with a date-only `start` or `end`** (same function) — an all-day value
  has no time component for a zone to qualify.
- **`timeZone` that is `null`, empty, or whitespace-only** (`validateCallerTimezone`, `src/coerce.ts`)
  — a ratified decision, not an incidental gap: there is no way to ask this server to force a
  FLOATING write via `timeZone`, on either tool, ever. Omitting the argument is how a caller
  reaches step 2/4 above, and a caller who explicitly sends `null` meaning "make it floating"
  gets a clear rejection instead of a silent floating write that looks identical to the
  omitted-argument case.
- **`timeZone` that does not contain a region-qualifying slash, and is not exactly `"UTC"`**
  (`zoneRejectionReason`, `src/coerce.ts`, [#157](https://github.com/JonathanGodley/fastmail-mcp/issues/157)
  amendment) - a bare abbreviation or alias such as `"EST"`, `"NZ"`, `"PST"`, `"MST"`, `"GMT"` or
  `"Zulu"` is rejected even though ICU resolves every one of them to a real zone, because the
  resolution is ambiguous and one shape is actively dangerous: `"EST"` names a **fixed-offset**
  zone with no daylight-saving rule, not US Eastern, and every other bare alias is exactly as
  easy to misread. There is deliberately no safe-list of "harmless" abbreviations - the rule is
  the same string test for all of them, checked after the leading-slash strip and before
  canonicalisation, so `"US/Pacific"` (slash-qualified) still passes and still canonicalises to
  `"America/Los_Angeles"`. The check order matters: a totally unresolvable string (`"Blah"`)
  reports as unresolvable, never as "shorthand", because ICU never got to say whether it's
  ambiguous or simply wrong; only a string ICU *can* resolve, but with no slash and no `"UTC"`
  match, is a shorthand rejection. The comparison against `"UTC"` is case-insensitive
  (`"utc"`/`"Utc"`/`"UTC"` all pass) - it is the rule's one deliberate exception, not an
  oversight the rule forgot to close.
- **`timeZone` on an update that touches neither `start` nor `end`** — `timeZone` alone has
  nothing to qualify (there is no designator-less value in the call at all), so this is rejected
  before any patching happens, naming the fix: re-send `start` and/or `end`.
- **`timeZone` on a SINGLE-sided update whose untouched side is stored in a DIFFERENT named
  zone** (`rejectStrandedZoneMismatch`, both directions — updating `start` alone with the stored
  `end` elsewhere, and updating `end` alone with the stored `start` elsewhere) — writing only
  one side into the caller's zone while the other stays wherever it was stored would silently
  strand the pair across two zones the caller never asked to create. The comparison uses
  `zoneNamesEqual` (case- and separator-normalising, the same helper the read-half's
  provenance/comparison machinery already uses — `Australia/Sydney` and `australia/sydney` are
  the same zone here), so this rule is never tripped by a spelling difference alone. That
  reuse fixed a real bug during this work: `validateDateConsistency`'s own zoned/zoned
  stand-down (the flight-lands-elsewhere exemption, [#140](https://github.com/JonathanGodley/fastmail-mcp/issues/140))
  used to compare TZIDs with a raw `!==`, so two differently-spelled names for the SAME zone
  read as two DIFFERENT zones and stood the ordering check down — silently accepting a
  backwards pair that a same-spelling stored/caller pair would have correctly rejected. It now
  reads `zoneNamesEqual`, so ordering is checked whenever the two sides genuinely name the same
  zone, however each was spelled.
  `zoneNamesEqual` is also **link/alias-aware**, not just case- and separator-normalising:
  after the trim and leading-slash strip, each side routes through `canonicalZoneName`
  (`src/coerce.ts`, the same seam `validateCallerTimezone`/`resolveUsableTimezone` use to decide
  what gets written) whenever ICU can resolve it, so `NZ` and `Pacific/Auckland` compare equal -
  not just two case variants of one string. This closed a regression `validateCallerTimezone`'s
  own canonicalisation introduced: once the read half started emitting a stored `NZ` TZID
  verbatim and the write half started canonicalising a caller's `timeZone` to `Pacific/Auckland`,
  an ordinary read-modify-write caller echoing `timeZone: "NZ"` straight back was rejected by
  `rejectStrandedZoneMismatch` as a false two-zone mismatch - the stored side really was `NZ`,
  the caller really did name `NZ`, and the raw-string comparison called that "different" anyway.
  Since the slash-rule amendment above, a caller cannot reach this exact scenario with a bare
  alias any more - `validateCallerTimezone` now rejects `timeZone: "NZ"` before
  `rejectStrandedZoneMismatch` ever runs - but the identical link/alias-awareness is still what a
  slash-qualified pair needs (`US/Pacific` echoed back still has to compare equal to a stored
  `America/Los_Angeles`), and a foreign client's stored bare-alias TZID is still read back
  verbatim and still has to compare correctly against whatever slash-qualified spelling the
  caller now has to send instead.
  A name ICU cannot resolve (a Windows zone id, a vendor-prefixed TZID) falls back to today's
  plain string comparison, so it is a real rejection, not a guess. Cached by exact input string
  in `src/coerce.ts` - `zoneNamesEqual` runs once per event on every calendar list read, and
  constructing an `Intl.DateTimeFormat` per call is not free.
- **A GENUINE two-zone pair is still legal and still not rejected** — `timeZone` cannot itself
  produce one (it applies identically to whichever side(s) are designator-less in a single
  call), so the only way to reach it is the pattern `rejectStrandedZoneMismatch` exists to catch:
  single-sided updates are exactly where it would otherwise happen by accident, and #140's flight
  case is exactly where it should still be allowed on purpose. Passing both `start` and `end`
  together with one `timeZone` always lands both sides in the same zone by construction, so the
  two rules do not collide.

**The response states which zone was actually written.** `createCalendarEvent`
and `updateCalendarEvent` return a structured `CalendarZoneWriteInfo` per side — `{ kind: 'zoned',
zone }` / `{ kind: 'utc' }` / `{ kind: 'floating' }` / `{ kind: 'allday' }`, produced by
`classifyWrittenLine` reading back the very line just written — rather than the caller having to
re-derive it from what they sent; `create`'s result
always names both sides (both are always written), `update`'s names only the side(s) actually
touched. This closes the same gap a caller hits without it: a designator-less create silently
landing in the configured zone, or an update silently inheriting a stored zone, would otherwise
be discoverable only by reading the event back.

**The VTIMEZONE residual.** This server writes `DTSTART;TZID=<name>` / `DTEND;TZID=<name>` with
no accompanying `VTIMEZONE` component — `createCalendarEvent`'s iCal assembly never emits one,
and never has. Read from Cyrus's own source, not
verified against a live probe: `caldav_store_resource` (the function behind every `caldav_put`
path in `imap/http_caldav.c`) has no VTIMEZONE-presence precondition, so a bare `TZID=` reference
with nothing defining it is accepted at write time. On the server's OWN read/export paths (the
`GET`/`multiget` handlers in `imap/http_caldav.c`, and the JMAP/JSCalendar converters in
`imap/jmap_calendar.c` and `imap/jmap_ical.c`), Cyrus re-attaches the matching `VTIMEZONE` itself
via `icalcomponent_add_required_timezones` (`imap/ical_support.c`) before handing the object back
out — so a client reading the event back from Fastmail (including this server's own read path)
always sees a complete, self-describing object regardless of what was actually stored. This is
also the mechanism `ALLOW_CAL_NOTZ`'s `tzbyref` mode formalises server-side: `strip_vtimezones`
(`imap/caldav_util.c`) actively REMOVES a client-supplied `VTIMEZONE` on write when the namespace
allows it, on the same premise — the component is reconstructible from the `TZID` name alone, so
storing it is pure overhead Cyrus elects not to keep. The residual this leaves: an IANA name ICU
resolves but the SERVER's own tzdata does not recognise would round-trip as an unresolvable
reference with nothing here to catch it before the write — another argument, alongside the ones
in `validateCallerTimezone` itself, for keeping that gate narrow rather than widening it to
accept anything ICU-shaped.

### A calendar window's DAY is a local day

`coerceCalendarWindowStart` / `coerceCalendarWindowEnd` resolve a date-only bound — and a
datetime carrying no zone designator — in the **configured timezone**, not UTC. A value
carrying `Z` or a numeric offset is never re-read: it named an instant, so it stays that
instant, and it is the escape hatch from the local-day rule.

Three divergences meet here, and each is right for its own question:

| | date-only value means | why |
| --- | --- | --- |
| `search_emails` `before`/`after` (`coerceUtcDate`) | midnight **UTC** | compares against `receivedAt`, a JMAP UTCDate — an instant, not a day |
| `list_calendar_events` `startDate`/`endDate` | the whole **local** day | "what is on the 12th?" asks about the asker's own day |
| `create_calendar_event` `end` | exclusive at the **start** of that day | RFC 5545 DTEND, one edge of a single event |

The two coercions share `classifyDateValue`, so they reject an identical set of bad values with
identical wording and diverge only on what an accepted value resolves to. Do not "unify" them
back into one function: the shared half already is one function, and the half that differs is
the answer to a different question.

**The shared half does not cover the TIME components, and that gap has bitten once.**
`classifyDateValue` never reads the hour and minute out; `coerceUtcDate` gets its range check
for free from `new Date()` refusing `25:00:00`, and the calendar pair reads the components
itself with a shape-only pattern and hands them to `Date.UTC`, which **rolls** rather than
refusing. So `2026-08-12T99:99:99` silently became a window starting three and a half days
later, and `create_calendar_event` refused on a write exactly what the read accepted.
`isWallClockInRange` restores the parity (`24:00:00` is deliberately allowed, because the
ECMAScript date format allows it and the UTC coercion takes it). When you add a value the two
sides read differently, check the divergence rather than assuming the shared function covers
it.

Mechanics worth knowing before touching it:

- **The zone comes from `getDefaultTimezone()`**, the value `setDefaultTimezone` stores — the
  same one every email `date` renders in. It is read rather than re-derived from the
  environment on purpose: a second lookup is a second answer, and the two drifting would mean
  a calendar query's day and an email's displayed date disagree with nothing to say so.
- **The offset is sampled either side, and the candidate is CHECKED.** The offset has to be
  read at the instant the wall clock names, and that instant is what is being solved for — so
  `wallClockToUtcMs` reads the offsets in force a day either side and verifies a candidate
  answer against the offset actually in force where it lands. Without that a bound within a
  day of a DST transition lands an hour out.
- **A skipped or repeated wall clock resolves the way RFC 5545 and Temporal's `compatible`
  rule resolve it**: a repeated one takes the EARLIER instant, a skipped one resolves FORWARD
  BY THE LENGTH OF THE GAP. That equals the transition instant only for a clock sitting at the
  very start of the gap — `America/New_York 2026-03-08T02:00:00` is 07:00:00Z, but `…T02:29:59`
  is 07:29:59Z, half an hour past it. Forward matters. Taking whatever a blind second pass produced
  resolved a skipped time backwards, and for the exclusive END of a window that quietly
  dropped the last hour of the requested day — in a zone whose transition is at midnight
  (`America/Santiago` 2026-09-05, `America/Havana` 2026-03-07) a single-day window ran local
  00:00 to 23:00 and an event at 23:30 was never searched for. Sydney and New York transition
  at 02:00, so the deployment's own zone hid it; `FASTMAIL_TIMEZONE` is a config value.
- **A date-only end advances by a whole LOCAL day, not by 24 hours.** A day a DST change makes
  23 or 25 hours long would otherwise end an hour inside itself or an hour into the next day.
  The 366-day clamp on a one-sided window is the deliberate exception: it advances in fixed
  24-hour days because it invents a span nobody named, where a DST hour either side of an
  arbitrary bound is not a wrong answer to anything, and the clamp note states the instant it
  landed on.
- **Years below 0100 are handled explicitly.** `Date.UTC(26, …)` is the year 1926 — legacy
  two-digit-year mapping — so `0026-08-12` resolved to a window in the 1900s while
  `coerceUtcDate` correctly returned the year 26. `utcMsFromComponents` shifts by one whole
  Gregorian cycle (400 years, exactly 146097 days) to step over the mapping without disturbing
  the leap arithmetic. Unreachable in practice; it is the same silent-different-window class as
  the rest of this section, which is why it is fixed rather than noted.
- **An unusable IANA name falls back to the host zone** rather than throwing, matching
  `toLocalIso` on the same kind of bad zone string. `describeTimezone` names BOTH in that case -
  the host zone that resolved and the value that did not - because naming only the unresolved
  one would put the disclosure and the behaviour in disagreement on the one call where a caller
  is trying to work out why their days look wrong. `FASTMAIL_TIMEZONE` itself can no longer
  reach this branch in production: `resolveConfiguredTimezone`
  ([#157](https://github.com/JonathanGodley/fastmail-mcp/issues/157), see "Writing a zone"
  above) validates it at server startup - including the near-unreachable case where the HOST
  zone itself fails the rule - so `getDefaultTimezone()`, the only value every current caller of
  `describeTimezone` passes, is already guaranteed usable by the time a request runs this code.
  The fallback stays in `describeTimezone` itself rather than being deleted: it is a general
  utility with its own tests, not something entitled to assume every future caller pre-validates
  its argument the way today's callers happen to.
- **Test with an INJECTED zone.** The coercions take the zone as an argument for exactly this
  reason. A test that leaves it to the machine passes under both the UTC-day and the local-day
  reading whenever the host sits in the zone asserted — which is how the wrong-day window sat
  under a green suite. The suite pins Sydney and New York so a sign error cannot pass both.

### A one-sided window is bounded, and the bound is disclosed

`startDate` and `endDate` are both optional, so one of them may have to be invented. It used
to be invented as 1970-01-01 / 2099-12-31, which was harmless while the window was only a
filter and stopped being harmless the moment `expand` was added: the window is now the range
the SERVER materialises occurrences over, so `startDate: <today>` alone asked Fastmail to
generate every occurrence of every repeating event for 73 years. Nothing caps that on either
side — Cyrus's `expand_caldata` has no iteration limit, and here the whole response is
buffered, regex-split, parsed, filtered and sorted before `limit` ever applies, so `limit` is
not a bound on the work. Calendar content is also attacker-authored in this deployment:
anyone who can send an invitation can put a `FREQ=MINUTELY` series in the account.

So the INVENTED half is clamped to `CALENDAR_OPEN_WINDOW_DAYS` (366) from the bound that was
given, and the clamp is surfaced in `CalendarEventQueryResult.windowClamp` — STRUCTURE, not
finished prose. That is the shape the email listings already established for a disclosure of
this kind (`QueryResult.exclusion` -> `buildExclusionNote` -> handler), and the calendar path
briefly diverged from it: it built the sentence inside the client, which put the wording where
no formatter test can reach it and gave the blank-line separator a second home. The note
builder owns both.
A caller silently handed a narrower window than it asked for would read "nothing after that
date" as an empty calendar — the same never-silently-degrade rule the exclusion note and
`unresolvedMailboxIds` exist for. A window whose bounds the caller named is never clamped:
that span is the caller's own decision.

**Saturation is the other narrowing, and it names an EDGE.** A bound that resolves outside the
four-digit-year range every consumer of these values can express is pulled back to that range's
edge rather than rejected — the invented half, and a caller-named one too, since the local-day
rule resolves the caller's value through a zone and an offset alone can push `9999-12-31` over.
Both ends saturate, and the disclosure is an opposite statement at each, so `windowClamp.saturated`
carries `{ bound, edge }` rather than a bare bound name: a `startDate` pulled UP to year 0000 was
otherwise reported as having "resolved past the last date this server can express", the reverse
of what happened, and a window that saturates at both ends at once needs both sentences.

**The bound therefore covers ONE of the two ways to ask for an unbounded expansion**, and the
justification above does not stretch to cover the other. A caller naming BOTH bounds
(`2000-01-01 .. 2100-01-01`) still asks the server to materialise every occurrence over a
century, bounded by nothing — and while the caller chooses the SPAN, an attacker chooses the
DENSITY, which is the same `FREQ=MINUTELY` hazard the paragraph above just described. Latent
rather than active: a live 10-year window on this account measured 6.4s for 1237 events. Left
as it is deliberately, because a cap on a span the caller named is a behaviour change to a
user-visible contract rather than a fix.

## iCalendar structure is decided on WHOLE CONTENT LINES, never with a `/m` regex

Two different characters can end a line, and only one of them is a line break.

RFC 5545 §3.1 knows exactly one: CRLF (this server tolerates a bare LF too, because real
servers emit it). JavaScript's `^`/`$` under `/m` know three more — **U+2028 LINE SEPARATOR,
U+2029 PARAGRAPH SEPARATOR and a bare CR** — and none of those three has to be escaped inside
an iCalendar TEXT value. So all three survive verbatim into a `SUMMARY` or `DESCRIPTION`
written by whoever authored the event, and under a `/m` anchor every TEXT value becomes a
structure-editing primitive. Calendar content is attacker-authored in this deployment: anyone
who can send an invitation can put whatever they like in a description.

Three measured consequences, in ascending order of how bad they are:

1. `DESCRIPTION:hi<U+2028>END:VEVENT<U+2028>BEGIN:VEVENT…` cuts one component in two. The real
   `DTSTART`/`DTEND` land in the discarded tail, so the event comes back with **no dates at
   all** — and `eventIntersectsWindow` keeps a dateless event (it has nothing to judge), so it
   displays inside a window nothing placed it in.
2. On the expanded path the same payload yields two events, one wholly attacker-authored with
   an attacker-chosen `start`, and `blockCountProvesSeries` then marks the **real** event
   recurring because two blocks "prove" a series.
3. **A property read takes the FIRST match in the block**, so a `SUMMARY` of
   `Lunch<U+2028>UID:board-meeting@victim.example<U+2028>DTSTART:…` makes the attacker's own
   resource report **another resource's UID**. `delete_calendar_event` and
   `update_calendar_event` resolve an id through `findCalendarObjectByUID`, which returns the
   first object in the ACCOUNT whose UID matches — so an agent following the tool descriptions
   ("the id names the series, act on it") irreversibly destroys a record the caller never
   named, and mails its attendees a cancellation. The same trick places a fabricated
   appointment on any date of the user's calendar.

The write path inherits all of it: the RRULE test and the ORGANIZER/ATTENDEE presence gates
steered an in-place patch from `/m` tests over the stored payload. The RRULE one now decides
whether `update_calendar_event` / `delete_calendar_event` refuse the call outright
(`isRecurringSeriesResource`), so a forged `RRULE:` line in a `SUMMARY` would either block a
legitimate edit or, read the other way, hide a real rule inside a folded line — which is why
that detector reads whole content lines and takes no `/m` shortcut.

**The rule, then.** `src/caldav-client.ts` splits a payload with `icalContentLines` — CRLF and
LF only — and matches WHOLE lines for `BEGIN:VEVENT` / `END:VEVENT` / `KEY[;:]`. U+2028,
U+2029 and a bare CR are thereby inert characters inside a value, which is what the RFC says
they are. Every structural question goes through one of three helpers (`extractVEventBlocks`,
`hasICalProperty`, `structuralLine`), so a new caller cannot reintroduce the `/m` form by
copying an old line. **Do not "simplify" any of it back to a regex over the blob.**

**Unfolding stays per RFC and stays LATER than the component split.** A continuation line is
one beginning with a space or a tab (`isFoldedContinuation`), and every structural scan skips
them, so a marker at the head of a continuation is text rather than structure. That closes the
fold-injection half of the same class: libical folds at a fixed octet count, so padding a
description to put a chosen fold in a chosen place is deterministic to construct — measured,
one VEVENT in and two rows out, the phantom carrying the real event's `SUMMARY` and `DTSTART`
(the author chooses the property order). The mirror image is a `BEGIN:VEVENT` in text *before*
the first real component — inside a `VTIMEZONE`'s `TZNAME`, say — which opens the block early
enough to swallow the zone rule's own `DTSTART` and report the event dated 1970.

**"Every structural scan" includes the one that decides where to INSERT.** `replaceICalProperty`
runs four scans and the insert-position one was the last to be converted, which is the shape
this class keeps taking: it looks for the first sub-component so a new property lands before a
`VALARM` (RFC 5545's `eventprop *alarmc` order), and a trimmed compare read ` BEGIN:phase two
of the agenda` — the head of a folded `DESCRIPTION` — as that sub-component. The new line was
then spliced into the MIDDLE of the description: the description lost its tail and the inserted
property swallowed it, two stored records damaged in one write with nothing reported. It is
reachable from every `update_calendar_event` that ADDS a property the event does not already
have, which is the ordinary case for setting a location or participants. When a scan in this
file compares a line, it uses `structuralLine`; a `.trim()` there is a bug even when the
function around it already looks converted.

**Test with the code points explicitly.** The fold tests do NOT cover this: a `/m`-anchored
parser and a line-based one both handle a folded terminator correctly, so nothing in the
folding tests would notice a regression back to `/m`. `src/caldav-client.test.ts` carries a
suite that writes U+2028/U+2029 as `\uXXXX` escapes for exactly that reason — a literal one is
invisible in an editor and indistinguishable from a space in review.

**Can such a value actually reach us? For U+2028/U+2029, yes.** Cyrus parses and re-serialises
every stored resource through libical (`icalcomponent_as_ical_string`), so whatever libical
preserves is what a read gets back. Reading libical's TEXT quoting: on output it escapes only
the LINE FEED and leaves every other byte alone, and on input it decodes only the five
single-letter escapes n, t, r, b and f. U+2028 and U+2029 are ordinary UTF-8 bytes to it and
pass through both directions untouched. So an invitation, a shared calendar or any client's
PUT can put one in a SUMMARY and this client will read it verbatim. A **bare CR** is the
exception: libical DISCARDS one on output, so that third variant is unlikely to survive a round trip through Fastmail. It is still parsed as text here,
because a defence that rests on another product's serialiser quietly dropping a byte is not a
defence, and the JMAP/JSCalendar path is not the same code as the CalDAV one.

**What was NOT done, and why.** U+2028/U+2029 are not stripped from calendar data before
parsing. The parser no longer depends on it; stripping would silently alter a legitimate TEXT
value, which is data loss of the kind the omit-empty convention exists to avoid; and the same
code points already reach a caller through every email subject and body, so scrubbing calendar
text alone would be an inconsistent half-measure rather than a defence. If output scrubbing is
ever wanted it belongs at the serialisation seam, applied to everything, not here.

## The expanded-recurrence payload shape is TOLD, never sniffed

`parseCalendarObjects` takes an `expanded` flag from the caller that decided to send
`expand`, because the payload cannot be read backwards to recover that decision. The
plausible-looking content sniff — "every block carries a RECURRENCE-ID, therefore the server
expanded this" — is false against Cyrus. `expand_cb` (`imap/http_caldav.c`) sets a
RECURRENCE-ID only on instances *after* the series' first; the first instance is emitted with
its RRULE stripped and no RECURRENCE-ID at all. So any window containing a series' original
DTSTART returns `[first-instance, occurrence, occurrence, …]`, the sniff identifies block 0 as
a master, and every sibling is discarded — measured live at 5 occurrences reported as 1, and
102 blocks reduced to 75. The loss was invisible because `total` is counted after it.

The same platform fact leaves one residue that cannot be closed from the payload: a lone
expanded block with neither marker is a one-off event AND a series whose only in-window
instance is its first. Cyrus emits both identically. `isRecurring` is therefore set from the
block list — more than one block, or any RECURRENCE-ID, proves a series — and left off in the
ambiguous case rather than guessed at in either direction, since claiming it would mark real
one-off events as repeating. The tool description and README name `get_calendar_event` as the
one-call way to settle it, because that path fetches without a window and returns the master.

## Local-time formatting and the U+202F trap

Date rendering for humans (`toLocalIso` and `formatReplyDate` in
`src/email-formatter.ts`) has two traps:

- **Render in an explicit timezone.** Use `timezone || defaultTimezone || host`. A bare
  `toLocaleString()` with no explicit `timeZone` silently emits GMT+0, not local time.
- **Normalise U+202F.** Node 20+ ICU inserts a narrow no-break space (U+202F), not an
  ASCII space, before `AM` / `PM`. So `Intl` returns e.g. `1:29` + U+202F + `PM`. The
  attribution strings (and their exact-match tests, which assert U+202F and U+00A0 are
  absent) use a plain ASCII space, so these functions normalise U+202F and U+00A0 to
  ASCII space before returning. Do not "simplify" the normalisation away: it is
  invisible in a diff and breaks exact-match / byte-compare verification.

## Re-sending sanitised content: the two-pass quote sanitiser

`reply_email` re-sends the original message's HTML as a quote under the user's own
`From`, so the quoted HTML is active content we originate, not passive display. The
`sanitizeQuoteHtml` choices in `src/inline-images.ts` are load-bearing security, not
incidental config:

- **No global `'*'` attribute key.** Drops `style=` / `class=` / `on*=` on every tag.
  `style` is the classic CSS-exfil / mXSS vector.
- **`allowedSchemes: ['http','https','mailto']` + `allowProtocolRelative: false`.** The
  library defaults add `ftp` and allow `//host`.
- **`exclusiveFilter` drops any `<img>` whose `src` did not survive sanitising**, so a
  quote never renders a broken-image placeholder. Intentionally narrow (drop
  unusable-src images); it is not a tracker-pixel arms race — mainstream clients do not
  strip trackers from quotes either, and a partial filter just makes the quote less
  faithful.

### Why two passes, and why in that order

Embedded (`cid:`) images used to die in the sanitiser: `cid` is not an allowed scheme, so
the `src` was stripped to empty and `exclusiveFilter` removed the element. Carrying those
images into the quote (#13) means the same html has to be sanitised twice, in two modes:

- **`collect`** reports which references the html makes and rewrites nothing. Its output is
  byte-for-byte what the sanitiser alone would emit — `cid` is still not admitted — so it is
  exactly the string a quotability check should read.
- **`map`** rewrites each reference that resolved to a part into the Content-ID this draft
  attaches for it, with `allowedSchemesByTag` admitting `cid` **on `<img>` and nowhere
  else** (a per-tag list replaces the global one for that tag, so no other attribute can
  ever carry a `cid:` URL).

The order is the whole point: minting an identifier commits the call to attaching a part,
and a part nothing references is a stray file on the finished message. So pass one decides
whether an html quote ships at all, and pass two runs only on a branch that really ships
one. A text-only reply mints nothing.

`map` mode is **default-deny**: an `<img>` survives only when the transform affirmatively
emits a `src` — a mapped identifier, or a value whose normalised scheme is http/https.
Everything else has its `src` deleted and the element removed. That matters because the
classifier reimplements the sanitiser's own URL normalisation (see the launder note under
*Dependency / build gotchas*): if the two ever drift, an unrecognised spelling lands in the
"not affirmatively emitted" bucket and is dropped, rather than sliding through as the
scheme-less URL the sanitiser would have passed. A visible consequence, deliberately
accepted: a relative or scheme-less `<img src>` no longer survives a quote. Such a
reference is already broken in mail — there is no base URL to resolve it against — and
admitting it would mean trusting the classifier's *negative* answer, which is the thing
this design refuses to do. The drop is counted, never silent: `droppedUnsupportedImages` is
its own counter with its own sentence, kept apart from a `data:` image (content this server
declines to re-encode) and from an unmatched embedded-image reference (which named a part
that was not there). It is a **map-mode-only** count, because the collecting pass drops none
of these — it leaves them to the sanitiser, which passes a relative URL through. That
asymmetry is precisely why the count exists: the pass that ships the quote loses an image
the other pass would have kept. An `<img>` with no `src` at all is not counted; there was no
image to lose.

Quotability follows from the same pass. An original whose only content is embedded images
sanitises in `collect` mode to something visually empty (`<div></div>`), so it is judged
quotable by whether at least one of its references would really embed — resolved to exactly
one part, that part declaring itself an image and carrying a blob. Testing mere *resolution*
would open an attribution over a quote showing nothing; testing the sanitised string would
call an image-only message unquotable, which is what it used to be.

Accepted threat floor (documented in README): `sanitize-html` is a string-to-string
sanitiser (roughly the bar Gmail / Apple Mail emit) and does not fully eliminate exotic
mutation-XSS. Stripping script / `on*` / `style` / unscoped wrappers plus pinned schemes
is the deliberate safety floor, not an oversight. It governs the markup this server writes
and says nothing about the bytes of a part carried by reference — see
`docs/security-model.md` for what carrying an image discloses.

## Process lifecycle: the exit on stdin EOF is implicit

The server has no shutdown handler, and that is deliberate. The SDK's
`StdioServerTransport` attaches only `data` / `error` listeners to `process.stdin` and
never registers an EOF handler, so nothing calls `transport.close()` and the server's
`onclose` never runs. The process ends anyway: stdin reaching EOF releases the last
handle holding the event loop open, and node exits with code 0. Measured on Windows,
that takes tens of milliseconds.

The post-request case was **measured once and is an accepted residual, not
regression-covered**: after a tool call had opened an HTTPS connection to Fastmail,
EOF still exited in ~35ms, so an idle keep-alive socket from `fetch` did not hold the
loop open. The lifecycle tests deliberately do not cover this - reaching a tool call
means credentials and a live network round trip, which `npm test` must not require -
so a future undici or node change here would not be caught automatically.

Two consequences worth knowing before changing startup code:

- **Anything that keeps the event loop alive turns a clean exit into a hang.** A bare
  `setInterval`, a retained socket, a file watcher, an unresolved handle - any of these
  and an embedding caller that closes stdin waits forever, most damagingly on Windows
  where the orphan outlives its parent silently. Keep long-lived handles out of the
  startup path, or `unref()` them.
- **An explicit `process.exit()` on EOF was considered and rejected.** It would trade
  the graceful drain (an in-flight tool call still gets to write its response) for a
  truncated one, and it would mask a leaked handle rather than expose it.

`src/server-lifecycle.test.ts` is the guard: it spawns the built `dist/index.js`, closes
stdin, and fails if the process does not exit on its own. The consumer-facing side of
this (spawn directly with node, never via `cmd /c` on Windows, expected handshake
latency) is in the README's embedding section.

## Dependency / build gotchas

- **`html-to-text` v10 ships no type declarations.** Types come from a separate
  `@types/html-to-text@^9` devDependency (its `index.d.ts` exports `convert`, matching
  v10's runtime API). The import is `import { convert } from 'html-to-text'` (v9+ removed
  the default export).
- **`sanitize-html` uses `export =`.** Import it as a default: `import sanitizeHtml from
  'sanitize-html'` (works under the repo's NodeNext / esModuleInterop), not
  `import * as`.
- **The URL normalisation is reimplemented, not imported.** `launderUrlValue` in
  `src/inline-images.ts` mirrors what `sanitize-html` does to a URL attribute before its
  scheme gate runs: strip every character of code 0x20 and below (browsers ignore those
  inside URLs in more places than is comfortable), then clobber embedded `<!--…-->`
  comments. That logic lives in `launder`, which reaches this project only as a
  **transitive dependency** of `sanitize-html` — `package.json` declares `sanitize-html`
  alone — so importing it directly would take a hard dependency on another package's
  dependency tree. The mirror can therefore drift when either package updates. The
  tripwire is the obfuscated-spelling property tests, which assert that `c id:x`,
  `cid&#9;:x` and `c<!--z-->id:x` classify the same way a plain `cid:x` does and that no
  such spelling survives sanitisation unmapped. Do not weaken them, and do not "simplify"
  the character class to `\s` — that excludes NUL and the other C0 controls, which are
  exactly what an obfuscated spelling would use.
