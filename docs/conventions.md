# Cross-cutting conventions and gotchas

Developer-facing conventions and non-obvious traps that span multiple tools or are
properties of the toolchain. Per-tool behaviour rationale lives in the relevant GitHub
issue; this file is the shared stuff a developer reading the code needs to know.

## Lenient input coercion

MCP clients (especially LLMs) send sloppy parameter shapes: a comma-joined string where
an array is expected, a single bare id, a stringified boolean, `"to": ""`. The server
coerces rather than rejects, so a reasonable-looking call does not crash. This spans
most tools, so the helpers are centralised in `src/coerce.ts`:

- `coerceStringArray` — array / comma-string / single value to `string[]` (or
  `undefined`). `""` coerces to `[]`.
- `coerceStringArrayStrict` — the same coercion for a parameter that must **fail closed**:
  a value that is present but uncoercible (a number, an object, a boolean) is an
  `InvalidInputError` naming the parameter, instead of the plain coercer's `undefined`.
  `null`/`undefined` still read as absent. The dividing line is what a dropped value
  costs: on a *content* field it means "leave this unchanged", which is harmless; on a
  field that **narrows what the call touches** it silently widens the operation the
  argument was passed to restrict, and the caller reads the wider answer as complete.
  `search_emails`' `requiredMailboxes` / `excludeMailboxes` are the first users.
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
name fallback and no destination parameter. These are fixed, membership-replacing defaults, and a
folder can be created with any name a caller likes, "Trash" or "archive" included. Resolving the
name would make the destination of a destructive default something text the model merely read could
aim; a role is assigned by the server and cannot be minted that way. A caller who wants a different
destination has one: `move_email`, where naming it is the whole point.

`move_email`, `bulk_move` and `archive_email` all set `mailboxIds` **whole-value** in a single
`Email/set` (RFC 8620 §5.3) rather than reading the current membership and patching each id away.
It states the promised contract directly ("replaces all mailbox membership") and has no read/write
window in which a newly-added mailbox survives the move. None of the three writes a keyword: a move
changes where a message is filed and nothing about its read or flagged state, so archiving does not
mark a message read. Marking read is `mark_email_read`. This is a deliberate divergence from
upstream PR `MadLlama25/fastmail-mcp#67`, whose `archive_email` writes `$seen` as part of the move:
folding two effects into one verb means a caller who wanted only the filing cannot get it, while a
caller who wanted both can still make the second call.

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
explicit scope has two edit sites. Naming folders to look **in** (`mailbox`,
`requiredMailboxes`) is an explicit scope and turns the default off; naming folders to
exclude is not, because narrowing by exclusion says nothing about wanting Trash/Spam back.

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

A **missing Archive mailbox is the one that switches sides**: `archive_email` on an account
with no archive-role mailbox throws `InvalidInputError` (`InvalidParams`), not the plain
`Error` its Trash counterpart throws. The split is recoverability, and it really does differ
here. "Archive" is a filing convention, so the caller substitutes any folder they like via
`move_email` and gets the same outcome, which is a re-formed call and so the definition of
caller-fixable. A missing Drafts/Sent/Trash has no substitute: nothing else *is* the trash,
and no argument to any tool produces a delete without it. The error message therefore names
`move_email`, because the classification is only honest if the route it implies exists.

This rule is **tool-family-agnostic.** Because the calendar tools share the same
`requireNonEmpty` / `validateClearFields` helpers from `src/coerce.ts`, their input
rejects (`create_calendar_event` / `update_calendar_event`) are `InvalidParams` too — the
classification is a property of the shared helpers, not of email specifically. The
calendar rejects raised in `src/caldav-client.ts` itself follow the same rule and throw
`InvalidInputError` directly: the start/end frame and ordering checks below, the date
format and control-character rejects, the participant-address rejects, a `calendarId` or
`eventId` that resolves to nothing, and the recurring-exception confirmation prompt (the
one asking for `confirmRecurring: true`). `src/contacts-calendar.ts` follows it too — its
input rejects (a contact with neither name nor address, an update naming no field, an
empty entry array, an ambiguous entry edit, an update aimed at a contact group) and its
not-found rejects are `InvalidInputError`, and its create/update/delete set-errors route
through the same `throwSingleSetError` classifier the mail writes use. That classifier puts
a `forbidden` on the operational side, so a contacts token issued read-only surfaces as an
`InternalError` whose message names the read-only scope — the refusal is the only place a
caller can learn that, since the session reports an identical capability either way.

Two nearby throws stay a plain `Error` **on purpose**, and both look like exceptions to the
rule until you see what they carry. The ORGANIZER address check validates the operator's
configured CalDAV username through the participant validator — a configuration fault no
argument can fix, so it is re-thrown in the plain class with its own wording. And the
recurrence-expansion guard's `skip-pruning:` throws are internal control-flow signals for a
`catch` that re-raises only `InvalidInputError`; tagging them would let a resource guard
escape to the caller dressed as a rejection of their input.
That last one is the clearest case in the file — the caller resolves it by re-sending the
same call with one more argument, which is the definition of caller-fixable, so surfacing
it as `InternalError` was telling them a bare retry might work when it never would. Its
`catch` re-throws on `instanceof InvalidInputError` rather than on a message substring, so
the routing no longer depends on the wording of the prompt.

`assertICalTextLimits` (`src/ical-limits.ts`) follows the same rule from the other end of
the call chain: it runs in the `create_calendar_event` / `update_calendar_event` handlers,
before anything measures or serializes the values, and throws `InvalidInputError` naming
the field, its size and the limit. The classification matters more than usual there. The
whole point of the bound is that the work it refuses is expensive (see **Bounding a
quadratic serializer** below), so an `InternalError` would not merely be inaccurate, it
would invite the bare retry that repeats the cost.

**One deliberate carve-out:** `download_attachment` returns `InternalError` for a bad
`emailId`/`attachmentId`. Its local catch collapses non-path errors to a generic message
on purpose, so it does not leak attachment metadata (see `docs/security-model.md`). So a
bad id is `InvalidParams` on `get_email`/`get_thread` but `InternalError` on
`download_attachment` — an accepted, documented asymmetry.

**Inside that carve-out, one further line: INPUT-FORM errors vs EXISTENCE errors.** The
generic message exists to avoid confirming what a mailbox contains. That reasoning covers
"this reference matched nothing" and nothing else — it does not cover "this reference is
not a shape this tool accepts", which is answerable from the caller's own string. So
`attachmentId` splits:

- **Input-form failures** — a bare `cid:` with no value, a number with junk in it (`3a`,
  `-1`, `1.5`) — throw `InvalidInputError` and reach the caller intact, naming what to
  pass instead. They quote back only the caller's own input (through `describePart`, so
  a hostile value cannot overrun the sentence) and never enumerate blobIds or names.
- **Two more throw `InvalidInputError` on a weaker justification, stated honestly:** a
  `cid:` value matching more than one part, and a resolved part carrying no blobId.
  Neither is answerable from the caller's string alone — both depend on message content,
  so each is a narrow existence oracle (it confirms that a reference resolves, and in the
  first case that at least two parts share that Content-ID). They are classified as input
  errors anyway because the alternative is worse: a generic failure on a reference that
  DOES resolve leaves the caller retrying a request that can never succeed, with no way
  to learn that a different reference form is required. Neither message reveals a blobId,
  a filename, or a count.
- **Existence failures** — a well-formed reference that simply matches nothing — stay a
  plain `Error` and collapse to the generic "Attachment download failed" message.

Mechanically, `download_attachment`'s local catch **re-throws** `InvalidInputError`
rather than mapping it there, so the top-level branch applies the `InvalidParams` mapping
*with* `redactBearerTokens`. A verbatim local re-throw would skip redaction the way the
`PathAccessError` route does, and unlike path messages these echo caller input.

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
  `list_emails` and `search_emails` inherit it together; `getRecentEmails` builds its own
  batch and takes the same parameter. The filters (including the default Trash/Spam
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
- **The refusal is honest about the opt-in.** When `FASTMAIL_ATTACH_DIR` is unset there is
  no way to supply the missing file at all, so the wording drops the "add it to
  attachments" repair rather than pointing at a parameter that cannot work. The flag is
  threaded into the refusal builders for exactly that reason.

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
  clocks read backwards. Resolving that ordering needs a timezone database this server
  does not carry, so the check declines to guess rather than reject valid travel events.

The frame check is deliberately *not* applied when the caller touches neither `start` nor
`end`: it exists to stop us writing a broken pair, not to hold a title edit hostage to an
inconsistency some other client left in the event.

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
