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
- `coerceRecipients` — fans `coerceStringArray` over `to` / `cc` / `bcc` / `replyTo` so
  no recipient field can reach `.map(parseAddress)` as a bare string (the original
  `cc:""` / `bcc:""` crash class).
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
- `requireNonEmpty` / `validateClearFields` — the loud-reject + `clearFields` machinery
  shared by `update_calendar_event` and `edit_draft`.

Leniency has a limit: a value is coerced when the intent is unambiguous, and rejected when
guessing would change the message. `textBody` / `htmlBody` are the reject side — a
non-string body, an entirely HTML-escaped `htmlBody`, and a CDATA-wrapped body are refused
by `assertBodyInputs` (`src/body-format.ts`) rather than repaired, because unescaping or
unwrapping would guess at what the caller meant to send. See `docs/email-bodies.md`.

### Verifying coercion

The normal MCP tool harness validates the declared `inputSchema` before the call
reaches the handler, so it will reject the malformed inputs these coercions are meant to
accept. To verify coercion you must drive a raw JSON-RPC request against the built
server (`dist/index.js`) with `FASTMAIL_API_TOKEN` set, bypassing the schema-validating
harness. (See the `verify-lenient-client-coercion` note in project memory.)

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

## Mailbox resolution is uniform (id / role / name, exact)

Every mailbox-taking parameter resolves through one exact matcher (`findMailboxExact` in
`src/jmap-client.ts`: exact id, then role, then name, case-insensitive, no substring). This
now spans the label tools' `mailboxIds` arrays too (`add_labels` / `remove_labels` /
`bulk_add_labels` / `bulk_remove_labels`), closing the last asymmetry from the #12
single-mailbox consolidation (#50): a caller can label by the same id/role/name that works on
`move_email`, not opaque ids only. The array resolver is **all-or-nothing** — if any entry
can't be resolved it names every unresolved value in one error (so all typos are fixable in a
single retry) and applies no labels, rather than half-applying a mutation the caller must
reconcile. `resolveMailbox` is the throwing single-input wrapper over the same core.

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
  identity), and a server-side `notFound` SetError on a mutation. These throw the tagged
  `InvalidInputError` (`src/coerce.ts`), which the top-level CallTool catch maps to
  `InvalidParams` (after `redactBearerTokens`).
- **Operational / server → `InternalError`.** A failure the caller cannot fix by changing
  input: zero sending identities, a missing system mailbox (Drafts/Sent/Trash), a
  transport error, a `notCreated`/non-`notFound` set-error (server refusal), or a
  post-condition like "returned no ID." These stay a plain `Error`.

This rule is **tool-family-agnostic.** Because the calendar tools share the same
`requireNonEmpty` / `validateClearFields` helpers from `src/coerce.ts`, their input
rejects (`create_calendar_event` / `update_calendar_event`) are `InvalidParams` too — the
classification is a property of the shared helpers, not of email specifically.

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

The JMAP set-error reason itself is surfaced (not just the code): every throwing
`Email/set` failure routes its `SetError` through `describeSetError` in
`src/jmap-client.ts` so the server's `type`/`description` reaches the caller, and bulk
mutators additionally report success/fail counts and the caller's failing ids grouped by
reason. The helper concatenates only server-authored text — we add no message body of our
own.

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

## Re-sending sanitised content: `sanitizeForQuote`

`reply_email` re-sends the original message's HTML as a quote under the user's own
`From`, so the quoted HTML is active content we originate, not passive display. The
`sanitizeForQuote` choices in `src/reply-quote.ts` are load-bearing security, not
incidental config:

- **No global `'*'` attribute key.** Drops `style=` / `class=` / `on*=` on every tag.
  `style` is the classic CSS-exfil / mXSS vector.
- **`allowedSchemes: ['http','https','mailto']` + `allowProtocolRelative: false`.** The
  library defaults add `ftp` and allow `//host`.
- **`exclusiveFilter` drops any `<img>` whose `src` did not survive sanitising.** A
  `cid:` / `data:` image gets scheme-stripped to an empty `src` and would otherwise
  render as a broken-image placeholder; inline `cid:` logos and signatures are very
  common in replies. This filter is intentionally narrow (drop unusable-src images); it
  is not a tracker-pixel arms race (mainstream clients do not strip trackers from quotes
  either, and a partial filter just makes the quote less faithful).

Accepted threat floor (documented in README): `sanitize-html` is a string-to-string
sanitiser (roughly the bar Gmail / Apple Mail emit) and does not fully eliminate exotic
mutation-XSS. Stripping script / `on*` / `style` / unscoped wrappers plus pinned schemes
is the deliberate safety floor, not an oversight.

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
