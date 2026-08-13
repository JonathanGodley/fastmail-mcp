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
  call's arguments OR by editing the object (e.g. the draft) the call operates on. This
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

**A flag that only a simplified response can honour is REJECTED with `raw`, not ignored.**
`stripQuoted` (#73) rewrites the simplified `bodyText`; there is no honest way to apply it
to a pure-JMAP payload, and silently dropping it would leave a caller believing a response
was stripped when it was not. So `assertStripQuotedNotRaw` (`src/quote-strip.ts`) rejects
the pair from both `get_email` and `get_thread`, naming the two ways out. Same reasoning as
the unknown-parameter guard: a parameter that quietly does nothing produces confident wrong
answers. Contrast `includeBodies` (#74), which is a *fetch*-level knob — it changes the
`Email/get` property set, so `raw` faithfully returns the richer JMAP object and the two
compose fine.

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
