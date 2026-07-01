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
- `requireNonEmpty` / `validateClearFields` — the loud-reject + `clearFields` machinery
  shared by `update_calendar_event` and `edit_draft`.

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

## Dependency / build gotchas

- **`html-to-text` v10 ships no type declarations.** Types come from a separate
  `@types/html-to-text@^9` devDependency (its `index.d.ts` exports `convert`, matching
  v10's runtime API). The import is `import { convert } from 'html-to-text'` (v9+ removed
  the default export).
- **`sanitize-html` uses `export =`.** Import it as a default: `import sanitizeHtml from
  'sanitize-html'` (works under the repo's NodeNext / esModuleInterop), not
  `import * as`.
