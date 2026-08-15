# Path-confinement security model

Two tools touch the local filesystem: attachment download (writes a file) and
send-with-attachment (reads a file). Both are constrained to a configured directory and
can never be told to escape it. This spans both features, so the model lives here; the
per-feature rationale is in issues #5 (download) and #1 (attachments).

## Confinement is always on, never bypassable

Path confinement is lexical plus symlink/realpath-safe and is permanently on. There is
no disable flag. The write side is `safeWritePath` (`src/jmap-client.ts:1352`): it
lexically pre-checks, realpaths the allowed directory, walks up to the longest existing
ancestor, verifies that ancestor lives under the canonical allowed root, and refuses to
overwrite an existing symlink at the target.

Scope is widened by configuration, not by a bypass. You set the allowed directory as
broadly as you like (a configurable `FASTMAIL_DOWNLOAD_DIR`, even a drive root); per-call
absolute paths are honoured within that root and stay symlink-safe. "I want anywhere"
is an explicit config choice, not a `FASTMAIL_ALLOW_ANY_PATH` flag. This is the same
conclusion upstream reached when they rejected a bypass flag in favour of a configurable
dir.

## Reads are an exfiltration vector

Sending a file as an attachment reads a local file and emails it out, so it is treated
as opt-in capability, not a default:

- Attach-from-path is disabled until the user sets `FASTMAIL_ATTACH_DIR`; until then the
  tool returns a clear self-documenting error rather than reading anything. A file-read-
  and-email capability is textbook opt-in; an injected agent could otherwise attach
  `~/.ssh/id_rsa` and send it.
- A missing `FASTMAIL_ATTACH_DIR` is an error, not an auto-create. Auto-creating an
  exfiltration root is a footgun.
- No coupling to the download directory. Reusing the download dir would let a just-
  downloaded file be auto-emailed straight back out (the exact round-trip this closes).
  The attach dir is resolved independently, via the same env-alias key set as the
  download dir.

### Filenames derived from message content

`download_attachment` builds a download URL whose `{name}` slot declares a filename, and
that name comes from the part's **sender-supplied** `name` — attacker-controlled on any
received message, and a value a receiving client may reuse as a save name. It goes
through `sanitizeDownloadFilename` (`src/inline-images.ts`) first: control and format
characters are stripped (a bidi override such as U+202E can otherwise make `exe.txt`
render as `txt.exe`), path separators and the Windows drive/ADS colon become underscores,
leading dots are dropped along with any whitespace shielding them (so ` .hidden` cannot
smuggle a dotfile name past the strip), the length is capped by code point so a surrogate
pair is never split, and a Windows device name gains an underscore on its stem (`CON.png`
→ `CON_.png`) — including one padded with the trailing spaces and dots Win32 strips before
it matches device names, so `CON .png` is defused too. A name that sanitizes to nothing
becomes `attachment`, so the value is never an empty path segment.

This is deliberately stricter than `forward_email`'s `sanitizeEmlFilename`, which applies
a similar character treatment but lets device names through: that helper always appends
`.eml`, which neutralizes them, and its output is a name a *remote* recipient's client
saves. This one is a name a local client may write, so the inherited posture does not
transfer. That helper also still trims after dropping leading dots, so whitespace can
shield one there; its unconditional `.eml` suffix means the result is a named file either
way, and changing it would alter what forwarded mail declares on the wire. The two are
kept as separate functions for exactly these divergences.

The guarantee stops at the name. The **path** an attachment is written to is never
derived from message content — `download_attachment` writes only where the caller's
`path` says, under the confinement rules above, and there is no "save with the sender's
filename" default. That is what keeps a hostile `name` from being a path decision at all.

## The read-shaped `safeReadPath` (built, issue #1)

The attachment-send feature reads a local file and emails it out, so it needs a
read-shaped guard distinct from the write-shaped `safeWritePath` (which `mkdir -p`s the
allowed root, walks *missing* path segments, and checks the target only for an overwrite
symlink — none of which is right for a read). `safeReadPath` (`src/jmap-client.ts`) is
handle-based:

- **Hard opt-in gate first.** If `FASTMAIL_ATTACH_DIR` is unset, it throws the
  self-documenting opt-in error **before any filesystem syscall** — nothing is read.
- **Reject Windows escape shapes** on the raw input (applied on every platform): device
  namespaces (`\\?\`, `\\.\`), UNC roots (`\\server\share`), drive-relative `C:foo` (no
  separator, resolves against the drive's own CWD), an NTFS alternate-data-stream `:`
  past the drive letter, and an 8.3 short-name `~` segment (which can alias a long name
  past the containment compare).
- **Lexical containment** against the resolved attach root. A bare filename or relative
  path resolves *inside* the root in one step; an absolute path must already be within it.
  The containment compare is **case-insensitive on Win32** (NTFS folds case, so a
  case-sensitive `startsWith` is bypassable); the write side keeps its byte-exact compare
  so download behaviour is unchanged. This case-fold is a *parameter* of the
  containment helper shared by both guards.
- **Open once, then validate the open file.** `open(path, 'r')`, `fstat` the handle and
  require a **regular file**, then `realpath` the **full target** (not merely an ancestor)
  and re-verify it is contained under `realpath(attachDir)`. The caller reads from the
  returned handle, so the bytes uploaded are the bytes of the file that was validated.
- A missing attach **root** is reported as a distinct config error (not a raw `realpath`
  ENOENT, and distinct from the opt-in gate).

The resolved attach root is **disclosed in the tool schema** (parallel to how the download
directory is interpolated into `download_attachment`'s `path` description). It is
operator-chosen and low-sensitivity; within-root path specifics in boundary errors are
therefore actionable rather than a sensitive oracle (caveat: a drive-root config makes
them a broad probe — acceptable as an explicit operator choice).

### Accepted residual risks (not claimed closed)

- **Same-inode swap race.** The `realpath` re-verification runs after `open`, so a
  component swapped between open and realpath is a narrowed-but-nonzero TOCTOU window.
- **Win32 has no fd→path binding.** There is no syscall to canonicalize the *open handle*
  itself on Windows, so a symlink/junction race there is residual.
- **Hardlinks inside the root** pointing at outside content defeat any path-based guard.

These are the honest limits of a path guard; the opt-in gate and confinement are the
primary defense, not a claim that exfiltration is impossible once enabled.

### edit_draft attachment model

`edit_draft` carries the existing (non-inline) attachments across the immutable-email
recreate and then applies the requested change: `attachments` **appends**;
`removeAttachments` drops carried parts by `blobId` (or a unique non-null `name`), rejecting
a ref that matches nothing or a name matching more than one; `clearFields:['attachments']`
removes all. Passing `attachments` together with `clearFields:['attachments']` is a rejected
conflict. An attachment-only edit stays body-invariant (it must not inject or strip a body).

**Accepted residual — orphaned blobs on a late reject.** The handlers upload new
attachment blobs *before* the draft create/recreate runs, so a rejection raised by a later
guard (e.g. `edit_draft`'s inline-cid reject, a non-text/html body part, an unresolvable
`removeAttachments` ref, or the no-body-result guard) leaves the just-uploaded blobs
unreferenced. `uploadAttachments` orphans zero blobs *within its own batch* (a two-pass
design), but that guarantee ends at its return; the upload-then-reject ordering reopens a
window. Accepted because Fastmail garbage-collects unreferenced blobs (the same GC the code
already relies on for a mid-batch upload failure) — no unbounded growth, no data exposure.

### forward_email attachment postures (#30)

- **Carried attachment names are relayed faithfully.** The original's attachment names are
  attacker-chosen, and the inline forward re-sends them verbatim under the user's From.
  Renaming a user's files in transit would be surprising, every mainstream client relays
  verbatim, and the consumer of the name is the *receiving* client's save dialog — which is
  the layer that sanitizes save names. Same posture as `edit_draft`'s pre-existing
  attachment carry (which also relays stored names verbatim across the recreate).
- **The tool-GENERATED `.eml` filename is sanitized** — the one name this server *creates*
  (`asAttachment`, derived from the original's attacker-controlled subject) is held to a
  higher bar than names it merely relays: control chars (`\p{Cc}`, C0+C1) and Unicode
  format/bidi controls (`\p{Cf}`, e.g. U+202E extension spoofing) stripped, path
  separators and the Windows drive/ADS colon replaced, leading dots stripped, length
  capped; a blank result falls back to `forwarded-message.eml`. Windows reserved device
  names (`CON`, `NUL`, …) deliberately survive as e.g. `CON.eml` — a save-time nuisance the
  receiving client handles, consistent with the relay posture above.
- **`asAttachment` raw-blob exposure (accepted, disclosed).** The attached `.eml` is the
  stored RFC 5322 message, byte-identical. Probed live 2026-07-05: a **Sent-copy blob
  retains the `Bcc` header** (forwarding your own sent mail as `.eml` discloses its Bcc
  recipients), and a received-message blob exposes the full transport record — the Received
  hop chain (8 hops on the probe, with internal host names/IPs), spam-filter verdict
  headers, ARC and Authentication-Results sets, and delivery-routing headers. None of this
  is visible in an inline forward, which reproduces only the body under a From/To/Cc/
  Subject/Date block. The tool's `asAttachment` description discloses this; it is the
  caller's deliberate trade for losslessness.
- **Outgoing size is unbounded by this server on the carry/.eml paths (accepted).**
  `MAX_ATTACHMENT_BYTES` caps only *local uploads* (files this server reads off disk);
  carried originals and the `.eml` are blobId **re-references** never read client-side, so
  no client-side cap applies — Fastmail's own message-size limit governs, and an oversized
  send fails loudly server-side. Parity with `edit_draft`'s carry, which re-references the
  same way.

## `originalEmailId` is an in-account read-and-embed primitive (accepted residual)

`reply_email`, `forward_email`, and `edit_draft`'s keep path all take an `originalEmailId`
and fetch that message's body and embed it (sanitized via `sanitizeForQuote`) into
a draft the caller may then send. Stated plainly: this lets a caller move one message's
content into outgoing mail addressed to arbitrary recipients under the user's own `From`. A
prompt-injected agent could use it to exfiltrate the content of any message in the account by
quoting it into a reply it sends to an attacker-chosen address.

The id is **trusted and unscoped within the connected account** — it may name *any* message,
deliberately, so a caller can correct a draft built against the wrong original. It is **never
re-resolved from the draft's `In-Reply-To`** (an attacker-controllable header), so there is no
confused-deputy / quote-spoofing surface from that direction, and there is **no cross-account
reach** (the fetch is scoped to `session.accountId`).

This introduces **no new capability class** versus the already-shipped `reply_email`, which
quotes any `originalEmailId` the same way; `edit_draft`'s keep path just reuses it. The
embedded html is run through `sanitizeForQuote` (script/style/handlers/unscoped attributes
stripped, schemes pinned) — a safety floor for re-sending under the user's `From`, not a
privacy control. Documented here as an accepted residual: the mitigation for misuse is the
same opt-in/authorization posture that governs sending mail at all, not a restriction on which
in-account message may be quoted.

Transmission also **writes two keyword flags** (`$answered`+`$seen`, or
`$forwarded`+`$seen`) after a send succeeds (#52/#54, #30, #60). The compose surface is
draft-first, so this write lives in `send_draft`: the target is resolved from the draft's
recorded provenance headers (see `docs/conventions.md` "Draft provenance"), which are
attacker-influenceable values on a foreign draft — but neither path marks an arbitrary
id. The exact-instance pointer (`X-Fastmail-MCP-Source-Id`) is honoured only after
validating that the named instance still carries the Message-ID the draft's kind header
names (a mismatched or dangling pointer falls back), and the Message-ID fallback only
ever marks the unique message that *owns* that Message-ID (no match, or more than one,
marks nothing). Either way it adds no capability class: two boolean keyword sets, no
move/delete/body write, scoped to `session.accountId`, and dominated by
`mark_email_read`, which already grants a standalone `$seen` write to any id. The write
is best-effort (a failure is swallowed so it can't mask the already-sent mail).
Accepted on the same footing as the read-and-embed primitive.

### `X-Fastmail-MCP-Source-Id` is transmitted to recipients (accepted, deliberate)

The exact-instance header that `reply_email`/`forward_email` stamp on a draft is **not
stripped at send** — EmailSubmission transmits stored headers verbatim (probed live
2026-08-14; see `docs/conventions.md` "Draft provenance" for the probe facts), so the
recipient's copy carries it. This was considered and consciously declined rather than
overlooked:

- **What it discloses:** an opaque, account-scoped JMAP id. It is meaningless outside
  the sending account's own session — it names no host, no folder, no address, and
  cannot be dereferenced by anyone but the account holder. The `In-Reply-To` header on
  the very same message already discloses strictly more (a globally meaningful
  Message-ID including the originating domain).
- **Precedent:** Fastmail's own mobile app transmits its private `X-PersonalityId`
  (an internal identity id) on sent replies — same class of value, shipped by the
  platform vendor.
- **Why not strip:** removing the header at send would mean recreating the message
  before submission (JMAP emails are immutable), turning every send into a
  destroy+recreate with its own failure modes, solely to withhold a value with no
  disclosure weight. The cure was strictly worse than the disease.

## Mailbox resolution + default Trash/Spam exclusion (accepted residuals)

The read surface gained one `mailbox` param (id/role/name) resolved **exactly** across the
read + single-mailbox-write tools, and `search_emails`/`list_emails` hide Trash and Spam by
default with a hidden-count note. Several residuals are accepted here, framed honestly rather
than overclaimed:

- **The default Trash/Spam exclusion is a product/noise default — NO security property is
  claimed.** It is *not* an anti-prompt-injection control: an injected agent simply passes
  `includeSpam:true` (or reads Spam via `get_recent_emails mailbox:"junk"`). Treating it as a
  security boundary would be the same overclaim as "redaction neutralizes the oracle" — so it
  isn't claimed. The `includeTrash`/`includeSpam` descriptions stay plain (no injection caution).
- **The hidden-count note is TRANSPARENCY for a cooperative reader, not an injection control.**
  `get_mailbox_stats mailbox:"junk"` returns Trash/Spam totals directly with zero friction, and
  `get_recent_emails mailbox:"trash"` reads them outright — so a determined/injected agent
  trivially bypasses the note. Its purpose is honesty (disclose what default-scope hid), not a
  boundary. The fail-closed degraded note exists so the published "no note ⇒ nothing in
  Trash/Spam matched" contract can be *trusted by a cooperative caller*, not to stop an attacker.
- **Count-into-Trash/Spam oracle, and the more direct `get_mailbox_stats` total-oracle.** The
  hidden-count discloses how many matches sit in Trash/Spam; `get_mailbox_stats mailbox:"trash"`/
  `"junk"` (single-mailbox = an explicit-scope override) hands those totals directly — the
  lowest-effort volume probe. Accepted; it's the caller's own account.
- **Exact resolution hardens *mis-resolution*, NOT deliberate steering.** Switching every
  read/delete/move target from substring (`findMailboxByRoleOrName`'s name fallback, which could
  mis-hit e.g. a custom "Junk mail rules" mailbox and silently hide real mail) to exact id/role/
  name removes *fuzzy* mis-targeting. It does **not** close *deliberate* steering: an injected
  agent with `move_email`/`bulk_move` access can still aim mail at `"trash"`/`"Archive"` by exact
  name. Move-to-any stays open **by design** (a move-target restriction is tracked as fork #43).
  Name/role resolution also **lowers the steering bar** from "must know a valid opaque id (needs a
  prior `list_mailboxes`)" to "blind one-shot by literal name" — a real, if modest, escalation.
  **The label tools join this class (#50):** `add_labels`/`remove_labels`/`bulk_add_labels`/
  `bulk_remove_labels` now resolve their `mailboxIds` arrays by exact id/role/name too, so an
  injected agent can label a message into e.g. `"trash"` blind-one-shot-by-name, the same modest
  escalation as move. Accepted on the same footing; not a new capability class.
- **The hidden-count note covers ONLY Trash/Spam.** It does **not** disclose a `move_email` to
  `Archive` or a custom folder — that conceals mail with **zero disclosure**. So move concealment
  is not "mitigated by the note" for non-Trash/Spam destinations; this is stated plainly rather
  than implied covered.
- **Resolver error message is an information oracle, reachable account-wide.** A bad `mailbox`/
  `targetMailbox`/`mailboxIds` to *any* swept tool (search, list, stats, move, compose, labels)
  reflects the caller's input and a capped list of mailbox names **reachable by the configured
  token** (the real boundary is the token's reach, not "the user's own account" — a delegated/
  scoped token sees only its slice). `InvalidInputError` messages are run through
  `redactBearerTokens` as defense-in-depth (a token can't actually appear in them), but that is
  **not** what makes the oracle acceptable — recoverability (naming valid mailboxes so a caller
  can retry) is, and it's the caller's own reachable names. Accepted, capped, framed honestly.
- **`get_mailbox_stats` and the label tools reject a real id that is absent from the fetched
  list.** Reading stats off the shared `getMailboxes()` list (and resolving label `mailboxIds`
  against it) means a hidden/role-less mailbox's id now throws `InvalidInputError` rather than
  returning data. Accepted; "resolvable" is defined as "matches some `mailbox.id`/role/name in
  the fetched list."
- **Per-message id-existence is a distinct oracle class.** A not-found id on `get_email`,
  `get_thread`, or `originalEmailId` now returns `InvalidParams` (a crisper signal than the prior
  `InternalError`), so it confirms whether a given *message/thread id* exists. This is a different
  class from the mailbox-resolver oracle above (which reflects reachable mailbox *names*) — it is
  per-message existence, and is likewise bounded by the **token's reach**, not "the user's own
  account." Accepted on the same footing: recoverability is the point, and it is dominated by the
  existing `get_email` read (the same probe already exists), so it adds no capability a caller with
  these tools lacked. The one read path that deliberately does NOT expose this is
  `download_attachment`, whose local catch keeps a generic `InternalError` for a bad
  `emailId`/`attachmentId` so it leaks no attachment metadata.
- **Two-query hidden-count race.** The visible query and the count query run in the same
  `makeRequest` (one atomic snapshot — no race) where possible; the derivation tolerates a missing/
  garbled count by failing closed to the degraded note. A message moving between *any* two reads is
  an inherent, accepted residual, not a temporal guarantee.

## Endpoint allowlist (where the bearer token may be sent)

The other half of the model: every URL that will carry the API token is checked by
`validateFastmailUrl` (`src/url-validation.ts`) — the configured `FASTMAIL_BASE_URL`, and the
`apiUrl`/`downloadUrl`/`uploadUrl` that JMAP session discovery hands back. HTTPS is mandatory
in all modes, including the `FASTMAIL_ALLOW_UNSAFE_BASE_URL` opt-in, because the token would
otherwise go over the wire in cleartext.

**The allowlist matches host shapes, not a fixed list of names**, because Fastmail pins some
accounts to a regional endpoint. Observed live: `apiUrl`/`uploadUrl` on
`phl.api.fastmail.com` (region as a leading label) and `downloadUrl` on
`phl-www.fastmailusercontent.com` (region hyphenated onto the `www` label). The two hosts
spell the region differently, so there are two anchored patterns rather than one. Session
discovery returning a regional host is normal Fastmail behaviour, not a redirect to be
distrusted — an exact-name allowlist rejected those accounts outright and made every tool
fail.

Region prefixes are limited to a single label (the pattern excludes `.`) and both patterns are
anchored at each end, so nothing can nest in front of a legitimate suffix
(`evil.phl-www.fastmailusercontent.com`) or append a foreign registrable domain behind one
(`phl.api.fastmail.com.attacker.com`). The prefix is allowed only in front of the two endpoint
hosts — `phl.www.fastmail.com` is still rejected.

`FASTMAIL_ALLOW_UNSAFE_BASE_URL` stays what it is: an opt-in for self-hosted JMAP servers that
drops host checking wholesale. It is not the answer to a regional Fastmail host, and using it
for that would disable the check for every URL in the session response.

## Credential logging is suppressed at the source (`DEBUG` and tsdav)

**Operator-visible behaviour: setting `DEBUG` no longer produces any tsdav output.** Every
other package's `DEBUG` logging is untouched, including under `DEBUG=*`. This is deliberate
and is not a knob.

tsdav (the CalDAV client) logs the HTTP Basic credential as bare base64 the moment `DEBUG`
is on:

```
tsdav:authHelper Basic auth token generated: <base64 of username:password>
```

That write goes straight to stderr from inside the library. It never passes through the
CallTool boundary in `src/index.ts`, so the redaction that covers every error egress
(`redactBearerTokens`, see `src/coerce.ts`) cannot reach it. Nothing downstream can scrub
it either, because the operator's own terminal or log collector is the destination.
Suppressing the logger at the source is the only control.

### The mechanism that actually ships

`src/index.ts` calls, at module scope:

```ts
createDebug.enable([process.env.DEBUG, '-tsdav*'].filter(Boolean).join(','));
```

Four things about that line are load-bearing, each established by running it:

- **Deleting `process.env.DEBUG` at startup cannot work.** The `debug` package reads the
  environment once, during its own module initialisation, and every logger created after
  that is enabled or not from the snapshot. Under ESM the whole import graph is evaluated
  before any importing module's body runs, so the chain `index.ts` -> `caldav-client.ts`
  -> `tsdav` -> `debug` has already initialised, and tsdav's loggers already exist, by the
  time the first statement in `index.ts` executes. There is no point early enough to win
  that race by editing the environment. `enable()` works because each logger's `enabled`
  getter recomputes from `createDebug.namespaces`, which `enable()` rewrites in place, so
  it applies to loggers that were created before the call.
- **The skip is `-tsdav*`, with no colon.** `-tsdav:*` matches only colon-prefixed
  children, which would leave a logger named bare `tsdav`, or a future `tsdavFoo`, still
  live. The bare glob covers every namespace the installed version creates
  (`tsdav:account`, `tsdav:addressBook`, `tsdav:authHelper`, `tsdav:calendar`,
  `tsdav:collection`, `tsdav:request`) and any sibling a later release adds.
- **The operator's own `DEBUG` is composed with, not replaced.** Skip entries beat enable
  entries, so `DEBUG=*` and `DEBUG=tsdav:*` are both suppressed for tsdav while
  `DEBUG=other:*` keeps logging normally. Only `DEBUG` gates the package; `NODE_DEBUG` is
  not consulted.
- **The control is instance-local.** It reaches only the `debug` module instance tsdav
  resolves. `package.json` therefore depends on `debug` at the exact version tsdav pins
  (`4.4.3`), so npm keeps one hoisted copy. A diverging range would let npm nest
  `node_modules/tsdav/node_modules/debug`, and the suppression would silently stop
  applying with no other symptom.

`src/built-server.test.ts` holds this down against the real library rather than a mock: a
control case asserts the credential *does* leak from a bare tsdav call under `DEBUG=*` (so
the suppression test cannot go vacuous the day tsdav stops logging it), a paired case
asserts it does not once the built server has loaded, and two further cases assert that
tsdav resolves the same `debug` copy and that the skip glob covers every namespace the
installed build actually creates.
