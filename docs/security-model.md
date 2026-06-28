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

## `originalEmailId` is an in-account read-and-embed primitive (accepted residual)

`reply_email` and `edit_draft`'s reply-quote keep path both take an `originalEmailId` and, on
the keep path, fetch that message's body and embed it (sanitized via `sanitizeForQuote`) into
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
  list.** Reading stats off the shared `getMailboxes()` list (and validating label `mailboxIds`
  against it) means a hidden/role-less mailbox's id now throws `InvalidInputError` rather than
  returning data. Accepted; "valid" is defined as "matches some `mailbox.id` in the fetched list."
- **Two-query hidden-count race.** The visible query and the count query run in the same
  `makeRequest` (one atomic snapshot — no race) where possible; the derivation tolerates a missing/
  garbled count by failing closed to the degraded note. A message moving between *any* two reads is
  an inherent, accepted residual, not a temporal guarantee.
