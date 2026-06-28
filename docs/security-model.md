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
