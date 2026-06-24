# Path-confinement security model

Two tools touch the local filesystem: attachment download (writes a file) and the
planned send-with-attachment (reads a file). Both are constrained to a configured
directory and can never be told to escape it. This spans both features, so the model
lives here; the per-feature rationale is in issues #5 (download) and #1 (attachments).

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

## Proposed: a read-shaped `safeReadPath` (not yet built)

Only `safeWritePath` exists today. The attachment-send feature (issue #1, unbuilt) needs
a read-shaped guard with different semantics, because `safeWritePath` is write-shaped:
it `mkdir -p`s the allowed root (`src/jmap-client.ts:1358`), walks missing path segments,
and checks the target only for an overwrite symlink. None of that is right for a read.

A `safeReadPath` should instead:

- require the target file to already exist (create nothing, no `mkdir`),
- `lstat` the final file and reject if it is a symlink (or realpath then re-verify
  containment),
- share the canonicalize-and-verify-containment core with `safeWritePath` rather than
  duplicating it.

Frame this as forward design when implementing #1; do not assume `safeReadPath` exists
in the current code.
