# Path-confinement security model

Two tools touch the local filesystem: attachment download (writes a file, under
`FASTMAIL_DOWNLOAD_DIR`) and send-with-attachment (reads a file, under
`FASTMAIL_ATTACH_DIR`). Both are constrained to a configured directory and can never be
told to escape it. This spans both features, so the model lives here; the per-feature
rationale is in issues #5 (download) and #1 (attachments).

An `attachments` item can also name content that is **already in the account** — a
`blobId`, or a part of an existing message — and neither of those crosses the local-disk
boundary at all. `FASTMAIL_ATTACH_DIR` therefore has nothing to say about them; they are
gated separately on `FASTMAIL_ALLOW_BLOB_ATTACH`, and the posture behind that second gate
is its own section below.

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

Two more residuals of a different kind, recorded here so the whole accepted set reads in one
place. Both belong to the embedded-image carry (#13), which moves message parts outward
without reading a byte off disk, so no path guard is in play at all:

- **No count or size cap on carried images.** A quote or forwarded block carries every
  image the body it reproduces displays. There is no ceiling on how many, and none on how
  large: the parts are re-referenced by `blobId` rather than uploaded, so
  `MAX_ATTACHMENT_BYTES` — which caps only local reads — never applies (same footing as the
  `forward_email` carry sizes above). Fastmail's own message-size limit is the only bound,
  and an oversized send fails loudly server-side. Not capped because a cap would silently
  mangle the one thing the feature exists to preserve, and because the caller already chose
  to quote or forward that specific message.
- **SVG is carried inline like any other image.** A part the sender declared `image/svg+xml`
  and the body displays is carried into the composed message. SVG is a scriptable document
  format, so this re-sends attacker-authored markup under the user's own From — the quote
  sanitizer governs the *quoting* html this server writes, never the bytes of a part it
  carries by reference. It is worth stating that reply carry makes this **default-on for
  every reply to a message that displays an SVG**: no flag turns it on, and only
  `quoteOriginal: false` turns it off. Accepted because the receiving client, not this
  server, decides whether to render an SVG attachment, and because singling the type out
  would be a content filter this server does not otherwise attempt.

### edit_draft attachment model

`edit_draft` carries the existing attachments across the immutable-email recreate and then
applies the requested change: `attachments` **appends**; `removeAttachments` drops carried
parts by `blobId` (or a unique non-null `name`), rejecting a ref that matches nothing or a
name matching more than one; `clearFields:['attachments']` removes all. Passing
`attachments` together with `clearFields:['attachments']` is a rejected conflict. An
attachment-only edit stays body-invariant (it must not inject or strip a body).

The carried set is the **union** of the JMAP `attachments` array and the media parts the
server routed into the body lists, which is what makes images embedded in the body carriable
at all — and it means both removal routes reach them. That reach is deliberate and
disclosed in the tool descriptions: a removal is refused whenever the surviving body would
still reference what it took away, so the destructive case a caller can reach silently is
the one where nothing displays the part any more.

**Accepted residual — orphaned blobs on a late reject.** The handlers upload new
attachment blobs *before* the draft create/recreate runs, so a rejection raised by a later
guard (e.g. `edit_draft`'s refusal over a body part the recreate cannot reproduce, an
unresolvable `removeAttachments` ref, a reference to an image nothing supplies, or the
no-body-result guard) leaves the just-uploaded blobs unreferenced. `uploadAttachments`
orphans zero blobs *within its own batch* (a two-pass design), but that guarantee ends at
its return; the upload-then-reject ordering reopens a window. Accepted because Fastmail
garbage-collects unreferenced blobs (the same GC the code already relies on for a
mid-batch upload failure) — no unbounded growth, no data exposure.

### send_draft reads the parts to report, never to refuse (#13)

`send_draft`'s pre-send fetch asks for the draft's part listing (`attachments`, plus each
part's `disposition`, `cid` and `name`) for one purpose: the sentence that reports how many
embedded images the transmitted message carried. It is a **receipt, computed after the
submission** — never a send-time vet. Nothing in that listing can refuse a send.

That is a deliberate split, not an oversight. `send_draft` submits the stored draft **by
reference**: it transmits exactly the bytes already saved, and it cannot rewrite them. A
refusal there would therefore strand a finished message with no in-place repair, which is
why every refusal over message *shape* lives on the edit path instead, where the draft can
actually be fixed or replaced. The corollary matters when reading the `edit_draft`
refusals above: **a draft this server declines to edit is still perfectly sendable.** Those
refusals say "this server cannot rebuild that shape faithfully", never "this message is
unsafe to send".

### Authored Content-IDs are vetted where the message is assembled (#13)

A caller may name a Content-ID on an `attachments` item to embed that file in the body. The
value lands in a MIME header, so it is vetted at coercion down to a simple token (letters,
digits, dot, dash, underscore, 64 characters) with the `cid:`/`<>` spellings pre-stripped,
and every later comparison uses that canonical form. The vet is the block on header
injection, not a cosmetic tidy-up: a value carrying CR or LF stores as a genuinely injected
header and does not show up in the JMAP read-back.

The compose tools then decide which of those files the body displays, and that decision is
the whole of this server's say over what a message embeds. It is made **before anything is
uploaded**, so a refused call leaves no orphaned blob behind, and it is made against the
caller's OWN html rather than the merged body, so a rebuilt quote's server-managed
identifiers are never held against them. Consistent with the split above, nothing is
re-checked at send time.

Two directions, two different answers, both deliberate: a body reference with no matching
file is **refused** (it would ship visibly broken and the caller can fix it from the message
alone), while a supplied file nothing displays is **attached anyway and reported as such**
(those bytes are the caller's, and dropping them silently would be the worse failure).

The refusals stay honest about the opt-in, and about **both** of them. When neither
`FASTMAIL_ATTACH_DIR` nor `FASTMAIL_ALLOW_BLOB_ATTACH` is set there is no way to supply the
missing image at all, so the "add it to attachments" repair is dropped from the wording
rather than pointing the caller at a parameter that cannot work on this server. Either gate
being open restores the repair, because either source can supply the image — so the flag
threaded into the refusal builders is the combination, not the attach directory alone.

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

## Calendar attendees are an outbound-mail path that bypasses draft-first

The draft-first surface is built so that one named verb transmits: `send_draft`. That is
true of everything a caller *composes*, and it is the basis on which a name-based
permission system can gate a single tool. It is not true of the account as a whole, and the
exception is worth stating where someone setting up that gate will read it.

Naming attendees on a calendar event causes mail to be sent. When the event is written the
server emails each attendee an invitation from this account; deleting that event later
emails them a cancellation; changing the attendee list notifies whoever the change affects.
This is CalDAV scheduling doing its job — the sending happens server-side in response to the
write, not in any code here — so there is no draft, no `send_draft` call, and no flag in
this server that suppresses it. Confirmed by observation, not inference: an event created
with a single attendee produced an outbound invitation, and deleting it produced the
cancellation.

**The reach.** A caller that can reach `create_calendar_event` can cause mail to be sent to
an address of its choosing, with attacker-influenced text in the event title, description
and location, under the account's own identity. That is a smaller surface than a composed
message — the recipient sees a meeting invitation rather than free-form mail, and the body
is the event — but it is a real one, and it is reachable with `send_draft` fully denied.

**Why it is documented rather than blocked.** Refusing attendees would remove the feature
outright, since there is no way to write an attendee without the server scheduling it, and
the tool's purpose includes inviting people. So the control is disclosure: the tool
description says plainly that naming an attendee emails them, and the parameter is optional
— an event written without `participants` notifies nobody. An operator who needs the
guarantee that nothing leaves the account has to deny `create_calendar_event` and
`update_calendar_event` alongside `send_draft`, and that requirement is stated here because
the alternative is an operator believing one denial covers it.

## Attaching in-account content is its own opt-in (`FASTMAIL_ALLOW_BLOB_ATTACH`)

An `attachments` item names its bytes in one of three ways, and the gates split by which
boundary the source crosses, not by which tool is calling:

- `path` reads a **local file** and emails it out. That is the exfiltration vector
  `FASTMAIL_ATTACH_DIR` exists for, and it is unchanged by this feature.
- `blobId` and `emailId` + `attachmentId` reference content the account **already holds**.
  No byte is read off disk, no path guard is in play, and nothing is uploaded — the part is
  re-referenced by blob, the same mechanism a quote or forwarded block already uses. So
  `FASTMAIL_ATTACH_DIR` has nothing to say about them, and gating them on it would be a
  guard over a boundary they never cross. They are gated on `FASTMAIL_ALLOW_BLOB_ATTACH`
  instead, off by default, parsed strictly (`true`/`1` only, so a literal `"false"` cannot
  enable what the operator wrote it to refuse), and refused per source with a message naming
  the variable. It is a declared `manifest.json` setting as well as an environment variable,
  so a DXT install offers it as an unchecked box — the strict parse is what makes an unchecked
  box (`"false"`) mean off. The endpoint-allowlist kill switch stays environment-only; that one
  decides where the API token may be sent, which is not a capability to put behind a checkbox.

**What `emailId` + `attachmentId` actually adds is provenance loss, not reach.** The reach
is already available with the gate closed: `forward_email` produces a draft carrying the
original's attachments, and `send_draft` transmits it — so a caller that can forward can
already put another message's attachment in front of an arbitrary recipient. What changes is
what the recipient and the account can see afterwards. A forward carries the visible
forwarded-message block and records `X-Forwarded-Message-Id` on the draft; a fresh draft with
a blob-attached part carries neither, so nothing on the sent message says where the file came
from. The honest statement is therefore *equal reach, weaker after-the-fact detectability* —
not "no new capability".

**`blobId` reaches further than an attachment listing, and this is stated because it is easy
to get wrong.** The whole-message `blobId` is in `EMAIL_PROPERTIES_COMPACT`: it is emitted in
the DEFAULT output of every list, search and get, with no `raw` needed. So with the gate open,
a caller holding ordinary read output can attach a **complete raw RFC822 message** — the same
bytes `forward_email`'s `asAttachment` mode produces, with the full transport-header and
Sent-copy-`Bcc` exposure documented above — to a fresh draft that carries none of a forward's
provenance. It is not true that blobIds only come from attachment reads.

**Decision: `blobId` is NOT restricted to ids surfaced by `get_email_attachments`.** The
restriction was considered and declined, on the grounds that it would cost real capability
while buying no security:

- It is not enforceable without new state. A blobId is an opaque server-assigned string;
  deciding whether a given one "came from an attachment listing" would mean either keeping a
  server-side memory of what this process has emitted (lost on restart, and a per-session
  allowlist is not a security boundary) or issuing an extra read per item to re-derive the
  legitimate set — and the caller can always run that read itself and pass the id back.
- It would not close the raw-message hole anyway: a message's own blobId is reachable from
  `get_email_attachments` on any message that has a `message/rfc822` part, and from a forward.
- It would break the legitimate case the parameter exists for — attaching content the caller
  obtained from any read — while the id it would refuse is one call away.

The mitigation is the gate itself: the capability is off until an operator turns it on, and
the reach it grants is documented here rather than implied. `emailId` + `attachmentId` is the
narrower form and the one to prefer, because it resolves through the single attachment
resolution path and cannot name a whole message.

**A positional `attachmentId` is refused on the way out.** `download_attachment` accepts an
entry number (`0`, `1`, …) because a wrong read is a wrong read and costs nothing. Attaching
is different: the mis-resolved file is baked into a draft that `send_draft` then transmits,
so an `attachmentId` that resolved **only through the entry-number fallback** is rejected on
the compose path. The refusal is decided from **how the resolver matched**, never from
whether the string looks numeric — `parseInt` accepts `"2abc"`, so a string test would both
over-match (a real `partId` of `"2"`) and miss.

## `originalEmailId` is an in-account read-and-embed primitive (accepted residual)

`reply_email`, `forward_email`, and `edit_draft`'s keep path all take an `originalEmailId`
and fetch that message's body and embed it (sanitized — see the quote sanitizer below) into
a draft the caller may then send. Stated plainly: this lets a caller move one message's
content into outgoing mail addressed to arbitrary recipients under the user's own `From`. A
prompt-injected agent could use it to exfiltrate the content of any message in the account by
quoting it into a reply it sends to an attacker-chosen address.

**The reply path now moves BYTES, not only text (#13).** Before embedded-image support, a
reply carried zero parts of the original — the quote was text and markup, and an embedded
image was simply lost. It now carries the image parts the quoted body displays, re-referenced
from the account's own blob store, so a reply can put binary content in front of recipients
that the caller never attached and this server never read off disk. `FASTMAIL_ATTACH_DIR`
does not gate it: that opt-in governs reading local files, and nothing local is read here.
This is a genuine widening of the primitive above and is called out as its own line rather
than folded into it. The escape is `quoteOriginal: false`, which drops the whole quote; both
`reply_email`'s description and its `quoteOriginal` parameter say so.

The id is **trusted and unscoped within the connected account** — it may name *any* message,
deliberately, so a caller can correct a draft built against the wrong original. It is **never
re-resolved from the draft's `In-Reply-To`** (an attacker-controllable header), so there is no
confused-deputy / quote-spoofing surface from that direction, and there is **no cross-account
reach** (the fetch is scoped to `session.accountId`).

This introduces **no new capability class** versus the already-shipped `reply_email`, which
quotes any `originalEmailId` the same way; `edit_draft`'s keep path just reuses it. The
embedded html is run through the quote sanitizer (script/style/handlers/unscoped attributes
stripped, schemes pinned) — a safety floor for re-sending under the user's `From`, not a
privacy control. Documented here as an accepted residual: the mitigation for misuse is the
same opt-in/authorization posture that governs sending mail at all, not a restriction on which
in-account message may be quoted.

### What a quote or forwarded block carries, and under what identity (#13)

**The bound on what is carried, in full.** A part is carried into a body this server composes
when the body it is reproducing references it with a `cid:` image reference AND the part
declares itself `image/*`. That is the entire filter. The content type is *sender-declared*
metadata, exactly like a filename: nothing is sniffed, nothing verifies the declaration, and
a sender who labels an arbitrary file `image/png` gets it carried. There is no size bound and
no count bound (see the residual above for why). On `forward_email` the carry happens even
with `includeOriginalAttachments: false`, because a referenced image is body content rather
than an attached file, and a forward missing it reproduces a message with a hole in it —
short of not forwarding the message at all, there is no way to reproduce the body and leave
those images behind. A part that is referenced but *not* declared an image is not force-
carried; it falls to the ordinary attachment set the flag governs.

**Identity is a keyless shape, deliberately.** Each carried part is attached under a
Content-ID this server mints: `ii-<32 hex>@inline.invalid`, a 128-bit CSPRNG label under a
domain RFC 2606 reserves as permanently unresolvable. "Is this a part I manage?" is answered
by testing that shape and nothing else — there is no signature and no server secret. Signing
it would fail in the dangerous direction: rotating or losing the key would reclassify every
previously-managed part as foreign, and a body edit that should have *removed* an image would
start *sending* it. The shape check fails the safe way instead. Any change to how the
identifier is built must change the `ii-` prefix, so old and new forms stay distinguishable.

**The foreign-part walk that residual buys.** A message this server did not compose can carry
a part whose Content-ID coincidentally — or deliberately — matches that shape. The odds of an
accident are 2^-128; a forger has to try. When such a draft is later rebuilt (an `edit_draft`
that rewrites or clears the body), the classifier treats that part as server-managed, so an
unreferenced one is **deleted rather than degraded to an attachment**. What a forger achieves
is therefore the removal of their own content from a draft, and the bytes survive in Trash
(#65) either way. Accepted at those odds, with the walk written out here so the next reader
does not have to rederive why the safe direction is the deleting one. The same reasoning is
why a foreign Content-ID of that shape is **never carried verbatim**: parts pooled onto a
forward have their Content-ID stripped entirely, so a planted identifier cannot ride into a
message this server composed.

**Sender-supplied strings in this server's prose.** Content-IDs, filenames and content types
all appear in the notes and refusals these paths emit, and all three are attacker-controlled
on received mail. They are rendered as *quoted data*: control and format characters (which
can reorder or hide surrounding text), line and paragraph separators, and the double quote
that would close the quoted span are removed, space runs collapse, and the value is capped by
code point with an explicit ellipsis. Every call site wraps the result in double quotes, so
hostile text reads as data and never as the server speaking. The one exception is deliberate:
a Content-ID that passes the authorable vet is echoed raw inside a copy-and-paste repair
clause, which is safe by construction — that vet admits only `[A-Za-z0-9._-]`.

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
read + single-mailbox-write tools, and `search_emails`/`list_emails`/`get_recent_emails` hide Trash and Spam by
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
  boundary. The fail-closed degraded note exists so the published "no note ⇒ nothing was
  withheld" contract can be *trusted by a cooperative caller*, not to stop an attacker. Read
  that contract precisely: the exclusion is solely-in, so what the silence promises is that no
  message filed **only** in Trash/Spam matched — a message cross-filed in Trash and a normal
  folder was never withheld and is already in the results.
- **Count-into-Trash/Spam oracle, and the more direct `get_mailbox_stats` total-oracle.** The
  hidden-count discloses how many matches sit in Trash/Spam; `get_mailbox_stats mailbox:"trash"`/
  `"junk"` (single-mailbox = an explicit-scope override) hands those totals directly — the
  lowest-effort volume probe. Accepted; it's the caller's own account.
- **`search_emails`' `excludeMailboxes` (#26) is a caller-directed narrowing, and adds no
  concealment reach.** It hides results from *the caller who asked for them to be hidden*, in
  that one response — it writes nothing and changes no default, so there is no state an
  injected agent could leave behind for a later reader. It also does not weaken the note: a
  Trash- or Spam-role mailbox named there is dropped from the default exclusion's
  `excludedRoles`, so the note never prescribes an `includeTrash`/`includeSpam` recovery that
  the caller's own exclusion would override, and the hidden count keeps subtracting the
  default ids only (see `docs/conventions.md`). The one honest caveat is scope, not security:
  the parameter maps to JMAP's solely-in `inMailboxOtherThan`, so it hides less than its name
  suggests — a message cross-filed outside the excluded set still comes back.
- **Exact resolution hardens *mis-resolution*, NOT deliberate steering.** Switching every
  read/delete/move target from substring (`findMailboxByRoleOrName`'s name fallback, which could
  mis-hit e.g. a custom "Junk mail rules" mailbox and silently hide real mail) to exact id/role/
  name/path removes *fuzzy* mis-targeting. It does **not** close *deliberate* steering: an injected
  agent with `move_email`/`bulk_move` access can still aim mail at `"trash"`/`"Archive"` by exact
  name. Move-to-any stays open **by design** (a move-target restriction is tracked as fork #43).
  Name/role resolution also **lowers the steering bar** from "must know a valid opaque id (needs a
  prior `list_mailboxes`)" to "blind one-shot by literal name" — a real, if modest, escalation.
  **The label tools join this class (#50):** `add_labels`/`remove_labels`/`bulk_add_labels`/
  `bulk_remove_labels` now resolve their `mailboxIds` arrays by exact id/role/name/path too, so an
  injected agent can label a message into e.g. `"trash"` blind-one-shot-by-name, the same modest
  escalation as move. Accepted on the same footing; not a new capability class.
- **The path form (#27), and the collision it made reachable on write paths.**
  A root-anchored path is still exact matching over mailboxes a bare name could already reach, so
  it mostly just disambiguates. But the tie rule (**an exact flat name wins over reading the same
  text as a path**) had a real consequence on write paths: if a top-level folder is literally
  named `A/B` while a real `A > B` nesting also exists, `move_email targetMailbox:"A/B"` filed the
  message into the flat folder, silently, even though the same text describes the nesting — and
  anyone who can create a folder can set that collision up, so treat it as reachable by a caller
  who wants it. **That silent resolution is now refused**: a reference matching one flat name AND
  a *different* mailbox by path is rejected as ambiguous, naming both with their ids, so the write
  does not land anywhere until the caller picks one. The tie-break survives only where nothing
  else answers to the same text, which is what keeps a folder whose own name contains the
  separator reachable by that name. What remains accepted is the retry cost: a caller that meant
  the flat folder now has to name it by id, which is the cheaper side of the trade against a
  message filed where it was not meant to go.
- **`create_mailbox` adds no concealment reach.** It lets a caller mint a destination rather than
  pick one, but concealment comes from the *move*, and `move_email` into an existing folder already
  conceals with zero disclosure (the bullet below). Creating the folder first only decides where
  the mail lands; nothing about whether it is reported changes either way. The one genuinely new
  thing is minting a mailbox *named* after a role ("Archive", "Trash"): that is why the fixed
  destinations resolve by exact role only (`archive_email`, `delete_email`) and remain unreachable
  this way.
- **The hidden-count note covers ONLY Trash/Spam.** It does **not** disclose a `move_email` to
  `Archive` or a custom folder — that conceals mail with **zero disclosure**. So move concealment
  is not "mitigated by the note" for non-Trash/Spam destinations; this is stated plainly rather
  than implied covered. **`archive_email` (#21) sits in exactly this gap and is the cheapest way
  into it**: one call, one required argument, and mail leaves the Inbox with nothing reporting it.
  Its fixed destination narrows only where the mail can be aimed. An injected agent cannot choose
  the folder, but Archive is where it wanted the mail anyway. It adds no capability `move_email`
  (`targetMailbox:"archive"`) did not already have; it is the same concealment at a lower bar.
  Accepted on the same footing as move-to-any, and for the same reason: the restriction that would
  change it is the deferred move-target guard (fork #43), not a disclosure note.
- **`archive_email`'s destination is EXACT-ROLE ONLY, and that is a real (small) hardening.**
  Its destination is not caller-supplied at all: no `targetMailbox`, no name fallback, role
  lookup only. It shares that shape with `delete_email`/`bulk_delete`, whose Trash is found the
  same way, and those three are the only destinations on the server that work like it (see the
  resolver section of `docs/conventions.md`). The reason is that a caller can
  create a mailbox literally *named* `archive`; had the tool resolved names, "archive this" (a
  phrase an untrusted message body can plant) would file mail into an attacker-chosen folder
  under the innocuous verb. Against a *deliberate* attacker holding `move_email` this changes
  nothing (they name the folder outright). It removes the case where a **cooperative** agent
  running an innocuous instruction is steered by a folder someone else created.
- **Resolver error message is an information oracle, reachable account-wide.** A bad `mailbox`/
  `targetMailbox`/`mailboxIds` to *any* swept tool (search, list, stats, move, compose, labels)
  reflects the caller's input and a capped list of mailbox **paths** reachable by the configured
  token - since #27 these are full paths rather than bare names, so the oracle now also discloses
  the *shape* of the folder tree (which folder nests under which), not just the set of names. That
  is a slightly richer disclosure of the same material, accepted for the same reason: it is what
  makes the error recoverable, and it is the caller's own reachable tree. The real boundary is the
  token's reach, not "the user's own account" — a delegated/scoped token sees only its slice. `InvalidInputError` messages are run through
  `redactBearerTokens` as defense-in-depth (a token can't actually appear in them), but that is
  **not** what makes the oracle acceptable — recoverability (naming valid mailboxes so a caller
  can retry) is, and it's the caller's own reachable tree. Accepted, capped, framed honestly.
- **`get_mailbox_stats` and the label tools reject a real id that is absent from the fetched
  list.** Reading stats off the shared `getMailboxes()` list (and resolving label `mailboxIds`
  against it) means a hidden/role-less mailbox's id now throws `InvalidInputError` rather than
  returning data. Accepted; "resolvable" is defined as "matches some `mailbox.id`/role/name, or a
  path built from that same fetched list."
- **Per-message id-existence is a distinct oracle class.** A not-found id on `get_email`,
  `get_thread`, or `originalEmailId` now returns `InvalidParams` (a crisper signal than the prior
  `InternalError`), so it confirms whether a given *message/thread id* exists. This is a different
  class from the mailbox-resolver oracle above (which reflects the reachable mailbox *tree*) — it is
  per-message existence, and is likewise bounded by the **token's reach**, not "the user's own
  account." Accepted on the same footing: recoverability is the point, and it is dominated by the
  existing `get_email` read (the same probe already exists), so it adds no capability a caller with
  these tools lacked. `download_attachment` is on the same footing and says so: a bad
  `emailId`/`attachmentId` there is `InvalidParams` naming what to pass instead, because
  `get_email_attachments` enumerates the same parts on request for the same caller.
- **The same per-id oracle now extends to calendar events and contacts.** `update_calendar_event`,
  `delete_calendar_event`, `update_contact` and `delete_contact` return `InvalidParams` for an id
  that resolves to nothing, so each confirms whether that event or contact exists. `create_calendar_event`
  does the same for a `calendarId`. Accepted on exactly the footing above: every one of these ids is
  already readable through `list_calendar_events` / `get_calendar_event` / `list_contacts` /
  `get_contact` with the same token, so the rejection discloses nothing a caller holding these tools
  could not already enumerate, and the recoverability it buys is the whole point — a caller who
  cannot tell "wrong id" from "server broke" retries a call that can never succeed.
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

### No request follows a redirect

An allowlist that only checks the URL the client *aims at* is defeated by a 302: the
runtime would replay the credential at whatever host the response named, and no check ever
sees that host. So every credential-bearing request in this server is made with
`redirect: 'error'` — the JMAP fetches (session discovery, `makeRequest`, blob upload and
attachment download, all carrying the bearer token) and the CalDAV requests (all carrying
HTTP Basic).

On the CalDAV side the option is set once, on the `DAVClient` constructor
(`fetchOptions: { redirect: 'error' }` in `src/caldav-client.ts`), rather than per call.
tsdav merges its client-level `fetchOptions` into the init of every underlying fetch, so
one setting covers every method — including the ones this code never calls explicitly but
tsdav issues during its own login and discovery handshake. `src/caldav-client.test.ts`
asserts both halves, because either alone can pass while the protection is gone: that the
client the production path builds carries the option, and that the option still reaches
`fetch` when a real tsdav method runs.

## Credential logging is suppressed at the source (`DEBUG` and tsdav)

**Operator-visible behaviour: setting `DEBUG` no longer produces any tsdav output.** Every
other package's `DEBUG` logging is untouched, including under `DEBUG=*`. This is deliberate
and is not a knob.

tsdav (the CalDAV client) writes account details to stderr the moment `DEBUG` is on. Up to
2.1.x that included the whole HTTP Basic credential as bare base64:

```
tsdav:authHelper Basic auth token generated: <base64 of username:password>
```

tsdav 2.3.1 narrowed that line to the account name alone (`Basic auth token generated for
user "<username>"`), which for a Fastmail account is the user's email address. The
suppression is kept at full strength regardless: the passphrase leak is one release away
from returning, the surrounding namespaces still log collection URLs and calendar names,
and the account identity is itself worth not printing.

Those writes go straight to stderr from inside the library. They never pass through the
CallTool boundary in `src/index.ts`, so the redaction that covers every error egress
(`redactBearerTokens`, see `src/coerce.ts`) cannot reach them. Nothing downstream can
scrub them either, because the operator's own terminal or log collector is the
destination. Suppressing the logger at the source is the only control.

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
control case asserts the account identity *does* reach stderr from a bare tsdav call under
`DEBUG=*` (so the suppression test cannot go vacuous the day tsdav stops logging), a paired
case asserts that neither the identity nor the base64 credential appears once the built
server has loaded, and two further cases assert that tsdav resolves the same `debug` copy
and that the skip glob covers every namespace the installed build actually creates. The
control keys on the identity rather than the passphrase deliberately — which of the two
tsdav prints is tsdav's choice and has already changed once, so pinning the exact secret
form would turn a routine dependency bump red for no security reason.
