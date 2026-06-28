# Email body handling

How this server composes, edits, and reasons about the `text/plain` and `text/html`
parts of an email. This spans every authoring path (`send_email`, `create_draft`,
`reply_email`, `edit_draft`, `send_draft`), so it lives here rather than in any one
tool's issue. The per-tool behaviour rationale lives in the closed GitHub issues
(#4, #7, #15, #16); this file is the shared model they all depend on.

## The body-format model

HTML is the source of truth; `text/plain` is a derived fallback.

- When a caller supplies only `htmlBody`, the `text/plain` part is auto-generated from
  the HTML (html to text). When a caller supplies `textBody` explicitly, it is stored
  verbatim.
- We never fabricate HTML from plain text. The reverse direction (text to html) is not
  done anywhere; a `text/plain`-only message is legitimate and ships untouched.
- Degrade gracefully. If the HTML yields no derivable text (an image-only newsletter),
  the message ships HTML-only rather than being rejected. Only a genuine no-body send
  (no readable text and no visible HTML content) is refused.

The model is implemented in `src/body-format.ts`:

- `isBlank` — the single emptiness predicate. Strips zero-width / invisible characters
  (ZWSP, ZWNJ, ZWJ, BOM, soft hyphen) plus `trim()`, so a `&zwnj;&#8203;`-only body
  reads as absent. Shared by every emit gate so `''` / whitespace / zero-width-only all
  read as "absent" consistently.
- `htmlToText` — converts HTML to the readable plain-text fallback. Never throws (on
  converter failure it falls back to a minimal tag-strip so a send is never blocked).
  May legitimately return `''` for image-only / empty HTML. Emits `<img>` alt-text only
  (not the src/filename), so an image-only no-alt newsletter converts to empty text and
  takes the html-only path rather than emitting junk like `[logo.png]`.
- `htmlHasVisibleContent` — the reject gate for the no-body case. True if the HTML
  converts to non-empty text OR carries any visible-media element (`<img>`, CSS
  `background-image`, `<svg>`, `<video>`, `<picture>`, `<object>`, `<embed>`). It errs
  toward shipping: a false positive sends a thin email, a false negative would block a
  real one, so an imperfect scan is safe-by-direction.
- `normalizeBodies` — derives the fallback. html-present + text-absent derives the text
  from the HTML; if that derives to empty it returns html-only (an internal `htmlOnly`
  flag, not a reject). text-only and both-supplied pass through untouched.
- `buildBodyParts` — pure JMAP shaping, no fallback derivation. Builds the body-part
  arrays + `bodyValues` keyed by the literal partIds `text`/`html`.

A consequence worth stating for future changes: a "tighten this up to require a text
part" change would wrongly refuse legitimate image-only sends. The no-body reject is
deliberately the only reject.

## The asymmetric edit coupling

Because the text part is an auto-managed fallback of the HTML, `edit_draft`'s
cross-format coupling is asymmetric, not symmetric. The guards live at
`src/jmap-client.ts:704-714`:

- Edit `htmlBody` alone: the text fallback is regenerated from the new HTML. No throw.
- Edit `textBody` alone while a non-empty `htmlBody` survives: rejected. Editing the
  text alone "won't change what most recipients see" (they render `htmlBody`).
- `clearFields: ['textBody']` while `htmlBody` is present: rejected. The text fallback
  is managed automatically, so clearing it on its own is meaningless.
- `clearFields: ['htmlBody']`: the draft becomes a plain-text email.
- A no-body result (everything cleared) is rejected.

This shipped in commits `8dde79c` / `8afbf68` and **supersedes** an earlier symmetric
design (the "option-D" guard, commit `2fc8283`) where a single-body edit threw whenever
it would discard a non-empty opposite partner, in either direction. The body-format
model made the text side auto-managed, so the symmetric throw was replaced with the
asymmetric rule above. (Issue #4's resolution comment describes the shipped asymmetric
model; do not reintroduce the symmetric option-D description or the `2fc8283` citation.)

## Reply-quote preservation on edit (#37, redesigned #42)

A reply draft carries the quoted original *inside* its body — `buildReplyBodies`
(`src/reply-quote.ts`) appends an attributed, cited `<blockquote type="cite">…` to the
`htmlBody` and an attribution + `> `-prefixed block to the `textBody`. Because a body edit
replaces the whole body (the recreate above writes it verbatim), an edit that rewrites or
clears the body would silently drop the quote. `edit_draft` guards against this.

**The decision is made on the EXISTING (stored) body, not the caller's new body.** A reply
draft (one with an `In-Reply-To` header) "has a quote" when its stored `htmlBody` matches
`hasQuoteMarker` OR its stored `textBody` matches `hasTextQuoteMarker`. When the draft has a
quote and the edit touches the body in a way that isn't quote-preserving by construction (see
the carve-outs below), the edit is **rejected** unless the caller resolves it one of two ways:

- `originalEmailId` — the JMAP id of the message the draft replies to (NOT the draft's own
  `emailId`). The body the edit is writing is regenerated by re-quoting that named original
  from scratch via `buildReplyBodies`, so the caller's new text is kept *and* the quote is
  restored. This is the keep path for BOTH body formats (html, or a text-only draft's text).
- `noQuote: true` — deliberately drop the quote and store the bare new body.

Design points, each load-bearing:

- **Detection is on the EXISTING body, never the caller's new body.** This supersedes the
  fork.8 #37 approach (which scanned the caller's *new* html). Scanning new content is
  fundamentally bypassable and noisy: it can't tell the real quote from any quote-shaped
  content (a caller who drops the real quote but includes a different quote-shaped block would
  pass), and ordinary prose ending in "wrote:" false-positives. The stored body, by contrast,
  is one *this server* generated, so its quote shape is reliable. The markers (`hasQuoteMarker`
  on html, `hasTextQuoteMarker` on text) are tolerant *presence* checks that only govern
  whether the guard fires; `originalEmailId` is the authoritative way to keep the quote.
- **The original is the caller-named `originalEmailId`, never re-resolved from the draft's
  `In-Reply-To`.** `In-Reply-To` is an attacker-controllable header; resolving it to fetch a
  message would be a confused-deputy / quote-spoofing surface. The id is trusted, not
  validated against the draft's `In-Reply-To` (such a check would false-reject legitimate
  cases, e.g. correcting a wrong original).
- **The guard error names only the keep path.** `noQuote` is deliberately omitted from the
  error message so a model is never nudged toward discarding the quote; it stays discoverable
  via the schema for a caller who genuinely wants a bare reply.
- This regenerates from an explicit source; it never reassembles or splices the stored body
  (consistent with the "regenerate, never reassemble" posture of the body-format model).
- **Format flip:** supplying `htmlBody` to a *text-only* reply draft + `originalEmailId`
  converts it to a dual-body (regenerated html + derived text) draft. This is the caller's
  choice (they supplied html), accepted and pinned by test.

Carve-outs — quote-preserving *by construction*, so no flag is required:

- **Metadata-only edit** (subject / recipients / attachments; no body written or cleared) —
  both bodies are preserved untouched.
- **Plain-text conversion** — `clearFields: ['htmlBody']` alone keeps the stored text, which
  already carries the `> ` quote. This is a clean carve-out **only when the stored text
  actually matches `hasTextQuoteMarker`** (always true for drafts this server made). If it
  does not (a foreign draft, or a future divergence in our text shape), the edit correctly
  falls through to the guard rather than asserting the carve-out unconditionally.
- **Text-side edits while a non-empty html survives** stay owned by the two pre-existing
  body-coupling guards (textBody-alone → "edit htmlBody instead"; `clearFields: ['textBody']`
  while html present → "the fallback is auto-managed"), which emit the correct remedy. The
  quote guard excludes those cases so it doesn't pre-empt them. On a *text-only* draft the
  stored html is blank, so this exclusion does not apply and a text edit there falls through
  to the guard — exactly the #42 case the guard exists to catch.

**Cross-session recovery.** In-session the caller already has `originalEmailId` (it was just
passed to `reply_email`). Cross-session, a saved reply draft exposes its `inReplyTo` only as a
*Message-ID* string, not the JMAP id `originalEmailId` needs; recovering the keep path then
requires resolving the original first (`search_emails` for that Message-ID, with
`includeTrash:true`/`includeSpam:true` so a filed-away original isn't hidden by the default
Trash/Spam exclusion) before passing `originalEmailId`. The redesign makes `originalEmailId` the only keep path (there is no
inline-keep shortcut, deliberately — see below), so this lookup is the standard cross-session
keep recipe.

**Why no inline-keep shortcut (consciously declined).** Letting the caller re-include the
quote in the new body count as "keep" was considered and rejected: presence-as-keep is
bypassable in the same class as the superseded new-body scan (a caller who drops the real
quote but includes a different/edited quote-shaped block would be silently accepted as
"kept"). Requiring `originalEmailId` is the accepted price of having no bypass.

**Keep path with a non-quotable original (loud-fail, not data-loss).** On the keep path the
quote is *rebuilt from the named original*. If `originalEmailId` names a message with no
quotable content (attachment-only / calendar-only / cid-image-only), `buildReplyBodies` returns
the body unquoted — so a keep request would yield a quote-less body. The guard checks for a
restored marker and rejects with an actionable error instead ("…has no quotable content… use
noQuote…"). This is **reachable only by naming the wrong/empty original**: a draft naming its
own original can't hit it, because a quote exists only if that original was quotable and JMAP
message content is immutable. It loses no caller input (the new body is preserved) — it just
turns a confusing quote-less result into a loud one. A UX safeguard on a self-inconsistent
request, not a data-loss fix.

**Recognition residual (accepted) — the widest edge.** If a stored quote is in a shape the
markers don't recognize, `draftHasQuote` is false and the edit isn't flagged → a silent drop
(the failure class this feature exists to kill). Two faces of the same coupling to the
`buildReplyBodies` shape: (a) a draft created by *another* client; (b) a future change to our
own format without updating the markers. The generation-side CI pin (markers tested against
live `buildReplyBodies` output) guards (b). For (a), `hasQuoteMarker` recognizes the two
common machine-emitted html shapes — `type="cite"` (this server, Apple Mail, Fastmail web) and
Gmail's `class="gmail_quote"` — and the text marker catches most clients incidentally (they
also use `… wrote:` + `> `). The remaining gap is html-only quoting that uses neither shape
(e.g. Outlook's `<div>`-based quoting). The foreign-client shapes are **reasoned about, not
probed** across clients — a one-time probe of a real foreign reply draft would upgrade this
from "recognized in principle" to "verified." This is still the **broadest** edge of the
feature — wider than the non-quotable-original corner above, which needs a wrong argument to
reach, whereas this needs only an unrecognized draft from another client. Documented and
accepted; surfaced to users in the README's `edit_draft` notes.

## Why destroy + recreate is mandatory

JMAP email body properties are immutable and server-set (RFC 8621 §4.1.4); only
`keywords` and `mailboxIds` are mutable. So editing a draft's subject or body is done
by recreating the email, not patching it.

This was confirmed live against Fastmail (see the server-behaviour facts below): an
in-place `Email/set update` of `subject` / `bodyStructure` / `bodyValues` returns
`updated: {id: null}` (i.e. success) but silently changes nothing. Recreate is a
stronger justification than a hard reject would be: an in-place edit falsely reports
success while leaving the draft unchanged.

The recreate is faithful (`8afbf68`): it carries `In-Reply-To` / `References`,
re-references attachments by `blobId`, and preserves keywords. Ordering is
create-then-delete (create the new draft, confirm, then destroy the old one) so there
is no data-loss window; the response returns the new id plus `orphanedOldDraftId` if the
cleanup delete fails. A draft carrying an inline `cid:` image (a `multipart/related`
tree that can't round-trip through the flat draft fields) is rejected rather than
silently flattened. That reconstruction is tracked as a follow-on in issue #13.

## Body extraction: matching by MIME type, not list membership

Reconstructing a draft's existing bodies on recreate has one non-obvious trap, settled
by live experiments against Fastmail.

The server does not auto-generate the missing partner body in either direction at draft
storage time. A single-format draft has its ONE part aliased into BOTH the `textBody`
and `htmlBody` lists. For example, a text-only draft lists its `text/plain` part under
`htmlBody` too, with `type: "text/plain"`. RFC 8621 §4.1.4 keys `bodyValues` by
`partId`; the `textBody` / `htmlBody` arrays are independent lists of body-part objects.

So `bodyValueForType` (`src/jmap-client.ts:539`) selects the value from the part whose
actual `type` matches (`text/plain` / `text/html`), then keys into `bodyValues` by that
part's `partId`. A naive "look up by list position / partId key" is insufficient:
because the single part aliases into both lists, it would read the text value into the
HTML slot and synthesise a phantom `text/html` part on recreate. (This was the original
`|| true` extraction bug: both `existingTextBody` and `existingHtmlBody` collapsed to
`Object.values(bodyValues)[0]`, so a trivial subject edit silently destroyed the HTML
body. Since recipients render HTML, they saw the wrong content.)

### Edit matrix (12 cells, confirmed live, post-fix)

Evidence that the extraction is correct: every cell matched the traced prediction (no
corruption, no cross-contamination, no phantom, nothing lost). Single-format edits that
stay single-format keep exactly one `bodyValue` aliased into both lists (the server
representation, not a phantom).

| Start | Edit | Result | Note |
|-------|------|--------|------|
| text-only | textBody | text updated, stays text-only | clean |
| text-only | htmlBody | dual: old text + new html | text stale (low harm) |
| text-only | subject/to | text preserved, no phantom | clean |
| text-only | text+html | dual, both new | clean |
| html-only | htmlBody | html updated, stays html-only | clean |
| html-only | textBody | dual: new text + old html | recipient renders stale html |
| html-only | subject/to | html preserved, no phantom | clean |
| html-only | text+html | dual, both new | clean |
| mixed | textBody | text updated, html preserved (old) | recipient renders stale html |
| mixed | htmlBody | html updated, text preserved (old) | text stale (low harm) |
| mixed | subject/to | both preserved | clean |
| mixed | both | both updated | clean |

The "recipient renders stale html" cells are exactly what the asymmetric edit coupling
above now prevents: a text-only edit while a non-empty HTML survives is rejected,
because RFC 2046 §5.1.4 says a receiver renders the last supported alternative (HTML
ordered last), so a text-only edit does not change what an HTML-rendering recipient
sees. The matrix is the evidence; the coupling is the policy built on it.

RFC 8621 §4.1.4 also notes the decomposition of `bodyStructure` into `textBody` /
`htmlBody` / `attachments` "is not mandated, as this is a quality-of-service
implementation issue" — so the spec does not require a server to auto-generate the
missing partner, which is why the live experiments (not the spec) settled the keep /
drop question. Client-side fabrication of the missing partner is rejected (see issue
#8); the honest options for an unedited body are preserve or regenerate-from-html,
never fabricate-html-from-text.

## Live-probed Fastmail server-behaviour facts (raw JMAP, 2026-06-24)

Hard to reproduce (need a live account + throwaway drafts), kept here as durable
reference.

- **In-place content edits silently no-op, not reject.** An `Email/set update` of a
  draft's `subject` / `bodyStructure` / `bodyValues` returns `updated: {id: null}`
  (success) but the server ignores it; a re-fetch shows subject, `blobId`, and body
  unchanged. Fastmail does NOT return `notUpdated` / `invalidProperties` for immutable-
  property edits. Only `keywords` and `mailboxIds` are mutable. Hence destroy+recreate;
  the code comment rationale is "server silently no-ops," not "server rejects."
- **`cid:` inline images surface in `attachments` with `disposition: 'inline'`** (plus
  `cid`, `partId`, `blobId`; `hasAttachment: false`). So the strict
  `disposition === 'inline'` reject detector is correct and fires. `bodyStructure`
  round-trips the full `multipart/related` tree, so the #13 inline-image reconstruction
  follow-on is feasible, not blocked.
- **Composed `text/plain` carries no `format=flowed`** (a bare `Content-Type:
  text/plain`). So uniform `> ` quoting is correct; no RFC 3676 §4 flow handling is
  needed.
- **No default body-value truncation; `maxBodyValueBytes: 0` is rejected.** An
  `Email/get` with `fetchTextBodyValues: true` and no `maxBodyValueBytes` returned a
  5 MB body whole (`isTruncated: false`). An explicit `maxBodyValueBytes: 0` is rejected
  with `invalidArguments` (contra RFC 8621, where `0` means no truncation), so it must
  never be sent. `getEmailById` therefore needs no fetch-knob; the reply-quote module
  keeps only an `isTruncated` elision marker as a cheap defensive net for a hypothetical
  truncating server.
- **Reply-quote markers survive store/fetch round-trip (2026-06-28).** Created html-only,
  text-only, and dual reply drafts via `reply_email`, fetched each back raw, and tested the
  markers against the *stored-and-returned* bodies. The html `<blockquote type="cite">`
  survives intact, so `hasQuoteMarker` matches the returned html in every html case. Two text
  shapes appear, and `hasTextQuoteMarker`'s blank-line tolerance is **load-bearing** for one of
  them: a caller-supplied text body (the text-only and dual cases) comes back as `wrote:\n> `
  (one newline), but the html-DERIVED text fallback (the html-only case, where the server adds
  the text part) comes back as `wrote:\n\n> ` (a blank line between attribution and the first
  `> ` line). The strict `wrote:\n>` would miss the derived case; the `([ \t]*\r?\n)*`
  tolerance catches both. A *text-only* reply draft returns **no** `text/html` part (its one
  `text/plain` part aliases into both lists), so `bodyValueForType('text/html')` is undefined
  and `existingHtmlValue` is blank — which is exactly why a text-side edit there falls through
  to the quote guard (the #42 case). An *html-only* `reply_email` is actually stored dual (the
  server derives and stores the text fallback); a genuinely text-part-less html reply draft
  only arises from another client. This is why detecting the quote on the OLD body (#42
  redesign) is reliable for drafts this server creates. Covers only drafts this server makes;
  foreign-client shapes are assumed, not probed.
