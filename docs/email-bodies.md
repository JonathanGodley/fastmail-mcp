# Email body handling

How this server composes, edits, and reasons about the `text/plain` and `text/html`
parts of an email. This spans every authoring path (`send_email`, `create_draft`,
`reply_email`, `forward_email`, `edit_draft`, `send_draft`), so it lives here rather than in any one
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

## Caller-supplied body validation (#62, #71/#77, #78)

Everything above assumes the caller handed us a body we can reason about. `assertBodyInputs`
(`src/body-format.ts`) is the gate that makes that true. It rejects three shapes with an
`InvalidInputError` (→ `InvalidParams` at the tool boundary), and they share a failure
signature: the call reports success, the tool result looks fine, and the defect is only
visible to a human who opens the draft or to the recipient.

- **Non-string body.** A present, non-string `textBody`/`htmlBody` used to reach `isBlank`
  and throw a raw `TypeError` (surfacing as `InternalError`, i.e. "retry" rather than "fix
  your input"). Now rejected by name. `undefined` **and `null`** both mean "omitted": null
  is how several lenient clients spell an unset optional field, and every downstream check
  already read it as absent, so accepting it preserves working calls rather than turning
  them into errors. Consistent with `coerceStringArray` / `coerceAttachments`, which also
  fold null into undefined.
- **Fully HTML-escaped `htmlBody`** — at least one escaped element tag (`&lt;p&gt;`) and
  zero real elements. Reject rather than unescape: unescaping guesses at intent, whereas
  the error teaches the right shape once. `htmlBody` only. Literal `<p>` characters in a
  plain-text body are ordinary content, and escaped markup inside real tags
  (`<pre>&lt;p&gt;</pre>`) is legitimate HTML.

  Both halves of the test are load-bearing, and the escaped half needs to be narrow. Because
  the guard only fires when there is **no** real markup, it is by definition judging prose —
  and in prose, escaped angle brackets are ordinary content. A loose "`&lt;` anything `&gt;`"
  test rejected real messages: `Hi &lt;name&gt;, see attached.`, `mail me at &lt;a@b.com&gt;`,
  `Please reply with &lt;approve&gt; or &lt;reject&gt;.` So the escaped tag NAME must be a
  known HTML element, followed by a genuine tag delimiter (whitespace, `/`, or the closing
  `&gt;`) — which is what tells `&lt;a href=…&gt;` from `&lt;a@b.com&gt;`.

  **Residual false-positive surface (accepted):** a body with no real markup whose escaped
  brackets happen to wrap a known element name *and* a tag-like delimiter — e.g. prose whose
  only markup-ish content is `&lt;code&gt;` or `&lt;table&gt;` used as a placeholder word.
  Much smaller than the original test's surface, but not nil. The remedy is in the error
  message either way (use `textBody`, or include real markup), and both keep the caller's
  words intact.
- **CDATA section**, with a deliberate asymmetry between the formats:
  - `htmlBody` — rejected **anywhere** in the body. The two parsers that read it disagree,
    and both outcomes are bad. Our `htmlToText` derivation (htmlparser2) recognizes the
    section and consumes it whole; measured: `<![CDATA[<p>Hi</p>]]>` derives to `''`,
    `<p>Before</p><![CDATA[<p>gone</p>]]><p>After</p>` derives to `Before\n\nAfter`, and an
    unclosed section swallows to the end of the string. So a mid-body section is as
    destructive as a wrapper, hence "anywhere" rather than "starts with". A browser takes a
    different route to the same mess: outside foreign content, HTML5 has no CDATA, so
    `<![CDATA[` starts a *bogus comment* that ends at the first `>` — the opening token and
    whatever precedes that `>` disappear, and the trailing `]]>` renders as visible text. The
    HTML half therefore fails visibly and the text half fails silently. The correct way to
    show a literal CDATA token in HTML is to escape it (`&lt;![CDATA[`), which passes.

    **Known false positive (consciously declined):** inline SVG that CDATA-wraps a `<style>`
    or `<script>` block — valid in the SVG/XML integration point, where a browser *does*
    honour CDATA. The reject stands: such a body is vanishingly rare in email, HTML5 mangles
    it anyway once the fragment is parsed as HTML rather than XML, and our text derivation
    would still swallow it. Recorded so the trade is visible rather than assumed absent.
  - `textBody` — rejected only when the trimmed body **starts with** `<![CDATA[`, i.e. the
    caller wrapped the whole body. A `text/plain` part is never markup-parsed, so an
    embedded CDATA token is inert, and a message quoting an XML snippet is real content that
    must keep working.
  - A bare `]]>` is **not** rejected in either format (consciously declined). The
    catastrophic swallow requires the opening token; on its own `]]>` renders as literal
    text and passes through the text derivation intact, so rejecting it would block ordinary
    prose and code snippets to prevent nothing.

**Where the gate sits, and why.** It runs on the CALLER's own body, at the seam of each of
the five compose paths, *before* a reply quote or forwarded-message block is merged in:
`buildReplyParams` (`src/reply-handler.ts`), `buildForwardParams`
(`src/forward-handler.ts`), `composeSend` / `composeDraft` (`src/compose-handler.ts`), and
the top of `updateDraft` (`src/jmap-client.ts`, edit_draft's only caller — which is why the
guard sits in the client method there, alongside the rest of the edit-body rules).

Two constraints pin that placement:

- **The merge masks the defect.** `jmap-client.ts`'s existing no-readable-body reject
  (`normalized.htmlOnly && !htmlHasVisibleContent`) would catch a bare CDATA-wrapped send,
  but a reply escapes it: the quoted original supplies the visible content the gate looks
  for, so `htmlOnly` is never set and the malformed new message rides through with its text
  part reduced to the quote alone. Same for a forward. The escaped-HTML test is masked the
  same way (the quote contributes the real tags the test looks for).
- **A merged body is not caller input.** `sendEmail` / `createDraft` cannot host this guard,
  because the reply and forward paths reach them with the quoted original folded in. A
  message that legitimately quotes an XML snippet would be rejected on reply, and the user
  has no way to edit the original to fix it. Validating only what the caller wrote keeps
  every reject actionable.

## The asymmetric edit coupling

Because the text part is an auto-managed fallback of the HTML, `edit_draft`'s
cross-format coupling is asymmetric, not symmetric. The guards live in `updateDraft`
in `src/jmap-client.ts` (the textBody-alone-while-html and clearFields:['textBody']-
while-html rejects, plus the no-body-result reject):

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

## Forwarding: the forwarded-message block + guard extension (#30)

`buildForwardBodies` (`src/reply-quote.ts`) reproduces the original *verbatim* below a
header block — no `> ` prefixing (that is reply quoting). The block matches the canonical
Fastmail shape, probed live 2026-07-05 by generating a native forward through Fastmail's
own official client and reading the draft back raw:

```
----- Original message -----
From: Ada Lovelace <ada@example.com>
To: Bob <bob@example.com>
Cc: Carol <carol@example.com>        (only when present)
Subject: Original subject
Date: 2026-07-01T09:14:00-04:00      (the JMAP sentAt string verbatim)
```

- **HTML form:** the same lines, each field escaped, in a plain `<div>` with `<br>`
  separators, followed by the original wrapped in `<div type="cite">…</div>` — the
  platform's own forward wrapper. Deliberately NOT a `<blockquote>`: reply markers are
  blockquote-anchored, so the two marker families are disjoint by tag name (Fastmail's
  official client confirms the split — replies get `<blockquote type="cite">`, forwards
  `<div type="cite">`).
- **Date line:** the `sentAt`/`receivedAt` string verbatim (ISO 8601 with offset), matching
  the platform block — deliberately not the humanized `formatReplyDate` shape replies use.
  Line-omission rule: a field with no usable value drops its whole line (no bare `Date:` /
  `To:` / `Subject:`).
- **Every interpolated field is attacker-controlled** (the original's From/To/Cc
  names+addresses, Subject) and re-sent under the user's From: the HTML form escapes each
  field; the text form whitespace-collapses each composed address/subject via
  `normalizeName`, whose class is `\s` **plus an explicit U+0085** — NEL is a mandatory
  line break per UAX #14 but is NOT in ECMAScript `\s` (verified empirically 2026-07-05).
- **Emission rules:** the caller's note (optional) goes above the block in each format the
  caller supplied; both supplied → both emitted with the caller's own text (a custom
  text alternative is never replaced by a derived fallback). Caller supplied neither →
  html only when the original has quotable html; a text-only original yields a TEXT
  forward — the "never fabricate HTML from plain text" rule above holds for the tool's
  own default choice. An attachment-only original gets the header block alone. The
  reproduced html runs through the same `sanitizeForQuote` floor as reply quotes
  (script/style/handlers stripped, real http(s) images kept, `cid:` images dropped).
- **Threading:** no In-Reply-To/References — a forward starts a new conversation
  (mainstream-client convention, confirmed by the official client). Instead the original's
  Message-ID is recorded as `X-Forwarded-Message-Id` (Thunderbird prior art; **Fastmail's
  official client sets the same header**, probed 2026-07-05).

**The edit_draft guard extends to forward drafts** (see the reply-quote section above for
the shared machinery; `guardVariant` dispatches mutually exclusively, reply-wins on a
pathological both-marker draft). Forward gating differs from reply gating in one deliberate
way: the variant engages when the header is present **OR** the body markers match
(`hasForwardMarker` = `<div type="cite">`; `hasTextForwardMarker` = the Fastmail or Gmail
dashed line, anchored so a `> `-quoted line never matches). Two postures produce that rule:

- **Marker-alone gating** protects forwards of a Message-ID-less original (no header could
  be set — and a malformed/hostile Message-ID is deliberately treated as absent, since
  Fastmail rejects CRLF/non-ASCII header values and mangles embedded angle brackets,
  probed 2026-07-05). Accepted false-positive cost, chosen data-loss-first: a draft whose
  body merely *carries* the conventional dashed shape (e.g. pasted forwarded content) gets
  challenged on a body edit — resolved in one step by `noQuote` (or `originalEmailId`),
  never lossy. Side benefit: header-less drafts in the conventional shape gain protection.
- **The header floor** challenges a forward draft whose block shape isn't recognizable
  (foreign header-setting clients — including Fastmail's own — or our marker lost to
  re-serialization). The challenge wording names the runnable recovery: `forwardedMessageId`
  via `get_email`, then `search_emails` on the **bare** id (the full-text lookup matches
  the bracket-less form; both probed working 2026-07-05, as is the RFC 8621 §4.4.1 `header`
  filter, unused). `noQuote` on a forward draft **also clears the header** — from either
  guard variant — so one step fully de-arms; otherwise the floor would re-challenge every
  later edit.
- The forward keep path has **no restored-marker check** (the block always regenerates, so
  it would be tautological). Flip side, intentional: a WRONG `originalEmailId` produces a
  valid-looking block for the wrong message with no loud fail — inherent to forwarding
  (any original yields a block); no caller data is lost.
- `asAttachment` forwards set no header and a deliberately non-marker filler body, so the
  guard is genuinely inert on them; the `.eml` is an ordinary carried attachment.

**Recognition residual (forward form, accepted).** Narrower than the reply residual: any
client that sets `X-Forwarded-Message-Id` is challenged via the header floor regardless of
its block shape. The remaining gap is a foreign forward draft with **neither** the header
nor a recognized dashed/div-cite shape — the guard is inert there, same accepted posture
(and same README surfacing) as the reply guard's foreign-quote residual. One structural
sub-case: Gmail's forward *html* cannot be marker-recognized at all — its wrapper is
class+text-keyed, and `hasForwardMarker` must key only on markup `sanitizeForQuote` strips
from embedded content (attribute-based), or pasted/quoted forwards would false-trip it.
Gmail's *text* dashed line IS recognized.

**Remote-image tracker note (accepted, inherited).** Forwarded HTML keeps real http(s)
images, same as reply quoting (the sanitizer is a safety floor, not a tracker filter) —
matching mainstream clients. Forwarding *broadens* the reply-quote posture's blast radius:
the new recipient's client loads the images, so the original sender learns the message was
forwarded (and roughly when/where). Accepted to match client convention; `asAttachment`
carries the original unrendered.

**Live-probed forward facts (2026-07-05, recorded here so they aren't re-derived):**

- `header:X-Forwarded-Message-Id:asMessageIds` is settable on `Email/set` create and
  round-trips store/fetch AND destroy+recreate exactly. Fastmail validates the value:
  embedded CRLF → rejected (`invalidProperties`); non-ASCII → rejected; embedded `<`/`>` →
  accepted but split into two mangled ids; a 1500-char id → accepted and folded. Hence the
  pre-vet in `forward-handler.ts` (printable ASCII, no whitespace/angles, ≤998 chars;
  malformed → treated as absent).
- Full-text `Email/query` finds a message by its **bare** Message-ID; the `<bracketed>`
  form finds nothing. The spec `header` FilterCondition also works on Fastmail.
- Attaching an existing Email's own `blobId` as a `message/rfc822` part stores a
  **byte-identical** `.eml` (the server dedupes to the same blobId).
- The official client's forward carries NO attachments (regular or inline) and dumps raw
  HTML into the text part of an html-only original — this server's attachment carry and
  `htmlToText` conversion are deliberate improvements, not parity.

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
