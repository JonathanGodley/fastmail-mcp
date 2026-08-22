# Email body handling

How this server composes, edits, reads, and reasons about the `text/plain` and `text/html`
parts of an email. This spans every authoring path (`create_draft`, `reply_email`,
`forward_email`, `edit_draft`, `send_draft`) and the read paths that undo
their quoting again (`get_email`, `get_thread`), so it lives here rather than in any one
tool's issue. The per-tool behaviour rationale lives in the closed GitHub issues
(#4, #7, #15, #16, #33, #73, #74); this file is the shared model they all depend on.

## The body-format model

HTML is the source of truth; `text/plain` is a derived fallback.

- When a caller supplies only `htmlBody`, the `text/plain` part is auto-generated from
  the HTML (html to text). When a caller supplies `textBody` explicitly, it is stored
  verbatim.
- We never fabricate HTML from plain text. The reverse direction (text to html) is not
  done anywhere; a `text/plain`-only message is legitimate and ships untouched.
- Degrade gracefully. If the HTML yields no derivable text (a newsletter that is nothing
  but remote images), the message ships HTML-only rather than being rejected. Only a genuine no-body message
  (no readable text and no visible HTML content) is refused.

The model is implemented in `src/body-format.ts`:

- `isBlank` — the single emptiness predicate. Strips zero-width / invisible characters
  (ZWSP, ZWNJ, ZWJ, BOM, soft hyphen) plus `trim()`, so a `&zwnj;&#8203;`-only body
  reads as absent. Shared by every emit gate so `''` / whitespace / zero-width-only all
  read as "absent" consistently.
- `htmlToText` — converts HTML to the readable plain-text fallback. Never throws (on
  converter failure it falls back to a minimal tag-strip so a send is never blocked).
  May legitimately return `''` for image-only / empty HTML. An `<img>` contributes its
  alt text and never its src or filename, so an image-only newsletter never emits junk
  like `[logo.png]`. What an alt-less image contributes is the caller's choice, made per
  call site through the image-policy parameter (below).
- **The image policy.** `htmlToText` takes one of three policies, because the same
  conversion serves two different jobs and they want opposite answers for an alt-less
  image. `suppress` emits nothing for it (the historical behaviour, and what a
  quotability probe wants, where a placeholder would make an image-only original look
  readable). `unconditional` emits `[image]` for an alt-less EMBEDDED (`cid:`) image, and
  is what the outgoing derivation uses: a picture-only message otherwise reaches a
  text-only reader as a blank page. `resolve` is the same but additionally requires the
  reference to resolve to a part the message actually carries, which is what a rebuilt
  quote needs — a reference whose image was dropped must not leave a placeholder standing
  for something no longer there. A remote (http/data/unknown) image with no alt contributes
  nothing under every policy, and the literal `cid:` token is never emitted in any of them.
  Non-image text derives byte-identically under all three.
- `htmlHasVisibleContent` — the reject gate for the no-body case. True if the HTML
  converts to non-empty text OR carries any visible-media element (`<img>`, CSS
  `background-image`, `<svg>`, `<video>`, `<picture>`, `<object>`, `<embed>`). It errs
  toward shipping: a false positive sends a thin email, a false negative would block a
  real one, so an imperfect scan is safe-by-direction. Its own answer is unchanged by the
  image policy: it already treats any `<img>` as visible content, so whether a placeholder
  is derived for one cannot move the gate either way.
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
  test rejected real messages: `Hi &lt;name&gt;, see attached.`, `mail me at &lt;a@b.example&gt;`,
  `Please reply with &lt;approve&gt; or &lt;reject&gt;.` So the escaped tag NAME must be a
  known HTML element, followed by a genuine tag delimiter (whitespace, `/`, or the closing
  `&gt;`) — which is what tells `&lt;a href=…&gt;` from `&lt;a@b.example&gt;`.

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
the four compose paths, *before* a reply quote or forwarded-message block is merged in:
`buildReplyParams` (`src/reply-handler.ts`), `buildForwardParams`
(`src/forward-handler.ts`), `composeDraft` (`src/compose-handler.ts`), and
the top of `updateDraft` (`src/jmap-client.ts`, edit_draft's only caller — which is why the
guard sits in the client method there, alongside the rest of the edit-body rules).

`editDraft` (`src/edit-draft-handler.ts`) runs the same check ahead of its attachment
coercion and upload. That is an ordering belt, not a fifth seam: `updateDraft` stays
authoritative, and because the guard is a pure idempotent check on the caller's own input,
running it earlier refuses nothing new — it only stops a body that was always going to be
rejected from orphaning freshly uploaded blobs first. Its position above the attachment
coercion matters as much as its presence: the compose paths report a body defect ahead of
an attachments-item defect, and edit_draft agrees with them on identical input.

Two constraints pin that placement:

- **The merge masks the defect.** `jmap-client.ts`'s existing no-readable-body reject
  (`normalized.htmlOnly && !htmlHasVisibleContent`) would catch a bare CDATA-wrapped body,
  but a reply escapes it: the quoted original supplies the visible content the gate looks
  for, so `htmlOnly` is never set and the malformed new message rides through with its text
  part reduced to the quote alone. Same for a forward. The escaped-HTML test is masked the
  same way (the quote contributes the real tags the test looks for).
- **A merged body is not caller input.** `createDraft` cannot host this guard, because the
  reply and forward paths reach it with the quoted original folded in. A message that
  legitimately quotes an XML snippet would be rejected on reply, and the user has no way
  to edit the original to fix it. Validating only what the caller wrote keeps every
  reject actionable.

**`send_draft` is deliberately outside the gate.** It takes no body parameter, so a draft
authored elsewhere that already carries an escaped-HTML or CDATA body can still be sent.
That is a consciously accepted gap, for the same caller-input principle as above: the
stored body is content this server didn't author, and rejecting it at send time would
block mail the caller may not be able to rewrite (the existing empty-part reject still
catches a body that derives to nothing). Drafts created or edited through this server's
own tools have already passed the gate on write.

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

## The identity signature in the body model (#33)

The sending identity's configured sign-off is a third thing this server puts into a body
the caller did not write, alongside the reply quote and the forwarded-message block. It is
opt-in per call (`appendSignature`) and appears in no message that did not ask for it. Two
properties of the body model above decide almost everything about how it behaves.

**The text form is derived, not the configured one.** An identity carries both
`textSignature` and `htmlSignature` (RFC 8621 §6; Fastmail writes both and keeps them in
sync). It would be natural to write each into the matching part — and it would be wrong.
The text part of an HTML message is a *derived fallback*: the first `htmlBody`-alone edit
regenerates it from the HTML, which regenerates the signature along with everything else.
A verbatim `textSignature` would therefore be correct when written and silently different
after that edit, with nothing reporting the change. So whenever HTML ships, the text form
is `htmlToText(htmlSignature)` — the same value that edit will produce — and the configured
`textSignature` is used only for a message that ships no HTML at all, where nothing will
ever derive it. This is the same "HTML is the source of truth" rule the rest of this file
describes, applied to a fragment rather than to a whole body.

The corollary for a half-configured identity: an identity with only one form still signs
either kind of body. An HTML-only signature derives its text form; a text-only signature is
escaped into an HTML block. That escape is not the forbidden "fabricate HTML from a
plain-text message" — the HTML body already exists and ships either way; this only decides
what goes inside it, and skipping it would drop a signature the user really has.

Deriving the text form of an HTML-only signature **suppresses image placeholders**, and that
is load-bearing rather than cosmetic. The `unconditional` policy writes `[image]` for an
embedded image, which is right for the text alternative of a body that ships that image —
and wrong for a message that ships no HTML at all, where the recipient's entire sign-off
would be a line describing something no part of the message carries. (That is the same rule
`buildReplyBodies` applies to quoted images.) So an images-only HTML signature derives to the
empty string on that branch and is reported as `no-text-form` rather than shipped as a
placeholder; alt text, which always wins, still derives normally.

The **other** branch deliberately does not suppress, and the asymmetry is the whole point.
When HTML ships, the embedded image ships with it, so `[image]` is exactly what
`unconditional` is for — and it has to agree with the derivation downstream, because an
HTML-only draft's text fallback is derived from the *whole* signed HTML under that same
policy. Suppressing here would give a caller who supplies both bodies a different text
sign-off from one who supplies HTML alone, which is the by-itself drift this split exists to
prevent. So the two spellings of "my signature is a logo" are genuinely different outcomes on
this branch, not an inconsistency: an **embedded** (`cid:`) logo derives `[image]` and the
sign-off is present, while a **remote** (`http(s)`) one derives nothing at all — a remote
image writes no placeholder under any policy — and that one is reported as
`text-part-unsigned`.

**Placement is above the quoted history**, which is why the insertion lives in
`src/reply-quote.ts` rather than in `createDraft`. By the time a compose path reaches
`createDraft`, the quote or forwarded block has already been concatenated onto the body, so
anything appended there lands *underneath* the quoted message and reads as part of it. The
same trap exists on the edit path: `updateDraft` replaces `updates.htmlBody` with the
quote-appended body when it rebuilds a quote, so the signature step runs before that.

**Preservation on edit is keyed on the draft, not on the flag.** The block is wrapped in a
`<div class="fm-mcp-signature">`, recognised by `hasSignatureMarker` beside the two quote
marker families. When an edit writes a body, the merge rule above drops the unwritten
partner and replaces the written one wholesale — so an `htmlBody`-alone edit, the commonest
edit there is, would drop a signature the draft carried. Reading that as intent would be
wrong: the caller asked to change a body, not to remove a sign-off. So an omitted
`appendSignature` *preserves* — the current signature is re-appended when the stored body
carried the marker and the new one does not — and because that is the one signature outcome
nobody asked for in the same call, it is the one the result announces. `appendSignature:false`
is the deliberate way to take one off. Note the asymmetry with the quote guard next door: a
dropped quote must be *challenged*, because only the caller knows which message it quoted,
while a dropped signature can simply be re-read from the identity.

**What the marker's survival is proven against, and what it is not.** The class round-trips
through the JMAP store — written on create, read back unchanged on the next `Email/get`, which
is what every test here and the live probe exercise. It has **not** been measured against a
Fastmail *web UI* round trip: a draft this server signed, opened and saved in the browser, and
then edited here again. Two interactions follow from that gap and neither is tested:

- if the UI re-appends its own signature on each save, a draft this server signed gains a
  second sign-off that nothing here put there — the accumulation shape reported in
  [#135](https://github.com/JonathanGodley/fastmail-mcp/issues/135), arriving from the other
  side of the same draft;
- if the UI re-serialises the `class` attribute away, `hasSignatureMarker` goes false and the
  preserve path stops firing — silently, because "the draft carried no signature" and "the
  marker was lost" are the same observation from in here.

Settling either needs a measurement against the live UI, not a reading of this code. Until
then, treat "the marker survives" as a claim about the JMAP store only.

**Residual: PRESERVATION is HTML-only.** The marker is a class, so a draft with no HTML body
carries none. State the residual at its real width: a signature *this server appended* to a
plain-text draft is exactly as invisible here as a hand-typed one, because nothing about the
detector looks at where the text came from. So a plain-text draft that was signed keeps its
signature through any edit that does not write a body (bodies are untouched), but an edit
that rewrites its text body without passing `appendSignature:true` loses it. Accepted rather
than fixed: the alternatives are matching the signature's *text* against the identity
(fragile the moment either is edited by a character) or storing state outside the message
(there is nowhere to put it that survives the recreate). The edit-time flag is the recovery,
and it is documented on the parameter.

**Idempotence is NOT the same thing, and the text path has it.** Naming
`appendSignature:true` as the recovery above only works if asking twice does not sign twice,
and the HTML marker cannot provide that for a body with no HTML. So `applySignature` also
declines to append when the body it is given **already carries one of the identity's two block
forms among its own lines**. That is not the fragile match rejected above: both forms are built
from the identity by this same call, so neither side is a remembered value that can drift.
Without it the read-modify-write loop the residual prescribes (read the draft, edit the words,
send the body back with the flag still set) appends another copy every time round, which is the
shape a duplicated sign-off actually takes in practice.

**The rule is SUBTRACTIVE: cut off what cannot be the draft's own text, then match over the
rest.** `updateDraft` hands `applySignature` the caller's **whole** `textBody`, quoted history
included, so the sign-off it is looking for is usually somewhere in the middle rather than at
either end. One removal builds the haystack, and it keys on a thing that *is* a line:

- **The forward separator line and everything below it** — this server's
  `----- Original message -----` and Gmail's `---------- Forwarded message ----------`,
  anchored at line start, the same two shapes `hasTextForwardMarker` recognises. Because the
  anchor sits on the dashes, a `> `-quoted separator inside a reply quote does not cut the body
  short, which is the same answer `hasTextForwardMarker` gives.

What is left is compared against each block form as a contiguous run of **whole lines**.

**A reply quote needs no removal, and removing one was a bug worth recording.** The whole-line
comparison already keeps quoted content out, because a quoted line is `'> '` + text and
`> Regards,` is not `Regards,`. An earlier version of this rule *also* dropped every
`>`-prefixed line from the body as belt and braces, and that was the opposite of safe: the
block on the other side of the comparison is not filtered, and `signatureTextBlock` derives the
text form through html-to-text, whose `<blockquote>` output **is** `> `-prefixed. So any
identity whose HTML signature contains a `<blockquote>` — a quote-of-the-day sign-off — had a
needle line that no surviving body line could ever equal, and every pass of the read-modify-write
loop appended another copy of the sign-off in silence. Dropping the filter cannot introduce a
wrong match either: a `> `-prefixed body line can only equal a `> `-prefixed needle line, which
is the signature's own quoted content and exactly the match wanted.

**Why subtractive, and why nothing here parses an attribution line.** Four earlier versions of
this guard tried to find where history *begins* by recognising the attribution above the quote
(`On <date>, <name> wrote:`). Attribution lines wrap, localise and differ per client, so each
version was a fresh guess about one client with an unbounded supply of others left to be wrong
about — a hard-wrapped Gmail attribution is what broke the last one, and each break shipped two
sign-offs. A separator line needs no parsing: it matches or it does not.

**The two errors are not symmetric, and the rule is tuned for that.** A false positive (read as
signed when it is not) appends nothing *and* reports `already-signed`, so the caller sees it and
can write the sign-off into the body themselves. A false negative (read as unsigned when it is
signed) ships two sign-offs and says nothing at all. Only the second is invisible from the
caller's side, so wherever the rule cannot be certain it falls on the side of not appending.
That trade is the whole justification for the design; a "smarter", fuzzier match would swap the
announced error for the silent one.

**Accepted residual — a forward this server does not recognise as one.** An Outlook-style
`From:/Sent:/To:/Subject:` header-block forward carries no dashed separator, so it is not
removed, and a forwarded message whose own text carries this identity's sign-off reads as this
draft's own: the call reports `already-signed` and the caller's note goes out bare. That is the
*cheap* error by construction, and it is recovered by writing the sign-off into the note.
Consistent rather than accidental: `hasTextForwardMarker` does not recognise that shape either,
so no path in this server treats such a body as carrying a forwarded block. Extending separator
recognition is tracked as issue #144, and it is one change to both markers.

A *recognised* forward has no such residual: the block is removed wherever the sign-off sits
inside it, including at its very end, so the caller's note above it is signed normally.

**BOTH forms, not just the one this call would write.** Testing only the outgoing block was a
duplication bug, not a simplification. Which form is sitting in a body depends on whether the
call that put it there shipped HTML — an HTML-shipping call writes the form derived from
`htmlSignature`, a text-only call writes the configured `textSignature` — and for the ordinary
identity that configures both, those are different strings. So the permitted HTML→text
conversion (`clearFields:['htmlBody']` with the text the draft just handed back) offered the
derived form to a call about to write the configured one, and the sign-off stacked: on the
`appendSignature:true` recovery this residual names, and on the *omitted*-flag preserve path,
where the result additionally announced a re-append that had not happened. Matching both forms
is what closes it, and it costs nothing: they are two strings this call already has.

**The comparison is normalised, not byte-exact.** The body being tested has usually made a
round trip through a mail store, and a store is free to change it in ways that mean nothing to
a reader and everything to an equality test. Both sides get the SAME normalisation — any line
ending to LF, then each line trimmed at both ends — so it can only ever make a match more
likely, which is the announced-error direction. Four things it answers, each of which would
otherwise ship a second sign-off in silence:

- **CRLF, and a bare CR.** RFC 5322 says CRLF; a store emitting bare CR would leave the whole
  body reading as one line, matching nothing.
- **Stripped trailing whitespace.** `-- `, the RFC 3676 signature delimiter, *ends in a space*,
  so losing it is not exotic.
- **A space replaced by NBSP.** The trim uses `\s`, which covers U+00A0 and the other Unicode
  spaces, so the same delimiter still matches when its trailing space arrives as U+00A0.
- **Leading whitespace**, because of RFC 3676 *space-stuffing*: a `format=flowed` sender
  prepends a space to any line starting with a space, `>` or `From `. Fastmail does not emit
  `format=flowed` (verified live), but `edit_draft` receives drafts composed by other clients.

**An append that lands nowhere is reported.** `appendSignature` is an input the caller
cannot verify without re-reading the draft, so a flag that silently no-ops is
indistinguishable from one that worked — the never-silently-drop rule, applied to an input.
Every path emits a note naming which reason applied, and there are five:

| reason | what happened |
| --- | --- |
| `no-signature` | the identity has none configured, or `from` names no verified identity |
| `no-body` | this call wrote no body for the sign-off to sit under (a blank body counts as none) |
| `already-signed` | the body supplied already carries a sign-off this call would otherwise duplicate (quoted and forwarded history excluded) |
| `no-text-form` | the message ships no HTML, and the signature has no plain-text form to write into the text part |
| `text-part-unsigned` | the HTML body *was* signed, but the signature has no text form for the message's plain-text alternative |

The last one is the only **partial** outcome: a recipient rendering the HTML sees the sign-off
and one reading the plain-text alternative does not, so its note deliberately does not open
with "nothing was appended", which would be a plainly false report.

**`text-part-unsigned` does not depend on which bodies the caller passed**, and scoping it to
"the caller also supplied a text part" was a sixth outcome hiding as a silence. An `htmlBody`
alone still ships a plain-text alternative — derived downstream from the now-signed HTML, under
the same policy this derives under — so when there is no text form to derive, the recipient
reading that alternative sees no sign-off, identically to the case where the caller supplied
the text part themselves. Reporting only the second made the HTML-only call the one place in
this feature where a requested flag landed nowhere without saying so.

`no-body` covers a **blank** body as well as an absent one, on both formats. `htmlShips` is
`!isBlank(htmlBody)` because `buildBodyParts` drops a blank body and signing a part no
recipient sees is not signing; the text side reads the same way for the same reason. Falling
through on a blank `textBody` instead made the signature the *whole* body — the caller's own
(blank) text discarded, a signature-only message stored, and nothing reported — which
contradicted every wording on every surface, all of which say a blank body counts as none. The
same rule applies to a blank text alternative beside a real HTML body: it is left alone rather
than replaced with a bare sign-off, and the alternative is derived from the signed HTML.

`edit_draft` reports the same set from both sides. `appendSignature:true` is a request made in
that call, so it gets exactly the compose wording for every reason — including
`already-signed` and `no-body`, which it used to swallow; a flag that reports on three
surfaces and stays silent on the fourth is worse than one that never reports. An **omitted**
flag is the stored draft's earlier decision being preserved, so a failure there is a loss
rather than a declined request and gets loss wording instead. Two outcomes are not losses on
that path and stay silent: a supplied body that already carries the sign-off (which is what
preservation wanted), and an edit that writes no body at all (both bodies, signature included,
are carried through verbatim).

Note that `no-signature` is not the only way an edit's *keep* can fail. An identity whose
signature is images-only HTML is not signature-less, so a gate keyed on "the identity has
none" arms nothing for it — and an edit converting that draft to plain text then dropped the
sign-off in silence. The reason is read off the same plan the append runs, which is what makes
that gate total rather than a list of remembered cases.

**A note-less forward signs; a body-less reply does not.** `buildForwardBodies` makes the
signature the whole content above the forwarded-message block when the caller supplied no
note, because a bare "FYI, see below" forward is the normal shape of that tool and a sign-off
is the only content such a message would otherwise lack. `buildReplyBodies` has no equivalent
arm: a reply whose entire content is a sign-off over a quote is not a message, and a body-less
reply draft is an intermediate state meant to be filled in via `edit_draft` before sending.
The asymmetry is deliberate; the reply reports the unsigned outcome rather than inventing a
body for it.

**The sign-off follows the address written into `from`, not the identity resolved for the
recreate.** On the edit path these can differ: when the edit writes no `from` and the stored
one matches no verified identity, `updateDraft`'s identity resolution falls back to the
account default while the address actually written into `from` stays the stored one. Signing
from that fallback would put one identity's sign-off under another identity's address, so the
signature is resolved against the written address instead, and the draft losing its signature
that way is reported.

The **display name** follows the same address, but resolves in the *opposite* order (#152):
the name the stored draft already carries against that address wins first, and the verified
identity that owns the address is only a fallback for a draft that carries none.
`edit_draft`'s contract is that only passed fields change, so a caller who deliberately set a
display name on their own address must not have it silently reverted to the identity's
configured name by a later edit that never even touched `from` — a metadata-only edit (say,
changing only the subject or a recipient) was doing exactly that before #152. The account
default's name is still deliberately *not* a fallback — pairing it with a foreign address is
the identical drift, one step to the left of the sign-off.

One consequence of this order is a documented residual, not a bug: a draft whose `from`
matches no verified identity (an address removed from the account, or a draft made
elsewhere) keeps whatever name it already carried against that foreign address, because
there is no identity name to fall back to and no verified identity to overwrite it with.
`edit_draft` never invents or strips a name for an address the account cannot send as.

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

**Keep path where nothing quotable reaches the written body (loud-fail, not data-loss).** On
the keep path the quote is *rebuilt from the named original*. If nothing quotable lands in any
body the edit writes, `buildReplyBodies` returns that body unquoted — so a keep request would
yield a quote-less body. The guard checks for a restored marker and rejects with an actionable
error instead ("…has no quotable content… use noQuote…"). Two routes reach it, and they are
worth telling apart:

- **The named message has nothing to quote at all** — attachment-only, calendar-only, or a
  body of embedded images whose parts the draft can no longer carry. Reachable only by naming
  the wrong/empty original: a draft naming its own original can't hit it, because a quote
  exists only if that original was quotable and JMAP message content is immutable.
- **The edit wrote only a plain-text body for an original whose content is images.** A
  picture has no plain-text form, so there is nothing to quote on that side even though the
  html quote would have carried it. The repair is to write `htmlBody` as well. Note this case
  did not exist before embedded-image carry: a cid-image-only original used to be
  unquotable outright, so it fell into the first route.

Either way it loses no caller input (the new body is preserved) — it just turns a confusing
quote-less result into a loud one. A UX safeguard on a self-inconsistent request, not a
data-loss fix.

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

## The read side: stripping quoted history (#73, #74)

Everything above is the **compose** side — building a quote, and recognizing our own quote
on a draft we are about to rewrite. `src/quote-strip.ts` is the **read** side: given a
message's `text/plain` body, remove the correspondence quoted inside it. The two live in
separate modules on purpose, because the same word ("marker") carries a different burden in
each:

| | compose side (`reply-quote.ts`) | read side (`quote-strip.ts`) |
|---|---|---|
| Input | a draft **we or a client of ours** produced | whatever a **foreign** client produced |
| A match means | a guard fires (a confirmation prompt) | text is **deleted from the output** |
| Miss cost | the guard doesn't challenge an edit | quoted bytes stay in the response |
| False-positive cost | a needless challenge, resolved in one step | the reader loses real content |

The asymmetry sets the posture: **recognize confidently or not at all.** Every marker is a
conventional, machine-emitted shape anchored at line start — a leading `>` run (nesting is
the same shape), an `On <date>, <someone> wrote:` attribution *directly above* such a run
(including Gmail's wrapped two-line form), an Outlook `From:`/`Sent:`/`To:`/`Subject:`
**block** (a lone `From:` line is not enough — it needs two more header lines within the
next few, with nothing but headers and blanks between), and `-----Original Message-----` as
a whole line. Nothing matches → the body is returned byte-identical.

**Regions, not one boundary.** The naive implementation cuts everything below the first
marker. That is right for a top-posted reply and wrong for the two other real shapes: a
bottom-posted reply (the new text is *under* the quote) and an inline reply (interleaved
between quoted paragraphs) would both be deleted entirely. So each quoted run is removed as
its own region and unquoted text is never removed for being *positioned* after a quote. The
two marker classes that have no end delimiter — the Outlook header block and
`-----Original Message-----` — do run to the end of the message, because that is what they
mean.

**The signal is the safety net, and it is never silent.** A caller that asked to strip gets
`quotedBytesStripped` (UTF-8 bytes removed; **`0` = nothing matched, this body is whole**)
or `quotedStripSkipped` (there was no plain-text body). Emitting the `0` is a deliberate
exception to the omit-empty-fields rule: without it "nothing was quoted" and "the shape
wasn't recognized" and "the flag did nothing" are indistinguishable, and a caller cannot
know whether re-reading verbatim is worth a round trip.

**Accepted residuals** (all documented in the README, all reported through the signal).
They run in **both directions**, and it matters that the list says so: an under-strip
returns duplicated bytes, an over-strip returns *less than the sender wrote*. Under-strip is
the one to prefer at every fork, and the marker rules are tuned that way — but the markers
are conventions, not syntax, so over-strip is real and the reader has to know to watch
`quotedBytesStripped` for a number that looks too large for a short message.

*Under-strip (quoted history survives; `quotedBytesStripped` is 0):*

- **HTML-only quoting is out of reach.** Outlook's `<div>` nesting, or any quote flattened
  from HTML without `>` prefixes, has no text-level boundary. This is the same
  foreign-client recognition residual as the compose-side guard above, seen from the other
  end. Deriving text from HTML in order to strip it was rejected: `get_email` returns what
  the message *is*, and swapping a verbatim `bodyHtml` for a lossy derived-then-cut plain
  text would be a bigger change to the read contract than the token saving is worth.
  Rejecting the combination outright was also rejected — a caller cannot know a message is
  HTML-only before reading it, and on a thread read one such message would fail the whole
  call. It reports `quotedStripSkipped` and returns unchanged.
- **A wrapped quoted line whose continuation lost its `>` prefix** survives as a fragment.
  Deleting an unquoted line on suspicion is the one thing this module will not do.
- **Localized attributions** ("schrieb:", "a écrit :") are not recognized, so the
  attribution line survives above a stripped `>` run.
- **An Outlook header block whose `From:` carries no address** (Outlook can render a known
  contact as a bare display name) does not fire the block rule — see the over-strip entry
  below for why the address is required.

*Over-strip (content the sender wrote is removed; `quotedBytesStripped` is the tell):*

- **A leading `>` is not only a quote.** It is also a markdown blockquote and the prompt of
  a pasted shell/REPL transcript, and both are stripped. This is the feature working to
  spec — `>` runs are the first marker #73 names — not a bug to be heuristically narrowed
  (distinguishing "> the value MUST be a string" from a quoted line is not decidable from
  the text). Pinned by test so it stays a decision.
- **A forward is stripped like a reply.** A forwarded message's content sits below the same
  `-----Original Message-----` marker (Fastmail's own forward block uses those words too),
  so `stripQuoted` leaves the covering note. The marker is in the feature's stated scope and
  Outlook uses it for replies; the alternative — casing/spacing heuristics to tell a
  forward's dashed line from a reply's — is exactly the guessing this module refuses to do.
- **Pasted email headers that carry an address** still fire the block rule, and that rule
  cuts to the END of the message, so anything the sender wrote below the paste goes with
  it. This is the sharpest over-strip and it is why the rule requires an address-shaped
  token in the `From:` value: probing found a pasted job posting ("From: The Hiring Team"
  over To:/Subject: lines) losing two thirds of a short message including the sender's own
  question underneath. Requiring an address keeps the display-name-only pastes — the common
  newsletter/job-posting shape — out of the rule entirely, at the cost of the under-strip
  noted above. A paste that reproduces a real address is indistinguishable from a quote and
  is accepted.
- **A prose line can be pulled into a wrapped attribution.** The walk-back that catches
  Gmail's two-line "On … / wrote:" form cannot tell it from two ordinary sentences that
  happen to end in "wrote:" directly above a genuine `>` block, and eats both. Narrowing it
  further would drop the wrapped-attribution case, which is common; this is the accepted
  trade.

In every over-strip case the remedy is the same and is stated in the README: the response
carries `quotedBytesStripped`, so a number that looks too large for the message is the cue
to re-read that message without the flag. That is the whole reason the count is emitted
rather than the stripping being silent.

**Thread bodies (#74)** ride on the same function. `getThread`'s `includeBodies` switches
the `Email/get` property set to the defined `EMAIL_PROPERTIES_VERBOSE` superset (not a third
ad-hoc list) and fetches **text values only** — a thread multiplies body size by the message
count and the HTML alternative is the expensive half, so a message with no plain-text part
is flagged `bodyTextUnavailable` rather than served HTML. The combined bodies are capped
(`THREAD_BODY_BYTE_CAP`, `src/thread-handler.ts`) and the cap **errors** rather than
truncates, for the reason #59 exists: a silently shortened body is indistinguishable from a
short message. The cap is measured on what would actually be returned, which is what makes
"retry with `stripQuoted:true`" a real remedy rather than advice.

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
  reproduced html runs through the same sanitiser floor as reply quotes
  (script/style/handlers stripped, real http(s) images kept).
- **Quotability now includes embedded images**, which shifts that default. An original
  whose body is nothing but `<img src="cid:…">` used to have no quotable html — the
  sanitiser dropped every image, leaving a visually empty string — so a note-less forward of
  one fell through to the TEXT branch and reproduced nothing at all. Such a message is
  quotable when at least one of its references would really embed (resolves to exactly one
  part, declared an image, carrying a blob), so a note-less forward of it now emits HTML and
  shows the picture. The flip is deliberate: the default should reproduce the message, and
  for that message the only faithful reproduction is html.
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
- `asAttachment` forwards record the header (for `send_draft`'s keyword maintenance) but
  keep a deliberately non-marker filler body, and the guard does not arm on the bare
  header when the draft carries a `message/rfc822` attachment — the forwarded content
  lives in the `.eml`, which body edits can't drop — so the guard is genuinely inert on
  them; the `.eml` is an ordinary carried attachment and the recreate carries the header.
  Residual: any draft with an `.eml` attachment loses the header-floor challenge — a
  foreign draft with an unrecognizable inline block, and also this server's own inline
  forward OF an original that itself carried an `.eml` attachment (there the carve-out's
  premise is false: that `.eml` is unrelated content, not the forwarded message). The
  floor's loss only bites when the body markers are ALSO absent/mangled (markers still
  arm on their own), and in the foreign case the `.eml` still preserves the forwarded
  content; accepted rather than fetching/matching the attachment's blob to the original.
  Related accepted quirk: `removeAttachments`-ing the `.eml` keeps the recorded source
  (silently dropping provenance would be worse), so sending that draft unedited still
  marks the original forwarded; `noQuote` on any edit is the deliberate de-forward.

**Recognition residual (forward form, accepted).** Narrower than the reply residual: any
client that sets `X-Forwarded-Message-Id` is challenged via the header floor regardless of
its block shape (unless the draft carries a `message/rfc822` attachment — the carve-out
above). The remaining gap is a foreign forward draft with **neither** the header
nor a recognized dashed/div-cite shape — the guard is inert there, same accepted posture
(and same README surfacing) as the reply guard's foreign-quote residual. One structural
sub-case: Gmail's forward *html* cannot be marker-recognized at all — its wrapper is
class+text-keyed, and `hasForwardMarker` must key only on markup the quote sanitiser strips
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
  round-trips store/fetch AND the edit recreate exactly. Fastmail validates the value:
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

## Why trash + recreate is mandatory

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
create-then-dispose (create the new draft, confirm, then dispose of the old one) so there
is no data-loss window. A draft carrying an inline `cid:` image (a `multipart/related`
tree that can't round-trip through the flat draft fields) is rejected rather than
silently flattened. That reconstruction is tracked as a follow-on in issue #13.

### Disposing of the replaced draft: Trash, never destroy (#65)

The old copy is disposed of by moving it to the mailbox with the `trash` role (an
`Email/set update` of `mailboxIds` only — `keywords` are left alone, so the copy is still
a `$draft` and restoring it is a move back to Drafts; Trash retention applies by mailbox,
not by keyword). It used to be an `Email/set destroy`, which cost real data: an assistant edited
a draft from a body it had cached earlier, the user had meanwhile rewritten that draft in
the web UI, and the recreate silently replaced the user's version with the stale one. The
destroy made that unrecoverable — the previous draft was gone from the account, leaving
only a support-side backup restore.

The server can never detect this on its own (it cannot know the caller's copy is stale),
so the mitigation is on the disposal side plus disclosure:

- **Trash, not destroy** turns the overwrite into a one-step undo. How long the copy
  survives is the account's business: Fastmail's Trash retention is a per-account setting
  and can be set to never auto-purge, so "recoverable until Trash is emptied or
  auto-purged" is the honest claim — not a fixed expiry window (not live-probed).
- **There is no destroy fallback.** If the Trash move can't happen (no `trash` role on the
  account, a `notUpdated` entry, a thrown transport error, or a response that reports the
  id in neither `updated` nor `notUpdated`), the old draft is left exactly where it was and
  reported as `orphanedOldDraftId` + `orphanedOldDraftReason`. Success is confirmed
  positively from `updated` (RFC 8620 §5.3 puts every id in exactly one map, and `updated`
  maps id → object|**null**, so it is a key test, not a truthiness test) — inferring it
  from the absence of a `notUpdated` entry would report "recoverable in Trash" for a draft
  still sitting in Drafts. The edit itself already succeeded, so this never throws. A
  visible duplicate the caller can delete is strictly cheaper than the unrecoverable act
  this path exists to avoid.
- **A trashed draft stops counting as an active draft.** `get_thread` hides drafts and
  reports how many it hid, keyed on the `$draft` keyword; since the replaced copy keeps
  that keyword, every edit would otherwise add one to that count for the life of the Trash
  copy, and a thread with one real draft reply would warn about three. So a `$draft`
  message whose ONLY mailbox is the `trash`-role one is neither shown nor counted (a
  caller who wants Trash content reads Trash). If the `trash` role can't be resolved,
  every draft is counted as before — fail toward over-warning, never toward missing a
  real draft reply.
- **The result echoes back what was replaced** (`replacedDraft`: id, subject, to/cc, and
  body character counts), so a caller comparing against its own copy sees an unintended
  overwrite immediately. Sizes rather than the previous bodies: the old draft is intact in
  Trash, so its full content is one `get_email` away.

Exactly one of `trashedOldDraftId` / `orphanedOldDraftId` is always set — the fate of the
replaced draft is never left unstated.

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
  property edits. Only `keywords` and `mailboxIds` are mutable. Hence recreate-on-edit;
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
