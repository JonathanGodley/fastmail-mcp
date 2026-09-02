# Email body handling

How this server composes, edits, reads, and reasons about the `text/plain` and `text/html`
parts of an email. This spans every authoring path (`draft_email`, `edit_draft`,
`send_draft`) and the read paths that undo their quoting again (`get_email`, `get_thread`),
so it lives here rather than in any one tool's issue. The per-tool behaviour rationale lives
in the closed GitHub issues (#4, #7, #15, #16, #33, #73, #74); this file is the shared model
they all depend on.

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

**Where the gate sits, and why.** It runs on the CALLER's own body, at two seams, and at
both of them it runs *before* this server's own blocks are put into that body:
`composeDraftEmail` (`src/draft-email-handler.ts`, which is all three of `draft_email`'s
modes) and the top of `updateDraft` (`src/jmap-client.ts`, edit_draft's only caller — which
is why the guard sits in the client method there, alongside the rest of the edit-body
rules).

`editDraft` (`src/edit-draft-handler.ts`) runs the same check ahead of its attachment
coercion and upload. That is an ordering belt, not a fifth seam: `updateDraft` stays
authoritative, and because the guard is a pure idempotent check on the caller's own input,
running it earlier refuses nothing new — it only stops a body that was always going to be
rejected from orphaning freshly uploaded blobs first. Its position above the attachment
coercion matters as much as its presence: the compose path reports a body defect ahead of
an attachments-item defect, and edit_draft agrees with it on identical input.

Two constraints pin that placement, and both are about the same moment — the point where a
`{{quote}}` or `{{forward}}` token is replaced by the block built from the original:

- **An expanded body masks the defect.** `jmap-client.ts`'s existing no-readable-body reject
  (`normalized.htmlOnly && !htmlHasVisibleContent`) would catch a bare CDATA-wrapped body,
  but a reply escapes it: the quoted original supplies the visible content the gate looks
  for, so `htmlOnly` is never set and the malformed new message rides through with its text
  part reduced to the quote alone. Same for a forward. The escaped-HTML test is masked the
  same way (the quote contributes the real tags the test looks for). This is #78, and it is
  why validation is step 2 of `composeDraftEmail` and expansion is step 8.
- **An expanded body is not caller input.** `createDraft` cannot host this guard, because a
  reply or forward reaches it with the quoted original already substituted in. A message
  that legitimately quotes an XML snippet would be rejected on reply, and the user has no
  way to edit the original to fix it. Validating only what the caller wrote keeps every
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

The sending identity's configured sign-off is the one block this server builds out of account
data rather than out of another message, alongside the reply quote and the forwarded-message
block. Where it goes is the caller's decision and nobody else's: a signature appears exactly
where the body says `{{signature}}`, and a body that writes no token gets no sign-off. Two
properties of the body model above decide almost everything about what gets written there.

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
`buildQuoteBlocks` applies to quoted images.) So an images-only HTML signature derives to the
empty string on that branch and the token is reported as `no-text-form` rather than expanded
into a placeholder; alt text, which always wins, still derives normally.

The **other** branch deliberately does not suppress, and the asymmetry is the whole point.
When HTML ships, the embedded image ships with it, so `[image]` is exactly what
`unconditional` is for — and it has to agree with the derivation downstream, because an
HTML-only draft's text fallback is derived from the *whole* signed HTML under that same
policy. Suppressing here would give a caller who supplies both bodies a different text
sign-off from one who supplies HTML alone, which is the by-itself drift this split exists to
prevent. So the two spellings of "my signature is a logo" are genuinely different outcomes on
this branch, not an inconsistency: an **embedded** (`cid:`) logo derives `[image]` and the
sign-off is present, while a **remote** (`http(s)`) one derives nothing at all — a remote
image writes no placeholder under any policy — and the text part's token is reported as
`no-text-form`.

**Placement is the caller's, and nothing is placed for them.** The three builders live
together in `src/reply-quote.ts` because they feed one substitution: `draft_email` expands
`{{signature}}` on the body it composes, `edit_draft` expands it on a flagged edit, and the
rule deciding which form a part gets is one rule. Placement used to be this server's problem
and a delicate one — a sign-off appended after the quote had already been concatenated onto
the body landed *underneath* the quoted message and read as part of it, so the insertion had
to run before the concatenation on both the compose and the edit paths. A token has no such
ordering to get wrong: a caller who wants the sign-off above the history writes
`{{signature}}` above `{{quote}}`.

**The block carries no marker class, and everything that hung off one went with it.** It used
to be wrapped in a `<div class="fm-mcp-signature">` so that a later edit could recognise a
sign-off this server had written. That recognition existed to serve an automatic append: if a
body might get signed without being asked, something has to decide whether it is signed
already, and whether an edit that rewrote the body meant to drop the sign-off or merely
forgot it. All of it — the marker check as an append gate, the preserve-on-omitted-flag path
that re-appended a sign-off an `htmlBody`-alone edit would otherwise have dropped, the
plain-text matcher that cut a body at its forward separator and compared whole lines of what
was left against both configured forms — was machinery for guessing an answer the caller can
now simply state. A token says where the sign-off goes; a body handed back without one says
there is none. Writing an identifying class into every signed body bought a reader nothing
and claimed the block was this server's to manage, which it no longer is.

That also retires a residual worth naming as closed rather than leaving readers to look for
it: the marker was a `class`, so it existed only in HTML, and a signature on a plain-text
draft was invisible to every rule keyed on it. The asymmetry is gone because the rules are.

**On `edit_draft` the trigger is a flag, never the token's presence.** Part of a body handed
back to that tool was authored by the original message's sender, so any in-band trigger — a
token, a spelling, an escape convention — can be planted by them; a flag (`expandSignature`)
cannot. The consequence is the intended one: a stored `{{signature}}`, planted at compose
time or escaped on purpose, is stable under every unflagged edit, with no rule for the caller
to re-apply. Passing the flag is the caller claiming the written part as its own, so the
compose-side refusals apply to it: a flagged edit that wrote no `{{signature}}` anywhere is
refused, and so is a part carrying more than one.

**A token that expands to nothing is reported, never dropped in silence.** `{{signature}}`
is an input the caller cannot verify without re-reading the draft, so a token that quietly
vanishes is indistinguishable from one that worked. Every unexpanded token emits a note
naming the part it sat in and the cause, and there are two:

| cause | what happened |
| --- | --- |
| `no-signature` | the identity has none configured, or `from` names no verified identity — and, on an HTML part, an identity whose signature has no form at all to write there |
| `no-text-form` | the identity has a signature, but no plain-text form this part can carry: an images-only HTML signature, in a text part |

The split is not cosmetic. An identity that has a sign-off but cannot put one in *this* part
is a different sentence from an identity that has none, and only the second is fixed by
configuring a signature. The `no-text-form` case is also the only **partial** outcome: a
recipient rendering the HTML sees the sign-off and one reading the plain-text alternative
does not, which is why it is reported on the part rather than on the message.

`no-text-form` does not depend on which bodies the caller passed. An `htmlBody` alone still
ships a plain-text alternative — derived downstream from the now-signed HTML, under the same
policy this derives under — so when there is no text form to derive, the recipient reading
that alternative sees no sign-off, identically to the case where the caller supplied the text
part themselves.

**The sign-off follows the address written into `from`, not the identity resolved for the
recreate.** On the edit path these can differ: when the edit writes no `from` and the stored
one matches no verified identity, `updateDraft`'s identity resolution falls back to the
account default while the address actually written into `from` stays the stored one. Signing
from that fallback would put one identity's sign-off under another identity's address, so the
signature is resolved against the written address instead, and a flagged edit that finds no
signature there says so rather than expanding the token into nothing.

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
elsewhere) keeps whatever non-blank name it already carried against that foreign address
(blank/whitespace-only counts as no name), because there is no identity name to fall back to
and no verified identity to overwrite it with. `edit_draft` never invents or strips a name
for an address the account cannot send as.

## How quoted history survives an edit (#37, #42, superseded by the body hash)

A reply draft carries the quoted original *inside* its body. At compose time the caller writes
`{{quote}}` (or `{{forward}}`) and `buildQuoteBlocks` / `buildForwardBlocks`
(`src/reply-quote.ts`) substitute an attributed, cited `<blockquote type="cite">…` into the
HTML part and an attribution plus a `> `-prefixed block into the text part. After that single
substitution the quoted history is ordinary body text. Nothing marks it as ours, and nothing
on the edit path looks for it.

**Because a body edit replaces the whole body, an edit that rewrites the body drops the
quote — and that is now the documented contract rather than a defect to be guarded against.**
`edit_draft` stores what it is handed, character for character. To keep a reply's quoted
original, read the draft and hand the whole body back with the edits made in it; the history
survives because the caller sent it, not because this server detected it. The tool's own
description says so in those terms, and says the converse just as plainly: a body sent
without the quote drops the quote, with no challenge and no warning.

**What replaced the guard, and why the trade is worth stating.** The earlier design
(#37, redesigned #42) recognised the stored quote by its shape and *refused* a body edit that
would drop it, unless the caller either named the original so the block could be rebuilt from
it or asked explicitly for a bare body. That guard is gone in every part: the shape
recognition on both formats, the two flags that resolved a challenge, the four refusals it
raised, and the check that the rebuilt block had actually landed. The reason is that it
answered "did this edit drop the quote?" by recognising a shape, and shape recognition is
lossy in both directions at once. A quote from a foreign client in a shape it did not know
was dropped in **silence** — the widest edge of the feature, and precisely the failure class
it existed to kill — while quote-shaped prose in a body this server had never written was
challenged for nothing.

`edit_draft` now answers a strictly weaker question and answers it exactly. Any edit that
writes or clears a body must carry the `bodyHash` that `get_email` issued for that draft,
which proves the caller is replacing the body it actually **read**. It does not prove the
caller kept any of that body: someone who reads a reply draft and deliberately sends back a
single line gets a draft holding a single line. What the hash removes is the *silent* drop —
you cannot overwrite quoted history you never saw — and it removes it for every draft
equally, foreign shapes included, because it is a fact about the bytes rather than a guess
about their meaning. The old guard was stronger wherever it recognised a quote and worthless
wherever it did not; the hash is uniformly weaker and uniformly total, and the second property
is the one that was missing.

**A `{{quote}}` handed back to `edit_draft` is text, not an instruction.** Neither history
token expands on the edit path, and neither may be removed either: the body may be a foreign
one handed back, so a token in it may have been planted by the original's author. It survives
the pass exactly as typed, and the result notes that it was stored as written — a note rather
than a refusal, because a refusal keyed on text somebody else wrote would recur on every edit
of that draft with nothing the caller could do about it.

**Where the hash cannot be issued or cannot be spent, the tool says so** rather than letting a
body edit through on a body nobody could have read faithfully. `get_email` withholds the hash
for a draft whose stored body no edit could reproduce — a part flagged truncated or with an
encoding problem, a part no read returns, or a body interleaving two parts of the same text
type (#85, #180) — and names recreating the draft as the way forward. Withholding rather than
issuing a hash that could never be spent is the same never-silently-drop rule applied to an
output field: the token the caller would have used is absent, and the reason is stated.

## The read side: stripping quoted history (#73, #74)

Everything above is the **compose** side — building a quote and substituting it into a body
the caller wrote. `src/quote-strip.ts` is the **read** side: given a message's `text/plain`
body, remove the correspondence quoted inside it. The word "marker" now belongs entirely to
this side. The compose side had markers of its own once, for recognizing a quote on a draft
it was about to rewrite; that guard is gone (see the section above), and with it the only
place in this server where matching a quote shape meant a *challenge* rather than a deletion.

What is left is the read side's stakes on their own, and they are the severe ones:

| | read side (`quote-strip.ts`) |
|---|---|
| Input | whatever a **foreign** client produced |
| A match means | text is **deleted from the output** |
| Miss cost | quoted bytes stay in the response |
| False-positive cost | the reader loses real content |

A miss leaves the response untidy; a false positive destroys content the reader will never
know was there. That gap sets the posture: **recognize confidently or not at all.** Every marker is
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

## Forwarding: the forwarded-message block (#30)

`buildForwardBlocks` (`src/reply-quote.ts`) reproduces the original *verbatim* below a
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
  platform's own forward wrapper. Deliberately NOT a `<blockquote>`: Fastmail's official
  client keeps the two shapes apart by tag name — replies get `<blockquote type="cite">`,
  forwards `<div type="cite">` — and matching that is what makes a forward written here
  render as a native one does.
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

**The forwarded-message header, and how a draft stops being a forward.** The original's
Message-ID is recorded as `X-Forwarded-Message-Id` (see the threading note above), which is
what `send_draft` reads to mark the original forwarded. A read exposes it as
`forwardedMessageId`, and `clearFields: ['forwardedMessageId']` is the deliberate way to
de-forward a draft: it stops `send_draft` marking the original, and on a forward draft it
drops the recorded `sourceEmailId` with it (the pointer refines the marking it rides on; on a
*reply* draft it stays, because the draft is still a reply to that instance). A malformed or
hostile Message-ID is deliberately treated as absent — Fastmail rejects CRLF/non-ASCII header
values and mangles embedded angle brackets, probed 2026-07-05 — so forwarding such an original
records no header at all.

Recovering a forwarded original across sessions is one lookup: `forwardedMessageId` from
`get_email`, then `search_emails` on the **bare** id (the full-text lookup matches the
bracket-less form; both probed working 2026-07-05, as is the RFC 8621 §4.4.1 `header` filter,
which stays unused).

**No guard extends to forward drafts, and the marker family that armed one is gone.** An
earlier design recognized the forwarded block by its shape — a `<div type="cite">` in html,
the Fastmail or Gmail dashed line in text — and challenged an edit that would drop it, arming
either on those markers or on the bare header, with a carve-out for a draft carrying the
forwarded message as a `message/rfc822` attachment and a floor that challenged any block shape
it could not recognize. All of it went with the reply guard and for the same reason: it
recognized shapes, so it was blind in silence to foreign forwards (Gmail's forward *html* was
structurally unrecognizable — its wrapper is class-and-text-keyed, and a marker may key only
on markup the quote sanitiser strips from embedded content, or pasted forwards would
false-trip it) and noisily wrong about bodies that merely looked like one. `edit_draft` now
stores the body it is handed and requires the `bodyHash` proving the caller read the body
being replaced; a forwarded block survives an edit because the caller sent it back.

`asAttachment` forwards are unaffected either way, and that is worth stating rather than
leaving to be rediscovered. Their forwarded content lives in the `.eml`, which no body edit
can drop; the `.eml` is an ordinary carried attachment and the recreate carries the header
alongside it. `removeAttachments`-ing that `.eml` deliberately leaves the recorded source in
place — silently dropping provenance would be worse — so sending such a draft unedited still
marks the original forwarded, and `clearFields: ['forwardedMessageId']` is how to stop that.

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

So `bodyValueForType` (`src/jmap-client.ts`) selects the value from the part in that list
whose type is the one being asked for, then keys into `bodyValues` by that part's `partId`.
"Is the one being asked for" is slightly wider than an equality test on `type`: a part that
declares **no** content type counts as the type of the list carrying it (RFC 8621 §4.1.4),
because that is how a recipient's client renders it, so such a part is found rather than
skipped past (#179). A naive "look up by list position / partId key" is insufficient:
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
- **Quoted-history shapes survive a store/fetch round-trip (2026-06-28).** Created html-only,
  text-only, and dual reply drafts, fetched each back raw, and compared the
  *stored-and-returned* bodies against what had been written. The html
  `<blockquote type="cite">` survives intact. Two text shapes appear, and the difference is
  the server's rather than the caller's: a caller-supplied text body (the text-only and dual
  cases) comes back as `wrote:\n> ` (one newline), but the html-DERIVED text fallback (the
  html-only case, where the server adds the text part) comes back as `wrote:\n\n> ` — a blank
  line between the attribution and the first `> ` line. Any rule written against one of those
  shapes has to tolerate the other. A *text-only* reply draft returns **no** `text/html` part
  (its one `text/plain` part aliases into both lists), so `bodyValueForType('text/html')` is
  undefined and `existingHtmlValue` is blank. An *html-only* reply draft is actually stored
  dual (the server derives and stores the text fallback); a genuinely text-part-less html
  reply draft only arises from another client. Covers only drafts this server makes;
  foreign-client shapes are assumed, not probed.
