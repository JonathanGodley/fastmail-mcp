# Email reference

## Label vs Move Operations
- **move_email/bulk_move**: Replaces ALL mailboxes for an email (folder behavior). Whatever else the message was filed under is dropped, so a message carrying three labels comes out of a move carrying one.
- **add_labels/remove_labels/bulk_add_labels/bulk_remove_labels**: Adds/removes SPECIFIC mailboxes while preserving others (label behavior). They work on **labels only**: the Inbox and the account's own user labels. A mailbox with any other JMAP role (archive, trash, junk/Spam, drafts, sent, snoozed, scheduled) is a **folder** and is rejected before anything is written. Removing the inbox label archives a message the way `archive_email` does, with the same caveats, and adding it puts a message back in the Inbox. Removing a message's last remaining label archives it, so a message is never left filed nowhere.
- **archive_email** is neither: it removes the Inbox membership and leaves everything else alone, adding the Archive folder only when doing that would leave the message filed nowhere. So a message filed in the Inbox plus a label keeps the label, and Archive is not added. See [Archiving does what the Fastmail client does](#archiving-does-what-the-fastmail-client-does).

## Mailbox fields

**Default**: `id`, `name`, `path`, `role`, `parentId`, `totalEmails`, `unreadEmails`, `totalThreads`, `unreadThreads`

**Verbose adds**: `myRights`, `sortOrder`, `isSubscribed`, `sort`, `autoLearn`, `autoPurge`, `purgeOlderThanDays`, `hidden`, `isCollapsed`, `identityRef`, `learnAsSpam`, `suppressDuplicates`, plus any other JMAP fields

`path` is the mailbox's **root-anchored full path**: every ancestor's name from the top-level folder down, joined with `/`, with no leading or trailing slash - `Archive/2026/Receipts`. A top-level mailbox's path is just its name. It is computed by this server, not a JMAP field, so it is absent from `raw` output. Every mailbox-taking parameter accepts a path (matched case-insensitively, segment by segment), so a path returned here can be pasted straight back into `mailbox`, `targetMailbox`, `parent`, or a `mailboxes` entry. See [Naming a mailbox](#naming-a-mailbox) for the full resolution order.

In the rare case a mailbox's parent chain never reaches a top-level folder (a loop, or a parent the account did not return), its path cannot be computed. That mailbox is still listed, without `path`, and a trailing note names its id - the field is never dropped in silence. Refer to such a mailbox by id.

Falsy `role` and `parentId` are stripped in default and verbose (use `raw` if you need `null` values).

## Naming a mailbox

Every parameter that names a mailbox - `mailbox` on the read tools, `targetMailbox` on `move_email`/`bulk_move`, `parent` on `list_mailboxes`/`create_mailbox`, each entry of a `mailboxes` array on the label tools, and each entry of `search_emails`' `requiredMailboxes`/`excludeMailboxes` - resolves through one matcher, in this order:

1. **Exact id**, matched case-**sensitively**.
2. **Role** - `inbox`, `archive`, `sent`, `drafts`, `trash`, `junk` - matched case-insensitively.
3. **Exact folder name**, case-insensitively.
4. **Root-anchored path**, `/`-separated with no leading or trailing slash, each segment matched case-insensitively.

There is no substring matching at any step. A folder name matching exactly one mailbox wins over reading the same text as a path, so a folder whose own name contains a `/` stays reachable by that name. That tie rule applies only where nothing else answers to the same text: if a top-level folder is literally named `Archive/2026` **and** a real `Archive` > `2026` nesting also exists, the reference is rejected as ambiguous rather than filed into the flat folder.

A `mailboxes`, `requiredMailboxes` or `excludeMailboxes` list given as one string splits on every comma, so a mailbox whose name or path contains a comma has to be passed in an array (or a JSON-encoded array).

Three failures are reported distinctly rather than as "not found", because they call for different corrections:

- A **name shared by several mailboxes** is rejected as ambiguous and the error lists the candidates by full path - retry with one of those, or with the id. Your spelling was not the problem.
- A **name that is also the path to a different mailbox** is rejected as ambiguous too, but a path cannot separate those two - the path is the text that failed - so the error names each candidate (which is the folder, which is the nesting) with its id, and the id is what you retry with. `list_mailboxes` returns each id next to its path.
- A **path that cannot be walked**, because some mailbox's parent chain never reaches a top-level folder, says so and points you at the id/role/name forms that still work.

Every array of mailboxes - the label tools' `mailboxes`, and `search_emails`' `requiredMailboxes`/`excludeMailboxes` - is all-or-nothing: nothing is applied and no search runs unless every entry resolves, and a single error names every failing entry, keeping typos and ambiguous names in separate buckets so one retry fixes them all.

## Scoping across several mailboxes

`mailbox` scopes a search to one folder. `search_emails` also takes two arrays for the multi-folder cases ([#26](https://github.com/JonathanGodley/fastmail-mcp/issues/26)); each entry accepts every form above, so a path or role you learned on `mailbox` works here too. Like every mailbox array, they are all-or-nothing: an entry that resolves to nothing rejects the whole call rather than being dropped.

- **`requiredMailboxes`** - the message must be in **all** of them. This is the "carrying several labels at once" query. `mailbox` is exactly its single-mailbox shorthand (`mailbox: "Receipts"` and `requiredMailboxes: ["Receipts"]` are the same query), and passing both folds them into one intersection.
- **`excludeMailboxes`** - **read this before using it.** It maps to JMAP's `inMailboxOtherThan`, which is *solely-in*: a message is hidden only when **every** mailbox it is filed in is on the list. A message in Inbox plus an excluded label still comes back. JMAP has no strict "not in this mailbox" filter at all, so absolute exclusion means filtering the returned results yourself.

The excluded entries are **unioned** with the default Trash/Spam exclusion into one excluded set, rather than applied as a second, separate condition. That is deliberately stronger: a message filed in {Trash, an excluded label} is hidden by the union, though neither exclusion on its own would hide it.

The two arrays sit on opposite sides of the default Trash/Spam exclusion, which is the easiest thing to guess wrong:

- `mailbox` and `requiredMailboxes` are **explicit scopes** - you named the folders to look in, so the default Trash/Spam exclusion and its note switch off.
- `excludeMailboxes` is **not** - the default exclusion and its note stay on. The one adjustment: naming Trash or Spam there takes that role out of the default set, so the note stops telling you to re-run with `includeTrash`/`includeSpam`, which could not override an exclusion you asked for yourself.

The same solely-in rule governs the default exclusion, since it uses the same JMAP key: a message cross-filed in Trash and a normal folder was never hidden, is already in your results, and is not in the withheld count. So "no note" promises that nothing filed **only** in Trash/Spam matched.

**Two notes say something other than a count**, and both mean re-run rather than trust the page:

- "…**the hidden count couldn't be confirmed**" - the exclusion ran, but the server did not return a number for what it withheld. Something was filed away; how much is unknown.
- "the Trash/Spam folder **couldn't be found, so it was NOT excluded**" - the role didn't resolve to a mailbox, so the exclusion never ran at all and these results **may include** Trash/Spam mail. This is the one case where the default exclusion is not in force despite the tool description saying it is.

Silence still means what it says: no note at all is the "nothing was withheld" signal.

```
// mail carrying both labels, wherever else it is filed
search_emails { "requiredMailboxes": ["Receipts", "2026"] }

// a sweep that skips a noisy folder (and still skips Trash/Spam by default)
search_emails { "query": "invoice", "excludeMailboxes": ["Newsletters"] }
```

## Identity fields

**Default**: `id`, `name`, `email`, `replyTo`, `mayDelete`, `textSignature`, `htmlSignature`

`replyTo`, and verbose's `bcc`, are passed through in JMAP's shape: an array of `{name, email}` objects, not `"Name <email>"` strings.

The signature fields are the identity's configured sign-off, closing line included when it has one. JMAP does not append it, so put `{{signature}}` where your message ends in a `draft_email` body (or an `edit_draft` body with `expandSignature:true`), with no closing line of your own above it, or write the field into the body yourself. An unset or blank signature is omitted like any other empty field ([#33](https://github.com/JonathanGodley/fastmail-mcp/issues/33)). See [Signing a message](#signing-a-message) for where the token goes and what happens when it has nothing to expand to.

**Verbose adds**: every other property the server returns for the identity, passed through as sent. Which ones appear is the server's choice; on Fastmail they have included `bcc`, `verificationState`, `showInCompose`, `saveSentToMailboxId`, `displayName`, `isAutoConfigured`, `enableExternalSMTP`, `server`, `port`, `ssl`, `addBccOnSMTP`, `saveOnSMTP`, `externalCredentialId`, `warnings`, `useForAutoReply` and `verificationCheckTime`, but none is promised

> **Drafts in reads.** `get_thread` excludes drafts by default; `search_emails`/`list_emails` include them. Drafts are identified by the `$draft` keyword, which survives a draft being moved out of Drafts, not by mailbox role, with one carve-out: a draft that now lives ONLY in Trash is neither shown nor counted (`edit_draft` and `delete_email` both leave a `$draft` copy there). Use `includeDrafts` (get_thread) / `excludeDrafts` (the two list/search tools) to override either default.

## Archiving does what the Fastmail client does

`archive_email` is not a move. Fastmail's Archive button **removes the message from the Inbox** and leaves every other folder and label untouched; the Archive folder is added only when removing the Inbox would otherwise leave the message filed nowhere. This server does the same thing, measured against the live client rather than inferred from the protocol:

| The message is filed in | What archiving does |
| --- | --- |
| Inbox + a label | The Inbox membership goes, the label stays, **Archive is not added** |
| Inbox only | It moves to Archive |
| Anywhere else, no Inbox | **Nothing.** It has already left the Inbox - the client calls this "already archived" - and that is a successful outcome |
| Trash, Spam, Drafts, Scheduled, Sent or Snoozed, no Inbox | **Refused**, saying why, and naming an alternative in this server where one exists (a scheduled send has none, and the refusal says so; for a snooze it names `move_email`, which also clears the snooze) |

Corroboration for the first row: of the 15,843 messages in Archive on the account this was measured against, **zero** are also in another mailbox. Fastmail never creates the Archive-plus-label state that a replace-then-add implementation produces.

The Inbox is checked **first**, so a message that is somehow in the Inbox *and* in Trash is archived (keeping Trash, gaining nothing) rather than refused - that is what the client does when shown the same state.

`archive_email` **never throws on a partial failure**: every id comes back in one of six buckets - `movedToArchive`, `removedFromInbox`, `notInInbox`, `refused`, `notFound`, `failed` - with counts that sum to the number of distinct ids you passed. The response carries a prose summary and, beside it, the result as JSON (`{ counts, results }`), so the per-id detail and the counts can both be parsed rather than read out of the sentence. Each entry carries `mailboxes`/`roles`: the **projected** filing for the two branches that wrote, the **observed, unchanged** filing for `notInInbox` and `refused` (which write nothing), and for `failed` the filing as **observed before** the attempted write - a weaker claim, because one `failed` sub-case is an id the server acknowledged in neither of its result maps, whose outcome nothing confirmed either way. Read `roles` for `archive` to tell whether a message is in Archive; the branch name will not tell you, because a message that was in Inbox *and* Archive reports `removedFromInbox` and is in Archive.

Two exceptions to "each entry carries them". A `notFound` entry has neither, because there is no filing to report - and so does the `failed` sub-case where the message's current filing could not be **read** - the server returned no `mailboxIds` object, or returned an empty one, which is not a filing a message can have. Its filing was never observed, which is precisely why it failed. And `roles` is omitted whenever nothing the message is filed in has a role. If an entry carries `unresolvedMailboxIds`, a mailbox id could not be resolved to a name: `mailboxes`/`roles` are **incomplete** for that message and those raw ids are the remainder. Those two read together - a missing `roles` with no `unresolvedMailboxIds` means "no role mailboxes"; a missing `roles` *with* them means the roles of the listed ids are unknown. Same caveat as the `unresolvedMailboxIds` on simplified email output: the full set reported for that entry is `mailboxes` ∪ `unresolvedMailboxIds`, and its presence is not an error. (For the two branches that wrote, that union is the projected destination set rather than an observed membership.)

Three things it does not promise:

- **It archives the messages you name, not the conversation.** Fastmail's Archive button acts on a single message or on the whole thread depending on a per-user display setting, and this server cannot see which one an account is using. Measured both ways on the same three-message thread: ungrouped, archiving one message left its two siblings sitting in the Inbox; grouped, one click cleared all three. The rule applied to each message is identical either way - only the fan-out differs. So to archive a conversation, pass every message id in it (`get_thread` lists them); pass one and the rest of the thread stays in the Inbox. Thread ids are not accepted ([#128](https://github.com/JonathanGodley/fastmail-mcp/issues/128)).
- **No keyword is written, which is not the same as the read state being untouched.** JMAP reports `$seen` only when *every* one of a message's per-mailbox copies carries it, so dropping an unread Inbox copy from a message whose other copy was already read flips it to read with no keyword write anywhere.
- **This describes an account in labels mode.** In folders mode a message has a single membership, so every archive is the move-to-Archive case. The server does not detect which mode an account is in ([#122](https://github.com/JonathanGodley/fastmail-mcp/issues/122)).

## Replying with a different subject

`draft_email` takes an optional `subject` in every mode. Omit it and the reply inherits `Re: <original subject>` (no double-prefixing, and `Re[2]:` or `RE :` counts as the prefix); pass it and your line is used verbatim, with no `Re:` added. A blank string is treated as omitted; a non-string is rejected. On `mode:'forward'` it behaves identically over the `Fwd:` default, so the two history modes do not differ on the same parameter. The `In-Reply-To`/`References` headers are built the same way either way, so **recipients thread the reply correctly whatever subject you choose**.

**Why this parameter exists, and the hazard it replaces.** Fastmail groups messages into a conversation by **subject as well as by the threading headers**, which has two consequences:

- **Changing the subject on a reply detaches the draft's display grouping.** The draft is given its own `threadId`, so it sits apart from the conversation it answers in the Fastmail UI. This is display-only and it resolves once the reply is sent.
- **Hand-rolling `inReplyTo`/`references` on `mode:'new'` under a mismatched subject is worse, and is not undoable.** Once such a draft exists, later drafts replying to that *same original message* are grouped onto the splinter thread too, including ones created afterwards with the correct `Re:` subject and the full reference chain. Deleting the offending drafts, emptying them from Trash, and recreating with a trimmed reference list all leave the mapping in place ([#68](https://github.com/JonathanGodley/fastmail-mcp/issues/68)).

So reply with `mode:'reply'`, with or without `subject`, rather than reconstructing a reply out of `mode:'new'` plus manual threading headers. Typing the prefix yourself is the half-way version of that mistake, and both compose and `edit_draft` say so: a `Re:`/`Fwd:`/`Fw:` prefix written into a `mode:'new'` subject, or edited onto a draft that carries no threading headers, draws a note naming the mode that would have threaded it ([#188](https://github.com/JonathanGodley/fastmail-mcp/issues/188)). It is a note and never a refusal — a fresh conversation reusing an old subject is legitimate — and each route has its own silencer: on compose, passing `inReplyTo`/`references` yourself; on `edit_draft`, a draft that already carries `In-Reply-To`, `References` or `X-Forwarded-Message-Id`.

## Reading long threads cheaply

Reading a conversation the obvious way costs one `get_email` per message, and each of those messages carries the whole earlier conversation quoted below it — at successively deeper quote depths. The longer a correspondence runs, the more expensive each additional message is to read, even though the new information per message stays constant. Two opt-in flags remove both halves of that:

- **`get_thread` with `includeBodies: true`** returns every message's `bodyText` in the same call. HTML bodies are never returned here (that is where the size risk lives), so a message with no plain-text part comes back flagged `bodyTextUnavailable: true` — fetch that one with `get_email` `verbose: true`. Hidden drafts are excluded *before* bodies are read, so an in-progress reply cannot land in a transcription.
- **`stripQuoted: true`** (on `get_email`, or on `get_thread` alongside `includeBodies`) removes the quoted history and leaves the new text. `get_thread({ includeBodies: true, stripQuoted: true })` is a whole conversation as each message's own words, once, in one call.

Both are opt-in and neither changes any default.

**What counts as a quote.** Detection is deliberately conservative and covers the conventional, machine-emitted markers: leading `>` quote runs (including nested `>>`), an `On <date>, <someone> wrote:` attribution directly above such a run (including Gmail's wrapped two-line form), an Outlook `From:`/`Sent:`/`To:`/`Subject:` header block, and `-----Original Message-----`. Quoted *regions* are removed rather than "everything below the first marker", so a bottom-posted reply (written under the quote) and an inline reply (interleaved between quoted paragraphs) both keep their text.

**How to read the result.** The response always says what happened, so you never have to re-fetch to check:

| Field | Meaning |
|-------|---------|
| `quotedBytesStripped: N` | N bytes of quoted history were removed |
| `quotedBytesStripped: 0` | Nothing matched — this body is **verbatim**, do not re-fetch |
| `quotedStripSkipped` | Stripping did not run; the string says why (there was no plain-text body) |

**Known limits, all by design.** They run in both directions, and `quotedBytesStripped` is the tell for both: a `0` where you expected a strip, or a number that looks too large for a short message. The remedy in every case is to re-read that message without the flag.

*Too little stripped:*

- **History flattened without `>` prefixes passes through unchanged.** The sender's own plain-text part carries the quoted history as an `On <date>, <name> wrote:` line followed by flush-left text, with no `>` lines; a message written in Front has a text part like this. With no `>` prefixes the quote has no end delimiter, so the only available cut runs to the end of the message and would delete a reply written below or between the quoted lines. Such a message comes back whole with `quotedBytesStripped` `0`, or loses only what a deeper recognised marker such as `-----Original Message-----` removes; this is the same foreign-client recognition residual as the compose-side quote guard (see [`docs/email-bodies.md`](../email-bodies.md)).
- **Localized attributions** ("schrieb:", "a écrit :") aren't recognised, so that one line survives above an otherwise stripped quote.

*Too much stripped:*

- **A leading `>` is not only a quote marker.** A markdown blockquote in the sender's own writing, or pasted shell/REPL/diff output whose prompt is `>`, is stripped as quoted text. Telling one from the other is not decidable from the text.
- **Pasted email headers count as a quoted section** when the `From:` line carries an address, and that rule cuts to the end of the message. A display-name-only paste (a job posting's "From: The Hiring Team", a newsletter) carries no address and passes through untouched.
- On a **forwarded** message, the forwarded content sits below `-----Original Message-----` too, so `stripQuoted` removes it and leaves your covering note.
- A **prose line ending in "wrote:"** directly above a genuine quote can be pulled in with the attribution (the same walk-back that catches Gmail's wrapped two-line attribution).
- An **inline answer that is both indented and pressed against the quote** — no blank line above or below it — is read as a quoted line whose wrap lost its `>` prefix, and goes with the quote. An answer written at the left margin, or one set off by a blank line, is kept; only a reply with neither habit matches.

*Other:*

- `stripQuoted` applies to **`bodyText` only**. A `bodyHtml` returned alongside it (`get_email` with `verbose`) is untouched.
- `stripQuoted` **cannot be combined with `raw`** (raw is unmodified JMAP), and on `get_thread` it requires `includeBodies`. Both combinations are rejected with an error rather than quietly ignored.
- `get_thread`'s combined bodies are capped at **100,000 bytes** (~25k tokens). Over that the call **errors**, naming the largest messages and telling you to add `stripQuoted: true` or fetch them individually, rather than truncating. The cap is measured on what would actually be returned, so `stripQuoted` genuinely brings an over-cap thread back under it.

## Message bodies

`draft_email` and `edit_draft` both take `textBody` / `htmlBody`. HTML is the source of truth and the plain-text part is derived from it when you don't supply one. In that derivation an image contributes its alt text, and an embedded (`cid:`) image with no alt contributes `[image]`; a remote image with no alt contributes nothing, so a body that is nothing but remote images still ships HTML-only. Three malformed body shapes are **rejected** with an `InvalidParams` error rather than stored, because each one otherwise ships and reaches the recipient looking wrong:

- **A non-string body.** Passing a number, array, or object where a body string belongs is rejected naming the parameter ([#62](https://github.com/JonathanGodley/fastmail-mcp/issues/62)). An omitted body, or an explicit `null`, just means "not provided".
- **An `htmlBody` that is entirely HTML-escaped** — escaped element tags (`&lt;p&gt;`) and no actual elements ([#71](https://github.com/JonathanGodley/fastmail-mcp/issues/71) / [#77](https://github.com/JonathanGodley/fastmail-mcp/issues/77)). Stored as-is it renders as literal `<p>` tags to the recipient, and the derived text part carries the same junk. The error names the fix: pass real markup, or use `textBody`. Escaped markup **inside** real elements (`<pre>&lt;p&gt;</pre>`) is legitimate and passes, as does an ordinary `&amp;`. Because the check only fires on a body with no real markup at all, it is deliberately narrow about what counts as an escaped tag — a known element name followed by a real tag delimiter — so ordinary prose like `Hi &lt;name&gt;`, `mail me at &lt;a@b.example&gt;`, or `reply with &lt;approve&gt;` still goes through.
- **A body wrapped in a CDATA section** ([#78](https://github.com/JonathanGodley/fastmail-mcp/issues/78)). The derived plain-text alternative loses the message entirely and a browser shows a stray `]]>`; on a reply or forward the quoted original hides the damage, so a text-only recipient sees only their own words quoted back. The mechanism is in [`docs/email-bodies.md`](../email-bodies.md).

The CDATA rule differs by format. In `htmlBody` a section is rejected **wherever** it appears (escape it as `&lt;![CDATA[` to show the token literally, which passes). In `textBody` only a body that **starts** with `<![CDATA[` is rejected, so an XML snippet quoted inside a plain-text message keeps working. A bare `]]>` with no opening token is left alone in both formats.

The `textBody` refusal names `htmlBody` as the place to fix it, because the caller that trips it usually never typed the token. An `htmlBody` may legally carry an escaped `&lt;![CDATA[`; that escape **unescapes** when the plain-text alternative is derived from the markup, so handing the derived text back on a later `edit_draft` is refused. Editing `textBody` alone cannot clear it — the part is regenerated from the HTML — so the markup is the only thing that can change.

All of these are checked against **your** body, before any reply quote or forwarded-message block is merged in, so the reply and forward paths are covered rather than shadowed by the quote.

## Signing a message

The sending identity's configured signature - the sign-off you set in Fastmail - is not applied by JMAP, so nothing signs a message unless you ask. In `draft_email` you ask by writing `{{signature}}` into `textBody`/`htmlBody` where your message ends; `edit_draft` expands the same token behind `expandSignature: true`. Nothing is added anywhere you did not put it. **The token is the sign-off**: it expands to the signature as configured, closing line included, so write no closing line of your own above it; only when `list_identities` shows a signature with no closing line does one go above the token ([#214](https://github.com/JonathanGodley/fastmail-mcp/issues/214)). The canonical placement is **above** the quoted original or the forwarded-message block; below one reads as part of the quote ([#33](https://github.com/JonathanGodley/fastmail-mcp/issues/33)).

**The result shows each join.** For every block inserted at a token (`{{signature}}`, `{{quote}}` and `{{forward}}` in `draft_email`, `{{signature}}` under `edit_draft`'s `expandSignature`), the result carries one line: the two lines above the token (or start of body), the block's first two and last two lines, and the two lines below it (or end of body). Between them is the spacing: a blank-line count in a text part, and in an HTML part the run of tags, comments and whitespace either side of the token. For the doubled sign-off in [#214](https://github.com/JonathanGodley/fastmail-mcp/issues/214) it reads:

```
{{signature}} joined in htmlBody: above "Attached is the drawing set, could you file it please?" / "Regards"; markup "</p><div>"; block "Kind regards," / "A. Sender" / "555 0100"; markup "</div>"; below: end of body
```

Look for a closing line of yours above the signature's own, spacing other than the canonical shape (one blank line in text; in HTML, the markup shown), the signature below the history rather than above it, and prose left below `{{forward}}`. To repair one, pass the `bodyHash` the same result carries to `edit_draft` with `bodyEdits`. Join lines are display text: each is neutralised and cut at 120 characters (a markup run at 80), so never use one as a `bodyEdits` `find`. In both tools the lines are shown as stored, a quoted or forwarded block's included.

When the token has nothing to expand to it is **removed, and the result says why** rather than shipping literal braces to the recipient. That is not an error - the draft is still saved - but it is not silence either. Two reasons arise:

- no signature is available for the address the message sends as - it has none configured, or (an `edit_draft` of a draft whose From is not one of your verified identities) there is no identity to take one from, and the note says which;
- the message ships no HTML and the identity's signature has no plain-text form (an images-only HTML signature: no HTML ships, so no image does either, and a bare `[image]` line would describe something the message does not carry).

In `draft_email`, a body with **no** `{{signature}}` in it is stored unsigned, and when the sending identity has a signature and a body you supplied shows something, the result carries a note saying so - so "this draft is unsigned" is never something you have to re-read the draft to find out. The receipt does not report it, and is omitted when it has nothing to list.

Which form you get follows the body, not which fields the identity has configured. An HTML body gets the HTML signature, and the plain-text alternative is **derived from that HTML** rather than copied from `textSignature`. A plain-text-only message gets `textSignature` as configured. An identity that has only one of the two forms still signs either kind of body; the missing form is derived.

An HTML signature that displays an **embedded** (`cid:`) image is refused when `{{signature}}` lands in an HTML body that ships, by both tools, unless an `attachments` item in the same call supplies that `cid`: the identity holds the signature's HTML but not the image. Write the sign-off into the body yourself, or remove the image from the signature in Fastmail's settings. A message that ships no HTML still signs from the text form.

`edit_draft` signs nothing on its own: a sign-off is there if you put it there and gone if you did not. `expandSignature: true` expands a `{{signature}}` **you wrote** in the body of that edit, resolved against the address the edit sends as; everything above about *which form* you get applies unchanged.

The flag exists because part of a body handed back to an edit was written by someone else, so an in-band token could be planted by them; see [docs/email-bodies.md](../email-bodies.md).

## Sending attachments

`draft_email` and `edit_draft` accept an `attachments` array (on `mode:'forward'` these are NEW attachments to add — the original's own attachments are carried separately, governed by `includeOriginalAttachments`/`asAttachment`, and the images its body displayed are carried as body content; see [Replying and forwarding with images](#replying-and-forwarding-with-images)).

Each item names its content in **exactly one** of three ways:

- `path` — a **local file** to read and upload. A bare filename or relative path resolves against the attach directory; an absolute path must fall within it. Needs `FASTMAIL_ATTACH_DIR`.
- `blobId` — content **already stored in the account**. Nothing is read off your disk; the blob is referenced. A blob carries no filename of its own, so `name` is **required** alongside it. Needs `FASTMAIL_ALLOW_BLOB_ATTACH`.
- `emailId` + `attachmentId` (together) — **one part of an existing message**, resolved to its blob. This is how you attach a subset of another message's attachments without forwarding the whole thing. `attachmentId` takes the same forms `download_attachment` accepts, except that a plain entry number is **rejected** here (see below). Needs `FASTMAIL_ALLOW_BLOB_ATTACH`.

Naming no source, naming two, or naming half of the `emailId`/`attachmentId` pair is rejected, by item index — as is a key belonging to a source the item did not choose, so `{ "blobId": ..., "attachmentId": ... }` is an error rather than a silent pick between them.

The remaining keys apply to every source:

- `name` — the filename recipients see. Optional on `path` (defaults to the file's basename) and on `emailId` + `attachmentId` (defaults to the part's own name); **required** on `blobId`, which has nothing to default to.
- `contentType` (optional) — a MIME type like `application/pdf`. Defaults to the type inferred from the file extension (`path`), inferred from `name` (`blobId`), or declared by the part itself (`emailId` + `attachmentId`). An explicit value is **echoed by Fastmail as-is** (not re-detected), so a wrong value rides out wrong.
- `cid` (optional) — a Content-ID that **embeds** the file in the body instead of hanging it off the end. Any simple token of up to 64 letters, digits, dot, dash or underscore is fine; a spelling copied out of HTML (`cid:logo`) or a header (`<logo>`) is accepted and normalised to the same value. Each item in a call needs a distinct one. See [Embedding an image in the body](#embedding-an-image-in-the-body).

Size caps (~25 MB/file, ~45 MB total) are a fail-fast guard on **local files only** — the other two sources are never read client-side, so there is nothing to measure. Fastmail's own limit ultimately governs in every case.

```json
{
  "mode": "new",
  "to": ["someone@example.com"],
  "subject": "Just the invoice from that thread",
  "textBody": "Only the one attachment, not the whole message.",
  "attachments": [{ "emailId": "M0f1...", "attachmentId": "Gd41f8..." }]
}
```

The `attachmentId` there is the `blobId` (or `partId`) of the part you want, read off `get_email_attachments`. A digit string that matches a real part is accepted, since Fastmail `partId`s are digit strings; one that matches no part falls through to the entry-number form, which is refused here with an error pointing you at `get_email_attachments`.

On `edit_draft`, `attachments` **appends**; removing attachments is covered under `edit_draft` in [Email Tools](tools.md#email-tools).

> **Two opt-ins, by design — one per boundary.** Both are off until you set them, and the server must be restarted to pick either up.
>
> **`path` needs `FASTMAIL_ATTACH_DIR`.** Attaching a local file reads it off the disk and emails it out — an exfiltration vector — so that source is **disabled until you set `FASTMAIL_ATTACH_DIR`** (see [Setup](configuration.md#configuration)); until then every attempt fails with a self-documenting error and no file is read. Files are confined to that directory: a path outside it, a missing file, or a symlink escaping the root is rejected. The attach directory is resolved **independently** of `FASTMAIL_DOWNLOAD_DIR` — pointing both at the same directory re-opens a download-then-email round-trip, so that is your explicit choice, not a default. The confinement narrows time-of-check/time-of-use races and blocks symlink escapes, but a same-inode swap race and hardlinks inside the root are residual.
>
> **`blobId` and `emailId` + `attachmentId` need `FASTMAIL_ALLOW_BLOB_ATTACH`.** They cross a different boundary: nothing is read off your disk, so the directory opt-in has nothing to say about them and they get their own flag, off by default and parsed strictly — only `true` or `1` enables it, so a literal `"false"` leaves it off rather than reading as "a value was set". A Desktop Extension install offers the same gate as an unchecked "Allow Attaching Content Already In The Account" box: the `"false"` an unchecked box sends is not an enable. What this gate protects is **provenance more than reach**: you could already put another message's attachment in front of an arbitrary recipient by forwarding it and sending the draft, but a forward carries a visible forwarded-message block and an `X-Forwarded-Message-Id` header, while a blob-attached part carries neither. Note also that a message's own `blobId` appears in ordinary read output, so with this gate open a caller can attach a **complete raw message** — transport headers and all — to a fresh draft; prefer `emailId` + `attachmentId`, which can only ever name one part.
>
> Full posture, including why `blobId` is not restricted to ids this server handed out, in [`docs/security-model.md`](../security-model.md).

## Embedding an image in the body

Give an attachments item a `cid` and reference it from `htmlBody`, in the same call ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)):

```json
{
  "mode": "new",
  "to": ["someone@example.com"],
  "subject": "The new logo",
  "htmlBody": "<p>Here it is:</p><img src=\"cid:logo\" alt=\"Our new logo\">",
  "attachments": [{ "path": "logo.png", "cid": "logo" }]
}
```

The same pair works in every `draft_email` mode and on `edit_draft`. What you can rely on:

- **A reference with no matching item is rejected**, rather than saved as a broken image. The error names the identifier and how to fix it.
- **A file you supply is never dropped.** If nothing displays it (no `htmlBody` at all, or no reference to that `cid`), it is still attached as an ordinary file and the result says so in plain words, so a silent demotion can't pass for success.
- **The saved draft is read back** whenever a call embeds something, and the result reports what it actually carries, with sizes. That is the one outcome you can't check from an email id alone.
- **Identifiers of the form `ii-<hex>@inline.invalid` are reserved** for images this server embeds on your behalf when it builds a quoted or forwarded block. Authoring one is rejected — it names an image that does not exist. Handing one **back** is not: on `edit_draft` such a reference is accepted whenever the draft actually carries a part under it, which is what lets you read an image-bearing draft and edit the prose around it. Drop the reference and the part comes off with it.
- The plain-text alternative derived from your HTML shows an image's alt text, and `[image]` for an **embedded** (`cid:`) image that has none — so a message of embedded pictures still reaches text-only readers. A **remote** (`http(s)`) image with no alt contributes nothing at all, so a body that is nothing but remote images reaches those readers empty. Give remote images alt text.

Embedding needs the same opt-in as any attachment, and either gate will do: with neither `FASTMAIL_ATTACH_DIR` nor `FASTMAIL_ALLOW_BLOB_ATTACH` set there is no image an `attachments` item could supply, and the refusal says so instead of pointing you at a parameter that can't work. With either one open, "add an `attachments` item supplying that `cid`" is a repair you can actually make.

## Replying and forwarding with images

When you reply to or forward a message whose body **displayed** an image, that image is carried into the quote or the forwarded block ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)).

This means a reply or a forward can send image data outward that you never attached and that `FASTMAIL_ATTACH_DIR` never sees — the parts are already in the account and are re-referenced by blob, so nothing is read off your disk and nothing is uploaded. What gets carried is bounded only by this:

- the body **references** the part, and
- the sender **declared** it `image/*`.

That declaration is metadata. Nothing is sniffed and nothing verifies the claim, so a part labelled `image/png` is carried whatever it actually contains. There is **no size limit and no count limit**, because re-referencing a blob costs nothing to send.

Your recourse is per-tool, and it is coarse on purpose:

- **`mode:'reply'`** — leaving `{{quote}}` out of the body omits the whole quote, and is the only way to send none of it. There is no quote-text-without-images option.
- **`mode:'forward'`** — body images are carried **even when `includeOriginalAttachments` is false**, because they are body content rather than attached files; that flag governs the original's attached *files*. There is no way to forward the message without its images: `asAttachment: true` re-composes nothing, but its `.eml` carries the whole original, images included. If they must not go, don't forward it.

Carrying needs an HTML body to put the images in:

- A **text-only reply** can't display them, so it drops them and the result says how many. Supply an `htmlBody` that places `{{quote}}` to keep them; a text-only reply never gets an HTML part built for it.
- A **text-only forward**, or an image the block can't display — a reference that didn't resolve to exactly one image part, or a duplicate identifier shared by several parts, or a `{{forward}}` placed inside a comment or an attribute so the block's images are never displayed — falls back to riding as a regular attachment. That fallback *is* governed by `includeOriginalAttachments`, and the result says what happened either way.

Images written as `data:` URIs are dropped and counted rather than converted, as are images referenced by a form a quote cannot carry — a relative or protocol-relative path, which resolves against an origin the new message does not have. The quote still ships, and the result says how many went. For how the identifiers this server mints for carried images (`ii-<hex>@inline.invalid`) behave on a later edit, see [Embedding an image in the body](#embedding-an-image-in-the-body).
