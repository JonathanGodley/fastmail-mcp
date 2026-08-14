# Fastmail MCP Server (Fork)

A fork of [MadLlama25/fastmail-mcp](https://github.com/MadLlama25/fastmail-mcp) — an MCP server for Fastmail's JMAP and CalDAV APIs.

This fork adds a **response simplification system** that reduces token usage when used with AI clients. All data-returning tools return a cleaned, curated format by default. Use `verbose` for all fields in the clean shape, or `raw` for the original JMAP response. See [Response Simplification](#response-simplification) for details.

## What this fork adds over upstream

- **Response simplification** — all data-returning tools return a token-lean shape by default (`verbose`/`raw` to opt out). See [Response Simplification](#response-simplification).
- **Calendar** — attendee/participant support, non-destructive event updates (no silent field wipes), and RFC 5545 date/TZID handling.
- **Local-time email dates** — the `date` field renders in your timezone with a UTC offset instead of raw UTC; `FASTMAIL_TIMEZONE` overrides the host zone.
- **Readable text/plain fallback** — when you supply only `htmlBody`, a plain-text alternative is generated automatically whenever one can be derived from the HTML (accessibility + deliverability). An image-only message with no derivable text still sends HTML-only; only a genuinely no-body send is rejected.
- **Faithful, reversible draft edits** — `edit_draft` preserves the draft's threading headers (In-Reply-To/References), attachments, and keywords across the immutable-email recreate, instead of silently dropping them. The draft it replaces goes to **Trash rather than being destroyed**, and the result echoes back what that draft contained, so an edit made from an out-of-date copy is both visible and undoable ([#65](https://github.com/JonathanGodley/fastmail-mcp/issues/65)).
- **Sending ergonomics** — drafts carry the identity's display name (parity with send); recipient strings like `"Name <email>"` are parsed across send/draft.
- **Outgoing attachments** — attach local files to a sent mail, draft, reply, forward, or edited draft (append / remove-by-ref / clear-all). Disabled until you opt in by setting `FASTMAIL_ATTACH_DIR`, and confined to that directory (reading a local file to email it out is an exfiltration vector). See [Sending attachments](#sending-attachments).
- **Forwarding** — `forward_email` reproduces the original under the Fastmail-native forwarded-message block and **carries its attachments** (the official server-side forward carries none), with an `asAttachment` mode that attaches the whole original as a lossless `.eml`. The edit-draft body guard extends to forward drafts.
- **Attachment paths** — relative `download_attachment` paths resolve inside the configured download dir, so a bare filename lands there in one step.

## Features

### Core Email Operations
- List mailboxes and get mailbox statistics
- List, search, and filter emails with advanced criteria
- Get specific emails by ID
- Send emails (text and HTML) with proper draft/sent handling
- Reply to emails with proper threading (In-Reply-To, References headers)
- Forward emails with attachment carry and a lossless `.eml` (`asAttachment`) mode
- Create, edit, and send email drafts (with or without threading)
- Email management: mark read/unread, pin/unpin, delete, move between folders

### Advanced Email Features
- **Attachment Handling**: List and download email attachments; attach local files to outgoing mail (opt-in via `FASTMAIL_ATTACH_DIR`)
- **Threading Support**: Get complete conversation threads
- **Advanced Search**: Multi-criteria filtering (sender, date range, attachments, read status)
- **Bulk Operations**: Process multiple emails simultaneously
- **Statistics & Analytics**: Account summaries and mailbox statistics

### Contacts Operations
- List all contacts
- Get specific contacts by ID
- Search contacts by name or email

### Calendar Operations
- List all calendars and calendar events
- Get specific calendar events by ID
- Create, update, and delete calendar events

### Label vs Move Operations
- **move_email/bulk_move**: Replaces ALL mailboxes for an email (folder behavior)
- **add_labels/remove_labels**: Adds/removes SPECIFIC mailboxes while preserving others (label behavior)

### Identity & Account Management
- List available sending identities
- Account summary with comprehensive statistics

## Setup

### Prerequisites
- Node.js 20+ 
- A Fastmail account with API access
- Fastmail API token

### Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Build the project:
   ```bash
   npm run build
   ```

### Configuration

1. Get your Fastmail API token:
   - Log in to Fastmail web interface
   - Go to Settings → Privacy & Security
   - Find "Connected apps & API tokens" section
   - Click "Manage API tokens"
   - Click "New API token"
   - Copy the generated token

2. Set environment variables:
   ```bash
   export FASTMAIL_API_TOKEN="your_api_token_here"
   # Optional: customize base URL (defaults to https://api.fastmail.com)
   # Only api.fastmail.com and www.fastmailusercontent.com are accepted by default,
   # each with an optional regional prefix (phl.api.fastmail.com,
   # phl-www.fastmailusercontent.com) as returned by JMAP session discovery.
   # For self-hosted JMAP servers, also set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true.
   export FASTMAIL_BASE_URL="https://api.fastmail.com"
   # Optional: customize attachment download directory (defaults to ~/Downloads/fastmail-mcp/).
   # download_attachment paths are confined to this directory; set it to the root
   # you want attachments saved under to write there directly in one step.
   export FASTMAIL_DOWNLOAD_DIR="/path/to/your/downloads"
   # Optional (opt-in): enable SENDING attachments. Until this is set, attaching a file
   # to an outgoing mail/draft is disabled and every attempt fails loudly — reading a
   # local file to email it out is an exfiltration vector, so it stays off by default.
   # When set, attachable files are confined to this directory (resolved independently
   # of FASTMAIL_DOWNLOAD_DIR; see the security note in "Sending attachments"). Restart
   # the server after setting it to enable the capability.
   export FASTMAIL_ATTACH_DIR="/path/to/attachable/files"
   # Optional: timezone for rendering email date fields in local time with a UTC
   # offset. Accepts an IANA name (e.g. America/New_York). Defaults to the server
   # host's timezone; set it if the server runs in a different timezone than you.
   export FASTMAIL_TIMEZONE="America/New_York"
   ```

### Running the Server

Start the MCP server:
```bash
npm start
```

For development with auto-reload:
```bash
npm run dev
```

### Run via npx (GitHub)

Default to `main` branch:

```bash
FASTMAIL_API_TOKEN="your_token" FASTMAIL_BASE_URL="https://api.fastmail.com" \
  npx --yes github:JonathanGodley/fastmail-mcp fastmail-mcp
```

Windows PowerShell:

```powershell
$env:FASTMAIL_API_TOKEN="your_token"
$env:FASTMAIL_BASE_URL="https://api.fastmail.com"
npx --yes github:JonathanGodley/fastmail-mcp fastmail-mcp
```

Pin to a tagged release:

```bash
FASTMAIL_API_TOKEN="your_token" \
  npx --yes github:JonathanGodley/fastmail-mcp@v1.9.4-fork.1 fastmail-mcp
```

## Embedding this server as a spawned child

The sections above describe registering the server with a terminal MCP client, which
keeps one long-lived connection. If instead your own program spawns this server per
task - spawn, initialize, run a few tool calls, close stdin - the following applies.

### Spawn node directly, never through a shell wrapper

Spawn the entry point with node itself:

```
node /path/to/fastmail-mcp/dist/index.js
```

Do **not** wrap it in a shell, and on Windows specifically **never** launch it via
`cmd /c node ...`. `cmd.exe` becomes an intermediary process: killing what you
spawned kills `cmd.exe` and leaves the node grandchild running, unparented and
invisible, holding its stdio pipes. The wrapper styles that are perfectly fine for
a terminal MCP client (`cmd /c`, `npx`, a `.cmd` shim, a shell one-liner) are all
wrong for an embedding caller, because an embedding caller has to be able to
terminate what it started.

For the same reason, prefer the absolute path to `dist/index.js` over `npx`: `npx`
adds its own process layer and a package-resolution step to every spawn.

### Closing stdin ends the process

Closing the child's stdin is the normal shutdown signal. On stdin EOF the server
finishes what it is doing and exits with code 0, typically within tens of
milliseconds; no kill is needed. This holds whether or not the handshake ever
completed, so a caller that gives up mid-spawn can still close stdin and walk away.

That behaviour is covered by tests (`src/server-lifecycle.test.ts`) so it cannot
regress unnoticed. Even so, a defensive caller should keep the usual belt: after
closing stdin, wait a short grace period (a second is ample) and force-kill the
child if it is still alive.

### Startup-to-ready latency

Measured on Windows 11 with Node 25, spawning `dist/index.js` directly: the
`initialize` response comes back in roughly **0.3 seconds** (15 warm runs: median
0.33s, slowest 0.43s). Nothing in startup touches the network or validates
credentials - authentication is resolved lazily on the first tool call - so the
figure is dominated by node's own startup and module loading, and does not vary
with account size or connectivity.

Set the handshake timeout to **10 seconds**. That is far above the observed range,
which leaves room for a cold module cache, an on-access virus scanner, or a loaded
machine, while still being short enough that a genuinely stuck child cannot hang
your own tool call. If `initialize` has not answered inside that window, treat the
spawn as failed and kill the child rather than waiting longer.

## Install as a Claude Desktop Extension (DXT)

You can install this server as a Desktop Extension for Claude Desktop using the packaged `.dxt` file.

1. Build and pack:
   ```bash
   npm run build
   npx @anthropic-ai/dxt pack
   ```
   This produces `fastmail-mcp.dxt` in the project root.

2. Install into Claude Desktop:
   - Open the `.dxt` file, or drag it into Claude Desktop
   - When prompted:
     - Fastmail API Token: paste your token (stored encrypted by Claude)
     - Fastmail Base URL: leave blank to use `https://api.fastmail.com` (default)

3. Use any of the tools (e.g. `get_recent_emails`).

## Response Simplification

All data-returning tools simplify responses by default to reduce token usage. Three optional parameters control how much data is returned:

- **Default** — a curated, cleaned response. Addresses are strings instead of objects, boolean flags replace keyword maps, null/empty fields are stripped, and only the most useful fields are included.
- **`verbose: true`** — all fields, still in the simplified shape. Use this when you need data the default omits (e.g. HTML body, mailbox permissions, contact addresses) without dealing with raw JMAP structures.
- **`raw: true`** — the original JMAP response with no transformation. Use this for debugging or when you need exact JMAP field names and structures.
- **`fields: [...]`** — the opposite direction: return *only* the named fields. See [Field projection](#field-projection-fields).

### What each tool returns

| Tool | `verbose` | `raw` | `fields` |
|------|-----------|-------|----------|
| `get_email` | ✅ | ✅ | ✅ |
| `list_emails`, `search_emails`, `get_recent_emails` | — | ✅ | ✅ |
| `get_thread` | — | ✅ | ✅ |
| `list_mailboxes` | ✅ | ✅ | — |
| `list_identities` | ✅ | ✅ | — |
| `list_contacts`, `get_contact`, `search_contacts` | ✅ | ✅ | — |

Email list/search tools don't support `verbose` — they always return metadata and preview. Use `get_email` for full email content, or `get_thread` with `includeBodies` for a whole conversation's plain-text bodies in one call (see [Reading long threads cheaply](#reading-long-threads-cheaply)). `preview` is a truncated snippet (~256 chars max), not the full body; these tools also return `bodyTextSize` (the full text-body size in bytes) so you can tell a short snippet apart from a long message — when `bodyTextSize` is much larger than the preview, fetch `get_email` before concluding content is absent. `size` is the whole-message size (including attachments and inline images), so it is not a body-length proxy. They do support `fields`, which is how you make a wide listing fit in one response.

### Parameter validation

Unknown or misspelled parameters are rejected with an `InvalidParams` error that lists the offending key(s) and the valid keys for that tool (e.g. passing `folder` instead of `mailbox`, or the now-renamed `mailboxId`). This is deliberate: a silently dropped parameter would let a tool run with defaults and return confident but wrong results. Note this is key-strictness only; values are still coerced leniently (e.g. a stringified `"true"`/`"20"` for a boolean/number param is accepted). The one place a *value* is checked for content rather than coerced is the message body, where a handful of malformed shapes would otherwise reach the recipient unnoticed — see [Message bodies](#message-bodies).

Error codes follow the same recoverability logic: a failure you can fix by changing input (a bad/empty field, a not-found id, a non-sendable draft) returns `InvalidParams`, while a server/operational failure you can't fix that way (a server refusal, a missing system mailbox, a transport error) returns `InternalError` — so a client knows whether to re-form the call or simply retry. A refused mutation also carries the server's stated reason. The one exception is `download_attachment`, which returns `InternalError` for a bad id to avoid leaking attachment metadata. Full rule: `docs/conventions.md`.

### Email fields

**Default fields** (all email tools): `id`, `subject`, `from`, `date`, `threadId`, `messageId`, `references`, `to`, `cc`, `bcc`, `replyTo`, `inReplyTo`, `isRead`, `isReply`, `isFlagged`, `isDraft`, `isAnswered`, `isForwarded`, `mailboxes`, `roles`, `keywords`, `preview`, `hasAttachment`, `attachments`, `listUnsubscribe`, `blobId`, `size`

**List/search tools also include**: `bodyTextSize` (byte-size hint for the text body, so a truncated `preview` isn't mistaken for the whole message; an upper bound — includes quoted history, excludes inline images)

**`get_email` also includes**: `bodyText`, `bodyHtmlSize` (character count hint — HTML omitted by default), and — on forward drafts — `forwardedMessageId` (the forwarded original's Message-ID, read from the draft's `X-Forwarded-Message-Id` header; set by `forward_email` and by Fastmail's own clients). List/search items don't carry it — they show forward-ness via `isForwarded`.

**`get_email` with `verbose`**: adds `bodyHtml` (WARNING: can produce very large responses for marketing/rich emails — only use when HTML content is specifically needed)

**With `stripQuoted`** (`get_email`, or `get_thread` with `includeBodies`): adds `quotedBytesStripped` (bytes of quoted history removed — `0` means nothing matched and the body is verbatim) **or** `quotedStripSkipped` (there was no non-empty plain-text body to strip — an HTML-only message, or one whose text part is empty). One of the two is always present, never omitted for being empty.

**`get_thread` with `includeBodies`**: adds `bodyText` per message, and `bodyTextUnavailable: true` on a message that has no plain-text part (thread reads never return HTML). Because this mode reads each message through the same property set `get_email` uses, thread messages also gain the `attachments` array (replacing the `hasAttachment` flag) and `forwardedMessageId` on a forward draft. See [Reading long threads cheaply](#reading-long-threads-cheaply).

**Simplification applied to all email output:**
- Addresses: `"Name <email>"` strings instead of `{name, email}` objects
- Flags: `isRead`, `isReply`, `isFlagged`, `isDraft`, `isAnswered`, `isForwarded` derived from JMAP keywords. `isRead` always included (unread is meaningful); `isReply`, `isFlagged`, `isDraft`, `isAnswered`, `isForwarded` omitted when false. `isAnswered` (already replied) and `isForwarded` are prime triage signals — they appear only when true
- Non-standard keywords (e.g. `$phishing`, `$important`) surfaced in a `keywords` field; standard keywords (`$seen`, `$flagged`, `$draft`, `$answered`, `$forwarded`) consumed by the boolean flags above
- **Two axes — status vs location.** The `is*` flags are the *status* axis (what the message is / what's been done to it). `mailboxes`/`roles` are the *location* axis (where it's filed). They normally agree — a draft shows `isDraft: true` and `roles: ["drafts"]` — and when they diverge both are still correct: a draft moved to Trash is `isDraft: true` with `roles: ["trash"]`. `isDraft` is *not* redundant with a `drafts` role.
- `roles`: an array of the **stable JMAP roles** of the mailboxes the message is in — lowercase `inbox`, `archive`, `sent`, `drafts`, `trash`, `junk` (`junk` is the role of the folder shown as "Spam"; there is no `spam` role). This is the rename-proof **location** signal: prefer it over `mailboxes` names for "where is this filed", because a user can rename a folder (a custom folder can even be *named* "Trash"). `roles` is a location signal, not a security/spam verdict. A custom folder has no role, so it contributes nothing to `roles` — a message filed only in custom folders has no `roles` field at all (omitted when empty).
- `mailboxes`: an array of the human-readable mailbox/label display **names** the message lives in (real mailbox `name`s, so custom labels resolve too). Multi-membership is real in Fastmail (labels are mailboxes), so the array can hold more than one name.
- **`roles` and `mailboxes` describe the SAME set of mailboxes but are NOT parallel arrays.** A custom folder appears in `mailboxes` with no `roles` entry, so the two can differ in length and position. Test membership (`roles.includes("trash")`), never `roles[0]` or `roles[i]` ↔ `mailboxes[i]`. Examples: a trashed message → `mailboxes: ["Trash"], roles: ["trash"]` (aligned); an inbox message also filed under a custom label → `mailboxes: ["Inbox", "Receipts"], roles: ["inbox"]` (divergent length — `Receipts` has no role).
- `unresolvedMailboxIds`: present only in the rare case that a mailbox id couldn't be resolved to a name (a just-created/racy folder, or a malformed/missing name). It carries the raw id(s) so the location is **never silently dropped** — its presence is **not** an error. A non-empty `unresolvedMailboxIds` means `roles`/`mailboxes` are **incomplete** for that message, so don't read "no `trash` role" as "not trashed"; true membership is `mailboxes` ∪ `unresolvedMailboxIds`.
- `raw: true` returns the underlying JMAP `keywords` and opaque `mailboxIds` (id-to-true map) instead of any of the simplified flag/location fields; map ids → names/roles via `list_mailboxes`.
- HTML-only emails (no plain text) auto-include `bodyHtml` as fallback
- `hasAttachment` omitted when false, and suppressed entirely when an `attachments` array is present (redundant)
- Attachments simplified to `{contentType, size, blobId, partId?, name?}`
- `listUnsubscribe` mapped from JMAP's `header:List-Unsubscribe:asURLs`
- `date` rendered in local time as ISO-8601 with a numeric UTC offset (e.g. `2026-03-02T08:00:00+10:00`), not UTC `Z`. The zone is the server host's by default, or `FASTMAIL_TIMEZONE` if set. Each email carries the offset for its own instant, so DST is handled per-message. Use `raw: true` to get the canonical JMAP UTC `receivedAt` instead.
- Empty and null fields omitted

### Field projection (`fields`)

`verbose`/`raw` ask for *more*. `fields` asks for **less**: pass an array of simplified field names and the response carries only those.

```jsonc
// A headers-only sweep of a two-week window, in one response instead of four
search_emails { "after": "2026-07-01", "limit": 100,
                "fields": ["id", "subject", "from", "date", "threadId"] }

// A large HTML draft's body alone - no metadata, no plain-text copy
get_email { "emailId": "M123", "fields": ["bodyHtml"] }
```

Available on `get_email`, `list_emails`, `search_emails`, `get_recent_emails` and `get_thread`. Why it exists: response size is otherwise decided by what happens to be in the mailbox rather than by anything you control — a 66-message sweep measured 84KB, of which thread `references`/`messageId`/`inReplyTo` were ~47% and `preview` another ~23%, while the five fields the caller actually wanted were 18% ([#79](https://github.com/JonathanGodley/fastmail-mcp/issues/79)); separately, an editor-inflated draft's HTML body pushed `get_email` past the same wall ([#69](https://github.com/JonathanGodley/fastmail-mcp/issues/69)). `limit` is not a substitute — per-message size varies by more than an order of magnitude, so no limit is both safe and useful.

The rules:

- **Names are the simplified field names, exactly** (see [Email fields](#email-fields) above), camelCase and case-sensitive. An unknown name is **rejected** with the full valid list rather than silently returning nothing — a typo must not read as "that field is always empty". One bad name rejects the whole call, so every typo is fixable in a single retry.
- **Omit the parameter** for the default shape. An **empty array is rejected**: "give me nothing" is never the intent, and treating it as "no projection" would hand back the full response the parameter exists to shrink.
- **`fields` cannot be combined with `raw: true`.** `raw` returns untransformed JMAP, whose field names differ (`receivedAt` not `date`, `mailboxIds` not `mailboxes`). Letting `raw` quietly win would return the largest response this server produces to a caller who asked for the smallest.
- **A field a message doesn't have is simply absent**, so a narrow projection can legitimately come back as `{}` (e.g. `fields: ["bodyText"]` on an HTML-only message — ask for `bodyHtml` too if either will do). Nothing is invented and no value's meaning changes; projection only subtracts.
- **`fields: ["bodyHtml"]` on `get_email` implies `verbose`** — you don't need to pass both. (With the HTML body included, the `bodyHtmlSize` hint is not emitted; it only appears when the body itself is omitted.)
- **Some names are valid everywhere but only populated where the tool fetches them.** The list/search tools fetch a narrower set of message properties, and the general rule is that **any field needing the full-message fetch is accepted as a name but comes back absent on a list result** — it is not rejected, because it isn't a typo. Today that is `bodyText`, `bodyHtml`, `bodyHtmlSize`, `attachments` and `forwardedMessageId`. The list tools carry populated substitutes for the common needs: `hasAttachment` for attachment presence, `isForwarded` for forward-ness, and `bodyTextSize` for how much body there is. For the content itself, fetch `get_email` — or, for a whole conversation, `get_thread` with `includeBodies`, where `bodyText` (never `bodyHtml`) is populated per message.
- **`unresolvedMailboxIds` rides along with `mailboxes`/`roles`.** If you project either location field and a mailbox id couldn't be resolved, `unresolvedMailboxIds` is included even if you didn't name it — otherwise a short `mailboxes` array would look complete when it isn't. Project neither and nothing rides along.
- **The `bodyText` signals ride along with `bodyText`** the same way: `quotedBytesStripped`/`quotedStripSkipped` (when `stripQuoted` ran) and `bodyTextUnavailable` (on thread reads) describe what happened to the body being returned, so they survive projection even unnamed — a stripped body must not read as verbatim. On `get_thread`, the 100,000-byte body cap is measured on the projected output, so projecting `bodyText` away also lifts the cap.
- **The summary line and the Trash/Spam exclusion note are not fields** and are never projected away. They describe the query, not the message — including the total match count and `nextPosition` (see [Result counts and paging](#result-counts-and-paging-position)).

Projection is applied to output only; it does not change what is fetched from the server.

### Result counts and paging (`position`)

Every `list_emails`, `search_emails` and `get_recent_emails` response opens with a summary line that states **how many results matched in total**, not just how many came back:

```
Showing 20 of 137 results. nextPosition: 20 (pass position:20 for the next page).
```

Pass that value back as `position` to read the next page, and keep going until the summary has no `nextPosition`:

```jsonc
search_emails { "after": "2026-07-01", "limit": 100 }                     // Showing 100 of 137 results. nextPosition: 100
search_emails { "after": "2026-07-01", "limit": 100, "position": 100 }    // Showing 37 of 137 results from position 100.
```

The rules:

- **The total is always stated.** A page that happens to fill the `limit` is otherwise indistinguishable from the complete answer, so a sweep asking "anything new this week?" could read a truncated page as "nothing else matched" — a false negative with no signal ([#51](https://github.com/JonathanGodley/fastmail-mcp/issues/51)). In the rare case the server declines to compute a total, the summary says the count was not returned rather than passing the page size off as the total.
- **`nextPosition` appears only while more results remain.** Its absence means the listing is complete — there is no `hasMore: false` to interpret, and no reason to re-run with a larger `position` to check. It is computed from the position actually served plus the items actually returned, so a final page shorter than `limit` ends the listing rather than advertising one more.
- **Take `nextPosition` from the response** instead of adding up `limit`s yourself. Both usually agree; the response value is the one that accounts for what the server actually did.
- **Filters apply to every page, server-side** — including the default Trash/Spam exclusion, so paging never changes what matches. The withheld-count note describes the **whole match set**, not the page: the same count repeats on every page, so don't sum the notes as you page.
- **`position: 0` is the same as omitting it.** A negative value is **rejected** (JMAP would read it as counting back from the end; to read from the oldest end pass `ascending: true`), as is a fraction or non-numeric text. A stringified `"40"` is accepted, like every other numeric parameter.
- **A position past the end is not an error.** It returns an empty page next to the real total (`Showing 0 of 137 results from position 500.`), so you can see you overshot.
- **The summary is identical on the `raw` path**, which already carried one; `raw` callers can also read JMAP's own `total`/`position` by querying directly.

`position` is how you read past the `limit` caps (`search_emails`/`list_emails` cap at 100, `get_recent_emails` at 50). Contacts and calendar listings do not take it yet: they state their total the same way, but never carry a `nextPosition`, since passing one back would be rejected as an unknown parameter.

### Mailbox fields

**Default**: `id`, `name`, `role`, `parentId`, `totalEmails`, `unreadEmails`, `totalThreads`, `unreadThreads`

**Verbose adds**: `myRights`, `sortOrder`, `isSubscribed`, `sort`, `autoLearn`, `autoPurge`, `purgeOlderThanDays`, `hidden`, `isCollapsed`, `identityRef`, `learnAsSpam`, `suppressDuplicates`, plus any other JMAP fields

Falsy `role` and `parentId` are stripped in default and verbose (use `raw` if you need `null` values).

### Identity fields

**Default**: `id`, `name`, `email`, `replyTo`, `mayDelete`, `textSignature`, `htmlSignature`

The signature fields are the identity's configured sign-off, the same text the Fastmail web UI appends for you. **Nothing appends it here**: JMAP does not do it server-side and this server does not either, so to sign a message, read it from `list_identities` and include it in the body you pass to `send_email` / `create_draft` / `reply_email` (above the quoted history on a reply). Reading it beats writing one from memory, which drifts from what the user actually configured. An unset or blank signature is omitted like any other empty field ([#33](https://github.com/JonathanGodley/fastmail-mcp/issues/33)).

**Verbose adds**: `bcc`, `verificationState`, `showInCompose`, `saveSentToMailboxId`, `displayName`, `isAutoConfigured`, `enableExternalSMTP`, `server`, `port`, `ssl`, `addBccOnSMTP`, `saveOnSMTP`, `externalCredentialId`, `warnings`, `useForAutoReply`, `verificationCheckTime`, plus any other JMAP fields

### Contact fields

**Default**: `id`, `name`, `emails`, `phones`, `organization`, `notes`

**Verbose adds**: `addresses`, `titles`, `online`, `photos`, `anniversaries`, plus any remaining JMAP fields

**Simplification applied:**
- Name resolved from `name.full` or `given + surname`
- Emails/phones flattened from JMAP's `{hash: {address}}` maps to string arrays
- Organization extracted from first entry
- Notes extracted from JMAP's `{hash: {note}}` object format
- Verbose: addresses as objects, titles as strings, online/URLs as URIs

## Available Tools (39 Total)

**🎯 Most Popular Tools:**
- **check_function_availability**: Check what's available and get setup guidance  
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
- **send_email**: Full-featured email sending with proper draft/sent handling
- **search_emails**: Free-text + structured email search (from/to/cc/bcc/subject/date/mailbox), with Trash and Spam excluded by default
- **get_recent_emails**: Quick access to recent emails from any mailbox

### Email Tools

> **Recipient format:** every recipient field (`to`/`cc`/`bcc`/`replyTo` on `send_email`, `reply_email`, `forward_email`, `create_draft`, `edit_draft`) accepts each entry as either a bare address (`a@x.com`) or the RFC 5322 `"Name <email>"` form (`Alice <a@x.com>`), which is parsed into a display name + address. The SMTP envelope always uses the bare address.
>
> **Draft sender name:** drafts created or edited via `create_draft`/`edit_draft` now carry the sending identity's display name (matching `send_email`), so the From shows your name rather than a bare address.

- **list_mailboxes**: Get all mailboxes in your account
  - Parameters: `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **list_emails**: List recent emails across all mailboxes (or one, via `mailbox`). **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam` to include them); drafts are included (set `excludeDrafts` to omit them). When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing in Trash/Spam matched, and you need not re-search to check.
  - Parameters: `mailbox` (optional — id, role, or name; scoping to a mailbox ignores the default exclusion), `limit` (default: 20, max: 100), `position` (optional offset — see [Result counts and paging](#result-counts-and-paging-position)), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
- **get_email**: Get a specific email by ID. Returns plain text body with HTML omitted (bodyHtmlSize hint provided). Only use `verbose` if you specifically need the HTML body — it can be very large for marketing emails. To read a large HTML body without the rest of the message alongside it, use `fields: ["bodyHtml"]`.
  - Parameters: `emailId` (required), `verbose` (optional, include HTML body — can be 50K+ chars for rich emails), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `stripQuoted` (optional, drop quoted reply history from `bodyText`), `raw` (optional, return original JMAP response)
  - `stripQuoted` is opt-in and reports what it did — see [Reading long threads cheaply](#reading-long-threads-cheaply).
- **send_email**: Send an email (supports threading via optional `inReplyTo` and `references` headers)
  - Parameters: `to` (required array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailbox` (optional — id/role/name to save into, defaults to Drafts), `subject` (required), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **reply_email**: Reply to an existing email with proper threading headers (automatically builds In-Reply-To and References). Saves a draft by default; set `send=true` to transmit immediately. The original is **quoted by default** (attributed, top-posted, matching the web client with a portable quote-bar style); set `quoteOriginal=false` to omit it. Quoted HTML is reproduced **sanitised** (script/style/event handlers stripped; formatting and real `http(s)` images kept; inline `cid:` images omitted — see [#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)) and is re-sent under your From address. On `send=true` the original message is marked answered and read (best-effort; the success message reports it when the mark succeeds); sending the saved draft later via `send_draft` marks it the same way, resolved from the draft's `In-Reply-To`. The subject defaults to `Re: <original subject>` (no double-prefixing); pass `subject` to override it (see [Replying with a different subject](#replying-with-a-different-subject)). (This tool returns a status string, not email data, so `raw`/simplification do not apply.)
  - Parameters: `originalEmailId` (required), `to` (optional array, defaults to the original sender — preserving their display name as `Name <email>`), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional override), `textBody` (optional), `htmlBody` (optional), `send` (optional boolean, default: false — saves a draft; set true to transmit), `quoteOriginal` (optional boolean, default: true), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **forward_email**: Forward an existing email to new recipients. Saves a draft by default; set `send=true` to transmit immediately. `to` is **required** — a forward has no default recipient. A note (`textBody`/`htmlBody`) is optional even when sending; it lands above the forwarded-message header block (`----- Original message -----` with From/To/Cc/Subject/Date — the Fastmail-native shape). The original's HTML is reproduced **sanitised** (same floor as reply quoting) and re-sent under your From address. No In-Reply-To/References are set (a forward starts a new conversation); instead the original's Message-ID is recorded in an `X-Forwarded-Message-Id` header on the draft (the same header Fastmail's own clients set), surfaced as `forwardedMessageId` by `get_email` and used by `edit_draft`'s guard. The original's regular attachments are **carried by default**, all-or-none (`includeOriginalAttachments:false` drops them; for per-part subsetting, save the default draft and use `edit_draft`'s `removeAttachments`). **Embedded inline (`cid:`) images are NOT carried** by an inline forward and their `<img>` references are stripped (see [#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)); the result message reports how many were dropped. `asAttachment:true` instead attaches the entire original as a raw `.eml` (`message/rfc822`) — lossless including inline images, but the raw message exposes its full transport headers and, for an original in Sent, any Bcc recipients (see [`docs/security-model.md`](docs/security-model.md)). Subject defaults to `Fwd: <original subject>` (no double-prefixing); pass `subject` to override (blank is treated as omitted, a non-string is rejected, same as [`reply_email`'s](#replying-with-a-different-subject)). On `send=true` the original is marked forwarded and read (best-effort; reported when it succeeds); sending the saved draft later via `send_draft` marks it the same way, resolved from the recorded `X-Forwarded-Message-Id`. (This tool returns a status string, not email data, so `raw`/simplification do not apply.)
  - Parameters: `originalEmailId` (required), `to` (required — array, or a comma-separated string), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional override), `textBody` (optional note), `htmlBody` (optional note), `send` (optional boolean, default: false — saves a draft), `includeOriginalAttachments` (optional boolean, default: true), `asAttachment` (optional boolean, default: false), `replyTo` (optional array), `attachments` (optional array of NEW files to add — see [Sending attachments](#sending-attachments))
- **create_draft**: Create an email draft without sending. Supports threading headers for replies, but **to reply to an existing message prefer `reply_email`**: hand-rolled `inReplyTo`/`references` under a mismatched subject permanently detach the draft from its conversation in the Fastmail UI (see [Replying with a different subject](#replying-with-a-different-subject)). Each call creates a new draft.
  - Parameters: `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailbox` (optional — id/role/name to save into, defaults to Drafts), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **edit_draft**: Edit an existing draft email. Since JMAP emails are immutable, this creates a replacement draft and moves the old one to Trash (so the returned email ID is new). The edit preserves the draft's threading headers (In-Reply-To/References), attachments, and other keywords. The replaced draft is **never destroyed** — it stays in Trash until Trash is emptied or auto-purged (Fastmail's Trash retention is a per-account setting, and can be set to never), so an edit made from a stale copy (the draft changed in the web UI or another client in between) can be undone. The result also reports what the replaced draft held (subject, recipients, body sizes); compare that against what you meant to replace to spot an unintended overwrite ([#65](https://github.com/JonathanGodley/fastmail-mcp/issues/65)). On the rare failure where the replacement is created but the old copy can't be moved to Trash, you are left with a duplicate draft rather than none, and the result says so. A draft containing inline (`cid:`) images, or a body part that isn't plain text or HTML, can't be preserved by editing and is **rejected** — recreate it instead (see [#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). Only fields you provide are changed; omit a field to leave it unchanged.
  - Parameters: `emailId` (required), `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `replyTo` (optional array), `clearFields` (optional array), `attachments` (optional array — **appends**), `removeAttachments` (optional array of blobId/name), `originalEmailId` (optional — for reply-quote preservation, below), `noQuote` (optional boolean)
  - **Attachments on edit.** `attachments` **appends** to the draft's existing attachments (they are kept). Remove specific ones with `removeAttachments` (pass the `blobId` from `get_email_attachments`; a unique attachment name also works — a ref matching nothing, or a name matching more than one, is rejected). Remove all with `clearFields:['attachments']`. Passing `attachments` together with `clearFields:['attachments']` is rejected as a conflict. See [Sending attachments](#sending-attachments).
  - **Empty values are rejected.** Passing an empty string or empty array for a field (e.g. `subject: ""`, `to: []`) is an error, not a silent clear — it's almost always an accidental clobber. To deliberately blank a field, name it in `clearFields`.
  - **`clearFields`**: list of field names to clear to empty/none. Allowed: `to`, `cc`, `bcc`, `replyTo`, `subject`, `textBody`, `htmlBody`, `attachments`. `from` cannot be cleared (a draft always has a sender, matching the Fastmail UI). You cannot both pass a field as a value and list it in `clearFields`. A cleared draft is still a valid draft; it just may not be sendable (e.g. with no recipients).
  - **The text body is an auto-managed fallback of the HTML.** Editing `htmlBody` alone **regenerates** `textBody` from the new HTML (so an html-alone edit discards any custom `textBody` the draft had). Editing `textBody` alone while `htmlBody` is present is **rejected** — it would not change what recipients render (the HTML), and the fallback is managed automatically; to change the message edit `htmlBody`, or supply both bodies to store a custom plain-text alternative. `clearFields:['textBody']` while `htmlBody` is present is **rejected** for the same reason; use `clearFields:['htmlBody']` to convert the draft to a plain-text email. A subject/recipient-only edit (no body written) leaves both bodies untouched. An edit that would leave the draft with no body at all is **rejected**.
  - **Reply-quote / forwarded-block preservation (`originalEmailId` / `noQuote`).** A reply draft keeps the quoted original inside its body — in the `htmlBody` for an HTML draft, or as a `> `-prefixed block in the `textBody` for a text-only one; a forward draft keeps the forwarded-message block the same way. Editing the body of a draft that **still carries** its quoted/forwarded original would silently discard it, so the edit is **rejected** unless you say what to do: pass `originalEmailId` (the id of the message the draft replies to or forwards — *not* the draft's own `emailId`) to regenerate the body and **keep** it, or `noQuote: true` to deliberately **drop** it. The check is made on the draft's *existing* body (the shape this server generated), not on your new text, so it can't be bypassed by including quote-shaped content. Two edits keep the quote/block automatically and need neither flag: a metadata-only edit (subject/recipients/attachments), and a plain-text conversion (`clearFields:['htmlBody']`, which leaves the text-side quote/block in place). Supplying `htmlBody` to a text-only reply draft converts it to HTML. The regenerated body re-quotes the named original from scratch (it never reassembles the stored quote). Each successive keep-edit must pass `originalEmailId` again.
    - **Forward drafts have a second trigger.** Besides the body markers, the draft's `X-Forwarded-Message-Id` header alone also arms the guard — so a forward draft whose block shape isn't recognizable (a foreign client's forward, or a marker lost in transit) is still challenged rather than silently stripped. The challenge names the runnable recovery: read the draft's `forwardedMessageId` via `get_email`, find that message with `search_emails` (search the bare id, without angle brackets), and pass its id as `originalEmailId`. On a forward draft, `noQuote: true` also **clears the forward marking** along with the block, so later edits aren't re-challenged. `asAttachment` forwards set no header and no block — their note edits freely, and the attached `.eml` is carried like any attachment.
    - **Cross-session keep.** In the same session you already have `originalEmailId` (you passed it to `reply_email`). For a draft saved earlier, the only thing the draft exposes is its `In-Reply-To` *Message-ID*, not the JMAP id `originalEmailId` needs — so resolve the original first with `search_emails` for that Message-ID (pass `includeTrash:true`/`includeSpam:true` so a filed-away original isn't hidden by the default exclusion), then pass its id as `originalEmailId`. (There is no shortcut that lets re-including the quote in your new body count as "keep": that would be bypassable, so the lookup is the deliberate keep path.)
    - **Limitation — drafts created by another client.** The guard recognizes the common machine-emitted shapes — for replies, `<blockquote type="cite">` (this server, Apple Mail, the Fastmail web UI) and Gmail's `class="gmail_quote"`, plus most clients' `> `-prefixed plain text; for forwards, `<div type="cite">` and the `----- Original message -----` / `---------- Forwarded message ----------` dashed lines — but not every shape (e.g. Outlook's `<div>`-based quoting). A reply draft authored in a client whose quote shape isn't recognized may have its quote dropped on a body edit **without** prompting for `originalEmailId`/`noQuote`. For forwards the exposure is narrower: any client that sets `X-Forwarded-Message-Id` (Fastmail's own clients do) is challenged via the header even when its block shape isn't recognized; only a foreign forward with neither the header nor a recognized shape is unprotected. If you're editing a draft you didn't create here, treat the quote/block as unprotected and pass `originalEmailId` explicitly to rebuild it.
- **send_draft**: Send an existing draft email. The draft must have recipients and a from address. Moves the email to the Sent folder. An **HTML-only draft with real content** (e.g. an image-only message) sends as-is — image-only/HTML-only mail is valid. Only a **genuinely empty body part** (e.g. a blank `htmlBody` alongside real text, which can happen for drafts created in other clients) is **rejected**: an empty `text/html` part renders blank and shadows a real `text/plain`, so the recipient would see nothing. Edit the draft to supply or clear that body first. (Drafts created by this server never carry an empty part; every send/draft path drops empty bodies on write.)
  - Parameters: `emailId` (required)
  - **Thread state is maintained on send** ([#60](https://github.com/JonathanGodley/fastmail-mcp/issues/60)). Because both compose tools save a draft by default, this is where most replies and forwards are actually transmitted, so it does the same marking the direct `send=true` paths do: a draft carrying `In-Reply-To` marks that original **answered and read**, and a draft carrying `X-Forwarded-Message-Id` (set by `forward_email`) marks that original **forwarded and read**. A draft with both — which this server never writes — is treated as a reply. A draft that references nothing (an ordinary compose) marks nothing.
  - The draft names its original by **Message-ID**, not by JMAP id, so the original is looked up first. When the mark succeeds the result says so; when the Message-ID matches **no** stored message, or more than one, nothing is marked and the result says which message went unmarked and why — rather than guessing at one. All of this happens **after** submission and can never fail or undo the send: the mail is already gone.
- **search_emails**: Email search. Free-text `query` matches subject/body/participants (plain words, **not** operator syntax — `from:alice@example.com` is matched literally; use the dedicated `from`/`to`/`cc`/`bcc`/`subject` params instead). All filters combine with AND. **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam`); drafts are included. `query` is optional — with no query it returns recent mail matching only the structural filters. When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing in Trash/Spam matched, and you need not re-search to check.
  - Parameters: `query` (optional), `from` (optional), `to` (optional), `cc` (optional), `bcc` (optional), `subject` (optional), `hasAttachment` (optional), `isUnread` (optional), `isPinned` (optional), `mailbox` (optional — id/role/name; scoping ignores the default exclusion), `after` (optional date or datetime), `before` (optional date or datetime), `limit` (default: 20, max: 100), `position` (optional offset — see [Result counts and paging](#result-counts-and-paging-position)), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `fields` (optional array — return only these fields; the way to keep a many-message sweep inside one response, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
  - **Date bounds (`after` / `before`).** Both accept a plain date (`2026-07-20`) or a full datetime (`2026-07-20T14:30:00Z`, or with an offset such as `2026-07-20T14:30:00+01:00`); a datetime with no zone is read as the server host's local time. A date-only value means **00:00:00 UTC on that date**, so `after:"2026-07-20"` includes all of July 20 while `before:"2026-07-20"` (the bound is exclusive) excludes it — pass `before:"2026-07-21"` to search up to and including July 20. Only those two shapes are accepted: an unpadded or slash-separated date (`2026-7-20`, `2026/07/20`), free text (`20 July 2026`), a partial date (`2026-07`), a day that doesn't exist in its month, and an empty string are all **rejected** with a message naming the parameter — omit the parameter to search without that bound. The strictness is deliberate: the loose forms are read as host-local midnight rather than UTC, which would silently move the search window. Before this, any of these reached the mail server as a bare `invalidArguments` that named no argument ([#70](https://github.com/JonathanGodley/fastmail-mcp/issues/70)).
- **get_recent_emails**: Get the most recent emails from a single mailbox (defaults to Inbox), max 50. Inbox-only with no Trash/Spam/draft flags; for an all-folder view (with the default Trash/Spam exclusion) use `list_emails`.
  - Parameters: `limit` (default: 10, max: 50), `position` (optional offset — the way to read past the 50 cap, see [Result counts and paging](#result-counts-and-paging-position)), `mailbox` (default: 'inbox' — id/role/name; e.g. `mailbox:"trash"` to read Trash directly), `ascending` (optional, oldest first), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
- **mark_email_read**: Mark an email as read or unread
  - Parameters: `emailId` (required), `read` (default: true)
- **pin_email**: Pin or unpin an email
  - Parameters: `emailId` (required), `pinned` (default: true)
- **delete_email**: Delete an email (move to trash)
  - Parameters: `emailId` (required)
- **move_email**: Move an email to a different mailbox (replaces all mailboxes)
  - Parameters: `emailId` (required), `targetMailbox` (required — id, role, or name)
- **add_labels**: Add labels (mailboxes) to an email without removing existing ones
  - Parameters: `emailId` (required), `mailboxIds` (required array - each entry an id, role, or name, resolved like every other mailbox tool; any unresolved entry rejects the whole call with the valid list)
- **remove_labels**: Remove specific labels (mailboxes) from an email
  - Parameters: `emailId` (required), `mailboxIds` (required array - each entry an id, role, or name, resolved like every other mailbox tool; any unresolved entry rejects the whole call with the valid list)

### Advanced Email Features

- **get_email_attachments**: Get list of attachments for an email
  - Parameters: `emailId` (required)
- **download_attachment**: Download an email attachment. If path is provided, saves the file to disk and returns the file path and size. Otherwise returns a download URL.
  - Parameters: `emailId` (required), `attachmentId` (required), `path` (optional)
  - `path` may be absolute or relative. Relative paths (including a bare filename) resolve against the download directory, so an attachment lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected. To save directly into your own location, set `FASTMAIL_DOWNLOAD_DIR` to that root (see [Setup](#setup)) — confinement stays on, scoped to the directory you choose.
- **get_thread**: Get all emails in a conversation thread. Returns metadata + preview for each email, or the full plain-text bodies with `includeBodies`.
  - Parameters: `threadId` (required), `includeDrafts` (optional, include in-progress drafts), `includeBodies` (optional, return each message's `bodyText`), `stripQuoted` (optional, requires `includeBodies` — return each message's new text only), `fields` (optional array — return only these fields per message, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
  - `includeBodies` turns an N-message conversation read into one call — see [Reading long threads cheaply](#reading-long-threads-cheaply).
  - Draft messages are **excluded by default** (an in-progress reply is noise when reading a conversation). When any are present, a trailing note reports **how many drafts are hidden** (so a draft reply you already started isn't missed) — the drafts themselves are not surfaced. Set `includeDrafts: true` to include them. (The note is on the simplified path only; `raw` output stays pure JSON.)

> **Draft handling is asymmetric by design.** `get_thread` excludes drafts by default while `search_emails`/`list_emails` include them: a draft reply is noise when reconstructing a conversation, but a search/list should still find everything you've written. `get_thread` does not hide them silently — its simplified output reports a hidden-draft count so you can tell a draft reply already exists (the `raw` output stays pure JSON, so a raw consumer passes `includeDrafts` itself). Drafts are identified by the `$draft` keyword (robust even if a draft is moved out of the Drafts mailbox), not by mailbox role - with one carve-out: a draft that now lives ONLY in Trash is neither shown nor counted, since it is not an active draft (`edit_draft` and `delete_email` both leave a `$draft` copy there, and counting those would inflate the warning). Use `includeDrafts` (get_thread) / `excludeDrafts` (search/list) to override either default.

### Replying with a different subject

`reply_email` takes an optional `subject`. Omit it and the reply inherits `Re: <original subject>` (no double-prefixing); pass it and your line is used verbatim, with no `Re:` added. A blank string is treated as omitted; a non-string is rejected. `forward_email`'s `subject` behaves identically over its own `Fwd:` default, so the two compose paths do not differ on the same parameter. The `In-Reply-To`/`References` headers are built the same way either way, so **recipients thread the reply correctly whatever subject you choose**.

**Why this parameter exists, and the hazard it replaces.** Fastmail groups messages into a conversation by **subject as well as by the threading headers**, which has two consequences:

- **Changing the subject on a reply detaches the draft's display grouping.** The draft is given its own `threadId`, so it sits apart from the conversation it answers in the Fastmail UI. This is display-only and it resolves once the reply is sent.
- **Hand-rolling `inReplyTo`/`references` on `create_draft` under a mismatched subject is worse, and is not undoable.** Once such a draft exists, later drafts replying to that *same original message* are grouped onto the splinter thread too, including ones created afterwards with the correct `Re:` subject and the full reference chain. Deleting the offending drafts, emptying them from Trash, and recreating with a trimmed reference list all leave the mapping in place ([#68](https://github.com/JonathanGodley/fastmail-mcp/issues/68)).

So reply with `reply_email`, with or without `subject`, rather than reconstructing a reply out of `create_draft` plus manual threading headers.

### Reading long threads cheaply

Reading a conversation the obvious way costs one `get_email` per message, and each of those messages carries the whole earlier conversation quoted below it — at successively deeper quote depths. The longer a correspondence runs, the more expensive each additional message is to read, even though the new information per message stays constant. Two opt-in flags remove both halves of that:

- **`get_thread` with `includeBodies: true`** returns every message's `bodyText` in the same call. HTML bodies are never returned here (that is where the size risk lives), so a message with no plain-text part comes back flagged `bodyTextUnavailable: true` — fetch that one with `get_email` `verbose: true`. Hidden drafts are excluded *before* bodies are read, so an in-progress reply cannot land in a transcription.
- **`stripQuoted: true`** (on `get_email`, or on `get_thread` alongside `includeBodies`) removes the quoted history and leaves the new text. `get_thread({ includeBodies: true, stripQuoted: true })` is a whole conversation as each message's own words, once, in one call.

Both are opt-in and neither changes any default: verbatim-by-default is the right behaviour for a mail tool.

**What counts as a quote.** Detection is deliberately conservative and covers the conventional, machine-emitted markers: leading `>` quote runs (including nested `>>`), an `On <date>, <someone> wrote:` attribution directly above such a run (including Gmail's wrapped two-line form), an Outlook `From:`/`Sent:`/`To:`/`Subject:` header block, and `-----Original Message-----`. Quoted *regions* are removed rather than "everything below the first marker", so a bottom-posted reply (written under the quote) and an inline reply (interleaved between quoted paragraphs) both keep their text.

**How to read the result.** The response always says what happened, so you never have to re-fetch to check:

| Field | Meaning |
|-------|---------|
| `quotedBytesStripped: N` | N bytes of quoted history were removed |
| `quotedBytesStripped: 0` | Nothing matched — this body is **verbatim**, do not re-fetch |
| `quotedStripSkipped` | Stripping did not run; the string says why (there was no plain-text body) |

**Known limits, all by design.** They run in both directions, and `quotedBytesStripped` is the tell for both: a `0` where you expected a strip, or a number that looks too large for a short message. The remedy in every case is to re-read that message without the flag.

*Too little stripped:*

- An **unrecognised quote shape passes through unchanged** rather than being guessed at, and says so with a `0`. Text derived from HTML-only quoting (Outlook's `<div>` nesting, a Gmail quote flattened without `>` prefixes) has no text-level boundary to find. This is the same foreign-client recognition residual as the compose-side quote guard; see [`docs/email-bodies.md`](docs/email-bodies.md).
- **Localized attributions** ("schrieb:", "a écrit :") aren't recognised, so that one line survives above an otherwise stripped quote.

*Too much stripped:*

- **A leading `>` is not only a quote marker.** A markdown blockquote in the sender's own writing, or pasted shell/REPL/diff output whose prompt is `>`, is stripped as quoted text. `>` runs are the primary marker the feature is built on, so this is it working as specified — telling one from the other is not decidable from the text.
- **Pasted email headers count as a quoted section** when the `From:` line carries an address, and that rule cuts to the end of the message. The address requirement exists precisely to keep display-name-only pastes (a job posting's "From: The Hiring Team", a newsletter) out of it — those now pass through untouched.
- On a **forwarded** message, the forwarded content sits below `-----Original Message-----` too, so `stripQuoted` removes it and leaves your covering note.
- A **prose line ending in "wrote:"** directly above a genuine quote can be pulled in with the attribution (the same walk-back that catches Gmail's wrapped two-line attribution).

*Other:*

- `stripQuoted` applies to **`bodyText` only**. A `bodyHtml` returned alongside it (`get_email` with `verbose`) is untouched.
- `stripQuoted` **cannot be combined with `raw`** (raw is unmodified JMAP), and on `get_thread` it requires `includeBodies`. Both combinations are rejected with an error rather than quietly ignored.
- `get_thread`'s combined bodies are capped at **100,000 bytes** (~25k tokens). Over that the call **errors**, naming the largest messages and telling you to add `stripQuoted: true` or fetch them individually — a silently truncated body is indistinguishable from a short message, which is the trap this feature exists to remove. The cap is measured on what would actually be returned, so `stripQuoted` genuinely brings an over-cap thread back under it.

### Message bodies

`send_email`, `create_draft`, `reply_email`, `forward_email`, and `edit_draft` all take `textBody` / `htmlBody`. HTML is the source of truth and the plain-text part is derived from it when you don't supply one (see [What this fork adds](#what-this-fork-adds-over-upstream)). Three malformed body shapes are **rejected** with an `InvalidParams` error rather than stored, because each one used to reach the recipient before anyone noticed:

- **A non-string body.** Passing a number, array, or object where a body string belongs now names the parameter instead of failing as an internal error ([#62](https://github.com/JonathanGodley/fastmail-mcp/issues/62)). An omitted body, or an explicit `null`, still just means "not provided".
- **An `htmlBody` that is entirely HTML-escaped** — escaped element tags (`&lt;p&gt;`) and no actual elements ([#71](https://github.com/JonathanGodley/fastmail-mcp/issues/71) / [#77](https://github.com/JonathanGodley/fastmail-mcp/issues/77)). Stored as-is it renders as literal `<p>` tags to the recipient, and the derived text part carries the same junk. Escaping is a reasonable-looking default for a client filling a markup field, so this is an easy mistake to make and an invisible one to catch. The error names the fix: pass real markup, or use `textBody`. Escaped markup **inside** real elements (`<pre>&lt;p&gt;</pre>`) is legitimate and passes, as does an ordinary `&amp;`. Because the check only fires on a body with no real markup at all, it is deliberately narrow about what counts as an escaped tag — a known element name followed by a real tag delimiter — so ordinary prose like `Hi &lt;name&gt;`, `mail me at &lt;a@b.com&gt;`, or `reply with &lt;approve&gt;` still goes through.
- **A body wrapped in a CDATA section** ([#78](https://github.com/JonathanGodley/fastmail-mcp/issues/78)). The plain-text alternative is derived by an HTML parser that recognises the section and consumes it whole, so the message is lost from it entirely; a browser, which has no CDATA in HTML, instead drops the `<![CDATA[` opener as a bogus comment and renders the trailing `]]>` as visible text. On a reply or forward it is worst of all: the quoted original supplies enough visible content that nothing downstream notices, and a text-only recipient sees only their own words quoted back.

The CDATA rule differs by format, because the damage does. In `htmlBody` a section is rejected **wherever** it appears, since a mid-body one swallows just as much as a wrapper (to show a literal CDATA token in HTML you have to escape it as `&lt;![CDATA[`, which passes). In `textBody` only a body that **starts** with `<![CDATA[` is rejected — a plain-text part is never markup-parsed, so an XML snippet quoted inside a plain-text message is real content and keeps working. A bare `]]>` with no opening token is left alone in both formats: it renders as ordinary text and survives the text derivation intact.

All of these are checked against **your** body, before any reply quote or forwarded-message block is merged in, so the reply and forward paths are covered rather than shadowed by the quote.

### Sending attachments

`send_email`, `create_draft`, `reply_email`, `forward_email`, and `edit_draft` accept an `attachments` array (on `forward_email` these are NEW files to add — the original's own attachments are carried separately, governed by `includeOriginalAttachments`/`asAttachment`). Each item is an object:

- `path` (required) — the file to attach. A bare filename or relative path resolves against the attach directory; an absolute path must fall within it.
- `name` (optional) — the filename recipients see; defaults to the file's basename.
- `contentType` (optional) — a MIME type like `application/pdf`. Inferred from the file extension when omitted. An explicit value is **echoed by Fastmail as-is** (not re-detected), so a wrong value rides out wrong.

Size caps (~25 MB/file, ~45 MB total) are a fail-fast guard; Fastmail's own limit ultimately governs.

On `edit_draft`, `attachments` **appends** (existing attachments are kept). Use `removeAttachments` (a `blobId` from `get_email_attachments`, or a unique name) to drop specific ones, or `clearFields:['attachments']` to remove all. Passing `attachments` and `clearFields:['attachments']` together is rejected as a conflict.

> **Opt-in, by design.** Attaching a file reads it off the local disk and emails it out — an exfiltration vector. So the capability is **disabled until you set `FASTMAIL_ATTACH_DIR`** (see [Setup](#setup)); until then every attach attempt fails with a self-documenting error and no file is read. Files are confined to that directory: a path outside it, a missing file, or a symlink escaping the root is rejected. The attach directory is resolved **independently** of `FASTMAIL_DOWNLOAD_DIR` — pointing both at the same directory re-opens a download-then-email round-trip, so that is your explicit choice, not a default. Restart the server after setting the variable to enable it. The confinement narrows time-of-check/time-of-use races and blocks symlink escapes, but a same-inode swap race and hardlinks inside the root are residual; see [`docs/security-model.md`](docs/security-model.md).

### Email Statistics & Analytics

- **get_mailbox_stats**: Get statistics for a mailbox (unread count, total emails, etc.)
  - Parameters: `mailbox` (optional — id, role, or name; defaults to all mailboxes)
- **get_account_summary**: Get overall account summary with statistics

### Bulk Operations

- **bulk_mark_read**: Mark multiple emails as read/unread
  - Parameters: `emailIds` (required array), `read` (default: true)
- **bulk_pin**: Pin or unpin multiple emails
  - Parameters: `emailIds` (required array), `pinned` (default: true)
- **bulk_move**: Move multiple emails to a mailbox
  - Parameters: `emailIds` (required array), `targetMailbox` (required — id, role, or name)
- **bulk_delete**: Delete multiple emails (move to trash)
  - Parameters: `emailIds` (required array)
- **bulk_add_labels**: Add labels to multiple emails simultaneously
  - Parameters: `emailIds` (required array), `mailboxIds` (required array - each entry an id, role, or name, resolved like every other mailbox tool; any unresolved entry rejects the whole call with the valid list)
- **bulk_remove_labels**: Remove labels from multiple emails simultaneously
  - Parameters: `emailIds` (required array), `mailboxIds` (required array - each entry an id, role, or name, resolved like every other mailbox tool; any unresolved entry rejects the whole call with the valid list)

### Contact Tools

- **list_contacts**: List all contacts. Returns simplified format by default.
  - Parameters: `limit` (default: 50), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **get_contact**: Get a specific contact by ID. Returns simplified format by default.
  - Parameters: `contactId` (required), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **search_contacts**: Search contacts by name or email. Returns simplified format by default.
  - Parameters: `query` (required), `limit` (default: 20), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)

### Calendar Tools

- **list_calendars**: List all calendars
- **list_calendar_events**: List calendar events (core fields only — no participants for token efficiency)
  - Parameters: `calendarId` (optional), `startDate` (optional, ISO 8601), `endDate` (optional, ISO 8601), `limit` (default: 50)
- **get_calendar_event**: Get a specific calendar event by ID. Returns organizer and participants when available.
  - Parameters: `eventId` (required)
- **create_calendar_event**: Create a new calendar event. Supports date-only (e.g. `2026-04-01`) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs `end: "2026-04-02"`.
  - Parameters: `calendarId` (required), `title` (required), `description` (optional), `start` (required, ISO 8601 or date-only), `end` (required, ISO 8601 or date-only), `location` (optional), `participants` (optional array of `{email, name?}`)
- **update_calendar_event**: Patch an existing calendar event. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty or whitespace-only string for `title`, `description`, or `location` is rejected (it won't silently blank the property). To delete `description` or `location`, list them in `clearFields`. Floating times (no Z/offset) preserve the original timezone. WARNING: providing `participants` replaces ALL existing attendee data; `participants: []` removes all attendees (and the now-orphaned ORGANIZER).
  - Parameters: `eventId` (required), `title`, `description`, `start`, `end`, `location`, `participants` (array of `{email, name?}`), `clearFields` (array of `"description"`/`"location"` to delete), `confirmRecurring` (boolean)
- **delete_calendar_event**: Delete a calendar event
  - Parameters: `eventId` (required)

#### Calendar known limitations

- **Recurring events**: Only "all events" modification is supported (master VEVENT). "This event only" or "this and future events" are not supported. Changing start/end on recurring events with exception overrides requires `confirmRecurring: true` — orphaned exceptions are pruned to prevent server errors.
- **Attendee parameters**: RSVP, ROLE, CUTYPE and other attendee parameters are parsed on read but not settable on create/update — only `email` and `name` are accepted.

### Identity & Testing Tools

- **list_identities**: List sending identities (email addresses that can be used for sending). Returns simplified format by default, including the identity's configured `textSignature`/`htmlSignature` when it has one (see [Identity fields](#identity-fields) for how to use them; nothing appends a signature for you).
  - Parameters: `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **check_function_availability**: Check which functions are available based on account permissions (includes setup guidance). Calendar tools run over CalDAV, so calendar is reported available when CalDAV credentials are configured, regardless of the JMAP calendar capability.
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
  - Parameters: `dryRun` (default: true), `limit` (default: 3)

## API Information

This server uses the JMAP (JSON Meta Application Protocol) API provided by Fastmail. JMAP is a modern, efficient alternative to IMAP for email access.

### Inspired by Fastmail JMAP-Samples

Many features in this MCP server are inspired by the official [Fastmail JMAP-Samples](https://github.com/fastmail/JMAP-Samples) repository, including:
- Recent emails retrieval (based on top-ten example)
- Email management operations
- Efficient chained JMAP method calls

### Authentication
The server uses bearer token authentication with Fastmail's API. API tokens provide secure access without exposing your main account password.

### Rate Limits
Fastmail applies rate limits to API requests. The server handles standard rate limiting, but excessive requests may be throttled.

## CalDAV Calendar Support

Fastmail does not currently expose calendar access via JMAP API tokens — the `urn:ietf:params:jmap:calendars` scope is not available because the JMAP Calendars specification is still an IETF Internet-Draft ([draft-ietf-jmap-calendars](https://datatracker.ietf.org/doc/draft-ietf-jmap-calendars/)). Fastmail has stated they will add JMAP calendar support once the spec becomes an RFC, but there is no public timeline.

However, Fastmail fully supports **CalDAV** for calendar access via `caldav.fastmail.com`. All calendar tools use CalDAV directly.

### Setup

1. Create an app-specific password on Fastmail:
   - Go to **Settings → Privacy & Security → Manage app passwords**
   - Create a new app password (you can name it "CalDAV MCP" or similar)

2. Set the following environment variables:
   ```bash
   export FASTMAIL_CALDAV_USERNAME="your-email@fastmail.com"
   export FASTMAIL_CALDAV_PASSWORD="your-app-specific-password"
   # Optional: display name for ORGANIZER when creating events with participants
   export FASTMAIL_CALDAV_DISPLAY_NAME="Your Name"
   ```

When these variables are set, all calendar tools are available. When they are not set, calendar tools will return an error with setup instructions.

## Development

### Project Structure
```
src/
├── index.ts                # Main MCP server implementation
├── auth.ts                 # Authentication handling
├── jmap-client.ts          # JMAP client wrapper
├── email-formatter.ts      # Simplified email format for AI consumption
├── quote-strip.ts          # Quoted-history detection and removal for the read path
├── thread-handler.ts       # get_thread orchestration (bodies, size cap, signals)
├── response-formatters.ts  # Mailbox/identity/contact simplifiers and query formatters
├── field-projection.ts     # `fields` output projection for the email read tools
├── contacts-calendar.ts    # Contacts and calendar extensions
└── caldav-client.ts        # CalDAV calendar client (fallback)
```

### Building
```bash
npm run build
```

### Development Mode
```bash
npm run dev
```

## License

MIT

## Contributing

Contributions are welcome! Please ensure that:
1. Code follows the existing style
2. All functions are properly typed
3. Error handling is implemented
4. Documentation is updated for new features

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Ensure your API token is valid and has the necessary permissions
2. **Missing Dependencies**: Run `npm install` to ensure all dependencies are installed  
3. **Build Errors**: Check that TypeScript compilation completes without errors using `npm run build`
4. **Calendar/Contacts "Forbidden" Errors**: Use `check_function_availability` to see setup guidance

### Calendar/Contacts Not Working?

If calendar and contacts functions return "Forbidden" errors, this is likely due to:

1. **Account Plan**: Calendar/contacts API may require business/professional Fastmail plans
2. **API Token Scope**: Your API token may need calendar/contacts permissions enabled
3. **Feature Enablement**: These features may need explicit activation in your account

**Solution**: Run `check_function_availability` for step-by-step setup guidance.

### Testing Your Setup

Use the built-in testing tools:
- **check_function_availability**: See what's available and get setup help
- **test_bulk_operations**: Safely test bulk operations without making changes

For more detailed error information, check the console output when running the server.

## Privacy & Security

- API tokens are stored encrypted by Claude Desktop when installed via the DXT and are never logged by this server.
- The server avoids logging raw errors and sensitive data (tokens, email addresses, identities, attachment names/blobIds) in error messages.
- Tool responses may include your email metadata/content by design (e.g., listing emails) but internal identifiers and credentials are not disclosed beyond what Fastmail returns for the requested data.
- If you encounter errors, messages are sanitized and summarized to prevent leaking personal information.
