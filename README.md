# Fastmail MCP Server (Unofficial Fork)

> **Not a Fastmail product.** This is a community project. It is **not affiliated with, endorsed by, or supported by Fastmail**, and nothing here comes from them. "Fastmail" is a trademark of Fastmail Pty Ltd, used only to say which APIs this server talks to. Run it at your own risk, under the terms of the [license](#license).

A fork of [MadLlama25/fastmail-mcp](https://github.com/MadLlama25/fastmail-mcp) — an MCP server for Fastmail's JMAP and CalDAV APIs.

This fork adds a **response simplification system** that reduces token usage when used with AI clients. All data-returning tools return a cleaned, curated format by default. Use `verbose` for all fields in the clean shape, or `raw` for the original JMAP response. See [Response Simplification](#response-simplification) for details.

## What this fork adds over upstream

- **Response simplification** — all data-returning tools return a token-lean shape by default (`verbose`/`raw` to opt out). See [Response Simplification](#response-simplification).
- **Draft-first compose** — every message you compose is saved as a draft first, and **`send_draft` is the only tool that transmits a composed message**: `send_email` is removed, and `reply_email`/`forward_email` only ever save drafts, so every outgoing message exists as stored, inspectable bytes before transmission ([#32](https://github.com/JonathanGodley/fastmail-mcp/issues/32), [#66](https://github.com/JonathanGodley/fastmail-mcp/issues/66)). **One path outside compose still sends mail:** naming attendees on a calendar event makes the server email them an invitation, and a later delete emails a cancellation — see [Calendar invitations send mail](#calendar-invitations-send-mail) before gating on `send_draft` alone.
- **Calendar** — attendee/participant support (which **sends invitations** — see [Calendar invitations send mail](#calendar-invitations-send-mail)), non-destructive event updates (no silent field wipes), and RFC 5545 date/TZID handling.
- **Local-time email dates** — the `date` field renders in your timezone with a UTC offset instead of raw UTC; `FASTMAIL_TIMEZONE` overrides the host zone. The same setting governs [calendar window dates](#calendar-window-dates), where `create_calendar_event` defaults a designator-less time (see [Writing calendar times](#writing-calendar-times)), and calendar sort order — one setting, applied consistently, not a recommendation that it be set at all.
- **Readable text/plain fallback** — when you supply only `htmlBody`, a plain-text alternative is generated automatically whenever one can be derived from the HTML (accessibility + deliverability). A body with no derivable text ships HTML-only; only a genuinely no-body message is rejected. See [Message bodies](#message-bodies) for what the derivation makes of an image.
- **Faithful, reversible draft edits** — `edit_draft` preserves the draft's threading headers (In-Reply-To/References), attachments, and keywords across the immutable-email recreate, instead of silently dropping them. The draft it replaces goes to **Trash rather than being destroyed**, and the result echoes back what that draft contained, so an edit made from an out-of-date copy is both visible and undoable ([#65](https://github.com/JonathanGodley/fastmail-mcp/issues/65)).
- **Compose ergonomics** — drafts carry a display name alongside the From address, preferring one the draft already carries over the sending identity's; recipient strings like `"Name <email>"` are parsed across every compose tool.
- **Outgoing attachments, including embedded images** — attach a local file, a `blobId`, or one part of an existing message to a draft, reply, forward, or edited draft (append / remove-by-ref / clear-all), and give an item a `cid` to show it *inside* the message body instead of hanging it off the end ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). Each source is opt-in: local files need `FASTMAIL_ATTACH_DIR` and are confined to it (reading a local file to email it out is an exfiltration vector), while the two in-account sources need `FASTMAIL_ALLOW_BLOB_ATTACH`. See [Sending attachments](#sending-attachments).
- **Forwarding** — `forward_email` reproduces the original under the Fastmail-native forwarded-message block and **carries its attachments** (the official server-side forward carries none), with an `asAttachment` mode that attaches the whole original as a lossless `.eml`. The edit-draft body guard extends to forward drafts.
- **Quotes and forwards keep their pictures** — a reply's quote and an inline forward's block carry the images the original displayed, instead of leaving broken references or silently blank space ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). This re-sends image data outward that you never attached and no `FASTMAIL_ATTACH_DIR` governs; `quoteOriginal=false` is the way to send none of it. See [Replying and forwarding with images](#replying-and-forwarding-with-images).
- **Attachment paths** — relative `download_attachment` paths resolve inside the configured download dir, so a bare filename lands there in one step.
- **Path-aware mailbox naming** - every mailbox parameter accepts a root-anchored path (`Archive/2026/Receipts`) alongside an id, role, or name, `list_mailboxes` returns the path it accepts, and `create_mailbox` makes the folder you want to file into ([#27](https://github.com/JonathanGodley/fastmail-mcp/issues/27), [#48](https://github.com/JonathanGodley/fastmail-mcp/issues/48)). A duplicated folder name is reported as ambiguous with its candidate paths instead of resolving to whichever one came back first. See [Naming a mailbox](#naming-a-mailbox).

## Features

### Core Email Operations
- List mailboxes (each with its full path) and get mailbox statistics; create new mailboxes
- List, search, and filter emails with advanced criteria
- Get specific emails by ID
- Compose emails (text and HTML) draft-first: every compose tool saves a draft, and `send_draft` is the only tool that transmits a composed message (calendar invitations are sent by the server — see [Calendar invitations send mail](#calendar-invitations-send-mail))
- Reply to emails with proper threading (In-Reply-To, References headers), carrying the images the original displayed into the quote
- Forward emails with attachment carry and a lossless `.eml` (`asAttachment`) mode; an inline forward carries the original's body images too
- Create, edit, and send email drafts (with or without threading)
- Email management: mark read/unread, pin/unpin, delete, move between folders, archive

### Advanced Email Features
- **Attachment Handling**: List and download email attachments; attach local files (opt-in via `FASTMAIL_ATTACH_DIR`) or content already in the account — a `blobId`, or one part of an existing message (opt-in via `FASTMAIL_ALLOW_BLOB_ATTACH`) — to outgoing mail, including images embedded in the body, and replies/forwards carry the images the original displayed
- **Threading Support**: Get complete conversation threads
- **Advanced Search**: Multi-criteria filtering (sender, date range, attachments, read status)
- **Bulk Operations**: Process multiple emails simultaneously
- **Statistics & Analytics**: Account summaries and mailbox statistics

### Contacts Operations
- List all contacts
- Get specific contacts by ID
- Search contacts by name or email
- Create, update and delete contacts, with a per-entry merge that keeps the card fields the simplified output doesn't show, and a full echo of the card as it stood before every write

### Calendar Operations
- List all calendars and calendar events
- Get specific calendar events by ID
- Create, update, and delete calendar events

### Label vs Move Operations
- **move_email/bulk_move**: Replaces ALL mailboxes for an email (folder behavior). Whatever else the message was filed under is dropped, so a message carrying three labels comes out of a move carrying one.
- **add_labels/remove_labels/bulk_add_labels/bulk_remove_labels**: Adds/removes SPECIFIC mailboxes while preserving others (label behavior). This is what you want when the message should keep what it already has. They work on **labels only**: the Inbox and the account's own user labels. A mailbox with any other JMAP role - archive, trash, junk/Spam, drafts, sent, snoozed, scheduled - is a **folder** in Fastmail's model and is rejected before anything is written, because Fastmail's own label picker does not offer it either (its "Move to" picker does). The Inbox is the one mailbox in both namespaces: removing the inbox label is exactly what archiving is, and adding it puts a message back in the Inbox. Removing a message's last remaining label archives it, matching what the Fastmail client does with a user label - a message is never left filed nowhere.
- **archive_email** is neither: it removes the Inbox membership and leaves everything else alone, adding the Archive folder only when doing that would leave the message filed nowhere. That is what Fastmail's own Archive button does, so a message filed in the Inbox plus a label keeps the label, and Archive is not added. See [Archiving does what the Fastmail client does](#archiving-does-what-the-fastmail-client-does).

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
   # For self-hosted JMAP servers, also set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true
   # (only "true" or "1" enables it). That opt-in widens the host allowlist and
   # nothing else: HTTPS is still required either way, because a plain-http base
   # URL would send the bearer token in cleartext.
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
   # Optional (opt-in, separate from the one above): allow attaching content that is
   # ALREADY in the account — an attachments item naming a blobId, or one part of an
   # existing message (emailId + attachmentId). Nothing is read off local disk on those
   # sources, so FASTMAIL_ATTACH_DIR does not govern them. Only "true" or "1" enables it;
   # anything else (including "false") leaves it off. Restart the server after setting it.
   # See the security note in "Sending attachments".
   export FASTMAIL_ALLOW_BLOB_ATTACH="true"
   # Optional: timezone for rendering email date fields in local time with a UTC
   # offset, reading calendar window dates as local days, and (on create_calendar_event)
   # the zone a designator-less start/end is written in when timeZone is omitted.
   # Accepts an IANA name (e.g. America/New_York). Defaults to the server host's
   # timezone; set it if the server runs in a different timezone than you. This is
   # an available option, not a recommendation — leave it unset and everything
   # still works, just in the host's zone. An unrecognised name is not an error -
   # it falls back to the host's timezone, with no warning if that turns out to be
   # the wrong one for you; check what actually got resolved if it matters.
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
  npx --yes github:JonathanGodley/fastmail-mcp@v1.13.4-fork.3 fastmail-mcp
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

A defensive caller should keep the usual belt: after closing stdin, wait a
short grace period (a second is ample) and force-kill the child if it is still
alive.

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
   - Fill in the settings it prompts for. These are the same settings the environment variables in [Setup](#setup) carry, so the descriptions there apply; only the API token is required.

   | Prompt | Environment equivalent | Notes |
   |---|---|---|
   | Fastmail API Token | `FASTMAIL_API_TOKEN` | Required. Stored encrypted by Claude Desktop. |
   | Fastmail Base URL | `FASTMAIL_BASE_URL` | Leave as the default `https://api.fastmail.com` unless you run your own JMAP server. |
   | Attachment Download Directory | `FASTMAIL_DOWNLOAD_DIR` | Where `download_attachment` writes. Blank uses `~/Downloads/fastmail-mcp/`. |
   | Attachment Send Directory | `FASTMAIL_ATTACH_DIR` | Blank leaves **sending** attachments off entirely. Set it only if you want it. |
   | Allow Attaching Content Already In The Account | `FASTMAIL_ALLOW_BLOB_ATTACH` | A checkbox, **unchecked** by default. Ticking it lets outgoing mail attach a `blobId` or one part of an existing message. Separate gate from the send directory — see [Sending attachments](#sending-attachments). |
   | Email Date Timezone | `FASTMAIL_TIMEZONE` | Blank uses the host's zone. Also governs calendar window dates and `create_calendar_event`'s default zone — an available option, not a requirement. If set, it must be a full IANA zone name (a region-qualifying slash, or exactly `UTC`) - the same rule `timeZone` enforces on a write, see [Writing calendar times](#writing-calendar-times). An abbreviation, alias or unresolvable value stops the server from starting rather than being silently substituted for. |
   | CalDAV Username / CalDAV App Password | `FASTMAIL_CALDAV_USERNAME` / `FASTMAIL_CALDAV_PASSWORD` | Both needed for the calendar tools; without them those tools report themselves unavailable and the rest are unaffected. |
   | Calendar Organizer Display Name | `FASTMAIL_CALDAV_DISPLAY_NAME` | Blank falls back to the CalDAV username. |

   `FASTMAIL_ALLOW_UNSAFE_BASE_URL` is deliberately not offered here — it decides where your API token may be sent, so it stays an environment variable you set on purpose.

3. Use any of the tools (e.g. `get_recent_emails`).

## Response Simplification

All data-returning tools simplify responses by default to reduce token usage. Three optional parameters control how much data is returned:

- **Default** — a curated, cleaned response. Addresses are strings instead of objects, boolean flags replace keyword maps, null/empty fields are stripped, and only the most useful fields are included.
- **`verbose: true`** — all fields, still in the simplified shape. Use this when you need data the default omits (e.g. HTML body, mailbox permissions, contact addresses) without dealing with raw JMAP structures.
- **`raw: true`** — the original JMAP response with no transformation. Use this for debugging or when you need exact JMAP field names and structures.
- **`fields: [...]`** — the opposite direction: return *only* the named fields. See [Field projection](#field-projection-fields).

### JSON payloads are compact

Every JSON payload this server returns is serialised **compact** — no indentation and no line breaks between fields — because pretty-printing is whitespace an LLM pays for by the token and reads no better for. Nothing about the data changes; only its spacing does. If you are eyeballing a payload, pipe it through your own formatter.

This is not a per-tool setting. Every handler and formatter goes through a single serialisation seam, and a guard in `src/tool-schema.test.ts` fails the build if a new call site reaches past it to `JSON.stringify` with an indent. That guard is about **spacing only** — it is not a redaction or a privacy control. See [Result serialisation in `docs/conventions.md`](docs/conventions.md) for the two seams and what the guard does and does not cover.

### What each tool returns

| Tool | `verbose` | `raw` | `fields` |
|------|-----------|-------|----------|
| `get_email` | ✅ | ✅ | ✅ |
| `list_emails`, `search_emails`, `get_recent_emails` | — | ✅ | ✅ |
| `get_thread` | — | ✅ | ✅ |
| `get_email_attachments` | — | ✅ (a set escape; see its entry) | — |
| `list_mailboxes`, `create_mailbox` | ✅ | ✅ | — |
| `list_identities` | ✅ | ✅ | — |
| `list_contacts`, `get_contact`, `search_contacts` | ✅ | ✅ | — |
| `create_contact`, `update_contact` | ✅ | ✅ | — |
| `delete_contact` | — | — | — |

`update_contact` returns `{contact, previousCard}` and `delete_contact` returns `{deleted, deletedCard}`. `verbose`/`raw` apply to `contact` only — both echoes are always the untransformed JMAP card, because they exist so a wrong write can be undone by hand and that needs every field. `delete_contact` therefore takes neither parameter: the only card it returns is the echo.

Email list/search tools don't support `verbose` — they always return metadata and preview. Use `get_email` for full email content, or `get_thread` with `includeBodies` for a whole conversation's plain-text bodies in one call (see [Reading long threads cheaply](#reading-long-threads-cheaply)). `preview` is a truncated snippet (~256 chars max), not the full body; these tools also return `bodyTextSize` (the full text-body size in bytes) so you can tell a short snippet apart from a long message — when `bodyTextSize` is much larger than the preview, fetch `get_email` before concluding content is absent. `size` is the whole-message size (including attachments and inline images), so it is not a body-length proxy. They do support `fields`, which is how you make a wide listing fit in one response.

The same loop applies to attachments. List/search results (and a default `get_thread`) fetch no attachment parts at all — they carry `hasAttachment`, which is a server heuristic that deliberately filters out small decorative images, so a message whose only picture is embedded in its body can read as `hasAttachment: false` there. `get_email` (or `get_thread` with `includeBodies`) returns the actual parts, and `get_email_attachments` returns them as raw JMAP objects.

### Parameter validation

Unknown or misspelled parameters are rejected with an `InvalidParams` error that lists the offending key(s) and the valid keys for that tool (e.g. passing `folder` instead of `mailbox`, or the now-renamed `mailboxId`). This is deliberate: a silently dropped parameter would let a tool run with defaults and return confident but wrong results. Note this is key-strictness only; values are still coerced leniently (e.g. a stringified `"true"`/`"20"` for a boolean/number param is accepted). The one place a *value* is checked for content rather than coerced is the message body, where a handful of malformed shapes would otherwise reach the recipient unnoticed — see [Message bodies](#message-bodies).

On booleans specifically:

- **A stringified `"true"`/`"false"` is accepted wherever a boolean is, like every other boolean parameter** - including `raw`, `verbose`, `read`, `pinned`, `dryRun`, `stripQuoted`, `includeBodies`, `includeDrafts`, `hasAttachment`, `isUnread`, `isPinned` and the scope flags. Every boolean also *declares* the string form in its schema (`"type": ["boolean", "string"]`), so a client that validates arguments against the schema sends it through instead of rejecting it before the server ever sees it.
- **Only those two spellings are recognised.** Anything else (`"1"`, `"yes"`, `"on"`) is not an error; it falls back to the parameter's documented default. So send the exact words rather than a near-miss, which would look accepted and do nothing.
- **List parameters take a bare string too** - a single value, a comma-separated string or a JSON-encoded array all coerce to an array. That covers `to`/`cc`/`bcc`/`replyTo`, `emailIds`/`mailboxIds`, `clearFields`, `fields`, and `inReplyTo`/`references` on `create_draft`.
- **Three parameters read their elements strictly, deliberately: `archive_email`'s `emailIds`, and `search_emails`' `requiredMailboxes` and `excludeMailboxes`.** The bare-string and comma-separated forms are still accepted, but a non-string, empty or whitespace-only *element* is rejected by index instead of being coerced. The common thread is that a dropped or mangled element there cannot be told apart from a legitimate outcome: on `archive_email` it comes back as a `notFound` entry, which that tool documents as meaning the server did not know the id, and on `search_emails` it silently widens or narrows which mailboxes the search touched. `emailIds` on `bulk_move`/`bulk_delete`, and `mailboxIds` on the label tools, keep the lenient reading. Every array parameter, lenient or strict, **trims** its elements, so `["e1", " e2"]` and `"e1, e2"` mean the same thing - an untrimmed id would otherwise reach the server and come back as a not-found, which reads as a wrong id rather than as stray whitespace.

Error codes follow the same recoverability logic: a failure you can fix by changing input (a bad/empty field, a not-found id, a non-sendable draft) returns `InvalidParams`, while a server/operational failure you can't fix that way (a permission refusal, a missing system mailbox such as Drafts/Sent/Trash, a transport error) returns `InternalError` — so a client knows whether to re-form the call or simply retry. A refused mutation also carries the server's stated reason. The one mailbox that sits on the other side is Archive: `archive_email`, `remove_labels` and `bulk_remove_labels` return `InvalidParams` when the account has no archive-role mailbox, because you can reach the same result yourself by naming any destination on `move_email`, while nothing substitutes for a missing Trash. They raise that only when a message actually **needs** Archive - a batch where every message keeps other filing never touches Archive and is served normally. When it does raise, it rejects the **whole batch** before writing anything, rather than failing that one message: the message says "Nothing was archived." (or "Nothing was changed." on the label tools) so you can tell no partial write happened. A missing *Inbox* role is `InternalError` by the same logic in reverse: nothing here substitutes for "remove this from the Inbox". There is a third whole-batch abort, also `InternalError`: a read that failed outright, or one that came back without accounting for an id in either its results or its not-found list. Those three all happen **before** anything is written, which is what lets the message say "Nothing was archived." A fourth exists and does not have that property: the `Email/set` itself failing, from a transport error or a method-level JMAP error. That throws **after** the write was dispatched, so the batch's outcome is unknown and the tell does not apply - re-read the messages rather than assume nothing changed. On `archive_email` all four are account-wide conditions rather than per-message ones, which is why they are the only cases where a tool documented as never throwing on a partial failure throws. `remove_labels`/`bulk_remove_labels` differ here: they abort the whole batch on a condition that is genuinely per-message - a message whose current filing the server did not report. (Archive being named for removal is no longer one of them: Archive is a folder, so it is rejected by the label namespace rule before any message's filing is consulted.) They return no per-message report, so "all of it" and "none of it" are the only honest answers available; serving the servable subset would leave you a bare success line and no way to learn what was skipped.

`archive_email` is also the one tool that reports a per-message failure **without** an error code at all: an unknown id comes back as a `notFound` entry in a successful response rather than as `InvalidParams`, because a batch tool that threw on one bad id would discard the outcome of every other id in the call ([#120](https://github.com/JonathanGodley/fastmail-mcp/issues/120) is where that contract is being settled for the tool family).

When the refusal comes from the server rather than from this server's own checks, the JMAP error type decides which of the two you get. A type you can act on — the id matched nothing (`notFound`), a property value was rejected (`invalidProperties`), an argument or patch was malformed (`invalidArguments`, `invalidPatch`), the object was over the size limit (`tooLarge`), the target's type forbids the operation (`singleton`) — returns `InvalidParams`, because re-forming the call is the route to success. A type describing the server, the account or its state — `forbidden`, `overQuota`, `rateLimit`, `serverFail`, `stateMismatch`, `accountReadOnly` and anything unrecognised — returns `InternalError`. A bulk operation is `InvalidParams` only when *every* failure in the batch is one of the first kind. The message text is the same either way (it always carries the server's `type` and description); only the code differs. Per-type reasoning: `docs/conventions.md`.

`download_attachment` is no exception: a bad `emailId`/`attachmentId` is `InvalidParams` there too, whether the reference is an unusable *form* (a bare `cid:` with no value, a `cid:` value matching several parts, a number with junk in it, a part carrying no blob) or a well-formed one that simply matches nothing. Either way the message names what to pass instead — a partId or blobId from `get_email_attachments`. Full rule: `docs/conventions.md`.

### Email fields

**Default fields** (all email tools): `id`, `subject`, `from`, `date`, `threadId`, `messageId`, `references`, `to`, `cc`, `bcc`, `replyTo`, `inReplyTo`, `isRead`, `isReply`, `isFlagged`, `isDraft`, `isAnswered`, `isForwarded`, `mailboxes`, `roles`, `keywords`, `preview`, `hasAttachment`, `attachments`, `listUnsubscribe`, `blobId`, `size`

`hasAttachment` and `attachments` are alternatives, never both: a tool that fetched the parts returns the `attachments` listing and drops the flag as redundant, and a tool that did not returns the flag. Empty fields are omitted, so a message with no parts at all carries **neither** — on a full-message read that means "no parts", and on a list result it means "not fetched, and `hasAttachment` was false". The two look alike; the tool you called tells them apart. Only the full-message reads (`get_email`, `get_thread` with `includeBodies`, `get_email_attachments`) fetch parts, so on list/search results `attachments` is a valid field name that is simply never populated — see the last two rules under [Field projection](#field-projection-fields).

**List/search tools also include** (and so does a default `get_thread`, which returns the same bodiless shape): `bodyTextSize` (byte-size hint for the text body, so a truncated `preview` isn't mistaken for the whole message; an upper bound — includes quoted history, excludes inline images)

**`get_email` also includes**: `bodyText`, `bodyHtmlSize` (character count hint — HTML omitted by default), and — on forward drafts — `forwardedMessageId` (the forwarded original's Message-ID, read from the draft's `X-Forwarded-Message-Id` header; set by `forward_email` and by Fastmail's own clients). On reply/forward drafts made by this server it also includes `sourceEmailId` (the JMAP id of the exact stored copy the draft was composed from, read from the draft's `X-Fastmail-MCP-Source-Id` header) — the copy `send_draft` will mark answered/forwarded, inspectable before sending. List/search items don't carry either — they show forward-ness via `isForwarded`.

**On a draft, `get_email` also issues `bodyHash`** — the token `edit_draft` requires before it will write or clear a body (see its entry below). It is computed over the body parts this read returned, so a read that did not return them all issues no hash and says so in **`bodyHashWithheld`** instead. Two of the four reasons name the read that *would* issue one — a `stripQuoted` read (the `bodyText` is not what the draft stores), and a read that omitted a part the draft carries (ask for `fields: ["bodyText", "bodyHtml", "bodyHash"]`, or `verbose: true`). The other two say to **recreate the draft**, because no read of it can issue a hash: a part the server returned truncated or with an encoding problem, and a body part no read returns at all (its declared type doesn't match the body list it sits in). Exactly one of the two fields is always there on a draft, never silence. A non-draft read carries neither, `get_thread` never issues one whatever its options, and `raw: true` returns neither (it is a derived field, and `raw` is untransformed JMAP).

**`get_email` with `verbose`**: adds `bodyHtml` (WARNING: can produce very large responses for marketing/rich emails — only use when HTML content is specifically needed)

**With `stripQuoted`** (`get_email`, or `get_thread` with `includeBodies`): adds `quotedBytesStripped` (bytes of quoted history removed — `0` means nothing matched and the body is verbatim) **or** `quotedStripSkipped` (there was no non-empty plain-text body to strip — an HTML-only message, or one whose text part is empty). One of the two is always present, never omitted for being empty.

**`get_thread` with `includeBodies`**: adds `bodyText` per message, and `bodyTextUnavailable: true` on a message that has no plain-text part (thread reads never return HTML). Because this mode reads each message through the same property set `get_email` uses, thread messages also gain the `attachments` array (replacing the `hasAttachment` flag, and carrying the same `isInline`/`cid` keys) and `forwardedMessageId`/`sourceEmailId` on this server's reply/forward drafts. Those entries are extra payload that the 100,000-byte body cap does **not** measure — it counts bodies only. Attachment bytes are never inlined, but each entry carries a sender-supplied `name` and `cid` of unbounded length, so on an attachment-heavy thread project them away with `fields`. See [Reading long threads cheaply](#reading-long-threads-cheaply).

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
- `hasAttachment` omitted when false, and suppressed entirely when the read fetched the parts (the listing is authoritative, so the heuristic flag is redundant). The two can legitimately disagree: `hasAttachment` answers "is something *attached*", so a message showing an embedded signature logo reports `false` and still lists that part
- Attachments simplified to `{contentType, size, blobId, partId?, name?, isInline?, cid?}`
- **The attachment listing covers embedded images, not just "attached" files.** For some MIME layouts the server files a body-displayed image under the message's body parts rather than its attachments, so the raw JMAP `attachments` array alone can be empty for a message that visibly shows a picture. The listing is the union of the two, in a stable order: the JMAP attachments first, then the body-routed parts (`textBody` before `htmlBody`), deduplicated
- `isInline: true` means **either** the server routed the part into the message body **or** the sender marked it `disposition: inline` — so it covers body-displayed images, and also an ordinary file the sender merely labelled inline. Omitted when false. `cid` is that part's Content-ID **verbatim**: the value a `cid:` reference in the HTML body points at, and the handle `download_attachment` takes as `cid:<value>`. A part the body actually references has one; omitted when the part has none
- **Both are sender-declared**, exactly like `name` and `contentType`. A sender chooses whether a part is marked inline and what it is called, so `isInline` is a rendering hint — never a reason to treat a part as harmless or to skip inspecting it
- `listUnsubscribe` mapped from JMAP's `header:List-Unsubscribe:asURLs`
- `date` rendered in local time as ISO-8601 with a numeric UTC offset (e.g. `2026-03-02T08:00:00+10:00`), not UTC `Z`. The zone is the server host's by default, or `FASTMAIL_TIMEZONE` if set. Each email carries the offset for its own instant, so DST is handled per-message. Use `raw: true` to get the canonical JMAP UTC `receivedAt` instead.
- Null fields and empty arrays omitted. Three fields are always present: `id`, and `subject`/`from`, which carry the placeholders `(no subject)` and `unknown` when the message has neither

### Field projection (`fields`)

`verbose`/`raw` ask for *more*. `fields` asks for **less**: pass an array of simplified field names and the response carries only those.

```jsonc
// A headers-only sweep of a two-week window, in one response instead of four
search_emails { "after": "2026-07-01", "limit": 100,
                "fields": ["id", "subject", "from", "date", "threadId"] }

// A large HTML draft's body alone - no metadata, no plain-text copy
get_email { "emailId": "M123", "fields": ["bodyHtml"] }
```

Available on `get_email`, `list_emails`, `search_emails`, `get_recent_emails` and `get_thread`. It is how you keep a response size under your own control: without it, size is decided by what happens to be in the mailbox — the thread headers and previews of a wide sweep, or one editor-inflated draft's HTML body ([#79](https://github.com/JonathanGodley/fastmail-mcp/issues/79), [#69](https://github.com/JonathanGodley/fastmail-mcp/issues/69)). `limit` is not a substitute — per-message size varies by more than an order of magnitude, so no limit is both safe and useful.

The rules:

- **Names are the simplified field names, exactly** (see [Email fields](#email-fields) above), camelCase and case-sensitive. An unknown name is **rejected** with the full valid list rather than silently returning nothing — a typo must not read as "that field is always empty". One bad name rejects the whole call, so every typo is fixable in a single retry.
- **Omit the parameter** for the default shape. An **empty array is rejected**: "give me nothing" is never the intent, and treating it as "no projection" would hand back the full response the parameter exists to shrink.
- **`fields` cannot be combined with `raw: true`.** `raw` returns untransformed JMAP, whose field names differ (`receivedAt` not `date`, `mailboxIds` not `mailboxes`).
- **A field a message doesn't have is simply absent**, so a narrow projection can legitimately come back as `{}` (e.g. `fields: ["bodyText"]` on an HTML-only message — ask for `bodyHtml` too if either will do). Nothing is invented and no value's meaning changes; projection only subtracts.
- **`fields: ["bodyHtml"]` on `get_email` implies `verbose`** — you don't need to pass both. (With the HTML body included, the `bodyHtmlSize` hint is not emitted; it only appears when the body itself is omitted.)
- **Some names are valid everywhere but only populated where the tool fetches them.** The list/search tools fetch a narrower set of message properties, and the general rule is that **any field needing the full-message fetch is accepted as a name but comes back absent on a list result** — it is not rejected, because it isn't a typo. Today that is `bodyText`, `bodyHtml`, `bodyHtmlSize`, `attachments`, `forwardedMessageId` and `sourceEmailId`. The same holds for the signals that describe what an opt-in flag did, where the list tools have no such flag: `quotedBytesStripped`/`quotedStripSkipped` (no `stripQuoted`) and `bodyTextUnavailable` (no `includeBodies`) are accepted names there and never populated. `bodyHash`/`bodyHashWithheld` are the same: valid names everywhere, issued only by a `get_email` of a draft, so a list or search of Drafts carries neither. It runs the other way once: `bodyTextSize` is a valid name on `get_email`, which returns the body itself and so emits no size hint. The list tools carry populated substitutes for the common needs: `hasAttachment` for attachment presence, `isForwarded` for forward-ness, and `bodyTextSize` for how much body there is. For the content itself, fetch `get_email` — or, for a whole conversation, `get_thread` with `includeBodies`, where `bodyText` (never `bodyHtml`) is populated per message.
- **`unresolvedMailboxIds` rides along with `mailboxes`/`roles`.** If you project either location field and a mailbox id couldn't be resolved, `unresolvedMailboxIds` is included even if you didn't name it — otherwise a short `mailboxes` array would look complete when it isn't. Project neither and nothing rides along.
- **The `bodyText` signals ride along with `bodyText`** the same way: `quotedBytesStripped`/`quotedStripSkipped` (when `stripQuoted` ran) and `bodyTextUnavailable` (on thread reads) describe what happened to the body being returned, so they survive projection even unnamed — a stripped body must not read as verbatim. On `get_thread`, the 100,000-byte body cap is measured on the projected output, so projecting `bodyText` away also lifts the cap.
- **The draft body hash rides along with either body field.** Project `bodyText` or `bodyHtml` on a `get_email` of a draft and you also get `bodyHash`, or `bodyHashWithheld` saying why there is none — unnamed, because a projection that hands back the body and drops the token `edit_draft` demands would give you the thing you asked for and quietly withhold what makes it usable. The two ride along with **each other** as well, so `fields: ["bodyHash"]` on its own is answerable: it names no body field, so no hash can be issued honestly, and what comes back is the `bodyHashWithheld` reason naming the fields to add — never `{}`.
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

`position` is how you read past the `limit` caps (`search_emails`/`list_emails` cap at 100, `get_recent_emails` at 50). Contacts and calendar listings do not take it: passing one back would be rejected as an unknown parameter. `list_contacts`/`search_contacts` state their total the same way but never carry a `nextPosition`, so a total larger than the returned count means the rest is out of reach behind the hard cap. `list_calendar_events` goes over CalDAV rather than JMAP but states its total the same way, and its total counts every matching occurrence across every calendar queried, before `limit` trimmed the list — so a total larger than the returned count means raising `limit` (hard cap 500) will reach the rest, and only a total above that cap calls for a narrower window.

### Mailbox fields

**Default**: `id`, `name`, `path`, `role`, `parentId`, `totalEmails`, `unreadEmails`, `totalThreads`, `unreadThreads`

**Verbose adds**: `myRights`, `sortOrder`, `isSubscribed`, `sort`, `autoLearn`, `autoPurge`, `purgeOlderThanDays`, `hidden`, `isCollapsed`, `identityRef`, `learnAsSpam`, `suppressDuplicates`, plus any other JMAP fields

`path` is the mailbox's **root-anchored full path**: every ancestor's name from the top-level folder down, joined with `/`, with no leading or trailing slash - `Archive/2026/Receipts`. A top-level mailbox's path is just its name. It is computed by this server, not a JMAP field, so it is absent from `raw` output. Every mailbox-taking parameter accepts a path (matched case-insensitively, segment by segment), so a path returned here can be pasted straight back into `mailbox`, `targetMailbox`, `parent`, or a `mailboxIds` entry. See [Naming a mailbox](#naming-a-mailbox) for the full resolution order.

In the rare case a mailbox's parent chain never reaches a top-level folder (a loop, or a parent the account did not return), its path cannot be computed. That mailbox is still listed, without `path`, and a trailing note names its id - the field is never dropped in silence. Refer to such a mailbox by id.

Falsy `role` and `parentId` are stripped in default and verbose (use `raw` if you need `null` values).

### Naming a mailbox

Every parameter that names a mailbox - `mailbox` on the read tools, `targetMailbox` on `move_email`/`bulk_move`, `parent` on `list_mailboxes`/`create_mailbox`, each entry of a `mailboxIds` array on the label tools, and each entry of `search_emails`' `requiredMailboxes`/`excludeMailboxes` - resolves through one matcher, in this order:

1. **Exact id**, matched case-**sensitively** (a JMAP id is an opaque server token; folding its case could make two distinct ids collide).
2. **Role** - `inbox`, `archive`, `sent`, `drafts`, `trash`, `junk` - matched case-insensitively.
3. **Exact folder name**, case-insensitively.
4. **Root-anchored path**, `/`-separated with no leading or trailing slash, each segment matched case-insensitively.

There is no substring matching at any step. A folder name matching exactly one mailbox wins over reading the same text as a path, so a folder whose own name contains a `/` stays reachable by that name. That tie rule applies only where nothing else answers to the same text: if a top-level folder is literally named `Archive/2026` **and** a real `Archive` > `2026` nesting also exists, the reference is rejected as ambiguous rather than filed into the flat folder, because both are plausible and a wrong destination is worth a retry to avoid.

Three failures are reported distinctly rather than as "not found", because they call for different corrections:

- A **name shared by several mailboxes** is rejected as ambiguous and the error lists the candidates by full path - retry with one of those, or with the id. Your spelling was not the problem.
- A **name that is also the path to a different mailbox** is rejected as ambiguous too, but a path cannot separate those two - the path is the text that failed - so the error names each candidate (which is the folder, which is the nesting) with its id, and the id is what you retry with. `list_mailboxes` returns each id next to its path.
- A **path that cannot be walked**, because some mailbox's parent chain never reaches a top-level folder, says so and points you at the id/role/name forms that still work.

Every array of mailboxes - the label tools' `mailboxIds`, and `search_emails`' `requiredMailboxes`/`excludeMailboxes` - is all-or-nothing: nothing is applied and no search runs unless every entry resolves, and a single error names every failing entry, keeping typos and ambiguous names in separate buckets so one retry fixes them all.

### Scoping across several mailboxes

`mailbox` scopes a search to one folder. `search_emails` also takes two arrays for the multi-folder cases ([#26](https://github.com/JonathanGodley/fastmail-mcp/issues/26)); each entry accepts every form above, so a path or role you learned on `mailbox` works here too. An entry that resolves to nothing rejects the whole call rather than being dropped - a dropped scope entry would silently widen the search it was passed to narrow - and the rejection names every failing entry at once, so one retry fixes them all.

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

### Identity fields

**Default**: `id`, `name`, `email`, `replyTo`, `mayDelete`, `textSignature`, `htmlSignature`

The signature fields are the identity's configured sign-off, the same text the Fastmail web UI appends for you. JMAP does not append it server-side, so signing is a per-call choice here: pass `appendSignature:true` to `create_draft` / `reply_email` / `forward_email` and this server places the identity's own signature in the body for you (above the quoted or forwarded history), or read the field from here and write it into the body yourself. Either way beats writing one from memory, which drifts from what the user actually configured. An unset or blank signature is omitted like any other empty field, and an identity with none configured appends nothing ([#33](https://github.com/JonathanGodley/fastmail-mcp/issues/33)). See [Signing a message](#signing-a-message) for what the flag does, and for `edit_draft`'s `expandSignature`, which signs only where you wrote a `{{signature}}` yourself.

**Verbose adds**: `bcc`, `verificationState`, `showInCompose`, `saveSentToMailboxId`, `displayName`, `isAutoConfigured`, `enableExternalSMTP`, `server`, `port`, `ssl`, `addBccOnSMTP`, `saveOnSMTP`, `externalCredentialId`, `warnings`, `useForAutoReply`, `verificationCheckTime`, plus any other JMAP fields

### Contact fields

**Default**: `id`, `name`, `emails`, `phones`, `organization`, `notes`, and `kind` when it is not `individual`

**Verbose adds**: `addresses`, `titles`, `online`, `photos`, `anniversaries`, plus any remaining JMAP fields (including `kind` on an ordinary card) — *and* it widens `emails`/`phones` from the hybrid shape below to the whole stored entry objects, so it adds detail to fields the default already returns, not just extra fields.

**Simplification applied:**
- Name resolved from `name.full` or `given + surname`
- Emails and phones flattened from JMAP's Id-map to an array — see the hybrid shape below
- Organization extracted from first entry
- Notes extracted from JMAP's `{hash: {note}}` object format
- `kind` omitted when it is the default `individual` - see below
- Verbose: addresses as objects, titles as strings, online/URLs as URIs

#### `kind` tells you which cards the write tools will refuse

A JSContact card carries a `kind` ([RFC 9553](https://www.rfc-editor.org/rfc/rfc9553.html) §2.1.4), and the server always emits one - a card whose stored vCard has no `KIND` line still comes back as `"individual"` (`docs/conventions.md` cites the server source for this). **The default view shows `kind` only when it is something else**:

```json
[{"id": "C1", "name": "Ada Lovelace"},
 {"id": "G1", "kind": "group", "name": "Book club"}]
```

No `kind` field means an individual. `kind: "group"` is a **contact group**, which `update_contact` and `delete_contact` both refuse (see [Writing contacts](#writing-contacts)) - so a listing shows you that before you spend a call finding out. RFC 9553 also defines `org`, `location`, `device` and `application`, and the server passes an unregistered value through as-is, so treat this as a value to read rather than a group-or-not flag: none of those kinds is a person, and `create_contact` cannot produce any of them.

`verbose: true` and `raw: true` pass `kind` through untouched, `individual` included - so under either one the presence of the field says nothing about what the card is, and you read its value. Should a card ever arrive without a `kind` at all, the default view reads that as an individual too ([#113](https://github.com/JonathanGodley/fastmail-mcp/issues/113)).

#### `emails` and `phones` are a hybrid array — handle both shapes

Each entry is **either a bare string** (the address / the number) when it carries no label, **or an object** when it does:

```json
"emails": ["plain@example.com", {"address": "work@example.com", "label": "work"}],
"phones": [{"number": "+1 555 0100", "label": "mobile"}, "+1 555 0199"]
```

A label is resolved from the entry itself, never from its map key (the keys are opaque server ids — 40-character hashes on older cards, short 6-character ids on newer ones). Two properties can carry one, and both are in live use in the same address book: a scalar `label` string wins when it is non-empty, otherwise a single-key `contexts` set (`{"private": true}`) supplies the name. An **empty** `label` — which every entry on some older imported cards carries — counts as no label at all. A `contexts` set naming more than one context resolves to no label rather than picking one.

Everything the hybrid shape folds away (`contexts`, `pref`, `@type`, …) is still reachable with `verbose: true` or `raw: true`.

### Writing contacts

`create_contact`, `update_contact` and `delete_contact` take the same entry arrays, and accept both shapes on input too: `emails` may be `["a@b.example"]` or `[{address, label}]`, `phones` may be `["+1 555 0100"]` or `[{number, label}]`. `addresses` take `{full, label?}` objects only. Any array may also arrive as a JSON string. An unknown per-item key, a key of the wrong type, or a repeated value is rejected naming its position (e.g. `emails[2]`).

**`update_contact` merges per entry rather than overwriting the card.** A JMAP patch replaces a top-level property outright, so writing `emails` from a flat array alone would silently discard the `contexts` and `pref` that sit on nearly every real entry. Instead the tool reads the card first, and:

- **The array you pass decides which entries exist.** An entry whose address/number matches one already stored keeps everything the simplified output does not show, under its existing entry id; only the properties you supply are written over it. Resending an entry exactly as you read it writes nothing at all.
- **Labels are add-and-override, not a clean rewrite.** A label that differs from the one you read is written as the entry's scalar `label` property, which is what this server resolves first — but the label you read may have come from a `contexts` set, and that set is left exactly as it was. The two can then disagree outside this server, and a Fastmail client may keep showing the old context. There is no way to *remove* a label here: a labelled entry can be relabelled, not unlabelled. Writing `contexts` the way the Fastmail apps do is not implemented.
- **A call that both drops a stored entry and adds an unknown one is rejected as ambiguous** — it reads equally as "correct this entry" and as "delete it and add an unrelated one", which produce different cards. The rejection prints the dropped entries **in full**, hidden fields included, so resending them losslessly is the cheap path. `allowEntryReplace: true` proceeds anyway, and is **scoped to the array that was actually ambiguous**: every entry of that one array is rewritten from what you supplied and **not** carrying those hidden fields over, while another array in the same call that merged cleanly still merges.
- **`name` merges into the stored structured name.** A bare string sets the full name and keeps the given/surname components; `{given}`/`{surname}` update just that part, leaving other components (titles, middle names) alone. Several real cards carry components and no full name, which is why this is not a whole-value replace.
- **`addresses` do NOT merge.** A postal entry has no matchable key, so supplying them replaces the whole set.
- **An empty array is rejected**, not read as "clear". `emails: []` is indistinguishable from a mapping bug that produced no entries, and the two outcomes differ by a whole field. Use `clearFields: ['emails']`, which can only be written on purpose. Clearable: `emails`, `phones`, `addresses`, `notes`. **`name` is not clearable** — it is how the contact is identified in every listing, so delete the card and create a new one instead. Passing a field as a value *and* in `clearFields` in the same call is rejected. An empty or whitespace-only `notes` string is rejected on the same reasoning as an empty array. `create_contact` rejects both for the same reason, except that there the fix is to omit the field — a contact being created has nothing to clear. `update_contact` also rejects a call that names no field and no `clearFields` at all.
- **A card storing more than one note is rejected** rather than collapsed. `notes` writes a single note, so a multi-note card would lose the others silently; the rejection names the count and leaves the card alone.
- **A contact group cannot be updated** by this tool. Its `members` are not editable here, and none of the tool's parameters describe a group.

**A contact group cannot be deleted here either, and that is the same rule from the other side.** `create_contact` has no `kind` and no `members` parameter, so this server cannot make a group - and a destroy path will not remove a record it could never put back, since the echo it hands you is only as good as the create surface that consumes it. `delete_contact` rejects a group card and points you at the Fastmail web interface. The rule is about the *kind* of record, not about fields the create tool cannot set: nearly every real card carries titles, organizations or photos this server cannot write, and those still delete normally.

**You can see which cards those are before you write to them.** `list_contacts`, `get_contact` and `search_contacts` show `kind: "group"` on a group card - see [`kind` tells you which cards the write tools will refuse](#kind-tells-you-which-cards-the-write-tools-will-refuse).

**Both writes echo the card as it stood before them.** `update_contact` returns `previousCard` and `delete_contact` returns `deletedCard`, the full untransformed JMAP card. That is what makes a wrong write legible: a contact delete is irreversible (a card does not go to Trash the way an email does), so the echo is the only copy left, and an update made from a stale copy is visible in the response. Neither tool has a confirmation parameter.

**The echo is not an undo.** This server writes a name, emails, phones, addresses and a note, and nothing else — so `create_contact` can rebuild that much of a deleted card, but **not** its photos, titles, organizations, nicknames, URLs, anniversaries, group membership, `uid`, or the per-entry `contexts`/`pref` that the merge exists to protect in the first place. The card you get back is a similar contact, not the one that was there. Keep the echo if there is any chance the write was wrong, and restore the rest in a Fastmail client.

`update_contact` also returns the merged card as `contact`, read back in the same request as the write. On the rare occasion the server does not return it, `contact` is absent and the response says so rather than letting the field vanish.

`deletedCard` degrades the same way, and that is the worst case of the two: `previousCard` is read before the update and the update does not proceed without it, but the delete's copy is read alongside the destroy, so on the rare occasion the server does not return it the card is gone and no copy came back. The response carries an explicit warning saying exactly that instead of returning `{deleted}` on its own. Nothing can reconstruct the card at that point.

## Available Tools (45 Total)

**🎯 Most Popular Tools:**
- **check_function_availability**: Check what's available and get setup guidance  
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
- **send_draft**: The one tool that transmits a composed message — compose with `create_draft`/`reply_email`/`forward_email`, then send here
- **search_emails**: Free-text + structured email search (from/to/cc/bcc/subject/date/mailbox, plus multi-mailbox scoping), with Trash and Spam excluded by default
- **get_recent_emails**: The small, cheap read across all mailboxes (Sent and drafts included), with Trash and Spam excluded by default

### Email Tools

> **Draft-first compose:** there is no send-in-one-call tool. `create_draft`, `reply_email`, and `forward_email` all save a draft, and **`send_draft` is the only tool that transmits a composed message** — so every outgoing message exists as stored, inspectable bytes before it is sent, and a name-based permission system can gate exactly one compose verb ([#32](https://github.com/JonathanGodley/fastmail-mcp/issues/32), [#66](https://github.com/JonathanGodley/fastmail-mcp/issues/66)). To compose and send: `create_draft` (or `reply_email`/`forward_email`) → `send_draft`. Gating `send_draft` does **not** gate all outbound mail: see [Calendar invitations send mail](#calendar-invitations-send-mail).
>
> **Recipient format:** every recipient field (`to`/`cc`/`bcc`/`replyTo` on `reply_email`, `forward_email`, `create_draft`, `edit_draft`) accepts each entry as either a bare address (`a@x.example`) or the RFC 5322 `"Name <email>"` form (`Alice <a@x.example>`), which is parsed into a display name + address. The SMTP envelope always uses the bare address.
>
> **Draft sender name:** drafts created or edited via `create_draft`/`edit_draft` carry a display name alongside the From address, so the From shows your name rather than a bare address. On an edit the name is resolved in the *opposite* order from the address: the name the stored draft already carries against the address it writes wins first, and the verified identity that owns that address is only a fallback for a draft that carries none — so a display name you set yourself is never silently reverted to your account's configured name by a later edit that didn't even touch `from` ([#152](https://github.com/JonathanGodley/fastmail-mcp/issues/152)). A stored `from` matching no verified identity (an address removed from the account, or a draft made elsewhere) keeps both its own address and its own display name (blank/whitespace-only counts as no name), rather than borrowing the default identity's name and putting it in front of a foreign address. See [`docs/email-bodies.md`](docs/email-bodies.md).

- **list_mailboxes**: Get the mailboxes in your account. Each one carries its `path` - the root-anchored, `/`-separated location - which you can paste straight back into any mailbox parameter. Pass `parent` to list one folder's direct children instead of the whole account.
  - Parameters: `parent` (optional - id, role, name, or path; lists that mailbox's direct children only), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response - untransformed JMAP carries no `path`)
- **create_mailbox**: Create a mailbox (a Fastmail folder, which is also what a label is). Returns the new mailbox in the same shape `list_mailboxes` returns, `path` included, so you can use it as a move destination or a label straight away without a second lookup. `name` is a leaf name and must not contain `/`; to nest the mailbox, pass `parent` (which itself accepts an id, role, name, or path).
  - Parameters: `name` (required - leaf name, no `/`), `parent` (optional - id, role, name, or path; omit to create at the top level), `verbose` (optional, include all fields), `raw` (optional, return original JMAP object)
- **list_emails**: List recent emails across all mailboxes (or one, via `mailbox`). **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam` to include them); drafts are included (set `excludeDrafts` to omit them). When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing filed only in Trash/Spam matched, and you need not re-search to check. One mailbox is all this tool scopes to; to intersect several mailboxes or exclude some, use `search_emails` with no query (see [Scoping across several mailboxes](#scoping-across-several-mailboxes)). `get_recent_emails` runs the same query with a smaller default and cap (10, max 50) for a quick check of what has come in.
  - Parameters: `mailbox` (optional — [id, role, name, or path](#naming-a-mailbox); scoping to a mailbox ignores the default exclusion), `limit` (default: 20, max: 100), `position` (optional offset — see [Result counts and paging](#result-counts-and-paging-position)), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
- **get_email**: Get a specific email by ID. Returns plain text body with HTML omitted (bodyHtmlSize hint provided). Only use `verbose` if you specifically need the HTML body — it can be very large for marketing emails. To read a large HTML body without the rest of the message alongside it, use `fields: ["bodyHtml"]`.
  - Parameters: `emailId` (required), `verbose` (optional, include HTML body — can be 50K+ chars for rich emails), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `stripQuoted` (optional, drop quoted reply history from `bodyText`), `raw` (optional, return original JMAP response)
  - `stripQuoted` is opt-in and reports what it did — see [Reading long threads cheaply](#reading-long-threads-cheaply).
- **draft_email**: Compose a draft in one of three modes — `new`, `reply` or `forward` — placing this server's generated blocks yourself with **body tokens**. It supersedes `create_draft`, `reply_email` and `forward_email`, which are removed in a later release. Three tokens are recognised in `textBody`/`htmlBody`: `{{signature}}` (the sending identity's configured signature), `{{quote}}` (the quoted original, `reply` only) and `{{forward}}` (the forwarded-message block, `forward` only). **Nothing is added for you** — a body with no `{{signature}}` is stored unsigned and the result says so, and a reply with no `{{quote}}` is stored without the conversation. Each mode accepts exactly its own history token and refuses the other; a near-miss spelling (`{{Quote}}`, `{{{quote}}}`) is refused rather than shipped as braces, while whitespace inside the braces is trimmed, so `{{ quote }}` is the token and a single brace each side (`{quote}`) is ordinary prose. The same token twice in one part is refused as well — expanding it twice would store the block twice. Tokens are expanded **once, at write**, so the stored draft is the finished message: `get_email` and the Fastmail web app both show what will be sent, and a body read back carries no token to re-expand. A block is substituted at the token's exact position with nothing added around it, so the spacing at the join is yours — the canonical shape is the sign-off above the history with one blank line (text) or one empty block element (HTML) between them. A token in one supplied part but not in the other supplied part is refused; supply one part and let the other be derived, or place the token in both. A body that expands to nothing visible is refused rather than stored. To write the braces literally, escape them with a backslash (`\{{signature}}`, which in JSON is `"\\{{signature}}"`); the backslash is consumed only before one of the three names. **Always saves a draft, never transmits** — review it, then send it with `send_draft`, which on `reply`/`forward` marks exactly the stored copy passed as `originalEmailId`. Quoted and forwarded HTML is reproduced **sanitised** and re-sent under your From address, and **images the original displayed are carried into the block** on the same bound as `reply_email`/`forward_email` below — a quoted image is only carried when the HTML part is non-blank **and** carries the history token, otherwise the block ships as text and the result says how many images went with it — dropped on a reply, riding as regular attachments on a forward, where the result also says the original ships HTML and that putting `{{forward}}` in `htmlBody` keeps the formatting and the images both. The result carries a **receipt**: which tokens expanded in which part, how many times, in what order they landed, which were removed for having nothing to expand to (and why), any `{{…}}` spelling in your body it left unexpanded, and whether the `asAttachment` filler body was inserted. `appendSignature` and `quoteOriginal` have no counterpart here: the token *is* the placement. (This tool returns a status string plus the receipt, not email data, so `raw`/simplification do not apply.)
  - Parameters: `mode` (**required**, one of `new`/`reply`/`forward`, no default), `originalEmailId` (required for `reply`/`forward`, refused on `new`), `to` (required on `forward`; optional on `reply`, defaulting to the original sender; optional array or a single bare string), `cc` (optional), `bcc` (optional), `from` (optional — also picks the identity whose signature `{{signature}}` expands to), `replyTo` (optional), `mailbox` (`new` only — [id, role, name, or path](#naming-a-mailbox)), `subject` (optional; `reply` inherits `Re: …` and `forward` inherits `Fwd: …`, neither double-prefixing), `textBody` (optional), `htmlBody` (optional — supplying it non-blank and carrying the history token is what lets the block show the original's images), `inReplyTo` (`new` only; array or a single bare string), `references` (`new` only; array or a single bare string), `asAttachment` (`forward` only — refuses `{{forward}}` alongside it, since the original is already carried whole; a forward with **neither** `{{forward}}` nor `asAttachment` is refused), `includeOriginalAttachments` (`forward` only, default true), `attachments` (optional).
- **reply_email**: **Superseded by `draft_email` (`mode:'reply'`)**, and removed in a later release. Reply to an existing email with proper threading headers (automatically builds In-Reply-To and References). **Always saves a draft, never transmits** — review it, then send it with `send_draft`, which marks the original answered and read on transmission — exactly the stored copy passed as `originalEmailId`, recorded on the draft (see the thread-state notes under `send_draft`). The original is **quoted by default** (attributed, top-posted, matching the web client with a portable quote-bar style); set `quoteOriginal=false` to omit it. Quoted HTML is reproduced **sanitised** (script/style/event handlers stripped; formatting and real `http(s)` images kept) and is re-sent under your From address. **Images the original displayed are carried into the quote** ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)) — so a reply re-sends image data outward that you never attached, with no `FASTMAIL_ATTACH_DIR` involved (the parts are already in the account). What is carried is bounded only by this: the body references it and the sender declared it `image/*`; the content type is sender-declared, nothing is sniffed, and there is no size or count limit. `quoteOriginal=false` omits the whole quote and is the only way to send none of it — there is no quote-text-without-images option. Carrying needs an HTML body: a text-only reply drops the images and the result says how many. See [Replying and forwarding with images](#replying-and-forwarding-with-images), and [`docs/security-model.md`](docs/security-model.md) for the full posture. The subject defaults to `Re: <original subject>` (no double-prefixing); pass `subject` to override it (see [Replying with a different subject](#replying-with-a-different-subject)). `appendSignature:true` adds the sending identity's configured signature above the quote (off by default — see [Signing a message](#signing-a-message)). (This tool returns a status string, not email data, so `raw`/simplification do not apply.)
  - Parameters: `originalEmailId` (required), `to` (optional array, defaults to the original sender — preserving their display name as `Name <email>`), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional override), `textBody` (optional), `htmlBody` (optional), `quoteOriginal` (optional boolean, default: true), `replyTo` (optional array), `appendSignature` (optional boolean, default: false — see [Signing a message](#signing-a-message)), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **forward_email**: **Superseded by `draft_email` (`mode:'forward'`)**, and removed in a later release. Forward an existing email to new recipients. **Always saves a draft, never transmits** — review it, then send it with `send_draft`, which marks the original forwarded and read on transmission — exactly the stored copy passed as `originalEmailId`, recorded on the draft on both inline and `asAttachment` forwards (see the thread-state notes under `send_draft`). `to` is **required** — a forward has no default recipient. A note (`textBody`/`htmlBody`) is optional; it lands above the forwarded-message header block (`----- Original message -----` with From/To/Cc/Subject/Date — the Fastmail-native shape). The original's HTML is reproduced **sanitised** (same floor as reply quoting) and re-sent under your From address. No In-Reply-To/References are set (a forward starts a new conversation); instead the original's Message-ID is recorded in an `X-Forwarded-Message-Id` header on the draft (the same header Fastmail's own clients set), surfaced as `forwardedMessageId` by `get_email` and droppable with `edit_draft`'s `clearFields:['forwardedMessageId']`. The original's attached **files** are carried by default, all-or-none (`includeOriginalAttachments:false` drops them). For a subset, either trim the saved draft with `edit_draft`'s `removeAttachments`, or drop them all and name just the parts you want as `attachments` items (`emailId` + `attachmentId`, which needs `FASTMAIL_ALLOW_BLOB_ATTACH` — see [Sending attachments](#sending-attachments)); the second route sends the parts with no forwarded-message block or `X-Forwarded-Message-Id` to say where they came from. **Images the original's body displayed are carried too** ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)), and are carried **even when `includeOriginalAttachments` is false** — they are body content rather than attached files, and a forward without them reproduces a message with holes in it. Same bound as `reply_email` above (referenced by the body, and sender-declared `image/*`; no size or count limit), and short of not forwarding the message there is no way to leave them behind. An image the block cannot display — a text-only forward, or a reference that did not resolve to exactly one image part — rides as a regular attachment instead (that one *is* governed by `includeOriginalAttachments`), and the result says so. See [Replying and forwarding with images](#replying-and-forwarding-with-images). `asAttachment:true` instead attaches the entire original as a raw `.eml` (`message/rfc822`) — lossless including inline images, but the raw message exposes its full transport headers and, for an original in Sent, any Bcc recipients (see [`docs/security-model.md`](docs/security-model.md)). That shape has **no forwarded-message block**: your note is the whole body, and with no note the draft ships the literal text `Forwarded message attached.` so the message reads as something rather than as an empty page with a file on it — pass your own note to replace it. Subject defaults to `Fwd: <original subject>` (no double-prefixing); pass `subject` to override (blank is treated as omitted, a non-string is rejected, same as [`reply_email`'s](#replying-with-a-different-subject)). `appendSignature:true` adds the sending identity's configured signature above the forwarded-message block — or, on an `asAttachment` forward, **below** the note, since there is no block for it to sit above (off by default — see [Signing a message](#signing-a-message)). (This tool returns a status string, not email data, so `raw`/simplification do not apply.)
  - Parameters: `originalEmailId` (required), `to` (required — array, or a comma-separated string), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional override), `textBody` (optional note), `htmlBody` (optional note), `includeOriginalAttachments` (optional boolean, default: true), `asAttachment` (optional boolean, default: false), `replyTo` (optional array), `appendSignature` (optional boolean, default: false — see [Signing a message](#signing-a-message)), `attachments` (optional array of NEW files to add — see [Sending attachments](#sending-attachments))
- **create_draft**: **Superseded by `draft_email` (`mode:'new'`)**, and removed in a later release. Create an email draft without sending (transmit it later with `send_draft`). Supports threading headers for replies, but **to reply to an existing message prefer `reply_email`**: hand-rolled `inReplyTo`/`references` under a mismatched subject permanently detach the draft from its conversation in the Fastmail UI (see [Replying with a different subject](#replying-with-a-different-subject)). Each call creates a new draft. `appendSignature:true` adds the sending identity's configured signature to the body (off by default — see [Signing a message](#signing-a-message)).
  - Parameters: `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailbox` (optional — id/role/name to save into, defaults to Drafts), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `appendSignature` (optional boolean, default: false — see [Signing a message](#signing-a-message)), `attachments` (optional array — see [Sending attachments](#sending-attachments))
  - `inReplyTo` and `references` each take a bare Message-ID string as well as an array, so a single `"<abc@example.com>"` needs no wrapping. They coerce the same way as the recipient lists, which matters here because an uncoerced string would be spread one character per header value.
- **edit_draft**: Edit an existing draft email. Since JMAP emails are immutable, this creates a replacement draft and moves the old one to Trash (so the returned email ID is new). The edit preserves the draft's threading headers (In-Reply-To/References), attachments, and other keywords. The replaced draft is **never destroyed** — it stays in Trash until Trash is emptied or auto-purged (Fastmail's Trash retention is a per-account setting, and can be set to never), so an edit made from a stale copy (the draft changed in the web UI or another client in between) can be undone. The result also reports what the replaced draft held (subject, recipients, body sizes); compare that against what you meant to replace to spot an unintended overwrite ([#65](https://github.com/JonathanGodley/fastmail-mcp/issues/65)). On the rare failure where the replacement is created but the old copy can't be moved to Trash, you are left with a duplicate draft rather than none, and the result says so. Only fields you provide are changed; omit a field to leave it unchanged. **The body you supply is stored exactly as written** — nothing is appended, removed or rebuilt in it, and the one exception is opt-in and explicit (`expandSignature: true`, below) — and any edit that writes or clears a body must pass `bodyHash` from a `get_email` read of that draft.
  - **Embedded (`cid:`) images survive an edit** ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). An image the edited body still displays keeps its identifier, so a client that has already rendered the draft doesn't renumber it — hand back the `<img src="cid:...">` references you read and the images stay. An image the body no longer displays comes **off** the draft when this server put it there and becomes a **regular attachment** when it came from elsewhere — those bytes aren't this server's to discard. The result says what the draft ended up embedding, and what came off it. A `cid:` reference to one of this server's own identifiers naming **no part the draft carries** is rejected: passing one back is how images keep displaying, but authoring a new one points at nothing that can ever exist. Two body shapes still can't be rebuilt faithfully and are **rejected**: a body part that is neither text nor a carriable image, audio, video or attached message; and a body that interleaves two parts of the same text type (the Apple Mail text-image-text layout, [#85](https://github.com/JonathanGodley/fastmail-mcp/issues/85)) — recreate those drafts instead. And if the draft's stored body already references an image that isn't attached, editing its **body** is rejected until the edit resolves that — replace the body, or add an `attachments` item supplying the missing `cid`; metadata and attachment edits still work on such a draft.
  - Parameters: `emailId` (required), `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `replyTo` (optional array), `bodyHash` (**required for any edit that writes or clears a body**, below), `expandSignature` (optional boolean, default false — expands a `{{signature}}` you wrote, below), `clearFields` (optional array), `attachments` (optional array — **appends**), `removeAttachments` (optional array of blobId/name)
  - **Attachments on edit.** `attachments` **appends** to the draft's existing attachments (they are kept). Remove specific ones with `removeAttachments` (pass the `blobId` from `get_email_attachments`; a unique attachment name also works — a ref matching nothing, or a name matching more than one, is rejected). Remove all with `clearFields:['attachments']`. Passing `attachments` together with `clearFields:['attachments']` is rejected as a conflict. Both removal routes reach **everything `get_email_attachments` lists**, embedded body images included, and both are rejected when the surviving body would still reference what they took away — so drop the `<img>` from the body you supply in the same call. See [Sending attachments](#sending-attachments).
  - **Empty values are rejected.** Passing an empty string or empty array for a field (e.g. `subject: ""`, `to: []`) is an error, not a silent clear — it's almost always an accidental clobber. To deliberately blank a field, name it in `clearFields`.
  - **`clearFields`**: list of field names to clear to empty/none. Allowed: `to`, `cc`, `bcc`, `replyTo`, `subject`, `textBody`, `htmlBody`, `attachments`, `forwardedMessageId`. `from` cannot be cleared (a draft always has a sender, matching the Fastmail UI). You cannot both pass a field as a value and list it in `clearFields`. Clearing `textBody` or `htmlBody` requires `bodyHash`. `clearFields:['forwardedMessageId']` **de-forwards** the draft: it drops the recorded `X-Forwarded-Message-Id`, so `send_draft` will not mark the original forwarded. On a forward draft it drops the recorded `sourceEmailId` too — that pointer names the exact instance the marking is about, so it goes with it; on a *reply* draft it is kept, since the draft is still a reply to that instance. Clearing is metadata, so it works on a body edit and a metadata-only edit alike — and it does not touch the body, so a forwarded-message block already in there stays until you replace the body yourself. **The converse also holds, and it is the easier one to get wrong:** deleting the forwarded-message block from the body by hand does *not* de-forward the draft. The marking lives in the header, so `send_draft` still marks the original forwarded until you name the field here. A cleared draft is still a valid draft; it just may not be sendable (e.g. with no recipients).
  - **The text body is an auto-managed fallback of the HTML.** Editing `htmlBody` alone **regenerates** `textBody` from the new HTML (so an html-alone edit discards any custom `textBody` the draft had — the result says so). Editing `textBody` alone while `htmlBody` is present is **rejected** — it would not change what recipients render (the HTML), and the fallback is managed automatically; to change the message edit `htmlBody`, or supply both bodies to store a custom plain-text alternative. `clearFields:['textBody']` while `htmlBody` is present is **rejected** for the same reason; use `clearFields:['htmlBody']` to convert the draft to a plain-text email. A subject/recipient-only edit (no body written) leaves both bodies untouched. An edit that would leave the draft with no body at all is **rejected**.
  - **The body is stored exactly as written.** Whatever you hand back is what the draft holds, character for character. This server appends nothing, removes nothing, and recognises nothing in it. So a reply's quoted original or a forward's forwarded-message block survives an edit **because you handed it back** — read the draft, change the words, send the whole body — and a body sent without it drops it, with no challenge and no warning. That is the trade the design makes on purpose: the old behaviour tried to detect a quote in the stored body and rebuild it from the original message, which meant guessing at foreign clients' markup, refusing edits it could not classify, and appending a freshly-quoted copy that could not merge with the one already in the body you sent.
  - **Proving your read: `bodyHash`.** Because an edit replaces the body wholesale, one written from a stale read silently overwrites whatever changed in between. So every edit that writes `textBody`/`htmlBody` or clears one of them must pass `bodyHash` — the value `get_email` returns for that draft. A missing or stale hash is **rejected**; re-read the draft and edit from that. It is a lost-update guard and nothing more: it proves you *saw* the body, never that you kept any of it, so it passes on an edit that replaces the body entirely. A metadata-only edit is body-invariant and needs none. A successful body edit hands the **next** hash back in its result, read off the saved draft, so a run of edits costs one `get_email` at the start rather than one between each pair. When a hash cannot be issued honestly the result says which case it hit and you re-read: the expansion rewrote the body, the surviving body is one this server derived from HTML rather than one you supplied (clearing `htmlBody` on a draft that never had one leaves your own stored text standing, so that edit *does* get a hash), or the read-back failed. There is a fourth group, and it is the one that does not end in a re-read. The saved draft is judged by exactly the rule `get_email` applies to any draft — one rule, so the two tools never disagree about the same stored body — and where that rule says no read can issue a hash, neither can the edit: a part the server flagged as truncated or with an encoding problem, or a body part no read returns at all (its declared type doesn't match the body list it sits in). Those notes say to **recreate** the draft. A returned hash always comes from re-reading the saved message — never from the bytes the call sent.
  - **Expanding a signature: `expandSignature`.** Off by default, and off is what makes the byte-for-byte promise hold: nothing this server finds in a body can make it rewrite one, which matters because part of a draft body you hand back was written by someone else (a quoted original, a forwarded message) and could carry a planted placeholder. Pass `expandSignature: true` and a `{{signature}}` you wrote is replaced with the sending identity's configured sign-off; see [Signing a message](#signing-a-message). `{{quote}}` and `{{forward}}` never expand here and are stored as the literal text you typed — this tool never places history for you. Passing the flag with no `{{signature}}` in the body is rejected, as is the same token twice in one part. A near-miss spelling (`{{Signature}}`, `{{{signature}}}`) is stored as written and **reported** rather than refused, since it may be text the original's author wrote — as is a `{{signature}}` you store on an *unflagged* edit, which is how you keep one as literal text. To write the braces as text *under the flag*, escape them: `\{{signature}}` (in JSON, `"\\{{signature}}"`). The backslash is consumed and the bare token is stored, so it will not expand again unless a later edit passes the flag too. On an unflagged edit there is nothing to escape — that body is stored byte for byte, backslash included — and the result says so when *this* edit added the escape, because a backslash you meant as an escape has just shipped as content. It follows the same count-rise rule as the token note below: an escape the body you handed back already carried passes in **silence**, or it would be reported to you on every edit of that draft for as long as it exists.
- **send_draft**: Send an existing draft email — **the only tool that transmits a composed message** (see the draft-first note above; calendar invitations are sent by the server, outside compose — see [Calendar invitations send mail](#calendar-invitations-send-mail)). The draft must have recipients and a from address. Moves the email to the Sent folder. An **HTML-only draft with real content** sends as-is — HTML-only mail is valid. That is a draft whose HTML yields no derivable text at all, e.g. one showing a remote image with no alt text; a draft displaying an embedded (`cid:`) image is not in that group, since it derives `[image]` and carries a text part. Only a **genuinely empty body part** (e.g. a blank `htmlBody` alongside real text, which can happen for drafts created in other clients) is **rejected**: an empty `text/html` part renders blank and shadows a real `text/plain`, so the recipient would see nothing. Edit the draft to supply or clear that body first. (Drafts created by this server never carry an empty part; every send/draft path drops empty bodies on write.)
  - Parameters: `emailId` (required)
  - **The result reports what went out.** It says how many embedded (`cid:`) images the transmitted message carried, read off the draft as submitted. That is a **receipt, not a check** — `send_draft` never refuses a draft over its images ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)).
  - **Thread state is maintained on send** ([#60](https://github.com/JonathanGodley/fastmail-mcp/issues/60)). Because every reply and forward is transmitted here, this is where the marking happens: a draft carrying `In-Reply-To` marks that original **answered and read**, and a draft carrying `X-Forwarded-Message-Id` (set by `forward_email`, on both inline and `asAttachment` forwards) marks that original **forwarded and read**. A draft with both — which this server never writes — is treated as a reply. A draft that records neither header marks nothing (an ordinary compose).
  - **Which copy gets marked.** A Message-ID names a *message*, but an account can hold several stored copies of one (a duplicate delivery, or a self-addressed message filed into another folder). Drafts made by `reply_email`/`forward_email` record the JMAP id of the exact copy they were composed from (an `X-Fastmail-MCP-Source-Id` header on the draft, surfaced as `sourceEmailId` by `get_email` so the copy to be marked is inspectable before sending), and that copy is what gets marked — validated first: the recorded instance must still exist and still carry the draft's source Message-ID, otherwise it falls back. The fallback (drafts without the record — older drafts, foreign clients) resolves the Message-ID by lookup: when it matches exactly one stored message, that one is marked; **no** match, or more than one, marks nothing and the result says which message went unmarked and why — rather than guessing at one. This matches Fastmail's own client, which also marks only the instance replied to. All of this happens **after** submission and can never fail or undo the send: the mail is already gone. (The recorded header is transmitted with the message; it is an opaque account-scoped id, disclosing less than the `In-Reply-To` beside it — see [`docs/security-model.md`](docs/security-model.md).)
- **search_emails**: Email search. Free-text `query` matches subject/body/participants (plain words, **not** operator syntax — `from:alice@example.com` is matched literally; use the dedicated `from`/`to`/`cc`/`bcc`/`subject` params instead). All filters combine with AND. **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam`); drafts are included. `query` is optional — with no query it returns recent mail matching only the structural filters. When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing filed only in Trash/Spam matched, and you need not re-search to check.
  - Parameters: `query` (optional), `from` (optional), `to` (optional), `cc` (optional), `bcc` (optional), `subject` (optional), `hasAttachment` (optional), `isUnread` (optional), `isPinned` (optional), `mailbox` (optional — id/role/name/path; scoping ignores the default exclusion), `requiredMailboxes` (optional array — see [Scoping across several mailboxes](#scoping-across-several-mailboxes)), `excludeMailboxes` (optional array — same section), `after` (optional date or datetime), `before` (optional date or datetime), `limit` (default: 20, max: 100), `position` (optional offset — see [Result counts and paging](#result-counts-and-paging-position)), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `fields` (optional array — return only these fields; the way to keep a many-message sweep inside one response, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
  - **`requiredMailboxes` / `excludeMailboxes`** scope across several mailboxes at once, and are the only place this server does — see [Scoping across several mailboxes](#scoping-across-several-mailboxes) for what exclusion actually means here (it is weaker than the name suggests).
  - **`hasAttachment` filters on the server's own heuristic**, compared directly (RFC 8621 §4.4.1) — it is the same value the results report, not a count of parts. It answers "is there content attached", so a message whose only image is a small embedded one is filtered *out* by `hasAttachment: true`. There is no server-side filter for embedded images: narrow with the other filters, then read the parts with `get_email`.
  - **Date bounds (`after` / `before`).** Both accept a plain date (`2026-07-20`) or a full datetime (`2026-07-20T14:30:00Z`, or with an offset such as `2026-07-20T14:30:00+01:00`); a datetime with no zone is read as the server host's local time. A date-only value means **00:00:00 UTC on that date**, so `after:"2026-07-20"` includes all of July 20 while `before:"2026-07-20"` (the bound is exclusive) excludes it — pass `before:"2026-07-21"` to search up to and including July 20. Only those two shapes are accepted: an unpadded or slash-separated date (`2026-7-20`, `2026/07/20`), free text (`20 July 2026`), a partial date (`2026-07`), a day that doesn't exist in its month, and an empty string are all **rejected** with a message naming the parameter — omit the parameter to search without that bound. The strictness is deliberate: the loose forms are read as host-local midnight rather than UTC, which would silently move the search window ([#70](https://github.com/JonathanGodley/fastmail-mcp/issues/70)).
- **get_recent_emails**: Get the newest emails across all mailboxes (or one, via `mailbox`) — the small, cheap read. **It is not an arrivals feed**: "all mailboxes" is every folder except Trash and Spam, so your own Sent copies, drafts and any custom folder come back alongside newly received mail. Pass `mailbox:"inbox"` for delivered mail only, and `excludeDrafts` to drop drafts. **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam` to include them). When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing filed only in Trash/Spam matched, and you need not re-run to check. It filters exactly as `list_emails` does; what separates them is size, not scope — this returns 10 by default and caps at 50, `list_emails` returns 20 and caps at 100, so use `list_emails` when you are working through a folder rather than taking a quick look ([#29](https://github.com/JonathanGodley/fastmail-mcp/issues/29)).
  - Parameters: `limit` (default: 10, max: 50), `position` (optional offset — the way to read past the 50 cap, see [Result counts and paging](#result-counts-and-paging-position)), `mailbox` (optional — [id, role, name, or path](#naming-a-mailbox); scoping to a mailbox ignores the default exclusion, e.g. `mailbox:"trash"` to read Trash directly), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `fields` (optional array — return only these fields, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
- **mark_email_read**: Mark an email as read or unread
  - Parameters: `emailId` (required), `read` (default: true)
- **pin_email**: Pin or unpin an email
  - Parameters: `emailId` (required), `pinned` (default: true)
- **delete_email**: Delete an email (move to trash)
  - Parameters: `emailId` (required)
- **move_email**: Move an email to a different mailbox. **Replaces the message's entire mailbox membership** - every other label or folder it was filed under is removed. To file it somewhere while keeping those, use `add_labels` - which takes the Inbox and the account's own labels only, so a folder is reachable by moving alone (archiving is the exception: `archive_email` keeps the rest of the filing). No keyword is changed, so the message keeps its read/unread and flagged state.
  - Parameters: `emailId` (required), `targetMailbox` (required — [id, role, name, or path](#naming-a-mailbox))
- **archive_email**: Archive one or more emails the way the Fastmail client does - **remove the Inbox membership and leave everything else in place**, adding the Archive folder only when doing that would leave the message filed nowhere. It handles a batch, so there is no `bulk_archive`. It writes **no keyword**, which is not the same as leaving the read state untouched - see the caveat in the section linked below. Once a message has **left the Inbox** it **refuses** one in Trash, Spam, Drafts, Scheduled, Sent or Snoozed, because the client offers no Archive action there either. The Inbox is tested first, so a message in the Inbox *and* one of those is archived rather than refused, keeping that membership - except that a message also in Scheduled may come back as `failed`, because the server appears to reject re-asserting a scheduled membership outside a send request ([#130](https://github.com/JonathanGodley/fastmail-mcp/issues/130)). The destination is found by JMAP `archive` role, never by folder name, so a folder merely *named* "archive" is not it - a name-resolved destination would let text the model merely read decide where mail lands. For any other destination use `move_email`. Full behaviour, including the six per-message outcomes it reports: [Archiving does what the Fastmail client does](#archiving-does-what-the-fastmail-client-does) ([#21](https://github.com/JonathanGodley/fastmail-mcp/issues/21), [#104](https://github.com/JonathanGodley/fastmail-mcp/issues/104)).
  - Parameters: `emailIds` (required array of **message** ids, not thread ids — pass an array even for one message; duplicates are collapsed; to archive a conversation pass every message id in it). **This replaced the old single `emailId` parameter**, so a call written against an earlier version is rejected with the valid key list.
- **add_labels**: Add labels (mailboxes) to an email without removing existing ones. Labels means the **Inbox and the account's own labels**; a folder (any other role mailbox) is rejected - use `move_email`. Adding the inbox label is how a message is put back in the Inbox
  - Parameters: `emailId` (required), `mailboxIds` (required array - each entry [id, role, name, or path](#naming-a-mailbox), resolved like every other mailbox tool; any unresolved or ambiguous entry rejects the whole call with the valid list, as does any entry that resolves to a folder rather than a label)
- **remove_labels**: Remove specific labels (mailboxes) from an email. Labels means the **Inbox and the account's own labels**; a folder (any other role mailbox, Archive included) is rejected before anything is written - use `move_email`. Removing the inbox label is what archiving does. Removing the message's last label archives it instead of leaving it filed nowhere
  - Parameters: `emailId` (required), `mailboxIds` (required array - each entry [id, role, name, or path](#naming-a-mailbox), resolved like every other mailbox tool; any unresolved or ambiguous entry rejects the whole call with the valid list, as does any entry that resolves to a folder rather than a label)

### Advanced Email Features

- **get_email_attachments**: List an email's parts as **raw JMAP part objects** (`partId`, `blobId`, `type`, `size`, `name`, `disposition`, `cid`) rather than the simplified shape the read tools return. Like `get_email`, the listing includes images embedded in the message body, not only "attached" files. Nothing in this raw listing distinguishes the two — a body-embedded part usually reports `disposition: null` — so cross-check `get_email`'s derived `isInline` before acting on an entry, for instance before handing its `blobId` to `edit_draft`'s `removeAttachments`, which would strip an image the body still displays. This listing is where an `attachmentId` comes from: pass an entry's `partId` or `blobId` back to `download_attachment`, or to an `attachments` item that attaches that part to a new message (`emailId` + `attachmentId`, which needs `FASTMAIL_ALLOW_BLOB_ATTACH` — see [Sending attachments](#sending-attachments)). This listing is also what `download_attachment`'s entry numbers count from: its first entry is `attachmentId: "0"`.
  - Parameters: `emailId` (required), `raw` (optional)
  - **`raw` here is a SET escape, not a shape escape.** The entries are already raw JMAP objects, so `raw: true` changes only *which* parts are listed: it returns the JMAP `attachments` array alone, dropping the body-routed parts. How many were withheld is reported as a **separate second message**, never appended to the JSON, which stays parseable. `download_attachment` entry numbers always count from the full listing, so a `raw` listing is not an index basis.
- **download_attachment**: Download an email attachment. If path is provided, saves the file to disk and returns the file path and size. Otherwise returns a download URL.
  - Parameters: `emailId` (required), `attachmentId` (required), `path` (optional)
  - **`attachmentId` accepts four forms**, resolved in this fixed order: a `partId` from `get_email_attachments`; a `blobId`; `cid:<value>` for an embedded image, using the `cid` from `get_email`; or a plain entry number (`0`, `1`, `2`, …) counting from the start of the `get_email_attachments` listing. Digits resolve as a **partId first** — Fastmail partIds are themselves digit strings — so the entry-number form applies only when no part claims that value. The `cid:` prefix is required for the cid form (a bare Content-ID is not one), only the first prefix is stripped (`cid:cid:x` looks up the Content-ID `cid:x`), the value is matched literally before any percent-decoding is tried, and a value matching more than one part is **rejected** rather than guessed at. A number with anything else in it (`3a`, `-1`, `1.5`) is rejected rather than silently read as an entry number. Entry numbers are positional and shift whenever the listing does, so prefer a `partId`/`blobId`/`cid` for a reference you will reuse. **The entry-number form is read-only**: the same grammar is accepted by an `attachments` item that attaches a part to outgoing mail, and there a reference that resolved *only* by entry number is rejected — a wrong download costs a retry, while a wrong attach is baked into a draft you may then send.
  - `path` here is a **destination written** under `FASTMAIL_DOWNLOAD_DIR` — the opposite direction from the compose tools' `attachments[].path`, which is a **source read** under `FASTMAIL_ATTACH_DIR`. It may be absolute or relative. Relative paths (including a bare filename) resolve against the download directory, so an attachment lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected. To save directly into your own location, set `FASTMAIL_DOWNLOAD_DIR` to that root (see [Setup](#setup)) — confinement stays on, scoped to the directory you choose. Parent directories under that root are created as needed.
- **get_thread**: Get all emails in a conversation thread. Returns metadata + preview for each email, or the full plain-text bodies with `includeBodies`.
  - Parameters: `threadId` (required), `includeDrafts` (optional, include in-progress drafts), `includeBodies` (optional, return each message's `bodyText`), `stripQuoted` (optional, requires `includeBodies` — return each message's new text only), `fields` (optional array — return only these fields per message, see [Field projection](#field-projection-fields)), `raw` (optional, return original JMAP response)
  - `includeBodies` turns an N-message conversation read into one call — see [Reading long threads cheaply](#reading-long-threads-cheaply).
  - Draft messages are **excluded by default** (an in-progress reply is noise when reading a conversation). When any are present, a trailing note reports **how many drafts are hidden** (so a draft reply you already started isn't missed) — the drafts themselves are not surfaced. Set `includeDrafts: true` to include them. (The note is on the simplified path only; `raw` output stays pure JSON, so a raw consumer passes `includeDrafts` itself.)

> **Draft handling is asymmetric by design.** `get_thread` excludes drafts by default while `search_emails`/`list_emails`/`get_recent_emails` include them: a draft reply is noise when reconstructing a conversation, but a search/list should still find everything you've written. Drafts are identified by the `$draft` keyword (robust even if a draft is moved out of the Drafts mailbox), not by mailbox role - with one carve-out: a draft that now lives ONLY in Trash is neither shown nor counted, since it is not an active draft (`edit_draft` and `delete_email` both leave a `$draft` copy there, and counting those would inflate the warning). Use `includeDrafts` (get_thread) / `excludeDrafts` (the three list/search tools) to override either default.

### Archiving does what the Fastmail client does

`archive_email` is not a move. Fastmail's Archive button **removes the message from the Inbox** and leaves every other folder and label untouched; the Archive folder is added only when removing the Inbox would otherwise leave the message filed nowhere. This server does the same thing, measured against the live client rather than inferred from the protocol:

| The message is filed in | What archiving does |
| --- | --- |
| Inbox + a label | The Inbox membership goes, the label stays, **Archive is not added** |
| Inbox only | It moves to Archive |
| Anywhere else, no Inbox | **Nothing.** It has already left the Inbox - the client calls this "already archived" - and that is a successful outcome |
| Trash, Spam, Drafts, Scheduled, Sent or Snoozed, no Inbox | **Refused**, saying why, and naming an alternative in this server where one exists (a scheduled send and a snooze have none, and the refusal says so) |

Corroboration for the first row: of the 15,843 messages in Archive on the account this was measured against, **zero** are also in another mailbox. Fastmail never creates the Archive-plus-label state that a replace-then-add implementation produces.

The Inbox is checked **first**, so a message that is somehow in the Inbox *and* in Trash is archived (keeping Trash, gaining nothing) rather than refused - that is what the client does when shown the same state.

`emailIds` takes an array, so one call handles one message or a batch, and there is no `bulk_archive`. It **never throws on a partial failure**: every id comes back in one of six buckets - `movedToArchive`, `removedFromInbox`, `notInInbox`, `refused`, `notFound`, `failed` - with counts that sum to the number of distinct ids you passed. The response carries a prose summary and, beside it, the result as JSON (`{ counts, results }`), so the per-id detail and the counts can both be parsed rather than read out of the sentence. Each entry carries `mailboxes`/`roles`: the **projected** filing for the two branches that wrote, the **observed, unchanged** filing for `notInInbox` and `refused` (which write nothing), and for `failed` the filing as **observed before** the attempted write - a weaker claim, because one `failed` sub-case is an id the server acknowledged in neither of its result maps, whose outcome nothing confirmed either way. Read `roles` for `archive` to tell whether a message is in Archive; the branch name will not tell you, because a message that was in Inbox *and* Archive reports `removedFromInbox` and is in Archive.

Two exceptions to "each entry carries them". A `notFound` entry has neither, because there is no filing to report - and so does the `failed` sub-case where the message's current filing could not be **read** - the server returned no `mailboxIds` object, or returned an empty one, which is not a filing a message can have. Its filing was never observed, which is precisely why it failed. And `roles` is omitted whenever nothing the message is filed in has a role. If an entry carries `unresolvedMailboxIds`, a mailbox id could not be resolved to a name: `mailboxes`/`roles` are **incomplete** for that message and those raw ids are the remainder. Those two read together - a missing `roles` with no `unresolvedMailboxIds` means "no role mailboxes"; a missing `roles` *with* them means the roles of the listed ids are unknown. Same caveat as the `unresolvedMailboxIds` on simplified email output: the full set reported for that entry is `mailboxes` ∪ `unresolvedMailboxIds`, and its presence is not an error. (For the two branches that wrote, that union is the projected destination set rather than an observed membership.)

Three things it does not promise:

- **It archives the messages you name, not the conversation.** Fastmail's Archive button acts on a single message or on the whole thread depending on a per-user display setting, and this server cannot see which one an account is using. Measured both ways on the same three-message thread: ungrouped, archiving one message left its two siblings sitting in the Inbox; grouped, one click cleared all three. The rule applied to each message is identical either way - only the fan-out differs. So to archive a conversation, pass every message id in it (`get_thread` lists them); pass one and the rest of the thread stays in the Inbox. Thread ids are not accepted ([#128](https://github.com/JonathanGodley/fastmail-mcp/issues/128)).
- **No keyword is written, which is not the same as the read state being untouched.** JMAP reports `$seen` only when *every* one of a message's per-mailbox copies carries it, so dropping an unread Inbox copy from a message whose other copy was already read flips it to read with no keyword write anywhere.
- **This describes an account in labels mode.** In folders mode a message has a single membership, so every archive is the move-to-Archive case. The server does not detect which mode an account is in ([#122](https://github.com/JonathanGodley/fastmail-mcp/issues/122)).


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

- An **unrecognised quote shape passes through unchanged** rather than being guessed at, and says so with a `0`. Text derived from HTML-only quoting (Outlook's `<div>` nesting, a Gmail quote flattened without `>` prefixes) has no text-level boundary to find. This is the same foreign-client recognition residual as the compose-side quote guard; see [`docs/email-bodies.md`](docs/email-bodies.md).
- **Localized attributions** ("schrieb:", "a écrit :") aren't recognised, so that one line survives above an otherwise stripped quote.

*Too much stripped:*

- **A leading `>` is not only a quote marker.** A markdown blockquote in the sender's own writing, or pasted shell/REPL/diff output whose prompt is `>`, is stripped as quoted text. `>` runs are the primary marker the feature is built on, so this is it working as specified — telling one from the other is not decidable from the text.
- **Pasted email headers count as a quoted section** when the `From:` line carries an address, and that rule cuts to the end of the message. A display-name-only paste (a job posting's "From: The Hiring Team", a newsletter) carries no address and passes through untouched.
- On a **forwarded** message, the forwarded content sits below `-----Original Message-----` too, so `stripQuoted` removes it and leaves your covering note.
- A **prose line ending in "wrote:"** directly above a genuine quote can be pulled in with the attribution (the same walk-back that catches Gmail's wrapped two-line attribution).

*Other:*

- `stripQuoted` applies to **`bodyText` only**. A `bodyHtml` returned alongside it (`get_email` with `verbose`) is untouched.
- `stripQuoted` **cannot be combined with `raw`** (raw is unmodified JMAP), and on `get_thread` it requires `includeBodies`. Both combinations are rejected with an error rather than quietly ignored.
- `get_thread`'s combined bodies are capped at **100,000 bytes** (~25k tokens). Over that the call **errors**, naming the largest messages and telling you to add `stripQuoted: true` or fetch them individually — a silently truncated body is indistinguishable from a short message, which is the trap this feature exists to remove. The cap is measured on what would actually be returned, so `stripQuoted` genuinely brings an over-cap thread back under it.

### Message bodies

`create_draft`, `reply_email`, `forward_email`, and `edit_draft` all take `textBody` / `htmlBody`. HTML is the source of truth and the plain-text part is derived from it when you don't supply one. In that derivation an image contributes its alt text, and an embedded (`cid:`) image with no alt contributes `[image]`; a remote image with no alt contributes nothing, so a body that is nothing but remote images still ships HTML-only. Three malformed body shapes are **rejected** with an `InvalidParams` error rather than stored, because each one otherwise ships and reaches the recipient looking wrong:

- **A non-string body.** Passing a number, array, or object where a body string belongs is rejected naming the parameter ([#62](https://github.com/JonathanGodley/fastmail-mcp/issues/62)). An omitted body, or an explicit `null`, just means "not provided".
- **An `htmlBody` that is entirely HTML-escaped** — escaped element tags (`&lt;p&gt;`) and no actual elements ([#71](https://github.com/JonathanGodley/fastmail-mcp/issues/71) / [#77](https://github.com/JonathanGodley/fastmail-mcp/issues/77)). Stored as-is it renders as literal `<p>` tags to the recipient, and the derived text part carries the same junk. Escaping is a reasonable-looking default for a client filling a markup field, so this is an easy mistake to make and an invisible one to catch. The error names the fix: pass real markup, or use `textBody`. Escaped markup **inside** real elements (`<pre>&lt;p&gt;</pre>`) is legitimate and passes, as does an ordinary `&amp;`. Because the check only fires on a body with no real markup at all, it is deliberately narrow about what counts as an escaped tag — a known element name followed by a real tag delimiter — so ordinary prose like `Hi &lt;name&gt;`, `mail me at &lt;a@b.example&gt;`, or `reply with &lt;approve&gt;` still goes through.
- **A body wrapped in a CDATA section** ([#78](https://github.com/JonathanGodley/fastmail-mcp/issues/78)). The plain-text alternative is derived by an HTML parser that recognises the section and consumes it whole, so the message is lost from it entirely; a browser, which has no CDATA in HTML, instead drops the `<![CDATA[` opener as a bogus comment and renders the trailing `]]>` as visible text. On a reply or forward it is worst of all: the quoted original supplies enough visible content that nothing downstream notices, and a text-only recipient sees only their own words quoted back.

The CDATA rule differs by format, because the damage does. In `htmlBody` a section is rejected **wherever** it appears, since a mid-body one swallows just as much as a wrapper (to show a literal CDATA token in HTML you have to escape it as `&lt;![CDATA[`, which passes). In `textBody` only a body that **starts** with `<![CDATA[` is rejected — a plain-text part is never markup-parsed, so an XML snippet quoted inside a plain-text message is real content and keeps working. A bare `]]>` with no opening token is left alone in both formats: it renders as ordinary text and survives the text derivation intact.

The `textBody` refusal names `htmlBody` as the place to fix it, because the caller that trips it usually never typed the token. An `htmlBody` may legally carry an escaped `&lt;![CDATA[`; that escape **unescapes** when the plain-text alternative is derived from the markup, so handing the derived text back on a later `edit_draft` is refused. Editing `textBody` alone cannot clear it — the part is regenerated from the HTML — so the markup is the only thing that can change.

All of these are checked against **your** body, before any reply quote or forwarded-message block is merged in, so the reply and forward paths are covered rather than shadowed by the quote.

### Signing a message

The sending identity's configured signature - the sign-off you set in Fastmail - is not applied by JMAP, so nothing signs a message unless you ask. `create_draft`, `reply_email` and `forward_email` take `appendSignature:true`, which reads the signature off the identity the message sends as (the `from` address, or the account default) and places it in the body. It goes **above** the quoted original or the forwarded-message block, where a sign-off belongs; below one reads as part of the quote. The default is off ([#33](https://github.com/JonathanGodley/fastmail-mcp/issues/33)).

When you ask and the sign-off does not land where it should, the result **says so and names the reason**. That is not an error - the draft is still saved - but it is not silence either, because `appendSignature` is an input you cannot check without re-reading the draft. There are five reasons:

- no signature is available for the address the message sends as - it has none configured, or it is not one of your verified identities;
- the call wrote no body for the sign-off to sit under (`create_draft` with only `to`/`subject`, or a `reply_email` with no `textBody`/`htmlBody`). **A blank body counts as none**, so `textBody:"   "` is reported here rather than quietly becoming a message whose entire content is your sign-off;
- the body you passed already carries a signature;
- the message ships no HTML and the identity's signature has no plain-text form (an images-only HTML signature: no HTML ships, so no image does either, and a bare `[image]` line would describe something the message does not carry);
- the HTML body *was* signed but the message's plain-text alternative was not, because the signature's only content is a **remote** image, which derives no readable text. This is the one **partial** outcome - a recipient rendering the HTML sees the sign-off, one reading the text alternative does not - so its note says exactly that rather than claiming nothing was appended. It is reported whether you supplied that plain-text alternative yourself or left it to be derived from the HTML; the recipient's experience is the same either way, so the report does not depend on which fields you passed. An **embedded** (`cid:`) image signature is *not* this case: the image ships with the HTML, so the text alternative gets an `[image]` placeholder, which is what the derived text of any body carrying an image gets.

One asymmetry worth knowing: an **inline** `forward_email` with no note at all *is* signed - the signature becomes the whole of the content above the forwarded-message block, because "FYI, see below" is the normal shape of a forward. A `reply_email` with no body is **not**; a reply whose entire content is a sign-off over a quote is not a message. The reply says so in its result rather than signing. (So on `forward_email` the "no body written" reason above means a note you *passed* that was blank, never an omitted one - its `appendSignature` description says so in those terms rather than repeating the shared rule. On an `asAttachment` forward that reason cannot arise at all: a note-less one writes the filler body `Forwarded message attached.` and signs that.)

Which form you get follows the body, not which fields the identity has configured. An HTML body gets the HTML signature, and the plain-text alternative is **derived from that HTML** rather than copied from `textSignature` - because the text part of an HTML message is regenerated from the HTML on the first HTML-only edit, so a verbatim copy would look right and then change by itself with nothing reporting it. A plain-text-only message gets `textSignature` as configured. An identity that has only one of the two forms still signs either kind of body; the missing form is derived.

`edit_draft` signs nothing on its own. It has no `appendSignature`: the body you give it is stored exactly as written, so a sign-off is there if you put it there and gone if you did not. What it offers instead is `expandSignature: true`, which expands a `{{signature}}` **you wrote** in the body of that edit into the same sign-off the compose tools place, resolved against the address the edit sends as. Everything above about *which form* you get applies unchanged: the HTML body gets the HTML signature, a plain-text-only message gets `textSignature`, and the plain-text alternative of an HTML message is derived from that HTML rather than copied.

The flag exists rather than the token alone because part of a draft body handed back to an edit was written by someone else — the quoted original, the forwarded message — so an in-band trigger could be **planted** by them. A flag cannot be. The consequence is the useful one: with the flag absent, a `{{signature}}` sitting in a stored body is stable under every edit. What the result reports is what *you* added — a token the body you hand back carries more of than the stored body did is named in a note; one that was already in the stored body passes in **silence**, because reporting it would mean reporting the original author's text back to you on every edit of that draft for as long as it exists.

Two things follow for a read-edit-write loop. A sign-off already in the body you read comes back **because you handed it back**, not because anything recognised it — so an edit that rewrites the body around it keeps it, and one that replaces the body drops it, with no note either way. And an edit that expands anything returns **no `bodyHash`**, because what landed is not what you wrote; re-read the draft before the next edit.

The "body you passed already carries a signature" reason above is decided by looking for the block this server writes. In HTML that is a `fm-mcp-signature` class; in plain text there is no marker, so the body is searched for **either** form of that identity's block - the HTML-derived one or the configured plain-text one - because which one is sitting in a body depends on whether the call that put it there shipped HTML.

Where the block sits is checked as carefully as which form it is. A body carrying a quoted original or a forwarded-message block has the sign-off somewhere in the **middle** rather than at either end. The check works by subtraction: it cuts the body at the forwarded-message separator, dropping that line and everything below it, then looks for the sign-off as a run of **whole lines** in what is above. Nothing tries to interpret the `On <date>, <name> wrote:` attribution above a quote, which is the part that varies by client and by language; a dashed separator does not.

A quoted original needs no special handling, because the comparison is line-by-line and `> Regards,` is not `Regards,` - so a thread that quotes an earlier signed message of yours is not mistaken for a signature on the draft. That also means a signature that **itself** contains a quotation (a quote-of-the-day sign-off, which renders as `> `-prefixed lines in the plain-text form) is recognised normally.

The accepted cost: a message you forward in a shape this server does not recognise as a forward - an Outlook-style `From:/Sent:/To:/Subject:` header block, with no dashed separator - is not removed, so if that forwarded message carries this identity's own sign-off, your note reports "already carries a signature" and goes out unsigned. That is the cheap error on purpose: it is announced, and you can write the sign-off into the note yourself. The expensive error is the other one - two sign-offs shipped in silence - so wherever the check cannot be sure, it declines to append and tells you.

The match is normalised for the things a round trip through a mail store changes without changing what the message says: any line ending becomes a newline, and each line is trimmed at both ends. That covers a stripped trailing space (`-- `, the standard signature delimiter, ends in one), that space arriving back as a non-breaking space, and the leading space a `format=flowed` sender adds to a line. The same idempotence covers HTML from the other side - a body you pass that already contains one of this server's signature blocks is left alone rather than signed twice.

The marker is stripped when the message is later quoted in someone's reply (the quote sanitiser removes `class` from every quoted element), so a signature quoted back can never be mistaken for that draft's own.

### Sending attachments

`create_draft`, `reply_email`, `forward_email`, and `edit_draft` accept an `attachments` array (on `forward_email` these are NEW attachments to add — the original's own attachments are carried separately, governed by `includeOriginalAttachments`/`asAttachment`, and the images its body displayed are carried as body content; see [Replying and forwarding with images](#replying-and-forwarding-with-images)).

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

**A positional `attachmentId` is refused here.** `download_attachment` accepts a plain entry number (`0`, `1`, ...); attaching does not. A reference that resolved **only** through the entry-number fallback is rejected, and the error points you at `get_email_attachments` for the part's `partId` or `blobId`. The refusal is decided by how the reference actually resolved, not by whether it looks numeric.

```json
{
  "to": ["someone@example.com"],
  "subject": "Just the invoice from that thread",
  "textBody": "Only the one attachment, not the whole message.",
  "attachments": [{ "emailId": "M0f1...", "attachmentId": "Gd41f8..." }]
}
```

The `attachmentId` there is the `blobId` (or `partId`) of the part you want, read off `get_email_attachments`. A digit string is not automatically wrong — Fastmail `partId`s *are* digit strings, and one that matches a real part is accepted — but a digit that matches no part falls through to the entry-number form, and that is the one this refuses.

On `edit_draft`, `attachments` **appends** (existing attachments are kept). Use `removeAttachments` (a `blobId` from `get_email_attachments`, or a unique name) to drop specific ones, or `clearFields:['attachments']` to remove all. Passing `attachments` and `clearFields:['attachments']` together is rejected as a conflict.

> **Two opt-ins, by design — one per boundary.** Both are off until you set them, and the server must be restarted to pick either up.
>
> **`path` needs `FASTMAIL_ATTACH_DIR`.** Attaching a local file reads it off the disk and emails it out — an exfiltration vector — so that source is **disabled until you set `FASTMAIL_ATTACH_DIR`** (see [Setup](#setup)); until then every attempt fails with a self-documenting error and no file is read. Files are confined to that directory: a path outside it, a missing file, or a symlink escaping the root is rejected. The attach directory is resolved **independently** of `FASTMAIL_DOWNLOAD_DIR` — pointing both at the same directory re-opens a download-then-email round-trip, so that is your explicit choice, not a default. The confinement narrows time-of-check/time-of-use races and blocks symlink escapes, but a same-inode swap race and hardlinks inside the root are residual.
>
> **`blobId` and `emailId` + `attachmentId` need `FASTMAIL_ALLOW_BLOB_ATTACH`.** They cross a different boundary: nothing is read off your disk, so the directory opt-in has nothing to say about them and they get their own flag, off by default and parsed strictly — only `true` or `1` enables it, so a literal `"false"` leaves it off rather than reading as "a value was set". A Desktop Extension install offers the same gate as an unchecked "Allow Attaching Content Already In The Account" box: the `"false"` an unchecked box sends is not an enable. What this gate protects is **provenance more than reach**: you could already put another message's attachment in front of an arbitrary recipient by forwarding it and sending the draft, but a forward carries a visible forwarded-message block and an `X-Forwarded-Message-Id` header, while a blob-attached part carries neither. Note also that a message's own `blobId` appears in ordinary read output, so with this gate open a caller can attach a **complete raw message** — transport headers and all — to a fresh draft; prefer `emailId` + `attachmentId`, which can only ever name one part.
>
> Full posture, including why `blobId` is not restricted to ids this server handed out, in [`docs/security-model.md`](docs/security-model.md).

### Embedding an image in the body

Give an attachments item a `cid` and reference it from `htmlBody`, in the same call ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)):

```json
{
  "to": ["someone@example.com"],
  "subject": "The new logo",
  "htmlBody": "<p>Here it is:</p><img src=\"cid:logo\" alt=\"Our new logo\">",
  "attachments": [{ "path": "logo.png", "cid": "logo" }]
}
```

The same pair works on `reply_email`, `forward_email` and `edit_draft`. What you can rely on:

- **A reference with no matching item is rejected**, rather than saved as a broken image. The error names the identifier and how to fix it.
- **A file you supply is never dropped.** If nothing displays it (no `htmlBody` at all, or no reference to that `cid`), it is still attached as an ordinary file and the result says so in plain words, so a silent demotion can't pass for success.
- **The saved draft is read back** whenever a call embeds something, and the result reports what it actually carries, with sizes. That is the one outcome you can't check from an email id alone.
- **Identifiers of the form `ii-<hex>@inline.invalid` are reserved** for images this server embeds on your behalf when it builds a quoted or forwarded block. Authoring one is rejected — it names an image that does not exist. Handing one **back** is not: on `edit_draft` such a reference is accepted whenever the draft actually carries a part under it, which is what lets you read an image-bearing draft and edit the prose around it. Drop the reference and the part comes off with it.
- The plain-text alternative derived from your HTML shows the image's alt text, or `[image]` when it has none, so a picture-only message still reaches text-only readers.

Embedding needs the same opt-in as any attachment, and either gate will do: with neither `FASTMAIL_ATTACH_DIR` nor `FASTMAIL_ALLOW_BLOB_ATTACH` set there is no image an `attachments` item could supply, and the refusal says so instead of pointing you at a parameter that can't work. With either one open, "add an `attachments` item supplying that `cid`" is a repair you can actually make.

### Replying and forwarding with images

When you reply to or forward a message whose body **displayed** an image, that image is carried into the quote or the forwarded block ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)).

This means a reply or a forward can send image data outward that you never attached and that `FASTMAIL_ATTACH_DIR` never sees — the parts are already in the account and are re-referenced by blob, so nothing is read off your disk and nothing is uploaded. What gets carried is bounded only by this:

- the body **references** the part, and
- the sender **declared** it `image/*`.

That declaration is metadata. Nothing is sniffed and nothing verifies the claim, so a part labelled `image/png` is carried whatever it actually contains. There is **no size limit and no count limit**, because re-referencing a blob costs nothing to send.

Your recourse is per-tool, and it is coarse on purpose:

- **`reply_email`** — `quoteOriginal: false` omits the whole quote, and is the only way to send none of it. There is no quote-text-without-images option.
- **`forward_email`** — body images are carried **even when `includeOriginalAttachments` is false**, because they are body content rather than attached files; that flag governs the original's attached *files*. If you must forward without the images, use `asAttachment: true` (the original rides as a `.eml` and nothing is re-composed) or don't forward it.

Carrying needs an HTML body to put the images in:

- A **text-only reply** can't display them, so it drops them and the result says how many. Supply `htmlBody` (or let the quote build one) to keep them.
- A **text-only forward**, or an image the block can't display — a reference that didn't resolve to exactly one image part, or a duplicate identifier shared by several parts — falls back to riding as a regular attachment. That fallback *is* governed by `includeOriginalAttachments`, and the result says what happened either way.

Images written as `data:` URIs are dropped and counted rather than converted, as are images referenced by a form a quote cannot carry — a relative or protocol-relative path, which resolves against an origin the new message does not have. The quote still ships, and the result says how many went. Identifiers this server mints for carried images (`ii-<hex>@inline.invalid`) survive an edit of the same draft for as long as the body you hand back keeps referencing them; drop the reference and the part comes off. Never author a new one — it names an image that does not exist, and is rejected.

### Email Statistics & Analytics

- **get_mailbox_stats**: Get statistics for a mailbox (unread count, total emails, etc.)
  - Parameters: `mailbox` (optional — [id, role, name, or path](#naming-a-mailbox); defaults to all mailboxes)
- **get_account_summary**: Get overall account summary with statistics

### Bulk Operations

- **bulk_mark_read**: Mark multiple emails as read/unread
  - Parameters: `emailIds` (required array), `read` (default: true)
- **bulk_pin**: Pin or unpin multiple emails
  - Parameters: `emailIds` (required array), `pinned` (default: true)
- **bulk_move**: Move multiple emails to a mailbox. **Replaces each message's entire mailbox membership** - every other label or folder each one was filed under is removed. To file them somewhere while keeping those, use `bulk_add_labels` - which takes the Inbox and the account's own labels only, so a folder is reachable by moving alone (archiving is the exception: `archive_email` keeps the rest of the filing). No keyword is changed, so each message keeps its read/unread and flagged state.
  - Parameters: `emailIds` (required array), `targetMailbox` (required — [id, role, name, or path](#naming-a-mailbox))
- **bulk_delete**: Delete multiple emails (move to trash)
  - Parameters: `emailIds` (required array)
- **bulk_add_labels**: Add labels to multiple emails simultaneously. Labels means the **Inbox and the account's own labels**; a folder (any other role mailbox) is rejected - use `bulk_move`
  - Parameters: `emailIds` (required array), `mailboxIds` (required array - each entry [id, role, name, or path](#naming-a-mailbox), resolved like every other mailbox tool; any unresolved or ambiguous entry rejects the whole call with the valid list, as does any entry that resolves to a folder rather than a label)
- **bulk_remove_labels**: Remove labels from multiple emails simultaneously. Labels means the **Inbox and the account's own labels**; a folder (any other role mailbox, Archive included) is rejected before anything is written - use `bulk_move`. The last-label archive fallback is decided per message, so one batch can archive some and merely unlabel others; a rejection is not, and aborts the whole batch
  - Parameters: `emailIds` (required array), `mailboxIds` (required array - each entry [id, role, name, or path](#naming-a-mailbox), resolved like every other mailbox tool; any unresolved or ambiguous entry rejects the whole call with the valid list, as does any entry that resolves to a folder rather than a label)

There is **no `bulk_archive`**. `archive_email` already takes an `emailIds` array and handles a batch in one call, reporting each id's outcome separately - see [Archiving does what the Fastmail client does](#archiving-does-what-the-fastmail-client-does). Reaching for `bulk_move` with a target of `"archive"` instead does the wrong thing: it replaces each message's whole membership, dropping every label, which is the behaviour the archive rewrite exists to stop. Its destination is also resolved [the ordinary way](#naming-a-mailbox), which falls through to matching on name; `archive_email` and `delete_email` do not, resolving by role and nothing else, so a folder named after a role can never become their destination on any account.

### Contact Tools

All three read tools carry `kind` on a card that is not an ordinary person - see [`kind` tells you which cards the write tools will refuse](#kind-tells-you-which-cards-the-write-tools-will-refuse).

- **list_contacts**: List all contacts. Returns simplified format by default.
  - Parameters: `limit` (default: 20, hard cap 100 — no paging), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **get_contact**: Get a specific contact by ID. Returns simplified format by default. Throws if the ID is not found.
  - Parameters: `contactId` (required), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **search_contacts**: Search contacts by name or email. Returns simplified format by default.
  - Parameters: `query` (required), `limit` (default: 20, hard cap 100 — no paging), `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **create_contact**: Create a contact. Needs at least a name or one email address. Returns the created card, read back after the write, in the same shape `get_contact` returns. See [Writing contacts](#writing-contacts) for the accepted entry shapes.
  - Parameters: `name` (full-name string or `{given?, surname?, full?}`), `emails` (array; bare address strings or `{address, label?}`), `phones` (array; bare number strings or `{number, label?}`), `addresses` (array of `{full, label?}` objects), `notes`, `addressBookId` (optional, defaults to the account's address book), `verbose`, `raw`
- **update_contact**: Update a contact, **merging per entry** so the stored fields the simplified output doesn't show (`contexts`, `pref`, …) survive an edit. Returns `{contact, previousCard}`. An edit that both drops a stored entry and adds an unknown one is rejected as ambiguous unless `allowEntryReplace` is set (which is scoped to that one array); `addresses` replace wholesale; `name` merges into the stored components; a changed label is added, never removed. See [Writing contacts](#writing-contacts).
  - Parameters: `contactId` (required), `name`, `emails`, `phones`, `addresses`, `notes` (same shapes as create; `[]` is rejected — use `clearFields`), `clearFields` (array of `"emails"`/`"phones"`/`"addresses"`/`"notes"`; `name` is not clearable), `allowEntryReplace` (boolean), `verbose`, `raw` (both apply to `contact` only — `previousCard` is always the raw card)
- **delete_contact**: Delete a contact. **Irreversible** — a card does not go to Trash. Returns `{deleted, deletedCard}`, the id plus the full card as it stood immediately before the destroy, so a wrong delete is at least visible in full. `create_contact` can rebuild the name, emails, phones, addresses and note from it, but not the rest of the card — see [Writing contacts](#writing-contacts). A contact **group** is refused, because this server cannot create one and so cannot put one back. There is deliberately no confirmation parameter.
  - Parameters: `contactId` (required)

### Calendar Tools

#### Calendar invitations send mail

Naming attendees on an event is not a local write. When the event is created the server emails each attendee a meeting invitation from this account, and deleting that event later emails them a cancellation. Updating attendees does the same for whoever the change affects. None of this goes through a compose tool, so it happens with no `send_draft` call and produces no draft to inspect first.

This is the one hole in the draft-first posture, and it matters most if you are relying on tool names for permissions: gating `send_draft` gates every message you *compose*, and gates none of this. If a caller can reach `create_calendar_event`, it can cause mail to be sent to an address of its choosing.

Omit `participants` to write an event without notifying anyone. The behaviour is the server's, not this server's — attendee scheduling is what CalDAV does — so there is no flag here that turns it off. An operator who needs the guarantee that nothing leaves the account has to deny `create_calendar_event`, `update_calendar_event` **and `delete_calendar_event`** alongside `send_draft`. Delete belongs on that list for a reason that is easy to miss: it takes no `participants` of its own, so it reads like a purely local removal, but deleting an event someone else is already on mails all of them.

- **list_calendars**: List all calendars. Each entry carries `id`, `displayName`, `url`, and `description`/`color` when set. A calendar's `id` **is** its CalDAV URL, and every `calendarId` parameter accepts either that or the calendar's display name.
- **list_calendar_events**: List calendar events (core fields only — no participants for token efficiency), one entry per occurrence in the requested window ([#64](https://github.com/JonathanGodley/fastmail-mcp/issues/64), [#100](https://github.com/JonathanGodley/fastmail-mcp/issues/100)).
  - Parameters: `calendarId` (optional; the `id`/URL or the display name from `list_calendars`, omit to read every calendar), `startDate` (optional, see [Calendar window dates](#calendar-window-dates)), `endDate` (optional, same), `limit` (default: 50, hard cap 500)
  - **Recurring events are expanded.** Give a `startDate`/`endDate` window and the server applies each recurrence rule, returning one entry per occurrence *inside that window* with `start`/`end` set to the occurrence's real dates. A fortnightly event across three months is seven entries, not one. **There is no unwindowed call**: omit both bounds and you get the next month - the start of today in the configured timezone, for 31 days - expanded like any other window, with a trailing `Note:` naming the range actually searched ([#142](https://github.com/JonathanGodley/fastmail-mcp/issues/142)). An absent window is an *open-ended* one, not an empty one, and expanding recurrences across it would materialise every occurrence of every repeating event. To see a series unexpanded, at its original start date and with its `RRULE` intact, call `get_calendar_event` on the row's `id`.
  - **Four fields say what a date is**: `recurrenceId` present means `start`/`end` **are** the in-window occurrence; `recurrenceRule` present (the raw `RRULE`) means this is the series **master**, shown at its **original** start date, which may be years before the window; `recurrenceDates` present (the raw `RDATE` values, comma-separated) says the same for a series that **lists** its occurrences instead of stating a rule; `isRecurring` says the entry belongs to a repeating series. **`recurrenceDates` carries the values and not the parameters** - a `TZID`, `VALUE=DATE` or `VALUE=PERIOD` on the `RDATE` line is dropped - so a designator-less value in that list does **not** follow the `timeZone` rule that governs `start`, and RDATEs written in different zones are indistinguishable once joined into one list. It is proof that other dates exist, not a date you can place on a clock; pass a window and let the server expand the series if you need real occurrence times. Master and occurrence are the usual split, but the two fields *can* appear together — RFC 5545 lets an override block carry its own rule — and where they do, `recurrenceId` names the instance and `recurrenceRule` is that block's own rule rather than the series'. There is one case the expanded output cannot distinguish, and it is the server's doing rather than a choice made here: Fastmail sets a `RECURRENCE-ID` only on instances **after** the first, so a series' first occurrence arrives bare. It is still reported as recurring when a sibling occurrence shares the window, but a window holding **only** that first occurrence yields an entry indistinguishable from a one-off. Call `get_calendar_event` on its `id` to see the master and its `RRULE`.
  - **A one-sided window is bounded to 31 days.** Pass only `startDate` (or only `endDate`) and the missing half is filled in a month away, then the response carries a trailing `Note:` naming the range actually searched. The old open-ended defaults (1970/2099) were harmless while the window was only a filter; once recurrences are expanded server-side that range is what the server *materialises occurrences over*, and `startDate: <today>` alone asked it to generate every occurrence of every repeating event for 73 years. A month is what a calendar client shows at a time, and it is the question a missing bound is asking - the expanding APIs nearest to this one (Cyrus's own JMAP query, Microsoft Graph's `calendarView`) require a bound rather than invent one at all. Name both bounds to choose the span yourself.
  - **`limit` is a genuine earliest-N.** Every target calendar is queried, and only then is the combined set sorted by start time and trimmed — so a cross-calendar availability check can no longer lose whole calendars to an earlier one filling the quota. The sort compares *instants*, not the text of the dates: a calendar mixes floating wall clocks, UTC instants and named zones (see [Calendar times: a name, never an offset](#calendar-times-a-name-never-an-offset)), and comparing those spellings as strings put a genuinely earlier event after a later one, where `limit` then cut the earlier of the two.
  - **Every occurrence row of a series carries the SAME `id`, and it names the series — which is why the write tools refuse it.** Seven rows of a fortnightly event are seven identical ids, and acting on one would act on all seven. So `update_calendar_event` and `delete_calendar_event` **reject** any id that resolves to a repeating event, with no override parameter (see [Repeating events cannot be changed or deleted here](#repeating-events-cannot-be-changed-or-deleted-here)). Change or remove a repeating event, or a single occurrence of one, in the Fastmail web interface. A row's **`url` is the same record under another name**, and passing it in place of `id` reaches the same refusal — the check reads the stored record, not the shape of the argument. It is kept on the row because a UID is unique per collection rather than per account, so where two calendars hold the same UID the `url` is the only thing telling the records apart.
  - **Rows never carry `participants` or `organizer`**, on any event, because this call does not fetch them. Empty fields are omitted throughout, so an absent participant list here cannot be told apart from an event with no attendees at all. Call `get_calendar_event` before anything destructive: it is the only way to see who the server is about to mail.
  - **Rows are filtered exactly against the window you asked for, and all-day events are your LOCAL days.** An all-day event on a neighbouring day is no longer returned; a date-only value covers that whole day in the account's configured zone, and a multi-day date span covers all of its days. The range this server *requests* of Fastmail is up to 14 hours wider at each edge than the window you gave, because the server would otherwise withhold an all-day or floating event from a window narrower than a day - the extra rows that pulls in are trimmed before you see them, and the `Note:` line, when there is one, always describes the window you asked about rather than the widened request. Two kinds of row can still sit outside the window (see [Calendar known limitations](#calendar-known-limitations)), and they fail in **opposite directions**. A block still carrying its own `RRULE`/`RDATE` only ever **adds** a row, so check each `start` against the window you asked for rather than assuming every row is inside it. A **floating** timed event that expansion has stamped as UTC is judged on UTC, so it can be **absent** from the window it really belongs to and present in a neighbouring one - which means that on an account far from UTC a row you expected can be missing. Separately again, a row can fail to arrive at all: a series that lists its dates as `RDATE`s and states no `RRULE` is matched by Fastmail's own filter over the series start alone, so a window covering one of its listed dates and not that start returns nothing whatsoever for it, and nothing on this side can recover a row the server never sent. **An empty result from this call is therefore not proof of a free day** - on any account, not only one far from UTC.
  - **The response opens with a summary line** stating how many events matched in total (`Showing 50 of 137 results.`). When that total exceeds the number returned, `limit` cut the rest off; there is no paging, so **raise `limit`** (hard cap 500) to reach the rest, and narrow the window only when the total is above 500. Expansion makes the cap easier to hit than it used to be.
  - **A TOTAL discovery failure is an error, not an empty list.** If the CalDAV server fails to list the account's calendars at all, this raises rather than returning `[]` — an empty result to an availability question reads as "you are free", which is the dangerous way to be wrong. A `calendarId` that matches no calendar raises for the same reason, naming the calendars that do exist: the display name is matched case-sensitively, so `"work"` for a calendar called `"Work"` is a plausible miss and used to come back as an empty day. Surrounding whitespace is ignored on both sides, and `list_calendars` reports the trimmed name, so the name it gives you is one this parameter resolves. An **empty or whitespace-only** `calendarId` is such a value, not a way to ask for every calendar — omit the parameter for that. It narrows what the call touches, so it fails closed rather than silently widening to the whole account.
- **get_calendar_event**: Get a specific calendar event by ID. Returns organizer and participants when available. Throws if the ID is not found. For a repeating event this returns the **series master**: `isRecurring` is set along with `recurrenceRule` (the raw `RRULE`) and/or `recurrenceDates` (the raw `RDATE` values, for a series that lists its occurrences rather than stating a rule), and `start`/`end` are the series' original dates, **not** the next or nearest occurrence — use `list_calendar_events` with a window to get occurrence dates. This unexpanded path is where those fields actually show up: Fastmail strips them from an expanded block, and both halves are measured - `calendar-expand.probe.mjs` for `RRULE`, `calendar-rdate-expand.probe.mjs` for `RDATE`. `recurrenceDates` carries the RDATE **values only** - a `TZID`, `VALUE=DATE` or `VALUE=PERIOD` parameter on the line is dropped - so a designator-less value there does **not** follow the `timeZone` rule that governs `start`, and it is evidence that other dates exist rather than a date you can place on a clock. **And passing a window will not place it on one either, for a series that states no `RRULE`**: Fastmail's window filter matches such a record over its `DTSTART` alone, so a `list_calendar_events` window covering one of these `RDATE` values but not the series start returns nothing at all for it, with no sign in the response that anything was left out ([#167](https://github.com/JonathanGodley/fastmail-mcp/issues/167)). Widen the window until it reaches the series start - the `start` in this response - or confirm the date in the Fastmail web interface. The id is the series id, so `update_calendar_event`/`delete_calendar_event` on it act on every occurrence. One exception: where a record holds only overridden instances and no master at all, what comes back is one of those overrides and carries a `recurrenceId` — so a `recurrenceId` in this response means you are **not** looking at the master. This path returns the stored `DTSTART`/`DTEND` exactly as written, with `timeZone`/`endTimeZone` naming the zone when it isn't the account's configured one — see [Calendar times: a name, never an offset](#calendar-times-a-name-never-an-offset).
  - Parameters: `eventId` (required)
- **create_calendar_event**: Create a new calendar event. Supports date-only (e.g. `2026-04-01`) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs `end: "2026-04-02"`. See [Start/end agreement](#startend-agreement) for the form and ordering rules, [Writing calendar times](#writing-calendar-times) for `timeZone`, and [Calendar text size limits](#calendar-text-size-limits) for the size caps.
  - Parameters: `calendarId` (required; the `id`/URL or the display name from `list_calendars`, matched case-sensitively and resolved by the same rule as on a read — surrounding whitespace is ignored on both sides, and `list_calendars` reports the trimmed name, so the name it gives you is one this parameter resolves. A calendar `list_calendars` does not show cannot be written to either), `title` (required), `description` (optional), `start` (required, ISO 8601 or date-only), `end` (required, ISO 8601 or date-only), `location` (optional), `timeZone` (optional; IANA zone name for a designator-less `start`/`end` — omitted writes the account's configured zone, see [Writing calendar times](#writing-calendar-times)), `participants` (optional array; each entry is `{email, name?}` **or a bare email-address string**. The whole array may also be sent as a JSON string; a comma-joined string is *not* accepted. An unknown per-item key, a missing/non-string `email`, or a non-string `name` is rejected naming its position, e.g. `participants[2]`)
- **update_calendar_event**: Patch an existing calendar event. **Single (non-repeating) events only** — an id that resolves to a repeating event is rejected, and no parameter overrides that (see [Repeating events cannot be changed or deleted here](#repeating-events-cannot-be-changed-or-deleted-here)). Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged, though a call that names no field and no `clearFields` at all is rejected; passing an empty or whitespace-only string for `title`, `description`, or `location` is rejected (it won't silently blank the property). To delete `description` or `location`, list them in `clearFields`. Floating times (no Z/offset) preserve the original timezone. A new `start`/`end` is checked against the value it will sit beside — see [Start/end agreement](#startend-agreement). `timeZone` works differently here than on create — see [Writing calendar times](#writing-calendar-times). WARNING: providing `participants` replaces ALL existing attendee data; `participants: []` removes all attendees (and the now-orphaned ORGANIZER). The same [text size limits](#calendar-text-size-limits) apply as on create.
  - Parameters: `eventId` (required), `title`, `description`, `start`, `end`, `location`, `timeZone` (optional; IANA zone name for a designator-less `start`/`end` you are also passing this call — never defaults when omitted, see [Writing calendar times](#writing-calendar-times)), `participants` (array; each entry is `{email, name?}` **or a bare email-address string**, and the whole array may be sent as a JSON string — same rules and same index-naming rejections as on create), `clearFields` (array of `"description"`/`"location"` to delete)
- **delete_calendar_event**: Delete a calendar event. **Single (non-repeating) events only** — an id that resolves to a repeating event is rejected (below). **Sends mail** if the event has attendees — the server emails each of them a cancellation (see [Calendar invitations send mail](#calendar-invitations-send-mail)).
  - Parameters: `eventId` (required)
  - **It would delete a whole series, so it refuses to.** `eventId` names the calendar record, and every occurrence row `list_calendar_events` returns for a repeating event carries that same id — so deleting the row shown for one Thursday would destroy the entire series, every past and future occurrence with it, and send one cancellation to every attendee. Expansion made that much easier to reach: before it, a series appeared once at its original start date and "delete Thursday's meeting" found no row for Thursday; now it finds a plausible per-occurrence row carrying a usable id. So a repeating event is now **rejected outright** (see [Repeating events cannot be changed or deleted here](#repeating-events-cannot-be-changed-or-deleted-here)) — delete it, or one occurrence of it, in the Fastmail web interface.
  - **Irreversible, and nothing is echoed back.** A calendar event does not go to Trash the way an email does, and this tool reads nothing before destroying it — so unlike `delete_contact`, there is no `deletedCard`-style copy in the response to rebuild from. Call `get_calendar_event` first when the id might be wrong: that response is the only copy you will have, and it is also how you see whether anyone is about to be mailed.

#### Start/end agreement

An event's `start` and `end` are only meaningful together, so both `create_calendar_event` and `update_calendar_event` check the pair that will actually be written and reject it if it doesn't hold up. Two rules:

1. **Same form.** Both must land in the same one of: date-only (`2026-03-20`), a zone-designated time (`2026-03-20T09:30:00Z` or `2026-03-20T09:30:00+10:00`, both stored as UTC), or a time carried by a named `TZID` (the account's configured zone by default, an explicit `timeZone`, or an inherited stored zone — see [Writing calendar times](#writing-calendar-times)). A floating time with **no** zone at all can only arise on `update_calendar_event`, where a designator-less value with no stored `TZID` to inherit stays floating; `create_calendar_event` always resolves a designator-less value to a zone, so this form is unreachable there. A mixed pair has no single duration — `DTSTART:20260320T093000` (no zone) beside `DTEND:20260320T093000Z` (UTC) is zero-length in UTC and ends ten hours before it starts in UTC+10, and renders differently for every attendee.
2. **`end` after `start`.** Equal values are rejected too; RFC 5545 §3.8.2.2 requires DTEND to be strictly later. For all-day events, remember DTEND is exclusive.

On `update_calendar_event` the comparison is against **the value it will sit beside** — the other one you passed, or the stored one you left alone. So changing only `start` is fine when it stays in the same form and still lands before the stored `end`, but **moving an event to a different day, or converting one side to UTC, means passing both `start` and `end`.** A floating time on an event that carries a `TZID` still inherits that `TZID` (the long-standing behaviour), so it agrees with its partner and is accepted.

Two `TZID`-bearing values in *different* zones — a flight that departs in one zone and lands in another — are a legal shape and are accepted; the ordering check stands down there, so a genuinely backwards cross-zone pair is written rather than rejected. That is a deliberate stand-down rather than a limit: both zone names are present on the write path and can be resolved, and adding the check would newly reject input this tool accepts today, so it is tracked as [#140](https://github.com/JonathanGodley/fastmail-mcp/issues/140) instead of folded in. Both rejections are `InvalidParams` and name both values and both forms.

#### Accepted date spellings

`start` and `end` are read as strict ISO-8601 and nothing else: `2026-04-07`, `2026-04-07T14:00:00`, `2026-04-07T14:00:00Z`, or `2026-04-07T14:00:00+10:00` (surrounding whitespace is trimmed). Anything else is rejected as `InvalidParams` rather than guessed at. `2026/04/07` and `April 7 2026` are parseable by JavaScript's fallback date parser, but it resolves them against the *server's* time zone, so accepting them would put the event on a day that depends on where the server runs. A day that does not exist in its month (`2026-02-31`, in either the date-only or the datetime form) is rejected for the same reason — it would otherwise roll into the next month and the event would be created on a day nobody named. `create_calendar_event` and `update_calendar_event` apply the identical rule.

#### Calendar text size limits

Every text field that ends up in the event (`title`, `description`, `location`, and each participant's `name` and `email`) is written as an iCalendar content line, which has to be folded to 75 octets per RFC 5545. That folding is quadratic in the length of the field, so an unbounded value is a way to stall or kill the server process. Both `create_calendar_event` and `update_calendar_event` therefore bound the input up front:

| Bound | Limit |
| --- | --- |
| Any single field (`title`, `description`, `location`, participant `name`, participant `email`) | 64 KB |
| Number of `participants` | 500 |
| All of that text combined, in one call | 256 KB |

The combined bound is what stops many fields that are each just under the per-field cap from adding up to the same problem.

Oversized input is **rejected, never truncated**: a silently trimmed description is data loss you would not find out about until someone read the event. The rejection is `InvalidParams` and names the offending field, its actual size and the limit, so one retry with a shorter value fixes it.

#### Calendar window dates

`list_calendar_events`' `startDate` and `endDate` accept the same **spellings** as `create_calendar_event`'s `start`/`end` (see [Accepted date spellings](#accepted-date-spellings)): a date (`2027-03-01`) or a full datetime (`2027-03-01T09:00:00Z`, `2027-03-01T09:00:00+10:00`), with surrounding whitespace trimmed, and nothing else. `2027/03/01`, `"March 1 2027"` and a day its month does not have (`2027-02-30`) are all rejected naming the parameter, rather than guessed at or silently rolled into the next month — either would move the window off the days you asked about. **What a date-only value MEANS differs from the write tools**, and it is read in your timezone rather than UTC; both are covered below.

**A date is a LOCAL day.** `startDate: "2026-08-12"` means that calendar day where you are, not the UTC day. The zone is `FASTMAIL_TIMEZONE` (or `USER_CONFIG_FASTMAIL_TIMEZONE`), falling back to the zone this server itself runs in. It is the same single setting every email's `date` is rendered in, read from one place so the two can never disagree. `FASTMAIL_TIMEZONE`, if set, has to name a full IANA zone (see [Writing calendar times](#writing-calendar-times) for the exact rule); an unresolvable, offset-shaped or abbreviated value stops the server from starting at all, naming the value and the rule, rather than being silently substituted for at request time. Left unset, this server's own zone is used; on the near-unreachable case where that host zone is itself not a usable IANA name, dates fall back to UTC instead, with a one-time startup warning naming the rejected host zone. A datetime with **no** `Z` and **no** offset (`2026-08-12T09:00:00`) is local time too, deliberately; it used to resolve in the host's zone by accident of JavaScript date parsing, with nothing saying so. A value carrying `Z` or a numeric offset means exactly the instant it names and is never re-read.

That is not cosmetic. Reading a date as a UTC day answered a UTC+10 account with *10:00 on the 12th through 10:00 on the 13th*, so a day holding three appointments came back holding one — the missing two look exactly like a quiet morning. Every hour of offset from Greenwich is an hour of somebody's day answered from the wrong date.

The window runs from `startDate` **inclusive** to `endDate` **exclusive**, and there is one deliberate asymmetry between them:

- A date-only **`startDate`** is local midnight at the start of that day. (`search_emails`' `before`/`after` keep the UTC reading — they compare against a message's `receivedAt`, which is an instant, not a day on anyone's wall.)
- A date-only **`endDate`** covers the **whole of that day** — `2027-03-10` runs through to local midnight at the start of the 11th. So a single-day query is `startDate` and `endDate` on the *same* date. Read as the start of the 10th instead, that call would be a zero-length window returning nothing at all, which reads as an empty day rather than as a mistake. **This is deliberately the opposite of `create_calendar_event`'s date-only `end`**, which is exclusive at the *start* of that day (RFC 5545 DTEND, so a one-day event on April 1 takes `end: "2026-04-02"`). The two are not inconsistent, they answer different questions: on a write the value names one edge of a single event, on a read it names the last day you are asking about.
- A **datetime** `endDate`, with a designator or without, is taken literally as the exclusive end — no day is added, because a wall clock names a time of day rather than a day.
- A day that a daylight-saving change makes 23 or 25 hours long is still covered end to end; the bound advances by a whole *local* day, not by a fixed 24 hours. Where the change happens *at midnight* (Santiago, Havana) the day's own first or last hour is one the clocks skipped — a skipped wall clock resolves **forward by the length of the gap**, so the window still covers every instant of the day that exists (that lands on the moment the clocks jumped to only for a wall clock at the very start of the gap, and past it otherwise). A wall clock the change makes happen twice resolves to the earlier of the two.
- The time of day has to exist: `2026-08-12T25:00:00`, `12:75:00` or `99:99:99` is rejected rather than rolled into another day, exactly as `create_calendar_event` rejects them on a write. (`24:00:00` is accepted as the end of the day, which is what ISO-8601 says it means.)

A `startDate` that is not before the resulting `endDate` is rejected as `InvalidParams`, quoting the values you passed, the range they resolved to, and the zone they were read in. Two bounds that resolve to the *same* instant are rejected as a zero-length window and told so in those words, rather than being shown a range whose two ends print identically.

**Omitting one bound clamps the window to 31 days.** `startDate` alone runs a month forward from it, `endDate` alone a month back, and the response says which range it searched in a trailing `Note:`. That bound exists because of recurrence expansion, not tidiness: with `expand` on, the window is the range the server *generates occurrences over*, so an open-ended default asked Fastmail to materialise every occurrence of every repeating event through 2099 — and `limit` bounds none of that work, since the whole response is parsed and sorted before the list is trimmed. A window whose bounds you name yourself is never clamped, however long it is - the span is yours to choose. A bound that resolves past the ends of the four-digit-year range — the invented half, or one you named, since the local-day rule resolves your value through a zone and an offset alone can push `9999-12-31` over — saturates at that edge rather than running past it. So `startDate: "9999-12-30"` on its own is answered rather than erroring on an unrepresentable date, and a saturated bound **you** named is disclosed in the same trailing `Note:`. (`endDate` alone runs the other way and cannot saturate at that end.) Where saturation collapses a one-sided window to zero length, the call is rejected as `InvalidParams` rather than answered with nothing.

#### Calendar times: a name, never an offset

This server never puts a UTC offset in a `start`/`end`, and never asks you to compute one yourself. A calendar time comes back as a bare local wall clock (`2026-04-20T10:00:00` — no `Z`, no offset) plus, when it matters, a separate `timeZone` naming the IANA zone that wall clock is in — never as `2026-04-20T10:00:00+12:00`. An offset is only ever valid at the single instant it was computed for, and the instant an event actually happens is exactly what a stored value is not; naming the zone instead lets daylight saving be worked out from *your* zone database, at the moment you need the real instant, rather than baked in on read.

`timeZone` describes `start`:

| `start` is stored as | `timeZone` |
| --- | --- |
| A named zone (`TZID=...`) **matching** the account's configured zone | *(omitted)* |
| A named zone **differing** from the configured zone | the zone's IANA name, e.g. `"Pacific/Auckland"` |
| Floating — no `TZID` and no `Z` (RFC 5545 §3.3.5: a different instant for every reader) | `null` |
| `Z`-designated (a UTC instant), or date-only (all-day) | *(omitted)* |

`timeZone` is **omitted**, not set to the configured zone's own name, for the overwhelming majority of rows — most events are already in the account's configured zone (`FASTMAIL_TIMEZONE`/`USER_CONFIG_FASTMAIL_TIMEZONE`, falling back to the server's own zone — the same rule [Calendar window dates](#calendar-window-dates) states for window bounds), so read "no `timeZone`" as "the configured zone". A `Z`-designated or all-day value is also omitted rather than set to `null`: both already fully describe themselves, and `null` there would wrongly assert "floating".

`endTimeZone` describes `end`, but only **relative to `start`**: omitted whenever `end`'s zone matches `start`'s (the ordinary case), and otherwise set by the same rule as the table above. When `start` is absent entirely — no `DTSTART` at all — `end` is compared against the account's configured zone instead, the same fallback `timeZone` itself uses. RFC 5545 §3.8.5.3 and this server's own write-side check ([Start/end agreement](#startend-agreement), [#140](https://github.com/JonathanGodley/fastmail-mcp/issues/140)) both allow a start and an end in two different named zones — the ordinary reason is a flight, departing in one zone and landing in another. `endTimeZone` is also omitted when `end` is absent, or was computed from a stored `DURATION` rather than a real `DTEND`: a computed end shares `start`'s zone by construction, so there is nothing to disagree about.

Where the underlying value comes from: `list_calendar_events` always expands (there is no unwindowed call - see [Calendar window dates](#calendar-window-dates)), and expansion normalises most in-window *occurrences* to a UTC instant server-side, so those rows carry `Z` and no `timeZone`; an all-day row, and any row expansion did not touch, keeps whatever form the stored property had, judged by the same table. `get_calendar_event` is the call that returns the stored property exactly as written, because it fetches the resource itself rather than a window over it.

Non-IANA names pass straight through rather than being rejected or resolved. A calendar written by a non-Fastmail client can carry a Windows zone name (`AUS Eastern Standard Time`) in its `TZID`; this server reports it verbatim in `timeZone`/`endTimeZone` and sorts that event by the account's configured zone instead of that name (below), leaving interpreting the name itself to you.

**Sorting.** `list_calendar_events`' result order compares instants, not the text of the dates — a calendar mixes floating wall clocks, UTC instants and named zones, and comparing those spellings as strings put a genuinely earlier event after a later one, where `limit` then cut the earlier of the two. An event whose own `timeZone` names a zone this server can resolve now sorts in **that** zone; the account's configured zone is only the fallback, used for a genuinely floating `start` (there is no zone to sort it in, so your own clock is the least-wrong guess) and for a `timeZone` this server cannot resolve (a Windows name, passed through rather than rejected — see above).

#### Writing calendar times

`create_calendar_event` and `update_calendar_event` never write an offset either, and neither asks you to compute one: a `start`/`end` with no `Z` and no date-only marker is a bare wall clock, and `timeZone` names the IANA zone (e.g. `"Australia/Sydney"`) it's written in. `timeZone` only **qualifies** a designator-less `start`/`end` — the one shape with no zone of its own already. It cannot combine with a value that carries `Z`/an offset, or with a date-only value, since both already name themselves; either combination is rejected as `InvalidParams`, naming which argument to drop.

**`timeZone` must contain a region-qualifying slash** (`Continent/City`), with exactly one exception: `"UTC"` itself, matched case-insensitively. A bare abbreviation or alias - `"EST"`, `"NZ"`, `"PST"`, `"MST"`, `"GMT"`, `"Zulu"` and others like them - is rejected even though each one resolves to a real zone, because the resolution is ambiguous: `"EST"` names a fixed-offset zone with no daylight saving, **not** US Eastern, and the others are exactly as easy to misread. Write `"Pacific/Auckland"` rather than `"NZ"`, `"America/New_York"` rather than `"EST"`. There is no safe-list of "harmless" abbreviations - the rule is the same for all of them.

`timeZone` accepts a leading `/` on a TZID (RFC 5545 §3.2.19 - a zone registered by its own creator rather than plain IANA form): it's accepted and stripped for you, then checked against the slash rule above. What actually gets written is `timeZone`'s **canonical IANA spelling**, which can differ in case, or via a link, from what you passed - `"us/pacific"` is written as `"America/Los_Angeles"`, and `"AUSTRALIA/SYDNEY"` as `"Australia/Sydney"`. A non-IANA name a read can hand back - a Windows zone id such as `"AUS Eastern Standard Time"`, from an external invite - is rejected by `timeZone`; there is nothing to canonicalise it against. Omitting `timeZone` on `update_calendar_event` preserves the stored zone exactly as it is, which is the way to round-trip an event carrying one of those.

**Create and update treat an omitted `timeZone` differently, on purpose:**

- **`create_calendar_event`**: an omitted `timeZone` writes the account's configured zone (`FASTMAIL_TIMEZONE`/`USER_CONFIG_FASTMAIL_TIMEZONE`, falling back to the server's own zone — the same setting [Calendar window dates](#calendar-window-dates) and email `date` rendering use). This is a deliberate change from writing a bare floating value: a designator-less `start`/`end` used to be stored with no zone at all, a different instant for every reader, and now always resolves to a concrete zone instead.
- **`update_calendar_event`**: an omitted `timeZone` never defaults to the configured zone — it leaves `start`/`end` exactly as before this parameter existed. A designator-less value inherits the event's existing `TZID` when it has one, or stays floating when it doesn't.

**`timeZone` is rejected outright** (never silently ignored) in these cases, each naming the argument to drop:

- `null`, an empty string, or a whitespace-only string — this is a deliberate decision, not an oversight: neither tool can write a genuinely floating event through this parameter, ever. Note that a read returns `timeZone: null` for a floating event, so a read-modify-write call that echoes that straight back will hit this rejection; pass no `timeZone` at all to leave an already-floating value floating (`update_calendar_event` only - `create_calendar_event` has no existing value to leave alone, so an omitted `timeZone` there always writes the configured zone, per above).
- On `update_calendar_event`, `timeZone` with **neither** `start` nor `end` supplied — it has nothing to qualify. Still reachable: re-send `start` and/or `end` unchanged alongside `timeZone` to re-zone them.
- On `update_calendar_event`, `timeZone` with **only one** of `start`/`end` supplied, when the side you left alone is already stored in a *different* named zone. Re-zoning just one side would silently strand the other, producing a two-zone event nobody asked for. Pass **both** `start` and `end` (re-sending the untouched one unchanged) to change the zone, or omit `timeZone`.

**The response states what was actually written.** Because a caller-named zone, an inherited stored zone, and create's configured-zone default all arrive as the same "no designator" input, only the written result can say which one happened — so both tools' responses report the zone (or `UTC`/all-day/floating, as applicable) that ended up on `start` and `end`.

#### Repeating events cannot be changed or deleted here

`update_calendar_event` and `delete_calendar_event` **refuse** any `eventId` that resolves to a repeating event. There is no flag, parameter or confirmation that overrides this, and the error message says so - if you are looking for one, there isn't one. `get_calendar_event` is read-only and still works on the same id.

**What counts as repeating** is read off the stored record, not off the row you passed, and any one of four markers is enough: an `RRULE`; an `RDATE`, which is how a series **lists** its occurrences instead of stating a rule ([#162](https://github.com/JonathanGodley/fastmail-mcp/issues/162)); more than one `VEVENT` in the record, since one CalDAV resource is one UID; or a `RECURRENCE-ID` on any block, which makes that block an edited occurrence. The test fails **closed** on purpose - miscalling a one-off a series costs one edit in the web client, and the reverse costs a destroyed series.

The reason is that the create side cannot match it. `create_calendar_event` has no recurrence-rule parameter and no `RDATE` parameter, so this server cannot make a repeating event at all - which means it must not destroy or rewrite one it has no way to put back. Concretely:

- A **delete** removes the whole calendar record: every occurrence, past and future, irreversibly, and the server then mails a cancellation to every attendee. Nothing is echoed back that could rebuild it.
- An **update** patches the series master, which moves every occurrence. Where the series has occurrences that were edited individually, RFC 5545 does not settle whether those follow the master or stay where they are - so there is no correct answer to implement, and guessing would move somebody's calendar without saying which way.

Recurrence expansion made this reachable in a way it was not before. A series used to appear once, at its original start date, so "delete Thursday's meeting" found no row for Thursday. It now returns a plausible per-occurrence row carrying the **series** id (and the row's `url` is the same record again), so a caller acting on what looks like one occurrence was one step from destroying the lot.

What to do instead: use the Fastmail web interface, which offers "this event", "this and following", and "all events" properly. Restoring per-occurrence delete, whole-series delete with a recovery artifact, and a recurrence-rule parameter on create is tracked as [#146](https://github.com/JonathanGodley/fastmail-mcp/issues/146).

Single (non-repeating) events are unaffected: both tools work on them exactly as before.

#### Calendar known limitations

- **Recurring events, on read**: `list_calendar_events` always expands recurrences server-side - a caller naming no bounds is given the next month rather than an unwindowed listing ([#142](https://github.com/JonathanGodley/fastmail-mcp/issues/142)) - so it reports real occurrence dates; `get_calendar_event` fetches a single resource with no window and therefore always returns the series master at its original date. Rows are filtered **exactly** against the window you asked for ([#162](https://github.com/JonathanGodley/fastmail-mcp/issues/162)): a `Z`/offset value is the instant it names, an all-day value is the account's LOCAL day (a `DTSTART..DTEND` date span being the full multi-day span), and a wall clock resolves in the zone name the parsed event carries (`timeZone`/`endTimeZone`, [#139](https://github.com/JonathanGodley/fastmail-mcp/issues/139)) where that name resolves, and in the configured zone otherwise. The ±14 hours that used to be added here now widens the range **requested of the server** instead, at both edges, which is the only place it does any good: Fastmail matches an all-day value on its UTC day and reads a floating time as UTC, so a window narrower than a day could touch either kind of event without the server returning it at all - and no client-side filter can keep what was never sent. Two kinds of row can still sit outside the window. A block that still carries its own recurrence - **`RRULE` or `RDATE`** - is never dropped whatever its dates say, because an unexpanded master shows the series' original `DTSTART` and judging that date would turn a wrongly-dated row into a missing one; such a row carries `recurrenceRule` and/or `recurrenceDates` so you can see why it is there. And a **floating** timed event comes back from expansion stamped as `Z` with the floating marker destroyed, so nothing on this side can move it to your clock: it is judged on UTC and can land in the wrong day for an account far from UTC. That one is documented rather than fixed - there is no information left to fix it with. **The two residuals fail in opposite directions, and that matters for how you read a result.** A recurrence carrier only ever **adds** a row, so check each `start` against the window you asked for rather than assuming every row is inside it. A mis-judged floating event is the other way round: it is **missing** from the window it really belongs to and present in a neighbouring one - so on an account far from UTC, an empty result from `list_calendar_events` is **not** proof of a free day. Confirm a free day in the Fastmail web interface, or widen the window by a day at each edge - the mis-judgement is bounded by the account's offset, so widening surfaces the event and tells you the day is not provably free, but not where it actually sits: `start` is the value that is wrong. **A third case is not a row this filter mis-judges but a row that never arrives.** For a series that lists its occurrences as `RDATE`s and states no `RRULE`, Fastmail's own time-range filter indexes only `DTSTART..DTSTART+DURATION`, so a window covering one of the listed dates and not the series start matches the resource not at all, with expansion and without it - measured against an `RRULE` control series whose occurrence falls at the identical instant and *is* returned by the very same request ([#165](https://github.com/JonathanGodley/fastmail-mcp/issues/165)). No client-side filter can recover that row, because the server never sends it, and widening the window only surfaces it once the wider window reaches the series start. So an empty result is not proof of a free day even on an account sitting at UTC. Whether this server should compensate is open ([#167](https://github.com/JonathanGodley/fastmail-mcp/issues/167)). One residual ambiguity comes from the server: an expanded series' **first** occurrence carries no `RECURRENCE-ID` (Fastmail marks only instances after the first), so where a window holds that occurrence and no sibling, the entry cannot be told apart from a one-off and `isRecurring` is left off rather than guessed at. `get_calendar_event` on the same `id` settles it.
- **Recurring events, on write**: not supported at all — see below.
- **Attendee parameters**: RSVP, ROLE, CUTYPE and other attendee parameters are parsed on read but not settable on create/update — only `email` and `name` are accepted.

### Identity & Testing Tools

- **list_identities**: List sending identities (email addresses that can be used for sending). Returns simplified format by default, including the identity's configured `textSignature`/`htmlSignature` when it has one (see [Identity fields](#identity-fields) for how to use them, and [Signing a message](#signing-a-message) for the `appendSignature` flag that applies them for you).
  - Parameters: `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **check_function_availability**: Check which functions are available based on account permissions (includes setup guidance). Calendar tools run over CalDAV, so calendar is reported available only when CalDAV credentials are configured, regardless of the JMAP calendar capability. Contacts is reported available only when the session carries both the JMAP contacts capability and a primary account for it. That confirms **read** access only: a read-only contacts token reports exactly the same capability, so `create_contact`/`update_contact`/`delete_contact` are listed as available and only refuse when a write is attempted.
- **test_bulk_operations**: Safely test bulk operations with dry-run mode
  - Parameters: `dryRun` (default: true), `limit` (default: 3, max: 10)

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
├── url-validation.ts       # Host/scheme allowlist for every URL the bearer token reaches
├── coerce.ts               # Lenient input coercion, key-strictness, and path confinement
├── jmap-client.ts          # JMAP client wrapper
├── email-formatter.ts      # Simplified email format for AI consumption
├── response-formatters.ts  # Mailbox/identity/contact simplifiers and query formatters
├── field-projection.ts     # `fields` output projection for the email read tools
├── quote-strip.ts          # Quoted-history detection and removal for the read path
├── thread-handler.ts       # get_thread orchestration (bodies, size cap, signals)
├── mailbox-handler.ts      # list_mailboxes/create_mailbox behind an injected client
├── compose-handler.ts      # create_draft orchestration behind an injected client
├── reply-handler.ts        # reply_email orchestration behind an injected client
├── forward-handler.ts      # forward_email orchestration behind an injected client
├── edit-draft-handler.ts   # edit_draft orchestration behind an injected client
├── send-draft-handler.ts   # send_draft, and the thread-state marking it performs
├── reply-quote.ts          # Builds the reply quote, the forwarded-message block and the signature
├── identity.ts             # Which identity a message sends as, and the sign-off it carries
├── subject.ts              # The Re:/Fwd: subject derivation and its override
├── body-format.ts          # HTML as source of truth, and the derived text/plain fallback
├── inline-images.ts        # Embedded (cid:) image identity, vetting and reconciliation
├── compose-inline.ts       # The embedded-image checks shared by the compose tools
├── inline-notes.ts         # The wording used when an image is carried, demoted or refused
├── contacts-calendar.ts    # Contacts and calendar extensions
├── contact-card.ts         # Contact card algebra: label resolution and the per-entry merge
├── contacts-handler.ts     # create/update/delete_contact orchestration behind an injected client
├── ical-limits.ts          # Size bounds on calendar text, checked before serialization
└── caldav-client.ts        # CalDAV calendar client (the only calendar path)
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
