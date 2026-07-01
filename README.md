# Fastmail MCP Server (Fork)

A fork of [MadLlama25/fastmail-mcp](https://github.com/MadLlama25/fastmail-mcp) — an MCP server for Fastmail's JMAP and CalDAV APIs.

This fork adds a **response simplification system** that reduces token usage when used with AI clients. All data-returning tools return a cleaned, curated format by default. Use `verbose` for all fields in the clean shape, or `raw` for the original JMAP response. See [Response Simplification](#response-simplification) for details.

## What this fork adds over upstream

- **Response simplification** — all data-returning tools return a token-lean shape by default (`verbose`/`raw` to opt out). See [Response Simplification](#response-simplification).
- **Calendar** — attendee/participant support, non-destructive event updates (no silent field wipes), and RFC 5545 date/TZID handling.
- **Local-time email dates** — the `date` field renders in your timezone with a UTC offset instead of raw UTC; `FASTMAIL_TIMEZONE` overrides the host zone.
- **Readable text/plain fallback** — when you supply only `htmlBody`, a plain-text alternative is generated automatically whenever one can be derived from the HTML (accessibility + deliverability). An image-only message with no derivable text still sends HTML-only; only a genuinely no-body send is rejected.
- **Faithful draft edits** — `edit_draft` preserves the draft's threading headers (In-Reply-To/References), attachments, and keywords across the immutable-email recreate, instead of silently dropping them.
- **Sending ergonomics** — drafts carry the identity's display name (parity with send); recipient strings like `"Name <email>"` are parsed across send/draft.
- **Outgoing attachments** — attach local files to a sent mail, draft, reply, or edited draft (append / remove-by-ref / clear-all). Disabled until you opt in by setting `FASTMAIL_ATTACH_DIR`, and confined to that directory (reading a local file to email it out is an exfiltration vector). See [Sending attachments](#sending-attachments).
- **Attachment paths** — relative `download_attachment` paths resolve inside the configured download dir, so a bare filename lands there in one step.

## Features

### Core Email Operations
- List mailboxes and get mailbox statistics
- List, search, and filter emails with advanced criteria
- Get specific emails by ID
- Send emails (text and HTML) with proper draft/sent handling
- Reply to emails with proper threading (In-Reply-To, References headers)
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
   # Only api.fastmail.com and www.fastmailusercontent.com are accepted by default.
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

All data-returning tools simplify responses by default to reduce token usage. Two optional parameters control how much data is returned:

- **Default** — a curated, cleaned response. Addresses are strings instead of objects, boolean flags replace keyword maps, null/empty fields are stripped, and only the most useful fields are included.
- **`verbose: true`** — all fields, still in the simplified shape. Use this when you need data the default omits (e.g. HTML body, mailbox permissions, contact addresses) without dealing with raw JMAP structures.
- **`raw: true`** — the original JMAP response with no transformation. Use this for debugging or when you need exact JMAP field names and structures.

### What each tool returns

| Tool | `verbose` | `raw` |
|------|-----------|-------|
| `get_email` | ✅ | ✅ |
| `list_emails`, `search_emails`, `get_recent_emails`, `get_thread` | — | ✅ |
| `list_mailboxes` | ✅ | ✅ |
| `list_identities` | ✅ | ✅ |
| `list_contacts`, `get_contact`, `search_contacts` | ✅ | ✅ |

Email list/search tools don't support `verbose` — they always return metadata and preview. Use `get_email` for full email content.

### Parameter validation

Unknown or misspelled parameters are rejected with an `InvalidParams` error that lists the offending key(s) and the valid keys for that tool (e.g. passing `folder` instead of `mailbox`, or the now-renamed `mailboxId`). This is deliberate: a silently dropped parameter would let a tool run with defaults and return confident but wrong results. Note this is key-strictness only; values are still coerced leniently (e.g. a stringified `"true"`/`"20"` for a boolean/number param is accepted).

Error codes follow the same recoverability logic: a failure you can fix by changing input (a bad/empty field, a not-found id, a non-sendable draft) returns `InvalidParams`, while a server/operational failure you can't fix that way (a server refusal, a missing system mailbox, a transport error) returns `InternalError` — so a client knows whether to re-form the call or simply retry. A refused mutation also carries the server's stated reason. The one exception is `download_attachment`, which returns `InternalError` for a bad id to avoid leaking attachment metadata. Full rule: `docs/conventions.md`.

### Email fields

**Default fields** (all email tools): `id`, `subject`, `from`, `date`, `threadId`, `messageId`, `references`, `to`, `cc`, `bcc`, `replyTo`, `inReplyTo`, `isRead`, `isReply`, `isFlagged`, `isDraft`, `isAnswered`, `isForwarded`, `mailboxes`, `roles`, `keywords`, `preview`, `hasAttachment`, `attachments`, `listUnsubscribe`, `blobId`, `size`

**`get_email` also includes**: `bodyText`, `bodyHtmlSize` (character count hint — HTML omitted by default)

**`get_email` with `verbose`**: adds `bodyHtml` (WARNING: can produce very large responses for marketing/rich emails — only use when HTML content is specifically needed)

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

### Mailbox fields

**Default**: `id`, `name`, `role`, `parentId`, `totalEmails`, `unreadEmails`, `totalThreads`, `unreadThreads`

**Verbose adds**: `myRights`, `sortOrder`, `isSubscribed`, `sort`, `autoLearn`, `autoPurge`, `purgeOlderThanDays`, `hidden`, `isCollapsed`, `identityRef`, `learnAsSpam`, `suppressDuplicates`, plus any other JMAP fields

Falsy `role` and `parentId` are stripped in default and verbose (use `raw` if you need `null` values).

### Identity fields

**Default**: `id`, `name`, `email`, `replyTo`, `mayDelete`

**Verbose adds**: `textSignature`, `htmlSignature`, `bcc`, `verificationState`, `showInCompose`, `saveSentToMailboxId`, `displayName`, `isAutoConfigured`, `enableExternalSMTP`, `server`, `port`, `ssl`, `addBccOnSMTP`, `saveOnSMTP`, `externalCredentialId`, `warnings`, `useForAutoReply`, `verificationCheckTime`, plus any other JMAP fields

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

> **Recipient format:** every recipient field (`to`/`cc`/`bcc`/`replyTo` on `send_email`, `reply_email`, `create_draft`, `edit_draft`) accepts each entry as either a bare address (`a@x.com`) or the RFC 5322 `"Name <email>"` form (`Alice <a@x.com>`), which is parsed into a display name + address. The SMTP envelope always uses the bare address.
>
> **Draft sender name:** drafts created or edited via `create_draft`/`edit_draft` now carry the sending identity's display name (matching `send_email`), so the From shows your name rather than a bare address.

- **list_mailboxes**: Get all mailboxes in your account
  - Parameters: `verbose` (optional, include all fields), `raw` (optional, return original JMAP response)
- **list_emails**: List recent emails across all mailboxes (or one, via `mailbox`). **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam` to include them); drafts are included (set `excludeDrafts` to omit them). When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing in Trash/Spam matched, and you need not re-search to check.
  - Parameters: `mailbox` (optional — id, role, or name; scoping to a mailbox ignores the default exclusion), `limit` (default: 20), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `raw` (optional, return original JMAP response)
- **get_email**: Get a specific email by ID. Returns plain text body with HTML omitted (bodyHtmlSize hint provided). Only use `verbose` if you specifically need the HTML body — it can be very large for marketing emails.
  - Parameters: `emailId` (required), `verbose` (optional, include HTML body — can be 50K+ chars for rich emails), `raw` (optional, return original JMAP response)
- **send_email**: Send an email (supports threading via optional `inReplyTo` and `references` headers)
  - Parameters: `to` (required array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailbox` (optional — id/role/name to save into, defaults to Drafts), `subject` (required), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **reply_email**: Reply to an existing email with proper threading headers (automatically builds In-Reply-To and References). Saves a draft by default; set `send=true` to transmit immediately. The original is **quoted by default** (attributed, top-posted, matching the web client with a portable quote-bar style); set `quoteOriginal=false` to omit it. Quoted HTML is reproduced **sanitised** (script/style/event handlers stripped; formatting and real `http(s)` images kept; inline `cid:` images omitted — see [#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)) and is re-sent under your From address. On `send=true` the original message is marked answered and read (best-effort; the success message reports it when the mark succeeds). (This tool returns a status string, not email data, so `raw`/simplification do not apply.)
  - Parameters: `originalEmailId` (required), `to` (optional array, defaults to the original sender — preserving their display name as `Name <email>`), `cc` (optional array), `bcc` (optional array), `from` (optional), `textBody` (optional), `htmlBody` (optional), `send` (optional boolean, default: false — saves a draft; set true to transmit), `quoteOriginal` (optional boolean, default: true), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **create_draft**: Create an email draft without sending. Supports threading headers for replies. Each call creates a new draft.
  - Parameters: `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `mailbox` (optional — id/role/name to save into, defaults to Drafts), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `inReplyTo` (optional array), `references` (optional array), `replyTo` (optional array), `attachments` (optional array — see [Sending attachments](#sending-attachments))
- **edit_draft**: Edit an existing draft email. Since JMAP emails are immutable, this creates a replacement draft and then deletes the old one (so the returned email ID is new). The edit preserves the draft's threading headers (In-Reply-To/References), attachments, and other keywords. On the rare failure where the replacement is created but the old copy can't be removed, you may be left with a duplicate draft rather than none. A draft containing inline (`cid:`) images, or a body part that isn't plain text or HTML, can't be preserved by editing and is **rejected** — recreate it instead (see [#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). Only fields you provide are changed; omit a field to leave it unchanged.
  - Parameters: `emailId` (required), `to` (optional array), `cc` (optional array), `bcc` (optional array), `from` (optional), `subject` (optional), `textBody` (optional), `htmlBody` (optional), `replyTo` (optional array), `clearFields` (optional array), `attachments` (optional array — **appends**), `removeAttachments` (optional array of blobId/name), `originalEmailId` (optional — for reply-quote preservation, below), `noQuote` (optional boolean)
  - **Attachments on edit.** `attachments` **appends** to the draft's existing attachments (they are kept). Remove specific ones with `removeAttachments` (pass the `blobId` from `get_email_attachments`; a unique attachment name also works — a ref matching nothing, or a name matching more than one, is rejected). Remove all with `clearFields:['attachments']`. Passing `attachments` together with `clearFields:['attachments']` is rejected as a conflict. See [Sending attachments](#sending-attachments).
  - **Empty values are rejected.** Passing an empty string or empty array for a field (e.g. `subject: ""`, `to: []`) is an error, not a silent clear — it's almost always an accidental clobber. To deliberately blank a field, name it in `clearFields`.
  - **`clearFields`**: list of field names to clear to empty/none. Allowed: `to`, `cc`, `bcc`, `replyTo`, `subject`, `textBody`, `htmlBody`, `attachments`. `from` cannot be cleared (a draft always has a sender, matching the Fastmail UI). You cannot both pass a field as a value and list it in `clearFields`. A cleared draft is still a valid draft; it just may not be sendable (e.g. with no recipients).
  - **The text body is an auto-managed fallback of the HTML.** Editing `htmlBody` alone **regenerates** `textBody` from the new HTML (so an html-alone edit discards any custom `textBody` the draft had). Editing `textBody` alone while `htmlBody` is present is **rejected** — it would not change what recipients render (the HTML), and the fallback is managed automatically; to change the message edit `htmlBody`, or supply both bodies to store a custom plain-text alternative. `clearFields:['textBody']` while `htmlBody` is present is **rejected** for the same reason; use `clearFields:['htmlBody']` to convert the draft to a plain-text email. A subject/recipient-only edit (no body written) leaves both bodies untouched. An edit that would leave the draft with no body at all is **rejected**.
  - **Reply-quote preservation (`originalEmailId` / `noQuote`).** A reply draft keeps the quoted original inside its body — in the `htmlBody` for an HTML draft, or as a `> `-prefixed block in the `textBody` for a text-only one. Editing the body of a reply draft that **still has its quoted original** would silently discard the quote, so the edit is **rejected** unless you say what to do: pass `originalEmailId` (the id of the message the draft replies to — *not* the draft's own `emailId`) to regenerate the body and **keep** the quote, or `noQuote: true` to deliberately **drop** it. The check is made on the draft's *existing* body (the quote this server generated), not on your new text, so it can't be bypassed by including quote-shaped content. It applies uniformly to HTML and text-only reply drafts. Two edits keep the quote automatically and need neither flag: a metadata-only edit (subject/recipients/attachments), and a plain-text conversion (`clearFields:['htmlBody']`, which leaves the `> ` text quote in place). Supplying `htmlBody` to a text-only reply draft converts it to HTML. The regenerated body re-quotes the named original from scratch (it never reassembles the stored quote). Each successive quote-keeping body edit must pass `originalEmailId` again.
    - **Cross-session keep.** In the same session you already have `originalEmailId` (you passed it to `reply_email`). For a draft saved earlier, the only thing the draft exposes is its `In-Reply-To` *Message-ID*, not the JMAP id `originalEmailId` needs — so resolve the original first with `search_emails` for that Message-ID (pass `includeTrash:true`/`includeSpam:true` so a filed-away original isn't hidden by the default exclusion), then pass its id as `originalEmailId`. (There is no shortcut that lets re-including the quote in your new body count as "keep": that would be bypassable, so the lookup is the deliberate keep path.)
    - **Limitation — reply drafts created by another client.** The guard recognizes the common machine-emitted quote shapes — `<blockquote type="cite">` (this server, Apple Mail, the Fastmail web UI) and Gmail's `class="gmail_quote"`, plus most clients' `> `-prefixed plain text — but not every shape (e.g. Outlook's `<div>`-based quoting). A reply draft authored in a client whose quote shape isn't recognized may have its quote dropped on a body edit **without** prompting for `originalEmailId`/`noQuote`. This is the widest edge of this protection — wider than the in-session cases above. If you're editing a reply draft you didn't create here, treat the quote as unprotected and pass `originalEmailId` explicitly to rebuild it.
- **send_draft**: Send an existing draft email. The draft must have recipients and a from address. Moves the email to the Sent folder. An **HTML-only draft with real content** (e.g. an image-only message) sends as-is — image-only/HTML-only mail is valid. Only a **genuinely empty body part** (e.g. a blank `htmlBody` alongside real text, which can happen for drafts created in other clients) is **rejected**: an empty `text/html` part renders blank and shadows a real `text/plain`, so the recipient would see nothing. Edit the draft to supply or clear that body first. (Drafts created by this server never carry an empty part; every send/draft path drops empty bodies on write.)
  - Parameters: `emailId` (required)
- **search_emails**: Email search. Free-text `query` matches subject/body/participants (plain words, **not** operator syntax — `from:alice@example.com` is matched literally; use the dedicated `from`/`to`/`cc`/`bcc`/`subject` params instead). All filters combine with AND. **Trash and Spam are excluded by default** (set `includeTrash`/`includeSpam`); drafts are included. `query` is optional — with no query it returns recent mail matching only the structural filters. When a Trash/Spam match is withheld, a trailing note reports how many — so no note means nothing in Trash/Spam matched, and you need not re-search to check.
  - Parameters: `query` (optional), `from` (optional), `to` (optional), `cc` (optional), `bcc` (optional), `subject` (optional), `hasAttachment` (optional), `isUnread` (optional), `isPinned` (optional), `mailbox` (optional — id/role/name; scoping ignores the default exclusion), `after` (optional ISO 8601), `before` (optional ISO 8601), `limit` (default: 20, max: 100), `ascending` (optional, oldest first), `excludeDrafts` (optional), `includeTrash` (optional), `includeSpam` (optional), `raw` (optional, return original JMAP response)
- **get_recent_emails**: Get the most recent emails from a single mailbox (defaults to Inbox), max 50. Inbox-only with no Trash/Spam/draft flags; for an all-folder view (with the default Trash/Spam exclusion) use `list_emails`.
  - Parameters: `limit` (default: 10, max: 50), `mailbox` (default: 'inbox' — id/role/name; e.g. `mailbox:"trash"` to read Trash directly), `ascending` (optional, oldest first), `raw` (optional, return original JMAP response)
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
- **get_thread**: Get all emails in a conversation thread. Returns metadata + preview for each email.
  - Parameters: `threadId` (required), `includeDrafts` (optional, include in-progress drafts), `raw` (optional, return original JMAP response)
  - Draft messages are **excluded by default** (an in-progress reply is noise when reading a conversation). When any are present, a trailing note reports **how many drafts are hidden** (so a draft reply you already started isn't missed) — the drafts themselves are not surfaced. Set `includeDrafts: true` to include them. (The note is on the simplified path only; `raw` output stays pure JSON.)

> **Draft handling is asymmetric by design.** `get_thread` excludes drafts by default while `search_emails`/`list_emails` include them: a draft reply is noise when reconstructing a conversation, but a search/list should still find everything you've written. `get_thread` does not hide them silently — its simplified output reports a hidden-draft count so you can tell a draft reply already exists (the `raw` output stays pure JSON, so a raw consumer passes `includeDrafts` itself). Drafts are identified by the `$draft` keyword (robust even if a draft is moved out of the Drafts mailbox), not by mailbox role. Use `includeDrafts` (get_thread) / `excludeDrafts` (search/list) to override either default.

### Sending attachments

`send_email`, `create_draft`, `reply_email`, and `edit_draft` accept an `attachments` array. Each item is an object:

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

- **list_identities**: List sending identities (email addresses that can be used for sending). Returns simplified format by default.
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
├── response-formatters.ts  # Mailbox/identity/contact simplifiers and query formatters
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
