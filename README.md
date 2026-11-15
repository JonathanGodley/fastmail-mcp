# Fastmail MCP Server (Unofficial Fork)

> **Not a Fastmail product.** This is a community project. It is **not affiliated with, endorsed by, or supported by Fastmail**, and nothing here comes from them. "Fastmail" is a trademark of Fastmail Pty Ltd, used only to say which APIs this server talks to. Run it at your own risk, under the terms of the [license](#license).

A fork of [MadLlama25/fastmail-mcp](https://github.com/MadLlama25/fastmail-mcp) — an MCP server for Fastmail's JMAP and CalDAV APIs.

## What this fork adds over upstream

- **Response simplification** — all data-returning tools return a token-lean shape by default (`verbose`/`raw` to opt out). See [Response Simplification](docs/reference/responses.md#response-simplification).
- **Draft-first compose**: every compose tool saves a draft, and **`send_draft` is the only tool that transmits a composed message** (upstream's `send_email` is not offered here), so every outgoing message exists as stored, inspectable bytes before it is sent ([#32](https://github.com/JonathanGodley/fastmail-mcp/issues/32), [#66](https://github.com/JonathanGodley/fastmail-mcp/issues/66)). **Calendar invitations still send mail**: naming attendees on an event emails them an invitation, and a later delete emails a cancellation. See [Calendar invitations send mail](docs/reference/calendar.md#calendar-invitations-send-mail) before gating on `send_draft` alone.
- **Calendar**: attendee/participant support, non-destructive event updates (no silent field wipes), and RFC 5545 date/TZID handling.
- **Local-time email dates** — the `date` field renders in your timezone with a UTC offset instead of raw UTC; `FASTMAIL_TIMEZONE` overrides the host zone. The same setting governs [calendar window dates](docs/reference/calendar.md#calendar-window-dates), where `create_calendar_event` defaults a designator-less time (see [Writing calendar times](docs/reference/calendar.md#writing-calendar-times)), and calendar sort order.
- **Readable text/plain fallback** — when you supply only `htmlBody`, a plain-text alternative is generated automatically whenever one can be derived from the HTML (accessibility + deliverability). A body with no derivable text ships HTML-only; only a genuinely no-body message is rejected. See [Message bodies](docs/reference/email.md#message-bodies) for what the derivation makes of an image.
- **Faithful, reversible draft edits** — `edit_draft` preserves the draft's threading headers (In-Reply-To/References), attachments, and keywords across the immutable-email recreate, instead of silently dropping them. The draft it replaces goes to **Trash rather than being destroyed**, and the result echoes back what that draft contained, so an edit made from an out-of-date copy is both visible and undoable ([#65](https://github.com/JonathanGodley/fastmail-mcp/issues/65)).
- **Compose ergonomics** — drafts carry a display name alongside the From address: state it yourself by passing `from` as `"Name <email>"`, or leave it to the name an edited draft already carries, and then the sending identity's; recipient strings like `"Name <email>"` are parsed across every compose tool.
- **Outgoing attachments, including embedded images** — attach a local file, a `blobId`, or one part of an existing message to a draft, reply, forward, or edited draft (append / remove-by-ref / clear-all), and give an item a `cid` to show it *inside* the message body instead of hanging it off the end ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). Each source is opt-in: local files need `FASTMAIL_ATTACH_DIR` and are confined to it, and the two in-account sources need `FASTMAIL_ALLOW_BLOB_ATTACH`. See [Sending attachments](docs/reference/email.md#sending-attachments).
- **Forwarding** — `draft_email`'s `forward` mode reproduces the original under the Fastmail-native forwarded-message block and **carries its attachments** (the official server-side forward carries none), with an `asAttachment` mode that attaches the whole original as a lossless `.eml`.
- **Quotes and forwards keep their pictures** — a reply's quote and an inline forward's block carry the images the original displayed, instead of leaving broken references or silently blank space ([#13](https://github.com/JonathanGodley/fastmail-mcp/issues/13)). This re-sends image data outward that you never attached and no `FASTMAIL_ATTACH_DIR` governs; leaving `{{quote}}` out of the body is the way to send none of it. See [Replying and forwarding with images](docs/reference/email.md#replying-and-forwarding-with-images).
- **Attachment paths** — relative `download_attachment` paths resolve inside the configured download dir, so a bare filename lands there in one step.
- **Path-aware mailbox naming** - every mailbox parameter accepts a root-anchored path (`Archive/2026/Receipts`) alongside an id, role, or name, `list_mailboxes` returns the path it accepts, and `create_mailbox` makes the folder you want to file into ([#27](https://github.com/JonathanGodley/fastmail-mcp/issues/27), [#48](https://github.com/JonathanGodley/fastmail-mcp/issues/48)). A duplicated folder name is reported as ambiguous with its candidate paths instead of resolving to whichever one came back first. See [Naming a mailbox](docs/reference/email.md#naming-a-mailbox).

## Features

### Core Email Operations
- List mailboxes (each with its full path) and get mailbox statistics; create new mailboxes
- List, search, and filter emails with advanced criteria
- Get specific emails by ID
- Compose emails (text and HTML) as drafts, sent with `send_draft`
- Reply to emails with proper threading (In-Reply-To, References headers), carrying the images the original displayed into the quote
- Forward emails with attachment carry and a lossless `.eml` (`asAttachment`) mode; an inline forward carries the original's body images too
- Create, edit, and send email drafts (with or without threading)
- Email management: mark read/unread, pin/unpin, delete, move between folders, archive

### Advanced Email Features
- **Attachment Handling**: List and download email attachments; attach local files or content already in the account to outgoing mail, including images embedded in the body (see [Sending attachments](docs/reference/email.md#sending-attachments)); replies and forwards carry the images the original displayed
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

### Identity & Account Management
- List available sending identities
- Account summary with comprehensive statistics

## Setup

### Prerequisites
- Node.js 20.12+ 
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

Configuration (the API token, environment variables and the settings file) is in [docs/reference/configuration.md](docs/reference/configuration.md).

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
  npx --yes github:JonathanGodley/fastmail-mcp@v1.13.4-fork.9 fastmail-mcp
```

The tool reference is in [docs/reference/tools.md](docs/reference/tools.md); response format, email, contacts, calendar, Desktop Extension install and embedding details are in [responses.md](docs/reference/responses.md), [email.md](docs/reference/email.md), [contacts.md](docs/reference/contacts.md), [calendar.md](docs/reference/calendar.md), [configuration.md](docs/reference/configuration.md#install-as-a-claude-desktop-extension-dxt) and [embedding.md](docs/reference/embedding.md).

## API Information

This server uses Fastmail's JMAP (JSON Meta Application Protocol) API.

### Inspired by Fastmail JMAP-Samples

Many features in this MCP server are inspired by the official [Fastmail JMAP-Samples](https://github.com/fastmail/JMAP-Samples) repository, including:
- Recent emails retrieval (based on top-ten example)

### Authentication
The server uses bearer token authentication with Fastmail's API.

## CalDAV Calendar Support

Calendar setup is in [docs/reference/configuration.md](docs/reference/configuration.md#setup); how calendar access works is in [docs/reference/calendar.md](docs/reference/calendar.md).

## License

MIT

## Contributing

Contributions are welcome. The development rules, including the documentation that ships with each change, are in [CONTRIBUTING.md](CONTRIBUTING.md).

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Ensure your API token is valid and has the necessary permissions
2. **Calendar/Contacts "Forbidden" Errors**: Use `check_function_availability` to see setup guidance

### Calendar/Contacts Not Working?

If calendar and contacts functions return "Forbidden" errors, this is likely due to:

1. **Calendar credentials**: Calendar runs over CalDAV, not the API token. Set `FASTMAIL_CALDAV_USERNAME` (your Fastmail address) and `FASTMAIL_CALDAV_PASSWORD` (a Fastmail app password with calendar (CalDAV) access)
2. **API Token Scope (contacts)**: Your API token may need contacts permissions enabled; a contacts write that comes back forbidden needs a token with read-write contacts access

**Solution**: Run `check_function_availability` for step-by-step setup guidance.

For more detailed error information, check the console output when running the server.

## Privacy & Security

- API tokens are stored encrypted by Claude Desktop when installed via the DXT and are never logged by this server.
- The server avoids logging raw errors and sensitive data (tokens, email addresses, identities, attachment names/blobIds) in error messages.
- Tool responses may include your email metadata/content by design (e.g., listing emails) but internal identifiers and credentials are not disclosed beyond what Fastmail returns for the requested data.
