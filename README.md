# Fastmail MCP Server (Unofficial Fork)

> **Vibe coded. Run it at your own risk.**
> I made this primarily as a tool for myself, and I use it almost every day. I'm happy to share it, but I'll fully admit it is into the realms of a vibe-coded product, and I have relied heavily on autonomous AI coding and review for significant chunks of it. I am a software engineer and believe I am steering it competently. I take pains to use existing libraries, match good practices, and not blindly trust the first thing Claude says, but I have most certainly not reviewed this in as much depth as I would give code in my day job. It acts on your real mailbox, and it can send mail. I'm reasonably confident there are no significant issues, but I make no warranty or claims about its quality or adherence to standards; see [LICENSE](LICENSE).

> **Not a Fastmail product.** This is a community project. It is **not affiliated with, endorsed by, or supported by Fastmail**, and nothing here comes from them. "Fastmail" is a trademark of Fastmail Pty Ltd, used only to say which APIs this server talks to. Run it at your own risk, under the terms of the [license](#license).

A fork of [MadLlama25/fastmail-mcp](https://github.com/MadLlama25/fastmail-mcp) — an MCP server for Fastmail's JMAP and CalDAV APIs.

## What this fork adds over upstream

- **Response simplification**: a token-lean shape by default, with `verbose`/`raw` to opt out. See [Response Simplification](docs/reference/responses.md#response-simplification).
- **Draft-first compose**: `send_draft` is the only tool that transmits a composed message, but calendar invitations still send mail. See [Calendar invitations send mail](docs/reference/calendar.md#calendar-invitations-send-mail).
- **Calendar**: attendees, non-destructive event updates, and RFC 5545 date/TZID handling. See [docs/reference/calendar.md](docs/reference/calendar.md).
- **Local-time email dates**: `date` is in your timezone with a UTC offset, set by `FASTMAIL_TIMEZONE`. See [Email fields](docs/reference/responses.md#email-fields).
- **Readable text/plain fallback**: an `htmlBody` on its own gets a plain-text part derived from it when one can be derived. See [Message bodies](docs/reference/email.md#message-bodies).
- **Faithful draft edits**: `edit_draft` keeps threading headers, attachments and keywords across the recreate. See [Email Tools](docs/reference/tools.md#email-tools).
- **Compose ergonomics**: `"Name <email>"` is accepted for `from` and for every recipient field. See [Email Tools](docs/reference/tools.md#email-tools).
- **Outgoing attachments, including embedded images**: local files, `blobId`s or parts of a message, each source opt-in. See [Sending attachments](docs/reference/email.md#sending-attachments).
- **Forwarding**: carries the original's attachments, or attaches the whole original as a `.eml`. See [Email Tools](docs/reference/tools.md#email-tools).
- **Quotes and forwards keep their pictures**: images the original displayed are sent on, though you never attached them. See [Replying and forwarding with images](docs/reference/email.md#replying-and-forwarding-with-images).
- **Attachment paths**: a relative `download_attachment` path lands inside the download directory. See [Advanced Email Features](docs/reference/tools.md#advanced-email-features).
- **Path-aware mailbox naming**: every mailbox parameter accepts a path such as `Archive/2026/Receipts`. See [Naming a mailbox](docs/reference/email.md#naming-a-mailbox).

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

### Running the Server

Start the MCP server:
```bash
npm start
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

### Configuration

Get your Fastmail API token:
- Log in to Fastmail web interface
- Go to **Settings → Privacy & Security**, and in the **Connected apps & API tokens** section click **Manage API tokens**
- Click **New API token**
- Copy the generated token

Then set the environment variables:

| Variable | Required | Meaning |
|---|---|---|
| `FASTMAIL_API_TOKEN` | Required | Your Fastmail API token. |
| `FASTMAIL_BASE_URL` | Optional | JMAP base URL. Defaults to `https://api.fastmail.com`; only Fastmail's API hosts are accepted. |
| `FASTMAIL_ALLOW_UNSAFE_BASE_URL` | Optional | `true` or `1` widens the host allowlist for a self-hosted JMAP server. HTTPS is still required. |
| `FASTMAIL_DOWNLOAD_DIR` | Optional | Where `download_attachment` saves, confined to this directory. Defaults to `~/Downloads/fastmail-mcp/`. |
| `FASTMAIL_ATTACH_DIR` | Optional | Turns on attaching local files to outgoing mail, confined to this directory. |
| `FASTMAIL_ALLOW_BLOB_ATTACH` | Optional | `true` or `1` allows attaching content already in the account (a `blobId`, or one part of a message). |
| `FASTMAIL_TIMEZONE` | Optional | Full IANA zone name for email dates and calendar times. Defaults to the host's zone; an invalid name stops the server starting. |
| `FASTMAIL_CALDAV_USERNAME` | For calendar tools | Your Fastmail address. See [CalDAV setup](docs/reference/configuration.md#caldav-setup). |
| `FASTMAIL_CALDAV_PASSWORD` | For calendar tools | A Fastmail app password with CalDAV access. |
| `FASTMAIL_CALDAV_DISPLAY_NAME` | Optional | Organizer name on events you create with participants. Defaults to the CalDAV username. |
| `FASTMAIL_EXPERIMENTAL_SNOOZE` | Optional | `1` turns on the experimental snooze tools. |
| `FASTMAIL_WEB_COOKIE` | For snooze tools | The session cookie of a Fastmail web login. Read [Experimental: snooze](docs/reference/tools.md#experimental-snooze) before setting it. |

The same variables can go in `~/.fastmail-mcp/.env` instead; see [The settings file](docs/reference/configuration.md#the-settings-file). To install as a Claude Desktop Extension, see [Install as a Claude Desktop Extension (DXT)](docs/reference/configuration.md#install-as-a-claude-desktop-extension-dxt).

If the calendar or contacts tools report themselves unavailable or forbidden, run `check_function_availability` for setup guidance.

## Tools

The server has 42 tools. Each one's parameters and behaviour are in [docs/reference/tools.md](docs/reference/tools.md).

| Group | Tools |
|---|---|
| Mail reading | `list_emails`, `search_emails`, `get_email`, `get_thread` |
| Composing and sending | `draft_email`, `edit_draft`, `send_draft` |
| Filing | `list_mailboxes`, `create_mailbox`, `delete_mailbox`, `move_email`, `archive_email`, `add_labels`, `remove_labels`, `mark_email_read`, `pin_email`, `delete_email` |
| Bulk | `bulk_mark_read`, `bulk_pin`, `bulk_move`, `bulk_delete`, `bulk_add_labels`, `bulk_remove_labels` |
| Attachments | `get_email_attachments`, `download_attachment` |
| Contacts | `list_contacts`, `get_contact`, `search_contacts`, `create_contact`, `update_contact`, `delete_contact` |
| Calendar | `list_calendars`, `list_calendar_events`, `get_calendar_event`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event` |
| Account and diagnostics | `list_identities`, `get_mailbox_stats`, `get_account_summary`, `check_function_availability`, `test_bulk_operations` |

The experimental snooze tools are off by default and not counted above. See [Experimental: snooze](docs/reference/tools.md#experimental-snooze).

## Safety notes

- **Calendar invitations send mail.** Naming attendees on an event emails them an invitation, and a later delete emails them a cancellation, with no `send_draft` call. Gating `send_draft` does not gate all outbound mail; see [Calendar invitations send mail](docs/reference/calendar.md#calendar-invitations-send-mail).
- **The snooze tools run on a whole web login.** `FASTMAIL_WEB_COOKIE` is the session cookie of a logged-in Fastmail web login, not a scoped token: it carries everything that login can do. See [Experimental: snooze](docs/reference/tools.md#experimental-snooze).

## Reference

- [tools.md](docs/reference/tools.md): every tool's parameters and behaviour.
- [responses.md](docs/reference/responses.md): the simplified response shape, field projection, paging, errors and rate limits.
- [email.md](docs/reference/email.md): mailboxes, identities, archiving, replies, message bodies, signatures and attachments.
- [contacts.md](docs/reference/contacts.md): contact fields and how contact writes merge.
- [calendar.md](docs/reference/calendar.md): invitations, event ids, dates and time zones, free/busy and repeating events.
- [configuration.md](docs/reference/configuration.md): environment variables, the settings file, the Desktop Extension and CalDAV setup.
- [embedding.md](docs/reference/embedding.md): spawning the server from your own program.

The design rationale files for maintainers are listed in [CONTRIBUTING.md](CONTRIBUTING.md#where-design-rationale-lives).

## Privacy & Security

- API tokens are stored encrypted by Claude Desktop when installed via the DXT and are never logged by this server.
- The server avoids logging raw errors and sensitive data (tokens, email addresses, identities, attachment names/blobIds) in error messages.
- Tool responses may include your email metadata/content by design (e.g., listing emails) but internal identifiers and credentials are not disclosed beyond what Fastmail returns for the requested data.

## License

MIT

## Contributing

Contributions are welcome. The development rules, including the documentation that ships with each change, are in [CONTRIBUTING.md](CONTRIBUTING.md).
