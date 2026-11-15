# Configuration reference

## Configuration

1. Get your Fastmail API token:
   - Log in to Fastmail web interface
   - Go to **Settings → Privacy & Security**, and in the **Connected apps & API tokens** section click **Manage API tokens**
   - Click **New API token**
   - Copy the generated token

2. Set environment variables:
   ```bash
   export FASTMAIL_API_TOKEN="your_api_token_here"
   # Optional: customize base URL (defaults to https://api.fastmail.com).
   # Only api.fastmail.com and www.fastmailusercontent.com are accepted, each with an
   # optional regional prefix as JMAP session discovery returns it. For a self-hosted
   # JMAP server also set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true ("true" or "1" only); it
   # widens the host allowlist and nothing else, and HTTPS is still required.
   export FASTMAIL_BASE_URL="https://api.fastmail.com"
   # Optional: customize attachment download directory (defaults to ~/Downloads/fastmail-mcp/).
   # download_attachment paths are confined to this directory; set it to the root
   # you want attachments saved under to write there directly in one step.
   export FASTMAIL_DOWNLOAD_DIR="/path/to/your/downloads"
   # Optional (opt-in): enable attaching LOCAL files to outgoing mail. Unset, every such
   # attempt fails. Files are confined to this directory, resolved independently of
   # FASTMAIL_DOWNLOAD_DIR (see "Sending attachments"). Restart the server after setting it.
   export FASTMAIL_ATTACH_DIR="/path/to/attachable/files"
   # Optional (opt-in, separate from the one above): allow attaching content ALREADY in
   # the account (a blobId, or one part of an existing message). FASTMAIL_ATTACH_DIR does
   # not govern these. Only "true" or "1" enables it. Restart the server after setting it.
   # See the security note in "Sending attachments".
   export FASTMAIL_ALLOW_BLOB_ATTACH="true"
   # Optional: IANA timezone (e.g. America/New_York) for email dates, calendar window
   # dates, and create_calendar_event's default zone. Defaults to the server host's zone.
   # If set, it must be a full IANA name (a region-qualifying slash, or UTC in any case);
   # a bare abbreviation, an alias or an unrecognised name stops the server from starting.
   export FASTMAIL_TIMEZONE="America/New_York"
   # Optional (opt-in, experimental): the snooze tools. They need a second setting,
   # FASTMAIL_WEB_COOKIE, which is a whole Fastmail web login rather than a scoped
   # token: read "Experimental: snooze" below before setting either. Left commented out
   # so that copying this block does not turn the tools on without their cookie.
   # export FASTMAIL_EXPERIMENTAL_SNOOZE="1"
   ```

3. Or put the same variables in a file. See [The settings file](#the-settings-file).

### The settings file

At startup the server reads `~/.fastmail-mcp/.env` (on Windows,
`%USERPROFILE%\.fastmail-mcp\.env`), if it exists:

```
FASTMAIL_API_TOKEN=your_api_token_here
FASTMAIL_TIMEZONE=America/New_York
```

- One `KEY=value` per line. Lines that are not `KEY=value` are ignored, and on some
  Node 20 releases they also stop the lines after them loading.
- Save it as UTF-8 without a byte-order mark. With a BOM the first line is lost,
  and a UTF-16 file (what `>` writes in Windows PowerShell 5.1) loads nothing,
  with no error. PowerShell 7 writes BOM-less UTF-8 by default.
- A variable already in the environment wins, even when it is blank. A client
  config that sets `FASTMAIL_API_TOKEN` to an empty string, or leaves a DXT
  `${user_config...}` placeholder in it, blocks the file's token.
- The file is read once, at startup. Restart the server after changing it.
- A missing file is fine. A file that exists but cannot be read stops the server
  with an error naming the path.
- It holds your API token, so make it readable only by you
  (`chmod 600 ~/.fastmail-mcp/.env`).

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
   - Fill in the settings it prompts for. These are the same settings the environment variables in [Setup](#configuration) carry, so the descriptions there apply; only the API token is required.

   | Prompt | Environment equivalent | Notes |
   |---|---|---|
   | Fastmail API Token | `FASTMAIL_API_TOKEN` | Required. Stored encrypted by Claude Desktop. |
   | Fastmail Base URL | `FASTMAIL_BASE_URL` | Leave as the default `https://api.fastmail.com` unless you run your own JMAP server. |
   | Attachment Download Directory | `FASTMAIL_DOWNLOAD_DIR` | Where `download_attachment` writes. Blank uses `~/Downloads/fastmail-mcp/`. |
   | Attachment Send Directory | `FASTMAIL_ATTACH_DIR` | Blank leaves **sending** attachments off entirely. Set it only if you want it. |
   | Allow Attaching Content Already In The Account | `FASTMAIL_ALLOW_BLOB_ATTACH` | A checkbox, **unchecked** by default. Ticking it lets outgoing mail attach a `blobId` or one part of an existing message. Separate gate from the send directory — see [Sending attachments](email.md#sending-attachments). |
   | Email Date Timezone | `FASTMAIL_TIMEZONE` | Blank uses the host's zone. Also governs calendar window dates and `create_calendar_event`'s default zone. A value that is not a full IANA zone name stops the server from starting; see [Writing calendar times](calendar.md#writing-calendar-times) for the rule. |
   | CalDAV Username / CalDAV App Password | `FASTMAIL_CALDAV_USERNAME` / `FASTMAIL_CALDAV_PASSWORD` | Both needed for the calendar tools; without them those tools report themselves unavailable and the rest are unaffected. |
   | Calendar Organizer Display Name | `FASTMAIL_CALDAV_DISPLAY_NAME` | Blank falls back to the CalDAV username. |

   `FASTMAIL_ALLOW_UNSAFE_BASE_URL` is deliberately not offered here — it decides where your API token may be sent, so it stays an environment variable you set on purpose. The [snooze](tools.md#experimental-snooze) settings, `FASTMAIL_EXPERIMENTAL_SNOOZE` and `FASTMAIL_WEB_COOKIE`, are not offered either: the cookie is a whole Fastmail web login, so it goes in the environment or `~/.fastmail-mcp/.env`.

3. Use any of the tools (e.g. `list_emails`).

## CalDAV setup

1. Create an app-specific password on Fastmail:
   - Go to **Settings → Privacy & Security**, and in the **Connected apps & API tokens** section click **Manage app passwords and access**
   - Create a new app password (you can name it "CalDAV MCP" or similar)

2. Set the following environment variables:
   ```bash
   export FASTMAIL_CALDAV_USERNAME="your-email@fastmail.com"
   export FASTMAIL_CALDAV_PASSWORD="your-app-specific-password"
   # Optional: display name for ORGANIZER when creating events with participants
   export FASTMAIL_CALDAV_DISPLAY_NAME="Your Name"
   ```

When these variables are set, all calendar tools are available. When they are not set, calendar tools will return an error with setup instructions.
