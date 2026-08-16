# Live probes

On-demand live verification scripts for externally-observable behavior that unit
tests cannot prove (Fastmail's blob store, MIME assembly, keyword writes). They
run the built server (`dist/index.js`) over real JMAP against the configured
account, via `scripts/mcp-harness.mjs`.

These are **not** durable regression coverage - the unit suite is. A probe run
proves the real external path once, on demand (typically before a release or
after touching an area a probe covers). See CLAUDE.md "Testing".

## Running

1. `npm run build` (the server runs from `dist/`, not `src/`).
2. Run through the token launcher, which injects `FASTMAIL_API_TOKEN` from the
   local MCP client config into the child environment without printing it:

   ```
   python scripts/probes/run-probe.py inline-read.smoke.mjs
   ```

Each probe prints one PASS/FAIL line per check and exits non-zero on any
failure.

## Safety rules

- Probes create their own fixtures (messages/drafts in the configured account)
  and move every artifact to Trash before exiting, including on failure.
- One probe (`inline-quotecarry.smoke.mjs`) performs a single send-to-self to
  verify the transmit receipt; the sent and received copies are swept to Trash.
- Never print or persist token values; scripts reference env var names only.

### Creating a calendar event with a participant sends real mail

`create_calendar_event` with a non-empty `participants` is **not** a local
write. Adding an attendee makes the server's scheduling layer send that address
an iTIP meeting invitation, from the account under test, the moment the event is
created - and a `delete_calendar_event` afterwards sends the matching
cancellation. Neither goes through a compose tool, so "the probe never called
`send_draft`" does not mean the probe sent nothing.

That makes it the one live check here that reaches outside the account, and the
consequence is a real invitation in a real person's calendar if the address
belongs to one. Two ways to stay inside the account:

- Omit `participants` entirely when the point of the check is the CalDAV write
  path itself (creation, properties, parsing, deletion). It exercises everything
  except the scheduling hop.
- If the attendee handling is what you are verifying, address it into a domain
  that cannot receive mail. RFC 2606 reserves `example.com`, which publishes a
  null MX, so nothing is delivered and no stranger is contacted.

Note what the second option costs, because it is not free: the invitation is
still *sent*, the null MX merely refuses it, and the refusal arrives back in the
account as an "Undelivered Mail Returned to Sender" bounce - one for the
invitation and one for the cancellation. Those bounces are real messages that
outlive the probe; nothing sweeps them, because they arrive as ordinary inbound
mail rather than as an artifact the probe created and can track. Expect them,
and clear them by hand if you care.

## Inventory

| Probe | Covers |
| --- | --- |
| `inline-read.smoke.mjs` | Read surfacing: `isInline`/`cid` in `get_email`, raw purity, `get_email_attachments`, download by `cid:` (including an `@`-bearing cid), compact lists unchanged |
| `inline-author.smoke.mjs` | Authoring: cid embed + note, lenient cid spellings, text-only degrade, dangling-ref and bad-cid rejects, `[image]` text derivation. Needs `FASTMAIL_ATTACH_DIR` (the probe sets it to the OS temp dir) |
| `inline-quotecarry.smoke.mjs` | Reply/forward quote carry: minted `ii-...@inline.invalid` cids, keep-rebuild reuse, drop/degrade/exclusion notes, asAttachment untouched, `send_draft` transmit receipt |
| `foreign-draft-roundtrip.mjs` | Edit round-trip of a foreign-shape draft (`alternative[text, related[html, inline image]]`, `@`-bearing Content-ID): metadata edit, body-keep edit, ref-dropping edit |
| `probe-exact-instance.mjs` | Exact-instance thread-state marking on duplicated messages (see its header; needs `FASTMAIL_PROBE_TEST_ADDR`) |
| `archive-parity.smoke.mjs` | Archive semantics end to end: the `mailboxIds/` patch form is accepted, an Inbox+label message keeps its label without gaining Archive, an Inbox-only message reaches Archive, a message already out of the Inbox is untouched, a refusing role writes nothing, and a mixed batch sends both patch shapes in one `Email/set`. Creates and destroys its own label folder |

`jmaplib.mjs` is a minimal raw-JMAP helper (session, Email/set, blob upload,
tiny PNG generator) used to build fixtures outside the server under test.
`probelib.mjs` holds the shared check harness and response parsing.
