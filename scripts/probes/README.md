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

## Inventory

| Probe | Covers |
| --- | --- |
| `inline-read.smoke.mjs` | Read surfacing: `isInline`/`cid` in `get_email`, raw purity, `get_email_attachments`, download by `cid:` (including an `@`-bearing cid), compact lists unchanged |
| `inline-author.smoke.mjs` | Authoring: cid embed + note, lenient cid spellings, text-only degrade, dangling-ref and bad-cid rejects, `[image]` text derivation. Needs `FASTMAIL_ATTACH_DIR` (the probe sets it to the OS temp dir) |
| `inline-quotecarry.smoke.mjs` | Reply/forward quote carry: minted `ii-...@inline.invalid` cids, keep-rebuild reuse, drop/degrade/exclusion notes, asAttachment untouched, `send_draft` transmit receipt |
| `foreign-draft-roundtrip.mjs` | Edit round-trip of a foreign-shape draft (`alternative[text, related[html, inline image]]`, `@`-bearing Content-ID): metadata edit, body-keep edit, ref-dropping edit |
| `probe-exact-instance.mjs` | Exact-instance thread-state marking on duplicated messages (see its header; needs `FASTMAIL_PROBE_TEST_ADDR`) |

`jmaplib.mjs` is a minimal raw-JMAP helper (session, Email/set, blob upload,
tiny PNG generator) used to build fixtures outside the server under test.
`probelib.mjs` holds the shared check harness and response parsing.
