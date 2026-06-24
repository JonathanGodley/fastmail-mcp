# Development Rules

## Documentation is mandatory

Every change that modifies tool behavior, parameters, response format, or adds/removes features MUST update:

1. **Tool descriptions in `src/index.ts`** — the `description` and `inputSchema` that MCP clients see
2. **README.md** — the tool reference section and any relevant format/feature sections
3. **Both at the same time as the code change** — not as a follow-up

Do not mark work as complete until documentation is verified.

## Response format consistency

All email tools that return email data must use the simplified format from `src/email-formatter.ts`:
- `simplifyEmail()` for full emails and list items
- Empty/null/false fields are omitted to save tokens
- Unknown JMAP fields go to `_extra`
- Every tool returning email data must support `raw: true` to bypass simplification

## JMAP property consistency

All email list/search methods in `src/jmap-client.ts` must request the same set of Email/get properties. If you add a property to one, add it to all:
- `getEmails()`
- `getRecentEmails()`
- `searchEmails()`
- `advancedSearch()`
- `getThread()` (full mode) and `getEmailById()` request additional body properties and must be a superset of the list set, so that `raw: true` returns a complete JMAP response.

## Version

The version string lives in three places — update all when bumping:
- `package.json` (line 3)
- `manifest.json` (line 4)
- `src/index.ts` (Server constructor)

## Building

The MCP server runs from `dist/index.js`, not `src/`. After making changes, run `npm run build` to compile. Connected MCP clients will need to reconnect to pick up the new code.

## Testing

Run `npx tsc --noEmit` and `npm test` before committing. All tests must pass.

## Releasing

Releases live on the fork (`origin` = `JonathanGodley/fastmail-mcp`). `gh` defaults to the upstream `MadLlama25/fastmail-mcp`, so pass `--repo JonathanGodley/fastmail-mcp` on every release, tag, and issue command. Cut a release only when the user asks for it.

1. Bump the version (see **Version**), then verify clean: `npx tsc --noEmit`, `npm test`, `npm run build`.
2. Tag and publish the GitHub release on `origin`.
3. **The release notes AND the git tag annotation message must be consumer-facing** — describe each change and cite its public `#issue`; never use internal plan codenames (e.g. `B4`, `B7`). Match the style of the existing fork releases.

## Where design rationale lives

The *why* behind shipped behaviour lives in two places, split by scope. Look here before re-deriving a decision:

- **Per-feature behaviour rationale → the relevant GitHub issue** (fork repo `JonathanGodley/fastmail-mcp`). Why one tool behaves the way it does sits in that tool's closed issue, next to the work — e.g. `edit_draft` coupling (#4), reply-quote sanitiser posture (#7), the html→text fallback reject rule (#15), faithful draft recreate (#16), attachment confinement (#1).
- **Cross-cutting rationale → the `docs/*.md` files** (checked in, version-controlled). Facts and models that span multiple tools, or that are properties of the JMAP/Fastmail platform or the shared codebase:
  - `docs/email-bodies.md` — the body-format model (HTML as source of truth, text/plain as a derived fallback), the asymmetric `edit_draft` coupling, MIME-matched body extraction + the 12-cell edit matrix, destroy+recreate, and live-probed Fastmail body facts.
  - `docs/security-model.md` — path confinement for download/attachment (always-on, configurable scope, the read-vs-write guard distinction).
  - `docs/conventions.md` — lenient input coercion, the U+202F local-time trap, the `sanitizeForQuote` posture, and dependency/build gotchas.

The dividing line: **an issue explains why ONE tool behaves as it does; a docs file captures a fact or model spanning multiple tools, or a property of the platform/codebase.** When you add a durable decision, file it on the side of that line — don't leave it in a local scratch file.
