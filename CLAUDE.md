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

When you append an extra method call to an existing batch (e.g. a trailing `Mailbox/get` to resolve mailbox names), read its result **defensively** with `readListResultIfPresent`, not a hard index. `getMethodResult`/`getListResult` throw on a missing index, and existing tests stub only the original method responses, so a hard index would make them throw — and a real server that drops the trailing method would error in production. Degrade gracefully (empty result → feature simply absent) instead.

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

Prefer batching related changes into one release: every shipped change pays the documentation + 3-file version-bump tax (see **Version**), so bundling a cluster of related work amortizes it — for a two-line fix the tax is most of the work.

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

## Working with upstream

`upstream` = `MadLlama25/fastmail-mcp` (the fork's base); `origin` = `JonathanGodley/fastmail-mcp`. `gh` defaults to upstream, so pass `--repo JonathanGodley/fastmail-mcp` for every fork-side issue/PR/release command.

**Strategy.** Track upstream by *generally merging it into the fork whenever that is doable* — a periodic mainline sync that re-bases the fork's differentiators (response simplification, the calendar work) on top of upstream's latest. Supplement that baseline with the fork's own fixes carried ahead of upstream as open PRs *against* upstream. The upstream maintainer is intermittent (active in bursts, absent in between), so never block fork progress on their review: land and release on the fork, offer the general fixes back, and move on whether or not they respond.

**Adopting an upstream PR (their work → ours).** **File a fork issue for every open third-party upstream PR** — every PR authored by someone other than us, excluding bot dependency bumps — one issue per PR, titled so it names the PR. Do NOT pre-filter by whether a PR looks worth carrying: raising the issue is just bookkeeping, and the adopt-or-decline call is made *in the issue*, never by judging which PRs deserve one. In the issue, capture what the PR adds and how it interacts with the fork's differentiators (especially response simplification — the fork trims the body from *output* but still *fetches* it, so "metadata-only / never-fetch" PRs are NOT redundant with us). Where the fork's structure has diverged (simplification, `advanced_search`, descriptions), **reimplement in the fork's style rather than cherry-pick the patch verbatim.** Link the upstream PR with the fully-qualified `MadLlama25/fastmail-mcp#NN` form (a bare `#NN` in a fork issue links to a fork issue, not upstream). That reference also auto-publishes a backlink on the upstream PR's timeline, so our adoption tracking is visible upstream without us commenting on the thread.

**Offering a fix back (our work → theirs).** Any fix that addresses an upstream issue or a general bug (not a fork-differentiator feature) should be offered back as a focused, single-purpose PR once it lands and tests pass on the fork: cut a branch with just that fix (a `git worktree` off `upstream/main` keeps the fork's tree clean), reference the issue it closes, and don't drag in fork-only changes. Fork-only differentiators (the simplification system, etc.) are not auto-offered — upstream wants their own (issue #40).

**⛔ Never comment on an upstream PR or issue directly.** Drafting the text is fine; a human posts it. The rule is about writing in *someone else's* repo. The fork's OWN issues are fine for Claude to open, comment on, and close (close only when the user says so — see the `github-comments-human-only` memory).

## Artifacts read as standalone work

Everything durable — commit messages, issue/PR titles and bodies, issue/PR comments, **and code comments and test names** — must stand on its own to a reader with no access to our planning. No internal plan codenames or slice labels (`B4`, `B7`, `Slice 2`), no review/session jargon (`Karen`, plan names, "the law"), and no AI-workflow meta (e.g. "drafted for a human to post," "we don't comment on upstream threads"). Explain the change or the *why* on its own terms; in GitHub artifacts, cite the public `#issue`/PR it relates to. For code comments this governs *content*, not prose style — they still follow ordinary code-comment conventions; the point is that a future reader sees the reason, not a label they can't decode. (Generalizes the release-notes codename rule above; see the `artifacts-stand-on-own-merit` memory.)
