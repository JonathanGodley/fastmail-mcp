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
- `getEmails()`, `searchEmails()` and `getRecentEmails()` (all three run through the shared `runFilteredQuery` helper, which sets `EMAIL_PROPERTIES_COMPACT` once — so they stay in sync by construction; `getRecentEmails()` reaches it by delegating to `getEmails()` and assembles no batch of its own)
- `getThread()` (full mode) and `getEmailById()` request additional body properties and must be a superset of the list set, so that `raw: true` returns a complete JMAP response.

When you append an extra method call to an existing batch (e.g. a trailing `Mailbox/get` to resolve mailbox names), read its result **defensively** with `readListResultIfPresent`, not a hard index. `getMethodResult`/`getListResult` throw on a missing index, and existing tests stub only the original method responses, so a hard index would make them throw — and a real server that drops the trailing method would error in production. The tolerant read **stays**: a dropped trailing method is a benign degrade.

But "tolerant read" is **not** "silently drop the promised field." A read tool promises `mailboxes`/`roles`; if a mailbox id can't be resolved to a name, the id is surfaced in `unresolvedMailboxIds` (never omitted with no trace). That was the #53 bug — `readListResultIfPresent` returning `[]` made the resolver silently emit nothing. The resolver (`attachMailboxInfo`) is therefore **never-silent but non-throwing**: resolved → name/role, unresolved → raw id in `unresolvedMailboxIds`. See its in-code comment for the full reasoning (why not throw, why not omit).

**Never silently drop a promised field.** When a resolution/enrichment can't complete, surface the degradation explicitly (e.g. the `unresolvedMailboxIds` fallback) OR error on a genuine failure — never let a promised output field vanish with no trace. And never weaken a production behavior just to satisfy an under-stubbed test: fix the test instead.

**Index-read caveat (RFC 8620 §3.4).** A JMAP response returns one entry per method call, in order, but errors are `error` entries — not absences. Our `getMethodResult`/`getListResult` index reads are safe ONLY because `Email/get`/`Mailbox/get`/`Thread/get` are each single-response in our batches; match method responses by their **call-id**, not a positional index, before generalizing index reads to any batch where a method could appear more than once or be reordered.

## A destroy must not remove what this server cannot recreate

A tool that **irreversibly destroys** a record must refuse a record whose **kind** this server's create surface cannot produce. The recovery artifact a destroy hands back — `deletedCard`, or any other pre-write echo — is only as good as the create tool that would consume it: where nothing here can make that kind of record at all, the echo cannot rebuild it, so the destroy has no recovery path and is not offered. That is why `delete_contact` rejects a contact GROUP (`create_contact` has no `kind` and no `members` parameter), matching the refusal `update_contact` already had; both raise it through one shared message so they read as a single rule.

**Get the granularity right: the test is the record KIND, not the fields.** Almost every real record carries fields the create tool cannot set — a contact card's titles, organizations, photos — and refusing to destroy all of those would break the tool. A field the echo cannot rebuild is a documented limit (see the echo bound in `docs/conventions.md`); a *kind of thing* that cannot be made at all is a refusal. Write the in-code comment so the next reader applies the narrow reading, because that is the one that stays correct as the create surface grows.

This governs delete paths not yet written. When adding one, check it against the create surface first, and if the two disagree, either refuse the unmakeable kind or say plainly why the destroy is still safe (e.g. `delete_email` moves to Trash rather than destroying, so no content is lost to recreate — note that this is true of CONTENT and false of FILING: it writes `mailboxIds` whole-value, so a message's other labels are dropped and the Trash copy cannot show you what they were, per #123).

## Version

The version string lives in **three hand-edited sites plus a regenerated lockfile** — update all when bumping:
- `package.json` (line 3)
- `manifest.json` (line 4)
- `src/index.ts` (Server constructor)
- then `npm install --package-lock-only` to carry it into `package-lock.json` (which holds it in two places — the root and the `""` package entry). Never hand-edit the lock; it is regenerated. The `version sync` test asserts all four.

## Building

The MCP server runs from `dist/index.js`, not `src/`. After making changes, run `npm run build` to compile. Connected MCP clients will need to reconnect to pick up the new code. A running server keeps serving the build it started with, so test a change to server code by invoking that code directly through this repo's own CLI or test entry point - going through the MCP tools answers from the old process and makes a correct change look broken.

## Testing

Run `npx tsc --noEmit`, `npm run typecheck:tests` and `npm test` before committing. All tests must pass.

`npm test` compiles first, via a `pretest` script. That is not a convenience: `built-server.test.ts` spawns `dist/index.js` as a real external process and asserts in a `before` hook that the build exists and is newer than `src/`. A `before` hook that throws makes the test runner report its suite as **cancelled**, not failed - so testing without a build gives `fail 0` and a green-looking exit path, with the real reason buried in a stack trace. `pretest` removes that trap from the normal route; the assertion stays as the backstop for anyone invoking `tsx --test` on those files directly.

**Handler logic must be unit-testable.** The `index.ts` CallTool `switch` has no test harness, so logic left inline there can only be exercised by a live run — which is not durable regression protection. When a handler does more than trivially destructure-and-delegate (orchestration, branching, threading uploaded data into multiple paths), extract it into a function that takes an **injected client** and unit-test it with a mock. The reply path is the pattern: `composeReply(args, client, attachDir)` in `src/reply-handler.ts` takes a `ReplyClient` interface (which `JmapClient` satisfies structurally), so the attachment-threading branches are covered in `npm test` with no credentials and no network. The handler is then a thin result→text wrapper.

**A live harness is on-demand proof of the real external path, never the sole coverage.** The only thing that proves the real upload/send path is a raw JSON-RPC harness spawning `dist/index.js` against a live account (with `FASTMAIL_API_TOKEN` + `FASTMAIL_ATTACH_DIR`) — Fastmail's blob store can't be meaningfully mocked. Use it to verify externally-observable behavior (byte-identical round-trip, server accept/reject), but it is a manual check that runs once; it must not be the only thing testing logic that could be unit-tested. "Verified once, live" ≠ "tested going forward." The reusable harness lives at `scripts/mcp-harness.mjs` (`createClient({ env })` → `init`/`call`/`close`, matches responses by JSON-RPC `id`); use it rather than hand-writing a new client each time.

`scripts/mutate-commit.mjs <commit>` runs Stryker over only the lines that commit changed, against only the matching test files; a survived mutant is an untested line. Use it in place of hand-reverting a change to prove a test fails without it. Two limits are inherent to the scoping, not bugs: a line the commit did not touch is never mutated (a boundary condition just above a changed block is invisible to the run), and the test files are chosen by name — `src/X.ts` → `src/X*.test.ts` plus whatever tests the commit itself changed — so coverage that lives in a differently-named file reads as a survivor. Pass `--tests` to set the list by hand; that is also the escape hatch for the drift guards that read `src/*.ts` as text (`index-env.test.ts`, `readme-inventory.test.ts`), which see Stryker's instrumentation in the file and fail its initial run.

## Review findings: where each disposition lands here

The global rule (every finding is fixed, tracked, consciously declined with a written home, or surfaced) applies; in this repo the homes are:

- **Tracked** → a fork GitHub issue (`--repo JonathanGodley/fastmail-mcp`), never a code `TODO`.
- **Consciously declined** → an in-code comment for a local call, or the `docs/*` rationale files for a cross-cutting/accepted residual (e.g. the inherent path-guard TOCTOU limit).

The sharpest local example of a scope call that wasn't a disposition: #37 guarded a quote-dropping `htmlBody` edit but the symmetric `textBody`-on-a-text-only-reply-draft quote-drop was quietly scoped out, with a justification that was factually wrong.

## Parallel work: one worktree per concern, kept alive through review

**One worktree per implementer, a lone agent included** - the shared-index reason and the incident
are in `~/.claude/rules/agents.md`; the change it swallowed here was #133. The main instance
orchestrates: it partitions the work, routes findings, and merges. It does not implement.

**Keep each worktree alive until its work is reviewed AND its findings are fixed.** Merging as
soon as a branch is feature-complete is the mistake — review then finds defects and the fixes have
nowhere to go but a shared checkout, where several agents edit the same files at once and the
result is one diff that maps to no issue.

**Route each finding to the worktree that owns it and RESUME that worktree's agent.** It already
holds the context for its area: the decisions it made, what it tried, why the code is shaped the
way it is. A fresh agent on a shared tree has none of that and re-derives it badly. Give the
resumed agent its list, and let it verify and commit in its own worktree.

**Land a branch by MERGING it, and sweep for the worktree afterwards.** Re-applying a branch's
changes as fresh commits on `main` looks equivalent and is not: the content arrives, but the branch
tip stays unreachable from `main` forever, so nothing will ever report the work as landed and the
worktree can never be cleaned up on that evidence. When a branch is landed, remove its worktree and
delete the branch in the same breath. **Nothing will remind you** — `.gitignore` excludes
`.claude/*`, which is where the harness puts an isolated agent's worktree, so a leftover never
appears in `git status` and the harness itself only auto-removes a worktree that has no commits in
it. `git worktree list` is the only thing that shows them; run it before ending a session. (Three
worktrees from 21 Aug 2026 sat undetected for a day that way, each holding an orphaned twin of work
that had been re-committed onto `main` under a different hash.)

**Split commits by concern**, so each commit maps to the issue it closes.

## Releasing

Releases live on the fork (`origin` = `JonathanGodley/fastmail-mcp`). `gh` defaults to the upstream `MadLlama25/fastmail-mcp`, so pass `--repo JonathanGodley/fastmail-mcp` on every release, tag, and issue command. Cut a release only when the user asks for it. The step-by-step checklist (with the outward steps grouped behind a checkpoint) is the `/release` skill (`.claude/skills/release/SKILL.md`); this section is the rationale behind it.

Prefer batching related changes into one release: every shipped change pays the documentation + 3-file version-bump tax (see **Version**), so bundling a cluster of related work amortizes it — for a two-line fix the tax is most of the work. **Exception: a safety or security fix warrants its own immediate release even when small** — the value of getting the safer default in front of users now outweighs the amortization. Batching is the default for ordinary fixes/features, not for a fix whose whole point is reducing a footgun.

1. Bump the version (see **Version**), then verify clean: `npx tsc --noEmit`, `npm test`, `npm run build`.
2. Tag and publish the GitHub release on `origin`.
3. **The release notes AND the git tag annotation message must be consumer-facing** — describe each change and cite its public `#issue`; never use internal plan codenames (e.g. `B4`, `B7`). Match the style of the existing fork releases.

## Where design rationale lives

The *why* behind shipped behaviour lives in two places, split by scope. Look here before re-deriving a decision:

- **Per-feature behaviour rationale → the relevant GitHub issue** (fork repo `JonathanGodley/fastmail-mcp`). Why one tool behaves the way it does sits in that tool's closed issue, next to the work — e.g. `edit_draft` coupling (#4), reply-quote sanitiser posture (#7), the html→text fallback reject rule (#15), faithful draft recreate (#16), attachment confinement (#1).
- **Cross-cutting rationale → the `docs/*.md` files** (checked in, version-controlled). Facts and models that span multiple tools, or that are properties of the JMAP/Fastmail platform or the shared codebase:
  - `docs/email-bodies.md` — the body-format model (HTML as source of truth, text/plain as a derived fallback), the asymmetric `edit_draft` coupling, the identity signature in the body model (derived text form, placement above the quoted history, HTML-only preservation and the text path's idempotence, and every reason an append is reported as landing nowhere), MIME-matched body extraction + the 12-cell edit matrix, destroy+recreate, and live-probed Fastmail body facts.
  - `docs/security-model.md` — path confinement for download/attachment (always-on, configurable scope, the read-vs-write guard distinction).
  - `docs/fastmail-action-availability.md` — what the Fastmail client offers per screen, what each action actually does, and what the client's own calendar writes look like on the wire, measured rather than inferred. The authority for any "what does Fastmail mean by this verb" question, because Cyrus implements almost none of the policy and JMAP permits far more than the client does. Extend it by measuring a view, never by inferring from a role's name.
  - `docs/conventions.md` — lenient input coercion (and its fail-closed variant for arguments that narrow what a call touches), mailbox-query scoping (JMAP's singular `inMailbox`, the solely-in `inMailboxOtherThan`, the single unioned excluded set, and the two sites that decide whether the default Trash/Spam exclusion runs), the U+202F local-time trap, result serialisation (the two compact seams, and what the drift guard over them does and does not buy - it is not redaction), calendar window bounds (a date-only window is the caller's LOCAL day, and a one-sided window is bounded and says so), the two-pass quote sanitiser (its posture, and why a quote is rebuilt in two passes), and dependency/build gotchas.

The dividing line: **an issue explains why ONE tool behaves as it does; a docs file captures a fact or model spanning multiple tools, or a property of the platform/codebase.** When you add a durable decision, file it on the side of that line — don't leave it in a local scratch file.

## Working with upstream

`upstream` = `MadLlama25/fastmail-mcp` (the fork's base); `origin` = `JonathanGodley/fastmail-mcp`. `gh` resolves bare commands to the **fork**: `gh repo set-default JonathanGodley/fastmail-mcp` is stored in this checkout (`remote.origin.gh-resolved`), so `gh issue`/`gh pr`/`gh release` act on the fork with no flag. Pass `--repo MadLlama25/fastmail-mcp` only when upstream is deliberately the target, such as reading their PRs for the adopt issues below — and remember that ⛔ below forbids writing there. (Restore the untargeted behaviour with `gh repo set-default --unset` if this checkout's role ever changes.)

**Strategy.** Track upstream by *generally merging it into the fork whenever that is doable* — a periodic mainline sync that re-bases the fork's differentiators (response simplification, the calendar work) on top of upstream's latest. Supplement that baseline with the fork's own fixes carried ahead of upstream as open PRs *against* upstream. Never block fork progress on upstream review: land and release on the fork, offer the general fixes back, and move on whether or not they respond.

**Doing the sync itself.** The per-file resolution policy, why it is a real two-parent merge, the audit ledger that keeps an `--ours` resolution from silently dropping upstream work, and the list of surgery that must be re-done every time all live in `docs/upstream-sync.md`. **The trigger is a release** — the `/release` skill carries a drift check (`git log $(git merge-base HEAD upstream/main)..upstream/main`), because releases are the fork's natural cadence and a two-digit drift list is the moment to schedule a sync rather than discover one. The per-PR issue rule above and that drift check are the same mechanism from two ends: the issues catch incoming work one PR at a time, the drift check catches what merged while nobody was looking.

**Adopting an upstream PR (their work → ours).** **File a fork issue for every open third-party upstream PR** — every PR authored by someone other than us, excluding bot dependency bumps — one issue per PR, titled so it names the PR. Do NOT pre-filter by whether a PR looks worth carrying: raising the issue is just bookkeeping, and the adopt-or-decline call is made *in the issue*, never by judging which PRs deserve one. In the issue, capture what the PR adds and how it interacts with the fork's differentiators (especially response simplification — the fork trims the body from *output* but still *fetches* it, so "metadata-only / never-fetch" PRs are NOT redundant with us). Where the fork's structure has diverged (simplification, the consolidated `search_emails` with default Trash/Spam exclusion, descriptions), **reimplement in the fork's style rather than cherry-pick the patch verbatim.** Link the upstream PR with the fully-qualified `MadLlama25/fastmail-mcp#NN` form (a bare `#NN` in a fork issue links to a fork issue, not upstream). That reference also auto-publishes a backlink on the upstream PR's timeline, so our adoption tracking is visible upstream without us commenting on the thread.

**Offering a fix back (our work → theirs).** Any fix that addresses an upstream issue or a general bug (not a fork-differentiator feature) should be offered back as a focused, single-purpose PR once it lands and tests pass on the fork: cut a branch with just that fix (a `git worktree` off `upstream/main` keeps the fork's tree clean), reference the issue it closes, and don't drag in fork-only changes. Fork-only differentiators (the simplification system, etc.) are not auto-offered — upstream wants their own (issue #40).

⚠️ **Write the closing keyword in the fully-qualified `Closes MadLlama25/fastmail-mcp#NN` form, never a bare `Closes #NN`.** A bare number in a commit destined for upstream is written against *their* numbering. It closes their issue correctly when they merge it, and then closes **this** repository's unrelated issue of the same number the moment upstream's history is merged back here — stamped COMPLETED, with nothing to warn anyone. It has already fired: `d2b71da`, our own CalDAV-availability fix carried upstream as their PR #77, closed fork issue #76 (`.env` credential sourcing), which sat closed with none of its work done until someone went looking for the code. The qualified form closes upstream's issue exactly the same way and is inert coming back. Same rule and same reason as the fully-qualified linking above. The author of the closing commit is no guide (that one was ours) and the repository's auto-close setting governs merged linked PRs, not commit keywords, so nothing catches this today: **a CLOSED/COMPLETED state on a fork issue is not trustworthy without checking that the closing commit is actually about it.** The mechanical guard is tracked as #158.

**⛔ Never comment on an upstream PR or issue directly.** Drafting the text is fine; a human posts it. The rule is about writing in *someone else's* repo. The fork's OWN issues are fine for Claude to open, comment on, and close. **Close a fork issue as part of shipping its fix — but validate first:** confirm the fix is complete, genuinely resolves the issue, and is **pushed to `origin/main`** (carried in the built `dist/`), then close with a commit-citing comment. A *tagged release is NOT a precondition* for closing — the issue is resolved when the fix is done and on the mainline, not when a release happens to be cut (releases batch on their own cadence). If the fix is pushed with a `Closes #N` commit, GitHub auto-closes it on push; otherwise close it manually. (The general outbound rule is in the global CLAUDE.md: nothing is posted outside the repo without a per-message instruction.)

## Artifacts read as standalone work

The global rule applies (no plan codenames, session jargon or AI-workflow meta in anything durable). Here that specifically means: GitHub artifacts cite the public `#issue`/PR, and the release-notes codename rule above is the same rule applied to tags and release bodies.
