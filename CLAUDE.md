# Development Rules

@CONTRIBUTING.md

CONTRIBUTING.md carries the rules every contributor follows: documentation ships with the change, the simplified response format, JMAP property consistency and never silently dropping a promised field, the destroy-versus-create rule, the version-bump sites, build and test commands, unit-testable handlers, the Linux CI matrix, and where design rationale lives. What follows is guidance for working in this repo with Claude Code on top of that.

## Building and testing, in practice

A running MCP server keeps serving the build it started with, so test a change to server code by invoking that code directly through this repo's own CLI, tests or `scripts/mcp-harness.mjs`. Going through the connected MCP tools answers from the old process and makes a correct change look broken.

JMAP index reads (RFC 8620 §3.4): our `getMethodResult`/`getListResult` positional reads are safe only while `Email/get`, `Mailbox/get` and `Thread/get` each appear once per batch. Match responses by call id before generalising to a batch where a method could repeat or be reordered. When a resolver cannot complete, it follows the `attachMailboxInfo` model: non-throwing, never silent.

A live harness run is on-demand proof of the real external path (Fastmail's blob store cannot be meaningfully mocked), never the sole coverage. "Verified once, live" is not "tested going forward." Use `scripts/mcp-harness.mjs` rather than hand-writing a client.

A new pin is proved by a lever that turns it red at that assertion. A lever that trips an earlier assertion proves nothing about a later one. A claim about a set (every call site, every tool, every field) is proved by enumerating the set, never by sampling it.

A guard that reads source as text is proved red by hand, with three levers: delete the pinned sentence from one of the sites it is asserted at; re-word the shared constant so the pinned phrase no longer appears while the site still carries a description; delete a whole property block so the site is gone rather than wrong. Run each lever until the guard fails at the assertion aimed at that case; a lever that trips an earlier assertion, or the vacuity floor, proves nothing about the pin it was meant for. Say in the commit message which levers went red and where.

CI is the finish line, not the local suite. Nothing is required to pass before a push lands, so a red `main` is discovered rather than prevented: check the run after pushing (`gh run list --repo JonathanGodley/fastmail-mcp --limit 3`; `--commit $(git rev-parse HEAD)` narrows it to the push, and a short SHA there matches nothing).

## Review findings: where each disposition lands here

The global rule (every finding is fixed, tracked, consciously declined with a written home, or surfaced) applies; in this repo the homes are:

- Tracked → a fork GitHub issue (`--repo JonathanGodley/fastmail-mcp`), never a code `TODO`.
- Consciously declined → an in-code comment for a local call, or the `docs/*` rationale files for a cross-cutting/accepted residual (e.g. the inherent path-guard TOCTOU limit).

A comment earns its place by changing what the next reader would do. `attachMailboxInfo` is the model: its comment says why the resolver neither throws nor omits. State a fact once, at the definition, and reference it from call sites; a decision nobody challenged needs no defence beside it. How the code got here belongs in the commit message and the issue; anything spanning several tools belongs in `docs/`. `node scripts/comment-share.mjs <commit>` measures a change against the file's own comment density; a figure far above it is nearly always restatement or history, not domain complexity - `/tidy-comments` is the pass that cuts it.

**Fix it rather than filing it.** A small defect with an obvious fix is fixed in the change in front of you, or, where nothing is in flight, in the pass that found it; a build that touches an area owns the defects it finds there. A defect that would leave the merged build broken is never trackable: if it genuinely cannot be fixed now, the build does not merge and it goes to the operator. File a finding only when the work genuinely cannot land now: the fix is design work of its own (judged on the minimal fix, not the ideal one), it needs a decision with a real trade-off nobody has made, it is the residual half of a partly-fixed change, or the area is fenced off from the current session. A review that surfaces a question answers it or surfaces it to the operator; it does not file it. A symmetric case scoped out of a fix is a finding like any other and needs its own disposition. Upstream bookkeeping is exempt: the per-PR adoption issues and the `docs/upstream-sync.md` ledger are filed unfiltered.

## Parallel work: one worktree per concern, kept alive through review

One worktree per implementer, a lone agent included, because agents sharing one checkout share one git index and can commit each other's staged work. The main instance orchestrates: it partitions the work, routes findings, and merges. It does not implement.

Keep each worktree alive until its work is reviewed AND its findings are fixed. Merging as soon as a branch is feature-complete leaves review fixes nowhere to go but a shared checkout, where several agents edit the same files at once and the result is one diff that maps to no issue.

Route each finding to the worktree that owns it and RESUME that worktree's agent, which already holds the context for its area. Give it its list, and let it verify and commit in its own worktree.

**Land a branch by MERGING it, and sweep for the worktree afterwards.** Re-applying a branch's changes as fresh commits on `main` leaves the branch tip unreachable from `main` forever, so nothing will report the work as landed and the worktree can never be cleaned up on that evidence. When a branch is landed, remove its worktree and delete the branch in the same breath. `.gitignore` excludes `.claude/*`, where the harness puts an isolated agent's worktree, so a leftover never appears in `git status`; `git worktree list` is the only thing that shows them. Run it before ending a session.

Split commits by concern, so each commit maps to the issue it closes.

## Releasing

Releases live on the fork (`origin` = `JonathanGodley/fastmail-mcp`). Pass `--repo JonathanGodley/fastmail-mcp` on every release, tag, and issue command. Cut a release only when the user asks for it. The step-by-step checklist (with the outward steps grouped behind a checkpoint) is the `/release` skill (`.claude/skills/release/SKILL.md`); this section is the rationale behind it.

Prefer batching related changes into one release: every shipped change pays the documentation and version-bump tax (see CONTRIBUTING.md), so bundling a cluster of related work amortizes it. Exception: a safety or security fix warrants its own immediate release even when small.

Release notes AND the git tag annotation message are consumer-facing: describe each change and cite its public `#issue`; never use internal plan codenames. Match the style of the existing fork releases.

## Where design rationale lives

Per-feature behaviour rationale lives in the relevant fork GitHub issue, e.g. `edit_draft` coupling (#4), reply-quote sanitiser posture (#7), the html→text fallback reject rule (#15), faithful draft recreate (#16), attachment confinement (#1). Cross-cutting rationale lives in `docs/*.md`, listed in CONTRIBUTING.md. `docs/fastmail-action-availability.md` is the authority for any "what does Fastmail mean by this verb" question; extend it by measuring a view, never by inferring from a role's name. Live measurements are made with the probes in `scripts/probes/` (see its README).

## Working with upstream

`upstream` = `MadLlama25/fastmail-mcp` (the fork's base); `origin` = `JonathanGodley/fastmail-mcp`. `gh` resolves bare commands to the fork: `gh repo set-default JonathanGodley/fastmail-mcp` is stored in this checkout (`remote.origin.gh-resolved`). Pass `--repo MadLlama25/fastmail-mcp` only when upstream is deliberately the target, such as reading their PRs for the adopt issues below, and remember that ⛔ below forbids writing there.

Strategy. Track upstream by *generally merging it into the fork whenever that is doable*: a periodic mainline sync that re-bases the fork's differentiators (response simplification, the calendar work) on top of upstream's latest, supplemented by the fork's own fixes carried ahead of upstream as open PRs *against* upstream. Never block fork progress on upstream review.

Doing the sync itself: the method lives in `docs/upstream-sync.md`. The trigger is a release; the `/release` skill carries the drift check.

Adopting an upstream PR (their work → ours). File a fork issue for every open third-party upstream PR (every PR authored by someone other than us, excluding bot dependency bumps), one issue per PR, titled so it names the PR. Do NOT pre-filter by whether a PR looks worth carrying: the adopt-or-decline call is made *in the issue*. In the issue, capture what the PR adds and how it interacts with the fork's differentiators (especially response simplification: the fork trims the body from *output* but still *fetches* it, so "metadata-only / never-fetch" PRs are NOT redundant with us). Where the fork's structure has diverged, reimplement in the fork's style rather than cherry-pick verbatim. Link the upstream PR with the fully-qualified `MadLlama25/fastmail-mcp#NN` form (a bare `#NN` in a fork issue links to a fork issue).

Offering a fix back (our work → theirs). Any fix that addresses an upstream issue or a general bug (not a fork-differentiator feature) should be offered back as a focused, single-purpose PR once it lands and tests pass on the fork: cut a branch with just that fix (a `git worktree` off `upstream/main`), reference the issue it closes, and don't drag in fork-only changes. Fork-only differentiators are not auto-offered (issue #40).

⚠️ **Write the closing keyword in the fully-qualified `Closes MadLlama25/fastmail-mcp#NN` form, never a bare `Closes #NN`.** A bare number in a commit destined for upstream closes their issue when they merge it, and then closes this repository's unrelated issue of the same number the moment upstream's history is merged back here. The qualified form is inert coming back. For the same reason, a CLOSED/COMPLETED state on a fork issue is not trustworthy without checking that the closing commit is actually about it. The mechanical guard is tracked as #158.

**⛔ Never comment on an upstream PR or issue directly.** Drafting the text is fine; a human posts it. The fork's OWN issues are fine for Claude to open, comment on, and close. Close a fork issue as part of shipping its fix, but validate first: confirm the fix is complete, genuinely resolves the issue, and is pushed to `origin/main`, then close with a commit-citing comment. A tagged release is NOT a precondition for closing.

## Artifacts read as standalone work

No plan codenames, session jargon or AI-workflow meta in anything durable. GitHub artifacts cite the public `#issue`/PR, and the release-notes codename rule above is the same rule applied to tags and release bodies.
