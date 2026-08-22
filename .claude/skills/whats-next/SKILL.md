---
name: whats-next
description: Triage what to work on next in fastmail-mcp. Scans the fork's open GitHub issues (JonathanGodley/fastmail-mcp) and local git state, then produces a prioritised briefing; add "upstream" to the invocation to also check upstream drift and third-party PRs. Use when the user asks what's next, what's urgent, or what needs doing, or runs /whats-next. Read-only by default.
argument-hint: "[optional focus, e.g. 'upstream' or an issue topic]"
---

# What's Next — fastmail-mcp edition

**First: invoke the `whats-next-core` skill and follow its procedure. If the
Skill tool is unavailable, Read `C:/Users/JG/.claude/skills/whats-next-core/SKILL.md`
in full instead.** It carries the shared triage discipline and the output
contract; the sections below are this repo's half — sources, tiers, judgement,
hygiene. Improvements to the shared discipline go in the core skill;
repo-specific lessons go here.

## Sources

1. **Fork issues** — `gh issue list --repo JonathanGodley/fastmail-mcp
   --state open --limit 200`. (⚠️ `gh` defaults to upstream in this clone; the
   `--repo` flag is mandatory on every command.) **Top tiers = bugs, and
   anything blocking the next release** — judged from content, since issues
   carry no priority labels: `gh issue view --repo JonathanGodley/fastmail-mcp
   <n> --comments` on every candidate before judging it. Never triage off a
   title.
2. **Local git state** — `git status`, current branch, unpushed commits.
   Uncommitted or unpushed work is a "Needs you" candidate.
3. **Upstream — only when the invocation asks** (e.g. `/whats-next upstream`):
   `git fetch upstream`, then the drift check from CLAUDE.md —
   `git log $(git merge-base HEAD upstream/main)..upstream/main --oneline` —
   and `gh pr list --repo MadLlama25/fastmail-mcp` for third-party PRs. Report
   the drift size and any third-party PR with no fork issue yet (the per-PR
   issue rule in CLAUDE.md). Recommend; don't merge, sync, or file anything.

**Out of scope by default** (still named in the "Not scanned" line): upstream
drift and PRs (when not asked for), the fork's own PRs offered against
upstream, and post-release feedback channels.

## Judgement (repo-specific)

- **Nothing is ever a "chase" against the upstream maintainer** — they are
  active in bursts and absent in between, and the fork deliberately never
  blocks on their review. Upstream work is scheduled by the fork's release
  cadence (the drift check runs at release), not by silence.
- **The issues are the tracker AND the rationale record.** An old open issue
  may be deliberately parked with a stated decision inside it — read for one
  before ranking it urgent. A closed issue is settled; it is not a candidate.

## Hygiene

- An open issue whose fix has plainly shipped (named in a release or merged
  commit) — offer to close it with a comment citing the release; never close
  unasked.
- On an upstream pass: a third-party upstream PR with no matching fork issue —
  list them; filing the issues is offered, not done.
