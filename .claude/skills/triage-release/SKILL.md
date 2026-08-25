---
name: triage-release
description: Pre-cut release audit for the fastmail-mcp fork. Answers "if we cut a release right now, would it ship anything we KNOW is defective?" — enumerates the unreleased delta since the last tag, re-derives every open issue's load-bearing claim from the code rather than the issue text, decides shipped-vs-new mechanically against the tag, and produces a banded triage table (P1 fix-first / P2 note-in-release / P3 backlog / open decisions) ending in a recommended cut. Read-only — changes no code, cuts nothing. Use before a large or feature release, or when the user asks what a release would ship or runs /triage-release.
---

# Release triage

One question drives everything: **if we cut a release right now, would it ship anything we KNOW is
defective?** The output is a decision document, not a fix list — the user decides the cut; this
skill makes the backlog decidable. It is read-only: no code changes, no version bump, no tag.
Cutting is `/release`, on its own explicit ask.

## 1. Baseline — what would ship

- Last tag: `git describe --tags --abbrev=0`. Delta: `git log <tag>..HEAD --oneline` and the count.
- Repo health snapshot, stated in the document: tree clean and in sync with `origin/main`;
  `npx tsc --noEmit`; `npm test` with actual numbers; upstream drift
  (`git log --oneline $(git merge-base HEAD upstream/main)..upstream/main`); version still at the
  old number at all sites (the bump belongs to `/release`, so "unbumped" is expected here, not a
  finding).

## 2. Candidates — everything that could make the answer "yes"

- Every **open fork issue**: `gh issue list --state open --limit 100` (this checkout defaults `gh`
  to the fork). The closed list is not swept — but see the #158 caveat: a CLOSED state here is not
  trustworthy without checking the closing commit is actually about it, so a closed issue that a
  delta commit *names* gets its closing commit read.
- Anything raised verbally this session or sitting in a prior triage document.
- New defects noticed while reading the delta itself.

## 3. Verdicts — from code, never from tracker text

Two rules, each of which caught real errors on the first run (23 Aug 2026):

- **Re-derive each issue's load-bearing claim from the code.** Read the actual gate, guard or
  format the issue describes; run a quick probe where reading is ambiguous. Issue text, issue
  status and any recorded verdict are candidates, not evidence - see `~/.claude/rules/working-with-jg.md`
  §Trusting the record.
- **Shipped-vs-NEW is decided mechanically against the tag, never inferred from where the work
  "belongs".** Does the defective code exist at the last tag? `git show <tag>:<file>` or
  `git grep -n "<symbol>" <tag> -- src/` versus the same at `HEAD`. `SHIPPED` means already in
  users' hands (a release now does not introduce it); `NEW` means this release is the thing that
  carries it — the column that decides whether a defect blocks the cut. On the first run this test
  overturned two verbal-briefing claims (one "new" defect predated the tag; one "reachable from
  create" defect was not reachable at all).

## 4. The document

Home: `.claude/triage/YYYY-MM-DD-release-triage.md` (`.claude/` is gitignored — a working snapshot
whose durable homes stay the GitHub issues; the document says so itself, and nothing in it is
quoted back as authority once an issue moves).

Start by copying `TEMPLATE.md` from this skill's directory and filling in its placeholders — it
carries the column meanings, the band structure and the per-section guidance. Two rules the
template states but cannot enforce: every effort letter and every number carries its reason, so
the value is checkable rather than asserted; and as rows close later, strike and annotate with a
dated note rather than rewriting history.

## 5. Hand back

Report to the user: the recommended cut; each call that is theirs (ship-as-known wording, holds,
anything changing a user-visible contract) with a recommendation apiece; and which findings were
filed as fork issues (a defect that blocks nothing still gets an issue — the triage file is not a
tracker). Then stop. The cut itself is `/release`, only when the user asks for it.
