---
name: triage-release
description: Pre-cut release audit for the fastmail-mcp fork. Answers "if we cut a release right now, would it ship anything we KNOW is defective?" — enumerates the unreleased delta since the last tag, re-derives every open issue's load-bearing claim from the code rather than the issue text, decides shipped-vs-new mechanically against the tag, and produces a banded triage table (P1 fix-first / P2 note-in-release / P3 backlog / open decisions) ending in a recommended cut. Read-only: changes no code, cuts nothing. Use before a large or feature release, or when the user asks what a release would ship or runs /triage-release.
---

# Release triage

One question drives everything: **if we cut a release right now, would it ship anything we KNOW is
defective?** The output is a decision document, not a fix list — the user decides the cut; this
skill makes the backlog decidable. It is read-only: no code changes, no version bump, no tag.
Cutting is `/release`, on its own explicit ask.

This procedure ran in full on 23 Aug 2026 and produced
`.claude/triage/2026-08-23-release-triage.md`; read that file as the worked example of every step
below.

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

Two rules, each of which caught real errors on the first run:

- **Re-derive each issue's load-bearing claim from the code.** Read the actual gate, guard or
  format the issue describes; run a quick probe where reading is ambiguous. Issue text, issue
  status and any recorded verdict are candidates, not evidence — a conclusion written down once
  gets quoted as fact while the test behind it never gets re-read.
- **Shipped-vs-NEW is decided mechanically against the tag, never inferred from where the work
  "belongs".** Does the defective code exist at the last tag? `git show <tag>:<file>` or
  `git grep -n "<symbol>" <tag> -- src/` versus the same at `HEAD`. `SHIPPED` means already in
  users' hands (a release now does not introduce it); `NEW` means this release is the thing that
  carries it — the column that decides whether a defect blocks the cut. On 23 Aug 2026 this test
  overturned two verbal-briefing claims (one "new" defect predated the tag; one "reachable from
  create" defect was not reachable at all).

## 4. The document

Home: `.claude/triage/YYYY-MM-DD-release-triage.md` (`.claude/` is gitignored — this is a working
snapshot whose durable homes stay the GitHub issues; the document says so itself, and nothing in it
is quoted back as authority once an issue moves).

Shape — follow the 2026-08-23 file:

- **Scope line** naming the audited tip and tag, and the commit count between them.
- **Repo health** at time of writing (step 1).
- **Column meanings**, stated in the document so the table is self-contained:
  - `Priority` — meaning for THIS release. `P1` fix before cutting (a user hits it on the new
    surface and gets a wrong result with no signal). `P2` ship it, but say so in the release
    notes. `P3` backlog; the release neither worsens nor depends on it.
  - `Severity` — how bad when it fires, independent of likelihood. `High` silent wrong output,
    data loss, or an unbounded operation. `Med` visibly wrong or degraded. `Low` cosmetic or
    narrow.
  - `Effort` — `S`/`M`/`L` **with the reason**, so the letter is checkable rather than asserted
    (`L` means a design decision has to be made first).
  - `Shipped` — `SHIPPED` at the last tag, or `NEW` (step 3's mechanical test).
- **Bands**: `P1 — fix before cutting`, `P2 — ships as known, needs release-note wording`,
  `P3 — backlog`, and `Open decisions, not defects` (design questions the release does not force,
  each marked as the user's call). One reasoned paragraph per row under each table — what was
  verified, how, and why the letters are what they are. Every number carries its reason.
- **Corrections to the verbal briefing**, if the audit contradicted anything the user was told —
  stated plainly, with what changed and whether it moves the recommendation.
- **Recommended cut** — the concrete shape of the release this audit supports: which P1s block,
  what the notes must say, what ships as known. As rows close later, strike and annotate with a
  dated note (the file's own convention) rather than rewriting history.

## 5. Hand back

Report to the user: the recommended cut; each call that is theirs (ship-as-known wording, holds,
anything changing a user-visible contract) with a recommendation apiece; and which findings were
filed as fork issues (a defect that blocks nothing still gets an issue — the triage file is not a
tracker). Then stop. The cut itself is `/release`, only when the user asks for it.
