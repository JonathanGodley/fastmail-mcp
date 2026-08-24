# Release triage - {YYYY-MM-DD}

Scope: everything raised in the "what is unreleased / would a release ship a known defect" review of
`main` at `{tip sha}`, {N} commits ahead of `{last tag}` (tagged {tag date}).

Repo health at the time of writing: {tree clean? in sync with `origin/main`? `npx tsc --noEmit`
result, `npm test` numbers, upstream drift count, version still `{old version}` at all sites (the
bump belongs to `/release`, so "unbumped" is expected, not a finding)}.

This file is a working snapshot, not a durable record. The durable homes are the GitHub issues
themselves; nothing here should be quoted back as authority once an issue moves.

## Column meanings

**Priority** - what it means for the release currently being considered.

- `P1` fix before cutting. A user hits it on the new surface and gets a wrong result with no signal.
- `P2` ship it, but say so in the release notes.
- `P3` backlog. A release does not make it worse and does not depend on it.

**Severity** - how bad the outcome is when it fires, independent of how likely.

- `High` silent wrong output, data loss, or an unbounded operation.
- `Med` visibly wrong or degraded, but the caller is told or can see it.
- `Low` cosmetic, narrow, or only reachable through an unusual path.

**Effort** - rough size with the reason, so the letter is checkable rather than asserted.

- `S` a few lines in one file, existing tests cover the shape.
- `M` one file plus a new test, or a small cross-file change.
- `L` a design decision has to be made first, or several files move together.

**Shipped** - whether the defective code is already in users' hands.

- `SHIPPED` present at `{last tag}`. A release now does not introduce it.
- `NEW` first reaches users in this release. This is the column that decides whether the release
  itself is the thing carrying the defect.

## P1 - fix before cutting

| ID | What it is | Sev | Effort | Shipped |
|---|---|---|---|---|
| {#NN} | {one sentence: the defect and its silent/loud character} | {Sev} | {Effort} | {Shipped} |

{One reasoned paragraph per row, under the table: what was verified, HOW it was verified (the code
read, the probe run, the tag-grep), and why each letter is what it is. When a row closes later,
strike its table text with `~~ ~~`, mark it **DONE {date}**, and append a dated closure note here -
never rewrite the original entry.}

## P2 - ships as known, needs release-note wording

| ID | What it is | Sev | Effort | Shipped |
|---|---|---|---|---|

{Reasoned paragraphs as above, plus: what the release notes must actually say for each row.}

## P3 - backlog, a release does not make these worse

| ID | What it is | Sev | Effort | Shipped |
|---|---|---|---|---|

{A line or two per row is enough here; flag any row where fixing would newly reject input the tool
accepts today - that is a behaviour change and the user's call, not a cleanup.}

## Open decisions, not defects

| ID | The question | Effort | Shipped |
|---|---|---|---|

{Design questions the release does not force. State the options and the trade-off; mark each as the
user's call. Do not resolve them here.}

## Corrections to the verbal briefing

{Anything the audit contradicted in what the user was told, stated plainly: what changed, what
evidence changed it, and whether it moves the recommendation. Omit the section only if there were
no corrections.}

## Recommended cut

{The concrete shape of the release this audit supports: which P1 rows block, what the notes must
say, what ships as known. As rows close, strike and annotate with a dated note (same convention as
the bands) rather than rewriting.}
