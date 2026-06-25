---
name: release
description: Cut a fastmail-mcp fork release on origin (JonathanGodley/fastmail-mcp). An ordered checklist that bumps the version at all three sites, verifies clean, and groups the outward, hard-to-undo steps (commit, push, tag, GitHub release, issue-close) behind an explicit checkpoint. Only run when the user has explicitly asked to release this session.
---

# Cut a fork release

This encodes the **Releasing** and **Version** sections of `CLAUDE.md` as a runnable checklist. Releases live on the fork (`origin` = `JonathanGodley/fastmail-mcp`); `gh` defaults to the upstream `MadLlama25/fastmail-mcp`, so **every** release/tag/issue command needs `--repo JonathanGodley/fastmail-mcp`.

The dangerous steps (anything that pushes, publishes, or closes a public issue) are grouped AFTER the checkpoint in step 4. Do the verification steps first; do not cross the checkpoint until its precondition holds.

## 1. Preconditions

- **The user explicitly asked to release THIS session.** Releases are never automatic. If they did not ask, the job of this skill is to STOP and report the procedure — a speculative or dry-read invocation must never trigger a release.
- **Confirm what is bundled.** Prefer batching related changes: every shipped change pays the documentation + 3-file version-bump tax, so a cluster amortizes it.
- **Confirm the documentation tax was paid.** Each bundled change must already have shipped its README + tool-description (`src/index.ts`) updates (`CLAUDE.md` "Documentation is mandatory"). A release does not retroactively excuse a missed doc update — if one is outstanding, fix it before tagging.

## 2. Bump the version at all three sites

Named in `CLAUDE.md` **Version** — match by content, line numbers drift:
- `package.json`
- `manifest.json`
- the `Server` constructor in `src/index.ts`

## 3. Verify clean

Run all three:
- `npx tsc --noEmit`
- `npm test` — this **enforces version sync**: the `version sync` test asserts the three files match, so a missed bump fails the suite.
- `npm run build`

## 4. ⛔ CHECKPOINT — do not cross until this holds

- Restate the step-1 precondition: the user asked for a release *this session*. Require a fresh explicit yes/no immediately before the first push.
- Precheck the outward steps: `gh auth status` succeeds, and the new tag does **not** already exist on origin.
- **Honest residual:** this checklist is prose, not an enforceable interlock. The gate is procedural (precondition + post-checkpoint grouping + a push-time confirm), and that is the whole protection. Treat it as such.

## 5. Commit, then push (two steps — the push is the irreversible one)

Releases land on `main` directly (the fork's documented flow), not a feature branch.

1. **Reset the index first:** `git reset` — start from a known-empty staging area.
2. **Stage explicit per-FILE paths only.** The intended set is the version-bumped source files plus any doc/changelog. NEVER stage a directory (`git add <dir>/` re-honors a nested `.gitignore` that could re-include a secret), NEVER a glob, NEVER `-A` / `.` (build output like `dist/` is gitignored and must stay out).
3. **Assert the staged set is EXACTLY the intended files.** `git diff --cached --name-only`, sorted, must *equal* the intended list, sorted — an equality check, not a subset (a subset check misses a pre-staged extra file). This assertion, not the staging verb, is the durable defense against committing an unintended file under `.claude/`. It fails closed: an empty staged set means "nothing changed, investigate," never "force it."
4. Commit: `Release: vX.Y.Z-fork.N`.
5. **Only on a passing assertion**, push: `git push origin main`.

## 6. Tag (annotated) and push the tag

- `git tag -a vX.Y.Z-fork.N` with a **consumer-facing** message (annotated, not lightweight).
- `git push origin vX.Y.Z-fork.N`.

## 7. Publish the GitHub release

```
gh release create vX.Y.Z-fork.N --repo JonathanGodley/fastmail-mcp --verify-tag --notes-file -
```
- Notes are **consumer-facing**: describe each change, cite its public `#issue`, no internal codenames, match the style of prior fork releases.
- `--repo` goes adjacent to the command — a mis-defaulted `gh` would publish to upstream.
- `--verify-tag` is why step 6 pushes the tag first.

## 8. Close shipped fork issues (close-on-ship, validate first)

Closing a fork issue is part of shipping its fix — it is the default, not a separate ask. But **validate before closing**, both conditions:
- the release is live and the fix is genuinely present in the released build (carried in the built `dist/`), AND
- the issue is actually resolved by what shipped (don't infer from a bare `(#N)` commit mention — that does not auto-close; only `Closes/Fixes/Resolves #N` on a default-branch push does).

This applies to the **fork's own** issues only. Never close (or comment on) an UPSTREAM `MadLlama25/fastmail-mcp` issue/PR autonomously — draft any such text and let the user post it.

```
gh issue close <N> --repo JonathanGodley/fastmail-mcp --comment "<cite the release + commit>"
```

## 9. Post-release

- Reconnect note: connected MCP clients load `dist/index.js`, so they must reconnect to pick up the new build.
- Update the fork-status memory.

## Gotchas

- `gh` defaults to **upstream** → always pass `--repo JonathanGodley/fastmail-mcp`.
- A bare `(#N)` in a commit message does NOT auto-close an issue — only `Closes/Fixes/Resolves #N` on a default-branch push does.
- Tag must be **annotated** (`-a`), not lightweight.
- Release notes and the tag message must be **codename-free** and consumer-facing.
- **Batch unless asked otherwise** — bundle related work into one release to amortize the version + doc tax.
