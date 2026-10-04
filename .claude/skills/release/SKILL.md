---
name: release
description: Cut a fastmail-mcp fork release on origin (JonathanGodley/fastmail-mcp). An ordered checklist that bumps the version at every site, verifies clean, drafts the notes, and groups the outward, hard-to-undo steps (commit, push, tag, GitHub release, issue-close) behind an explicit checkpoint. Only run when the user has explicitly asked to release this session.
---

# Cut a fork release

This encodes the **Releasing** section of `CLAUDE.md` and **Bumping the version** in `CONTRIBUTING.md` as a runnable checklist.

The dangerous steps (anything that pushes, publishes, or closes a public issue) are grouped AFTER the checkpoint in step 5. Do the verification steps first; do not cross the checkpoint until its precondition holds.

## 1. Preconditions

- **The user explicitly asked to release THIS session.** Releases are never automatic. If they did not ask, the job of this skill is to STOP and report the procedure — a speculative or dry-read invocation must never trigger a release.
- **Large or feature release? Run `/triage-release` first, strongly suggested,** before bundling more than a handful of commits or any new feature surface; a two-line fix release does not need it.
- **Confirm what is bundled** (batching, and its security-fix exception: `CLAUDE.md` Releasing).
- **Check upstream drift.** `git fetch upstream && git log --oneline $(git merge-base HEAD upstream/main)..upstream/main`. A two-digit list means schedule a sync — the method is `docs/upstream-sync.md`. This check never blocks a release; it just makes the drift visible while someone is looking.
- **Confirm the documentation tax was paid.** Each bundled change must already have shipped its README + tool-description (`src/index.ts`) updates (`CONTRIBUTING.md` "Documentation ships with the change"); fix any outstanding one before tagging.

## 2. Bump the version: three hand-edited sites, the lockfile and the README pin

Named in `CONTRIBUTING.md` **Bumping the version**; match by content, line numbers drift:
- `package.json`
- `manifest.json`
- the `Server` constructor in `src/index.ts`

Then `npm install --package-lock-only`; never hand-edit the lockfile.

Then bump the **README `npx` pin example** (`github:JonathanGodley/fastmail-mcp@v…`) to the tag this release will push; the `version sync` test checks it.

## 3. Verify clean

Run all four:
- `npx tsc --noEmit`
- `npm run typecheck:tests`
- `npm test` (builds first via `pretest`; its `version sync` test fails a missed bump, including the README pin)
- `npm run scan:secrets`: "Build and Release DXT" runs it before packing, so a failure after the tag ships a release with no `.dxt`. Triage a hit per `CONTRIBUTING.md` "Triaging a hit".

## 4. Draft the release notes and the tag message

Delegate the draft to one opus subagent, read-only on the repo and GitHub, writing only two scratchpad files: the notes and the tag message. Its brief carries:
- the last tag, the commit being released, and the prior two fork releases to match in voice and shape (`gh release view <tag> --repo JonathanGodley/fastmail-mcp`);
- every issue closed since the tag. For each one, confirm that a commit after the tag fixed it (`git log <tag>..HEAD --grep "#N"`, `git merge-base --is-ancestor`), and drop accidental closes (see the upstream closing-keyword hazard in `CLAUDE.md`), fixes that shipped in an earlier release, and not-planned closes;
- a sweep of `git log <tag>..HEAD --no-merges` for changes a user would notice that have no closed issue;
- if `/triage-release` ran, the P2 (note-in-release) rows of its `.claude/triage/YYYY-MM-DD-release-triage.md`.

Both are consumer-facing: describe each change, cite its public `#issue`, no internal codenames (`CLAUDE.md` Releasing). The notes lead with breaking changes and behaviour a caller will be surprised by, and end with `**Full Changelog**: https://github.com/JonathanGodley/fastmail-mcp/compare/<last tag>...<new tag>`. The tag message's first line is a title that does not repeat the version, which the tag name already carries.

The subagent reports the issues it cited, each closed issue it left out with the reason, and every claim it could not verify. Check each unverified claim and every breaking-change claim against the code at the tag and at HEAD before showing the draft to the user at the checkpoint.

## 5. ⛔ CHECKPOINT — do not cross until this holds

- Restate the step-1 precondition: the user asked for a release *this session*. Require a fresh explicit yes/no immediately before the first push, with the notes and tag message in front of them.
- Step 3's commands passed on the tree being committed.
- Precheck the outward steps: `gh auth status` succeeds, and the new tag does **not** already exist on origin.

## 6. Commit, then push (two steps — the push is the irreversible one)

Releases land on `main` directly (the fork's documented flow), not a feature branch.

1. **Reset the index first:** `git reset` — start from a known-empty staging area.
2. **Stage explicit per-FILE paths only.** The intended set is the five files step 2 changed: `package.json`, `manifest.json`, `src/index.ts`, `package-lock.json`, `README.md`. NEVER stage a directory, a glob, `-A` or `.`.
3. **Assert the staged set is EXACTLY the intended files.** `git diff --cached --name-only`, sorted, must *equal* the intended list, sorted — an equality check, not a subset. It fails closed: an empty staged set means "nothing changed, investigate," never "force it."
4. Commit: `Release: vX.Y.Z-fork.N`.
5. **Only on a passing assertion**, push: `git push origin main`.

## 7. Tag (annotated) and push the tag

- `git tag -a vX.Y.Z-fork.N -F <absolute path of the step-4 tag message>` (annotated, not lightweight).
- `git push origin vX.Y.Z-fork.N`.

## 8. Publish the GitHub release

```
gh release create vX.Y.Z-fork.N --repo JonathanGodley/fastmail-mcp --verify-tag --notes-file <absolute path of the step-4 notes file>
```
- `--repo` goes adjacent to the command — a mis-defaulted `gh` would publish to upstream.
- `--verify-tag` is why step 7 pushes the tag first.
- The tag push starts "Build and Release DXT", whose release job (`softprops/action-gh-release`) creates the release itself if it gets there first, about a minute after the push. So run step 8 straight after step 7, and if the release already exists use `gh release edit vX.Y.Z-fork.N --repo JonathanGodley/fastmail-mcp --notes-file <path>` instead.

## 9. Close shipped fork issues (close-on-ship, validate first)

Close any issue this release fixed that is still open, on the rule in `CLAUDE.md` Working with upstream (validate first; a tagged release is not a precondition). A bare `(#N)` in a commit does not auto-close; only `Closes/Fixes/Resolves #N` on a default-branch push does.

This applies to the **fork's own** issues only; never close or comment on a `MadLlama25/fastmail-mcp` issue or PR (`CLAUDE.md`, Working with upstream).

```
gh issue close <N> --repo JonathanGodley/fastmail-mcp --comment "<cite the release + commit>"
```

## 10. Post-release

- Confirm the tag's "Build and Release DXT" run succeeded and the release carries `fastmail-mcp-<version>.dxt` (`gh run list --repo JonathanGodley/fastmail-mcp --workflow build-dxt.yml --limit 1`), and that the push's `test.yml` run is green. Reconnect MCP clients to load the new `dist/`.
