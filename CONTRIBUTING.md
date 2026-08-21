# Contributing

## Secret & PII protection

This repo has layered guards to keep credentials and personal information out of
commits and published artifacts. Please keep them working.

### Enable the git hooks (one-time, per clone)

```bash
git config core.hooksPath .githooks
```

That switches on three hooks, all of which run `scripts/scan-secrets.mjs`:

- **pre-commit** scans your staged content. It reads the staged blobs, not your
  working tree, so a secret you `git add` and then edit out of the working copy
  is still caught.
- **commit-msg** scans the commit message, which is the text `git log`, GitHub
  and every release note will show.
- **pre-push** scans everything the push would publish: the message and content
  of every commit not yet reachable from any remote-tracking ref, and the
  annotation of every tag. It is the backstop for the two things the commit
  hooks cannot see: a commit made with `--no-verify`, and a tag message (git
  has no hook that runs when a tag is created).

The hooks find the scanner relative to their own location, so `core.hooksPath`
can also be an absolute path. That is useful with `git worktree`: a relative
`.githooks` resolves inside each worktree and so runs whatever version of the
hooks that worktree's branch has, while an absolute path to the main checkout's
`.githooks` gives every worktree the same, current hooks.

### What the scanner checks

- **Credentials**: Fastmail API tokens (`fmu…`), `Bearer`/`Basic` auth values,
  and hardcoded `token`/`secret`/`password`/`api_key` assignments.
- **Personal information**: email addresses on any domain outside a small
  allowlist of placeholder/service domains (`example.com`, `fastmail.com`, …),
  and Australian mobile numbers (the shape a signature block carries).
- **Your own identifiers**, if you give it a local denylist (below).

A finding names the file (or message) and line and the rule that matched,
never the matched value: printing the value would put another copy of the
thing being protected into terminal scrollback and anything that captures it.
Open the location named and you will see it. A denylist hit is reported
differently from a shape hit - it means a string you have declared must never
leave the machine reached a commit or message, so the output says to stop and
escalate rather than to rewrite and retry.

Run it manually anytime:

```bash
npm run scan:secrets
```

The same scan runs in CI (`.github/workflows/secret-scan.yml`) on every push and
pull request, so it catches anything the local hooks missed, and again as a gate
before the `.dxt` is packed.

Every run prints its own coverage: how many files it scanned, how many it
excluded by policy, and how many it could not read. A file it could not read
fails the run and is named in the output, because a gate that cannot see a file
should say so rather than report clean. If a genuinely binary file has to be
tracked, add its exact path to `BINARY_ALLOWANCES` in the script; that keeps
each blind spot written down instead of inferred from a filename.

### What the scanner does not cover

A clean run is a useful signal, not a guarantee. These are its known limits.

**It never reads git history.** It only ever looks at the current staged content
or the current checkout. "Scan clean" says nothing about what is in earlier
commits, and a secret that was committed and later removed stays in history
until the history itself is rewritten. Removing a leaked credential from the
tree is not a substitute for revoking it.

**It does not read `dist/`, lockfiles, a packed `.dxt`, or untracked files.**
Build output and the lockfile are excluded by policy, and the file list comes
from git, so anything untracked is invisible to it. This matters most at release
time: the release workflow builds from a clean checkout, so untracked files
never exist there to be scanned in the first place. Anything you want checked
has to be tracked.

**The "example" skip is per matched substring, within one rule.** The hardcoded
assignment rule ignores a match containing an obvious placeholder (`process.env`,
`your-`, `example`, a run of `x` or `0`, and similar). That test runs against the
matched text alone and only for that rule, so a placeholder in one part of a
line does nothing for a real credential elsewhere on it, and nothing at all for
the token, `Bearer`, `Basic` or email rules.

**Unquoted assignments never match.** The assignment rule requires the value to
be in quotes, so `API_KEY=<value>` in shell, `.env` or YAML style is not
detected. Treat env-style files as unscanned.

**The `allowlist-secret` marker is line-total.** One marker silences every rule
on that line, including the personal-email and local-denylist checks, not only
the match you added it for. It also has to sit in a comment: the script looks
for `//`, `#`, `/*` or `<!--` earlier on the line, and ignores a marker that
falls inside the matched credential itself, so a value that happens to contain
the marker string cannot suppress itself. That check is textual rather than a
real parse of each language, so a `//` inside a string literal earlier on the
line still makes the rest of that line look like a comment. Put the marker at
the end of the line in a real comment and it behaves as documented.

**Two exempt domains are registered names.** `SAFE_EMAIL_DOMAINS` includes
`evil.com` and `other.com`, kept because attacker-domain fixtures only mean
something if the domain reads as hostile and real. Both are registered domains
that someone could hold an address at, so each one is a standing blind spot: a
genuine address at either would pass the scan. This is a deliberate trade, and
it is the reason the reserved suffixes in `RESERVED_TLDS` (`.example`,
`.invalid`, `.test`, `.localhost`) are preferred for new fixtures. Those
suffixes can never be delegated to anyone, so exempting them closes the whole
space rather than one name at a time, and costs no coverage. Do not add another
registered domain to the exempt list without weighing what it blinds.

**The local denylist is per-clone.** `.secret-scan-local.txt` is gitignored and
`git config secretscan.denylist` is local config, so your own strings are only
checked on the machine that has them. CI and other contributors' hooks run
without them.

**The push scan decides what is "new" by your remote-tracking refs.** A commit
already reachable from any fetched remote branch is treated as published and is
not rescanned, so a stale fetch can make the push scan skip commits that are in
fact new; `git fetch` before a push keeps it accurate. A merge commit is scanned
for its own resolutions only (the paths whose merged result differs from every
parent), since each merged commit is scanned as itself. And `git push
--no-verify` bypasses the push scan entirely, just as `git commit --no-verify`
bypasses the commit scans.

**Messages have no suppression marker.** The `allowlist-secret` marker works in
file content only. A commit or tag message that trips a rule is rewritten.

**It matches shapes, not meaning.** Every rule is a regular expression over a
single line. A credential in a format it has no rule for, or one split across
lines, passes.

### Triaging a hit

Every hit gets triaged, and there are only two outcomes. A true positive is
scrubbed from the tree, and if it was ever committed, the credential is revoked
or rotated as well. A hit is suppressed only once it is proven to be a synthetic
fixture, and the suppression carries a comment saying why it is safe, so the next
reader can check the reasoning instead of trusting the marker.

Never suppress a hit to get a commit through.

### Test fixtures must be synthetic

Never paste a real token, password, or personal email into a test, not even a
revoked one. Use obviously-fake values: addresses under `example.com` or one of
the reserved suffixes above, and zero-filled or `a`-filled token shapes. If a
synthetic value is unavoidably credential-shaped and the scanner flags it, put
`allowlist-secret` in a comment on that line, along with the reason it is safe:

```ts
const sample = 'fmu0-00000000-0000…'; // allowlist-secret (synthetic token shape)
```

### Local denylist for your own identifiers (optional but recommended)

To make the scanner also flag *your* real names, domains, addresses and
identifiers without publishing them, give it a local denylist: one literal
string per line, `#` for comments, matched as a case-insensitive substring.
There are two places it can live, and both are read if both exist:

```bash
# 1. A gitignored file in the repo root:
cp .secret-scan-local.txt.example .secret-scan-local.txt
# then add your strings, one per line

# 2. Any file outside the repo, named in local git config - useful when one
#    list is shared with other tooling on the machine, and because local git
#    config is shared by every worktree of the clone:
git config secretscan.denylist /path/to/your/denylist.txt
```

A configured list that cannot be read fails the scan rather than being skipped,
so a moved file shows up as a loud failure, not a quiet loss of coverage.
Because the match is a substring, keep entries specific: a short surname that
is also an ordinary word will flag unrelated text.

### Packaging

The published `.dxt` is built from `dist/` plus runtime dependencies only.
`.dxtignore` excludes `src/`, all `*.test.*` files, scripts, and CI config, so
source and test files never ship inside a release binary.

## Reporting problems

Issues and pull requests for this fork go to
[JonathanGodley/fastmail-mcp](https://github.com/JonathanGodley/fastmail-mcp).
If you have found a credential or personal datum that reached a published
artifact, please report it there rather than opening a public pull request that
points at it.
