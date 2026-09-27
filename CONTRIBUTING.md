# Contributing

## Development rules

### Documentation ships with the change

Any change to a tool's behaviour, parameters or response format, and any added
or removed feature, updates two places in the same change, not as a follow-up:

1. the tool's `description` and `inputSchema` in `src/index.ts`, which is what
   MCP clients see;
2. `README.md`: the tool reference section and any format or feature section the
   change touches.

### Response format

Every tool that returns email data uses the simplified format from
`src/email-formatter.ts`:

- `simplifyEmail()` for full emails and list items;
- empty, null and false fields are omitted to save tokens;
- the output is an allow-list: `simplifyEmail()` emits only the fields it
  names, so a JMAP property that is fetched but not mapped is dropped (a
  fetched `sentAt`, for example, never reaches the simplified output; #84);
- every such tool accepts `raw: true` to bypass simplification and return the
  JMAP objects as fetched.

### JMAP properties

All email list and search methods in `src/jmap-client.ts` request the same set
of `Email/get` properties. `getEmails()` and `searchEmails()` both run through
the shared `runFilteredQuery` helper, which sets `EMAIL_PROPERTIES_COMPACT`
once, so they stay in sync by construction. `getThread()` (full mode) and
`getEmailById()` request extra body properties and must stay a superset of the
list set, so that `raw: true` returns a complete JMAP response. If you add a
property to one, add it to all.

When you append a method call to an existing batch (for example a trailing
`Mailbox/get` to resolve mailbox names), read its result with
`readListResultIfPresent` rather than a hard index. `getMethodResult` and
`getListResult` throw on a missing index, so a hard index would fail every test
that stubs only the original responses, and would fail in production against a
server that drops the trailing method.

A tolerant read must still never silently drop a promised field. A read tool
promises `mailboxes` and `roles`; when a mailbox id cannot be resolved to a
name, the id is reported in `unresolvedMailboxIds` rather than omitted (#53).
In general, when a resolution or enrichment cannot complete, either surface the
degradation explicitly or raise an error on a genuine failure. Do not weaken a
production behaviour to satisfy an under-stubbed test; fix the test.

The positional index reads are safe only because `Email/get`, `Mailbox/get`
and `Thread/get` each appear once per batch. JMAP (RFC 8620 section 3.4) returns
errors as `error` entries in place, so before using index reads on a batch where
a method could appear more than once or be reordered, match responses by call
id instead.

### A destroy must not remove what the server cannot recreate

A tool that irreversibly destroys a record refuses any record whose kind this
server's create tools cannot produce. The recovery echo a destroy returns
(`deletedCard`, for example) is only useful if a create tool can consume it.
That is why `delete_contact` refuses every card whose kind is not `individual`
(a contact group, an org, a location): `create_contact` has no `kind` or
`members` parameter, so it makes individuals only. `update_contact` refuses the
same kinds for the same reason, through the same shared message.

The test is the record kind, not its fields. Most real records carry fields the
create tool cannot set, such as a contact card's titles or photos; those are a
documented limit (see the echo bound in `docs/conventions.md`), not a reason to
refuse. When you add a delete path, check it against the create surface first,
and either refuse the kind that cannot be made or say plainly why the destroy is
still safe. For example, `delete_email` moves to Trash, so no content is lost,
but it writes `mailboxIds` as a whole value, so the message's other labels are
dropped (#123).

### Bumping the version

The version string is hand-edited in three places:

- `package.json`
- `manifest.json`
- the `Server` constructor in `src/index.ts`

Then run `npm install --package-lock-only` to carry it into
`package-lock.json`, which holds it twice. Never hand-edit the lockfile. The
`version sync` test asserts all four agree.

### Building and testing

The server runs from `dist/index.js`, not `src/`, so run `npm run build` after a
change. A running MCP client keeps serving the build it started with; reconnect
it, or exercise the change through the tests or `scripts/mcp-harness.mjs`.

Before committing, run:

```bash
npx tsc --noEmit
npm run typecheck:tests
npm test
```

`npm test` builds first (via `pretest`), because `built-server.test.ts` spawns
`dist/index.js` and checks in a `before` hook that the build exists and is newer
than `src/`. When that hook throws, the runner reports the suite as cancelled
rather than failed, so a run without a build can look green.

Handler logic must be unit-testable. The `CallTool` switch in `src/index.ts`
has no test harness, so a handler that does more than destructure and delegate
belongs in a function that takes an injected client, tested with a mock. The
model is `composeDraftEmail(args, client, attachDir)` in
`src/draft-email-handler.ts`: it takes a `DraftEmailClient` interface, which
`JmapClient` satisfies structurally, so its branches are covered by `npm test`
with no credentials or network.

A live run against a real account is a one-off check of externally observable
behaviour (for example a byte-identical attachment round-trip), never the only
test of logic that could be unit-tested. `scripts/mcp-harness.mjs` is the
reusable client for that: it spawns `dist/index.js` with `FASTMAIL_API_TOKEN` in
its environment and matches JSON-RPC responses by id.

### Mutation testing

Mutation testing changes a source line (a flipped condition, an emptied string)
and checks that some test fails. `npm run mutation -- <commit>` mutates only the
`src/` lines that commit changed; the commit must be `HEAD` and the tree clean.
`npm run mutation -- --all` mutates all of `src/`, reusing earlier results
through Stryker's incremental file in `reports/`. A surviving mutant is a line
no test notices changing: add or tighten a test rather than deleting the line.
Each mutant takes about 0.25 s, so a commit usually finishes in under a minute;
a cold `--all` run (about 15,700 mutants) takes roughly 1 to 1.5 hours on an
8-core desktop. `src/index.ts` is never mutated. Each test process's heap is
capped at 1 GB, so a mutant that grows the heap without bound fails as a
RuntimeError instead of exhausting the machine.

`.github/workflows/mutation.yml` runs the full set weekly, and on manual
dispatch, as four `--all --shard <i>/4` jobs that each carry their incremental
file forward from the newest run that uploaded one, whether or not that run
passed. Read the results in the run's
summary (counts, score and each surviving mutant) or in its report artifacts.
Survivors never fail the workflow; only an error running Stryker does.

### CI runs on Linux

`.github/workflows/test.yml` runs the build, a test-file count check,
`typecheck:tests` and `npm test` on ubuntu-latest across Node 20, 22 and 24, for
every push to `main` and every pull request. `npx-smoke.yml` packs the tarball
and boots the installed binary on the same matrix, and `secret-scan.yml` runs
the secret scanner.

A local pass on Windows or macOS is therefore not the finish line. Any fixture
whose arithmetic depends on path length, the path separator or line endings
must derive its values at run time (for example from the real `tmpdir()`)
rather than assume them; a Windows temp path is far longer than `/tmp`.

### Where design rationale lives

Why one tool behaves as it does is recorded in that tool's GitHub issue. Facts
that span several tools, or that describe the JMAP/Fastmail platform, live in
`docs/`:

- `docs/email-bodies.md`: the body-format model (HTML as the source of truth,
  text/plain as a derived fallback), `edit_draft` coupling, the identity
  signature, MIME-matched body extraction, and destroy-and-recreate.
- `docs/security-model.md`: path confinement for downloads and attachments.
- `docs/fastmail-action-availability.md`: what the Fastmail client offers on
  each screen and what each action actually does, measured rather than
  inferred.
- `docs/conventions.md`: the sending-identity model, lenient input coercion,
  mailbox-query scoping, result serialisation, untrusted values in messages,
  calendar window bounds, free/busy handling, the quote sanitiser, and
  dependency and build notes.
- `docs/upstream-sync.md`: how this fork merges from upstream.

Read the relevant file before re-deriving a decision, and add a new cross-tool
decision there rather than in a local note.

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
A denylist hit is reported
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

A clean run is a useful signal, not a guarantee.

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
