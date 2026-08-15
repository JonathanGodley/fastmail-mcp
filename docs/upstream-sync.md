# Syncing with upstream

This fork tracks `MadLlama25/fastmail-mcp`. The two trees have diverged enough
that a sync is a deliberate operation rather than a `git pull`, and the method
matters more than any single sync: done one way each sync starts from scratch,
done the other way each sync only has to look at what changed since the last one.

This file is the method. Per-feature rationale lives in the relevant fork issue;
cross-cutting models live in the other `docs/*.md` files.

## When to sync

The trigger is a **release**. Releases are this fork's natural cadence, so the
`/release` skill carries a drift check:

```
git fetch upstream
git log --oneline $(git merge-base HEAD upstream/main)..upstream/main
```

If that list is long, a sync is due before the release rather than after it. The
v1.13.4 sync ran at 97 commits of drift, which is well past the point where the
work is comfortable — a two-digit list is a good moment to schedule one.

Separately, `CLAUDE.md` §Working with upstream requires a fork issue for every
open third-party upstream PR. That rule and this drift check are the same
mechanism seen from two ends: the issues catch incoming work one PR at a time,
the drift check catches what merged while nobody was looking.

## Why a real two-parent merge

The fork's differentiators (response simplification, the consolidated search
surface, the draft-first compose surface, path confinement) mean many files
cannot take upstream's version. The temptation is to hand-port the changes worth
having and skip the merge commit entirely.

Do not. Without a merge commit, git has no record that upstream's history was
considered, so the *next* sync re-offers every one of those commits and the whole
adjudication runs again from zero. One real two-parent merge — even one where
most files resolve `--ours` — is what makes the next sync incremental.

For the same reason: **never rebase the sync branch.** Rebasing flattens the
two-parent shape and throws away exactly the thing the merge exists to record.
If `main` moved while the sync was in flight, merge `main` into the sync branch,
re-run the gate, and integrate from there.

## The method

1. **Prep.** `npm ci`, a green baseline (`npx tsc --noEmit && npm test`) with the
   SHA recorded, and a branch off `main`. File issues for anything the sync will
   knowingly not adopt, before the merge rather than after.
2. **One merge commit** with a per-file resolution policy (below), including the
   ports of upstream's own code into files resolved `--ours` — those ports *are*
   the resolution. Fork-originated work that goes beyond upstream rides as
   ordinary follow-up commits, so the audit in step 3 can tell "we kept ours"
   apart from "we wrote new".
3. **The audit ledger, as a hard gate.** See below. No feature work until it is
   clean.
4. **Feature phases** as ordinary commits on `main`, each carrying its own docs.
5. **Close with a tool-surface diff.** The sync is done when every upstream tool
   and every parameter on a shared tool resolves to exactly one of: adopted,
   adopted as a different shape, declined with the reason recorded where the next
   reader hits it, or deferred to a named open issue. Derive both surfaces from
   the code (`git show upstream/main:src/index.ts` against `src/index.ts`), not
   from memory — the count alone will not tell you that a parameter was renamed
   or that a tool was folded into another. Post the diff as the ledger issue's
   closing comment; it is the record that nothing arrived and quietly vanished.

## Per-file resolution policy

| File | Policy | Why |
|---|---|---|
| `src/index.ts` | `--ours` + port | Tool surface is fork-shaped (draft-first, consolidated search); upstream's security edits ported in |
| `src/jmap-client.ts` | `--ours` + port | Simplification, path confinement, the fork's own query helpers |
| `src/coerce.ts` | `--ours` + splice | Fork-dominant; upstream's redaction additions spliced in |
| `src/caldav-client.ts` | `--theirs` + surgery | Upstream is ahead here; the fork's `coerce.js` import and error-class discipline are restored after |
| `src/contacts-calendar.ts` | `--theirs` + surgery | Same, plus the fork's account-routing and `properties` deletions restored |
| `src/url-validation.ts` | `--theirs` + delete | Upstream already carries the fork's regional-host fix; `validateHttpsUrl` dies with the WebDAV client |
| `README.md`, `manifest.json` | `--ours` | Rewritten per-phase, never hunk-merged |
| `package.json`, `package-lock.json` | `--ours` | Not a union — adopt upstream's real changes by hand, in their owning phase |
| `.gitignore` | 3-way, not union | The fork's deletion of the `CLAUDE.md` line must survive; upstream's bare `.claude/` must not land (it would kill the skills un-ignore) |
| `.github/**` | merges clean, then neutralize | Upstream's auto-release trigger is declined; restore this fork's `build-dxt.yml` wholesale |

## The audit ledger (the hard gate)

`--ours` is the dangerous half of the policy. Git records those files as merged,
so anything upstream changed in them is marked "already handled" and **will never
be offered again**. A drop there is invisible and permanent — no compile error,
no failing test, no future conflict.

So every hunk upstream changed in an `--ours` file gets classified into exactly
one of four buckets:

1. **Fork differentiator** — keep ours.
2. **Declined**, with the reason recorded where the next reader will hit it (an
   in-code comment, a `docs/*` file, or a closed issue).
3. **Scheduled** — a checkbox on the ledger issue, ticked by the phase that lands it.
4. **Unclassified** — the gate. Must be empty before any feature work starts.

Run the diff in the direction that catches loss for that file's policy:
`git diff <merge-base>..upstream/main -- <file>` for `--ours` files (what upstream
changed that we could drop), and the mirror `git diff HEAD:<f> upstream/main:<f>`
for `--theirs` files (what the fork loses).

Two things the ledger is not:

- **It is not a substitute for reading the arriving code.** A `--theirs` file
  arrives as third-party code this fork then republishes under the Fastmail name.
  A presence checklist confirms a guard is *there*; it does not confirm the guard
  *works*. Read it adversarially. The v1.13.4 sync found an arriving guard that
  did not bound what it claimed to bound, exactly this way.
- **A ticked checkbox is not the same as an adjudicated item.** Some upstream
  work is neither declined nor scheduled but genuinely undecided, sitting behind
  an open adopt-or-decline issue. Closing the ledger does not decide those.

## The recurring list

These are re-done every sync, because the policy that requires them re-creates
the problem each time. This list is the sync's real cost, so keep it honest.

- **`src/contacts-calendar.ts`: re-delete the `properties:` arrays** on
  `getContacts` and `searchContacts`. Both are two-sided edits present at the
  merge-base, so a plain 3-way merge would carry the fork's deletion
  automatically — `--theirs` throws that away. There is a durable test asserting
  the request objects carry no `properties` key; if it goes red after a sync,
  this is why.
- **`src/contacts-calendar.ts`: re-apply the contact write redesign.** Upstream's
  `updateContact` issues a `properties: ['id']` existence probe and then sends the
  caller's flat arrays as the patch, which destroys the `contexts`/`pref` on every
  entry; its description claims `[] clears`, which its own code does not do. The
  fork reads the whole card, merges per entry, rejects an ambiguous drop-plus-add,
  and rejects `[]` in favour of `clearFields`. Upstream's `deleteContact` destroys
  without reading, so the fork's pre-destroy echo goes too. Both are whole-method
  rewrites, so a `--theirs` resolution silently reverts a data-loss fix — the merge
  helpers live in the fork-only `src/contact-card.ts`, which survives untouched and
  will be left with no callers if this is missed.
- **`src/caldav-client.ts`: restore the `coerce.js` import** and delete
  upstream's local `requireNonEmpty` / `validateClearFields`, including the
  arriving test file's import bindings of those names. The fork's versions throw
  `InvalidInputError` (→ InvalidParams); upstream's throw plain `Error`
  (→ InternalError, "server bug"). There is a test pinning the error class
  specifically because upstream's replacements assert by message regex and pass
  under either class.
- **`src/caldav-client.test.ts`: re-add the fork-only tests** that `--theirs`
  deletes — `toICalUTC`, `foldICalLine`, the iCal round-trip, and the error-class
  pin above.
- **WebDAV paths stay deleted.** `src/webdav-files-client.ts` and its test were
  removed on security grounds (#90). Future syncs will raise modify/delete
  conflicts on both; the resolution is keep-deleted. Two things travel with that
  decision and are easy to miss because neither raises a conflict:
  `validateHttpsUrl` in `src/url-validation.ts` has to be re-deleted (it is dead
  the moment the WebDAV client is gone, and `src/url-validation.ts` is a
  `--theirs` file, so it arrives again every time), and `.mcp.json` has to have
  its three `FASTMAIL_WEBDAV_*` passthroughs re-stripped — upstream's copy hands
  the server a WebDAV password that nothing here reads.
- **`.mcp.json` is rewritten, not merged.** The key list is this server's
  `findEnvValue` call sites, so it carries fork-only keys upstream has never seen
  (`FASTMAIL_ATTACH_DIR`, `FASTMAIL_TIMEZONE`, `FASTMAIL_ALLOW_BLOB_ATTACH`) and
  omits `FASTMAIL_ALLOW_UNSAFE_BASE_URL` deliberately — that one is a kill switch
  and stays env-only, out of both this file and the manifest.
- **`.github/` policy strip.** Upstream's auto-release-on-push is declined; this
  fork's `build-dxt.yml` is restored wholesale rather than patched, because
  deleting only the `decide` job leaves the `release` job referencing a
  nonexistent one. Two smaller strips in the same directory:
  `create-tag.yml`'s tag regex is widened to accept the `-fork.N` suffix this
  fork's versions carry (upstream's accepts plain `vX.Y.Z` only, so a fork tag is
  rejected outright), and its 20-line header comment is rewritten, because
  upstream's documents the auto-release flow as the normal path. And
  `dependabot.yml` needs its `debug` ignore entry re-added: upstream carries only
  the `@types/node` one. That entry is a security control, not a stability
  preference — see the comment in the file — and losing it lets a routine bump
  nest a second `debug` copy under `tsdav` and silently stop the credential-log
  suppression while every test stays green.
- **`package.json`, `.gitignore` and `.dxtignore` are never a union.** All three
  look mergeable and all three carry deliberate fork deletions or fork-only
  stanzas — `.gitignore`'s `!.claude/skills/**/SKILL.md` un-ignore, `.dxtignore`'s
  exclusion of `/docs/`, `/AGENTS.md`, `/.claude/` and the build droppings,
  `package.json`'s extra dependencies and its exact `debug` pin.
- **Test-fixture email domains diverge in `src/coerce.test.ts`** (~13 lines) and in
  `scripts/scan-secrets.mjs`'s `SAFE_EMAIL_DOMAINS`. Upstream's fixtures use real
  registered placeholder domains (`b.com`, `d.com`, `x.com`, …) and safelist them by
  name; the fork moved those fixtures to RFC 2606 reserved forms (`@b.example`) and
  dropped the corresponding safelist entries. Expect a mechanical conflict on those
  lines whenever upstream touches them — the resolution is the fork's form.

  The rule going forward, so this does not grow: **match upstream's fixtures in
  upstream-shared files; use reserved-TLD forms in fork-authored ones.** The
  reserved-TLD suffix rule in the scanner is additive and touches nothing upstream
  wrote, so it survives any resolution. Reverting the existing divergence was
  considered and declined — 13 mechanical lines is cheaper than the churn of undoing
  it, and the fixtures are equally correct either way.

## What the sync must not silently change

Grep these after every merge. Each is a trap a plausible-looking resolution
springs quietly:

- `send_email` absent from `TOOLS`, and no compose schema carrying a `send`
  param. Upstream still ships both, so the risk runs toward *reintroducing* a
  transmit path — the opposite direction from a normal drop.
- Every credential-bearing `fetch` carries `redirect: 'error'`.
- Both attachment writes use `flag: 'wx'` (the exclusive-create TOCTOU control).
- Contacts getters carry no `properties` filter, and all three route through
  `contactsAccountId()`.
- `updateContact` reads the whole card before patching and merges per entry;
  `deleteContact` reads the card in the same request as the destroy and returns it.
  Both echoes reach tool output regardless of `verbose`/`raw`.
- No error text reaches tool output unredacted — returned fields count, not just
  throws.
- No new caller-fixable rejection throws a plain `Error` at the boundary.
- `assertICalTextLimits` is still called in BOTH the `create_calendar_event` and
  `update_calendar_event` handlers. `src/ical-limits.ts` is fork-only so it
  survives any resolution untouched, but its two call sites live in `index.ts`
  and a `--theirs` resolution there drops them silently: the tools keep working,
  and the quadratic `foldICalLine` goes back to being reachable without a bound
  (see `docs/conventions.md`, "Bounding a quadratic serializer").

Then the full gate — `npx tsc --noEmit && npm test && npm run build` — and, because
the live harness spawns `dist/index.js`, at least one live check of a path that
cannot be mocked. After the v1.13.4 merge those were a CalDAV event write and a
contacts read: roughly 1,100 lines of un-mockable write path, and a routing change
that would make every contact read return empty against a real account while every
mock stayed green.
