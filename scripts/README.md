# Scripts

Developer tooling. None of it ships with the server; `dist/` is built from `src/` alone.

| Script | What it does |
| --- | --- |
| `mcp-harness.mjs` | Raw JSON-RPC client that spawns `dist/index.js` as a real MCP server. `createClient({ env })` → `init`/`call`/`close`, matching responses by JSON-RPC `id`. `node scripts/mcp-harness.mjs --list` dumps the advertised tool schemas. Use it rather than hand-writing a client. |
| `mutate-commit.mjs` | `node scripts/mutate-commit.mjs <commit>` (or `npm run mutate -- <commit>`) runs Stryker over only the source lines that commit changed, against only the test files matching those sources. Prints the mutants that survived — each one a changed line no test notices. Everything it writes goes outside the repo (`MUTATE_OUT_DIR`, default `os.tmpdir()`). **Exits 1 when any mutant survives**, so a run that found something ends in npm's error banner under `npm run mutate`; that is the finding, not a crash. |
| `scan-secrets.mjs` | The PII/credential scanner the git hooks run (`npm run scan:secrets` for a full sweep). See `CONTRIBUTING.md`. |
| `comment-share.mjs` | `node scripts/comment-share.mjs <commit>` (or `--staged`, `--working`, `<a>..<b>`) says when a change added far more comment than the file it landed in carries - at least 20 comment lines at twice the file's own comment-per-code ratio. The pre-commit hook prints it; it never gates. Quiet unless something is over the bar; `--all` prints every measured file. |
| `find-raw-interpolations.mjs` | `node scripts/find-raw-interpolations.mjs [dir]` lists every `'${...}'` under `src/` rendered with no echo helper — the untrusted-values-in-prose class that the `echo-quoting convention` guard in `src/coerce.test.ts` provably **cannot** cover, because whether a value is untrusted depends on where it came from and nothing lexical separates the two. A hand-run inventory, not a test: the output is candidates to trace, never a defect list, and it always exits 0. |
| `dump-official-surface.mjs` | Snapshots Fastmail's official MCP tool surface into `docs/official-mcp-*` for the comparison doc. The output is deliberately not committed; the generator is. |
| `probes/` | On-demand live checks against a real account. See `probes/README.md` — they are not regression coverage, and some of them send mail. |
