# Scripts

Developer tooling. None of it ships with the server; `dist/` is built from `src/` alone.

| Script | What it does |
| --- | --- |
| `mcp-harness.mjs` | Raw JSON-RPC client that spawns `dist/index.js` as a real MCP server. `createClient({ env })` → `init`/`call`/`close`, matching responses by JSON-RPC `id`. `node scripts/mcp-harness.mjs --list` dumps the advertised tool schemas. Use it rather than hand-writing a client. |
| `mutate-commit.mjs` | `node scripts/mutate-commit.mjs <commit>` (or `npm run mutate -- <commit>`) runs Stryker over only the source lines that commit changed, against only the test files matching those sources. Prints the mutants that survived — each one a changed line no test notices. Everything it writes goes outside the repo (`MUTATE_OUT_DIR`, default `os.tmpdir()`). |
| `scan-secrets.mjs` | The PII/credential scanner the git hooks run (`npm run scan:secrets` for a full sweep). See `CONTRIBUTING.md`. |
| `dump-official-surface.mjs` | Snapshots Fastmail's official MCP tool surface into `docs/official-mcp-*` for the comparison doc. The output is deliberately not committed; the generator is. |
| `probes/` | On-demand live checks against a real account. See `probes/README.md` — they are not regression coverage, and some of them send mail. |
