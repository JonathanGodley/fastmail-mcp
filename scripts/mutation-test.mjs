#!/usr/bin/env node
// Mutation testing with Stryker. A surviving mutant is a source line no test notices changing.
//
//   node scripts/mutation-test.mjs <commit>   mutate only the src/ lines that commit changed
//   node scripts/mutation-test.mjs --all      mutate all of src/, incrementally (reports/)
//   node scripts/mutation-test.mjs --all --shard <i>/<n>
//                                             mutate shard i of n, with an html report as well
//
// Every mode writes a json report listing each mutant's status, the only place an errored
// mutant is named: reports/mutation-commit.json, reports/mutation.json or
// reports/mutation-<i>-of-<n>.json.

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const USAGE = 'Usage: node scripts/mutation-test.mjs <commit> | --all [--shard <i>/<n>]';

// These read src/ as text or run the built dist/, so under Stryker they either fail its initial
// run (its instrumentation reads process.env and reprints the mutated file whole) or measure
// the last build instead of the mutant.
export const EXCLUDED_TESTS = ['index-env', 'readme-inventory', 'tool-schema', 'config-surface', 'built-server', 'server-lifecycle']
  .map((name) => `src/${name}.test.ts`);

// A mutant that grows the JavaScript heap without bound would otherwise take the whole machine
// before Stryker's timeout, and a CI runner with it. Capped, it crashes its own test process
// and is reported as a RuntimeError. A test process's old space peaks under 50 MB, and four of
// these fit a 16 GB runner. The cap does not reach ArrayBuffer memory.
export const TEST_HEAP_MB = 1024;

/** Returns { all: true, shard? }, { commit }, or { error } for anything else. */
export function parseArgs(argv) {
  if (argv.length === 1 && argv[0] === '--all') return { all: true };
  if (argv.length === 3 && argv[0] === '--all' && argv[1] === '--shard') {
    const m = /^(\d+)\/(\d+)$/.exec(argv[2]);
    const i = Number(m?.[1]);
    const n = Number(m?.[2]);
    if (!m || i < 1 || i > n) return { error: `--shard takes <i>/<n> with 1 <= i <= n, got ${argv[2]}` };
    return { all: true, shard: { i, n } };
  }
  if (argv.length === 1 && !argv[0].startsWith('-')) return { commit: argv[0] };
  return { error: argv.length === 0 ? 'A commit or --all is required.' : `Unexpected arguments: ${argv.join(' ')}` };
}

// Instrumenting src/index.ts at all breaks jmap-client.test.ts's version-sync check, and every
// other test pinning its content is one of the EXCLUDED_TESTS, so it is never mutated.
export const isMutable = (file) => file.startsWith('src/') && file.endsWith('.ts') && file !== 'src/index.ts'
  && !file.endsWith('.test.ts') && !file.startsWith('src/testing/');

export const MUTATE_ALL = ['src/**/*.ts', '!src/index.ts', '!src/**/*.test.ts', '!src/testing/**'];

export const testFiles = (srcNames) => srcNames.filter((n) => n.endsWith('.test.ts'))
  .map((n) => `src/${n}`).filter((f) => !EXCLUDED_TESTS.includes(f)).sort();

/** `git diff -U0` output -> Stryker mutate ranges (`src/x.ts:5-7`) for the mutable files. */
export function diffToRanges(diff) {
  const ranges = [];
  let file = null;
  for (const line of diff.split('\n')) {
    // A real header only: an added content line can itself start with "++ ".
    const header = /^\+\+\+ (?:b\/(.+)|\/dev\/null)$/.exec(line);
    if (header) { file = header[1] ?? null; continue; }
    const hunk = /^@@ -\S+ \+(\d+)(?:,(\d+))? @@/.exec(line);
    if (!hunk || !file || !isMutable(file)) continue;
    const start = Number(hunk[1]);
    const count = hunk[2] === undefined ? 1 : Number(hunk[2]);
    if (count > 0) ranges.push(`${file}:${start}-${start + count - 1}`);
  }
  return ranges;
}

/**
 * Deals [path, bytes] pairs into n shards of similar total size: largest first, each to the
 * lightest shard so far (lowest index on a tie). Returns n sorted path lists.
 */
export function partition(files, n) {
  if (!Number.isInteger(n) || n < 1) throw new Error(`shard count must be a positive integer, got ${n}`);
  const shards = Array.from({ length: n }, () => ({ bytes: 0, paths: [] }));
  const bySize = [...files].sort(([a, x], [b, y]) => y - x || (a < b ? -1 : a > b ? 1 : 0));
  for (const [file, bytes] of bySize) {
    const lightest = shards.reduce((min, s) => (s.bytes < min.bytes ? s : min));
    lightest.bytes += bytes;
    lightest.paths.push(file);
  }
  return shards.map((s) => s.paths.sort());
}

/** The Stryker reporters for a parseArgs result, and the tag its report file names carry. */
export function reportOutputs(args) {
  const { shard } = args;
  return {
    reporters: ['clear-text', 'progress', 'json', ...(shard ? ['html'] : [])],
    tag: shard ? `-${shard.i}-of-${shard.n}` : args.commit ? '-commit' : '',
  };
}

function main() {
  const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const git = (...args) => execFileSync('git', args, { cwd: REPO, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  const resolve = (rev) => {
    try { return git('rev-parse', '--verify', '--quiet', '--end-of-options', `${rev}^{commit}`).trim(); } catch { return null; }
  };
  const fail = (msg) => { console.error(msg); process.exit(2); };

  const args = parseArgs(process.argv.slice(2));
  if (args.error) fail(`${USAGE}\n${args.error}`);

  let mutate = MUTATE_ALL;
  if (args.commit) {
    const sha = resolve(args.commit);
    if (!sha) fail(`${args.commit} does not resolve to a commit.`);
    // Stryker copies the tree on disk, not the commit, so anything else there would be measured
    // as if the commit contained it.
    if (sha !== resolve('HEAD')) fail(`${args.commit} is not HEAD. Check it out (detached, or in a worktree) and re-run.`);
    const status = git('status', '--porcelain').trim();
    if (status) fail(`The tree is not clean, so it would not describe ${args.commit}:\n${status}`);
    if (!resolve(`${sha}~1`)) fail(`${args.commit} is a root commit; there is no parent to diff against.`);
    mutate = diffToRanges(git('diff', `${sha}~1`, sha, '-U0', '--', 'src'));
    if (mutate.length === 0) fail(`${args.commit} changed no mutable src/ lines.`);
  }
  const { shard } = args;
  const { reporters, tag } = reportOutputs(args);
  if (shard) {
    // Blob sizes at HEAD, not on-disk sizes, so every platform and line-ending setting deals
    // the same shards and each shard's incremental file keeps matching its files.
    const files = git('ls-tree', '-r', '-l', 'HEAD', '--', 'src').trim().split('\n')
      .map((line) => /^\S+ blob \S+\s+(\d+)\t(.+)$/.exec(line))
      .filter((m) => m && isMutable(m[2]))
      .map((m) => [m[2], Number(m[1])]);
    mutate = partition(files, shard.n)[shard.i - 1];
    if (mutate.length === 0) fail(`Shard ${shard.i}/${shard.n} has no files; use fewer shards.`);
  }

  // A git worktree has no node_modules; `npm install` runs once, in the primary checkout. Rather
  // than link anything, the sandboxes live in the primary's .stryker-tmp, so Node's ordinary
  // walk-up from any sandboxed file reaches its node_modules. From the primary this is Stryker's
  // default location.
  const primary = path.dirname(git('rev-parse', '--path-format=absolute', '--git-common-dir').trim());
  const depsRoot = [REPO, primary].find((dir) => existsSync(path.join(dir, 'node_modules')));
  if (!depsRoot) fail(`No node_modules in ${REPO} or ${primary}. Run \`npm install\` in the primary checkout.`);

  const config = {
    testRunner: 'tap',
    coverageAnalysis: 'perTest',
    tap: {
      testFiles: testFiles(readdirSync(path.join(REPO, 'src'))),
      nodeArgs: [`--max-old-space-size=${TEST_HEAP_MB}`, '--test-reporter=tap', '--import', 'tsx', '-r', '{{hookFile}}', '{{testFile}}'],
    },
    mutate,
    ignorePatterns: ['/dist', '/.claude', '/coverage'],
    tempDirName: path.join(depsRoot, '.stryker-tmp'),
    symlinkNodeModules: false,
    cleanTempDir: 'always',
    incremental: Boolean(args.all),
    incrementalFile: `reports/stryker-incremental${tag}.json`,
    reporters,
    jsonReporter: { fileName: `reports/mutation${tag}.json` },
    htmlReporter: { fileName: `reports/mutation${tag}.html` },
    // Stryker's tsconfig rewriter calls an API TypeScript 7 removed; a missing file skips it.
    tsconfigFile: 'tsconfig.stryker-skip.json',
  };
  const configFile = path.join(tmpdir(), `fastmail-mcp-stryker-${process.pid}.json`);
  writeFileSync(configFile, JSON.stringify(config, null, 2));

  const corePkg = createRequire(path.join(depsRoot, 'package.json')).resolve('@stryker-mutator/core/package.json');
  const bin = path.join(path.dirname(corePkg), JSON.parse(readFileSync(corePkg, 'utf8')).bin.stryker);
  console.log(`mutating ${REPO}${depsRoot === REPO ? '' : ` with dependencies from ${depsRoot}`}`);
  const run = spawnSync(process.execPath, [bin, 'run', configFile], { cwd: REPO, stdio: 'inherit' });
  rmSync(configFile, { force: true });
  if (run.error) fail(`Could not start Stryker: ${run.error.message}`);
  process.exit(run.status ?? 1);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) main();
