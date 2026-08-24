#!/usr/bin/env node
// Mutation-test only the source lines a single commit changed.
//
// Asks, for every line a commit touched, "would any test notice if this line did
// something else?". A survived mutant is a line no test covers, which is the same
// question as hand-reverting a change to watch its new test fail - asked of every
// changed line at once instead of one test at a time.
//
// Usage: node scripts/mutate-commit.mjs <commit> [--tests src/a.test.ts,src/b.test.ts]
//
// Test files are chosen from the commit by default. `--tests` overrides that, which is
// needed when the mechanical choice picks up a test Stryker cannot run: a drift guard that
// reads src/*.ts as TEXT (index-env.test.ts, readme-inventory.test.ts) sees Stryker's
// instrumentation in the file and fails the initial run before any mutant is tried.
//
// Writes the Stryker config, sandbox and report OUTSIDE the repo tree (os.tmpdir by
// default; override with MUTATE_OUT_DIR). Re-running overwrites; nothing needs deleting.

import { execFileSync, spawnSync } from 'node:child_process';
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = path.resolve(process.env.MUTATE_OUT_DIR || path.join(tmpdir(), 'fastmail-mcp-mutate'));
const CONFIG_FILE = path.join(OUT_DIR, 'stryker.config.json');
const REPORT_FILE = path.join(OUT_DIR, 'mutation-report.json');

const argv = process.argv.slice(2);
const testsFlag = argv.indexOf('--tests');
const testsOverride = testsFlag === -1 ? null : (argv[testsFlag + 1] || '').split(',').filter(Boolean);
const commit = argv.filter((_, i) => testsFlag === -1 || (i !== testsFlag && i !== testsFlag + 1))[0];
if (!commit || (testsFlag !== -1 && !testsOverride?.length)) {
  console.error('Usage: node scripts/mutate-commit.mjs <commit> [--tests src/a.test.ts,src/b.test.ts]');
  process.exit(2);
}

const git = (...args) => execFileSync('git', args, { cwd: REPO, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

/** Changed line ranges per file, from a zero-context diff of the commit against its parent. */
function changedRanges(rev) {
  const diff = git('diff', `${rev}~1`, rev, '-U0', '--', 'src');
  const ranges = new Map();
  let file = null;
  for (const line of diff.split('\n')) {
    const header = /^\+\+\+ b\/(.+)$/.exec(line);
    if (header) {
      file = header[1].trim();
      continue;
    }
    const hunk = /^@@ -\S+ \+(\d+)(?:,(\d+))? @@/.exec(line);
    if (!hunk || !file) continue;
    const start = Number(hunk[1]);
    const count = hunk[2] === undefined ? 1 : Number(hunk[2]);
    if (count === 0) continue; // pure deletion: nothing left in the new file to mutate
    if (!ranges.has(file)) ranges.set(file, []);
    ranges.get(file).push([start, start + count - 1]);
  }
  return ranges;
}

const isTest = (f) => f.endsWith('.test.ts');
const isSource = (f) => f.startsWith('src/') && f.endsWith('.ts') && !isTest(f);

const ranges = changedRanges(commit);
const sources = [...ranges.keys()].filter(isSource).sort();
if (sources.length === 0) {
  console.error(`${commit} changed no non-test src/*.ts file - nothing to mutate.`);
  process.exit(1);
}

// Test selection: for each changed src/X.ts, every src/X*.test.ts that exists, plus any
// test file the commit itself changed.
const srcFiles = readdirSync(path.join(REPO, 'src'));
const tests = new Set([...ranges.keys()].filter(isTest));
for (const source of sources) {
  const stem = path.basename(source, '.ts');
  for (const candidate of srcFiles) {
    if (candidate.startsWith(stem) && candidate.endsWith('.test.ts')) tests.add(`src/${candidate}`);
  }
}
const testFiles = testsOverride ?? [...tests].sort();
if (testFiles.length === 0) {
  console.error(`No test file matches ${sources.join(', ')} - every mutant would survive by default.`);
  process.exit(1);
}

const mutate = sources.flatMap((f) => ranges.get(f).map(([a, b]) => `${f}:${a}-${b}`));

mkdirSync(OUT_DIR, { recursive: true });
writeFileSync(CONFIG_FILE, JSON.stringify({
  $schema: './node_modules/@stryker-mutator/core/schema/stryker-schema.json',
  packageManager: 'npm',
  testRunner: 'command',
  commandRunner: { command: `npx tsx --test ${testFiles.join(' ')}` },
  coverageAnalysis: 'off', // the command runner reports one exit code, not per-test coverage
  mutate,
  // Everything Stryker writes goes to OUT_DIR: the sandbox copies (tempDirName) and the
  // machine-readable report. 'html' is left out because it writes into reports/ in the repo.
  tempDirName: path.join(OUT_DIR, 'sandbox'),
  cleanTempDir: 'always',
  reporters: ['progress', 'json'],
  jsonReporter: { fileName: REPORT_FILE },
  timeoutMS: 60000,
  // Point Stryker's tsconfig rewriter at a file that does not exist, so it skips. It rewrites
  // `extends`/`references` paths that would break in the sandbox; this repo's tsconfig has
  // neither, and the rewriter calls a TypeScript API that typescript@7 no longer exposes
  // (`ts.parseConfigFileTextToJson`), which aborts the whole run.
  tsconfigFile: 'tsconfig.stryker-skip.json',
}, null, 2) + '\n');

console.log(`commit      ${commit}`);
console.log(`mutating    ${mutate.join('\n            ')}`);
console.log(`tests       ${testFiles.join(' ')}`);
console.log(`scratch     ${OUT_DIR}\n`);

// Blank the previous report BEFORE the run, so a Stryker that dies without writing one is
// reported as a failure instead of silently re-reporting the last run's numbers.
writeFileSync(REPORT_FILE, '');

const run = spawnSync('npx', ['stryker', 'run', CONFIG_FILE], {
  cwd: REPO,
  stdio: ['ignore', 'inherit', 'inherit'],
  shell: process.platform === 'win32',
});

let report;
try {
  report = JSON.parse(readFileSync(REPORT_FILE, 'utf8'));
} catch {
  console.error(`\nStryker produced no report at ${REPORT_FILE} (exit ${run.status}).`);
  console.error('If it failed the initial test run, one of the selected tests reads src/*.ts as');
  console.error('text and is seeing Stryker\'s instrumentation. Re-run with --tests to drop it.');
  process.exit(run.status || 1);
}

const trim = (s, n) => {
  const one = String(s ?? '').replace(/\s+/g, ' ').trim();
  return one.length > n ? `${one.slice(0, n)}...` : one;
};

const counts = {};
const survivors = [];
for (const [file, entry] of Object.entries(report.files || {})) {
  const lines = (entry.source || '').split('\n');
  for (const m of entry.mutants || []) {
    counts[m.status] = (counts[m.status] || 0) + 1;
    if (m.status !== 'Survived' && m.status !== 'NoCoverage') continue;
    const { start, end } = m.location;
    // Two literals on one line mutate to the same text, so print what was replaced as well.
    const original = (lines[start.line - 1] || '').slice(start.column - 1, end.line === start.line ? end.column - 1 : undefined);
    survivors.push({ file, line: start.line, column: start.column, mutator: m.mutatorName, status: m.status, original, replacement: m.replacement });
  }
}

const total = Object.values(counts).reduce((a, b) => a + b, 0);
const killed = (counts.Killed || 0) + (counts.Timeout || 0);
const scored = killed + (counts.Survived || 0) + (counts.NoCoverage || 0);
const score = scored === 0 ? 'n/a' : `${((killed / scored) * 100).toFixed(2)}%`;

console.log(`\n=== ${total} mutants: ${Object.entries(counts).map(([k, v]) => `${k} ${v}`).join(', ')} ===`);
console.log(`Mutation score ${score}\n`);

if (survivors.length === 0) {
  console.log('No survivors: every mutable line this commit changed is noticed by a test.');
} else {
  console.log(`${survivors.length} survived - each is a line no selected test notices:\n`);
  for (const s of survivors.sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line || a.column - b.column)) {
    console.log(`  ${s.file}:${s.line}:${s.column}  ${s.mutator} [${s.status}]`);
    console.log(`      ${trim(s.original, 120)}`);
    console.log(`      -> ${trim(s.replacement, 120)}`);
  }
}

process.exit(survivors.length > 0 ? 1 : run.status || 0);
