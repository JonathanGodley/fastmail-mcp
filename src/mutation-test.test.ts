/**
 * Tests for scripts/mutation-test.mjs: the pure helpers directly, and the refusals that run
 * before Stryker by spawning a committed copy of the script in a throwaway git repo. Stryker
 * itself is never started.
 */

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import {
  copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, unlinkSync, writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  EXCLUDED_TESTS, MUTATE_ALL, TEST_HEAP_MB, diffToRanges, isMutable, parseArgs, partition, reportOutputs, testFiles,
} from '../scripts/mutation-test.mjs';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');

test('parseArgs accepts one commit or --all, and nothing else', () => {
  assert.deepEqual(parseArgs(['abc123']), { commit: 'abc123' });
  assert.deepEqual(parseArgs(['--all']), { all: true });
  for (const argv of [[], ['a', 'b'], ['--all', 'a'], ['--tests=x'], ['--help']]) {
    assert.ok('error' in parseArgs(argv), argv.join(' '));
  }
});

test('parseArgs accepts --all --shard i/n and rejects a bad i/n', () => {
  assert.deepEqual(parseArgs(['--all', '--shard', '2/4']), { all: true, shard: { i: 2, n: 4 } });
  assert.deepEqual(parseArgs(['--all', '--shard', '1/1']), { all: true, shard: { i: 1, n: 1 } });
  for (const s of ['0/4', '5/4', '1/0', '0/0', '-1/4', '2', '2/', '/4', 'a/b', '1.5/4', '2/4/6', '']) {
    assert.ok('error' in parseArgs(['--all', '--shard', s]), s);
  }
  for (const argv of [['--shard', '1/4'], ['abc', '--shard', '1/4'], ['--all', '--shard'], ['--shard', '1/4', '--all']]) {
    assert.ok('error' in parseArgs(argv), argv.join(' '));
  }
});

test('every mode writes a json report under its own tag, and only a shard writes html', () => {
  const commit = reportOutputs({ commit: 'abc123' });
  assert.ok(commit.reporters.includes('json'), commit.reporters.join());
  assert.ok(!commit.reporters.includes('html'), commit.reporters.join());
  assert.equal(commit.tag, '-commit');
  const all = reportOutputs({ all: true });
  assert.ok(all.reporters.includes('json'), all.reporters.join());
  assert.ok(!all.reporters.includes('html'), all.reporters.join());
  assert.equal(all.tag, '');
  const shard = reportOutputs({ all: true, shard: { i: 2, n: 4 } });
  assert.ok(shard.reporters.includes('json') && shard.reporters.includes('html'), shard.reporters.join());
  assert.equal(shard.tag, '-2-of-4');
});

test('partition covers every file exactly once, balanced by size, deterministically', () => {
  const files: [string, number][] = [['src/a.ts', 900], ['src/b.ts', 500], ['src/c.ts', 400], ['src/d.ts', 300], ['src/e.ts', 100], ['src/f.ts', 100]];
  const shards = partition(files, 3);
  assert.deepEqual(shards, [['src/a.ts'], ['src/b.ts', 'src/e.ts', 'src/f.ts'], ['src/c.ts', 'src/d.ts']]);
  assert.deepEqual(shards.flat().sort(), files.map(([f]) => f).sort());
  assert.deepEqual(partition([...files].reverse(), 3), shards, 'input order must not matter');
  assert.deepEqual(partition(files, 1), [files.map(([f]) => f)]);
  assert.deepEqual(partition(files.slice(0, 2), 3), [['src/a.ts'], ['src/b.ts'], []]);
  for (const n of [0, -1, 1.5, NaN]) assert.throws(() => partition(files, n), /positive integer/);
});

test('partition over the real src tree keeps every mutable file in exactly one shard', () => {
  const real = readdirSync(join(REPO, 'src')).map((n) => `src/${n}`).filter(isMutable);
  const shards = partition(real.map((f) => [f, statSync(join(REPO, f)).size] as [string, number]), 4);
  assert.equal(shards.flat().length, real.length);
  assert.deepEqual(new Set(shards.flat()), new Set(real));
  assert.ok(shards.every((s) => s.length > 0));
});

test('the test-process heap cap leaves four runners room on a 16 GB CI runner', () => {
  assert.ok(Number.isInteger(TEST_HEAP_MB) && TEST_HEAP_MB >= 256, String(TEST_HEAP_MB));
  assert.ok(TEST_HEAP_MB * 4 <= 8 * 1024, String(TEST_HEAP_MB));
});

test('isMutable keeps src/*.ts and drops index.ts, tests, testing/ and non-src files', () => {
  for (const f of ['src/coerce.ts', 'src/api/x.ts']) assert.ok(isMutable(f), f);
  for (const f of ['src/index.ts', 'src/coerce.test.ts', 'src/testing/mock.ts', 'scripts/a.ts', 'src/a.js', 'README.md']) {
    assert.ok(!isMutable(f), f);
  }
});

test('MUTATE_ALL excludes the same files isMutable does', () => {
  assert.deepEqual(MUTATE_ALL, ['src/**/*.ts', '!src/index.ts', '!src/**/*.test.ts', '!src/testing/**']);
});

test('testFiles keeps every src test except the seven excluded ones, sorted', () => {
  const names = ['b.test.ts', 'a.test.ts', 'a.ts', 'index-env.test.ts', 'built-server.test.ts', 'testing'];
  assert.deepEqual(testFiles(names), ['src/a.test.ts', 'src/b.test.ts']);
  assert.equal(EXCLUDED_TESTS.length, 7);
});

test('every excluded test file exists, so a rename cannot silently re-include it', () => {
  for (const f of EXCLUDED_TESTS) assert.ok(existsSync(join(REPO, f)), f);
  const real = testFiles(readdirSync(join(REPO, 'src')));
  assert.ok(real.length > 20, `only ${real.length} test files selected`);
  for (const f of EXCLUDED_TESTS) assert.ok(!real.includes(f), f);
});

test('diffToRanges turns zero-context hunks into mutate ranges for mutable files only', () => {
  const diff = [
    'diff --git a/src/a.ts b/src/a.ts',
    '--- a/src/a.ts',
    '+++ b/src/a.ts',
    '@@ -3,0 +4,2 @@ context',
    '+one',
    '+++ two, an added line whose own text starts with "++ "',
    '@@ -10 +12 @@',
    '-old',
    '+new',
    '@@ -20,3 +21,0 @@',
    '-pure',
    '-deletion',
    '-skipped',
    'diff --git a/src/gone.ts b/src/gone.ts',
    '--- a/src/gone.ts',
    '+++ /dev/null',
    '@@ -1,2 +0,0 @@',
    'diff --git a/src/index.ts b/src/index.ts',
    '+++ b/src/index.ts',
    '@@ -1 +1 @@',
    '+++ b/src/a.test.ts',
    '@@ -1 +1,4 @@',
    '+++ b/src/b.ts',
    '@@ -7,2 +7,3 @@',
  ].join('\n');
  assert.deepEqual(diffToRanges(diff), ['src/a.ts:4-5', 'src/a.ts:12-12', 'src/b.ts:7-9']);
});

let root: string;
let work: string;
let rootOnly: string;
const env = () => ({ ...process.env, GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' });

function git(cwd: string, args: string[]) {
  const r = spawnSync('git', ['-c', 'user.email=t@example.invalid', '-c', 'user.name=T', '-c', 'commit.gpgsign=false', ...args], {
    cwd, encoding: 'utf8', env: env(),
  });
  assert.equal(r.status, 0, `git ${args.join(' ')} failed:\n${r.stdout}${r.stderr}`);
  return r.stdout.trim();
}

function run(cwd: string, args: string[]) {
  const r = spawnSync(process.execPath, [join(cwd, 'scripts', 'mutation-test.mjs'), ...args], { cwd, encoding: 'utf8', env: env() });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

// The script copy is committed, since an untracked copy would make every tree dirty.
function initRepo(dir: string) {
  mkdirSync(join(dir, 'src'), { recursive: true });
  mkdirSync(join(dir, 'scripts'), { recursive: true });
  copyFileSync(join(REPO, 'scripts', 'mutation-test.mjs'), join(dir, 'scripts', 'mutation-test.mjs'));
  writeFileSync(join(dir, 'src', 'foo.ts'), 'export const foo = 1;\n');
  git(dir, ['init', '-b', 'main']);
  git(dir, ['add', '.']);
  git(dir, ['commit', '-m', 'root']);
}

before(() => {
  root = mkdtempSync(join(tmpdir(), 'mutation-test-'));
  work = join(root, 'work');
  initRepo(work);
  writeFileSync(join(work, 'src', 'foo.ts'), 'export const foo = 1;\nexport const bar = 2;\n');
  git(work, ['commit', '-am', 'second']);
  rootOnly = join(root, 'root-only');
  initRepo(rootOnly);
});

after(() => rmSync(root, { recursive: true, force: true, maxRetries: 3 }));

test('the script refuses bad arguments with its usage line', () => {
  const r = run(work, ['a', 'b']);
  assert.equal(r.status, 2, r.out);
  assert.match(r.out, /Usage: node scripts\/mutation-test\.mjs/);
});

test('the script refuses a commit that is not HEAD, an unresolvable rev and a root commit', () => {
  assert.match(run(work, ['HEAD~1']).out, /is not HEAD/);
  const bad = run(work, ['no-such-ref-xyz']);
  assert.equal(bad.status, 2, bad.out);
  assert.match(bad.out, /no-such-ref-xyz does not resolve/);
  assert.doesNotMatch(bad.out, /fatal:/);
  assert.match(run(rootOnly, ['HEAD']).out, /root commit/);
});

test('the script refuses a dirty tree, including an untracked file', () => {
  writeFileSync(join(work, 'untracked.txt'), 'scratch\n');
  try {
    const r = run(work, ['HEAD']);
    assert.equal(r.status, 2, r.out);
    assert.match(r.out, /not clean[\s\S]*untracked\.txt/);
  } finally {
    unlinkSync(join(work, 'untracked.txt'));
  }
});

test('a clean HEAD commit passes every refusal and reaches the dependency check', () => {
  const r = run(work, ['main']);
  assert.equal(r.status, 2, r.out);
  assert.match(r.out, /No node_modules/);
});
