/**
 * Tests for scripts/mutate-commit.mjs's caller-facing refusals: naming a commit that is not
 * the tree's current HEAD, a dirty tree, `--tests` misuse, an unresolvable rev, and a root
 * commit. None of these reach Stryker, so no node_modules is needed - a bare git fixture is
 * enough, and every case's current-code behaviour is the same dependency-gate message
 * ("No node_modules in ...") since that gate runs before any of these refusals exist today.
 *
 * The script is copied into a throwaway git repo (as scripts/mutate-commit.mjs, so its
 * self-located REPO resolves there) and, unlike scripts/comment-share.mjs's fixture, that
 * copy is COMMITTED: the refusal this suite exists to test is "the tree must be clean at
 * HEAD", so an untracked copy of the script itself would make every case dirty before its
 * own scenario does.
 */

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import {
  mkdtempSync, mkdirSync, writeFileSync, appendFileSync, copyFileSync, rmSync, unlinkSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');

let root: string;
let work: string; // multi-commit fixture: root commit, then two more, HEAD at the third
let rootWork: string; // single-commit fixture, used only for the root-commit case: HEAD IS the root

function git(cwd: string, args: string[]) {
  const r = spawnSync('git', ['-c', 'user.email=t@example.invalid', '-c', 'user.name=T', '-c', 'commit.gpgsign=false', ...args], {
    cwd,
    encoding: 'utf8',
    env: { ...process.env, GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' },
  });
  assert.equal(r.status, 0, `git ${args.join(' ')} failed:\n${r.stdout}${r.stderr}`);
  return r.stdout.trim();
}

function run(cwd: string, args: string[]) {
  const r = spawnSync(process.execPath, [join(cwd, 'scripts', 'mutate-commit.mjs'), ...args], {
    cwd,
    encoding: 'utf8',
    env: { ...process.env, GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' },
  });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

function initRepo(dir: string) {
  mkdirSync(join(dir, 'src'), { recursive: true });
  mkdirSync(join(dir, 'scripts'), { recursive: true });
  copyFileSync(join(REPO, 'scripts', 'mutate-commit.mjs'), join(dir, 'scripts', 'mutate-commit.mjs'));
  writeFileSync(join(dir, 'src', 'foo.ts'), 'export const foo = 1;\n');
  writeFileSync(join(dir, 'src', 'foo.test.ts'), "import { test } from 'node:test';\ntest('noop', () => {});\n");
  git(dir, ['init', '-b', 'main']);
  git(dir, ['add', '.']);
  git(dir, ['commit', '-m', 'root']);
}

before(() => {
  root = mkdtempSync(join(tmpdir(), 'mutate-commit-'));

  work = join(root, 'work');
  initRepo(work);
  const rootSha = git(work, ['rev-parse', 'HEAD']);
  writeFileSync(join(work, 'src', 'foo.ts'), 'export const foo = 1;\nexport const bar = 2;\n');
  git(work, ['add', '.']);
  git(work, ['commit', '-m', 'second']);
  const secondSha = git(work, ['rev-parse', 'HEAD']);
  writeFileSync(join(work, 'src', 'foo.ts'), 'export const foo = 1;\nexport const bar = 2;\nexport const baz = 3;\n');
  git(work, ['add', '.']);
  git(work, ['commit', '-m', 'third']);
  const headSha = git(work, ['rev-parse', 'HEAD']);
  (globalThis as any).__fixtureShas = { rootSha, secondSha, headSha };

  // Separate single-commit repo: the only way to reach the root-commit refusal is a root
  // commit that is ALSO HEAD, since a root commit that is not HEAD is refused for that
  // reason first (see D4's ordering). A second repo keeps that untangled from `work`'s
  // multi-commit non-HEAD case rather than checking `work`'s own root out detached.
  rootWork = join(root, 'root-work');
  initRepo(rootWork);
});

after(() => {
  // Windows briefly holds a handle on a just-used repo; the retry option covers it.
  rmSync(root, { recursive: true, force: true, maxRetries: 3 });
});

test('refuses a commit that is not the tree\'s current HEAD', () => {
  const { secondSha, headSha } = (globalThis as any).__fixtureShas;
  const r = run(work, [secondSha]);
  assert.equal(r.status, 2, r.out);
  assert.ok(r.out.includes(work), r.out);
  assert.ok(r.out.includes(secondSha), r.out);
  assert.ok(r.out.includes(headSha), r.out);
  assert.ok(/HEAD/.test(r.out), r.out);
  assert.ok(/detached|worktree/.test(r.out), r.out);
  assert.ok(!r.out.includes('No node_modules'), r.out);
});

test('refuses a dirty tree at HEAD, including an untracked file', () => {
  const { headSha } = (globalThis as any).__fixtureShas;
  writeFileSync(join(work, 'untracked.txt'), 'scratch\n');
  try {
    const r = run(work, [headSha]);
    assert.equal(r.status, 2, r.out);
    assert.ok(/porcelain/i.test(r.out), r.out);
    assert.ok(/stash/i.test(r.out) && /untracked/i.test(r.out), r.out);
    assert.ok(!r.out.includes('No node_modules'), r.out);
  } finally {
    unlinkSync(join(work, 'untracked.txt'));
  }
});

test('two bare commit arguments is refused as a plain typo, not a --tests mistake', () => {
  const { headSha } = (globalThis as any).__fixtureShas;
  const r = run(work, [headSha, 'some-other-arg']);
  assert.equal(r.status, 2, r.out);
  assert.ok(/exactly one commit/i.test(r.out), r.out);
  assert.ok(!r.out.includes('--tests=a,b') && !/comma-separated --tests/i.test(r.out), r.out);
  assert.ok(!r.out.includes('No node_modules'), r.out);
});

test('an unresolvable rev is refused by name, not a git stack trace', () => {
  const r = run(work, ['no-such-ref-xyz']);
  assert.equal(r.status, 2, r.out);
  assert.ok(r.out.includes('no-such-ref-xyz'), r.out);
  assert.ok(!/at Object|node:internal|throw/.test(r.out), r.out);
  assert.ok(!r.out.includes('No node_modules'), r.out);
});

test('a root commit that is HEAD is refused - nothing to diff against', () => {
  const rootSha = git(rootWork, ['rev-parse', 'HEAD']);
  const r = run(rootWork, [rootSha]);
  assert.equal(r.status, 2, r.out);
  assert.ok(/root commit/i.test(r.out), r.out);
  assert.ok(!r.out.includes('No node_modules'), r.out);
});

test('an empty --tests= value is refused, not silently name-matched', () => {
  const { headSha } = (globalThis as any).__fixtureShas;
  const r = run(work, [headSha, '--tests=']);
  assert.equal(r.status, 2, r.out);
  assert.ok(/--tests/.test(r.out) && /value|empty/i.test(r.out), r.out);
  assert.ok(!r.out.includes('No node_modules'), r.out);
});

test('--tests=a,b is parsed as an override, echoed before the dependency gate', () => {
  const { headSha } = (globalThis as any).__fixtureShas;
  const r = run(work, [headSha, '--tests=src/whatever-a.test.ts,src/whatever-b.test.ts']);
  // Bare fixture, no node_modules: this run cannot get past the dependency gate, so its exit
  // status is the gate's (1), same as an unrelated dependency failure would give. What is
  // asserted is that the gate is REACHED at all (proving --tests parsed and every refusal
  // above passed) and that the parsed list was echoed before it - observable only because
  // that echo runs ahead of the dependency gate.
  assert.equal(r.status, 1, r.out);
  assert.ok(r.out.includes('src/whatever-a.test.ts'), r.out);
  assert.ok(r.out.includes('src/whatever-b.test.ts'), r.out);
  assert.ok(r.out.includes('No node_modules'), r.out);
});
