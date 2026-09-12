/**
 * Tests for scripts/comment-share.mjs, the comment-density report the
 * pre-commit hook prints. A throwaway git repo is built under the OS temp dir
 * with the script copied in, so the added-line counts and baselines are what
 * git actually reports. The repo's own working tree is never touched.
 */

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import {
  mkdtempSync, mkdirSync, writeFileSync, appendFileSync, copyFileSync, rmSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');

let root: string;
let work: string;

function git(args: string[]) {
  const r = spawnSync('git', ['-c', 'user.email=t@example.invalid', '-c', 'user.name=T', '-c', 'commit.gpgsign=false', ...args], {
    cwd: work,
    encoding: 'utf8',
    env: { ...process.env, GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' },
  });
  assert.equal(r.status, 0, `git ${args.join(' ')} failed:\n${r.stdout}${r.stderr}`);
  return r.stdout.trim();
}

function report(args: string[]) {
  const r = spawnSync(process.execPath, [join(work, 'scripts', 'comment-share.mjs'), ...args], {
    cwd: work,
    encoding: 'utf8',
    env: { ...process.env, GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' },
  });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

const comments = (n: number) => Array.from({ length: n }, (_, i) => `// note ${i}`);
const code = (n: number) => Array.from({ length: n }, (_, i) => `const v${i} = ${i};`);
const block = (c: number, k: number) => comments(c).concat(code(k)).join('\n') + '\n';

before(() => {
  root = mkdtempSync(join(tmpdir(), 'comment-share-'));
  work = join(root, 'work');
  mkdirSync(join(work, 'scripts'), { recursive: true });
  copyFileSync(join(REPO, 'scripts', 'comment-share.mjs'), join(work, 'scripts', 'comment-share.mjs'));
  git(['init', '-b', 'main']);
  writeFileSync(join(work, 'a.ts'), block(20, 20));
  git(['add', 'a.ts']);
  git(['commit', '-m', 'base']);
});

after(() => {
  rmSync(root, { recursive: true, force: true });
});

test('--staged reports a change far above the file\'s own density, exit 0', () => {
  appendFileSync(join(work, 'a.ts'), block(25, 5));
  git(['add', 'a.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: a.ts: +25 comment lines against +5 code (5.0 per code line; the file runs 1.0).'), r.out);
  assert.ok(r.out.includes('/tidy-comments'), r.out);
});

test('a commit is measured against its first parent', () => {
  git(['commit', '-m', 'over the bar']);
  const sha = git(['rev-parse', 'HEAD']);
  const r = report([sha]);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: a.ts: +25 comment lines against +5 code'), r.out);
});

test('prints nothing when the change is under the bar', () => {
  appendFileSync(join(work, 'a.ts'), block(10, 10));
  git(['add', 'a.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.equal(r.out, '');
});

test('--all prints the under-the-bar figures too', () => {
  const r = report(['--staged', '--all']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share (under the bar): a.ts: +10 comment lines against +10 code (1.0 per code line; the file runs'), r.out);
  assert.ok(!r.out.includes('/tidy-comments'), 'no nudge when nothing is over the bar');
  git(['reset', '--hard', 'HEAD']);
});

test('a new file with more comment than code fires at 1.0', () => {
  writeFileSync(join(work, 'b.ts'), block(25, 10));
  git(['add', 'b.ts']);
  const r = report(['--staged']);
  assert.ok(r.out.includes('comment-share: b.ts: +25 comment lines against +10 code (2.5 per code line; new file, no baseline).'), r.out);
  git(['reset', '--hard', 'HEAD']);
  rmSync(join(work, 'b.ts'), { force: true });
});

test('markdown is not measured', () => {
  writeFileSync(join(work, 'notes.md'), comments(40).join('\n') + '\n');
  git(['add', 'notes.md']);
  const r = report(['--staged']);
  assert.equal(r.out, '');
  git(['reset', '--hard', 'HEAD']);
  rmSync(join(work, 'notes.md'), { force: true });
});

test('no arguments prints usage and still exits 0', () => {
  const r = report([]);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('usage:'), r.out);
});
