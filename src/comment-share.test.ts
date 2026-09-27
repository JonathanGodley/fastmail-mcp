/**
 * Tests for scripts/comment-share.mjs, the comment-density report the
 * pre-commit hook prints. A throwaway git repo is built under the OS temp dir
 * with the script copied in, so the added-line counts and baselines are what
 * git actually reports. The repo's own working tree is never touched.
 */

import {
  test, before, beforeEach, after,
} from 'node:test';
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
// Labelled variants so a rewrite's old and new lines never share text with
// each other, and git's diff reports them as real removes and real adds
// rather than matching them as unchanged.
const labelled = (n: number, label: string) => Array.from({ length: n }, (_, i) => `// ${label} ${i}`);
const code2 = (n: number) => Array.from({ length: n }, (_, i) => `let w${i} = ${i};`);

before(() => {
  root = mkdtempSync(join(tmpdir(), 'comment-share-'));
  work = join(root, 'work');
  mkdirSync(join(work, 'scripts'), { recursive: true });
  copyFileSync(join(REPO, 'scripts', 'comment-share.mjs'), join(work, 'scripts', 'comment-share.mjs'));
  git(['init', '-b', 'main']);
  writeFileSync(join(work, 'a.ts'), block(20, 20));
  // Track the script fixture too, so beforeEach's clean (untracked files only)
  // never removes the very thing report() spawns.
  git(['add', 'a.ts', 'scripts/comment-share.mjs']);
  git(['commit', '-m', 'base']);
});

after(() => {
  // Windows holds handles on a just-used repo briefly; the retry option covers it.
  rmSync(root, { recursive: true, force: true, maxRetries: 3 });
});

// Clears staged and untracked leftovers before every case, so a case that
// throws mid-test can't corrupt the next. Commits survive it.
beforeEach(() => {
  git(['reset', '--hard', 'HEAD']);
  git(['clean', '-fdx']);
});

test('--staged reports a change far above the file\'s own density, exit 0', () => {
  appendFileSync(join(work, 'a.ts'), block(25, 5));
  git(['add', 'a.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: a.ts: +25 / -0 comment lines, +5 / -0 code (net +25 comment against +5 code; 5.0 per code line; the file runs 1.0).'), r.out);
  assert.ok(r.out.includes('/trim pass'), r.out);
});

test('a commit is measured against its first parent', () => {
  // beforeEach wipes uncommitted state, so this stages its own change; the
  // commit does carry forward to later cases.
  appendFileSync(join(work, 'a.ts'), block(25, 5));
  git(['add', 'a.ts']);
  git(['commit', '-m', 'over the bar']);
  const sha = git(['rev-parse', 'HEAD']);
  const r = report([sha]);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: a.ts: +25 / -0 comment lines, +5 / -0 code (net +25 comment against +5 code'), r.out);
});

test('prints nothing when the change is under the bar', () => {
  appendFileSync(join(work, 'a.ts'), block(10, 10));
  git(['add', 'a.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.equal(r.out, '');
});

test('--all prints the under-the-bar figures too', () => {
  appendFileSync(join(work, 'a.ts'), block(10, 10));
  git(['add', 'a.ts']);
  const r = report(['--staged', '--all']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share (under the bar): a.ts: +10 / -0 comment lines, +10 / -0 code (net +10 comment against +10 code; 1.0 per code line; the file runs'), r.out);
  assert.ok(!r.out.includes('/trim pass'),'no nudge when nothing is over the bar');
});

test('a new file with more comment than code fires at 1.0', () => {
  writeFileSync(join(work, 'b.ts'), block(25, 10));
  git(['add', 'b.ts']);
  const r = report(['--staged']);
  assert.ok(r.out.includes('comment-share: b.ts: +25 / -0 comment lines, +10 / -0 code (net +25 comment against +10 code; 2.5 per code line; new file, no baseline).'), r.out);
});

test('markdown is not measured', () => {
  writeFileSync(join(work, 'notes.md'), comments(40).join('\n') + '\n');
  git(['add', 'notes.md']);
  const r = report(['--staged']);
  assert.equal(r.out, '');
});

test('no arguments prints usage and still exits 0', () => {
  const r = report([]);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('usage:'), r.out);
});

test('a pure trim (removes more comment than it adds) stays silent', () => {
  // 60 old comment lines replaced with 25 new ones, code untouched. The added
  // comment alone would be over the bar; net comment is negative, so this must
  // never fire.
  writeFileSync(join(work, 'trim.ts'), labelled(60, 'old').concat(code(20)).join('\n') + '\n');
  git(['add', 'trim.ts']);
  git(['commit', '-m', 'trim baseline']);
  writeFileSync(join(work, 'trim.ts'), labelled(25, 'new').concat(code(20)).join('\n') + '\n');
  git(['add', 'trim.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.equal(r.out, '');
});

test('a rewrite that nets +25 comment and 0 code fires, with a breakdown', () => {
  // 5 old comment lines replaced with 30 new ones, code untouched.
  writeFileSync(join(work, 'netcomment.ts'), labelled(5, 'old').concat(code(20)).join('\n') + '\n');
  git(['add', 'netcomment.ts']);
  git(['commit', '-m', 'netcomment baseline']);
  writeFileSync(join(work, 'netcomment.ts'), labelled(30, 'new').concat(code(20)).join('\n') + '\n');
  git(['add', 'netcomment.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: netcomment.ts: +30 / -5 comment lines, +0 / -0 code (net +25 comment and no added code; the file runs 0.3).'), r.out);
});

test('mixed adds and removes are judged on the net against the baseline', () => {
  // Removes 15 old comment lines, adds 40 new comment lines and 5 new code
  // lines; the 25 existing code lines are untouched. Judged on the net
  // (+25 comment, +5 code) against the file's own baseline (15/25 = 0.6).
  writeFileSync(join(work, 'mixed.ts'), labelled(15, 'old').concat(code(25)).join('\n') + '\n');
  git(['add', 'mixed.ts']);
  git(['commit', '-m', 'mixed baseline']);
  writeFileSync(join(work, 'mixed.ts'), code(25).concat(labelled(40, 'new')).concat(code2(5)).join('\n') + '\n');
  git(['add', 'mixed.ts']);
  const r = report(['--staged']);
  assert.equal(r.status, 0);
  assert.ok(r.out.includes('comment-share: mixed.ts: +40 / -15 comment lines, +5 / -0 code (net +25 comment against +5 code; 5.0 per code line; the file runs 0.6).'), r.out);
});
