/**
 * End-to-end tests for scripts/scan-secrets.mjs and the .githooks/ that call
 * it. A throwaway git repo is built under the OS temp dir with the hooks and
 * scanner copied in, a bare "remote" beside it, and commits, tags and pushes
 * are driven through real git so that what is tested is what a developer's
 * git actually does: the commit-msg hook refusing a message, the pre-commit
 * hook refusing staged content, and the pre-push hook refusing a commit that
 * slipped past --no-verify or a tag whose annotation carries personal data.
 *
 * Every value here is synthetic. The "personal" address is under a domain that
 * is not on the scanner's safe list because that is the condition under test;
 * the phone number is nobody's; the denylisted name is a placeholder.
 *
 * The repo's own working tree is never touched: nothing here reads or writes
 * outside the temp dir except to copy the scanner and hooks out of this repo.
 */

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import {
  mkdtempSync, mkdirSync, writeFileSync, copyFileSync, chmodSync, rmSync, readdirSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');

const PERSONAL = 'sample.person@fictional-domain.org'; // allowlist-secret (synthetic: invented address on an unregistered-looking domain, deliberately off the safe list)
const PHONE = '0400 111 222'; // allowlist-secret (synthetic: nobody's number, the shape under test)
const DENIED = 'placeholdersurname';

let root: string;
let work: string;

function git(args: string[], opts: { cwd?: string; input?: string } = {}) {
  const r = spawnSync('git', args, {
    cwd: opts.cwd ?? work,
    input: opts.input,
    encoding: 'utf8',
    env: {
      ...process.env,
      // Keep the developer's own git config out of the test: no templates,
      // signing, or hooks path of theirs, and no autocrlf surprises.
      GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'),
      GIT_CONFIG_NOSYSTEM: '1',
    },
  });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

function scanner(args: string[], input?: string) {
  const r = spawnSync(process.execPath, [join(work, 'scripts', 'scan-secrets.mjs'), ...args], {
    cwd: work,
    input,
    encoding: 'utf8',
    env: {
      ...process.env,
      GIT_CONFIG_GLOBAL: join(root, 'no-global-gitconfig'),
      GIT_CONFIG_NOSYSTEM: '1',
    },
  });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

function ok(r: { status: number | null; out: string }, what: string) {
  assert.equal(r.status, 0, `${what} should succeed, got exit ${r.status}:\n${r.out}`);
}

function refused(r: { status: number | null; out: string }, what: string, ...needles: string[]) {
  assert.notEqual(r.status, 0, `${what} should be refused, but exit was 0:\n${r.out}`);
  for (const n of needles) {
    assert.ok(r.out.includes(n), `${what}: expected output to mention ${JSON.stringify(n)}:\n${r.out}`);
  }
  // The matched values must never be echoed back.
  for (const secret of [PERSONAL, PHONE, DENIED]) {
    assert.ok(!r.out.toLowerCase().includes(secret.toLowerCase()),
      `${what}: output must not echo the protected value:\n${r.out}`);
  }
}

before(() => {
  root = mkdtempSync(join(tmpdir(), 'scan-secrets-test-'));
  work = join(root, 'work');
  const bare = join(root, 'remote.git');
  mkdirSync(work);

  ok(git(['init', '-q', '--initial-branch=main'], { cwd: work }), 'git init');
  ok(git(['init', '-q', '--bare', bare], { cwd: root }), 'git init --bare');
  for (const [k, v] of [
    ['user.name', 'Test'], ['user.email', 'test@example.com'],
    ['commit.gpgsign', 'false'], ['tag.gpgsign', 'false'],
    ['core.hooksPath', '.githooks'],
  ]) ok(git(['config', k, v]), `git config ${k}`);
  ok(git(['remote', 'add', 'origin', bare]), 'git remote add');

  // The scanner and hooks, laid out as they are in this repo, because the
  // hooks locate the scanner relative to themselves.
  mkdirSync(join(work, 'scripts'));
  mkdirSync(join(work, '.githooks'));
  copyFileSync(join(REPO, 'scripts', 'scan-secrets.mjs'), join(work, 'scripts', 'scan-secrets.mjs'));
  for (const hook of readdirSync(join(REPO, '.githooks'))) {
    const dest = join(work, '.githooks', hook);
    copyFileSync(join(REPO, '.githooks', hook), dest);
    chmodSync(dest, 0o755);
  }

  // A denylist outside the repo, wired in the way CONTRIBUTING.md describes.
  const deny = join(root, 'deny.txt');
  writeFileSync(deny, `# test denylist\n${DENIED}\n`);
  ok(git(['config', 'secretscan.denylist', deny]), 'git config secretscan.denylist');

  writeFileSync(join(work, 'README.md'), 'A clean file. Contact alice@example.com.\n');
  ok(git(['add', 'README.md', 'scripts/scan-secrets.mjs', '.githooks']), 'git add');
  ok(git(['commit', '-q', '-m', 'Initial commit'], {}), 'initial commit through the hooks');
  ok(git(['push', '-q', '-u', 'origin', 'main']), 'initial push through the pre-push hook');
});

after(() => {
  rmSync(root, { recursive: true, force: true });
});

test('commit-msg hook: a clean message commits', () => {
  ok(git(['commit', '-q', '--allow-empty', '-m', 'Tidy the README wording']), 'clean commit');
});

test('commit-msg hook: a phone number in the message is refused without being echoed', () => {
  refused(git(['commit', '--allow-empty', '-m', `Call me on ${PHONE} about this`]),
    'commit with a phone number', '[possible personal phone number]', 'message:1');
});

test('commit-msg hook: a personal address in the message body is refused', () => {
  refused(git(['commit', '--allow-empty', '-m', 'Fix reply quoting', '-m', `Reported by ${PERSONAL}`]),
    'commit with an address in the body', '[possible personal email]');
});

test('commit-msg hook: a denylisted string escalates rather than asking for a rewrite', () => {
  refused(git(['commit', '--allow-empty', '-m', `Handle mail from ${DENIED.toUpperCase()}`]),
    'commit with a denylisted name', '[local denylist match]', 'ESCALATE TO THE USER');
});

test('pre-commit hook: staged content with a token is refused, staged content that is clean commits', () => {
  writeFileSync(join(work, 'notes.md'), 'token = "fmu1-abcdefghijklmnopqrstuvwxyz012345"\n'); // allowlist-secret (synthetic token shape)
  ok(git(['add', 'notes.md']), 'git add dirty file');
  refused(git(['commit', '-m', 'Add notes']), 'commit of staged token', 'notes.md:1');
  writeFileSync(join(work, 'notes.md'), 'Nothing to see.\n');
  ok(git(['add', 'notes.md']), 'git add clean file');
  ok(git(['commit', '-q', '-m', 'Add notes']), 'commit of clean file');
});

test('pre-push hook: a commit that bypassed commit-msg is caught when pushed', () => {
  ok(git(['commit', '-q', '--allow-empty', '--no-verify', '-m', `Reported by ${PERSONAL}`]),
    'commit --no-verify (deliberate bypass under test)');
  refused(git(['push', 'origin', 'main']), 'push of a commit with a bad message',
    '[possible personal email]', 'message:1');
  ok(git(['reset', '-q', '--hard', 'HEAD~1']), 'drop the bad commit');
  ok(git(['push', '-q', 'origin', 'main']), 'push of the remaining clean commits');
});

test('pre-push hook: content that bypassed pre-commit is caught when pushed', () => {
  writeFileSync(join(work, 'fixture.txt'), `From: Someone <${PERSONAL}>\n`);
  ok(git(['add', 'fixture.txt']), 'git add');
  ok(git(['commit', '-q', '--no-verify', '-m', 'Add a fixture']), 'commit --no-verify (deliberate bypass under test)');
  refused(git(['push', 'origin', 'main']), 'push of a commit with bad content', ':fixture.txt:1');
  ok(git(['reset', '-q', '--hard', 'HEAD~1']), 'drop the bad commit');
});

test('pre-push hook: an annotated tag whose message carries personal data is refused', () => {
  ok(git(['tag', '-a', 'v0.0.1-test', '-m', `Release notes. Thanks to ${PERSONAL}.`]), 'git tag -a');
  refused(git(['push', 'origin', 'v0.0.1-test']), 'push of a tag with a bad annotation',
    'tag v0.0.1-test:1', '[possible personal email]');
  ok(git(['tag', '-d', 'v0.0.1-test']), 'delete the tag');
  ok(git(['tag', '-a', 'v0.0.2-test', '-m', 'Release notes. Closes #1.']), 'git tag -a clean');
  ok(git(['push', '-q', 'origin', 'v0.0.2-test']), 'push of a clean tag');
});

test('pre-push hook: deleting a remote ref publishes nothing and is not refused', () => {
  ok(git(['push', '-q', 'origin', '--delete', 'v0.0.2-test']), 'push --delete');
});

test('--message mode reads the file it is given and fails closed on a missing one', () => {
  const msg = join(root, 'msg.txt');
  writeFileSync(msg, 'A perfectly ordinary message.\n');
  ok(scanner(['--message', msg]), '--message on a clean file');
  writeFileSync(msg, `Signature: ${PHONE}\n`);
  refused(scanner(['--message', msg]), '--message on a dirty file', '[possible personal phone number]');
  refused(scanner(['--message', join(root, 'does-not-exist.txt')]), '--message on a missing file', 'could not be read');
});

test('a configured denylist that cannot be read fails the run instead of scanning without it', () => {
  ok(git(['config', 'secretscan.denylist', join(root, 'missing-deny.txt')]), 'point config at a missing file');
  try {
    const r = scanner(['--all']);
    assert.notEqual(r.status, 0, `expected failure, got:\n${r.out}`);
    assert.ok(r.out.includes('could not be read'), r.out);
  } finally {
    ok(git(['config', 'secretscan.denylist', join(root, 'deny.txt')]), 'restore config');
  }
});
