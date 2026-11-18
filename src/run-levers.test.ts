/**
 * Tests for scripts/run-levers.mjs. Each test builds a throwaway git repo under
 * the OS temp dir holding a tiny target file and a fixture test that reads it,
 * and runs the repo's script with that repo as its working directory. No real
 * source file is ever a target.
 */

import { test, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { constants, tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const SCRIPT = join(dirname(fileURLToPath(import.meta.url)), '..', 'scripts', 'run-levers.mjs');

// CRLF on purpose: the restore must give back the bytes, not a normalised copy.
const TARGET = Buffer.from('first line\r\nstatus = ok\r\nxaaay\r\nlast line\r\n');

// Line 8 is the assertion a RED lever must be reported at, and line 18 the call that throws.
const FIXTURE_TEST = [
  "import { test } from 'node:test';",
  "import assert from 'node:assert/strict';",
  "import { readFileSync, writeFileSync } from 'node:fs';",
  "writeFileSync('ran', '');",
  "test('target is ok', async () => {",
  "  const text = readFileSync('target.txt', 'utf8');",
  "  if (text.includes('HANG')) { writeFileSync('hanging', ''); await new Promise((r) => setTimeout(r, 60000)); }",
  "  assert.ok(text.includes('status = ok'), 'status changed');",
  '});',
  "test('always fails', () => {",
  "  assert.fail('broken before any lever');",
  '});',
  "test('fails without a stack when told', () => {",
  "  if (readFileSync('target.txt', 'utf8').includes('STACKLESS')) { const e = new assert.AssertionError({ message: 'x' }); e.stack = 'AssertionError: x'; throw e; }",
  '});',
  "test('throws when told', () => {",
  "  const text = readFileSync('target.txt', 'utf8');",
  "  if (text.includes('THROW')) JSON.parse('{');",
  "  assert.ok(text.includes('status'));",
  '});',
  '',
].join('\n');

let work: string;

function git(args: string[]) {
  const r = spawnSync('git', ['-c', 'user.email=t@example.invalid', '-c', 'user.name=T', '-c', 'commit.gpgsign=false', ...args], {
    cwd: work,
    encoding: 'utf8',
    env: { ...process.env, GIT_CONFIG_GLOBAL: join(work, 'no-global-gitconfig'), GIT_CONFIG_NOSYSTEM: '1' },
  });
  assert.equal(r.status, 0, `git ${args.join(' ')} failed:\n${r.stdout}${r.stderr}`);
}

function lever(over: Record<string, string | number> = {}) {
  return {
    label: 'status lever', target: 'target.txt', find: 'status = ok', replace: 'status = no',
    test: 'fixture.test.mjs', pattern: 'target is ok', ...over,
  };
}

function writeLevers(levers: unknown[]) {
  const path = join(work, 'levers.json');
  writeFileSync(path, JSON.stringify(levers));
  return path;
}

function run(levers: unknown[]) {
  const r = spawnSync(process.execPath, [SCRIPT, writeLevers(levers)], { cwd: work, encoding: 'utf8' });
  return { status: r.status, out: `${r.stdout ?? ''}${r.stderr ?? ''}` };
}

beforeEach(() => {
  work = mkdtempSync(join(tmpdir(), 'run-levers-'));
  writeFileSync(join(work, 'target.txt'), TARGET);
  writeFileSync(join(work, 'fixture.test.mjs'), FIXTURE_TEST);
  writeFileSync(join(work, '.gitignore'), 'ran\nhanging\nlevers.json\n');
  git(['init', '-q']);
  git(['add', 'target.txt', 'fixture.test.mjs', '.gitignore']);
  git(['commit', '-q', '-m', 'fixture']);
});

afterEach(() => {
  rmSync(work, { recursive: true, force: true });
});

test('a lever that turns the test red reports the test and the assertion line, and restores the bytes', () => {
  const r = run([lever()]);
  assert.equal(r.status, 0, r.out);
  assert.match(r.out, /^status lever: RED - "target is ok" at fixture\.test\.mjs:8$/m);
  assert.ok(readFileSync(join(work, 'target.txt')).equals(TARGET));
});

test('a lever the test does not notice is GREEN and fails the run', () => {
  const r = run([lever({ find: 'last line', replace: 'final line' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: GREEN - /m);
  assert.ok(readFileSync(join(work, 'target.txt')).equals(TARGET));
});

test('a lever that makes the code throw is THREW, not RED, and fails the run', () => {
  const r = run([lever({ replace: 'status = THROW', pattern: 'throws when told' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: THREW - "throws when told" threw SyntaxError at fixture\.test\.mjs:18, not at an assertion$/m);
  assert.ok(readFileSync(join(work, 'target.txt')).equals(TARGET));
});

test('a find that does not occur exactly once is INVALID and the run goes on to the next lever', () => {
  const r = run([lever({ label: 'twice', find: 'line' }), lever({ label: 'never', find: 'absent' }), lever()]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^twice: INVALID - find occurs 2 times in target\.txt, not once$/m);
  assert.match(r.out, /^never: INVALID - find occurs 0 times in target\.txt, not once$/m);
  assert.match(r.out, /^status lever: RED - /m);
});

test('overlapping occurrences of find are counted', () => {
  const r = run([lever({ find: 'aa', replace: 'b' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - find occurs 2 times in target\.txt, not once$/m);
});

test('a pattern that matches no test is INVALID', () => {
  const r = run([lever({ pattern: 'no such test' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - no test in fixture\.test\.mjs matches "no such test"$/m);
});

test('a lever after which the pattern matches no test is INVALID, not GREEN', () => {
  const r = run([lever({ target: 'fixture.test.mjs', find: "test('target is ok'", replace: "test('renamed'" })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - with the lever in place no test matches "target is ok"$/m);
});

test('a test that already fails without the lever is INVALID, not RED', () => {
  const r = run([lever({ pattern: 'always fails' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - the test fails without the lever$/m);
});

test('a lever that stops the test file loading is INVALID, not RED', () => {
  const r = run([lever({ target: 'fixture.test.mjs', find: "test('target is ok', async () => {", replace: "test('target is ok', async () => {{" })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - the run failed \(exit 1\) but no matching test did/m);
});

test('a failure with no stack frame in the repo is INVALID, not RED', () => {
  const r = run([lever({ replace: 'status = STACKLESS', pattern: 'fails without a stack' })]);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: INVALID - "fails without a stack when told" failed with no stack frame in the repo$/m);
});

test('a run past timeoutMs is TIMEOUT, the script exits promptly, and the target is restored', () => {
  const started = Date.now();
  const r = run([lever({ replace: 'status = HANG', timeoutMs: 5000 })]);
  // The fixture hangs for 60 s; a test runner left running would hold the script that long.
  assert.ok(Date.now() - started < 30000, `the script took ${Date.now() - started} ms to exit`);
  assert.equal(r.status, 1, r.out);
  assert.match(r.out, /^status lever: TIMEOUT - the test ran past 5000 ms with the lever in place$/m);
  assert.ok(readFileSync(join(work, 'target.txt')).equals(TARGET));
});

test('it refuses to start when a target has uncommitted changes, and runs no test', () => {
  writeFileSync(join(work, 'target.txt'), 'status = ok\nwork in progress\n');
  const r = run([lever()]);
  assert.equal(r.status, 2, r.out);
  assert.match(r.out, /refusing to start/);
  assert.ok(!existsSync(join(work, 'ran')), 'the fixture test ran');
});

// On Windows, kill() ends a child outright, with no signal for it to catch.
for (const signal of ['SIGINT', 'SIGTERM'] as const) {
  test(`${signal} mid-run restores the target`, { skip: process.platform === 'win32' }, async () => {
    const child = spawn(process.execPath, [SCRIPT, writeLevers([lever({ replace: 'status = HANG' })])], { cwd: work });
    let out = '';
    child.stdout.on('data', (b) => { out += b; });
    child.stderr.on('data', (b) => { out += b; });
    const closed = new Promise<number | null>((done) => child.on('close', done));
    const deadline = Date.now() + 30000;
    while (!existsSync(join(work, 'hanging'))) {
      assert.ok(Date.now() < deadline, `the lever run never started:\n${out}`);
      await new Promise((r) => setTimeout(r, 50));
    }
    assert.ok(!readFileSync(join(work, 'target.txt')).equals(TARGET), 'the lever should be in place');
    child.kill(signal);
    assert.equal(await closed, 128 + constants.signals[signal], out);
    assert.ok(readFileSync(join(work, 'target.txt')).equals(TARGET));
  });
}
