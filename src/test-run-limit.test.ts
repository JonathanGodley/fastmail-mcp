/**
 * `npm test` must end on its own, however a test file misbehaves, and leave no process behind
 * (scripts/run-tests.mjs says why the limit sits outside the runner). These run the `test`
 * script from package.json with its glob swapped for one fixture in a temp directory, where
 * the suite's own glob can never pick it up.
 */

import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { basename, delimiter, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DETACHED, killTree } from '../scripts/kill-tree.mjs';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');
const TEST_GLOB = 'src/*.test.ts';
const LIMIT_MS = 3000;
// Past the limit by far more than tsx's start-up, so only a run nothing ended reaches it.
const DEADLINE_MS = LIMIT_MS + 30_000;

const FIXTURES = {
  'never-ends.test.mjs': `import { test } from 'node:test';
test('never ends', () => new Promise(() => { setInterval(() => {}, 1000); }));
`,
  // A V8 serializer header and a size promising 1000 bytes, then 3 bytes and exit: what the
  // runner sees when a file's process stops partway through a report message.
  'cut-short.test.mjs': `import v8 from 'node:v8';
const s = new v8.Serializer(); s.writeHeader();
const size = Buffer.alloc(4); size.writeUInt32BE(1000);
process.stdout.write(Buffer.concat([s.releaseBuffer(), size, Buffer.from([1, 2, 3])]), () => process.exit(0));
`,
  'passes.test.mjs': `import { test } from 'node:test';
test('passes', () => {});
`,
  'fails.test.mjs': `import { test } from 'node:test';
test('fails', () => { throw new Error('fixture failure'); });
`,
};

let dir: string;
before(() => {
  dir = mkdtempSync(join(tmpdir(), 'test-run-limit-'));
  for (const [name, body] of Object.entries(FIXTURES)) writeFileSync(join(dir, name), body);
});
// A failing run here can leave a runner spinning; it must not outlive the suite.
after(() => {
  for (const pid of survivors()) {
    try { process.kill(Number(pid), 'SIGKILL'); } catch { /* already gone */ }
  }
  rmSync(dir, { recursive: true, force: true });
});

// npm puts node_modules/.bin on PATH for a script; a worktree's node_modules is an ancestor's.
function envForScript(limited: boolean): NodeJS.ProcessEnv {
  const bin = join(dirname(createRequire(import.meta.url).resolve('tsx/cli')), '..', '..', '.bin');
  // NODE_TEST_CONTEXT would make the inner runner report to this one in its own protocol.
  const { NODE_TEST_CONTEXT: _context, ...env } = process.env;
  const pathKey = Object.keys(env).find((k) => k.toUpperCase() === 'PATH') ?? 'PATH';
  const out: NodeJS.ProcessEnv = { ...env, [pathKey]: `${bin}${delimiter}${env[pathKey] ?? ''}` };
  // Only a run meant to hang gets the short limit: on a loaded machine a healthy run can pass it.
  if (limited) out.TEST_RUN_LIMIT_MS = String(LIMIT_MS);
  else delete out.TEST_RUN_LIMIT_MS;
  return out;
}

type Run = { ended: boolean; code: number | null; output: string };

function runTestScript(fixture: string, limited: boolean): Promise<Run> {
  const script: string = JSON.parse(readFileSync(join(REPO, 'package.json'), 'utf8')).scripts.test;
  assert.ok(script.includes(TEST_GLOB), `the test script no longer names ${TEST_GLOB}: ${script}`);
  const command = script.replace(TEST_GLOB, `"${join(dir, fixture)}"`);
  const child = spawn(command, { cwd: REPO, shell: true, env: envForScript(limited), detached: DETACHED, stdio: ['ignore', 'pipe', 'pipe'] });
  let output = '';
  child.stdout.on('data', (b) => { output += b; });
  child.stderr.on('data', (b) => { output += b; });
  const release = () => {
    child.stdout.destroy();
    child.stderr.destroy();
    child.unref();
  };
  return new Promise((done) => {
    const timer = setTimeout(() => {
      killTree(child);
      release();
      done({ ended: false, code: null, output });
    }, DEADLINE_MS);
    // On exit, not close: a process left behind would hold the pipes open, and is what
    // `survivors` is there to report. The wait lets output already written arrive.
    child.on('exit', (code) => {
      clearTimeout(timer);
      const finish = () => { release(); done({ ended: true, code, output }); };
      const grace = setTimeout(finish, 2000);
      child.on('close', () => { clearTimeout(grace); finish(); });
    });
  });
}

// Process ids whose command line names the fixture directory, other than the query itself.
function survivors(): string[] {
  const marker = basename(dir);
  if (process.platform === 'win32') {
    const r = spawnSync('powershell', ['-NoProfile', '-Command',
      `Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*${marker}*' -and $_.ProcessId -ne $PID } | ForEach-Object { $_.ProcessId }`],
    { encoding: 'utf8' });
    return r.stdout.split(/\s+/).filter(Boolean);
  }
  const r = spawnSync('ps', ['-eo', 'pid=,args='], { encoding: 'utf8' });
  return r.stdout.split('\n').filter((l) => l.includes(marker)).map((l) => l.trim().split(/\s+/)[0]);
}

const opts = { timeout: DEADLINE_MS + 20_000 };

test('npm test ends a run that would never end by itself, and leaves no process behind', opts, async () => {
  const run = await runTestScript('never-ends.test.mjs', true);
  assert.ok(run.ended, `the run was still going ${DEADLINE_MS} ms in, with a ${LIMIT_MS} ms limit set`);
  assert.notEqual(run.code, 0);
  assert.match(run.output, new RegExp(`${LIMIT_MS} ms`));
  assert.deepEqual(survivors(), []);
});

test('npm test ends a run whose file process stopped partway through a report', opts, async () => {
  const run = await runTestScript('cut-short.test.mjs', true);
  assert.ok(run.ended, `the run was still going ${DEADLINE_MS} ms in, with a ${LIMIT_MS} ms limit set`);
  assert.deepEqual(survivors(), []);
});

test('npm test exits 0 when the tests pass and 1 when one fails', opts, async () => {
  assert.equal((await runTestScript('passes.test.mjs', false)).code, 0);
  const failed = await runTestScript('fails.test.mjs', false);
  assert.equal(failed.code, 1);
  assert.match(failed.output, /fixture failure/);
});
