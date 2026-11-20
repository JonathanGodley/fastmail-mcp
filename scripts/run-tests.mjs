#!/usr/bin/env node
// `npm test`: runs `tsx --test` over the files given and kills the whole run, exiting 1, if it
// has not ended within TEST_RUN_LIMIT_MS (default 600000).
//
//   node scripts/run-tests.mjs src/*.test.ts
//
// The limit has to be enforced from a separate process. Node's test runner (25.2.1, at least)
// spins forever in its drain loop when a file's process stops partway through a report
// message, holding a core and growing without bound; that loop is synchronous, so no timer
// inside the runner, `--test-timeout` included, ever fires. The suite takes under a minute, so
// ten minutes ends only a run that is stuck.
//
// Off Windows, a SIGKILL to this process leaves the run's process group running: look for it
// before treating the run as gone.

import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { DETACHED, killTree } from './kill-tree.mjs';

const DEFAULT_LIMIT_MS = 600_000;
const limitMs = Number(process.env.TEST_RUN_LIMIT_MS) || DEFAULT_LIMIT_MS;

const tsxCli = createRequire(import.meta.url).resolve('tsx/cli');
// NODE_TEST_CONTEXT, inherited when this runs under a test, would make the run report to that
// parent runner in its own protocol instead of to this terminal.
const { NODE_TEST_CONTEXT: _context, ...env } = process.env;
const child = spawn(process.execPath, [tsxCli, '--test', ...process.argv.slice(2)], { stdio: 'inherit', env, detached: DETACHED });

function stop(message) {
  clearTimeout(timer);
  const killError = killTree(child);
  console.error(killError ? `${message}; ${killError}` : message);
  process.exit(1);
}

const timer = setTimeout(() => {
  stop(`test run passed its ${limitMs} ms limit (TEST_RUN_LIMIT_MS) and was killed with its processes`);
}, limitMs);

// A detached run is outside the terminal's process group, so Ctrl-C reaches only this process.
for (const signal of ['SIGINT', 'SIGTERM', 'SIGHUP']) {
  process.on(signal, () => stop(`test run stopped by ${signal}`));
}

child.on('error', (e) => stop(`test run could not start: ${e.message}`));
child.on('close', (code, signal) => {
  clearTimeout(timer);
  process.exit(code ?? (signal ? 1 : 0));
});
