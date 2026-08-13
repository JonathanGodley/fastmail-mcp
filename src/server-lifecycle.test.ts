// Process-lifecycle tests for the built server (dist/index.js).
//
// Why these exist (#80): an embedding caller spawns this server as a short-lived
// stdio child - spawn, initialize, a few tool calls, close stdin, expect the
// process to end. If it does not end, the caller silently accumulates orphaned
// node processes (most damaging on Windows, where an orphan outlives its parent
// with no session cleanup).
//
// How the exit actually happens: the SDK's StdioServerTransport attaches only
// 'data'/'error' listeners to process.stdin and never registers an EOF handler,
// so nothing calls transport.close() and the server's onclose never fires. The
// process ends instead because stdin reaching EOF releases the last handle
// holding the libuv event loop open, and node exits naturally with code 0.
//
// That is graceful - an in-flight tool call still gets to write its response
// before the loop drains - but it is *implicit*: any future long-lived handle
// (a setInterval, an unclosed socket or file watcher) would silently turn a
// clean exit into a hang. That is precisely why an explicit exit handler was
// NOT added: forcing process.exit() on stdin EOF would trade the graceful
// drain for a truncated response, and would also mask a leaked handle rather
// than reveal it. These tests are the guard instead - they fail the moment the
// implicit exit stops working.
//
// The tests need no credentials and make no network calls: authentication is
// resolved lazily on the first tool call, and no tool is called here.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, readdirSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const SERVER_ENTRY = join(SRC_DIR, '..', 'dist', 'index.js');

/**
 * Newest mtime across the compiled sources, or null if src/ cannot be read.
 *
 * Deliberately every non-test .ts in src/, not just index.ts: the entry point is
 * one module among many, so a change confined to any other source file leaves
 * index.ts untouched and a comparison against it alone would call a stale dist/
 * fresh.
 */
function newestSourceMtimeMs(): number | null {
  try {
    let newest = 0;
    for (const name of readdirSync(SRC_DIR)) {
      if (!name.endsWith('.ts') || name.endsWith('.test.ts')) continue;
      newest = Math.max(newest, statSync(join(SRC_DIR, name)).mtimeMs);
    }
    return newest === 0 ? null : newest;
  } catch {
    return null;
  }
}

// Generous relative to the measured behaviour on a warm developer machine
// (~300ms to the initialize response, ~10-40ms from stdin EOF to exit). The
// budgets are sized for a slow or loaded CI box, not tuned to the local numbers.
const READY_BUDGET_MS = 15_000;
const EXIT_BUDGET_MS = 10_000;

// `npm ci` / `npm install` run the `prepare` script, which builds dist/, so a
// freshly installed checkout always has the entry point. A missing dist/ means
// someone cleaned it by hand; skip loudly rather than fail a build-independent
// `npm test`.
const DIST_MISSING = !existsSync(SERVER_ENTRY);

/** Child env with every Fastmail credential stripped, so the test cannot touch a real account. */
function credentialFreeEnv(): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = { ...process.env };
  for (const key of Object.keys(env)) {
    if (/fastmail/i.test(key)) delete env[key];
  }
  return env;
}

interface RunOutcome {
  /** Milliseconds from stdin EOF to process exit, or null if it never exited. */
  exitAfterEofMs: number | null;
  exitCode: number | null;
  signal: NodeJS.Signals | null;
  timedOut: boolean;
  /** True once the initialize response was seen (only meaningful when handshake is requested). */
  handshakeCompleted: boolean;
}

/**
 * Spawn the built server, optionally complete the initialize handshake, then close
 * stdin and wait for the process to exit on its own. The child is killed only if it
 * overruns a budget, so a kill in the outcome is always a test failure, never cleanup.
 */
function runUntilExit(options: { handshake: boolean }): Promise<RunOutcome> {
  return new Promise<RunOutcome>((resolve) => {
    const child = spawn(process.execPath, [SERVER_ENTRY], {
      env: credentialFreeEnv(),
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    let stdoutBuf = '';
    let eofAt: bigint | null = null;
    let handshakeCompleted = false;
    let settled = false;
    let readyTimer: NodeJS.Timeout | undefined;
    let exitTimer: NodeJS.Timeout | undefined;

    const finish = (outcome: RunOutcome) => {
      if (settled) return;
      settled = true;
      clearTimeout(readyTimer);
      clearTimeout(exitTimer);
      resolve(outcome);
    };

    const closeStdin = () => {
      eofAt = process.hrtime.bigint();
      child.stdin.end();
      exitTimer = setTimeout(() => {
        finish({ exitAfterEofMs: null, exitCode: null, signal: null, timedOut: true, handshakeCompleted });
        child.kill();
      }, EXIT_BUDGET_MS);
    };

    // The server logs a startup line to stderr; drain it so a full pipe buffer can
    // never be the thing that keeps the child alive.
    child.stderr.resume();

    child.stdout.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => {
      stdoutBuf += chunk;
      let nl: number;
      while ((nl = stdoutBuf.indexOf('\n')) !== -1) {
        const line = stdoutBuf.slice(0, nl).trim();
        stdoutBuf = stdoutBuf.slice(nl + 1);
        if (!line) continue;
        let msg: { id?: unknown };
        try {
          msg = JSON.parse(line);
        } catch {
          continue;
        }
        if (msg.id === 1 && !handshakeCompleted) {
          handshakeCompleted = true;
          clearTimeout(readyTimer);
          closeStdin();
        }
      }
    });

    child.on('error', () => {
      finish({ exitAfterEofMs: null, exitCode: null, signal: null, timedOut: false, handshakeCompleted });
    });

    child.on('exit', (code, signal) => {
      finish({
        exitAfterEofMs: eofAt === null ? null : Number(process.hrtime.bigint() - eofAt) / 1e6,
        exitCode: code,
        signal,
        timedOut: false,
        handshakeCompleted,
      });
    });

    readyTimer = setTimeout(() => {
      finish({ exitAfterEofMs: null, exitCode: null, signal: null, timedOut: true, handshakeCompleted });
      child.kill();
    }, READY_BUDGET_MS);

    if (options.handshake) {
      child.stdin.write(
        JSON.stringify({
          jsonrpc: '2.0',
          id: 1,
          method: 'initialize',
          params: {
            protocolVersion: '2024-11-05',
            capabilities: {},
            clientInfo: { name: 'server-lifecycle-test', version: '1.0.0' },
          },
        }) + '\n',
      );
    } else {
      clearTimeout(readyTimer);
      closeStdin();
    }
  });
}

describe('spawned server process lifecycle', { skip: DIST_MISSING ? 'dist/index.js not built - run `npm run build` first' : false }, () => {
  it('exits on its own when stdin closes after the initialize handshake', async (t) => {
    // A stale dist/ would test yesterday's server, so say so rather than let a
    // pass be quietly meaningless.
    const newestSource = newestSourceMtimeMs();
    if (newestSource !== null && newestSource > statSync(SERVER_ENTRY).mtimeMs) {
      t.diagnostic('dist/ is older than src/ - run `npm run build` for a meaningful result');
    }

    const outcome = await runUntilExit({ handshake: true });

    assert.equal(outcome.handshakeCompleted, true, 'server never answered initialize');
    assert.equal(outcome.timedOut, false, `server did not exit within ${EXIT_BUDGET_MS}ms of stdin closing`);
    assert.equal(outcome.signal, null, 'server had to be killed instead of exiting on its own');
    assert.equal(outcome.exitCode, 0, 'server exited with a non-zero code');
  });

  it('exits on its own when stdin closes before any handshake', async () => {
    // A caller that spawns the server and then gives up (or times out) must not
    // leave a live child behind either.
    const outcome = await runUntilExit({ handshake: false });

    assert.equal(outcome.timedOut, false, `server did not exit within ${EXIT_BUDGET_MS}ms of stdin closing`);
    assert.equal(outcome.signal, null, 'server had to be killed instead of exiting on its own');
    assert.equal(outcome.exitCode, 0, 'server exited with a non-zero code');
  });
});
