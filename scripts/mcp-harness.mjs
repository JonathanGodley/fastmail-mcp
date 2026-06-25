// Raw JSON-RPC (MCP-over-stdio) client for the fastmail-mcp server.
//
// Purpose: a reusable harness for the on-demand live-verification path, so it is
// not hand-rewritten (and re-bugged) each time. The server speaks newline-delimited
// JSON over stdio and logs only to stderr, so stdout is pure protocol.
//
// Before use:
//   1. `npm run build`  (the server runs from dist/index.js, not src/)
//   2. set FASTMAIL_API_TOKEN in the environment (and FASTMAIL_ATTACH_DIR for
//      attachment tests). This harness references those var *names* only and never
//      prints their values.
//
// Usage as a module:
//   import { createClient } from './scripts/mcp-harness.mjs';
//   const c = createClient({ env: process.env });
//   await c.init();
//   const result = await c.call('list_mailboxes', {});   // read-only smoke
//   c.close();
//
// Usage from the CLI (read-only smoke call):
//   FASTMAIL_API_TOKEN=… node scripts/mcp-harness.mjs list_mailboxes
//   (the token is a placeholder above; never paste a real value into a shared shell.)

import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const SERVER_ENTRY = join(HERE, '..', 'dist', 'index.js');

const PROTOCOL_VERSION = '2024-11-05';

/**
 * Spawn the built server and return a small JSON-RPC client.
 * @param {{ env?: NodeJS.ProcessEnv }} opts
 * @returns {{ init: () => Promise<object>, call: (name: string, args?: object) => Promise<object>, close: () => void }}
 */
export function createClient({ env } = {}) {
  // Inherit the caller's env (so FASTMAIL_API_TOKEN etc. flow through). We never
  // read individual secrets here and never print env on any path.
  const child = spawn('node', [SERVER_ENTRY], {
    env: env ?? process.env,
    stdio: ['pipe', 'pipe', 'inherit'], // stderr inherited: server logs pass through, untouched
  });

  let nextId = 1;
  const pending = new Map(); // id -> { resolve, reject }
  let stdoutBuf = '';

  // Newline-delimited JSON framing. A single 'data' chunk can carry multiple
  // objects, and one object can span two chunks — so split on every '\n' and
  // retain the trailing partial line for the next chunk.
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (chunk) => {
    stdoutBuf += chunk;
    let nl;
    while ((nl = stdoutBuf.indexOf('\n')) !== -1) {
      const line = stdoutBuf.slice(0, nl).trim();
      stdoutBuf = stdoutBuf.slice(nl + 1);
      if (!line) continue;
      let msg;
      try {
        msg = JSON.parse(line);
      } catch {
        // Not a JSON-RPC line (defensive — the server should only emit protocol
        // on stdout). Ignore rather than crash the matcher.
        continue;
      }
      // Match strictly by JSON-RPC id — never by scraping substrings, which was
      // the historical bug (a greedy pattern over-captured the id token).
      if (msg.id !== undefined && msg.id !== null && pending.has(msg.id)) {
        const { resolve, reject } = pending.get(msg.id);
        pending.delete(msg.id);
        if (msg.error) reject(new Error(`JSON-RPC error: ${JSON.stringify(msg.error)}`));
        else resolve(msg.result);
      }
      // Notifications (no id) and unmatched ids are ignored.
    }
  });

  // Settle every in-flight request with the same error — used by all three
  // failure paths so none of them can leave a promise hanging.
  const failAll = (err) => {
    for (const { reject } of pending.values()) reject(err);
    pending.clear();
  };

  child.on('error', (err) => {
    // Spawn failure (e.g. `node` not on PATH) or a transport-level error. Without
    // a listener, Node throws on the unhandled 'error' event and init()/call()
    // hang forever — so reject every pending request loudly instead.
    failAll(err);
  });

  child.on('exit', (code, signal) => {
    const how = signal ? `signal ${signal}` : `code ${code}`;
    failAll(new Error(`server exited (${how}) with ${pending.size} request(s) pending`));
  });

  // A write that races the pipe closing can emit an async 'error' (EPIPE) on
  // stdin; with no listener that becomes an uncaughtException and kills the
  // harness. Swallow it — the 'exit' handler settles anything still pending.
  child.stdin.on('error', () => {});

  // No per-request timeout by design: a server that is alive but never replies
  // would hang here, but this is a manual one-shot harness (Ctrl-C it), and a
  // blanket timeout would wrongly abort legitimately slow calls (a large sync, a
  // big attachment upload). A caller that wants one can race send() against their
  // own timer. The settled failure paths above cover the cases that actually
  // recur: the process dying or the pipe breaking.
  function send(method, params) {
    // The child may already be gone (exited or killed). Writing to a dead pipe
    // would otherwise leave the request to hang forever, since the 'exit' handler
    // has already drained `pending`. Reject up front instead.
    if (child.exitCode !== null || child.killed) {
      return Promise.reject(new Error('cannot send: server process is not running'));
    }
    const id = nextId++;
    const req = { jsonrpc: '2.0', id, method, params };
    return new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject });
      // Do NOT log `req` — a tool-call request body can contain account data or,
      // for some tools, sensitive arguments.
      try {
        child.stdin.write(JSON.stringify(req) + '\n');
      } catch (err) {
        pending.delete(id);
        reject(err);
      }
    });
  }

  function notify(method, params) {
    if (child.exitCode !== null || child.killed) return;
    try {
      child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params }) + '\n');
    } catch {
      /* child gone; nothing is waiting on a notification */
    }
  }

  return {
    async init() {
      const result = await send('initialize', {
        protocolVersion: PROTOCOL_VERSION,
        capabilities: {},
        clientInfo: { name: 'mcp-harness', version: '1.0.0' },
      });
      notify('notifications/initialized', {});
      return result;
    },
    // NOTE: call() is generic and CAN mutate the account (send_email, delete_email,
    // bulk_* …). The copy-paste default below is the read-only one on purpose —
    // pick the tool name deliberately.
    call(name, args = {}) {
      return send('tools/call', { name, arguments: args });
    },
    close() {
      try {
        child.stdin.end();
      } catch {
        /* already closed */
      }
      child.kill();
    },
  };
}

// CLI entry: `node scripts/mcp-harness.mjs <tool> [jsonArgs]`. Defaults to the
// read-only list_mailboxes smoke call.
const INVOKED_DIRECTLY = process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1];
if (INVOKED_DIRECTLY) {
  const [, , toolName = 'list_mailboxes', rawArgs] = process.argv;
  if (toolName === '--help' || toolName === '-h') {
    // Intentionally prints only the var NAME, never a value.
    console.log('Usage: FASTMAIL_API_TOKEN=… node scripts/mcp-harness.mjs <tool> [jsonArgs]');
    console.log('Example (read-only): FASTMAIL_API_TOKEN=… node scripts/mcp-harness.mjs list_mailboxes');
    console.log('Build first: npm run build');
    process.exit(0);
  }
  let args;
  try {
    args = rawArgs ? JSON.parse(rawArgs) : {};
  } catch {
    // The bad value is the operator's own CLI arg — safe to surface; print only a
    // generic message and exit before spawning a server.
    console.error('harness error: invalid JSON args');
    process.exit(1);
  }
  const client = createClient({ env: process.env });
  try {
    await client.init();
    const result = await client.call(toolName, args);
    console.log(JSON.stringify(result, null, 2));
  } catch (err) {
    // Print only the message — never the token, env, or request body — and exit
    // non-zero so a failed smoke call is scriptable.
    console.error(`harness error: ${err.message}`);
    process.exitCode = 1;
  } finally {
    client.close();
  }
}
