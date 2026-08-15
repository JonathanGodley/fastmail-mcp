// Credential-egress tests that can only be made against the BUILT server
// (dist/index.js) running as a real external process.
//
// Two output channels carry credential risk, and neither is reachable from an
// in-process unit test:
//
//   1. Tool output. The CallTool catch in index.ts is the single choke point where
//      an error message becomes text the MCP caller reads. There is no exported
//      seam for it — the mapping lives inside the request handler — so the only way
//      to prove a given error path is redacted is to drive the real server over
//      JSON-RPC and read what comes back.
//   2. stderr. tsdav logs the HTTP Basic credential as bare base64 whenever DEBUG
//      is set. That write never passes through index.ts's redaction boundary, so
//      the only control is suppressing the logger — and the only honest proof is
//      loading the real built module and calling the real tsdav function.
//
// Neither group needs credentials or network: every tool call below is rejected by
// input validation before a request is issued, and the tsdav call is pure local
// string work. The fake values are synthetic and shaped like credentials on purpose
// — a placeholder that no pattern could match would prove nothing about redaction.

import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
// The reusable JSON-RPC client, rather than a hand-rolled one: it matches responses
// by JSON-RPC id and settles every failure path, which a fresh client re-bugs.
import { createClient } from '../scripts/mcp-harness.mjs';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = join(SRC_DIR, '..');
const SERVER_ENTRY = join(REPO_ROOT, 'dist', 'index.js');
const require = createRequire(join(REPO_ROOT, 'package.json'));

// A synthetic API credential. Deliberately matches NONE of the shape patterns
// (no "Bearer " prefix, no `fmu…` form): the only thing that can scrub it is
// value-based redaction of the literal registered at startup, so if it comes back
// clean, registration-based redaction is what did it.
const FAKE_API_VALUE = 'probe-value-not-a-real-credential';

// `npm test` runs tsx over src/ and never builds, so every assertion in this file
// is made against whatever dist/ happened to be lying around. A stale dist would
// pass these tests using the previous build's code - the exact false green they
// exist to prevent, and it would be loudest on the change most likely to break
// them (a newly edited source file). So refuse to run against a stale artifact
// rather than reporting a pass that means nothing.
function assertDistIsCurrent(): void {
  let built: number;
  try {
    built = statSync(SERVER_ENTRY).mtimeMs;
  } catch {
    assert.fail(`${SERVER_ENTRY} does not exist. Run "npm run build" before npm test.`);
  }
  const newest = readdirSync(SRC_DIR)
    .filter((f) => f.endsWith('.ts') && !f.endsWith('.test.ts'))
    .map((f) => ({ f, m: statSync(join(SRC_DIR, f)).mtimeMs }))
    .reduce((a, b) => (b.m > a.m ? b : a));
  assert.ok(
    built >= newest.m,
    `dist/index.js is older than src/${newest.f}. These tests assert against the built ` +
      `server, so a stale dist would report a pass for code that is not running. ` +
      `Run "npm run build".`,
  );
}

// ---------------------------------------------------------------------------
// 1. Error text reaching tool output
// ---------------------------------------------------------------------------

describe('every error path reaching tool output is redacted', () => {
  let client: ReturnType<typeof createClient>;
  let outsidePath: string;

  before(async () => {
    assertDistIsCurrent();

    // A real directory for the confinement roots, and a path OUTSIDE it whose
    // filename carries the credential — so the rejection message quotes the
    // credential back and there is something for redaction to do.
    const allowedDir = mkdtempSync(join(tmpdir(), 'fastmail-mcp-allowed-'));
    outsidePath = join(tmpdir(), `fastmail-mcp-outside-${FAKE_API_VALUE}.txt`);

    // Build the child env explicitly: strip every FASTMAIL_* name the server
    // consults so a developer's real credentials in the ambient environment can
    // never be the thing under test, then inject the synthetic one.
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;
    env.FASTMAIL_ATTACH_DIR = allowedDir;
    env.FASTMAIL_DOWNLOAD_DIR = allowedDir;

    client = createClient({ env });
    await client.init();
  });

  after(() => client?.close());

  // The credential is registered for value-based redaction inside initializeClient(),
  // which the request handler reaches only AFTER the unknown-parameter check. A test
  // whose very first call is rejected by that check would therefore see an
  // unregistered credential and pass for the wrong reason. This makes registration
  // explicit instead of relying on the order tests happen to run in.
  async function registerCredential(): Promise<void> {
    try {
      await client.call('download_attachment', { emailId: 'e', attachmentId: 'a', path: outsidePath });
    } catch {
      // Expected to reject; the point is that it got past the parameter check to
      // initializeClient() first.
    }
  }

  // Both assertions matter. "Credential absent" alone would pass if the message
  // never contained it (e.g. a different guard fired first and named no path), so
  // the presence of the redaction marker is what proves the message really did
  // carry the credential and really was scrubbed.
  function assertRedacted(message: string): void {
    assert.ok(!message.includes(FAKE_API_VALUE), `credential leaked into tool output: ${message}`);
    assert.ok(message.includes('[REDACTED]'), `expected a redaction marker in: ${message}`);
  }

  async function callAndCaptureError(tool: string, args: object): Promise<string> {
    try {
      const result = await client.call(tool, args);
      assert.fail(`${tool} unexpectedly succeeded: ${JSON.stringify(result).slice(0, 200)}`);
    } catch (e: any) {
      return String(e?.message ?? e);
    }
  }

  it("redacts download_attachment's own PathAccessError branch", async () => {
    await registerCredential();
    // This branch maps and throws locally, short-circuiting the top-level catch, so
    // it has to redact for itself.
    const message = await callAndCaptureError('download_attachment', {
      emailId: 'e1',
      attachmentId: 'a1',
      path: outsidePath,
    });
    assertRedacted(message);
  });

  it('redacts the top-level PathAccessError branch', async () => {
    await registerCredential();
    // create_draft has no local try/catch, so an attachment path rejection travels
    // to the CallTool catch and is mapped there.
    const message = await callAndCaptureError('create_draft', {
      subject: 'redaction probe',
      attachments: [{ path: outsidePath }],
    });
    assertRedacted(message);
  });

  it('redacts the top-level McpError rethrow', async () => {
    await registerCredential();
    // The unknown-parameter rejection echoes the offending KEY back to the caller,
    // so a credential-shaped key lands inside an McpError message that the catch
    // used to rethrow untouched.
    const message = await callAndCaptureError('list_emails', { [FAKE_API_VALUE]: 1 });
    assertRedacted(message);
    // Redacting in place rather than rebuilding the error: a rebuild would run the
    // already-prefixed message through the McpError constructor a second time.
    assert.equal(message.match(/MCP error /g)?.length, 1, `prefix duplicated in: ${message}`);
  });
});

// ---------------------------------------------------------------------------
// 2. Credential logging reaching stderr
// ---------------------------------------------------------------------------

describe('tsdav credential logging is suppressed', () => {
  before(() => assertDistIsCurrent());

  // Credential-shaped, and on a placeholder domain. The base64 of "user:pass" is
  // exactly what tsdav prints, so it is computed here rather than hardcoded.
  const PROBE_USER = 'probe-user@example.com';
  const PROBE_PASSPHRASE = 'probe-pass-value-0123456789';
  const LEAKED_FORM = Buffer.from(`${PROBE_USER}:${PROBE_PASSPHRASE}`).toString('base64');

  // Run tsdav's real credential-header function in a child with DEBUG turned fully
  // on, optionally loading the built server first so its suppression applies.
  // Nothing about the namespace is hardcoded here — the child exercises whatever
  // tsdav actually logs.
  function runTsdavCall({ loadServer }: { loadServer: boolean }) {
    const preamble = loadServer
      ? `await import(${JSON.stringify(pathToFileURL(SERVER_ENTRY).href)});\n`
      : '';
    const source =
      preamble +
      `const tsdav = await import('tsdav');\n` +
      `const headers = tsdav.getBasicAuthHeaders ?? tsdav.default?.getBasicAuthHeaders;\n` +
      `if (typeof headers !== 'function') { console.error('MISSING_TSDAV_EXPORT'); process.exit(2); }\n` +
      `headers({ username: ${JSON.stringify(PROBE_USER)}, password: ${JSON.stringify(PROBE_PASSPHRASE)} });\n` +
      `process.exit(0);\n`;
    return spawnSync(process.execPath, ['--input-type=module', '--eval', source], {
      cwd: REPO_ROOT,
      encoding: 'utf8',
      env: { ...process.env, DEBUG: '*' },
    });
  }

  it('leaks the credential without the suppression (control)', () => {
    // Without this control the suppression test would silently go vacuous the day
    // tsdav stops logging credentials, and would keep passing while a real
    // regression elsewhere went unnoticed.
    const child = runTsdavCall({ loadServer: false });
    assert.equal(child.status, 0, `probe failed: ${child.stderr}`);
    assert.ok(
      child.stderr.includes(LEAKED_FORM),
      'tsdav no longer logs the Basic credential under DEBUG — the suppression below is now proving nothing; re-derive it against the installed version.',
    );
  });

  it('does not leak the credential once the built server has loaded', () => {
    const child = runTsdavCall({ loadServer: true });
    assert.equal(child.status, 0, `probe failed: ${child.stderr}`);
    assert.ok(
      !child.stderr.includes(LEAKED_FORM),
      `tsdav logged the Basic credential to stderr despite the suppression: ${child.stderr}`,
    );
  });

  it('resolves the same debug instance as tsdav does', () => {
    // The suppression reaches only the `debug` module instance tsdav itself
    // resolves. If npm ever nests node_modules/tsdav/node_modules/debug (which a
    // version range diverging from tsdav's exact pin would cause), the control
    // still runs but applies to a different copy and silently stops working.
    const tsdavDir = dirname(require.resolve('tsdav/package.json'));
    assert.equal(
      require.resolve('debug', { paths: [tsdavDir] }),
      require.resolve('debug', { paths: [REPO_ROOT] }),
      'tsdav resolves a different copy of `debug` than the top level; the suppression no longer reaches it',
    );
  });

  it('suppresses every debug namespace the installed tsdav actually creates', async () => {
    // Read the namespaces out of the tsdav build Node loads, rather than asserting
    // against a hardcoded list: a tsdav upgrade that renames or adds a namespace
    // outside the skip glob has to turn this red, not quietly reopen the leak.
    const entry = require.resolve('tsdav');
    const source = readFileSync(entry, 'utf8');

    // Find whatever identifier tsdav binds the debug factory to, then collect every
    // string literal it is called with. Keyed off the import rather than off the
    // string "tsdav", so a namespace renamed to something else is still collected.
    const binding =
      source.match(/(?:var|const|let)\s+(\w+)\s*=\s*require\(['"]debug['"]\)/) ??
      source.match(/import\s+(\w+)\s+from\s*['"]debug['"]/);
    assert.ok(binding, `could not find tsdav's debug import in ${entry}`);

    const callRe = new RegExp(`\\b${binding[1]}\\(\\s*['"\`]([^'"\`]+)['"\`]\\s*\\)`, 'g');
    const namespaces = [...source.matchAll(callRe)].map((m) => m[1]);
    assert.ok(namespaces.length > 0, `found no debug namespaces in ${entry}`);

    // Ask the real matcher whether the skip pattern covers them, rather than
    // reimplementing the glob.
    const { default: createDebug } = await import('debug');
    // `disable()` returns the namespace string currently in force and clears it, so
    // it both captures the state to restore and puts the matcher in a known one. The
    // alternative read, `createDebug.namespaces`, is a real runtime property that
    // @types/debug does not declare, and reaching it would need a cast; this is the
    // declared API for the same job.
    const restore = createDebug.disable();
    try {
      createDebug.enable('*,-tsdav*');
      for (const ns of namespaces) {
        assert.equal(createDebug.enabled(ns), false, `namespace '${ns}' is not covered by the -tsdav* skip`);
      }
      // The skip must not be narrowed to '-tsdav:*': that form matches only
      // colon-prefixed children, leaving a bare or differently-suffixed logger live.
      createDebug.enable('*,-tsdav:*');
      assert.equal(createDebug.enabled('tsdav'), true, 'expected -tsdav:* to MISS a bare `tsdav` logger');
      assert.equal(createDebug.enabled('tsdavFoo'), true, 'expected -tsdav:* to MISS a `tsdavFoo` logger');
      createDebug.enable('*,-tsdav*');
      assert.equal(createDebug.enabled('tsdav'), false);
      assert.equal(createDebug.enabled('tsdavFoo'), false);
      // The operator's own DEBUG survives for every other package.
      assert.equal(createDebug.enabled('other:thing'), true);
    } finally {
      createDebug.enable(restore);
    }
  });
});
