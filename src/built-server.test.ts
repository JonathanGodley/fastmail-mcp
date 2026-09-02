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
import { spawn, spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
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
    // edit_draft has no local try/catch, so an attachment path rejection travels
    // to the CallTool catch and is mapped there. It is the vehicle because its
    // attachment upload runs before it touches the network: draft_email resolves the
    // sending identity first, so with an unusable token that call dies on the session
    // fetch and never reaches the path guard this test is about.
    const message = await callAndCaptureError('edit_draft', {
      emailId: 'e1',
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
// 1b. The capability flag as the real process parses it
// ---------------------------------------------------------------------------
//
// FASTMAIL_ALLOW_BLOB_ATTACH opens a send capability, so how its value is PARSED is a
// security property, not a formatting detail: under a truthy-string test the value "false"
// would enable the capability the operator wrote it to refuse. That parse runs once at
// module load in the real process, against the real environment — an in-process unit test
// cannot reach it, because the flag is resolved before any exported function is called. So
// the check spawns the built server and reads the clause it advertises in tools/list, which
// is derived from the same resolved value the handlers use.

describe('FASTMAIL_ALLOW_BLOB_ATTACH is parsed strictly', () => {
  before(() => assertDistIsCurrent());

  // The advertised attachments description of draft_email, from a server started with the
  // given flag value. Every FASTMAIL_* name is stripped from the child environment first, so
  // an ambient setting cannot be what the assertion sees.
  async function attachmentsClause(value: string | undefined): Promise<string> {
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;
    if (value !== undefined) env.FASTMAIL_ALLOW_BLOB_ATTACH = value;

    const client = createClient({ env });
    try {
      await client.init();
      const result: any = await client.list();
      const tool = result.tools.find((t: any) => t.name === 'draft_email');
      return String(tool.inputSchema.properties.attachments.description);
    } finally {
      client.close();
    }
  }

  it('treats "false" as disabled — the parse this strictness exists for', async () => {
    const clause = await attachmentsClause('false');
    assert.match(clause, /blobId and emailId\+attachmentId are disabled until FASTMAIL_ALLOW_BLOB_ATTACH=true/);
  });

  it('leaves the capability off when the variable is unset', async () => {
    const clause = await attachmentsClause(undefined);
    assert.match(clause, /disabled until FASTMAIL_ALLOW_BLOB_ATTACH=true/);
  });

  it('enables it on "true" and on "1"', async () => {
    for (const value of ['true', '1']) {
      const clause = await attachmentsClause(value);
      assert.match(clause, /blobId and emailId\+attachmentId are ENABLED/, `value ${value} did not enable it`);
    }
  });

  // The values an operator reaching for "on" actually types. All of them fail CLOSED today,
  // which is the right direction for a capability gate — a flag that half-works is worse
  // than one that plainly did not take. Pinned so a later "be more helpful about spellings"
  // change has to be a deliberate edit here, and so the same widening cannot arrive by
  // accident and drag "FALSE"/"False" in with it.
  it('leaves every other spelling disabled, including the ones that look like yes', async () => {
    for (const value of ['TRUE', 'True', 'yes', 'on', '0', '2', ' ', '']) {
      const clause = await attachmentsClause(value);
      assert.match(
        clause,
        /blobId and emailId\+attachmentId are disabled until FASTMAIL_ALLOW_BLOB_ATTACH=true/,
        `value ${JSON.stringify(value)} unexpectedly enabled the capability`,
      );
    }
  });
});

// ---------------------------------------------------------------------------
// 1b. The lenient array parameters advertise the string form they accept
// ---------------------------------------------------------------------------
//
// This server coerces a stringified array (coerceStringArray takes a JSON or comma-separated
// string), but a client validates the ADVERTISED schema first, so a parameter declared
// `type: 'array'` has its string form rejected before any handler runs — leniency that
// cannot be exercised. The two halves have to be checked against each other, and only the
// advertised schema says what a client will see, which is why this reads tools/list off the
// real process rather than asserting over the source.

describe('edit_draft advertises the stringified-array form its handler accepts', () => {
  before(() => assertDistIsCurrent());

  // The advertised inputSchema of one tool, from a server started with no FASTMAIL_* setting
  // beyond the token, so an ambient value cannot be what the assertion sees.
  async function toolSchema(name: string): Promise<any> {
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;

    const client = createClient({ env });
    try {
      await client.init();
      const result: any = await client.list();
      return result.tools.find((t: any) => t.name === name).inputSchema;
    } finally {
      client.close();
    }
  }

  // What a client can SEND is the fact under test, not how the schema spells it. A
  // `type: ['array', 'string']` union and a two-branch `oneOf` admit the same values, and
  // the declarations read here have been written both ways, so the assertions read the
  // types the declaration admits and the constraint on its array elements, whichever shape
  // carries them. Pinning one spelling would fail a rewrite into the other while the
  // client-visible behaviour was unchanged.
  function admittedTypes(declared: any): string[] {
    const out = new Set<string>();
    const add = (t: any) => {
      for (const one of Array.isArray(t) ? t : (t ? [t] : [])) out.add(one);
    };
    add(declared.type);
    for (const branch of declared.oneOf ?? declared.anyOf ?? []) add(branch.type);
    return [...out].sort();
  }

  // The element constraint sits in `items`, which applies to array instances only — so it
  // reads the same off a type union (property-level `items`) as off the array branch of a
  // `oneOf`.
  function arrayItems(declared: any): any {
    if (declared.items) return declared.items;
    const branch = (declared.oneOf ?? declared.anyOf ?? []).find((b: any) => b.type === 'array');
    return branch?.items;
  }

  // Both parameters run through coerceStringArray in edit-draft-handler.ts, and
  // edit-draft-handler.test.ts pins that the handler reads `clearFields: 'cc'` and
  // `removeAttachments: 'blob-9'`.
  for (const param of ['clearFields', 'removeAttachments']) {
    it(`declares ${param} as array-or-string`, async () => {
      const schema = await toolSchema('edit_draft');
      const declared = schema.properties[param];
      assert.deepEqual(
        admittedTypes(declared),
        ['array', 'string'],
        `${param} declares ${JSON.stringify(declared.type ?? declared.oneOf ?? declared.anyOf)} ` +
          'with no string alternative, so a validating client rejects the stringified form ' +
          'before coerceStringArray runs',
      );
    });
  }

  // The enum is what turns a bad field name into an error naming the nine valid ones. It
  // constrains the array elements and is the easiest thing to lose to a widening, which is
  // why it is asserted separately from the admitted types above.
  it('keeps the clearFields enum on the array elements', async () => {
    const schema = await toolSchema('edit_draft');
    assert.deepEqual(arrayItems(schema.properties.clearFields).enum, [
      'to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments', 'forwardedMessageId',
    ]);
  });
});

// ---------------------------------------------------------------------------
// 1c. FASTMAIL_TIMEZONE refuses to start the real process (#157 amendment)
// ---------------------------------------------------------------------------
//
// resolveConfiguredTimezone's throw path (coerce.test.ts) proves the pure logic; this proves
// the thing that actually matters operationally — that runServer() in index.ts really does
// call process.exit(1) with the message on stderr, rather than starting up and silently
// working out of the wrong zone. That wiring lives in runServer() itself and is not covered by
// any in-process unit test (importing src/index.ts would run runServer() for real and try to
// open a stdio transport), so a spawned real process is the only way to prove it.

describe('an unusable FASTMAIL_TIMEZONE refuses to start the built server', () => {
  before(() => assertDistIsCurrent());

  // No FASTMAIL_API_TOKEN in either env helper below: the timezone gate runs before any
  // credential is needed, so a startup failure here must not be mistaken for a missing-token
  // failure.
  function envWithTimezone(value: string | undefined): Record<string, string> {
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    if (value !== undefined) env.FASTMAIL_TIMEZONE = value;
    return env;
  }

  function spawnWithTimezone(value: string | undefined) {
    // A rejected config exits on its own almost immediately, so spawnSync (blocking until
    // exit) is safe here — unlike the "starts normally" cases below, which never exit by
    // themselves.
    return spawnSync(process.execPath, [SERVER_ENTRY], { encoding: 'utf8', env: envWithTimezone(value), timeout: 10_000 });
  }

  it('exits non-zero and names the value and the rule for a shorthand configured zone', () => {
    const result = spawnWithTimezone('EST');
    assert.equal(result.status, 1, `expected exit 1, got ${result.status}; stderr: ${result.stderr}`);
    assert.match(result.stderr, /FASTMAIL_TIMEZONE is set to "EST"/);
    assert.match(result.stderr, /"EST" resolves to a fixed-offset zone with no daylight saving, not US Eastern/);
    assert.doesNotMatch(result.stdout, /running on stdio/);
  });

  it('exits non-zero for an offset-shaped configured zone', () => {
    const result = spawnWithTimezone('+10:00');
    assert.equal(result.status, 1, `expected exit 1, got ${result.status}; stderr: ${result.stderr}`);
    assert.match(result.stderr, /FASTMAIL_TIMEZONE is set to "\+10:00"/);
  });

  it('exits non-zero for an unresolvable configured zone', () => {
    const result = spawnWithTimezone('Not/AZone');
    assert.equal(result.status, 1, `expected exit 1, got ${result.status}; stderr: ${result.stderr}`);
    assert.match(result.stderr, /FASTMAIL_TIMEZONE is set to "Not\/AZone"/);
    assert.match(result.stderr, /is not a time zone this server can resolve/);
  });

  // A server that starts successfully keeps running (it awaits requests on stdio), so it never
  // exits on its own — spawnSync would just hang until its timeout and report a kill, not a
  // clean exit. These spawn asynchronously instead, wait for the "running on stdio" line on
  // stderr (or a timeout), then kill the process explicitly — the same shape mcp-harness.mjs
  // uses for its own client, but scoped down to only what a stderr assertion needs.
  function spawnAndCaptureStartupLine(env: Record<string, string>): Promise<{ stderr: string; exited: boolean }> {
    return new Promise((resolve) => {
      const child = spawn(process.execPath, [SERVER_ENTRY], { env });
      let stderr = '';
      let settled = false;
      child.stderr.setEncoding('utf8');
      child.stderr.on('data', (chunk) => {
        stderr += chunk;
        if (!settled && /running on stdio/.test(stderr)) {
          settled = true;
          clearTimeout(timer);
          child.kill();
          resolve({ stderr, exited: false });
        }
      });
      child.on('exit', () => {
        if (!settled) {
          settled = true;
          clearTimeout(timer);
          resolve({ stderr, exited: true });
        }
      });
      const timer = setTimeout(() => {
        if (!settled) {
          settled = true;
          child.kill();
          resolve({ stderr, exited: false });
        }
      }, 5_000);
    });
  }

  it('starts normally with a valid configured zone', async () => {
    const { stderr, exited } = await spawnAndCaptureStartupLine(envWithTimezone('Australia/Sydney'));
    assert.match(stderr, /running on stdio/, `server did not report starting; exited early: ${exited}; stderr: ${stderr}`);
  });

  it('starts normally with FASTMAIL_TIMEZONE unset (host zone)', async () => {
    const { stderr, exited } = await spawnAndCaptureStartupLine(envWithTimezone(undefined));
    assert.match(stderr, /running on stdio/, `server did not report starting; exited early: ${exited}; stderr: ${stderr}`);
  });
});

// ---------------------------------------------------------------------------
// 2. Credential logging reaching stderr
// ---------------------------------------------------------------------------

describe('tsdav credential logging is suppressed', () => {
  before(() => assertDistIsCurrent());

  // Credential-shaped, and on a placeholder domain.
  const PROBE_USER = 'probe-user@example.com';
  const PROBE_PASSPHRASE = 'probe-pass-value-0123456789';
  // tsdav 2.1.x printed the whole Basic credential as bare base64. 2.3.1 replaced
  // that with the username alone. Both forms are asserted absent under the
  // suppression: the base64 is what a regression would look like, the username is
  // what the current version actually prints. Computed, never hardcoded.
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

  it('logs the account identity without the suppression (control)', () => {
    // Without this control the suppression test would silently go vacuous the day
    // tsdav stops logging anything from the auth helper, and would keep passing
    // while a real regression elsewhere went unnoticed. It deliberately asserts on
    // the account identity rather than the passphrase: which of the two tsdav
    // prints is tsdav's choice and has already changed once, so pinning the exact
    // secret form would make a routine bump red for no security reason.
    const child = runTsdavCall({ loadServer: false });
    assert.equal(child.status, 0, `probe failed: ${child.stderr}`);
    assert.ok(
      child.stderr.includes(PROBE_USER),
      'tsdav no longer logs the account identity under DEBUG — the suppression below is now proving nothing; re-derive it against the installed version.',
    );
  });

  it('does not leak the credential or the account identity once the built server has loaded', () => {
    const child = runTsdavCall({ loadServer: true });
    assert.equal(child.status, 0, `probe failed: ${child.stderr}`);
    assert.ok(
      !child.stderr.includes(LEAKED_FORM),
      `tsdav logged the Basic credential to stderr despite the suppression: ${child.stderr}`,
    );
    assert.ok(
      !child.stderr.includes(PROBE_USER),
      `tsdav logged the account identity to stderr despite the suppression: ${child.stderr}`,
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
    //
    // The ESM build specifically. `require.resolve('tsdav')` returns the CommonJS
    // entry, which this server never loads and which spells the namespaces
    // differently — reading it produced zero namespaces and a vacuous pass.
    const pkgPath = require.resolve('tsdav/package.json');
    const pkg = JSON.parse(readFileSync(pkgPath, 'utf8'));
    const esmRelative = pkg.exports?.['.']?.import ?? pkg.module;
    assert.ok(esmRelative, `tsdav declares no ESM entry point in ${pkgPath}`);
    const entry = resolve(dirname(pkgPath), esmRelative);
    const source = readFileSync(entry, 'utf8');

    // Find whatever identifier tsdav binds the debug factory to, then collect every
    // string literal it is called with. Keyed off the import rather than off the
    // string "tsdav", so a namespace renamed to something else is still collected.
    // Both binding forms stay recognised: the ESM build imports, and a future build
    // shape could go back to a require.
    const binding =
      source.match(/import\s+(\w+)\s+from\s*['"]debug['"]/) ??
      source.match(/(?:var|const|let)\s+(\w+)\s*=\s*require\(['"]debug['"]\)/);
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
