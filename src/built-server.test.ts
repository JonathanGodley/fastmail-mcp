// Credential-egress tests that can only be made against the BUILT server
// (dist/index.js) running as a real external process.
//
// Two output channels carry credential risk, and neither is reachable from an
// in-process unit test:
//
//   1. Tool output. The CallTool catch in index.ts is the single choke point where
//      an error message becomes text the MCP caller reads, and it has no exported
//      seam, so the only proof a given error path is redacted is to drive the real
//      server over JSON-RPC.
//   2. stderr. tsdav logs the HTTP Basic credential whenever DEBUG is set. That
//      write never passes through index.ts's redaction boundary, so the only control
//      is suppressing the logger, proved by loading the real built module and
//      calling the real tsdav function.
//
// Neither group needs credentials or network: every tool call below is rejected by
// input validation before a request is issued, and the tsdav call is pure local
// string work. The fake values are shaped like credentials on purpose: a
// placeholder that no pattern could match would prove nothing about redaction.

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

// `npm test`'s `pretest` builds first, but `tsx --test src/built-server.test.ts` run
// directly skips it, and a stale dist/ would then pass these tests using the previous
// build's code. So refuse to run against a stale artifact.
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
    // so a credential-shaped key lands inside an McpError message.
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
// FASTMAIL_ALLOW_BLOB_ATTACH opens a send capability, so its parse is a security property:
// under a truthy-string test, "false" would enable it. The flag is resolved at module load,
// before any exported function is called, so the check spawns the built server and reads
// the clause it advertises in tools/list, derived from the same value the handlers use.

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

  // The values an operator reaching for "on" actually types. They fail CLOSED on purpose, so
  // accepting more spellings has to be a deliberate edit here, and cannot arrive by accident
  // and drag "FALSE"/"False" in with it.
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
// This server coerces a stringified array, but a client validates the ADVERTISED schema
// first, so a parameter declared `type: 'array'` has its string form rejected before any
// handler runs. Only the advertised schema says what a client will see, so this reads
// tools/list off the real process rather than the source.

// The element constraint sits in `items`, which applies to array instances only, so it
// reads the same off a type union (property-level `items`) as off the array branch of a
// `oneOf`.
function arrayItems(declared: any): any {
  if (declared.items) return declared.items;
  const branch = (declared.oneOf ?? declared.anyOf ?? []).find((b: any) => b.type === 'array');
  return branch?.items;
}

// What a client can SEND is the fact under test, not how the schema spells it. A
// `type: ['array', 'string']` union and a two-branch `oneOf` admit the same values, so a
// rewrite of the union into `oneOf` must not fail this guard for a schema that behaves the
// same.
function admittedTypes(declared: any): string[] {
  const out = new Set<string>();
  const add = (t: any) => {
    for (const one of Array.isArray(t) ? t : (t ? [t] : [])) out.add(one);
  };
  add(declared.type);
  for (const branch of declared.oneOf ?? declared.anyOf ?? []) add(branch.type);
  return [...out].sort();
}

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

  // The enum is what turns a bad field name into an error naming the nine valid ones. It
  // constrains the array elements and is the easiest thing to lose to a widening. Not covered
  // by the general drift guard below (#98), which pins the string-form sentence and the
  // coercer table but says nothing about an enum's contents.
  it('keeps the clearFields enum on the array elements', async () => {
    const schema = await toolSchema('edit_draft');
    assert.deepEqual(arrayItems(schema.properties.clearFields).enum, [
      'to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments', 'forwardedMessageId',
    ]);
  });
});

// ---------------------------------------------------------------------------
// 1b-ter. Array-side schema drift guard (#98)
// ---------------------------------------------------------------------------
//
// The array-side counterpart of the `lenient-boolean convention` in src/tool-schema.test.ts:
// a narrow `type: 'array'` makes a coercer unreachable through a validating client for the
// same reason a narrow `type: 'boolean'` does. Unlike the boolean guard, this reads tools/list
// off the built server rather than scanning source, because the ADVERTISED schema is what a
// client actually sees.
//
// Four things, each lost independently by a different mistake:
//
//   1. Every top-level parameter admitting `array` also admits `string`.
//   2. No top-level parameter carries `items` with no declared `type` at all: that shape
//      admits no `array`, so it slips past assertion 1 unseen.
//   3. Every array-admitting parameter keeps `items` and states which string forms it accepts,
//      in one of the sentences below, matched by SENTENCE TEXT rather than parameter name, so
//      `participants`' own wording is accepted on the same footing as the shared constants.
//   4. Every array-admitting parameter is named in the table below against the coercer that
//      reads it, checked in BOTH directions. Coercion is wired by hand per call site
//      (`assertKnownParams` in src/coerce.ts is key-strictness only), so a widened type with
//      no coercer is a silent fail-OPEN: a handler doing `for (const id of "abc")` iterates
//      characters. A name-pattern scan of the handlers cannot replace the table:
//      `coerceRecipients(a)` names none of to/cc/bcc/replyTo, and `fields` is read by
//      `parseEmailFields`.
//
// Scoped to TOP-LEVEL tool parameters only. Nothing below the top level declared an array type
// when this was written, and no test enforces that bound: a nested array parameter would need
// this walk extended to reach it.

// The accepted string-form sentences, matched by substring. LENIENT_LIST_DESC and
// LENIENT_OBJECT_LIST_DESC (src/index.ts) are copied verbatim because index.ts has no
// exported seam for them, and copying pins the exact prose a widening must not drop.
const ARRAY_STRING_FORM_SENTENCES = [
  'Accepts an array, or a single value, comma-separated string or JSON-encoded array as one string.',
  'Accepts an array, or a JSON-encoded array as one string; a comma-joined string is NOT accepted.',
  'The whole array may also be sent as a JSON string',
];

// Every array-admitting top-level parameter, named against the coercer that reads its string
// form, derived by reading each call site in index.ts and the handler modules.
const ARRAY_PARAM_COERCERS: Record<string, string> = {
  'list_emails.fields': 'index.ts parseEmailFields() -> coerceStringArray',
  'get_email.fields': 'index.ts parseEmailFields() -> coerceStringArray',
  'get_thread.fields': 'thread-handler.ts parseEmailFields() -> coerceStringArray',
  'search_emails.fields': 'index.ts parseEmailFields() -> coerceStringArray',
  'search_emails.requiredMailboxes': 'index.ts coerceStringArrayStrict',
  'search_emails.excludeMailboxes': 'index.ts coerceStringArrayStrict',
  'draft_email.to': 'draft-email-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'draft_email.cc': 'draft-email-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'draft_email.bcc': 'draft-email-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'draft_email.replyTo': 'draft-email-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'draft_email.inReplyTo': 'draft-email-handler.ts coerceStringArray',
  'draft_email.references': 'draft-email-handler.ts coerceStringArray',
  'draft_email.attachments': 'draft-email-handler.ts coerceAttachments',
  'edit_draft.to': 'edit-draft-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'edit_draft.cc': 'edit-draft-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'edit_draft.bcc': 'edit-draft-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'edit_draft.replyTo': 'edit-draft-handler.ts coerceRecipients() -> coerceStringArrayStrict',
  'edit_draft.attachments': 'edit-draft-handler.ts coerceAttachments',
  'edit_draft.removeAttachments': 'edit-draft-handler.ts coerceStringArray',
  'edit_draft.clearFields': 'edit-draft-handler.ts coerceStringArray',
  'create_contact.emails': 'contacts-handler.ts coerceContactEmails',
  'create_contact.phones': 'contacts-handler.ts coerceContactPhones',
  'create_contact.addresses': 'contacts-handler.ts coerceContactAddresses',
  'update_contact.emails': 'contacts-handler.ts coerceContactEmails',
  'update_contact.phones': 'contacts-handler.ts coerceContactPhones',
  'update_contact.addresses': 'contacts-handler.ts coerceContactAddresses',
  'update_contact.clearFields': 'contacts-handler.ts coerceStringArray',
  'create_calendar_event.participants': 'index.ts coerceParticipants',
  'update_calendar_event.participants': 'index.ts coerceParticipants',
  'update_calendar_event.clearFields': 'index.ts coerceStringArray',
  'archive_email.emailIds': 'index.ts coerceStringArrayStrict',
  'add_labels.mailboxes': 'index.ts coerceStringArray',
  'remove_labels.mailboxes': 'index.ts coerceStringArray',
  'bulk_mark_read.emailIds': 'index.ts coerceStringArray',
  'bulk_pin.emailIds': 'index.ts coerceStringArray',
  'bulk_move.emailIds': 'index.ts coerceStringArray',
  'bulk_delete.emailIds': 'index.ts coerceStringArray',
  'bulk_add_labels.emailIds': 'index.ts coerceStringArray',
  'bulk_add_labels.mailboxes': 'index.ts coerceStringArray',
  'bulk_remove_labels.emailIds': 'index.ts coerceStringArray',
  'bulk_remove_labels.mailboxes': 'index.ts coerceStringArray',
};

describe('array-side schema drift guard (#98)', () => {
  let tools: any[];
  let bootError: unknown;

  // Spawned ONCE and shared by every test below, with the same env scrub as toolSchema(), so
  // an ambient setting cannot be what the assertions see.
  //
  // The spawn/init/list has its own try/catch: a throwing before() makes node:test report the
  // whole suite CANCELLED, not failed (fail 0, a green-looking exit). Storing the error and
  // asserting it in the first test turns a boot or tools/list failure into a failing
  // assertion instead.
  before(async () => {
    assertDistIsCurrent();
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;

    try {
      const client = createClient({ env });
      try {
        await client.init();
        const result: any = await client.list();
        tools = result.tools;
      } finally {
        client.close();
      }
    } catch (err) {
      bootError = err;
      tools = [];
    }
  });

  function topLevelParams(): { key: string; declared: any }[] {
    const out: { key: string; declared: any }[] = [];
    for (const tool of tools) {
      for (const [param, declared] of Object.entries<any>(tool.inputSchema.properties ?? {})) {
        out.push({ key: `${tool.name}.${param}`, declared });
      }
    }
    return out;
  }

  it('sees all 41 tools on the wire', () => {
    // Checked first, so a boot/list failure caught in before() reads as this test's failure
    // rather than every test below blaming an empty tool set on the wrong thing.
    assert.equal(
      bootError,
      undefined,
      `the built server failed to boot or answer tools/list, so every assertion in this ` +
        `describe block saw an empty tool set rather than a real check: ${bootError}`,
    );

    // A floor on the TOOL count, separate from the array-admitting floor below: a tool
    // removed outright shrinks `tools` itself, which every assertion below would absorb
    // silently. Pinned to the exact count so a single removal trips it. `>=` cannot see an
    // added tool, so raise this floor in the same change that adds one, or it loses
    // single-removal sensitivity.
    assert.ok(
      tools.length >= 41,
      `found only ${tools.length} tools (expected 41); either tools/list stopped returning the ` +
        'real surface, or a tool was removed with this floor not yet lowered to match',
    );
  });

  it('never advertises array without also admitting string', () => {
    const params = topLevelParams();
    // A floor on the ARRAY-ADMITTING set, not the total walked: re-spelling every array
    // parameter as a narrow `type: 'string'` still walks every parameter, so a floor on the
    // walk would stay green through exactly the regression this guard catches. Pinned to the
    // exact count, like the tool floor above.
    const arrayAdmitting = params.filter((p) => admittedTypes(p.declared).includes('array'));
    assert.ok(
      arrayAdmitting.length >= 41,
      `found only ${arrayAdmitting.length} array-admitting top-level parameters (expected 41); ` +
        'either admittedTypes() has stopped matching the live schema, or an array-admitting ' +
        'parameter was removed or narrowed with this floor not yet lowered to match',
    );

    const offenders = arrayAdmitting
      .filter((p) => !admittedTypes(p.declared).includes('string'))
      .map((p) => `${p.key} (declares ${JSON.stringify(p.declared.type ?? p.declared.oneOf ?? p.declared.anyOf)})`);
    assert.deepEqual(
      offenders,
      [],
      'these parameters admit array with no string alternative, so a validating client rejects ' +
        'the stringified form before any coercer runs. Cure, both halves: widen the type to ' +
        "admit 'string' too (LENIENT_LIST_DESC / LENIENT_OBJECT_LIST_DESC in src/index.ts are " +
        'the usual route, not a requirement — participants deliberately writes its own ' +
        'sentence) AND wire a coercer that actually reads the string form, or the widening ' +
        `ships with nothing behind it: ${offenders.join(', ')}`,
    );

    // No "admits string" filter needed: the assertion above has already thrown on any
    // array-admitting parameter that does not.
    const extraTypeOffenders = arrayAdmitting
      .filter((p) => admittedTypes(p.declared).some((t) => t !== 'array' && t !== 'string'))
      .map((p) => {
        const extra = admittedTypes(p.declared).filter((t) => t !== 'array' && t !== 'string');
        return `${p.key} (also admits ${extra.join(', ')})`;
      });
    assert.deepEqual(
      extraTypeOffenders,
      [],
      'these parameters admit array, string, AND a further type. The lenient-list convention ' +
        "is exactly the pair ['array', 'string'] — a client and a reader both learn the whole " +
        'admitted set from that pair alone, with no third type to account for. Whoever widened ' +
        'this needs to say why a third type belongs here, and either narrow back to the pair or ' +
        "document its own coercion path the same way the pair's two halves are documented: " +
        `${extraTypeOffenders.join(', ')}`,
    );
  });

  it('never advertises items with no declared type at all', () => {
    const offenders = topLevelParams()
      .filter(({ declared }) => declared.items !== undefined && admittedTypes(declared).length === 0)
      .map((p) => p.key);
    assert.deepEqual(
      offenders,
      [],
      "these parameters carry 'items' with no 'type' at all — array-shaped to every reader and " +
        "to the schema's own intent, but admitting no 'array' means the previous assertion " +
        "skips them silently. Declare a type (['array', 'string'] if a coercer reads the " +
        `string form): ${offenders.join(', ')}`,
    );
  });

  it('keeps items and a documented string-form sentence on every array-admitting parameter', () => {
    const offenders: string[] = [];
    for (const { key, declared } of topLevelParams()) {
      if (!admittedTypes(declared).includes('array')) continue;
      const missing: string[] = [];
      if (!arrayItems(declared)) missing.push('items');
      const description = String(declared.description ?? '');
      if (!ARRAY_STRING_FORM_SENTENCES.some((s) => description.includes(s))) missing.push('string-form sentence');
      if (missing.length > 0) offenders.push(`${key} (missing ${missing.join(' and ')})`);
    }
    assert.deepEqual(
      offenders,
      [],
      'a widening can drop items (degrading a named rejection into a generic one, per ' +
        "index.ts's clearFields comment) or the string-form sentence (leaving the type union " +
        'with no clue which strings actually work) without tripping the previous two ' +
        `assertions: ${offenders.join(', ')}`,
    );
  });

  it('names a coercer for every array-admitting parameter, matched to the wire in both directions', () => {
    const wireKeys = new Set(
      topLevelParams()
        .filter((p) => admittedTypes(p.declared).includes('array'))
        .map((p) => p.key),
    );
    const tableKeys = new Set(Object.keys(ARRAY_PARAM_COERCERS));

    const untabled = [...wireKeys].filter((k) => !tableKeys.has(k)).sort();
    assert.deepEqual(
      untabled,
      [],
      'these array-admitting parameters are not named in ARRAY_PARAM_COERCERS, so nothing ' +
        'states which coercer is supposed to read their string form — add a row naming it ' +
        `(and confirm the coercer actually runs): ${untabled.join(', ')}`,
    );

    const stale = [...tableKeys].filter((k) => !wireKeys.has(k)).sort();
    assert.deepEqual(
      stale,
      [],
      'ARRAY_PARAM_COERCERS names a coercer for a parameter no longer on the wire (renamed, ' +
        `narrowed back to a plain array, or removed) — delete its row: ${stale.join(', ')}`,
    );
  });
});

// ---------------------------------------------------------------------------
// 1b-bis. The calendar tools advertise `transparency` (#194)
// ---------------------------------------------------------------------------
//
// A parameter the handler threads but the schema never declares is rejected by a validating
// client before any handler runs, and the caldav-client tests, which call the client
// directly, cannot tell that apart from a wired-up parameter. This is the one assertion that
// spans the `TOOLS` literal, the build, and the real process's tools/list response.

describe('the calendar write tools advertise transparency with its closed value set', () => {
  before(() => assertDistIsCurrent());

  // Both tools off ONE listing: the fact under test is that each declares the parameter, and
  // spawning the server twice to learn it would double the cost for nothing.
  async function listedTools(): Promise<any[]> {
    const env: Record<string, string> = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
    }
    env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;

    const client = createClient({ env });
    try {
      await client.init();
      const result: any = await client.list();
      return result.tools;
    } finally {
      client.close();
    }
  }

  it('declares transparency as a string enumerating exactly busy and free, on both tools', async () => {
    const tools = await listedTools();
    // Enumerated, not sampled: the claim is about the pair of tools that write an event, and
    // a parameter on only one of them is the failure this exists to catch.
    for (const name of ['create_calendar_event', 'update_calendar_event']) {
      const tool = tools.find((t: any) => t.name === name);
      assert.ok(tool, `${name} is not in tools/list`);
      const declared = tool.inputSchema.properties.transparency;
      assert.ok(declared, `${name} does not advertise transparency, so a validating client cannot send it`);
      assert.equal(declared.type, 'string', name);
      // The caller vocabulary, and nothing else. `OPAQUE`/`TRANSPARENT` are the iCalendar
      // spellings and stay inside caldav-client.ts; finding either here would mean the storage
      // format had reached the tool surface.
      assert.deepEqual(declared.enum, ['busy', 'free'], name);
    }
  });

  it('lists transparency among the clearable fields on update_calendar_event', async () => {
    // The other half of the parameter: `clearFields: ["transparency"]` removes the property.
    // The enum is what turns a bad field name into an error naming the valid ones, so a
    // missing entry here refuses a documented call rather than silently ignoring it.
    const tools = await listedTools();
    const update = tools.find((t: any) => t.name === 'update_calendar_event');
    assert.deepEqual(arrayItems(update.inputSchema.properties.clearFields).enum, [
      'description', 'location', 'transparency',
    ]);
  });
});

// ---------------------------------------------------------------------------
// 1c. FASTMAIL_TIMEZONE refuses to start the real process (#157 amendment)
// ---------------------------------------------------------------------------
//
// coerce.test.ts covers resolveConfiguredTimezone's throw; this proves runServer() really
// calls process.exit(1) with the message on stderr rather than starting in the wrong zone.
// Importing src/index.ts would run runServer() for real and open a stdio transport, so only
// a spawned process can prove that wiring.

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

  // A server that starts successfully never exits on its own, so spawnSync would hang until
  // its timeout. These spawn asynchronously, wait for the "running on stdio" line on stderr
  // (or a timeout), then kill the process.
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
    // Without this control the suppression test would go vacuous the day tsdav stops
    // logging from the auth helper. It asserts on the account identity rather than the
    // passphrase: which of the two tsdav prints is tsdav's choice and has changed
    // before, so pinning the secret form would make a routine bump red for no reason.
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
    // The require form stays for a future build shape that goes back to require.
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
    // it both captures the state to restore and puts the matcher in a known one.
    // `createDebug.namespaces` would need a cast: @types/debug does not declare it.
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
