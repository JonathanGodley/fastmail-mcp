import { describe, it, before, after, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

// load-env-file.ts loads the real home's file as a side effect of being imported, so the
// home is pointed at a temp dir first and the module is imported dynamically after. That
// home holds a file setting IMPORT_SENTINEL, which pins the load done at import.
const SANDBOX = mkdtempSync(join(tmpdir(), 'fastmail-mcp-env-file-'));
const IMPORT_SENTINEL = 'FASTMAIL_MCP_ENV_FILE_IMPORT_SENTINEL';
delete process.env[IMPORT_SENTINEL];
mkdirSync(join(SANDBOX, '.fastmail-mcp'));
writeFileSync(join(SANDBOX, '.fastmail-mcp', '.env'), `${IMPORT_SENTINEL}=loaded-at-import\n`);
process.env.HOME = SANDBOX;
process.env.USERPROFILE = SANDBOX;
const { loadHomeEnvFile, envFilePath } = await import('./load-env-file.js');
const sentinelAfterImport = process.env[IMPORT_SENTINEL];
delete process.env[IMPORT_SENTINEL];
after(() => rmSync(SANDBOX, { recursive: true, force: true }));

it('loads the home .env when imported', () => {
  assert.equal(sentinelAfterImport, 'loaded-at-import');
});

const VAR = 'FASTMAIL_MCP_ENV_FILE_TEST_VAR';

// A fresh home. `null` leaves out the .fastmail-mcp directory; a string is written as the file.
function homeWith(content: string | null): string {
  const home = mkdtempSync(join(SANDBOX, 'home-'));
  if (content !== null) {
    mkdirSync(join(home, '.fastmail-mcp'));
    writeFileSync(envFilePath(home), content);
  }
  return home;
}

// chmod cannot remove read access on Windows, so there an ACL entry denies it instead.
function denyRead(file: string): void {
  if (process.platform === 'win32') icacls(file, '/deny', `${process.env.USERNAME}:(R)`);
  else chmodSync(file, 0o000);
}

function allowRead(file: string): void {
  if (process.platform === 'win32') icacls(file, '/remove:d', `${process.env.USERNAME}`);
  else chmodSync(file, 0o600);
}

function icacls(...args: string[]): void {
  const result = spawnSync('icacls', args, { encoding: 'utf8' });
  assert.equal(result.status, 0, `icacls ${args.join(' ')} failed: ${result.stdout}${result.stderr}`);
}

describe('loadHomeEnvFile', () => {
  beforeEach(() => {
    delete process.env[VAR];
  });
  afterEach(() => {
    delete process.env[VAR];
  });

  it('names <home>/.fastmail-mcp/.env as the file', () => {
    assert.equal(envFilePath(SANDBOX), join(SANDBOX, '.fastmail-mcp', '.env'));
  });

  it('fills a variable the environment does not set', () => {
    loadHomeEnvFile(homeWith(`${VAR}=from-file\n`));
    assert.equal(process.env[VAR], 'from-file');
  });

  it('leaves a variable the environment already sets', () => {
    process.env[VAR] = 'from-environment';
    loadHomeEnvFile(homeWith(`${VAR}=from-file\n`));
    assert.equal(process.env[VAR], 'from-environment');
  });

  it('leaves a variable the environment sets to an empty string', () => {
    process.env[VAR] = '';
    // Guards the premise: an empty value must survive as a set variable, or this test
    // would pass by checking an unset one.
    assert.equal(process.env[VAR], '');
    loadHomeEnvFile(homeWith(`${VAR}=from-file\n`));
    assert.equal(process.env[VAR], '');
  });

  it('leaves a variable the environment sets to an unsubstituted placeholder', () => {
    process.env[VAR] = '${user_config.fastmail_api_token}';
    loadHomeEnvFile(homeWith(`${VAR}=from-file\n`));
    assert.equal(process.env[VAR], '${user_config.fastmail_api_token}');
  });

  it('does nothing when the .fastmail-mcp directory holds no .env', () => {
    const home = homeWith(null);
    mkdirSync(join(home, '.fastmail-mcp'));
    assert.doesNotThrow(() => loadHomeEnvFile(home));
    assert.equal(process.env[VAR], undefined);
  });

  it('does nothing when the home has no .fastmail-mcp directory at all', () => {
    assert.doesNotThrow(() => loadHomeEnvFile(homeWith(null)));
  });

  it('throws saying plainly that the path is a directory', () => {
    const home = homeWith(null);
    mkdirSync(envFilePath(home), { recursive: true });
    assert.throws(() => loadHomeEnvFile(home), (err: Error) => {
      assert.equal(err.message, `${envFilePath(home)} is a directory, not a file`);
      assert.ok(err.cause instanceof Error, "Node's own error is kept as the cause");
      return true;
    });
  });

  it(
    'throws naming the path when the file cannot be read',
    { skip: process.platform !== 'win32' && process.getuid?.() === 0 ? 'root reads a file whatever its mode' : false },
    () => {
      const home = homeWith(`${VAR}=from-file
`);
      const file = envFilePath(home);
      denyRead(file);
      try {
        assert.throws(
          () => loadHomeEnvFile(home),
          (err: Error) => err.message.includes(file) && err.cause instanceof Error,
        );
      } finally {
        allowRead(file);
      }
    },
  );
});

// What the module above relies on, measured rather than assumed, so a Node release that
// changes it fails here on whichever version CI runs.
describe('process.loadEnvFile', () => {
  let home: string;
  before(() => {
    home = homeWith(`${VAR}=from-file\n`);
  });
  afterEach(() => {
    delete process.env[VAR];
  });

  it('reports a missing file as ENOENT', () => {
    assert.throws(
      () => process.loadEnvFile(join(home, 'absent.env')),
      (err: NodeJS.ErrnoException) => err.code === 'ENOENT',
    );
  });

  it('never overwrites a variable already set, even to an empty string', () => {
    process.env[VAR] = 'from-environment';
    process.loadEnvFile(envFilePath(home));
    assert.equal(process.env[VAR], 'from-environment');
    process.env[VAR] = '';
    process.loadEnvFile(envFilePath(home));
    assert.equal(process.env[VAR], '');
  });
});

describe('load order', () => {
  const SRC_DIR = dirname(fileURLToPath(import.meta.url));

  it("is index.ts's first import, so it runs before any other module reads a setting", () => {
    const source = readFileSync(join(SRC_DIR, 'index.ts'), 'utf8');
    // `export ... from` loads a module too, so it counts as an import here.
    const firstImport = source.split(/\r?\n/).find((line) => /^(import|export)\b/.test(line));
    assert.equal(firstImport, "import './load-env-file.js';");
  });

  it('imports only Node built-ins, so nothing it pulls in can read a setting first', () => {
    const source = readFileSync(join(SRC_DIR, 'load-env-file.ts'), 'utf8');
    // Every module specifier, in any quote style, static, dynamic or require(), on one line or several.
    const specifiers = [...source.matchAll(/\b(?:from|import|require)\s*\(?\s*(['"`])([^'"`]+)\1/g)].map((m) => m[2]);
    assert.ok(specifiers.length > 0, 'no imports found in load-env-file.ts');
    for (const specifier of specifiers) {
      assert.ok(specifier.startsWith('node:'), `load-env-file.ts imports ${specifier}`);
    }
  });
});
