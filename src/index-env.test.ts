// Drift guard for the env-resolution convention.
//
// Every env-derived setting is resolved through findEnvValue() in src/index.ts, which
// tries a list of names in order so that a DXT extension's user_config key reaches the
// server, and which rejects an unsubstituted "${...}" placeholder rather than passing the
// literal through. A module that reads process.env directly gets none of that: it sees one
// spelling, so a DXT install silently fails to configure the setting, and a placeholder is
// taken as a real value. That is exactly how the CalDAV ORGANIZER display name behaved
// before it was moved into the constructor config — nothing complained, because nothing
// was checking.
//
// This reads the sources as TEXT rather than importing them, for the same reason the
// lenient-boolean guard does: `npm test` runs tsx over src/ and never builds, so a check
// that inspected dist/ would read whatever was compiled last and miss a module added since.
//
// Test files are deliberately out of scope — a test may legitimately set or stub
// process.env for a fixture, and spawning-the-server tests copy the whole environment.
// The convention constrains the shipped server, not its harnesses.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));

// Every direct process.env read the shipped sources are allowed to contain. Each entry is
// one exact expression in one file with the reason it is not a convention violation; a
// loose pattern would re-admit the whole class of bypasses this guard exists to catch.
const ALLOWED_READS: Array<{ file: string; expression: string; reason: string }> = [
  {
    file: 'index.ts',
    expression: 'process.env[key]',
    reason:
      'The single lookup inside findEnvValue that every setting goes through. Only ' +
      'permitted within that function body, which is asserted separately below.',
  },
  {
    file: 'index.ts',
    expression: 'process.env.DEBUG',
    reason:
      "Not a server setting: the debug package's own switch, re-read at module load to " +
      'silence tsdav request logging before any configuration is resolved.',
  },
];

function sourceFiles(): string[] {
  return readdirSync(SRC_DIR)
    .filter((f) => f.endsWith('.ts') && !f.endsWith('.test.ts') && !f.endsWith('.d.ts'))
    .sort();
}

// Drop comments so prose that quotes a process.env expression (including the comments in
// this convention's own code) is not mistaken for a read. A "//" inside a string literal
// would blank the rest of that line, so protocol-relative "://" is left alone; the residual
// risk is a missed read sitting after a string literal on the same line, which is a far
// smaller hole than counting every comment as a violation.
function stripComments(source: string): string {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .split('\n')
    .map((line) => line.replace(/(^|[^:])\/\/.*$/, '$1'))
    .join('\n');
}

interface EnvRead {
  file: string;
  line: number;
  expression: string;
}

function collectEnvReads(): EnvRead[] {
  const reads: EnvRead[] = [];
  for (const file of sourceFiles()) {
    const lines = stripComments(readFileSync(join(SRC_DIR, file), 'utf8')).split('\n');
    lines.forEach((line, i) => {
      const pattern = /process\.env\s*(?:\.\s*[A-Za-z_$][\w$]*|\[[^\]]*\])/g;
      for (const match of line.matchAll(pattern)) {
        reads.push({ file, line: i + 1, expression: match[0].replace(/\s+/g, '') });
      }
    });
  }
  return reads;
}

// The inclusive line range of findEnvValue's body in index.ts, located by its declaration
// and the first closing brace at column 0 after it.
function findEnvValueRange(): { start: number; end: number } {
  const lines = readFileSync(join(SRC_DIR, 'index.ts'), 'utf8').split('\n');
  const start = lines.findIndex((l) => l.startsWith('function findEnvValue('));
  assert.notEqual(start, -1, 'findEnvValue() was not found in src/index.ts');
  const end = lines.findIndex((l, i) => i > start && l.startsWith('}'));
  assert.notEqual(end, -1, 'the end of findEnvValue() was not found in src/index.ts');
  return { start: start + 1, end: end + 1 };
}

describe('env-resolution convention', () => {
  it('resolves every setting through findEnvValue rather than reading process.env', () => {
    const reads = collectEnvReads();
    // A floor, so a regex that silently stops matching cannot make this pass vacuously.
    assert.ok(
      reads.length >= ALLOWED_READS.length,
      `expected at least the ${ALLOWED_READS.length} known process.env reads, found ${reads.length}`,
    );

    const unexpected = reads.filter(
      (r) => !ALLOWED_READS.some((a) => a.file === r.file && a.expression === r.expression),
    );
    assert.deepEqual(
      unexpected,
      [],
      'Direct process.env reads outside findEnvValue: ' +
        unexpected.map((r) => `${r.expression} (src/${r.file}:${r.line})`).join(', ') +
        '. Resolve the setting in index.ts through findEnvValue with the full name list ' +
        "(FASTMAIL_X, USER_CONFIG_FASTMAIL_X, USER_CONFIG_fastmail_x, fastmail_x) and pass " +
        'the value into the module, or add an entry to ALLOWED_READS stating why it is exempt.',
    );

    // Every listed exception must still exist, or the list is stale and is quietly widening
    // what the guard permits.
    for (const allowed of ALLOWED_READS) {
      assert.ok(
        reads.some((r) => r.file === allowed.file && r.expression === allowed.expression),
        `ALLOWED_READS lists ${allowed.expression} in src/${allowed.file}, which no longer exists — remove the entry`,
      );
    }
  });

  it("keeps the process.env[key] lookup inside findEnvValue's own body", () => {
    const { start, end } = findEnvValueRange();
    const bracketReads = collectEnvReads().filter((r) => r.expression.startsWith('process.env['));
    assert.ok(bracketReads.length > 0, 'the computed process.env lookup was not found');
    for (const read of bracketReads) {
      assert.ok(
        read.file === 'index.ts' && read.line >= start && read.line <= end,
        `computed process.env lookup at src/${read.file}:${read.line} is outside findEnvValue (src/index.ts:${start}-${end})`,
      );
    }
  });

  it('resolves the CalDAV settings and the blob-attach opt-in under all four configuration spellings', () => {
    // The four-name list is what carries a DXT user_config key through to the server. The
    // CalDAV credentials and display name each resolved fewer names than that, so a DXT
    // install configured CalDAV and got calendar tools that reported themselves
    // unavailable with nothing to explain why.
    const source = readFileSync(join(SRC_DIR, 'index.ts'), 'utf8');
    // allow_blob_attach rides along here rather than in its own test: it is a declared
    // manifest.json user_config key like the three CalDAV settings, so the same four-name
    // resolution is what carries a DXT host's answer through to the server, whichever
    // spelling that host uses. (The base-URL kill switch is deliberately NOT in this list —
    // it resolves one name on purpose, and is not in the manifest either; see getAuthConfig.)
    for (const setting of ['caldav_username', 'caldav_password', 'caldav_display_name', 'allow_blob_attach']) {
      const upper = setting.toUpperCase();
      for (const name of [
        `'FASTMAIL_${upper}'`,
        `'USER_CONFIG_FASTMAIL_${upper}'`,
        `'USER_CONFIG_fastmail_${setting}'`,
        `'fastmail_${setting}'`,
      ]) {
        assert.ok(source.includes(name), `src/index.ts does not resolve ${name}`);
      }
    }
  });
});
