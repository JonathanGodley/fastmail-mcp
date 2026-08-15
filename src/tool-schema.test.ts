// Drift guard for the lenient-boolean convention (#54).
//
// Every boolean tool parameter has to satisfy BOTH halves of that convention, and each
// half fails silently on its own:
//
//   1. The schema declares `type: ['boolean', 'string']`. A narrow `type: 'boolean'`
//      makes a validating client reject "true"/"false" before the request is dispatched,
//      so whatever coercion the handler runs is unreachable.
//   2. The handler reads the value through coerceBool, not `!!`. Under `!!`, the string
//      "false" is truthy — the parameter silently inverts.
//
// This asserts against src/index.ts and the handler modules as TEXT rather than by
// spawning the built server and reading tools/list. tools/list would prove what is
// actually shipped, which is the stronger claim, but `npm test` runs tsx over src/ and
// never builds first: a guard read out of a stale dist/ would not see a tool added since
// the last build, which is precisely the drift it exists to catch. tsc does not rewrite
// string literals, so the source and the shipped schema cannot disagree on these, and the
// source check can never go stale. (scripts/mcp-harness.mjs grew a list() for the
// on-demand check against a freshly built server; `node scripts/mcp-harness.mjs --list`.)

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const SCHEMA_FILE = join(SRC_DIR, 'index.ts');

// The files that read tool arguments. The formatters and the JMAP client take already-
// coerced values, and email-formatter.ts uses `raw` as the name of a JMAP object, so
// scanning them would only generate false positives.
const HANDLER_FILES = [
  'index.ts',
  'thread-handler.ts',
  'reply-handler.ts',
  'forward-handler.ts',
  'compose-handler.ts',
  'send-draft-handler.ts',
  'mailbox-handler.ts',
  'contacts-handler.ts',
];

function readLines(file: string): string[] {
  return readFileSync(join(SRC_DIR, file), 'utf8').split('\n').map((l) => l.replace(/\r$/, ''));
}

// Collect the boolean parameters out of the TOOLS schemas. Matches the declaration line
// exactly (leading whitespace, trailing comma) so the prose in nearby comments — which
// quotes both forms — is not picked up.
function collectBooleanParams(): { unionNames: string[]; narrow: string[] } {
  const lines = readLines('index.ts');
  const unionNames = new Set<string>();
  const narrow: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const isUnion = line === "type: ['boolean', 'string'],";
    const isNarrow = line === "type: 'boolean',";
    if (!isUnion && !isNarrow) continue;
    // The property name is the line that opened this block.
    const name = lines[i - 1].trim().replace(/:\s*\{$/, '');
    if (isUnion) unionNames.add(name);
    else narrow.push(`${name} (src/index.ts:${i + 1})`);
  }
  return { unionNames: [...unionNames], narrow };
}

describe('lenient-boolean convention', () => {
  it('declares every boolean tool parameter as type: [boolean, string]', () => {
    const { unionNames, narrow } = collectBooleanParams();
    // A sanity floor: if the scan ever matches nothing (the schema literal is reformatted,
    // TOOLS moves out of this file) the assertion below would pass vacuously.
    assert.ok(
      unionNames.length >= 15,
      `found only ${unionNames.length} boolean parameters; the schema scan has probably stopped matching`,
    );
    assert.deepEqual(
      narrow,
      [],
      `these boolean parameters declare a narrow type: 'boolean', which makes their coerceBool ` +
        `unreachable from a validating client. Widen to type: ['boolean', 'string'] and wrap the ` +
        `description in lenientBool(): ${narrow.join(', ')}`,
    );
  });

  it('reads every boolean tool parameter through coerceBool, never bare !!', () => {
    const { unionNames } = collectBooleanParams();
    const offenders: string[] = [];
    for (const file of HANDLER_FILES) {
      const lines = readLines(file);
      for (const name of unionNames) {
        // `!!name`, `!!(args as any).name` and `!!args?.name`. The negative lookahead
        // keeps `!!raw.hasAttachment` (a property read off a JMAP object that happens to
        // be called `raw`) from matching.
        const patterns = [
          new RegExp(`!!\\s*${name}\\b(?![.\\[])`),
          new RegExp(`!!\\s*\\(\\s*args\\s+as\\s+any\\s*\\)\\.${name}\\b`),
          new RegExp(`!!\\s*args\\?\\.${name}\\b`),
        ];
        lines.forEach((line, i) => {
          if (line.trimStart().startsWith('//')) return;
          if (patterns.some((p) => p.test(line))) offenders.push(`${name} (src/${file}:${i + 1})`);
        });
      }
    }
    assert.deepEqual(
      offenders,
      [],
      `these boolean parameters are read with bare !!, so a lenient client's "false" reads as ` +
        `true. Use coerceBool(...) ?? <default>: ${offenders.join(', ')}`,
    );
  });
});
