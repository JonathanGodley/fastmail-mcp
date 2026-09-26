/**
 * Tests for scripts/mutation-summary.mjs, which turns Stryker JSON reports into the Markdown the
 * weekly mutation workflow writes to its run summary.
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { summarise } from '../scripts/mutation-summary.mjs';

const mutant = (status: string, line: number, mutatorName = 'StringLiteral') => ({
  id: `${status}-${line}`, mutatorName, replacement: '""', status,
  location: { start: { line, column: 1 }, end: { line, column: 5 } },
});

const shardA = {
  schemaVersion: '2', thresholds: { high: 80, low: 60 },
  files: {
    'src/b.ts': { language: 'typescript', source: '', mutants: [mutant('Killed', 1), mutant('Survived', 9, 'ConditionalExpression')] },
    'src/a.ts': { language: 'typescript', source: '', mutants: [mutant('Timeout', 2), mutant('NoCoverage', 4, 'BlockStatement')] },
  },
};
const shardB = {
  schemaVersion: '2', thresholds: { high: 80, low: 60 },
  files: { 'src/c.ts': { language: 'typescript', source: '', mutants: [mutant('Killed', 3), mutant('Survived', 7), mutant('CompileError', 8)] } },
};

test('totals, score and the untested list are summed across shard reports', () => {
  const md = summarise([shardA, shardB]);
  assert.match(md, /Reports read: 2/);
  // 7 mutants; detected 3 (2 killed + 1 timeout) of 6 scored, so 50%.
  assert.match(md, /\| 7 \| 2 \| 2 \| 1 \| 1 \| 50\.00% \|/);
  assert.match(md, /### Surviving mutants \(3\)/);
  const list = md.split('\n').filter((l) => l.startsWith('- '));
  assert.deepEqual(list, [
    '- `src/a.ts:4` BlockStatement (no coverage)',
    '- `src/b.ts:9` ConditionalExpression',
    '- `src/c.ts:7` StringLiteral',
  ]);
});

test('the list is capped with a count of the rest', () => {
  const md = summarise([shardA, shardB], 2);
  assert.equal(md.split('\n').filter((l) => l.startsWith('- ')).length, 2);
  assert.match(md, /\.\.\.and 1 more/);
});

test('no survivors and no mutants read cleanly', () => {
  const clean = { files: { 'src/a.ts': { mutants: [mutant('Killed', 1)] } } };
  assert.match(summarise([clean]), /100\.00%[\s\S]*No surviving mutants\./);
  assert.match(summarise([{ files: {} }]), /\| 0 \| 0 \| 0 \| 0 \| 0 \| n\/a \|/);
});
