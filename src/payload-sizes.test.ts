import { it } from 'node:test';
import assert from 'node:assert/strict';
import { PAYLOAD_NAMES, formatTable, producePayloads, readBaseline } from './testing/payload-sizes/measure.js';

// Keeps the payload-size fixtures producible as the read paths change, and prints the figures
// on every CI run. It never compares a size or a digest: a size change is not a failure.
it('produces every payload-size row, each with a parseable JSON block', async (t) => {
  const produced = await producePayloads();
  assert.deepEqual(produced.map((p) => p.name), PAYLOAD_NAMES);
  for (const p of produced) {
    assert.ok('text' in p, `${p.name} could not be produced: ${'error' in p ? p.error : ''}`);
    // Every payload serialises compact, so its JSON is the one line that opens a value.
    const json = p.text.split('\n').find((line) => line.startsWith('[') || line.startsWith('{'));
    assert.ok(json, `${p.name} has no JSON block`);
    const parsed = JSON.parse(json);
    assert.ok(Array.isArray(parsed) ? parsed.length > 0 : Object.keys(parsed).length > 0, `${p.name}'s JSON block is empty`);
  }
  let baseline;
  try {
    baseline = readBaseline();
  } catch (err) {
    t.diagnostic(`baseline unreadable, table skipped: ${err instanceof Error ? err.message : String(err)}`);
    return;
  }
  for (const line of formatTable(baseline, produced).split('\n')) t.diagnostic(line);
});
