// Typed reads of the call records a `node:test` mock keeps.
//
// `mock.calls[i]` is correctly typed as possibly undefined: a mock that was never
// called has no entry there. Reading straight through it is the single most common
// shape in this repo's tests, and doing it unguarded costs twice — the typechecker
// rejects it, and at runtime a mock that never fired surfaces as "Cannot read
// properties of undefined" somewhere below the assertion, naming neither the mock
// nor the expectation. These helpers assert the call happened first, so a code path
// that stopped calling the collaborator fails on that fact.
//
// This module is test-only. It lives outside the build (see the exclude list in
// tsconfig.json) so it is not shipped in dist/, and outside the top level of src/
// so the "is dist/ stale?" guards, which scan src/*.ts, do not treat it as a
// production source file.

import assert from 'node:assert/strict';

/**
 * The slice of a `node:test` mock these helpers read. Declared structurally rather
 * than as `Mock<F>` so it accepts anything that records calls the same way, and so
 * the argument tuple is inferred from the mock's own signature — a mock declared
 * with typed parameters gives typed `arguments`, with no cast at the read site.
 */
export interface RecordedCalls<Args extends readonly unknown[]> {
  mock: { calls: ReadonlyArray<{ arguments: Args }> };
}

/** The arguments of the `index`-th recorded call, asserting that the call happened. */
export function callArguments<Args extends readonly unknown[]>(
  fn: RecordedCalls<Args>,
  index = 0,
): Args {
  const { calls } = fn.mock;
  const call = calls[index];
  assert.ok(
    call,
    `expected the mock to have been called at least ${index + 1} time(s), ` +
      `but it recorded ${calls.length}`,
  );
  return call.arguments;
}

/**
 * The arguments of the first recorded call matching `match`, asserting that one did.
 *
 * For the batches whose call order is not fixed — where an extra request on one
 * branch shifts the call under test to a different index — so the test names the
 * call it means instead of depending on a position that a nearby change moves.
 * `description` completes the sentence "expected a call ..." in the failure.
 */
export function findCallArguments<Args extends readonly unknown[]>(
  fn: RecordedCalls<Args>,
  match: (args: Args) => boolean,
  description: string,
): Args {
  const { calls } = fn.mock;
  const call = calls.find((c) => match(c.arguments));
  assert.ok(call, `expected a call ${description}, but none of the ${calls.length} recorded matched`);
  return call.arguments;
}
