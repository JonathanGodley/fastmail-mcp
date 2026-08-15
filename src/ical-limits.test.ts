import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  assertICalTextLimits,
  MAX_ICAL_FIELD_BYTES,
  MAX_ICAL_PARTICIPANTS,
  MAX_ICAL_TOTAL_BYTES,
} from './ical-limits.js';
import { InvalidInputError } from './coerce.js';

// Capture the error a call throws, so every assertion can check the CLASS first.
// Asserting on the message alone would pass just as happily against a plain Error —
// which maps to InternalError ("server fault, retry") instead of InvalidParams
// ("your input, fix it"), i.e. it would miss the point of the rejection entirely.
function captureThrow(fn: () => void): unknown {
  try {
    fn();
  } catch (e) {
    return e;
  }
  assert.fail('expected the call to throw');
}

function overFieldCap(): string {
  return 'x'.repeat(MAX_ICAL_FIELD_BYTES + 1);
}

describe('assertICalTextLimits: per-field cap', () => {
  it('rejects an oversized title', () => {
    const err = captureThrow(() => assertICalTextLimits({ title: overFieldCap() }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /title/);
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_FIELD_BYTES)));
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_FIELD_BYTES + 1)));
  });

  it('rejects an oversized description', () => {
    const err = captureThrow(() => assertICalTextLimits({ title: 'Standup', description: overFieldCap() }));
    assert.ok(err instanceof InvalidInputError);
    assert.equal((err as Error).name, 'InvalidInputError');
    assert.match((err as Error).message, /description/);
  });

  it('rejects an oversized location', () => {
    const err = captureThrow(() => assertICalTextLimits({ title: 'Standup', location: overFieldCap() }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /location/);
  });

  it('measures bytes, not characters, so multi-byte text cannot exceed the octet bound', () => {
    // Each of these is 3 UTF-8 bytes but one JS character: a length check on .length
    // would let ~3x the intended payload through to the octet-based fold.
    const value = '中'.repeat(Math.ceil(MAX_ICAL_FIELD_BYTES / 3) + 1);
    assert.ok(value.length < MAX_ICAL_FIELD_BYTES);
    const err = captureThrow(() => assertICalTextLimits({ title: 'x', description: value }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /description/);
  });

  it('accepts a field exactly at the cap', () => {
    assert.doesNotThrow(() => assertICalTextLimits({ description: 'x'.repeat(MAX_ICAL_FIELD_BYTES) }));
  });
});

describe('assertICalTextLimits: participant-count cap', () => {
  it('rejects more participants than the cap allows', () => {
    const participants = Array.from({ length: MAX_ICAL_PARTICIPANTS + 1 }, (_, i) => ({
      email: `person${i}@example.com`,
    }));
    const err = captureThrow(() => assertICalTextLimits({ title: 'All hands', participants }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /participants/);
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_PARTICIPANTS)));
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_PARTICIPANTS + 1)));
  });

  it('accepts a participant list exactly at the cap', () => {
    const participants = Array.from({ length: MAX_ICAL_PARTICIPANTS }, (_, i) => ({
      email: `person${i}@example.com`,
      name: `Person ${i}`,
    }));
    assert.doesNotThrow(() => assertICalTextLimits({ title: 'All hands', participants }));
  });
});

describe('assertICalTextLimits: participant fields reach the same fold', () => {
  it('rejects one oversized participant name and names its index', () => {
    const participants = [
      { email: 'a@example.com', name: 'Ada' },
      { email: 'b@example.com', name: overFieldCap() },
    ];
    const err = captureThrow(() => assertICalTextLimits({ title: 'Review', participants }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /participants\[1\]\.name/);
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_FIELD_BYTES)));
  });

  it('rejects one oversized participant email and names its index', () => {
    const participants = [{ email: `${'x'.repeat(MAX_ICAL_FIELD_BYTES)}@example.com` }];
    const err = captureThrow(() => assertICalTextLimits({ title: 'Review', participants }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /participants\[0\]\.email/);
  });
});

describe('assertICalTextLimits: total-payload cap', () => {
  it('rejects many fields that are each under the per-field cap', () => {
    // Every value here is legal on its own; only the combined bound catches this.
    const perName = 60 * 1024;
    const count = Math.ceil(MAX_ICAL_TOTAL_BYTES / perName) + 1;
    assert.ok(count <= MAX_ICAL_PARTICIPANTS, 'the count cap must not be what rejects this case');
    const participants = Array.from({ length: count }, (_, i) => ({
      email: `person${i}@example.com`,
      name: 'n'.repeat(perName),
    }));
    const err = captureThrow(() => assertICalTextLimits({ title: 'Congress', participants }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /total/i);
    assert.match((err as Error).message, new RegExp(String(MAX_ICAL_TOTAL_BYTES)));
  });

  it('rejects a title plus description plus location that together exceed the total', () => {
    const chunk = 'y'.repeat(MAX_ICAL_FIELD_BYTES);
    const participants = Array.from({ length: 40 }, (_, i) => ({
      email: `person${i}@example.com`,
      name: 'z'.repeat(3 * 1024),
    }));
    const err = captureThrow(() => assertICalTextLimits({
      title: chunk,
      description: chunk,
      location: chunk,
      participants,
    }));
    assert.ok(err instanceof InvalidInputError);
    assert.match((err as Error).message, /total/i);
  });
});

describe('assertICalTextLimits: legitimate events pass', () => {
  it('accepts a real large event — a 10KB agenda with attendees', () => {
    const agenda = Array.from(
      { length: 200 },
      (_, i) => `${i + 1}. Agenda item ${i + 1}: review the quarterly numbers and agree next steps.`,
    ).join('\n');
    assert.ok(Buffer.byteLength(agenda, 'utf8') > 10 * 1024);
    const participants = Array.from({ length: 30 }, (_, i) => ({
      email: `attendee${i}@example.com`,
      name: `Attendee Number ${i}`,
    }));
    assert.doesNotThrow(() => assertICalTextLimits({
      title: 'Quarterly business review',
      description: agenda,
      location: 'Meeting Room 4, 12 Example Street, London',
      participants,
    }));
  });

  it('accepts an ordinary small event', () => {
    assert.doesNotThrow(() => assertICalTextLimits({
      title: 'Coffee',
      start: undefined,
      participants: [{ email: 'ada@example.com', name: 'Ada Lovelace' }],
    } as any));
  });

  it('accepts an input with no text fields at all', () => {
    assert.doesNotThrow(() => assertICalTextLimits({}));
  });

  it('ignores non-string and malformed values, leaving them to the validators that own them', () => {
    assert.doesNotThrow(() => assertICalTextLimits({
      title: 42,
      description: null,
      participants: [null, 'not-an-object', { email: 7 }],
    }));
  });

  it('cannot bound a participants list that is still a string, which is why coercion runs first', () => {
    // This guard measures an ARRAY. A lenient client's JSON string has no length it can
    // read, so 600 stringified participants pass unmeasured — the count cap and the
    // per-name byte cap both see nothing. The calendar handlers therefore run
    // coerceParticipants (src/coerce.ts) BEFORE this call, so what arrives here is the
    // real array. If that ordering is ever swapped, this is the bound that goes silent.
    const oversized = JSON.stringify(
      Array.from({ length: MAX_ICAL_PARTICIPANTS + 100 }, (_, i) => ({ email: `p${i}@example.com` })),
    );
    assert.doesNotThrow(() => assertICalTextLimits({ participants: oversized }));
    assert.throws(
      () => assertICalTextLimits({ participants: JSON.parse(oversized) }),
      InvalidInputError,
    );
  });
});
