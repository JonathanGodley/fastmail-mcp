import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { generateVTimezone, foldVTimezoneLine } from './vtimezone.js';

function utc(iso: string): number {
  return Date.parse(iso);
}

function expectedUtcStamp(ms: number): string {
  const d = new Date(ms);
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  return `${pad(d.getUTCFullYear(), 4)}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

/** Unfold RFC 5545 continuation lines back into one logical line each. */
function unfold(block: string): string[] {
  return block.split('\r\n').reduce<string[]>((lines, raw) => {
    if (raw.startsWith(' ') && lines.length > 0) lines[lines.length - 1] += raw.slice(1);
    else lines.push(raw);
    return lines;
  }, []);
}

interface ParsedObservance {
  kind: string;
  dtstart: string;
  from: string;
  to: string;
  name: string;
}

/** Every STANDARD/DAYLIGHT sub-component in a generated block, in order. */
function observances(block: string): ParsedObservance[] {
  const result: ParsedObservance[] = [];
  let current: Partial<ParsedObservance> | null = null;
  for (const line of unfold(block)) {
    if (line === 'BEGIN:STANDARD' || line === 'BEGIN:DAYLIGHT') {
      current = { kind: line.slice('BEGIN:'.length) };
    } else if (current) {
      if (line.startsWith('DTSTART:')) current.dtstart = line.slice('DTSTART:'.length);
      else if (line.startsWith('TZOFFSETFROM:')) current.from = line.slice('TZOFFSETFROM:'.length);
      else if (line.startsWith('TZOFFSETTO:')) current.to = line.slice('TZOFFSETTO:'.length);
      else if (line.startsWith('TZNAME:')) current.name = line.slice('TZNAME:'.length);
      else if (line === 'END:STANDARD' || line === 'END:DAYLIGHT') {
        result.push(current as ParsedObservance);
        current = null;
      }
    }
  }
  return result;
}

describe('generateVTimezone', () => {
  it('wraps a TZID matching the zone passed in, BEGIN to END', () => {
    const block = generateVTimezone('Asia/Hong_Kong', utc('2026-06-01T00:00:00+08:00'), utc('2026-06-02T00:00:00+08:00'));
    const lines = unfold(block);
    assert.equal(lines[0], 'BEGIN:VTIMEZONE');
    assert.equal(lines[1], 'TZID:Asia/Hong_Kong');
    assert.equal(lines[lines.length - 1], 'END:VTIMEZONE');
  });

  it('emits the DAYLIGHT onset for Australia/Sydney crossing the October transition', () => {
    const block = generateVTimezone(
      'Australia/Sydney',
      utc('2026-09-20T00:00:00+10:00'),
      utc('2026-10-20T00:00:00+11:00'),
    );
    const daylight = observances(block).find(o => o.kind === 'DAYLIGHT' && o.dtstart === '20261004T020000');
    assert.ok(daylight, block);
    assert.equal(daylight!.from, '+1000');
    assert.equal(daylight!.to, '+1100');
  });

  it('emits the STANDARD onset for America/New_York crossing the November fall-back', () => {
    const block = generateVTimezone(
      'America/New_York',
      utc('2026-10-20T00:00:00-04:00'),
      utc('2026-11-15T00:00:00-05:00'),
    );
    const standard = observances(block).find(o => o.kind === 'STANDARD' && o.dtstart === '20261101T020000');
    assert.ok(standard, block);
    assert.equal(standard!.from, '-0400');
    assert.equal(standard!.to, '-0500');
  });

  it('emits a single STANDARD +0800 observance for Asia/Hong_Kong, which has had no DST since 1979', () => {
    const block = generateVTimezone('Asia/Hong_Kong', utc('2026-06-01T00:00:00+08:00'), utc('2026-06-02T00:00:00+08:00'));
    const obs = observances(block);
    assert.equal(obs.length, 1);
    assert.equal(obs[0].kind, 'STANDARD');
    assert.equal(obs[0].from, '+0800');
    assert.equal(obs[0].to, '+0800');
  });

  it('formats a half-hour offset as +0530 for Asia/Kolkata, a fixed-offset zone', () => {
    const block = generateVTimezone('Asia/Kolkata', utc('2026-06-01T00:00:00+05:30'), utc('2026-06-02T00:00:00+05:30'));
    const obs = observances(block);
    assert.equal(obs.length, 1);
    assert.equal(obs[0].from, '+0530');
    assert.equal(obs[0].to, '+0530');
  });

  it('emits the STANDARD onset for Europe/London crossing the October clock change, offset +0000/+0100', () => {
    const block = generateVTimezone('Europe/London', utc('2026-10-01T00:00:00+01:00'), utc('2026-11-01T00:00:00+00:00'));
    const standard = observances(block).find(o => o.kind === 'STANDARD' && o.dtstart === '20261025T020000');
    assert.ok(standard, block);
    assert.equal(standard!.from, '+0100');
    assert.equal(standard!.to, '+0000');
  });

  it('yields exactly one observance for a span that crosses no transition', () => {
    const block = generateVTimezone('Australia/Sydney', utc('2026-11-01T00:00:00+11:00'), utc('2026-11-02T00:00:00+11:00'));
    assert.equal(observances(block).length, 1);
  });

  it('sets TZUNTIL to the span end, in UTC', () => {
    const spanEndMs = utc('2026-11-02T00:00:00+11:00');
    const block = generateVTimezone('Australia/Sydney', utc('2026-11-01T00:00:00+11:00'), spanEndMs);
    assert.ok(block.includes(`TZUNTIL:${expectedUtcStamp(spanEndMs)}`), block);
  });

  it('names every observance (present), without pinning ICU\'s own spelling of it', () => {
    const block = generateVTimezone(
      'Australia/Sydney',
      utc('2026-09-20T00:00:00+10:00'),
      utc('2026-10-20T00:00:00+11:00'),
    );
    for (const obs of observances(block)) {
      assert.ok(obs.name && obs.name.length > 0, block);
    }
  });
});

describe('foldVTimezoneLine', () => {
  it('returns short lines unchanged', () => {
    assert.equal(foldVTimezoneLine('TZID:Australia/Sydney', '\r\n'), 'TZID:Australia/Sydney');
  });

  it('folds lines longer than 75 octets, continuation lines leading with a space', () => {
    const long = 'TZID:' + 'x'.repeat(80);
    const folded = foldVTimezoneLine(long, '\r\n');
    const lines = folded.split('\r\n');
    assert.ok(Buffer.byteLength(lines[0], 'utf8') <= 75);
    assert.ok(lines[1].startsWith(' '));
  });

  it('keeps every segment within 75 octets for a much longer line', () => {
    const long = 'TZID:' + 'y'.repeat(200);
    const folded = foldVTimezoneLine(long, '\r\n');
    for (const line of folded.split('\r\n')) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75);
    }
  });
});
