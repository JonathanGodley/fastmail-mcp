import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { generateVTimezone, findTransitions, bisectTransition } from './vtimezone.js';
import { InvalidInputError, utcMsFromComponents } from './coerce.js';

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

  it('detects a transition even when the span itself is under a day — an overnight event straddling the October transition', () => {
    // A sub-day span, unlike the month-wide spans the other transition tests use: the day loop
    // never samples inside it, so this pins the tail check that catches the transition anyway.
    const block = generateVTimezone(
      'Australia/Sydney',
      utc('2026-10-03T23:00:00+10:00'),
      utc('2026-10-04T05:00:00+11:00'),
    );
    const obs = observances(block);
    assert.equal(obs.length, 2, block);
    assert.ok(obs.some(o => o.kind === 'STANDARD' && o.to === '+1000'), block);
    assert.ok(obs.some(o => o.kind === 'DAYLIGHT' && o.to === '+1100'), block);
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

  it('adds no extra observance for a sub-day span with no transition — the day-loop never runs, only the tail check does', () => {
    // Unlike the exact-24h Hong Kong span above (whose single day-loop iteration lands exactly
    // on toMs, so the tail check's own guard skips it), this 3-hour span leaves the day-loop with
    // nothing to do at all — every offset comparison here happens in the tail check itself. A
    // zone with no DST ever guarantees the two offsets it compares are equal, so this is the one
    // case that shows the tail check adding nothing when there is truly nothing to add.
    const block = generateVTimezone('Asia/Hong_Kong', utc('2026-06-01T00:00:00+08:00'), utc('2026-06-01T03:00:00+08:00'));
    assert.equal(observances(block).length, 1, block);
  });

  it('finds no transitions and emits no bogus observance for a span entirely after Sydney\'s October change (#166)', () => {
    // Both ends of this span are already past the spring-forward: findTransitions itself must
    // report none, not just generateVTimezone folding a spurious one away downstream.
    const from = utc('2026-10-03T17:00:00Z');
    const to = utc('2026-10-05T17:00:00Z');
    assert.equal(findTransitions('Australia/Sydney', from, to).length, 0);
    const block = generateVTimezone('Australia/Sydney', from, to);
    assert.equal(observances(block).length, 1, block);
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

  it('folds a long TZID line at 75 octets, per RFC 5545 §3.1 (shared foldICalLine, #166)', () => {
    // No real IANA zone name is remotely this long; the fold itself is `ical-fold.ts`'s own
    // unit-tested responsibility (shared with caldav-client.ts), so this only needs to prove the
    // generator actually wires its output through it, with a continuation line to show for it.
    const zone = 'Fake/' + 'x'.repeat(200);
    const block = generateVTimezone(zone, utc('2026-06-01T00:00:00Z'), utc('2026-06-01T01:00:00Z'));
    const lines = block.split('\r\n');
    for (const line of lines) {
      assert.ok(Buffer.byteLength(line, 'utf8') <= 75, line);
    }
    assert.ok(lines.some(l => l.startsWith(' ')), block);
  });

  it('labels a lone observance DAYLIGHT when it is the higher of the two offsets a prior transition set up, even with no transition IN the span', () => {
    // January is peak daylight saving in Sydney (+1100), reached by the October transition
    // outside this span. Classifying by `to`-offsets alone sees one distinct value here and
    // calls the zone fixed-offset, mislabelling this DAYLIGHT observance STANDARD.
    const block = generateVTimezone('Australia/Sydney', utc('2027-01-05T00:00:00+11:00'), utc('2027-01-06T00:00:00+11:00'));
    const obs = observances(block);
    assert.equal(obs.length, 1, block);
    assert.equal(obs[0].kind, 'DAYLIGHT');
    assert.equal(obs[0].to, '+1100');
  });

  it('labels every Pacific/Apia observance STANDARD across its 2011 date-line move, the three-or-more-offsets labelling limit (#166)', () => {
    // See generateVTimezone's own "two known limits" comment for why: this zone crosses FOUR
    // distinct offsets in this window (-1100, -1000, +1400, +1300), so the lowest-offset-reversion
    // check the labelling relies on never fires here. -1000 and +1400 were both real DST periods,
    // so labelling them STANDARD is wrong in the real world — this test pins that wrong label
    // deliberately, as a known limit, and should flip to expect DAYLIGHT on those two observances
    // if that limit is ever lifted.
    const block = generateVTimezone('Pacific/Apia', utc('2011-03-01T00:00:00Z'), utc('2012-06-01T00:00:00Z'));
    const obs = observances(block);
    assert.ok(obs.length > 0, block);
    assert.ok(obs.every(o => o.kind === 'STANDARD'), block);
  });

  it('keeps Europe/Moscow STANDARD across its 2011 permanent step from +0300 to +0400', () => {
    // Moscow abolished its fall-back after 2011's spring-forward: the offset touches exactly two
    // values, the same shape a real DST cycle has, but never reverts. The lookback window (a
    // year plus a day before the span) stays after 2010's last REAL fall-back (31 Oct 2010), so
    // only the permanent step is in range — nothing here reverts +0400 back to +0300.
    const block = generateVTimezone('Europe/Moscow', utc('2011-12-01T00:00:00Z'), utc('2011-12-02T00:00:00Z'));
    const obs = observances(block);
    assert.equal(obs.length, 1, block);
    assert.equal(obs[0].kind, 'STANDARD', block);
    assert.equal(obs[0].to, '+0400');
  });

  it('keeps Asia/Pyongyang STANDARD across its 2018 permanent step from +0830 to +0900', () => {
    // Pyongyang reversed its 2015 shift to +08:30 in 2018, moving back to +09:00 permanently.
    // The lookback window stays after the 2015 change, so only the 2018 step is in range, and it
    // never reverts either.
    const block = generateVTimezone('Asia/Pyongyang', utc('2018-06-01T00:00:00Z'), utc('2018-06-02T00:00:00Z'));
    const obs = observances(block);
    assert.equal(obs.length, 1, block);
    assert.equal(obs[0].kind, 'STANDARD', block);
    assert.equal(obs[0].to, '+0900');
  });

  it('refuses a span starting in year 1, whose 366-day lookback reaches into year 0', () => {
    // `Date.UTC(1, 0, 15)` is NOT year 1 — JS maps a year in 0..99 to 1900+n, so that call is
    // 1901-01-15. `utcMsFromComponents` (shared with the production constant) defeats the same
    // mapping, giving a genuine year-1 instant whose 366-day lookback reaches into year 0.
    const spanStart = utcMsFromComponents(1, 1, 15, 0, 0, 0); // 0001-01-15
    assert.throws(
      () => generateVTimezone('UTC', spanStart, spanStart + 1000),
      InvalidInputError,
    );
  });

  it('does not refuse an 1850 span, whose lookback stays comfortably after the boundary', () => {
    // Pins that the guard above is a real boundary rather than something that rejects every early
    // date. This is a comfortable case, not an edge one: the actual boundary for a span start is 3
    // January of year 2 (pinned exactly in the test below), far earlier than this 1850 fixture.
    const spanStart = utcMsFromComponents(1850, 1, 15, 0, 0, 0);
    assert.doesNotThrow(() => generateVTimezone('UTC', spanStart, spanStart + 1000));
  });

  it('pins the exact UTC boundary the thrown message states: accepts 0002-01-03T00:00:00Z, refuses one second earlier', () => {
    // MIN_LOOKBACK_START_MS and LOOKBACK_MS are arithmetic, not directly asserted elsewhere; this
    // pins the boundary they produce against the exact instant the thrown message names.
    const accepted = utc('0002-01-03T00:00:00Z');
    assert.doesNotThrow(() => generateVTimezone('UTC', accepted, accepted + 1000));
    const refused = utc('0002-01-02T23:59:59Z');
    assert.throws(
      () => generateVTimezone('UTC', refused, refused + 1000),
      /Cannot generate a VTIMEZONE this far back: the event must start on or after 3 January of year 2 \(UTC\)\./,
    );
  });

  it('refuses a span of 101 years (#166)', () => {
    // findTransitions samples once per day and generateVTimezone emits one observance per
    // transition found — both linear in the span itself, which is caller-controlled up to a
    // DTEND in year 9999 (validateAndFormatICalDate's own ceiling). MAX_VTIMEZONE_SPAN_DAYS
    // bounds that before either cost is paid.
    const spanStart = utcMsFromComponents(2000, 1, 1, 0, 0, 0);
    const spanEnd = utcMsFromComponents(2101, 1, 1, 0, 0, 0);
    assert.throws(() => generateVTimezone('Australia/Sydney', spanStart, spanEnd), InvalidInputError);
  });

  it('does not refuse a span of 99 years', () => {
    // Guards against an over-tight constant: a span comfortably inside the limit must still
    // succeed.
    const spanStart = utcMsFromComponents(2000, 1, 1, 0, 0, 0);
    const spanEnd = utcMsFromComponents(2099, 1, 1, 0, 0, 0);
    assert.doesNotThrow(() => generateVTimezone('Australia/Sydney', spanStart, spanEnd));
  });

  it('pins the exact MAX_VTIMEZONE_SPAN_DAYS boundary the thrown message states: 36600 days is accepted, 36601 is refused', () => {
    const DAY_MS = 24 * 60 * 60 * 1000;
    const spanStart = utcMsFromComponents(2000, 1, 1, 0, 0, 0);
    assert.doesNotThrow(() => generateVTimezone('UTC', spanStart, spanStart + 36600 * DAY_MS));
    assert.throws(
      () => generateVTimezone('UTC', spanStart, spanStart + 36600 * DAY_MS + 1000),
      /Event spans 36601 days; VTIMEZONE generation is limited to 36600 days\. Shorten the event\./,
    );
  });

  it('pins the exact year-10000 boundary the thrown message states: accepts 9999-12-31T23:59:59Z, refuses 10000-01-01T00:00:00Z', () => {
    // MAX_VTIMEZONE_SPAN_DAYS bounds the span's LENGTH only; nothing else bounded the span's END
    // instant, so a short span whose end still lands at or after year 10000 would otherwise reach
    // toUtcStamp and print a 5-digit year, which RFC 5545's fixed-4-digit-year DATE-TIME form
    // cannot represent (#166).
    const spanStart = utcMsFromComponents(9999, 12, 30, 0, 0, 0);
    const accepted = utcMsFromComponents(9999, 12, 31, 23, 59, 59);
    assert.doesNotThrow(() => generateVTimezone('UTC', spanStart, accepted));
    const refused = utcMsFromComponents(10000, 1, 1, 0, 0, 0);
    assert.throws(
      () => generateVTimezone('UTC', spanStart, refused),
      /Cannot generate a VTIMEZONE this far ahead: the event must end before year 10000 \(UTC\)\./,
    );
  });

  it('gives Sydney\'s STANDARD and DAYLIGHT observances distinct TZNAMEs, each a plausible shape (#166)', () => {
    const block = generateVTimezone(
      'Australia/Sydney',
      utc('2026-09-20T00:00:00+10:00'),
      utc('2026-10-20T00:00:00+11:00'),
    );
    const obs = observances(block);
    assert.equal(obs.length, 2, block);
    for (const o of obs) {
      assert.ok(/^[A-Za-z]+$/.test(o.name) || /^GMT[+-]\d+(:\d{2})?$/.test(o.name), `unexpected TZNAME shape: ${o.name}`);
    }
    assert.notEqual(obs[0].name, obs[1].name, block);
  });

  it('pins the initial observance\'s DTSTART for a Sydney span entirely after the April 2026 fallback', () => {
    // A span with no transition of its own still has to find the RIGHT prior onset by scanning
    // backwards — this pins that onset to the exact date/offsets rather than just its count.
    const block = generateVTimezone('Australia/Sydney', utc('2026-04-10T00:00:00+10:00'), utc('2026-04-11T00:00:00+10:00'));
    const obs = observances(block);
    assert.equal(obs.length, 1, block);
    assert.equal(obs[0].dtstart, '20260405T030000', block);
    assert.equal(obs[0].from, '+1100');
    assert.equal(obs[0].to, '+1000');
  });

  it('finds both transitions in a six-month span crossing Sydney\'s April and October changes', () => {
    const block = generateVTimezone('Australia/Sydney', utc('2026-04-01T00:00:00+11:00'), utc('2026-10-05T01:00:00+11:00'));
    const obs = observances(block);
    assert.equal(obs.length, 3, block); // in-force-at-start, April fallback, October springforward
    const april = obs.find(o => o.dtstart === '20260405T030000');
    const october = obs.find(o => o.dtstart === '20261004T020000');
    assert.ok(april, block);
    assert.equal(april!.from, '+1100');
    assert.equal(april!.to, '+1000');
    assert.ok(october, block);
    assert.equal(october!.from, '+1000');
    assert.equal(october!.to, '+1100');
  });

  it('formats Africa/Monrovia\'s pre-1972 -00:44:30 offset with seconds, per the utc-offset ABNF', () => {
    // Liberia ran 44 minutes 30 seconds behind UTC until 1972 — one of the few IANA zones whose
    // historical offset is not a whole minute.
    const block = generateVTimezone('Africa/Monrovia', utc('1970-06-01T00:00:00Z'), utc('1970-06-02T00:00:00Z'));
    const obs = observances(block);
    assert.equal(obs.length, 1, block);
    assert.equal(obs[0].to, '-004430', block);
  });
});

describe('findTransitions', () => {
  it('returns the right transition for a fractional toMs, because findTransitions floors it before bisectTransition\'s precondition check (#166)', () => {
    // Sydney's 2026 spring-forward instant is 2026-10-03T16:00:00Z (+10:00 -> +11:00); toMs here
    // is half a second past it. This pins that findTransitions' whole-second floor on toMs keeps
    // bisectTransition's own whole-second precondition satisfied end to end, rather than letting a
    // fractional bound reach it and be rejected.
    const fromMs = Date.parse('2026-10-03T12:00:00Z');
    const toMs = Date.parse('2026-10-03T16:00:00Z') + 500;
    const transitions = findTransitions('Australia/Sydney', fromMs, toMs);
    assert.equal(transitions.length, 1, JSON.stringify(transitions));
    assert.equal(transitions[0].utcMs, Date.parse('2026-10-03T16:00:00Z'));
    assert.equal(transitions[0].fromOffsetMs, 10 * 3600 * 1000);
    assert.equal(transitions[0].toOffsetMs, 11 * 3600 * 1000);
  });
});

describe('bisectTransition', () => {
  // findTransitions always floors before calling in, so nothing exercises this check through that
  // path. Pinned directly here instead. Bounds sit an hour either side of Sydney's real 2026
  // spring-forward transition rather than tight against it (unlike the findTransitions test above):
  // with the check removed, bounds tight against a transition bisect forever once the gap narrows
  // to a fractional (1, 2) seconds, while these wider bounds return without hanging
  // — keep them wide if these tests are ever changed, or a red run here hangs instead of failing.
  const wholeSecondMs = Date.parse('2026-10-03T16:00:00Z');
  const lowOffsetMs = 10 * 3600 * 1000;

  it('refuses a fractional low bound', () => {
    assert.throws(
      () => bisectTransition('Australia/Sydney', wholeSecondMs - 3600000 + 0.7, wholeSecondMs + 3600000, lowOffsetMs),
      /Timezone transition bisection requires whole-second bounds\./,
    );
  });

  it('refuses a fractional high bound', () => {
    assert.throws(
      () => bisectTransition('Australia/Sydney', wholeSecondMs - 3600000, wholeSecondMs + 3600000 + 0.7, lowOffsetMs),
      /Timezone transition bisection requires whole-second bounds\./,
    );
  });
});
