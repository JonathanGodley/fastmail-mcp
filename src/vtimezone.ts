// Generates a VTIMEZONE component (RFC 5545 §3.6.5) for an IANA zone, from Node's own ICU
// timezone data, so `caldav-client.ts` can attach one to every `TZID` it writes (#166). This
// server has no timezone database of its own; the offsets below all come from `Intl` through
// `zoneOffsetMsAt`.
//
// The block is not a byte-for-byte reproduction of what Cyrus (or `vzic`) would generate: those
// carry the zone's real `RRULE` observances with their `UNTIL` trimmed to the event's span. This
// generator instead emits one explicit, RRULE-free observance per offset change inside the span,
// plus the observance already in force when the span starts. Both are valid RFC 5545 and resolve
// to identical offsets for every instant the event touches, which is the only property a
// `VTIMEZONE` needs for CalDAV round-tripping — reproducing a zone's recurrence RULE is a
// separate, harder problem this does not attempt.

import { zoneOffsetMsAt } from './coerce.js';
import { foldICalLine } from './ical-fold.js';

const SECOND_MS = 1000;
const DAY_MS = 24 * 60 * 60 * 1000;

// A year plus a day: wide enough that the observance in force at the start of the span is
// findable even for a zone whose transitions are six months apart, narrow enough that a
// day-by-day scan over it stays cheap. 366 covers a leap year's extra day.
const LOOKBACK_MS = 366 * DAY_MS;

interface Transition {
  utcMs: number;
  fromOffsetMs: number;
  toOffsetMs: number;
}

interface Observance {
  /** The transition instant this observance begins at (or the lookback boundary, for the
   * observance already in force with nothing earlier examined). */
  onsetUtcMs: number;
  fromOffsetMs: number;
  toOffsetMs: number;
}

/**
 * Every offset change `zone` makes between `fromMs` and `toMs`, in chronological order.
 *
 * Sampled once a day: no IANA zone has ever changed offset twice within 24 hours, so a change
 * between two day-apart samples can only have one transition behind it. Coarse only in WHERE it
 * looks — each change found is then refined to the exact SECOND by `bisectTransition`.
 *
 * A real calendar event's own span is almost always well under a day, unlike the day-spaced
 * sampling grid: `toMs - fromMs < DAY_MS` left the loop below with no sample point inside the
 * span at all, so a transition strictly between the two endpoints went undetected regardless of
 * how close either sat to it — an overnight event straddling a clock change generated a
 * single-observance VTIMEZONE with the wrong (or right-by-luck) offset for whichever end the
 * lookback happened to match. The tail check closes exactly that gap: once the day-stepping
 * loop stops short of `toMs` (it always does, unless `toMs - fromMs` happens to be an exact
 * multiple of `DAY_MS`), one last comparison against `toMs` itself catches a transition in the
 * remaining partial day the loop never got to sample.
 *
 * `fromMs` is floored to a whole second first, and stays whole-second-aligned at every sample
 * after that (`DAY_MS` is itself a whole number of seconds) — see `bisectTransition` for why that
 * matters.
 */
function findTransitions(zone: string, fromMs: number, toMs: number): Transition[] {
  const transitions: Transition[] = [];
  let prevMs = Math.floor(fromMs / SECOND_MS) * SECOND_MS;
  let prevOffsetMs = zoneOffsetMsAt(prevMs, zone);
  for (let t = prevMs + DAY_MS; t <= toMs; t += DAY_MS) {
    const offsetMs = zoneOffsetMsAt(t, zone);
    if (offsetMs !== prevOffsetMs) {
      transitions.push({ utcMs: bisectTransition(zone, prevMs, t, prevOffsetMs), fromOffsetMs: prevOffsetMs, toOffsetMs: offsetMs });
      prevOffsetMs = offsetMs;
    }
    prevMs = t;
  }
  if (prevMs < toMs) {
    const offsetMs = zoneOffsetMsAt(toMs, zone);
    if (offsetMs !== prevOffsetMs) {
      transitions.push({ utcMs: bisectTransition(zone, prevMs, toMs, prevOffsetMs), fromOffsetMs: prevOffsetMs, toOffsetMs: offsetMs });
    }
  }
  return transitions;
}

/**
 * The exact second `zone` moves away from `lowOffsetMs`, somewhere in `(lowMs, highMs]`. `lowMs`
 * is known to still be at `lowOffsetMs`, `highMs` is known not to be, and (per `findTransitions`)
 * there is exactly one transition between them — so ordinary bisection on the step function
 * `zoneOffsetMsAt` finds the boundary exactly.
 *
 * Bisection stays on WHOLE-SECOND instants throughout (both inputs are already whole-second —
 * see `findTransitions` — and every midpoint computed here is too), never probing a sub-second
 * instant. `zoneOffsetMsAt` reads the offset by formatting an instant to WHOLE SECONDS and
 * comparing that reconstructed instant back to the original `utcMs`: fed a `utcMs` that itself
 * carries a sub-second remainder, the remainder leaks into the "offset" it returns, and bisecting
 * past a two-second gap did exactly that — the noise made the step function look like it crossed
 * the boundary up to several minutes early, at whichever millisecond happened to zero out its own
 * remainder. Every real IANA transition lands on a whole minute regardless, so second resolution
 * loses nothing.
 */
function bisectTransition(zone: string, lowMs: number, highMs: number, lowOffsetMs: number): number {
  let lowSec = lowMs / SECOND_MS;
  let highSec = highMs / SECOND_MS;
  while (highSec - lowSec > 1) {
    const midSec = lowSec + Math.floor((highSec - lowSec) / 2);
    if (zoneOffsetMsAt(midSec * SECOND_MS, zone) === lowOffsetMs) lowSec = midSec;
    else highSec = midSec;
  }
  return highSec * SECOND_MS;
}

function pad(n: number, width = 2): string {
  return String(n).padStart(width, '0');
}

/** A UTC offset as RFC 5545's `utc-offset` (§3.3.14): `+HHMM`, seconds appended only when the
 * offset itself carries them (real for some pre-modern zones, never for one this span will hit,
 * but the ABNF allows it so there is no reason to lose the precision if it ever occurs). */
function formatOffset(offsetMs: number): string {
  const sign = offsetMs < 0 ? '-' : '+';
  const totalSeconds = Math.round(Math.abs(offsetMs) / 1000);
  const hh = Math.floor(totalSeconds / 3600);
  const mm = Math.floor((totalSeconds % 3600) / 60);
  const ss = totalSeconds % 60;
  return ss === 0 ? `${sign}${pad(hh)}${pad(mm)}` : `${sign}${pad(hh)}${pad(mm)}${pad(ss)}`;
}

/** The wall-clock reading `utcMs` has when shifted by `offsetMs`, as an unfolded local
 * DATE-TIME (`YYYYMMDDTHHMMSS`) — the shift-then-read-UTC-components trick every wall-clock
 * conversion in `coerce.ts` uses, in the opposite direction from `zoneOffsetMsAt`. */
function localWallClockString(utcMs: number, offsetMs: number): string {
  const d = new Date(utcMs + offsetMs);
  return `${pad(d.getUTCFullYear(), 4)}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}`;
}

/** `utcMs` as an RFC 5545 UTC DATE-TIME (`YYYYMMDDTHHMMSSZ`). */
function toUtcStamp(utcMs: number): string {
  const d = new Date(utcMs);
  return `${pad(d.getUTCFullYear(), 4)}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}` +
    `T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
}

const abbreviationFormatterCache = new Map<string, Intl.DateTimeFormat | null>();

/** ICU's `en-US` short name for `zone` at `utcMs` (e.g. `AEDT`, or `GMT+11` where ICU has no
 * abbreviation for it) — the `TZNAME` value. Cached per zone for the same reason
 * `zoneOffsetMsAt`'s formatter is: one generated block can look this up several times.
 *
 * Construction failure (a name ICU's `Intl.DateTimeFormat` rejects outright) falls back to
 * `zone` itself, the same fallback already used below for a formatter that built but has no
 * abbreviation to offer — mirroring `zoneOffsetMsAt`'s own formatter cache in `coerce.ts`,
 * which guards the identical construction the same way. */
function zoneAbbreviation(zone: string, utcMs: number): string {
  let formatter = abbreviationFormatterCache.get(zone);
  if (formatter === undefined) {
    try {
      formatter = new Intl.DateTimeFormat('en-US', {
        timeZone: zone, timeZoneName: 'short',
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
      });
    } catch {
      formatter = null;
    }
    abbreviationFormatterCache.set(zone, formatter);
  }
  if (!formatter) return zone;
  const name = formatter.formatToParts(new Date(utcMs)).find(p => p.type === 'timeZoneName')?.value;
  return name ?? zone;
}

/**
 * Generate a `VTIMEZONE` component for `zone`, valid across `[spanStartUtcMs, spanEndUtcMs]`
 * (the event's own start and end, as UTC instants) — folded content lines joined by
 * `lineEnding`, `BEGIN:VTIMEZONE` through `END:VTIMEZONE` inclusive.
 *
 * One observance covers the instant the span starts at, found by scanning backwards over a
 * lookback of a year plus a day: a zone with no transition anywhere in that lookback is
 * fixed-offset for this purpose, and that observance's `DTSTART` is then the lookback boundary
 * itself (there being no real onset in range to report). Every further offset change up to
 * `spanEndUtcMs` gets its own observance. `TZOFFSETFROM`/`TZOFFSETTO` are exact per observance;
 * `TZNAME` is whatever ICU calls the zone at that observance's onset. Per RFC 5545 §3.6.5, an
 * observance's `DTSTART` is the transition instant read in the OLD (`TZOFFSETFROM`) offset — the
 * wall clock reading immediately before the jump, e.g. Sydney's `20261004T020000` immediately
 * before the local clock skips to 3am.
 *
 * The larger UTC offset among the observances found is `DAYLIGHT` and the rest `STANDARD` — DST
 * advances the clock in every hemisphere, so this holds regardless of which side of the equator
 * `zone` is on — and a zone with only one offset in range is `STANDARD` outright.
 *
 * `TZUNTIL` (a Cyrus/Apple extension, not in RFC 5545 itself) is always set to `spanEndUtcMs`,
 * matching the bound Cyrus's own `icalcomponent_add_required_timezones` puts on the block it
 * attaches server-side — see `docs/conventions.md`'s VTIMEZONE section.
 */
export function generateVTimezone(
  zone: string,
  spanStartUtcMsInput: number,
  spanEndUtcMsInput: number,
  lineEnding: string = '\r\n',
): string {
  // Floored to a whole second so every instant this function samples is one — see
  // `bisectTransition`'s comment for why a sub-second instant fed to `zoneOffsetMsAt` reads a
  // corrupted offset. RFC 5545 has no sub-second datetime form, so every real caller's span is
  // already whole-second; this only guards a caller that isn't.
  const spanStartUtcMs = Math.floor(spanStartUtcMsInput / SECOND_MS) * SECOND_MS;
  const spanEndUtcMs = Math.floor(spanEndUtcMsInput / SECOND_MS) * SECOND_MS;
  const lookbackStartMs = spanStartUtcMs - LOOKBACK_MS;
  const priorTransitions = findTransitions(zone, lookbackStartMs, spanStartUtcMs);

  let initial: Observance;
  if (priorTransitions.length > 0) {
    const onset = priorTransitions[priorTransitions.length - 1];
    initial = { onsetUtcMs: onset.utcMs, fromOffsetMs: onset.fromOffsetMs, toOffsetMs: onset.toOffsetMs };
  } else {
    const offsetMs = zoneOffsetMsAt(lookbackStartMs, zone);
    initial = { onsetUtcMs: lookbackStartMs, fromOffsetMs: offsetMs, toOffsetMs: offsetMs };
  }

  const spanTransitions = findTransitions(zone, spanStartUtcMs, spanEndUtcMs);
  const observances: Observance[] = [
    initial,
    ...spanTransitions.map(t => ({ onsetUtcMs: t.utcMs, fromOffsetMs: t.fromOffsetMs, toOffsetMs: t.toOffsetMs })),
  ];

  // Both offsets of every observance, not just what it changes TO: a span with a single
  // observance (no transition inside it) still has a `fromOffsetMs` inherited from whatever
  // observance preceded it — a Sydney span sitting entirely inside daylight saving carries only
  // one observance, `to: +1100`, but its `from: +1000` is what makes the +1000/+1100 PAIR
  // visible at all. Reading `to` alone saw one offset, called the zone fixed-offset, and
  // labelled that lone DAYLIGHT observance STANDARD.
  const distinctOffsets = Array.from(new Set(observances.flatMap(o => [o.fromOffsetMs, o.toOffsetMs])));
  const daylightOffsetMs = distinctOffsets.length > 1 ? Math.max(...distinctOffsets) : null;

  const lines: string[] = ['BEGIN:VTIMEZONE', `TZID:${zone}`, `TZUNTIL:${toUtcStamp(spanEndUtcMs)}`];
  for (const obs of observances) {
    const kind = obs.toOffsetMs === daylightOffsetMs ? 'DAYLIGHT' : 'STANDARD';
    lines.push(
      `BEGIN:${kind}`,
      `DTSTART:${localWallClockString(obs.onsetUtcMs, obs.fromOffsetMs)}`,
      `TZOFFSETFROM:${formatOffset(obs.fromOffsetMs)}`,
      `TZOFFSETTO:${formatOffset(obs.toOffsetMs)}`,
      `TZNAME:${zoneAbbreviation(zone, obs.onsetUtcMs)}`,
      `END:${kind}`,
    );
  }
  lines.push('END:VTIMEZONE');

  return lines.map(l => foldICalLine(l, lineEnding)).join(lineEnding);
}
