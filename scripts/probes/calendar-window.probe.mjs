// What this probe settles
// -----------------------
// `calendar-expand.probe.mjs` measured the PLATFORM: what Fastmail's CalDAV server returns
// for an expanded time-range query. It talks to tsdav directly and says nothing about what
// this server does with the answer. This probe settles the other half:
//
//   does the shipped `list_calendar_events` tool now answer a window question with dates
//   that are actually in that window?
//
// The two halves of #64 failed at different layers (no expand was requested, and the parser
// read only the first VEVENT), and either alone reproduces the symptom (1-10 March 2027
// returned a single event dated 2020-08-28), so only the end-to-end path proves both are
// closed.
//
// This spawns the BUILT server (dist/index.js) through the shared MCP harness and asserts on
// the tool's real response:
//   - every returned `start` falls inside the requested window
//   - a recurring entry is marked as an occurrence (`recurrenceId`, no `RRULE`), not as a
//     series master (#64 request 3)
//   - the response opens with a summary line stating the total (#100)
//   - a window containing a series' OWN start date reports EVERY occurrence in it, not one:
//     Fastmail leaves the first instance without a RECURRENCE-ID
//   - the rows come back in ascending order of the INSTANT each start names: a real calendar
//     mixes zone-stripped wall clocks with UTC values, and `limit` slices this list
//   - a window given only ONE bound is bounded, and the response says so
//   - a call given NO bounds is bounded to the next month, and the note says neither bound
//     was given (#142)
//
// READ-ONLY: creating or deleting an event with participants sends real iTIP mail (see the
// README).
//
// WHAT IT DOES NOT COVER: a non-UTC zone, unless one is supplied. With no FASTMAIL_TIMEZONE
// the server reads dates in the host zone, so on a UTC host the "date-only window is the
// caller's LOCAL day" assertion holds trivially. The real zone coverage is src/coerce.test.ts
// and src/caldav-client.test.ts; export FASTMAIL_TIMEZONE=Australia/Sydney (or set it in the
// MCP client config, which wins over the shell) to exercise a non-zero offset live.
//
// Run: python scripts/probes/run-probe.py calendar-window.probe.mjs [startDate endDate [singleDay]]
// Requires FASTMAIL_API_TOKEN plus FASTMAIL_CALDAV_USERNAME/PASSWORD; the launcher injects
// all three from the local MCP client config. Build first: npm run build.

import { createClient } from '../mcp-harness.mjs';
import { makeChecker, jsonOf, text } from './probelib.mjs';

const START = process.argv[2] || '2027-03-01';
const END = process.argv[3] || '2027-03-10';
// The date the local-day check asks about. Defaults to the date the wrong-day window was
// reported on; override it to point at any day whose contents you know.
const SINGLE_DAY = process.argv[4] || '2026-08-12';

if (!process.env.FASTMAIL_CALDAV_USERNAME || !process.env.FASTMAIL_CALDAV_PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-window.probe.mjs');
  process.exit(1);
}

const { check, failures } = makeChecker();

// The same fallback the server takes, so both agree about which day is being asked about.
const ZONE = (process.env.FASTMAIL_TIMEZONE || '').trim() || Intl.DateTimeFormat().resolvedOptions().timeZone;

function offsetAt(ms) {
  const p = new Intl.DateTimeFormat('en-US', {
    timeZone: ZONE, hour12: false,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(new Date(ms));
  const g = (t) => Number(p.find((x) => x.type === t)?.value);
  return Date.UTC(g('year'), g('month') - 1, g('day'), g('hour') % 24, g('minute'), g('second')) - ms;
}

/**
 * The UTC instant a wall clock names in ZONE — the server's own rule, mirrored.
 *
 * A repeated wall clock takes the earlier instant, a skipped one resolves forward to the
 * transition. Mirroring rather than hard-coding instants makes a change to that rule show up
 * as a failing assertion here.
 */
function wallClockMs(y, mo, d, h = 0, mi = 0, s = 0) {
  const naive = Date.UTC(y, mo - 1, d, h, mi, s);
  const DAY = 86400000;
  const before = offsetAt(naive - DAY);
  const after = offsetAt(naive + DAY);
  if (before === after) return naive - before;
  const early = naive - before;
  if (offsetAt(early) === before) return early;
  const late = naive - after;
  if (offsetAt(late) === after) return late;
  return early;
}

/** The UTC instant local midnight of `YYYY-MM-DD` names in ZONE. */
function localMidnightMs(day, dayOffset = 0) {
  const [y, mo, d] = day.split('-').map(Number);
  return wallClockMs(y, mo, d + dayOffset);
}

/** A returned `start` with no zone designator, read in ZONE as the server's ordering does. */
function localWallClockMs(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(String(value ?? '').trim());
  if (!m) return NaN;
  return wallClockMs(Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0));
}

// A date-only window covers whole LOCAL days, the end exclusive at the next local midnight.
const windowStart = localMidnightMs(START);
const windowEnd = localMidnightMs(END, 1);

/** Split the tool's text response into its summary line and the JSON array beneath it. */
function parseResponse(body) {
  const newline = body.indexOf('\n');
  if (newline === -1) return { summary: body, events: null };
  return { summary: body.slice(0, newline), events: jsonOf(body) };
}

const client = createClient({ env: process.env });
try {
  await client.init();

  const result = await client.call('list_calendar_events', {
    startDate: START,
    endDate: END,
    limit: 50,
  });

  const { summary, events } = parseResponse(text(result));

  console.log(`\nWindow: ${START} .. ${END} (exclusive end ${new Date(windowEnd).toISOString()})`);
  console.log(`Summary line: ${summary}\n`);

  check('the response carries a JSON array of events', Array.isArray(events));
  check(
    'the response opens with a summary line stating the total',
    /^Showing \d+ of \d+ results\.?/.test(summary),
    summary,
  );
  // nextPosition would be an instruction the caller cannot follow: this tool declares no
  // `position` parameter, so passing one back is rejected by the unknown-parameter guard.
  check('the summary offers no nextPosition on this unpaged tool', !summary.includes('nextPosition'));

  if (!Array.isArray(events)) {
    console.log('\nNo event array to inspect; stopping.');
  } else {
    console.log(`Events returned: ${events.length}`);
    for (const e of events) {
      const marker = e.recurrenceId
        ? `occurrence ${e.recurrenceId}`
        : e.recurrenceRule
          ? `MASTER rrule=${e.recurrenceRule}`
          : 'one-off';
      console.log(`  ${String(e.start).padEnd(22)} ${marker.padEnd(34)} ${e.title}`);
    }

    const outOfWindow = events.filter(e => {
      if (!e.start) return false;
      // An all-day start is placed at local midnight, as the window is; UTC midnight would
      // drift by the zone offset.
      if (!/\d{2}:\d{2}/.test(e.start)) {
        const t = localMidnightMs(e.start);
        return !Number.isNaN(t) && (t < windowStart || t >= windowEnd);
      }
      const iso = /(Z|[+-]\d{2}:?\d{2})$/.test(e.start) ? e.start : `${e.start}Z`;
      const t = Date.parse(iso);
      return !Number.isNaN(t) && (t < windowStart || t >= windowEnd);
    });
    check(
      'every returned start falls inside the requested window',
      outOfWindow.length === 0,
      outOfWindow.length ? `out of window: ${outOfWindow.map(e => `${e.title}@${e.start}`).join(', ')}` : '',
    );

    // The half that a window filter alone would not catch.
    const recurring = events.filter(e => e.isRecurring);
    check('at least one recurring event was returned', recurring.length > 0, `recurring: ${recurring.length}`);
    check(
      'every recurring entry is an expanded occurrence, not a series master',
      recurring.every(e => e.recurrenceId && !e.recurrenceRule),
    );
    // Every occurrence of a series carries the SERIES id, which is what makes update and
    // delete act on all of them. Reported as a fact of the shape rather than asserted as a
    // defect: it is what the tool descriptions now warn about.
    const ids = new Set(recurring.map(e => e.id));
    console.log(`  ${recurring.length} recurring row(s) across ${ids.size} distinct id(s) — an id names the series, not the occurrence.`);

    const summarised = Number(summary.match(/of (\d+) results/)?.[1]);
    check(
      'the stated total is consistent with the number of events returned',
      Number.isFinite(summarised) && summarised >= events.length,
      `total=${summarised} returned=${events.length}`,
    );
    // The order is the thing `limit` slices, and the spellings it has to compare are mixed:
    // a zone-stripped wall clock beside a UTC value sorts by the wrong key as text.
    const ordered = events
      .map(e => (/(Z|[+-]\d{2}:?\d{2})$/.test(String(e.start)) ? Date.parse(e.start) : localWallClockMs(e.start)))
      .filter(t => Number.isFinite(t));
    check(
      'the returned events are in ascending order of their resolved instants',
      ordered.every((t, i) => i === 0 || ordered[i - 1] <= t),
      ordered.length ? `first=${new Date(ordered[0]).toISOString()}` : 'nothing to order',
    );

    // A window containing the series' OWN start date, which the window above does not hold.
    // get_calendar_event on a recurring entry returns the SERIES MASTER, whose DTSTART is the
    // date to build it from.
    console.log('\n--- a window containing the series\' own start date ---');
    const seed = recurring[0];
    if (!seed) {
      console.log('  SKIP  no recurring event in the first window, so there was no series to follow.');
    } else {
      const master = jsonOf(text(await client.call('get_calendar_event', { eventId: seed.id })));
      const rule = master.recurrenceRule || '';
      // Sub-daily frequencies are skipped, not sized: expanding a FREQ=MINUTELY series over
      // any useful span is the work the one-sided window clamp exists to prevent, and this
      // probe has to stay safe to run unattended.
      const spanDays = { YEARLY: 5 * 366, MONTHLY: 5 * 31, WEEKLY: 5 * 7, DAILY: 5 }[/FREQ=([A-Z]+)/.exec(rule)?.[1]];
      const masterDay = String(master.start || '').slice(0, 10);
      if (!spanDays || !/^\d{4}-\d{2}-\d{2}$/.test(masterDay)) {
        console.log(`  SKIP  master start "${master.start}" / rule "${rule}" gives no safe window to expand.`);
      } else {
        const seriesEnd = new Date(Date.parse(`${masterDay}T00:00:00Z`) + spanDays * 86400000)
          .toISOString().slice(0, 10);
        console.log(`  master: ${master.start}  ${rule}`);
        console.log(`  window: ${masterDay} .. ${seriesEnd}\n`);

        const seriesText = text(await client.call('list_calendar_events', {
          startDate: masterDay,
          endDate: seriesEnd,
          limit: 100,
        }));
        const mine = (parseResponse(seriesText).events || []).filter(e => e.id === master.id);
        for (const e of mine) {
          console.log(`    ${String(e.start).padEnd(22)} ${e.recurrenceId ? `occurrence ${e.recurrenceId}` : '(no recurrenceId)'} isRecurring=${!!e.isRecurring}`);
        }
        check('every occurrence in the window is reported, not just the first', mine.length > 1, `got ${mine.length}`);
        const first = mine.find(e => String(e.start).slice(0, 10) === masterDay);
        check("the series' own start date is among them", !!first);
        // And it is not reported as a one-off: its siblings prove the series even though the
        // server left that block without a RECURRENCE-ID.
        check('the first instance is marked recurring, not reported as a one-off', !!first && first.isRecurring === true);
      }
    }
  }

  // A date-only single-day window covers the caller's LOCAL day. Asserted by equivalence, so it
  // holds whatever the account has that day: one date must return exactly what that local
  // day's instants return. Under the old UTC-day reading, a +10:00 account's `2026-08-12`
  // reported one of three appointments that day.
  console.log(`\n--- a date-only single-day window, in ${ZONE} ---`);
  const localStartIso = new Date(localMidnightMs(SINGLE_DAY)).toISOString().replace(/\.\d{3}Z$/, 'Z');
  const localEndIso = new Date(localMidnightMs(SINGLE_DAY, 1)).toISOString().replace(/\.\d{3}Z$/, 'Z');
  console.log(`  ${SINGLE_DAY} is ${localStartIso} .. ${localEndIso}`);

  const titlesOf = async (args) => {
    const body = text(await client.call('list_calendar_events', { limit: 100, ...args }));
    return (parseResponse(body).events || []).map(e => `${e.start} ${e.title}`).sort();
  };

  const byDate = await titlesOf({ startDate: SINGLE_DAY, endDate: SINGLE_DAY });
  const byInstant = await titlesOf({ startDate: localStartIso, endDate: localEndIso });
  for (const row of byDate) console.log(`    ${row}`);
  check(
    'a date-only single-day window returns exactly the caller\'s local day',
    JSON.stringify(byDate) === JSON.stringify(byInstant),
    `byDate=${byDate.length} byInstant=${byInstant.length}`,
  );
  // The same query read as a UTC day, to show the two are genuinely different questions
  // wherever the account is not on UTC. Reported, not asserted: on a UTC deployment they
  // legitimately coincide.
  // Both bounds are instants here, so the exclusive end is the NEXT UTC midnight — passing
  // the same instant twice would be a zero-length window and is rejected as one.
  const nextUtcDay = new Date(Date.parse(`${SINGLE_DAY}T00:00:00Z`) + 86400000).toISOString().slice(0, 10);
  const byUtcDay = await titlesOf({ startDate: `${SINGLE_DAY}T00:00:00Z`, endDate: `${nextUtcDay}T00:00:00Z` });
  console.log(`  the UTC-day reading of the same date would return ${byUtcDay.length} (local: ${byDate.length}).`);
  if (byDate.length === 0) {
    console.log(`  NOTE: nothing on ${SINGLE_DAY}; the equivalence above held trivially.`);
    console.log('        Re-run naming a busy date to exercise it: calendar-window.probe.mjs <start> <end> <day>.');
  }

  console.log('\n--- a window with only one bound ---');
  const oneSided = await client.call('list_calendar_events', { startDate: START, limit: 5 });
  const oneSidedText = text(oneSided);
  const clampNote = oneSidedText.split('\n').find(l => l.startsWith('Note:')) || '';
  console.log(`  ${clampNote || '(no Note line)'}`);
  check('a one-sided window is disclosed in a trailing Note line', !!clampNote);
  check(
    'the note names the range actually searched, so the narrowing is not silent',
    /bounded to \d+ days/.test(clampNote) && /\.\. \d{4}-/.test(clampNote),
  );

  console.log('\n--- a window with no bounds at all ---');
  const noBounds = await client.call('list_calendar_events', { limit: 5 });
  const noBoundsText = text(noBounds);
  const defaultNote = noBoundsText.split('\n').find(l => l.startsWith('Note:')) || '';
  console.log(`  ${defaultNote || '(no Note line)'}`);
  check('a bounds-free call is disclosed in a trailing Note line', !!defaultNote);
  check(
    'the note says neither bound was given, rather than blaming one the caller never passed',
    /no startDate or endDate/.test(defaultNote),
  );
  check(
    'the note names the invented span and the range actually searched',
    /bounded to \d+ days/.test(defaultNote) && /\.\. \d{4}-/.test(defaultNote),
  );
} finally {
  client.close();
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
