// What this probe settles
// -----------------------
// `calendar-expand.probe.mjs` measured the PLATFORM: what Fastmail's CalDAV server returns
// for an expanded time-range query. It talks to tsdav directly and says nothing about what
// this server does with the answer. This probe settles the other half:
//
//   does the shipped `list_calendar_events` tool now answer a window question with dates
//   that are actually in that window?
//
// It matters because the two halves of #64 failed at different layers and only the
// end-to-end path proves both are closed at once. Asking for 1-10 March 2027 used to return
// a single event dated 2020-08-28: the server had correctly matched a recurring series on an
// occurrence inside the window, but nothing asked it to expand the recurrence, and the
// parser read only the first VEVENT of the resource and reported the master's DTSTART. Either
// defect alone reproduces the same symptom, so a fix has to be verified through the whole
// stack rather than at the parser.
//
// This spawns the BUILT server (dist/index.js) through the shared MCP harness and asserts on
// the tool's real response:
//   - every returned `start` falls inside the requested window
//   - a recurring entry is marked as an occurrence (`recurrenceId`, no `RRULE`) rather than
//     as a series master, which is what makes the date unambiguous (#64 request 3)
//   - the response opens with a summary line stating the total, so a `limit` that trimmed
//     the list is visible instead of being inferred from the array length (#100)
//   - a window containing a series' OWN start date reports EVERY occurrence in it, not one.
//     That case is not covered by the window above and it is where the read path broke:
//     Fastmail marks only instances after the first with a RECURRENCE-ID, so a window that
//     includes a series' original DTSTART returns one bare block plus its siblings, and a
//     parser identifying the bare block as a series master discarded all of them
//   - the rows come back in ascending order of the INSTANT each start names, resolved the way
//     the server resolves them. A real calendar mixes zone-stripped wall clocks with UTC
//     values, and `limit` slices this list, so an order computed from the spellings drops a
//     genuinely earlier event and keeps a later one
//   - a window given only ONE bound is bounded rather than run open-ended, and the response
//     says so — a caller handed a narrower window than it asked for must be told, or
//     "nothing after that date" reads as an empty calendar
//   - a call given NO bounds is bounded the same way, to the next month from today, and the
//     note says neither bound was given rather than blaming one the caller never passed
//     (#142). That call used to go out with no time range and therefore no `expand`, so the
//     one call most likely to be asked "what is on?" answered with series masters
//
// It is READ-ONLY: it lists and reads, and creates nothing. That is deliberate beyond
// ordinary tidiness — creating an event with participants makes the server send real iTIP
// invitations, and deleting one sends cancellations (see the README), so a calendar probe
// that writes is not a safe default.
//
// WHAT IT DOES NOT COVER: a timezone other than the host's, unless one is deliberately supplied.
// `run-probe.py` forwards FASTMAIL_TIMEZONE from the MCP client config, and separately, as with
// every value here, a shell-exported FASTMAIL_TIMEZONE or TZ reaches the child too, since the
// launcher starts from a copy of its own environment before layering the config's FASTMAIL_TIMEZONE
// on top - so when the config and a shell-exported FASTMAIL_TIMEZONE are both set, the config value
// wins, not the other way round. (An exported TZ is different: the launcher never touches it, so
// any precedence between TZ and FASTMAIL_TIMEZONE is the server's own, not the launcher's layering.)
// But the local config sets no timezone today, so unless one of those routes is used, the built
// server reads dates in whatever zone this machine runs in, which makes the "a date-only window
// is the caller's LOCAL day" assertion below hold trivially on a UTC host: local midnight IS UTC
// midnight there, so a broken zone resolution would still pass. The real zone coverage is the unit
// tests, which drive named zones and DST edges directly (src/coerce.test.ts,
// src/caldav-client.test.ts); export FASTMAIL_TIMEZONE=Australia/Sydney before running this (or set
// it in the MCP client config) if you want the live path to exercise a non-zero offset too.
//
// Run: python scripts/probes/run-probe.py calendar-window.probe.mjs [startDate endDate [singleDay]]
// Requires FASTMAIL_API_TOKEN plus FASTMAIL_CALDAV_USERNAME/PASSWORD; the launcher injects
// all three from the local MCP client config. Build first: npm run build.

import { createClient } from '../mcp-harness.mjs';
// The PASS/FAIL harness and the JSON extractor are the shared ones. This file used to carry
// its own copies, and the local `check` took its arguments in the opposite order to the
// shared one — two probes in the same directory, two signatures, one of them silently
// printing a label where an assertion belonged.
import { makeChecker, jsonOf, text } from './probelib.mjs';

const START = process.argv[2] || '2027-03-01';
const END = process.argv[3] || '2027-03-10';
// The single date the local-day check below asks about. Defaults to the date the wrong-day
// window was reported on, so a re-run reproduces the original report rather than a synthetic
// case; override it to point the check at any day whose contents you know.
const SINGLE_DAY = process.argv[4] || '2026-08-12';

if (!process.env.FASTMAIL_CALDAV_USERNAME || !process.env.FASTMAIL_CALDAV_PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-window.probe.mjs');
  process.exit(1);
}

const { check, failures } = makeChecker();

// The zone the server reads a date-only window in: FASTMAIL_TIMEZONE if the launcher injected
// one, otherwise the host zone — the same fallback the server itself takes, so the probe and
// the server agree about which day is being asked about.
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
 * The offsets a day either side bracket the answer, and a candidate is checked against the
 * offset actually in force where it lands: a repeated wall clock takes the earlier instant, a
 * skipped one resolves forward to the transition. Mirroring rather than restating a
 * hard-coded instant is what makes a change to that rule show up as a failing assertion here.
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

/**
 * The UTC instant a returned `start` names when it carries no zone designator — a bare date,
 * or a wall clock whose TZID the read dropped. Read in ZONE, which is what the server's own
 * ordering does with the same value.
 */
function localWallClockMs(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(String(value ?? '').trim());
  if (!m) return NaN;
  return wallClockMs(Number(m[1]), Number(m[2]), Number(m[3]), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0));
}

// A date-only window covers whole LOCAL days: the start is local midnight, and the exclusive
// end is local midnight of the following day. The probe mirrors the server's rule rather than
// restating a hard-coded instant, so a change to that rule shows up as a failing assertion
// instead of a silently looser test.
const windowStart = localMidnightMs(START);
const windowEnd = localMidnightMs(END, 1);

/**
 * Split the tool's text response into its summary line and the JSON array beneath it.
 *
 * The array is extracted by `jsonOf`, which tracks bracket depth and string state, rather
 * than by slicing to the last `]` in the response: a window note rides AFTER the JSON (the
 * same shape the email listings use for their Trash/Spam note), and one bracket in that prose
 * — or in an event title — moved where the slice ended.
 */
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
      // An all-day event's start is a DATE, so it is placed at local midnight — the same
      // reading the window itself now uses. Comparing it as UTC midnight would drift by the
      // zone offset and mislabel an event on either edge of the window.
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

    // The half that a window filter alone would not catch. Before the fix a recurring series
    // reported its original DTSTART, so an in-window date and a recurring event were mutually
    // exclusive outcomes rather than the same one.
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

    // ---------------------------------------------------------------------
    // A window that contains the series' OWN start date.
    // ---------------------------------------------------------------------
    // The window above deliberately holds no series start, which is why it could not see
    // the defect this pass exists for. get_calendar_event on any recurring entry returns
    // the SERIES MASTER — its original DTSTART and its RRULE — which is exactly the date to
    // build the second window from.
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
        // The defect, stated as an assertion: the first instance used to be mistaken for a
        // series master and every sibling in the same resource thrown away, leaving 1.
        check('every occurrence in the window is reported, not just the first', mine.length > 1, `got ${mine.length}`);
        const first = mine.find(e => String(e.start).slice(0, 10) === masterDay);
        check("the series' own start date is among them", !!first);
        // And it is not reported as a one-off: its siblings prove the series even though the
        // server left that block without a RECURRENCE-ID.
        check('the first instance is marked recurring, not reported as a one-off', !!first && first.isRecurring === true);
      }
    }
  }

  // -----------------------------------------------------------------------
  // A date-only single-day window covers the caller's LOCAL day.
  // -----------------------------------------------------------------------
  // Asserted by equivalence rather than against a hard-coded expected count, so it holds
  // whatever the account happens to have on the day: asking for one date must return exactly
  // what asking for that local day's instants returns. Under the UTC-day reading the two
  // disagreed by the account's whole UTC offset — for a +10:00 account, `2026-08-12` searched
  // 12 Aug 10:00 to 13 Aug 10:00 local, so a day with three appointments reported one and the
  // 08:00 one whose own title said "Wednesday 12 Aug 2026" was simply absent.
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

  // -----------------------------------------------------------------------
  // A one-sided window is bounded, and the response says so.
  // -----------------------------------------------------------------------
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

  // -----------------------------------------------------------------------
  // A window with NO bounds is the next month, and the response says so.
  // -----------------------------------------------------------------------
  // This is the call that used to go out with no time range at all, and therefore with no
  // `expand` either — so the one call most likely to be asked "what is on?" answered with
  // series masters at their original DTSTART. It is now bounded like any other invented
  // window (#142). Only the live server can show that the range really is sent and really is
  // expanded; the note is the caller-facing half of the same fact.
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
