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
//
// It is READ-ONLY: it lists and reads, and creates nothing. That is deliberate beyond
// ordinary tidiness — creating an event with participants makes the server send real iTIP
// invitations, and deleting one sends cancellations (see the README), so a calendar probe
// that writes is not a safe default.
//
// Run: python scripts/probes/run-probe.py calendar-window.probe.mjs [startDate endDate]
// Requires FASTMAIL_API_TOKEN plus FASTMAIL_CALDAV_USERNAME/PASSWORD; the launcher injects
// all three from the local MCP client config. Build first: npm run build.

import { createClient } from '../mcp-harness.mjs';

const START = process.argv[2] || '2027-03-01';
const END = process.argv[3] || '2027-03-10';

if (!process.env.FASTMAIL_CALDAV_USERNAME || !process.env.FASTMAIL_CALDAV_PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-window.probe.mjs');
  process.exit(1);
}

let failures = 0;
const check = (ok, label, detail) => {
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}${detail ? ` - ${detail}` : ''}`);
  if (!ok) failures++;
};

// A date-only endDate covers the whole of that day, so the exclusive bound is the following
// midnight. The probe mirrors the server's rule rather than restating a hard-coded instant,
// so a change to that rule shows up as a failing assertion instead of a silently looser test.
const windowStart = Date.parse(`${START}T00:00:00Z`);
const windowEnd = Date.parse(`${END}T00:00:00Z`) + 24 * 60 * 60 * 1000;

/** Split the tool's text response into its summary line and the JSON array beneath it. */
function parseResponse(text) {
  const newline = text.indexOf('\n');
  if (newline === -1) return { summary: text, events: null };
  return { summary: text.slice(0, newline), events: JSON.parse(text.slice(newline + 1)) };
}

const client = createClient({ env: process.env });
try {
  await client.init();

  const result = await client.call('list_calendar_events', {
    startDate: START,
    endDate: END,
    limit: 50,
  });

  const text = result?.content?.[0]?.text ?? '';
  const { summary, events } = parseResponse(text);

  console.log(`\nWindow: ${START} .. ${END} (exclusive end ${new Date(windowEnd).toISOString()})`);
  console.log(`Summary line: ${summary}\n`);

  check(Array.isArray(events), 'the response carries a JSON array of events');
  check(
    /^Showing \d+ of \d+ results\.?/.test(summary),
    'the response opens with a summary line stating the total',
    summary,
  );
  // nextPosition would be an instruction the caller cannot follow: this tool declares no
  // `position` parameter, so passing one back is rejected by the unknown-parameter guard.
  check(!summary.includes('nextPosition'), 'the summary offers no nextPosition on this unpaged tool');

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
      const iso = /\d{2}:\d{2}/.test(e.start)
        ? (/(Z|[+-]\d{2}:?\d{2})$/.test(e.start) ? e.start : `${e.start}Z`)
        : `${e.start}T00:00:00Z`;
      const t = Date.parse(iso);
      return !Number.isNaN(t) && (t < windowStart || t >= windowEnd);
    });
    check(
      outOfWindow.length === 0,
      'every returned start falls inside the requested window',
      outOfWindow.length ? `out of window: ${outOfWindow.map(e => `${e.title}@${e.start}`).join(', ')}` : undefined,
    );

    // The half that a window filter alone would not catch. Before the fix a recurring series
    // reported its original DTSTART, so an in-window date and a recurring event were mutually
    // exclusive outcomes rather than the same one.
    const recurring = events.filter(e => e.isRecurring);
    check(recurring.length > 0, 'at least one recurring event was returned', `recurring: ${recurring.length}`);
    check(
      recurring.every(e => e.recurrenceId && !e.recurrenceRule),
      'every recurring entry is an expanded occurrence, not a series master',
    );

    const summarised = Number(summary.match(/of (\d+) results/)?.[1]);
    check(
      Number.isFinite(summarised) && summarised >= events.length,
      'the stated total is consistent with the number of events returned',
      `total=${summarised} returned=${events.length}`,
    );
  }
} finally {
  client.close();
}

console.log(`\n${failures === 0 ? 'ALL CHECKS PASSED' : `${failures} CHECK(S) FAILED`}`);
process.exit(failures === 0 ? 0 : 1);
