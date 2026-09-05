// What this probe settles
// -----------------------
// `findCalendarObjectByUID` finds an event by scanning every object in every calendar and
// comparing UIDs locally (#137). The fix proposed for #137 would replace that with one CalDAV
// `calendar-query` per selectable calendar, filtered on the UID itself with `match-type="equals"`,
// and no fallback to a full scan. That rests on one platform fact no amount of reading settles:
//
//   does Fastmail's CalDAV server accept a `prop-filter` on UID with a
//   `text-match match-type="equals"`, and does it HONOR `equals` rather than
//   quietly applying CalDAV's default `contains` semantics?
//
// The two halves matter for different reasons. Acceptance decides whether the targeted query
// works at all: a server that rejects the filter, or that ignores `prop-filter` and returns
// the whole collection, leaves the fix with nothing to stand on. Honoring `equals` decides
// whether the answer can be TRUSTED: a server matching more loosely than asked returns every
// resource whose UID merely CONTAINS the one asked for, which reads back as several "copies"
// of the event. Two destructive tools rest on that ambiguity count, so a loose match does not
// merely degrade the lookup — it manufactures duplicates that were never there. That is why
// the substring query below, which must return nothing, is the load-bearing check here rather
// than a nicety.
//
// The probe is READ-ONLY: it creates, updates and deletes nothing. It queries for an event
// that already exists in the account, so it needs no fixture, and a calendar probe that
// writes is not a safe default (see the README — an event with participants makes the server
// send real iTIP invitations).
//
// OUTPUT IS COUNTS AND PASS/FAIL ONLY. No collection URL, UID, title or any other value read
// from the account is ever printed, so a run can be quoted verbatim into a public issue or
// commit message. Anything that could carry such a value (a server error string) is redacted
// before it reaches the log.
//
// Run: python scripts/probes/run-probe.py calendar-uid-query.probe.mjs

import { DAVClient } from 'tsdav';
// The shared PASS/FAIL harness — check(label, ok, extra), label first. Taken from probelib so
// the calendar probes cannot disagree about the argument order.
import { makeChecker } from './probelib.mjs';

const USERNAME = process.env.FASTMAIL_CALDAV_USERNAME;
const PASSWORD = process.env.FASTMAIL_CALDAV_PASSWORD;

if (!USERNAME || !PASSWORD) {
  console.error('FAIL  FASTMAIL_CALDAV_USERNAME / FASTMAIL_CALDAV_PASSWORD not set.');
  console.error('      Run through: python scripts/probes/run-probe.py calendar-uid-query.probe.mjs');
  process.exit(1);
}

// Appended to a real UID to build one that cannot exist. Kept to characters that need no XML
// or iCalendar escaping, so the check measures the server's matching rather than our
// serialisation.
const IMPOSSIBLE_SUFFIX = 'zzz-no-such-uid-137';

const { check, failures } = makeChecker();

// Every account-derived string that must never reach the log, filled in as they are read.
// `redact` is applied to anything whose content this probe does not control — server error
// text in particular, which echoes the request back on some failures.
const secrets = [];
// Beyond the exact-substring pass over `secrets` (absolute calendar urls/displayName, the uid,
// the substring, the impossible uid, and every discovered href/path pushed below), a server
// error can echo the account username on its own, or a path-form url the substring pass would
// miss because only the ABSOLUTE form was pushed. Both extra passes match the sibling probes'
// own redaction (calendar-window-frames.probe.mjs, client-authored-events.probe.mjs).
const redact = s => {
  const substringPass = secrets
    .filter(Boolean)
    .reduce((acc, v) => acc.split(v).join('<redacted>'), String(s ?? ''));
  return substringPass
    .split(USERNAME).join('<account>')
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, '<email>');
};

/** Unfold RFC 5545 continuation lines, so a wrapped UID reads as one value. */
const unfold = data => String(data ?? '').replace(/\r?\n[ \t]/g, '');

/** The UID of every VEVENT inside one calendar-data blob (a resource may hold overrides). */
function uidsIn(data) {
  const text = unfold(data);
  const blocks = text.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];
  return blocks
    .map(b => b.match(/^UID:(.*)$/m)?.[1]?.trim())
    .filter(v => v !== undefined && v !== '');
}

// tsdav forwards `filters` verbatim into the calendar-query REPORT body (xml-js compact form:
// `_attributes` for attributes, `_text` for character data; unprefixed element names take the
// `c:` CalDAV namespace). These two shapes are therefore exactly the XML the server sees.
const VEVENT_ONLY = [{
  'comp-filter': {
    _attributes: { name: 'VCALENDAR' },
    'comp-filter': { _attributes: { name: 'VEVENT' } },
  },
}];

const uidEquals = value => [{
  'comp-filter': {
    _attributes: { name: 'VCALENDAR' },
    'comp-filter': {
      _attributes: { name: 'VEVENT' },
      'prop-filter': {
        _attributes: { name: 'UID' },
        // No `collation`, so the server's RFC 4791 default (i;ascii-casemap) applies — which is
        // ASCII case-INSENSITIVE, so `equals` under it may still match a case-variant UID. Step 5
        // below measures whether that reaches equals matching; it is not a gate here, because the
        // narrowing this probe exists to confirm is match-type itself (CalDAV's own text-match
        // has contains semantics by default, and `equals` is the attribute under test).
        'text-match': { _attributes: { 'match-type': 'equals' }, _text: value },
      },
    },
  },
}];

const client = new DAVClient({
  serverUrl: 'https://caldav.fastmail.com/dav/',
  credentials: { username: USERNAME, password: PASSWORD },
  authMethod: 'Basic',
  defaultAccountType: 'caldav',
});

/**
 * Run one UID-filtered `calendar-query` and report only its resource count. Returns the
 * matched hrefs, or null when the server refused the filter — which is itself a FAIL, since
 * the fix proposed for #137 has no fallback path to take.
 *
 * Uses `calendarQuery` alone, NOT `fetchCalendarObjects({ filters })`: under tsdav that helper
 * issues TWO requests — a filtered calendar-query for etags, then an unfiltered multiget by
 * href — so its object count measures the multiget succeeding, not the filter. A throw from
 * the multiget half would then misreport as "the server refused the filtered query" when the
 * filter itself was fine. `calendarQuery`'s own row count is what this probe means by "the
 * filtered query's resource count"; fetching the matched resource's body is done separately
 * (step 2 only), with its own distinct failure mode.
 */
async function uidQuery(calendar, value, label) {
  try {
    const rows = await client.calendarQuery({
      url: calendar.url,
      props: { 'd:getetag': {} },
      filters: uidEquals(value),
      depth: '1',
    });
    const hrefs = (rows ?? []).map(r => r.href ?? '').filter(Boolean);
    for (const h of hrefs) secrets.push(h);
    console.log(`  ${label}: resources returned = ${hrefs.length}`);
    return hrefs;
  } catch (err) {
    console.log(`  ${label}: the server REFUSED the filtered query — ${redact(err?.message ?? String(err)).slice(0, 300)}`);
    check(`${label}: the server accepted a UID prop-filter with match-type="equals"`, false);
    return null;
  }
}

console.log('\nUID-targeted calendar-query, match-type="equals" (#137)');
console.log('Counts and PASS/FAIL only — no account value is printed.\n');

// The whole run is wrapped so any unguarded rejection (login, discovery, a multiget) lands as
// one FAIL line — redacted — rather than a raw stack, per the pattern in
// calendar-rdate-expand.probe.mjs. The summary print and exit below always run, whether the
// try completes or is caught.
try {
  await client.login();
  const calendars = await client.fetchCalendars();
  console.log(`Calendars discovered: ${calendars.length}`);
  check('calendar discovery returned at least one calendar', calendars.length > 0);
  for (const cal of calendars) {
    secrets.push(cal.url, cal.displayName, new URL(cal.url, 'https://caldav.fastmail.com/').pathname);
  }

  // ------------------------------------------------------------------------------------
  // Step 1: find an event that already exists. Discovery asks for etags only and then fetches
  // a matching resource by URL, rather than pulling every object in the collection: the probe
  // needs one UID, and a whole-calendar fetch on a real account is a lot of traffic for it.
  // Up to 10 hrefs per calendar are tried before moving on, so one unparseable resource does
  // not skip a calendar that holds a usable one further down the listing.
  // ------------------------------------------------------------------------------------

  let target = null;
  let scanned = 0;

  for (const cal of calendars) {
    scanned++;
    let rows;
    try {
      rows = await client.calendarQuery({
        url: cal.url,
        props: { 'd:getetag': {} },
        filters: VEVENT_ONLY,
        depth: '1',
      });
    } catch {
      // A collection that refuses a VEVENT query (a task-only or contact collection) is not a
      // finding: move on to the next one.
      continue;
    }
    const hrefs = (rows ?? []).map(r => r.href ?? '').filter(h => h.includes('.ics')).slice(0, 10);

    let foundUid = null;
    for (const href of hrefs) {
      secrets.push(href);
      try {
        const [object] = await client.fetchCalendarObjects({ calendar: cal, objectUrls: [href] });
        const uidHere = uidsIn(object?.data)[0];
        if (uidHere) { foundUid = uidHere; break; }
      } catch {
        // This one resource could not be fetched by href — not a finding, try the next href.
        continue;
      }
    }
    if (!foundUid) continue;

    target = { calendar: cal, uid: foundUid };
    break;
  }

  console.log(`Calendars scanned before one held an event: ${scanned}`);
  check('an existing event was found to query for', target !== null);

  if (!target) {
    console.log('\nNo calendar in this account holds a VEVENT, so there is nothing to look up.');
    console.log('This probe does not create one: it is read-only by design. Point it at an');
    console.log('account with at least one event.');
    console.log(`\n${failures()} CHECK(S) FAILED`);
    process.exit(1);
  }

  const { calendar, uid } = target;
  secrets.push(uid);

  const substring = uid.slice(0, -2);
  const impossible = `${uid}${IMPOSSIBLE_SUFFIX}`;
  secrets.push(substring, impossible);

  // ------------------------------------------------------------------------------------
  // Step 2: the exact UID. The happy path the fix proposed for #137 rests on.
  // ------------------------------------------------------------------------------------

  console.log('\n=== Step 2: query the exact UID ===');
  const exactHrefs = await uidQuery(calendar, uid, 'exact UID');
  if (exactHrefs) {
    check('step 2: the query returned at least one resource', exactHrefs.length >= 1, `count=${exactHrefs.length}`);
    if (exactHrefs.length >= 1) {
      let matched = null;
      try {
        matched = await client.fetchCalendarObjects({ calendar, objectUrls: exactHrefs });
      } catch (err) {
        // The filtered query above already succeeded — a throw here is the matched resource
        // failing to fetch, a distinct failure from a refused filter, and must read as one.
        check('step 2: the matched resource could not be fetched', false, redact(err?.message ?? String(err)).slice(0, 300));
      }
      if (matched) {
        check(
          'step 2: every returned resource parses to the target UID exactly',
          matched.length > 0 && matched.every(o => {
            const uids = uidsIn(o.data);
            return uids.length > 0 && uids.every(u => u === uid);
          }),
        );
        check(
          'step 2: every returned resource carries non-empty data AND non-empty etag',
          matched.length > 0 && matched.every(o => String(o.data ?? '') !== '' && String(o.etag ?? '') !== ''),
        );
      }
    }
  }

  // ------------------------------------------------------------------------------------
  // Step 3: a STRICT substring of that UID. THE LOAD-BEARING CHECK. A server treating
  // match-type="equals" as CalDAV's default `contains` returns the target here; anything that
  // comes back is a resource the caller did not ask for, and counting them is how a targeted
  // lookup manufactures the ambiguity two destructive tools would refuse on.
  // ------------------------------------------------------------------------------------

  console.log('\n=== Step 3: query a strict substring of that UID ===');
  const substringUsable = substring.length > 0 && substring !== uid;
  check('a strict, non-empty substring of the UID could be built', substringUsable);
  if (substringUsable) {
    const loose = await uidQuery(calendar, substring, 'strict substring');
    check(
      'step 3: the substring query returned ZERO resources — match-type="equals" is honored',
      loose !== null && loose.length === 0,
      `count=${loose === null ? 'n/a' : loose.length}`,
    );
  }

  // ------------------------------------------------------------------------------------
  // Step 4: a UID that cannot exist. Guards against a server that ignores the prop-filter and
  // answers every query with the whole collection — which step 2 alone would read as a pass.
  // ------------------------------------------------------------------------------------

  console.log('\n=== Step 4: query a UID that cannot exist ===');
  const absent = await uidQuery(calendar, impossible, 'nonexistent UID');
  check(
    'step 4: the nonexistent-UID query returned ZERO resources',
    absent !== null && absent.length === 0,
    `count=${absent === null ? 'n/a' : absent.length}`,
  );

  // ------------------------------------------------------------------------------------
  // Step 5: a case-variant of the UID. MEASUREMENT, NOT A GATE — the default collation
  // (i;ascii-casemap) is case-insensitive, so either count is a legitimate answer and the
  // consumer's own client-side exact-equality filter handles both. This only records which
  // way the server actually answers, so a refusal of the query itself is the only failure
  // mode (handled by uidQuery's own check, reusing the same label pattern as the other steps).
  // ------------------------------------------------------------------------------------

  console.log('\n=== Step 5: query a case-variant of the UID (measurement, not a gate) ===');
  const upper = uid.toUpperCase();
  if (upper === uid) {
    console.log('  UID has no letters to vary — case-variant step is not applicable.');
  } else {
    secrets.push(upper);
    const caseVariant = await uidQuery(calendar, upper, 'case-variant UID');
    if (caseVariant !== null) {
      const insensitive = caseVariant.length > 0;
      console.log(
        insensitive
          ? `  count=${caseVariant.length} -> equals is ASCII-case-insensitive under the default collation`
          : '  count=0 -> equals is case-sensitive under the default collation',
      );
      check(
        'step 5: the case-variant query completed (measurement, not a gate on the count)',
        true,
        insensitive ? 'case-insensitive' : 'case-sensitive',
      );
    }
  }
} catch (err) {
  check('probe ran to completion', false, redact(err?.message ?? String(err)));
}

console.log(`\n${failures() === 0 ? 'ALL CHECKS PASSED' : `${failures()} CHECK(S) FAILED`}`);
process.exit(failures() === 0 ? 0 : 1);
