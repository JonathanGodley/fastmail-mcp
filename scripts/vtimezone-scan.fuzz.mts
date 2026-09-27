// What this settles
// -----------------
// #166: the VTIMEZONE block scan behind `removeOrphanedVTimezones` and `regenerateVTimezones`
// (src/caldav-client.ts) decides which line ranges of a stored calendar resource are a VTIMEZONE
// and may be deleted. A scan that mis-identifies a block boundary deletes lines that belong to
// the event itself. The unit tests pin the malformed shapes someone thought of; this sweeps the
// ones nobody did.
//
// Invariant: both functions only ever remove VTIMEZONE block ranges, so a line that can never be
// legal inside a VTIMEZONE (`UID:`, `SUMMARY:`) must NEVER disappear from a payload they accept.
// RFC 5545 §3.6.5's timezonec carries tzprop plus standardc/daylightc only, and UID and SUMMARY
// appear in no timezone grammar, so the invariant has no legitimate exception. A refusal
// (InvalidInputError) is an acceptable outcome; any other throw is not.
//
// Each iteration mutates a WELL-FORMED calendar (0-2 random mutations): purely random line
// sequences are almost never accepted, so the side under test would go unexercised. The sentinel
// lines are never mutated or moved, because moving one INTO a timezone block would make its
// removal correct; any disappearance is therefore the scan getting a boundary wrong. The PRNG is
// seeded, so a given iteration count always generates the same payloads. No credentials, no
// network.
//
// Run: npx tsx scripts/vtimezone-scan.fuzz.mts [iterations]   (default 200000)

import { removeOrphanedVTimezones, regenerateVTimezones } from '../src/caldav-client.js';

const DEFAULT_ITERATIONS = 200000;
const arg = process.argv[2];
const N = arg === undefined ? DEFAULT_ITERATIONS : Number(arg);
if (!Number.isInteger(N) || N < 1) {
  console.error(`Usage: npx tsx scripts/vtimezone-scan.fuzz.mts [iterations]  (a positive integer, default ${DEFAULT_ITERATIONS})`);
  process.exit(2);
}

const CRLF = '\r\n';
const MARKERS = ['UID:sentinel@example.com', 'SUMMARY:sentinel subject'];

let seed = 20260920;
const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
const pick = <T>(a: T[]): T => a[rnd(a.length)];

function wellFormed(): string[] {
  const zones = ['Australia/Sydney', 'Europe/London', 'America/New_York'];
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//x//EN'];
  const used: string[] = [];
  for (let b = 0, nb = rnd(3); b <= nb; b++) {
    const z = pick(zones);
    used.push(z);
    lines.push('BEGIN:VTIMEZONE', `TZID:${z}`);
    if (rnd(2)) lines.push('BEGIN:STANDARD', 'DTSTART:20250405T030000', 'TZOFFSETFROM:+1100', 'TZOFFSETTO:+1000', 'END:STANDARD');
    if (rnd(2)) lines.push('BEGIN:DAYLIGHT', 'DTSTART:20251005T020000', 'TZOFFSETFROM:+1000', 'TZOFFSETTO:+1100', 'END:DAYLIGHT');
    lines.push('END:VTIMEZONE');
  }
  const ez = used.length && rnd(2) ? used[rnd(used.length)] : 'Australia/Sydney';
  lines.push('BEGIN:VEVENT', ...MARKERS, 'DTSTAMP:20260301T000000Z',
    `DTSTART;TZID=${ez}:20260320T090000`, `DTEND;TZID=${ez}:20260320T100000`);
  if (rnd(2)) lines.push('BEGIN:VALARM', 'ACTION:DISPLAY', 'TRIGGER:-PT15M', 'DESCRIPTION:d', 'END:VALARM');
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines;
}

const isSentinel = (line: string) => MARKERS.includes(line.trim());
const freeIdx = (l: string[]): number => {
  for (let tries = 0; tries < 20; tries++) { const i = rnd(l.length); if (!isSentinel(l[i])) return i; }
  return -1;
};

const MUTATE: Array<(l: string[]) => void> = [
  l => { const i = freeIdx(l); if (i >= 0) l[i] = l[i].toLowerCase(); },                       // case flip
  l => { const i = freeIdx(l); if (i >= 0 && /^(BEGIN|END):/i.test(l[i])) l.splice(i, 1); },   // drop a boundary
  l => l.splice(rnd(l.length + 1), 0, pick(['BEGIN:VTIMEZONE', 'END:VTIMEZONE', 'BEGIN:VEVENT', 'END:VEVENT', 'BEGIN:STANDARD', 'END:STANDARD', 'BEGIN:X-VENDOR', 'END:X-VENDOR', 'BEGIN:', 'END:'])),
  l => { const i = freeIdx(l); if (i < 0) return; const j = freeIdx(l); if (j < 0) return; const [x] = l.splice(i, 1); l.splice(j, 0, x); }, // move a non-sentinel line
  l => { const i = freeIdx(l); if (i >= 0) l[i] = l[i] + ' '; },                               // trailing space
  l => { const i = freeIdx(l); if (i < 0 || l[i].length < 2) return; const c = 1 + rnd(l[i].length - 1); // fold
         const head = l[i].slice(0, c), tail = l[i].slice(c); l.splice(i, 1, head, (rnd(2) ? ' ' : '\t') + tail); },
];

const counts = { accepted: 0, refused: 0, unexpected: 0, violations: 0 };
const failures: Array<{ p: string; o: string; via: string }> = [];

for (let n = 0; n < N; n++) {
  const lines = wellFormed();
  for (let m = rnd(3); m > 0; m--) pick(MUTATE)(lines);
  const payload = lines.join(rnd(4) === 0 ? '\n' : CRLF);
  const le = payload.includes(CRLF) ? CRLF : '\n';

  for (const [via, fn] of [
    ['sweep', () => removeOrphanedVTimezones(payload)],
    ['regen+sweep', () => removeOrphanedVTimezones(regenerateVTimezones(payload, le))],
  ] as Array<[string, () => string]>) {
    try {
      const out = fn();
      counts.accepted++;
      if (MARKERS.some(mk => !out.includes(mk))) {
        counts.violations++;
        if (failures.length < 3) failures.push({ p: payload, o: out, via });
      }
    } catch (e: any) {
      if (e?.constructor?.name === 'InvalidInputError') counts.refused++;
      else {
        counts.unexpected++;
        if (failures.length < 3) failures.push({ p: payload, o: `UNEXPECTED ${e?.constructor?.name}: ${e?.message}`, via });
      }
    }
  }
}

console.log(`iterations=${N}  accepted=${counts.accepted}  refused=${counts.refused}  unexpected-throw=${counts.unexpected}`);
console.log(`invariant violations (accepted payload that lost a UID: or SUMMARY: line): ${counts.violations}`);
for (const f of failures) console.log(`\n--- via ${f.via} ---\nINPUT:\n${f.p.replace(/\r\n/g, '\n')}\nOUTPUT:\n${f.o.replace(/\r\n/g, '\n')}`);
console.log(counts.violations === 0 && counts.unexpected === 0 ? 'PASS' : 'FAIL');
process.exit(counts.violations === 0 && counts.unexpected === 0 ? 0 : 1);
