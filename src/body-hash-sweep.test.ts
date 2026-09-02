/**
 * Characterisation sweep over `resolveDraftBodyHash`: every draft body shape this module
 * distinguishes, against every combination of read options, tallied by outcome.
 *
 * WHAT THIS IS FOR. `resolveDraftBodyHash` is a cascade of five refusals over a body that
 * can be malformed in many independent ways at once, and its promise is that a draft read
 * carries either a hash or a reason — never silence, and never the wrong reason. The
 * individual cases are pinned by name in `body-hash.test.ts`; this file pins the SHAPE OF
 * THE WHOLE FUNCTION, so a change that fixes one branch by quietly capturing traffic from
 * another shows up as a moved count rather than passing unnoticed.
 *
 * RECONSTRUCTION. An earlier sweep of this function, reported as 260 combinations, was not
 * preserved and cannot be reproduced. This enumeration is built fresh rather than restored,
 * and the counts below are its own honest figures, measured by running it — 4,922 draft
 * shapes against 8 read-option combinations, 39,376 evaluations. They are not the earlier
 * number and are not reconciled to it.
 *
 * WHY A TALLY IS ASSERTED RATHER THAN DESCRIBED. A count written into a comment is prose,
 * and prose cannot fail. Asserted, the same count is checked on every `npm test`, so the
 * rule "re-run this when the function changes" costs nobody anything to remember. The
 * counts are not sacred: a deliberate change to the function is expected to move them, and
 * updating them is part of making that change — but it has to be done knowingly, which is
 * the whole point.
 */
import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { resolveDraftBodyHash } from './body-hash.js';

// ---------------------------------------------------------------------------
// The enumeration
// ---------------------------------------------------------------------------

// One part per axis the function branches on: both text types, a parameterised and an
// upper-case spelling (which must normalise), every way a `type` can be absent (missing,
// empty, null, not a string), an uncarriable text type, media with and without a blobId, a
// part the server flagged as truncated, a part with no stored value at all, a part keyed by
// blobId rather than partId, and two junk entries a real response could contain.
const P = {
  plain: { partId: 'a', type: 'text/plain' },
  plain2: { partId: 'b', type: 'text/plain' },
  html: { partId: 'c', type: 'text/html' },
  html2: { partId: 'd', type: 'text/html' },
  param: { partId: 'e', type: 'text/plain; charset=utf-8' },
  upper: { partId: 'f', type: 'TEXT/HTML' },
  noType: { partId: 'g' },
  emptyType: { partId: 'h', type: '' },
  nullType: { partId: 'i', type: null },
  numType: { partId: 'j', type: 7 },
  cal: { partId: 'k', type: 'text/calendar' },
  image: { partId: 'l', type: 'image/png', blobId: 'blob-1' },
  imageNoBlob: { partId: 'm', type: 'image/png' },
  degraded: { partId: 'n', type: 'text/plain' },
  valueless: { partId: 'o', type: 'text/plain' },
  blobOnly: { blobId: 'blob-2', type: 'text/plain' },
  empty: {},
  nul: null,
} as const;

// 'o' is deliberately absent: a part the draft lists but stores no value for.
const BODY_VALUES: Record<string, any> = {
  a: { value: 'A' }, b: { value: 'B' }, c: { value: '<p>C</p>' }, d: { value: '<p>D</p>' },
  e: { value: 'E' }, f: { value: '<p>F</p>' }, g: { value: 'G' }, h: { value: 'H' },
  i: { value: 'I' }, j: { value: 'J' }, k: { value: 'BEGIN:VCALENDAR' },
  l: { value: '' }, m: { value: '' }, n: { value: 'N', isTruncated: true },
};

const KEYS = Object.keys(P) as (keyof typeof P)[];

// Body-list shapes: the absent and malformed forms a response can carry, every singleton,
// and every ORDERED pair from the pool — order matters, because the whole #179 defect was a
// lookup that took the first part it matched out of a list.
const LISTS: any[] = [undefined, null, [], 'not-an-array'];
for (const k of KEYS) LISTS.push([P[k]]);
for (const k1 of KEYS) for (const k2 of KEYS) if (k1 !== k2) LISTS.push([P[k1], P[k2]]);

// The full LISTS x LISTS cross product is mostly redundant, so each list shape is paired
// against a spread of the other side instead — including AGAINST ITSELF, which is the
// RFC 8621 §4.1.4 aliasing case a single-format draft actually produces.
const OTHERS = [undefined, null, [], [P.plain], [P.html], [P.noType], [P.plain, P.html]];

type Row = { text: any; html: any };
const DRAFTS: Row[] = [];
for (const list of LISTS) {
  for (const other of OTHERS) DRAFTS.push({ text: list, html: other });
  for (const other of OTHERS) DRAFTS.push({ text: other, html: list });
  DRAFTS.push({ text: list, html: list });
}
// The two shapes this sweep exists to keep covered, added explicitly so they survive any
// future trimming of the pool above (see the REQUIRED assertions).
DRAFTS.push({ text: [P.plain, P.plain2], html: null });
DRAFTS.push({ text: [P.plain, P.noType], html: null });

const READS = [false, true].flatMap((bodyText) =>
  [false, true].flatMap((bodyHtml) =>
    [false, true].map((stripQuoted) => ({ bodyText, bodyHtml, stripQuoted }))));

// ---------------------------------------------------------------------------
// Classifying an outcome
// ---------------------------------------------------------------------------

// The outcome type carries no reason CODE — a withheld read is `{ bodyHashWithheld }` and
// the reason is the message itself — so a reason is keyed by its message IN FULL, matched
// exactly. Not a substring: two reasons that happened to share a phrase would merge into
// one bucket and the tally would stay green through a real regression. The cost of the
// exact match is that rewording a message fails this test, which is correct — the wording
// is what the tool promises its caller, and changing it is a change worth being told about.
const REASONS: Record<string, string> = {
  degraded:
    'the server flagged part of this draft\'s stored body as truncated or as having '
    + 'encoding problems, so this read did not return it whole and no read can prove you '
    + 'saw it. Recreate the draft rather than editing its body.',
  unreadable:
    'this draft carries a body part no read returns (a part whose declared type does not '
    + 'match the body list it sits in), so no read can prove you saw the whole body. '
    + 'Recreate the draft rather than editing its body.',
  'uneditable:text/plain':
    'this draft\'s body interleaves multiple text/plain parts, a layout editing cannot '
    + 'preserve, so every edit of this draft is refused and a bodyHash could never be '
    + 'spent. Recreate the draft rather than editing its body (see issue #85).',
  'uneditable:text/html':
    'this draft\'s body interleaves multiple text/html parts, a layout editing cannot '
    + 'preserve, so every edit of this draft is refused and a bodyHash could never be '
    + 'spent. Recreate the draft rather than editing its body (see issue #85).',
  stripQuoted:
    'this read stripped quoted history out of bodyText, so it does not show the body as '
    + 'stored. Read the draft again without stripQuoted to get a bodyHash.',
  'projection:both':
    'this read did not return the draft\'s stored body whole, so it cannot prove you saw '
    + 'it. Read the draft with fields: ["bodyText", "bodyHtml", "bodyHash"] (or '
    + 'verbose:true) to get a bodyHash.',
  'projection:html':
    'this read did not return the draft\'s stored body whole, so it cannot prove you saw '
    + 'it. Read the draft with fields: ["bodyHtml", "bodyHash"] (or verbose:true) to get a '
    + 'bodyHash.',
  'projection:text':
    'this read did not return the draft\'s stored body whole, so it cannot prove you saw '
    + 'it. Read the draft with fields: ["bodyText", "bodyHash"] to get a bodyHash.',
};
const REASON_NAME = new Map(Object.entries(REASONS).map(([name, text]) => [text, name]));

function classify(out: any): string {
  if (out === undefined) return 'not-a-draft';
  if ('bodyHash' in out) return 'hash';
  return REASON_NAME.get(out.bodyHashWithheld) ?? `UNRECOGNISED REASON: ${out.bodyHashWithheld}`;
}

function tally(): Map<string, number> {
  const counts = new Map<string, number>();
  for (const row of DRAFTS) {
    const email = {
      keywords: { $draft: true },
      textBody: row.text, htmlBody: row.html, bodyValues: BODY_VALUES,
    };
    for (const read of READS) {
      const key = classify(resolveDraftBodyHash(email, read));
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
  }
  return counts;
}

// ---------------------------------------------------------------------------

describe('resolveDraftBodyHash sweep', () => {
  it('enumerates the shapes this sweep exists to cover', () => {
    const has = (pred: (r: Row) => boolean) => DRAFTS.some(pred);
    // Two parts of one text type in one list — the layout every edit is refused over, and
    // so the layout no bodyHash can be spent on (#180).
    assert.ok(has((r) => Array.isArray(r.text) && r.text[0] === P.plain && r.text[1] === P.plain2));
    // A part declaring no type beside a typed one. The interleaved check does not count a
    // typeless part, so this is NOT that layout and is not refused as one — the covered
    // shape is the pair the two rules could be made to disagree about (#179).
    assert.ok(has((r) => Array.isArray(r.text) && r.text[0] === P.plain && r.text[1] === P.noType));
    // A part declaring no type on its own, which is an ordinary body and gets a hash.
    assert.ok(has((r) => Array.isArray(r.text) && r.text.length === 1 && r.text[0] === P.noType));
    // A single part aliased into BOTH lists: the shape a real one-format draft has.
    assert.ok(has((r) => Array.isArray(r.text) && r.text === r.html));
  });

  // An enumeration that silently built nothing would tally nothing and read green forever,
  // which no per-shape check above would catch. Its size is asserted so the counts below
  // are known to have been measured over something.
  it('has an enumeration of the size the counts below were measured over', () => {
    assert.equal(LISTS.length, 328); // 4 absent/junk forms, 18 singletons, 18x17 ordered pairs
    assert.equal(DRAFTS.length, 4922);
    assert.equal(READS.length, 8);
  });

  // Frozen from a run of this enumeration, after checking each figure against the branch it
  // belongs to: the three read-independent refusals (degraded, unreadable, uneditable) are
  // each an exact multiple of 8, the stripQuoted refusal is a multiple of 4 because it can
  // only fire on the four reads that set the flag, and the hash and the three projection
  // reasons divide the remaining 6,668 evaluations of those same drafts on their other four
  // reads. Every bucket is reachable and none is empty.
  it('resolves the whole enumeration into the expected outcomes', () => {
    const counts = tally();
    assert.deepEqual(Object.fromEntries([...counts].sort()), {
      'hash': 3548,
      'degraded': 4200,
      'unreadable': 20552,
      'uneditable:text/html': 128,
      'uneditable:text/plain': 1160,
      'stripQuoted': 6668,
      'projection:both': 1276,
      'projection:html': 908,
      'projection:text': 936,
    });
  });

  it('accounts for every evaluation, and issues an outcome for all of them', () => {
    const counts = tally();
    const total = [...counts.values()].reduce((a, b) => a + b, 0);
    assert.equal(total, DRAFTS.length * READS.length);
    assert.equal(total, 39376);
    // Never silence: every draft read is a hash or a reason, and every reason is one this
    // file knows by name.
    assert.equal(counts.get('not-a-draft'), undefined);
    assert.deepEqual([...counts.keys()].filter((k) => k.startsWith('UNRECOGNISED')), []);
  });
});
