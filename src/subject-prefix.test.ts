import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  matchSubjectPrefix,
  noteComposeSubjectPrefix,
  noteEditSubjectPrefix,
} from './subject-prefix.js';

// The matcher is the whole of what #188 calls a prefix, on both routes, so its accept and
// reject sets are ENUMERATED here rather than sampled: every form the description in the
// tool surface promises is listed, and so is every near-miss that must not draw a note. A
// case added to the matcher without a row here is a claim nobody checked.

describe('matchSubjectPrefix — what counts as a reply prefix', () => {
  const REPLY: Array<[string, string]> = [
    ['Re: hi', 'the ordinary form'],
    ['RE: hi', 'all caps'],
    ['re: hi', 'all lower case'],
    ['rE: hi', 'mixed case'],
    ['Re:hi', 'no space after the colon'],
    ['Re:', 'a subject that is the prefix and nothing else'],
    ['Re : hi', 'a space before the colon'],
    ['Re[2]: hi', 'a counter'],
    ['RE[10]: hi', 'a counter above nine'],
    ['re [3] : hi', 'a counter with whitespace on both sides'],
    ['  Re: hi', 'leading whitespace before the prefix'],
  ];

  for (const [subject, why] of REPLY) {
    it(`matches ${JSON.stringify(subject)} — ${why}`, () => {
      assert.equal(matchSubjectPrefix(subject), 'reply');
    });
  }
});

describe('matchSubjectPrefix — what counts as a forward prefix', () => {
  const FORWARD: Array<[string, string]> = [
    ['Fwd: hi', 'the ordinary three-letter form'],
    ['FWD: hi', 'all caps'],
    ['fwd: hi', 'all lower case'],
    ['Fw: hi', 'the two-letter form'],
    ['FW: hi', 'the two-letter form in caps'],
    ['Fwd:hi', 'no space after the colon'],
    ['Fwd : hi', 'a space before the colon'],
    ['Fwd[2]: hi', 'a counter'],
    ['Fw [3] : hi', 'the two-letter form with a spaced counter'],
    ['  Fwd: hi', 'leading whitespace before the prefix'],
  ];

  for (const [subject, why] of FORWARD) {
    it(`matches ${JSON.stringify(subject)} — ${why}`, () => {
      assert.equal(matchSubjectPrefix(subject), 'forward');
    });
  }
});

describe('matchSubjectPrefix — what does not count', () => {
  // Each row is a form that would draw a note if the set were widened by one obvious step:
  // dropping the colon, dropping the anchor, admitting a longer word, or admitting a
  // counter that is not a number. The note is cheap but it is not free — it tells a caller
  // their subject will not thread — so a false hit is the failure mode worth pinning.
  const NEITHER: Array<[string | undefined, string]> = [
    ['Reply: hi', 'a longer word starting with "re"'],
    ['Reference: hi', 'another longer word starting with "re"'],
    ['Rest: today', 'a short word starting with "re"'],
    ['Rex: hi', 'one letter past the prefix, then the colon'],
    ['Forward: hi', 'the spelt-out word, which no client writes as a prefix'],
    ['Fwded: hi', 'a longer word starting with "fwd"'],
    ['Re hi', 'the letters with no colon at all'],
    ['Fwd hi', 'the forward letters with no colon at all'],
    ['Re', 'the letters alone'],
    ['Notes on Re: pricing', 'a prefix that is not leading'],
    ['[EXTERNAL] Re: hi', 'a prefix behind a tag some gateways prepend'],
    ['Re[]: hi', 'an empty counter'],
    ['Re[a]: hi', 'a counter that is not a number'],
    ['Re[x2]: hi', 'a counter with something in front of the number'],
    ['Re[2x]: hi', 'a counter with something after the number'],
    ['AW: hi', 'the German reply prefix, deliberately out of the set'],
    ['', 'an empty subject'],
    ['   ', 'a whitespace-only subject'],
    [undefined, 'no subject at all'],
  ];

  for (const [subject, why] of NEITHER) {
    it(`does not match ${JSON.stringify(subject)} — ${why}`, () => {
      assert.equal(matchSubjectPrefix(subject), undefined);
    });
  }
});

describe('matchSubjectPrefix — a long subject cannot stall the server', () => {
  // A subject is caller-supplied and nothing caps its length before the matcher sees it,
  // on a server that is one stdio process: a pattern that backtracks quadratically here
  // holds up every other tool call. The shape that provokes it is the prefix, a long
  // whitespace run, and then something that is not the colon the pattern needs, so the
  // engine has to give the run back one character at a time.
  //
  // THE BOUND IS DELIBERATELY ENORMOUS. This input takes well under a millisecond with a
  // linear pattern and about 27 SECONDS with the quadratic one, so a second separates the
  // two by three orders of magnitude in one direction and thirty in the other. It fails
  // only for a genuine reintroduction of the backtracking, never for a slow or loaded CI
  // box, and it cannot be made to flake by tightening the machine.
  it('answers a 200,000-character whitespace run in well under a second', () => {
    const pathological = 'Re' + ' '.repeat(200_000) + 'x';

    const started = performance.now();
    const answer = matchSubjectPrefix(pathological);
    const elapsed = performance.now() - started;

    // No colon, so no prefix - asserted as well as timed, so the bound is measured over
    // the real answer rather than over an early exit.
    assert.equal(answer, undefined);
    assert.ok(elapsed < 1000, `matching took ${elapsed.toFixed(1)}ms`);
  });
});

describe('the two sentences', () => {
  // THE EXACT SENTENCES, written out here rather than read back off the functions that
  // build them. Both wiring suites assert with the builder's own output, which cannot
  // notice a fragment going missing - the two sides would change together. These four
  // are what a caller actually reads, and the tool descriptions promise they exist, so
  // they are pinned whole, once.
  it('says exactly this on a compose', () => {
    assert.equal(
      noteComposeSubjectPrefix('reply'),
      "This subject reads as a reply, but mode:'new' writes no In-Reply-To or References, "
        + "so the message starts its own conversation rather than joining the one it names. "
        + "Compose with mode:'reply' and originalEmailId to thread it — passing to as well if "
        + "you do not want reply-all."
    );
    assert.equal(
      noteComposeSubjectPrefix('forward'),
      "This subject reads as a forward, but mode:'new' records no forwarded message, so "
        + "nothing connects this draft to the one it says it forwards. Compose with "
        + "mode:'forward' and originalEmailId to forward it properly."
    );
  });

  it('says exactly this on an edit', () => {
    assert.equal(
      noteEditSubjectPrefix('reply'),
      "This subject reads as a reply, but the draft carries no In-Reply-To or References "
        + "and an edit cannot add them, so the message starts its own conversation. Compose a "
        + "fresh draft with draft_email mode:'reply' and originalEmailId, then delete this one."
    );
    assert.equal(
      noteEditSubjectPrefix('forward'),
      "This subject reads as a forward, but the draft records no forwarded message and an "
        + "edit cannot add one, so nothing connects it to the message it says it forwards. "
        + "Compose a fresh draft with draft_email mode:'forward' and originalEmailId, then "
        + "delete this one."
    );
  });
  // Each route names the mode that would have done what the prefix claims, and each names
  // ITS OWN remedy: compose can still switch mode, an edit cannot and has to start over.
  it('the compose note names the matching mode and originalEmailId', () => {
    assert.match(noteComposeSubjectPrefix('reply'), /mode:'reply' and originalEmailId/);
    assert.match(noteComposeSubjectPrefix('forward'), /mode:'forward' and originalEmailId/);
  });

  it('the compose note offers no reply remedy for a forward prefix, and no forward remedy for a reply one', () => {
    assert.equal(noteComposeSubjectPrefix('reply').includes("mode:'forward'"), false);
    assert.equal(noteComposeSubjectPrefix('forward').includes("mode:'reply'"), false);
  });

  it('the compose reply note names to, since the reply mode defaults to reply-all', () => {
    assert.match(noteComposeSubjectPrefix('reply'), /passing to as well/);
  });

  it('the edit note sends the caller to a fresh draft rather than another edit', () => {
    for (const kind of ['reply', 'forward'] as const) {
      const note = noteEditSubjectPrefix(kind);
      assert.match(note, /Compose a fresh draft with draft_email/, `kind: ${kind}`);
      assert.match(note, /then delete this one/, `kind: ${kind}`);
      assert.match(note, new RegExp(`mode:'${kind}' and originalEmailId`), `kind: ${kind}`);
    }
  });

  it('the edit note says an edit cannot add the headers, which is why the remedy differs', () => {
    assert.match(noteEditSubjectPrefix('reply'), /an edit cannot add them/);
    assert.match(noteEditSubjectPrefix('forward'), /an edit cannot add one/);
  });
});
