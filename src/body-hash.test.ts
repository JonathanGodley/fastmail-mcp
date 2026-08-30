// The draft body hash — the lost-update guard `edit_draft` requires before it writes a body.
//
// Three properties are worth pinning here and each fails silently on its own:
//
//  1. WHAT IS HASHED. RFC 8621 §4.1.4 puts one displayed part into both body lists, so the
//     part set has to be deduplicated or an ordinary plain-text draft hashes its one part
//     twice. Two different draft bodies that hash alike would let a stale edit through.
//  2. THAT THE CANONICAL FORM IS INJECTIVE. The parts are joined, so without a length prefix
//     two parts could spell one, and a part with no stored value could spell an empty one.
//  3. WHEN A READ MAY ISSUE ONE. The hash certifies that the caller SAW the bytes it is about
//     to replace, so a read that did not show them whole must withhold it — and say so,
//     never fall silent, because a missing field reads exactly like "this draft has no hash".

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  attachDraftBodyHash, bodyHash, classifyPartType, collectDraftBodyParts,
  draftInterleavedTextType, draftPartKey, draftTextBodyType, isDraftEmail, isTextBodyType,
  resolveDraftBodyHash,
} from './body-hash.js';
import type { SimplifiedEmail } from './email-formatter.js';

/** A read that returns both body fields and strips nothing — the ordinary get_email. */
const FULL_READ = { bodyText: true, bodyHtml: true, stripQuoted: false };

const draft = (over: any) => ({ keywords: { $draft: true }, ...over });

/** A text-only draft, whose single part both body lists alias (§4.1.4). */
const TEXT_ONLY = draft({
  textBody: [{ partId: 't', type: 'text/plain' }],
  htmlBody: [{ partId: 't', type: 'text/plain' }],
  bodyValues: { t: { value: 'the body' } },
});

/** A dual-body draft: two distinct parts, one in each list. */
const DUAL = draft({
  textBody: [{ partId: 't', type: 'text/plain' }],
  htmlBody: [{ partId: 'h', type: 'text/html' }],
  bodyValues: { t: { value: 'the text' }, h: { value: '<p>the html</p>' } },
});

describe('draftPartKey', () => {
  it('prefers partId, falls back to blobId, then to the part\'s position', () => {
    assert.equal(draftPartKey({ partId: 't', blobId: 'b' }, 7), 'p:t');
    assert.equal(draftPartKey({ blobId: 'b' }, 7), 'b:b');
    assert.equal(draftPartKey({}, 7), 'i:7');
  });

  // A blank string is not an identifier. Treating one as a key would make every part that
  // carries it the SAME part, collapsing a multi-part body to one entry.
  it('treats an empty partId or blobId as absent', () => {
    assert.equal(draftPartKey({ partId: '', blobId: 'b' }, 7), 'b:b');
    assert.equal(draftPartKey({ partId: '', blobId: '' }, 7), 'i:7');
    assert.equal(draftPartKey(undefined, 3), 'i:3');
  });
});

describe('classifyPartType', () => {
  it('drops the parameters and lowercases, and answers empty for a non-string', () => {
    assert.equal(classifyPartType('TEXT/Plain; charset=utf-8'), 'text/plain');
    assert.equal(classifyPartType('  text/html  '), 'text/html');
    assert.equal(classifyPartType(undefined), '');
    assert.equal(classifyPartType(42), '');
  });
});

describe('draftTextBodyType', () => {
  it('answers the part\'s own type when it declares one a body list displays', () => {
    assert.equal(draftTextBodyType('text/plain', 'text/plain'), 'text/plain');
    // The list it sits in does not override a type the part declares: an html part listed
    // under textBody (§4.1.4 aliases a single-format draft's one part into both) is still
    // html, and calling it plain would make an ordinary html-only draft look interleaved.
    assert.equal(draftTextBodyType('text/html', 'text/plain'), 'text/html');
  });

  it('answers the LIST\'s type for a part that declares none', () => {
    // RFC 8621 §4.1.4 puts a part into a body list to say a client should display it there,
    // so list membership is the authority when the type is absent — which is how extractBody
    // has always read one.
    assert.equal(draftTextBodyType(undefined, 'text/plain'), 'text/plain');
    assert.equal(draftTextBodyType(undefined, 'text/html'), 'text/html');
    assert.equal(draftTextBodyType(null, 'text/html'), 'text/html');
    assert.equal(draftTextBodyType('', 'text/html'), 'text/html');
  });

  it('answers undefined for anything that is not displayed text', () => {
    assert.equal(draftTextBodyType('image/png', 'text/html'), undefined);
    assert.equal(draftTextBodyType('application/pdf', 'text/plain'), undefined);
    // Taken verbatim: a caller that wants parameters stripped classifies first. A truthy
    // non-string is not an absent type and must not fall through to the list's type.
    assert.equal(draftTextBodyType('text/plain; charset=utf-8', 'text/plain'), undefined);
    assert.equal(draftTextBodyType(42, 'text/plain'), undefined);
  });

  it('agrees with isTextBodyType about which types a body list displays', () => {
    assert.equal(isTextBodyType('text/plain'), true);
    assert.equal(isTextBodyType('text/html'), true);
    assert.equal(isTextBodyType('image/png'), false);
    assert.equal(isTextBodyType(undefined), false);
  });
});

// The shape a flat recreate has no spelling for (#85). One expression, asked by BOTH the
// edit-side guard and the read that issues a bodyHash — pinned here rather than beside
// either caller, so a change to it is seen as a change to both.
describe('draftInterleavedTextType', () => {
  it('names the type when two DISTINCT parts count as one text type', () => {
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a', type: 'text/plain' }, { partId: 'b', type: 'text/plain' }],
    })), 'text/plain');
  });

  it('counts across the two lists, not within one', () => {
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a', type: 'text/html' }],
      htmlBody: [{ partId: 'b', type: 'text/html' }],
    })), 'text/html');
  });

  // Deduping FIRST is what keeps an ordinary single-format draft editable: §4.1.4 aliases
  // its one part into both lists, so a raw count would see two text/plain parts.
  it('answers undefined for the ordinary draft shapes', () => {
    assert.equal(draftInterleavedTextType(TEXT_ONLY), undefined);
    assert.equal(draftInterleavedTextType(DUAL), undefined);
    assert.equal(draftInterleavedTextType(draft({})), undefined);
    assert.equal(draftInterleavedTextType(draft({ textBody: 'not an array' })), undefined);
  });

  it('classifies the type, so two spellings of one type still interleave', () => {
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a', type: 'text/plain' }, { partId: 'b', type: 'TEXT/Plain; charset=utf-8' }],
    })), 'text/plain');
  });

  it('ignores parts a body list does not display as text', () => {
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a', type: 'text/plain' }, { partId: 'i', type: 'image/png', blobId: 'B' }],
      htmlBody: [{ partId: 'j', type: 'image/png', blobId: 'C' }],
    })), undefined);
  });

  // The state this leaves #179 in, pinned so the repair is visible as a change: a part that
  // declares no type is skipped here, where every other reader of a body list treats it as
  // the content of the list it sits in.
  it('today skips a part that declares no content type', () => {
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a', type: 'text/plain' }, { partId: 'b' }],
    })), undefined);
    assert.equal(draftInterleavedTextType(draft({
      textBody: [{ partId: 'a' }, { partId: 'b' }],
    })), undefined);
  });
});

describe('collectDraftBodyParts', () => {
  it('counts a part both body lists carry exactly once', () => {
    const parts = collectDraftBodyParts(TEXT_ONLY);
    assert.equal(parts.length, 1);
    assert.equal(parts[0].value, 'the body');
    // It is displayed by the text read, and the html read shows it too — a text/plain part in
    // the htmlBody list is not carriable there, so only showsInText is set.
    assert.equal(parts[0].showsInText, true);
    assert.equal(parts[0].showsInHtml, false);
  });

  it('returns the two parts of a dual-body draft in stored order, text list first', () => {
    const parts = collectDraftBodyParts(DUAL);
    assert.deepEqual(parts.map((p) => p.value), ['the text', '<p>the html</p>']);
    assert.deepEqual(parts.map((p) => p.showsInText), [true, false]);
    assert.deepEqual(parts.map((p) => p.showsInHtml), [false, true]);
  });

  it('carries a part with no declared type in whichever list it sits in', () => {
    const parts = collectDraftBodyParts(draft({
      textBody: [{ partId: 't' }], htmlBody: [{ partId: 't' }],
      bodyValues: { t: { value: 'x' } },
    }));
    assert.equal(parts.length, 1);
    assert.equal(parts[0].showsInText, true);
    assert.equal(parts[0].showsInHtml, true);
  });

  it('records a part with no fetched value, and one the server flagged as degraded', () => {
    const parts = collectDraftBodyParts(draft({
      textBody: [{ partId: 't', type: 'text/plain' }, { blobId: 'img', type: 'image/png' }],
      bodyValues: { t: { value: 'x', isTruncated: true } },
    }));
    assert.equal(parts[0].degraded, true);
    assert.equal(parts[1].value, undefined); // an embedded image routed into a body list
    assert.equal(parts[1].degraded, false);
  });

  it('flags an encoding problem the same way as a truncation', () => {
    const parts = collectDraftBodyParts(draft({
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'x', isEncodingProblem: true } },
    }));
    assert.equal(parts[0].degraded, true);
  });

  it('returns nothing for a message with no body lists at all', () => {
    assert.deepEqual(collectDraftBodyParts(draft({})), []);
    assert.deepEqual(collectDraftBodyParts(undefined), []);
  });
});

describe('bodyHash', () => {
  it('is a bh1- token of 32 hex characters', () => {
    assert.match(bodyHash(collectDraftBodyParts(DUAL)), /^bh1-[0-9a-f]{32}$/);
  });

  it('gives the same token for the same stored body, every time', () => {
    assert.equal(bodyHash(collectDraftBodyParts(DUAL)), bodyHash(collectDraftBodyParts(DUAL)));
  });

  it('changes when any part\'s value changes', () => {
    const edited = draft({ ...DUAL, bodyValues: { t: { value: 'the text' }, h: { value: '<p>edited</p>' } } });
    assert.notEqual(bodyHash(collectDraftBodyParts(DUAL)), bodyHash(collectDraftBodyParts(edited)));
  });

  // The byte-length prefix is what makes the canonical form injective. Without it these two
  // bodies join to the same string and a stale edit against either passes the guard.
  it('cannot be spelled by a different split of the same bytes', () => {
    const a = [
      { key: 'p:1', value: 'ab', degraded: false, showsInText: true, showsInHtml: false },
      { key: 'p:2', value: 'c', degraded: false, showsInText: true, showsInHtml: false },
    ];
    const b = [
      { key: 'p:1', value: 'a', degraded: false, showsInText: true, showsInHtml: false },
      { key: 'p:2', value: 'bc', degraded: false, showsInText: true, showsInHtml: false },
    ];
    assert.notEqual(bodyHash(a), bodyHash(b));
  });

  it('distinguishes a part with no stored value from one whose value is empty', () => {
    const none = [{ key: 'p:1', degraded: false, showsInText: true, showsInHtml: false }];
    const empty = [{ key: 'p:1', value: '', degraded: false, showsInText: true, showsInHtml: false }];
    assert.notEqual(bodyHash(none), bodyHash(empty));
  });

  it('counts a length in BYTES, not characters', () => {
    // Two values of equal character length but different byte length must not collide with
    // each other through the prefix; the prefix is what the multi-byte case is about.
    const one = [{ key: 'p:1', value: 'é', degraded: false, showsInText: true, showsInHtml: false }];
    const two = [{ key: 'p:1', value: 'ee', degraded: false, showsInText: true, showsInHtml: false }];
    assert.notEqual(bodyHash(one), bodyHash(two));
  });
});

describe('isDraftEmail', () => {
  it('reads the $draft keyword and nothing else', () => {
    assert.equal(isDraftEmail(TEXT_ONLY), true);
    assert.equal(isDraftEmail({ keywords: { $seen: true } }), false);
    assert.equal(isDraftEmail({ mailboxIds: { 'mb-drafts': true } }), false);
    assert.equal(isDraftEmail(undefined), false);
  });
});

describe('resolveDraftBodyHash', () => {
  it('issues the hash of the collected parts on a full read of a draft', () => {
    assert.deepEqual(resolveDraftBodyHash(DUAL, FULL_READ), { bodyHash: bodyHash(collectDraftBodyParts(DUAL)) });
  });

  // Not a degradation to report: nothing was promised on a message that is not a draft, and
  // the field would be dead weight on every other read.
  it('returns nothing at all for a message that is not a draft', () => {
    assert.equal(resolveDraftBodyHash({ ...DUAL, keywords: { $seen: true } }, FULL_READ), undefined);
  });

  it('withholds with a reason when the server flagged a stored value as degraded', () => {
    const degraded = draft({
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'x', isTruncated: true } },
    });
    const out = resolveDraftBodyHash(degraded, FULL_READ) as any;
    assert.match(out.bodyHashWithheld, /truncated or as having encoding problems/);
    assert.match(out.bodyHashWithheld, /Recreate the draft/);
    assert.equal(out.bodyHash, undefined);
  });

  it('withholds when the draft carries a part no read returns', () => {
    // A part whose declared type does not match the list it sits in is displayed by neither
    // body field, so no read can show the caller the bytes the hash would cover.
    const odd = draft({
      textBody: [{ partId: 't', type: 'text/plain' }, { partId: 'x', type: 'application/pdf' }],
      bodyValues: { t: { value: 'x' }, x: { value: 'invisible' } },
    });
    const out = resolveDraftBodyHash(odd, FULL_READ) as any;
    assert.match(out.bodyHashWithheld, /a body part no read returns/);
  });

  it('withholds on a stripQuoted read, and names the read that would issue one', () => {
    const out = resolveDraftBodyHash(DUAL, { ...FULL_READ, stripQuoted: true }) as any;
    assert.match(out.bodyHashWithheld, /stripped quoted history/);
    assert.match(out.bodyHashWithheld, /without stripQuoted/);
  });

  // Order: a degraded part is reported ahead of stripQuoted, because re-reading without
  // stripQuoted would issue no hash either and the caller would have been sent nowhere.
  it('reports the degraded body ahead of the stripQuoted read', () => {
    const degraded = draft({
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: 'x', isTruncated: true } },
    });
    const out = resolveDraftBodyHash(degraded, { ...FULL_READ, stripQuoted: true }) as any;
    assert.match(out.bodyHashWithheld, /truncated/);
  });

  it('withholds when the projection dropped the only field that shows the body', () => {
    const out = resolveDraftBodyHash(TEXT_ONLY, { bodyText: false, bodyHtml: true, stripQuoted: false }) as any;
    assert.match(out.bodyHashWithheld, /did not return the draft's stored body whole/);
    // The remedy names the whole read. A text-only draft needs no html field, so none is
    // asked for and verbose:true is not offered — it would not help.
    assert.match(out.bodyHashWithheld, /fields: \["bodyText", "bodyHash"\]/);
    assert.equal(/verbose/.test(out.bodyHashWithheld), false);
  });

  it('names BOTH body fields, and verbose:true, when the draft needs both', () => {
    const out = resolveDraftBodyHash(DUAL, { bodyText: true, bodyHtml: false, stripQuoted: false }) as any;
    // Not just the field this read lacked: the hash needs every part shown at once, so a
    // remedy naming only bodyHtml would send the caller to a read that issues no hash.
    assert.match(out.bodyHashWithheld, /fields: \["bodyText", "bodyHtml", "bodyHash"\] \(or verbose:true\)/);
  });

  it('issues the hash when the one field the draft needs is present', () => {
    const out = resolveDraftBodyHash(TEXT_ONLY, { bodyText: true, bodyHtml: false, stripQuoted: false }) as any;
    assert.equal(out.bodyHash, bodyHash(collectDraftBodyParts(TEXT_ONLY)));
  });

  // An empty-string part is not content — extractBody skips it, so no read displays it and no
  // read has to. It still hashes distinctly from a valueless part (see bodyHash above).
  it('issues a hash for a draft whose only stored value is empty, whatever the read carries', () => {
    const blank = draft({
      textBody: [{ partId: 't', type: 'text/plain' }],
      bodyValues: { t: { value: '' } },
    });
    const out = resolveDraftBodyHash(blank, { bodyText: false, bodyHtml: false, stripQuoted: false }) as any;
    assert.equal(out.bodyHash, bodyHash(collectDraftBodyParts(blank)));
  });

  // A part carried by BOTH lists is satisfied by either field, so a read that dropped one of
  // them still shows the caller every stored byte.
  it('accepts either field for a part both lists carry', () => {
    const typeless = draft({
      textBody: [{ partId: 't' }], htmlBody: [{ partId: 't' }],
      bodyValues: { t: { value: 'x' } },
    });
    for (const read of [
      { bodyText: true, bodyHtml: false, stripQuoted: false },
      { bodyText: false, bodyHtml: true, stripQuoted: false },
    ]) {
      assert.equal((resolveDraftBodyHash(typeless, read) as any).bodyHash, bodyHash(collectDraftBodyParts(typeless)));
    }
    const neither = resolveDraftBodyHash(typeless, { bodyText: false, bodyHtml: false, stripQuoted: false }) as any;
    assert.match(neither.bodyHashWithheld, /did not return the draft's stored body whole/);
  });
});

// The read-side decision: which get_email responses may carry the token at all. This is the
// half that depends on what the RESPONSE ended up holding rather than on the message, so it
// cannot be derived from the draft alone and has to be pinned separately.
describe('attachDraftBodyHash', () => {
  const simplifiedOf = (over: Partial<SimplifiedEmail> = {}): SimplifiedEmail =>
    ({ id: 'd1', bodyText: 'the body', ...over } as SimplifiedEmail);

  it('attaches the hash to an ordinary unprojected read of a draft', () => {
    const out = simplifiedOf();
    attachDraftBodyHash(TEXT_ONLY, out, { raw: false, fields: undefined, stripQuoted: false });
    assert.equal(out.bodyHash, bodyHash(collectDraftBodyParts(TEXT_ONLY)));
  });

  // raw is unmodified JMAP. A field of this server's own invention in it would stop it being
  // raw — and the caller has the stored bodyValues in front of it anyway.
  it('attaches NEITHER field on a raw read, hash or reason', () => {
    const out = simplifiedOf();
    attachDraftBodyHash(TEXT_ONLY, out, { raw: true, fields: undefined, stripQuoted: true });
    assert.equal('bodyHash' in out, false);
    assert.equal('bodyHashWithheld' in out, false);
  });

  it('attaches neither field to a message that is not a draft', () => {
    const out = simplifiedOf();
    attachDraftBodyHash({ ...TEXT_ONLY, keywords: { $seen: true } }, out, { raw: false, fields: undefined, stripQuoted: false });
    assert.equal('bodyHash' in out, false);
    assert.equal('bodyHashWithheld' in out, false);
  });

  // The two ways a body can fail to reach the caller, and they must count the same: the
  // simplifier never produced the field, or the projection dropped it. Either way the
  // response does not show the bytes the hash is over.
  it('withholds when the simplifier produced no body field', () => {
    const out = simplifiedOf({ bodyText: undefined });
    attachDraftBodyHash(TEXT_ONLY, out, { raw: false, fields: undefined, stripQuoted: false });
    assert.match(out.bodyHashWithheld!, /did not return the draft's stored body whole/);
    assert.equal(out.bodyHash, undefined);
  });

  it('withholds when the projection dropped the body field the simplifier did produce', () => {
    const out = simplifiedOf();
    attachDraftBodyHash(TEXT_ONLY, out, { raw: false, fields: new Set(['id', 'subject']), stripQuoted: false });
    assert.match(out.bodyHashWithheld!, /did not return the draft's stored body whole/);
  });

  it('attaches the hash when the projection kept a body field', () => {
    const out = simplifiedOf();
    attachDraftBodyHash(TEXT_ONLY, out, { raw: false, fields: new Set(['bodyText']), stripQuoted: false });
    assert.equal(out.bodyHash, bodyHash(collectDraftBodyParts(TEXT_ONLY)));
  });

  it('passes stripQuoted straight through, so a shortened body withholds', () => {
    const out = simplifiedOf();
    attachDraftBodyHash(TEXT_ONLY, out, { raw: false, fields: undefined, stripQuoted: true });
    assert.match(out.bodyHashWithheld!, /stripped quoted history/);
  });

  // Exactly one of the two, always, on a draft read that is not raw. A response carrying
  // neither is indistinguishable from a non-draft, and the caller could not tell which.
  it('leaves a draft read carrying exactly one of the two fields', () => {
    for (const options of [
      { raw: false, fields: undefined, stripQuoted: false },
      { raw: false, fields: undefined, stripQuoted: true },
      { raw: false, fields: new Set(['subject']), stripQuoted: false },
    ]) {
      const out = simplifiedOf();
      attachDraftBodyHash(TEXT_ONLY, out, options);
      assert.equal(('bodyHash' in out ? 1 : 0) + ('bodyHashWithheld' in out ? 1 : 0), 1, JSON.stringify(options));
    }
  });
});
