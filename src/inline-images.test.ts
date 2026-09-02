import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  buildCidMap,
  buildUnionParts,
  checkInlineClosure,
  cidKey,
  collectImgCidRefs,
  decodeCidSrc,
  describePart,
  extractCidRefs,
  InlineClosureError,
  isAuthorableCid,
  isOurMint,
  isRecreatableCid,
  isReservedCid,
  launderUrlValue,
  mintCid,
  reconcileInlineParts,
  resolveCidRefs,
  sanitizeDownloadFilename,
  sanitizeQuoteHtml,
  stripCidSpelling,
  urlScheme,
} from './inline-images.js';

const TEXT_PART = { partId: '1', type: 'text/plain', size: 10 };
const HTML_PART = { partId: '2', type: 'text/html', size: 20 };
const IMAGE_PART = { partId: '3', type: 'image/png', size: 30, blobId: 'B3', cid: 'logo@host' };
const FILE_PART = { partId: '4', type: 'application/pdf', size: 40, blobId: 'B4', name: 'a.pdf' };

describe('buildUnionParts gating', () => {
  it('returns nothing when attachments was not fetched (the compact property set)', () => {
    // list/search fetch textBody for the body-size hint but no attachments, so there is
    // no listing to complete — emitting half a listing there would be worse than none.
    assert.deepEqual(buildUnionParts({ textBody: [TEXT_PART, IMAGE_PART] }), []);
  });

  it('returns nothing for a null or absent email', () => {
    assert.deepEqual(buildUnionParts(undefined), []);
    assert.deepEqual(buildUnionParts(null), []);
    assert.deepEqual(buildUnionParts({}), []);
  });

  it('unions from an EMPTY attachments array — the shape an embedded-only message has', () => {
    const union = buildUnionParts({
      attachments: [],
      textBody: [TEXT_PART, IMAGE_PART],
      htmlBody: [HTML_PART, IMAGE_PART],
    });
    assert.equal(union.length, 1);
    assert.equal(union[0].part, IMAGE_PART);
    assert.equal(union[0].inBodyList, true);
  });
});

describe('buildUnionParts membership and order', () => {
  it('keeps attachments in server order, then body-list additions', () => {
    const other = { partId: '5', type: 'image/gif', size: 5, blobId: 'B5' };
    const union = buildUnionParts({
      attachments: [FILE_PART, IMAGE_PART],
      textBody: [TEXT_PART, other],
      htmlBody: [HTML_PART],
    });
    assert.deepEqual(union.map((u) => u.part.partId), ['4', '3', '5']);
  });

  it('adds textBody parts before htmlBody parts', () => {
    const fromText = { partId: '8', type: 'image/png', blobId: 'B8' };
    const fromHtml = { partId: '9', type: 'image/png', blobId: 'B9' };
    const union = buildUnionParts({
      attachments: [],
      textBody: [TEXT_PART, fromText],
      htmlBody: [HTML_PART, fromHtml],
    });
    assert.deepEqual(union.map((u) => u.part.partId), ['8', '9']);
  });

  it('never lists the rendered body parts themselves', () => {
    const union = buildUnionParts({
      attachments: [],
      textBody: [TEXT_PART],
      htmlBody: [HTML_PART],
    });
    assert.deepEqual(union, []);
  });

  it('ignores content-type case and parameters when classifying a body part', () => {
    const shouty = { partId: '6', type: 'TEXT/PLAIN; charset=UTF-8', size: 1 };
    const image = { partId: '7', type: 'IMAGE/PNG', size: 1, blobId: 'B7' };
    const union = buildUnionParts({ attachments: [], textBody: [shouty, image] });
    assert.deepEqual(union.map((u) => u.part.partId), ['7']);
  });

  it('emits the part object verbatim, never a copy', () => {
    const union = buildUnionParts({ attachments: [IMAGE_PART] });
    assert.equal(union[0].part, IMAGE_PART);
  });

  it('leaves the source lists unmutated', () => {
    const attachments = [FILE_PART];
    const textBody = [TEXT_PART, IMAGE_PART];
    buildUnionParts({ attachments, textBody });
    assert.deepEqual(attachments, [FILE_PART]);
    assert.deepEqual(textBody, [TEXT_PART, IMAGE_PART]);
  });
});

describe('buildUnionParts dedup', () => {
  it('lists a part once when it appears in attachments and both body lists', () => {
    const union = buildUnionParts({
      attachments: [IMAGE_PART],
      textBody: [TEXT_PART, IMAGE_PART],
      htmlBody: [HTML_PART, IMAGE_PART],
    });
    assert.equal(union.length, 1);
  });

  it('reports body-list routing even when attachments listed the part first', () => {
    // The routing signal must not depend on which list won the dedup.
    const union = buildUnionParts({ attachments: [IMAGE_PART], htmlBody: [HTML_PART, IMAGE_PART] });
    assert.equal(union[0].inBodyList, true);
  });

  it('reports inBodyList false for an attachments-only part', () => {
    const union = buildUnionParts({ attachments: [FILE_PART], textBody: [TEXT_PART] });
    assert.equal(union[0].inBodyList, false);
  });

  it('falls back to blobId when a part carries no partId', () => {
    const noPartId = { type: 'image/png', blobId: 'SHARED', size: 1 };
    const alias = { type: 'image/png', blobId: 'SHARED', size: 1 };
    const union = buildUnionParts({ attachments: [noPartId], htmlBody: [HTML_PART, alias] });
    assert.equal(union.length, 1);
    assert.equal(union[0].inBodyList, true);
  });

  it('keeps a part that has neither partId nor blobId rather than folding it away', () => {
    const anon1 = { type: 'image/png', size: 1 };
    const anon2 = { type: 'image/png', size: 2 };
    const union = buildUnionParts({ attachments: [anon1, anon2] });
    assert.equal(union.length, 2);
  });

  it('lists a key-less part once even when the same object is in both body lists', () => {
    // A displayed part is routed into textBody AND htmlBody, so without an identity
    // backstop an unidentifiable part would be listed twice and shift entry numbers.
    const anon = { type: 'image/png', size: 1 };
    const union = buildUnionParts({
      attachments: [],
      textBody: [TEXT_PART, anon],
      htmlBody: [HTML_PART, anon],
    });
    assert.equal(union.length, 1);
    assert.equal(union[0].inBodyList, true);
  });

  it('skips null entries in any list', () => {
    const union = buildUnionParts({
      attachments: [null, FILE_PART],
      textBody: [undefined, IMAGE_PART],
    } as any);
    assert.deepEqual(union.map((u) => u.part.partId), ['4', '3']);
  });

  it('ignores a body list that is not an array', () => {
    const union = buildUnionParts({ attachments: [FILE_PART], textBody: 'nope' } as any);
    assert.deepEqual(union.map((u) => u.part.partId), ['4']);
  });

  it('does not list a body part that carries no type at all', () => {
    // Matching the body extractor, which reads an untyped part as body text.
    const untyped = { partId: '9', size: 1, blobId: 'B9' };
    assert.deepEqual(buildUnionParts({ attachments: [], textBody: [untyped] }), []);
  });
});

describe('decodeCidSrc', () => {
  it('decodes a percent escape once', () => {
    assert.equal(decodeCidSrc('%78'), 'x');
    assert.equal(decodeCidSrc('logo%40host'), 'logo@host');
  });

  it('decodes only once, so a double escape cannot collapse further', () => {
    assert.equal(decodeCidSrc('%2578'), '%78');
  });

  it('returns a malformed escape verbatim instead of throwing', () => {
    assert.equal(decodeCidSrc('%'), '%');
    assert.equal(decodeCidSrc('%zz'), '%zz');
    assert.equal(decodeCidSrc('100% sure'), '100% sure');
  });

  it('leaves a value with no escapes untouched', () => {
    assert.equal(decodeCidSrc('logo@host'), 'logo@host');
    assert.equal(decodeCidSrc(''), '');
  });
});

describe('cidKey', () => {
  it('strips the cid: scheme and decodes the reference', () => {
    assert.equal(cidKey('cid:logo%40host'), 'logo@host');
  });

  it('accepts the scheme in any case', () => {
    assert.equal(cidKey('CID:logo'), 'logo');
    assert.equal(cidKey('Cid:logo'), 'logo');
  });

  it('strips only the first scheme, so a Content-ID beginning with cid: stays reachable', () => {
    assert.equal(cidKey('cid:cid:x'), 'cid:x');
  });

  it('is a no-op on a value that carries no scheme', () => {
    assert.equal(cidKey('logo@host'), 'logo@host');
  });

  it('does not decode the part side: a literal %78 cid is not the same as x', () => {
    // Reference-side only. Decoding a part's Content-ID too would let a reference
    // spelled "x" resolve to a part whose id genuinely contains "%78".
    assert.notEqual(cidKey('cid:x'), '%78');
    assert.equal(cidKey('cid:%78'), 'x');
  });
});

describe('describePart', () => {
  it('passes an ordinary value through unchanged', () => {
    assert.equal(describePart('logo@example.com'), 'logo@example.com');
  });

  it('strips control characters, including newlines', () => {
    assert.equal(describePart('a\nb\r\tc\u0000d'), 'abcd');
  });

  it('strips format and bidi-override characters', () => {
    assert.equal(describePart('inv\u202Efdp.exe'), 'invfdp.exe');
    assert.equal(describePart('a\u200Bb'), 'a\u200Bb'.replace('\u200B', ''));
  });

  it('strips line and paragraph separators', () => {
    assert.equal(describePart('a\u2028b\u2029c'), 'abc');
  });

  it('collapses runs of space separators to one plain space', () => {
    assert.equal(describePart('a\u00a0\u2003 b'), 'a b');
  });

  it('replaces a double quote so the value cannot close its quoted span', () => {
    assert.equal(describePart('a" then instructions'), "a' then instructions");
  });

  it('caps at 64 code points and marks the truncation', () => {
    const out = describePart('x'.repeat(200));
    assert.equal(out, `${'x'.repeat(64)}…`);
  });

  it('caps by code point, never splitting a surrogate pair', () => {
    const out = describePart('\u{1F600}'.repeat(100));
    assert.equal([...out].length, 65); // 64 code points plus the truncation mark
    // No unpaired surrogate anywhere: a UTF-16 unit slice would have left one.
    assert.ok(!/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(out));
  });

  it('leaves a value of exactly the cap unmarked', () => {
    assert.equal(describePart('y'.repeat(64)), 'y'.repeat(64));
  });

  it('renders a non-string or absent value as a string', () => {
    assert.equal(describePart(undefined), '');
    assert.equal(describePart(null), '');
    assert.equal(describePart(42), '42');
  });
});

describe('sanitizeDownloadFilename', () => {
  it('passes an ordinary filename through', () => {
    assert.equal(sanitizeDownloadFilename('logo.png'), 'logo.png');
  });

  it('does not append .eml — that belongs to the forwarded-message helper alone', () => {
    assert.equal(sanitizeDownloadFilename('logo.png'), 'logo.png');
    assert.ok(!sanitizeDownloadFilename('logo.png').endsWith('.eml'));
  });

  it('falls back to "attachment" when the name is absent or sanitizes to nothing', () => {
    assert.equal(sanitizeDownloadFilename(undefined), 'attachment');
    assert.equal(sanitizeDownloadFilename(null), 'attachment');
    assert.equal(sanitizeDownloadFilename(''), 'attachment');
    assert.equal(sanitizeDownloadFilename('\u0000\u0007'), 'attachment');
    assert.equal(sanitizeDownloadFilename('...'), 'attachment');
  });

  it('maps path separators and the drive colon to underscores', () => {
    // The separators go first, then the leading-dot strip runs on the result. Dots
    // themselves survive, but with every separator gone they can no longer form a path
    // component, so a traversal attempt cannot describe a directory.
    assert.equal(sanitizeDownloadFilename('../../etc/passwd'), '_.._etc_passwd');
    assert.equal(sanitizeDownloadFilename('C:\\evil.exe'), 'C__evil.exe');
  });

  it('strips leading dots so the result is never a hidden or relative segment', () => {
    assert.equal(sanitizeDownloadFilename('.hidden'), 'hidden');
  });

  it('strips a leading dot that whitespace would otherwise shield', () => {
    // Trimming after the dot strip would leave ".hidden" here, reinstating the name
    // the strip exists to prevent.
    assert.equal(sanitizeDownloadFilename(' .hidden'), 'hidden');
  });

  it('falls back when a name is nothing but whitespace and dots', () => {
    assert.equal(sanitizeDownloadFilename(' ..'), 'attachment');
  });

  it('strips control and bidi-override characters', () => {
    assert.equal(sanitizeDownloadFilename('inv\u202Efdp.exe'), 'invfdp.exe');
  });

  it('caps the length by code point', () => {
    assert.equal([...sanitizeDownloadFilename('a'.repeat(300))].length, 80);
    assert.equal([...sanitizeDownloadFilename('\u{1F600}'.repeat(300))].length, 80);
  });

  it('defuses a Windows device name, extension included', () => {
    assert.equal(sanitizeDownloadFilename('CON'), 'CON_');
    assert.equal(sanitizeDownloadFilename('nul'), 'nul_');
    assert.equal(sanitizeDownloadFilename('CON.png'), 'CON_.png');
    assert.equal(sanitizeDownloadFilename('CoM1.txt'), 'CoM1_.txt');
    assert.equal(sanitizeDownloadFilename('LPT9'), 'LPT9_');
    assert.equal(sanitizeDownloadFilename('aux.tar.gz'), 'aux_.tar.gz');
  });

  it('defuses a device name padded with the characters Win32 strips first', () => {
    // Win32 removes trailing spaces and dots from a path component BEFORE matching
    // device names, so these are the console too.
    assert.equal(sanitizeDownloadFilename('CON .png'), 'CON _.png');
    assert.equal(sanitizeDownloadFilename('com1 .txt'), 'com1 _.txt');
    assert.equal(sanitizeDownloadFilename('CON  '), 'CON_');
    assert.equal(sanitizeDownloadFilename('nul..txt'), 'nul_..txt');
  });

  it('leaves a name that merely resembles a device name alone', () => {
    assert.equal(sanitizeDownloadFilename('CONTRACT.pdf'), 'CONTRACT.pdf');
    assert.equal(sanitizeDownloadFilename('COM10.txt'), 'COM10.txt');
    assert.equal(sanitizeDownloadFilename('report.con'), 'report.con');
  });
});

// A stand-in identifier of the shape the server mints, fixed so assertions can name it.
const MINT_A = `ii-${'a'.repeat(32)}@inline.invalid`;
const MINT_B = `ii-${'b'.repeat(32)}@inline.invalid`;

// A deterministic mint for the mapping tests, so a claim decision is visible rather than
// hidden behind a random value.
function sequentialMint(): () => string {
  let n = 0;
  return () => `ii-${String(n++).padStart(32, '0')}@inline.invalid`;
}
const MINT_0 = `ii-${'0'.repeat(32)}@inline.invalid`;
const MINT_1 = `ii-${'0'.repeat(31)}1@inline.invalid`;

describe('isAuthorableCid', () => {
  it('accepts a simple token of letters, digits, dot, dash and underscore', () => {
    for (const value of ['logo', 'LOGO', 'a', '0', 'logo.png', 'logo-1', 'logo_1', 'a.b-c_d9']) {
      assert.equal(isAuthorableCid(value), true, value);
    }
  });

  it('accepts exactly 64 characters and rejects 65', () => {
    assert.equal(isAuthorableCid('a'.repeat(64)), true);
    assert.equal(isAuthorableCid('a'.repeat(65)), false);
  });

  it('rejects an empty value', () => {
    assert.equal(isAuthorableCid(''), false);
  });

  it('rejects every character outside the allowlist, including the ones a header uses', () => {
    for (const value of [
      'logo@host', 'logo host', 'logo:1', 'logo;1', 'logo,1', 'logo(1)', 'logo<1>',
      'logo"1', 'logo/1', 'logo\\1', 'logo%40host', 'logoé', 'logo\t',
    ]) {
      assert.equal(isAuthorableCid(value), false, value);
    }
  });

  it('rejects a value carrying a line break, the header-injection shape', () => {
    // A Content-ID goes into a MIME header verbatim; a break in it ends that header and
    // starts one of the sender's choosing, and a read of the stored message shows only the
    // fragment before the break.
    assert.equal(isAuthorableCid('logo\r\nX-Injected: yes'), false);
    assert.equal(isAuthorableCid('logo\rX'), false);
    assert.equal(isAuthorableCid('logo\nX'), false);
  });

  it('rejects a non-string', () => {
    assert.equal(isAuthorableCid(undefined), false);
    assert.equal(isAuthorableCid(null), false);
    assert.equal(isAuthorableCid(42), false);
  });
});

describe('isRecreatableCid', () => {
  it('accepts the identifiers real mail clients write', () => {
    for (const value of [
      'image001.png@01CDD4E9.F5F9A6E0',
      'part1.2.3@example.com',
      'ii-0123456789abcdef0123456789abcdef@inline.invalid',
      'logo',
    ]) {
      assert.equal(isRecreatableCid(value), true, value);
    }
  });

  it('admits colon, semicolon and comma', () => {
    // These were measured to round-trip exactly through a store-and-recreate cycle, so
    // excluding them would refuse to edit drafts this server can reproduce faithfully.
    assert.equal(isRecreatableCid('a:b'), true);
    assert.equal(isRecreatableCid('a;b'), true);
    assert.equal(isRecreatableCid('a,b'), true);
  });

  it('excludes parentheses, angle brackets and the double quote', () => {
    for (const value of ['a(b)', 'a)b(', 'a<b', 'a>b', 'a"b', '<a@b>']) {
      assert.equal(isRecreatableCid(value), false, value);
    }
  });

  it('excludes whitespace of every kind', () => {
    // `\u00A0` is a non-breaking space, written as the escape because a raw one is
    // indistinguishable from the plain space in the entry beside it: the two cases would read
    // as a duplicate, and a re-encode could turn one into the other unnoticed.
    for (const value of ['a b', 'a\tb', 'a\u00A0b', ' a', 'a ']) {
      assert.equal(isRecreatableCid(value), false, value);
    }
  });

  it('excludes a line break, the header-injection shape', () => {
    // The printable-ASCII range is what blocks this: a carriage return in a Content-ID is
    // stored as a genuine extra header and is invisible in a normal read of the message.
    // Every control character here is an escape, DEL included: a raw one is invisible in the
    // source too, so an editor or a re-encode could drop it and leave the case testing 'ab'.
    assert.equal(isRecreatableCid('a\r\nX-Injected: yes'), false);
    assert.equal(isRecreatableCid('a\rb'), false);
    assert.equal(isRecreatableCid('a\nb'), false);
    assert.equal(isRecreatableCid('a\x00b'), false);
    assert.equal(isRecreatableCid('a\x7Fb'), false);
  });

  it('excludes non-ASCII', () => {
    assert.equal(isRecreatableCid('café@host'), false);
  });

  it('accepts up to 998 characters and rejects 999', () => {
    assert.equal(isRecreatableCid('a'.repeat(998)), true);
    assert.equal(isRecreatableCid('a'.repeat(999)), false);
  });

  it('rejects an empty value and a non-string', () => {
    assert.equal(isRecreatableCid(''), false);
    assert.equal(isRecreatableCid(null), false);
    assert.equal(isRecreatableCid(12), false);
  });
});

describe('mintCid and the reserved shape', () => {
  it('mints an identifier the reserved-shape test recognizes', () => {
    const cid = mintCid();
    assert.match(cid, /^ii-[0-9a-f]{32}@inline\.invalid$/);
    assert.equal(isReservedCid(cid), true);
    assert.equal(isOurMint(cid), true);
  });

  it('mints a distinct identifier each time', () => {
    const seen = new Set(Array.from({ length: 50 }, () => mintCid()));
    assert.equal(seen.size, 50);
  });

  it('mints an identifier both vets accept, so a minted part can be recreated', () => {
    assert.equal(isRecreatableCid(mintCid()), true);
  });

  it('uses a domain reserved as permanently unresolvable, so it cannot name a real host', () => {
    assert.ok(mintCid().endsWith('@inline.invalid'));
  });

  it('rejects values that only resemble the shape', () => {
    for (const value of [
      `ii-${'a'.repeat(31)}@inline.invalid`,
      `ii-${'a'.repeat(33)}@inline.invalid`,
      `ii-${'g'.repeat(32)}@inline.invalid`,
      `xx-${'a'.repeat(32)}@inline.invalid`,
      `ii-${'a'.repeat(32)}@inline.invalid.evil.com`,
      `prefix-ii-${'a'.repeat(32)}@inline.invalid`,
      `ii-${'a'.repeat(32)}@example.com`,
      'logo@host',
      '',
    ]) {
      assert.equal(isReservedCid(value), false, value);
    }
  });

  it('is case-insensitive on the hex label, matching how a header may be re-cased', () => {
    assert.equal(isReservedCid(`ii-${'A'.repeat(32)}@INLINE.INVALID`), true);
  });

  it('rejects a non-string', () => {
    assert.equal(isReservedCid(undefined), false);
    assert.equal(isReservedCid(null), false);
  });

  it('exposes one predicate under both names, so the two call sites cannot diverge', () => {
    assert.equal(isOurMint, isReservedCid);
  });
});

describe('stripCidSpelling', () => {
  it('normalizes the two spellings an identifier is realistically copied from', () => {
    assert.equal(stripCidSpelling('cid:logo'), 'logo');
    assert.equal(stripCidSpelling('<logo>'), 'logo');
    assert.equal(stripCidSpelling('logo'), 'logo');
  });

  it('strips the angle brackets first, so an angle-quoted reference normalizes', () => {
    assert.equal(stripCidSpelling('<cid:logo>'), 'logo');
    assert.equal(isAuthorableCid(stripCidSpelling('<cid:logo>')), true);
  });

  it('leaves a prefix-then-brackets spelling failing the vet', () => {
    // Not one of the two real copy sources, and the order is what keeps it out: after the
    // prefix comes off there is no second angle-bracket pass.
    assert.equal(stripCidSpelling('cid:<logo>'), '<logo>');
    assert.equal(isAuthorableCid(stripCidSpelling('cid:<logo>')), false);
  });

  it('strips each spelling at most once', () => {
    assert.equal(stripCidSpelling('<<logo>>'), '<logo>');
    assert.equal(stripCidSpelling('cid:cid:logo'), 'cid:logo');
    assert.equal(isAuthorableCid(stripCidSpelling('<<logo>>')), false);
    assert.equal(isAuthorableCid(stripCidSpelling('cid:cid:logo')), false);
  });

  it('accepts the prefix in any case', () => {
    assert.equal(stripCidSpelling('CID:logo'), 'logo');
    assert.equal(stripCidSpelling('<Cid:logo>'), 'logo');
  });

  it('does not widen what the vet accepts', () => {
    // Authored identifiers are simple local tokens; the strip makes two spellings work, it
    // does not admit a foreign identifier.
    assert.equal(stripCidSpelling('<logo@host>'), 'logo@host');
    assert.equal(isAuthorableCid(stripCidSpelling('<logo@host>')), false);
  });

  it('leaves a value with neither spelling untouched', () => {
    assert.equal(stripCidSpelling(''), '');
    assert.equal(stripCidSpelling('<'), '<');
    assert.equal(stripCidSpelling('>'), '>');
    assert.equal(stripCidSpelling('a>b<c'), 'a>b<c');
  });

  it('canonicalizes two spellings of one identifier to one value', () => {
    // Which is what lets a duplicate check see them as one identifier rather than two
    // parts that would end up sharing a Content-ID.
    assert.equal(stripCidSpelling('cid:logo'), stripCidSpelling('<cid:logo>'));
    assert.equal(stripCidSpelling('logo'), stripCidSpelling('cid:logo'));
  });
});

describe('launderUrlValue and urlScheme', () => {
  it('strips every character of code 0x20 and below, anywhere in the value', () => {
    assert.equal(launderUrlValue('c id:x'), 'cid:x');
    assert.equal(launderUrlValue('cid\t:x'), 'cid:x');
    assert.equal(launderUrlValue('cid\x00:x'), 'cid:x');
    assert.equal(launderUrlValue('cid\x1f:x'), 'cid:x');
    assert.equal(launderUrlValue(' cid:x '), 'cid:x');
    assert.equal(launderUrlValue('c\ni\rd:x'), 'cid:x');
  });

  it('clobbers an embedded comment', () => {
    assert.equal(launderUrlValue('c<!--z-->id:x'), 'cid:x');
    assert.equal(launderUrlValue('<!--a-->c<!--b-->id:x'), 'cid:x');
  });

  it('leaves an unterminated comment marker in place, as a browser would', () => {
    assert.equal(launderUrlValue('c<!--id:x'), 'c<!--id:x');
  });

  it('leaves an ordinary value untouched', () => {
    assert.equal(launderUrlValue('http://example.com/a.png'), 'http://example.com/a.png');
  });

  it('reports the scheme in lower case, after normalization', () => {
    assert.equal(urlScheme('CID:x'), 'cid');
    assert.equal(urlScheme('c id:x'), 'cid');
    assert.equal(urlScheme('HtTp://a'), 'http');
    assert.equal(urlScheme('data:image/png;base64,AA'), 'data');
  });

  it('reports no scheme for a relative or scheme-less value', () => {
    assert.equal(urlScheme('logo.png'), null);
    assert.equal(urlScheme('/logo.png'), null);
    assert.equal(urlScheme('//example.com/logo.png'), null);
    assert.equal(urlScheme(''), null);
    assert.equal(urlScheme(undefined), null);
  });
});

describe('sanitizeQuoteHtml, collecting pass', () => {
  it('reports the references an <img> carries, in first-seen order', () => {
    const out = sanitizeQuoteHtml('<img src="cid:b"><img src="cid:a"><img src="cid:b">', {
      mode: 'collect',
    });
    assert.deepEqual(out.refs, ['b', 'a']);
  });

  it('drops every embedded image, exactly as the sanitizer alone would', () => {
    // `cid` is not in the collecting configuration's scheme list, so the html this pass
    // produces is the html a quotability check has always read.
    const out = sanitizeQuoteHtml('<div><img src="cid:logo" alt="Logo"></div>', {
      mode: 'collect',
    });
    assert.equal(out.html, '<div></div>');
    assert.equal(out.droppedCidImages, 1);
  });

  it('keeps a remote image', () => {
    const out = sanitizeQuoteHtml('<img src="https://example.com/a.png" alt="a">', {
      mode: 'collect',
    });
    assert.match(out.html, /src="https:\/\/example\.com\/a\.png"/);
  });

  it('keeps a relative image, which the sanitizer has always passed through', () => {
    const out = sanitizeQuoteHtml('<img src="/logo.png">', { mode: 'collect' });
    assert.match(out.html, /src="\/logo\.png"/);
  });

  it('counts data: images separately', () => {
    const out = sanitizeQuoteHtml(
      '<img src="data:image/png;base64,AA"><img src="data:image/gif;base64,BB">',
      { mode: 'collect' },
    );
    assert.equal(out.droppedDataImages, 2);
    assert.equal(out.droppedCidImages, 0);
    assert.equal(out.html, '');
  });

  it('embeds nothing, because it decides nothing', () => {
    assert.deepEqual(sanitizeQuoteHtml('<img src="cid:a">', { mode: 'collect' }).embedded, []);
  });

  it('still strips scripts, handlers and unscoped attributes', () => {
    const out = sanitizeQuoteHtml(
      '<div style="x" onclick="y" class="z"><script>bad()</script><b>hi</b></div>',
      { mode: 'collect' },
    );
    assert.equal(out.html, '<div><b>hi</b></div>');
  });
});

describe('sanitizeQuoteHtml, mapping pass', () => {
  const cidMap = new Map([['logo', MINT_A]]);

  it('rewrites a resolved reference to the identifier being attached', () => {
    const out = sanitizeQuoteHtml('<div><img src="cid:logo" alt="Logo"></div>', {
      mode: 'map',
      cidMap,
    });
    assert.match(out.html, new RegExp(`src="cid:${MINT_A}"`));
    assert.match(out.html, /alt="Logo"/);
    assert.deepEqual(out.embedded, [MINT_A]);
    assert.equal(out.droppedCidImages, 0);
  });

  it('drops an unresolved reference and counts it', () => {
    const out = sanitizeQuoteHtml('<div><img src="cid:missing"></div>', { mode: 'map', cidMap });
    assert.equal(out.html, '<div></div>');
    assert.equal(out.droppedCidImages, 1);
    assert.deepEqual(out.embedded, []);
  });

  it('reports each emitted identifier once even when several references share it', () => {
    const out = sanitizeQuoteHtml('<img src="cid:logo"><img src="cid:logo">', {
      mode: 'map',
      cidMap,
    });
    assert.deepEqual(out.embedded, [MINT_A]);
    assert.deepEqual(out.refs, ['logo']);
  });

  it('keeps a remote image and drops a data: image', () => {
    const out = sanitizeQuoteHtml(
      '<img src="https://example.com/a.png"><img src="data:image/png;base64,AA">',
      { mode: 'map', cidMap },
    );
    assert.equal(out.html, '<img src="https://example.com/a.png" />');
    assert.equal(out.droppedDataImages, 1);
  });

  it('drops a relative or scheme-less image, because nothing affirmatively emitted it', () => {
    // Such a reference is already broken in mail — there is no base URL to resolve it
    // against — and keeping it would mean trusting the classifier's negative answer.
    assert.equal(sanitizeQuoteHtml('<img src="/logo.png">', { mode: 'map', cidMap }).html, '');
    assert.equal(sanitizeQuoteHtml('<img src="logo.png">', { mode: 'map', cidMap }).html, '');
    assert.equal(sanitizeQuoteHtml('<img>', { mode: 'map', cidMap }).html, '');
  });

  it('counts each dropped reference form, so no image disappears silently', () => {
    for (const src of ['/logo.png', 'logo.png', '//cdn.example.com/logo.png', 'ftp://a/b']) {
      const out = sanitizeQuoteHtml(`<img src="${src}">`, { mode: 'map', cidMap });
      assert.equal(out.droppedUnsupportedImages, 1, src);
      // Not confused with the two losses that have their own sentences.
      assert.equal(out.droppedDataImages, 0, src);
      assert.equal(out.droppedCidImages, 0, src);
    }
  });

  it('counts no loss for an image that never named a source', () => {
    for (const tag of ['<img>', '<img src="">', '<img src="  ">']) {
      assert.equal(sanitizeQuoteHtml(tag, { mode: 'map', cidMap }).droppedUnsupportedImages, 0, tag);
    }
  });

  it('the collecting pass counts none of them: it drops nothing on its own', () => {
    const out = sanitizeQuoteHtml(
      '<img src="/logo.png"><img src="//cdn.example.com/a.png"><img src="ftp://a/b">',
      { mode: 'collect' },
    );
    assert.equal(out.droppedUnsupportedImages, 0);
  });

  it('drops an unknown or dangerous scheme', () => {
    for (const src of ['javascript:alert(1)', 'vbscript:x', 'file:///etc/passwd', 'ftp://a/b']) {
      assert.equal(sanitizeQuoteHtml(`<img src="${src}">`, { mode: 'map', cidMap }).html, '', src);
    }
  });

  it('resolves the reference side by its decoded value, never the part side', () => {
    // A reference in html is a URL and is percent-decoded once; a part's own Content-ID is
    // not a URL and is matched literally.
    const decoded = new Map([['x', MINT_A]]);
    assert.match(
      sanitizeQuoteHtml('<img src="cid:%78">', { mode: 'map', cidMap: decoded }).html,
      new RegExp(`src="cid:${MINT_A}"`),
    );
    const literal = new Map([['%78', MINT_A]]);
    assert.equal(sanitizeQuoteHtml('<img src="cid:%78">', { mode: 'map', cidMap: literal }).html, '');
  });
});

describe('sanitizeQuoteHtml obfuscation properties', () => {
  // Every spelling a browser resolves to the same reference. The normalization these rely
  // on mirrors the sanitizer's own; if the two ever drift, these are what fail.
  const SPELLINGS = [
    'cid:x',
    'CID:x',
    'Cid:x',
    'c id:x',
    'cid :x',
    ' cid:x ',
    'cid\t:x',
    'c<!--z-->id:x',
    'cid\x00:x',
    'cid\x1f:x',
    'cid:%78',
  ];

  for (const spelling of SPELLINGS) {
    it(`maps ${JSON.stringify(spelling)} when the reference resolves`, () => {
      const out = sanitizeQuoteHtml(`<img src="${spelling}">`, {
        mode: 'map',
        cidMap: new Map([['x', MINT_A]]),
      });
      assert.deepEqual(out.refs, ['x']);
      assert.match(out.html, new RegExp(`src="cid:${MINT_A}"`));
    });

    it(`drops ${JSON.stringify(spelling)} when nothing resolves it`, () => {
      const out = sanitizeQuoteHtml(`<img src="${spelling}">`, { mode: 'map', cidMap: new Map() });
      assert.equal(out.html, '');
      assert.equal(out.droppedCidImages, 1);
      assert.ok(!/cid/i.test(out.html));
    });
  }

  it('sees a reference spelled with a character entity', () => {
    // The parser decodes entities before the transform sees the attribute.
    assert.deepEqual(sanitizeQuoteHtml('<img src="cid&#9;:x">', { mode: 'collect' }).refs, ['x']);
    assert.deepEqual(sanitizeQuoteHtml('<img src="c&#105;d:x">', { mode: 'collect' }).refs, ['x']);
  });

  it('never lets an embedded-image scheme reach a link', () => {
    // The per-tag scheme list admits it on <img> alone. A global entry would apply to
    // href, cite, poster and every other URL-bearing attribute the sanitizer knows.
    const cidMap = new Map([['x', MINT_A]]);
    const out = sanitizeQuoteHtml('<a href="cid:x">click</a>', { mode: 'map', cidMap });
    assert.equal(out.html, '<a>click</a>');
    assert.ok(!/cid:/i.test(out.html));
  });

  it('keeps an ordinary link working', () => {
    const out = sanitizeQuoteHtml('<a href="https://example.com/">x</a>', { mode: 'map' });
    assert.equal(out.html, '<a href="https://example.com/">x</a>');
  });

  it('never lets an unmapped reference survive in any element', () => {
    const html =
      '<div><a href="cid:x">a</a><img src="cid:x"><p>cid:x</p></div>';
    const out = sanitizeQuoteHtml(html, { mode: 'map', cidMap: new Map() });
    assert.ok(!/src="cid/i.test(out.html));
    assert.ok(!/href="cid/i.test(out.html));
  });
});

// The sanitizer calls a tag transform with exactly (tagName, attribs) — that is both
// what sanitize-html's Transformer type declares and what its traversal passes — so
// these call it the same way. Anything a test added beyond those two would be
// discarded before the transform saw it, and would prove nothing about the hook.
describe('collectImgCidRefs', () => {
  it('reports a reference and returns the tag untouched', () => {
    const seen: string[] = [];
    const transform = collectImgCidRefs({ onCidRef: (key) => seen.push(key) });
    const attribs = { src: 'cid:logo', alt: 'a' };
    const out = transform('img', attribs);
    assert.deepEqual(seen, ['logo']);
    assert.deepEqual(out, { tagName: 'img', attribs });
    // The collecting pass must hand the sanitizer back the SAME attributes object,
    // not a copy of it: its whole contract is that the html it produces is
    // byte-for-byte what the sanitizer would have emitted with no transform at all.
    assert.equal(out.attribs, attribs);
  });

  it('reports a data: image separately and no reference', () => {
    const refs: string[] = [];
    let data = 0;
    const transform = collectImgCidRefs({
      onCidRef: (key) => refs.push(key),
      onDataImage: () => { data++; },
    });
    transform('img', { src: 'data:image/png;base64,AA' });
    assert.deepEqual(refs, []);
    assert.equal(data, 1);
  });

  it('reports nothing for a remote, relative or missing src', () => {
    const refs: string[] = [];
    let data = 0;
    const transform = collectImgCidRefs({
      onCidRef: (key) => refs.push(key),
      onDataImage: () => { data++; },
    });
    // Remote, root-relative, and an <img> with no src at all.
    const cases: Record<string, string>[] = [{ src: 'https://a/b.png' }, { src: '/b.png' }, {}];
    for (const attribs of cases) {
      transform('img', attribs);
    }
    assert.deepEqual(refs, []);
    assert.equal(data, 0);
  });
});

describe('extractCidRefs', () => {
  it('finds references the <img> collector cannot see', () => {
    const html =
      '<div style="background:url(cid:bg.png)">' +
      '<p>the logo is cid:logo, see below</p>' +
      '<video poster="cid:frame"></video></div>';
    assert.deepEqual(extractCidRefs(html), ['bg.png', 'logo', 'frame']);
  });

  it('finds a reference written with a character entity', () => {
    assert.deepEqual(extractCidRefs('&#99;id:logo'), ['logo']);
    assert.deepEqual(extractCidRefs('cid&#58;logo'), ['logo']);
  });

  it('decodes entities only once, so an escaped entity stays text', () => {
    // `&amp;#99;` is the literal text `&#99;`, not the letter c, so nothing here is a
    // reference and a value could not be spelled two ways that both match.
    assert.deepEqual(extractCidRefs('&amp;#99;id:logo'), []);
  });

  it('strips sentence punctuation from the end but never from the middle', () => {
    assert.deepEqual(extractCidRefs('see cid:logo.'), ['logo']);
    assert.deepEqual(extractCidRefs('see cid:a:b, then'), ['a:b']);
    assert.deepEqual(extractCidRefs('see cid:a;b!'), ['a;b']);
  });

  it('decodes a percent escape once, matching the precise collector', () => {
    assert.deepEqual(extractCidRefs('<p>cid:%78</p>'), ['x']);
  });

  it('reports each value once, in first-seen order', () => {
    assert.deepEqual(extractCidRefs('cid:b cid:a cid:b'), ['b', 'a']);
  });

  it('is case-insensitive on the scheme', () => {
    assert.deepEqual(extractCidRefs('CID:logo'), ['logo']);
  });

  it('reports nothing for html with no reference, or no html', () => {
    assert.deepEqual(extractCidRefs('<p>hello</p>'), []);
    assert.deepEqual(extractCidRefs(''), []);
    assert.deepEqual(extractCidRefs(null), []);
    assert.deepEqual(extractCidRefs(undefined), []);
  });

  it('matches prose that merely mentions a reference, which is why a hit only warns', () => {
    assert.deepEqual(extractCidRefs('<p>Use cid:yourname to embed an image.</p>'), ['yourname']);
  });
});

describe('resolveCidRefs', () => {
  const logo = { cid: 'logo', blobId: 'B1', type: 'image/png', name: 'logo.png', size: 100 };
  const icon = { cid: 'icon', blobId: 'B2', type: 'image/gif', name: 'icon.gif', size: 50 };

  it('reports the distinct references in first-seen order', () => {
    const out = resolveCidRefs(['logo', 'icon', 'logo'], [logo, icon]);
    assert.deepEqual(out.distinctRefs, ['logo', 'icon']);
  });

  it('matches a Content-ID literally, never by a decoded form', () => {
    const out = resolveCidRefs(['x'], [{ cid: '%78', blobId: 'B1', type: 'image/png' }]);
    assert.deepEqual(out.unresolvedRefs, ['x']);
    assert.deepEqual(out.resolvedParts, []);
  });

  it('counts a reference that embeds separately from one that merely resolves', () => {
    const noBlob = { cid: 'logo', type: 'image/png' };
    const out = resolveCidRefs(['logo', 'icon', 'gone'], [noBlob, icon]);
    assert.deepEqual(out.resolvedParts, [noBlob, icon]);
    assert.deepEqual(out.unresolvedRefs, ['gone']);
    // Only `icon` could really be carried, so only it makes an image-only body quotable.
    assert.deepEqual(out.embeddableRefs, ['icon']);
  });

  it('treats a shared Content-ID as embeddable by nothing, and resolves to every colliding part', () => {
    const one = { cid: 'dup', blobId: 'B1', type: 'image/png' };
    const two = { cid: 'dup', blobId: 'B2', type: 'image/png' };
    const out = resolveCidRefs(['dup'], [one, two]);
    assert.deepEqual([...out.ambiguousCids], ['dup']);
    assert.deepEqual(out.resolvedParts, [one, two]);
    assert.deepEqual(out.embeddableRefs, []);
  });

  it('ignores a part with no Content-ID at all', () => {
    const out = resolveCidRefs(['logo'], [{ blobId: 'B9', type: 'image/png' }, logo]);
    assert.deepEqual(out.resolvedParts, [logo]);
  });

  it('handles empty inputs', () => {
    const out = resolveCidRefs([], []);
    assert.deepEqual(out.distinctRefs, []);
    assert.deepEqual(out.resolvedParts, []);
    assert.deepEqual(out.embeddableRefs, []);
  });
});

describe('buildCidMap', () => {
  const logoPart = { cid: 'logo', blobId: 'B1', type: 'image/png', name: 'logo.png', size: 100 };
  const iconPart = { cid: 'icon', blobId: 'B2', type: 'image/gif', name: 'icon.gif', size: 50 };

  it('mints an identifier and a part for a reference with no stored match', () => {
    const out = buildCidMap({
      refs: ['logo'],
      sourceParts: [logoPart],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_0]]);
    assert.deepEqual(out.minted, [
      { blobId: 'B1', type: 'image/png', name: 'logo.png', cid: MINT_0, disposition: 'inline' },
    ]);
    assert.equal(out.mappings[0].reused, false);
    assert.equal(out.mappings[0].source, logoPart);
    assert.equal(out.resolvedPartCount, 1);
  });

  it('reports a reference that matched no part, without minting for it', () => {
    const out = buildCidMap({ refs: ['gone'], sourceParts: [logoPart], mint: sequentialMint() });
    assert.deepEqual(out.unresolvedRefs, ['gone']);
    assert.deepEqual(out.minted, []);
    assert.equal(out.cidMap.size, 0);
    assert.equal(out.resolvedPartCount, 0);
  });

  it('matches a part by its Content-ID literally, never by a decoded form', () => {
    const percent = { cid: '%78', blobId: 'B9', type: 'image/png' };
    assert.equal(buildCidMap({ refs: ['x'], sourceParts: [percent] }).cidMap.size, 0);
    assert.equal(buildCidMap({ refs: ['%78'], sourceParts: [percent] }).cidMap.size, 1);
  });

  it('reuses a surviving part with the same blob, keeping its stored identifier', () => {
    const survivor = { cid: MINT_A, blobId: 'B1', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['logo'],
      sourceParts: [logoPart],
      survivors: [survivor],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_A]]);
    assert.deepEqual(out.minted, []);
    assert.deepEqual(out.reusedCids, [MINT_A]);
    assert.equal(out.mappings[0].reused, true);
    assert.deepEqual(out.unclaimedSurvivors, []);
  });

  it('mints afresh when no survivor carries the blob', () => {
    const survivor = { cid: MINT_A, blobId: 'OTHER', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['logo'],
      sourceParts: [logoPart],
      survivors: [survivor],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_0]]);
    assert.equal(out.minted.length, 1);
    assert.deepEqual(out.unclaimedSurvivors, [survivor]);
  });

  it('ignores a survivor whose identifier is not one this server minted', () => {
    // Reusing a foreign identifier would write someone else's value into a body this
    // server composed, and later classify that part as server-managed.
    const foreign = { cid: 'logo@sender.example', blobId: 'B1', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['logo'],
      sourceParts: [logoPart],
      survivors: [foreign],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_0]]);
    assert.deepEqual(out.unclaimedSurvivors, []);
  });

  it('claims each survivor at most once, so two references over one blob take two', () => {
    const a = { cid: 'a', blobId: 'SHARED', type: 'image/png' };
    const b = { cid: 'b', blobId: 'SHARED', type: 'image/png' };
    const s1 = { cid: MINT_A, blobId: 'SHARED', type: 'image/png', disposition: 'inline' };
    const s2 = { cid: MINT_B, blobId: 'SHARED', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['a', 'b'],
      sourceParts: [a, b],
      survivors: [s1, s2],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['a', MINT_A], ['b', MINT_B]]);
    assert.deepEqual(out.minted, []);
    assert.deepEqual(out.unclaimedSurvivors, []);
  });

  it('mints for the second of two references over one blob when only one survivor exists', () => {
    // Never a silent collapse of two images into one part.
    const a = { cid: 'a', blobId: 'SHARED', type: 'image/png' };
    const b = { cid: 'b', blobId: 'SHARED', type: 'image/png' };
    const s1 = { cid: MINT_A, blobId: 'SHARED', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['a', 'b'],
      sourceParts: [a, b],
      survivors: [s1],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['a', MINT_A], ['b', MINT_0]]);
    assert.equal(out.minted.length, 1);
    assert.deepEqual(out.unclaimedSurvivors, []);
  });

  it('treats a repeated reference as one, so it cannot claim two survivors', () => {
    // The collecting pass already reports distinct keys, but a repeat here would take a
    // second survivor for the same image and leave that part attached with nothing
    // pointing at it — and the closure check covers only freshly minted identifiers, so
    // nothing downstream would catch it.
    const s1 = { cid: MINT_A, blobId: 'B1', type: 'image/png', disposition: 'inline' };
    const s2 = { cid: MINT_B, blobId: 'B1', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['logo', 'logo'],
      sourceParts: [logoPart],
      survivors: [s1, s2],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_A]]);
    assert.equal(out.mappings.length, 1);
    assert.deepEqual(out.unclaimedSurvivors, [s2]);
    assert.deepEqual(out.minted, []);
  });

  it('reports a repeated unresolved reference once', () => {
    const out = buildCidMap({ refs: ['gone', 'gone'], sourceParts: [logoPart] });
    assert.deepEqual(out.unresolvedRefs, ['gone']);
  });

  it('claims survivors in stored order', () => {
    const s1 = { cid: MINT_A, blobId: 'X', type: 'image/png', disposition: 'inline' };
    const s2 = { cid: MINT_B, blobId: 'B1', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({
      refs: ['logo'],
      sourceParts: [logoPart],
      survivors: [s1, s2],
      mint: sequentialMint(),
    });
    assert.deepEqual([...out.cidMap], [['logo', MINT_B]]);
    assert.deepEqual(out.unclaimedSurvivors, [s1]);
  });

  it('reports survivors nothing referenced, in stored order', () => {
    const s1 = { cid: MINT_A, blobId: 'OLD1', type: 'image/png', disposition: 'inline' };
    const s2 = { cid: MINT_B, blobId: 'OLD2', type: 'image/png', disposition: 'inline' };
    const out = buildCidMap({ refs: [], sourceParts: [], survivors: [s1, s2] });
    assert.deepEqual(out.unclaimedSurvivors, [s1, s2]);
  });

  it('treats a Content-ID shared by two parts as unusable rather than picking one', () => {
    const one = { cid: 'dup', blobId: 'B1', type: 'image/png' };
    const two = { cid: 'dup', blobId: 'B2', type: 'image/png' };
    const out = buildCidMap({ refs: ['dup'], sourceParts: [one, two], mint: sequentialMint() });
    assert.equal(out.cidMap.size, 0);
    assert.deepEqual(out.minted, []);
    // BOTH colliding parts, not just the first: the reader loses two images, and a report
    // naming one of them would understate what the collision cost.
    assert.deepEqual(out.unembeddableParts, [one, two]);
    assert.deepEqual(out.unresolvedRefs, []);
    assert.equal(out.resolvedPartCount, 2);
  });

  it('treats a part with no blob as unusable, since it cannot be re-referenced', () => {
    const noBlob = { cid: 'logo', type: 'image/png' };
    const out = buildCidMap({ refs: ['logo'], sourceParts: [noBlob], mint: sequentialMint() });
    assert.equal(out.cidMap.size, 0);
    assert.deepEqual(out.unembeddableParts, [noBlob]);
    assert.equal(out.resolvedPartCount, 1);
  });

  it('counts references with no part apart from parts that could not be embedded', () => {
    // The two are disjoint and must never be summed, or one lost image reads as two.
    const noBlob = { cid: 'logo', type: 'image/png' };
    const out = buildCidMap({ refs: ['logo', 'gone'], sourceParts: [noBlob, iconPart] });
    assert.deepEqual(out.unresolvedRefs, ['gone']);
    assert.equal(out.unembeddableParts.length, 1);
    assert.equal(out.resolvedPartCount, 1);
  });

  it('keeps the minted parts on their own channel and out of the source parts', () => {
    const sourceParts = [logoPart, iconPart];
    const out = buildCidMap({ refs: ['logo'], sourceParts, mint: sequentialMint() });
    assert.deepEqual(sourceParts, [logoPart, iconPart]);
    assert.equal(out.minted.length, 1);
    assert.notEqual(out.minted[0] as any, logoPart as any);
  });

  it('omits a name a part does not have', () => {
    const bare = { cid: 'bare', blobId: 'B7', type: 'image/png' };
    const out = buildCidMap({ refs: ['bare'], sourceParts: [bare], mint: sequentialMint() });
    assert.deepEqual(out.minted, [
      { blobId: 'B7', type: 'image/png', cid: MINT_0, disposition: 'inline' },
    ]);
  });

  it('carries only what the sender declared an image, whatever the <img> claims', () => {
    const notAnImage = { cid: 'doc', blobId: 'B8', type: 'application/pdf', name: 'doc.pdf' };
    const typeless = { cid: 'bare', blobId: 'B9' };
    const out = buildCidMap({
      refs: ['doc', 'bare'], sourceParts: [notAnImage, typeless], mint: sequentialMint(),
    });
    assert.deepEqual(out.minted, []);
    // Resolved, so never reported as a reference that matched nothing — the parts are there,
    // they are just not ones this server pulls into a body it composes.
    assert.deepEqual(out.unresolvedRefs, []);
    assert.deepEqual(out.unembeddableParts, [notAnImage, typeless]);
  });

  it('ignores content-type parameters and case when deciding a part is an image', () => {
    const odd = { cid: 'logo', blobId: 'B1', type: 'IMAGE/PNG; name=x' };
    const out = buildCidMap({ refs: ['logo'], sourceParts: [odd], mint: sequentialMint() });
    assert.deepEqual(out.minted, [
      { blobId: 'B1', type: 'IMAGE/PNG; name=x', cid: MINT_0, disposition: 'inline' },
    ]);
  });

  it('emits every minted part as an inline disposition', () => {
    const out = buildCidMap({
      refs: ['logo', 'icon'],
      sourceParts: [logoPart, iconPart],
      mint: sequentialMint(),
    });
    assert.deepEqual(out.minted.map((p) => p.disposition), ['inline', 'inline']);
    assert.deepEqual(out.minted.map((p) => p.cid), [MINT_0, MINT_1]);
  });

  it('handles an empty call', () => {
    const out = buildCidMap({ refs: [], sourceParts: [] });
    assert.equal(out.cidMap.size, 0);
    assert.deepEqual(out.minted, []);
    assert.deepEqual(out.mappings, []);
    assert.equal(out.resolvedPartCount, 0);
  });
});

describe('reconcileInlineParts', () => {
  const stored = (over: Record<string, unknown>) => ({
    blobId: 'B', type: 'image/png', ...over,
  });

  it('keeps a part the final bodies still reference', () => {
    const part = stored({ cid: MINT_A, disposition: 'inline' });
    const out = reconcileInlineParts({ storedParts: [part], referencedCids: [MINT_A] });
    assert.deepEqual(out.parts, [{ part, action: 'kept' }]);
    assert.deepEqual(out.removed, []);
  });

  it('removes an unreferenced part this server minted', () => {
    // This server put it there to display an image in a body that no longer shows it, so
    // leaving it behind would attach a file the user never asked to send.
    const part = stored({ cid: MINT_A, disposition: 'inline' });
    const out = reconcileInlineParts({ storedParts: [part], referencedCids: [] });
    assert.deepEqual(out.removed, [part]);
    assert.deepEqual(out.kept, []);
  });

  it('degrades an unreferenced inline part carrying someone else\'s identifier', () => {
    const part = stored({ cid: 'logo@sender.example', disposition: 'inline' });
    const out = reconcileInlineParts({ storedParts: [part], referencedCids: [] });
    assert.deepEqual(out.degraded, [part]);
    assert.deepEqual(out.removed, []);
  });

  it('leaves an ordinary unreferenced attachment alone', () => {
    const part = stored({ name: 'a.pdf', type: 'application/pdf', disposition: 'attachment' });
    const out = reconcileInlineParts({ storedParts: [part], referencedCids: [] });
    assert.deepEqual(out.kept, [part]);
  });

  it('never degrades an attachment-labelled part that happens to carry an identifier', () => {
    const part = stored({ cid: 'logo@sender.example', disposition: 'attachment' });
    assert.deepEqual(reconcileInlineParts({ storedParts: [part], referencedCids: [] }).kept, [part]);
  });

  it('recognizes an inline disposition regardless of case or padding', () => {
    const part = stored({ cid: 'x@y', disposition: ' INLINE ' });
    assert.deepEqual(reconcileInlineParts({ storedParts: [part], referencedCids: [] }).degraded, [part]);
  });

  it('references nothing when no html body ships', () => {
    // With no html there is nothing to display an embedded image, and the mail server
    // rejects an inline disposition outright.
    const ours = stored({ cid: MINT_A, disposition: 'inline' });
    const theirs = stored({ cid: 'x@y', disposition: 'inline' });
    const out = reconcileInlineParts({
      storedParts: [ours, theirs],
      referencedCids: [MINT_A, 'x@y'],
      htmlShips: false,
    });
    assert.deepEqual(out.removed, [ours]);
    assert.deepEqual(out.degraded, [theirs]);
  });

  it('keeps the stored order', () => {
    const a = stored({ cid: MINT_A, disposition: 'inline' });
    const b = stored({ name: 'b.pdf', disposition: 'attachment' });
    const c = stored({ cid: 'c@y', disposition: 'inline' });
    const out = reconcileInlineParts({ storedParts: [a, b, c], referencedCids: [MINT_A] });
    assert.deepEqual(out.parts.map((p) => p.action), ['kept', 'kept', 'degraded']);
  });

  it('passes the minted parts through on their own channel, untouched', () => {
    const minted = [{ blobId: 'N1', type: 'image/png', cid: MINT_B, disposition: 'inline' as const }];
    const part = stored({ name: 'a.pdf', disposition: 'attachment' });
    const out = reconcileInlineParts({ storedParts: [part], referencedCids: [MINT_B], minted });
    assert.deepEqual(out.minted, minted);
    assert.deepEqual(out.kept, [part]);
    assert.ok(!out.kept.includes(minted[0] as any));
  });

  it('handles an empty draft', () => {
    const out = reconcileInlineParts({ storedParts: [], referencedCids: [] });
    assert.deepEqual(out.parts, []);
    assert.deepEqual(out.minted, []);
  });
});

describe('checkInlineClosure', () => {
  it('passes when every reference resolves and every minted part is referenced', () => {
    assert.doesNotThrow(() =>
      checkInlineClosure({
        htmlBodies: [`<div><img src="cid:${MINT_A}"></div>`],
        finalPartCids: [MINT_A],
        attachedMintedCids: [MINT_A],
      }),
    );
  });

  it('fails when a body references an image no part supplies', () => {
    assert.throws(
      () => checkInlineClosure({
        htmlBodies: [`<img src="cid:${MINT_A}">`],
        finalPartCids: [],
      }),
      (err: any) => err instanceof InlineClosureError &&
        err.name === 'InlineClosureError' &&
        /references embedded image/.test(err.message),
    );
  });

  it('fails when a part this call minted is referenced by nothing', () => {
    assert.throws(
      () => checkInlineClosure({
        htmlBodies: ['<p>no images here</p>'],
        finalPartCids: [MINT_A],
        attachedMintedCids: [MINT_A],
      }),
      (err: any) => err instanceof InlineClosureError && /was attached/.test(err.message),
    );
  });

  it('skips entirely when asked to', () => {
    assert.doesNotThrow(() =>
      checkInlineClosure({
        htmlBodies: [`<img src="cid:${MINT_A}">`],
        finalPartCids: [],
        skip: true,
      }),
    );
  });

  it('ignores a part the caller attached and never referenced', () => {
    // An ordinary attachment is not a loose end; only this call's own minted parts are.
    assert.doesNotThrow(() =>
      checkInlineClosure({
        htmlBodies: ['<p>hi</p>'],
        finalPartCids: ['logo@sender.example'],
        attachedMintedCids: [],
      }),
    );
  });

  it('reads only the bodies it was given, not any the call carried through', () => {
    assert.doesNotThrow(() => checkInlineClosure({ htmlBodies: [], finalPartCids: [] }));
    assert.doesNotThrow(() => checkInlineClosure({}));
  });

  it('sees a reference however it is spelled', () => {
    assert.throws(
      () => checkInlineClosure({ htmlBodies: ['<img src="c id:missing">'], finalPartCids: [] }),
      InlineClosureError,
    );
  });

  it('renders a hostile identifier as bounded, quoted data', () => {
    const hostile = `${'z'.repeat(200)}`;
    assert.throws(
      () => checkInlineClosure({ htmlBodies: [`<img src="cid:${hostile}">`], finalPartCids: [] }),
      (err: any) => err.message.includes(`"${'z'.repeat(64)}…"`),
    );
  });
});
