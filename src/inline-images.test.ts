import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  buildUnionParts,
  cidKey,
  decodeCidSrc,
  describePart,
  sanitizeDownloadFilename,
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
