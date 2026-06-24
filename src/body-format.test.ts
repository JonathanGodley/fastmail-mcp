import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { isBlank, htmlToText, htmlHasVisibleContent, normalizeBodies, buildBodyParts } from './body-format.js';

describe('isBlank', () => {
  it('treats empty / whitespace / zero-width-only as blank', () => {
    assert.equal(isBlank(''), true);
    assert.equal(isBlank('   '), true);
    assert.equal(isBlank(undefined), true);
    assert.equal(isBlank(null), true);
    assert.equal(isBlank('\u200B\u200C\uFEFF\u00AD'), true); // zero-width only
    assert.equal(isBlank(' \u200B \n '), true);
  });
  it('treats real content as non-blank', () => {
    assert.equal(isBlank('x'), false);
    assert.equal(isBlank('  hi  '), false);
  });
});

describe('htmlToText', () => {
  it('renders lists as markers', () => {
    assert.match(htmlToText('<ul><li>one</li><li>two</li></ul>'), /\*\s*one/);
  });
  it('renders links readably and decodes entities', () => {
    const t = htmlToText('<p>A &amp; B</p><a href="http://example.com">click</a>');
    assert.match(t, /A & B/);
    assert.match(t, /click/);
  });
  it('prefixes blockquotes and preserves nesting depth', () => {
    const t = htmlToText('<blockquote>outer<blockquote>inner</blockquote></blockquote>');
    assert.match(t, /> outer/);
    assert.match(t, /> > inner/); // nested → per-level depth, NOT flattened
  });
  it('leaves no raw tags', () => {
    const t = htmlToText('<div><p>Hello <b>world</b></p><span>x</span></div>');
    assert.equal(/[<>]/.test(t), false);
  });
  it('derives text from img alt, and nothing from a no-alt img', () => {
    assert.match(htmlToText('<p>see <img src="c.png" alt="the chart"></p>'), /see the chart/);
    assert.equal(isBlank(htmlToText('<div><img src="banner.jpg"></div>')), true);
  });
  it('never throws on malformed input and always returns a string', () => {
    assert.equal(typeof htmlToText('<<<>>></ '), 'string');
    assert.equal(typeof htmlToText('<img alt='), 'string');
  });
});

describe('htmlHasVisibleContent', () => {
  it('is true for any visible media element', () => {
    assert.equal(htmlHasVisibleContent('<div><img src="x.png"></div>'), true);
    assert.equal(htmlHasVisibleContent('<div style="background-image: url(x.png)"></div>'), true);
    assert.equal(htmlHasVisibleContent('<svg><circle/></svg>'), true);
    assert.equal(htmlHasVisibleContent('<video src="v.mp4"></video>'), true);
    assert.equal(htmlHasVisibleContent('<picture><source srcset="x"></picture>'), true);
  });
  it('is true when there is readable text', () => {
    assert.equal(htmlHasVisibleContent('<p>hello</p>'), true);
  });
  it('is false for genuinely empty / invisible markup', () => {
    assert.equal(htmlHasVisibleContent('<p></p>'), false);
    assert.equal(htmlHasVisibleContent('   '), false);
    assert.equal(htmlHasVisibleContent('<p>\u200B</p>'), false);
    assert.equal(htmlHasVisibleContent('<!-- <img src="x"> --><div></div>'), false); // commented-out tag ignored
    assert.equal(htmlHasVisibleContent('<div style="background-image: none"></div>'), false);
  });
});

describe('normalizeBodies', () => {
  it('derives a readable text fallback from html when text is absent', () => {
    const r = normalizeBodies({ htmlBody: '<p>Hello <b>world</b></p>' });
    assert.match(r.textBody!, /Hello world/);
    assert.equal(r.htmlBody, '<p>Hello <b>world</b></p>');
    assert.equal(r.htmlOnly, undefined);
  });
  it('derives text from image alt (no htmlOnly)', () => {
    const r = normalizeBodies({ htmlBody: '<img src="x" alt="Company Logo">' });
    assert.match(r.textBody!, /Company Logo/);
    assert.equal(r.htmlOnly, undefined);
  });
  it('flags htmlOnly for image-only html with no derivable text', () => {
    const r = normalizeBodies({ htmlBody: '<div><img src="banner.jpg"></div>' });
    assert.equal(r.textBody, undefined);
    assert.equal(r.htmlOnly, true);
    assert.equal(r.htmlBody, '<div><img src="banner.jpg"></div>');
  });
  it('flags htmlOnly for zero-width-only html (treated empty)', () => {
    const r = normalizeBodies({ htmlBody: '<p>\u200B\u200C</p>' });
    assert.equal(r.textBody, undefined);
    assert.equal(r.htmlOnly, true);
  });
  it('treats a blank textBody alongside html as absent → derives', () => {
    const r = normalizeBodies({ htmlBody: '<p>x</p>', textBody: '' });
    assert.match(r.textBody!, /x/);
  });
  it('passes text-only through untouched', () => {
    const r = normalizeBodies({ textBody: 'just text' });
    assert.deepEqual(r, { textBody: 'just text' });
  });
  it('passes both bodies through untouched (distinct content preserved)', () => {
    const r = normalizeBodies({ textBody: 'my own text', htmlBody: '<p>different html</p>' });
    assert.deepEqual(r, { textBody: 'my own text', htmlBody: '<p>different html</p>' });
  });
});

describe('buildBodyParts', () => {
  it('shapes both parts + bodyValues keyed by literal text/html partIds', () => {
    const r = buildBodyParts({ textBody: 'T', htmlBody: '<p>H</p>' });
    assert.deepEqual(r.textBody, [{ partId: 'text', type: 'text/plain' }]);
    assert.deepEqual(r.htmlBody, [{ partId: 'html', type: 'text/html' }]);
    assert.deepEqual(r.bodyValues, { text: { value: 'T' }, html: { value: '<p>H</p>' } });
  });
  it('emits only the text part for text-only', () => {
    const r = buildBodyParts({ textBody: 'T' });
    assert.deepEqual(r.textBody, [{ partId: 'text', type: 'text/plain' }]);
    assert.equal(r.htmlBody, undefined);
    assert.deepEqual(r.bodyValues, { text: { value: 'T' } });
  });
  it('drops a blank body (no part emitted)', () => {
    const r = buildBodyParts({ textBody: '   ', htmlBody: '<p>H</p>' });
    assert.equal(r.textBody, undefined);
    assert.deepEqual(r.htmlBody, [{ partId: 'html', type: 'text/html' }]);
    assert.deepEqual(r.bodyValues, { html: { value: '<p>H</p>' } });
  });
});
