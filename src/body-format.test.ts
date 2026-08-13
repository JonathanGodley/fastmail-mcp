import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { isBlank, htmlToText, htmlHasVisibleContent, normalizeBodies, buildBodyParts, assertBodyInputs } from './body-format.js';
import { InvalidInputError } from './coerce.js';

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

describe('assertBodyInputs — body type (#62)', () => {
  it('rejects a present non-string body, naming the parameter', () => {
    assert.throws(() => assertBodyInputs({ textBody: 42 }), (e: any) =>
      e instanceof InvalidInputError && /textBody must be a string; received number/.test(e.message));
    assert.throws(() => assertBodyInputs({ htmlBody: 42 }), /htmlBody must be a string; received number/);
    assert.throws(() => assertBodyInputs({ htmlBody: { p: 'hi' } }), /htmlBody must be a string; received object/);
    assert.throws(() => assertBodyInputs({ htmlBody: ['<p>hi</p>'] }), /htmlBody must be a string; received array/);
    assert.throws(() => assertBodyInputs({ textBody: true }), /textBody must be a string; received boolean/);
  });

  it('treats an omitted body and an explicit null alike (both mean "not provided")', () => {
    assert.doesNotThrow(() => assertBodyInputs({}));
    assert.doesNotThrow(() => assertBodyInputs({ textBody: null, htmlBody: null }));
    assert.doesNotThrow(() => assertBodyInputs(undefined as any));
  });

  it('accepts an empty string (emptiness is handled by the body rules, not this guard)', () => {
    assert.doesNotThrow(() => assertBodyInputs({ textBody: '', htmlBody: '' }));
  });
});

describe('assertBodyInputs — HTML-escaped htmlBody (#71/#77)', () => {
  it('rejects a body that is entirely escaped markup', () => {
    assert.throws(
      () => assertBodyInputs({ htmlBody: '&lt;p&gt;Hi there&lt;/p&gt;\n&lt;p&gt;Regards&lt;br&gt;Someone&lt;/p&gt;' }),
      (e: any) => e instanceof InvalidInputError && /htmlBody appears to be HTML-escaped/.test(e.message),
    );
    assert.throws(() => assertBodyInputs({ htmlBody: '&lt;br&gt;' }), /appears to be HTML-escaped/);
    assert.throws(() => assertBodyInputs({ htmlBody: 'Hi&lt;br&gt;there' }), /appears to be HTML-escaped/);
  });

  it('names the fix in the error (real markup, or textBody)', () => {
    assert.throws(() => assertBodyInputs({ htmlBody: '&lt;p&gt;x&lt;/p&gt;' }), /Pass real markup \(<p>\.\.\.<\/p>\), or use textBody/);
  });

  it('accepts escaped markup that sits inside real elements', () => {
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<pre>&lt;p&gt;</pre>' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<p>To bold text write &lt;b&gt;word&lt;/b&gt;.</p>' }));
  });

  it('accepts real markup that merely contains entities', () => {
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<p>Tom &amp; Jerry</p>' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<p>if a &lt; b then c &gt; d</p>' }));
  });

  it('accepts escaped angle brackets that are not tag-shaped, even with no real elements', () => {
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'a &lt; b and c &gt; d' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'no markup at all' }));
  });

  // The escaped test only fires on a body with no real markup, i.e. on prose — where
  // escaped angle brackets are ordinary content. These are real messages, and a looser
  // "&lt; anything &gt;" test rejected every one of them.
  it('accepts a tag-less body whose escaped brackets are placeholders or addresses', () => {
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'Hi &lt;name&gt;, see attached.' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'mail me at &lt;a@b.com&gt;' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'Please reply with &lt;approve&gt; or &lt;reject&gt;.' }));
  });

  it('still rejects escaped real element tags in every delimiter form', () => {
    for (const html of [
      '&lt;p&gt;Hi&lt;/p&gt;',
      '&lt;br&gt;',
      '&lt;br/&gt;',
      '&lt;div class=&quot;x&quot;&gt;body&lt;/div&gt;',
      '&lt;a href=&quot;http://x&quot;&gt;link&lt;/a&gt;',
      '&lt;H1&gt;Title&lt;/H1&gt;',
    ]) {
      assert.throws(() => assertBodyInputs({ htmlBody: html }), /appears to be HTML-escaped/, `should reject: ${html}`);
    }
  });

  it('does not treat an escaped element name without a tag delimiter as markup', () => {
    // "a" is an element name, but "a@b.com" and "ary" are not tags — the lookahead
    // requires whitespace, "/", or "&gt;" straight after the name.
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'contact &lt;a@b.com&gt; today' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'the &lt;primary&gt; option' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: 'sizes &lt;big-item&gt; here' }));
  });

  it('is htmlBody-only: literal tags in a plain-text body are legitimate content', () => {
    assert.doesNotThrow(() => assertBodyInputs({ textBody: '&lt;p&gt;Hi&lt;/p&gt;' }));
    assert.doesNotThrow(() => assertBodyInputs({ textBody: 'Use <p> to open a paragraph' }));
  });
});

describe('assertBodyInputs — CDATA (#78)', () => {
  it('rejects a CDATA-wrapped htmlBody', () => {
    assert.throws(
      () => assertBodyInputs({ htmlBody: '<![CDATA[<p>Hi Lachlan,</p><p>Just checking in</p>]]>' }),
      (e: any) => e instanceof InvalidInputError && /htmlBody contains a CDATA section/.test(e.message),
    );
  });

  it('rejects a CDATA section anywhere in htmlBody, not just at the start', () => {
    // A mid-body section swallows its contents just as completely: html-to-text renders
    // '<p>Before</p><![CDATA[<p>gone</p>]]><p>After</p>' as "Before\n\nAfter".
    assert.throws(() => assertBodyInputs({ htmlBody: '<p>Before</p><![CDATA[<p>gone</p>]]><p>After</p>' }), /contains a CDATA section/);
    assert.throws(() => assertBodyInputs({ htmlBody: '<p>Before</p><![cdata[<p>gone</p>' }), /contains a CDATA section/);
  });

  it('rejects a CDATA-wrapped textBody (leading whitespace tolerated)', () => {
    assert.throws(() => assertBodyInputs({ textBody: '<![CDATA[Hi there]]>' }), /textBody is wrapped in a CDATA section/);
    assert.throws(() => assertBodyInputs({ textBody: '\n  <![CDATA[Hi there]]>' }), /textBody is wrapped in a CDATA section/);
  });

  it('leaves an embedded CDATA token in a plain-text body alone (never markup-parsed)', () => {
    assert.doesNotThrow(() => assertBodyInputs({ textBody: 'The config uses <![CDATA[ raw ]]> around the query.' }));
  });

  it('leaves a bare ]]> alone in both formats (inert without an opening token)', () => {
    assert.doesNotThrow(() => assertBodyInputs({ textBody: 'End the section with ]]> and save.' }));
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<p>End the section with ]]> and save.</p>' }));
  });

  it('accepts an escaped CDATA token in htmlBody (the correct way to show one)', () => {
    assert.doesNotThrow(() => assertBodyInputs({ htmlBody: '<pre>&lt;![CDATA[ x ]]&gt;</pre>' }));
  });
});

// The html-to-text behaviour the CDATA rule is built on, pinned so a dependency
// upgrade that changed it would surface here rather than silently make the guard
// look arbitrary.
describe('CDATA in HTML — the failure the guard prevents', () => {
  it('swallows everything from the opening token to the next ]]> (or to the end)', () => {
    assert.equal(htmlToText('<![CDATA[<p>Hi there</p>]]>'), '');
    assert.equal(htmlToText('<p>Before</p><![CDATA[<p>gone</p>]]><p>After</p>'), 'Before\n\nAfter');
    assert.equal(htmlToText('<p>Before</p><![CDATA[<p>rest of message</p>'), 'Before');
  });
  it('leaves the new message out of a reply, keeping only the quote', () => {
    const derived = htmlToText('<![CDATA[<p>new msg</p>]]><blockquote type="cite"><p>quoted original</p></blockquote>');
    assert.doesNotMatch(derived, /new msg/);
    assert.match(derived, /quoted original/);
  });
  it('passes a bare ]]> through as ordinary text', () => {
    assert.match(htmlToText('<p>Some XML uses ]]> to end</p>'), /Some XML uses \]\]> to end/);
  });
});
