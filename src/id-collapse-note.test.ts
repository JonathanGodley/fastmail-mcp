import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { buildIdCollapseNote } from './id-collapse-note.js';

describe('buildIdCollapseNote', () => {
  it('is silent when every submitted id is distinct', () => {
    assert.equal(buildIdCollapseNote(['e1', 'e2', 'e3']), '');
  });

  it('is silent on an empty list', () => {
    assert.equal(buildIdCollapseNote([]), '');
  });

  it('names duplication as the cause, and says nothing was skipped', () => {
    const text = buildIdCollapseNote(['e1', 'e1']);
    assert.match(text, /duplicates collapsed/);
    assert.match(text, /nothing was skipped/);
  });

  it('pluralises the submitted-id count', () => {
    assert.match(buildIdCollapseNote(['e1', 'e1']), /^2 ids were given/);
  });

  it('pluralises the distinct-email count', () => {
    const text = buildIdCollapseNote(['e1', 'e1', 'e2']);
    assert.match(text, /collapsed them to 2 distinct emails/);
  });

  it('uses the singular form for one distinct email', () => {
    const text = buildIdCollapseNote(['e1', 'e1']);
    assert.match(text, /collapsed them to 1 distinct email;/);
  });

  it('states the exact sentence for a simple duplicate', () => {
    assert.equal(
      buildIdCollapseNote(['e1', 'e1']),
      '2 ids were given, but duplicates collapsed them to 1 distinct email; nothing was skipped.',
    );
  });
});
