import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  assertUnambiguousEntryEdit,
  buildEntryMap,
  contactCardKind,
  isAmbiguousEntryEdit,
  isContactGroupCard,
  nonDefaultContactKind,
  mergeContactName,
  mergeContactNotes,
  mergeEntryMap,
  resolveEntryLabel,
  simplifyEntryMap,
} from './contact-card.js';

// ---------- resolveEntryLabel ----------

// Every case below is a shape read off real cards in a live address book. The two
// label-carrying properties are different kinds of thing and both are in use in the same
// account, so both have to be read.
describe('resolveEntryLabel', () => {
  it('uses the scalar label when it says something', () => {
    assert.equal(resolveEntryLabel({ address: 'a@b.example', label: 'work' }), 'work');
  });

  it('treats an empty scalar label as no label at all', () => {
    // Older imported cards carry `label: ""` on every entry. Reading that as a label would
    // wrap a whole address book in {address, label: ""} objects saying nothing.
    assert.equal(resolveEntryLabel({ address: 'a@b.example', label: '' }), undefined);
    assert.equal(resolveEntryLabel({ address: 'a@b.example', label: '   ' }), undefined);
  });

  it('falls back to a single contexts key', () => {
    // The shape every card written by a recent Fastmail interface carries: a contexts SET
    // and no `label` key at all.
    assert.equal(resolveEntryLabel({ address: 'a@b.example', contexts: { private: true } }), 'private');
  });

  it('prefers a non-empty scalar label over contexts', () => {
    assert.equal(
      resolveEntryLabel({ address: 'a@b.example', label: 'work', contexts: { private: true } }),
      'work',
    );
  });

  it('falls back to contexts when the scalar label is empty', () => {
    assert.equal(
      resolveEntryLabel({ address: 'a@b.example', label: '', contexts: { private: true } }),
      'private',
    );
  });

  it('resolves no label from a multi-key contexts set', () => {
    // Two contexts name no single label, and picking one would be arbitrary.
    assert.equal(resolveEntryLabel({ address: 'a@b.example', contexts: { private: true, work: true } }), undefined);
  });

  it('resolves no label from an empty or absent contexts set', () => {
    assert.equal(resolveEntryLabel({ address: 'a@b.example', contexts: {} }), undefined);
    assert.equal(resolveEntryLabel({ address: 'a@b.example' }), undefined);
  });

  it('never reads the entry-map key as a label', () => {
    // The keys are opaque server ids — 40-char sha1-shaped on older cards, 6-char on newer
    // ones — so an entry under any key still resolves by its own properties.
    const map = { '0dd713ddb7cdcb0fbc17f59321f227fabcddc7da': { address: 'a@b.example' }, ba5ddd: { address: 'c@d.example', label: 'work' } };
    assert.deepEqual(simplifyEntryMap(map, 'address'), ['a@b.example', { address: 'c@d.example', label: 'work' }]);
  });
});

// ---------- simplifyEntryMap ----------

describe('simplifyEntryMap', () => {
  it('emits a bare string for an unlabelled entry and an object for a labelled one', () => {
    const map = {
      k1: { address: 'plain@example.com', pref: 1 },
      k2: { address: 'work@example.com', contexts: { work: true } },
      k3: { address: 'home@example.com', label: 'home' },
    };
    assert.deepEqual(simplifyEntryMap(map, 'address'), [
      'plain@example.com',
      { address: 'work@example.com', label: 'work' },
      { address: 'home@example.com', label: 'home' },
    ]);
  });

  it('keys phones by number', () => {
    assert.deepEqual(simplifyEntryMap({ k: { number: '+1 555 0100', contexts: { mobile: true } } }, 'number'), [
      { number: '+1 555 0100', label: 'mobile' },
    ]);
  });

  it('skips an entry with no value and returns undefined for an empty result', () => {
    assert.equal(simplifyEntryMap({ k: { contexts: { work: true } } }, 'address'), undefined);
    assert.equal(simplifyEntryMap({}, 'address'), undefined);
    assert.equal(simplifyEntryMap(undefined, 'address'), undefined);
  });
});

// ---------- mergeContactName ----------

describe('mergeContactName', () => {
  it('sets the full name and keeps the stored components', () => {
    const stored = { full: 'Ada Lovelace', components: [{ kind: 'given', value: 'Ada' }] };
    assert.deepEqual(mergeContactName(stored, { full: 'Augusta Ada King' }), {
      full: 'Augusta Ada King',
      components: [{ kind: 'given', value: 'Ada' }],
    });
  });

  it('updates one component in place on a card that has no full name', () => {
    const stored = { '@type': 'Name', components: [{ kind: 'given', value: 'Ada' }, { kind: 'surname', value: 'Byron' }] };
    assert.deepEqual(mergeContactName(stored, { surname: 'Lovelace' }), {
      '@type': 'Name',
      components: [{ kind: 'given', value: 'Ada' }, { kind: 'surname', value: 'Lovelace' }],
    });
  });

  it('leaves components of other kinds untouched', () => {
    const stored = { components: [{ kind: 'title', value: 'Dr' }, { kind: 'given', value: 'Ada' }] };
    assert.deepEqual(mergeContactName(stored, { given: 'Augusta' }), {
      components: [{ kind: 'title', value: 'Dr' }, { kind: 'given', value: 'Augusta' }],
    });
  });

  it('appends a component the card did not have', () => {
    assert.deepEqual(mergeContactName({ full: 'Ada' }, { surname: 'Lovelace' }), {
      full: 'Ada',
      components: [{ kind: 'surname', value: 'Lovelace' }],
    });
  });

  it('does not mutate the stored name it was handed', () => {
    // The same object is echoed back to the caller as `previousCard`, so a merge that wrote
    // through it would make the echo report the post-edit state.
    const stored = { full: 'Ada Lovelace', components: [{ kind: 'given', value: 'Ada' }] };
    mergeContactName(stored, { full: 'Augusta', given: 'Augusta' });
    assert.deepEqual(stored, { full: 'Ada Lovelace', components: [{ kind: 'given', value: 'Ada' }] });
  });

  it('builds a name from nothing when the card had none', () => {
    assert.deepEqual(mergeContactName(undefined, { full: 'Ada' }), { full: 'Ada' });
    assert.deepEqual(mergeContactName(undefined, { given: 'Ada' }), { components: [{ kind: 'given', value: 'Ada' }] });
  });
});

// ---------- mergeContactNotes ----------

describe('mergeContactNotes', () => {
  it('keeps the stored key and the entry\'s other properties', () => {
    assert.deepEqual(mergeContactNotes({ abc: { '@type': 'Note', note: 'old', created: 'x' } }, 'new'), {
      abc: { '@type': 'Note', note: 'new', created: 'x' },
    });
  });

  it('writes a fresh note when the card had none', () => {
    assert.deepEqual(mergeContactNotes(undefined, 'hello'), { n0: { note: 'hello' } });
  });

  it('rejects a card storing several notes rather than collapsing them to one', () => {
    // `notes` is a single string, so writing it over a multi-note map would delete the others
    // — a field the caller was never shown, lost silently. Refuse and name the two-step route.
    assert.throws(
      () => mergeContactNotes({ a: { note: '1' }, b: { note: '2' } }, 'hello'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /2 separate notes/);
        assert.match(err.message, /clearFields/);
        return true;
      },
    );
  });
});

// ---------- mergeEntryMap ----------

describe('mergeEntryMap', () => {
  const stored = {
    sha1key: { address: 'a@example.com', contexts: { private: true }, pref: 1 },
    short1: { address: 'b@example.com', pref: 2 },
  };

  it('preserves the map key and every unsurfaced field of a matched entry', () => {
    const outcome = mergeEntryMap(stored, [{ address: 'a@example.com', label: 'home' }], 'address');
    assert.deepEqual(outcome.map, {
      sha1key: { address: 'a@example.com', contexts: { private: true }, pref: 1, label: 'home' },
    });
    assert.deepEqual(outcome.added, []);
    assert.deepEqual(outcome.dropped, [{ key: 'short1', entry: stored.short1 }]);
  });

  it('leaves a matched entry untouched when nothing but its value was supplied', () => {
    const outcome = mergeEntryMap(stored, [{ address: 'b@example.com' }], 'address');
    assert.deepEqual(outcome.map.short1, { address: 'b@example.com', pref: 2 });
  });

  it('builds a fresh entry for an unmatched value, avoiding an existing key', () => {
    const outcome = mergeEntryMap({ e0: { address: 'a@example.com' } }, [{ address: 'a@example.com' }, { address: 'z@example.com' }], 'address');
    assert.deepEqual(outcome.added, ['z@example.com']);
    assert.equal(Object.keys(outcome.map).length, 2);
    // e0 was taken by the matched entry, so the new one cannot reuse it.
    assert.deepEqual(outcome.map.e1, { address: 'z@example.com' });
  });

  it('reserves an existing key for its own entry even when the addition is handled first', () => {
    // The discriminating order. Above, the matched entry claims `e0` before the new one asks
    // for a key, so a generator that only avoided keys already written would still pass. Here
    // the NEW entry is minted first, while `e0` is still unwritten — and it must still skip
    // `e0`, because the matched entry is about to be stored under it. Getting this wrong
    // overwrites one of the two entries and the card silently loses an address.
    const outcome = mergeEntryMap(
      { e0: { address: 'a@example.com', pref: 1 } },
      [{ address: 'z@example.com' }, { address: 'a@example.com' }],
      'address',
    );
    assert.deepEqual(outcome.map, {
      e1: { address: 'z@example.com' },
      e0: { address: 'a@example.com', pref: 1 },
    });
    assert.deepEqual(outcome.added, ['z@example.com']);
    assert.deepEqual(outcome.dropped, []);
  });

  it('writes nothing when a labelled entry is read and sent straight back', () => {
    // The round-trip that must be a no-op: the read shape resolved "private" out of the
    // `contexts` SET, and the entry has no scalar `label` key. Echoing that label back must
    // not stamp one on, or every unchanged resend would mutate the card.
    const contextsCard = { k: { address: 'a@example.com', contexts: { private: true }, pref: 1 } };
    const outcome = mergeEntryMap(contextsCard, [{ address: 'a@example.com', label: 'private' }], 'address');
    assert.deepEqual(outcome.map, { k: { address: 'a@example.com', contexts: { private: true }, pref: 1 } });
    assert.equal('label' in outcome.map.k, false);
  });

  it('writes the scalar label when it genuinely differs, leaving contexts as it was', () => {
    // The documented residual: a relabel lands in `label` (which wins on the next read) while
    // the `contexts` set the Fastmail apps read is untouched, so the two can disagree.
    const contextsCard = { k: { address: 'a@example.com', contexts: { private: true } } };
    const outcome = mergeEntryMap(contextsCard, [{ address: 'a@example.com', label: 'work' }], 'address');
    assert.deepEqual(outcome.map.k, { address: 'a@example.com', contexts: { private: true }, label: 'work' });
  });

  it('does not mutate the stored map or its entries', () => {
    const snapshot = JSON.parse(JSON.stringify(stored));
    mergeEntryMap(stored, [{ address: 'a@example.com', label: 'home' }], 'address');
    assert.deepEqual(stored, snapshot);
  });

  it('matches exactly, so a case-differing value reads as a drop plus an add', () => {
    // Deliberate: resolving A@example.com onto a@example.com would be a guess about the
    // caller's intent. Surfacing it as an ambiguous edit gets it in front of the caller.
    const outcome = mergeEntryMap(stored, [{ address: 'A@example.com' }, { address: 'b@example.com' }], 'address');
    assert.deepEqual(outcome.added, ['A@example.com']);
    assert.deepEqual(outcome.dropped.map((d) => d.key), ['sha1key']);
  });

  it('treats a card with no such property as an empty map', () => {
    const outcome = mergeEntryMap(undefined, [{ address: 'a@example.com' }], 'address');
    assert.deepEqual(outcome.map, { e0: { address: 'a@example.com' } });
    assert.deepEqual(outcome.dropped, []);
  });
});

// ---------- isAmbiguousEntryEdit / assertUnambiguousEntryEdit ----------

describe('isAmbiguousEntryEdit', () => {
  it('answers the question without throwing, so the override can be scoped per field', () => {
    const dropped = [{ key: 'k', entry: { address: 'a@b.example' } }];
    assert.equal(isAmbiguousEntryEdit({ map: {}, dropped, added: ['n@b.example'] }), true);
    assert.equal(isAmbiguousEntryEdit({ map: {}, dropped, added: [] }), false);
    assert.equal(isAmbiguousEntryEdit({ map: {}, dropped: [], added: ['n@b.example'] }), false);
    assert.equal(isAmbiguousEntryEdit({ map: {}, dropped: [], added: [] }), false);
  });
});

describe('assertUnambiguousEntryEdit', () => {
  it('allows a pure addition and a pure removal', () => {
    assert.doesNotThrow(() => assertUnambiguousEntryEdit('emails', { map: {}, dropped: [], added: ['a@b.example'] }));
    assert.doesNotThrow(() =>
      assertUnambiguousEntryEdit('emails', { map: {}, dropped: [{ key: 'k', entry: { address: 'a@b.example' } }], added: [] }),
    );
  });

  it('rejects a simultaneous drop and add, echoing the dropped entry in full', () => {
    assert.throws(
      () =>
        assertUnambiguousEntryEdit('emails', {
          map: {},
          dropped: [{ key: 'k', entry: { address: 'old@b.example', contexts: { work: true }, pref: 1 } }],
          added: ['new@b.example'],
        }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /"contexts":\{"work":true\}/);
        assert.match(err.message, /"pref":1/);
        assert.match(err.message, /new@b\.example/);
        assert.match(err.message, /allowEntryReplace/);
        return true;
      },
    );
  });

  it('caps the echoed entries and says the list is partial', () => {
    const dropped = Array.from({ length: 8 }, (_, i) => ({ key: `k${i}`, entry: { address: `d${i}@b.example` } }));
    assert.throws(
      () => assertUnambiguousEntryEdit('phones', { map: {}, dropped, added: ['x'] }),
      /…and 3 more/,
    );
  });
});

// ---------- buildEntryMap ----------

describe('buildEntryMap', () => {
  it('numbers entries from the prefix', () => {
    assert.deepEqual(buildEntryMap([{ address: 'a' }, { address: 'b' }], 'e'), {
      e0: { address: 'a' },
      e1: { address: 'b' },
    });
  });
});

// ---------- the card kind ----------

// One reading of the property serves both the `kind` the read tools show and the refusal the
// write tools raise, so the two cannot end up disagreeing about what a group is.
describe('contactCardKind', () => {
  it('reads the declared kind', () => {
    assert.equal(contactCardKind({ kind: 'group' }), 'group');
    assert.equal(contactCardKind({ kind: 'individual' }), 'individual');
  });

  it('treats a missing, blank or non-string kind as undeclared', () => {
    assert.equal(contactCardKind({}), undefined);
    assert.equal(contactCardKind({ kind: '' }), undefined);
    assert.equal(contactCardKind({ kind: 42 }), undefined);
    assert.equal(contactCardKind(undefined), undefined);
  });
});

describe('nonDefaultContactKind', () => {
  it('drops the individual default, whether declared or merely implied', () => {
    assert.equal(nonDefaultContactKind({ kind: 'individual' }), undefined);
    assert.equal(nonDefaultContactKind({}), undefined);
  });

  it('keeps every other kind, so it is not a group-or-not flag', () => {
    for (const kind of ['group', 'org', 'location', 'device', 'application', 'x-robot']) {
      assert.equal(nonDefaultContactKind({ kind }), kind, kind);
    }
  });
});

describe('isContactGroupCard', () => {
  it('recognises a group', () => {
    assert.equal(isContactGroupCard({ kind: 'group' }), true);
  });

  it('does not treat other non-person kinds as groups', () => {
    // Only a group is refused by the write tools: it is the one kind whose whole content is a
    // members list this server has no surface for. An org card is an ordinary card that
    // update_contact and delete_contact can still handle.
    for (const kind of ['org', 'location', 'device', 'application', 'individual']) {
      assert.equal(isContactGroupCard({ kind }), false, kind);
    }
    assert.equal(isContactGroupCard({}), false);
    assert.equal(isContactGroupCard(undefined), false);
  });

  it('agrees with the kind the read surface surfaces', () => {
    // The read shows a value and the write refuses on one: the same card must not read as a
    // group in one place and not the other.
    const group = { id: 'G1', kind: 'group' };
    assert.equal(isContactGroupCard(group), true);
    assert.equal(nonDefaultContactKind(group), 'group');

    const person = { id: 'C1', kind: 'individual' };
    assert.equal(isContactGroupCard(person), false);
    assert.equal(nonDefaultContactKind(person), undefined);
  });
});
