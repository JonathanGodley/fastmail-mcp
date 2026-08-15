import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { createContactTool, updateContactTool, deleteContactTool, type ContactsWriteClient } from './contacts-handler.js';
import { InvalidInputError } from './coerce.js';
import type { UpdateContactPatch } from './contacts-calendar.js';

// A card in the shape a live address book actually returns: opaque entry-map keys, a
// contexts set instead of a label, and pref on every entry.
const CARD = {
  id: 'C1',
  '@type': 'Card',
  uid: 'uid-c1',
  name: { full: 'Ada Lovelace', components: [{ kind: 'given', value: 'Ada' }] },
  emails: {
    '0dd713ddb7cdcb0fbc17f59321f227fabcddc7da': { address: 'ada@example.com', contexts: { private: true }, pref: 1 },
  },
};

interface StubCalls {
  created: any[];
  updated: Array<{ id: string; patch: UpdateContactPatch }>;
  deleted: string[];
}

// What the card looks like AFTER the update. Deliberately a different object from CARD: a
// stub that returned the same one for `contact` and `previousCard` would pass even if the
// tool wired the echo to the post-edit card, which is the one thing the envelope must not do.
const UPDATED_CARD = { ...CARD, notes: { n0: { note: 'hi' } } };

function makeClient(opts: { card?: any; updateResult?: any; deleteResult?: any } = {}): {
  client: ContactsWriteClient;
  calls: StubCalls;
} {
  const card = opts.card ?? CARD;
  const calls: StubCalls = { created: [], updated: [], deleted: [] };
  const client: ContactsWriteClient = {
    async createContact(input) {
      calls.created.push(input);
      return 'C1';
    },
    async getContactById() {
      return card;
    },
    async updateContact(id, patch) {
      calls.updated.push({ id, patch });
      return opts.updateResult ?? { previousCard: card, contact: UPDATED_CARD };
    },
    async deleteContact(id) {
      calls.deleted.push(id);
      return opts.deleteResult ?? { deletedCard: card };
    },
  };
  return { client, calls };
}

/** The JSON payload of a tool result's first content item. */
function payload(content: Array<{ type: 'text'; text: string }>): any {
  return JSON.parse(content[0].text);
}

// ---------- create_contact ----------

describe('createContactTool', () => {
  it('reads the created card back and returns the simplified shape', async () => {
    const { client } = makeClient();
    const content = await createContactTool({ name: 'Ada Lovelace', emails: ['ada@example.com'] }, client);
    assert.equal(content.length, 1);
    assert.deepEqual(payload(content), {
      id: 'C1',
      name: 'Ada Lovelace',
      emails: [{ address: 'ada@example.com', label: 'private' }],
    });
  });

  it('coerces every input array before handing it to the client', async () => {
    const { client, calls } = makeClient();
    await createContactTool(
      { name: '{"given":"Ada","surname":"Lovelace"}', emails: '["ada@example.com"]', phones: [{ number: '+1 555 0100' }] },
      client,
    );
    assert.deepEqual(calls.created[0].name, { given: 'Ada', surname: 'Lovelace' });
    assert.deepEqual(calls.created[0].emails, [{ address: 'ada@example.com' }]);
    assert.deepEqual(calls.created[0].phones, [{ number: '+1 555 0100' }]);
  });

  it('returns the raw card under raw, and the whole entries under verbose', async () => {
    const { client } = makeClient();
    assert.deepEqual(payload(await createContactTool({ name: 'Ada', raw: true }, client)), CARD);
    assert.deepEqual(
      payload(await createContactTool({ name: 'Ada', verbose: 'true' }, client)).emails,
      Object.values(CARD.emails),
    );
  });

  it('rejects a non-string addressBookId, which would become a nonsense card key', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => createContactTool({ name: 'Ada', addressBookId: 42 }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /addressBookId must be a non-empty string/);
        return true;
      },
    );
  });

  it('rejects an empty notes string rather than dropping it', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => createContactTool({ name: 'Ada', notes: '  ' }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /notes cannot be empty/);
        return true;
      },
    );
  });
});

// ---------- update_contact ----------

describe('updateContactTool', () => {
  it('returns the {contact, previousCard} envelope', async () => {
    const { client } = makeClient();
    const body = payload(await updateContactTool({ contactId: 'C1', notes: 'hi' }, client));
    assert.deepEqual(Object.keys(body), ['contact', 'previousCard']);
    // `contact` is simplified by default, and is the card AFTER the write…
    assert.deepEqual(body.contact.emails, [{ address: 'ada@example.com', label: 'private' }]);
    assert.equal(body.contact.notes, 'hi');
    // …while the echo is the raw card as it stood BEFORE it, which had no note.
    assert.deepEqual(body.previousCard, CARD);
    assert.equal('notes' in body.previousCard, false);
  });

  it('keeps the envelope under raw, with previousCard still the pre-edit raw card', async () => {
    // raw governs the CARD the tool returns, never the echo: keeping what the write took away
    // visible is the echo's whole purpose, and a caller asking for exact JMAP is the last one
    // that should lose it.
    const { client } = makeClient();
    const body = payload(await updateContactTool({ contactId: 'C1', notes: 'hi', raw: true }, client));
    assert.deepEqual(Object.keys(body), ['contact', 'previousCard']);
    assert.deepEqual(body.contact, UPDATED_CARD);
    assert.deepEqual(body.previousCard, CARD);
  });

  it('keeps the envelope under verbose, with previousCard unaffected by it', async () => {
    const { client } = makeClient();
    const body = payload(await updateContactTool({ contactId: 'C1', notes: 'hi', verbose: 'true' }, client));
    assert.deepEqual(body.contact.emails, Object.values(UPDATED_CARD.emails));
    assert.deepEqual(body.previousCard, CARD);
  });

  it('says so, rather than dropping the field, when the read-back did not come back', async () => {
    const { client } = makeClient({ updateResult: { previousCard: CARD } });
    const content = await updateContactTool({ contactId: 'C1', notes: 'hi' }, client);
    const body = payload(content);
    assert.equal('contact' in body, false);
    assert.deepEqual(body.previousCard, CARD);
    assert.equal(content.length, 2);
    assert.match(content[1].text, /did not return the updated card/);
    assert.match(content[1].text, /get_contact/);
  });

  it('coerces clearFields and allowEntryReplace from a lenient client', async () => {
    const { client, calls } = makeClient();
    await updateContactTool({ contactId: 'C1', clearFields: 'notes', allowEntryReplace: 'true' }, client);
    assert.deepEqual(calls.updated[0].patch.clearFields, ['notes']);
    assert.equal(calls.updated[0].patch.allowEntryReplace, true);
  });

  it('defaults allowEntryReplace to false, including for a stringified "false"', async () => {
    const { client, calls } = makeClient();
    await updateContactTool({ contactId: 'C1', notes: 'x', allowEntryReplace: 'false' }, client);
    assert.equal(calls.updated[0].patch.allowEntryReplace, false);
  });

  it('requires a contactId', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => updateContactTool({ notes: 'x' }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /contactId is required/);
        return true;
      },
    );
  });

  it('rejects an empty notes string, naming clearFields', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => updateContactTool({ contactId: 'C1', notes: '' }, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /clearFields:\['notes'\]/);
        return true;
      },
    );
  });
});

// ---------- delete_contact ----------

describe('deleteContactTool', () => {
  it('returns the id and the full pre-destroy card', async () => {
    const { client, calls } = makeClient();
    const body = payload(await deleteContactTool({ contactId: 'C1' }, client));
    assert.deepEqual(calls.deleted, ['C1']);
    assert.deepEqual(body, { deleted: 'C1', deletedCard: CARD });
  });

  it('requires a contactId', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => deleteContactTool({}, client),
      (err: Error) => {
        assert.ok(err instanceof InvalidInputError);
        assert.match(err.message, /contactId is required/);
        return true;
      },
    );
  });

  it('states the degrade loudly when the destroy landed but no card came back', async () => {
    // The one degrade with no retry: the card is gone and nothing described it. Reporting the
    // delete with an absent `deletedCard` and no explanation would read as "the card was
    // empty", so the tool says what actually happened instead.
    const { client } = makeClient({ deleteResult: {} });
    const content = await deleteContactTool({ contactId: 'C1' }, client);
    assert.deepEqual(payload(content), { deleted: 'C1' });
    assert.equal(content.length, 2);
    assert.match(content[1].text, /WARNING/);
    assert.match(content[1].text, /irreversible/);
  });

  it('lets a not-found from the client through unchanged', async () => {
    const { client } = makeClient();
    client.deleteContact = async () => {
      throw new InvalidInputError('Contact not found: ghost');
    };
    await assert.rejects(() => deleteContactTool({ contactId: 'ghost' }, client), /Contact not found: ghost/);
  });
});
