import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { createContactTool, getContactTool, updateContactTool, deleteContactTool, listContactsTool, searchContactsTool, type ContactsReadClient, type ContactsWriteClient } from './contacts-handler.js';
import { InvalidInputError } from './coerce.js';
import { simplifyContact } from './response-formatters.js';
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

  it('reports the created id, not a failure, when only the read-back fails', async () => {
    const { client, calls } = makeClient();
    client.getContactById = async () => { throw new Error('read failed'); };
    const content = await createContactTool({ name: 'Ada Lovelace' }, client);
    assert.equal(calls.created.length, 1);
    assert.deepEqual(payload(content), { id: 'C1' });
    assert.equal(content.length, 2);
    assert.deepEqual(content.map((c) => c.type), ['text', 'text']);
    assert.match(content[1].text, /was created/);
    assert.match(content[1].text, /get_contact/);
    assert.match(content[1].text, /duplicate/);
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

  for (const clearFields of [{ notes: true }, 42, ['notes', 7]]) {
    it(`refuses an unparseable clearFields (${JSON.stringify(clearFields)}) rather than ignoring it`, async () => {
      const { client, calls } = makeClient();
      await assert.rejects(
        () => updateContactTool({ contactId: 'C1', notes: 'hi', clearFields }, client),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError);
          assert.match(err.message, /clearFields/);
          return true;
        },
      );
      assert.equal(calls.updated.length, 0);
    });
  }

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

describe('getContactTool', () => {
  it('trims contactId before the lookup, as update_contact and delete_contact do', async () => {
    const { client } = makeClient();
    const seen: string[] = [];
    client.getContactById = async (id: string) => { seen.push(id); return CARD; };
    await getContactTool({ contactId: '  C1 ' }, client);
    assert.deepEqual(seen, ['C1']);
  });

  it('returns one text item holding the simplified card by default', async () => {
    const { client } = makeClient();
    const content = await getContactTool({ contactId: 'C1' }, client);
    assert.equal(content.length, 1);
    assert.equal(content[0].type, 'text');
    assert.deepEqual(payload(content), simplifyContact(CARD, { verbose: false }));
  });

  it('returns the card untransformed under raw, and the verbose shape under verbose', async () => {
    const { client } = makeClient();
    assert.deepEqual(payload(await getContactTool({ contactId: 'C1', raw: true }, client)), CARD);
    const verbose = simplifyContact(CARD, { verbose: true });
    assert.notDeepEqual(verbose, simplifyContact(CARD, { verbose: false }));
    assert.deepEqual(payload(await getContactTool({ contactId: 'C1', verbose: 'true' }, client)), verbose);
  });

  for (const flag of ['raw', 'verbose']) {
    it(`names ${flag} when it cannot read it`, async () => {
      const { client } = makeClient();
      await assert.rejects(
        () => getContactTool({ contactId: 'C1', [flag]: 'garbage' }, client),
        (err: Error) => err instanceof InvalidInputError && err.message.startsWith(`${flag} must be true or false`),
      );
    });
  }

  it('refuses absent arguments as a missing contactId, not a TypeError', async () => {
    const { client } = makeClient();
    await assert.rejects(
      () => getContactTool(undefined, client),
      (err: Error) => err instanceof InvalidInputError && /contactId is required/.test(err.message),
    );
  });

  it('refuses a whitespace-only contactId without a lookup', async () => {
    const { client } = makeClient();
    let looked = false;
    client.getContactById = async () => { looked = true; return CARD; };
    await assert.rejects(
      () => getContactTool({ contactId: '   ' }, client),
      (err: Error) => err instanceof InvalidInputError && /contactId is required/.test(err.message),
    );
    assert.equal(looked, false);
  });
});

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

describe('contact tool boolean flags', () => {
  it('names the parameter when a boolean flag cannot be read', async () => {
    await assert.rejects(() => createContactTool({ name: 'Ada', raw: 'yes' }, {} as any), /raw must be true or false/);
    await assert.rejects(() => createContactTool({ name: 'Ada', verbose: 'yes' }, {} as any), /verbose must be true or false/);
    await assert.rejects(() => updateContactTool({ contactId: 'C1', notes: 'x', raw: 'yes' }, {} as any), /raw must be true or false/);
    await assert.rejects(() => updateContactTool({ contactId: 'C1', notes: 'x', verbose: 'yes' }, {} as any), /verbose must be true or false/);
    await assert.rejects(() => updateContactTool({ contactId: 'C1', notes: 'x', allowEntryReplace: 'yes' }, {} as any), /allowEntryReplace must be true or false/);
  });
});

// ---------- list_contacts / search_contacts ----------

// A read client that records every query, so a test can prove a refused call sent none.
function makeReadClient(page: { ids: string[]; total: number; position: number }): {
  client: ContactsReadClient;
  queries: Array<{ query?: string; limit: number; position?: number }>;
} {
  const queries: Array<{ query?: string; limit: number; position?: number }> = [];
  const result = () => ({ items: page.ids.map((id) => ({ ...CARD, id })), total: page.total, position: page.position });
  const client: ContactsReadClient = {
    async getContacts(limit, position) {
      queries.push({ limit, position });
      return result();
    },
    async searchContacts(query, limit, position) {
      queries.push({ query, limit, position });
      return result();
    },
  };
  return { client, queries };
}

const READ_TOOLS = [
  { tool: 'listContactsTool', run: (args: any, limit: number, client: ContactsReadClient) => listContactsTool(args, limit, client) },
  { tool: 'searchContactsTool', run: (args: any, limit: number, client: ContactsReadClient) => searchContactsTool({ query: 'ada', ...args }, limit, client) },
];

for (const { tool, run } of READ_TOOLS) {
  describe(`${tool} paging`, () => {
    for (const bad of [-1, 1.5, 'abc']) {
      it(`refuses position ${JSON.stringify(bad)} before any query`, async () => {
        const { client, queries } = makeReadClient({ ids: ['C1'], total: 1, position: 0 });
        await assert.rejects(
          () => run({ position: bad }, 20, client),
          (err: Error) => err instanceof InvalidInputError && /^position /.test(err.message),
        );
        assert.equal(queries.length, 0, 'a refused position must not reach the server');
      });
    }

    it('refuses a negative position without the email-only ascending hint', async () => {
      const { client } = makeReadClient({ ids: ['C1'], total: 1, position: 0 });
      await assert.rejects(
        () => run({ position: -1 }, 20, client),
        (err: Error) => {
          assert.ok(err instanceof InvalidInputError);
          assert.doesNotMatch(err.message, /ascending/);
          return true;
        },
      );
    });

    it('hands the coerced position and the given limit to the client', async () => {
      const { client, queries } = makeReadClient({ ids: ['C1'], total: 1, position: 40 });
      await run({ position: '40' }, 25, client);
      assert.equal(queries.length, 1);
      assert.equal(queries[0].position, 40);
      assert.equal(queries[0].limit, 25);
    });

    it('omits position when the caller gave none', async () => {
      const { client, queries } = makeReadClient({ ids: ['C1'], total: 1, position: 0 });
      await run({}, 20, client);
      assert.equal(queries[0].position, undefined);
    });

    for (const raw of [false, true]) {
      it(`offers nextPosition while more remain${raw ? ' under raw' : ''}`, async () => {
        const { client } = makeReadClient({ ids: ['C1', 'C2'], total: 10, position: 4 });
        const content = await run({ position: 4, raw }, 2, client);
        assert.ok(
          content[0].text.startsWith('Showing 2 of 10 results from position 4. nextPosition: 6'),
          content[0].text.split('\n')[0],
        );
      });
    }
  });
}

describe('searchContactsTool', () => {
  it('refuses a missing query before any query is sent', async () => {
    const { client, queries } = makeReadClient({ ids: [], total: 0, position: 0 });
    await assert.rejects(
      () => searchContactsTool({}, 20, client),
      (err: Error) => err instanceof InvalidInputError && /query is required/.test(err.message),
    );
    assert.equal(queries.length, 0);
  });
});
