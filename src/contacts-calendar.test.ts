import { describe, it, beforeEach, mock } from 'node:test';
import assert from 'node:assert/strict';
import { ContactsCalendarClient } from './contacts-calendar.js';
import type { JmapRequest } from './jmap-client.js';
import { FastmailAuth } from './auth.js';
import { callArguments, findCallArguments, type RecordedCalls } from './testing/mock-calls.js';

// ---------- helpers ----------

const MAIL_ACCOUNT = 'acct-mail';
const CONTACTS_ACCOUNT = 'acct-contacts';

function makeClient(opts: { contactsPrimary?: boolean } = { contactsPrimary: true }): ContactsCalendarClient {
  const auth = new FastmailAuth({ apiToken: 'fake-token' });
  const client = new ContactsCalendarClient(auth);

  mock.method(client, 'getSession', async () => ({
    apiUrl: 'https://api.example.com/jmap/api/',
    accountId: MAIL_ACCOUNT,
    capabilities: { 'urn:ietf:params:jmap:contacts': {} },
    primaryAccounts: opts.contactsPrimary
      ? { 'urn:ietf:params:jmap:contacts': CONTACTS_ACCOUNT, 'urn:ietf:params:jmap:mail': MAIL_ACCOUNT }
      : { 'urn:ietf:params:jmap:mail': MAIL_ACCOUNT },
  }));

  return client;
}

// The parameter is declared even though the stub ignores it: node:test types the
// returned mock from BOTH the real method and the implementation, so an
// implementation taking no arguments makes the recorded arguments a union with the
// empty tuple, and every read of the request goes back to being possibly undefined.
function stubMakeRequest(client: ContactsCalendarClient, response: any) {
  return mock.method(client, 'makeRequest', async (_request: JmapRequest) => response);
}

function queryAndGetResponse(list: any[]) {
  return {
    methodResponses: [
      ['ContactCard/query', { ids: list.map((c) => c.id), total: list.length, position: 0 }, 'query'],
      ['ContactCard/get', { list }, 'contacts'],
    ],
  };
}

/** Every ContactCard/get in a batch, as [method, params] pairs. */
function contactCardGets(makeReq: any): any[] {
  return makeReq.mock.calls
    .flatMap((call: any) => call.arguments[0].methodCalls)
    .filter((mc: any) => mc[0] === 'ContactCard/get')
    .map((mc: any) => mc[1]);
}

// ---------- contacts reads ----------

describe('contacts reads', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('addresses the contacts account from getContacts', async () => {
    const makeReq = stubMakeRequest(client, queryAndGetResponse([{ id: 'C1' }]));
    await client.getContacts(10);
    const [query, get] = callArguments(makeReq)[0].methodCalls;
    assert.equal(query[1].accountId, CONTACTS_ACCOUNT);
    assert.equal(get[1].accountId, CONTACTS_ACCOUNT);
  });

  it('addresses the contacts account from searchContacts', async () => {
    const makeReq = stubMakeRequest(client, queryAndGetResponse([{ id: 'C1' }]));
    await client.searchContacts('ada', 10);
    const [query, get] = callArguments(makeReq)[0].methodCalls;
    assert.equal(query[1].accountId, CONTACTS_ACCOUNT);
    assert.equal(get[1].accountId, CONTACTS_ACCOUNT);
  });

  it('addresses the contacts account from getContactById', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'contact']],
    });
    await client.getContactById('C1');
    assert.equal(callArguments(makeReq)[0].methodCalls[0][1].accountId, CONTACTS_ACCOUNT);
  });

  it('throws from every reader when no contacts primary exists', async () => {
    client = makeClient({ contactsPrimary: false });
    const makeReq = stubMakeRequest(client, queryAndGetResponse([{ id: 'C1' }]));

    await assert.rejects(() => client.getContacts(10), /primary account/i);
    await assert.rejects(() => client.searchContacts('ada', 10), /primary account/i);
    await assert.rejects(() => client.getContactById('C1'), /primary account/i);
    // Reading from the mail account would silently return the wrong account's
    // contacts, so no reader may issue a request at all.
    assert.equal(makeReq.mock.calls.length, 0);
  });

  it('requests no properties filter, so verbose reads keep every contact field', async () => {
    const makeReq = stubMakeRequest(client, queryAndGetResponse([{ id: 'C1' }]));
    await client.getContacts(10);
    await client.searchContacts('ada', 10);

    const gets = contactCardGets(makeReq);
    assert.equal(gets.length, 2);
    for (const params of gets) {
      // A properties filter here drops fields the verbose/raw contact output
      // promises. Re-adding one is a silent output regression, so it is pinned.
      assert.ok(!('properties' in params), 'ContactCard/get must not filter properties');
    }
  });

  it('rejects an unknown contact id as caller-fixable input', async () => {
    stubMakeRequest(client, {
      methodResponses: [['ContactCard/get', { list: [], notFound: ['ghost'] }, 'contact']],
    });
    await assert.rejects(
      () => client.getContactById('ghost'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Contact not found: ghost/);
        return true;
      },
    );
  });

  it('surfaces the server-reported total alongside the returned page', async () => {
    // The page is capped at the caller's limit, so `total` is the only signal
    // that the address book holds more than the caller just received.
    stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/query', { ids: ['C1'], total: 250, position: 0 }, 'query'],
        ['ContactCard/get', { list: [{ id: 'C1' }] }, 'contacts'],
      ],
    });

    const result = await client.getContacts(1);
    assert.equal(result.total, 250);
    assert.equal(result.items.length, 1);
  });

  it('omits total on the AddressBook fallback rather than inventing one', async () => {
    // AddressBook/get returns no count, so a fabricated total (say items.length)
    // would read as "this is the whole address book" on a truncated page.
    let call = 0;
    mock.method(client, 'makeRequest', async () => {
      call += 1;
      if (call === 1) throw new Error('ContactCard/query not supported');
      return {
        methodResponses: [['AddressBook/get', { list: [{ id: 'AB1' }] }, 'addressbooks']],
      };
    });

    const result = await client.getContacts(10);
    assert.equal('total' in result, false);
    assert.equal(result.items.length, 1);
  });
});

// ---------- createContact ----------

describe('createContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('builds an RFC 9610 Card and returns the created id', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C1', uid: 'u-1' } } }, 'createContact'],
      ],
    });

    const id = await client.createContact({
      name: { given: 'Ada', surname: 'Lovelace', full: 'Ada Lovelace' },
      emails: [{ address: 'ada@example.com', label: 'work' }],
      phones: [{ number: '+1 555 0100' }],
      notes: 'test note',
    });

    assert.equal(id, 'C1');
    const [method, params] = callArguments(makeReq)[0].methodCalls[0];
    assert.equal(method, 'ContactCard/set');
    assert.equal(params.accountId, CONTACTS_ACCOUNT);
    const card = params.create.newContact;
    assert.equal(card['@type'], 'Card');
    assert.equal(card.name.full, 'Ada Lovelace');
    assert.deepEqual(card.name.components, [
      { kind: 'given', value: 'Ada' },
      { kind: 'surname', value: 'Lovelace' },
    ]);
    assert.deepEqual(card.emails, { e0: { address: 'ada@example.com', label: 'work' } });
    assert.deepEqual(card.phones, { p0: { number: '+1 555 0100' } });
    assert.deepEqual(card.notes, { n0: { note: 'test note' } });
  });

  it('throws rather than writing to the mail account when no contacts primary exists', async () => {
    client = makeClient({ contactsPrimary: false });
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C2' } } }, 'createContact'],
      ],
    });

    await assert.rejects(
      () => client.createContact({ name: { full: 'Solo Mail' } }),
      (err: Error) => {
        // The message has to teach the fix, not read as a bare account error.
        assert.match(err.message, /primary account/i);
        assert.match(err.message, /check_function_availability/);
        return true;
      },
    );
    // The load-bearing assertion: no request at all. A fall back to the mail
    // account would have created the contact in the wrong account instead.
    assert.equal(makeReq.mock.calls.length, 0);
  });

  it('passes addressBookIds when addressBookId is supplied', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { created: { newContact: { id: 'C3' } } }, 'createContact'],
      ],
    });
    await client.createContact({ name: { full: 'Booked' }, addressBookId: 'ab-1' });
    const card = callArguments(makeReq)[0].methodCalls[0][1].create.newContact;
    assert.deepEqual(card.addressBookIds, { 'ab-1': true });
  });

  it('rejects empty input client-side', async () => {
    await assert.rejects(
      () => client.createContact({}),
      /name or at least one email/,
    );
  });

  it('classes empty input as caller-fixable', async () => {
    // A missing required field is corrected in one retry, so it maps to InvalidParams
    // rather than the InternalError that reads as "nothing you can do".
    await assert.rejects(
      () => client.createContact({}),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /name or at least one email/);
        return true;
      },
    );
  });

  it('rejects an empty entry array, pointing at omission rather than clearFields', async () => {
    // The same rule update_contact applies, with the only difference the two tools can have:
    // a card being created has nothing to clear, so the route out is to leave the field off.
    // Accepting [] here while rejecting it there would make the pair inconsistent for nothing.
    for (const field of ['emails', 'phones', 'addresses'] as const) {
      await assert.rejects(
        () => client.createContact({ name: { full: 'Ada' }, [field]: [] }),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, new RegExp(`${field}: \\[\\] is not accepted`));
          assert.match(err.message, new RegExp(`Omit ${field}`));
          return true;
        },
      );
    }
  });

  it('surfaces notCreated errors with type and description', async () => {
    stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/set', { notCreated: { newContact: { type: 'invalidProperties', description: 'bad email' } } }, 'createContact'],
      ],
    });
    await assert.rejects(
      () => client.createContact({ name: { full: 'X' } }),
      /invalidProperties.*bad email/s,
    );
  });
});

// ---------- updateContact ----------

describe('updateContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  it('checks existence then sends a top-level patch', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { updated: { C1: null } }, 'u']] };
    });

    await client.updateContact('C1', { emails: [{ address: 'new@example.com' }] });

    const [setRequest] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][0] === 'ContactCard/set',
      'issuing ContactCard/set',
    );
    const update = setRequest.methodCalls[0][1].update.C1;
    assert.deepEqual(update, { emails: { e0: { address: 'new@example.com' } } });
  });

  it('passes expectState through as ifInState', async () => {
    const makeReq = mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { updated: { C1: null } }, 'u']] };
    });
    await client.updateContact('C1', { notes: 'x', expectState: 'state-42' });
    const [setRequest] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][0] === 'ContactCard/set',
      'issuing ContactCard/set',
    );
    assert.equal(setRequest.methodCalls[0][1].ifInState, 'state-42');
  });

  it('throws not-found before attempting the update', async () => {
    const makeReq = stubMakeRequest(client, {
      methodResponses: [
        ['ContactCard/get', { list: [], notFound: ['ghost'] }, 'g'],
      ],
    });
    await assert.rejects(() => client.updateContact('ghost', { notes: 'x' }), /Contact not found: ghost/);
    assert.equal(makeReq.mock.calls.length, 1);
  });

  it('classes the not-found rejection as caller-fixable input', async () => {
    // A message regex passes under either class. A wrong id is corrected and re-sent by
    // the caller, so it must arrive as InvalidParams, not the InternalError a plain
    // Error maps to.
    stubMakeRequest(client, {
      methodResponses: [['ContactCard/get', { list: [], notFound: ['ghost'] }, 'g']],
    });
    await assert.rejects(
      () => client.updateContact('ghost', { notes: 'x' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Contact not found: ghost/);
        return true;
      },
    );
  });

  it('surfaces notUpdated errors', async () => {
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { notUpdated: { C1: { type: 'stateMismatch' } } }, 'u']] };
    });
    await assert.rejects(() => client.updateContact('C1', { notes: 'x' }), /stateMismatch/);
  });

  it('keeps an operational set failure a server-side error', async () => {
    // The other half of the split: a stateMismatch is not fixed by re-forming the
    // arguments, so it must NOT be tagged as caller-fixable input.
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [{ id: 'C1' }] }, 'g']] };
      }
      return { methodResponses: [['ContactCard/set', { notUpdated: { C1: { type: 'stateMismatch' } } }, 'u']] };
    });
    await assert.rejects(
      () => client.updateContact('C1', { notes: 'x' }),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /Failed to update contact: stateMismatch/);
        return true;
      },
    );
  });

  it('rejects an empty patch client-side', async () => {
    await assert.rejects(() => client.updateContact('C1', {}), /at least one field/i);
  });

  it('classes an empty patch as caller-fixable input', async () => {
    await assert.rejects(
      () => client.updateContact('C1', {}),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /at least one field/i);
        return true;
      },
    );
  });
});

// ---------- updateContact: the per-entry merge ----------

// A card shaped like the ones a live address book actually holds: opaque entry-map keys, a
// `contexts` set instead of a `label`, `pref` on every entry, and a structured name.
function storedCard(overrides: Record<string, any> = {}): any {
  return {
    id: 'C1',
    '@type': 'Card',
    version: '1.0',
    uid: 'uid-c1',
    name: {
      full: 'Ada Lovelace',
      components: [
        { kind: 'given', value: 'Ada' },
        { kind: 'surname', value: 'Lovelace' },
      ],
    },
    emails: {
      '0dd713ddb7cdcb0fbc17f59321f227fabcddc7da': { '@type': 'EmailAddress', address: 'ada@example.com', contexts: { private: true }, pref: 1 },
      '506539': { '@type': 'EmailAddress', address: 'ada@work.example', contexts: { work: true }, pref: 2 },
    },
    phones: {
      ba5ddd: { '@type': 'Phone', number: '+1 555 0100', contexts: { mobile: true }, pref: 1 },
    },
    ...overrides,
  };
}

describe('updateContact merge', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  /**
   * Stub the two round trips an update makes: the whole-card read, then the
   * [set, read-back] batch.
   */
  function stubUpdate(
    target: ContactsCalendarClient,
    card: any,
    opts: { setResult?: any; updatedCard?: any; omitReadBack?: boolean } = {},
  ) {
    return mock.method(target, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: card ? [card] : [] }, 'card']] };
      }
      const responses: any[] = [
        ['ContactCard/set', opts.setResult ?? { updated: { C1: null } }, 'updateContact'],
      ];
      if (!opts.omitReadBack) {
        responses.push(['ContactCard/get', { list: [opts.updatedCard ?? card] }, 'updatedCard']);
      }
      return { methodResponses: responses };
    });
  }

  /** The PatchObject the update sent for `id`. */
  function patchFrom(makeReq: RecordedCalls<[any]>, id = 'C1'): any {
    const [setRequest] = findCallArguments(
      makeReq,
      ([req]) => req.methodCalls[0][0] === 'ContactCard/set',
      'issuing ContactCard/set',
    );
    return setRequest.methodCalls[0][1].update[id];
  }

  it('fetches the whole card, not an id-only existence probe', async () => {
    // An existence probe answers "does it exist" and nothing else. The merge needs every
    // field the entries carry, so a `properties` filter here would silently defeat it.
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', { notes: 'hello' });
    const readParams = callArguments(makeReq)[0].methodCalls[0][1];
    assert.deepEqual(readParams.ids, ['C1']);
    assert.ok(!('properties' in readParams), 'the pre-merge read must not filter properties');
  });

  it('keeps contexts and pref on an entry whose address the edit retains', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', {
      emails: [{ address: 'ada@example.com', label: 'personal' }, { address: 'ada@work.example' }],
    });

    const emails = patchFrom(makeReq).emails;
    // Existing map keys preserved, hidden fields intact, only `label` written over.
    assert.deepEqual(emails['0dd713ddb7cdcb0fbc17f59321f227fabcddc7da'], {
      '@type': 'EmailAddress',
      address: 'ada@example.com',
      contexts: { private: true },
      pref: 1,
      label: 'personal',
    });
    assert.deepEqual(emails['506539'], {
      '@type': 'EmailAddress',
      address: 'ada@work.example',
      contexts: { work: true },
      pref: 2,
    });
  });

  it('adds an entry the card does not have without disturbing the ones it does', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', {
      emails: [{ address: 'ada@example.com' }, { address: 'ada@work.example' }, { address: 'ada@new.example' }],
    });

    const emails = patchFrom(makeReq).emails;
    assert.equal(Object.keys(emails).length, 3);
    assert.equal(emails['0dd713ddb7cdcb0fbc17f59321f227fabcddc7da'].pref, 1);
    const addedKey = Object.keys(emails).find((k) => emails[k].address === 'ada@new.example')!;
    // A fresh entry carries only what the caller supplied — nothing is invented for it.
    assert.deepEqual(emails[addedKey], { address: 'ada@new.example' });
  });

  it('drops an entry the edit leaves out, with no ambiguity to resolve', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', { emails: [{ address: 'ada@example.com' }] });
    const emails = patchFrom(makeReq).emails;
    assert.deepEqual(Object.keys(emails), ['0dd713ddb7cdcb0fbc17f59321f227fabcddc7da']);
  });

  it('rejects an edit that both drops a stored entry and adds an unknown one', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await assert.rejects(
      () => client.updateContact('C1', { emails: [{ address: 'ada@example.com' }, { address: 'typo@example.com' }] }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        // The dropped entry is echoed WHOLE, hidden fields included, so the lossless retry
        // is cheaper than reaching for the override.
        assert.match(err.message, /ada@work\.example/);
        assert.match(err.message, /"pref":2/);
        assert.match(err.message, /"work":true/);
        assert.match(err.message, /typo@example\.com/);
        assert.match(err.message, /allowEntryReplace/);
        return true;
      },
    );
    // Rejected before the write: only the read happened.
    assert.equal(makeReq.mock.calls.length, 1);
  });

  it('proceeds under allowEntryReplace, writing every entry fresh', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', {
      emails: [{ address: 'ada@example.com' }, { address: 'typo@example.com' }],
      allowEntryReplace: true,
    });

    const emails = patchFrom(makeReq).emails;
    // Whole-entry-replace semantics: new keys, and the hidden fields deliberately gone —
    // which is exactly what the rejection warned would happen.
    assert.deepEqual(emails, {
      e0: { address: 'ada@example.com' },
      e1: { address: 'typo@example.com' },
    });
  });

  it('scopes allowEntryReplace to the array that was actually ambiguous', async () => {
    // The flag says "I accept the loss on the edit you refused", not "rewrite everything".
    // emails here is the ambiguous one; phones merges cleanly in the same call and must keep
    // its map key and hidden fields, because nothing about it was ever in doubt.
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', {
      emails: [{ address: 'ada@example.com' }, { address: 'typo@example.com' }],
      phones: [{ number: '+1 555 0100' }],
      allowEntryReplace: true,
    });

    const patch = patchFrom(makeReq);
    assert.deepEqual(patch.emails, {
      e0: { address: 'ada@example.com' },
      e1: { address: 'typo@example.com' },
    });
    assert.deepEqual(patch.phones, {
      ba5ddd: { '@type': 'Phone', number: '+1 555 0100', contexts: { mobile: true }, pref: 1 },
    });
  });

  it('still merges when allowEntryReplace is set but nothing was ambiguous', async () => {
    // The flag is an intent marker for a refusal that happened, so it must not turn an
    // ordinary edit into a lossy one just by being present.
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', {
      emails: [{ address: 'ada@example.com' }, { address: 'ada@work.example' }],
      allowEntryReplace: true,
    });
    const emails = patchFrom(makeReq).emails;
    assert.equal(emails['0dd713ddb7cdcb0fbc17f59321f227fabcddc7da'].pref, 1);
    assert.equal(emails['506539'].contexts.work, true);
  });

  it('patches only the fields the caller supplied, leaving the rest of the card unnamed', async () => {
    // A JMAP PatchObject replaces every property it names, so a property that has no business
    // in this edit must not appear at all — naming `titles` with the value just read would
    // turn a note edit into a rewrite of fields the caller never touched.
    const card = storedCard({
      titles: { t0: { name: 'Countess' } },
      organizations: { o0: { name: 'Analytical Engine Ltd' } },
      nicknames: { k0: { name: 'Ada' } },
    });
    const makeReq = stubUpdate(client, card);
    await client.updateContact('C1', { notes: 'hello' });

    const patch = patchFrom(makeReq);
    assert.deepEqual(Object.keys(patch), ['notes']);
    for (const untouched of ['titles', 'organizations', 'nicknames', 'name', 'emails', 'phones', 'uid']) {
      assert.equal(untouched in patch, false, `${untouched} must not appear in the patch`);
    }
  });

  it('rejects a card that stores several notes rather than collapsing them', async () => {
    const card = storedCard({ notes: { a: { note: 'first' }, b: { note: 'second' } } });
    const makeReq = stubUpdate(client, card);
    await assert.rejects(
      () => client.updateContact('C1', { notes: 'replacement' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /2 separate notes/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 1, 'rejected before the write');
  });

  it('merges phones by number, the same way emails merge by address', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', { phones: [{ number: '+1 555 0100', label: 'cell' }] });
    assert.deepEqual(patchFrom(makeReq).phones, {
      ba5ddd: { '@type': 'Phone', number: '+1 555 0100', contexts: { mobile: true }, pref: 1, label: 'cell' },
    });
  });

  it('merges a bare-string name into the stored components rather than replacing them', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', { name: { full: 'Augusta Ada King' } });
    assert.deepEqual(patchFrom(makeReq).name, {
      full: 'Augusta Ada King',
      components: [
        { kind: 'given', value: 'Ada' },
        { kind: 'surname', value: 'Lovelace' },
      ],
    });
  });

  it('merges per component on a card that has components and no full name', async () => {
    // A real shape: several stored cards carry `components` with no `full` at all.
    const card = storedCard({
      name: { '@type': 'Name', components: [{ kind: 'given', value: 'Ada' }, { kind: 'surname', value: 'Byron' }] },
    });
    const makeReq = stubUpdate(client, card);
    await client.updateContact('C1', { name: { surname: 'Lovelace' } });
    assert.deepEqual(patchFrom(makeReq).name, {
      '@type': 'Name',
      components: [
        { kind: 'given', value: 'Ada' },
        { kind: 'surname', value: 'Lovelace' },
      ],
    });
  });

  it('replaces addresses whole, since an address entry has no matchable key', async () => {
    const card = storedCard({ addresses: { old: { full: '1 Old Road', pref: 1 } } });
    const makeReq = stubUpdate(client, card);
    await client.updateContact('C1', { addresses: [{ full: '2 New Street', label: 'home' }] });
    assert.deepEqual(patchFrom(makeReq).addresses, {
      a0: { full: '2 New Street', label: 'home' },
    });
  });

  it('keeps the stored note key when replacing the note text', async () => {
    const card = storedCard({ notes: { 'note-abc': { '@type': 'Note', note: 'old', created: '2020-01-01T00:00:00Z' } } });
    const makeReq = stubUpdate(client, card);
    await client.updateContact('C1', { notes: 'new' });
    assert.deepEqual(patchFrom(makeReq).notes, {
      'note-abc': { '@type': 'Note', note: 'new', created: '2020-01-01T00:00:00Z' },
    });
  });

  it('rejects an empty array and names clearFields', async () => {
    for (const field of ['emails', 'phones', 'addresses'] as const) {
      await assert.rejects(
        () => client.updateContact('C1', { [field]: [] }),
        (err: Error) => {
          assert.equal(err.name, 'InvalidInputError');
          assert.match(err.message, new RegExp(`clearFields:\\['${field}'\\]`));
          return true;
        },
      );
    }
  });

  it('rejects an empty notes string and names clearFields', async () => {
    await assert.rejects(
      () => client.updateContact('C1', { notes: '  ' }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /clearFields:\['notes'\]/);
        return true;
      },
    );
  });

  it('clears a field named in clearFields by patching it to null', async () => {
    const makeReq = stubUpdate(client, storedCard());
    await client.updateContact('C1', { clearFields: ['emails', 'notes'] });
    const patch = patchFrom(makeReq);
    assert.equal(patch.emails, null);
    assert.equal(patch.notes, null);
  });

  it('rejects a field that is both supplied and listed in clearFields', async () => {
    await assert.rejects(
      () => client.updateContact('C1', { emails: [{ address: 'a@b.example' }], clearFields: ['emails'] }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /both set and clear emails/);
        return true;
      },
    );
  });

  it('rejects clearing a field that is not clearable, naming the ones that are', async () => {
    // `name` is deliberately absent from the clearable set: it is how the contact is
    // identified in every listing.
    await assert.rejects(
      () => client.updateContact('C1', { clearFields: ['name'] }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Cannot clear "name"/);
        assert.match(err.message, /emails, phones, addresses, notes/);
        return true;
      },
    );
  });

  it('refuses to update a contact group', async () => {
    const makeReq = stubUpdate(client, { id: 'G1', kind: 'group', name: { full: 'Team' }, members: { 'uid-1': true } });
    await assert.rejects(
      () => client.updateContact('G1', { name: { full: 'Renamed' } }),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /GROUP/);
        return true;
      },
    );
    assert.equal(makeReq.mock.calls.length, 1, 'a group must be refused before any write');
  });

  it('echoes the pre-edit card and returns the read-back card', async () => {
    const before = storedCard();
    const after = storedCard({ notes: { n0: { note: 'hello' } } });
    stubUpdate(client, before, { updatedCard: after });

    const result = await client.updateContact('C1', { notes: 'hello' });
    assert.deepEqual(result.previousCard, before);
    assert.deepEqual(result.contact, after);
  });

  it('still reports the pre-edit card when the server drops the read-back', async () => {
    // A dropped trailing method is a benign degrade: the write landed. `contact` is then
    // absent rather than invented, and the tool layer says so.
    stubUpdate(client, storedCard(), { omitReadBack: true });
    const result = await client.updateContact('C1', { notes: 'hello' });
    assert.equal(result.contact, undefined);
    assert.deepEqual(result.previousCard, storedCard());
  });

  it('still reports the pre-edit card when the read-back comes back as an error entry', async () => {
    // RFC 8620 section 3.4: a method that fails returns an `error` ENTRY in its slot, not an
    // absence. A hard read of that slot would throw AFTER the write landed, reporting a
    // failure that did not happen and destroying `previousCard` on the way out.
    mock.method(client, 'makeRequest', async (req: any) => {
      if (req.methodCalls[0][0] === 'ContactCard/get') {
        return { methodResponses: [['ContactCard/get', { list: [storedCard()] }, 'card']] };
      }
      return {
        methodResponses: [
          ['ContactCard/set', { updated: { C1: null } }, 'updateContact'],
          ['error', { type: 'invalidArguments' }, 'updatedCard'],
        ],
      };
    });

    const result = await client.updateContact('C1', { notes: 'hello' });
    assert.equal(result.contact, undefined);
    assert.deepEqual(result.previousCard, storedCard());
  });

  it('refuses to call an update successful when the server acknowledged neither outcome', async () => {
    // RFC 8620 section 5.3 puts the id in `updated` or in `notUpdated`. Neither means the
    // server did not say what happened, and reporting success there would invent one.
    stubUpdate(client, storedCard(), { setResult: {} });
    await assert.rejects(
      () => client.updateContact('C1', { notes: 'x' }),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /neither confirmed nor refused/);
        assert.match(err.message, /get_contact/);
        return true;
      },
    );
  });

  it('never issues a request when no contacts primary account exists', async () => {
    client = makeClient({ contactsPrimary: false });
    const makeReq = stubUpdate(client, storedCard());
    await assert.rejects(() => client.updateContact('C1', { notes: 'x' }), /primary account/i);
    assert.equal(makeReq.mock.calls.length, 0);
  });

  it('names the read-only scope in a refused update, so the fix is stated', async () => {
    stubUpdate(client, storedCard(), {
      setResult: { notUpdated: { C1: { type: 'forbidden', description: 'read-only scope' } } },
    });
    await assert.rejects(
      () => client.updateContact('C1', { notes: 'x' }),
      (err: Error) => {
        assert.match(err.message, /Failed to update contact: forbidden - read-only scope/);
        return true;
      },
    );
  });
});

// ---------- deleteContact ----------

describe('deleteContact', () => {
  let client: ContactsCalendarClient;

  beforeEach(() => {
    client = makeClient();
  });

  // The destroy request reads the card and destroys it in ONE batch, get first, so the
  // response carries a ContactCard/get before the ContactCard/set.
  function destroyResponse(card: any, setResult: any) {
    return {
      methodResponses: [
        ['ContactCard/get', { list: card ? [card] : [], ...(card ? {} : { notFound: ['ghost'] }) }, 'doomedCard'],
        ['ContactCard/set', setResult, 'deleteContact'],
      ],
    };
  }

  it('destroys by id', async () => {
    const makeReq = stubMakeRequest(client, destroyResponse({ id: 'C1' }, { destroyed: ['C1'] }));
    await client.deleteContact('C1');
    assert.deepEqual(callArguments(makeReq)[0].methodCalls[1][1].destroy, ['C1']);
  });

  it('reads the card in the same request as the destroy, before it', async () => {
    // A card fetched in an earlier round trip could have changed before the destroy landed,
    // so the echo would not be what was actually deleted. Ordering the get ahead of the set
    // inside one request is what makes the echo the destroyed card.
    const makeReq = stubMakeRequest(client, destroyResponse({ id: 'C1' }, { destroyed: ['C1'] }));
    await client.deleteContact('C1');
    assert.equal(makeReq.mock.calls.length, 1);
    const [get, set] = callArguments(makeReq)[0].methodCalls;
    assert.equal(get[0], 'ContactCard/get');
    assert.equal(set[0], 'ContactCard/set');
    // No properties filter: a recreate needs every field the account stores.
    assert.ok(!('properties' in get[1]), 'the pre-destroy read must not filter properties');
  });

  it('returns the full pre-destroy card, so a wrong delete leaves the card visible', async () => {
    const card = {
      id: 'C1',
      '@type': 'Card',
      uid: 'u-1',
      name: { full: 'Ada Lovelace', components: [{ kind: 'given', value: 'Ada' }] },
      emails: { '0dd713dd': { address: 'ada@example.com', contexts: { private: true }, pref: 1 } },
      notes: { n0: { note: 'keep me' } },
    };
    stubMakeRequest(client, destroyResponse(card, { destroyed: ['C1'] }));

    const result = await client.deleteContact('C1');
    // Deep-equal against the whole card, not a field spot-check: the point of the echo is
    // that NOTHING was trimmed on the way out.
    assert.deepEqual(result.deletedCard, card);
  });

  it('maps notFound destroy errors to the not-found convention', async () => {
    stubMakeRequest(client, destroyResponse(null, { notDestroyed: { ghost: { type: 'notFound' } } }));
    await assert.rejects(() => client.deleteContact('ghost'), /Contact not found: ghost/);
  });

  it('classes a notFound destroy error as caller-fixable input', async () => {
    stubMakeRequest(client, destroyResponse(null, { notDestroyed: { ghost: { type: 'notFound' } } }));
    await assert.rejects(
      () => client.deleteContact('ghost'),
      (err: Error) => {
        assert.equal(err.name, 'InvalidInputError');
        assert.match(err.message, /Contact not found: ghost/);
        return true;
      },
    );
  });

  it('reports a completed destroy with no readable card as a degrade, never as not-found', async () => {
    // The destroy SUCCEEDED — the server listed the id in `destroyed`, so the id was real.
    // Throwing "Contact not found" here would report a failure for a completed irreversible
    // write and tell the caller its id was wrong about a card that was found and destroyed.
    // The honest result is the delete plus an absent echo, which the tool states loudly.
    stubMakeRequest(client, destroyResponse(null, { destroyed: ['ghost'] }));
    const result = await client.deleteContact('ghost');
    assert.deepEqual(result, { deletedCard: undefined });
  });

  it('survives an error entry in the leading read rather than throwing after the destroy', async () => {
    // RFC 8620 section 3.4: a failed method occupies its slot with an `error` entry. A hard
    // read of slot 0 would throw here — after the card was already destroyed.
    stubMakeRequest(client, {
      methodResponses: [
        ['error', { type: 'invalidArguments' }, 'doomedCard'],
        ['ContactCard/set', { destroyed: ['C1'] }, 'deleteContact'],
      ],
    });
    const result = await client.deleteContact('C1');
    assert.equal(result.deletedCard, undefined);
  });

  it('refuses to call a destroy successful when the server acknowledged neither outcome', async () => {
    stubMakeRequest(client, destroyResponse({ id: 'C1' }, {}));
    await assert.rejects(
      () => client.deleteContact('C1'),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /neither confirmed nor refused/);
        return true;
      },
    );
  });

  it('surfaces other notDestroyed errors with their type', async () => {
    stubMakeRequest(
      client,
      destroyResponse({ id: 'C1' }, { notDestroyed: { C1: { type: 'forbidden', description: 'read-only scope' } } }),
    );
    await assert.rejects(() => client.deleteContact('C1'), /forbidden.*read-only scope/s);
  });

  it('names the read-only scope in a refused delete, so the fix is stated', async () => {
    // A token issued with read-only contacts access looks identical to a read-write one on
    // the session, so the refusal itself is the only place the caller learns why.
    stubMakeRequest(
      client,
      destroyResponse({ id: 'C1' }, { notDestroyed: { C1: { type: 'forbidden', description: 'read-only scope' } } }),
    );
    await assert.rejects(
      () => client.deleteContact('C1'),
      (err: Error) => {
        assert.match(err.message, /read-only scope/);
        return true;
      },
    );
  });

  it('keeps a forbidden destroy error a server-side error', async () => {
    // A permissions failure is not fixed by re-forming the arguments, so it stays out of
    // the caller-fixable class even though the sibling notFound case is in it.
    stubMakeRequest(
      client,
      destroyResponse({ id: 'C1' }, { notDestroyed: { C1: { type: 'forbidden', description: 'read-only scope' } } }),
    );
    await assert.rejects(
      () => client.deleteContact('C1'),
      (err: Error) => {
        assert.notEqual(err.name, 'InvalidInputError');
        assert.match(err.message, /Failed to delete contact: forbidden - read-only scope/);
        return true;
      },
    );
  });

  it('passes expectState through as ifInState', async () => {
    const makeReq = stubMakeRequest(client, destroyResponse({ id: 'C1' }, { destroyed: ['C1'] }));
    await client.deleteContact('C1', 'state-7');
    assert.equal(callArguments(makeReq)[0].methodCalls[1][1].ifInState, 'state-7');
  });
});
