import { JmapClient, JmapRequest, QueryResult } from './jmap-client.js';
import { InvalidInputError } from './coerce.js';

export class ContactsCalendarClient extends JmapClient {
  
  private async checkContactsPermission(): Promise<boolean> {
    const session = await this.getSession();
    return !!session.capabilities['urn:ietf:params:jmap:contacts'];
  }
  
  /**
   * Contacts may live on a different primary account than mail, so every
   * contacts method addresses the contacts primary account.
   *
   * There is deliberately no fall back to the mail account when the session
   * reports no contacts primary: the writes (create/update/delete) would then
   * silently target the mail account, and deleteContact has no existence
   * pre-check to catch it. Throwing keeps a misrouted write impossible.
   */
  private async contactsAccountId(): Promise<string> {
    const session = await this.getSession();
    const accountId = session.primaryAccounts?.['urn:ietf:params:jmap:contacts'];
    if (!accountId) {
      throw new Error(
        'No contacts account is available: this JMAP session reports no primary account for ' +
        '"urn:ietf:params:jmap:contacts", so there is no account for contacts operations to address. ' +
        'This usually means the API token lacks the contacts scope. Run check_function_availability ' +
        'to see the reported contacts status, then enable contacts for the token under Fastmail ' +
        'Settings > Privacy & Security > Connected Apps & API tokens.'
      );
    }
    return accountId;
  }

  async getContacts(limit: number = 50): Promise<QueryResult> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const accountId = await this.contactsAccountId();

    // Try CardDAV namespace first, then Fastmail specific
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId,
          limit,
          calculateTotal: true
        }, 'query'],
        ['ContactCard/get', {
          accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
          // No properties filter — return all fields so verbose mode works
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getQueryResult(response, 0, 1);
    } catch (error) {
      // Fallback: try to get contacts using AddressBook methods
      const fallbackRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
        methodCalls: [
          ['AddressBook/get', {
            accountId
          }, 'addressbooks']
        ]
      };

      try {
        const fallbackResponse = await this.makeRequest(fallbackRequest);
        const items = this.getListResult(fallbackResponse, 0);
        return { items };
      } catch (fallbackError) {
        throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
      }
    }
  }

  async getContactById(id: string): Promise<any> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const accountId = await this.contactsAccountId();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/get', {
          accountId,
          ids: [id]
        }, 'contact']
      ]
    };

    let contact;
    try {
      const response = await this.makeRequest(request);
      contact = this.getListResult(response, 0)[0];
    } catch (error) {
      throw new Error(`Contact access not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
    // ContactCard/get reports unknown ids via notFound, leaving list empty — a
    // bare undefined here used to serialize as a successful empty tool response.
    // InvalidInputError (not a plain Error) because a wrong id is the caller's
    // to fix: it maps to InvalidParams at the MCP boundary rather than the
    // InternalError that reads as a server bug.
    if (!contact) {
      throw new InvalidInputError(`Contact not found: ${id}`);
    }
    return contact;
  }

  async searchContacts(query: string, limit: number = 20): Promise<QueryResult> {
    // Check permissions first
    const hasPermission = await this.checkContactsPermission();
    if (!hasPermission) {
      throw new Error('Contacts access not available. This account may not have JMAP contacts permissions enabled. Please check your Fastmail account settings or contact support to enable contacts API access.');
    }

    const accountId = await this.contactsAccountId();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/query', {
          accountId,
          filter: { text: query },
          limit,
          calculateTotal: true
        }, 'query'],
        ['ContactCard/get', {
          accountId,
          '#ids': { resultOf: 'query', name: 'ContactCard/query', path: '/ids' },
          // No properties filter — return all fields so verbose mode works
        }, 'contacts']
      ]
    };

    try {
      const response = await this.makeRequest(request);
      return this.getQueryResult(response, 0, 1);
    } catch (error) {
      throw new Error(`Contact search not supported: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  // There are deliberately no JMAP calendar methods in this class: the calendar
  // tools run over CalDAV (CalDAVCalendarClient in caldav-client.ts), which is the
  // only calendar path index.ts routes to. The JMAP Calendar/CalendarEvent methods
  // that used to sit here had no callers, so they are removed rather than kept as a
  // second, untested implementation a reader could mistake for the live one.

  // ---------- contacts write (JMAP ContactCard/set, RFC 9610) ----------
  //
  // A live probe confirmed Fastmail accepts ContactCard/set with an RFC 9610
  // Card shape; the server assigns the default address book, uid, and prodId.
  // Note: creation-id references ("#id") are NOT resolved in destroy arrays by
  // Fastmail's backend — always destroy by real id.

  /** Map the flat tool-facing input onto an RFC 9610 Card (arrays -> Id-maps). */
  private buildCardProperties(input: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
  }): Record<string, any> {
    const card: Record<string, any> = {};

    if (input.name) {
      const components: Array<{ kind: string; value: string }> = [];
      if (input.name.given) components.push({ kind: 'given', value: input.name.given });
      if (input.name.surname) components.push({ kind: 'surname', value: input.name.surname });
      card.name = {
        ...(components.length && { components }),
        ...(input.name.full && { full: input.name.full }),
      };
    }
    const toIdMap = (items: any[] | undefined, prefix: string) => {
      if (!items?.length) return undefined;
      const map: Record<string, any> = {};
      items.forEach((item, i) => { map[`${prefix}${i}`] = item; });
      return map;
    };
    const emails = toIdMap(input.emails, 'e');
    const phones = toIdMap(input.phones, 'p');
    const addresses = toIdMap(input.addresses, 'a');
    if (emails) card.emails = emails;
    if (phones) card.phones = phones;
    if (addresses) card.addresses = addresses;
    if (input.notes) card.notes = { n0: { note: input.notes } };

    return card;
  }

  async createContact(input: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
    addressBookId?: string;
  }): Promise<string> {
    const hasName = !!(input.name?.full || input.name?.given || input.name?.surname);
    if (!hasName && !input.emails?.length) {
      throw new Error('A contact needs a name or at least one email address');
    }

    const accountId = await this.contactsAccountId();
    const card: Record<string, any> = {
      '@type': 'Card',
      version: '1.0',
      ...this.buildCardProperties(input),
      ...(input.addressBookId && { addressBookIds: { [input.addressBookId]: true } }),
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', { accountId, create: { newContact: card } }, 'createContact'],
      ],
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    if (result.notCreated?.newContact) {
      const err = result.notCreated.newContact;
      throw new Error(`Failed to create contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
    const id = result.created?.newContact?.id;
    if (!id) {
      throw new Error('Contact creation returned no id');
    }
    return id;
  }

  async updateContact(id: string, patch: {
    name?: { given?: string; surname?: string; full?: string };
    emails?: Array<{ address: string; label?: string }>;
    phones?: Array<{ number: string; label?: string }>;
    addresses?: Array<{ full: string; label?: string }>;
    notes?: string;
    expectState?: string;
  }): Promise<void> {
    const { expectState, ...fields } = patch;
    const patchObject = this.buildCardProperties(fields);
    if (Object.keys(patchObject).length === 0) {
      throw new Error('At least one field to update must be provided (name, emails, phones, addresses, or notes)');
    }

    const accountId = await this.contactsAccountId();

    // Existence check first, for a clean not-found error (repo convention).
    const getResponse = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [['ContactCard/get', { accountId, ids: [id], properties: ['id'] }, 'g']],
    });
    if (!this.getListResult(getResponse, 0)[0]) {
      throw new Error(`Contact not found: ${id}`);
    }

    // JMAP PatchObject semantics: each provided top-level field wholly
    // replaces the stored value (e.g. emails: [] clears all emails).
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId,
          update: { [id]: patchObject },
          ...(expectState && { ifInState: expectState }),
        }, 'updateContact'],
      ],
    });
    const result = this.getMethodResult(response, 0);
    if (result.notUpdated?.[id]) {
      const err = result.notUpdated[id];
      throw new Error(`Failed to update contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
  }

  async deleteContact(id: string, expectState?: string): Promise<void> {
    const accountId = await this.contactsAccountId();
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId,
          destroy: [id],
          ...(expectState && { ifInState: expectState }),
        }, 'deleteContact'],
      ],
    });
    const result = this.getMethodResult(response, 0);
    if (result.notDestroyed?.[id]) {
      const err = result.notDestroyed[id];
      if (err.type === 'notFound') {
        throw new Error(`Contact not found: ${id}`);
      }
      throw new Error(`Failed to delete contact: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }
  }
}
