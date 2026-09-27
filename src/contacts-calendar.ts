import { JmapClient, JmapRequest, QueryResult } from './jmap-client.js';
import {
  InvalidInputError,
  validateClearFields,
  type ContactAddressSpec,
  type ContactEmailSpec,
  type ContactNameSpec,
  type ContactPhoneSpec,
} from './coerce.js';
import {
  assertUnambiguousEntryEdit,
  buildEntryMap,
  contactKindRefusal,
  isAmbiguousEntryEdit,
  refusedContactKind,
  mergeContactName,
  mergeContactNotes,
  mergeEntryMap,
  type EntryKeyField,
} from './contact-card.js';

/** The fields `update_contact` can blank, mirrored by the tool's `clearFields` schema enum. */
export const CLEARABLE_CONTACT_FIELDS: ReadonlySet<string> = new Set(['emails', 'phones', 'addresses', 'notes']);

/** The patch `update_contact` applies, after coercion. */
export interface UpdateContactPatch {
  name?: ContactNameSpec;
  emails?: ContactEmailSpec[];
  phones?: ContactPhoneSpec[];
  addresses?: ContactAddressSpec[];
  notes?: string;
  clearFields?: string[];
  allowEntryReplace?: boolean;
}

export interface UpdateContactResult {
  /**
   * The card exactly as it stood BEFORE the update, untransformed. Always present: the merge
   * has to fetch it anyway, so echoing it costs nothing and makes an unintended overwrite
   * both visible and undoable.
   */
  previousCard: any;
  /**
   * The card after the update, untransformed. Absent only if the server did not return the
   * read-back that rides in the same request as the write — the write still succeeded, and
   * the caller is told so rather than being handed a silently missing field.
   */
  contact?: any;
}

export interface DeleteContactResult {
  /**
   * The card exactly as it stood before the destroy, untransformed. Absent only when the
   * read that rides ahead of the destroy in the same request produced nothing — the contact
   * is gone either way, so this is a degraded result the tool reports loudly, never a
   * failure and never a not-found.
   */
  deletedCard?: any;
}

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
        'Settings → Privacy & Security, in the "Connected apps & API tokens" section under "Manage API tokens".'
      );
    }
    return accountId;
  }

  async getContacts(limit: number = 50): Promise<QueryResult> {
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
      // No AddressBook/get fallback: it would hide the real ContactCard/query error behind
      // address books dressed up as contacts. A failed contacts query surfaces as a failure,
      // never as a different kind of record.
      throw new Error(`Contacts not supported or accessible: ${error instanceof Error ? error.message : String(error)}. Try checking account permissions or enabling contacts API access in Fastmail settings.`);
    }
  }

  async getContactById(id: string): Promise<any> {
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
    // ContactCard/get reports an unknown id via notFound with an empty list; a bare undefined
    // would serialize as a successful empty tool response. InvalidInputError because a wrong
    // id is the caller's to fix (InvalidParams, not InternalError).
    if (!contact) {
      throw new InvalidInputError(`Contact not found: ${id}`);
    }
    return contact;
  }

  async searchContacts(query: string, limit: number = 20): Promise<QueryResult> {
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

  // There are deliberately no JMAP calendar methods in this class: the calendar tools run
  // over CalDAV (CalDAVCalendarClient in caldav-client.ts), the only calendar path index.ts
  // routes to.

  // ---------- contacts write (JMAP ContactCard/set, RFC 9610) ----------
  //
  // Fastmail accepts ContactCard/set with an RFC 9610 Card shape (live probe) and assigns
  // the default address book, uid, and prodId. Creation-id references ("#id") are NOT
  // resolved in destroy arrays by Fastmail's backend: always destroy by real id.

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
    // An empty array is refused for the same reason updateContact refuses one, and
    // `buildCardProperties` would silently omit the field rather than say so. There is nothing
    // to clear on a create, so the route out is to omit the parameter.
    for (const field of ['emails', 'phones', 'addresses'] as const) {
      const value = input[field];
      if (value && value.length === 0) {
        throw new InvalidInputError(
          `${field}: [] is not accepted. An empty array is indistinguishable from a mistake that ` +
            `produced no entries. Omit ${field} to create the contact without any.`,
        );
      }
    }

    const hasName = !!(input.name?.full || input.name?.given || input.name?.surname);
    if (!hasName && !input.emails?.length) {
      // A missing required field is caller-fixable, so it must reach the MCP
      // boundary as InvalidParams rather than the InternalError a plain Error maps to.
      throw new InvalidInputError('A contact needs a name or at least one email address');
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
      this.throwSingleSetError(result.notCreated.newContact, 'create contact');
    }
    const id = result.created?.newContact?.id;
    if (!id) {
      throw new Error('Contact creation returned no id');
    }
    return id;
  }

  /**
   * Fetch one card WHOLE (no `properties` filter, so every field the account stores comes
   * back). Both writes below need this: the update merges against it and echoes it, and the
   * delete echoes it. A `properties: ['id']` existence probe would answer "does it exist"
   * and nothing else, which is not enough to merge with, nor to show a caller what a write
   * took off the card.
   *
   * `card` is undefined for an id the account does not hold; callers decide what that means.
   * `state` is the ContactCard state the card was read at. Both writes send it as `ifInState`,
   * so a card changed between this read and the write is refused rather than overwritten
   * from a stale merge. Fastmail's state is account-wide, so a change to ANY card in that
   * window refuses too; a retry re-reads and succeeds.
   */
  private async fetchCard(accountId: string, id: string): Promise<{ card: any | undefined; state?: string }> {
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [['ContactCard/get', { accountId, ids: [id] }, 'card']],
    });
    const card = this.getListResult(response, 0)[0];
    const state = this.getMethodResult(response, 0).state;
    return { card, state: typeof state === 'string' ? state : undefined };
  }

  /** Refuse before an unguarded write: without the read's state, a stale merge would go through. */
  private requireReadState(state: string | undefined, id: string, outcome: string): string {
    if (state === undefined) {
      throw new Error(
        `The server returned no ContactCard state with contact ${id}, so the write could not be guarded ` +
          `against a change made since the read; nothing was ${outcome}.`,
      );
    }
    return state;
  }

  /** Throw the retry refusal when a write's `ifInState` no longer matched (RFC 8620 section 5.3). */
  private assertStateStillMatched(response: any, index: number, id: string, tool: string, outcome: string): void {
    const entry = response.methodResponses?.[index];
    if (entry?.[0] === 'error' && entry[1]?.type === 'stateMismatch') {
      throw new Error(
        `The contact ${id} changed since it was read; nothing was ${outcome}. Retry the ${tool} call: ` +
          `it re-reads the contact first. The contacts state is account-wide on Fastmail, so a change ` +
          `to any contact in between also causes this refusal.`,
      );
    }
  }

  /**
   * Build the Id-map to write for `emails` or `phones`.
   *
   * Default (merge): the supplied array says which entries exist, and every entry that
   * matches one already on the card keeps the fields the simplified read shape never showed
   * — `contexts`, `pref`, `@type` and anything else — under its existing map key.
   *
   * `allowEntryReplace`: every entry is written FRESH from what the caller supplied, so those
   * hidden fields do not carry. That is the whole meaning of the flag, and the ambiguity
   * rejection says so before a caller reaches for it.
   */
  private buildEntryPatch(
    field: 'emails' | 'phones',
    existing: any,
    incoming: Array<{ address?: string; number?: string; label?: string }>,
    keyField: EntryKeyField,
    allowEntryReplace: boolean,
  ): Record<string, any> {
    const prefix = keyField === 'address' ? 'e' : 'p';
    // Rebuild each entry from the validated keys rather than passing the caller's object
    // through, so a key added to the spec later has to be handled here deliberately.
    const fresh = incoming.map((item) => {
      const entry: Record<string, any> = { [keyField]: item[keyField] };
      if (item.label !== undefined) entry.label = item.label;
      return entry;
    });

    // Merge FIRST, always, and consult the flag only if this field's merge is actually
    // ambiguous. `allowEntryReplace` is scoped to the field that could not be resolved, not
    // to the whole call: a caller that hits the rejection on `emails` and resends the same
    // call with the flag would otherwise silently lose the contexts/pref on a `phones` array
    // it was never warned about. The rejection text already promises this scoping ("to
    // REPLACE the <field>").
    const outcome = mergeEntryMap(existing, fresh, keyField);
    if (!isAmbiguousEntryEdit(outcome)) return outcome.map;

    if (!allowEntryReplace) assertUnambiguousEntryEdit(field, outcome);

    // `allowEntryReplace` is an intent MARKER, not the safety control, and must not be
    // mistaken for one. A model retries a rejected call with a flag as readily as it
    // retries with corrected arguments, so a flag can record that a lossy write was meant
    // — it cannot prevent one. What actually makes a flagged-through mistake recoverable
    // is `previousCard`: every update returns the whole pre-edit card, hidden entry fields
    // included, so whatever the replace discarded can be put back. Do not "strengthen"
    // this into a confirmation handshake (it would add friction and prevent nothing), and
    // do not drop the pre-edit echo as redundant — the echo is the mitigation.
    return buildEntryMap(fresh, prefix);
  }

  /**
   * Update a contact by MERGING per entry, and echo the card as it stood beforehand.
   *
   * The card is fetched whole first because a JMAP PatchObject replaces a top-level property
   * outright: writing `emails` from the tool's flat array alone would silently discard the
   * `contexts` and `pref` that sit on nearly every real entry. So the supplied arrays decide
   * which entries exist, and this decides what each surviving entry still carries.
   *
   * An empty array is REJECTED rather than read as "clear" — see the rejection below.
   */
  async updateContact(id: string, patch: UpdateContactPatch): Promise<UpdateContactResult> {
    const { clearFields, allowEntryReplace = false, name, emails, phones, addresses, notes } = patch;

    const provided = new Set<string>();
    for (const [field, value] of Object.entries({ name, emails, phones, addresses, notes })) {
      if (value !== undefined) provided.add(field);
    }
    // Same conflict rule as edit_draft: a field cannot be both supplied and cleared in one
    // call, and only the fields in the runtime set are clearable. `name` is deliberately not
    // in that set — see the tool description.
    validateClearFields(clearFields, CLEARABLE_CONTACT_FIELDS, provided);

    // An empty array is refused rather than treated as "remove them all". JMAP would accept
    // it, but the caller cannot tell an intended wipe from a mapping bug that produced no
    // entries, and the two outcomes differ by an entire address book field. `clearFields`
    // says the same thing and can only be written on purpose.
    for (const field of ['emails', 'phones', 'addresses'] as const) {
      const value = patch[field];
      if (value && value.length === 0) {
        throw new InvalidInputError(
          `${field}: [] is not accepted. An empty array is indistinguishable from a mistake that ` +
            `produced no entries, so it does not clear the field. To remove every ${field} entry, ` +
            `pass clearFields:['${field}'].`,
        );
      }
    }
    if (notes !== undefined && notes.trim() === '') {
      throw new InvalidInputError(
        `notes cannot be empty; to remove the note pass clearFields:['notes'], or omit notes to leave it unchanged.`,
      );
    }

    if (provided.size === 0 && !clearFields?.length) {
      throw new InvalidInputError('At least one field to update must be provided (name, emails, phones, addresses, notes, or clearFields)');
    }

    const accountId = await this.contactsAccountId();

    const { card: previousCard, state: readState } = await this.fetchCard(accountId, id);
    if (!previousCard) {
      throw new InvalidInputError(`Contact not found: ${id}`);
    }

    const refusedKind = refusedContactKind(previousCard);
    if (refusedKind) {
      throw new InvalidInputError(contactKindRefusal({
        id,
        kind: refusedKind,
        tool: 'update_contact',
        because:
          'name/emails/phones/addresses/notes describe a person card, this server can create only individual ' +
          'cards, and group members are not editable here.',
        recovery: 'Edit it in the Fastmail web interface instead.',
      }));
    }
    const guardState = this.requireReadState(readState, id, 'written');

    const patchObject: Record<string, any> = {};
    if (name) patchObject.name = mergeContactName(previousCard.name, name);
    if (emails) patchObject.emails = this.buildEntryPatch('emails', previousCard.emails, emails, 'address', allowEntryReplace);
    if (phones) patchObject.phones = this.buildEntryPatch('phones', previousCard.phones, phones, 'number', allowEntryReplace);
    if (addresses) {
      // Addresses have no matchable key — an entry is `{full, label}`-shaped, and two
      // addresses can differ only in punctuation — so there is nothing to match a supplied
      // entry against and the whole property is replaced. Stated plainly in the tool's
      // parameter description, because it is the one array that does not merge.
      patchObject.addresses = buildEntryMap(
        addresses.map((a) => {
          const entry: Record<string, any> = { full: a.full };
          if (a.label !== undefined) entry.label = a.label;
          return entry;
        }),
        'a',
      );
    }
    if (notes !== undefined) patchObject.notes = mergeContactNotes(previousCard.notes, notes);
    // RFC 8620 section 5.3: a PatchObject value of null removes the property.
    for (const field of clearFields ?? []) patchObject[field] = null;

    // The read-back rides in the SAME request as the write, so the merged card comes home in
    // one round trip and is the server's own copy rather than our guess at what it stored.
    // Two method calls, two different single-response methods, so the positional reads are
    // safe (RFC 8620 section 3.4).
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/set', {
          accountId,
          update: { [id]: patchObject },
          ifInState: guardState,
        }, 'updateContact'],
        ['ContactCard/get', { accountId, ids: [id] }, 'updatedCard'],
      ],
    });
    this.assertStateStillMatched(response, 0, id, 'update_contact', 'written');
    const result = this.getMethodResult(response, 0);
    if (result.notUpdated?.[id]) {
      this.throwSingleSetError(result.notUpdated[id], 'update contact');
    }
    // RFC 8620 section 5.3 requires the id in exactly one of `updated`/`notUpdated`. A
    // response carrying it in neither is not a success this client may report as one. Note
    // `updated[id]` is legitimately `null` (the server changed nothing extra), so this asks
    // whether the KEY is present, never whether the value is truthy.
    if (!result.updated || !Object.prototype.hasOwnProperty.call(result.updated, id)) {
      throw new Error(
        `The server neither confirmed nor refused the update of contact ${id}: it reported the id in ` +
          `neither its updated nor its notUpdated map. Re-read the contact with get_contact to see ` +
          `whether the change was applied.`,
      );
    }

    // Read the trailing method tolerantly: it tolerates both a dropped method and an `error`
    // entry, because the write has already happened either way and turning an auxiliary
    // read-back failure into a thrown error would report a failure that did not occur.
    // It is NOT dropped silently — `contact` is then absent and the tool says so.
    const contact = this.readListResultIfPresent(response, 1)[0];
    return contact ? { previousCard, contact } : { previousCard };
  }

  /**
   * Delete a contact, returning the card exactly as it stood immediately before the destroy.
   *
   * There is deliberately NO confirmation parameter. A JMAP destroy is irreversible — a
   * ContactCard does not go to a trash the way an email does — but a confirm flag would not
   * prevent a wrong delete: a caller that got the id wrong passes a confirmation as readily as
   * it passed the id. The mitigation that actually works is this echo. The read is issued in
   * the SAME request as the destroy and ordered before it, so what comes back is the card that
   * was destroyed (not a copy fetched earlier that something else could have changed since).
   *
   * The echo is best-effort BY DESIGN, and never a reason to fail the call. Once the destroy
   * has happened it cannot be taken back, so a leading read that came home absent or as an
   * `error` entry must not throw: that would report a failure for a completed irreversible
   * write and discard the only thing the caller could still act on. `deletedCard` is then
   * undefined and the tool states the degrade.
   */
  async deleteContact(id: string): Promise<DeleteContactResult> {
    const accountId = await this.contactsAccountId();

    // The kind refusal (CONTRIBUTING.md, "A destroy must not remove what the server cannot
    // recreate") costs its own round trip: a JMAP batch cannot make one method conditional on
    // another's result, so the card has to be read in a request that completes before the
    // destroy is sent. The echo still comes from the read inside the destroy batch, so it
    // remains the card as it stood at the moment it was destroyed. A card that cannot be read
    // at ALL fails here, before anything is destroyed, where there is nothing yet to lose by
    // throwing. A card the account simply does not hold reads as undefined and falls through
    // to the destroy, whose own `notFound` is the authoritative answer for a bad id.
    const { card: doomedCard, state: readState } = await this.fetchCard(accountId, id);
    const refusedKind = refusedContactKind(doomedCard);
    if (refusedKind) {
      throw new InvalidInputError(contactKindRefusal({
        id,
        kind: refusedKind,
        tool: 'delete_contact',
        because:
          'this server can create only individual cards (it cannot create a group or any other kind; ' +
          'create_contact has no kind or members parameter), so it will not destroy one it could never ' +
          'put back, and the deletedCard echo could not rebuild it either.',
        recovery: 'Delete it in the Fastmail web interface instead.',
      }));
    }
    const guardState = this.requireReadState(readState, id, 'deleted');

    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/get', { accountId, ids: [id] }, 'doomedCard'],
        ['ContactCard/set', {
          accountId,
          destroy: [id],
          ifInState: guardState,
        }, 'deleteContact'],
      ],
    });
    this.assertStateStillMatched(response, 1, id, 'delete_contact', 'deleted');

    const result = this.getMethodResult(response, 1);
    if (result.notDestroyed?.[id]) {
      const err = result.notDestroyed[id];
      if (err.type === 'notFound') {
        // Caller-fixable bad id: InvalidParams. A genuinely unknown id lands HERE (RFC 8620
        // section 5.3 puts it in notDestroyed), which is why an unreadable card below is never
        // reported as not-found: by then the destroy has succeeded, so the id was real.
        throw new InvalidInputError(`Contact not found: ${id}`);
      }
      this.throwSingleSetError(err, 'delete contact');
    }
    // Same acknowledgement check as the update: the id must appear in one of the two maps.
    if (!Array.isArray(result.destroyed) || !result.destroyed.includes(id)) {
      throw new Error(
        `The server neither confirmed nor refused the deletion of contact ${id}: it reported the id in ` +
          `neither its destroyed list nor its notDestroyed map. Re-read the contact with get_contact to ` +
          `see whether it still exists.`,
      );
    }

    // Tolerant read, never a throw (see the doc comment). Nor is an empty echo "not found":
    // that would call the id wrong for a card that was found and destroyed.
    const deletedCard = this.readListResultIfPresent(response, 0)[0];
    return { deletedCard };
  }
}
