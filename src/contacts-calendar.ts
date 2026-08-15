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
  contactGroupRefusal,
  isAmbiguousEntryEdit,
  isContactGroupCard,
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
  expectState?: string;
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
    // An empty array is refused here for the same reason update_contact refuses one: it is
    // indistinguishable from a mapping bug that produced no entries, and `buildCardProperties`
    // would silently omit the field rather than say so. There is nothing to clear on a create,
    // so the route out is simply to omit the parameter. Keeping both tools on one rule is
    // what makes "[] is never accepted on a contact entry array" a rule rather than a
    // per-tool detail.
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
      // Same SetError classification the mail writes use: a `notFound` is a bad id the
      // caller can fix (InvalidParams), anything else is an operational failure
      // (InternalError). The rendered message is unchanged.
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
   * Returns undefined for an id the account does not hold; callers decide what that means.
   */
  private async fetchCardOrUndefined(accountId: string, id: string): Promise<any | undefined> {
    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [['ContactCard/get', { accountId, ids: [id] }, 'card']],
    });
    return this.getListResult(response, 0)[0];
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
    // that had nothing ambiguous about it — the flag would then destroy data on a field the
    // caller was never warned about, which is precisely the failure the merge exists to
    // prevent. The rejection text already promises this scoping ("to REPLACE the <field>").
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
    const { expectState, clearFields, allowEntryReplace = false, name, emails, phones, addresses, notes } = patch;

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
      // An empty patch is caller-fixable input, not a server fault.
      throw new InvalidInputError('At least one field to update must be provided (name, emails, phones, addresses, notes, or clearFields)');
    }

    const accountId = await this.contactsAccountId();

    const previousCard = await this.fetchCardOrUndefined(accountId, id);
    if (!previousCard) {
      // A wrong id is the caller's to fix, so this maps to InvalidParams — matching
      // getContactById and the not-found convention across the client.
      throw new InvalidInputError(`Contact not found: ${id}`);
    }

    // A group card holds a `members` map and no emails/phones at all. None of this tool's
    // parameters describe a group, and there is no members surface here to edit one through,
    // so an update aimed at a group is refused rather than half-applied to a record whose
    // shape it does not fit. (The group's membership is untouched either way — a PatchObject
    // only replaces the properties named in it — but a call that reads as "set this group's
    // phone number" has no correct outcome.) deleteContact refuses the same card kind from
    // the same rule; both raise it through contactGroupRefusal so they read as one.
    if (isContactGroupCard(previousCard)) {
      throw new InvalidInputError(contactGroupRefusal({
        id,
        tool: 'update_contact',
        because:
          'its members are not editable through this server, and name/emails/phones/addresses/notes ' +
          'do not describe a group.',
        recovery: 'Edit it in the Fastmail web interface instead.',
      }));
    }

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
          ...(expectState && { ifInState: expectState }),
        }, 'updateContact'],
        ['ContactCard/get', { accountId, ids: [id] }, 'updatedCard'],
      ],
    });
    const result = this.getMethodResult(response, 0);
    if (result.notUpdated?.[id]) {
      this.throwSingleSetError(result.notUpdated[id], 'update contact');
    }
    // RFC 8620 section 5.3 requires the id in exactly one of `updated`/`notUpdated`. A
    // response carrying it in neither is not a success this client may report as one — the
    // difference between "the write landed" and "we assumed it did" is the whole value of
    // saying so. Note `updated[id]` is legitimately `null` (the server changed nothing extra),
    // so this asks whether the KEY is present, never whether the value is truthy.
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
   *
   * A contact GROUP is refused outright — see the guard below.
   */
  async deleteContact(id: string, expectState?: string): Promise<DeleteContactResult> {
    const accountId = await this.contactsAccountId();

    // A destroy must not remove a record this server has no way to recreate, and a group is
    // exactly that: `create_contact` has no `kind` and no `members` parameter, so the echoed
    // card — the whole safety net on this irreversible call — could not rebuild a membership
    // list that on a real card runs to a hundred-odd uids. `update_contact` already refuses a
    // group; this is the same rule from the other side.
    //
    // The check is scoped to the record KIND, not to fields create_contact cannot set. Almost
    // every real card carries titles, organizations or photos that this server cannot write,
    // and refusing to delete all of those would break the tool. What makes a group different
    // is that there is no way to make one at all. Keep it that way as create_contact grows.
    //
    // It costs its own round trip: a JMAP batch cannot make one method conditional on
    // another's result, so the card has to be read in a request that completes before the
    // destroy is sent. The echo still comes from the read inside the destroy batch, so it
    // remains the card as it stood at the moment it was destroyed. A card that cannot be read
    // at ALL fails here, before anything is destroyed — the safe direction on an irreversible
    // call, and unlike the post-destroy read there is nothing yet to lose by throwing. A card
    // the account simply does not hold reads as undefined and falls through to the destroy,
    // whose own `notFound` is the authoritative answer for a bad id.
    const doomedCard = await this.fetchCardOrUndefined(accountId, id);
    if (isContactGroupCard(doomedCard)) {
      throw new InvalidInputError(contactGroupRefusal({
        id,
        tool: 'delete_contact',
        because:
          'this server cannot create a group — create_contact has no kind or members parameter — so it ' +
          'will not destroy one it could never put back, and the deletedCard echo could not rebuild its ' +
          'members either.',
        recovery: 'Delete it in the Fastmail web interface instead.',
      }));
    }

    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:contacts'],
      methodCalls: [
        ['ContactCard/get', { accountId, ids: [id] }, 'doomedCard'],
        ['ContactCard/set', {
          accountId,
          destroy: [id],
          ...(expectState && { ifInState: expectState }),
        }, 'deleteContact'],
      ],
    });

    const result = this.getMethodResult(response, 1);
    if (result.notDestroyed?.[id]) {
      const err = result.notDestroyed[id];
      if (err.type === 'notFound') {
        // Caller-fixable bad id: InvalidParams, not the InternalError a plain Error
        // would map to. The wording matches the other not-found paths. A genuinely unknown
        // id lands HERE (RFC 8620 section 5.3 puts it in notDestroyed), which is why an
        // unreadable card below is never reported as not-found: by then the destroy has
        // succeeded, so the id was real.
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

    // Tolerant read, and deliberately NOT a throw when it comes back empty. The destroy has
    // already happened and cannot be undone, so an unreadable echo is a degraded result, not a
    // failed call — and it is certainly not "not found", which would tell the caller its id
    // was wrong about a card that was found and destroyed.
    const deletedCard = this.readListResultIfPresent(response, 0)[0];
    return { deletedCard };
  }
}
