import {
  InvalidInputError,
  coerceBool,
  coerceContactAddresses,
  coerceContactEmails,
  coerceContactName,
  coerceContactPhones,
  coerceStringArray,
  toolJson,
  type ContactAddressSpec,
  type ContactEmailSpec,
  type ContactNameSpec,
  type ContactPhoneSpec,
} from './coerce.js';
import { simplifyContact } from './response-formatters.js';
import type { DeleteContactResult, UpdateContactPatch, UpdateContactResult } from './contacts-calendar.js';

/**
 * The slice of the contacts client the three write tools need, so each handler can be
 * exercised with a stub instead of a live address book. `ContactsCalendarClient` satisfies it
 * structurally.
 */
export interface ContactsWriteClient {
  createContact(input: {
    name?: ContactNameSpec;
    emails?: ContactEmailSpec[];
    phones?: ContactPhoneSpec[];
    addresses?: ContactAddressSpec[];
    notes?: string;
    addressBookId?: string;
  }): Promise<string>;
  getContactById(id: string): Promise<any>;
  updateContact(id: string, patch: UpdateContactPatch): Promise<UpdateContactResult>;
  deleteContact(id: string, expectState?: string): Promise<DeleteContactResult>;
}

export type ToolContent = Array<{ type: 'text'; text: string }>;

// `expectState` is deliberately NOT a parameter of any of these tools, even though the client
// methods accept it and pass it through as `ifInState`.
//
// It guards a write against a concurrent change by naming the JMAP state string the caller
// expects the account to still be in — but no read tool on this server surfaces that string,
// so a caller has no way to obtain a correct value. The only things it could pass are a guess
// (which fails every write with a stateMismatch) or nothing at all, and a parameter whose only
// honest value is "omit it" is a parameter that misleads. The pre-edit `previousCard` echo
// covers the case it was reached for anyway: an overwrite made from a stale copy is visible in
// the response rather than prevented.
//
// The client argument and its tests stay, so exposing it later is one line. Doing so means
// surfacing `state` on the contacts READS first, then accepting it here — in that order,
// because a guard nobody can supply a value for is worse than no guard.

// The pre-edit and pre-destroy echoes are ALWAYS the untransformed JMAP card, whatever
// `verbose` or `raw` say. Their purpose is to keep visible whatever the write took away, which
// needs every field, including the per-entry `contexts` and `pref` the simplified shape folds
// into a bare string. Those are precisely the fields the merge exists to protect, so a
// simplified echo would drop exactly what the caller most needs to see it lost. `raw`/`verbose`
// therefore govern the CARD the tool returns, never the echo.
//
// What the echo is NOT is a restore. This server writes a name, emails, phones, addresses and
// a note; nothing here can put back photos, titles, organizations, nicknames, URLs,
// anniversaries, group membership, the uid, or a per-entry `contexts`/`pref`. So the echo
// makes a bad write legible and partly repairable, and the rest is a job for a Fastmail
// client. Descriptions and docs must say that, rather than promising a clean recreate.

function coerceContactNotes(value: unknown, hint: string): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    throw new InvalidInputError(`notes must be a string, not ${Array.isArray(value) ? 'an array' : `a ${typeof value}`}.`);
  }
  if (value.trim() === '') {
    throw new InvalidInputError(`notes cannot be empty; ${hint}`);
  }
  return value;
}

function coerceAddressBookId(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  // Same reason the entry keys are type-checked: this reaches the card as
  // `addressBookIds: {[value]: true}`, so a non-string would be stringified into a
  // nonsense key rather than rejected.
  if (typeof value !== 'string' || value.trim() === '') {
    throw new InvalidInputError('addressBookId must be a non-empty string; omit it to use the account default.');
  }
  return value.trim();
}

function requireContactId(args: any): string {
  const contactId = args?.contactId;
  if (typeof contactId !== 'string' || contactId.trim() === '') {
    throw new InvalidInputError('contactId is required');
  }
  return contactId.trim();
}

/** Render a card through the mode the caller asked for. */
function renderCard(card: any, raw: boolean, verbose: boolean): any {
  return raw ? card : simplifyContact(card, { verbose });
}

/**
 * create_contact. The card is read back after the write so the tool returns the same shape
 * `get_contact` does — including the id, uid and prodId the server assigns — rather than a
 * bare id the caller then has to fetch.
 */
export async function createContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw) ?? false;
  const verbose = coerceBool(args?.verbose) ?? false;

  const id = await client.createContact({
    name: coerceContactName(args?.name),
    emails: coerceContactEmails(args?.emails),
    phones: coerceContactPhones(args?.phones),
    addresses: coerceContactAddresses(args?.addresses),
    notes: coerceContactNotes(args?.notes, 'omit it to create the contact without a note.'),
    addressBookId: coerceAddressBookId(args?.addressBookId),
  });

  const card = await client.getContactById(id);
  return [{ type: 'text', text: toolJson(renderCard(card, raw, verbose)) }];
}

/**
 * update_contact. Returns `{contact, previousCard}` in every mode — see the note above for why
 * the echo is not subject to `raw`/`verbose`.
 */
export async function updateContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw) ?? false;
  const verbose = coerceBool(args?.verbose) ?? false;
  const contactId = requireContactId(args);

  const result = await client.updateContact(contactId, {
    name: coerceContactName(args?.name),
    emails: coerceContactEmails(args?.emails),
    phones: coerceContactPhones(args?.phones),
    addresses: coerceContactAddresses(args?.addresses),
    notes: coerceContactNotes(args?.notes, `to remove the note pass clearFields:['notes'].`),
    // Same lenient-client reason as edit_draft's clearFields: a stringified array has to
    // coerce back before the allowed/conflict rules can see it.
    clearFields: coerceStringArray(args?.clearFields),
    allowEntryReplace: coerceBool(args?.allowEntryReplace) ?? false,
  });

  const envelope: Record<string, any> = {};
  if (result.contact !== undefined) envelope.contact = renderCard(result.contact, raw, verbose);
  envelope.previousCard = result.previousCard;

  const content: ToolContent = [{ type: 'text', text: toolJson(envelope) }];
  if (result.contact === undefined) {
    // Never-silent degrade: the write landed, the read-back that rides with it did not come
    // back, so the promised `contact` field is absent and says why rather than vanishing.
    content.push({
      type: 'text',
      text:
        `The contact was updated, but the server did not return the updated card alongside the ` +
        `write, so \`contact\` is absent above. Read it with get_contact (contactId ${contactId}). ` +
        `\`previousCard\` is the card as it stood before this update.`,
    });
  }
  return content;
}

/**
 * delete_contact. Takes no `raw`/`verbose`: the only card it returns is `deletedCard`, which is
 * always untransformed, so either parameter could only be a no-op — and a parameter that
 * quietly does nothing is what the unknown-parameter guard exists to prevent.
 */
export async function deleteContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const contactId = requireContactId(args);
  const { deletedCard } = await client.deleteContact(contactId);

  const content: ToolContent = [
    { type: 'text', text: toolJson({ deleted: contactId, deletedCard }) },
  ];
  if (deletedCard === undefined) {
    // The one degrade that cannot be retried out of: the contact is gone and no copy of it
    // came back. Saying so loudly is all that is left — the alternative, throwing, would
    // report a failure for a completed irreversible write and lose the fact that it happened.
    content.push({
      type: 'text',
      text:
        `WARNING: contact ${contactId} was deleted, but the server did not return the card ` +
        `alongside the destroy, so \`deletedCard\` is absent. The delete is irreversible and ` +
        `no copy of the card came back, so nothing here can tell you what was on it.`,
    });
  }
  return content;
}
