import {
  InvalidInputError,
  coerceBool,
  coerceContactAddresses,
  coerceContactEmails,
  coerceContactName,
  coerceContactPhones,
  coerceStringArrayStrict,
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
  deleteContact(id: string): Promise<DeleteContactResult>;
}

export type ToolContent = Array<{ type: 'text'; text: string }>;

// The pre-edit and pre-destroy echoes are always the raw card, and are not a restore; see
// docs/conventions.md.

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
 * create_contact. Reads the card back so the tool returns the `get_contact` shape, including
 * the server-assigned id, uid and prodId.
 */
export async function createContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const verbose = coerceBool(args?.verbose, 'verbose') ?? false;

  const id = await client.createContact({
    name: coerceContactName(args?.name),
    emails: coerceContactEmails(args?.emails),
    phones: coerceContactPhones(args?.phones),
    addresses: coerceContactAddresses(args?.addresses),
    notes: coerceContactNotes(args?.notes, 'omit it to create the contact without a note.'),
    addressBookId: coerceAddressBookId(args?.addressBookId),
  });

  let card: any;
  try {
    card = await client.getContactById(id);
  } catch {
    // Not a throw: the create has happened, and a reported failure invites a duplicating retry.
    return [
      { type: 'text', text: toolJson({ id }) },
      {
        type: 'text',
        text:
          `The contact was created (id ${id}), but reading it back failed, so only its id is shown. ` +
          `Do not create it again, which would make a duplicate; read it with get_contact.`,
      },
    ];
  }
  return [{ type: 'text', text: toolJson(renderCard(card, raw, verbose)) }];
}

/** update_contact. Returns `{contact, previousCard}` in every mode. */
export async function updateContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const verbose = coerceBool(args?.verbose, 'verbose') ?? false;
  const contactId = requireContactId(args);

  const result = await client.updateContact(contactId, {
    name: coerceContactName(args?.name),
    emails: coerceContactEmails(args?.emails),
    phones: coerceContactPhones(args?.phones),
    addresses: coerceContactAddresses(args?.addresses),
    notes: coerceContactNotes(args?.notes, `to remove the note pass clearFields:['notes'].`),
    // Strict: an ignored clearFields would report a clear that never happened.
    clearFields: coerceStringArrayStrict(args?.clearFields, 'clearFields'),
    allowEntryReplace: coerceBool(args?.allowEntryReplace, 'allowEntryReplace') ?? false,
  });

  const envelope: Record<string, any> = {};
  if (result.contact !== undefined) envelope.contact = renderCard(result.contact, raw, verbose);
  envelope.previousCard = result.previousCard;

  const content: ToolContent = [{ type: 'text', text: toolJson(envelope) }];
  if (result.contact === undefined) {
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

export async function getContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const verbose = coerceBool(args?.verbose, 'verbose') ?? false;
  const contactId = requireContactId(args);
  const contact = await client.getContactById(contactId);
  return [{ type: 'text', text: toolJson(raw ? contact : simplifyContact(contact, { verbose })) }];
}

/**
 * delete_contact. Takes no `raw`/`verbose`: its only card, `deletedCard`, is always
 * untransformed, so either parameter would be a silent no-op.
 */
export async function deleteContactTool(args: any, client: ContactsWriteClient): Promise<ToolContent> {
  const contactId = requireContactId(args);
  const { deletedCard } = await client.deleteContact(contactId);

  const content: ToolContent = [
    { type: 'text', text: toolJson({ deleted: contactId, deletedCard }) },
  ];
  if (deletedCard === undefined) {
    // Not a throw: that would report a failure for a completed irreversible write.
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
