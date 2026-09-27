import { InvalidInputError, coerceStringArray } from './coerce.js';
import type { SimplifiedEmail } from './email-formatter.js';

// Caller-directed output projection for the email read tools (#69, #79). A 66-message
// sweep measured 84KB, of which the five fields the caller wanted were 18%.
//
// It runs AFTER simplifyEmail and is subtractive only, so every guarantee of the
// simplified shape holds for the fields that survive; `raw: true` is never projected.
// Dropping a field the caller excluded is not a "silently dropped promised field": the
// caller asked. The exceptions are the ride-along signals below.

// Typed Record<keyof SimplifiedEmail, true> so a field added to SimplifiedEmail without
// being added here is a compile error. Declaration order matches the interface, so the
// rejection message reads as a field list.
const EMAIL_FIELD_MAP: Record<keyof SimplifiedEmail, true> = {
  id: true,
  subject: true,
  from: true,
  date: true,
  threadId: true,
  messageId: true,
  references: true,
  to: true,
  cc: true,
  bcc: true,
  replyTo: true,
  inReplyTo: true,
  isReply: true,
  forwardedMessageId: true,
  sourceEmailId: true,
  isRead: true,
  isFlagged: true,
  isDraft: true,
  isAnswered: true,
  isForwarded: true,
  mailboxes: true,
  roles: true,
  unresolvedMailboxIds: true,
  preview: true,
  listUnsubscribe: true,
  hasAttachment: true,
  bodyText: true,
  bodyHtml: true,
  bodyHtmlSize: true,
  bodyTextSize: true,
  quotedBytesStripped: true,
  quotedStripSkipped: true,
  bodyTextUnavailable: true,
  bodyHash: true,
  bodyHashWithheld: true,
  blobId: true,
  size: true,
  keywords: true,
  attachments: true,
};

export const EMAIL_FIELD_NAMES: readonly string[] = Object.keys(EMAIL_FIELD_MAP);

const EMAIL_FIELD_SET = new Set<string>(EMAIL_FIELD_NAMES);

// Longest caller-supplied field name echoed back in a rejection, so a pasted blob
// can't become the error message (mirrors the date-echo limit in coerce.ts).
const FIELD_ECHO_LIMIT = 40;

function echoField(value: string): string {
  return value.length > FIELD_ECHO_LIMIT ? `${value.slice(0, FIELD_ECHO_LIMIT)}...` : value;
}

function validFieldList(): string {
  return `Valid fields: ${EMAIL_FIELD_NAMES.join(', ')}.`;
}

/**
 * Parse and validate the `fields` tool parameter into the set the formatters project
 * with, or undefined when the caller didn't ask for a projection.
 *
 * Every failure throws rather than falling back, because every quiet fallback here
 * returns the FULL response the caller was trying to avoid: `raw` with `fields` (raw
 * quietly winning would return the largest response with no signal, the #11 posture),
 * an unknown name (a typo would come back as an empty object), and an empty array
 * (a dynamically built list that came out empty, not a request for nothing).
 */
export function parseEmailFields(value: unknown, options?: { raw?: boolean }): Set<string> | undefined {
  if (value === undefined || value === null) return undefined;

  if (options?.raw) {
    throw new InvalidInputError(
      'fields cannot be combined with raw:true. raw returns the untransformed JMAP response, whose field names differ from the simplified ones fields selects (receivedAt not date, mailboxIds not mailboxes). Drop raw to project the simplified shape, or drop fields to get the raw response.',
    );
  }

  const names = coerceStringArray(value);
  if (!names) {
    throw new InvalidInputError(
      `fields must be an array of simplified field names (a comma-separated string is also accepted), not ${typeof value}. ${validFieldList()}`,
    );
  }

  if (names.length === 0) {
    throw new InvalidInputError(
      `fields cannot be empty; omit the parameter to get the default fields. ${validFieldList()}`,
    );
  }

  const selected = new Set<string>();
  const unknown: string[] = [];
  for (const name of names) {
    const trimmed = name.trim();
    if (EMAIL_FIELD_SET.has(trimmed)) {
      selected.add(trimmed);
    } else {
      unknown.push(`"${echoField(trimmed)}"`);
    }
  }

  if (unknown.length > 0) {
    throw new InvalidInputError(
      `Unknown field name(s) in fields: ${unknown.join(', ')}. ${validFieldList()}`,
    );
  }

  return selected;
}

/**
 * True when the projection asks for the HTML body, so get_email can turn on includeHtml
 * without verbose; otherwise `fields: ["bodyHtml"]` would return `{}` (#69).
 */
export function wantsHtmlBody(fields: ReadonlySet<string> | undefined): boolean {
  return !!fields?.has('bodyHtml');
}

// The bodyText signals (#73) describe what happened TO the returned bodyText, so they ride
// along with it: a projected bodyText without its strip signal would read as verbatim.
// Attached by key presence, not truthiness, since quotedBytesStripped 0 is an answer.
const BODY_TEXT_SIGNALS = ['quotedBytesStripped', 'quotedStripSkipped', 'bodyTextUnavailable'] as const;

// The draft body hash rides along with EITHER body field: a body without the token
// edit_draft demands is unusable for an edit. Both halves ride, and each rides with the
// other, so `fields: ["bodyHash"]` (which names no body, so no hash can be issued) returns
// the withheld reason naming the fields to add, rather than `{}`.
const BODY_HASH_SIGNALS = ['bodyHash', 'bodyHashWithheld'] as const;

/**
 * Keep only the projected fields of an already-simplified email; no field is invented.
 *
 * `unresolvedMailboxIds` rides along uninvited whenever `mailboxes` or `roles` is
 * projected: it is their degradation half, and dropping it would make a short
 * `mailboxes` array look complete (#53).
 */
export function projectEmail(
  email: SimplifiedEmail,
  fields: ReadonlySet<string> | undefined,
): Partial<SimplifiedEmail> {
  if (!fields) return email;

  const source = email as unknown as Record<string, unknown>;
  const projected: Record<string, unknown> = {};
  for (const key of Object.keys(source)) {
    if (fields.has(key)) projected[key] = source[key];
  }

  if (
    !fields.has('unresolvedMailboxIds') &&
    (fields.has('mailboxes') || fields.has('roles')) &&
    email.unresolvedMailboxIds?.length
  ) {
    projected.unresolvedMailboxIds = email.unresolvedMailboxIds;
  }

  if (fields.has('bodyText')) {
    for (const signal of BODY_TEXT_SIGNALS) {
      if (!fields.has(signal) && Object.prototype.hasOwnProperty.call(source, signal)) {
        projected[signal] = source[signal];
      }
    }
  }

  if (fields.has('bodyText') || fields.has('bodyHtml') || BODY_HASH_SIGNALS.some((s) => fields.has(s))) {
    for (const signal of BODY_HASH_SIGNALS) {
      if (!fields.has(signal) && Object.prototype.hasOwnProperty.call(source, signal)) {
        projected[signal] = source[signal];
      }
    }
  }

  return projected as Partial<SimplifiedEmail>;
}
