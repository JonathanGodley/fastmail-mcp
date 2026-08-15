import { InvalidInputError, coerceStringArray } from './coerce.js';
import type { SimplifiedEmail } from './email-formatter.js';

// Caller-directed output projection for the email read tools (#69, #79).
//
// The read tools return a fixed, fairly wide per-message shape, and response size is
// therefore decided by what happens to be in the mailbox rather than by anything the
// caller controls: a 66-message sweep measured 84KB, of which the five fields the
// caller actually wanted were 18%; an editor-inflated draft's HTML body pushed
// get_email past the same wall. `fields` lets a caller ask for LESS.
//
// Three properties make this safe to bolt onto the existing shape:
//   1. It runs AFTER simplifyEmail, on the simplified object. `raw: true` is a
//      promise of the untransformed JMAP response and is never projected.
//   2. It is subtractive only. It cannot invent a field or change one's meaning, so
//      every other guarantee the simplified shape makes still holds for the fields
//      that survive.
//   3. It is caller-directed. Omitting a field the caller explicitly excluded is not
//      the "silently dropped promised field" failure — the caller asked. The one
//      field that still rides along uninvited is unresolvedMailboxIds; see below.

// Exhaustive runtime vocabulary of the simplified email shape. Typing the map as
// Record<keyof SimplifiedEmail, true> is what keeps it honest: adding a field to
// SimplifiedEmail without adding it here is a compile error, so a new field can
// never become permanently unselectable. Declaration order matches the interface
// (and roughly the emitted order), so the rejection message reads as a field list.
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
 * Every failure is a loud InvalidInputError (mapped to InvalidParams at the index
 * boundary) rather than a quiet fallback, because every quiet fallback here returns
 * the FULL response — which is the exact refusal the parameter exists to avoid:
 *
 *   - `raw: true` together with `fields` is rejected. raw returns untransformed JMAP,
 *     whose field names differ from the simplified ones (`receivedAt` not `date`,
 *     `mailboxIds` not `mailboxes`), so projecting raw would either drop everything or
 *     require a second vocabulary. Letting raw quietly win is worse than rejecting:
 *     the caller who asked for a small response would get the largest one this server
 *     produces, with no signal that its request was ignored (the #11 posture).
 *   - An unknown name is rejected naming the valid set, so a typo can't come back as a
 *     plausible-looking empty object.
 *   - An empty array is rejected. "Give me nothing" is never the intent; it is a caller
 *     that built the list dynamically and got nothing, and treating it as "no
 *     projection" would silently hand back the full shape.
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
 * True when the projection asks for the HTML body, so get_email can turn on
 * simplifyEmail's includeHtml without the caller also having to pass verbose. Without
 * this, `fields: ["bodyHtml"]` would project a field the simplifier never emitted and
 * return `{}` — the trap #69 is trying to escape.
 */
export function wantsHtmlBody(fields: ReadonlySet<string> | undefined): boolean {
  return !!fields?.has('bodyHtml');
}

/**
 * Keep only the projected fields of an already-simplified email. Fields the message
 * doesn't have are simply absent (a projection can legitimately come back empty, e.g.
 * bodyText on an email with no plain-text part); no field is invented.
 *
 * One field rides along uninvited: `unresolvedMailboxIds` is emitted whenever the
 * caller projected `mailboxes` or `roles` and a mailbox id failed to resolve. It is
 * not an independent field — it is the degradation half of those two, and dropping it
 * would reproduce the #53 bug through a new door: a short `mailboxes` array that looks
 * complete but silently isn't, which is precisely the "never silently drop a promised
 * field" rule. A caller that projects neither location field is promised no location,
 * so nothing is owed and nothing rides along.
 *
 * The bodyText signals ride along the same way: `quotedBytesStripped` /
 * `quotedStripSkipped` (#73) and `bodyTextUnavailable` describe what happened TO the
 * bodyText being returned — a projected bodyText with its strip signal dropped would
 * read as a verbatim body when it isn't. They are attached by key presence, not
 * truthiness: quotedBytesStripped 0 is a meaningful answer ("no marker matched, the
 * body is whole") and is deliberately emitted despite the omit-empties norm.
 */
const BODY_TEXT_SIGNALS = ['quotedBytesStripped', 'quotedStripSkipped', 'bodyTextUnavailable'] as const;

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

  return projected as Partial<SimplifiedEmail>;
}
