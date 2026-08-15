import { InvalidInputError } from './coerce.js';

// Size bounds for the caller-supplied text that gets serialized into an iCalendar
// VEVENT by `create_calendar_event` / `update_calendar_event`.
//
// WHY THIS EXISTS: every one of those values is emitted through `foldICalLine`
// (src/caldav-client.ts), which folds a content line to 75 octets per RFC 5545 §3.1 by
// repeatedly re-slicing the REMAINDER of the line — so its cost grows with the square of
// the field length, and each fold allocates a fresh copy of the tail. Measured on this
// code: ~135ms to fold a 200KB value, and the process runs out of memory somewhere near
// 800KB. Nothing upstream bounded these fields, so a single oversized `description` —
// a pasted document, or a model that decided to inline a transcript — was enough to
// stall or kill the server for every other request on the same process. The cap is not
// an iCalendar rule; it is a bound on that quadratic serializer. Do not remove it
// without first making the folding linear.
//
// The bounds are deliberately generous: they exist to stop a denial of service, not to
// referee how long an agenda may be.

// Per-field cap. 64KB of text is far beyond any real event summary, location or agenda
// (a dense page of prose is ~3KB), while keeping the worst-case fold on one field in the
// tens of milliseconds.
export const MAX_ICAL_FIELD_BYTES = 64 * 1024;

// Cap on the number of participants. One participant is one folded ATTENDEE line, so an
// unbounded array is an unbounded number of folds. 500 comfortably covers an all-hands
// or a large distribution list.
export const MAX_ICAL_PARTICIPANTS = 500;

// Cap on the TOTAL of all folded text in one call. Per-field caps alone are defeated by
// many fields that each sit just under the bound — 500 participant names at 60KB apiece
// are individually legal and collectively 30MB. Sized for the largest event that is
// still a real event: a 64KB description (long minutes or a full agenda) plus a title
// and location, plus 500 participants averaging ~100 bytes of name and address each,
// lands around 120KB. 256KB leaves roughly double that as headroom.
export const MAX_ICAL_TOTAL_BYTES = 256 * 1024;

// The calendar-event shape as it arrives from a tool call, before any downstream
// validation has run. Everything is `unknown` because this guard runs FIRST — ahead of
// the type and emptiness checks — so that an oversized value is rejected before it can
// be measured, escaped or folded. Non-string values are skipped here and left to the
// validators that own them (`requireNonEmpty`, `validateAttendeeEmail`).
export interface ICalTextInput {
  title?: unknown;
  description?: unknown;
  location?: unknown;
  participants?: unknown;
}

function byteLength(value: string): number {
  return Buffer.byteLength(value, 'utf8');
}

function describeBytes(limit: number): string {
  return `${limit} bytes (${Math.round(limit / 1024)} KB)`;
}

// Measure one field, add it to the running total, and reject if it alone is over the
// per-field cap. The message names the field, its actual size and the limit, so a single
// retry with a shorter value fixes it.
function measureField(name: string, value: unknown, totals: { bytes: number }): void {
  if (typeof value !== 'string') return;
  const size = byteLength(value);
  totals.bytes += size;
  if (size > MAX_ICAL_FIELD_BYTES) {
    throw new InvalidInputError(
      `${name} is too large: ${size} bytes; the limit is ${describeBytes(MAX_ICAL_FIELD_BYTES)}. ` +
      `Shorten it, or keep the long text somewhere the event can link to.`,
    );
  }
}

/**
 * Reject a calendar-event input whose text fields exceed the serialization bounds above.
 *
 * REJECTS rather than truncates. A silently trimmed description is data loss the caller
 * never learns about — the event would be created, reported as created, and be missing
 * the half of the agenda nobody thought to re-read. An error naming the field and the
 * limit is recoverable in one retry.
 *
 * Throws `InvalidInputError`, which the CallTool boundary maps to `InvalidParams`: this
 * is the caller's input and the caller can fix it. A plain `Error` would surface as
 * `InternalError` — "server-side, a bare retry might work" — which would be false here,
 * and would invite exactly the retry that repeats the expensive serialization.
 */
export function assertICalTextLimits(input: ICalTextInput): void {
  const totals = { bytes: 0 };

  measureField('title', input.title, totals);
  measureField('description', input.description, totals);
  measureField('location', input.location, totals);

  const participants = input.participants;
  if (Array.isArray(participants)) {
    if (participants.length > MAX_ICAL_PARTICIPANTS) {
      throw new InvalidInputError(
        `participants has ${participants.length} entries; the limit is ${MAX_ICAL_PARTICIPANTS}. ` +
        `Split the invitation, or invite a mailing list address instead of its members.`,
      );
    }
    for (let i = 0; i < participants.length; i++) {
      const p = participants[i];
      if (typeof p !== 'object' || p === null) continue;
      const { email, name } = p as { email?: unknown; name?: unknown };
      // Both halves reach the same fold — `name` through the CN parameter and `email`
      // through the mailto: value — so one oversized participant name is the same
      // failure as one oversized description.
      measureField(`participants[${i}].email`, email, totals);
      measureField(`participants[${i}].name`, name, totals);
    }
  }

  if (totals.bytes > MAX_ICAL_TOTAL_BYTES) {
    throw new InvalidInputError(
      `The event's text fields total ${totals.bytes} bytes; the combined limit is ` +
      `${describeBytes(MAX_ICAL_TOTAL_BYTES)}. Shorten the description or reduce the participant list.`,
    );
  }
}
