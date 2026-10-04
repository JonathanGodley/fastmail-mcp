import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { InvalidInputError } from './coerce.js';

/** One exact find/replace op of edit_draft's `bodyEdits`. */
export interface BodyEdit {
  find: string;
  replace: string;
}

export type BodyEditPart = 'htmlBody' | 'textBody';

/** Where one op's `find` sits in the stored part, in UTF-16 code units. */
export interface LocatedBodyEdit {
  offset: number;
  matchedSize: number;
}

/** What a `bodyEdits` call did, op by op in the caller's order; offsets are into the pre-edit part. */
export interface BodyEditsReceipt {
  part: BodyEditPart;
  ops: { offset: number; matchedSize: number; replacementSize: number }[];
}

export const REJECT_BODY_EDITS_WITH_BODY =
  'bodyEdits cannot be combined with htmlBody, textBody, or a clearFields entry naming either: ' +
  'bodyEdits changes part of the stored body, and those replace or remove the whole of it. Use one or the other.';

/** In place of the html-alone note, whose remedy (send textBody alongside) bodyEdits refuses. */
export const NOTE_BODY_EDITS_DISCARDED_TEXT_PART =
  'This edit changed htmlBody through bodyEdits, so the draft\'s stored plain-text part was replaced by a ' +
  'fallback derived from the edited html. If that part was hand-written, re-send both parts as htmlBody ' +
  'and textBody instead of using bodyEdits.';

export const REJECT_BODY_EDITS_WITHOUT_SIGNATURE =
  'expandSignature: true was passed but no bodyEdits replace carries {{signature}}, so there is nothing ' +
  'to expand; a {{signature}} already in the stored body is not expanded. Put {{signature}} in the replace ' +
  'where the sign-off goes, or drop the flag and the replacements are stored as written.';

/** Under expandSignature: tokens formed across a replace's edge, which no expansion pass sees. */
export function noteBodyEditsSplitSignature(part: BodyEditPart, count: number): string {
  const one = count === 1;
  return (
    `${count} {{signature}} token${one ? '' : 's'} in ${part} ${one ? 'was' : 'were'} formed where a bodyEdits ` +
    `replace meets the text beside it, so ${one ? 'it is' : 'they are'} stored as literal text: only a ` +
    'token wholly inside a replace expands. Put the whole {{signature}} in one replace.'
  );
}

export const REJECT_BODY_EDITS_NO_BODY =
  'bodyEdits would leave the draft with no body; keep some text in the edited part, or replace the body ' +
  'whole with htmlBody or textBody.';

const BODY_EDIT_KEYS = new Set(['find', 'replace']);
const BODY_EDITS_SHAPE = 'bodyEdits must be an array of {find, replace} objects.';

/**
 * Lenient read of `bodyEdits`, on coerceAttachments' rules: the whole value or any element
 * may arrive JSON-encoded, and a blank string is an omitted value. `find` and `replace` are
 * never trimmed, since whitespace is part of an exact match.
 */
export function coerceBodyEdits(value: unknown): BodyEdit[] | undefined {
  if (value === undefined || value === null) return undefined;

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    try {
      arr = JSON.parse(trimmed);
    } catch {
      throw new McpError(ErrorCode.InvalidParams, BODY_EDITS_SHAPE);
    }
  }
  if (!Array.isArray(arr)) throw new McpError(ErrorCode.InvalidParams, BODY_EDITS_SHAPE);
  if (arr.length === 0) {
    throw new McpError(ErrorCode.InvalidParams, 'bodyEdits cannot be empty; omit it to leave the body unchanged.');
  }

  return arr.map((raw, i) => {
    let item: unknown = raw;
    if (typeof item === 'string') {
      const t = item.trim();
      if (!(t.startsWith('{') && t.endsWith('}'))) {
        throw new McpError(ErrorCode.InvalidParams, `bodyEdits[${i}] must be a {find, replace} object, not a bare string.`);
      }
      try {
        item = JSON.parse(t);
      } catch {
        throw new McpError(ErrorCode.InvalidParams, `bodyEdits[${i}] is a string that isn't valid JSON; pass a {find, replace} object.`);
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new McpError(ErrorCode.InvalidParams, `bodyEdits[${i}] must be a {find, replace} object.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter((k) => !BODY_EDIT_KEYS.has(k));
    if (unknownKeys.length > 0) {
      throw new McpError(ErrorCode.InvalidParams, `bodyEdits[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: find, replace`);
    }
    for (const key of ['find', 'replace'] as const) {
      if (typeof obj[key] !== 'string') {
        throw new McpError(ErrorCode.InvalidParams, `bodyEdits[${i}].${key} must be a string.`);
      }
    }
    return { find: obj.find as string, replace: obj.replace as string };
  });
}

/**
 * Locate every op's `find` in `stored`, matched raw and exactly. Each must occur exactly
 * once, counting overlapping occurrences, and no two ops' matches may intersect (touching is
 * fine). Refusals name the op by index and never echo its text.
 */
export function locateBodyEdits(stored: string, ops: readonly BodyEdit[], part: BodyEditPart): LocatedBodyEdit[] {
  ops.forEach((op, i) => {
    if (op.find === '') {
      throw new InvalidInputError(`bodyEdits[${i}].find is empty; give the exact text to replace.`);
    }
  });

  const located = ops.map((op, i) => {
    const first = stored.indexOf(op.find);
    if (first === -1) {
      throw new InvalidInputError(
        `bodyEdits[${i}].find not found; re-read the draft and copy the text exactly as its ${part} stores it (no entity decoding or whitespace normalisation is applied).`,
      );
    }
    let count = 1;
    for (let at = stored.indexOf(op.find, first + 1); at !== -1; at = stored.indexOf(op.find, at + 1)) count++;
    if (count > 1) {
      throw new InvalidInputError(
        `bodyEdits[${i}].find occurs ${count} times in ${part}: ambiguous; include more context so it matches exactly once.`,
      );
    }
    return { offset: first, matchedSize: op.find.length };
  });

  for (let i = 0; i < located.length; i++) {
    for (let j = i + 1; j < located.length; j++) {
      const a = located[i];
      const b = located[j];
      if (a.offset < b.offset + b.matchedSize && b.offset < a.offset + a.matchedSize) {
        throw new InvalidInputError(
          `bodyEdits[${i}] and bodyEdits[${j}] match overlapping text in ${part}; combine them into one edit.`,
        );
      }
    }
  }
  return located;
}

/**
 * Apply located ops to the string they were located in, all at once: every offset is into
 * `stored`, so no op sees another's replacement. `replacements[i]` replaces `located[i]`.
 */
export function spliceBodyEdits(stored: string, located: readonly LocatedBodyEdit[], replacements: readonly string[]): string {
  const order = located.map((_, i) => i).sort((x, y) => located[x].offset - located[y].offset);
  let out = '';
  let cursor = 0;
  for (const i of order) {
    out += stored.slice(cursor, located[i].offset) + replacements[i];
    cursor = located[i].offset + located[i].matchedSize;
  }
  return out + stored.slice(cursor);
}

/** The stretches of `stored` that no located op matched, in order. */
export function unmatchedSegments(stored: string, located: readonly LocatedBodyEdit[]): string[] {
  const ordered = [...located].sort((x, y) => x.offset - y.offset);
  const segments: string[] = [];
  let cursor = 0;
  for (const op of ordered) {
    segments.push(stored.slice(cursor, op.offset));
    cursor = op.offset + op.matchedSize;
  }
  segments.push(stored.slice(cursor));
  return segments;
}
