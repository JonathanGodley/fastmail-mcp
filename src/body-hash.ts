import { createHash } from 'node:crypto';
import type { SimplifiedEmail } from './email-formatter.js';

// ---------------------------------------------------------------------------
// The draft body hash: one lost-update guard, computed the same way on both sides
// ---------------------------------------------------------------------------
//
// `get_email` issues the hash over a draft's stored body; `edit_draft`, which stores the body
// it is handed byte for byte, refuses a body edit whose hash is absent or stale. It proves the
// caller SAW the body it replaces and preserves none of it.
//
// The hash is over EVERY stored body part, deduplicated across the two lists, whatever its
// type: no part type is consulted, so a typeless part or several parts of one type are
// covered rather than withheld.

/**
 * The identity a part is deduped by across the body lists. RFC 8621 §4.1.4 puts one displayed
 * part into BOTH lists, so raw entries would count a plain-text draft's one part twice. Shared
 * with `classifyDraftBodyShape`.
 */
export function draftPartKey(part: any, fallback: number): string {
  if (typeof part?.partId === 'string' && part.partId) return `p:${part.partId}`;
  if (typeof part?.blobId === 'string' && part.blobId) return `b:${part.blobId}`;
  return `i:${fallback}`;
}

// ---------------------------------------------------------------------------
// What a body list carries, as one rule
// ---------------------------------------------------------------------------
//
// The read, the hash over what it showed, and the edit-side guard must agree on which part is
// a list's displayed text, so the rule lives here, in the lowest of their modules.

// Content type without its parameters, lowercased. For CLASSIFYING only — the value stored
// or sent for a part is always the server's own string (RFC 2045 §5.1).
export function classifyPartType(type: unknown): string {
  if (typeof type !== 'string') return '';
  const semicolon = type.indexOf(';');
  return (semicolon === -1 ? type : type.slice(0, semicolon)).trim().toLowerCase();
}

export type DraftTextType = 'text/plain' | 'text/html';

export function isTextBodyType(type: unknown): type is DraftTextType {
  return type === 'text/plain' || type === 'text/html';
}

/**
 * The text body type a part counts as inside one of a draft's two body lists, or undefined
 * when it is not displayed text at all.
 *
 * A part that declares NO content type counts as the list it sits in: RFC 8621 §4.1.4 puts a
 * part in a list to say it is displayed there, and `extractBody` displays it so. `type` is
 * taken as the caller holds it, classified or verbatim.
 *
 * SCOPE: THIS ANSWERS WHAT A READ DISPLAYS, AND ONLY THAT. The edit side does NOT widen to
 * match: `bodyValueForType` in jmap-client.ts matches the declared type exactly (#179).
 */
export function draftTextBodyType(type: unknown, listType: DraftTextType): DraftTextType | undefined {
  if (type === undefined || type === null || type === '') return listType;
  return isTextBodyType(type) ? type : undefined;
}

/**
 * The text body type a draft's body alternates between: two DISTINCT parts counting as one
 * text type, the Apple Mail text-image-text layout whose ordering a flat rebuild cannot
 * express (issue #85). Undefined when the body has no such pair.
 *
 * ONE EXPRESSION, TWO CONSUMERS: `updateDraft` refuses every edit of this shape and the read
 * withholds its `bodyHash` (#180), and they agree because both ask this function.
 */
export function draftInterleavedTextType(email: any): string | undefined {
  const seen = new Set<string>();
  const counts = new Map<string, number>();
  let index = 0;

  for (const list of [
    { parts: email?.textBody, listType: 'text/plain' as const },
    { parts: email?.htmlBody, listType: 'text/html' as const },
  ]) {
    if (!Array.isArray(list.parts)) continue;
    for (const part of list.parts) {
      if (!part) continue;
      // The fallback counter advances for every part examined, duplicates included, so it
      // matches collectDraftBodyParts' numbering on the same message.
      const key = draftPartKey(part, index++);
      if (seen.has(key)) continue;
      seen.add(key);

      const type = classifyPartType(part.type);
      // A typeless part is not counted, the one place this module does not fall back to the
      // list. The shape cannot arise (RFC 8621 §4.1.4 makes `type` mandatory, Cyrus fills a
      // missing one in, and every `bodyProperties` request in jmap-client.ts asks for it), and
      // widening would pair a lone typeless part with the typed one beside it and refuse an
      // edit for no visible reason (compare #179).
      if (!type) continue;
      const countsAs = draftTextBodyType(type, list.listType);
      if (countsAs === undefined) continue;

      const count = (counts.get(countsAs) ?? 0) + 1;
      counts.set(countsAs, count);
      if (count > 1) return countsAs;
    }
  }

  return undefined;
}

/** One deduplicated body part, with what the read returned for it. */
export interface CollectedBodyPart {
  key: string;
  /** Declared content type, verbatim. */
  type?: string;
  /**
   * Undefined both for a part with no body value (an embedded image routed into a body list)
   * and for one whose value this read did not fetch; `showsIn*` separates the two.
   */
  value?: string;
  degraded: boolean;
  /**
   * Whether `simplifyEmail`'s `bodyText` / `bodyHtml` would carry this part, by
   * `draftTextBodyType`. `extractBody` mirrors that test with its own copy.
   */
  showsInText: boolean;
  showsInHtml: boolean;
}

/**
 * The deduplicated body part set of one JMAP email, in stored order (the textBody list,
 * then anything the htmlBody list adds).
 *
 * Both callers fetch with `fetchTextBodyValues` and `fetchHTMLBodyValues`, which is what lets
 * a hash issued by a read match one recomputed at edit time.
 */
export function collectDraftBodyParts(email: any): CollectedBodyPart[] {
  const bodyValues: Record<string, any> = email?.bodyValues || {};
  const parts = new Map<string, CollectedBodyPart>();
  let index = 0;

  for (const list of [
    { parts: email?.textBody, inText: true },
    { parts: email?.htmlBody, inText: false },
  ]) {
    if (!Array.isArray(list.parts)) continue;
    for (const part of list.parts) {
      if (!part) continue;
      // Kept in step BY HAND with `draftInterleavedTextType`'s walk: same dedupe, and the
      // counter advances for every part, duplicates included.
      const key = draftPartKey(part, index++);
      let entry = parts.get(key);
      if (!entry) {
        const type = typeof part.type === 'string' ? part.type : undefined;
        const bv = typeof part.partId === 'string' ? bodyValues[part.partId] : undefined;
        entry = {
          key,
          ...(type !== undefined && { type }),
          ...(typeof bv?.value === 'string' && { value: bv.value }),
          degraded: !!bv?.isTruncated || !!bv?.isEncodingProblem,
          showsInText: false,
          showsInHtml: false,
        };
        parts.set(key, entry);
      }
      const listType = list.inText ? 'text/plain' : 'text/html';
      const carriable = draftTextBodyType(entry.type, listType) === listType;
      if (list.inText) entry.showsInText ||= carriable;
      else entry.showsInHtml ||= carriable;
    }
  }

  return [...parts.values()];
}

/**
 * The hash of a draft's stored body, as an opaque token. `bh1-` is a version marker: changing
 * what is hashed means bumping it, so an old token is rejected as stale rather than colliding.
 *
 * The byte-length prefix stops one concatenation of parts spelling another, and the `-`
 * sentinel keeps "no body value" distinct from "empty body value".
 */
export function bodyHash(parts: readonly CollectedBodyPart[]): string {
  const canonical = parts
    .map((p) => (p.value === undefined ? '-' : `${Buffer.byteLength(p.value, 'utf8')}:${p.value}`))
    .join('\n');
  return `bh1-${createHash('sha256').update(canonical, 'utf8').digest('hex').slice(0, 32)}`;
}

export function isDraftEmail(email: any): boolean {
  return !!email?.keywords?.$draft;
}

/** What the response this hash would ride on actually carries, after `fields` projection. */
export interface DraftBodyHashRead {
  bodyText: boolean;
  bodyHtml: boolean;
  stripQuoted: boolean;
}

export type DraftBodyHashOutcome =
  | { bodyHash: string }
  | { bodyHashWithheld: string };

/**
 * Whether this read may issue a `bodyHash`, and the reason when it may not.
 *
 * NEVER SILENT: a draft read carries the hash or a reason naming the read that would issue
 * one. The response must have shown every stored byte the hash covers. Undefined for a
 * non-draft, where nothing was promised.
 */
export function resolveDraftBodyHash(email: any, read: DraftBodyHashRead): DraftBodyHashOutcome | undefined {
  if (!isDraftEmail(email)) return undefined;

  const parts = collectDraftBodyParts(email);
  // An empty-string value is not content: extractBody skips it, so no read displays it and
  // no read has to. It still hashes distinctly from a valueless part (see bodyHash).
  const withContent = parts.filter((p) => p.value !== undefined && p.value !== '');

  if (withContent.some((p) => p.degraded)) {
    return {
      bodyHashWithheld:
        'the server flagged part of this draft\'s stored body as truncated or as having ' +
        'encoding problems, so this read did not return it whole and no read can prove you ' +
        'saw it. Recreate the draft rather than editing its body.',
    };
  }

  const unreadable = withContent.filter((p) => !p.showsInText && !p.showsInHtml);
  if (unreadable.length > 0) {
    return {
      bodyHashWithheld:
        'this draft carries a body part no read returns (a part whose declared type does not ' +
        'match the body list it sits in), so no read can prove you saw the whole body. ' +
        'Recreate the draft rather than editing its body.',
    };
  }

  // Reported ahead of stripQuoted, like the degraded case: no second read would issue a hash
  // either, so naming one would send the caller nowhere.
  const interleaved = draftInterleavedTextType(email);
  if (interleaved) {
    return {
      bodyHashWithheld:
        `this draft's body interleaves multiple ${interleaved} parts, a layout editing ` +
        'cannot preserve, so every edit of this draft is refused and a bodyHash could never ' +
        'be spent. Recreate the draft rather than editing its body (see issue #85).',
    };
  }

  if (read.stripQuoted) {
    return {
      bodyHashWithheld:
        'this read stripped quoted history out of bodyText, so it does not show the body as ' +
        'stored. Read the draft again without stripQuoted to get a bodyHash.',
    };
  }

  // A part both lists carry is satisfied by either field.
  const needsText = withContent.some((p) => p.showsInText && !p.showsInHtml);
  const needsHtml = withContent.some((p) => p.showsInHtml && !p.showsInText);
  const eitherOnly = withContent.filter((p) => p.showsInText && p.showsInHtml);

  const missing =
    (needsText && !read.bodyText) ||
    (needsHtml && !read.bodyHtml) ||
    (eitherOnly.length > 0 && !read.bodyText && !read.bodyHtml);

  if (missing) {
    // The remedy names the WHOLE read, not just the field this one lacked: the hash needs
    // every part shown at once, so telling a caller who read only bodyHtml to "read
    // bodyText" would send them to a second read that issues no hash either.
    const wanted = ['bodyText', 'bodyHtml'].filter((f) =>
      f === 'bodyText' ? needsText || eitherOnly.length > 0 : needsHtml,
    );
    const list = [...wanted, 'bodyHash'].map((f) => `"${f}"`).join(', ');
    const verbose = wanted.includes('bodyHtml') ? ' (or verbose:true)' : '';
    return {
      bodyHashWithheld:
        `this read did not return the draft's stored body whole, so it cannot prove you saw ` +
        `it. Read the draft with fields: [${list}]${verbose} to get a bodyHash.`,
    };
  }

  return { bodyHash: bodyHash(parts) };
}

/** What the get_email call being answered asked for. */
export interface DraftBodyHashReadOptions {
  raw: boolean;
  /** The parsed `fields` projection, or undefined for an unprojected read. */
  fields?: ReadonlySet<string>;
  stripQuoted: boolean;
}

/**
 * Attach `bodyHash` / `bodyHashWithheld` to a simplified `get_email` result, in place.
 *
 * `raw` attaches nothing: a field of this server's invention would stop it being raw. A body
 * field the simplifier never produced, or the projection dropped, counts as NOT returned.
 */
export function attachDraftBodyHash(
  email: any,
  simplified: SimplifiedEmail,
  options: DraftBodyHashReadOptions,
): void {
  if (options.raw) return;
  const { fields } = options;
  const outcome = resolveDraftBodyHash(email, {
    bodyText: simplified.bodyText !== undefined && (!fields || fields.has('bodyText')),
    bodyHtml: simplified.bodyHtml !== undefined && (!fields || fields.has('bodyHtml')),
    stripQuoted: options.stripQuoted,
  });
  if (outcome) Object.assign(simplified, outcome);
}
