import { createHash } from 'node:crypto';
import type { SimplifiedEmail } from './email-formatter.js';

// ---------------------------------------------------------------------------
// The draft body hash: one lost-update guard, computed the same way on both sides
// ---------------------------------------------------------------------------
//
// `edit_draft` stores the body it is handed byte for byte. Nothing about the draft's
// existing body is preserved, inferred or rebuilt, so an edit written against a stale
// read silently overwrites whatever changed in between. The hash is what makes that
// loud: `get_email` issues one over the draft's stored body, `edit_draft` recomputes it
// before any body write and refuses a body edit whose hash is absent or no longer
// current.
//
// READ IT AS A LOST-UPDATE GUARD AND NOTHING ELSE. It proves the caller SAW the body it
// is replacing; it does not preserve any of it, and it is not a quote-preservation
// mechanism (the guard that used to be one is gone). What the caller hands back is its
// edit, whatever that drops.
//
// The hash is over EVERY body part the draft stores, in stored order, whatever the part's
// type — the `draftPartKey`-deduplicated union of the two body lists, each part
// contributing its stored `bodyValue`. No part type is consulted, deliberately: a typeless
// part, and several parts of one type, are covered by the hash rather than withheld from
// it. `bodyValueForType` is the merge's and `sendDraft`'s selector and plays no part here.

/**
 * The identity a part is deduped by across the body lists.
 *
 * RFC 8621 §4.1.4 puts one displayed part into BOTH body lists, and a single-format draft
 * aliases its one text part into both, so counting raw array entries would hash every part
 * of an ordinary plain-text draft twice. This is the same rule `classifyDraftBodyShape`
 * dedupes by — one rule, one place, so a part counted once there is counted once here.
 */
export function draftPartKey(part: any, fallback: number): string {
  if (typeof part?.partId === 'string' && part.partId) return `p:${part.partId}`;
  if (typeof part?.blobId === 'string' && part.blobId) return `b:${part.blobId}`;
  return `i:${fallback}`;
}

/** One deduplicated body part, with what the read returned for it. */
export interface CollectedBodyPart {
  key: string;
  /** The part's declared content type, verbatim, or undefined when it declares none. */
  type?: string;
  /**
   * The part's stored value, when the read fetched one. Undefined for a part that carries
   * no body value at all — an embedded image the server routed into a body list — and for
   * a text part whose value this read did not fetch, which is why the two are separated by
   * `showsIn*` below rather than by this field alone.
   */
  value?: string;
  /** The server flagged the fetched value as truncated or as having encoding problems. */
  degraded: boolean;
  /**
   * Whether `simplifyEmail`'s `bodyText` / `bodyHtml` would carry this part's value. The
   * test mirrors `extractBody`'s exactly (a part with NO declared type is carried by
   * whichever list it sits in); it is duplicated rather than shared because the two answer
   * different questions — that one builds a string, this one decides whether the caller
   * has seen the bytes it is about to hand back.
   */
  showsInText: boolean;
  showsInHtml: boolean;
}

/**
 * The deduplicated body part set of one JMAP email, in stored order (the textBody list,
 * then anything the htmlBody list adds).
 *
 * Both callers reach it from an `Email/get` made with `fetchTextBodyValues` and
 * `fetchHTMLBodyValues`, so the values are the whole stored ones on both sides — which is
 * what lets a hash issued by a read be compared against one recomputed at edit time.
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
      // The fallback counter advances for every part examined, duplicates included, so it
      // matches classifyDraftBodyShape's numbering on the same message.
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
      const carriable = !entry.type || entry.type === (list.inText ? 'text/plain' : 'text/html');
      if (list.inText) entry.showsInText ||= carriable;
      else entry.showsInHtml ||= carriable;
    }
  }

  return [...parts.values()];
}

/**
 * The hash of a draft's stored body, as an opaque token.
 *
 * Opaque on purpose: it is a token to hand back, never something a caller reconstructs, so
 * nothing here is part of the contract except that the same stored body always produces
 * the same string. The `bh1-` prefix is a version marker — changing what is hashed means
 * bumping it, so an old token is rejected as stale rather than silently colliding.
 *
 * Each part contributes a byte-length prefix and its value, so no concatenation of parts
 * can spell another; a part with no stored value contributes a sentinel that no value can
 * spell, so "a part with no body value" and "a part whose body value is empty" are
 * distinct bodies.
 */
export function bodyHash(parts: readonly CollectedBodyPart[]): string {
  const canonical = parts
    .map((p) => (p.value === undefined ? '-' : `${Buffer.byteLength(p.value, 'utf8')}:${p.value}`))
    .join('\n');
  return `bh1-${createHash('sha256').update(canonical, 'utf8').digest('hex').slice(0, 32)}`;
}

/** True when a JMAP email carries the `$draft` keyword. */
export function isDraftEmail(email: any): boolean {
  return !!email?.keywords?.$draft;
}

/** What the response this hash would ride on actually carries. */
export interface DraftBodyHashRead {
  /** The response emits the draft's plain-text body (after `fields` projection). */
  bodyText: boolean;
  /** The response emits the draft's html body (after `fields` projection). */
  bodyHtml: boolean;
  /** The caller asked for quoted history to be stripped out of `bodyText`. */
  stripQuoted: boolean;
}

export type DraftBodyHashOutcome =
  | { bodyHash: string }
  | { bodyHashWithheld: string };

/**
 * Whether this read may issue a `bodyHash`, and the reason when it may not.
 *
 * NEVER SILENT: a draft read either carries the hash or carries a reason naming the read
 * that would issue one. The rule is that the response has to have shown the caller every
 * stored byte the hash covers — a hash issued beside a body the caller cannot see would
 * certify a read that never happened, which is the whole thing this guard is for.
 *
 * Returns undefined for a message that is not a draft: the field is dead weight on every
 * other read, and there is no degradation to report because nothing was promised.
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

  if (read.stripQuoted) {
    return {
      bodyHashWithheld:
        'this read stripped quoted history out of bodyText, so it does not show the body as ' +
        'stored. Read the draft again without stripQuoted to get a bodyHash.',
    };
  }

  // Which body fields this draft's stored bytes need in order to be shown whole. A part is
  // shown by whichever field carries it; a part both lists carry is satisfied by either.
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
  /** `raw: true` — the response is unmodified JMAP. */
  raw: boolean;
  /** The parsed `fields` projection, or undefined for an unprojected read. */
  fields?: ReadonlySet<string>;
  /** `stripQuoted: true` — bodyText was shortened. */
  stripQuoted: boolean;
}

/**
 * Attach `bodyHash` / `bodyHashWithheld` to a simplified `get_email` result, in place.
 *
 * This lives here rather than in the CallTool switch because the DECISION is the whole
 * feature: which fields this particular response ends up carrying is what says whether a
 * hash would be honest, and none of that is visible to `resolveDraftBodyHash` on its own.
 *
 *  - `raw` attaches nothing at all. Raw is unmodified JMAP; a field of this server's own
 *    invention in it would stop it being raw, and the caller asking for raw has the stored
 *    `bodyValues` in front of it anyway.
 *  - A body field the simplifier never produced, or one the projection dropped, counts as
 *    NOT returned — a hash beside a body the caller cannot see certifies a read that did
 *    not happen. `fields === undefined` is the unprojected read, which emits whatever the
 *    simplifier produced.
 *  - Everything else — draft or not, whole or degraded — is `resolveDraftBodyHash`'s call.
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
