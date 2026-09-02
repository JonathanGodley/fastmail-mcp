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

// ---------------------------------------------------------------------------
// What a body list carries, as one rule
// ---------------------------------------------------------------------------
//
// Three places have to agree about which part of a body list is that list's displayed
// text: the read that shows the caller a body, the hash over what it showed, and the
// edit-side guard that decides whether a flat recreate can express the draft. They live in
// different modules, so the rule itself lives here — the lowest of the three — and the
// others call down to it.

// Content type without its parameters, lowercased. For CLASSIFYING only — the value stored
// or sent for a part is always the server's own string (RFC 2045 §5.1).
export function classifyPartType(type: unknown): string {
  if (typeof type !== 'string') return '';
  const semicolon = type.indexOf(';');
  return (semicolon === -1 ? type : type.slice(0, semicolon)).trim().toLowerCase();
}

/** The two content types a draft's body lists carry as displayed text. */
export type DraftTextType = 'text/plain' | 'text/html';

/** Whether a content type is one a body list carries as displayed text. */
export function isTextBodyType(type: unknown): type is DraftTextType {
  return type === 'text/plain' || type === 'text/html';
}

/**
 * The text body type a part counts as inside one of a draft's two body lists, or undefined
 * when it is not displayed text at all.
 *
 * A part that declares NO content type counts as the list it sits in. That is not leniency:
 * RFC 8621 §4.1.4 puts a part into `textBody` or `htmlBody` precisely to say a client should
 * display it there, so LIST MEMBERSHIP IS THE AUTHORITY WHEN THE TYPE IS ABSENT. It is also
 * how the reader has always behaved — `extractBody` displays a typeless part in whichever
 * list carries it — so anything that disagrees is reasoning about a part the caller can see.
 *
 * `type` is taken as the caller holds it: a caller that classifies (strips parameters,
 * lowercases) passes the classified string, one that compares the server's string verbatim
 * passes that. The absent-type rule is the same either way, which is the part that has to
 * be shared.
 *
 * SCOPE: THIS ANSWERS WHAT A READ DISPLAYS, AND ONLY THAT. It is the rule the body reader
 * and the hash over what it read share, so the hash covers exactly the bytes the caller was
 * shown. The edit side does NOT widen to match: the selector that picks a part to rebuild a
 * draft from matches the declared type exactly, because a typeless part matching whichever
 * list asks for it makes a one-part draft's single part answer BOTH formats, and the
 * recreate then writes the same bytes into a `text/plain` and a `text/html` part — plain
 * text rendered as markup, silently, on the metadata-only path that promises to leave the
 * body alone (#179). Its own comment carries the rest of the reasoning.
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
 * A part that declares no content type is not counted at all here — see the skip in the
 * loop for why that is right rather than merely conservative.
 *
 * ONE EXPRESSION, TWO CONSUMERS, DELIBERATELY. `updateDraft` refuses every edit of this
 * shape, metadata-only included, and a read that issued a `bodyHash` for it would hand out
 * a lost-update guard that can never be spent (#180). The read withholds and the write
 * refuses for exactly the same drafts because they ask this one function, not because two
 * conditions are kept in step by hand.
 *
 * Deduping FIRST is load-bearing, not an optimization: a single-format draft lists its one
 * text part under both `textBody` and `htmlBody`, so a raw count would see two text/plain
 * parts on an ordinary plain-text draft and call it interleaved.
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
      // A part that declares no type is not counted as either format, deliberately, and
      // this is the one place in this module that does not fall back to the list.
      //
      // It costs nothing, because the shape cannot arise. RFC 8621 §4.1.4 makes `type`
      // mandatory on a body part, and Cyrus — the server Fastmail runs — enforces that on
      // the way out: `Email/get` fills a missing Content-Type in with `text/plain` (or
      // `multipart/related`, or `message/rfc822`), lowercased and stripped of parameters,
      // before it ever reaches a client. This module issues no `Email/get` of its own — it
      // is handed parts that `jmap-client.ts` fetched — and every `bodyProperties` request
      // there asks for `type` (it is in `EMAIL_BODY_PROPERTIES`, and in the one hand-written
      // list beside it), so no read feeding this can drop it either, and a message appended over IMAP
      // with no Content-Type gets RFC 2045's `text/plain` default, which is the same value.
      //
      // And it costs something to widen. Counting a typeless part as its list would pair a
      // lone one with the typed part beside it and refuse an edit no reader could see a
      // reason for; the matching widening on the edit-side selector put unescaped plain
      // text into a text/html part on the metadata-only path (#179). A refusal and a
      // corruption, both for an input the server cannot emit.
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
   * test is `draftTextBodyType` above, over the server's string verbatim, which mirrors
   * `extractBody`'s exactly (a part with NO declared type is carried by whichever list it
   * sits in). `extractBody` keeps its own copy rather than calling down, because the two
   * answer different questions — that one builds a string, this one decides whether the
   * caller has seen the bytes it is about to hand back.
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
      // TWO WALKS, ON PURPOSE, and this is the whole disposition of that — there is no
      // issue tracking it. This loop and `draftInterleavedTextType`'s above cross the same
      // two lists with the same `draftPartKey` dedupe and the same fallback-index
      // convention, and those two things are still coordinated BY HAND: the counter
      // advances for every part examined, duplicates included, so the numbering matches.
      // What is NO LONGER coordinated by hand is the rule that decides what a body list
      // carries — `draftTextBodyType` is one expression and both walks ask it.
      // They are deliberately NOT merged into one traversal. They produce different things
      // (this one a part map with values, degradation and per-list visibility; that one a
      // single verdict), so sharing the walk would mean a shared iterator plus two
      // consumers — machinery, not a simplification — and it would reshape the traversal
      // the hash's own coverage rests on in order to fix something that is not broken here.
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

  // The hash is a token to SPEND on an edit, so a draft no edit of which can be made gets
  // none. `updateDraft` refuses every edit of this shape, a metadata-only one included, so
  // a hash issued here could never be spent and would send the caller off to compose an
  // edit against a draft that will refuse it (#180). One expression decides both — the read
  // withholds for exactly the drafts the write refuses because they ask the same function.
  // Reported ahead of stripQuoted for the same reason the degraded case is: no second read
  // would issue a hash either, so naming one would send the caller nowhere.
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
