// Pure helpers for embedded (cid:) image support (#13). No I/O, no JMAP client, no
// server state — everything here is a total function over values, so the read paths
// that consume it stay unit-testable without credentials or a network.

/**
 * A part of an email, paired with where the server routed it.
 *
 * `part` is the JMAP EmailBodyPart VERBATIM — never copied and never mutated, so a
 * caller that hands these straight back (get_email_attachments) emits real JMAP
 * objects, and a caller that derives simplified fields from them (simplifyEmail)
 * cannot leak a derived value into a `raw: true` response.
 *
 * `inBodyList` records the routing itself: RFC 8621 §4.1.4 puts an image that the
 * body displays into `textBody`/`htmlBody` rather than `attachments` for some MIME
 * shapes, and into `attachments` for others. That routing is the only structural
 * signal that a part is body content rather than a separate file.
 */
export interface UnionPart {
  part: any;
  inBodyList: boolean;
}

interface UnionSourceEmail {
  attachments?: any[] | null;
  textBody?: any[] | null;
  htmlBody?: any[] | null;
}

// The two media types a JMAP body list uses to carry the RENDERED body itself. RFC 8621
// §4.1.4 also routes image/*, audio/* and video/* parts into those lists when the body
// displays them, and that is exactly the set the union has to pick up. Excluding the two
// body types (rather than allowlisting the three media ones) is deliberate: it is a
// superset, so a shape the RFC predicate does not enumerate is surfaced rather than
// silently dropped from a listing this server promises is complete.
const BODY_TEXT_TYPES = new Set(['text/plain', 'text/html']);

// Content types are compared case-insensitively and without their parameters, per
// RFC 2045 §5.1. This normalization is for CLASSIFYING only: the value emitted or
// stored for a part is always the server's own string, never this one.
function classifyType(type: unknown): string {
  if (typeof type !== 'string') return '';
  const semicolon = type.indexOf(';');
  return (semicolon === -1 ? type : type.slice(0, semicolon)).trim().toLowerCase();
}

// The identity used to dedupe a part across the three lists. `partId` is the stable
// per-message identity; `blobId` is the fallback for a part the server reports without
// one. A part with neither cannot be matched to anything, so it is kept as its own
// entry rather than silently folded into another.
function partKey(part: any): string | null {
  if (typeof part?.partId === 'string' && part.partId) return `p:${part.partId}`;
  if (typeof part?.blobId === 'string' && part.blobId) return `b:${part.blobId}`;
  return null;
}

/**
 * The part set an attachment-aware read should work from: the JMAP `attachments`
 * array UNION the media parts the server routed into `textBody`/`htmlBody`.
 *
 * Why a union at all: the same embedded image lands in `attachments` for one MIME
 * shape and in the body lists for another (RFC 8621 §4.1.4), so `attachments` alone
 * is not a complete part listing — an image-only message can list nothing at all.
 *
 * GATED on `attachments` being present. The compact list/search property set fetches
 * `textBody` (for the body-size hint) but not `attachments`, so a compact result has
 * no attachment basis to union with; returning `[]` there keeps those responses
 * byte-identical rather than emitting half a listing from the one list they happen
 * to fetch. The gate lives here, not at the call sites, so it cannot be forgotten.
 *
 * Order is canonical and stable: `attachments` in server order first, then the
 * additions from `textBody`, then from `htmlBody`. Callers index into this order
 * (download_attachment's entry-number form), so it must not depend on iteration
 * accidents.
 */
export function buildUnionParts(email: UnionSourceEmail | null | undefined): UnionPart[] {
  const attachments = email?.attachments;
  if (!Array.isArray(attachments)) return [];

  const bodyLists = [email?.textBody, email?.htmlBody].filter(Array.isArray) as any[][];

  // Body-list membership is resolved BEFORE the walk, so a part that appears in both
  // `attachments` and a body list reports the routing either way — otherwise the
  // signal would depend on which list happened to win the dedup.
  const bodyKeys = new Set<string>();
  for (const list of bodyLists) {
    for (const part of list) {
      const key = partKey(part);
      if (key !== null) bodyKeys.add(key);
    }
  }

  const union: UnionPart[] = [];
  const seen = new Set<string>();
  // Identity backstop for a part the server reported with neither partId nor blobId.
  // Such a part is kept as its own entry (it cannot be matched to anything), but RFC
  // 8621 §4.1.4 puts one displayed part into BOTH body lists, so without this the same
  // object would be listed twice — inflating the listing and shifting entry numbers.
  const seenObjects = new Set<any>();

  const add = (part: any, inBodyList: boolean): void => {
    const key = partKey(part);
    if (key !== null) {
      if (seen.has(key)) return;
      seen.add(key);
    } else {
      if (seenObjects.has(part)) return;
      seenObjects.add(part);
    }
    union.push({ part, inBodyList: inBodyList || (key !== null && bodyKeys.has(key)) });
  };

  for (const part of attachments) {
    if (part) add(part, false);
  }

  for (const list of bodyLists) {
    for (const part of list) {
      if (!part) continue;
      // A part with no type at all is treated as body text, matching how the body
      // extractor reads such a part; only a positively-classified non-body type joins.
      const type = classifyType(part.type);
      if (!type || BODY_TEXT_TYPES.has(type)) continue;
      add(part, true);
    }
  }

  return union;
}

/**
 * Percent-decode a `cid:` URL's value once, per RFC 2392 (the value in a `cid:` URL
 * is percent-encoded; the Content-ID it names is not).
 *
 * SINGLE decode, deliberately: decoding repeatedly would let `%2525` reach a
 * comparison as `%`, so a reference could be spelled several ways and resolve to the
 * same part. A malformed escape makes `decodeURIComponent` throw URIError; that is
 * not a caller error worth surfacing, so the input is handed back verbatim and the
 * literal comparison decides.
 */
export function decodeCidSrc(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/**
 * The comparison key for a cid REFERENCE — an `<img src>` value, or any other place
 * a `cid:` URL appears: strip one leading `cid:` scheme, then decode once.
 *
 * Reference side ONLY. A part's own `cid` is a Content-ID, not a URL, and is compared
 * LITERALLY against this key — never decoded. Decoding both sides would make a part
 * whose Content-ID genuinely contains `%78` reachable by a reference spelled `x`.
 *
 * STAGING differs by provenance, and the difference is deliberate. A reference lifted
 * out of HTML is decoded FIRST, because it is a URL. download_attachment's `cid:`
 * parameter is a handle that round-tripped from get_email's literal echo, so it
 * compares LITERALLY first and consults this key only when the literal matched
 * nothing — otherwise pasting back a Content-ID that really contains `%78` would
 * silently resolve to a different part. See docs/conventions.md; do not unify them.
 */
export function cidKey(ref: string): string {
  return decodeCidSrc(ref.replace(/^cid:/i, ''));
}

// How many code points of a foreign value an error or note may echo. Long enough to
// identify a real Content-ID or filename, short enough that a hostile one cannot bury
// the server's own sentence.
const DESCRIBE_PART_MAX = 64;

/**
 * Render a sender- or caller-supplied part value (a cid, a name, a content type) for
 * interpolation into server prose.
 *
 * These values are attacker-controlled on received mail, so they are treated as DATA:
 * control and format characters (which can reorder or hide the text around them) are
 * removed, line and paragraph separators with them, runs of space separators collapse
 * to one plain space, and a double quote becomes a single quote so the value cannot
 * close the quoted span the caller renders it inside. Every call site wraps the result
 * in double quotes, so hostile text reads as quoted data, never as the server speaking.
 *
 * Truncation is marked with an ellipsis rather than being silent: two long values that
 * differ only past the cap must not print identically.
 *
 * A value made entirely of stripped characters renders as the empty string, and no
 * placeholder is substituted: this function renders DATA inside a quoted span, so
 * inventing "(empty)" would put the server's words where the sender's are supposed to
 * be. The surrounding sentence always names the form the value came from, so an empty
 * pair of quotes still reads unambiguously.
 */
export function describePart(value: unknown): string {
  const source = typeof value === 'string' ? value : value == null ? '' : String(value);
  const cleaned = source
    .replace(/[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]/gu, '')
    .replace(/\p{Zs}+/gu, ' ')
    .replace(/"/g, "'");
  // Cap by CODE POINTS, not UTF-16 units — a unit slice can split a surrogate pair
  // and leave a lone surrogate in the message.
  const points = [...cleaned];
  return points.length > DESCRIBE_PART_MAX
    ? `${points.slice(0, DESCRIBE_PART_MAX).join('')}…`
    : cleaned;
}

// Windows device names, which cannot be used as a file's stem on that platform even
// with an extension ("CON.png" is the console, not a file). The trailing [ .]* is
// load-bearing: Win32 strips trailing spaces and dots from a path component BEFORE
// matching device names, so "CON .png" is the console too and an anchored bare-name
// test would miss it.
const WINDOWS_RESERVED_STEM = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])[ .]*$/i;

// The character treatment a filename derived from message content needs: strip control
// and format characters (\p{Cc} covers C0/C1 including U+0085; \p{Cf} covers bidi
// overrides such as U+202E), map path separators and the Windows drive/ADS colon to
// underscores, drop leading dots and the whitespace that would otherwise shield them,
// and cap by code point so a surrogate pair is never split.
//
// forward_email's sanitizeEmlFilename applies a similar treatment and is deliberately
// NOT folded into this one: the two diverge on both ends. That helper appends ".eml"
// (which neutralizes a reserved device name, so it lets them through) and falls back
// to "forwarded-message"; this one produces the name a download declares, where a
// reserved stem is a real hazard and the fallback is "attachment".
function sanitizeFilenameChars(value: string | null | undefined): string {
  const stripped = (value ?? '')
    .replace(/[\p{Cc}\p{Cf}]/gu, '')
    .replace(/[/\\:]/g, '_')
    // Leading dots and leading whitespace come off TOGETHER, in one pass. Stripping
    // dots first and trimming afterwards lets a single leading space shield the dot,
    // so " .hidden" would survive as ".hidden" - the exact name the strip exists to
    // prevent, since the sender chooses this value.
    .replace(/^[\s.]+/u, '')
    .replace(/\s+$/u, '');
  return [...stripped].slice(0, 80).join('').trim();
}

/**
 * The filename a download declares for an attachment, derived from the part's
 * sender-supplied `name`.
 *
 * Total by construction: a name that sanitizes to nothing (all controls, or absent)
 * becomes "attachment" rather than an empty segment, and a Windows device name gains
 * a trailing underscore on its stem ("CON.png" -> "CON_.png") so the value is usable
 * as a save name on every platform.
 */
export function sanitizeDownloadFilename(name: string | null | undefined): string {
  const cleaned = sanitizeFilenameChars(name) || 'attachment';
  const dot = cleaned.indexOf('.');
  const stem = dot === -1 ? cleaned : cleaned.slice(0, dot);
  if (!WINDOWS_RESERVED_STEM.test(stem)) return cleaned;
  return dot === -1 ? `${cleaned}_` : `${stem}_${cleaned.slice(dot)}`;
}
