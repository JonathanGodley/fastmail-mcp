// Helpers for embedded (cid:) image support (#13). No I/O, no JMAP client, no server
// state: everything here is a function over values, so the paths that consume it stay
// unit-testable without credentials or a network. The one non-deterministic function is
// `mintCid`, which draws from the CSPRNG; every other function is total and pure, and the
// map/reconcile helpers take an injectable mint so their callers' tests stay deterministic.
import sanitizeHtml from 'sanitize-html';
import { randomBytes } from 'node:crypto';

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

// ---------------------------------------------------------------------------
// URL normalization for classifying an <img src>
// ---------------------------------------------------------------------------

// Characters browsers ignore inside a URL. Stripping them is what stops `c id:x` and
// `cid&#9;:x` from smuggling a scheme past a naive `startsWith('cid:')` test.
const URL_IGNORED_CHARS = /[\x00-\x20]+/g;

// A URL scheme, in the shape the sanitizer's own gate recognizes: a letter followed by
// letters, digits, dot, plus or hyphen, up to the colon.
const URL_SCHEME = /^([a-zA-Z][a-zA-Z0-9.+-]*):/;

/**
 * Normalize a URL attribute value the way the HTML sanitizer's scheme gate does before it
 * decides whether a scheme is allowed: remove every character of code 0x20 and below
 * (browsers ignore those inside URLs in a surprising number of places), then clobber any
 * embedded `<!--…-->` comment (which a browser may drop inside an XML data island).
 *
 * WHY THIS IS REIMPLEMENTED rather than imported: the normalization lives in `launder`,
 * which reaches this project only as a transitive dependency of sanitize-html —
 * package.json declares sanitize-html alone. Importing it directly would take a hard
 * dependency on another package's dependency tree. So this is a deliberate mirror, and a
 * mirror can drift when either package updates.
 *
 * The tripwire for that drift is the obfuscated-spelling property tests: they assert that
 * a reference spelled `c id:x`, `cid&#9;:x` or `c<!--z-->id:x` classifies the same way a
 * plain `cid:x` does, and that no such spelling survives sanitization unmapped. If the
 * upstream normalization changes, those are what fail — do not weaken them, and do not
 * "simplify" the character class to `\s`, which excludes NUL and the other C0 controls
 * this one covers. Those controls are exactly what an obfuscated spelling would use.
 */
export function launderUrlValue(value: string): string {
  let out = value.replace(URL_IGNORED_CHARS, '');
  for (;;) {
    const open = out.indexOf('<!--');
    if (open === -1) break;
    const close = out.indexOf('-->', open + 4);
    // An unterminated comment marker stays in place, matching the upstream behaviour: a
    // browser would not treat it as a comment either.
    if (close === -1) break;
    out = out.slice(0, open) + out.slice(close + 3);
  }
  return out;
}

/**
 * The lowercase scheme of a URL attribute value after normalization, or null when the
 * value carries no scheme at all (a relative or scheme-less URL).
 */
export function urlScheme(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const match = URL_SCHEME.exec(launderUrlValue(value));
  return match ? match[1].toLowerCase() : null;
}

/** How an `<img src>` classifies once normalized. */
export type ImgSrcClass =
  | { kind: 'cid'; key: string }
  | { kind: 'data' }
  | { kind: 'remote'; scheme: string }
  | { kind: 'other' };

// The remote schemes an image may legitimately load from. Kept separate from the
// sanitizer's own scheme list because this classifier decides what the transform EMITS,
// and only these two make sense on an <img>.
const REMOTE_IMAGE_SCHEMES = new Set(['http', 'https']);

/**
 * Classify an `<img src>` after the same normalization the sanitizer's scheme gate applies.
 *
 * Exported because two very different consumers need the SAME answer: the quote sanitizer,
 * which decides what to emit, and the html-to-text derivation, which decides whether an
 * alt-less image deserves a placeholder in the plain-text alternative. Classifying twice
 * with two tests is how the derivation would end up disagreeing with what actually shipped.
 */
export function classifyImgSrc(src: unknown): ImgSrcClass {
  if (typeof src !== 'string' || src === '') return { kind: 'other' };
  const laundered = launderUrlValue(src);
  const match = URL_SCHEME.exec(laundered);
  if (!match) return { kind: 'other' };
  const scheme = match[1].toLowerCase();
  // The reference key comes from the NORMALIZED value, so every obfuscated spelling of
  // one reference produces one key.
  if (scheme === 'cid') return { kind: 'cid', key: cidKey(laundered) };
  if (scheme === 'data') return { kind: 'data' };
  if (REMOTE_IMAGE_SCHEMES.has(scheme)) return { kind: 'remote', scheme };
  return { kind: 'other' };
}

// ---------------------------------------------------------------------------
// Identity: the server's own embedded-image identifiers
// ---------------------------------------------------------------------------

// A Content-ID this server mints for an image it embeds on the caller's behalf. The
// domain is one RFC 2606 reserves as permanently unresolvable, so the identifier cannot
// collide with a real host, and the 128-bit random label makes an accidental collision
// with a foreign identifier a non-event.
//
// Deliberately KEYLESS. The alternative — signing the identifier with a server secret —
// fails in the dangerous direction: rotating or losing the key would turn "this is a part
// I manage, so remove it when the quote no longer shows it" into "this is a foreign part,
// so send it", and a key rotation would silently start mailing images that should have
// been dropped. A pure shape check fails the safe way instead: the most a forger achieves
// by copying the shape onto their own attachment is having their own content removed from
// a draft. Any change to how the identifier is built must change the `ii-` prefix too, so
// old and new forms stay distinguishable.
const MINTED_CID_SHAPE = /^ii-[0-9a-f]{32}@inline\.invalid$/i;

/** Mint a fresh Content-ID for an image this server embeds. */
export function mintCid(): string {
  return `ii-${randomBytes(16).toString('hex')}@inline.invalid`;
}

/**
 * True when a Content-ID has the shape this server mints.
 *
 * Used at two kinds of call site with two meanings, which is why the predicate carries two
 * exported names. As `isReservedCid` it is a CARRY-BOUNDARY check: an identifier of this
 * shape arriving on someone else's message is never carried verbatim, because a sender who
 * copies the shape onto an ordinary attachment would otherwise plant a part that a later
 * edit classifies as server-managed and removes from outgoing mail under the wrong
 * explanation. As `isOurMint` it classifies the parts this server itself put on a draft.
 * Same test on purpose — see the keyless rationale above.
 */
export function isReservedCid(value: unknown): boolean {
  return typeof value === 'string' && MINTED_CID_SHAPE.test(value);
}

/** The same reserved-shape test, named for the call sites that classify our own parts. */
export const isOurMint = isReservedCid;

// ---------------------------------------------------------------------------
// The two Content-ID vets
// ---------------------------------------------------------------------------

// An identifier a caller may author. Deliberately a narrow allowlist rather than a
// denylist of the characters known to cause trouble.
//
// SAFETY CONTROL, not tidiness. A Content-ID is written into a MIME header, and a value
// containing a carriage return or line feed is stored by the mail server as a REAL
// injected header: it ends the Content-ID header and begins one of the sender's choosing.
// Worse, reading the stored message back shows only the fragment before the break, so the
// injected header is invisible to anything short of the raw MIME source. This allowlist
// admits nothing outside [A-Za-z0-9._-], so no line break, no whitespace and no
// header-structural character can reach the header. Do not relax it.
const AUTHORABLE_CID = /^[A-Za-z0-9._-]{1,64}$/;

/** True when a caller-supplied Content-ID is one this server will author. */
export function isAuthorableCid(value: unknown): boolean {
  return typeof value === 'string' && AUTHORABLE_CID.test(value);
}

/**
 * Normalize the two spellings an agent realistically copies a Content-ID out of — an HTML
 * reference (`cid:logo`) and a raw header (`<logo>`) — before the authorable vet runs.
 *
 * Order is deliberate and each strip happens AT MOST ONCE: the enclosing angle-bracket
 * pair first, then one leading `cid:` prefix. So `<cid:logo>` (an HTML reference quoted the
 * way a header would be) normalizes to `logo`, while `cid:<logo>` normalizes to `<logo>`
 * and then fails the vet — that spelling is not one of the two real copy sources, and a
 * second angle-bracket pass would start accepting arbitrarily nested spellings. Authored
 * identifiers are simple local tokens by design, so `<logo@host>` still fails too: the
 * strip makes the two common spellings work, it does not widen what is accepted.
 *
 * Callers compare the NORMALIZED value when checking for duplicates, so two spellings of
 * one identifier count as one identifier rather than two parts sharing a Content-ID.
 */
export function stripCidSpelling(value: string): string {
  let out = value;
  if (out.length >= 2 && out.startsWith('<') && out.endsWith('>')) {
    out = out.slice(1, -1);
  }
  if (/^cid:/i.test(out)) {
    out = out.slice('cid:'.length);
  }
  return out;
}

// A stored Content-ID this server can reproduce faithfully when it recreates a draft. The
// length bound mirrors the RFC 5322 line-length limit already used for message identifiers.
//
// SAFETY CONTROL, for the same reason as the authorable vet: the printable-ASCII range
// excludes carriage return and line feed, which are a working MIME header-injection vector
// — a Content-ID carrying them is stored as a genuine extra header, and reading the message
// back shows only the pre-break fragment, so nothing but the raw MIME reveals it. This vet
// is what stops such a value being copied forward onto a recreated draft.
//
// The five excluded printable characters are excluded for fidelity rather than injection:
// angle brackets and the double quote are structural in a header and in an HTML attribute
// slot, and parentheses are comment delimiters that the RFC 5322 addr-spec a Content-ID is
// does not permit. Colon, semicolon and comma ARE admitted — they were measured to
// round-trip exactly through a store-and-recreate cycle, and excluding them would buy
// nothing. So are `@` and `.`, which are not optional: every Content-ID a real mail client
// writes contains an `@`, so excluding it would make this vet reject essentially every
// message composed elsewhere — exactly the population that needs to survive an edit.
const RECREATABLE_CID_RANGE = /^[\x21-\x7e]{1,998}$/;
const RECREATABLE_CID_EXCLUDED = /[<>"()]/;

/** True when a stored Content-ID can be reproduced verbatim on a recreated draft. */
export function isRecreatableCid(value: unknown): boolean {
  return (
    typeof value === 'string' &&
    RECREATABLE_CID_RANGE.test(value) &&
    !RECREATABLE_CID_EXCLUDED.test(value)
  );
}

// ---------------------------------------------------------------------------
// Quote sanitization: the two-pass factory
// ---------------------------------------------------------------------------

// The tag and attribute floor a quoted original is reduced to. Formatting survives;
// script/style/handlers and ALL unscoped attributes do not (there is no global '*' key, so
// style=/class=/on*= are removed — style being the classic CSS-exfiltration and mXSS
// vector). This is a safety floor for content re-sent under the user's own From address,
// matching what mainstream clients emit; it is not a tracker-pixel filter.
const QUOTE_ALLOWED_TAGS = [
  'p', 'div', 'span', 'br', 'b', 'i', 'strong', 'em', 'u', 'a', 'ul', 'ol', 'li',
  'blockquote', 'pre', 'code', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'table', 'thead', 'tbody', 'tr', 'td', 'th', 'img',
];

const QUOTE_ALLOWED_ATTRIBUTES = { a: ['href'], img: ['src', 'alt'] };

// The schemes any allowed attribute may carry. `cid` is NOT here and must never be added:
// this list governs href/cite/poster and every other URL-bearing attribute the sanitizer
// knows about, so a global entry would let an embedded-image reference land far outside an
// <img src>.
const QUOTE_ALLOWED_SCHEMES = ['http', 'https', 'mailto'];

// An <img> whose src the sanitizer stripped is removed entirely, so a quote never carries
// a broken-image placeholder.
const dropSrclessImages = (frame: sanitizeHtml.IFrame): boolean =>
  frame.tag === 'img' && !frame.attribs.src;

/** What one sanitization pass is for. */
export type SanitizeQuoteMode = 'collect' | 'map';

export interface SanitizeQuoteOptions {
  mode: SanitizeQuoteMode;
  /** Map mode only: reference key (see `cidKey`) to the Content-ID to emit for it. */
  cidMap?: Map<string, string>;
}

export interface SanitizeQuoteResult {
  /** The sanitized html. */
  html: string;
  /** Distinct reference keys found on an `<img src>`, in first-seen order. */
  refs: string[];
  /** Map mode: the distinct Content-IDs actually written into the html, in emit order. */
  embedded: string[];
  /** `<img>` elements whose src was a `data:` URI. */
  droppedDataImages: number;
  /** Embedded-image `<img>` elements this pass emitted no src for. */
  droppedCidImages: number;
  /**
   * MAP MODE ONLY. `<img>` elements carrying a real src that is neither an embedded-image
   * reference nor an http(s) URL — a relative or scheme-less path, a protocol-relative
   * `//host/…`, or an exotic scheme. Map mode is default-deny, so it drops these; the
   * collecting pass leaves them to the sanitizer and never counts any (a relative path is
   * not a scheme violation, so the scheme filter passes it through).
   *
   * That difference is the reason this is counted rather than dropped quietly: what ships
   * loses an image the other pass would have kept, and a reference relative to the
   * original's own origin cannot resolve from a message sent by someone else anyway.
   * An `<img>` with no src at all is not counted — there was no image to lose.
   */
  droppedUnsupportedImages: number;
}

/** Notified as the traversal walks each `<img>`. */
export interface ImgRefObserver {
  onCidRef?(key: string): void;
  onDataImage?(): void;
}

/**
 * The `<img>` transform the collecting pass installs.
 *
 * A HOOK on the single traversal, not a second parse: the sanitizer runs tag transforms
 * before it filters attributes, so this sees every `<img>` — including the ones whose src
 * the scheme filter is about to delete — and records the reference without the cost, or the
 * divergence risk, of walking the html twice. The attributes are returned untouched, so the
 * collecting pass emits exactly the html the sanitizer would have produced on its own.
 */
export function collectImgCidRefs(observer: ImgRefObserver): sanitizeHtml.Transformer {
  return (tagName, attribs) => {
    const classified = classifyImgSrc(attribs?.src);
    if (classified.kind === 'cid') observer.onCidRef?.(classified.key);
    else if (classified.kind === 'data') observer.onDataImage?.();
    return { tagName, attribs };
  };
}

/**
 * Sanitize an original's html for quoting, in one of two passes.
 *
 * COLLECT reports which embedded images the html references and decides nothing. Its html
 * is what a quotability check reads, and it is byte-for-byte what the sanitizer alone would
 * emit: `cid` is not in its scheme list, so every embedded image is dropped exactly as it
 * was before embedded images were supported. Quotability for an original whose only content
 * is embedded images therefore comes from the references resolving to real parts, never
 * from this html.
 *
 * MAP rewrites each reference that resolves through `cidMap` to the Content-ID the server
 * is attaching, and is DEFAULT-DENY: an `<img>` survives only when this transform
 * affirmatively emits a src for it — a mapped identifier, or a src whose normalized scheme
 * is http/https. Everything else has its src deleted and the element removed. That ordering
 * matters more than it looks. The classifier above reimplements the sanitizer's own URL
 * normalization, and if the two ever drift, an unrecognized spelling falls into the "not
 * affirmatively emitted" bucket and is dropped, instead of sliding through as the
 * scheme-less URL the sanitizer would have passed. The drop decision fails closed
 * independently of the reimplementation being right.
 *
 * A consequence worth knowing: a relative or scheme-less `<img src>` does NOT survive map
 * mode. Such a reference is already broken in mail — there is no base URL to resolve it
 * against — and admitting it would mean trusting the classifier's negative answer, which is
 * the thing this design refuses to do.
 */
export function sanitizeQuoteHtml(
  html: string,
  options: SanitizeQuoteOptions,
): SanitizeQuoteResult {
  const refs: string[] = [];
  const seenRefs = new Set<string>();
  const embedded: string[] = [];
  const seenEmbedded = new Set<string>();
  let droppedDataImages = 0;
  let droppedCidImages = 0;
  let droppedUnsupportedImages = 0;

  const recordRef = (key: string): void => {
    if (seenRefs.has(key)) return;
    seenRefs.add(key);
    refs.push(key);
  };

  let transformer: sanitizeHtml.Transformer;
  if (options.mode === 'collect') {
    transformer = collectImgCidRefs({
      onCidRef: (key) => {
        recordRef(key);
        // This configuration emits no embedded-image src at all, so every reference it
        // sees is one this pass drops.
        droppedCidImages++;
      },
      onDataImage: () => { droppedDataImages++; },
    });
  } else {
    const cidMap = options.cidMap;
    transformer = (tagName, attribs) => {
      const next: sanitizeHtml.Attributes = { ...attribs };
      const classified = classifyImgSrc(attribs?.src);
      if (classified.kind === 'cid') {
        recordRef(classified.key);
        const mapped = cidMap?.get(classified.key);
        if (mapped) {
          next.src = `cid:${mapped}`;
          if (!seenEmbedded.has(mapped)) {
            seenEmbedded.add(mapped);
            embedded.push(mapped);
          }
        } else {
          delete next.src;
          droppedCidImages++;
        }
      } else if (classified.kind === 'remote') {
        // Emitted verbatim: the scheme filter runs after this and re-checks the value, so
        // what ships is the original attribute, exactly as for any other allowed URL.
        next.src = typeof attribs.src === 'string' ? attribs.src : '';
      } else {
        if (classified.kind === 'data') droppedDataImages++;
        // A src this pass will not emit, but that named SOMETHING: counted so the drop is
        // visible. An <img> with no src (or a blank one) had no image to lose, and the
        // sanitizer's own srcless filter takes it either way.
        else if (hasWrittenSrc(attribs?.src)) droppedUnsupportedImages++;
        delete next.src;
      }
      return { tagName, attribs: next };
    };
  }

  const sanitized = sanitizeHtml(html, {
    allowedTags: QUOTE_ALLOWED_TAGS,
    allowedAttributes: QUOTE_ALLOWED_ATTRIBUTES,
    allowedSchemes: QUOTE_ALLOWED_SCHEMES,
    // A per-tag scheme list REPLACES the global list for that tag, so this admits `cid` on
    // <img> and nowhere else. Map mode needs it because the identifiers it writes are cid
    // URLs, which the scheme filter would otherwise strip straight back out.
    ...(options.mode === 'map'
      ? { allowedSchemesByTag: { img: ['http', 'https', 'cid'] } }
      : {}),
    allowProtocolRelative: false,
    transformTags: { img: transformer },
    exclusiveFilter: dropSrclessImages,
  });

  return {
    html: sanitized, refs, embedded, droppedDataImages, droppedCidImages, droppedUnsupportedImages,
  };
}

// Whether an `<img src>` names anything at all, read after the same laundering the scheme
// gate applies so a value made of control characters is not mistaken for a real reference.
function hasWrittenSrc(src: unknown): boolean {
  return typeof src === 'string' && launderUrlValue(src).trim() !== '';
}

// ---------------------------------------------------------------------------
// The broad collector
// ---------------------------------------------------------------------------

// Entity spellings that could write "cid:" (or the characters around it) without writing
// it literally. Deliberately a small named set plus the numeric forms rather than the full
// HTML5 table: a hit from this collector only ever produces a warning, so a spelling it
// misses costs a note, never a safety property.
const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  colon: ':', semi: ';', comma: ',', period: '.', commat: '@', num: '#',
  lowbar: '_', sol: '/', bsol: '\\', excl: '!', quest: '?', equals: '=',
  dollar: '$', percnt: '%', ast: '*', plus: '+', lpar: '(', rpar: ')',
};

// Decode ONCE, matching the single-decode posture used for references: `&amp;#58;` must
// decode to the text `&#58;` and not to a colon, or one value could be spelled several
// ways and still compare equal.
function decodeHtmlEntitiesOnce(value: string): string {
  return value.replace(
    /&(#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[a-zA-Z][a-zA-Z0-9]{1,31});/g,
    (whole, body: string) => {
      if (body[0] === '#') {
        const code = body[1] === 'x' || body[1] === 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
        if (!Number.isFinite(code)) return whole;
        try {
          return String.fromCodePoint(code);
        } catch {
          // Past the last code point: leave the source text alone. A surrogate value does
          // NOT land here — it is substituted, producing an unpaired surrogate. Harmless
          // for this collector, whose only job afterwards is to look for the literal text
          // of a reference, which no surrogate can spell part of.
          return whole;
        }
      }
      const named = NAMED_ENTITIES[body.toLowerCase()];
      return named === undefined ? whole : named;
    },
  );
}

// The run of characters that can follow `cid:` before something obviously ends the value.
// Whitespace, quotes, angle brackets, parentheses and the bracket families all terminate
// it: none of them can appear in an identifier this server would recreate.
const BROAD_CID_REF = /cid:([^\s"'<>()[\]{}\\]+)/gi;

// Punctuation that ends a sentence rather than an identifier. Stripped from the END only,
// so an identifier that legitimately contains a colon, semicolon or comma keeps it.
const TRAILING_SENTENCE_PUNCTUATION = /[.,;:!?]+$/;

/**
 * Every `cid:`-looking reference anywhere in some html, not only the ones on an `<img>`.
 *
 * BROAD on purpose, and a hit must never reject a message. It exists to notice a reference
 * in a place the precise `<img>` collector cannot see — a CSS `url()`, an SVG href, a
 * poster attribute — so the caller can say plainly that something was left as it was. It
 * also matches ordinary prose that merely mentions a reference: someone discussing this
 * feature, a pasted MIME fragment, this server's own error text quoted back. That
 * false-positive class is unbounded, which is precisely why a hit here is a warning and
 * only a real `<img>` reference is ever an error.
 *
 * The html is entity-decoded first so `&#99;id:x` is seen, and each value is percent
 * decoded once so the keys line up with the ones the precise collector produces.
 */
export function extractCidRefs(html: string | null | undefined): string[] {
  if (!html) return [];
  const decoded = decodeHtmlEntitiesOnce(html);
  const out: string[] = [];
  const seen = new Set<string>();
  for (const match of decoded.matchAll(BROAD_CID_REF)) {
    const raw = match[1].replace(TRAILING_SENTENCE_PUNCTUATION, '');
    if (!raw) continue;
    const key = decodeCidSrc(raw);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(key);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Resolving references to parts: the map, the reuse claim, the mint
// ---------------------------------------------------------------------------

/**
 * The fields of a message part these helpers reason about. Structurally satisfied by both
 * a raw JMAP body part and the attachment shape the client sends back, so nothing here
 * needs to know which one it was handed.
 */
export interface CidPart {
  cid?: string | null;
  blobId?: string | null;
  type?: string | null;
  name?: string | null;
  size?: number | null;
  disposition?: string | null;
}

/**
 * Whether a part declares itself an image, which is the only kind this server carries into a
 * body it composes.
 *
 * A bound worth stating precisely because it is weaker than it looks: the content type is
 * SENDER-DECLARED, exactly like a filename, so it says what the sender called the bytes and
 * nothing about what they are. Nothing is sniffed, and nothing verifies the claim. The gate
 * is still worth having — it keeps a reference from pulling an arbitrary file into a message
 * the recipient never asked for — but it is a declaration filter, not a content check. See
 * docs/security-model.md.
 */
export function isImageType(type: unknown): boolean {
  return classifyType(type).startsWith('image/');
}

/** A part this call newly attaches so the body it writes can display it. */
export interface MintedInlinePart {
  blobId: string;
  type: string;
  name?: string;
  cid: string;
  disposition: 'inline';
}

/** One reference that resolved to a part, and the Content-ID the rewritten body uses. */
export interface CidMapping {
  /** The reference key from the original's body. */
  ref: string;
  /** The Content-ID the rewritten body emits for it. */
  cid: string;
  /** True when an existing part on the draft supplied both the Content-ID and the bytes. */
  reused: boolean;
  /** The part the reference resolved to. */
  source: CidPart;
}

/** What a set of references resolved to, before anything is minted or rewritten. */
export interface CidRefResolution {
  /** The references, deduped, in first-seen order. */
  distinctRefs: string[];
  /** Reference key to the first part carrying that Content-ID. */
  byRef: Map<string, CidPart>;
  /** Content-IDs carried by more than one part; no reference can identify one of them. */
  ambiguousCids: Set<string>;
  /**
   * Distinct parts the references resolved to, in first-reference order.
   *
   * A reference onto a Content-ID two parts share resolves to BOTH of them, not to the
   * first: neither can be embedded, and a count that named only one would report a single
   * lost image where the reader lost two.
   */
  resolvedParts: CidPart[];
  /** References that matched no part at all. Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  /**
   * References that would really embed: they resolved to exactly one part, that part declares
   * itself an image, and it has a blob to re-reference.
   *
   * Separate from `resolvedParts` because a body whose only content is embedded images is
   * worth quoting only if at least one of them can actually be carried — quoting on
   * "resolved" alone would emit an attribution over a quote that shows nothing.
   */
  embeddableRefs: string[];
}

/**
 * Match a body's embedded-image references against the parts of the message that carries it.
 *
 * Pure and mint-free, so a caller can ask what a quote WOULD carry before deciding whether to
 * quote at all — the decision that has to be made before any identifier is minted, since a
 * minted part with nothing referencing it is a loose attachment on the finished message.
 */
export function resolveCidRefs(refs: string[], sourceParts: CidPart[]): CidRefResolution {
  const distinctRefs = [...new Set(refs ?? [])];
  const parts = sourceParts ?? [];

  const groups = new Map<string, CidPart[]>();
  const byRef = new Map<string, CidPart>();
  for (const part of parts) {
    if (typeof part?.cid !== 'string' || !part.cid) continue;
    const group = groups.get(part.cid);
    if (group) group.push(part);
    else groups.set(part.cid, [part]);
    if (!byRef.has(part.cid)) byRef.set(part.cid, part);
  }
  const ambiguousCids = new Set(
    [...groups.entries()].filter(([, g]) => g.length > 1).map(([cid]) => cid),
  );

  const resolvedParts: CidPart[] = [];
  const seen = new Set<CidPart>();
  const unresolvedRefs: string[] = [];
  const embeddableRefs: string[] = [];
  for (const ref of distinctRefs) {
    const group = groups.get(ref);
    if (!group) {
      unresolvedRefs.push(ref);
      continue;
    }
    for (const part of group) {
      if (seen.has(part)) continue;
      seen.add(part);
      resolvedParts.push(part);
    }
    const part = group[0];
    const blobId = typeof part.blobId === 'string' && part.blobId ? part.blobId : null;
    if (blobId && isImageType(part.type) && !ambiguousCids.has(ref)) embeddableRefs.push(ref);
  }

  return { distinctRefs, byRef, ambiguousCids, resolvedParts, unresolvedRefs, embeddableRefs };
}

export interface BuildCidMapInput {
  /** Reference keys from the original's body, in first-seen order. Repeats are tolerated. */
  refs: string[];
  /** The original's parts. Their own Content-IDs are compared LITERALLY. */
  sourceParts: CidPart[];
  /**
   * Parts already on the draft being edited that survived this call's removals and carry
   * a Content-ID of this server's own shape. Empty when composing, where no draft exists.
   */
  survivors?: CidPart[];
  /** Injected so callers' tests are deterministic. */
  mint?: () => string;
}

export interface BuildCidMapResult {
  /** Reference key to the Content-ID to emit — the input the mapping pass takes. */
  cidMap: Map<string, string>;
  /** One entry per reference that will be embedded, in reference order. */
  mappings: CidMapping[];
  /** Parts to attach on their OWN channel. Never merged into a carried attachment set here. */
  minted: MintedInlinePart[];
  /** Content-IDs of the surviving parts a reference claimed; these ride the normal carry. */
  reusedCids: string[];
  /** Survivors no reference claimed, in stored order. */
  unclaimedSurvivors: CidPart[];
  /** References that matched no part at all. Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  /** Distinct parts a reference resolved to but which cannot be embedded. */
  unembeddableParts: CidPart[];
  /** Distinct parts the references resolved to — the denominator a shortfall reports against. */
  resolvedPartCount: number;
}

/**
 * Decide, for every reference in a body about to be quoted, which part supplies it and
 * under which Content-ID.
 *
 * REUSE comes first and is scoped to survivors — the parts already on the draft that this
 * call's removals left in place. A survivor whose blob matches the resolved part supplies
 * both the bytes and its stored Content-ID, so an ordinary edit does not renumber the
 * images a client has already rendered once. Matching is ONE-TO-ONE: candidates are taken
 * in stored order and each is claimed at most once, so two references over one blob claim
 * two survivors when two exist and the second falls back to a fresh mint when only one
 * does. Two images never collapse into one part, and a many-to-one match never leaves a
 * survivor looking unreferenced.
 *
 * A reference with no surviving match carries the original's blob under a freshly minted
 * identifier. Those parts come back on `minted` and are the caller's to attach as a
 * separate assembly step — this function never folds them into an existing attachment set,
 * because doing so would make a minted part indistinguishable from one the caller carried.
 */
export function buildCidMap(input: BuildCidMapInput): BuildCidMapResult {
  const mint = input.mint ?? mintCid;
  const sourceParts = input.sourceParts ?? [];

  // The reference-to-part match, and with it the ambiguous-Content-ID set. Shared with the
  // callers that have to decide whether an original is worth quoting BEFORE anything is
  // minted, so the quotability decision and this mapping can never disagree about which
  // references resolve. Distinct references, first-seen order: the collecting pass already
  // dedupes, but a repeated reference would claim two survivors for one image and leave the
  // second reused part attached with nothing pointing at it, and the closure check covers
  // only freshly minted identifiers, so nothing downstream would notice.
  const resolution = resolveCidRefs(input.refs ?? [], sourceParts);
  const refs = resolution.distinctRefs;
  const byCid = resolution.byRef;

  // Only a survivor carrying an identifier of this server's own shape is a reuse candidate.
  // Reusing a foreign one would write someone else's identifier into a body this server
  // composed, and later classify that part as server-managed.
  const survivors = (input.survivors ?? []).filter((s) => isReservedCid(s?.cid));
  const claimed = new Set<number>();

  const cidMap = new Map<string, string>();
  const mappings: CidMapping[] = [];
  const minted: MintedInlinePart[] = [];
  const reusedCids: string[] = [];

  for (const ref of refs) {
    const part = byCid.get(ref);
    // A reference that matched no part is already recorded by the resolution above.
    if (!part) continue;

    // The three ways a reference that DID find a part still fails to embed: a Content-ID
    // naming more than one part cannot identify any one of them; a part with no blob cannot
    // be re-referenced on a new message; and a part that does not declare itself an image is
    // not something this server pulls into a body it composes, whatever an <img> claims about
    // it. Naming them here is what makes a shortfall report explicable rather than mysterious.
    const ambiguous = resolution.ambiguousCids.has(ref);
    const blobId = typeof part.blobId === 'string' && part.blobId ? part.blobId : null;
    if (ambiguous || !blobId || !isImageType(part.type)) continue;

    let cid: string | null = null;
    let reused = false;
    for (let i = 0; i < survivors.length; i++) {
      if (claimed.has(i)) continue;
      if (survivors[i].blobId !== blobId) continue;
      claimed.add(i);
      cid = survivors[i].cid as string;
      reused = true;
      reusedCids.push(cid);
      break;
    }

    if (cid === null) {
      cid = mint();
      minted.push({
        blobId,
        // Non-empty by construction: the gate above admits only a declared image type.
        type: part.type as string,
        ...(typeof part.name === 'string' && part.name ? { name: part.name } : {}),
        cid,
        disposition: 'inline',
      });
    }

    cidMap.set(ref, cid);
    mappings.push({ ref, cid, reused, source: part });
  }

  const unclaimedSurvivors = survivors.filter((_, i) => !claimed.has(i));

  // Derived from the resolution rather than accumulated in the loop above, so it cannot
  // disagree with the count it is the shortfall against: every part a reference landed on,
  // less the ones a mapping actually embedded.
  const embeddedSources = new Set(mappings.map((m) => m.source));
  const unembeddableParts = resolution.resolvedParts.filter((p) => !embeddedSources.has(p));

  return {
    cidMap,
    mappings,
    minted,
    reusedCids,
    unclaimedSurvivors,
    unresolvedRefs: resolution.unresolvedRefs,
    unembeddableParts,
    resolvedPartCount: resolution.resolvedParts.length,
  };
}

// ---------------------------------------------------------------------------
// Reconciling the parts a rebuilt draft carries
// ---------------------------------------------------------------------------

/** What happens to one part already on the draft. */
export type InlinePartAction =
  /** Rides the rebuilt draft as it stands. */
  | 'kept'
  /** Rides the rebuilt draft, but as a regular attachment rather than an embedded image. */
  | 'degraded'
  /** Comes off the rebuilt draft entirely. */
  | 'removed';

export interface ReconciledPart {
  part: CidPart;
  action: InlinePartAction;
}

export interface ReconcileInlinePartsInput {
  /** Parts on the draft that survived this call's explicit removals, in stored order. */
  storedParts: CidPart[];
  /** Every Content-ID the FINAL bodies reference. Compared literally against a part's own. */
  referencedCids: string[];
  /** Parts this call minted, passed through untouched onto their own channel. */
  minted?: MintedInlinePart[];
  /**
   * Whether an html body ships. Defaults to true. With no html body there is nothing to
   * display an embedded image, and the mail server rejects an inline disposition outright,
   * so an image can only ride as a regular attachment.
   */
  htmlShips?: boolean;
}

export interface ReconcileInlinePartsResult {
  /** Every stored part with its outcome, in stored order. */
  parts: ReconciledPart[];
  /** The minted channel, verbatim and separate. */
  minted: MintedInlinePart[];
  /** Convenience views over `parts`. */
  kept: CidPart[];
  degraded: CidPart[];
  removed: CidPart[];
}

function isInlineDisposition(part: CidPart): boolean {
  return typeof part.disposition === 'string' && part.disposition.trim().toLowerCase() === 'inline';
}

/**
 * Decide what becomes of each part already on a draft once the rebuilt bodies are known.
 *
 * A part the final bodies still reference stays as it is. An unreferenced part carrying an
 * identifier of this server's own shape is REMOVED: this server put it there to display an
 * image in a body that no longer shows it, so leaving it behind would attach a file the
 * user never asked to send. An unreferenced part that was inline but carries someone
 * else's identifier is DEGRADED to a regular attachment instead of removed — the bytes
 * came from the caller or from a message being carried, so silently dropping them would
 * lose content this server did not create. Everything else is an ordinary attachment and
 * is left alone.
 *
 * Minted parts are returned on their own field and are never merged into the stored set
 * here: keeping the channels separate is what lets the caller attach them as an explicit
 * assembly step and lets a later call tell a minted part from a carried one.
 */
export function reconcileInlineParts(
  input: ReconcileInlinePartsInput,
): ReconcileInlinePartsResult {
  const htmlShips = input.htmlShips !== false;
  const referenced = new Set(htmlShips ? input.referencedCids ?? [] : []);

  const parts: ReconciledPart[] = [];
  const kept: CidPart[] = [];
  const degraded: CidPart[] = [];
  const removed: CidPart[] = [];

  for (const part of input.storedParts ?? []) {
    if (!part) continue;
    const cid = typeof part.cid === 'string' ? part.cid : '';
    let action: InlinePartAction;
    if (cid && referenced.has(cid)) action = 'kept';
    else if (isReservedCid(cid)) action = 'removed';
    else if (isInlineDisposition(part)) action = 'degraded';
    else action = 'kept';

    parts.push({ part, action });
    if (action === 'kept') kept.push(part);
    else if (action === 'degraded') degraded.push(part);
    else removed.push(part);
  }

  return { parts, minted: input.minted ?? [], kept, degraded, removed };
}

// ---------------------------------------------------------------------------
// The closure invariant
// ---------------------------------------------------------------------------

/**
 * A self-check failure: the message this call assembled is internally inconsistent. Not a
 * caller error — a caller cannot cause it by supplying bad input, because every input-shaped
 * problem is rejected before assembly. Callers map it to an internal error.
 */
export class InlineClosureError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InlineClosureError';
  }
}

export interface InlineClosureInput {
  /** The html bodies this call wrote or rebuilt. Bodies it merely carried are not checked. */
  htmlBodies?: (string | null | undefined)[];
  /** The Content-IDs of every part the assembled message carries. */
  finalPartCids?: (string | null | undefined)[];
  /** Content-IDs this call minted and attached. */
  attachedMintedCids?: string[];
  /** Skip the check entirely — for a call that assembled no body and minted nothing. */
  skip?: boolean;
}

/**
 * Assert that the message this call assembled closes over its own embedded images, in both
 * directions.
 *
 * Every reference in a body this call wrote must resolve to a part the message carries, or
 * the recipient sees a broken image. And every part this call minted must be referenced by
 * one of those bodies, or the recipient receives an unexplained file attachment. Neither
 * failure is reachable from caller input — the rejects upstream see to that — so a failure
 * here means the assembly itself is wrong, which is why it raises rather than degrading.
 *
 * SCOPE is deliberately narrow on both arms. Arm one reads only the bodies this call
 * produced, so a pre-existing broken reference in a body being carried through untouched is
 * not this call's problem to fail on. Arm two covers only identifiers this call minted, so a
 * part the caller attached and never referenced — a perfectly ordinary attachment — is not
 * mistaken for a loose end.
 */
export function checkInlineClosure(input: InlineClosureInput): void {
  if (input.skip) return;

  const bodies = (input.htmlBodies ?? []).filter(
    (b): b is string => typeof b === 'string' && b !== '',
  );
  const attachedMinted = input.attachedMintedCids ?? [];
  if (bodies.length === 0 && attachedMinted.length === 0) return;

  const refs = new Set<string>();
  for (const body of bodies) {
    for (const ref of sanitizeQuoteHtml(body, { mode: 'collect' }).refs) refs.add(ref);
  }

  const finalCids = new Set(
    (input.finalPartCids ?? []).filter((c): c is string => typeof c === 'string' && c !== ''),
  );

  for (const ref of refs) {
    if (finalCids.has(ref)) continue;
    throw new InlineClosureError(
      `The composed message body references embedded image "${describePart(ref)}", ` +
      'but no part of the assembled message supplies it.',
    );
  }

  for (const cid of attachedMinted) {
    if (refs.has(cid)) continue;
    throw new InlineClosureError(
      `Embedded image "${describePart(cid)}" was attached to the composed message, ` +
      'but no body written by this call references it.',
    );
  }
}
