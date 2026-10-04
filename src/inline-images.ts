// Helpers for embedded (cid:) image support (#13). Pure except `mintCid`; the map/reconcile
// helpers take an injectable mint so their callers' tests stay deterministic.
import sanitizeHtml from 'sanitize-html';
import { randomBytes } from 'node:crypto';
import { trimEnd, isWhitespace } from './trim-end.js';

/**
 * A part of an email, paired with where the server routed it.
 *
 * `part` is the JMAP EmailBodyPart VERBATIM, never copied or mutated, so a derived value
 * cannot leak into a `raw: true` response. `inBodyList` is the only structural signal that
 * a part is body content rather than a separate file (RFC 8621 §4.1.4).
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

// Excluding the two body types rather than allowlisting image/audio/video (RFC 8621 §4.1.4)
// is deliberate: a superset, so a shape the RFC does not enumerate is surfaced, not dropped.
const BODY_TEXT_TYPES = new Set(['text/plain', 'text/html']);

// For CLASSIFYING only (RFC 2045 §5.1): the value emitted for a part is the server's own.
function classifyType(type: unknown): string {
  if (typeof type !== 'string') return '';
  const semicolon = type.indexOf(';');
  return (semicolon === -1 ? type : type.slice(0, semicolon)).trim().toLowerCase();
}

// A part with neither id cannot be matched, so it is kept as its own entry.
function partKey(part: any): string | null {
  if (typeof part?.partId === 'string' && part.partId) return `p:${part.partId}`;
  if (typeof part?.blobId === 'string' && part.blobId) return `b:${part.blobId}`;
  return null;
}

/**
 * The part set an attachment-aware read should work from: the JMAP `attachments`
 * array UNION the media parts the server routed into `textBody`/`htmlBody`.
 *
 * The same embedded image lands in `attachments` for one MIME shape and in the body lists
 * for another (RFC 8621 §4.1.4), so `attachments` alone is not a complete listing.
 *
 * GATED on `attachments` being present: the compact list/search set fetches `textBody` but
 * not `attachments`, and must not emit half a listing. The gate lives here so it cannot be
 * forgotten at a call site.
 *
 * The order is stable (attachments, then textBody additions, then htmlBody), because
 * download_attachment's entry-number form indexes into it.
 */
export function buildUnionParts(email: UnionSourceEmail | null | undefined): UnionPart[] {
  const attachments = email?.attachments;
  if (!Array.isArray(attachments)) return [];

  const bodyLists = [email?.textBody, email?.htmlBody].filter(Array.isArray) as any[][];

  // Resolved BEFORE the walk, so the routing signal does not depend on which list wins the dedup.
  const bodyKeys = new Set<string>();
  for (const list of bodyLists) {
    for (const part of list) {
      const key = partKey(part);
      if (key !== null) bodyKeys.add(key);
    }
  }

  const union: UnionPart[] = [];
  const seen = new Set<string>();
  // For a part with no ids: RFC 8621 §4.1.4 puts one displayed part into BOTH body lists, so
  // without this the same object would be listed twice, shifting entry numbers.
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
      // A typeless part is body text, matching the body extractor.
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
 * SINGLE decode, deliberately: repeated decoding would let `%2525` reach a comparison as
 * `%`. A malformed escape is handed back verbatim for the literal comparison to decide.
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
 * Reference side ONLY: a part's own `cid` is a Content-ID, compared LITERALLY, never
 * decoded, or a Content-ID containing `%78` would be reachable as `x`.
 *
 * download_attachment's `cid:` parameter compares literally FIRST and consults this key only
 * on a miss, since it round-tripped from a literal echo. See docs/conventions.md; do not
 * unify the two stagings.
 */
export function cidKey(ref: string): string {
  return decodeCidSrc(ref.replace(/^cid:/i, ''));
}

// Long enough to identify a real Content-ID or filename, short enough that a hostile one
// cannot bury the server's own sentence.
export const DESCRIBE_PART_MAX = 64;

/**
 * Render an attacker-controlled part value (a cid, a name, a content type) as DATA for
 * server prose. The `"` to `'` swap protects a DOUBLE-quoted span only (#190); inside `'…'`
 * the value's own `'` closes the span.
 *
 * So: A CALLER THAT QUOTES QUOTES WITH `"…"`. A caller that renders the value BARE is judged
 * on the whole sentence: A NEW `'…'` SPAN IN ANY SENTENCE THAT RENDERS A BARE ONE REOPENS
 * THIS. No list of bare callers is kept, deliberately; the mechanical half is a drift guard
 * in coerce.test.ts.
 *
 * Truncation is marked, so two values differing past the cap never print identically. An
 * all-stripped value renders as '' with no placeholder: the server's words must not stand
 * where the sender's belong.
 */
export function describePart(value: unknown, max: number = DESCRIBE_PART_MAX): string {
  const source = typeof value === 'string' ? value : value == null ? '' : String(value);
  const cleaned = source
    .replace(/[\p{Cc}\p{Cf}\p{Zl}\p{Zp}]/gu, '')
    .replace(/\p{Zs}+/gu, ' ')
    .replace(/"/g, "'");
  // Code points, so a surrogate pair is never split.
  const points = [...cleaned];
  return points.length > max ? `${points.slice(0, max).join('')}…` : cleaned;
}

// The trailing [ .]* is load-bearing: Win32 strips trailing spaces and dots before
// matching device names, so "CON .png" is the console too.
const WINDOWS_RESERVED_STEM = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])[ .]*$/i;

// draft_email's sanitizeEmlFilename is similar and deliberately NOT folded into this: it
// leaves a device name for the receiving client to handle, and its fallback differs.
function sanitizeFilenameChars(value: string | null | undefined): string {
  const stripped = trimEnd(
    (value ?? '')
      .replace(/[\p{Cc}\p{Cf}]/gu, '')
      .replace(/[/\\:]/g, '_')
      // TOGETHER, in one pass: trimming afterwards would let a leading space shield the dot,
      // so " .hidden" would survive as ".hidden".
      .replace(/^[\s.]+/u, ''),
    isWhitespace,
  );
  return [...stripped].slice(0, 80).join('').trim();
}

/**
 * The filename a download declares, from the sender-supplied `name`. Total: never empty,
 * and a Windows device stem gains an underscore ("CON.png" -> "CON_.png").
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

// In the shape the sanitizer's own gate recognizes.
const URL_SCHEME = /^([a-zA-Z][a-zA-Z0-9.+-]*):/;

/**
 * Normalize a URL attribute value the way the HTML sanitizer's scheme gate does: remove
 * every character of code 0x20 and below, then clobber any embedded `<!--…-->` comment.
 *
 * A deliberate MIRROR of `launder`, which reaches this project only as a transitive
 * dependency of sanitize-html. The obfuscated-spelling property tests are the drift
 * tripwire: do not weaken them, and do not "simplify" the class to `\s`, which misses NUL
 * and the other C0 controls an obfuscated spelling would use.
 */
export function launderUrlValue(value: string): string {
  let out = value.replace(URL_IGNORED_CHARS, '');
  for (;;) {
    const open = out.indexOf('<!--');
    if (open === -1) break;
    const close = out.indexOf('-->', open + 4);
    // Unterminated stays, as upstream: a browser would not treat it as a comment either.
    if (close === -1) break;
    out = out.slice(0, open) + out.slice(close + 3);
  }
  return out;
}

export function urlScheme(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const match = URL_SCHEME.exec(launderUrlValue(value));
  return match ? match[1].toLowerCase() : null;
}

export type ImgSrcClass =
  | { kind: 'cid'; key: string }
  | { kind: 'data' }
  | { kind: 'remote'; scheme: string }
  | { kind: 'other' };

// Separate from the sanitizer's scheme list: this decides what the transform EMITS on an <img>.
const REMOTE_IMAGE_SCHEMES = new Set(['http', 'https']);

/**
 * Classify an `<img src>` after the sanitizer's own normalization. Shared by the quote
 * sanitizer and the html-to-text derivation so the two cannot disagree about what shipped.
 */
export function classifyImgSrc(src: unknown): ImgSrcClass {
  if (typeof src !== 'string' || src === '') return { kind: 'other' };
  const laundered = launderUrlValue(src);
  const match = URL_SCHEME.exec(laundered);
  if (!match) return { kind: 'other' };
  const scheme = match[1].toLowerCase();
  // From the NORMALIZED value, so every obfuscated spelling produces one key.
  if (scheme === 'cid') return { kind: 'cid', key: cidKey(laundered) };
  if (scheme === 'data') return { kind: 'data' };
  if (REMOTE_IMAGE_SCHEMES.has(scheme)) return { kind: 'remote', scheme };
  return { kind: 'other' };
}

// ---------------------------------------------------------------------------
// Identity: the server's own embedded-image identifiers
// ---------------------------------------------------------------------------

// The domain is RFC 2606-reserved, so it cannot collide with a real host.
//
// Deliberately KEYLESS: a signed identifier fails dangerously, since losing or rotating the
// key would turn "a part I manage, remove it" into "a foreign part, send it". A shape check
// fails safe: a forger copying the shape only gets their own content removed from a draft.
// Any change to how the identifier is built must change the `ii-` prefix too.
const MINTED_CID_SHAPE = /^ii-[0-9a-f]{32}@inline\.invalid$/i;

export function mintCid(): string {
  return `ii-${randomBytes(16).toString('hex')}@inline.invalid`;
}

/**
 * True when a Content-ID has the shape this server mints. Two names for two meanings: as
 * `isReservedCid` it is a CARRY-BOUNDARY check (a foreign identifier of this shape is never
 * carried verbatim); as `isOurMint` it classifies parts this server put on a draft.
 */
export function isReservedCid(value: unknown): boolean {
  return typeof value === 'string' && MINTED_CID_SHAPE.test(value);
}

export const isOurMint = isReservedCid;

// ---------------------------------------------------------------------------
// The two Content-ID vets
// ---------------------------------------------------------------------------

// SAFETY CONTROL, not tidiness: a CR or LF in a Content-ID is stored as a REAL injected MIME
// header, invisible on read-back short of the raw source. A narrow allowlist, not a
// denylist. Do not relax it.
const AUTHORABLE_CID = /^[A-Za-z0-9._-]{1,64}$/;

export function isAuthorableCid(value: unknown): boolean {
  return typeof value === 'string' && AUTHORABLE_CID.test(value);
}

/**
 * Normalize the two spellings an agent realistically copies a Content-ID out of — an HTML
 * reference (`cid:logo`) and a raw header (`<logo>`) — before the authorable vet runs.
 *
 * Each strip happens AT MOST ONCE, brackets then `cid:`, so `cid:<logo>` still fails the
 * vet: a second pass would accept arbitrarily nested spellings. Callers compare the
 * NORMALIZED value for duplicates.
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

// SAFETY CONTROL, as for the authorable vet: the printable-ASCII range excludes CR and LF,
// so an injected header is never copied forward onto a recreated draft. The length bound is
// RFC 5322's line limit.
//
// The excluded printables are structural (`<>"`) or RFC 5322 comment delimiters (`()`).
// Colon, semicolon and comma round-trip exactly (measured), and `@` must stay: every
// Content-ID a real client writes contains one.
const RECREATABLE_CID_RANGE = /^[\x21-\x7e]{1,998}$/;
const RECREATABLE_CID_EXCLUDED = /[<>"()]/;

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

// A safety floor for content re-sent under the user's own From, not a tracker-pixel filter.
// No global '*' attribute key, so style= (the classic CSS-exfil/mXSS vector) is removed.
const QUOTE_ALLOWED_TAGS = [
  'p', 'div', 'span', 'br', 'b', 'i', 'strong', 'em', 'u', 'a', 'ul', 'ol', 'li',
  'blockquote', 'pre', 'code', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'table', 'thead', 'tbody', 'tr', 'td', 'th', 'img',
];

const QUOTE_ALLOWED_ATTRIBUTES = { a: ['href'], img: ['src', 'alt'] };

// `cid` must NEVER be added here: this list governs every URL-bearing attribute, so it would
// let a reference land far outside an <img src>.
const QUOTE_ALLOWED_SCHEMES = ['http', 'https', 'mailto'];

// So a quote never carries a broken-image placeholder.
const dropSrclessImages = (frame: sanitizeHtml.IFrame): boolean =>
  frame.tag === 'img' && !frame.attribs.src;

export type SanitizeQuoteMode = 'collect' | 'map';

export interface SanitizeQuoteOptions {
  mode: SanitizeQuoteMode;
  /** Map mode only: reference key (see `cidKey`) to the Content-ID to emit for it. */
  cidMap?: Map<string, string>;
}

export interface SanitizeQuoteResult {
  html: string;
  /** Distinct reference keys found on an `<img src>`, in first-seen order. */
  refs: string[];
  /** Map mode: the distinct Content-IDs actually written, in emit order. */
  embedded: string[];
  droppedDataImages: number;
  droppedCidImages: number;
  /**
   * MAP MODE ONLY: a real src that is not cid, data: or http(s), such as a relative path, which
   * the sanitizer alone would pass. Map mode drops it, so it is counted to keep the drop
   * visible. An `<img>` with no src is not counted.
   */
  droppedUnsupportedImages: number;
}

export interface ImgRefObserver {
  onCidRef?(key: string): void;
  onDataImage?(): void;
}

/**
 * The collecting pass's `<img>` transform: a HOOK on the single traversal, not a second
 * parse. Tag transforms run before attribute filtering, so it sees srcs the scheme filter is
 * about to delete; attributes are returned untouched.
 *
 * Only `<img src>` is read. A cid reached any other way (`srcset`, `<input type="image">`,
 * SVG `<image href>`, a `background` attribute, CSS `url(cid:…)`) is not collected, so the
 * checks built on this pass (the signature image refusal, the dangling-reference checks on a
 * body) do not see it, and such an image ships unresolved.
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
 * COLLECT reports references and decides nothing; its html is byte-for-byte what the
 * sanitizer alone emits, with every embedded image dropped.
 *
 * MAP is DEFAULT-DENY: an `<img>` survives only when this transform affirmatively emits a src
 * (a mapped identifier, or http/https). So if the mirrored URL normalization ever drifts, an
 * unrecognized spelling is dropped rather than sliding through; a relative src does not
 * survive either, since admitting it would mean trusting the classifier's negative answer.
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
        // Verbatim: the scheme filter runs after this and re-checks the value.
        next.src = typeof attribs.src === 'string' ? attribs.src : '';
      } else {
        if (classified.kind === 'data') droppedDataImages++;
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
    // A per-tag list REPLACES the global one for that tag, so `cid` is admitted on <img> only.
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

// Laundered first, so a value made of control characters is not a real reference.
function hasWrittenSrc(src: unknown): boolean {
  return typeof src === 'string' && launderUrlValue(src).trim() !== '';
}

// ---------------------------------------------------------------------------
// The broad collector
// ---------------------------------------------------------------------------

// Deliberately not the full HTML5 table: a hit here only ever produces a warning, so a
// missed spelling costs a note, never a safety property.
const NAMED_ENTITIES: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
  colon: ':', semi: ';', comma: ',', period: '.', commat: '@', num: '#',
  lowbar: '_', sol: '/', bsol: '\\', excl: '!', quest: '?', equals: '=',
  dollar: '$', percnt: '%', ast: '*', plus: '+', lpar: '(', rpar: ')',
};

// ONCE, as for references: `&amp;#58;` decodes to the text `&#58;`, not a colon.
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
          // Past the last code point. A surrogate value does not land here; the unpaired
          // surrogate it produces is harmless, since none can spell part of a reference.
          return whole;
        }
      }
      const named = NAMED_ENTITIES[body.toLowerCase()];
      return named === undefined ? whole : named;
    },
  );
}

// None of the terminators can appear in an identifier this server would recreate.
const BROAD_CID_REF = /cid:([^\s"'<>()[\]{}\\]+)/gi;

// From the END only, so an identifier containing a colon or comma keeps it.
const SENTENCE_PUNCTUATION = new Set(['.', ',', ';', ':', '!', '?']);

/**
 * Every `cid:`-looking reference anywhere in some html, not only on an `<img>` (a CSS
 * `url()`, an SVG href). BROAD on purpose, so a hit must NEVER reject a message: it also
 * matches prose, an unbounded false-positive class. Keys line up with the precise collector's.
 */
export function extractCidRefs(html: string | null | undefined): string[] {
  if (!html) return [];
  const decoded = decodeHtmlEntitiesOnce(html);
  const out: string[] = [];
  const seen = new Set<string>();
  for (const match of decoded.matchAll(BROAD_CID_REF)) {
    const raw = trimEnd(match[1], (ch) => SENTENCE_PUNCTUATION.has(ch));
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

/** Satisfied by both a raw JMAP body part and the attachment shape the client sends back. */
export interface CidPart {
  cid?: string | null;
  blobId?: string | null;
  type?: string | null;
  name?: string | null;
  size?: number | null;
  disposition?: string | null;
}

/**
 * Whether a part DECLARES itself an image, the only kind carried into a composed body. The
 * type is sender-declared and nothing is sniffed: a declaration filter, not a content check.
 * See docs/security-model.md.
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

export interface CidMapping {
  ref: string;
  /** The Content-ID the rewritten body emits for it. */
  cid: string;
  /** True when an existing part on the draft supplied both the Content-ID and the bytes. */
  reused: boolean;
  source: CidPart;
}

/** What a set of references resolved to, before anything is minted or rewritten. */
export interface CidRefResolution {
  distinctRefs: string[];
  /** Reference key to the first part carrying that Content-ID. */
  byRef: Map<string, CidPart>;
  ambiguousCids: Set<string>;
  /**
   * In first-reference order. A shared Content-ID resolves to BOTH parts, or the count would
   * report one lost image where the reader lost two.
   */
  resolvedParts: CidPart[];
  /** Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  /**
   * References that resolved to exactly one image part with a blob. Separate from
   * `resolvedParts` because an image-only body is worth quoting only if one can be carried.
   */
  embeddableRefs: string[];
}

/**
 * Match a body's embedded-image references against the parts of the message that carries it.
 * Mint-free, so a caller can decide whether to quote before any identifier is minted.
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
  refs: string[];
  /** Their own Content-IDs are compared LITERALLY. */
  sourceParts: CidPart[];
  /** Parts on the draft being edited that survived this call's removals. Empty when composing. */
  survivors?: CidPart[];
  /** Injected so callers' tests are deterministic. */
  mint?: () => string;
}

export interface BuildCidMapResult {
  cidMap: Map<string, string>;
  mappings: CidMapping[];
  /** Parts to attach on their OWN channel. Never merged into a carried attachment set here. */
  minted: MintedInlinePart[];
  /** Content-IDs of the surviving parts a reference claimed; these ride the normal carry. */
  reusedCids: string[];
  unclaimedSurvivors: CidPart[];
  /** Counted separately from parts, never summed. */
  unresolvedRefs: string[];
  unembeddableParts: CidPart[];
  /** The denominator a shortfall reports against. */
  resolvedPartCount: number;
}

/**
 * Decide, for every reference in a body about to be quoted, which part supplies it and
 * under which Content-ID.
 *
 * REUSE of a survivor with a matching blob comes first, so an ordinary edit does not renumber
 * images a client has already rendered. Matching is ONE-TO-ONE: each survivor is claimed at
 * most once, so two images never collapse into one part.
 *
 * Otherwise the blob is carried under a fresh mint, returned on `minted` and never folded into
 * an existing attachment set, where it would be indistinguishable from a carried part.
 */
export function buildCidMap(input: BuildCidMapInput): BuildCidMapResult {
  const mint = input.mint ?? mintCid;
  const sourceParts = input.sourceParts ?? [];

  // Shared with the quotability decision so the two cannot disagree. Distinct references
  // matter: a repeat would claim a second survivor left attached with nothing pointing at it,
  // which the closure check (minted identifiers only) would not notice.
  const resolution = resolveCidRefs(input.refs ?? [], sourceParts);
  const refs = resolution.distinctRefs;
  const byCid = resolution.byRef;

  // Reusing a foreign identifier would later classify that part as server-managed.
  const survivors = (input.survivors ?? []).filter((s) => isReservedCid(s?.cid));
  const claimed = new Set<number>();

  const cidMap = new Map<string, string>();
  const mappings: CidMapping[] = [];
  const minted: MintedInlinePart[] = [];
  const reusedCids: string[] = [];

  for (const ref of refs) {
    const part = byCid.get(ref);
    // Already recorded as unresolved.
    if (!part) continue;

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

  // Derived from the resolution, so it cannot disagree with the count it is the shortfall against.
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

export type InlinePartAction =
  | 'kept'
  /** Rides the rebuilt draft as a regular attachment rather than an embedded image. */
  | 'degraded'
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
  /** Passed through untouched onto their own channel. */
  minted?: MintedInlinePart[];
  /**
   * Defaults to true. Without an html body the mail server rejects an inline disposition
   * outright, so an image can only ride as a regular attachment.
   */
  htmlShips?: boolean;
}

export interface ReconcileInlinePartsResult {
  parts: ReconciledPart[];
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
 * An unreferenced part of this server's minted shape is REMOVED: leaving it would attach a
 * file the user never asked to send. An unreferenced inline part with a foreign identifier is
 * DEGRADED instead, since dropping it would lose content this server did not create.
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
 * A self-check failure, never a caller error: a dangling authored reference is rejected and an
 * unreferenced minted part dropped before assembly, so anything reaching this check is this
 * code's own inconsistency. Callers map it to an internal error.
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
  finalPartCids?: (string | null | undefined)[];
  attachedMintedCids?: string[];
  skip?: boolean;
}

/**
 * Assert the assembled message closes over its own embedded images, both ways: every
 * reference in a body this call wrote resolves to a carried part, and every minted part is
 * referenced.
 *
 * Deliberately narrow: a broken reference in a body carried through untouched, or a caller's
 * unreferenced ordinary attachment, is not this call's loose end.
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
