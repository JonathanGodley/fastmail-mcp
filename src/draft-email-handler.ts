import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceBool, coerceAttachments, describeUntrusted, parseAddress } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs, isBlank, htmlHasVisibleContent } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import {
  buildQuoteBlocks, buildForwardBlocks, emptyQuoteImages, rejectSignatureEmbeddedImage,
  signatureBlock, signatureCidRefs,
} from './reply-quote.js';
import type { QuoteImageOutcome } from './reply-quote.js';
import { expandBodyTokens, scanBodyTokens } from './body-tokens.js';
import type {
  BlockUnavailableCause, BodyBlock, BodyBlocks, BodyTokenExpansion, BodyTokenName, BodyTokenScan,
} from './body-tokens.js';
import { matchesIdentity, selectIdentity, signatureOf } from './identity.js';
import type { ResolvedSignature } from './identity.js';
import { formatAddress } from './email-formatter.js';
import {
  planAuthoredInlineImages, recordQuoteImages, reportAuthoredInlineImages,
} from './compose-inline.js';
import type { AuthoredInlinePlan } from './compose-inline.js';
import {
  buildUnionParts, checkInlineClosure, isImageType, sanitizeQuoteHtml,
} from './inline-images.js';
import type { CidPart } from './inline-images.js';
import { CAUSE_SENTENCE, InlineNoteLedger, describePartNames, noteTokenEmpty } from './inline-notes.js';
import type { AttachmentPart, UploadAttachmentsOptions } from './jmap-client.js';
import { matchSubjectPrefix, noteComposeSubjectPrefix } from './subject-prefix.js';

// ---------------------------------------------------------------------------
// draft_email — one compose tool, three modes
// ---------------------------------------------------------------------------
//
// The caller says WHERE this server's generated blocks belong by writing `{{signature}}`,
// `{{quote}}` and `{{forward}}` into the body; nothing is added that they did not place. A
// token with nothing behind it is removed and reported, never silently dropped, and a body
// that was content before expansion and empty after it is refused rather than stored.
//
// SECURITY: expansion is body-tokens.ts's single pass, over the CALLER'S OWN AUTHORED BODY,
// before any fetched content is joined in. The quote and forwarded blocks are built from an
// attacker-authored original, so a `{{signature}}` inside it must stay inert text. Never
// re-run `expandBodyTokens` (or any other token-aware transform) over expanded output, and
// never pass it a joined body: there is exactly one call per part, on `a.textBody` /
// `a.htmlBody` verbatim.
// ---------------------------------------------------------------------------

export type DraftEmailMode = 'new' | 'reply' | 'forward';

const MODES: readonly DraftEmailMode[] = ['new', 'reply', 'forward'];

const HISTORY_TOKEN: Record<DraftEmailMode, BodyTokenName | undefined> = {
  new: undefined,
  reply: 'quote',
  forward: 'forward',
};

type PartName = 'textBody' | 'htmlBody';

// One shape for all three modes, so the assembly below has one exit.
export interface DraftEmailParams {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  from?: string;
  mailbox?: string;
  subject?: string;
  textBody?: string;
  htmlBody?: string;
  inReplyTo?: string[];
  references?: string[];
  replyTo?: string[];
  /** Forward only: the original's Message-ID, recorded as X-Forwarded-Message-Id. */
  forwardedMessageId?: string[];
  /** The exact stored instance this draft came from, for send_draft's thread marking (#60). */
  sourceEmailId?: string;
  attachments?: AttachmentPart[];
}

/** What one part's tokens did, in the order their positions ran. */
export interface TokenPartReceipt {
  part: PartName;
  /**
   * Token names in POSITION order, so a sign-off placed below the history is visible on the
   * receipt without this server judging whether that was intended.
   */
  order: BodyTokenName[];
  /** Tokens whose block was substituted, with how many times. */
  expanded: { token: BodyTokenName; count: number }[];
  /** Tokens removed because there was nothing to expand to, with the cause. */
  removed: { token: BodyTokenName; count: number; cause: BlockUnavailableCause }[];
}

/**
 * What this call did with the caller's tokens. Every entry is read off `expandBodyTokens`'
 * own per-site report, so it cannot claim an expansion that did not happen.
 */
export interface DraftEmailReceipt {
  /** One entry per supplied part that carried at least one `{{…}}` spelling. */
  parts: TokenPartReceipt[];
  /**
   * The caller's own `{{…}}` spellings this call left as written — a `{{sig}}` typo ships
   * braces otherwise, with nothing said. Bounded, because on a forward the spelling can have
   * been copied out of the original.
   */
  unexpanded?: string;
  /** True when an asAttachment forward had no body of its own and got the filler note. */
  fillerBody?: true;
}

export interface ComposeDraftEmailResult {
  emailId: string;
  mode: DraftEmailMode;
  subject?: string;
  to?: string[];
  cc?: string[];
  /**
   * The bcc actually stored — the caller's own, or the list a reply carried out of the
   * original. Reported because a blind list the caller never named would otherwise reach
   * recipients with nothing said.
   */
  bcc?: string[];
  /** What the tokens did. Absent when the call wrote no token at all. */
  tokens?: DraftEmailReceipt;
  /** What the draft embedded, could not, or was told about. Absent when there is nothing to say. */
  notes?: string[];
}

// JmapClient satisfies this structurally; declared here so the tests can pass a mock.
export interface DraftEmailClient {
  getEmailById(id: string): Promise<any>;
  getIdentities(): Promise<any[]>;
  uploadAttachments(
    specs: AttachmentSpec[],
    attachDir: string | undefined,
    allowBlobAttach: boolean,
    options?: UploadAttachmentsOptions,
  ): Promise<AttachmentPart[]>;
  createDraft(params: DraftEmailParams): Promise<string>;
}

// ---------------------------------------------------------------------------
// Refusal wording
// ---------------------------------------------------------------------------

/**
 * The clause naming the spellings THIS mode accepts, for the near-miss refusal.
 *
 * Scoped to the mode, never all three: the wrong-mode gate refuses the other mode's history
 * token in every spelling, so listing `{{forward}}` on a reply would offer a spelling this
 * same call rejects. `new` has no history token, which is why the clause carries its own verb
 * and article rather than being interpolated into a fixed plural sentence.
 */
function acceptedSpellings(mode: DraftEmailMode): string {
  const history = HISTORY_TOKEN[mode];
  return history
    ? `the exact spellings are {{signature}} and {{${history}}}`
    : 'the exact spelling is {{signature}}';
}

// Mode-agnostic on purpose: it states what the backslash does, and `{{signature}}` is a valid
// token in all three modes, so the example never offers a spelling any mode would refuse.
const ESCAPE_HINT =
  'To write braces as text, escape them: \\{{signature}} ships the literal token. ' +
  'The backslash is consumed only before a token name.';

function bad(message: string): McpError {
  return new McpError(ErrorCode.InvalidParams, message);
}

/** What a part is called in a sentence, so the two never drift. */
function partWord(part: PartName): string {
  return part === 'htmlBody' ? 'htmlBody' : 'textBody';
}

function partHasContent(part: PartName, body: string): boolean {
  return part === 'htmlBody' ? htmlHasVisibleContent(body) : !isBlank(body);
}

// ---------------------------------------------------------------------------
// Mode and mode-only parameters
// ---------------------------------------------------------------------------

/**
 * The mode, exactly as spelled. NOT coerced or defaulted (`docs/conventions.md` on
 * leniency): a defaulted mode would turn a forgotten parameter into a silently unthreaded new
 * message. Any other value reaches here only from a client that skipped schema validation.
 */
function readMode(value: unknown): DraftEmailMode {
  if (typeof value === 'string' && (MODES as readonly string[]).includes(value)) {
    return value as DraftEmailMode;
  }
  throw bad(
    `mode is required and must be exactly one of "new", "reply" or "forward" ` +
    `(got ${value === undefined ? 'nothing' : `"${describeUntrusted(value)}"`}). ` +
    'It is not defaulted: a forgotten mode would store an unthreaded new message.',
  );
}

/**
 * A parameter that belongs to one mode only. One schema serves three modes, so the schema
 * cannot refuse it and this does.
 */
function assertModeOnly(
  present: boolean, param: string, allowed: DraftEmailMode, mode: DraftEmailMode,
): void {
  if (present && mode !== allowed) {
    throw bad(`${param} applies to mode:'${allowed}' only (this call is mode:'${mode}').`);
  }
}

// ---------------------------------------------------------------------------
// Token acceptance — decided from the SCAN, before anything is built or expanded
// ---------------------------------------------------------------------------

interface PartScan {
  part: PartName;
  /** The caller's authored text, exactly as supplied. */
  authored: string;
  scan: BodyTokenScan;
}

const TOKEN_ORDER = ['signature', 'quote', 'forward'] as const;

/**
 * Refuse everything about the caller's tokens that is refusable, from the scan alone, before
 * any block is built: a refused call has fetched an original and nothing else — no blob
 * written, no draft stored.
 *
 * THE ORDER IS OVER THE WHOLE CALL, NOT PER PART: wrong-mode token, near-miss, repeat,
 * one-part, then the two forward gates, each a separate pass across every supplied part. A
 * per-part loop would let part A's near-miss beat part B's wrong-mode token, so the refusal
 * would depend on which body the caller wrote first.
 *
 * Wrong-mode leads because it also matches every near-miss spelling of a history token this
 * mode does not accept. Run the near-miss pass first and `{{{forward}}}` on a reply is
 * refused for its spelling, by a message offering `{{signature}}` and `{{quote}}` that never
 * says a reply has no forwarded block to place. Every other test here is over unescaped EXACT
 * tokens only.
 */
function assertTokensAcceptable(
  parts: PartScan[], mode: DraftEmailMode, asAttachment: boolean,
): void {
  const history = HISTORY_TOKEN[mode];

  // --- 1. A history token this mode does not accept, in any spelling -------
  // A near-miss's `name` is the token it near-missed, so one scan answers the mode question.
  for (const { scan } of parts) {
    for (const site of [...scan.tokens, ...scan.nearMisses]) {
      if (site.name === 'signature' || site.name === history) continue;
      throw bad(
        `{{${site.name}}} does not apply to mode:'${mode}'` +
        (history ? `; use {{${history}}} instead.` : ' (a new message has no history to place).'),
      );
    }
  }

  // --- 2. A near-miss spelling of a token this mode DOES accept ------------
  // REFUSED, not coerced: `docs/conventions.md` rejects a guess that would change the message,
  // and this one would put a signature block into a body that spelled something else.
  for (const { part, scan } of parts) {
    const miss = scan.nearMisses[0];
    if (miss) {
      throw bad(
        `${partWord(part)} carries "${describeUntrusted(miss.text)}", which is not a token: ` +
        `on mode:'${mode}' ${acceptedSpellings(mode)}, lower case, with two braces each side. ` +
        ESCAPE_HINT,
      );
    }
  }

  // --- 3. The same token twice in one part ---------------------------------
  // Expansion substitutes every site, so a repeat would store the block twice; a caller who
  // meant the braces as text has the escape.
  for (const { part, scan } of parts) {
    for (const name of TOKEN_ORDER) {
      if (scan.counts[name] < 2) continue;
      throw bad(
        `{{${name}}} appears ${scan.counts[name]} times in ${partWord(part)}; a token may be ` +
        'placed once per part, and expanding it twice would store the block twice. Remove the ' +
        `extra one, or escape it (\\{{${name}}}) to ship the braces as text there.`,
      );
    }
  }

  // --- 4. A token in one SUPPLIED part but not the other -------------------
  // The caller's slip. A SOURCE that has one form and not the other is a different thing,
  // reported per part as a note, not refused here.
  if (parts.length === 2) {
    const [a, b] = parts as [PartScan, PartScan];
    for (const name of TOKEN_ORDER) {
      const inA = a.scan.counts[name] > 0;
      const inB = b.scan.counts[name] > 0;
      if (inA === inB) continue;
      const has = inA ? a : b;
      const lacks = inA ? b : a;
      throw bad(
        `{{${name}}} is in ${partWord(has.part)} but not in ${partWord(lacks.part)}. ` +
        'When you supply both parts, place each token in both, or supply only one part.',
      );
    }
  }

  // --- 5. The two forward gates, last ---------------------------------------
  if (mode === 'forward' && asAttachment && parts.some((p) => p.scan.counts.forward > 0)) {
    throw bad(
      '{{forward}} does not apply to an asAttachment forward: the original rides whole ' +
      'as a .eml attachment, so there is no block to place. Drop the token, or drop ' +
      'asAttachment to forward inline.',
    );
  }

  // A forward with no {{forward}} and no .eml forwards nothing while still carrying the
  // original's attachments and marking it forwarded on send. Unlike a reply without {{quote}},
  // which is simply a message, that has no honest reading, so it is refused rather than noted.
  // Presence only: whether the block has content is step 9's question.
  if (mode === 'forward' && !asAttachment && !parts.some((p) => p.scan.counts.forward > 0)) {
    throw bad(
      'A forward must place {{forward}} in a body part, or pass asAttachment:true to send ' +
      'the original whole as a .eml. Without one of those the draft forwards nothing while ' +
      "still carrying the original's attachments. The minimal spelling is " +
      'htmlBody: "{{forward}}".',
    );
  }
}

// ---------------------------------------------------------------------------
// Blocks
// ---------------------------------------------------------------------------

const unavailable = (cause: BlockUnavailableCause): BodyBlock => ({ available: false, cause });

/**
 * A block from a builder's output: available when it produced non-blank content, otherwise
 * carrying the cause the caller is owed.
 *
 * The cause is decided against the source THIS part would quote, not a combined either-form
 * gate: an images-only original passes the combined gate and fails the text one, so a
 * text-only reply to it would otherwise store no history and say nothing at all.
 */
function quoteBlock(content: string | undefined, anyForm: boolean): BodyBlock {
  if (content !== undefined && !isBlank(content)) return { available: true, content };
  return unavailable(anyForm ? 'nothing-quotable-in-this-form' : 'nothing-quotable');
}

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

/** The identity has a signature and the caller placed none. */
function noteSignatureNotPlaced(identityEmail: string | undefined): string {
  return (
    `Identity ${identityEmail ? `${describeUntrusted(identityEmail)} ` : ''}has a signature; ` +
    'this body has no {{signature}}, so it was stored as written. To add one on edit_draft, ' +
    'place the token and pass expandSignature: true.'
  );
}

/**
 * A `{{forward}}` whose block ships only in the TEXT form, over an original that really has
 * html to reproduce.
 *
 * A note of its own rather than a line on the image sentence, because the loss is the
 * FORMATTING first: a formatted message forwarded as plain text is degraded even with no
 * images at all.
 */
function noteForwardTextForm(imagesRode: boolean): string {
  return (
    '{{forward}} ships in the text form only and the original ships HTML, so this forward loses ' +
    (imagesRode
      ? 'its formatting and its inline images ride as attachments; put {{forward}} in ' +
        'htmlBody to keep both.'
      : 'its formatting; put {{forward}} in htmlBody to keep it.')
  );
}

/**
 * What the pooled-media sentence ends on. Every forward that pools a part has `{{forward}}`
 * in its body, so re-running as .eml means dropping the token first. Either fix makes a new
 * draft, so both say to delete this one.
 */
const POOLED_REMEDY_PLACE_IN_HTML =
  'put {{forward}} in htmlBody to embed them, or drop the token and pass asAttachment: true ' +
  'to forward the original whole, then delete this draft.';
const POOLED_REMEDY_DROP_TOKEN =
  'drop {{forward}} and pass asAttachment: true to forward the original whole, then delete ' +
  'this draft.';

/** A reply that placed no {{quote}}: a forgotten token would otherwise be silent. */
const NOTE_REPLY_UNQUOTED =
  'This reply was stored without the original: place {{quote}} in the body to include it.';

/**
 * A reply that carried the original's own Bcc list into its bcc. Said out loud because a
 * blind list is invisible in the composed draft, the one recipient field a caller cannot see
 * it has widened.
 */
const NOTE_BCC_CARRIED =
  "The original's Bcc list was carried into this reply — pass bcc to replace it, or to or " +
  'cc to reply to fewer people and turn the carry off.';

/** An image the block minted that no part of the expanded body ends up referencing. */
function noteMintedDropped(names: (string | null | undefined)[], total: number): string {
  const listed = describePartNames(names, total);
  return (
    `${total} image(s) the quoted original displayed ${listed ? `(${listed}) ` : ''}` +
    'were dropped: after expansion no body written by this call references them. ' +
    'A token placed inside a comment or an attribute is the usual cause.'
  );
}

/** The forward counterpart of noteMintedDropped. */
function noteForwardUnreferenced(
  names: (string | null | undefined)[], total: number, carried: boolean,
): string {
  const listed = describePartNames(names, total);
  return (
    `${total} image(s) the forwarded original displayed ${listed ? `(${listed}) ` : ''}` +
    (carried ? 'ride as regular attachments' : 'were left out') +
    ': after expansion no body written by this call references them' +
    (carried ? '' : ', and includeOriginalAttachments is false') +
    '. A token placed inside a comment or an attribute is the usual cause.'
  );
}

// ---------------------------------------------------------------------------
// Forward-mode hygiene on values taken from the forwarded message
// ---------------------------------------------------------------------------

// Fastmail validates header:…:asMessageIds values on Email/set (probed live
// 2026-07-05): embedded CR/LF and non-ASCII are REJECTED (failing the whole create),
// and embedded angle brackets round-trip MANGLED (split into two ids). The value
// comes verbatim from the forwarded — attacker-controlled — message, so pre-vet it
// and treat a malformed id as absent: the forward still works, and only the
// recorded-source affordance is lost.
//
// `edit_draft` carries the stored value (or drops it via clearFields) and does NOT re-vet
// it, though a draft composed in another client was never vetted here: a value Fastmail
// rejects fails the recreate's CREATE loudly with the old draft intact (it creates before it
// disposes), so the caller gets an error and one recovery — clear the marking — rather than
// a silent drop.
export function isSettableMessageId(id: unknown): id is string {
  return (
    typeof id === 'string' &&
    id.length > 0 &&
    id.length <= 998 && // RFC 5322 line limit; anything longer is garbage, not an id
    /^[\x21-\x7e]+$/.test(id) && // printable ASCII only — no spaces, controls, or non-ASCII
    !id.includes('<') &&
    !id.includes('>')
  );
}

// Filename for the attached .eml, from the ORIGINAL's subject (a caller subject override
// does not affect it). Receiving clients use it as a SAVE NAME, so strip control chars (\p{Cc} covers C0 and C1 incl. U+0085), Unicode
// format/bidi controls (\p{Cf}, e.g. U+202E right-to-left override), path
// separators and the Windows drive/ADS colon, and leading dots; cap the length.
// Windows reserved device names (CON, NUL, …) deliberately survive as e.g.
// "CON.eml" — a save-time nuisance the receiving client handles, same
// receiver-sanitizes posture as carried attachment names (docs/security-model.md).
export function sanitizeEmlFilename(subject: string | null | undefined): string {
  const stripped = (subject ?? '')
    .replace(/[\p{Cc}\p{Cf}]/gu, '')
    .replace(/[/\\:]/g, '_')
    .replace(/^\.+/, '')
    .trim();
  // Cap by CODE POINTS, not UTF-16 units — a unit slice can split a surrogate pair
  // and leave a lone surrogate, invalid on the wire.
  const cleaned = [...stripped].slice(0, 80).join('').trim();
  return `${cleaned || 'forwarded-message'}.eml`;
}

// ---------------------------------------------------------------------------
// Reply recipients
// ---------------------------------------------------------------------------

/** One of a fetched message's address headers, reduced to the entries that name an address. */
function addressList(value: unknown): { name?: string; email: string }[] {
  if (!Array.isArray(value)) return [];
  return value.filter(
    (x: any): x is { name?: string; email: string } => typeof x?.email === 'string' && x.email !== '',
  );
}

/**
 * The reply-all cc: everyone the original's To and CC named, in that order, minus the
 * addresses this reply is already going to and minus every address the account sends as.
 *
 * formatAddress, never coerceStringArray (#31): a display name carrying a comma would
 * otherwise re-split into a bogus second recipient.
 *
 * Exclusions and dedupe use the case-folded ADDRESS alone — "D. Fox" in To and "Dana" in CC
 * is one recipient — while the whole formatted address is kept. Self is tested with
 * matchesIdentity, so a wildcard identity (`*@example.com`) excludes its whole domain.
 */
function replyAllCc(
  original: any, addressed: { email: string }[], identities: any[],
): string[] {
  const seen = new Set(addressed.map((a) => a.email.toLowerCase()));
  const out: string[] = [];
  for (const entry of [...addressList(original.to), ...addressList(original.cc)]) {
    const key = entry.email.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const isSelf = identities.some(
      (id: any) => typeof id?.email === 'string' && matchesIdentity(id.email, entry.email),
    );
    if (isSelf) continue;
    out.push(formatAddress(entry));
  }
  return out;
}

/**
 * The bcc a reply carries: the original's own Bcc entries, whole and in order, because that
 * is what pressing Reply in Fastmail's mobile app produces (measured 2026-09-10;
 * docs/fastmail-action-availability.md, "what a reply prefills"). So NOTHING is excluded:
 * not the account's own identities, and not an address the reply's to or cc already names.
 *
 * Deliberately no "is this the account's own message" check: a received message carries no
 * Bcc header (the submitting server strips it), so presence already marks the account's own
 * copy. That is derived, not measured, and such a check could only silently refuse an
 * imported or malformed message, whose shape is recorded as unmeasured in the same doc.
 *
 * formatAddress and a case-folded ADDRESS dedupe, as in replyAllCc (#31).
 */
function replyBcc(original: any): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const entry of addressList(original.bcc)) {
    const key = entry.email.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(formatAddress(entry));
  }
  return out;
}

// ---------------------------------------------------------------------------
// The orchestration
// ---------------------------------------------------------------------------

/**
 * Orchestrate draft_email end to end.
 *
 * Reads no environment: the caller resolves `attachDir` and `allowBlobAttach`. Only ever
 * drafts — send_draft transmits, and does the thread-state maintenance from the provenance
 * headers recorded here (#60).
 *
 * The order of the numbered steps is load-bearing:
 *   1  mode and mode-only parameters      — cheapest refusals first
 *   2  assertBodyInputs + contentless      — VALIDATE, before any expansion
 *   3  fetch the original                  — the only I/O before a refusal is possible
 *   4  scan and refuse                     — every token refusal, from the scan alone
 *   5  identity, once                      — the signature and its warning
 *   6  planAuthoredInlineImages            — on the PRE-expansion html
 *   7  build blocks                        — only for a token that is actually present
 *   8  expandBodyTokens                    — THE single pass, per part
 *   9  empty-after-expansion refusal       — per part, before the filler
 *  10  upload, carry, assemble
 *  11  checkInlineClosure                  — on the POST-expansion html
 *  12  createDraft, then the receipt
 *
 * Steps 6 and 11 straddle step 8 deliberately (a unit test pins it): the plan reads the
 * caller's OWN html, since the blocks legitimately carry identifiers the caller never wrote,
 * while the closure check reads what actually ships. Reversed, every image-bearing reply is
 * refused.
 */
export async function composeDraftEmail(
  args: any,
  client: DraftEmailClient,
  attachDir: string | undefined,
  allowBlobAttach: boolean,
): Promise<ComposeDraftEmailResult> {
  const a = args ?? {};

  // --- 1. Mode, and the parameters that belong to one mode -----------------
  const mode = readMode(a.mode);
  // Read off `args?.`, NOT the `a` alias: tool-schema.test.ts's lenient-boolean guard matches
  // `!!args?.asAttachment` but not `!!a.asAttachment`, so the alias would hide a future
  // bare-`!!` read of these flags from the only check that looks for one.
  const asAttachment = coerceBool(args?.asAttachment) ?? false;
  const includeOriginalAttachments = coerceBool(args?.includeOriginalAttachments) ?? true;

  // `!= null`, not `!== undefined`: a lenient client sends null for every declared key, and
  // null reads as absent (docs/conventions.md).
  assertModeOnly(args?.asAttachment != null, 'asAttachment', 'forward', mode);
  assertModeOnly(
    args?.includeOriginalAttachments != null, 'includeOriginalAttachments', 'forward', mode,
  );
  assertModeOnly(a.mailbox != null, 'mailbox', 'new', mode);
  assertModeOnly(a.inReplyTo != null, 'inReplyTo', 'new', mode);
  assertModeOnly(a.references != null, 'references', 'new', mode);

  const originalEmailId = a.originalEmailId;
  if (mode === 'new') {
    if (originalEmailId != null) {
      throw bad("originalEmailId applies to mode:'reply' and mode:'forward' only.");
    }
  } else if (!originalEmailId) {
    throw bad(`originalEmailId is required for mode:'${mode}'.`);
  }

  // --- 2. Validate the caller's own bodies, BEFORE anything is expanded ----
  // The blocks would otherwise mask a malformed body, supplying the real tags an
  // escaped-markup body lacks and the visible content an empty one lacks, and a forwarded
  // original containing `<![CDATA[` would refuse the call (#78, docs/email-bodies.md).
  assertBodyInputs(a);

  const { from, subject: rawSubject, textBody, htmlBody } = a;
  // `from` may carry a display name (#161). selectIdentity's wildcard match accepts a bare
  // addr-spec only, so the identity lookup and the not-placed note take the ADDRESS half, or
  // a named `from` would lose its signature. `from` itself goes to createDraft RAW, which is
  // what puts the caller's name into the stored From header.
  const fromAddress: string | undefined = from ? parseAddress(from).email : from;
  const { to: toArg, cc, bcc, replyTo } = coerceRecipients(a);
  // Coerced before the contentless guard, so an attachment-only stash counts as content; a
  // lenient client may send a JSON-string array, so the guard tests the coerced specs.
  const specs = coerceAttachments(a.attachments);

  const subjectOverride = coerceSubjectOverride(
    rawSubject,
    mode === 'new'
      ? 'Omit it for a subject-less draft.'
      : `Omit it to inherit "${mode === 'reply' ? 'Re' : 'Fwd'}: <original subject>".`,
  );

  if (mode === 'new') {
    // isBlank, as createDraft's copy of this guard does: truthiness would let a
    // whitespace-only body through to its generic message, which names no parameter.
    if (!toArg?.length && !subjectOverride && isBlank(textBody) && isBlank(htmlBody)
        && !specs?.length) {
      throw bad('At least one of to, subject, textBody, htmlBody, or attachments must be provided');
    }
  }

  // --- 3. Fetch the original (reply and forward only) ----------------------
  const original = mode === 'new' ? undefined : await client.getEmailById(originalEmailId);

  // --- 4. Scan the supplied parts, and refuse from the scan alone ----------
  const supplied: PartScan[] = [];
  if (typeof textBody === 'string') {
    supplied.push({ part: 'textBody', authored: textBody, scan: scanBodyTokens(textBody) });
  }
  if (typeof htmlBody === 'string') {
    supplied.push({ part: 'htmlBody', authored: htmlBody, scan: scanBodyTokens(htmlBody) });
  }
  assertTokensAcceptable(supplied, mode, asAttachment);

  const history = HISTORY_TOKEN[mode];
  const htmlPart = supplied.find((p) => p.part === 'htmlBody');
  const textPart = supplied.find((p) => p.part === 'textBody');

  // Does this message ship an html part at all? A present-but-blank htmlBody emits none
  // (buildBodyParts drops it), so it must not carry the only sign-off.
  const messageShipsHtml = !isBlank(htmlBody);
  // Kept apart from it: the html part ships AND carries the history token (the builders add
  // that the original has quotable html). It decides whether the quoted images are minted,
  // and so whether the text alternative may describe an image the reader can look at.
  const historyHtmlShips =
    messageShipsHtml && !!history && (htmlPart?.scan.counts[history] ?? 0) > 0;
  const historyPlaced = !!history && supplied.some((p) => p.scan.counts[history] > 0);

  // --- 5. The identity, fetched once ---------------------------------------
  // Always fetched: the not-placed note needs to know whether the identity HAS a signature
  // even when no token was placed, and the reply-all cc excludes every identity, not just the
  // selected one. The `?? []` is unpinnable (any non-empty stand-in matches nothing at either
  // consumer); what the test beside it pins is that a client returning no list does not throw
  // the compose away.
  const identities = (await client.getIdentities()) ?? [];
  const identity = selectIdentity(identities, fromAddress);
  const signature = signatureOf(identity);

  // --- 6. The caller's embedded images, read PRE-expansion -----------------
  const inlinePlan: AuthoredInlinePlan = planAuthoredInlineImages({
    callerHtml: htmlBody,
    htmlShips: messageShipsHtml,
    specs,
    attachmentsEnabled: !!attachDir || allowBlobAttach,
    surface: mode === 'new' ? 'compose' : 'note',
  });

  // --- 7. Build the blocks, only for a token that is actually there --------
  // Block construction resolves the original's image references, so building an unplaced
  // block would give an unquoted reply "dropped image" notes.
  let quoteImages: QuoteImageOutcome = emptyQuoteImages();
  const htmlBlocks: BodyBlocks = { signature: undefined, quote: undefined, forward: undefined };
  const textBlocks: BodyBlocks = { signature: undefined, quote: undefined, forward: undefined };

  if (supplied.some((p) => p.scan.counts.signature > 0)) {
    htmlBlocks.signature = signatureBlock(signature, 'htmlBody', messageShipsHtml);
    textBlocks.signature = signatureBlock(signature, 'textBody', messageShipsHtml);
  }

  if (historyPlaced && mode === 'reply') {
    // No `timezone`: this tool takes none, so the attribution line uses the server's zone.
    const built = buildQuoteBlocks({
      original,
      htmlShips: historyHtmlShips,
      quoteImages: { sourceParts: buildUnionParts(original).map((u) => u.part) },
    });
    quoteImages = built.images;
    const anyForm = built.textBlock !== undefined || built.htmlBlock !== undefined;
    htmlBlocks.quote = quoteBlock(built.htmlBlock, anyForm);
    textBlocks.quote = quoteBlock(built.textBlock, anyForm);
  }

  let forwardSourceParts: { part: CidPart; inBodyList: boolean }[] = [];
  // The noteForwardTextForm case. Read off the builder's own `htmlQuotable`, so it cannot
  // disagree with the arm that chose the form.
  let forwardTextFormOnly = false;
  if (historyPlaced && mode === 'forward') {
    forwardSourceParts = buildUnionParts(original).filter((u) => u.part?.blobId);
    const built = buildForwardBlocks({
      original,
      htmlShips: historyHtmlShips,
      quoteImages: { sourceParts: forwardSourceParts.map((u) => u.part) },
    });
    quoteImages = built.images;
    forwardTextFormOnly = !historyHtmlShips && built.htmlQuotable;
    // Always `true`: the forwarded block has a header block, so it is never "nothing
    // quotable" in any form, only in the form of the part the token was placed in.
    htmlBlocks.forward = quoteBlock(built.htmlBlock, true);
    textBlocks.forward = quoteBlock(built.textBlock, true);
  }

  // --- 8. THE SINGLE PASS, per part, over the caller's authored body -------
  // Read the security rule at the top of this file before touching these two lines.
  const expansions = new Map<PartName, BodyTokenExpansion>();
  if (htmlPart) expansions.set('htmlBody', expandBodyTokens(htmlPart.authored, htmlBlocks));
  if (textPart) expansions.set('textBody', expandBodyTokens(textPart.authored, textBlocks));

  const expandedHtml = expansions.get('htmlBody')?.text ?? (htmlPart ? '' : undefined);
  const expandedText = expansions.get('textBody')?.text ?? (textPart ? '' : undefined);

  // --- 9. A part that was content before expansion and is empty after it ---
  // Refused per part, in every mode, winning over the empty-token note: a forward whose
  // {{forward}} has nothing quotable passes step 4's presence gate and would otherwise store
  // a body-less forward that still carries attachments and marks the original forwarded.
  // Raised here, not left to createDraft's generic message, so it can name causes and fixes.
  for (const { part, authored } of supplied) {
    const before = partHasContent(part, authored);
    if (!before) continue;
    const after = part === 'htmlBody'
      ? htmlHasVisibleContent(expandedHtml ?? '')
      : !isBlank(expandedText ?? '');
    if (after) continue;
    const causes = [...new Set(
      (expansions.get(part)?.tokens ?? [])
        .filter((t) => !t.expanded && t.cause)
        .map((t) => CAUSE_SENTENCE[t.cause!]),
    )];
    throw bad(
      `${partWord(part)} is empty after expansion: it was nothing but tokens, and ` +
      `${causes.length ? causes.join('; ') : 'the block had no content for this part'}. ` +
      'Write prose beside the token — the block skips and the result says so — or, on a ' +
      'forward, drop {{forward}} and pass asAttachment:true.',
    );
  }

  // A sign-off displaying an image no part carries is refused before anything is uploaded.
  // Tested on the EXPANDED html, so a token the markup hides displays nothing to refuse, and
  // an attachments item supplying the identifier resolves the reference like any other.
  const suppliedCids = new Set((specs ?? []).map((s) => s.cid).filter((c) => !!c));
  const liveAfterExpansion = new Set(expandedHtml ? extractLiveCidRefs(expandedHtml) : []);
  if (htmlBlocks.signature?.available === true) {
    for (const ref of signatureCidRefs(signature)) {
      if (liveAfterExpansion.has(ref) && !suppliedCids.has(ref)) {
        throw bad(rejectSignatureEmbeddedImage(ref));
      }
    }
  }

  // --- 10. Assemble ---------------------------------------------------------
  const params: DraftEmailParams = { from, replyTo };
  if (toArg?.length) params.to = toArg;
  // On LENGTH: coerceRecipients returns [] for '' and [], and a truthy [] would reach the
  // result.
  if (cc?.length) params.cc = cc;
  if (bcc?.length) params.bcc = bcc;

  let fillerBody: true | undefined;
  let bccCarried = false;
  params.textBody = expandedText;
  params.htmlBody = expandedHtml;

  if (mode === 'new') {
    if (a.mailbox != null) params.mailbox = a.mailbox;
    params.inReplyTo = coerceStringArray(a.inReplyTo);
    params.references = coerceStringArray(a.references);
    params.subject = subjectOverride;
  } else {
    if (typeof original?.id === 'string' && original.id !== '') params.sourceEmailId = original.id;
  }

  if (mode === 'reply') {
    const originalMessageId = original?.messageId?.[0];
    if (!originalMessageId) {
      throw new McpError(
        ErrorCode.InternalError,
        'Original email does not have a Message-ID; cannot thread reply',
      );
    }
    params.inReplyTo = [originalMessageId];
    params.references = [...(original.references || []), originalMessageId];

    let subject = subjectOverride ?? (original.subject || '');
    if (subjectOverride === undefined && !/^Re:/i.test(subject)) subject = `Re: ${subject}`;
    params.subject = subject;

    // Reply-To if the original named one, else From, via formatAddress and never
    // coerceStringArray (#31). The address objects are kept so the cc carry can exclude them
    // without re-parsing a formatted string. This branch IS the "caller named no `to`" case,
    // which is why the carry is nested inside it rather than re-testing `toArg`.
    if (!params.to?.length) {
      const replyToHeader = addressList(original.replyTo);
      const addressed = replyToHeader.length ? replyToHeader : addressList(original.from);
      params.to = addressed.map(formatAddress);
      if (!params.to.length) {
        throw bad('Could not determine reply recipient. Please provide "to" explicitly.');
      }

      // Reply-all by default (#184), ONLY when the caller named no `to` or `cc`: either is a
      // deliberately narrowed reply this server must not widen. An explicit `cc` suppresses
      // the CARRY alone; the `to` fallback above still runs.
      if (!cc?.length) {
        const carried = replyAllCc(original, addressed, identities);
        if (carried.length) params.cc = carried;

        // The original's Bcc list, under the same condition (#189). A caller `bcc` is
        // additive rather than narrowing, so it displaces only this carry; an EMPTY one is no
        // bcc and displaces nothing, as with an empty `cc`.
        if (!bcc?.length) {
          const carriedBcc = replyBcc(original);
          if (carriedBcc.length) {
            params.bcc = carriedBcc;
            bccCarried = true;
          }
        }
      }
    }
  }

  if (mode === 'forward') {
    if (!params.to?.length) {
      throw bad('to is required for a forward; there is no default recipient');
    }
    if (subjectOverride !== undefined) {
      params.subject = subjectOverride;
    } else {
      const orig = original?.subject || '';
      params.subject = /^fwd?:/i.test(orig.trim()) ? orig : `Fwd: ${orig}`;
    }
    // Recorded on BOTH forward shapes: send_draft resolves it to mark the original
    // forwarded on transmit, and the attached .eml is not machine-resolvable as provenance.
    const originalMessageId = original?.messageId?.[0];
    if (isSettableMessageId(originalMessageId)) params.forwardedMessageId = [originalMessageId];
  }

  // A minted part nothing references AFTER expansion (a token inside a comment or an
  // attribute) is dropped here, before anything is recorded, so it is never also reported as
  // embedded; a forward then treats it like any image it cannot embed.
  const liveRefs = new Set(
    expandedHtml ? extractLiveCidRefs(expandedHtml) : [],
  );
  const droppedMinted = quoteImages.minted.filter((p) => !liveRefs.has(p.cid));
  const droppedSources = new Set(
    quoteImages.mappings.filter((m) => !liveRefs.has(m.cid)).map((m) => m.source),
  );
  const forwardReferenced = new Set(quoteImages.resolvedParts);
  if (droppedMinted.length > 0) {
    quoteImages = {
      ...quoteImages,
      minted: quoteImages.minted.filter((p) => liveRefs.has(p.cid)),
      mappings: quoteImages.mappings.filter((m) => liveRefs.has(m.cid)),
      resolvedParts: quoteImages.resolvedParts.filter((p) => !droppedSources.has(p)),
    };
  }

  const carried: AttachmentPart[] = [];
  const pooled: CidPart[] = [];
  const droppedCarried: CidPart[] = [];
  const droppedExcluded: CidPart[] = [];
  const attachedFiles: CidPart[] = [];
  const notIncluded: CidPart[] = [];

  if (mode === 'forward' && asAttachment) {
    // Lossless form: the Email's own blobId is the raw RFC 5322 message. The filler goes in
    // AFTER step 9, so a supplied part is tested on its own terms. It is ordinary prose that
    // nothing reads back; an edit replaces it like any other body.
    if (isBlank(params.textBody) && isBlank(params.htmlBody)) {
      params.textBody = 'Forwarded message attached.';
      params.htmlBody = undefined;
      fillerBody = true;
    }
    if (!original?.blobId) {
      throw new McpError(
        ErrorCode.InternalError, 'Original email has no blobId; cannot attach it as .eml',
      );
    }
    carried.push({
      blobId: original.blobId,
      type: 'message/rfc822',
      name: sanitizeEmlFilename(original?.subject),
      disposition: 'attachment',
    });
  } else if (mode === 'forward') {
    // An image the forwarded block displays is BODY CONTENT and is carried whatever
    // includeOriginalAttachments says; the flag governs the original's FILES.
    const embedded = new Set(quoteImages.mappings.map((m) => m.source));
    for (const entry of forwardSourceParts) {
      if (embedded.has(entry.part)) continue;
      if (!includeOriginalAttachments) {
        (droppedSources.has(entry.part) ? droppedExcluded : notIncluded).push(entry.part);
        continue;
      }
      const part: AttachmentPart = { blobId: entry.part.blobId!, type: entry.part.type! };
      if (entry.part.name != null) part.name = entry.part.name;
      if ((entry.part as any).disposition != null) {
        part.disposition = (entry.part as any).disposition === 'inline'
          ? 'attachment'
          : (entry.part as any).disposition;
      }
      carried.push(part);
      if (droppedSources.has(entry.part)) {
        droppedCarried.push(entry.part);
        continue;
      }
      const bodyMedia = entry.inBodyList
        || ((entry.part as any)?.disposition === 'inline' && !!entry.part?.cid);
      if (forwardReferenced.has(entry.part) || bodyMedia) pooled.push(entry.part);
      else attachedFiles.push(entry.part);
    }
  }

  const ledger = new InlineNoteLedger();
  const carry = recordQuoteImages(ledger, quoteImages, mode === 'forward' ? 'forward' : 'reply');

  const uploaded = specs?.length
    ? await client.uploadAttachments(specs, attachDir, allowBlobAttach, {
      inlineCids: inlinePlan.inlineCids,
    })
    : undefined;

  const attachments = [...carried, ...(uploaded ?? []), ...carry.minted];
  if (attachments.length > 0) params.attachments = attachments;

  // --- 11. Closure, on what actually ships ---------------------------------
  checkInlineClosure({
    htmlBodies: [params.htmlBody],
    finalPartCids: attachments.map((part) => part.cid),
    attachedMintedCids: carry.minted.map((part) => part.cid).filter((c): c is string => !!c),
  });

  // --- 12. Store, then report ----------------------------------------------
  const emailId = await client.createDraft(params);

  pooled.forEach((part, i) => {
    ledger.record({ key: `pool:${i}`, outcome: 'pooled', name: part.name, isImage: isImageType(part.type) });
  });
  attachedFiles.forEach((part, i) => {
    ledger.record({ key: `carry:${i}`, outcome: 'attached', name: part.name });
  });
  notIncluded.forEach((part, i) => {
    ledger.record({ key: `excluded:${i}`, outcome: 'notIncluded', name: part.name, isImage: isImageType(part.type) });
  });

  const receipt = buildReceipt(expansions, fillerBody);
  const signaturePlaced = supplied.some((p) => p.scan.counts.signature > 0);

  // A "Re:" or "Fwd:" typed into a mode:'new' subject does not thread the message (#188), so
  // it reads as part of a conversation and arrives as a new one. Noted, never refused:
  // reusing an old subject for a fresh conversation is legitimate.
  //
  // Silent when the caller's own threading headers make the draft thread (the documented
  // route for replying to a message this account does not hold). Tested on what they
  // COERCED to, not whether they were mentioned: `inReplyTo: []` writes no header.
  const prefixTyped = mode === 'new' && !params.inReplyTo?.length && !params.references?.length
    ? matchSubjectPrefix(params.subject)
    : undefined;

  const notes = [
    ...ledger.emit({
      surface: mode === 'forward' ? 'forward' : 'reply',
      ...(carry.resolvedPartCount !== undefined && { resolvedPartCount: carry.resolvedPartCount }),
      pooledRemedy: forwardTextFormOnly ? POOLED_REMEDY_PLACE_IN_HTML : POOLED_REMEDY_DROP_TOKEN,
    }),
    ...(mode === 'reply' && droppedMinted.length > 0
      ? [noteMintedDropped(droppedMinted.map((p) => p.name), droppedMinted.length)]
      : []),
    ...(droppedCarried.length > 0
      ? [noteForwardUnreferenced(droppedCarried.map((p) => p.name), droppedCarried.length, true)]
      : []),
    ...(droppedExcluded.length > 0
      ? [noteForwardUnreferenced(droppedExcluded.map((p) => p.name), droppedExcluded.length, false)]
      : []),
    ...await reportAuthoredInlineImages({
      uploaded,
      mintedCids: carry.minted.map((p) => p.cid).filter((c): c is string => !!c),
      plan: inlinePlan,
      emailId,
      readBack: (id) => client.getEmailById(id),
    }),
    ...emptyTokenNotes(expansions),
    ...(forwardTextFormOnly ? [noteForwardTextForm(pooled.some((p) => isImageType(p.type)))] : []),
    // Presence on the PRE-expansion scan of a SUPPLIED body with content, so it cannot
    // false-fire and an attachment-only stash, a body-less reply (a blank or visually empty
    // part included) and an asAttachment filler are silent. It fires on every deliberately
    // unsigned message: the accepted cost of never storing an unsigned body with nothing said.
    ...(!signaturePlaced && signature && supplied.some((p) => partHasContent(p.part, p.authored))
      ? [noteSignatureNotPlaced(identity?.email ?? fromAddress)]
      : []),
    ...(mode === 'reply' && !historyPlaced ? [NOTE_REPLY_UNQUOTED] : []),
    ...(bccCarried ? [NOTE_BCC_CARRIED] : []),
    ...(prefixTyped ? [noteComposeSubjectPrefix(prefixTyped)] : []),
  ];

  return {
    emailId,
    mode,
    ...(params.subject !== undefined && { subject: params.subject }),
    ...(params.to && { to: params.to }),
    // cc and bcc are read off `params`, not the caller's arguments, so a carried list is
    // reported rather than the draft quietly going to more people than the result names.
    ...(params.cc && { cc: params.cc }),
    ...(params.bcc && { bcc: params.bcc }),
    ...(receipt && { tokens: receipt }),
    ...(notes.length > 0 && { notes }),
  };
}

/**
 * Content-IDs an html body really references, read with THE SAME collector the closure check
 * uses, so the minted-part drop can never disagree with the throw it exists to prevent.
 */
function extractLiveCidRefs(html: string): string[] {
  return sanitizeQuoteHtml(html, { mode: 'collect' }).refs;
}

/** One note per token that was placed and had nothing to expand to, per part. */
function emptyTokenNotes(expansions: Map<PartName, BodyTokenExpansion>): string[] {
  const out: string[] = [];
  for (const [part, expansion] of expansions) {
    const seen = new Set<string>();
    for (const site of expansion.tokens) {
      if (site.expanded || !site.cause) continue;
      const key = `${site.name}:${site.cause}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(noteTokenEmpty(site.name, partWord(part), site.cause));
    }
  }
  return out;
}

/** The receipt; absent when the call wrote no `{{…}}` spelling and produced no filler. */
function buildReceipt(
  expansions: Map<PartName, BodyTokenExpansion>, fillerBody: true | undefined,
): DraftEmailReceipt | undefined {
  const parts: TokenPartReceipt[] = [];
  // DISTINCT SPELLINGS, gathered across every part rather than per occurrence: this is a
  // token-spelling caller of `describePartNames`, and that is the contract it is under.
  const unexpandedSpellings: string[] = [];

  for (const [part, expansion] of expansions) {
    for (const s of expansion.otherSpellings) {
      if (!unexpandedSpellings.includes(s.text)) unexpandedSpellings.push(s.text);
    }
    if (expansion.tokens.length === 0) continue;
    const expanded = new Map<BodyTokenName, number>();
    const removed = new Map<string, { token: BodyTokenName; count: number; cause: BlockUnavailableCause }>();
    for (const site of expansion.tokens) {
      if (site.expanded) {
        expanded.set(site.name, (expanded.get(site.name) ?? 0) + 1);
      } else if (site.cause) {
        const key = `${site.name}:${site.cause}`;
        const row = removed.get(key);
        if (row) row.count++;
        else removed.set(key, { token: site.name, count: 1, cause: site.cause });
      }
    }
    parts.push({
      part,
      // expandBodyTokens reports sites in body order; never re-infer it from the text.
      order: expansion.tokens.map((t) => t.name),
      expanded: [...expanded.entries()].map(([token, count]) => ({ token, count })),
      removed: [...removed.values()],
    });
  }

  if (parts.length === 0 && unexpandedSpellings.length === 0 && !fillerBody) return undefined;
  const listed = describePartNames(unexpandedSpellings);
  return {
    parts,
    ...(listed && { unexpanded: listed }),
    ...(fillerBody && { fillerBody }),
  };
}
