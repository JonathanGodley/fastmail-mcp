import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { coerceRecipients, coerceStringArray, coerceBool, coerceAttachments, describeUntrusted } from './coerce.js';
import type { AttachmentSpec } from './coerce.js';
import { assertBodyInputs, isBlank, htmlHasVisibleContent } from './body-format.js';
import { coerceSubjectOverride } from './subject.js';
import {
  buildQuoteBlocks, buildForwardBlocks, emptyQuoteImages, signatureBlock,
} from './reply-quote.js';
import type { QuoteImageOutcome } from './reply-quote.js';
import { expandBodyTokens, scanBodyTokens } from './body-tokens.js';
import type {
  BlockUnavailableCause, BodyBlock, BodyBlocks, BodyTokenExpansion, BodyTokenName, BodyTokenScan,
} from './body-tokens.js';
import { selectIdentity, signatureOf } from './identity.js';
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

// ---------------------------------------------------------------------------
// draft_email — one compose tool, three modes
// ---------------------------------------------------------------------------
//
// The caller says WHERE this server's generated blocks belong by writing `{{signature}}`,
// `{{quote}}` and `{{forward}}` into the body; nothing is added that they did not place.
// That is the whole difference from the three tools this supersedes, and every rule below
// falls out of it: the mode decides which history token is even offered, a token with
// nothing behind it is removed and reported rather than silently dropped, and a body that
// was content before expansion and empty after it is refused rather than stored.
//
// ---------------------------------------------------------------------------
// THE SECURITY RULE, restated here because this is the file that could break it.
//
// Expansion is a SINGLE PASS over the CALLER'S OWN AUTHORED BODY, run BEFORE any fetched
// content is joined into it. The quote and forwarded blocks are built from an
// attacker-authored original: a `{{signature}}` sitting inside a stranger's email must land
// on the REPLACEMENT side of the substitution, where `String.prototype.replace` never
// rescans it, and stay inert as literal text.
//
// So: never re-run `expandBodyTokens` — or any other token-aware transform — over expanded
// output, and never pass a joined body to it. There is exactly one `expandBodyTokens` call
// per part in this file and it takes `a.textBody` / `a.htmlBody` verbatim.
//
// Validation runs BEFORE expansion for its own reason (`assertBodyInputs` plus the
// contentless guard): the other order lets a forwarded original that happens to contain
// `<![CDATA[` refuse the whole call, and lets the quote's real tags mask an escaped-markup
// body — the #78 defect.
// ---------------------------------------------------------------------------

/** The three modes. `mode` is required and is never coerced or defaulted. */
export type DraftEmailMode = 'new' | 'reply' | 'forward';

const MODES: readonly DraftEmailMode[] = ['new', 'reply', 'forward'];

/** Which history token each mode offers. `new` offers none. */
const HISTORY_TOKEN: Record<DraftEmailMode, BodyTokenName | undefined> = {
  new: undefined,
  reply: 'quote',
  forward: 'forward',
};

/** The two body parts, named as the caller names them. */
type PartName = 'textBody' | 'htmlBody';

// Parameters passed to createDraft. One shape for all three modes — the mode decides which
// fields are populated, not which type is used, so the assembly below has one exit.
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
   * receipt without this server judging whether that was intended. Read off the token
   * indices, never re-inferred from the body.
   */
  order: BodyTokenName[];
  /** Tokens whose block was substituted, with how many times. */
  expanded: { token: BodyTokenName; count: number }[];
  /** Tokens removed because there was nothing to expand to, with the cause. */
  removed: { token: BodyTokenName; count: number; cause: BlockUnavailableCause }[];
}

/**
 * What this call did with the caller's tokens.
 *
 * It cannot claim an expansion that did not happen: every entry is read off
 * `expandBodyTokens`' own per-site report, which records what the single pass actually
 * substituted.
 */
export interface DraftEmailReceipt {
  /** One entry per supplied part that carried at least one `{{…}}` spelling. */
  parts: TokenPartReceipt[];
  /**
   * The caller's own `{{…}}` spellings this call left as written — a `{{sig}}` typo ships
   * braces otherwise, with nothing said. Bounded like every other echoed listing, because
   * on a forward the spelling can have been copied out of the original.
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
  /** What the tokens did. Absent when the call wrote no token at all. */
  tokens?: DraftEmailReceipt;
  /** What the draft embedded, could not, or was told about. Absent when there is nothing to say. */
  notes?: string[];
}

// The minimal client surface this orchestration needs; JmapClient satisfies it structurally.
// Declared here (rather than importing JmapClient) so the handler stays unit-testable with a
// mock, and unit-tested against one.
export interface DraftEmailClient {
  getEmailById(id: string): Promise<any>;
  /** The sending identities. Fetched once per compose, for the signature and the warning. */
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
 * Scoped to the mode, never the full three: the wrong-mode gate that runs before the
 * near-miss pass refuses the other mode's history token in every spelling, so a message that
 * listed `{{forward}}` on a reply would hand the caller a spelling this same call is about to
 * reject. `{{signature}}` applies in every mode; the history token is the mode's own, and
 * `new` has none — which is why the clause carries its own verb and article rather than
 * being interpolated into a fixed plural sentence.
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

// ---------------------------------------------------------------------------
// Mode and mode-only parameters
// ---------------------------------------------------------------------------

/**
 * The mode, exactly as spelled. NOT coerced: `docs/conventions.md` makes leniency a
 * schema-plus-coercion pair, and a defaulted mode would turn a forgotten parameter into a
 * silently unthreaded new message. A value that is not one of the three reaches here only
 * from a client that skipped schema validation, and is refused naming all three.
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
 * A parameter that belongs to one mode only.
 *
 * One schema serves three modes, so a parameter that only makes sense in one of them cannot
 * be made `required` in the schema and is refused here instead. That trade is accepted; what
 * it buys is one message shape for every mode-only parameter, so a caller learns the rule
 * once.
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
 * Refuse everything about the caller's tokens that is refusable, from the scan alone.
 *
 * All of it runs BEFORE any block is built and before any substitution, so a refused call
 * has fetched an original and nothing else — no blob written, no draft stored.
 *
 * THE ORDER IS FIXED AND IS OVER THE WHOLE CALL, NOT PER PART: wrong-mode token, near-miss,
 * repeat, one-part, then the two forward gates. Each is a separate pass across every supplied
 * part, and that is the point — a per-part loop lets part A's near-miss beat part B's
 * wrong-mode token, so the same call refuses differently depending on which body the caller
 * happened to write first.
 *
 * The wrong-mode pass leads because of what it spans. It matches the exact token AND every
 * near-miss spelling of a history token this mode does not accept, so `{{forward}}`,
 * `{{Forward}}` and `{{{forward}}}` on a reply are all refused FOR THE MODE. Run the
 * near-miss pass first and `{{{forward}}}` on a reply is refused for its spelling instead,
 * by a message that answers a caller reaching for the forwarded block with this mode's own
 * spellings — `{{signature}}` and `{{quote}}` — and never says the thing that is actually
 * wrong, which is that a reply has no forwarded block to place. Every other count and
 * presence test here (repeat, one-part, the forward gates) is over unescaped EXACT tokens
 * only.
 */
function assertTokensAcceptable(
  parts: PartScan[], mode: DraftEmailMode, asAttachment: boolean,
): void {
  const history = HISTORY_TOKEN[mode];

  // --- 1. A history token this mode does not accept, in any spelling -------
  // `BodyTokenSite.name` carries the name a near-miss near-missed, so the near-misses can be
  // read for the mode question without a second scan surface.
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
  // REFUSED, not coerced. `docs/conventions.md` bounds leniency at "rejected when guessing
  // would change the message", and guessing here would put a signature block into a body
  // that spelled something else.
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
  // Expansion is a single pass over every site, so a repeat really would store the block
  // twice — a second copy of a stranger's whole message, or a second sign-off. There is no
  // reading of a body that wants that, and the caller who meant the braces as text has the
  // escape.
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
  // The caller's slip — a message whose html carries the sign-off and whose text alternative
  // silently does not. A SOURCE that has one form and not the other is a different thing and
  // is reported per part as a note, not refused here.
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
        'When you supply both parts, place each token in both, or supply only one part and ' +
        'let the other be derived from it.',
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

  // A forward that places no {{forward}} and does not ride as .eml forwards nothing, while
  // still carrying the original's attachments, recording the source, and marking the
  // original forwarded on send. Unlike a reply without {{quote}} — which is simply a
  // message — that is a shape with no honest reading, so it is refused rather than noted.
  // Presence test only: whether the block turns out to have content is a later question.
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

/** An unavailable block, so the cause travels with it to the note. */
const unavailable = (cause: BlockUnavailableCause): BodyBlock => ({ available: false, cause });

/**
 * A block from a builder's output: available when it produced non-blank content, otherwise
 * carrying the cause the caller is owed.
 *
 * `nothing-quotable` versus `nothing-quotable-in-this-form` is decided against the source
 * THIS part would quote, not against a combined either-form gate. An images-only original
 * passes the combined gate and fails the text one, so a text-only reply to it would
 * otherwise store no history and say nothing at all.
 */
function quoteBlock(content: string | undefined, anyForm: boolean): BodyBlock {
  if (content !== undefined && !isBlank(content)) return { available: true, content };
  return unavailable(anyForm ? 'nothing-quotable-in-this-form' : 'nothing-quotable');
}

// signatureBlock lives in reply-quote.ts beside the two block builders it chooses between:
// edit_draft expands {{signature}} too, and the rule for WHICH form a part gets is one rule.

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

// CAUSE_SENTENCE and noteTokenEmpty live in inline-notes.ts: edit_draft emits the same two
// sentences from its own expansion, and a second copy of the cause wording is how the two
// surfaces drift apart.

/**
 * The identity has a signature and the caller placed none.
 *
 * Presence only, tested against the PRE-expansion body, so it cannot false-fire; and it
 * tests a SUPPLIED body, so an attachment-only stash, a body-less reply and an asAttachment
 * filler get no warning. It fires on every deliberately unsigned message, which is the
 * accepted cost of covering a body stored with no sign-off and nothing said.
 */
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
 * A token note of its own rather than a line on the image sentence, because the loss is the
 * FORMATTING first: the images riding as attachments is the visible half, but a forward of a
 * formatted message reproduced as plain text is degraded even when it carries no images at
 * all. The remedy has to terminate, so it names the one move that fixes both halves.
 */
const NOTE_FORWARD_TEXT_FORM =
  '{{forward}} ships in the text form only and the original ships HTML, so this forward loses ' +
  'its formatting and its inline images ride as attachments; put {{forward}} in htmlBody to ' +
  'keep both.';

/**
 * What the pooled-media sentence ends on for THIS tool, on the path above.
 *
 * The shared default tells the caller to re-run with `asAttachment: true`, which on
 * `draft_email` is a call the token gate refuses while `{{forward}}` is still in the body —
 * a remedy that does not terminate. Named here so the two sentences give one instruction.
 */
const POOLED_REMEDY_PLACE_IN_HTML =
  'put {{forward}} in htmlBody to embed them, or drop the token and pass asAttachment: true ' +
  'to forward the original whole.';

/** A reply that placed no {{quote}}: the default flipped, so a forgotten token is reported. */
const NOTE_REPLY_UNQUOTED =
  'This reply was stored without the original: place {{quote}} in the body to include it.';

/** An image the block minted that no part of the expanded body ends up referencing. */
function noteMintedDropped(names: (string | null | undefined)[], total: number): string {
  const listed = describePartNames(names, total);
  return (
    `${total} image(s) the quoted original displayed ${listed ? `(${listed}) ` : ''}` +
    'were dropped: after expansion no body written by this call references them. ' +
    'A token placed inside a comment or an attribute is the usual cause.'
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
// recorded-source affordance is lost. `edit_draft` never writes this header — it
// carries whatever the draft already stores, or drops it whole when the caller names
// forwardedMessageId in clearFields — and it does NOT re-vet what it carries. That is
// not because the value was vetted before: a draft composed in another client was never seen
// by this function. It is because a value Fastmail will not accept fails the CREATE loudly,
// with the old draft still intact (the recreate creates before it disposes), so the caller
// gets an error and one recovery step — clear the marking — rather than a silent drop.
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

// Filename for the attached .eml, derived from the ORIGINAL's subject (the name
// describes the attached message, not the new wrapper — a caller subject override
// does not affect it). The value is consumed as a SAVE NAME by receiving clients,
// so strip control chars (\p{Cc} covers C0 and C1 incl. U+0085), Unicode
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
// The orchestration
// ---------------------------------------------------------------------------

/**
 * Orchestrate draft_email end to end.
 *
 * Reads no environment: `attachDir` and `allowBlobAttach` are resolved by the caller and
 * between them say which attachment SOURCES this server accepts. Only ever drafts —
 * send_draft is the single tool that transmits mail, and it does the thread-state
 * maintenance from the provenance headers recorded here (#60).
 *
 * The order below is load-bearing and is the order of the numbered steps:
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
 * Steps 6 and 11 straddle step 8 deliberately, and a unit test pins the order: the plan
 * reads the caller's OWN html (the blocks are this server's markup and legitimately carry
 * identifiers the caller never wrote), while the closure check has to read what actually
 * ships. Reversed, every image-bearing reply is refused.
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
  // BOTH FLAGS ARE READ OFF `args?.` RATHER THAN THE `a` ALIAS, and must stay that way. The
  // lenient-boolean guard in tool-schema.test.ts matches `!!asAttachment` and
  // `!!args?.asAttachment` but not `!!a.asAttachment`, so tidying these back to the alias
  // would put any future bare-`!!` read of them outside the only check that looks for one.
  const asAttachment = coerceBool(args?.asAttachment) ?? false;
  const includeOriginalAttachments = coerceBool(args?.includeOriginalAttachments) ?? true;

  assertModeOnly(args?.asAttachment !== undefined, 'asAttachment', 'forward', mode);
  assertModeOnly(
    args?.includeOriginalAttachments !== undefined, 'includeOriginalAttachments', 'forward', mode,
  );
  assertModeOnly(a.mailbox !== undefined, 'mailbox', 'new', mode);
  assertModeOnly(a.inReplyTo !== undefined, 'inReplyTo', 'new', mode);
  assertModeOnly(a.references !== undefined, 'references', 'new', mode);

  const originalEmailId = a.originalEmailId;
  if (mode === 'new') {
    if (originalEmailId !== undefined) {
      throw bad("originalEmailId applies to mode:'reply' and mode:'forward' only.");
    }
  } else if (!originalEmailId) {
    throw bad(`originalEmailId is required for mode:'${mode}'.`);
  }

  // --- 2. Validate the caller's own bodies, BEFORE anything is expanded ----
  // The quote and forwarded blocks would otherwise mask a malformed body: they supply the
  // real tags an escaped-markup body lacks and the visible content an empty one lacks (#78).
  assertBodyInputs(a);

  const { from, subject: rawSubject, textBody, htmlBody } = a;
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
    // Bodies are tested with isBlank, matching the second copy of this guard in
    // createDraft. Truthiness would let a whitespace-only body through to the generic
    // message, which names none of the parameters that would fix it.
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

  // THE MESSAGE QUESTION: does this message ship an html part at all? A present-but-blank
  // htmlBody emits none — buildBodyParts drops it — so it must not carry the only sign-off.
  const messageShipsHtml = !isBlank(htmlBody);
  // THE HTML QUOTE QUESTION, kept apart from it: the html part is non-blank AND carries the
  // history token. (The builders add the third conjunct, that the original has quotable
  // html.) It decides whether the quoted images are minted, and therefore whether the text
  // alternative may describe an image the reader can actually look at.
  const historyHtmlShips =
    messageShipsHtml && !!history && (htmlPart?.scan.counts[history] ?? 0) > 0;
  const historyPlaced = !!history && supplied.some((p) => p.scan.counts[history] > 0);

  // --- 5. The identity, fetched once ---------------------------------------
  // Always fetched on this tool, because the warning below needs to know whether the
  // identity HAS a signature even when the caller placed no token.
  const identity = selectIdentity(await client.getIdentities(), from);
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
  // Block construction runs the collect pass, which is what resolves the original's image
  // references — so an unquoted reply gains no "dropped image" notes by doing work nobody
  // asked for.
  let quoteImages: QuoteImageOutcome = emptyQuoteImages();
  const htmlBlocks: BodyBlocks = { signature: undefined, quote: undefined, forward: undefined };
  const textBlocks: BodyBlocks = { signature: undefined, quote: undefined, forward: undefined };

  if (supplied.some((p) => p.scan.counts.signature > 0)) {
    htmlBlocks.signature = signatureBlock(signature, 'htmlBody', messageShipsHtml);
    textBlocks.signature = signatureBlock(signature, 'textBody', messageShipsHtml);
  }

  if (historyPlaced && mode === 'reply') {
    // No `timezone`: this tool takes no such parameter, so the attribution line is
    // formatted in the server's local zone, which is what buildQuoteBlocks defaults to.
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
  // True when the forwarded block ships in its TEXT form while the original really has html
  // worth reproducing — the cell NOTE_FORWARD_TEXT_FORM is about. Read off the builder's own
  // `htmlQuotable`, so it cannot disagree with the arm that chose the form.
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
    // The forwarded block always has a header block, so it is never "nothing quotable" in
    // the way a reply quote can be — but a caller can still place the token in a part whose
    // form carries nothing below the headers, and the note above says which.
    htmlBlocks.forward = quoteBlock(built.htmlBlock, true);
    textBlocks.forward = quoteBlock(built.textBlock, true);
  }

  // --- 8. THE SINGLE PASS, per part, over the caller's authored body -------
  // Read the security rule at the top of this file before touching these two lines. Each
  // takes the caller's OWN string; neither takes anything a builder produced.
  const expansions = new Map<PartName, BodyTokenExpansion>();
  if (htmlPart) expansions.set('htmlBody', expandBodyTokens(htmlPart.authored, htmlBlocks));
  if (textPart) expansions.set('textBody', expandBodyTokens(textPart.authored, textBlocks));

  const expandedHtml = expansions.get('htmlBody')?.text ?? (htmlPart ? '' : undefined);
  const expandedText = expansions.get('textBody')?.text ?? (textPart ? '' : undefined);

  // --- 9. A part that was content before expansion and is empty after it ---
  // Refused per part, in every mode, and this refusal wins over the "token had nothing to
  // expand to" note below: `mode:'forward'` with {{forward}} and an original with nothing
  // quotable passes the presence gate in step 4 and would otherwise store a body-less
  // forward that still carries attachments and marks the original forwarded. The message is
  // raised here rather than left to createDraft's generic contentless one so it can name
  // the causes and the remedies.
  for (const { part, authored } of supplied) {
    const before = part === 'htmlBody' ? htmlHasVisibleContent(authored) : !isBlank(authored);
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

  // --- 10. Assemble ---------------------------------------------------------
  const params: DraftEmailParams = { from, replyTo };
  if (toArg?.length) params.to = toArg;
  if (cc) params.cc = cc;
  if (bcc) params.bcc = bcc;

  let fillerBody: true | undefined;
  params.textBody = expandedText;
  params.htmlBody = expandedHtml;

  if (mode === 'new') {
    if (a.mailbox !== undefined) params.mailbox = a.mailbox;
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

    // Default the recipient to the original sender, keeping the display name via
    // formatAddress. This array bypasses coerceStringArray, so a comma inside a name is
    // never re-split into a bogus second recipient (#31).
    if (!params.to?.length) {
      params.to = Array.isArray(original.from)
        ? original.from.filter((x: any) => x?.email).map(formatAddress)
        : [];
    }
    if (!params.to?.length) {
      throw bad('Could not determine reply recipient. Please provide "to" explicitly.');
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

  const carried: AttachmentPart[] = [];
  const pooled: CidPart[] = [];
  const attachedFiles: CidPart[] = [];
  const notIncluded: CidPart[] = [];

  if (mode === 'forward' && asAttachment) {
    // Lossless form: the Email's own blobId is the raw RFC 5322 message. Inserted AFTER the
    // empty-after-expansion check above, which is what keeps a caller-supplied part being
    // tested on its own terms while a body-less forward still ships something readable. The
    // filler is ordinary prose rather than a forwarded-message block, and nothing downstream
    // reads it back: an edit of this draft replaces it like any other body.
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
    const referenced = new Set(quoteImages.resolvedParts);
    for (const entry of forwardSourceParts) {
      if (embedded.has(entry.part)) continue;
      if (!includeOriginalAttachments) {
        notIncluded.push(entry.part);
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
      const bodyMedia = entry.inBodyList
        || ((entry.part as any)?.disposition === 'inline' && !!entry.part?.cid);
      if (referenced.has(entry.part) || bodyMedia) pooled.push(entry.part);
      else attachedFiles.push(entry.part);
    }
  }

  const ledger = new InlineNoteLedger();
  const carry = recordQuoteImages(ledger, quoteImages, mode === 'forward' ? 'forward' : 'reply');

  // A minted part that nothing references AFTER expansion is dropped before assembly, and
  // the result names it. New behaviour, and reachable from caller input for the first time:
  // a token placed inside a comment or an attribute expands there, so the block is in the
  // body but its image references are not live. The old path for an unreferenced minted part
  // was checkInlineClosure's second arm THROWING, which is the wrong answer for something a
  // caller can cause.
  const liveRefs = new Set(
    expandedHtml ? extractLiveCidRefs(expandedHtml) : [],
  );
  const keptMinted = carry.minted.filter((p) => !!p.cid && liveRefs.has(p.cid));
  const droppedMinted = carry.minted.filter((p) => !keptMinted.includes(p));

  const uploaded = specs?.length
    ? await client.uploadAttachments(specs, attachDir, allowBlobAttach, {
      inlineCids: inlinePlan.inlineCids,
    })
    : undefined;

  const attachments = [...carried, ...(uploaded ?? []), ...keptMinted];
  if (attachments.length > 0) params.attachments = attachments;

  // --- 11. Closure, on what actually ships ---------------------------------
  checkInlineClosure({
    htmlBodies: [params.htmlBody],
    finalPartCids: attachments.map((part) => part.cid),
    attachedMintedCids: keptMinted.map((part) => part.cid).filter((c): c is string => !!c),
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

  const notes = [
    ...ledger.emit({
      surface: mode === 'forward' ? 'forward' : 'reply',
      ...(carry.resolvedPartCount !== undefined && { resolvedPartCount: carry.resolvedPartCount }),
      // Only on the path where placing the token in htmlBody really is the fix; anywhere
      // else the shared "re-run with asAttachment" remedy is the right one.
      ...(forwardTextFormOnly && { pooledRemedy: POOLED_REMEDY_PLACE_IN_HTML }),
    }),
    ...(droppedMinted.length > 0
      ? [noteMintedDropped(droppedMinted.map((p) => p.name), droppedMinted.length)]
      : []),
    ...await reportAuthoredInlineImages({
      uploaded,
      mintedCids: keptMinted.map((p) => p.cid).filter((c): c is string => !!c),
      plan: inlinePlan,
      emailId,
      readBack: (id) => client.getEmailById(id),
    }),
    ...emptyTokenNotes(expansions),
    ...(forwardTextFormOnly ? [NOTE_FORWARD_TEXT_FORM] : []),
    // Presence only, on a SUPPLIED body, so a body-less reply and an attachment-only stash
    // are silent.
    ...(!signaturePlaced && signature && supplied.length > 0
      ? [noteSignatureNotPlaced(identity?.email ?? from)]
      : []),
    ...(mode === 'reply' && !historyPlaced ? [NOTE_REPLY_UNQUOTED] : []),
  ];

  return {
    emailId,
    mode,
    ...(params.subject !== undefined && { subject: params.subject }),
    ...(params.to && { to: params.to }),
    ...(cc && { cc }),
    ...(receipt && { tokens: receipt }),
    ...(notes.length > 0 && { notes }),
  };
}

/**
 * Content-IDs an html body really references, read with THE SAME collector the closure check
 * uses — so "referenced" means the same thing in both places, and the drop below can never
 * disagree with the throw it exists to prevent. A token expanded inside a comment or an
 * attribute leaves the block's markup in the body but its `<img>` outside the document, and
 * this is what tells the two apart.
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

/**
 * The receipt, read off the expansion's own per-site report.
 *
 * Absent when the call wrote no `{{…}}` spelling at all, so an ordinary body's result is
 * unchanged in shape.
 */
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
      // Position order: expandBodyTokens reports sites in the order the single pass met
      // them, which is the order they sit in the body.
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
