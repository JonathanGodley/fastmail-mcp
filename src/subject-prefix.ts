// The reply/forward prefix a caller can type into a subject line, and what each route says
// when it finds one (#188).
//
// A subject that opens "Re:" or "Fwd:" reads as part of a conversation, but a subject line
// is only text. What actually threads a message is the headers this server writes for
// mode:'reply' (In-Reply-To/References) and the forwarded-message id it records for
// mode:'forward'. Type the prefix by hand onto a mode:'new' draft, or edit one onto a draft
// that carries neither, and the result reads as a reply while threading as a brand-new
// conversation — in the caller's own client and in every recipient's.
//
// It is a NOTE on both routes and a refusal on neither. Reusing an old subject for a fresh
// conversation is legitimate ("Re: your invoice" written deliberately to a new thread), and
// a refusal would leave that caller no way through at all.
//
// ONE matcher serves both routes, from this module, so compose and edit cannot come to
// disagree about what counts as a prefix — a caller warned on the way in and not on the way
// out (or the reverse) would read the difference as a rule rather than as drift.

/** Which mode the prefix claims: `Re` is a reply, `Fwd`/`Fw` a forward. */
export type SubjectPrefixKind = 'reply' | 'forward';

// A LEADING `Re`, `Fwd` or `Fw`, in any case, an optional `[n]` counter (Fastmail and
// several other clients write "Re[2]:" on a subject that has been round the loop), loose
// whitespace anywhere between the pieces, then the colon that ends the prefix.
//
// NOTHING ELSE, deliberately. A miss costs a note nobody sees, while a false hit tells a
// caller their perfectly ordinary subject will not thread — so the set stays narrow enough
// to state in one sentence and obvious enough that a reader can tell at a glance which
// subjects are in it. That rules out the localised prefixes ("AW:", "RE~:", "SV:", "Odp:")
// every mail client spells differently, and the trailing-prefix forms ("... (fwd)"): both
// would widen the set past what the description in the tool surface can honestly say.
//
// The colon is required, so a subject that merely BEGINS with those letters — "Reference
// pricing", "Fwd of the notes" — is not a prefix. Anchored, so a prefix that appears later
// in the line ("Notes on Re: pricing") is just text, which is what it looks like to a mail
// client too.
// The trailing whitespace run sits INSIDE the optional counter group on purpose. With it
// outside, a subject carrying no counter matches two adjacent runs (`\s*\s*:`), and a long
// whitespace run with no colon after it makes the engine hand characters back one at a
// time while the second run re-scans: quadratic, and measured at 265ms for 20,000 spaces
// against 0.02ms for this form. Nothing caps a subject before it reaches here and the
// server is one stdio process, so a stall here stalls every other call. This spelling
// accepts exactly the same subjects - proved row by row against the enumerated sets.
const SUBJECT_PREFIX = /^\s*(re|fwd|fw)\s*(?:\[\s*\d+\s*\]\s*)?:/i;

/**
 * The kind of prefix a subject opens with, or undefined for a subject that carries none.
 *
 * Takes `undefined` as well as a string so both callers can hand it whatever they hold: a
 * subject nobody passed, and a blank one (which `coerceSubjectOverride` has already turned
 * into `undefined`), both come back as no match rather than needing a guard of their own.
 *
 * The type test is for the compiler and for a lenient caller, and no test can observe it
 * being removed: `exec` stringifies whatever it is handed, and the pattern is anchored, so
 * an absent subject that reached it would fail to match anyway. It stays because reading a
 * subject that is not a string is a caller bug, not a subject without a prefix.
 */
export function matchSubjectPrefix(subject: string | undefined): SubjectPrefixKind | undefined {
  if (typeof subject !== 'string') return undefined;
  const match = SUBJECT_PREFIX.exec(subject);
  if (!match) return undefined;
  return match[1].toLowerCase() === 're' ? 'reply' : 'forward';
}

/**
 * What a fresh compose says: the prefix is claiming something the mode does not do, and the
 * mode that does it is one parameter away.
 *
 * The remedy follows the prefix rather than offering both, because the caller has already
 * said which one they meant by typing it.
 */
export function noteComposeSubjectPrefix(kind: SubjectPrefixKind): string {
  return kind === 'reply'
    ? "This subject reads as a reply, but mode:'new' writes no In-Reply-To or References, "
      + 'so the message starts its own conversation rather than joining the one it names. '
      + "Compose with mode:'reply' and originalEmailId to thread it — passing to as well if "
      + 'you do not want reply-all.'
    : "This subject reads as a forward, but mode:'new' records no forwarded message, so "
      + 'nothing connects this draft to the one it says it forwards. Compose with '
      + "mode:'forward' and originalEmailId to forward it properly.";
}

/**
 * What an edit says. The remedy differs from the compose one and has to: threading headers
 * are written when a draft is created and an edit cannot add them to a draft that has none,
 * so the way out is a fresh draft in the right mode rather than another edit of this one.
 */
export function noteEditSubjectPrefix(kind: SubjectPrefixKind): string {
  return kind === 'reply'
    ? 'This subject reads as a reply, but the draft carries no In-Reply-To or References and '
      + 'an edit cannot add them, so the message starts its own conversation. Compose a fresh '
      + "draft with draft_email mode:'reply' and originalEmailId, then delete this one."
    : 'This subject reads as a forward, but the draft records no forwarded message and an '
      + 'edit cannot add one, so nothing connects it to the message it says it forwards. '
      + "Compose a fresh draft with draft_email mode:'forward' and originalEmailId, then "
      + 'delete this one.';
}
