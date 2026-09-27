import { isBlank } from './body-format.js';

// Which identity a message sends as, and what sign-off that identity carries (#33).
//
// Shared by JmapClient (which writes `from`) and the compose handlers (which hand the
// signature to the pure body builders), because a signature picked under a different rule
// than the one that picks `from` would sign a message with someone else's sign-off.

/** Match an email address against an identity, supporting wildcard identities (e.g. *@example.com). */
export function matchesIdentity(identityEmail: string, address: string): boolean {
  const identity = identityEmail.toLowerCase();
  const addr = address.toLowerCase();
  if (identity === addr) return true;
  if (identity.startsWith('*@')) {
    const domain = identity.slice(1); // "@example.com"
    // A wildcard identity is only honoured for a single well-formed addr-spec. Without
    // this, a composite value like "a@evil.com,b@example.com" (or one carrying CR/LF or
    // a quoted local part) satisfies the endsWith test and lands unparsed in the
    // outgoing `from`/`mailFrom`, turning the "verified identity" check into a pass.
    // Note the pattern admits a BARE addr-spec only — a "Name <a@b.example>" form is
    // rejected on purpose, because the display name is supplied separately and is never
    // part of the value matched here. Do not widen it to accept angle-addr shapes.
    if (!/^[^\s@,;"]+@[^\s@,;"]+$/.test(addr)) return false;
    return addr.endsWith(domain);
  }
  return false;
}

/**
 * The identity a compose call will send as: the one matching an explicit `from`, else the
 * account's default. JmapClient.createDraft picks through the same two helpers below, and
 * that is the rule that actually decides the `from` header.
 *
 * Returns undefined when `from` names nothing verified. Deliberately NOT an error here:
 * createDraft raises the real "not verified for sending" refusal a moment later, and a
 * signature lookup that threw its own version first would replace an accurate message with
 * an oblique one.
 */
export function selectIdentity(identities: any[] | undefined | null, from?: string): any | undefined {
  return from ? identityFor(identities, from) : defaultIdentity(identities);
}

/**
 * The identity that verifies `address`. An exact-address identity wins over a wildcard one
 * whatever order the server lists them in, since its name and signature are the ones set up
 * for that address. Every site that picks an identity for an address goes through here.
 */
export function identityFor(identities: any[] | undefined | null, address: string): any | undefined {
  const list = (identities ?? []).filter((id: any) => typeof id?.email === 'string');
  const addr = address.toLowerCase();
  return list.find((id: any) => id.email.toLowerCase() === addr)
    ?? list.find((id: any) => matchesIdentity(id.email, address));
}

/** The account's default identity: the one that cannot be deleted, else the first listed. */
export function defaultIdentity(identities: any[] | undefined | null): any | undefined {
  const list = identities ?? [];
  return list.find((id: any) => id?.mayDelete === false) ?? list[0];
}

/**
 * An identity's configured sign-off, in the forms it was configured in.
 *
 * Both halves are optional because a server is free to store either alone. Which one a
 * message uses is decided by the body it ships, not by which was configured — see
 * signatureBlock in src/reply-quote.ts, and the signature section of docs/email-bodies.md.
 */
export interface ResolvedSignature {
  /** The identity's `htmlSignature`, unwrapped; `signatureHtmlBlock` wraps it in a plain div. */
  html?: string;
  /** The identity's `textSignature`, used for a body that ships no HTML at all. */
  text?: string;
}

/**
 * Read the sign-off off one identity. Undefined when it has none configured (a blank
 * signature counts as none); `signatureBlock` in src/reply-quote.ts turns that into its
 * `no-signature` cause.
 */
export function signatureOf(identity: any): ResolvedSignature | undefined {
  const html = typeof identity?.htmlSignature === 'string' && !isBlank(identity.htmlSignature)
    ? identity.htmlSignature : undefined;
  const text = typeof identity?.textSignature === 'string' && !isBlank(identity.textSignature)
    ? identity.textSignature : undefined;
  if (html === undefined && text === undefined) return undefined;
  return { ...(html !== undefined && { html }), ...(text !== undefined && { text }) };
}

// Deliberately no `resolveSignature(identities, from)` wrapper: every caller needs the
// identity object as well as its sign-off, because the note that reports an empty expansion
// names the address the message sends as.
