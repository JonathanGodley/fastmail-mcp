import { FastmailAuth } from './auth.js';
import { validateFastmailUrl } from './url-validation.js';
import { parseAddress, requireNonEmpty, validateClearFields, PathAccessError, InvalidInputError } from './coerce.js';
import { normalizeBodies, htmlHasVisibleContent, buildBodyParts, isBlank } from './body-format.js';
import { buildReplyBodies, hasQuoteMarker, hasTextQuoteMarker } from './reply-quote.js';
import { writeFile, mkdir, realpath, stat, lstat, open } from 'fs/promises';
import type { FileHandle } from 'fs/promises';
import { dirname, resolve, normalize, sep, basename, join } from 'path';
import { homedir } from 'os';

// A JMAP Email "attachment" body part referencing an uploaded blob. blobId is the
// stable handle the server assigns; type is the stored MIME type. New uploads carry
// only blobId/type/name/disposition; a carried (re-referenced) part may also pass
// through cid/disposition from the existing draft.
export interface AttachmentPart {
  blobId: string;
  type: string;
  name?: string;
  disposition?: string;
  cid?: string;
}

// True if `child` is `parent` itself or nested beneath it. Case-fold is a parameter:
// the read guard compares case-insensitively on Win32 (NTFS folds case, so a
// case-sensitive startsWith is bypassable), while the write guard keeps its existing
// byte-exact compare so download behaviour is unchanged.
function isPathContained(child: string, parent: string, caseInsensitive: boolean): boolean {
  let c = child, p = parent;
  if (caseInsensitive) { c = c.toLowerCase(); p = p.toLowerCase(); }
  return c === p || c.startsWith(p + sep);
}

// Shared lexical pre-check: resolve `inputPath` against `allowedDir` (relative incl.
// a bare filename lands inside the root in one step; absolute must already be within),
// reject null bytes, and verify lexical containment. Returns the resolved absolute
// path. Throws PathAccessError on null bytes or escape. The canonical (realpath)
// re-verification is the caller's job — this is only the cheap first gate.
function lexicalContainedPath(inputPath: string, allowedDir: string, caseInsensitive: boolean): string {
  const resolved = resolve(allowedDir, normalize(inputPath));
  if (resolved.includes('\0')) {
    throw new PathAccessError('path contains null bytes');
  }
  if (!isPathContained(resolved, allowedDir, caseInsensitive)) {
    throw new PathAccessError(`path must be within ${allowedDir}. Received: ${inputPath}`);
  }
  return resolved;
}

// Reject Windows path forms that can dodge the lexical containment compare. Applied
// to the raw input on every platform: these shapes are never a legitimate attachment
// path, and resolving them first would mask the escape. (Device namespaces and UNC
// roots jump outside the drive-relative root; a drive-relative `C:foo` resolves
// against the drive's own CWD; a `:` past the drive names an NTFS alternate data
// stream; a `~` segment can be an 8.3 short name aliasing a long name past the compare.)
function rejectWindowsPathEscapes(input: string): void {
  if (/^\\\\[?.]\\/.test(input)) {
    throw new PathAccessError('path uses a Windows device namespace (\\\\?\\ or \\\\.\\), which is not allowed.');
  }
  if (/^(\\\\|\/\/)/.test(input)) {
    throw new PathAccessError('path is a UNC network path, which is not allowed.');
  }
  if (/^[A-Za-z]:(?![\\/])/.test(input)) {
    throw new PathAccessError('path is drive-relative (e.g. C:foo); use an absolute path or a name under the attach directory.');
  }
  // Strip a leading drive letter's own colon before scanning for a stream colon.
  if (input.replace(/^[A-Za-z]:/, '').includes(':')) {
    throw new PathAccessError("path contains a ':' (NTFS alternate data stream), which is not allowed.");
  }
  // 8.3 short-name segment (e.g. PROGRA~1) can alias a long name past the compare. Match
  // the 8.3 form specifically — a tilde followed by a digit — so legitimate filenames that
  // merely contain a tilde (e.g. report~final.pdf) are NOT rejected.
  if (/~\d/.test(input)) {
    throw new PathAccessError("path contains an 8.3 short-name segment (e.g. PROGRA~1), which is not allowed; use the full name.");
  }
}

// Extension -> MIME map for the Content-Type we POST when the caller omits one. Only
// sets the upload header; the recipient sees whatever type the server echoes back.
const EXT_CONTENT_TYPES: Record<string, string> = {
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  zip: 'application/zip',
  txt: 'text/plain',
  csv: 'text/csv',
  json: 'application/json',
  html: 'text/html',
};

function guessContentType(path: string): string {
  const ext = basename(path).split('.').pop()?.toLowerCase() ?? '';
  return EXT_CONTENT_TYPES[ext] ?? 'application/octet-stream';
}

// RFC 2045 token grammar for type/subtype, with a length cap. A positive grammar
// (reject anything outside the token set) closes header injection via the
// Content-Type we POST — defense in depth alongside undici's own header validation.
// MIME parameters (e.g. "; charset=utf-8") are intentionally not accepted: only the
// type/subtype is needed to upload a blob.
const MIME_TYPE_PATTERN = /^[A-Za-z0-9!#$%&'*+.^_`|~-]+\/[A-Za-z0-9!#$%&'*+.^_`|~-]+$/;

function validateContentType(value: string, index: number): string {
  const v = value.trim();
  if (v.length > 255 || !MIME_TYPE_PATTERN.test(v)) {
    throw new PathAccessError(`attachments[${index}] has an invalid contentType '${value}'. Use a MIME type like application/pdf.`);
  }
  return v;
}

export interface JmapSession {
  apiUrl: string;
  accountId: string;
  capabilities: Record<string, any>;
  downloadUrl?: string;
  uploadUrl?: string;
}

export interface JmapRequest {
  using: string[];
  methodCalls: [string, any, string][];
}

export interface JmapResponse {
  methodResponses: Array<[string, any, string]>;
  sessionState: string;
}

export interface QueryResult<T = any> {
  items: T[];
  total?: number;
  // Out-of-band metadata for the default Trash/Spam exclusion. Populated by
  // searchEmails/getEmails when an exclusion was active; read by the handlers to
  // emit a trailing note. NEVER serialized into the JSON body or the raw path
  // (same discipline as getThread's hiddenDraftCount). `hidden:null` = the hidden
  // count could not be computed (degraded — emit the fail-closed note).
  exclusion?: {
    hidden: number | null;
    excludedRoles: string[];   // roles actually excluded (note fires iff hidden>0)
    unresolvedRoles: string[]; // roles intended-but-NOT-excluded (fail-loud note)
  };
}

// Result of computeExclusion: the mailbox ids to exclude via inMailboxOtherThan,
// the display labels of roles actually excluded, and the labels of roles we meant
// to exclude but couldn't resolve (fail-loud, never silently included).
export interface ExclusionResult {
  excludeIds: string[];
  excludedRoles: string[];
  unresolvedRoles: string[];
}

// Shared Email/get property lists — keep in sync per CLAUDE.md rules.
// COMPACT: used by the list/search tools and getThread (metadata + preview, no bodies)
export const EMAIL_PROPERTIES_COMPACT = [
  'id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'receivedAt',
  'preview', 'keywords', 'threadId', 'messageId', 'references', 'inReplyTo',
  'hasAttachment', 'header:List-Unsubscribe:asURLs', 'blobId', 'size', 'mailboxIds',
] as const;

// VERBOSE: superset with body properties — used by verbose mode and getEmailById
// `sentAt` is here (a get-path superset addition, allowed by the property-consistency rule)
// for the reply-quote attribution (when the original was actually written), not in COMPACT.
export const EMAIL_PROPERTIES_VERBOSE = [
  ...EMAIL_PROPERTIES_COMPACT,
  'textBody', 'htmlBody', 'attachments', 'bodyValues', 'sentAt',
] as const;

export const EMAIL_BODY_PROPERTIES = ['partId', 'blobId', 'type', 'size', 'name'] as const;

export interface MailboxInfo {
  name: string;
  // The stable JMAP role (lowercased), or null for a custom folder/label. Only
  // the Mailbox object carries a role; the Email object never does (RFC 8621).
  role: string | null;
}

// Build an id -> {name, role} lookup from a Mailbox/get list. We require a real
// string `name` (so a malformed/missing name leaves the id unresolvable — surfaced
// later, not silently dropped) but key on the id, so custom labels — which have
// role:null — still resolve their name; default mailboxes carry names like
// "Trash"/"Archive" AND a role. Role is lowercased here so callers get the
// docs-promised lowercase form regardless of server casing. (#10, #49)
export function buildMailboxInfoMap(mailboxes: any[]): Map<string, MailboxInfo> {
  const map = new Map<string, MailboxInfo>();
  for (const mb of mailboxes || []) {
    if (mb && mb.id && typeof mb.name === 'string') {
      const role = typeof mb.role === 'string' && mb.role ? mb.role.toLowerCase() : null;
      map.set(mb.id, { name: mb.name, role });
    }
  }
  return map;
}

// Attach the resolved mailbox/label location onto each raw email as NON-enumerable
// properties (`_mailboxNames`, `_mailboxRoles`, `_unresolvedMailboxIds`) so
// JSON.stringify — the raw:true paths — omits all three while simplifyEmail can
// still read them. Raw output therefore stays pure JMAP (opaque mailboxIds), and
// the simplified output carries friendly names + stable roles.
//
// NEVER-SILENT, NON-THROWING resolution (#53). For each of the email's mailboxIds:
//   - resolves to a name  -> name goes to `_mailboxNames`; its role (if any) to `_mailboxRoles`.
//   - does NOT resolve    -> the raw id goes to `_unresolvedMailboxIds` (the fallback +
//                            the explicit "this couldn't be named" indicator).
// So a promised location is never silently dropped: it appears either as a friendly
// name or as a raw id flagged in `unresolvedMailboxIds`. This function does NOT throw
// on an unresolved id, and it never fabricates a name.
//
// WHY not throw, and WHY not silently omit (do not "fix" this back to either):
//   - An unresolved id is rare and benign — a just-created custom folder, a TOCTOU
//     mailbox-creation race on the separately-fetched list/search mailbox list, or a
//     mailbox with a malformed/missing `name`. Trash/Spam (and all role mailboxes)
//     ALWAYS resolve, so the location that motivated #49/#53 is never the unresolved one.
//   - Strict-throw was rejected as disproportionate: it would fail a whole
//     list_emails/search_emails page over one rare/benign id, against the codebase's
//     documented best-effort resolver posture.
//   - Silently omitting the id (the prior behaviour) WAS the #53 bug — a promised
//     field vanished with no trace.
// A genuine Mailbox/get `error` response is a different thing and still throws via the
// callers' existing catches (a real failure stays loud); that is not this path.
//
// `_mailboxRoles` and `_unresolvedMailboxIds` are attached only when non-empty (so the
// formatter's addIf omits them); `_mailboxNames` keeps its existing attach-when-non-empty
// behaviour. `roles` and `mailboxes` are INDEPENDENT sets, not parallel arrays — a custom
// folder contributes a name but no role, so their lengths can differ. (#10, #49, #53)
export function attachMailboxInfo(emails: any[], map: Map<string, MailboxInfo>): void {
  for (const email of emails || []) {
    if (!email || !email.mailboxIds) continue;
    const ids = Object.keys(email.mailboxIds);
    if (ids.length === 0) continue;
    const names: string[] = [];
    const roles: string[] = [];
    const unresolved: string[] = [];
    for (const id of ids) {
      const info = map.get(id);
      if (info) {
        names.push(info.name);
        if (info.role) roles.push(info.role);
      } else {
        unresolved.push(id);
      }
    }
    if (names.length > 0) {
      Object.defineProperty(email, '_mailboxNames', { value: names, enumerable: false, configurable: true });
    }
    if (roles.length > 0) {
      Object.defineProperty(email, '_mailboxRoles', { value: roles, enumerable: false, configurable: true });
    }
    if (unresolved.length > 0) {
      Object.defineProperty(email, '_unresolvedMailboxIds', { value: unresolved, enumerable: false, configurable: true });
    }
  }
}

// Cap the mailbox names listed in a not-found error so a large account doesn't
// produce a huge message; the list_mailboxes pointer keeps a truncated list actionable.
const MAILBOX_LIST_CAP = 30;

function formatMailboxNotFound(input: string, mailboxes: any[]): string {
  const entries = (mailboxes || [])
    .filter(mb => mb && typeof mb.name === 'string')
    .map(mb => (mb.role ? `${mb.name} (${mb.role})` : mb.name));
  const shown = entries.slice(0, MAILBOX_LIST_CAP);
  let list = shown.join(', ');
  if (entries.length > shown.length) {
    list += `, …and ${entries.length - shown.length} more — call list_mailboxes for the full list`;
  }
  return `Mailbox '${input}' not found. Use a name, or a role (inbox/archive/sent/drafts/trash/junk). Valid: ${list}`;
}

// Exact-only mailbox resolution: exact id -> role (case-insensitive) -> exact name
// (case-insensitive); else throw InvalidInputError with a (capped) valid list. NO
// substring matching (substring is an injection-steering primitive on write paths and
// can mis-resolve). Edge: a custom mailbox literally named after a role (e.g. "Archive")
// is shadowed by the role branch — acceptable, and the reason the docs use "Receipts"
// not "Archive" as the name example. Exported pure for unit testing.
export function resolveMailbox(mailboxes: any[], input: string): any {
  const list = mailboxes || [];
  const raw = String(input).trim();
  const byId = list.find(mb => mb && mb.id === raw);
  if (byId) return byId;
  const lower = raw.toLowerCase();
  const byRole = list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === lower);
  if (byRole) return byRole;
  const byName = list.find(mb => mb && typeof mb.name === 'string' && mb.name.toLowerCase() === lower);
  if (byName) return byName;
  throw new InvalidInputError(formatMailboxNotFound(raw, list));
}

// Compute the default Trash/Spam exclusion. Resolves trash/junk by EXACT role only
// (case-insensitive) — NEVER findMailboxByRoleOrName, whose substring name fallback
// could mis-hit a custom mailbox (e.g. "Junk mail rules") and silently hide real mail.
// When an explicit mailbox is set, exclusion is off (the explicit scope wins). When we
// intend to exclude a role we can't resolve (role absent, OR an empty/degraded mailbox
// list), DO NOT silently include it: flag it in unresolvedRoles so the handler emits a
// fail-loud "not excluded" note — never run a default search/list with zero exclusion
// ids and zero disclosure. Exported pure for unit testing.
export function computeExclusion(
  mailboxes: any[],
  opts: { includeTrash?: boolean; includeSpam?: boolean; hasExplicitMailbox?: boolean },
): ExclusionResult {
  const excludeIds: string[] = [];
  const excludedRoles: string[] = [];
  const unresolvedRoles: string[] = [];
  if (opts.hasExplicitMailbox) {
    return { excludeIds, excludedRoles, unresolvedRoles };
  }
  const list = mailboxes || [];
  const findRole = (role: string) => list.find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === role);
  if (!opts.includeTrash) {
    const tb = findRole('trash');
    if (tb) { excludeIds.push(tb.id); excludedRoles.push('Trash'); }
    else unresolvedRoles.push('Trash');
  }
  if (!opts.includeSpam) {
    const jb = findRole('junk');
    if (jb) { excludeIds.push(jb.id); excludedRoles.push('Spam'); }
    else unresolvedRoles.push('Spam');
  }
  return { excludeIds, excludedRoles, unresolvedRoles };
}

/** Match an email address against an identity, supporting wildcard identities (e.g. *@example.com). */
function matchesIdentity(identityEmail: string, address: string): boolean {
  const identity = identityEmail.toLowerCase();
  const addr = address.toLowerCase();
  if (identity === addr) return true;
  if (identity.startsWith('*@')) {
    const domain = identity.slice(1); // "@example.com"
    return addr.endsWith(domain) && addr.indexOf('@') > 0;
  }
  return false;
}

export class JmapClient {
  private auth: FastmailAuth;
  private session: JmapSession | null = null;

  constructor(auth: FastmailAuth) {
    this.auth = auth;
  }

  /**
   * Extract the result from a JMAP method response, throwing on method-level errors.
   */
  protected getMethodResult(response: JmapResponse, index: number): any {
    if (!response.methodResponses || index >= response.methodResponses.length) {
      throw new Error(
        `JMAP response missing expected method at index ${index} (got ${response.methodResponses?.length ?? 0} responses)`
      );
    }
    const entry = response.methodResponses[index];
    if (!Array.isArray(entry) || entry.length < 2) {
      throw new Error(`JMAP response entry at index ${index} is malformed`);
    }
    const [tag, result] = entry;
    if (tag === 'error') {
      throw new Error(`JMAP error: ${result.type}${result.description ? ' - ' + result.description : ''}`);
    }
    return result;
  }

  /**
   * Extract the .list array from a JMAP method response, with null safety.
   */
  protected getListResult(response: JmapResponse, index: number): any[] {
    const result = this.getMethodResult(response, index);
    return result?.list || [];
  }

  /**
   * Like getListResult, but returns [] when the method at `index` is absent
   * instead of throwing. Used for an appended Mailbox/get (the trailing #10
   * mailbox-name resolver): a server that drops the trailing method, or an
   * older 2-entry test stub, degrades to "no names resolved" rather than an error.
   */
  protected readListResultIfPresent(response: JmapResponse, index: number): any[] {
    if (!response.methodResponses || index >= response.methodResponses.length) return [];
    return this.getListResult(response, index);
  }

  /**
   * Build a QueryResult from a query + get pair.
   * queryIndex is the /query response; listIndex is the /get response.
   */
  protected getQueryResult(response: JmapResponse, queryIndex: number, listIndex: number): QueryResult {
    const queryResult = this.getMethodResult(response, queryIndex);
    const items = this.getListResult(response, listIndex);
    const total = queryResult?.total;
    return total != null ? { items, total } : { items };
  }

  async getSession(): Promise<JmapSession> {
    if (this.session) {
      return this.session;
    }

    const response = await fetch(this.auth.getSessionUrl(), {
      method: 'GET',
      headers: this.auth.getAuthHeaders()
    });

    if (!response.ok) {
      throw new Error(`Failed to get session: ${response.statusText}`);
    }

    const sessionData = await response.json() as any;

    // Validate every URL the server hands us before we send the bearer token to it.
    // The downloadUrl/uploadUrl are URL templates with {accountId}/{blobId}/etc.
    // placeholders, so we strip those for parsing and validate origin only.
    const allowUnsafe = this.auth.getAllowUnsafe();
    const stripTemplate = (url: string) => url.replace(/\{[^}]+\}/g, 'x');
    if (typeof sessionData.apiUrl !== 'string') {
      throw new Error('Invalid session response: apiUrl missing');
    }
    validateFastmailUrl(sessionData.apiUrl, 'session.apiUrl', allowUnsafe);
    if (typeof sessionData.downloadUrl === 'string') {
      validateFastmailUrl(stripTemplate(sessionData.downloadUrl), 'session.downloadUrl', allowUnsafe);
    }
    if (typeof sessionData.uploadUrl === 'string') {
      validateFastmailUrl(stripTemplate(sessionData.uploadUrl), 'session.uploadUrl', allowUnsafe);
    }

    this.session = {
      apiUrl: sessionData.apiUrl,
      accountId: sessionData.primaryAccounts?.['urn:ietf:params:jmap:mail']
        || sessionData.primaryAccounts?.['urn:ietf:params:jmap:core']
        || Object.keys(sessionData.accounts)[0],
      capabilities: sessionData.capabilities,
      downloadUrl: sessionData.downloadUrl,
      uploadUrl: sessionData.uploadUrl
    };

    return this.session;
  }

  async makeRequest(request: JmapRequest): Promise<JmapResponse> {
    const session = await this.getSession();
    
    const response = await fetch(session.apiUrl, {
      method: 'POST',
      headers: this.auth.getAuthHeaders(),
      body: JSON.stringify(request)
    });

    if (!response.ok) {
      throw new Error(`JMAP request failed: ${response.statusText}`);
    }

    const data = await response.json();
    if (!data || !Array.isArray(data.methodResponses)) {
      throw new Error('Invalid JMAP response: missing or malformed methodResponses');
    }
    return data as JmapResponse;
  }

  // Resolve a fixed role with a SUBSTRING name fallback. The substring fallback is an
  // injection-steering / mis-resolution hazard on any exclusion/delete/move target, so
  // this is kept ONLY for the compose path (drafts/sent save target), where it resolves
  // a benign save destination. Default-exclusion uses computeExclusion (exact role),
  // delete/move/the #12 sweep use resolveMailbox / resolveMailboxId (exact only).
  protected findMailboxByRoleOrName(mailboxes: any[], role: string, nameFallback?: string): any | undefined {
    return mailboxes.find(mb => mb.role === role) ||
           (nameFallback ? mailboxes.find(mb => mb.name.toLowerCase().includes(nameFallback)) : undefined);
  }

  // Resolve trash/junk for the default Trash/Spam exclusion by EXACT role only
  // (case-insensitive) — used by both searchEmails and getEmails. Fixed-role lookup
  // with a private helper to share the resolved id between the visible filter and the
  // hidden-count query.
  private findByExactRole(mailboxes: any[], role: string): any | undefined {
    const target = role.toLowerCase();
    return (mailboxes || []).find(mb => mb && typeof mb.role === 'string' && mb.role.toLowerCase() === target);
  }

  // Resolve an optional mailbox input to an id. undefined/blank -> undefined (no filter).
  // Else resolve EXACTLY against a passed-in list (shared, no double-fetch) or one
  // getMailboxes(); throws InvalidInputError on no match. Used by every swept tool
  // (reads + writes) — safe to share now that matching is exact.
  private async resolveMailboxId(input?: string, mailboxes?: any[]): Promise<string | undefined> {
    if (input === undefined || input === null || String(input).trim() === '') return undefined;
    const list = mailboxes ?? await this.getMailboxes();
    return resolveMailbox(list, input).id;
  }

  async getMailboxes(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Mailbox/get', { accountId: session.accountId }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0);
  }

  // Options-object signature (a positional add of the three scope bools would be
  // fragile). When no explicit `mailbox` is set, applies the same default Trash/Spam
  // exclusion + hidden-count as searchEmails. Its only structural filter is the
  // excludeDrafts keyword — isUnread/isPinned are NOT exposed on list_emails.
  async getEmails(opts: {
    mailbox?: string;
    limit?: number;
    ascending?: boolean;
    includeTrash?: boolean;
    includeSpam?: boolean;
    excludeDrafts?: boolean;
  } = {}): Promise<QueryResult> {
    const mailboxes = await this.getMailboxes();
    const resolvedMailboxId = await this.resolveMailboxId(opts.mailbox, mailboxes);

    const base: any = {};
    if (resolvedMailboxId) base.inMailbox = resolvedMailboxId;

    const conds: any[] = [];
    if (opts.excludeDrafts) conds.push({ notKeyword: '$draft' });

    const hasExplicitMailbox = !!resolvedMailboxId;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: opts.includeTrash,
      includeSpam: opts.includeSpam,
      hasExplicitMailbox,
    });
    const exclusionIntended = !hasExplicitMailbox && (!opts.includeTrash || !opts.includeSpam);

    return this.runFilteredQuery({
      base,
      conds,
      exclusion,
      exclusionIntended,
      limit: opts.limit ?? 20,
      ascending: opts.ascending ?? false,
      mailboxes,
    });
  }

  async getEmailById(id: string): Promise<any> {
    const session = await this.getSession();

    // No maxBodyValueBytes: verified live (2026-06-24) that Fastmail does NOT truncate body
    // values by default (a 5 MB body returned whole, isTruncated=false), so reply_email gets
    // the complete original to quote. (An explicit maxBodyValueBytes:0 is REJECTED by Fastmail
    // with invalidArguments, so we must not send it.) The reply-quote module still appends an
    // elision marker if any bodyValue ever reports isTruncated, as a defensive net.
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [id],
          properties: [...EMAIL_PROPERTIES_VERBOSE],
          bodyProperties: [...EMAIL_BODY_PROPERTIES],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'email'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name', 'role'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notFound && result.notFound.includes(id)) {
      throw new Error(`Email with ID '${id}' not found`);
    }

    const email = result.list?.[0];
    if (!email) {
      throw new Error(`Email with ID '${id}' not found or not accessible`);
    }

    attachMailboxInfo([email], buildMailboxInfoMap(this.readListResultIfPresent(response, 1)));
    return email;
  }

  async getIdentities(): Promise<any[]> {
    const session = await this.getSession();
    
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Identity/get', {
          accountId: session.accountId
        }, 'identities']
      ]
    };

    const response = await this.makeRequest(request);
    return this.getListResult(response, 0);
  }

  async getDefaultIdentity(): Promise<any> {
    const identities = await this.getIdentities();
    
    // Find the default identity (usually the one that can't be deleted)
    return identities.find((id: any) => id.mayDelete === false) || identities[0];
  }

  async sendEmail(email: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    mailbox?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
    attachments?: AttachmentPart[];
  }): Promise<string> {
    const session = await this.getSession();

    // Get all identities to validate from address
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    // Determine which identity to use
    let selectedIdentity;
    if (email.from) {
      // Validate that the from address matches an available identity
      selectedIdentity = identities.find(id => matchesIdentity(id.email, email.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      // Use default identity
      selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
    }

    // Use the requested from address (not the identity email, which may be a wildcard like *@domain)
    const fromEmail = email.from || selectedIdentity.email;

    // Get the mailbox IDs we need
    const mailboxes = await this.getMailboxes();
    const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');

    if (!draftsMailbox) {
      throw new Error('Could not find Drafts mailbox to save email');
    }
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox to move email after sending');
    }

    // Use provided mailbox (resolved id/role/name) or default to drafts for initial
    // creation. Resolving shares the already-fetched list (no double-fetch); an unknown
    // mailbox throws InvalidInputError, and an id is validated against the list too.
    const initialMailboxId = email.mailbox ? resolveMailbox(mailboxes, email.mailbox).id : draftsMailbox.id;

    // Ensure we have at least one body type (zero-width/whitespace-only counts as absent).
    if (isBlank(email.textBody) && isBlank(email.htmlBody)) {
      throw new Error('Either textBody or htmlBody must be provided');
    }

    const initialMailboxIds: Record<string, boolean> = {};
    initialMailboxIds[initialMailboxId] = true;

    const sentMailboxIds: Record<string, boolean> = {};
    sentMailboxIds[sentMailbox.id] = true;

    // Generate the body parts: html-only input gets an auto text/plain fallback where one
    // is derivable (ships html-only otherwise; a no-body message is rejected by shapeBodies).
    const emailObject = {
      mailboxIds: initialMailboxIds,
      keywords: { $draft: true },
      from: [{ name: selectedIdentity.name, email: fromEmail }],
      to: email.to.map(parseAddress),
      cc: email.cc?.map(parseAddress) || [],
      bcc: email.bcc?.map(parseAddress) || [],
      subject: email.subject,
      ...(email.inReplyTo && { inReplyTo: email.inReplyTo }),
      ...(email.references && { references: email.references }),
      ...(email.replyTo?.length && { replyTo: email.replyTo.map(parseAddress) }),
      ...(email.attachments?.length && { attachments: email.attachments }),
      ...this.shapeBodies(email.textBody, email.htmlBody),
    };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject }
        }, 'createEmail'],
        ['EmailSubmission/set', {
          accountId: session.accountId,
          create: {
            submission: {
              emailId: '#draft',
              identityId: selectedIdentity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: [
                  ...email.to.map(addr => ({ email: parseAddress(addr).email })),
                  ...(email.cc || []).map(addr => ({ email: parseAddress(addr).email })),
                  ...(email.bcc || []).map(addr => ({ email: parseAddress(addr).email })),
                ]
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              keywords: { $seen: true }
            }
          }
        }, 'submitEmail']
      ]
    };

    const response = await this.makeRequest(request);

    const emailResult = this.getMethodResult(response, 0);
    if (emailResult.notCreated?.draft) {
      const err = emailResult.notCreated.draft;
      throw new Error(`Failed to create email: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const emailId = emailResult.created?.draft?.id;
    if (!emailId) {
      throw new Error('Email creation returned no email ID');
    }

    const submissionResult = this.getMethodResult(response, 1);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(`Failed to submit email: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Email submission returned no submission ID');
    }

    return submissionId;
  }

  async createDraft(email: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    mailbox?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
    attachments?: AttachmentPart[];
  }): Promise<string> {
    const session = await this.getSession();

    // Validate at least one meaningful field is present (zero-width/whitespace-only
    // bodies count as absent). Attachments count as content too: an attachment-only
    // draft is a valid artifact (and is consistent with edit_draft, which preserves a
    // body-less draft that carries attachments).
    if (!email.to?.length && !email.subject && isBlank(email.textBody) && isBlank(email.htmlBody) && !email.attachments?.length) {
      throw new Error('At least one of to, subject, textBody, htmlBody, or attachments must be provided');
    }

    // Get all identities to resolve from address
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (email.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, email.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
    }

    const fromEmail = email.from || selectedIdentity.email;

    // Resolve the save target. Fetch the mailbox list unconditionally now (a name/role
    // needs it, and an explicit id is validated against it too) and share it. An unknown
    // mailbox throws InvalidInputError; otherwise default to the Drafts mailbox.
    const mailboxes = await this.getMailboxes();
    let draftMailboxId: string;
    if (email.mailbox) {
      draftMailboxId = resolveMailbox(mailboxes, email.mailbox).id;
    } else {
      const draftsMailbox = this.findMailboxByRoleOrName(mailboxes, 'drafts', 'draft');
      if (!draftsMailbox) {
        throw new Error('Could not find Drafts mailbox');
      }
      draftMailboxId = draftsMailbox.id;
    }

    const mailboxIds: Record<string, boolean> = {};
    mailboxIds[draftMailboxId] = true;

    const emailObject: any = {
      mailboxIds,
      keywords: { $draft: true },
      from: [{ name: selectedIdentity.name, email: fromEmail }],
    };

    if (email.to?.length) emailObject.to = email.to.map(parseAddress);
    if (email.cc?.length) emailObject.cc = email.cc.map(parseAddress);
    if (email.bcc?.length) emailObject.bcc = email.bcc.map(parseAddress);
    if (email.subject) emailObject.subject = email.subject;
    if (email.inReplyTo?.length) emailObject.inReplyTo = email.inReplyTo;
    if (email.references?.length) emailObject.references = email.references;
    if (email.replyTo?.length) emailObject.replyTo = email.replyTo.map(parseAddress);
    if (email.attachments?.length) emailObject.attachments = email.attachments;
    // Generate the body parts (auto text/plain fallback for html-only input where
    // derivable; ships html-only otherwise; no-body html is rejected by shapeBodies). A
    // draft with neither body is allowed — shapeBodies returns empty shaping in that case.
    Object.assign(emailObject, this.shapeBodies(email.textBody, email.htmlBody));

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject }
        }, 'createDraft']
      ]
    };

    const response = await this.makeRequest(request);

    const result = this.getMethodResult(response, 0);

    // Propagate server-provided error details from notCreated
    if (result.notCreated?.draft) {
      const err = result.notCreated.draft;
      throw new Error(`Failed to create draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    // Throw if created ID is missing instead of returning silently
    const emailId = result.created?.draft?.id;
    if (!emailId) {
      throw new Error('Draft creation returned no email ID');
    }

    return emailId;
  }

  // Extract a stored body value by MIME type, keyed into bodyValues by partId.
  //
  // Server behaviour (verified live against Fastmail, 2026-06-23):
  //  - The server does NOT auto-generate the missing text/html partner in either
  //    direction; the client owns keeping the pair in sync.
  //  - A single-format draft has its ONE part aliased into BOTH the textBody and
  //    htmlBody lists (e.g. a text-only draft lists the text/plain part under htmlBody
  //    too, with type "text/plain"). So we select by the part's actual MIME type — not
  //    mere presence in a list — otherwise we'd read the text value into the html slot
  //    and synthesise a phantom text/html part on recreate.
  //  - JMAP body properties are immutable (RFC 8621 §4.1), which is why updateDraft
  //    rebuilds and re-sends the bodies via destroy+recreate rather than patching.
  // Takes the first part of the given type (drafts here carry at most one per type). If a
  // value were ever elided from bodyValues, that format reads as undefined rather than a
  // partial body (callers fetch full values, so this won't occur in practice).
  private bodyValueForType(parts: any[] | undefined, mimeType: string, bodyValues: Record<string, any>): string | undefined {
    const part = parts?.find((p: any) => p.type === mimeType && p.partId != null && bodyValues[p.partId]);
    return part ? bodyValues[part.partId].value : undefined;
  }

  // Generate the JMAP body-part shaping for an authoring path (sendEmail/createDraft) to
  // splat into the email object. normalizeBodies derives the text/plain fallback from html
  // when none was supplied; we degrade gracefully — ship html-only when the html has
  // visible media but no derivable text — and reject ONLY a genuinely no-body message
  // (html present that renders to nothing AND has no image). A message with neither body
  // returns empty shaping (a body-less draft is
  // allowed; the no-body reject only fires when an html body was actually provided).
  private shapeBodies(textBody?: string, htmlBody?: string) {
    const normalized = normalizeBodies({ textBody, htmlBody });
    if (normalized.htmlOnly && !htmlHasVisibleContent(htmlBody!)) {
      throw new Error('This message has no readable body; add text or visible content.');
    }
    return buildBodyParts(normalized);
  }

  async updateDraft(emailId: string, updates: {
    to?: string[];
    cc?: string[];
    bcc?: string[];
    subject?: string;
    textBody?: string;
    htmlBody?: string;
    from?: string;
    replyTo?: string[];
    clearFields?: string[];
    attachments?: AttachmentPart[];
    removeAttachments?: string[];
    // Reply-quote preservation on body edit (#37, redesigned #42). If a reply draft already
    // carries the quoted original (detected on its EXISTING body, which this server generated)
    // and the edit would touch the body in a way that could drop the quote, the caller must
    // say what to do: originalEmailId = regenerate and keep the quote from that (caller-named)
    // message; noQuote = deliberately drop it. Absent both, the edit is rejected (no silent loss).
    originalEmailId?: string;
    noQuote?: boolean;
  }): Promise<{ id: string; orphanedOldDraftId?: string }> {
    const session = await this.getSession();

    // Fetch the existing email
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'textBody', 'htmlBody', 'bodyValues', 'mailboxIds', 'keywords', 'inReplyTo', 'references', 'attachments'],
          // Inline list (NOT the module-level EMAIL_BODY_PROPERTIES) — extended with
          // name/disposition/cid so the faithful recreate can carry attachment metadata
          // and detect inline (cid:) images.
          bodyProperties: ['partId', 'blobId', 'type', 'size', 'name', 'disposition', 'cid'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const existingEmail = this.getListResult(getResponse, 0)[0];
    if (!existingEmail) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }

    // Verify it's a draft
    if (!existingEmail.keywords?.$draft) {
      throw new Error('Cannot edit a non-draft email');
    }

    // Faithful-recreate guards. The recreate below rebuilds the message from flat
    // convenience props (textBody/htmlBody/attachments), which can't round-trip an
    // inline (cid:) image's multipart/related linkage, nor a non-text/non-html body
    // part (e.g. an externally-created multipart/signed or text/calendar draft). Rather
    // than silently drop or mangle those, reject loudly. (Verified live 2026-06-24: a
    // cid: inline image surfaces in `attachments` with disposition:'inline'; a regular
    // attachment that merely carries a cid keeps disposition 'attachment'/absent and is
    // carried fine. Inline-image authoring/editing is tracked as fork issue #13.)
    const existingAttachments: any[] = existingEmail.attachments || [];
    if (existingAttachments.some((a: any) => a.disposition === 'inline')) {
      throw new Error('This draft has inline images, which editing can\'t preserve yet. Recreate the draft instead (see issue #13).');
    }
    // Alias-aware: a single-format draft aliases its one part into BOTH lists with its
    // real MIME type, so a text-only draft lists text/plain twice — not a reject. Only a
    // genuinely non-text/non-html typed part trips this; a typeless part is left alone.
    const allBodyParts = [...(existingEmail.textBody || []), ...(existingEmail.htmlBody || [])];
    if (allBodyParts.some((p: any) => p.type && p.type !== 'text/plain' && p.type !== 'text/html')) {
      throw new Error('This draft has a body part that isn\'t plain text or HTML, which editing can\'t preserve. Recreate the draft instead.');
    }

    // Resolve identity
    const identities = await this.getIdentities();
    if (!identities || identities.length === 0) {
      throw new Error('No sending identities found');
    }

    let selectedIdentity;
    if (updates.from) {
      selectedIdentity = identities.find(id => matchesIdentity(id.email, updates.from!));
      if (!selectedIdentity) {
        throw new Error('From address is not verified for sending. Choose one of your verified identities.');
      }
    } else {
      // Use existing from, or fall back to default identity
      const existingFrom = existingEmail.from?.[0]?.email;
      if (existingFrom) {
        selectedIdentity = identities.find(id => matchesIdentity(id.email, existingFrom))
          || identities.find(id => id.mayDelete === false) || identities[0];
      } else {
        selectedIdentity = identities.find(id => id.mayDelete === false) || identities[0];
      }
    }

    // Extract existing body values by MIME type (see bodyValueForType for the
    // MIME-match-not-list-presence rationale; bodies are immutable so we destroy+recreate).
    const bodyValues = existingEmail.bodyValues || {};
    const existingTextValue = this.bodyValueForType(existingEmail.textBody, 'text/plain', bodyValues);
    const existingHtmlValue = this.bodyValueForType(existingEmail.htmlBody, 'text/html', bodyValues);

    // Strict empty-reject + explicit clearFields. A provided-but-empty value is a
    // loud error (it's almost always an accidental clobber); deliberately blanking
    // a field is done by naming it in clearFields. Every field is clearable EXCEPT
    // `from` (identity-resolved; a draft always has a sender, matching the Fastmail UI).
    const CLEARABLE = new Set(['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments']); // NOT 'from'
    const SETTABLE = ['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'from'] as const;
    const provided = new Set<string>(SETTABLE.filter(f => (updates as any)[f] !== undefined));
    // `attachments` isn't a string SETTABLE field, so add it to `provided` explicitly
    // when an attachment add/remove was requested. This is the seam that makes
    // validateClearFields throw "can't set and clear attachments" if the caller also
    // passes clearFields:['attachments'] (clear-then-append in one call is ambiguous).
    if (updates.attachments?.length || updates.removeAttachments?.length) provided.add('attachments');
    validateClearFields(updates.clearFields, CLEARABLE, provided);
    const clear = new Set(updates.clearFields ?? []);

    // The `!clear.has(f)` guard is belt-and-suspenders: validateClearFields already
    // throws when a field is BOTH provided and cleared, so a cleared field can't also
    // reach these checks — the skip just keeps each check self-evidently correct.
    const clearHint = 'omit to leave it unchanged, or list it in clearFields to clear it';
    if (updates.subject  !== undefined && !clear.has('subject'))  requireNonEmpty(updates.subject,  'subject',  clearHint);
    if (updates.textBody !== undefined && !clear.has('textBody')) requireNonEmpty(updates.textBody, 'textBody', clearHint);
    if (updates.htmlBody !== undefined && !clear.has('htmlBody')) requireNonEmpty(updates.htmlBody, 'htmlBody', clearHint);
    if (updates.from     !== undefined) requireNonEmpty(updates.from, 'from'); // not clearable; no hint about clearFields
    for (const f of ['to', 'cc', 'bcc', 'replyTo'] as const) {
      if (updates[f] !== undefined && !clear.has(f) && updates[f]!.length === 0) {
        throw new Error(`${f} cannot be empty; ${clearHint}`);
      }
    }
    // Body requireNonEmpty calls above are GUARDS ONLY — their trimmed return is
    // discarded so stored bodies keep their exact (untrimmed) value below.

    // ---- Reply-quote preservation guard (#37, redesigned #42) ----
    // A reply draft keeps the quoted original in its body (buildReplyBodies appends a cited
    // <blockquote> to html and a "> "-quoted block to text). An edit that rewrites or clears
    // that body would silently drop the quote. We decide on the EXISTING (stored) body — which
    // THIS server generated, so its quote shape is reliable — never on the caller's NEW body
    // (untrusted: it can't tell the real quote from any quote-shaped content, and prose can
    // false-positive). When the draft has a quote and the edit touches the body in a way that
    // isn't quote-preserving by construction, force the caller to choose: regenerate+keep from
    // a caller-named originalEmailId, or deliberately noQuote. Supersedes the fork.8 new-body-
    // scan (bypassable + html-only); see docs/email-bodies.md and #42.
    const isReply = !!existingEmail.inReplyTo?.length;
    const oldHtmlQuoted = hasQuoteMarker(existingHtmlValue);
    const oldTextQuoted = hasTextQuoteMarker(existingTextValue);
    const draftHasQuote = isReply && (oldHtmlQuoted || oldTextQuoted);

    // Pre-merge signals: what the edit does to each body. existingHtmlValue/existingTextValue
    // were fetched above; these are the same inputs the coupling guards below use.
    const wroteHtml = updates.htmlBody !== undefined;
    const wroteText = updates.textBody !== undefined;
    const clearedHtml = clear.has('htmlBody');
    const clearedText = clear.has('textBody');
    const touchesBody = wroteHtml || wroteText || clearedHtml || clearedText;

    // The quote survives WITHOUT inspecting any new content in exactly two shapes:
    //  - a metadata-only edit (no body written or cleared) leaves both bodies untouched;
    //  - a plain-text conversion (clearFields:['htmlBody'] alone) keeps the old text, but only
    //    counts as quote-preserving when that surviving text actually carries the "> " quote
    //    (oldTextQuoted — always true for drafts this server made). If it doesn't, this is NOT
    //    a clean carve-out and the edit correctly falls through to the guard below.
    const quoteKeptByConstruction =
         !touchesBody
      || (clearedHtml && !wroteHtml && !wroteText && !clearedText && oldTextQuoted);

    // Text-side edits while a non-empty html survives are owned by the two coupling guards
    // further down (textBody-alone; clearFields:['textBody']-while-html), which emit the
    // correct remedy and (for guard ii) need the post-merge mergedHtmlRaw, so they can't move
    // up here. Exclude those cases so this pre-merge guard doesn't pre-empt them. On a text-
    // only draft existingHtmlValue is blank, so this is false → a text edit there correctly
    // falls through to the guard (exactly the #42 case this guard exists to catch).
    const coupledTextEdit =
      !wroteHtml && !clearedHtml && !isBlank(existingHtmlValue) && (wroteText || clearedText);

    if (draftHasQuote && touchesBody && !quoteKeptByConstruction && !coupledTextEdit) {
      if (updates.originalEmailId && updates.noQuote === true) {
        throw new Error('Pass either originalEmailId (keep the quote) or noQuote (discard it), not both.');
      } else if (updates.originalEmailId) {
        // Regenerate from the caller-named original — never re-resolved from the draft's
        // In-Reply-To (which is attacker-controllable), so there's no spoof surface. The id is
        // trusted, not validated against the draft's In-Reply-To (that check would false-reject
        // legitimate cases, e.g. correcting a wrong original). getEmailById throws on not-found;
        // rethrow with a message naming the param so the caller can fix it (it surfaces via
        // index's error wrap like this function's other guards).
        let original: any;
        try {
          original = await this.getEmailById(updates.originalEmailId);
        } catch {
          throw new Error(`originalEmailId '${updates.originalEmailId}' could not be fetched (no such message, or not accessible). Pass the id of the message this draft replies to.`);
        }
        // Regenerate the quote into EVERY body the edit is writing — both, when the caller
        // supplies both (a new html + a custom text alternative), so neither side silently
        // loses the quote on the keep path. buildReplyBodies quotes exactly the formats passed.
        // A clear-only edit writes neither body, so there's nowhere to regenerate into — reject
        // loudly rather than silently no-op the keep intent. This pre-empts the downstream
        // no-body reject for a clear-the-last-body edit (the caller sees the regenerate message,
        // not the no-body one); both are loud and lose no data, and the throw means no double-fire.
        if (wroteHtml || wroteText) {
          const rebuilt = buildReplyBodies({
            original,
            ...(wroteHtml && { htmlBody: updates.htmlBody }),
            ...(wroteText && { textBody: updates.textBody }),
            quoteOriginal: true,
          });
          if (wroteHtml) updates.htmlBody = rebuilt.htmlBody;
          if (wroteText) updates.textBody = rebuilt.textBody;
          // Loud-fail a self-inconsistent keep request: the caller asked to KEEP via
          // originalEmailId, but the named message has no quotable content (attachment-only /
          // calendar-only / cid-image-only), so buildReplyBodies passed the body through
          // unquoted. This is reachable only by naming the WRONG/empty original — a draft naming
          // its own original can't hit it (a quote exists only if that original was quotable, and
          // JMAP message content is immutable). It loses no caller input (the new body is kept);
          // it just turns a confusing quote-less result into an actionable error instead of a
          // silent one. The `||` accepts the edit if ANY written format kept a marker, so a
          // partially-quotable original still keeps.
          const restored = (wroteHtml && hasQuoteMarker(updates.htmlBody))
            || (wroteText && hasTextQuoteMarker(updates.textBody));
          if (!restored) {
            throw new Error(`originalEmailId '${updates.originalEmailId}' has no quotable content (e.g. an attachment-only or calendar-only message), so the quote can't be restored. Check the id, or use noQuote to drop the quote deliberately.`);
          }
        } else {
          throw new Error("originalEmailId can't regenerate a quote on a body you're not writing — edit the body (htmlBody or textBody) to keep the quote, or use noQuote to drop it.");
        }
      } else if (updates.noQuote === true) {
        // Proceed: the quote is dropped on explicit request.
      } else {
        // Error names ONLY the data-preserving keep path; noQuote is deliberately omitted so
        // the model is never nudged toward discarding the quote (it stays in the schema).
        throw new Error("Editing this reply draft's body would drop the quoted original. Pass originalEmailId (the message it replies to) to keep the quote. If you only have the draft, resolve the original from its In-Reply-To Message-ID via search_emails first.");
      }
    }

    // Merge non-body fields: updates override existing; clearFields force the empty value.
    const mergedSubject = clear.has('subject') ? '' : (updates.subject !== undefined ? updates.subject : (existingEmail.subject || ''));
    const mergedTo      = clear.has('to')      ? [] : (updates.to      !== undefined ? updates.to.map(parseAddress)      : (existingEmail.to || []));
    const mergedCc      = clear.has('cc')      ? [] : (updates.cc      !== undefined ? updates.cc.map(parseAddress)      : (existingEmail.cc || []));
    const mergedBcc     = clear.has('bcc')     ? [] : (updates.bcc     !== undefined ? updates.bcc.map(parseAddress)     : (existingEmail.bcc || []));
    const mergedReplyTo = clear.has('replyTo') ? [] : (updates.replyTo !== undefined ? updates.replyTo.map(parseAddress) : (existingEmail.replyTo || null));

    // ---- Body pipeline: one-sided guard + text-fallback generation ----
    // The text part is a DERIVED fallback when html is present. So:
    //  - editing htmlBody alone REGENERATES the text fallback from the new html (no throw);
    //  - editing textBody alone (while a non-empty html survives) is rejected — it won't
    //    change what most recipients render (the html), and the fallback is auto-managed;
    //  - a metadata-only edit (no body written) stays body-invariant (both bodies kept).
    // wroteText/wroteHtml are computed once at the reply-quote guard above (same values; the
    // originalEmailId path only ever replaces an already-written body with another, so their
    // truth doesn't change). Reuse them here.
    const wroteAnyBody = wroteText || wroteHtml;

    // Raw merge: a written body drops the unwritten partner (single-format intent);
    // a no-body edit preserves both; clearFields force the body absent.
    const mergedTextRaw = clear.has('textBody') ? undefined
      : (updates.textBody !== undefined ? updates.textBody
      : (wroteAnyBody ? undefined : existingTextValue));
    const mergedHtmlRaw = clear.has('htmlBody') ? undefined
      : (updates.htmlBody !== undefined ? updates.htmlBody
      : (wroteAnyBody ? undefined : existingHtmlValue));

    // Guard: editing textBody alone while a non-empty htmlBody survives (checked against
    // the EXISTING html, since the raw merge has already dropped the unwritten partner).
    if (wroteText && !wroteHtml && !clear.has('htmlBody') && !isBlank(existingHtmlValue)) {
      throw new Error('editing textBody alone won\'t change what most recipients see (they render htmlBody). To change the message, edit htmlBody (the text fallback regenerates automatically); to save a custom plain-text alternative, supply htmlBody alongside it; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }

    // Guard: clearFields:['textBody'] while htmlBody survives — the text fallback is
    // managed automatically (regenerated from html, or html-only if none is derivable), so
    // clearing it on its own is rejected. Evaluated against the MERGED html and BEFORE the
    // fallback step runs (else that step would silently refill it). Allowed when html is also cleared.
    if (clear.has('textBody') && !clear.has('htmlBody') && !isBlank(mergedHtmlRaw)) {
      throw new Error('textBody can\'t be cleared on its own while htmlBody is present — the text fallback is managed automatically (regenerated from htmlBody, or html-only if none can be derived). Omit textBody from clearFields; or use clearFields:[\'htmlBody\'] to make this a plain-text email.');
    }

    // Generate the text fallback, but ONLY when a body was actually written — a
    // metadata-only edit must stay body-invariant (it must NOT inject a text part into an
    // html-only draft). When html was (re)written without text, regenerate the text fallback
    // from the new html; ship html-only if none is derivable but the html has visible
    // content; reject a genuinely no-body result.
    let textBodyValue = mergedTextRaw;
    let htmlBodyValue = mergedHtmlRaw;
    if (wroteAnyBody) {
      const normalized = normalizeBodies({ textBody: mergedTextRaw, htmlBody: mergedHtmlRaw });
      textBodyValue = normalized.textBody;
      htmlBodyValue = normalized.htmlBody;
      if (normalized.htmlOnly && !htmlHasVisibleContent(mergedHtmlRaw!)) {
        throw new Error('This message has no readable body; add text or visible content.');
      }
    }

    // Reject a body-less RESULT, but only when this edit actually touched the body —
    // a written body that came out empty, or a cleared body. An attachment-only (or
    // any metadata-only) edit must NOT trip this: it stays body-invariant and may run
    // against a draft that legitimately has no body yet. Gating on
    // `wroteAnyBody || clearedAnyBody` (not `wroteAnyBody` alone) keeps the throw firing
    // when the last body is cleared — incl. alongside an attachments change — so a
    // caller can't silently strip a draft down to no body. A draft keeps >=1 body.
    // Distinct from the clear-text-while-html guard (which only fires when merged html IS
    // present), so the two can't both match.
    const clearedAnyBody = clear.has('textBody') || clear.has('htmlBody');
    if ((wroteAnyBody || clearedAnyBody) && isBlank(textBodyValue) && isBlank(htmlBodyValue)) {
      throw new Error('a draft needs a body; supply textBody or htmlBody (this edit would leave it with neither).');
    }

    // Carry existing (non-inline) attachments by referencing their existing blobIds.
    // Whitelist exactly these fields — a blob-backed part is blobId XOR partId, and
    // `size` is server-set, so sending partId/size would be rejected by a strict server.
    const carriedAttachments: AttachmentPart[] = existingAttachments.map((a: any) => ({
      blobId: a.blobId,
      type: a.type,
      ...(a.name != null && { name: a.name }),
      ...(a.disposition != null && { disposition: a.disposition }),
      ...(a.cid != null && { cid: a.cid }),
    }));

    // Build the final attachment set: clear-all empties it; each removeAttachments ref
    // drops a carried part by blobId (the stable ref the caller sees), or by a UNIQUE
    // non-null name as a convenience; then the freshly uploaded parts are appended.
    // A ref that matches nothing, or an ambiguous name, is rejected loudly rather than
    // silently no-op'd (a silent no-match is the confident-wrong-result class this
    // codebase rejects). `attachments` + clearFields:['attachments'] was already rejected
    // as a conflict above, so clear-all never coexists with an append/remove here.
    let finalAttachments: AttachmentPart[] = clear.has('attachments') ? [] : carriedAttachments.slice();
    if (updates.removeAttachments?.length) {
      for (const ref of updates.removeAttachments) {
        const beforeLen = finalAttachments.length;
        const byBlob = finalAttachments.filter(a => a.blobId !== ref);
        if (byBlob.length < beforeLen) {
          finalAttachments = byBlob;
          continue;
        }
        const nameMatches = finalAttachments.filter(a => a.name != null && a.name === ref);
        if (nameMatches.length === 1) {
          finalAttachments = finalAttachments.filter(a => a !== nameMatches[0]);
          continue;
        }
        if (nameMatches.length > 1) {
          throw new PathAccessError(
            `removeAttachments ref '${ref}' matches ${nameMatches.length} attachments by name; pass the blobId instead (one of: ${finalAttachments.map(a => a.blobId).join(', ')}).`
          );
        }
        throw new PathAccessError(
          `removeAttachments ref '${ref}' matched no attachment on this draft. Carried blobIds: ${carriedAttachments.map(a => a.blobId).join(', ') || '(none)'}.`
        );
      }
    }
    if (updates.attachments?.length) {
      finalAttachments = finalAttachments.concat(updates.attachments);
    }

    const emailObject: any = {
      mailboxIds: existingEmail.mailboxIds,
      // Preserve all existing keywords (e.g. $flagged, custom labels), not just $draft.
      keywords: { ...(existingEmail.keywords || {}), $draft: true },
      from: [{ name: selectedIdentity.name, email: updates.from || existingEmail.from?.[0]?.email || selectedIdentity.email }],
      to: mergedTo,
      cc: mergedCc,
      bcc: mergedBcc,
      subject: mergedSubject,
      ...(mergedReplyTo?.length && { replyTo: mergedReplyTo }),
      // Threading: carry inReplyTo/references as JMAP structured properties so the
      // In-Reply-To/References headers regenerate (fixes silent threading loss on reply
      // drafts this client creates via reply_email send=false).
      ...(existingEmail.inReplyTo && { inReplyTo: existingEmail.inReplyTo }),
      ...(existingEmail.references && { references: existingEmail.references }),
      ...(finalAttachments.length && { attachments: finalAttachments }),
    };

    Object.assign(emailObject, buildBodyParts({ textBody: textBodyValue, htmlBody: htmlBodyValue }));

    // Create-then-delete (NOT a single combined create+destroy call). JMAP content is
    // immutable — verified live 2026-06-24 that Fastmail SILENTLY NO-OPS an in-place
    // subject/body update (returns success but changes nothing), so a recreate is
    // mandatory. RFC 8620 §6.3 guarantees blob lifetime within a call but says NOTHING
    // about create/destroy atomicity: a server MAY apply the destroy even when the create
    // lands in notCreated, which would vanish the draft. So we create FIRST, confirm it
    // succeeded, and only THEN destroy the old draft. Worst case is a harmless duplicate
    // (recoverable), never a vanished draft (unrecoverable).
    // WARNING: do NOT "optimize" this back into one Email/set call; that reintroduces the
    // data-loss window.
    const createRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          create: { draft: emailObject },
        }, 'createDraft']
      ]
    };

    const createResponse = await this.makeRequest(createRequest);
    const createResult = this.getMethodResult(createResponse, 0);

    if (createResult.notCreated?.draft) {
      const err = createResult.notCreated.draft;
      // Create failed → old draft is untouched (no destroy was issued). This is the
      // data-loss-prevention path: surface the error, leave the draft as-is.
      throw new Error(`Failed to create updated draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const newEmailId = createResult.created?.draft?.id;
    if (!newEmailId) {
      throw new Error('Draft update returned no email ID');
    }

    // The new draft is valid → the edit has SUCCEEDED. Now remove the old copy. Any
    // failure here (structured notDestroyed OR a thrown transport/method error) leaves a
    // harmless duplicate holding the OLD pre-edit content — report it as an orphan
    // warning, but do NOT throw (throwing would tell the caller the edit failed when it
    // didn't).
    let orphanedOldDraftId: string | undefined;
    try {
      const destroyResponse = await this.makeRequest({
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/set', { accountId: session.accountId, destroy: [emailId] }, 'destroyDraft']
        ]
      });
      const destroyResult = this.getMethodResult(destroyResponse, 0);
      if (destroyResult.notDestroyed?.[emailId]) {
        orphanedOldDraftId = emailId;
      }
    } catch {
      orphanedOldDraftId = emailId;
    }

    return { id: newEmailId, ...(orphanedOldDraftId && { orphanedOldDraftId }) };
  }

  async sendDraft(emailId: string): Promise<string> {
    const session = await this.getSession();

    // Fetch the existing email to verify it's a draft
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['id', 'from', 'to', 'cc', 'bcc', 'replyTo', 'keywords', 'textBody', 'htmlBody', 'bodyValues'],
          bodyProperties: ['partId', 'blobId', 'type', 'size'],
          fetchTextBodyValues: true,
          fetchHTMLBodyValues: true,
        }, 'getEmail']
      ]
    };

    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];
    if (!email) {
      throw new Error(`Email with ID '${emailId}' not found`);
    }

    if (!email.keywords?.$draft) {
      throw new Error('Cannot send a non-draft email');
    }

    // Reject an empty body part before an irreversible send. sendDraft submits the draft by
    // reference WITHOUT recreating it, so unlike edit_draft it can't truthy-gate the body away.
    // An empty/whitespace body part renders blank; an empty text/html part even shadows a real
    // text/plain (RFC 2046: clients render the richest alternative), so the recipient sees
    // nothing. We never emit such a part, but an externally-created draft (Fastmail web UI, etc.)
    // can carry one — refuse to ship it. Reject, not silent sanitize: recreating to strip the part
    // would change the email id and rewrite the message, against this codebase's loud-reject
    // philosophy. The hardened edit_draft is the fix path.
    const textVal = this.bodyValueForType(email.textBody, 'text/plain', email.bodyValues || {});
    const htmlVal = this.bodyValueForType(email.htmlBody, 'text/html', email.bodyValues || {});
    if (htmlVal !== undefined && htmlVal.trim() === '') {
      throw new Error('This draft has an empty htmlBody that would render blank to recipients. Edit the draft to supply or clear htmlBody before sending.');
    }
    if (textVal !== undefined && textVal.trim() === '') {
      throw new Error('This draft has an empty textBody that would render blank for plain-text recipients. Edit the draft to supply or clear textBody before sending.');
    }

    // Collect all recipients for the envelope
    const allRecipients: { email: string }[] = [
      ...(email.to || []),
      ...(email.cc || []),
      ...(email.bcc || []),
    ];

    if (allRecipients.length === 0) {
      throw new Error('Draft has no recipients');
    }

    // Determine identity from the email's from field
    const fromEmail = email.from?.[0]?.email;
    if (!fromEmail) {
      throw new Error('Draft has no from address');
    }

    const identities = await this.getIdentities();
    const selectedIdentity = identities.find(id => matchesIdentity(id.email, fromEmail));
    if (!selectedIdentity) {
      throw new Error('From address on draft does not match any sending identity');
    }

    // Find the Sent mailbox
    const mailboxes = await this.getMailboxes();
    const sentMailbox = this.findMailboxByRoleOrName(mailboxes, 'sent', 'sent');
    if (!sentMailbox) {
      throw new Error('Could not find Sent mailbox');
    }

    const sentMailboxIds: Record<string, boolean> = {};
    sentMailboxIds[sentMailbox.id] = true;

    // Submit the draft
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', 'urn:ietf:params:jmap:submission'],
      methodCalls: [
        ['EmailSubmission/set', {
          accountId: session.accountId,
          create: {
            submission: {
              emailId,
              identityId: selectedIdentity.id,
              envelope: {
                mailFrom: { email: fromEmail },
                rcptTo: allRecipients.map(addr => ({ email: addr.email })),
              }
            }
          },
          onSuccessUpdateEmail: {
            '#submission': {
              mailboxIds: sentMailboxIds,
              'keywords/$draft': null,
              'keywords/$seen': true,
            }
          }
        }, 'submitDraft']
      ]
    };

    const response = await this.makeRequest(request);
    const submissionResult = this.getMethodResult(response, 0);
    if (submissionResult.notCreated?.submission) {
      const err = submissionResult.notCreated.submission;
      throw new Error(`Failed to submit draft: ${err.type}${err.description ? ' - ' + err.description : ''}`);
    }

    const submissionId = submissionResult.created?.submission?.id;
    if (!submissionId) {
      throw new Error('Draft submission returned no submission ID');
    }

    return submissionId;
  }

  async getRecentEmails(limit: number = 10, mailbox: string = 'inbox', ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    // Resolve the target mailbox EXACTLY (id/role/name) — replaces the old substring
    // match, so this stays consistent with the #12 sweep and carries no substring
    // injection-steering primitive. A blank/whitespace mailbox falls back to the inbox
    // default (matching the resolveMailboxId blank handling the swept tools use), rather
    // than throwing. Throws InvalidInputError on a non-blank unknown mailbox.
    const target = mailbox && mailbox.trim() ? mailbox : 'inbox';
    const mailboxes = await this.getMailboxes();
    const targetMailbox = resolveMailbox(mailboxes, target);

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
    };

    emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter: { inMailbox: targetMailbox.id },
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit: Math.min(limit, 50),
          calculateTotal: true
        }, 'query'],
        ['Email/get', emailGetParams, 'emails']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getQueryResult(response, 0, 1);
    // Reuse the mailbox list already fetched above to resolve names + roles — no extra methodCall.
    attachMailboxInfo(result.items, buildMailboxInfoMap(mailboxes));
    return result;
  }

  async markEmailRead(emailId: string, read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = read
      ? { 'keywords/$seen': true }
      : { 'keywords/$seen': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'updateEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to mark email as ${read ? 'read' : 'unread'}.`);
    }
  }

  async pinEmail(emailId: string, pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const update: Record<string, any> = {};
    update[emailId] = pinned
      ? { 'keywords/$flagged': true }
      : { 'keywords/$flagged': null };

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update
        }, 'pinEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error(`Failed to ${pinned ? 'pin' : 'unpin'} email.`);
    }
  }

  async deleteEmail(emailId: string): Promise<void> {
    const session = await this.getSession();

    // Find the trash mailbox by EXACT role only (case-insensitive). NOT the substring
    // findMailboxByRoleOrName: a custom "Trash bin rules" mailbox (no trash role) must
    // never be the delete destination, and computeExclusion's exact-role Trash would
    // then never count mail mis-filed there.
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findByExactRole(mailboxes, 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: {
              mailboxIds: trashMailboxIds
            }
          }
        }, 'moveToTrash']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to delete email.');
    }
  }

  async moveEmail(emailId: string, target: string): Promise<void> {
    const session = await this.getSession();

    // Resolve the destination EXACTLY (id/role/name) — a new capability (moveEmail
    // previously took a raw id with no resolution). Exact-only, so it carries no
    // substring mis-resolution; deliberate move-to-any-mailbox stays open by design
    // (a move-target restriction is tracked separately, fork #43). Throws
    // InvalidInputError on an unknown destination.
    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // Fetch current mailboxIds to build a proper JMAP patch
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['mailboxIds']
        }, 'getEmail']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const email = this.getListResult(getResponse, 0)[0];

    // Build patch: remove from all current mailboxes, add to target
    const patch: Record<string, boolean | null> = {};
    if (email?.mailboxIds) {
      for (const mbId of Object.keys(email.mailboxIds)) {
        patch[`mailboxIds/${mbId}`] = null;
      }
    }
    patch[`mailboxIds/${targetMailboxId}`] = true;

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'moveEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to move email.');
    }
  }

  // The label tools take mailbox IDs only (no name/role resolution — full name
  // resolution there is tracked as fork #50). Reject any element that isn't a real
  // mailbox id BEFORE the Email/set, so a caller who learned `mailbox:"Archive"` works
  // elsewhere gets a guided error here instead of a silent no-op or a cryptic JMAP
  // failure. "Valid" = matches some mailbox.id in the fetched list (a real id absent
  // from the list is rejected too — accepted residual, see docs/security-model.md).
  private async assertValidMailboxIds(mailboxIds: string[]): Promise<void> {
    const mailboxes = await this.getMailboxes();
    const validIds = new Set((mailboxes || []).map((mb: any) => mb.id));
    const invalid = mailboxIds.filter(id => !validIds.has(id));
    if (invalid.length > 0) {
      throw new InvalidInputError(
        `Not valid mailbox id(s): ${invalid.join(', ')}. The label tools accept mailbox IDs only (not names or roles) — use list_mailboxes to resolve a name to its id.`,
      );
    }
  }

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    await this.assertValidMailboxIds(mailboxIds);

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'addLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to add labels to email.');
    }
  }

  async removeLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    await this.assertValidMailboxIds(mailboxIds);

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: {
            [emailId]: patch
          }
        }, 'removeLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && result.notUpdated[emailId]) {
      throw new Error('Failed to remove labels from email.');
    }
  }

  async bulkAddLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    await this.assertValidMailboxIds(mailboxIds);

    // Build patch object to add specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = true;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkAddLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to add labels to some emails.');
    }
  }

  async bulkRemoveLabels(emailIds: string[], mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();
    await this.assertValidMailboxIds(mailboxIds);

    // Build patch object to remove specific mailboxIds
    const patch: Record<string, any> = {};
    mailboxIds.forEach(mailboxId => {
      patch[`mailboxIds/${mailboxId}`] = null;
    });

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkRemoveLabels']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to remove labels from some emails.');
    }
  }

  async getEmailAttachments(emailId: string): Promise<any[]> {
    const session = await this.getSession();

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['attachments']
        }, 'getAttachments']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];
    return email?.attachments || [];
  }

  async downloadAttachment(emailId: string, attachmentId: string): Promise<string> {
    const session = await this.getSession();

    // Get the email with full attachment details
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: [emailId],
          properties: ['attachments', 'bodyValues'],
          bodyProperties: ['partId', 'blobId', 'size', 'name', 'type']
        }, 'getEmail']
      ]
    };

    const response = await this.makeRequest(request);
    const email = this.getListResult(response, 0)[0];

    if (!email) {
      throw new Error('Email not found');
    }

    // Find attachment by partId or by index
    let attachment = email.attachments?.find((att: any) => 
      att.partId === attachmentId || att.blobId === attachmentId
    );

    // If not found, try by array index
    if (!attachment) {
      const index = parseInt(attachmentId, 10);
      if (!isNaN(index)) {
        attachment = email.attachments?.[index];
      }
    }
    
    if (!attachment) {
      throw new Error('Attachment not found.');
    }

    // Get the download URL from session
    const downloadUrl = session.downloadUrl;
    if (!downloadUrl) {
      throw new Error('Download capability not available in session');
    }

    // Build download URL
    const url = downloadUrl
      .replace('{accountId}', session.accountId)
      .replace('{blobId}', attachment.blobId)
      .replace('{type}', encodeURIComponent(attachment.type || 'application/octet-stream'))
      .replace('{name}', encodeURIComponent(attachment.name || 'attachment'));

    return url;
  }

  static readonly DEFAULT_DOWNLOADS_DIR = resolve(homedir(), 'Downloads', 'fastmail-mcp');

  static validateSavePath(savePath: string, downloadDir?: string): string {
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;
    // Resolve relative paths against the allowed download directory rather than
    // the process cwd (which is unpredictable for an MCP server launched by a
    // client). Absolute paths are taken as-is; either way the containment check
    // is the security boundary. So a bare filename lands safely in the configured
    // dir in one step, and an absolute path inside that dir writes exactly there.
    // Write side keeps a byte-exact (case-sensitive) compare so download behaviour
    // is unchanged; the read guard opts into case-insensitive containment on Win32.
    // Throws PathAccessError so the index layer maps it to InvalidParams uniformly.
    return lexicalContainedPath(savePath, allowedDir, false);
  }

  /**
   * Symlink-safe canonicalization of a save path. Walks up to the longest
   * existing ancestor, realpaths it, and verifies it lives under the canonical
   * allowed directory. Refuses to overwrite an existing symlink at the target.
   *
   * Returns the canonical path that is safe to write to. Throws on escape.
   */
  static async safeWritePath(savePath: string, downloadDir?: string): Promise<string> {
    // Lexical pre-check first (cheap and gives nice errors)
    const lexical = JmapClient.validateSavePath(savePath, downloadDir);
    const allowedDir = downloadDir ? resolve(normalize(downloadDir)) : JmapClient.DEFAULT_DOWNLOADS_DIR;

    // Ensure allowed dir exists so realpath can resolve it.
    await mkdir(allowedDir, { recursive: true });
    const canonicalAllowed = await realpath(allowedDir);

    // Walk up from the target until we find an existing ancestor.
    let ancestor = dirname(lexical);
    const missingSegments: string[] = [];
    while (true) {
      try {
        await stat(ancestor);
        break;
      } catch (e: any) {
        if (e.code !== 'ENOENT') throw e;
        missingSegments.unshift(basename(ancestor));
        const parent = dirname(ancestor);
        if (parent === ancestor) {
          throw new PathAccessError(`Could not find an existing ancestor for path: ${lexical}`);
        }
        ancestor = parent;
      }
    }

    // Canonicalize the existing ancestor — this is what catches symlink escapes.
    const canonicalAncestor = await realpath(ancestor);
    if (canonicalAncestor !== canonicalAllowed && !canonicalAncestor.startsWith(canonicalAllowed + sep)) {
      throw new PathAccessError(
        `path resolves to '${canonicalAncestor}' which is outside the allowed directory '${canonicalAllowed}'. ` +
        `Refusing to follow symlink escape.`,
      );
    }

    // Reconstruct the safe canonical path under the canonical ancestor.
    const safePath = join(canonicalAncestor, ...missingSegments, basename(lexical));

    // If a symlink already exists at the target, refuse — writing through it
    // would still escape the allowed directory.
    try {
      const lst = await lstat(safePath);
      if (lst.isSymbolicLink()) {
        throw new PathAccessError(`Refusing to overwrite an existing symlink at the target: ${safePath}`);
      }
    } catch (e: any) {
      if (e.code !== 'ENOENT') throw e;
    }

    return safePath;
  }

  // Per-file / aggregate fail-fast guards. NOT authoritative — Fastmail's own ceiling
  // governs; these just bound the in-memory read and reject obviously-too-large inputs
  // before we upload. The per-file cap also bounds the fd read in uploadAttachments.
  static readonly MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
  static readonly MAX_TOTAL_ATTACHMENT_BYTES = 45 * 1024 * 1024;

  /**
   * Read-shaped, handle-based path confinement for the attachment-send capability.
   * Distinct from the write-shaped safeWritePath (which mkdir -p's the root and walks
   * MISSING segments) — a read creates nothing and must validate the OPEN file:
   *
   *  - attachDir undefined → throw the opt-in error BEFORE any fs syscall (the hard
   *    gate; an exfiltration capability stays disabled until the operator sets the var);
   *  - reject the Windows escape shapes and lexically contain the path (case-insensitively
   *    on Win32) against the resolved root;
   *  - open(path,'r') ONCE, then fstat the handle (require a regular file) and realpath
   *    the FULL target (not an ancestor), re-verifying canonical containment. The caller
   *    reads from the returned handle, so the bytes uploaded are the bytes of the file we
   *    validated — TOCTOU is narrowed, not eliminated (see docs/security-model.md).
   *
   * Returns the open handle and its size; the CALLER must close the handle.
   */
  static async safeReadPath(inputPath: string, attachDir: string | undefined): Promise<{ handle: FileHandle; size: number }> {
    if (!attachDir) {
      throw new PathAccessError(
        'Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR to the directory attachable files live in, then restart the server to enable it.'
      );
    }

    rejectWindowsPathEscapes(inputPath);

    const allowedDir = resolve(normalize(attachDir));
    const caseInsensitive = process.platform === 'win32';
    const lexical = lexicalContainedPath(inputPath, allowedDir, caseInsensitive);

    // The attach root itself must exist — a missing root is a config error, reported
    // distinctly from the opt-in gate above (not a raw realpath ENOENT).
    let canonicalAllowed: string;
    try {
      canonicalAllowed = await realpath(allowedDir);
    } catch (e: any) {
      if (e.code === 'ENOENT') {
        throw new PathAccessError(`FASTMAIL_ATTACH_DIR (${allowedDir}) does not exist. Create it or fix the path, then restart.`);
      }
      throw e;
    }

    let handle: FileHandle;
    try {
      handle = await open(lexical, 'r');
    } catch (e: any) {
      if (e.code === 'ENOENT') {
        throw new PathAccessError(`File not found: ${inputPath} (resolved under ${allowedDir}).`);
      }
      if (e.code === 'EISDIR') {
        throw new PathAccessError(`Not a regular file: ${inputPath}.`);
      }
      throw e;
    }

    try {
      const st = await handle.stat();
      if (!st.isFile()) {
        throw new PathAccessError(`Not a regular file: ${inputPath}.`);
      }
      // Re-verify against the canonical full target — this catches a symlinked leaf or
      // an intermediate-dir symlink that escapes the root.
      const canonicalTarget = await realpath(lexical);
      if (!isPathContained(canonicalTarget, canonicalAllowed, caseInsensitive)) {
        throw new PathAccessError(
          `path resolves to '${canonicalTarget}' which is outside the allowed directory '${canonicalAllowed}'. Refusing to follow symlink escape.`
        );
      }
      return { handle, size: st.size };
    } catch (e) {
      await handle.close().catch(() => {});
      throw e;
    }
  }

  /**
   * Upload a single blob and return its server-assigned blobId. POSTs the raw bytes to
   * the {accountId}-substituted session uploadUrl with ONLY Authorization + the given
   * Content-Type — deliberately NOT spreading getAuthHeaders() (which hardcodes
   * application/json) and NOT JSON.stringifying the body (unlike every other call here).
   * The server-returned `type` is authoritative for the stored blob (it echoes the
   * Content-Type we sent — a best-effort hint, not content sniffing).
   */
  async uploadBlob(data: Buffer, contentType: string): Promise<{ blobId: string; type: string; size: number }> {
    const session = await this.getSession();
    if (!session.uploadUrl) {
      throw new Error('Upload capability not available in session');
    }

    const url = session.uploadUrl.replace('{accountId}', session.accountId);
    // POST the raw bytes — no JSON.stringify, unlike every other call here. The copy
    // constructor `new Uint8Array(data)` yields a concrete Uint8Array<ArrayBuffer> (not the
    // ArrayBufferLike-backed view a Buffer/`.subarray` carries), which IS assignable to
    // fetch's BodyInit — so this stays fully type-checked, no `any` escape hatch.
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': this.auth.getAuthHeaders()['Authorization'],
        'Content-Type': contentType,
      },
      body: new Uint8Array(data),
    });

    if (!response.ok) {
      throw new Error(`Blob upload failed: ${response.status} ${response.statusText}`);
    }

    const result = await response.json() as any;
    if (!result || typeof result.blobId !== 'string') {
      throw new Error('Blob upload returned no blobId');
    }
    // End-to-end integrity check: the stored blob size the server reports back must equal
    // the byte count we sent. A mismatch means a truncated/corrupted upload (proxy or
    // transport), so fail loudly rather than attach a corrupt file. (Only checked when the
    // server returns a numeric size.)
    if (typeof result.size === 'number' && result.size !== data.length) {
      throw new Error(`Blob upload size mismatch: sent ${data.length} bytes, server stored ${result.size}.`);
    }
    return { blobId: result.blobId, type: result.type || contentType, size: result.size ?? data.length };
  }

  /**
   * Confine, read, and upload each file spec, returning the JMAP attachment parts to
   * splat into an Email. Re-checks the opt-in gate (so a caller that skipped safeReadPath
   * can't bypass it), validates the caller contentType against the MIME token grammar,
   * rejects an over-cap file via fstat.size BEFORE reading, then does a bounded read from
   * the confined handle. Each part is a FRESH 4-key literal (NOT the carriedAttachments
   * shape, which passes through size/cid that a strict server rejects or that would
   * mislabel a plain file).
   */
  async uploadAttachments(
    specs: { path: string; name?: string; contentType?: string }[],
    attachDir: string | undefined,
  ): Promise<AttachmentPart[]> {
    if (!attachDir) {
      throw new PathAccessError(
        'Sending attachments is disabled. Set FASTMAIL_ATTACH_DIR to the directory attachable files live in, then restart the server to enable it.'
      );
    }

    // Two passes so a confinement/size failure orphans NO blobs. Pass 1 validates and
    // opens every file (path confinement + per-file/total size caps) before a single
    // upload; a bad path or oversize file anywhere in the batch rejects with zero blobs
    // uploaded. Pass 2 reads + uploads the already-validated handles. (A network failure
    // mid-upload can still orphan a blob — unavoidable without server-side transactions;
    // Fastmail garbage-collects unreferenced blobs.)
    const opened: { handle: FileHandle; size: number; contentType: string; name: string }[] = [];
    try {
      let totalBytes = 0;
      for (let i = 0; i < specs.length; i++) {
        const spec = specs[i];
        // contentType grammar is validated before any fs work (no handle to leak yet).
        const contentType = spec.contentType
          ? validateContentType(spec.contentType, i)
          : guessContentType(spec.path);
        const { handle, size } = await JmapClient.safeReadPath(spec.path, attachDir);
        // Push BEFORE the size checks so the finally closes this handle even if a cap throws.
        opened.push({ handle, size, contentType, name: spec.name ?? basename(spec.path) });
        if (size > JmapClient.MAX_ATTACHMENT_BYTES) {
          throw new PathAccessError(
            `attachments[${i}] (${basename(spec.path)}) is ${size} bytes, over the ${JmapClient.MAX_ATTACHMENT_BYTES}-byte per-file guard. Fastmail's own limit ultimately governs.`
          );
        }
        totalBytes += size;
        if (totalBytes > JmapClient.MAX_TOTAL_ATTACHMENT_BYTES) {
          throw new PathAccessError(
            `attachments total exceeds the ${JmapClient.MAX_TOTAL_ATTACHMENT_BYTES}-byte fail-fast guard. Fastmail's own limit ultimately governs.`
          );
        }
      }

      const parts: AttachmentPart[] = [];
      for (const o of opened) {
        // Bounded read of exactly `size` bytes from the validated handle (never read the
        // whole handle then check .length — that buffers a hostile oversize file first).
        const buffer = Buffer.alloc(o.size);
        const { bytesRead } = await o.handle.read(buffer, 0, o.size, 0);
        const data = bytesRead === o.size ? buffer : buffer.subarray(0, bytesRead);

        const uploaded = await this.uploadBlob(data, o.contentType);
        parts.push({
          blobId: uploaded.blobId,
          type: uploaded.type,
          name: o.name,
          disposition: 'attachment',
        });
      }
      return parts;
    } finally {
      for (const o of opened) await o.handle.close().catch(() => {});
    }
  }

  async downloadAttachmentToFile(emailId: string, attachmentId: string, savePath: string, downloadDir?: string): Promise<{ url: string; bytesWritten: number; savedPath: string }> {
    const safePath = await JmapClient.safeWritePath(savePath, downloadDir);
    const url = await this.downloadAttachment(emailId, attachmentId);

    const response = await fetch(url, {
      headers: { 'Authorization': this.auth.getAuthHeaders()['Authorization'] }
    });

    if (!response.ok) {
      throw new Error(`Download failed: ${response.status} ${response.statusText}`);
    }

    const buffer = Buffer.from(await response.arrayBuffer());

    await mkdir(dirname(safePath), { recursive: true });
    await writeFile(safePath, buffer);

    return { url, bytesWritten: buffer.length, savedPath: safePath };
  }

  // Shared engine for searchEmails + getEmails: assemble the filter from a flat base
  // FilterCondition plus a list of single-keyword condition objects, inject the default
  // Trash/Spam exclusion, run the visible query + a hidden-count query, and populate
  // QueryResult.exclusion. Both callers fetch `mailboxes` themselves (to resolve their
  // `mailbox` param + compute the exclusion) and pass it in, so names attach with no
  // extra round-trip and there is no in-batch Mailbox/get.
  private async runFilteredQuery(opts: {
    base: any;
    conds: any[];
    exclusion: ExclusionResult;
    exclusionIntended: boolean;
    limit: number;
    ascending: boolean;
    mailboxes: any[];
  }): Promise<QueryResult> {
    const session = await this.getSession();
    const { base, conds, exclusion, exclusionIntended, limit, ascending, mailboxes } = opts;

    const doExclude = exclusion.excludeIds.length > 0;
    // Inject the exclusion into `base` BEFORE computing baseEmpty — otherwise an
    // exclusion-only query (no text/from fields) would see base as {} and take the
    // conds[0]-alone branch, silently dropping the folder exclusion (fail-open).
    if (doExclude) base.inMailboxOtherThan = exclusion.excludeIds;

    // Combine the base FilterCondition with N single-keyword conditions. Each keyword
    // is its own condition object because a single JMAP FilterCondition allows only one
    // hasKeyword/notKeyword. baseEmpty alone + one cond -> the lone cond; else AND-wrap.
    const combine = (b: any, c: any[], bEmpty: boolean) =>
      c.length === 0 ? b
      : (bEmpty && c.length === 1) ? c[0]
      : { operator: 'AND', conditions: [...(bEmpty ? [] : [b]), ...c] };

    const baseEmpty = Object.keys(base).length === 0;
    const visibleFilter = combine(base, conds, baseEmpty);

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
      properties: [...EMAIL_PROPERTIES_COMPACT],
    };

    const methodCalls: [string, any, string][] = [
      ['Email/query', {
        accountId: session.accountId,
        filter: visibleFilter,
        sort: [{ property: 'receivedAt', isAscending: ascending }],
        limit,
        calculateTotal: true,
      }, 'query'],
      ['Email/get', emailGetParams, 'emails'],
    ];

    if (doExclude) {
      // Hidden-count query = the visible filter with ONLY inMailboxOtherThan removed, so
      // hidden = broaderTotal - visibleTotal = matches withheld to Trash/Spam (the
      // complement, which never overcounts a message cross-filed in {Trash, a visible
      // mailbox} — that message is in both totals). Reconstruct from a COPY of base
      // minus the key, then re-run the identical combine: a naive top-level delete on
      // the assembled filter would no-op when the key sits inside conditions[0] under
      // the AND-wrap (count == visible -> note never fires, fail-open). Issued in the
      // SAME makeRequest at a higher index: one atomic snapshot (no two-query race) and
      // one fewer round-trip; the visible indices 0/1 are unchanged.
      const countBase = { ...base };
      delete countBase.inMailboxOtherThan;
      const countBaseEmpty = Object.keys(countBase).length === 0;
      const countFilter = combine(countBase, conds, countBaseEmpty);
      methodCalls.push(['Email/query', {
        accountId: session.accountId,
        filter: countFilter,
        limit: 0,
        calculateTotal: true,
      }, 'count']);
    }

    const response = await this.makeRequest({
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls,
    });
    const result = this.getQueryResult(response, 0, 1);
    attachMailboxInfo(result.items, buildMailboxInfoMap(mailboxes));

    // Populate exclusion metadata whenever an exclusion was INTENDED — even if
    // excludeIds came back empty (a role couldn't be resolved), so the handler still
    // fires the fail-loud "not excluded" note rather than silently running unfiltered.
    if (exclusionIntended) {
      let hidden: number | null = 0;
      if (doExclude) {
        // FAIL-CLOSED: the published "no note => nothing hidden" contract is only safe
        // if a missing/garbled count fails loud. calculateTotal is server-discretionary
        // and the count methodCall can error. If either total is non-numeric, or the
        // count method errored, or hidden computes negative (a wrong total:0 on the
        // broader query), set hidden=null (degraded note) — never clamp to 0.
        const visibleTotal = result.total;
        let broaderTotal: number | undefined;
        try { broaderTotal = this.getMethodResult(response, 2)?.total; } catch { broaderTotal = undefined; }
        if (typeof visibleTotal !== 'number' || typeof broaderTotal !== 'number') {
          hidden = null;
        } else {
          const h = broaderTotal - visibleTotal;
          hidden = h < 0 ? null : h;
        }
      }
      result.exclusion = {
        hidden,
        excludedRoles: exclusion.excludedRoles,
        unresolvedRoles: exclusion.unresolvedRoles,
      };
    }

    return result;
  }

  async searchEmails(filters: {
    query?: string;
    from?: string;
    to?: string;
    cc?: string;
    bcc?: string;
    subject?: string;
    hasAttachment?: boolean;
    isUnread?: boolean;
    isPinned?: boolean;
    mailbox?: string;
    after?: string;
    before?: string;
    limit?: number;
    ascending?: boolean;
    excludeDrafts?: boolean;
    includeTrash?: boolean;
    includeSpam?: boolean;
  }): Promise<QueryResult> {
    const mailboxes = await this.getMailboxes();
    const resolvedMailboxId = await this.resolveMailboxId(filters.mailbox, mailboxes);

    const base: any = {};
    if (filters.query) base.text = filters.query;
    if (filters.from) base.from = filters.from;
    if (filters.to) base.to = filters.to;
    if (filters.cc) base.cc = filters.cc;
    if (filters.bcc) base.bcc = filters.bcc;
    if (filters.subject) base.subject = filters.subject;
    if (filters.hasAttachment !== undefined) base.hasAttachment = filters.hasAttachment;
    if (filters.after) base.after = filters.after;
    if (filters.before) base.before = filters.before;
    if (resolvedMailboxId) base.inMailbox = resolvedMailboxId;

    // Each keyword is its own condition (mixed polarities can't share one FilterCondition).
    const conds: any[] = [];
    if (filters.isUnread === true) conds.push({ notKeyword: '$seen' });
    else if (filters.isUnread === false) conds.push({ hasKeyword: '$seen' });
    if (filters.isPinned === true) conds.push({ hasKeyword: '$flagged' });
    else if (filters.isPinned === false) conds.push({ notKeyword: '$flagged' });
    if (filters.excludeDrafts) conds.push({ notKeyword: '$draft' });

    const hasExplicitMailbox = !!resolvedMailboxId;
    const exclusion = computeExclusion(mailboxes, {
      includeTrash: filters.includeTrash,
      includeSpam: filters.includeSpam,
      hasExplicitMailbox,
    });
    const exclusionIntended = !hasExplicitMailbox && (!filters.includeTrash || !filters.includeSpam);

    return this.runFilteredQuery({
      base,
      conds,
      exclusion,
      exclusionIntended,
      limit: Math.min(filters.limit || 20, 100),
      ascending: filters.ascending ?? false,
      mailboxes,
    });
  }

  async getThread(threadId: string, includeDrafts: boolean = false): Promise<{ emails: any[]; hiddenDraftCount: number }> {
    const session = await this.getSession();

    // First, check if threadId is actually an email ID and resolve the thread
    let actualThreadId = threadId;

    // Try to get the email first to see if we need to resolve thread ID
    try {
      const emailRequest: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Email/get', {
            accountId: session.accountId,
            ids: [threadId],
            properties: ['threadId']
          }, 'checkEmail']
        ]
      };

      const emailResponse = await this.makeRequest(emailRequest);
      const email = this.getListResult(emailResponse, 0)[0];

      if (email && email.threadId) {
        actualThreadId = email.threadId;
      }
    } catch (error) {
      // If email lookup fails, assume threadId is correct
    }

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'getThread', name: 'Thread/get', path: '/list/*/emailIds' },
    };

    emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];

    // Use Thread/get with the resolved thread ID
    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Thread/get', {
          accountId: session.accountId,
          ids: [actualThreadId]
        }, 'getThread'],
        ['Email/get', emailGetParams, 'emails'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name', 'role'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const threadResult = this.getMethodResult(response, 0);

    // Check if thread was found
    if (threadResult.notFound && threadResult.notFound.includes(actualThreadId)) {
      throw new Error(`Thread with ID '${actualThreadId}' not found`);
    }

    // Resolve mailbox names onto the FULL list before filtering, so the draft
    // filter below doesn't skip the attach for retained messages.
    const emails = this.getListResult(response, 1);
    attachMailboxInfo(emails, buildMailboxInfoMap(this.readListResultIfPresent(response, 2)));

    // Drafts (e.g. an in-progress reply) are noise when reading a conversation,
    // so exclude them by default. Identify by the $draft keyword (survives a
    // draft moved out of the Drafts mailbox); opt back in via includeDrafts.
    // Return the hidden count (no extra query — derived from the already-fetched
    // thread) so the handler can ANNOUNCE that drafts were hidden without surfacing
    // them (the duplicate-draft trap: an agent reading a thread to reply must not
    // miss that a draft reply already exists). Assumes the full thread is returned
    // (Thread/get -> Email/get, no limit); threads are small so this holds.
    if (includeDrafts) {
      return { emails, hiddenDraftCount: 0 };
    }
    const filtered = emails.filter((e: any) => !e.keywords?.$draft);
    return { emails: filtered, hiddenDraftCount: emails.length - filtered.length };
  }

  async getMailboxStats(mailbox?: string): Promise<any> {
    // Fetch the full mailbox list once and read counts off it (getMailboxes returns all
    // fields, including the stat fields — it must NOT be narrowed). Resolving a specific
    // mailbox by id/role/name shares this list rather than issuing a second Mailbox/get.
    const mailboxes = await this.getMailboxes();
    const toStats = (mb: any) => ({
      id: mb.id,
      name: mb.name,
      role: mb.role,
      totalEmails: mb.totalEmails || 0,
      unreadEmails: mb.unreadEmails || 0,
      totalThreads: mb.totalThreads || 0,
      unreadThreads: mb.unreadThreads || 0,
    });

    if (mailbox !== undefined && String(mailbox).trim() !== '') {
      // Exact resolution (id/role/name); throws InvalidInputError on unknown. A real id
      // present but absent from the fetched list (a hidden/role-less mailbox) now throws
      // rather than returning stats — accepted residual (see docs/security-model.md).
      const mb = resolveMailbox(mailboxes, mailbox);
      return toStats(mb);
    }
    // Stats for all mailboxes.
    return mailboxes.map(toStats);
  }

  async getAccountSummary(): Promise<any> {
    const session = await this.getSession();
    const mailboxes = await this.getMailboxes();
    const identities = await this.getIdentities();

    // Calculate totals
    const totals = mailboxes.reduce((acc, mb) => ({
      totalEmails: acc.totalEmails + (mb.totalEmails || 0),
      unreadEmails: acc.unreadEmails + (mb.unreadEmails || 0),
      totalThreads: acc.totalThreads + (mb.totalThreads || 0),
      unreadThreads: acc.unreadThreads + (mb.unreadThreads || 0)
    }), { totalEmails: 0, unreadEmails: 0, totalThreads: 0, unreadThreads: 0 });

    return {
      accountId: session.accountId,
      mailboxCount: mailboxes.length,
      identityCount: identities.length,
      ...totals,
      mailboxes: mailboxes.map(mb => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0
      }))
    };
  }

  async bulkMarkRead(emailIds: string[], read: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = read
        ? { 'keywords/$seen': true }
        : { 'keywords/$seen': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkUpdate']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);
    
    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to update some emails.');
    }
  }

  async bulkPinEmails(emailIds: string[], pinned: boolean = true): Promise<void> {
    const session = await this.getSession();

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = pinned
        ? { 'keywords/$flagged': true }
        : { 'keywords/$flagged': null };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkFlag']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to pin/unpin some emails.');
    }
  }

  async bulkMove(emailIds: string[], target: string): Promise<void> {
    const session = await this.getSession();

    // Resolve the destination EXACTLY (id/role/name) — see moveEmail for the rationale
    // (new capability, exact-only, deliberate move-to-any stays open per fork #43).
    const mailboxes = await this.getMailboxes();
    const targetMailboxId = resolveMailbox(mailboxes, target).id;

    // Fetch current mailboxIds for all emails to build proper JMAP patches
    const getRequest: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/get', {
          accountId: session.accountId,
          ids: emailIds,
          properties: ['id', 'mailboxIds']
        }, 'getEmails']
      ]
    };
    const getResponse = await this.makeRequest(getRequest);
    const emails: any[] = this.getListResult(getResponse, 0);
    const mailboxMap: Record<string, Record<string, boolean>> = {};
    emails.forEach((e: any) => { mailboxMap[e.id] = e.mailboxIds || {}; });

    // Build patch per email: remove all current mailboxes, add target
    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      const patch: Record<string, boolean | null> = {};
      for (const mbId of Object.keys(mailboxMap[id] || {})) {
        patch[`mailboxIds/${mbId}`] = null;
      }
      patch[`mailboxIds/${targetMailboxId}`] = true;
      updates[id] = patch;
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkMove']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to move some emails.');
    }
  }

  async bulkDelete(emailIds: string[]): Promise<void> {
    const session = await this.getSession();

    // Find the trash mailbox by EXACT role only (case-insensitive) — see deleteEmail.
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findByExactRole(mailboxes, 'trash');

    if (!trashMailbox) {
      throw new Error('Could not find Trash mailbox');
    }

    const trashMailboxIds: Record<string, boolean> = {};
    trashMailboxIds[trashMailbox.id] = true;

    const updates: Record<string, any> = {};
    emailIds.forEach(id => {
      updates[id] = { mailboxIds: trashMailboxIds };
    });

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/set', {
          accountId: session.accountId,
          update: updates
        }, 'bulkDelete']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getMethodResult(response, 0);

    if (result.notUpdated && Object.keys(result.notUpdated).length > 0) {
      throw new Error('Failed to delete some emails.');
    }
  }
}