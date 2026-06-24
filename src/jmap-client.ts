import { FastmailAuth } from './auth.js';
import { validateFastmailUrl } from './url-validation.js';
import { parseAddress, requireNonEmpty, validateClearFields } from './coerce.js';
import { normalizeBodies, htmlHasVisibleContent, buildBodyParts, isBlank } from './body-format.js';
import { writeFile, mkdir, realpath, stat, lstat } from 'fs/promises';
import { dirname, resolve, normalize, sep, basename, join } from 'path';
import { homedir } from 'os';

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
}

// Shared Email/get property lists — keep in sync per CLAUDE.md rules.
// COMPACT: used by list/search tools in default mode and getThread(compact=true)
export const EMAIL_PROPERTIES_COMPACT = [
  'id', 'subject', 'from', 'to', 'cc', 'bcc', 'replyTo', 'receivedAt',
  'preview', 'keywords', 'threadId', 'messageId', 'references', 'inReplyTo',
  'hasAttachment', 'header:List-Unsubscribe:asURLs', 'blobId', 'size', 'mailboxIds',
] as const;

// VERBOSE: superset with body properties — used by verbose mode, getEmailById, getThread(compact=false)
// `sentAt` is here (a get-path superset addition, allowed by the property-consistency rule)
// for the reply-quote attribution (when the original was actually written), not in COMPACT.
export const EMAIL_PROPERTIES_VERBOSE = [
  ...EMAIL_PROPERTIES_COMPACT,
  'textBody', 'htmlBody', 'attachments', 'bodyValues', 'sentAt',
] as const;

export const EMAIL_BODY_PROPERTIES = ['partId', 'blobId', 'type', 'size', 'name'] as const;

// Build an id -> human name lookup from a Mailbox/get list. We key on the real
// `name` (not `role`) so custom labels — which have role:null — resolve too;
// default mailboxes already carry names like "Trash"/"Archive". (#10)
export function buildMailboxNameMap(mailboxes: any[]): Map<string, string> {
  const map = new Map<string, string>();
  for (const mb of mailboxes || []) {
    if (mb && mb.id && typeof mb.name === 'string') map.set(mb.id, mb.name);
  }
  return map;
}

// Attach the resolved mailbox/label names onto each raw email as a NON-enumerable
// `_mailboxNames` so JSON.stringify (the raw:true paths) omits it while
// simplifyEmail can still read it. Attach only when the email actually has
// mailboxIds and at least one resolved to a name — otherwise omit, don't
// fabricate (an empty/unresolvable set leaves the field absent). (#10)
export function attachMailboxNames(emails: any[], map: Map<string, string>): void {
  for (const email of emails || []) {
    if (!email || !email.mailboxIds) continue;
    const ids = Object.keys(email.mailboxIds);
    if (ids.length === 0) continue;
    const names = ids.map(id => map.get(id)).filter(Boolean) as string[];
    if (names.length === 0) continue;
    Object.defineProperty(email, '_mailboxNames', { value: names, enumerable: false, configurable: true });
  }
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

  protected findMailboxByRoleOrName(mailboxes: any[], role: string, nameFallback?: string): any | undefined {
    return mailboxes.find(mb => mb.role === role) ||
           (nameFallback ? mailboxes.find(mb => mb.name.toLowerCase().includes(nameFallback)) : undefined);
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

  async getEmails(mailboxId?: string, limit: number = 20, ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const filter = mailboxId ? { inMailbox: mailboxId } : {};

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
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', emailGetParams, 'emails'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getQueryResult(response, 0, 1);
    attachMailboxNames(result.items, buildMailboxNameMap(this.readListResultIfPresent(response, 2)));
    return result;
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
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name'] }, 'mailboxes']
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

    attachMailboxNames([email], buildMailboxNameMap(this.readListResultIfPresent(response, 1)));
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
    mailboxId?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
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

    // Use provided mailboxId or default to drafts for initial creation
    const initialMailboxId = email.mailboxId || draftsMailbox.id;

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
    mailboxId?: string;
    inReplyTo?: string[];
    references?: string[];
    replyTo?: string[];
  }): Promise<string> {
    const session = await this.getSession();

    // Validate at least one meaningful field is present (zero-width/whitespace-only
    // bodies count as absent).
    if (!email.to?.length && !email.subject && isBlank(email.textBody) && isBlank(email.htmlBody)) {
      throw new Error('At least one of to, subject, textBody, or htmlBody must be provided');
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

    // Resolve drafts mailbox
    let draftMailboxId: string;
    if (email.mailboxId) {
      draftMailboxId = email.mailboxId;
    } else {
      const mailboxes = await this.getMailboxes();
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
    const CLEARABLE = new Set(['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody']); // NOT 'from'
    const SETTABLE = ['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'from'] as const;
    const provided = new Set(SETTABLE.filter(f => (updates as any)[f] !== undefined));
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
    const wroteText = updates.textBody !== undefined;
    const wroteHtml = updates.htmlBody !== undefined;
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

    // Reject a body-less result (clearing the only body, or both). A draft keeps >=1 body.
    // Distinct from the clear-text-while-html guard (which only fires when merged html IS
    // present), so the two can't both match.
    if (isBlank(textBodyValue) && isBlank(htmlBodyValue)) {
      throw new Error('a draft needs a body; supply textBody or htmlBody (this edit would leave it with neither).');
    }

    // Carry existing (non-inline) attachments by referencing their existing blobIds.
    // Whitelist exactly these fields — a blob-backed part is blobId XOR partId, and
    // `size` is server-set, so sending partId/size would be rejected by a strict server.
    const carriedAttachments = existingAttachments.map((a: any) => ({
      blobId: a.blobId,
      type: a.type,
      ...(a.name != null && { name: a.name }),
      ...(a.disposition != null && { disposition: a.disposition }),
      ...(a.cid != null && { cid: a.cid }),
    }));

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
      ...(carriedAttachments.length && { attachments: carriedAttachments }),
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

  async getRecentEmails(limit: number = 10, mailboxName: string = 'inbox', ascending: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    // Find the specified mailbox (default to inbox)
    const mailboxes = await this.getMailboxes();
    const targetMailbox = mailboxes.find(mb =>
      mb.role === mailboxName.toLowerCase() ||
      mb.name.toLowerCase().includes(mailboxName.toLowerCase())
    );

    if (!targetMailbox) {
      throw new Error(`Could not find mailbox: ${mailboxName}`);
    }

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
    // Reuse the mailbox list already fetched above to resolve names — no extra methodCall.
    attachMailboxNames(result.items, buildMailboxNameMap(mailboxes));
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
    
    // Find the trash mailbox
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findMailboxByRoleOrName(mailboxes, 'trash', 'trash');

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

  async moveEmail(emailId: string, targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

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

  async addLabels(emailId: string, mailboxIds: string[]): Promise<void> {
    const session = await this.getSession();

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
    // below is the security boundary. So a bare filename lands safely in the
    // configured dir in one step, and an absolute path inside that dir writes
    // exactly there.
    const resolved = resolve(allowedDir, normalize(savePath));

    if (resolved.includes('\0')) {
      throw new Error('Save path contains null bytes');
    }

    if (!resolved.startsWith(allowedDir + sep) && resolved !== allowedDir) {
      throw new Error(
        `Save path must be within ${allowedDir}. ` +
        `Received: ${savePath}`
      );
    }

    return resolved;
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
          throw new Error(`Could not find existing ancestor for save path: ${lexical}`);
        }
        ancestor = parent;
      }
    }

    // Canonicalize the existing ancestor — this is what catches symlink escapes.
    const canonicalAncestor = await realpath(ancestor);
    if (canonicalAncestor !== canonicalAllowed && !canonicalAncestor.startsWith(canonicalAllowed + sep)) {
      throw new Error(
        `Save path resolves to '${canonicalAncestor}' which is outside the allowed directory '${canonicalAllowed}'. ` +
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
        throw new Error(`Refusing to overwrite an existing symlink at the target: ${safePath}`);
      }
    } catch (e: any) {
      if (e.code !== 'ENOENT') throw e;
    }

    return safePath;
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

  async advancedSearch(filters: {
    query?: string;
    from?: string;
    to?: string;
    subject?: string;
    hasAttachment?: boolean;
    isUnread?: boolean;
    isPinned?: boolean;
    mailboxId?: string;
    after?: string;
    before?: string;
    limit?: number;
    ascending?: boolean;
  }): Promise<QueryResult> {
    const session = await this.getSession();

    // Build JMAP filter object
    const filter: any = {};

    if (filters.query) filter.text = filters.query;
    if (filters.from) filter.from = filters.from;
    if (filters.to) filter.to = filters.to;
    if (filters.subject) filter.subject = filters.subject;
    if (filters.hasAttachment !== undefined) filter.hasAttachment = filters.hasAttachment;
    if (filters.isUnread === true) filter.notKeyword = '$seen';
    else if (filters.isUnread === false) filter.hasKeyword = '$seen';
    if (filters.isPinned === true) filter.hasKeyword = '$flagged';
    if (filters.isPinned === false) filter.notKeyword = '$flagged';
    if (filters.mailboxId) filter.inMailbox = filters.mailboxId;
    if (filters.after) filter.after = filters.after;
    if (filters.before) filter.before = filters.before;

    // When both isUnread and isPinned are set, hasKeyword/notKeyword may conflict.
    // JMAP FilterCondition only supports one hasKeyword, so wrap in an AND operator.
    let finalFilter: any = filter;
    if (filters.isUnread !== undefined && filters.isPinned !== undefined) {
      delete filter.hasKeyword;
      delete filter.notKeyword;
      const conditions: any[] = [filter];
      conditions.push(filters.isUnread ? { notKeyword: '$seen' } : { hasKeyword: '$seen' });
      conditions.push(filters.isPinned ? { hasKeyword: '$flagged' } : { notKeyword: '$flagged' });
      finalFilter = { operator: 'AND', conditions };
    }

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
          filter: finalFilter,
          sort: [{ property: 'receivedAt', isAscending: filters.ascending ?? false }],
          limit: Math.min(filters.limit || 50, 100),
          calculateTotal: true
        }, 'query'],
        ['Email/get', emailGetParams, 'emails'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getQueryResult(response, 0, 1);
    attachMailboxNames(result.items, buildMailboxNameMap(this.readListResultIfPresent(response, 2)));
    return result;
  }

  async searchEmails(query: string, limit: number = 20, ascending: boolean = false, excludeDrafts: boolean = false): Promise<QueryResult> {
    const session = await this.getSession();

    const emailGetParams: any = {
      accountId: session.accountId,
      '#ids': { resultOf: 'query', name: 'Email/query', path: '/ids' },
    };

    emailGetParams.properties = [...EMAIL_PROPERTIES_COMPACT];

    // A JMAP FilterCondition ANDs its properties, so text + notKeyword means
    // "matches the query AND is not a draft". Server-side, so calculateTotal
    // stays honest (no post-filtering).
    const filter: any = { text: query };
    if (excludeDrafts) filter.notKeyword = '$draft';

    const request: JmapRequest = {
      using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
      methodCalls: [
        ['Email/query', {
          accountId: session.accountId,
          filter,
          sort: [{ property: 'receivedAt', isAscending: ascending }],
          limit,
          calculateTotal: true
        }, 'query'],
        ['Email/get', emailGetParams, 'emails'],
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name'] }, 'mailboxes']
      ]
    };

    const response = await this.makeRequest(request);
    const result = this.getQueryResult(response, 0, 1);
    attachMailboxNames(result.items, buildMailboxNameMap(this.readListResultIfPresent(response, 2)));
    return result;
  }

  async getThread(threadId: string, includeDrafts: boolean = false): Promise<any[]> {
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
        ['Mailbox/get', { accountId: session.accountId, properties: ['id', 'name'] }, 'mailboxes']
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
    attachMailboxNames(emails, buildMailboxNameMap(this.readListResultIfPresent(response, 2)));

    // Drafts (e.g. an in-progress reply) are noise when reading a conversation,
    // so exclude them by default. Identify by the $draft keyword (survives a
    // draft moved out of the Drafts mailbox); opt back in via includeDrafts.
    return includeDrafts ? emails : emails.filter((e: any) => !e.keywords?.$draft);
  }

  async getMailboxStats(mailboxId?: string): Promise<any> {
    const session = await this.getSession();
    
    if (mailboxId) {
      // Get stats for specific mailbox
      const request: JmapRequest = {
        using: ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail'],
        methodCalls: [
          ['Mailbox/get', {
            accountId: session.accountId,
            ids: [mailboxId],
            properties: ['id', 'name', 'role', 'totalEmails', 'unreadEmails', 'totalThreads', 'unreadThreads']
          }, 'mailbox']
        ]
      };

      const response = await this.makeRequest(request);
      return this.getListResult(response, 0)[0];
    } else {
      // Get stats for all mailboxes
      const mailboxes = await this.getMailboxes();
      return mailboxes.map(mb => ({
        id: mb.id,
        name: mb.name,
        role: mb.role,
        totalEmails: mb.totalEmails || 0,
        unreadEmails: mb.unreadEmails || 0,
        totalThreads: mb.totalThreads || 0,
        unreadThreads: mb.unreadThreads || 0
      }));
    }
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

  async bulkMove(emailIds: string[], targetMailboxId: string): Promise<void> {
    const session = await this.getSession();

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

    // Find the trash mailbox
    const mailboxes = await this.getMailboxes();
    const trashMailbox = this.findMailboxByRoleOrName(mailboxes, 'trash', 'trash');

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