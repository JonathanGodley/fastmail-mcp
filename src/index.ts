#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import { FastmailAuth, FastmailConfig } from './auth.js';
import { JmapClient, QueryResult } from './jmap-client.js';
import { ContactsCalendarClient } from './contacts-calendar.js';
import { CalDAVCalendarClient } from './caldav-client.js';
import { simplifyEmail, setDefaultTimezone } from './email-formatter.js';
import { formatQueryResult, formatRawEmailQueryResult, formatEmailQueryResult, buildExclusionNote, buildAttachmentListContent, simplifyMailbox, simplifyIdentity, simplifyContact, formatContactQueryResult, formatEditDraftResult, formatSendDraftResult, formatInlineNotes } from './response-formatters.js';
import { coerceRecipients, coerceStringArray, coerceBool, coercePosition, redactBearerTokens, assertKnownParams, coerceAttachments, PathAccessError, InvalidInputError } from './coerce.js';
import { parseEmailFields, projectEmail, wantsHtmlBody } from './field-projection.js';
import { composeReply } from './reply-handler.js';
import { composeForward } from './forward-handler.js';
import { sendDraftAndMaintainKeywords } from './send-draft-handler.js';
import { composeDraft } from './compose-handler.js';
import { assertBodyInputs } from './body-format.js';
import { assertStripQuotedNotRaw } from './quote-strip.js';
import { readThread } from './thread-handler.js';

const server = new Server(
  {
    name: 'fastmail-mcp',
    version: '1.9.4-fork.11',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

let jmapClient: JmapClient | null = null;
let contactsCalendarClient: ContactsCalendarClient | null = null;
let caldavClient: CalDAVCalendarClient | null = null;

function findEnvValue(keys: string[]): { value?: string; key?: string; wasPlaceholder: boolean } {
  const isPlaceholder = (val: string) => /\$\{[^}]+\}/.test(val.trim());
  for (const key of keys) {
    const raw = process.env[key];
    if (typeof raw === 'string' && raw.trim().length > 0) {
      if (isPlaceholder(raw)) {
        return { value: undefined, key, wasPlaceholder: true };
      }
      return { value: raw.trim(), key, wasPlaceholder: false };
    }
  }
  return { value: undefined, key: undefined, wasPlaceholder: false };
}

function maskSecret(value: string): string {
  if (value.length <= 6) return '***';
  return `${value.slice(0, 4)}…${value.slice(-2)} (len ${value.length})`;
}

function getAuthConfig(): FastmailConfig {
  const tokenInfo = findEnvValue([
    'FASTMAIL_API_TOKEN',
    'USER_CONFIG_FASTMAIL_API_TOKEN',
    'USER_CONFIG_fastmail_api_token',
    'fastmail_api_token',
  ]);
  const apiToken = tokenInfo.value;
  if (!apiToken) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      'FASTMAIL_API_TOKEN environment variable is required'
    );
  }

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  // Opt-in for self-hosted JMAP servers. Required to use any base URL outside
  // the api.fastmail.com / www.fastmailusercontent.com allowlist (which already
  // covers Fastmail's regional hosts, e.g. phl.api.fastmail.com).
  const unsafeInfo = findEnvValue([
    'FASTMAIL_ALLOW_UNSAFE_BASE_URL',
    'USER_CONFIG_FASTMAIL_ALLOW_UNSAFE_BASE_URL',
  ]);
  const allowUnsafeBaseUrl = unsafeInfo.value === 'true' || unsafeInfo.value === '1';

  return { apiToken, baseUrl: baseInfo.value, allowUnsafeBaseUrl };
}

function initializeClient(): JmapClient {
  if (jmapClient) {
    return jmapClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  jmapClient = new JmapClient(auth);
  return jmapClient;
}

function initializeContactsCalendarClient(): ContactsCalendarClient {
  if (contactsCalendarClient) {
    return contactsCalendarClient;
  }

  const auth = new FastmailAuth(getAuthConfig());
  contactsCalendarClient = new ContactsCalendarClient(auth);
  return contactsCalendarClient;
}

function initializeCalDAVClient(): CalDAVCalendarClient | null {
  if (caldavClient) return caldavClient;

  const username = findEnvValue([
    'FASTMAIL_CALDAV_USERNAME',
    'USER_CONFIG_FASTMAIL_CALDAV_USERNAME',
  ]).value;
  const password = findEnvValue([
    'FASTMAIL_CALDAV_PASSWORD',
    'USER_CONFIG_FASTMAIL_CALDAV_PASSWORD',
  ]).value;

  if (!username || !password) return null;

  caldavClient = new CalDAVCalendarClient({ username, password });
  return caldavClient;
}

function getDownloadDir(): string | undefined {
  return findEnvValue([
    'FASTMAIL_DOWNLOAD_DIR',
    'USER_CONFIG_FASTMAIL_DOWNLOAD_DIR',
    'USER_CONFIG_fastmail_download_dir',
    'fastmail_download_dir',
  ]).value;
}

// The directory attachable files must live within, for outgoing-attachment sending.
// Resolved independently of the download dir (a shared value re-opens a download->
// upload round-trip; that's the operator's explicit choice, not a default). No default
// and no auto-create: an unset value leaves the capability disabled. Blank collapses to
// undefined via findEnvValue, so FASTMAIL_ATTACH_DIR="" can't resolve to cwd.
function getAttachDir(): string | undefined {
  return findEnvValue([
    'FASTMAIL_ATTACH_DIR',
    'USER_CONFIG_FASTMAIL_ATTACH_DIR',
    'USER_CONFIG_fastmail_attach_dir',
    'fastmail_attach_dir',
  ]).value;
}

// Shared `attachments` schema + description for the compose tools. Read getAttachDir()
// at module load (same as the download `path` description reads getDownloadDir()), and
// render an honest disabled clause when it's unset — attachments have no fallback
// default the way downloads do, so we must not print a phantom root.
function attachmentsDescription(forEdit: boolean): string {
  const dir = getAttachDir();
  const gate = dir
    ? `Files must resolve within ${dir} (set via FASTMAIL_ATTACH_DIR); a bare filename or relative path resolves against it, and an absolute path must fall inside it.`
    : `Attachments are disabled until FASTMAIL_ATTACH_DIR is set (restart to enable); each path will then resolve within that directory.`;
  const base =
    `Files to attach, each an object { path (required), name?, contentType?, cid? }. ${gate} ` +
    `contentType is inferred from the file extension when omitted; an explicit contentType is echoed by Fastmail as-is (not re-detected), so a wrong value rides out wrong. ` +
    `Give an item a cid to EMBED it in the message body instead of hanging it off the end: reference it from htmlBody as <img src="cid:THE_CID">. ` +
    `An htmlBody reference with no matching item is rejected, and an item whose cid nothing displays (no htmlBody, or no reference to it) is still attached as an ordinary file and the result says so — a supplied file is never dropped. ` +
    `Size caps (~25 MB/file, ~45 MB total) are a fail-fast guard — Fastmail's own limit ultimately governs.`;
  if (!forEdit) return base;
  return base +
    ` On edit_draft this APPENDS to the draft's existing attachments (they are kept). ` +
    `To remove specific ones, use removeAttachments (pass the blobId from get_email_attachments; a unique attachment name also works). ` +
    `To remove all, use clearFields:['attachments']. Passing attachments together with clearFields:['attachments'] is rejected as a conflict.`;
}

// `leadIn` prepends tool-specific context to the shared description (defaulted to ''
// so send/reply/create/edit are untouched) — forward_email uses it to state how NEW
// uploads relate to the original's own carried attachments.
function attachmentsSchemaProperty(forEdit: boolean, leadIn = '') {
  return {
    type: 'array',
    items: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'Path to the file to attach: an absolute path within the attach directory, or a bare filename/relative path resolved against it.' },
        name: { type: 'string', description: "Filename recipients see (optional; defaults to the file's basename)." },
        contentType: { type: 'string', description: 'MIME type like application/pdf (optional; inferred from the file extension when omitted).' },
        cid: { type: 'string', description: 'Content-ID to embed this file under (optional), referenced from htmlBody as <img src="cid:THE_CID">. A simple token of up to 64 letters, digits, dot, dash or underscore, of your choosing — a spelling copied from HTML ("cid:logo") or a header ("<logo>") is accepted and normalised. Each item needs a distinct one. Omit it for an ordinary attachment; identifiers of the form ii-<hex>@inline.invalid are reserved for images this server embeds on your behalf and cannot be authored.' },
      },
      required: ['path'],
    },
    description: leadIn + attachmentsDescription(forEdit),
  };
}

function getTimezone(): string | undefined {
  return findEnvValue([
    'FASTMAIL_TIMEZONE',
    'USER_CONFIG_FASTMAIL_TIMEZONE',
    'USER_CONFIG_fastmail_timezone',
    'fastmail_timezone',
  ]).value;
}

// Shared scope-control descriptions for the read tools (search_emails + list_emails),
// defined once so the per-flag strings and the reliability-contract clause stay in sync
// between the two tools rather than drifting as hand-copied strings.
const EXCLUDE_DRAFTS_DESC =
  'Drafts are included by default; set true to omit them from results (and from the total count). (Note: get_thread differs on BOTH axes — it uses includeDrafts AND excludes drafts by default.)';
const INCLUDE_TRASH_DESC =
  'Trash is excluded by default; set true to also include Trash in the results.';
const INCLUDE_SPAM_DESC =
  'The Spam/Junk folder is excluded by default; set true to also include it in the results.';

// The reliability contract that makes silence trustworthy. Lead with the no-note
// guarantee as its own sentence (a skimming model must hit "no note => trustworthy"
// first), then the per-signal actions; scoped to the default all-mailbox scope.
const SCOPE_RELIABILITY_CONTRACT =
  'When you search the default scope (no mailbox set): NO note means no Trash/Spam message matched this search — do not re-run with includeTrash/includeSpam just to re-check the same query. ' +
  'A "N in Trash/Spam excluded" note means re-run (includeTrash:true / includeSpam:true, or mailbox:"trash"/"junk") to see those matches. ' +
  'A "count could not be confirmed" or "folder not found; not excluded" note means re-run to be sure. ' +
  'Setting mailbox searches only that folder (no note, by definition).';

const MAILBOX_PARAM_DESC =
  'Mailbox to scope to: an id, a role (inbox, trash, junk, sent, drafts, archive), or a folder name (e.g. Receipts). Setting it searches exactly that mailbox (incl. Trash/Spam) and ignores the default Trash/Spam exclusion. Unknown mailbox is rejected with the valid list.';

// One canonical explanation of the simplified location + status fields, shared
// verbatim by every read tool (get_email, get_thread, list_emails, search_emails,
// get_recent_emails) so the five can't drift. Carries the only-when-true semantics,
// the junk=Spam gloss, the "same set, not parallel arrays — test membership" rule,
// and the keyword-vs-location two-axis model. The rare unresolvedMailboxIds field is
// intentionally NOT here (documented in the README, not this per-call surface). (#49)
const LOCATION_FIELDS_DESC =
  'Use `roles` to tell where a message is filed — stable lowercase JMAP roles: inbox, archive, sent, drafts, trash, junk (junk is the role of the folder shown as "Spam"; there is no "spam" role). `mailboxes` holds folder display names, which the user can rename, so do not identify a folder by a `mailboxes` name (a custom folder can even be named "Trash"). `roles` and `mailboxes` describe the SAME set of mailboxes the message is in (a message can be in several at once) but are NOT positionally aligned — a custom folder appears in `mailboxes` with no `roles` entry — so test membership (roles.includes("trash")), never roles[0] or roles[i] vs mailboxes[i]. Separately, the is* flags (isRead/isFlagged/isDraft/isAnswered/isForwarded) are status, not location: isDraft and a drafts role normally agree, and when they diverge (a draft filed in Trash gives isDraft:true with roles:["trash"]) both are still correct. isAnswered/isForwarded appear only when true. Simplified-only — raw=true returns the underlying JMAP keywords and opaque mailboxIds.';

// Shared, verbatim across the compact-listing read tools (list_emails, search_emails,
// get_recent_emails, get_thread) so their preview/size guidance can't drift. Names the
// trap behind #59: an agent read a `preview` snippet, saw a large `size`, and wrongly
// concluded the body's real content was absent without ever fetching get_email.
const PREVIEW_SIZE_DESC =
  '`preview` is a truncated snippet (~256 chars max), NOT the full body. `bodyTextSize` is the full text-body size in bytes (it includes quoted history, so treat it as an upper bound); when it is much larger than the preview, fetch get_email before concluding content is absent. `size` is the whole-message size including attachments and inline images, so it is NOT a body-length proxy.';

// Shared, verbatim across create_draft's inReplyTo and references, so the hand-rolled
// threading hazard is stated on whichever one the caller reaches for. Fastmail groups a
// message into a conversation by subject as well as by the threading headers, so a draft
// carrying the headers under a different subject lands on a new threadId; observed after
// that: every later draft replying to the same original was assigned to that splinter
// thread too, and deleting the offending drafts did not restore the grouping (#68).
const THREAD_SPLINTER_DESC =
  'THREADING HAZARD: Fastmail groups a message into an existing conversation by SUBJECT as well as by these headers. A draft that carries them under a subject that does not match the thread\'s base subject is given a NEW threadId, and from then on later drafts replying to that same original message are grouped onto that splinter thread as well — including ones created afterwards with the correct "Re:" subject and full reference chain. Deleting the offending drafts does not undo it. The effect is display-only (the headers are correct, so recipients thread normally and sending resolves it), but the drafts stay detached from the conversation in the Fastmail UI. To reply on an existing thread prefer reply_email, which builds the headers and the matching subject for you and takes a deliberate `subject` override.';

// The `fields` projection, shared verbatim by every read tool that offers it
// (get_email, list_emails, search_emails, get_recent_emails, get_thread) so they
// can't drift. One sentence goes in the tool description (why you would reach for
// it), the full contract in the parameter description. (#69, #79)
const FIELDS_TOOL_DESC =
  'Use `fields` to return ONLY the fields you need (e.g. fields:["id","subject","from","date","threadId"]) when the default shape would be too large for one response.';

const FIELDS_PARAM_DESC =
  'Return ONLY these simplified fields, e.g. ["id","subject","from","date","threadId"] for a headers-only sweep. Response size otherwise depends on what is in the mailbox (thread references and previews dominate a wide listing), so this is the way to keep a many-message read inside one response instead of splitting it into several. Names must match the simplified field names EXACTLY (camelCase); an unknown name is rejected with the full valid list rather than silently returning nothing. Omit the parameter for the default shape - an empty array is rejected. Cannot be combined with raw:true (raw returns untransformed JMAP, whose field names differ). A field a message does not have is simply absent, so a narrow projection can come back as {}. Any field needing the full-message fetch (bodyText, bodyHtml, bodyHtmlSize, attachments, forwardedMessageId, sourceEmailId) is a valid name on list_emails/search_emails/get_recent_emails but is never populated there — those results carry hasAttachment, isForwarded and bodyTextSize instead; fetch get_email for the rest. Selecting `mailboxes` or `roles` also emits `unresolvedMailboxIds` in the rare case an id could not be resolved, so a partial location is never hidden. On get_thread, `bodyText` IS populated when includeBodies:true (`bodyHtml` never is), and a projected bodyText keeps its signals (quotedBytesStripped/quotedStripSkipped/bodyTextUnavailable) uninvited — without them a stripped body would read as verbatim.';

// The `fields` parameter, declared identically on every read tool that offers it.
// The string alternative is advertised because lenient clients stringify arrays
// (coerceStringArray accepts a JSON or comma-separated string). (#69, #79)
function fieldsSchemaProperty() {
  return {
    oneOf: [
      { type: 'array', items: { type: 'string' } },
      { type: 'string' },
    ],
    description: FIELDS_PARAM_DESC,
  };
}

// Paging, shared verbatim by the three list/search tools (list_emails, search_emails,
// get_recent_emails) so their offset contract can't drift. One sentence goes in each
// tool description (that a result is one page and how to tell there are more), the full
// contract in the parameter description. (#51)
const POSITION_TOOL_DESC =
  'Results are ONE PAGE: the summary line always states the total number of matches, and when more remain it carries a `nextPosition` to pass back as `position`. No `nextPosition` means you have seen every match — do not re-run to check.';

const POSITION_PARAM_DESC =
  'Skip this many results before returning the page — a 0-based offset into the full match set, and the way to read past this tool\'s `limit` cap (e.g. limit:50, then position:50, position:100). Take the value from the previous response\'s `nextPosition` rather than computing it: it is the position the server actually served plus what it actually returned, so a short final page ends the listing instead of advertising another one. Every response states the total match count; `nextPosition` appears only while more results remain. All filters (including the default Trash/Spam exclusion) are applied server-side to every page, so paging never changes what matches. The Trash/Spam withheld-count note describes the WHOLE match set, not the page: the same count repeats on every page, so never add up the notes across pages. Omit it, or pass 0, for the first page. A position past the end is not an error — it returns an empty page alongside the real total, so you can see you overshot. Must be a whole number, 0 or greater: a negative value is rejected (JMAP would read it as counting back from the end; to read from the oldest end use ascending:true) and so is a fraction.';

// The `position` parameter, declared identically on every list/search tool. Numbers are
// also accepted as strings, matching `limit`, because lenient clients stringify them.
function positionSchemaProperty() {
  return {
    type: ['number', 'string'],
    description: POSITION_PARAM_DESC,
  };
}

// One canonical statement that the attachment listing is not just "attached files",
// shared verbatim by every tool that emits one (get_email, get_email_attachments, and
// get_thread under includeBodies) so they can't drift on what the listing covers (#13).
const UNION_SCOPE_DESC =
  'Attachment entries include images embedded in the message body, not only "attached" files.';

// The two keys a simplified attachment entry gains for embedded images, shared verbatim
// by the tools that emit simplified entries (get_email, and get_thread under
// includeBodies, whose messages carry the same entry shape). get_email_attachments
// deliberately does NOT carry this: it returns raw JMAP parts, which have no derived
// flag to describe. The sender-declared caveat travels with the keys because both
// values come from the message itself. (#13)
const INLINE_PAIR_DESC =
  'Attachment entries carry two extra keys for embedded images. isInline:true means EITHER the server routed the part into the message body OR the sender marked it Content-Disposition: inline — so it covers body-displayed images, and also an ordinary file the sender merely labelled inline. `cid` is that part\'s Content-ID: the value a cid: reference in the HTML body points at, and the handle download_attachment accepts as cid:<value>; a part the body actually references has one. Both keys are omitted when they do not apply (isInline never appears as false). Both are SENDER-DECLARED metadata, exactly like `name` and `contentType`: a sender chooses whether a part is marked inline and what it is called, so isInline is a rendering hint, never a reason to treat a part as harmless or to skip inspecting it.';

// Shared verbatim by the compact list/search reads (list_emails, search_emails,
// get_recent_emails) and by get_thread's default mode, which fetch no attachment parts
// at all. Without this, hasAttachment:false reads as "no images" — and Fastmail's
// hasAttachment heuristic answers "content or decoration", so an embedded logo or a
// small pasted image is exactly what it filters out. (#13)
const COMPACT_ATTACHMENT_DESC =
  'These results carry hasAttachment but never the attachment entries themselves, and hasAttachment is a server heuristic that deliberately ignores small decorative images — so a message whose only picture is embedded in its body can read as hasAttachment:false here. Fetch get_email (or get_thread with includeBodies) to see the actual parts.';

// Shared by get_email and get_thread so the two can't drift on what stripping does,
// what it does NOT touch, and how to read the signal it returns (#73).
const STRIP_QUOTED_DESC =
  'Remove quoted reply history from the plain-text body, so a long thread is not re-read at every quote depth. Opt-in; the default output is verbatim. ' +
  'Detection is deliberately conservative and covers the conventional markers only: leading ">" quote runs (including nested ">>"), an "On <date>, <someone> wrote:" attribution directly above such a run, an Outlook From:/Sent:/To:/Subject: header block, and "-----Original Message-----". An unrecognised shape is returned UNCHANGED rather than guessed at. ' +
  'Read the result from the signals: `quotedBytesStripped` is how many bytes went, and 0 means no marker matched and the body is whole (do not re-fetch to check); `quotedStripSkipped` instead means there was no non-empty plain-text body to strip (an HTML-only message). ' +
  'Applies to `bodyText` only — a `bodyHtml` returned alongside it is NOT stripped. Cannot be combined with raw (raw is unmodified JMAP). ' +
  'Stripping can also over-reach, because these markers are conventions rather than syntax: a leading ">" is equally a markdown blockquote or a pasted shell prompt, a pasted email header block (From: with an address, plus To:/Sent:/Subject:) is treated as a quoted section and cuts to the end of the message, and a FORWARDED message\'s content sits below the same "-----Original Message-----" marker, leaving only the covering note. So treat a `quotedBytesStripped` that looks too large for the message as the cue to re-read that message without the flag.';

// Single source of truth for the tool catalog. Hoisted to module scope so the
// CallTool handler can derive each tool's declared parameter set for the
// unknown-parameter guard (#11) — no drift from what clients see via ListTools.
const TOOLS = [
      {
        name: 'list_mailboxes',
        description: 'List all mailboxes in the Fastmail account. Returns simplified format by default with core fields (name, role, counts). Use verbose=true only if you need extra fields like sortOrder or myRights. Use raw=true for original JMAP response.',
        inputSchema: {
          type: 'object',
          properties: {
            verbose: {
              type: 'boolean',
              description: 'Include extra mailbox fields (sortOrder, isSubscribed, myRights). Not needed for most tasks.',
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'list_emails',
        description: 'List recent emails across all mailboxes (or one, via mailbox). Trash and Spam are excluded by default (set includeTrash/includeSpam to include them); drafts are included (set excludeDrafts to omit them). Set mailbox to scope to a single mailbox (incl. Trash/Spam), which ignores the default exclusion. ' + SCOPE_RELIABILITY_CONTRACT + ' Spans all mailboxes; for just the Inbox\'s newest use get_recent_emails. Returns simplified format (metadata + preview, no bodies). Use raw=true for original JMAP response. For email bodies, use get_email. The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            mailbox: {
              type: 'string',
              description: MAILBOX_PARAM_DESC,
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of emails to return (default: 20, max: 100)',
              default: 20,
            },
            position: positionSchemaProperty(),
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
            excludeDrafts: {
              type: 'boolean',
              description: EXCLUDE_DRAFTS_DESC,
            },
            includeTrash: {
              type: 'boolean',
              description: INCLUDE_TRASH_DESC,
            },
            includeSpam: {
              type: 'boolean',
              description: INCLUDE_SPAM_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'get_email',
        description: 'Get a specific email by ID. Returns simplified format with plain text body (HTML omitted, bodyHtmlSize hint provided). Only use verbose=true if you specifically need the HTML body — it can be very large for marketing emails. Use raw=true for original JMAP response. Set stripQuoted=true to drop quoted reply history from bodyText when reading a message deep in a long thread (the quoted tail is duplicated from earlier messages). The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + UNION_SCOPE_DESC + ' ' + INLINE_PAIR_DESC + ' ' + FIELDS_TOOL_DESC + ' On this tool fields:["bodyHtml"] returns the HTML body ALONE (no verbose needed, no metadata, no plain-text copy) — the way to read a large HTML draft without the rest of the message pushing the response past the output limit.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to retrieve',
            },
            verbose: {
              type: 'boolean',
              description: 'Include HTML body in response. WARNING: can produce very large responses (50K+ chars) for marketing/rich emails. Only use when HTML content is specifically needed.',
            },
            fields: fieldsSchemaProperty(),
            stripQuoted: {
              type: 'boolean',
              description: STRIP_QUOTED_DESC,
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'reply_email',
        description: 'Reply to an existing email with proper threading headers (In-Reply-To, References). Automatically fetches the original email to build the reply chain. The subject defaults to \'Re: <original subject>\'; pass subject to override it (see that parameter for what a changed subject does to draft grouping). Use this rather than hand-rolling threading headers on create_draft. This tool always saves the reply as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the reply, it marks the original answered and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft), so when several copies of the original exist the right one is marked. The original message is quoted by default (attributed, top-posted, matching the web client with a portable quote-bar); set quoteOriginal=false to omit it. Quoted HTML is reproduced sanitised (script/style/event handlers stripped; formatting and real http(s) images kept; inline cid: images omitted) and is re-sent under your From address. You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; requires FASTMAIL_ATTACH_DIR).',
        inputSchema: {
          type: 'object',
          properties: {
            originalEmailId: {
              type: 'string',
              description: 'ID of the email to reply to',
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional, defaults to the original sender). Each entry may be "Name <email>" or a bare address.',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            subject: {
              type: 'string',
              description: 'Override the subject (optional). When omitted the reply inherits "Re: <original subject>", without double-prefixing an existing Re:; a blank string is treated as omitted. The In-Reply-To/References headers are built identically either way, so recipients still thread the reply correctly. NOTE: Fastmail groups messages by subject as well as by those headers, so a changed subject detaches the DRAFT from the original conversation in the Fastmail UI (it is given its own threadId) — and once such a draft exists, later drafts replying to that same original message may be grouped onto the detached thread too. Display only, and it resolves once the reply is sent.',
            },
            textBody: {
              type: 'string',
              description: 'Plain-text body (optional). Use it for genuinely plain messages, or alongside htmlBody to provide your own plain-text alternative in place of the auto-generated one. Must be a plain string: a body wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional), and the preferred format for outgoing mail. When both bodies are supplied, recipients\' clients render this one. Supplying htmlBody alone is fine: a readable plain-text alternative is generated automatically whenever one can be derived from the HTML. In that derivation an image contributes its alt text; an embedded (cid:) image with no alt contributes \"[image]\", so a picture-only message still has a readable text part, while a remote image with no alt contributes nothing. Pass REAL markup — a body that is entirely HTML-escaped (escaped element tags like &lt;p&gt; with no actual elements) is rejected, because recipients would see the tags as text; so is any body containing a CDATA section, whose contents are dropped from the derived plain-text alternative.',
            },
            quoteOriginal: {
              type: ['boolean', 'string'],
              description: 'Append the original message as an attributed quote (default true). Set false to omit it.',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender). Each entry may be "Name <email>" or a bare address.',
            },
            attachments: attachmentsSchemaProperty(false),
          },
          required: ['originalEmailId'],
        },
      },
      {
        name: 'forward_email',
        description: 'Forward an existing email to new recipients. This tool always saves the forward as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the forward, it marks the original forwarded and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft, on both inline and asAttachment forwards). `to` is required — a forward has no default recipient, unlike reply. A note (textBody/htmlBody) is optional: the forwarded message itself is the content. The original is reproduced below a forwarded-message header block (From/To/Cc/Subject/Date), its HTML sanitised (script/style/event handlers stripped; formatting and real http(s) images kept) and re-sent under your From address. The original\'s regular attachments are carried by default, but embedded inline (cid:) images are NOT carried by an inline forward (their references are stripped from the reproduced HTML) — use asAttachment for full fidelity. The subject defaults to \'Fwd: <original subject>\'. You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; requires FASTMAIL_ATTACH_DIR).',
        inputSchema: {
          type: 'object',
          properties: {
            originalEmailId: {
              type: 'string',
              description: 'ID of the email to forward',
            },
            to: {
              oneOf: [
                { type: 'array', items: { type: 'string' } },
                { type: 'string' },
              ],
              description: 'Recipient email addresses (array of strings, or a comma-separated string); required — a forward has no default recipient. Each entry may be "Name <email>" or a bare address.',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            subject: {
              type: 'string',
              description: "Override the subject (optional). When omitted: 'Fwd: <original subject>', without double-prefixing an existing Fwd:/Fw:/FW:. A blank string is treated as omitted; a non-string is rejected. (Same posture as reply_email's subject.)",
            },
            textBody: {
              type: 'string',
              description: 'Optional note placed ABOVE the forwarded-message block, in plain text — the original is reproduced below it automatically; omit for a bare FYI forward. NOTE: a text-only note produces a PLAIN-TEXT forward (the original\'s HTML formatting is reduced to text) — use htmlBody, or both, to preserve its formatting. (When asAttachment is set, this note is the whole body — the original rides as the attached .eml, with no inline block.) Must be a plain string: a note wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'Optional note placed ABOVE the forwarded-message block, in HTML (the preferred format; a plain-text alternative is derived automatically, using the alt text of each image, and \"[image]\" for an embedded (cid:) image that has none) — the original is reproduced below it. (When asAttachment is set, this note is the whole body — the original rides as the attached .eml, with no inline block.) Pass REAL markup — an entirely HTML-escaped note (escaped element tags like &lt;p&gt; with no actual elements) is rejected, as is one containing a CDATA section.',
            },
            includeOriginalAttachments: {
              type: ['boolean', 'string'],
              description: "Carry the original message's attachments on the forward (default true; all-or-none — to drop individual ones, save the default draft and use edit_draft's removeAttachments). Embedded inline (cid:) images are never carried by an inline forward; use asAttachment for those. Ignored when asAttachment is set — the .eml already embeds every original attachment.",
            },
            asAttachment: {
              type: ['boolean', 'string'],
              description: 'Instead of reproducing the original inline, attach the entire original as a raw .eml file (message/rfc822): lossless, including embedded inline images; supersedes includeOriginalAttachments. NOTE: the raw message carries its full transport headers (Received chain, authentication results) and — when forwarding a message from Sent — any Bcc recipients (see docs/security-model.md), which an inline forward would not expose.',
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender). Each entry may be "Name <email>" or a bare address.',
            },
            attachments: attachmentsSchemaProperty(false, "NEW files to upload and attach (the original's own attachments are carried automatically — see includeOriginalAttachments). "),
          },
          required: ['originalEmailId', 'to'],
        },
      },
      {
        name: 'create_draft',
        description: 'Create an email draft without sending it (transmit it later with send_draft, the only tool that sends mail). Supports threading headers for replies, but for a reply to an existing message prefer reply_email — hand-rolled inReplyTo/references under a mismatched subject permanently detach the draft from its conversation in the Fastmail UI (see the inReplyTo parameter). IMPORTANT: each call creates a new draft — do not call twice for the same message. You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; requires FASTMAIL_ATTACH_DIR).',
        inputSchema: {
          type: 'object',
          properties: {
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Recipient email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'CC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'BCC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            from: {
              type: 'string',
              description: 'Sender email address (optional, defaults to account primary email)',
            },
            mailbox: {
              type: 'string',
              description: 'Mailbox to SAVE the draft into — id, role, or name (optional, defaults to Drafts). Does not set From or recipients. Unknown mailbox is rejected with the valid list.',
            },
            subject: {
              type: 'string',
              description: 'Email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Plain-text body (optional). Use it for genuinely plain messages, or alongside htmlBody to provide your own plain-text alternative in place of the auto-generated one. Must be a plain string: a body wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional), and the preferred format for outgoing mail. When both bodies are supplied, recipients\' clients render this one. Supplying htmlBody alone is fine: a readable plain-text alternative is generated automatically whenever one can be derived from the HTML. In that derivation an image contributes its alt text; an embedded (cid:) image with no alt contributes \"[image]\", so a picture-only message still has a readable text part, while a remote image with no alt contributes nothing. Pass REAL markup — a body that is entirely HTML-escaped (escaped element tags like &lt;p&gt; with no actual elements) is rejected, because recipients would see the tags as text; so is any body containing a CDATA section, whose contents are dropped from the derived plain-text alternative.',
            },
            inReplyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs to reply to (optional, for threading). ' + THREAD_SPLINTER_DESC,
            },
            references: {
              type: 'array',
              items: { type: 'string' },
              description: 'Message-IDs for References header (optional, for threading). ' + THREAD_SPLINTER_DESC,
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender). Each entry may be "Name <email>" or a bare address.',
            },
            attachments: attachmentsSchemaProperty(false),
          },
        },
      },
      {
        name: 'edit_draft',
        description: 'Edit an existing draft email. Only fields you provide are changed; omit a field to leave it unchanged. Setting a field to an empty value is rejected: to deliberately clear a field, name it in `clearFields`. A cleared draft is still valid (it just may not be sendable, e.g. with no recipients). The plain-text body is an auto-managed fallback of the HTML: editing htmlBody alone regenerates textBody from the new HTML (an html-alone edit discards any custom textBody the draft had); editing textBody alone while htmlBody is present is rejected (it would not change what recipients render); clearFields:[\'textBody\'] while htmlBody is present is rejected (the fallback is auto-managed); clearFields:[\'htmlBody\'] converts the draft to plain text. An edit that would leave the draft with no body is rejected. Editing the body of a reply draft that still carries the quoted original — or of a forward draft that carries the forwarded-message block — requires you to say what happens to it: pass originalEmailId (the id of the message this draft replies to or forwards) to rebuild the body and keep it, or noQuote:true to drop it. Metadata-only edits (subject/recipients/attachments) and plain-text conversion (clearFields:[\'htmlBody\']) keep the quote automatically; each successive body edit that should keep the quote must pass originalEmailId again. Supplying htmlBody to a text-only reply draft converts it to HTML. Since JMAP emails are immutable, this creates a replacement draft and moves the old one to Trash (so the returned email ID is new); the edit preserves the draft\'s threading headers (In-Reply-To/References), attachments, and other keywords. The replaced draft is never destroyed: it stays recoverable in Trash until Trash is emptied or auto-purged (Trash retention is a per-account setting), so an edit made from an out-of-date copy of the draft can be undone. The result also reports what the replaced draft contained (subject, recipients, body sizes) — compare it against what you expected to replace, since a draft changed elsewhere (the web UI, another client) in the meantime will have been overwritten. On the rare failure where the replacement is created but the old copy can\'t be moved to Trash, you are left with a duplicate draft rather than none, and the result says so. Drafts with embedded (cid:) images can be edited: an image the edited body still displays keeps its identifier, an image the body no longer displays is taken off the draft if this server put it there and becomes a regular attachment if it came from elsewhere, and the result says what the draft ended up embedding. Two body shapes still can\'t be rebuilt faithfully and are rejected — a body part that is neither text nor a carriable image, audio, video or attached message, and a body that interleaves two parts of the same text type — recreate those drafts instead. If the draft\'s stored body already references an image that isn\'t attached, editing its body is rejected until the edit resolves that (replace the body, or add an attachments item supplying the missing cid); metadata and attachment edits still work on such a draft.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: "The ID of the draft email to edit (this draft's own id). The message it replies to or forwards, needed to regenerate a quote or forwarded block, is a separate param: originalEmailId.",
            },
            to: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated recipient email addresses (optional, keeps existing if omitted). Each entry may be "Name <email>" or a bare address.',
            },
            cc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated CC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            bcc: {
              type: 'array',
              items: { type: 'string' },
              description: 'Updated BCC email addresses (optional). Each entry may be "Name <email>" or a bare address.',
            },
            from: {
              type: 'string',
              description: 'Updated sender email address (optional)',
            },
            subject: {
              type: 'string',
              description: 'Updated email subject (optional)',
            },
            textBody: {
              type: 'string',
              description: 'Updated plain-text body (optional). Provide it for a genuinely plain message, or alongside htmlBody for a custom plain-text alternative in place of the auto-generated one. Editing textBody alone while htmlBody is present is rejected (the fallback is auto-managed). For a TEXT-ONLY reply or forward draft, editing textBody is a quote-bearing edit: pass originalEmailId to rebuild and keep the quoted/forwarded original, or noQuote:true to drop it. Must be a plain string: a body wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'Updated HTML body (optional), the preferred format. Editing it alone regenerates the plain-text fallback from the new HTML automatically, using the alt text of each image, and \"[image]\" for an embedded (cid:) image that has none. For a REPLY or FORWARD draft that carries the quoted/forwarded original, editing the body is rejected unless you pass originalEmailId (rebuilds and keeps it) or noQuote:true (drops it). Supplying htmlBody to a text-only reply draft converts it to HTML. Pass REAL markup — an entirely HTML-escaped body (escaped element tags like &lt;p&gt; with no actual elements) is rejected, as is any body containing a CDATA section.',
            },
            originalEmailId: {
              type: 'string',
              description: "When editing the body of a REPLY or FORWARD draft, the id of the message this draft replies to OR forwards (NOT this draft's own id, which is emailId). Pass it to rebuild the body and keep the quoted original / forwarded-message block. Without it, a body edit that would drop them is rejected.",
            },
            noQuote: {
              type: ['boolean', 'string'],
              description: "Set true to DISCARD the quoted original (reply draft) or the forwarded-message block (forward draft) when editing the body, instead of keeping it via originalEmailId. On a forward draft this also clears the forward marking (the recorded X-Forwarded-Message-Id), so later edits aren't re-challenged and send_draft will not mark the original forwarded — this applies on ANY edit that passes it, including a metadata-only edit or an asAttachment forward's note edit (the deliberate way to de-forward such a draft). Deliberate-discard escape; without it a quote-dropping body edit is rejected. Cannot be combined with originalEmailId.",
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender). Each entry may be "Name <email>" or a bare address.',
            },
            attachments: attachmentsSchemaProperty(true),
            removeAttachments: {
              type: 'array',
              items: { type: 'string' },
              description: "Attachments to remove from the draft, identified by blobId (from get_email_attachments) or, if unambiguous, by name. A ref that matches no attachment, or a name matching more than one, is rejected — use the blobId. This reaches everything get_email_attachments lists, including images embedded in the body. Removing an image the surviving body still displays is rejected: remove its <img> reference in the same call, or keep the attachment. Removing an image supplied by a quote you are keeping (originalEmailId) is also rejected, because rebuilding the quote re-embeds it — use noQuote with a replacement body to drop the quote and its images instead. To remove every attachment, use clearFields:['attachments'] instead.",
            },
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['to', 'cc', 'bcc', 'replyTo', 'subject', 'textBody', 'htmlBody', 'attachments'] },
              description: "Field names to deliberately clear (to empty/none). Allowed: to, cc, bcc, replyTo, subject, textBody, htmlBody, attachments. `from` cannot be cleared. Cannot also pass the same field as a value (e.g. attachments + clearFields:['attachments'] is rejected). clearFields:['attachments'] takes off every part, images embedded in the body included, and is rejected when the surviving body still references one of them (rewrite or clear that body in the same call). On an edit that also keeps a quote via originalEmailId, the rebuilt quote re-embeds the images it supplies and the result says how many.",
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'send_draft',
        description: 'Send an existing draft email. This is the ONLY tool that transmits mail: every compose tool (create_draft, reply_email, forward_email) saves a draft as stored, inspectable bytes, and this tool submits it. The draft must have recipients (to/cc/bcc) and a from address. After sending, the email is moved to the Sent folder and the draft keyword is removed. An HTML-only draft with real content sends as-is: that is a draft whose html yields no derivable text at all, e.g. one showing a remote image with no alt text (a draft displaying an embedded (cid:) image is not in that group — it derives \"[image]\" and carries a text part). Only a genuinely empty body part (e.g. a blank htmlBody alongside real text) is rejected, because it would render blank to recipients — edit the draft to supply or clear that body first. Thread state is maintained after sending: a draft that replies to a message (In-Reply-To) marks that original answered and read, and a draft that forwards one (X-Forwarded-Message-Id, set by forward_email) marks it forwarded and read. Drafts made by reply_email/forward_email also record WHICH stored copy they were composed from, so when several copies of the original exist (e.g. a self-addressed message filed in two folders) exactly that copy is marked. Best-effort — on a draft without that record the original is found from the Message-ID; when the mark succeeds the result says so, and when the message cannot be identified (no match, or several copies and no record of which) the result says it was not marked and why. A draft that records neither header marks nothing (an ordinary compose). The result also reports how many embedded (cid:) images the message carried out, read off the draft as submitted — a receipt on what was sent, not a check: this tool never refuses a draft over its images.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'The ID of the draft email to send',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'search_emails',
        description: 'Search emails. Provide a free-text query matched across subject, body, and participants (plain words — NOT operator syntax: "from:alice" is matched literally; for structured matching use this tool\'s own from/to/cc/bcc/subject params). All filters combine with AND. Trash and Spam are excluded by default (deleted mail lives in Trash; set includeTrash/includeSpam to include them); drafts are included. Set mailbox (incl. Trash/Spam) to search exactly that mailbox, which ignores the default exclusion. ' + SCOPE_RELIABILITY_CONTRACT + ' Recovery example: if a search returns a "2 in Trash excluded" note, re-run with mailbox:"trash" (or includeTrash:true) to find the deleted message. Returns simplified format (metadata + preview, no bodies); use raw=true for original JMAP, get_email for bodies. The date field is local time with a UTC offset (raw=true returns canonical JMAP UTC). ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' query is optional: search_emails with no query returns recent mail matching only the structural filters (for a plain folder listing use list_emails). limit default 20, max 100. ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Free-text terms matched across subject, body, and participants. Operator syntax like "from:"/"to:"/"subject:" (or boolean AND/OR) is matched literally, not parsed — use the dedicated from/to/cc/bcc/subject params for those. Optional.',
            },
            from: {
              type: 'string',
              description: 'Filter by sender email',
            },
            to: {
              type: 'string',
              description: 'Filter by recipient (To) email',
            },
            cc: {
              type: 'string',
              description: 'Filter by Cc email (mail where this address is cc\'d)',
            },
            bcc: {
              type: 'string',
              description: 'Filter by Bcc email',
            },
            subject: {
              type: 'string',
              description: 'Filter by subject',
            },
            hasAttachment: {
              type: 'boolean',
              description: 'Filter emails with attachments. This is the server\'s own hasAttachment property, compared directly (RFC 8621 §4.4.1) — the same heuristic value the results report, not a count of parts. It answers "is there content attached", so a message whose only image is a small embedded one (a signature logo) is filtered OUT by hasAttachment:true. There is no server-side filter for embedded images; to find those, fetch candidates with get_email and read the attachment entries.',
            },
            isUnread: {
              type: 'boolean',
              description: 'true = only unread; false = only read',
            },
            isPinned: {
              type: 'boolean',
              description: 'true = only pinned/flagged; false = only un-pinned',
            },
            mailbox: {
              type: 'string',
              description: MAILBOX_PARAM_DESC,
            },
            after: {
              type: 'string',
              description: 'Only emails received at or after this time. Accepts a date ("2026-07-20") or a full datetime ("2026-07-20T14:30:00Z", or with an offset such as "2026-07-20T14:30:00+01:00") - no other format, so no unpadded/slash-separated dates and no free text like "20 July 2026". A date-only value means 00:00:00 UTC on that date, so it includes the whole of that day. An empty string is rejected; omit the parameter to search without a start bound.',
            },
            before: {
              type: 'string',
              description: 'Only emails received before this time (exclusive). Accepts a date ("2026-07-20") or a full datetime ("2026-07-20T14:30:00Z", or with an offset) - no other format, so no unpadded/slash-separated dates and no free text like "20 July 2026". A date-only value means 00:00:00 UTC on that date, so it excludes that whole day; pass the following date to include it. An empty string is rejected; omit the parameter to search without an end bound.',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of results (default: 20, max: 100)',
              default: 20,
            },
            position: positionSchemaProperty(),
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
            excludeDrafts: {
              type: 'boolean',
              description: EXCLUDE_DRAFTS_DESC,
            },
            includeTrash: {
              type: 'boolean',
              description: INCLUDE_TRASH_DESC,
            },
            includeSpam: {
              type: 'boolean',
              description: INCLUDE_SPAM_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'list_contacts',
        description: 'List contacts from the address book. Returns simplified format by default (name, emails, phones, org). Use verbose=true only if you need extra fields like addresses, titles, or photos. Use raw=true for original JMAP response.',
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: 'number',
              description: 'Maximum number of contacts to return (default: 50)',
              default: 50,
            },
            verbose: {
              type: 'boolean',
              description: 'Include extra contact fields (addresses, titles, URLs, photos, anniversaries). Not needed for most tasks.',
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'get_contact',
        description: 'Get a specific contact by ID. Returns simplified format by default (name, emails, phones, org). Use verbose=true only if you need extra fields like addresses, titles, or photos. Use raw=true for original JMAP response.',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to retrieve',
            },
            verbose: {
              type: 'boolean',
              description: 'Include extra contact fields (addresses, titles, URLs, photos, anniversaries). Not needed for most tasks.',
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'search_contacts',
        description: 'Search contacts by name or email. Returns simplified format by default (name, emails, phones, org). Use verbose=true only if you need extra fields like addresses, titles, or photos. Use raw=true for original JMAP response.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of results (default: 20)',
              default: 20,
            },
            verbose: {
              type: 'boolean',
              description: 'Include extra contact fields (addresses, titles, URLs, photos, anniversaries). Not needed for most tasks.',
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'list_calendars',
        description: 'List all calendars',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'list_calendar_events',
        description: 'List events from a calendar',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar (optional, defaults to all calendars)',
            },
            startDate: {
              type: 'string',
              description: 'Filter events starting from this date (ISO 8601, e.g. 2026-03-23T00:00:00Z)',
            },
            endDate: {
              type: 'string',
              description: 'Filter events ending before this date (ISO 8601, e.g. 2026-03-30T00:00:00Z)',
            },
            limit: {
              type: 'number',
              description: 'Maximum number of events to return (default: 50)',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_calendar_event',
        description: 'Get a specific calendar event by ID. Returns organizer and participants when available.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to retrieve',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'create_calendar_event',
        description: 'Create a new calendar event. Supports date-only (e.g. 2026-04-01) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs end: 2026-04-02.',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar to create the event in',
            },
            title: {
              type: 'string',
              description: 'Event title',
            },
            description: {
              type: 'string',
              description: 'Event description (optional)',
            },
            start: {
              type: 'string',
              description: 'Start time in ISO 8601 format (e.g. 2026-04-07T14:00:00Z) or date-only for all-day events (e.g. 2026-04-07)',
            },
            end: {
              type: 'string',
              description: 'End time in ISO 8601 format. For all-day events, DTEND is exclusive — a one-day event on April 1 requires end: 2026-04-02',
            },
            location: {
              type: 'string',
              description: 'Event location (optional)',
            },
            participants: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  email: { type: 'string', description: 'Participant email address' },
                  name: { type: 'string', description: 'Participant display name (optional)' }
                },
                required: ['email'],
              },
              description: 'Event participants (optional). Automatically adds ORGANIZER from CalDAV username.',
            },
          },
          required: ['calendarId', 'title', 'start', 'end'],
        },
      },
      {
        name: 'update_calendar_event',
        description: 'Update an existing calendar event. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty/whitespace string for title, description, or location is rejected (use clearFields to delete description/location). Floating times preserve the original timezone; explicit UTC/offset times convert to UTC. WARNING: providing participants replaces ALL existing attendee data (acceptance status, roles, etc.). participants: [] removes all attendees.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to update',
            },
            title: {
              type: 'string',
              description: 'New event title',
            },
            description: {
              type: 'string',
              description: 'New event description',
            },
            start: {
              type: 'string',
              description: 'New start time in ISO 8601 format. Floating times (no Z/offset) preserve original timezone',
            },
            end: {
              type: 'string',
              description: 'New end time in ISO 8601 format. DTEND is exclusive per RFC 5545',
            },
            location: {
              type: 'string',
              description: 'New event location',
            },
            participants: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  email: { type: 'string', description: 'Participant email address' },
                  name: { type: 'string', description: 'Participant display name (optional)' }
                },
                required: ['email'],
              },
              description: 'Replaces ALL existing attendees. Empty array removes all attendees. Omit to preserve existing attendees.',
            },
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['description', 'location'] },
              description: 'Property names to delete from the event. Allowed: description, location. Cannot also pass the same field as a value.',
            },
            confirmRecurring: {
              type: 'boolean',
              description: 'Required when changing start/end on a recurring event with exceptions. Acknowledges that orphaned exception overrides will be removed.',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'delete_calendar_event',
        description: 'Delete a calendar event by ID',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to delete',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'list_identities',
        description: "List sending identities (email addresses that can be used for sending). Returns simplified format by default (name, email, replyTo, and the identity's configured signature as textSignature/htmlSignature when it has one; a blank signature is omitted). Nothing appends the signature for you — JMAP does not do it server-side and neither does this server, so to sign a message read it from here and include it in the body you pass to create_draft/reply_email/forward_email (above the quoted history on a reply). Use verbose=true only if you need extra fields like SMTP config or verification state. Use raw=true for original JMAP response.",
        inputSchema: {
          type: 'object',
          properties: {
            verbose: {
              type: 'boolean',
              description: 'Include extra identity fields (SMTP config, verification state). Not needed for most tasks.',
            },
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'get_recent_emails',
        description: 'Get the most recent emails from a single mailbox (defaults to Inbox), max 50. Pass mailbox:"trash" (or any id/role/name) to read that folder directly. This is Inbox-only with no Trash/Spam/draft flags; for an all-folder view (with the default Trash/Spam exclusion and a hidden-count note) use list_emails. Returns simplified format (metadata + preview, no bodies). Use raw=true for original JMAP response. For email bodies, use get_email. The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: ['number', 'string'],
              description: 'Number of recent emails to retrieve (default: 10, max: 50)',
              default: 10,
            },
            position: positionSchemaProperty(),
            mailbox: {
              type: 'string',
              description: 'Mailbox to read — id, role, or name (default: inbox). Unknown mailbox is rejected with the valid list.',
              default: 'inbox',
            },
            ascending: {
              type: 'boolean',
              description: 'Sort oldest first instead of newest first (default: false)',
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
        },
      },
      {
        name: 'mark_email_read',
        description: 'Mark an email as read or unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false to mark as unread',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'pin_email',
        description: 'Pin or unpin an email',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'delete_email',
        description: 'Delete an email (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to delete',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'move_email',
        description: 'Move an email to a different mailbox (replaces its mailbox membership). The destination accepts an id, role, or name. An unknown destination is rejected with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to move',
            },
            targetMailbox: {
              type: 'string',
              description: 'Destination mailbox — id, role (e.g. archive, trash), or name. Unknown mailbox is rejected with the valid list.',
            },
          },
          required: ['emailId', 'targetMailbox'],
        },
      },
      {
        name: 'add_labels',
        description: 'Add labels (mailboxes) to an email without removing existing ones. Each label mailbox may be given by id, role (e.g. archive, trash), or name; an unknown mailbox rejects the whole call with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailboxes to add as labels — each an id, role (e.g. archive, trash), or name. Any unresolved mailbox rejects the whole call with the valid list.',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'remove_labels',
        description: 'Remove specific labels (mailboxes) from an email. Each label mailbox may be given by id, role (e.g. archive, trash), or name; an unknown mailbox rejects the whole call with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailboxes to remove as labels — each an id, role (e.g. archive, trash), or name. Any unresolved mailbox rejects the whole call with the valid list.',
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'get_email_attachments',
        description: 'List an email\'s parts, as raw JMAP part objects (partId, blobId, type, size, name, disposition, cid) rather than the simplified shape the read tools return. ' + UNION_SCOPE_DESC + ' A body-embedded part usually reports disposition:null rather than "inline", and nothing in this raw listing tells it apart from a genuinely attached file — the derived isInline flag lives only in get_email (and get_thread with includeBodies), so cross-check there before acting on an entry, e.g. before handing its blobId to edit_draft removeAttachments, which would strip an image the body still displays. This listing is also what download_attachment counts from: its first entry is attachmentId "0".',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
            raw: {
              type: 'boolean',
              description: 'Return the JMAP attachments array alone. This is a SET escape, not a shape escape: the entries here are already raw JMAP objects, so raw changes only WHICH parts are listed — it drops the body-embedded parts that the JMAP attachments array does not contain. How many were withheld is reported as a separate second message, so the JSON itself stays parseable. download_attachment entry numbers always count from the full listing, so a raw listing is not an index basis.',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'download_attachment',
        description: 'Download an email attachment. If path is provided, saves the file to disk and returns the file path and size. Otherwise returns a download URL.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
            attachmentId: {
              type: 'string',
              description: 'Which part to download. Four accepted forms, resolved in this fixed order: (1) a partId from get_email_attachments; (2) a blobId; (3) cid:<value> for an embedded image, using the cid from get_email — the cid: prefix is REQUIRED for this form, only the first one is stripped (cid:cid:x looks up the Content-ID "cid:x"), and a value matching more than one part is rejected rather than guessed at; (4) a plain entry number (0, 1, 2, ...) counting from the start of the get_email_attachments listing. Digits resolve as a partId FIRST — Fastmail partIds are themselves digit strings — so the entry-number form applies only when no part claims that value. A number with anything else in it (3a, -1, 1.5) is rejected rather than silently read as an entry number. Entry numbers are positional: they shift whenever the listing does, so prefer a partId, blobId or cid when you will reuse the reference.',
            },
            path: {
              type: 'string',
              description: `File path to save the attachment to. May be absolute or relative; relative paths resolve against ${getDownloadDir() || '~/Downloads/fastmail-mcp/'} (configurable via FASTMAIL_DOWNLOAD_DIR), so a bare filename lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected for security. To save directly into your own location, set FASTMAIL_DOWNLOAD_DIR to that root. Parent directories will be created automatically.`,
            },
          },
          required: ['emailId', 'attachmentId'],
        },
      },
      {
        name: 'get_thread',
        description: 'Get all emails in a conversation thread. Returns simplified format (metadata + preview, no bodies) unless you set includeBodies=true, which returns each message\'s plain-text body in the SAME call — use it (ideally with stripQuoted=true) to read or transcribe a whole conversation instead of issuing one get_email per message. Use raw=true for original JMAP response. The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' Drafts are excluded by default (asymmetric by design — a draft reply is noise when reading a conversation); when any are present a note reports how many are hidden so you can tell a draft reply already exists. A draft that now lives only in Trash is neither shown nor counted (it is not an active draft). Set includeDrafts=true to include them. By default this tool fetches no attachment parts. ' + COMPACT_ATTACHMENT_DESC + ' Under includeBodies each message gains its attachment entries too. ' + UNION_SCOPE_DESC + ' ' + INLINE_PAIR_DESC + ' ' + FIELDS_TOOL_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            threadId: {
              type: 'string',
              description: 'ID of the thread/conversation',
            },
            includeDrafts: {
              type: 'boolean',
              description: 'Include draft messages in the thread (default: false, drafts excluded; a note still reports how many were hidden). Note: search_emails/list_emails differ on BOTH axes — they use excludeDrafts AND include drafts by default.',
            },
            includeBodies: {
              type: 'boolean',
              description: 'Return each message\'s plain-text body (bodyText) alongside its metadata, turning an N-message conversation read into one call. HTML bodies are never returned here (that is where the size risk lives) — a message with no plain-text part is flagged bodyTextUnavailable:true, fetch that one with get_email verbose=true. Hidden drafts are excluded before bodies are read, so an in-progress reply never lands in a transcription. The combined bodies are capped at 100000 bytes: over that the call fails with a message naming the largest messages and telling you to add stripQuoted=true or fetch them individually, rather than silently truncating a body. This mode also adds each message\'s attachment entries, which are extra payload the cap does NOT measure — it counts bodies only. Attachment bytes are never inlined, but the entries themselves carry sender-supplied names and Content-IDs of unbounded length, so on a thread with many attachment-heavy messages project them away with fields (e.g. fields:["id","from","date","bodyText"]).',
            },
            stripQuoted: {
              type: 'boolean',
              description: 'Requires includeBodies. Strips quoted history from every returned body, so the response is each message\'s new text only — the shape most read-a-conversation tasks want. ' + STRIP_QUOTED_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: 'boolean',
              description: 'Return original JMAP response instead of simplified format',
            },
          },
          required: ['threadId'],
        },
      },
      {
        name: 'get_mailbox_stats',
        description: 'Get statistics for a mailbox (unread count, total emails, etc.). Pass mailbox as an id, role, or name; omit it for stats across all mailboxes. An unknown mailbox is rejected with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            mailbox: {
              type: 'string',
              description: 'Mailbox to report on — id, role, or name (optional, defaults to all mailboxes). Unknown mailbox is rejected with the valid list.',
            },
          },
        },
      },
      {
        name: 'get_account_summary',
        description: 'Get overall account summary with statistics',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'bulk_mark_read',
        description: 'Mark multiple emails as read/unread',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to mark',
            },
            read: {
              type: 'boolean',
              description: 'true to mark as read, false as unread',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_pin',
        description: 'Pin or unpin multiple emails',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to pin/unpin',
            },
            pinned: {
              type: 'boolean',
              description: 'true to pin, false to unpin',
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_move',
        description: 'Move multiple emails to a mailbox (replaces each one\'s mailbox membership). The destination accepts an id, role, or name. An unknown destination is rejected with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to move',
            },
            targetMailbox: {
              type: 'string',
              description: 'Destination mailbox — id, role (e.g. archive, trash), or name. Unknown mailbox is rejected with the valid list.',
            },
          },
          required: ['emailIds', 'targetMailbox'],
        },
      },
      {
        name: 'bulk_delete',
        description: 'Delete multiple emails (move to trash)',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to delete',
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_add_labels',
        description: 'Add labels to multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. archive, trash), or name; an unknown mailbox rejects the whole call with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to add labels to',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailboxes to add as labels — each an id, role (e.g. archive, trash), or name. Any unresolved mailbox rejects the whole call with the valid list.',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'bulk_remove_labels',
        description: 'Remove labels from multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. archive, trash), or name; an unknown mailbox rejects the whole call with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of email IDs to remove labels from',
            },
            mailboxIds: {
              type: 'array',
              items: { type: 'string' },
              description: 'Array of mailboxes to remove as labels — each an id, role (e.g. archive, trash), or name. Any unresolved mailbox rejects the whole call with the valid list.',
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'check_function_availability',
        description: 'Check which MCP functions are available based on account permissions. Calendar tools run over CalDAV, so calendar is reported available when CalDAV credentials are configured, regardless of the JMAP calendar capability.',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      },
      {
        name: 'test_bulk_operations',
        description: 'Test bulk operations by finding recent emails and performing safe operations (mark read/unread)',
        inputSchema: {
          type: 'object',
          properties: {
            dryRun: {
              type: 'boolean',
              description: 'If true, only shows what would be done without making changes (default: true)',
              default: true,
            },
            limit: {
              type: 'number',
              description: 'Number of emails to test with (default: 3, max: 10)',
              default: 3,
            },
          },
        },
      },
];

// Per-tool allowed parameter keys + escape hatch, derived once from the same
// inputSchema clients see. Drives the unknown-parameter guard (#11).
const TOOL_SCHEMAS = new Map(
  TOOLS.map(t => [
    t.name,
    {
      keys: new Set(Object.keys((t.inputSchema as any).properties ?? {})),
      additional: (t.inputSchema as any).additionalProperties === true,
    },
  ])
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: TOOLS };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    // Reject unknown/misspelled parameters before touching credentials, so the
    // error doesn't depend on the client being initialized. Key-strictness only;
    // value coercion (coerceBool/coerceStringArray/...) is unaffected. (#11)
    const schema = TOOL_SCHEMAS.get(name);
    if (schema) assertKnownParams(name, args as any, schema.keys, schema.additional);

    const client = initializeClient();

    switch (name) {
      case 'list_mailboxes': {
        const { verbose, raw } = args as any;
        const mailboxes = await client.getMailboxes();
        const output = raw ? mailboxes : mailboxes.map(m => simplifyMailbox(m, { verbose: !!verbose }));
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'list_emails': {
        const { mailbox, limit, ascending, raw } = args as any;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw: !!raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position);
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.getEmails({
          mailbox,
          limit: validLimit,
          position,
          ascending: !!ascending,
          includeTrash: coerceBool((args as any).includeTrash) ?? false,
          includeSpam: coerceBool((args as any).includeSpam) ?? false,
          excludeDrafts: coerceBool((args as any).excludeDrafts) ?? false,
        });
        // Append the exclusion note (if any) to the formatter's string — same out-of-band
        // discipline on both raw + simplified; the JSON block stays parseable.
        const body = raw ? formatRawEmailQueryResult(result) : formatEmailQueryResult(result, { fields });
        return {
          content: [
            {
              type: 'text',
              text: body + buildExclusionNote(result.exclusion),
            },
          ],
        };
      }

      case 'get_email': {
        const { emailId, verbose, raw, stripQuoted } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        // Validated before the fetch so a typo'd field name costs no round trip.
        // Selecting bodyHtml implies verbose's includeHtml: without that, projecting a
        // field the simplifier never emitted would return {} — the trap the parameter
        // exists to avoid (#69).
        const fields = parseEmailFields((args as any).fields, { raw: !!raw });
        // Rejected before the fetch: raw is unmodified JMAP, so honouring stripQuoted
        // there would be impossible and ignoring it would be silent (#73).
        const strip = coerceBool(stripQuoted) ?? false;
        assertStripQuotedNotRaw(strip, !!raw);
        const email = await client.getEmailById(emailId);
        const simplified = simplifyEmail(email, { includeHtml: !!verbose || wantsHtmlBody(fields), stripQuoted: strip });
        return {
          content: [
            {
              type: 'text',
              text: raw ? JSON.stringify(email, null, 2) : JSON.stringify(projectEmail(simplified, fields), null, 2),
            },
          ],
        };
      }

      case 'reply_email': {
        // The orchestration (fetch original, assemble reply, upload + thread attachments,
        // save the draft) lives in composeReply so it is unit-testable with a mock client;
        // this handler just maps the result to the response text.
        const result = await composeReply(args, client, getAttachDir());
        const text = `Reply draft saved successfully (Email ID: ${result.emailId}). Use send_draft to transmit it. Subject: ${result.subject}${formatInlineNotes(result.notes)}`;
        return { content: [{ type: 'text', text }] };
      }

      case 'forward_email': {
        // The orchestration (fetch original, assemble the forwarded-message block,
        // carry/.eml attachments, upload new ones, save the draft) lives in
        // composeForward so it is unit-testable with a mock client; this handler just
        // maps the result to the response text.
        const result = await composeForward(args, client, getAttachDir());
        const inlineNote = result.droppedInlineImages
          ? ` ${result.droppedInlineImages} embedded image(s) were not carried — use asAttachment for full fidelity.`
          : '';
        const text = `Forward draft saved successfully (Email ID: ${result.emailId}).${inlineNote} Use send_draft to transmit it. Subject: ${result.subject}${formatInlineNotes(result.notes)}`;
        return { content: [{ type: 'text', text }] };
      }

      case 'create_draft': {
        // The orchestration (body validation, recipient coercion, the contentless-draft
        // guard, attachment upload, create) lives in composeDraft so it is unit-testable
        // with a mock client; this handler just maps the result to the response text.
        const { emailId, subject, to, cc, notes } = await composeDraft(args, client, getAttachDir());

        const summary = [
          `Draft created successfully (Email ID: ${emailId}).`,
          subject ? `Subject: ${subject}` : null,
          to?.length ? `To: ${to.join(', ')}` : null,
          cc?.length ? `CC: ${cc.join(', ')}` : null,
        ].filter(Boolean).join(' ') + formatInlineNotes(notes);

        return {
          content: [
            {
              type: 'text',
              text: summary,
            },
          ],
        };
      }

      case 'edit_draft': {
        const { emailId, from, subject, textBody, htmlBody, originalEmailId } = args as any;
        const { to, cc, bcc, replyTo } = coerceRecipients(args as any);
        const clearFields = coerceStringArray((args as any).clearFields);
        const removeAttachments = coerceStringArray((args as any).removeAttachments);
        // Coerce to a real boolean (lenient clients send "true"/"false"); a non-bool like
        // "garbage" yields undefined, never true — so it can never silently drop the quote.
        const noQuote = coerceBool((args as any).noQuote) === true;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }

        // Ordering belt only: updateDraft runs the same check and remains the authoritative
        // seam for this tool. Running it here too means a malformed body is refused before
        // any attachment is read off disk and pushed to the blob store, matching the other
        // four compose paths (the guard is a pure, idempotent input check).
        assertBodyInputs(args as any);

        const editAttachmentSpecs = coerceAttachments((args as any).attachments);
        // No inline decision is made here, unlike the compose tools: an edit's shipping body
        // is only settled inside updateDraft (the caller's html merged with the draft's own),
        // so that is where a supplied Content-ID is matched against the body and marked.
        const editAttachments = editAttachmentSpecs?.length
          ? await client.uploadAttachments(editAttachmentSpecs, getAttachDir())
          : undefined;

        const updateResult = await client.updateDraft(emailId, {
          to,
          cc,
          bcc,
          from,
          subject,
          textBody,
          htmlBody,
          replyTo,
          clearFields,
          attachments: editAttachments,
          removeAttachments,
          originalEmailId,
          noQuote,
        }, {
          // Whether this server can attach files at all, which decides whether a refusal
          // over a missing embedded image may suggest supplying it via `attachments`.
          attachmentsEnabled: !!getAttachDir(),
        });

        // JMAP content is immutable, so an edit creates a replacement draft and moves the
        // old one to Trash. The summary reports where that old copy went and what it held,
        // so an unintended overwrite is visible and undoable (#65).
        return {
          content: [
            {
              type: 'text',
              text: formatEditDraftResult(updateResult),
            },
          ],
        };
      }

      case 'send_draft': {
        // The orchestration (submit, then best-effort thread-state maintenance on the
        // message the draft replied to or forwarded) lives in sendDraftAndMaintainKeywords
        // so it is unit-testable with a mock client; this handler just maps the result to
        // the response text.
        const result = await sendDraftAndMaintainKeywords(args, client);

        return {
          content: [
            {
              type: 'text',
              text: formatSendDraftResult(result),
            },
          ],
        };
      }

      case 'list_contacts': {
        const { limit = 50, verbose, raw } = args as any;
        const contactsClient = initializeContactsCalendarClient();
        const result = await contactsClient.getContacts(limit);
        return {
          content: [
            {
              type: 'text',
              text: raw ? formatQueryResult(result) : formatContactQueryResult(result, { verbose: !!verbose }),
            },
          ],
        };
      }

      case 'get_contact': {
        const { contactId, verbose, raw } = args as any;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contact = await contactsClient.getContactById(contactId);
        const output = raw ? contact : simplifyContact(contact, { verbose: !!verbose });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'search_contacts': {
        const { query, limit = 20, verbose, raw } = args as any;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const result = await contactsClient.searchContacts(query, limit);
        return {
          content: [
            {
              type: 'text',
              text: raw ? formatQueryResult(result) : formatContactQueryResult(result, { verbose: !!verbose }),
            },
          ],
        };
      }

      // Calendar operations use CalDAV directly.
      // JMAP Calendars: spec not yet finalized, Fastmail has not enabled JMAP calendar support.
      // Existing JMAP calendar code in contacts-calendar.ts has known bugs and must not be used.
      // When Fastmail enables JMAP calendars: re-enable the path, fix to match finalized spec,
      // do a parity pass with CalDAV implementation, and test against live Fastmail.
      // CalDAV tests should be structured so they can serve as a basis for JMAP tests later.

      case 'list_calendars': {
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const calendars = await davClient.getCalendars();
        return { content: [{ type: 'text', text: JSON.stringify(calendars, null, 2) }] };
      }

      case 'list_calendar_events': {
        const { calendarId, limit = 50, startDate, endDate } = args as any;
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const events = await davClient.getCalendarEvents(calendarId, limit, startDate, endDate);
        return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
      }

      case 'get_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const event = await davClient.getCalendarEventById(eventId);
        return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
      }

      case 'create_calendar_event': {
        const { calendarId, title, description, start, end, location, participants } = args as any;
        if (!calendarId || !title || !start || !end) {
          throw new McpError(ErrorCode.InvalidParams, 'calendarId, title, start, and end are required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const eventId = await davClient.createCalendarEvent({
          calendarId, title, description, start, end, location, participants,
        });
        return { content: [{ type: 'text', text: `Calendar event created. Event ID: ${eventId}` }] };
      }

      case 'update_calendar_event': {
        const { eventId, title, description, start, end, location, participants, clearFields, confirmRecurring } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const hasClearFields = Array.isArray(clearFields) && clearFields.length > 0;
        if (title === undefined && description === undefined && start === undefined && end === undefined && location === undefined && participants === undefined && !hasClearFields) {
          throw new McpError(ErrorCode.InvalidParams, 'At least one field to update must be provided (title, description, start, end, location, participants, or clearFields)');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const fields = { title, description, start, end, location, participants, clearFields, confirmRecurring };
        await davClient.updateCalendarEvent(eventId, fields);
        return { content: [{ type: 'text', text: `Calendar event updated. Event ID: ${eventId}` }] };
      }

      case 'delete_calendar_event': {
        const { eventId } = args as any;
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        await davClient.deleteCalendarEvent(eventId);
        return { content: [{ type: 'text', text: `Calendar event deleted. Event ID: ${eventId}` }] };
      }

      case 'list_identities': {
        const { verbose, raw } = args as any;
        const client = initializeClient();
        const identities = await client.getIdentities();
        const output = raw ? identities : identities.map(i => simplifyIdentity(i, { verbose: !!verbose }));
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      case 'get_recent_emails': {
        const { limit = 10, mailbox = 'inbox', ascending, raw } = args as any;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw: !!raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position) ?? 0;
        const client = initializeClient();
        const result = await client.getRecentEmails(limit, mailbox, !!ascending, position);
        return {
          content: [
            {
              type: 'text',
              text: raw ? formatRawEmailQueryResult(result) : formatEmailQueryResult(result, { fields }),
            },
          ],
        };
      }

      case 'mark_email_read': {
        const { emailId } = args as any;
        const read = coerceBool((args as any).read) ?? true;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.markEmailRead(emailId, read);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'pin_email': {
        const { emailId } = args as any;
        const pinned = coerceBool((args as any).pinned) ?? true;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.pinEmail(emailId, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `Email ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'delete_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.deleteEmail(emailId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email deleted successfully (moved to trash)',
            },
          ],
        };
      }

      case 'move_email': {
        const { emailId, targetMailbox } = args as any;
        if (!emailId || !targetMailbox) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and targetMailbox are required');
        }
        const client = initializeClient();
        await client.moveEmail(emailId, targetMailbox);
        return {
          content: [
            {
              type: 'text',
              text: 'Email moved successfully',
            },
          ],
        };
      }

      case 'add_labels': {
        const { emailId } = args as any;
        const mailboxIds = coerceStringArray((args as any).mailboxIds);
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.addLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to email`,
            },
          ],
        };
      }

      case 'remove_labels': {
        const { emailId } = args as any;
        const mailboxIds = coerceStringArray((args as any).mailboxIds);
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        if (!mailboxIds || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.removeLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from email`,
            },
          ],
        };
      }

      case 'get_email_attachments': {
        const { emailId, raw } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        const result = await client.getEmailAttachments(emailId);
        return { content: buildAttachmentListContent(result, !!raw) };
      }

      case 'download_attachment': {
        const { emailId, attachmentId, path } = args as any;
        if (!emailId || !attachmentId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId and attachmentId are required');
        }
        const client = initializeClient();
        try {
          if (path) {
            const result = await client.downloadAttachmentToFile(emailId, attachmentId, path, getDownloadDir());
            return {
              content: [
                {
                  type: 'text',
                  text: `Saved to: ${result.savedPath} (${result.bytesWritten} bytes)`,
                },
              ],
            };
          } else {
            const downloadUrl = await client.downloadAttachment(emailId, attachmentId);
            return {
              content: [
                {
                  type: 'text',
                  text: `Download URL: ${downloadUrl}`,
                },
              ],
            };
          }
        } catch (error) {
          // Let path-confinement rejections through so the caller sees why their path was
          // rejected. PathAccessError is the tagged discriminator (the path guards no
          // longer prefix "Save path", so a substring match would miss them).
          if (error instanceof PathAccessError) {
            throw new McpError(ErrorCode.InvalidParams, error.message);
          }
          // An unusable attachmentId (a malformed cid: handle, a value that matches
          // several parts, a number with junk in it, a part with no blob) is a caller
          // input error, not a lookup result, so it survives the generic message below —
          // it names what to pass instead and reveals nothing about the mailbox. A
          // reference that is well-formed but simply matches nothing stays generic.
          // Re-thrown rather than mapped locally so the top-level branch applies its
          // InvalidParams mapping WITH redaction (these messages echo caller input).
          if (error instanceof InvalidInputError) {
            throw error;
          }
          // Sanitize other errors to avoid leaking attachment metadata. This branch is
          // retained deliberately: redactBearerTokens at the top level would not suppress
          // attachment metadata in a transport/JMAP error message, so we keep the local
          // generic message instead of letting such errors reach the top-level catch.
          throw new McpError(
            ErrorCode.InternalError,
            'Attachment download failed. Verify emailId and attachmentId and try again.'
          );
        }
      }

      case 'search_emails': {
        const { query, from, to, cc, bcc, subject, hasAttachment, isUnread, isPinned, mailbox, after, before, limit, ascending, raw } = args as any;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw: !!raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position);
        const client = initializeClient();
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.searchEmails({
          query, from, to, cc, bcc, subject,
          hasAttachment: coerceBool(hasAttachment),
          isUnread: coerceBool(isUnread),
          isPinned: coerceBool(isPinned),
          mailbox, after, before, limit: validLimit, position,
          ascending: !!ascending,
          excludeDrafts: coerceBool((args as any).excludeDrafts) ?? false,
          includeTrash: coerceBool((args as any).includeTrash) ?? false,
          includeSpam: coerceBool((args as any).includeSpam) ?? false,
        });
        const body = raw ? formatRawEmailQueryResult(result) : formatEmailQueryResult(result, { fields });
        return {
          content: [
            {
              type: 'text',
              text: body + buildExclusionNote(result.exclusion),
            },
          ],
        };
      }

      case 'get_thread': {
        const { threadId } = args as any;
        if (!threadId) {
          throw new McpError(ErrorCode.InvalidParams, 'threadId is required');
        }
        const client = initializeClient();
        try {
          // The flag guards, the body-size cap and the per-message signals live in
          // readThread so they are unit-testable with a mock client; this stays a
          // result-to-text wrapper.
          const text = await readThread(args, client);
          return {
            content: [
              {
                type: 'text',
                text,
              },
            ],
          };
        } catch (error) {
          // Re-raise the tagged caller-input errors BARE so the top-level catch applies
          // its InvalidParams mapping (with redaction for InvalidInputError) — otherwise
          // this local catch would collapse them to InternalError. getThread throws
          // InvalidInputError on a not-found threadId (a bad id is caller-fixable input);
          // PathAccessError is re-raised too for parity with download_attachment, though
          // get_thread has no path input today. Everything else is an operational thread
          // failure → InternalError (redacted here, since it doesn't reach a tagged branch).
          if (error instanceof PathAccessError || error instanceof InvalidInputError) {
            throw error;
          }
          throw new McpError(ErrorCode.InternalError, `Thread access failed: ${redactBearerTokens(error instanceof Error ? error.message : String(error))}`);
        }
      }

      case 'get_mailbox_stats': {
        const { mailbox } = args as any;
        const client = initializeClient();
        const stats = await client.getMailboxStats(mailbox);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(stats, null, 2),
            },
          ],
        };
      }

      case 'get_account_summary': {
        const client = initializeClient();
        const summary = await client.getAccountSummary();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(summary, null, 2),
            },
          ],
        };
      }

      case 'bulk_mark_read': {
        const read = coerceBool((args as any).read) ?? true;
        const emailIds = coerceStringArray((args as any).emailIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkMarkRead(emailIds, read);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${read ? 'marked as read' : 'marked as unread'} successfully`,
            },
          ],
        };
      }

      case 'bulk_pin': {
        const pinned = coerceBool((args as any).pinned) ?? true;
        const emailIds = coerceStringArray((args as any).emailIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkPinEmails(emailIds, pinned);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails ${pinned ? 'pinned' : 'unpinned'} successfully`,
            },
          ],
        };
      }

      case 'bulk_move': {
        const { targetMailbox } = args as any;
        const emailIds = coerceStringArray((args as any).emailIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!targetMailbox) {
          throw new McpError(ErrorCode.InvalidParams, 'targetMailbox is required');
        }
        const client = initializeClient();
        await client.bulkMove(emailIds, targetMailbox);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails moved successfully`,
            },
          ],
        };
      }

      case 'bulk_delete': {
        const emailIds = coerceStringArray((args as any).emailIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkDelete(emailIds);
        return {
          content: [
            {
              type: 'text',
              text: `${emailIds.length} emails deleted successfully (moved to trash)`,
            },
          ],
        };
      }

      case 'bulk_add_labels': {
        const emailIds = coerceStringArray((args as any).emailIds);
        const mailboxIds = coerceStringArray((args as any).mailboxIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkAddLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels added successfully to ${emailIds.length} emails`,
            },
          ],
        };
      }

      case 'bulk_remove_labels': {
        const emailIds = coerceStringArray((args as any).emailIds);
        const mailboxIds = coerceStringArray((args as any).mailboxIds);
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        if (!mailboxIds || mailboxIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'mailboxIds array is required and must not be empty');
        }
        const client = initializeClient();
        await client.bulkRemoveLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: `Labels removed successfully from ${emailIds.length} emails`,
            },
          ],
        };
      }

      case 'check_function_availability': {
        const client = initializeClient();
        const session = await client.getSession();

        // Calendar tools run on CalDAV, not JMAP (the JMAP calendar path is
        // disabled). So calendar is available if EITHER the JMAP calendar
        // capability is present OR CalDAV credentials are configured.
        const jmapCalendar = !!session.capabilities['urn:ietf:params:jmap:calendars'];
        const caldavConfigured = initializeCalDAVClient() !== null;
        const calendarAvailable = jmapCalendar || caldavConfigured;
        const calendarNote = jmapCalendar
          ? 'Calendar is available (JMAP)'
          : caldavConfigured
            ? 'Calendar is available via CalDAV'
            : 'Calendar access not available - set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD, or enable calendar scope in Fastmail account settings';

        const availability = {
          email: {
            available: true,
            functions: [
              'list_mailboxes', 'list_emails', 'get_email', 'reply_email', 'forward_email', 'create_draft', 'edit_draft', 'send_draft', 'search_emails',
              'get_recent_emails', 'mark_email_read', 'pin_email', 'delete_email', 'move_email',
              'get_email_attachments', 'download_attachment', 'get_thread',
              'get_mailbox_stats', 'get_account_summary', 'bulk_mark_read', 'bulk_pin', 'bulk_move', 'bulk_delete',
              'add_labels', 'remove_labels', 'bulk_add_labels', 'bulk_remove_labels'
            ]
          },
          identity: {
            available: true,
            functions: ['list_identities']
          },
          contacts: {
            available: !!session.capabilities['urn:ietf:params:jmap:contacts'],
            functions: ['list_contacts', 'get_contact', 'search_contacts'],
            note: session.capabilities['urn:ietf:params:jmap:contacts'] ? 
              'Contacts are available' : 
              'Contacts access not available - may require enabling in Fastmail account settings',
            enablementGuide: session.capabilities['urn:ietf:params:jmap:contacts'] ? null : {
              steps: [
                '1. Log into Fastmail web interface',
                '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
                '3. Check if contacts scope is enabled for your API token',
                '4. If not available, you may need to upgrade your Fastmail plan or contact support'
              ],
              documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
            }
          },
          calendar: {
            available: calendarAvailable,
            functions: ['list_calendars', 'list_calendar_events', 'get_calendar_event', 'create_calendar_event', 'update_calendar_event', 'delete_calendar_event'],
            note: calendarNote,
            enablementGuide: calendarAvailable ? null : {
              steps: [
                'Option A (CalDAV): set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD (app password) — calendar tools run over CalDAV',
                'Option B (JMAP scope): 1. Log into Fastmail web interface',
                '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
                '3. Check if calendar scope is enabled for your API token',
                '4. If not available, you may need to upgrade your Fastmail plan or contact support'
              ],
              documentation: 'https://www.fastmail.com/help/technical/jmap-api.html'
            }
          },
          capabilities: Object.keys(session.capabilities)
        };
        
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(availability, null, 2),
            },
          ],
        };
      }

      case 'test_bulk_operations': {
        const { limit = 3 } = args as any;
        // Coerce dryRun for lenient clients: a stringified "false" is otherwise truthy
        // and would silently keep the diagnostic in dry-run mode. Defaults to dry-run
        // (the safe, non-acting direction).
        const dryRun = coerceBool((args as any).dryRun) ?? true;
        const client = initializeClient();
        
        // Get some recent emails to test with
        const testLimit = Math.min(Math.max(limit, 1), 10);
        const { items: emails } = await client.getRecentEmails(testLimit, 'inbox');

        if (emails.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: 'No emails found for bulk operation testing. Try sending yourself a test email first.',
              },
            ],
          };
        }
        
        const emailIds = emails.slice(0, testLimit).map(email => email.id);
        const operations = [
          {
            name: 'bulk_mark_read',
            description: `Mark ${emailIds.length} emails as read`,
            parameters: { emailIds, read: true }
          },
          {
            name: 'bulk_mark_read (undo)',
            description: `Mark ${emailIds.length} emails as unread (undo previous)`,
            parameters: { emailIds, read: false }
          }
        ];
        
        const results = {
          testEmails: emails.map(email => ({
            id: email.id,
            subject: email.subject,
            from: email.from?.[0]?.email || 'unknown',
            receivedAt: email.receivedAt
          })),
          operations: [] as any[]
        };
        
        if (dryRun) {
          results.operations = operations.map(op => ({
            ...op,
            status: 'DRY RUN - Would execute but not actually performed',
            executed: false
          }));
          
          return {
            content: [
              {
                type: 'text',
                text: `BULK OPERATIONS TEST (DRY RUN)\n\n${JSON.stringify(results, null, 2)}\n\nTo actually execute the test, set dryRun: false`,
              },
            ],
          };
        } else {
          // Execute the test operations
          for (const operation of operations) {
            try {
              await client.bulkMarkRead(operation.parameters.emailIds, coerceBool(operation.parameters.read) ?? true);
              results.operations.push({
                ...operation,
                status: 'SUCCESS',
                executed: true,
                timestamp: new Date().toISOString()
              });
              
              // Small delay between operations
              await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
              results.operations.push({
                ...operation,
                status: 'FAILED',
                executed: false,
                // Folded into result JSON rather than raised, so the top-level catch's
                // redaction never sees it — redact here for defense-in-depth parity now
                // that richer #22 reasons flow through this path.
                error: redactBearerTokens(error instanceof Error ? error.message : String(error)),
                timestamp: new Date().toISOString()
              });
            }
          }
          
          return {
            content: [
              {
                type: 'text',
                text: `BULK OPERATIONS TEST (EXECUTED)\n\n${JSON.stringify(results, null, 2)}`,
              },
            ],
          };
        }
      }

      default:
        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${name}`);
    }
  } catch (error) {
    if (error instanceof McpError) {
      throw error;
    }
    // The four compose handlers (reply_email/forward_email/create_draft/edit_draft) have no
    // local try/catch, so an attachment opt-in/path/contentType rejection thrown as a
    // PathAccessError surfaces here. Map it to InvalidParams (actionable) rather than the
    // generic InternalError wrap below. (download_attachment maps its own PathAccessError
    // locally and never reaches here.)
    if (error instanceof PathAccessError) {
      throw new McpError(ErrorCode.InvalidParams, error.message);
    }
    // A semantically-invalid caller input (unresolvable mailbox, label id that is
    // really a name). Map to InvalidParams like PathAccessError above, but DO run it
    // through redactBearerTokens — unlike PathAccessError these messages reflect
    // caller input and mailbox names, so token-shaped redaction is cheap insurance.
    // This branch is placed after the PathAccessError branch and before the generic
    // wrap so an InvalidInputError can't fall through to InternalError.
    if (error instanceof InvalidInputError) {
      throw new McpError(ErrorCode.InvalidParams, redactBearerTokens(error.message));
    }
    const raw = error instanceof Error ? error.message : String(error);
    throw new McpError(
      ErrorCode.InternalError,
      `Tool execution failed: ${redactBearerTokens(raw)}`
    );
  }
});

async function runServer() {
  // Resolve the email-date render zone once, before any tool handler can fire.
  setDefaultTimezone(getTimezone());
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch(() => {
  // Avoid logging raw error objects to prevent accidental PII leakage
  console.error('Fastmail MCP server failed to start');
  process.exit(1);
});