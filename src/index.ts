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
import { formatQueryResult, formatRawEmailQueryResult, formatEmailQueryResult, buildExclusionNote, buildAttachmentListContent, simplifyIdentity, simplifyContact, formatContactQueryResult, formatEditDraftResult, formatSendDraftResult, formatInlineNotes } from './response-formatters.js';
import { coerceRecipients, coerceStringArray, coerceBool, coercePosition, clampLimit, redactBearerTokens, registerSecret, assertKnownParams, coerceAttachments, coerceParticipants, PathAccessError, InvalidInputError } from './coerce.js';
import { parseEmailFields, projectEmail, wantsHtmlBody } from './field-projection.js';
import { composeReply } from './reply-handler.js';
import { composeForward } from './forward-handler.js';
import { sendDraftAndMaintainKeywords } from './send-draft-handler.js';
import { composeDraft } from './compose-handler.js';
import { assertBodyInputs } from './body-format.js';
import { assertStripQuotedNotRaw } from './quote-strip.js';
import { assertICalTextLimits, MAX_ICAL_FIELD_BYTES, MAX_ICAL_PARTICIPANTS, MAX_ICAL_TOTAL_BYTES } from './ical-limits.js';
import { readThread } from './thread-handler.js';
import { listMailboxes, createMailbox } from './mailbox-handler.js';
import { createContactTool, updateContactTool, deleteContactTool } from './contacts-handler.js';
import createDebug from 'debug';

// The calendar text bounds, rendered once in KB for the tool descriptions below so the
// documented numbers can never drift from the numbers actually enforced.
const MAX_ICAL_FIELD_KB = MAX_ICAL_FIELD_BYTES / 1024;
const MAX_ICAL_TOTAL_KB = MAX_ICAL_TOTAL_BYTES / 1024;

// Silence tsdav's debug logging. tsdav logs the HTTP Basic credential as bare base64
// ("tsdav:authHelper Basic auth token generated: <base64 of user:password>") whenever
// DEBUG is set, straight to stderr — it never passes through this file's redaction
// boundary, so redactBearerTokens cannot help. Suppressing the namespaces is the only
// control. Each detail below was established by running it; re-establish before changing:
//
//   * The skip is `-tsdav*` with NO colon. `-tsdav:*` matches only colon-prefixed
//     children, so a logger created as bare `tsdav` — or a future `tsdavFoo` — would
//     still log. The bare glob covers every current namespace (tsdav:account,
//     tsdav:addressBook, tsdav:authHelper, tsdav:calendar, tsdav:collection,
//     tsdav:request) and any sibling a future release adds.
//   * Deleting process.env.DEBUG here does NOT work: the `debug` package snapshots the
//     environment once at its own module init, and under ESM the whole import chain
//     (caldav-client -> tsdav -> debug) evaluates before this module's body runs. This
//     call works instead because each logger's `enabled` getter recomputes from
//     createDebug.namespaces, which enable() rewrites.
//   * Composing with the operator's own DEBUG preserves their logging for every other
//     package; skip entries beat enables, so `DEBUG=*` and `DEBUG=tsdav:*` are both
//     suppressed for tsdav while `other:*` still logs.
//   * Only DEBUG gates the package; NODE_DEBUG is not consulted.
//   * The control is instance-local — it reaches only the `debug` module instance tsdav
//     resolves. package.json therefore depends on `debug` at the exact version tsdav
//     pins, so npm keeps one hoisted copy; a diverging range would let npm nest
//     node_modules/tsdav/node_modules/debug and silently revert this.
createDebug.enable([process.env.DEBUG, '-tsdav*'].filter(Boolean).join(','));

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
  // Register for value-based redaction so an exact token occurrence in any error
  // string is scrubbed even when it doesn't match the `fmu…` token-shape pattern
  // (self-hosted JMAP servers issue tokens of any shape).
  registerSecret(apiToken);

  const baseInfo = findEnvValue([
    'FASTMAIL_BASE_URL',
    'USER_CONFIG_FASTMAIL_BASE_URL',
    'USER_CONFIG_fastmail_base_url',
    'fastmail_base_url',
  ]);

  // Opt-in for self-hosted JMAP servers. Required to use any base URL outside
  // the api.fastmail.com / www.fastmailusercontent.com allowlist (which already
  // covers Fastmail's regional hosts, e.g. phl.api.fastmail.com).
  // Deliberately env-only: this kill switch is not exposed as a DXT user_config
  // key, so it resolves a single name rather than the four-name fallback the
  // configurable settings use.
  const unsafeInfo = findEnvValue([
    'FASTMAIL_ALLOW_UNSAFE_BASE_URL',
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
    'USER_CONFIG_fastmail_caldav_username',
    'fastmail_caldav_username',
  ]).value;
  const password = findEnvValue([
    'FASTMAIL_CALDAV_PASSWORD',
    'USER_CONFIG_FASTMAIL_CALDAV_PASSWORD',
    'USER_CONFIG_fastmail_caldav_password',
    'fastmail_caldav_password',
  ]).value;

  if (!username || !password) return null;

  // Register the CalDAV password for value-based redaction — a Basic-auth
  // credential matches none of the token-shape patterns. The username is
  // deliberately NOT registered: it is an email address, and exact-match
  // scrubbing of it would mangle legitimate output (every message from or to
  // that address). Note the scrubber ignores values under 8 characters, so a
  // very short password is not covered either.
  registerSecret(password);
  // tsdav transmits and logs the Basic credential as bare base64 with no
  // "Basic " prefix, which BASIC_PATTERN cannot match — register the encoded
  // form as its own literal secret.
  registerSecret(Buffer.from(`${username}:${password}`).toString('base64'));

  // The ORGANIZER display name is resolved here, not read from process.env inside the
  // CalDAV client, so every env-derived setting goes through the one four-name lookup
  // that lets a DXT user_config key reach the server. An unset value leaves it
  // undefined and the client falls back to the CalDAV username.
  const displayName = findEnvValue([
    'FASTMAIL_CALDAV_DISPLAY_NAME',
    'USER_CONFIG_FASTMAIL_CALDAV_DISPLAY_NAME',
    'USER_CONFIG_fastmail_caldav_display_name',
    'fastmail_caldav_display_name',
  ]).value;

  caldavClient = new CalDAVCalendarClient({ username, password, displayName });
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

// Shared `participants` schema for the two calendar write tools, so the accepted shapes
// stay identical between create and update. `leadIn` carries the per-tool sentence
// (create adds an ORGANIZER; update REPLACES the whole attendee list).
//
// The type is widened to ['array', 'string'] for the same reason every lenient boolean is
// widened: coerceParticipants accepts a JSON-stringified array from clients that
// stringify structured params, and a client validating against a narrow `type: 'array'`
// would reject that string before dispatch, making the coercion unreachable. The
// description spells out which strings work, because the type alone does not say.
function participantsSchemaProperty(leadIn: string) {
  return {
    type: ['array', 'string'],
    items: {
      type: ['object', 'string'],
      properties: {
        email: { type: 'string', description: `Participant email address (max ${MAX_ICAL_FIELD_KB}KB)` },
        name: { type: 'string', description: `Participant display name (optional, max ${MAX_ICAL_FIELD_KB}KB)` },
      },
      required: ['email'],
    },
    description:
      leadIn +
      ` Each entry is either an object { email, name? } or a bare email-address string (equivalent to { email }).` +
      ` An entry with any other key, a missing or non-string email, or a non-string name is rejected naming its position, e.g. participants[2] — nothing is silently dropped.` +
      ` The whole array may also be sent as a JSON string ('[{"email":"a@example.com"}]'), for clients that stringify structured parameters; a comma-joined list is NOT accepted.`,
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

// Appended to every boolean whose handler runs coerceBool. The schema declares
// `type: ['boolean', 'string']` alongside it, so a validating client can actually send
// the string form; declaring the pair together keeps the advertised type and the runtime
// coercion from drifting apart. A narrow `type: 'boolean'` makes the coercion
// unreachable, which is the failure this note exists to prevent. (#54)
//
// The prose earns its bytes on top of the widened type: `["boolean","string"]` says a
// string is accepted but not WHICH strings, and coerceBool recognises only "true"/"false"
// — anything else falls back to the parameter's default rather than erroring, so a caller
// guessing "1" or "yes" would get the default with no signal.
const LENIENT_BOOL_DESC =
  ' Also accepts the strings "true"/"false", for clients that stringify booleans.';

// Wrap a boolean parameter's description with the note above. Adds sentence-ending
// punctuation first when the description lacks it, so the two clauses don't run
// together — the descriptions here are not uniformly punctuated.
function lenientBool(description: string): string {
  // Look past a trailing bracket before deciding: "...by default.)" is already stopped,
  // while "...(default: true)" is not.
  const needsStop = !/[.!?]$/.test(description.trimEnd().replace(/[)\]]+$/, ''));
  return description + (needsStop ? '.' : '') + LENIENT_BOOL_DESC;
}

// Shared scope-control descriptions for the read tools (search_emails + list_emails),
// defined once so the per-flag strings and the reliability-contract clause stay in sync
// between the two tools rather than drifting as hand-copied strings.
const EXCLUDE_DRAFTS_DESC =
  lenientBool('Drafts are included by default; set true to omit them from results (and from the total count). (Note: get_thread differs on BOTH axes — it uses includeDrafts AND excludes drafts by default.)');
const INCLUDE_TRASH_DESC =
  lenientBool('Trash is excluded by default; set true to also include Trash in the results.');
const INCLUDE_SPAM_DESC =
  lenientBool('The Spam/Junk folder is excluded by default; set true to also include it in the results.');

// `ascending`, declared identically on the three list/search tools (list_emails,
// search_emails, get_recent_emails) so their sort contract can't drift.
const ASCENDING_DESC =
  lenientBool('Sort oldest first instead of newest first (default: false).');

// The reliability contract that makes silence trustworthy. Lead with the no-note
// guarantee as its own sentence (a skimming model must hit "no note => trustworthy"
// first), then the per-signal actions; scoped to the default all-mailbox scope.
const SCOPE_RELIABILITY_CONTRACT =
  'When you search the default scope (no mailbox set): NO note means no Trash/Spam message matched this search — do not re-run with includeTrash/includeSpam just to re-check the same query. ' +
  'A "N in Trash/Spam excluded" note means re-run (includeTrash:true / includeSpam:true, or mailbox:"trash"/"junk") to see those matches. ' +
  'A "count could not be confirmed" or "folder not found; not excluded" note means re-run to be sure. ' +
  'Setting mailbox searches only that folder (no note, by definition).';

// The forms every mailbox-taking parameter accepts, written ONCE and shared by all of
// them — the scalar ones (mailbox / targetMailbox / parent) and the mailboxIds arrays
// alike. They resolve through a single matcher, so a form documented on one tool and not
// another would be a documentation-only difference, and the path form is specified here
// and nowhere else in the schemas.
const MAILBOX_REF_FORMS =
  'Accepts an id, a role (inbox, archive, sent, drafts, trash, junk), a folder name (e.g. Receipts), or a root-anchored path (e.g. Archive/2026/Receipts): "/"-separated, no leading or trailing slash, segments matched case-insensitively. ' +
  'A folder name matching exactly one mailbox wins over reading the same text as a path, so a folder whose own name contains "/" stays reachable by that name. ' +
  'A name shared by several mailboxes is rejected as ambiguous, listing their full paths — retry with one of those, or with the id. An unknown mailbox is rejected with the valid list. ' +
  'list_mailboxes returns each mailbox\'s path, and a path it returns can be pasted straight back into this parameter.';

const MAILBOX_PARAM_DESC =
  'Mailbox to scope to. ' + MAILBOX_REF_FORMS +
  ' Setting it searches exactly that mailbox (incl. Trash/Spam) and ignores the default Trash/Spam exclusion.';

// The mailbox a draft is filed into, shared by nothing else: create_draft is the only tool
// that picks a save destination without moving anything.
const DRAFT_MAILBOX_PARAM_DESC =
  'Mailbox to SAVE the draft into (optional, defaults to Drafts). Does not set From or recipients. ' + MAILBOX_REF_FORMS;

const READ_MAILBOX_PARAM_DESC =
  'Mailbox to read (default: inbox). ' + MAILBOX_REF_FORMS;

const STATS_MAILBOX_PARAM_DESC =
  'Mailbox to report on (optional, defaults to all mailboxes). ' + MAILBOX_REF_FORMS;

// The parent narrowing on list_mailboxes and the nesting parent on create_mailbox are
// different jobs, so they get separate leading sentences over the shared forms.
const LIST_PARENT_PARAM_DESC =
  'Restrict the listing to the DIRECT children of this mailbox (grandchildren are not included). Omit to list every mailbox. ' + MAILBOX_REF_FORMS;

const CREATE_PARENT_PARAM_DESC =
  'Parent mailbox to nest the new mailbox under. Omit to create it at the top level. ' + MAILBOX_REF_FORMS;

// The label arrays, which take the same forms per entry. A function rather than two
// constants so the add/remove verb is the only thing that differs.
const labelMailboxIdsDesc = (verb: 'add' | 'remove') =>
  `Array of mailboxes to ${verb} as labels. Each entry resolves the same way: ` + MAILBOX_REF_FORMS +
  ' Any entry that fails to resolve rejects the whole call, and the error names every failing entry at once.';

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

// Shared verbatim by every tool that shows a caller a Content-ID this server manages
// (get_email and get_email_attachments both emit `cid` values, and a quoted image's is one
// of these), so the warning cannot appear on one read and be missing from the other. It
// matters because the value looks perfectly stable in a single response: it survives edits,
// and only a dropped-and-re-added quote regenerates it. The compose tools reject an authored
// reference to one outright; this is why. (#13)
const MINTED_CID_NONDURABILITY =
  'Server-managed identifiers for quoted images are reused across edits but regenerated when the quote is dropped and re-added — never author references to them.';

// The bound on what a quoted or forwarded body pulls along with it, shared verbatim by
// reply_email and forward_email. Stated on both because it is a bytes-out disclosure and
// the two tools differ only in the escape hatch they offer, never in the bound itself (#13).
const CARRIED_IMAGE_BOUND_DESC =
  'What is carried is bounded only by this: a part is carried when the body references it AND the sender declared it an image (image/*). The content type is sender-declared metadata — nothing is sniffed and nothing verifies the claim — and there is no size limit and no count limit, because the parts are re-referenced by blob rather than uploaded.';

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

// The destination parameter shared by move_email and bulk_move, declared once so the two
// cannot drift on what a destination accepts or on how an unknown one is refused.
const TARGET_MAILBOX_PARAM_DESC =
  'Destination mailbox. ' + MAILBOX_REF_FORMS;

// The membership warning carried by every tool that sets mailboxIds whole-value
// (move_email, bulk_move, archive_email). Written once because the consequence is the
// same on each: a message filed under several labels keeps only the destination, and the
// additive alternative is the one a caller usually wants. The additive tool is a
// parameter rather than a fixed word — a bulk caller sent to the single-email
// `add_labels` would find it rejects their `emailIds`, which is a worse outcome than no
// pointer at all.
const membershipReplaceDesc = (additiveTool: 'add_labels' | 'bulk_add_labels') =>
  'This REPLACES the message\'s entire mailbox membership: every other label/folder it was filed under is removed. ' +
  `To file it somewhere while KEEPING its existing labels, use ${additiveTool} instead.`;

// What the three contacts READ tools return, written once. All three had a hand-copied
// duplicate of this sentence and of the `verbose` parameter text below, which is how the
// wrong field name ("org" for `organization`) survived in all three at once.
const CONTACT_SHAPE_DESC =
  'Returns simplified format by default: id, name, emails, phones, organization, notes. ' +
  'Each emails/phones entry is EITHER a bare string (the address / the number, when the entry ' +
  'carries no label) OR an {address, label} / {number, label} object — so handle both shapes. ' +
  'Use verbose=true for the whole entry objects (contexts, pref, …) plus addresses, titles, ' +
  'URLs, photos and anniversaries. Use raw=true for the original JMAP response.';

// The `verbose` parameter text shared by the same three tools.
const CONTACT_VERBOSE_PARAM_DESC =
  'Return each emails/phones entry whole (contexts, pref and any other stored field) instead ' +
  'of the bare-string-or-{value,label} shape, and include the extra contact fields ' +
  '(addresses, titles, URLs, photos, anniversaries). Not needed for most tasks.';

// The write tools state their success shape, because every one of them returns something a
// caller has to parse rather than a sentence.
const CONTACT_ECHO_DESC =
  'The echo is the ORIGINAL JMAP card in full, never the simplified shape and never affected ' +
  'by verbose/raw: it exists so that whatever a write took away is still visible afterwards, ' +
  'including the per-entry fields (contexts, pref) the simplified shape folds away. Keeping ' +
  'the card is not the same as being able to put it back — this server can rewrite a name, ' +
  'emails, phones, addresses and the note, and nothing else, so photos, titles, ' +
  'organizations, nicknames, URLs, anniversaries, group membership, uid and per-entry ' +
  'contexts/pref would have to be restored in a Fastmail client.';

// Single source of truth for the tool catalog. Hoisted to module scope so the
// CallTool handler can derive each tool's declared parameter set for the
// unknown-parameter guard (#11) — no drift from what clients see via ListTools.
const TOOLS = [
      {
        name: 'list_mailboxes',
        // No `properties` projection parameter, deliberately. Upstream added one so that
        // accounts with hundreds of mailboxes could trim the payload; this server does not
        // expose it, so a client-side option would be surface nothing can reach, and a
        // narrowed set that dropped id/name/role/parentId would silently break the path
        // column and every path-form lookup. The full payload sits well inside the result
        // window on this account's mailbox count; a large-account trim would be a real
        // feature to design, not an option to leave unreachable.
        description: 'List the mailboxes in the Fastmail account. Returns simplified format by default with core fields (name, path, role, counts). Each mailbox carries `path`, its root-anchored "/"-separated location (e.g. Archive/2026/Receipts), which can be passed straight back to any mailbox parameter. Use parent to list one folder\'s direct children. Use verbose=true only if you need extra fields like sortOrder or myRights. Use raw=true for original JMAP response (no path — raw is untransformed JMAP).',
        inputSchema: {
          type: 'object',
          properties: {
            parent: {
              type: 'string',
              description: LIST_PARENT_PARAM_DESC,
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool('Include extra mailbox fields (sortOrder, isSubscribed, myRights). Not needed for most tasks.'),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
        },
      },
      {
        name: 'create_mailbox',
        description: 'Create a mailbox (a Fastmail folder, which is also what a label is). Returns the created mailbox in the same shape list_mailboxes returns, including its `path`, so it can be used as a move destination or a label immediately with no follow-up lookup. `name` is a LEAF name: it must not contain "/", and nesting is expressed with parent. Use raw=true for the original JMAP object, verbose=true for the extra mailbox fields.',
        inputSchema: {
          type: 'object',
          properties: {
            name: {
              type: 'string',
              description: 'Name of the new mailbox, as a single leaf name (e.g. "Receipts"). Must not contain "/" — to nest it, pass the parent parameter.',
            },
            parent: {
              type: 'string',
              description: CREATE_PARENT_PARAM_DESC,
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool('Include extra mailbox fields (sortOrder, isSubscribed, myRights). Not needed for most tasks.'),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
          required: ['name'],
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
              type: ['boolean', 'string'],
              description: ASCENDING_DESC,
            },
            excludeDrafts: {
              type: ['boolean', 'string'],
              description: EXCLUDE_DRAFTS_DESC,
            },
            includeTrash: {
              type: ['boolean', 'string'],
              description: INCLUDE_TRASH_DESC,
            },
            includeSpam: {
              type: ['boolean', 'string'],
              description: INCLUDE_SPAM_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
        },
      },
      {
        name: 'get_email',
        description: 'Get a specific email by ID. Returns simplified format with plain text body (HTML omitted, bodyHtmlSize hint provided). Only use verbose=true if you specifically need the HTML body — it can be very large for marketing emails. Use raw=true for original JMAP response. Set stripQuoted=true to drop quoted reply history from bodyText when reading a message deep in a long thread (the quoted tail is duplicated from earlier messages). The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + UNION_SCOPE_DESC + ' ' + INLINE_PAIR_DESC + ' ' + MINTED_CID_NONDURABILITY + ' ' + FIELDS_TOOL_DESC + ' On this tool fields:["bodyHtml"] returns the HTML body ALONE (no verbose needed, no metadata, no plain-text copy) — the way to read a large HTML draft without the rest of the message pushing the response past the output limit.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to retrieve',
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool('Include HTML body in response. WARNING: can produce very large responses (50K+ chars) for marketing/rich emails. Only use when HTML content is specifically needed.'),
            },
            fields: fieldsSchemaProperty(),
            stripQuoted: {
              type: ['boolean', 'string'],
              description: lenientBool(STRIP_QUOTED_DESC),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'reply_email',
        description: 'Reply to an existing email with proper threading headers (In-Reply-To, References). Automatically fetches the original email to build the reply chain. The subject defaults to \'Re: <original subject>\'; pass subject to override it (see that parameter for what a changed subject does to draft grouping). Use this rather than hand-rolling threading headers on create_draft. This tool always saves the reply as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the reply, it marks the original answered and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft), so when several copies of the original exist the right one is marked. The original message is quoted by default (attributed, top-posted, matching the web client with a portable quote-bar); set quoteOriginal=false to omit it. Quoted HTML is reproduced sanitised (script/style/event handlers stripped; formatting and real http(s) images kept) and is re-sent under your From address. IMAGES THE ORIGINAL DISPLAYED ARE CARRIED INTO THE QUOTE by default, and are re-sent to this reply\'s recipients — so a reply can send image data outward that you never attached. ' + CARRIED_IMAGE_BOUND_DESC + ' The only way to send none of it is quoteOriginal=false, which omits the whole quote; there is no setting for quote text without its images. Carrying needs no FASTMAIL_ATTACH_DIR: the parts are already in the account. A quoted image is only carried when the reply ships an HTML body — a text-only reply drops them, and the result says how many. You can also embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; that one does require FASTMAIL_ATTACH_DIR).',
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
              description: 'Plain-text body (optional). Use it for genuinely plain messages, or alongside htmlBody to provide your own plain-text alternative in place of the auto-generated one. NOTE: a reply with no htmlBody quotes the original as plain text, which cannot show the images the original displayed — they are dropped, and the result says how many. Must be a plain string: a body wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'HTML body (optional), and the preferred format for outgoing mail. When both bodies are supplied, recipients\' clients render this one. Supplying htmlBody alone is fine: a readable plain-text alternative is generated automatically whenever one can be derived from the HTML. In that derivation an image contributes its alt text; an embedded (cid:) image with no alt contributes \"[image]\", so a picture-only message still has a readable text part, while a remote image with no alt contributes nothing. Supplying it is also what lets the quote show the images the original displayed. Pass REAL markup — a body that is entirely HTML-escaped (escaped element tags like &lt;p&gt; with no actual elements) is rejected, because recipients would see the tags as text; so is any body containing a CDATA section, whose contents are dropped from the derived plain-text alternative.',
            },
            quoteOriginal: {
              type: ['boolean', 'string'],
              description: lenientBool('Append the original message as an attributed quote (default true). Set false to omit it — the WHOLE quote, its images included. There is no option for quote text without the images the original displayed: those images are the quoted body, and carrying them is what makes the quote show what the sender wrote. So false is also the only way to stop this reply re-sending them to its recipients (see the tool description for what is carried).'),
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
        description: 'Forward an existing email to new recipients. This tool always saves the forward as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the forward, it marks the original forwarded and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft, on both inline and asAttachment forwards). `to` is required — a forward has no default recipient, unlike reply. A note (textBody/htmlBody) is optional: the forwarded message itself is the content. The original is reproduced below a forwarded-message header block (From/To/Cc/Subject/Date), its HTML sanitised (script/style/event handlers stripped; formatting and real http(s) images kept) and re-sent under your From address. The original\'s regular attachments are carried by default (includeOriginalAttachments). Images the original\'s body DISPLAYED are carried too, and are carried even when includeOriginalAttachments is false, because they are body content rather than attached files — a forward without them would reproduce a message with holes in it. ' + CARRIED_IMAGE_BOUND_DESC + ' Short of not forwarding the message, there is no way to reproduce the body and leave those images behind. Carrying needs no FASTMAIL_ATTACH_DIR: the parts are already in the account. An image the block cannot display — a text-only forward, or a reference this server could not resolve to exactly one image part — rides as a regular attachment instead (subject to includeOriginalAttachments), and the result says so; asAttachment is the lossless alternative. The subject defaults to \'Fwd: <original subject>\'. You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; that one does require FASTMAIL_ATTACH_DIR).',
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
              description: 'Optional note placed ABOVE the forwarded-message block, in plain text — the original is reproduced below it automatically; omit for a bare FYI forward. NOTE: a text-only note produces a PLAIN-TEXT forward (the original\'s HTML formatting is reduced to text, and the images its body displayed cannot be shown — they ride as regular attachments instead, or are left behind when includeOriginalAttachments is false) — use htmlBody, or both, to preserve its formatting and its images. (When asAttachment is set, this note is the whole body — the original rides as the attached .eml, with no inline block.) Must be a plain string: a note wrapped in a CDATA section is rejected.',
            },
            htmlBody: {
              type: 'string',
              description: 'Optional note placed ABOVE the forwarded-message block, in HTML (the preferred format; a plain-text alternative is derived automatically, using the alt text of each image, and \"[image]\" for an embedded (cid:) image that has none) — the original is reproduced below it. (When asAttachment is set, this note is the whole body — the original rides as the attached .eml, with no inline block.) Pass REAL markup — an entirely HTML-escaped note (escaped element tags like &lt;p&gt; with no actual elements) is rejected, as is one containing a CDATA section.',
            },
            includeOriginalAttachments: {
              type: ['boolean', 'string'],
              description: lenientBool("Carry the original message's attached FILES on the forward (default true; all-or-none — to drop individual ones, save the default draft and use edit_draft's removeAttachments). This flag does not govern the images the original's body displayed: those are body content and are carried either way (see the tool description for the bound on that). What it does govern, besides the ordinary files, is an image the forwarded block could not display — that one rides as a regular attachment when this is true, and is left behind when it is false. Ignored when asAttachment is set — the .eml already embeds every original attachment."),
            },
            asAttachment: {
              type: ['boolean', 'string'],
              description: lenientBool('Instead of reproducing the original inline, attach the entire original as a raw .eml file (message/rfc822): lossless, including embedded inline images; supersedes includeOriginalAttachments. NOTE: the raw message carries its full transport headers (Received chain, authentication results) and — when forwarding a message from Sent — any Bcc recipients (see docs/security-model.md), which an inline forward would not expose.'),
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
              description: DRAFT_MAILBOX_PARAM_DESC,
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
              description: lenientBool("Set true to DISCARD the quoted original (reply draft) or the forwarded-message block (forward draft) when editing the body, instead of keeping it via originalEmailId. On a forward draft this also clears the forward marking (the recorded X-Forwarded-Message-Id), so later edits aren't re-challenged and send_draft will not mark the original forwarded — this applies on ANY edit that passes it, including a metadata-only edit or an asAttachment forward's note edit (the deliberate way to de-forward such a draft). Deliberate-discard escape; without it a quote-dropping body edit is rejected. Cannot be combined with originalEmailId."),
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
              type: ['boolean', 'string'],
              description: lenientBool('Filter emails with attachments. This is the server\'s own hasAttachment property, compared directly (RFC 8621 §4.4.1) — the same heuristic value the results report, not a count of parts. It answers "is there content attached", so a message whose only image is a small embedded one (a signature logo) is filtered OUT by hasAttachment:true. There is no server-side filter for embedded images; to find those, fetch candidates with get_email and read the attachment entries.'),
            },
            isUnread: {
              type: ['boolean', 'string'],
              description: lenientBool('true = only unread; false = only read'),
            },
            isPinned: {
              type: ['boolean', 'string'],
              description: lenientBool('true = only pinned/flagged; false = only un-pinned'),
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
              type: ['boolean', 'string'],
              description: ASCENDING_DESC,
            },
            excludeDrafts: {
              type: ['boolean', 'string'],
              description: EXCLUDE_DRAFTS_DESC,
            },
            includeTrash: {
              type: ['boolean', 'string'],
              description: INCLUDE_TRASH_DESC,
            },
            includeSpam: {
              type: ['boolean', 'string'],
              description: INCLUDE_SPAM_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
        },
      },
      {
        name: 'list_contacts',
        description: 'List contacts from the address book. ' + CONTACT_SHAPE_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of contacts to return (default: 20, max: 100). Hard cap, no paging — a larger value is silently reduced to 100 and there is no way to reach the rest.',
              default: 20,
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool(CONTACT_VERBOSE_PARAM_DESC),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
        },
      },
      {
        name: 'get_contact',
        description: 'Get a specific contact by ID. ' + CONTACT_SHAPE_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to retrieve',
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool(CONTACT_VERBOSE_PARAM_DESC),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'search_contacts',
        description: 'Search contacts by name or email. ' + CONTACT_SHAPE_DESC,
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search query string',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of results (default: 20, max: 100). Hard cap, no paging — a larger value is silently reduced to 100 and there is no way to reach the rest.',
              default: 20,
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool(CONTACT_VERBOSE_PARAM_DESC),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
          required: ['query'],
        },
      },
      {
        name: 'create_contact',
        description: 'Create a contact in the address book. Needs at least a name or one email address. Returns the created card, read back after the write, in the same shape get_contact returns (verbose/raw apply). Every entry array accepts both a bare string and an object: emails ["a@b.example"] or [{address, label}], phones ["+1…"] or [{number, label}]; addresses take objects only. An empty array is rejected in every one of them — omit the field instead, the same rule update_contact applies. An unknown per-item key, or a key of the wrong type, is rejected naming its position (e.g. emails[2]).',
        inputSchema: {
          type: 'object',
          properties: {
            name: {
              type: ['object', 'string'],
              description: 'The contact\'s name: a full-name string ("Ada Lovelace"), or an object {given?, surname?, full?}. Required unless at least one email is given.',
              properties: {
                given: { type: 'string', description: 'Given (first) name.' },
                surname: { type: 'string', description: 'Surname (family name).' },
                full: { type: 'string', description: 'The whole name as one string.' },
              },
            },
            emails: {
              type: 'array',
              items: {
                type: ['object', 'string'],
                properties: {
                  address: { type: 'string', description: 'The email address.' },
                  label: { type: 'string', description: 'What this address is for, e.g. "work" or "home".' },
                },
              },
              description: 'Email addresses. Each entry is a bare address string or {address, label?}. Each address may appear once. [] is rejected — omit the field instead.',
            },
            phones: {
              type: 'array',
              items: {
                type: ['object', 'string'],
                properties: {
                  number: { type: 'string', description: 'The phone number.' },
                  label: { type: 'string', description: 'What this number is for, e.g. "mobile" or "work".' },
                },
              },
              description: 'Phone numbers. Each entry is a bare number string or {number, label?}. Each number may appear once. [] is rejected — omit the field instead.',
            },
            addresses: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  full: { type: 'string', description: 'The whole postal address as one string.' },
                  label: { type: 'string', description: 'What this address is for, e.g. "home".' },
                },
              },
              description: 'Postal addresses, as {full, label?} objects. A bare string is NOT accepted here. [] is rejected — omit the field instead.',
            },
            notes: {
              type: 'string',
              description: 'A free-text note on the contact. An empty string is rejected — omit the field instead.',
            },
            addressBookId: {
              type: 'string',
              description: 'Address book to create the contact in. Omit to let the server use the account default.',
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool(CONTACT_VERBOSE_PARAM_DESC),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return the original JMAP card instead of the simplified format'),
            },
          },
        },
      },
      {
        name: 'update_contact',
        description:
          'Update a contact, MERGING per entry rather than overwriting the card. Returns {contact, previousCard}: the updated card (verbose/raw apply to it) and the card exactly as it stood before the write. ' +
          CONTACT_ECHO_DESC +
          ' Only the fields you pass are touched; omit a field to leave it alone. emails/phones merge by value: an entry whose address/number matches one already stored keeps everything the simplified output does not show (contexts, pref, and any other stored field), and only what you supply is written over it. Resending an entry exactly as you read it changes nothing. LABELS ARE ADD-AND-OVERRIDE, NOT A CLEAN REWRITE: a label that differs from the one you read is written as this card\'s `label` property, which then wins here — but Fastmail\'s own apps commonly store the label as a `contexts` set instead, and that set is left as it was, so the two can end up disagreeing outside this server. A label cannot currently be removed at all. A single call that BOTH drops a stored entry AND adds one the card does not have is rejected as ambiguous, and the rejection prints the dropped entries in full so you can resend them losslessly; pass allowEntryReplace:true to go ahead anyway, which rewrites every entry of THAT array from what you supplied and does NOT carry those hidden fields (arrays in the same call that merged cleanly are unaffected). addresses do NOT merge (an address entry has no matchable key) — supplying them replaces the whole set. name merges into the stored structured name: a bare string sets the full name and keeps the given/surname components, and {given}/{surname} update just that part. An empty array (emails: []) is rejected — use clearFields. notes sets a single note, so a card storing more than one is rejected rather than collapsed. A contact GROUP cannot be updated by this tool.',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to update',
            },
            name: {
              type: ['object', 'string'],
              description: 'New name, MERGED into the stored one: a full-name string sets the full name and keeps the stored given/surname components; an object {given?, surname?, full?} updates only the parts it names. Cannot be cleared.',
              properties: {
                given: { type: 'string', description: 'Given (first) name.' },
                surname: { type: 'string', description: 'Surname (family name).' },
                full: { type: 'string', description: 'The whole name as one string.' },
              },
            },
            emails: {
              type: 'array',
              items: {
                type: ['object', 'string'],
                properties: {
                  address: { type: 'string', description: 'The email address.' },
                  label: { type: 'string', description: 'What this address is for, e.g. "work" or "home".' },
                },
              },
              description: 'The complete set of email addresses the contact should end up with, each a bare address string or {address, label?}. Matched against the stored entries by address, so a repeated address keeps its hidden fields. Send an entry back with the label you read and nothing changes; send a DIFFERENT label and it is added as this card\'s `label` property while any `contexts` set the entry already carried stays put, so the label can only be changed or added, never removed. Each address may appear once. [] is rejected — use clearFields.',
            },
            phones: {
              type: 'array',
              items: {
                type: ['object', 'string'],
                properties: {
                  number: { type: 'string', description: 'The phone number.' },
                  label: { type: 'string', description: 'What this number is for, e.g. "mobile" or "work".' },
                },
              },
              description: 'The complete set of phone numbers the contact should end up with, each a bare number string or {number, label?}. Matched against the stored entries by number, so a repeated number keeps its hidden fields. Labels behave as they do for emails: resending the label you read changes nothing, a different label is added as this card\'s `label` property alongside any existing `contexts` set, and a label cannot be removed. Each number may appear once. [] is rejected — use clearFields.',
            },
            addresses: {
              type: 'array',
              items: {
                type: 'object',
                properties: {
                  full: { type: 'string', description: 'The whole postal address as one string.' },
                  label: { type: 'string', description: 'What this address is for, e.g. "home".' },
                },
              },
              description: 'Postal addresses, as {full, label?} objects. These REPLACE the stored set outright — a postal entry has no matchable key, so nothing is merged and any field the stored entries carried is lost. A bare string is NOT accepted here. [] is rejected — use clearFields.',
            },
            notes: {
              type: 'string',
              description: 'Replacement note text. An empty string is rejected — use clearFields:[\'notes\'].',
            },
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['emails', 'phones', 'addresses', 'notes'] },
              description: 'Field names to deliberately empty. Allowed: emails, phones, addresses, notes. `name` cannot be cleared — it is how the contact is identified in every listing, so delete and recreate the card instead. Passing a field as a value AND in clearFields in the same call is rejected.',
            },
            allowEntryReplace: {
              type: ['boolean', 'string'],
              description: lenientBool('Proceed with an emails/phones edit that both drops a stored entry and adds an unknown one, which is otherwise rejected as ambiguous. Scoped to the array that was actually ambiguous: every entry of THAT array is then written fresh from what you supplied, so contexts, pref and any other field the simplified output does not show are NOT carried over — while another array in the same call that merged cleanly still merges. previousCard in the response is what makes that recoverable.'),
            },
            verbose: {
              type: ['boolean', 'string'],
              description: lenientBool(CONTACT_VERBOSE_PARAM_DESC + ' Applies to `contact` only; `previousCard` is always the raw card.'),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return the original JMAP card as `contact` instead of the simplified format. `previousCard` is the raw card either way, and the {contact, previousCard} envelope is returned in every mode.'),
            },
          },
          required: ['contactId'],
        },
      },
      {
        name: 'delete_contact',
        description:
          'Delete a contact. Returns {deleted, deletedCard} — the id, and the full card as it stood immediately before the destroy. ' +
          CONTACT_ECHO_DESC +
          ' This is IRREVERSIBLE: unlike an email, a deleted contact does not go to Trash, so the echoed card is the only copy left — keep it if there is any chance the delete was wrong. create_contact can rebuild the name, emails, phones, addresses and note from it, but NOT the rest of the card (photos, titles, organizations, nicknames, URLs, anniversaries, group membership, the uid, or the per-entry contexts/pref), so recreating gives you a similar contact rather than the one that was deleted. There is deliberately no confirmation parameter.',
        inputSchema: {
          type: 'object',
          properties: {
            contactId: {
              type: 'string',
              description: 'ID of the contact to delete',
            },
          },
          required: ['contactId'],
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
              type: ['number', 'string'],
              description: 'Maximum number of events to return (default: 50, max: 500). Hard cap, no paging — narrow the window with startDate/endDate instead.',
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
        description: `Create a new calendar event. Supports date-only (e.g. 2026-04-01) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs end: 2026-04-02. start and end must use the SAME form — both date-only, both with a zone designator (Z or +HH:MM), or both without one — and end must be later than start; a mismatched or backwards pair is rejected. Write both as strict ISO-8601 (2026-04-07, 2026-04-07T14:00:00, 2026-04-07T14:00:00Z, or 2026-04-07T14:00:00+10:00): other spellings such as 2026/04/07 or "April 7 2026" are rejected rather than guessed at, since guessing would place the event on a day that depends on the server's own time zone, and so is a day that does not exist in its month (2026-02-31). participants entries may be { email, name? } objects or bare email-address strings. Text size limits: title, description, location and each participant name/email are capped at ${MAX_ICAL_FIELD_KB}KB each, at most ${MAX_ICAL_PARTICIPANTS} participants, and all of that text together must stay under ${MAX_ICAL_TOTAL_KB}KB. Oversized input is rejected naming the field and the limit — nothing is silently truncated.`,
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: 'ID of the calendar to create the event in',
            },
            title: {
              type: 'string',
              description: `Event title (max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            description: {
              type: 'string',
              description: `Event description (optional, max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            start: {
              type: 'string',
              description: 'Start time in ISO 8601 format (e.g. 2026-04-07T14:00:00Z) or date-only for all-day events (e.g. 2026-04-07). Must be in the same form as end. Only these spellings are accepted — 2026-04-07, 2026-04-07T14:00:00, ...Z, or ...+HH:MM — and the date must be a real day in its month.',
            },
            end: {
              type: 'string',
              description: 'End time in ISO 8601 format, in the SAME form as start (both date-only, both with a Z/offset, or both without one) and later than start. Same accepted spellings as start. For all-day events, DTEND is exclusive — a one-day event on April 1 requires end: 2026-04-02',
            },
            location: {
              type: 'string',
              description: `Event location (optional, max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            participants: participantsSchemaProperty(
              `Event participants (optional, at most ${MAX_ICAL_PARTICIPANTS}). Automatically adds ORGANIZER from CalDAV username.`,
            ),
          },
          required: ['calendarId', 'title', 'start', 'end'],
        },
      },
      {
        name: 'update_calendar_event',
        description: `Update an existing calendar event. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty/whitespace string for title, description, or location is rejected (use clearFields to delete description/location). Floating times (no Z/offset) preserve the original timezone; explicit UTC/offset times convert to UTC. A new start/end is checked against the value it will sit beside — the other one you passed, or the stored one you left alone: they must end up in the same form (both date-only, both UTC, both floating, or both in the same TZID) and end must be later than start, otherwise the update is rejected. So moving an event to a different day or converting only one side to UTC means passing BOTH start and end. Both are read as strict ISO-8601, matching create_calendar_event: a non-ISO spelling like 2026/04/07, or a day its month does not have (2026-02-31), is rejected rather than guessed at. WARNING: providing participants replaces ALL existing attendee data (acceptance status, roles, etc.). participants: [] removes all attendees, and its entries may be { email, name? } objects or bare email-address strings. Text size limits match create_calendar_event: title, description, location and each participant name/email are capped at ${MAX_ICAL_FIELD_KB}KB each, at most ${MAX_ICAL_PARTICIPANTS} participants, and all of that text together must stay under ${MAX_ICAL_TOTAL_KB}KB. Oversized input is rejected naming the field and the limit — nothing is silently truncated.`,
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to update',
            },
            title: {
              type: 'string',
              description: `New event title (max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            description: {
              type: 'string',
              description: `New event description (max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            start: {
              type: 'string',
              description: 'New start time in ISO 8601 format (2026-04-07, 2026-04-07T14:00:00, ...Z, or ...+HH:MM; the date must be a real day in its month). Floating times (no Z/offset) preserve the original timezone. Must end up in the same form as the end it sits beside (the one you pass, or the stored one) and before it — pass both start and end when moving the event.',
            },
            end: {
              type: 'string',
              description: 'New end time in ISO 8601 format, in the same form as the start it sits beside and later than it. Same accepted spellings as start. DTEND is exclusive per RFC 5545',
            },
            location: {
              type: 'string',
              description: `New event location (max ${MAX_ICAL_FIELD_KB}KB)`,
            },
            participants: participantsSchemaProperty(
              `Replaces ALL existing attendees (at most ${MAX_ICAL_PARTICIPANTS}). Empty array removes all attendees. Omit to preserve existing attendees.`,
            ),
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['description', 'location'] },
              description: 'Property names to delete from the event. Allowed: description, location. Cannot also pass the same field as a value.',
            },
            confirmRecurring: {
              type: ['boolean', 'string'],
              description: lenientBool('Required when changing start/end on a recurring event with exceptions. Acknowledges that orphaned exception overrides will be removed.'),
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
              type: ['boolean', 'string'],
              description: lenientBool('Include extra identity fields (SMTP config, verification state). Not needed for most tasks.'),
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
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
              description: READ_MAILBOX_PARAM_DESC,
              default: 'inbox',
            },
            ascending: {
              type: ['boolean', 'string'],
              description: ASCENDING_DESC,
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
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
              type: ['boolean', 'string'],
              description: lenientBool('true to mark as read, false to mark as unread'),
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
              type: ['boolean', 'string'],
              description: lenientBool('true to pin, false to unpin'),
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
        description: 'Move an email to a different mailbox. ' + membershipReplaceDesc('add_labels') + ' The destination accepts an id, role, name, or path; an unknown or ambiguous destination is rejected with the valid list. No keyword is changed: a moved message keeps its read/unread and flagged state.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to move',
            },
            targetMailbox: {
              type: 'string',
              description: TARGET_MAILBOX_PARAM_DESC,
            },
          },
          required: ['emailId', 'targetMailbox'],
        },
      },
      {
        name: 'archive_email',
        description: 'Archive an email: file it into the account\'s Archive folder, the mailbox carrying the JMAP archive role. ' + membershipReplaceDesc('add_labels') + ' It does NOT mark the message read — archiving and reading are separate actions, so call mark_email_read as well if you want both. The destination is fixed and takes no parameter: it is found by role, never by folder name, so a folder merely NAMED "archive" is not it. To file into any other mailbox use move_email. An account with no archive-role mailbox is rejected, pointing you at move_email. On success it returns a one-line confirmation, not the message — re-read it with get_email if you need its new mailbox membership.',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email to archive',
            },
          },
          required: ['emailId'],
        },
      },
      {
        name: 'add_labels',
        description: 'Add labels (mailboxes) to an email without removing existing ones. Each label mailbox may be given by id, role (e.g. archive, trash), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.',
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
              description: labelMailboxIdsDesc('add'),
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'remove_labels',
        description: 'Remove specific labels (mailboxes) from an email. Each label mailbox may be given by id, role (e.g. archive, trash), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.',
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
              description: labelMailboxIdsDesc('remove'),
            },
          },
          required: ['emailId', 'mailboxIds'],
        },
      },
      {
        name: 'get_email_attachments',
        description: 'List an email\'s parts, as raw JMAP part objects (partId, blobId, type, size, name, disposition, cid) rather than the simplified shape the read tools return. ' + UNION_SCOPE_DESC + ' A body-embedded part usually reports disposition:null rather than "inline", and nothing in this raw listing tells it apart from a genuinely attached file — the derived isInline flag lives only in get_email (and get_thread with includeBodies), so cross-check there before acting on an entry, e.g. before handing its blobId to edit_draft removeAttachments, which would strip an image the body still displays. ' + MINTED_CID_NONDURABILITY + ' This listing is also what download_attachment counts from: its first entry is attachmentId "0".',
        inputSchema: {
          type: 'object',
          properties: {
            emailId: {
              type: 'string',
              description: 'ID of the email',
            },
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return the JMAP attachments array alone. This is a SET escape, not a shape escape: the entries here are already raw JMAP objects, so raw changes only WHICH parts are listed — it drops the body-embedded parts that the JMAP attachments array does not contain. How many were withheld is reported as a separate second message, so the JSON itself stays parseable. download_attachment entry numbers always count from the full listing, so a raw listing is not an index basis.'),
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
              type: ['boolean', 'string'],
              description: lenientBool('Include draft messages in the thread (default: false, drafts excluded; a note still reports how many were hidden). Note: search_emails/list_emails differ on BOTH axes — they use excludeDrafts AND include drafts by default.'),
            },
            includeBodies: {
              type: ['boolean', 'string'],
              description: lenientBool('Return each message\'s plain-text body (bodyText) alongside its metadata, turning an N-message conversation read into one call. HTML bodies are never returned here (that is where the size risk lives) — a message with no plain-text part is flagged bodyTextUnavailable:true, fetch that one with get_email verbose=true. Hidden drafts are excluded before bodies are read, so an in-progress reply never lands in a transcription. The combined bodies are capped at 100000 bytes: over that the call fails with a message naming the largest messages and telling you to add stripQuoted=true or fetch them individually, rather than silently truncating a body. This mode also adds each message\'s attachment entries, which are extra payload the cap does NOT measure — it counts bodies only. Attachment bytes are never inlined, but the entries themselves carry sender-supplied names and Content-IDs of unbounded length, so on a thread with many attachment-heavy messages project them away with fields (e.g. fields:["id","from","date","bodyText"]).'),
            },
            stripQuoted: {
              type: ['boolean', 'string'],
              description: lenientBool('Requires includeBodies. Strips quoted history from every returned body, so the response is each message\'s new text only — the shape most read-a-conversation tasks want. ' + STRIP_QUOTED_DESC),
            },
            fields: fieldsSchemaProperty(),
            raw: {
              type: ['boolean', 'string'],
              description: lenientBool('Return original JMAP response instead of simplified format'),
            },
          },
          required: ['threadId'],
        },
      },
      {
        name: 'get_mailbox_stats',
        description: 'Get statistics for a mailbox (unread count, total emails, etc.). Pass mailbox as an id, role, name, or path; omit it for stats across all mailboxes. An unknown or ambiguous mailbox is rejected with the valid list.',
        inputSchema: {
          type: 'object',
          properties: {
            mailbox: {
              type: 'string',
              description: STATS_MAILBOX_PARAM_DESC,
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
              type: ['boolean', 'string'],
              description: lenientBool('true to mark as read, false as unread'),
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
              type: ['boolean', 'string'],
              description: lenientBool('true to pin, false to unpin'),
              default: true,
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'bulk_move',
        description: 'Move multiple emails to a mailbox. ' + membershipReplaceDesc('bulk_add_labels') + ' The destination accepts an id, role, name, or path; an unknown or ambiguous destination is rejected with the valid list. No keyword is changed: a moved message keeps its read/unread and flagged state.',
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
              description: TARGET_MAILBOX_PARAM_DESC,
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
        description: 'Add labels to multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. archive, trash), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.',
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
              description: labelMailboxIdsDesc('add'),
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'bulk_remove_labels',
        description: 'Remove labels from multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. archive, trash), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.',
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
              description: labelMailboxIdsDesc('remove'),
            },
          },
          required: ['emailIds', 'mailboxIds'],
        },
      },
      {
        name: 'check_function_availability',
        description: 'Check which MCP functions are available based on account permissions. Calendar tools run over CalDAV, so calendar is reported available only when CalDAV credentials are configured, regardless of the JMAP calendar capability. Contacts is reported available only when the session has both the JMAP contacts capability and a primary account for it.',
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
              type: ['boolean', 'string'],
              description: lenientBool('If true, only shows what would be done without making changes (default: true)'),
              default: true,
            },
            limit: {
              type: ['number', 'string'],
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
      // Both mailbox tools are thin result wrappers: their orchestration lives in
      // src/mailbox-handler.ts behind an injected client, so it is covered by npm test
      // rather than only by running the server.
      case 'list_mailboxes':
        return { content: await listMailboxes(args, client) };

      case 'create_mailbox':
        return { content: await createMailbox(args, client) };

      case 'list_emails': {
        const { mailbox, limit } = args as any;
        // coerceBool, not !!: a lenient client's stringified "false" is truthy and would
        // silently reverse the sort order. Defaults to false (newest first), matching the
        // documented default. Sits alongside the scope flags below, which coerce the same way.
        const ascending = coerceBool((args as any).ascending) ?? false;
        // raw takes coerceBool like every other flag on this server, not `!!`. A lenient
        // client's stringified "false" is truthy, so `!!` on raw:"false" returned
        // untransformed JMAP to a caller that had explicitly asked for the simplified shape
        // — a silent response-format flip, and raw/verbose sit on nearly every read tool.
        // Both default to false, which is what `!!undefined` already produced, so the
        // default shape is unchanged. An unrecognised value ("1", "yes") now lands on that
        // default instead of flipping the format. Repeated per handler rather than hoisted:
        // each read tool destructures its own args, and the surrounding code differs. (#54)
        const raw = coerceBool((args as any).raw) ?? false;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position);
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.getEmails({
          mailbox,
          limit: validLimit,
          position,
          ascending,
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
        const { emailId, stripQuoted } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        // Same coercion as list_mailboxes for both flags. It matters most here: `!!` on
        // verbose:"false" pulled the HTML body into the response, and `!!` on raw:"false"
        // ALSO made assertStripQuotedNotRaw reject a legitimate stripQuoted read.
        const raw = coerceBool((args as any).raw) ?? false;
        const verbose = coerceBool((args as any).verbose) ?? false;
        // Validated before the fetch so a typo'd field name costs no round trip.
        // Selecting bodyHtml implies verbose's includeHtml: without that, projecting a
        // field the simplifier never emitted would return {} — the trap the parameter
        // exists to avoid (#69).
        const fields = parseEmailFields((args as any).fields, { raw });
        // Rejected before the fetch: raw is unmodified JMAP, so honouring stripQuoted
        // there would be impossible and ignoring it would be silent (#73).
        const strip = coerceBool(stripQuoted) ?? false;
        assertStripQuotedNotRaw(strip, raw);
        const email = await client.getEmailById(emailId);
        const simplified = simplifyEmail(email, { includeHtml: verbose || wantsHtmlBody(fields), stripQuoted: strip });
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
        const text = `Forward draft saved successfully (Email ID: ${result.emailId}). Use send_draft to transmit it. Subject: ${result.subject}${formatInlineNotes(result.notes)}`;
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
        const { limit } = args as any;
        // Same coercion as list_emails - see there for why `!!` was wrong.
        const raw = coerceBool((args as any).raw) ?? false;
        const verbose = coerceBool((args as any).verbose) ?? false;
        const contactsClient = initializeContactsCalendarClient();
        // Hard cap: the contacts tools have no `position` param, so anything past
        // the cap is unreachable. Paging for contacts is tracked as issue #94.
        const result = await contactsClient.getContacts(clampLimit(limit, 20, 100));
        return {
          content: [
            {
              type: 'text',
              text: raw ? formatQueryResult(result) : formatContactQueryResult(result, { verbose }),
            },
          ],
        };
      }

      case 'get_contact': {
        const { contactId } = args as any;
        // Same coercion as list_emails - see there for why `!!` was wrong.
        const raw = coerceBool((args as any).raw) ?? false;
        const verbose = coerceBool((args as any).verbose) ?? false;
        if (!contactId) {
          throw new McpError(ErrorCode.InvalidParams, 'contactId is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        const contact = await contactsClient.getContactById(contactId);
        const output = raw ? contact : simplifyContact(contact, { verbose });
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
        const { query, limit } = args as any;
        // Same coercion as list_emails - see there for why `!!` was wrong.
        const raw = coerceBool((args as any).raw) ?? false;
        const verbose = coerceBool((args as any).verbose) ?? false;
        if (!query) {
          throw new McpError(ErrorCode.InvalidParams, 'query is required');
        }
        const contactsClient = initializeContactsCalendarClient();
        // Hard cap, same as list_contacts — no `position` param, so results past
        // the cap are unreachable. Paging for contacts is tracked as issue #94.
        const result = await contactsClient.searchContacts(query, clampLimit(limit, 20, 100));
        return {
          content: [
            {
              type: 'text',
              text: raw ? formatQueryResult(result) : formatContactQueryResult(result, { verbose }),
            },
          ],
        };
      }

      // The three contacts write tools are thin result wrappers: their coercion and
      // orchestration live in src/contacts-handler.ts behind an injected client, so they are
      // covered by npm test rather than only by running the server against a real account.
      case 'create_contact':
        return { content: await createContactTool(args, initializeContactsCalendarClient()) };

      case 'update_contact':
        return { content: await updateContactTool(args, initializeContactsCalendarClient()) };

      case 'delete_contact':
        return { content: await deleteContactTool(args, initializeContactsCalendarClient()) };

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
        const { calendarId, limit, startDate, endDate } = args as any;
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const events = await davClient.getCalendarEvents(calendarId, clampLimit(limit, 50, 500), startDate, endDate);
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
        const { calendarId, title, description, start, end, location } = args as any;
        // Coerce BEFORE the size guard below, so it measures the real array rather than a
        // lenient client's JSON string (which would slip through as a non-array and be
        // left unmeasured). Per-item shape and key checks live in coerceParticipants.
        const participants = coerceParticipants((args as any).participants);
        // Bound the text before anything measures, escapes or folds it: the iCal line
        // folding these values pass through is quadratic in the field length, so an
        // unbounded description or participant name stalls the whole process. Rejects
        // (never truncates) with the field, its size and the limit. See ical-limits.ts.
        assertICalTextLimits({ title, description, location, participants });
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
        const { eventId, title, description, start, end, location } = args as any;
        // Same coercion, and the same ordering, as create_calendar_event. An omitted
        // participants stays undefined here ("leave the attendees alone"), so the
        // no-field-to-update check and updateCalendarEvent still read it correctly.
        const participants = coerceParticipants((args as any).participants);
        // Same bound the create path applies, for the same reason: the update path folds
        // SUMMARY/DESCRIPTION/LOCATION and every ATTENDEE line through the same quadratic
        // folder, so it is the identical stall from the identical input. See ical-limits.ts.
        assertICalTextLimits({ title, description, location, participants });
        // Same lenient-client reason as edit_draft's clearFields: a stringified array
        // ('["location"]') would fail the Array.isArray test below, so the call would be
        // rejected as "no field to update" while the caller had named one — and if some
        // other field carried the update, the clear would be dropped silently on the way
        // to updateCalendarEvent. Coerce once and pass the coerced value on. (#54)
        const clearFields = coerceStringArray((args as any).clearFields);
        // Coerce for lenient clients, and default to false. updateCalendarEvent tests this
        // with a bare truthiness check, so an unconverted string "false" would read as true
        // — which both skips the confirmation prompt AND authorizes pruning the orphaned
        // recurrence exceptions the prompt exists to warn about. A non-bool like "garbage"
        // yields undefined and so falls to false, never to the destructive direction.
        const confirmRecurring = coerceBool((args as any).confirmRecurring) ?? false;
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
        // Same coercion as list_emails - see there for why `!!` was wrong.
        const raw = coerceBool((args as any).raw) ?? false;
        const verbose = coerceBool((args as any).verbose) ?? false;
        const client = initializeClient();
        const identities = await client.getIdentities();
        const output = raw ? identities : identities.map(i => simplifyIdentity(i, { verbose }));
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
        const { mailbox = 'inbox' } = args as any;
        // clampLimit, not the client's bare Math.min: the schema advertises a string
        // limit, and Math.min("abc", 50) is NaN, which JMAP serializes as
        // `"limit": null` — a query with no bound at all. Same failure the other
        // clamped handlers close; this was the one call site that still had it.
        const limit = clampLimit((args as any).limit, 10, 50);
        // coerceBool, not !!: a lenient client's stringified "false" is truthy and would
        // silently reverse the sort order. Defaults to false (newest first).
        const ascending = coerceBool((args as any).ascending) ?? false;
        // Same coercion for raw — see list_emails for why `!!` was wrong here.
        const raw = coerceBool((args as any).raw) ?? false;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position) ?? 0;
        const client = initializeClient();
        const result = await client.getRecentEmails(limit, mailbox, ascending, position);
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

      case 'archive_email': {
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        const client = initializeClient();
        await client.archiveEmail(emailId);
        return {
          content: [
            {
              type: 'text',
              text: 'Email archived successfully (moved to Archive; read state unchanged)',
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
        const { emailId } = args as any;
        if (!emailId) {
          throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
        }
        // coerceBool, not !!: a lenient client's stringified "false" is truthy and would
        // hand back the raw JMAP shape to a caller who asked for the simplified one.
        const raw = coerceBool((args as any).raw) ?? false;
        const client = initializeClient();
        const result = await client.getEmailAttachments(emailId);
        return { content: buildAttachmentListContent(result, raw) };
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
          // longer prefix "Save path", so a substring match would miss them). Redacted
          // like every other error egress — this branch returns caller-reflected path text
          // and short-circuits the top-level catch, so it has to do its own redaction for
          // the no-exemptions audit to hold.
          if (error instanceof PathAccessError) {
            throw new McpError(ErrorCode.InvalidParams, redactBearerTokens(error.message));
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
        const { query, from, to, cc, bcc, subject, hasAttachment, isUnread, isPinned, mailbox, after, before, limit } = args as any;
        // coerceBool, not !!: a lenient client's stringified "false" is truthy and would
        // silently reverse the sort order. Defaults to false (newest first), matching the
        // three-valued coerceBool the other flags on this call already use.
        const ascending = coerceBool((args as any).ascending) ?? false;
        // Same coercion for raw — see list_emails for why `!!` was wrong here.
        const raw = coerceBool((args as any).raw) ?? false;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw });
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
          ascending,
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
          // its InvalidParams mapping (both tagged branches there redact) — otherwise
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

        // Every calendar tool runs over CalDAV — the JMAP calendar path is disabled
        // and its client methods are gone. So availability is CalDAV configuration
        // alone: reporting available off the JMAP calendar capability would promise
        // tools that then throw "CalDAV not configured" on every call.
        const caldavConfigured = initializeCalDAVClient() !== null;
        const calendarAvailable = caldavConfigured;
        const calendarNote = caldavConfigured
          ? 'Calendar is available via CalDAV'
          : 'Calendar access not available - set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD (a Fastmail app password)';

        // Contacts need BOTH the JMAP contacts capability and a contacts primary
        // account on the session: every contacts method addresses the contacts
        // account and throws when the session reports none, so the capability on
        // its own would promise tools that cannot run.
        const contactsCapability = !!session.capabilities['urn:ietf:params:jmap:contacts'];
        const contactsAccount = !!session.primaryAccounts?.['urn:ietf:params:jmap:contacts'];
        const contactsAvailable = contactsCapability && contactsAccount;

        const availability = {
          email: {
            available: true,
            functions: [
              'list_mailboxes', 'create_mailbox', 'list_emails', 'get_email', 'reply_email', 'forward_email', 'create_draft', 'edit_draft', 'send_draft', 'search_emails',
              'get_recent_emails', 'mark_email_read', 'pin_email', 'delete_email', 'move_email', 'archive_email',
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
            available: contactsAvailable,
            functions: ['list_contacts', 'get_contact', 'search_contacts', 'create_contact', 'update_contact', 'delete_contact'],
            note: contactsAvailable
              // The session reports the contacts capability and an account for it, which is
              // what the READS need. It says nothing about whether the token may WRITE: a
              // read-only contacts scope looks identical here and only refuses at the
              // ContactCard/set, so the write tools are reported available and the caveat is
              // stated rather than implied.
              ? 'Contacts are available. Read access is confirmed by this session; whether create_contact/update_contact/delete_contact can write is not — a read-only contacts token reports exactly the same capability and only refuses when a write is attempted. If a write comes back forbidden, re-issue the API token with read-write contacts access.'
              : contactsCapability
                ? 'Contacts access not available - this session reports the contacts capability but no primary account for it, so there is no account for contacts operations to address'
                : 'Contacts access not available - may require enabling in Fastmail account settings',
            enablementGuide: contactsAvailable ? null : {
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
                '1. Log into Fastmail web interface',
                '2. Go to Settings → Privacy & Security → Connected Apps & API tokens',
                '3. Create an app password with calendar (CalDAV) access',
                '4. Set FASTMAIL_CALDAV_USERNAME (your Fastmail address) and FASTMAIL_CALDAV_PASSWORD (that app password), then restart the server'
              ],
              documentation: 'https://www.fastmail.com/help/technical/servernamesandports.html'
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
        const { limit } = args as any;
        // Coerce dryRun for lenient clients: a stringified "false" is otherwise truthy
        // and would silently keep the diagnostic in dry-run mode. Defaults to dry-run
        // (the safe, non-acting direction).
        const dryRun = coerceBool((args as any).dryRun) ?? true;
        const client = initializeClient();

        // Get some recent emails to test with. clampLimit, not a bare Math.min/max:
        // a non-numeric limit used to reach getRecentEmails as NaN, which JMAP
        // serializes as `"limit": null` — an unbounded metadata dump.
        const testLimit = clampLimit(limit, 3, 10);
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
    // EVERY branch below runs its message through redactBearerTokens, with no exemption.
    // This is the single choke point where error text becomes tool output, so making the
    // rule unconditional is what makes the audit — "no unredacted error text reaches tool
    // output" — a grep anyone can run and verify, rather than a claim resting on an
    // exemption list that has to be re-argued (and kept accurate) per error class. Each
    // exemption costs one function call to remove and one reader-hour to justify, so
    // there are none.
    if (error instanceof McpError) {
      // Redact in place rather than rebuilding: McpError's constructor prefixes the
      // message with "MCP error <code>: ", so re-wrapping an already-wrapped message
      // would double the prefix. Mutating preserves the code, the data, and any McpError
      // subclass identity the SDK constructed.
      // Only `.message` is redacted; `.data` is deliberately left alone. Nothing in this
      // codebase ever populates it, so redacting it would guard a channel that carries
      // nothing — and `.data` is arbitrary JSON, which a string-shaped scrubber cannot
      // walk without either flattening structure or recursing over untyped values.
      // Revisit if anything here starts setting it.
      error.message = redactBearerTokens(error.message);
      throw error;
    }
    // The four compose handlers (reply_email/forward_email/create_draft/edit_draft) have no
    // local try/catch, so an attachment opt-in/path/contentType rejection thrown as a
    // PathAccessError surfaces here. Map it to InvalidParams (actionable) rather than the
    // generic InternalError wrap below. (download_attachment maps its own PathAccessError
    // locally and never reaches here.)
    if (error instanceof PathAccessError) {
      throw new McpError(ErrorCode.InvalidParams, redactBearerTokens(error.message));
    }
    // A semantically-invalid caller input (unresolvable mailbox, label id that is
    // really a name). Map to InvalidParams like PathAccessError above. Placed after the
    // PathAccessError branch and before the generic wrap so an InvalidInputError can't
    // fall through to InternalError.
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