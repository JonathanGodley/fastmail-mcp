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
import { CalDAVCalendarClient, describeCreateCalendarEventResult, describeUpdateCalendarEventResult } from './caldav-client.js';
import { simplifyEmail, setDefaultTimezone } from './email-formatter.js';
import { formatQueryResult, formatRawEmailQueryResult, formatEmailQueryResult, buildExclusionNote, buildCalendarWindowNote, buildCalendarDensityNote, excludedCountPhrase, UNCONFIRMED_COUNT_PHRASE, NOT_EXCLUDED_PHRASE, buildAttachmentListContent, simplifyIdentity, simplifyContact, formatContactQueryResult, formatEditDraftResult, formatSendDraftResult, formatInlineNotes, formatArchiveResult, formatLabelRemoval } from './response-formatters.js';
import { coerceStringArray, coerceStringArrayStrict, coerceBool, coercePosition, clampLimit, redactBearerTokens, redactedJson, toolJson, registerSecret, assertKnownParams, coerceParticipants, PathAccessError, InvalidInputError, resolveUsableTimezone, resolveConfiguredTimezone } from './coerce.js';
import { parseEmailFields, projectEmail, wantsHtmlBody } from './field-projection.js';
import { composeReply } from './reply-handler.js';
import { composeForward } from './forward-handler.js';
import { sendDraftAndMaintainKeywords } from './send-draft-handler.js';
import { composeDraft } from './compose-handler.js';
import { editDraft } from './edit-draft-handler.js';
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
    version: '1.13.4-fork.3',
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

// Opt-in for attaching content that is ALREADY in the account — an `attachments` item
// naming a blobId, or a part of an existing message. Separate from FASTMAIL_ATTACH_DIR
// because it crosses a different boundary: nothing is read off local disk, so the
// local-disk opt-in has nothing to say about it (see docs/security-model.md).
//
// Parsed STRICTLY: only "true" or "1" enable it. A truthy-string test would read
// FASTMAIL_ALLOW_BLOB_ATTACH="false" as ON, turning an operator's explicit refusal into the
// capability they refused. Same parse as the base-URL kill switch — but resolved through the
// four-name fallback the configurable settings use, not that one's deliberate single name,
// so a host that only forwards USER_CONFIG_* spellings can still set it.
//
// Settable from a DXT install too: manifest.json declares `fastmail_allow_blob_attach` as a
// boolean in both user_config and server.mcp_config.env, so a host renders it as a checkbox
// and hands the answer over as "true"/"false". "false" is exactly what the strict parse
// above reads as off, so the unchecked box means what it looks like. The base-URL kill
// switch stays out of the manifest; this is a send capability, not a security control that
// decides where the token may be sent.
function getAllowBlobAttach(): boolean {
  const info = findEnvValue([
    'FASTMAIL_ALLOW_BLOB_ATTACH',
    'USER_CONFIG_FASTMAIL_ALLOW_BLOB_ATTACH',
    'USER_CONFIG_fastmail_allow_blob_attach',
    'fastmail_allow_blob_attach',
  ]);
  return info.value === 'true' || info.value === '1';
}

// Shared `attachments` schema + description for the compose tools. Read getAttachDir() and
// getAllowBlobAttach() at module load (same as the download `path` description reads
// getDownloadDir()), and render an honest disabled clause for whichever source is off —
// attachments have no fallback default the way downloads do, so we must not print a phantom
// root, and a caller must not be sent at a source this server would refuse.
function attachmentsDescription(forEdit: boolean): string {
  const dir = getAttachDir();
  const gate = dir
    ? `path: files must resolve within ${dir} (set via FASTMAIL_ATTACH_DIR); a bare filename or relative path resolves against it, and an absolute path must fall inside it.`
    : `path is disabled until FASTMAIL_ATTACH_DIR is set (restart to enable); each path will then resolve within that directory.`;
  const blobGate = getAllowBlobAttach()
    ? `blobId and emailId+attachmentId are ENABLED (FASTMAIL_ALLOW_BLOB_ATTACH is set to true): they attach content already in the account, so nothing is read off local disk and no size guard applies.`
    : `blobId and emailId+attachmentId are disabled until FASTMAIL_ALLOW_BLOB_ATTACH=true is set (restart to enable).`;
  const base =
    `Things to attach, each an object { path | blobId | emailId+attachmentId, name?, contentType?, cid? }. ` +
    `Each item names EXACTLY ONE source: path (a local file this server reads and uploads), blobId (content already in the account), or emailId + attachmentId together (a part of an existing message). ` +
    `Naming none, naming two, or naming half of the emailId/attachmentId pair is rejected by index; so is a key belonging to a source the item did not choose. ` +
    `A blobId item MUST also give a name — a stored blob carries no filename and none is invented for you. ` +
    `${gate} ${blobGate} ` +
    `contentType is inferred from the file extension when omitted (from the part's own declared type on an emailId+attachmentId item); an explicit contentType is echoed by Fastmail as-is (not re-detected), so a wrong value rides out wrong. ` +
    `Give an item a cid to EMBED it in the message body instead of hanging it off the end: reference it from htmlBody as <img src="cid:THE_CID">. ` +
    `An htmlBody reference with no matching item is rejected, and an item whose cid nothing displays (no htmlBody, or no reference to it) is still attached as an ordinary file and the result says so — a supplied file is never dropped. ` +
    `Size caps (~25 MB/file, ~45 MB total) are a fail-fast guard on LOCAL files only — nothing is read client-side for the other two sources — and Fastmail's own limit ultimately governs.`;
  if (!forEdit) return base;
  return base +
    ` On edit_draft this APPENDS to the draft's existing attachments (they are kept). ` +
    `To remove specific ones, use removeAttachments (pass the blobId from get_email_attachments; a unique attachment name also works). ` +
    `To remove all, use clearFields:['attachments']. Passing attachments together with clearFields:['attachments'] is rejected as a conflict.`;
}

// How an `attachmentId` names a part, shared verbatim by the two places a caller supplies
// one: download_attachment (read a part out) and an attachments item (attach a part to
// outgoing mail). Written once because the resolver is one function — a per-site copy would
// drift from it and from each other, and the entry-number rule differs between the sites in
// exactly one way, which is stated here rather than left to two half-descriptions.
const ATTACHMENT_REF_DESC =
  'Which part to use. Four accepted forms, resolved in this fixed order: (1) a partId from get_email_attachments; (2) a blobId; (3) cid:<value> for an embedded image, using the cid from get_email — the cid: prefix is REQUIRED for this form, only the first one is stripped (cid:cid:x looks up the Content-ID "cid:x"), and a value matching more than one part is rejected rather than guessed at; (4) a plain entry number (0, 1, 2, ...) counting from the start of the get_email_attachments listing. The order is a real precedence, not a single match: every part is checked for a matching partId before any is checked for a matching blobId, and digits therefore resolve as a partId FIRST — Fastmail partIds are themselves digit strings — so the entry-number form applies only when no part claims that value. A number with anything else in it (3a, -1, 1.5) is rejected rather than silently read as an entry number. ' +
  'The entry-number form is READ-ONLY: entry numbers are positional and shift whenever the listing does, so download_attachment accepts one for a one-off read, while attaching a part to outgoing mail rejects a reference that resolved only that way (the rejection is on how it resolved, not on how the string looks) — pass a partId or blobId there, and prefer one anywhere you will reuse the reference.';

// `leadIn` prepends tool-specific context to the shared description (defaulted to ''
// so send/reply/create/edit are untouched) — forward_email uses it to state how NEW
// uploads relate to the original's own carried attachments.
//
// The item schema is FLAT: all seven keys sit in one `properties` map, with the exactly-one-
// source rule stated in prose. Item-level `oneOf` branches would hide the key list from a
// reader for no enforcement gain — the SDK does not validate inputSchema at all, so
// coerceAttachments' source rules are the only real gate either way. (A TOP-LEVEL oneOf on a
// tool's inputSchema would be worse than useless: it leaves `properties` undefined, which
// empties that tool's TOOL_SCHEMAS key set and makes assertKnownParams reject every call.)
// There is no `required` for the same reason: no single key is required on every item.
function attachmentsSchemaProperty(forEdit: boolean, leadIn = '') {
  return {
    type: 'array',
    items: {
      type: 'object',
      properties: {
        path: { type: 'string', description: 'SOURCE file to READ off this machine and upload: an absolute path within the attach directory (FASTMAIL_ATTACH_DIR), or a bare filename/relative path resolved against it. This is the opposite direction from download_attachment\'s `path`, which is a DESTINATION written under FASTMAIL_DOWNLOAD_DIR. One of the three sources; omit it when using blobId or emailId+attachmentId.' },
        blobId: { type: 'string', description: 'Blob already stored in this account to attach, instead of reading a file off disk (requires FASTMAIL_ALLOW_BLOB_ATTACH). Nothing is uploaded — the blob is referenced. A blob carries no filename and no type, so `name` is REQUIRED alongside it and contentType is inferred from that name unless you give one. One of the three sources.' },
        emailId: { type: 'string', description: 'Message whose part you want to attach, used TOGETHER with attachmentId (requires FASTMAIL_ALLOW_BLOB_ATTACH). This attaches one specific part of an existing message — the way to forward a subset of another message\'s attachments. Name and content type default to the part\'s own. One of the three sources; passing it without attachmentId is rejected.' },
        attachmentId: { type: 'string', description: 'Which part of emailId to attach. ' + ATTACHMENT_REF_DESC },
        name: { type: 'string', description: "Filename recipients see. Optional on a path item (defaults to the file's basename) and on an emailId+attachmentId item (defaults to the part's own name); REQUIRED on a blobId item." },
        contentType: { type: 'string', description: 'MIME type like application/pdf (optional). Defaults to the type inferred from the file extension (path), inferred from `name` (blobId), or declared by the part itself (emailId+attachmentId).' },
        cid: { type: 'string', description: 'Content-ID to embed this file under (optional), referenced from htmlBody as <img src="cid:THE_CID">. A simple token of up to 64 letters, digits, dot, dash or underscore, of your choosing — a spelling copied from HTML ("cid:logo") or a header ("<logo>") is accepted and normalised. Each item needs a distinct one. Omit it for an ordinary attachment; identifiers of the form ii-<hex>@inline.invalid are reserved for images this server embeds on your behalf and cannot be authored.' },
      },
    },
    description: leadIn + attachmentsDescription(forEdit),
  };
}

// Shared `participants` schema for the two calendar write tools, so the accepted shapes
// stay identical between create and update. `leadIn` carries the per-tool sentence
// (create adds an ORGANIZER; update REPLACES the whole attendee list).
//
// The shared text has to say that naming an attendee SENDS them mail, because nothing
// else here does. The invitation is emitted by the server's scheduling layer when the
// event is written, not by any tool in this server, so a caller reasoning about which
// tools transmit — send_draft and nothing else — reaches the wrong answer about this one.
// Confirmed live: an event created with one attendee produced an outbound invitation, and
// deleting that event produced the matching cancellation.
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
      ` SENDS MAIL: naming an attendee makes the server email them a meeting invitation from this account as soon as the event is written, and deleting that event later emails them a cancellation. That is the server's scheduling layer, not a compose tool, so it happens with no send_draft call — treat naming a real person here as contacting them. Omit participants to write the event without notifying anyone.` +
      ` Each entry is either an object { email, name? } or a bare email-address string (equivalent to { email }).` +
      ` An entry with any other key, a missing or non-string email, or a non-string name is rejected naming its position, e.g. participants[2] — nothing is silently dropped.` +
      ` The whole array may also be sent as a JSON string ('[{"email":"a@example.com"}]'), for clients that stringify structured parameters; a comma-joined list is NOT accepted.`,
  };
}

/**
 * How `calendarId` is matched, said ONCE for the two tools that take one.
 *
 * The read and the write path share how the parameter RESOLVES — same filtered calendar list,
 * same trim, same fail-closed treatment of an empty value, and literally the same not-found
 * error — though not what they do with a tie: the read path queries every calendar a name
 * matches and the create path takes the first (issue #173). Only the descriptions had
 * diverged: 394 characters on the read side against 41 on the write side ("ID of the calendar
 * to create the event in"), so a caller reading the write tool could not learn that a display
 * name works, that it is matched case-sensitively with surrounding whitespace ignored on both
 * sides, or that a miss is rejected rather than answered emptily.
 *
 * `update_calendar_event` takes no calendarId, so these two are the whole set.
 */
const CALENDAR_ID_MATCHING_DESC =
  'Takes either the calendar\'s `id`/URL from list_calendars or its display name, matched CASE-SENSITIVELY; ' +
  'surrounding whitespace is ignored on both sides, and list_calendars reports the trimmed name. ' +
  'A value matching no calendar is rejected naming ' +
  'the available calendars rather than answered with an empty result, and an empty or whitespace-only string is ' +
  'such a value. Calendars list_calendars does not show cannot be named here either, on a read or on a write.';

function getTimezone(): string | undefined {
  return findEnvValue([
    'FASTMAIL_TIMEZONE',
    'USER_CONFIG_FASTMAIL_TIMEZONE',
    'USER_CONFIG_fastmail_timezone',
    'fastmail_timezone',
  ]).value;
}

// The IANA zone name actually in force for calendar reads, resolved once so the
// list_calendar_events/get_calendar_event descriptions below can say "no timeZone means
// THIS zone" instead of leaving a model to infer it from FASTMAIL_TIMEZONE's env-var name.
// Reads the environment directly rather than through `setDefaultTimezone`'s stored value, so
// it does not depend on that call having run first — TOOLS is built once at module load,
// before any request has reached the handler that calls `setDefaultTimezone`.
const CONFIGURED_TIMEZONE = resolveUsableTimezone(getTimezone());

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

// The signature parameter, shared by the three compose tools so they cannot drift on what
// the flag means (edit_draft's is different in kind — it is a tri-state that preserves what
// the draft already carries — and is written out separately). `context` names where the
// signature lands relative to whatever else that tool puts in the body.
//
// `noBodyReason` exists because the shared no-body clause is FALSE on forward_email: an
// INLINE forward with no note at all is signed, the signature becoming the whole of the
// content above the forwarded-message block, and an asAttachment forward with no note cannot
// reach no-body at all because it substitutes a filler body and signs that. Only a note that
// was supplied and blank yields no-body there, so that tool passes its own wording rather
// than reading a rule that does not apply to it. Every other reason is identical across the
// three, which is why they share this.
function appendSignatureDesc(context: string, noBodyReason?: string): string {
  const noBody = noBodyReason ?? 'the call wrote no body for the sign-off to sit under (a blank body counts as none)';
  return lenientBool(
    `Append the sending identity's configured signature — the sign-off set in Fastmail — to the body${context}. Default false: nothing is appended unless you ask, and nothing is said when you do not. The signature is read from the identity this message sends as (the \`from\` address, or the default identity), so it is always the configured one rather than a remembered copy. An HTML body gets the HTML signature, and its plain-text alternative is derived from it; a plain-text-only message gets the identity's plain-text signature, or — if only an HTML signature is configured — a plain-text form derived from that. WHEN YOU ASK AND THE SIGN-OFF DOES NOT LAND EVERYWHERE IT SHOULD, THE RESULT SAYS SO AND WHY — no signature is available for the address this message sends as (none configured, or not one of your verified identities); ${noBody}; the body you passed already carries a signature; the message ships no HTML and the identity's signature has no plain-text form to write instead (an images-only HTML signature — no HTML ships, so no image does either); or the HTML body WAS signed while the message's plain-text alternative could not be, because the signature's only content is a remote image, which derives no readable text. That last one is reported whether you supplied the plain-text alternative yourself or left it to be derived from the HTML — the recipient reading it sees no sign-off either way. The already-carries case is why asking twice does not sign twice: an HTML body carrying a signature block this server wrote is left alone, and a plain-text body already ending in either form of this identity's block — at the end, or above a quoted or forwarded original — is left alone too, so re-sending a body you read back from a draft is safe, including a body read back from a draft that used to be HTML.`,
  );
}

// Shared scope-control descriptions for the read tools, defined once so the per-flag
// strings and the reliability-contract clause stay in sync between the tools rather than
// drifting as hand-copied strings.
//
// The set is SPLIT by what is true of a given tool, not by tool name: the base constants
// say only what holds for any tool with a single `mailbox` scope parameter (today
// list_emails, search_emails and get_recent_emails), and the SEARCH_-prefixed constants
// append the clauses that exist only where the multi-mailbox scope arrays do — which is
// search_emails alone, because list_emails and get_recent_emails deliberately do not offer
// them. get_recent_emails is what that split was for: it took the Trash/Spam flags later
// (#29) and consumed the base constants unchanged. Do not fold a search-only clause back
// into a base constant, or the base ones start describing parameters their other consumers
// don't have.
const EXCLUDE_DRAFTS_DESC =
  lenientBool('Drafts are included by default; set true to omit them from results (and from the total count). (Note: get_thread differs on BOTH axes — it uses includeDrafts AND excludes drafts by default.)');
// What "excluded" actually means, stated once and carried by both flags so they cannot
// drift apart. JMAP's only exclusion operator is inMailboxOtherThan, which is SOLELY-IN:
// it withholds a message only when every mailbox the message is filed in is excluded.
// Said plainly here because the flag names ("include Trash") imply the opposite reading,
// under which a caller would conclude a cross-filed message must be missing and re-run.
const SOLELY_IN_CAVEAT =
  'Exclusion is "solely in": a message is withheld only when EVERY mailbox it is filed in is excluded, so one filed in both Trash and a normal folder is never withheld — it is already in the results and the withheld-count note does not count it.';
const INCLUDE_TRASH_DESC =
  lenientBool('Trash is excluded by default; set true to also include Trash in the results. ' + SOLELY_IN_CAVEAT);
const INCLUDE_SPAM_DESC =
  lenientBool('The Spam/Junk folder is excluded by default; set true to also include it in the results. ' + SOLELY_IN_CAVEAT);

// `ascending`, declared identically on the three list/search tools (list_emails,
// search_emails, get_recent_emails) so their sort contract can't drift.
const ASCENDING_DESC =
  lenientBool('Sort oldest first instead of newest first (default: false).');

// The reliability contract that makes silence trustworthy. Lead with the no-note
// guarantee as its own sentence (a skimming model must hit "no note => trustworthy"
// first), then the per-signal actions; scoped to the default all-mailbox scope.
//
// Each note is quoted through the phrase constants buildExclusionNote itself emits
// (src/response-formatters.ts), never as hand-copied text: telling a caller to look for a
// string the server never prints reads to it as "no note", which is the one conclusion
// this contract exists to make safe.
// The per-signal actions, shared verbatim; only the lead "what silence promises" sentence
// differs between the tools, because on search_emails the caller can have withheld
// messages itself and silence has to be qualified accordingly.
const SCOPE_NOTE_ACTIONS =
  `A note saying "${excludedCountPhrase('Trash/Spam')}" means re-run (includeTrash:true / includeSpam:true, or mailbox:"trash"/"junk") to see those matches. ` +
  `A note saying "${UNCONFIRMED_COUNT_PHRASE}" or "${NOT_EXCLUDED_PHRASE}" means re-run to be sure. ` +
  'Setting mailbox searches only that folder (no note, by definition).';

const SCOPE_RELIABILITY_CONTRACT =
  'When you search the default scope (no mailbox set): NO note means nothing was withheld — no message filed ONLY in Trash/Spam matched this search, so do not re-run with includeTrash/includeSpam just to re-check the same query. ' +
  SCOPE_NOTE_ACTIONS;

// search_emails additionally has the multi-mailbox scope arrays, and they land on
// opposite sides of the disable rule — which is exactly the thing a caller would guess
// wrong, so it is stated rather than left to symmetry. The lead sentence is also
// qualified rather than inherited: excluding Trash or Spam through excludeMailboxes takes
// that role out of the default set, so silence there means "nothing was withheld BEYOND
// what you excluded" — the unqualified promise would be literally false in that state.
const SEARCH_SCOPE_RELIABILITY_CONTRACT =
  'When you search the default scope (no mailbox and no requiredMailboxes set): NO note means nothing was withheld beyond what you excluded yourself — no message filed ONLY in Trash/Spam matched this search, other than any Trash/Spam you named in excludeMailboxes, so do not re-run with includeTrash/includeSpam just to re-check the same query. ' +
  SCOPE_NOTE_ACTIONS +
  ' requiredMailboxes is an explicit scope too and turns the default exclusion and its note off the same way; excludeMailboxes does NOT — the default exclusion and its note stay on, minus any Trash/Spam role you excluded yourself.';

// The forms every mailbox-taking parameter accepts, written ONCE and shared by all of
// them — the scalar ones (mailbox / targetMailbox / parent) and the mailboxIds arrays
// alike. They resolve through a single matcher, so a form documented on one tool and not
// another would be a documentation-only difference, and the path form is specified here
// and nowhere else in the schemas.
const MAILBOX_REF_FORMS =
  'Accepts an id, a role (inbox, archive, sent, drafts, trash, junk), a folder name (e.g. Receipts), or a root-anchored path (e.g. Archive/2026/Receipts): "/"-separated, no leading or trailing slash, segments matched case-insensitively. ' +
  'A folder name matching exactly one mailbox wins over reading the same text as a path, so a folder whose own name contains "/" stays reachable by that name — unless the same text ALSO reaches a different mailbox as a path (a folder named "A/B" alongside a real A > B nesting), which is rejected as ambiguous and answered with the id of each, since no path can tell those two apart. ' +
  'A name shared by several mailboxes is rejected as ambiguous, listing their full paths — retry with one of those, or with the id. An unknown mailbox is rejected with the valid list. ' +
  'list_mailboxes returns each mailbox\'s path, and a path it returns can be pasted straight back into this parameter.';

// True of every read tool that scopes by a single mailbox (list_emails, search_emails,
// get_recent_emails). Anything specific to the multi-mailbox arrays belongs in
// SEARCH_MAILBOX_PARAM_DESC below, not here — the other two do not have them.
const MAILBOX_PARAM_DESC =
  'Mailbox to scope to. ' + MAILBOX_REF_FORMS +
  ' Setting it searches exactly that mailbox (incl. Trash/Spam) and ignores the default Trash/Spam exclusion.';

// search_emails' `mailbox`, which is the one-element case of requiredMailboxes.
const SEARCH_MAILBOX_PARAM_DESC =
  MAILBOX_PARAM_DESC +
  ' It is exactly the single-mailbox shorthand for requiredMailboxes: mailbox:"X" and requiredMailboxes:["X"] are the same query, with no hidden difference. Passing both is allowed — they fold into one intersection (the message must be in all of them).';

// The two multi-mailbox scope arrays (#26). Both take every reference form the scalar
// `mailbox` takes, resolved by the same exact matcher as the label arrays and with the
// same all-or-nothing rejection, so a path or role learned on one works on the others and
// a bad array is fixed in one retry.
const SCOPE_ARRAY_REJECT_DESC =
  ' Any entry that fails to resolve rejects the whole call (no widened search runs), and the error names every failing entry at once. Every entry must be a string; a non-string entry is rejected by index.';

const REQUIRED_MAILBOXES_PARAM_DESC =
  'Mailboxes the message must be filed in — ALL of them (they intersect: each entry is a separate condition, so a message in only some of them does not match). ' +
  'Use it to find mail carrying several labels at once. Each entry: ' + MAILBOX_REF_FORMS + SCOPE_ARRAY_REJECT_DESC +
  ' Setting this is an explicit scope, so like mailbox it turns OFF the default Trash/Spam exclusion and its note — the folders you named are the scope. mailbox is its single-mailbox shorthand, and passing both intersects them.';

const EXCLUDE_MAILBOXES_PARAM_DESC =
  'Mailboxes to exclude — but read this first: it maps to JMAP\'s inMailboxOtherThan, which is SOLELY-IN. A message is hidden only when EVERY mailbox it is filed in is in this list, so a message in Inbox plus an excluded label is still returned. ' +
  'Nothing here offers strict exclusion: JMAP has no filter for "not in this mailbox", so absolute exclusion means filtering the returned results yourself. ' +
  'Each entry: ' + MAILBOX_REF_FORMS + SCOPE_ARRAY_REJECT_DESC +
  ' These entries are UNIONED with the default Trash/Spam exclusion into one excluded set rather than applied as a second, separate condition. That is deliberate and strictly stronger: a message filed in {Trash, an excluded label} is hidden by the union, though neither exclusion on its own would hide it. ' +
  'Excluding mailboxes does NOT disable the default Trash/Spam exclusion or its note (unlike mailbox and requiredMailboxes) — except that naming Trash or Spam here takes that role out of the default set, so the note stops prescribing includeTrash/includeSpam, which could not override your own exclusion. ' +
  'May be combined with requiredMailboxes: the query then reads "in every required mailbox, and in at least one mailbox outside the excluded set".';

// The mailbox a draft is filed into, shared by nothing else: create_draft is the only tool
// that picks a save destination without moving anything.
const DRAFT_MAILBOX_PARAM_DESC =
  'Mailbox to SAVE the draft into (optional, defaults to Drafts). Does not set From or recipients. ' + MAILBOX_REF_FORMS;

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
  ' Any entry that fails to resolve rejects the whole call, and the error names every failing entry at once. So does any entry that resolves to a FOLDER rather than a label (see the tool description): the check runs after resolution, so naming one by name or path is rejected exactly as naming it by role is.';

// Shared by all four label tools (#133). States the namespace the tools operate in, which a
// caller cannot infer from "label" alone: Fastmail's own label picker offers the Inbox and
// the account's user labels and nothing else, while every other role mailbox appears only
// under "Move to".
const LABEL_NAMESPACE_DESC =
  ' Labels here means the Inbox and the account\'s own user labels ONLY. A mailbox with any other JMAP role (archive, trash, junk/Spam, drafts, sent, snoozed, scheduled) is a FOLDER in Fastmail\'s model, not a label — Fastmail\'s label picker does not offer it — so naming one rejects the whole call before anything is written; use move_email or bulk_move to put a message in a folder. The Inbox is the one mailbox in both namespaces: removing the inbox label is exactly what archiving a message is, and adding it is how a message is put back in the Inbox.';

// Shared by remove_labels and bulk_remove_labels. A message must be filed somewhere, so a
// removal that would take away its last mailbox needs an answer; this states the one the
// Fastmail client gives, since a caller cannot otherwise predict where the message lands.
const LABEL_REMOVAL_RESCUE_DESC =
  ' If removing these labels would take away the LAST mailbox holding the message, the archive-role mailbox is added in the same write (found by ROLE — a folder merely NAMED "Archive" is not it), so removing a message\'s only label archives it rather than deleting it. One case is rejected instead of served: the account has no archive-role mailbox at all, so there is no fallback to reach for. It says so and points at move_email/bulk_move or delete_email/bulk_delete. (Removing Archive itself never reaches that question — Archive is a folder, so the namespace rule above rejects it whatever the message is filed under.)' +
  ' Naming a label the message does not carry changes nothing for that message.' +
  ' Every rejection here, and a message whose current filing the server does not report, aborts the WHOLE call before anything is written — the message says so. Per-message server failures are reported per message as usual.' +
  ' Surviving mailboxes are re-asserted in the same write, which is what stops the removal emptying the message; one consequence is that a message also in Scheduled may come back as a failure, because the server appears to reject re-asserting a scheduled membership outside a send request (see issue #130). That combination has not been measured.';

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
//
// The delete tool is a PARAMETER for the same reason membershipReplaceDesc's additive tool
// is: this text is attached to a bulk tool as well as a single-message one, and a bulk
// caller sent to the single-email `delete_email` would find it rejects their `emailIds`.
// `archive_email` needs no such treatment — it takes an array and serves both callers.
const targetMailboxParamDesc = (deleteTool: 'delete_email' | 'bulk_delete') =>
  'Destination mailbox. ' + MAILBOX_REF_FORMS +
  ' A role name resolves to the mailbox carrying that ROLE whenever the account has one, so a user-created folder of the same name does not capture it. That folder is then reachable by its id, or by its full path if it is nested — a TOP-LEVEL folder sharing a role name has a path identical to its name, and the path form is only tried for an input containing a separator, so its id is the only way to reach it. But the role branch is tried FIRST, not exclusively — on an account with no mailbox carrying the role, the name branch runs and a folder of that name is what you get. ' +
  `The dedicated verbs do not have that fall-through: archive_email and ${deleteTool} resolve by role and nothing else, so they never fall back to a folder of that name. ` +
  'For ARCHIVE there is a second reason: moving to the archive role replaces the whole membership and drops every other label, which is not what archiving means — archive_email patches the Inbox membership away and keeps the rest. Deleting has no such difference: delete_email and bulk_delete replace the whole membership too, exactly as moving to trash does, because that is what Fastmail\'s own Delete does.';

// The membership warning carried by every tool that sets mailboxIds whole-value
// (move_email, bulk_move). archive_email is deliberately NOT one of them — it patches the
// Inbox membership away and re-asserts the rest, so it never drops other filing. Written
// once because the consequence is the
// same on each: a message filed under several labels keeps only the destination, and the
// additive alternative is the one a caller usually wants. The additive tool is a
// parameter rather than a fixed word — a bulk caller sent to the single-email
// `add_labels` would find it rejects their `emailIds`, which is a worse outcome than no
// pointer at all.
const membershipReplaceDesc = (additiveTool: 'add_labels' | 'bulk_add_labels') =>
  'This REPLACES the message\'s entire mailbox membership: every other label/folder it was filed under is removed. ' +
  `To file it somewhere while KEEPING its existing labels, use ${additiveTool} instead — with one limit: the label tools take the Inbox and the account's own labels only, so a folder (any other role mailbox: Archive, Trash, Spam, Drafts, Sent, Snoozed, Scheduled) is reachable only by moving. Archiving is the exception, and archive_email is the tool for it: it drops the Inbox membership and keeps every other label.`;

// What the three contacts READ tools return, written once. All three had a hand-copied
// duplicate of this sentence and of the `verbose` parameter text below, which is how the
// wrong field name ("org" for `organization`) survived in all three at once.
const CONTACT_SHAPE_DESC =
  'Returns simplified format by default: id, name, emails, phones, organization, notes, and kind. ' +
  'In the DEFAULT view kind appears only when the card is NOT an ordinary person, so no kind ' +
  'field there means an individual — which is nearly every card. kind:"group" is a contact ' +
  'group, which update_contact and delete_contact both REFUSE (this server has no kind or ' +
  'members parameter, so it cannot make one or put one back); other values (org, location, ' +
  'device, application, …) are not people either. Check it before planning an edit or a cleanup ' +
  'rather than discovering it from a refused write. ' +
  'Each emails/phones entry is EITHER a bare string (the address / the number, when the entry ' +
  'carries no label) OR an {address, label} / {number, label} object — so handle both shapes. ' +
  'Use verbose=true for the whole entry objects (contexts, pref, …), addresses, titles, URLs, ' +
  'photos and anniversaries, and kind on EVERY card (individual included, so the "no field means ' +
  'individual" reading above applies to the default view only). Use raw=true for the original ' +
  'JMAP response, which also carries kind on every card.';

// The `verbose` parameter text shared by the same three tools.
const CONTACT_VERBOSE_PARAM_DESC =
  'Return each emails/phones entry whole (contexts, pref and any other stored field) instead ' +
  'of the bare-string-or-{value,label} shape, and include the extra contact fields ' +
  '(addresses, titles, URLs, photos, anniversaries). Also returns kind on EVERY card, ' +
  'including the individual the default view omits — so under verbose the presence of kind ' +
  'says nothing about what the card is; read its value. Not needed for most tasks.';

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
        description: 'List recent emails across all mailboxes (or one, via mailbox). Trash and Spam are excluded by default (set includeTrash/includeSpam to include them); drafts are included (set excludeDrafts to omit them). Set mailbox to scope to a single mailbox (incl. Trash/Spam), which ignores the default exclusion. ' + SCOPE_RELIABILITY_CONTRACT + ' One mailbox is all this tool scopes to: to intersect several mailboxes or exclude some, use search_emails (requiredMailboxes / excludeMailboxes) with no query. get_recent_emails filters identically and differs only in size — it returns 10 by default and caps at 50, for a quick look; use this one to work through a folder. Returns simplified format (metadata + preview, no bodies). Use raw=true for original JMAP response. For email bodies, use get_email. The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
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
        description: 'Reply to an existing email with proper threading headers (In-Reply-To, References). Automatically fetches the original email to build the reply chain. The subject defaults to \'Re: <original subject>\'; pass subject to override it (see that parameter for what a changed subject does to draft grouping). Use this rather than hand-rolling threading headers on create_draft. Set appendSignature=true to have the sending identity\'s configured signature added above the quote (off by default). This tool always saves the reply as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the reply, it marks the original answered and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft), so when several copies of the original exist the right one is marked. The original message is quoted by default (attributed, top-posted, matching the web client with a portable quote-bar); set quoteOriginal=false to omit it. Quoted HTML is reproduced sanitised (script/style/event handlers stripped; formatting and real http(s) images kept) and is re-sent under your From address. IMAGES THE ORIGINAL DISPLAYED ARE CARRIED INTO THE QUOTE by default, and are re-sent to this reply\'s recipients — so a reply can send image data outward that you never attached. ' + CARRIED_IMAGE_BOUND_DESC + ' The only way to send none of it is quoteOriginal=false, which omits the whole quote; there is no setting for quote text without its images. Carrying needs no FASTMAIL_ATTACH_DIR: the parts are already in the account. A quoted image is only carried when the reply ships an HTML body — a text-only reply drops them, and the result says how many. You can also embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; that one needs a source you have enabled — FASTMAIL_ATTACH_DIR to attach a local file, FASTMAIL_ALLOW_BLOB_ATTACH to attach content already in the account).',
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
            appendSignature: {
              type: ['boolean', 'string'],
              description: appendSignatureDesc(', above the quoted original'),
            },
            attachments: attachmentsSchemaProperty(false),
          },
          required: ['originalEmailId'],
        },
      },
      {
        name: 'forward_email',
        description: 'Forward an existing email to new recipients. This tool always saves the forward as a DRAFT and never transmits it — review it, then transmit with send_draft (the only tool that sends mail). When send_draft transmits the forward, it marks the original forwarded and read — exactly the stored copy this call was given as originalEmailId (recorded on the draft, on both inline and asAttachment forwards). `to` is required — a forward has no default recipient, unlike reply. A note (textBody/htmlBody) is optional: the forwarded message itself is the content. The original is reproduced below a forwarded-message header block (From/To/Cc/Subject/Date), its HTML sanitised (script/style/event handlers stripped; formatting and real http(s) images kept) and re-sent under your From address. The original\'s regular attachments are carried by default (includeOriginalAttachments). Images the original\'s body DISPLAYED are carried too, and are carried even when includeOriginalAttachments is false, because they are body content rather than attached files — a forward without them would reproduce a message with holes in it. ' + CARRIED_IMAGE_BOUND_DESC + ' Short of not forwarding the message, there is no way to reproduce the body and leave those images behind. Carrying needs no FASTMAIL_ATTACH_DIR: the parts are already in the account. An image the block cannot display — a text-only forward, or a reference this server could not resolve to exactly one image part — rides as a regular attachment instead (subject to includeOriginalAttachments), and the result says so; asAttachment is the lossless alternative. The subject defaults to \'Fwd: <original subject>\'. Set appendSignature=true to have the sending identity\'s configured signature added above the forwarded-message block (off by default). You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; that one needs a source you have enabled — FASTMAIL_ATTACH_DIR to attach a local file, FASTMAIL_ALLOW_BLOB_ATTACH to attach content already in the account).',
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
              description: lenientBool("Carry the original message's attached FILES on the forward (default true; all-or-none — for a subset, either trim the saved draft with edit_draft's removeAttachments, or set this false and name the parts you want as attachments items (emailId + attachmentId), which sends them with no forwarded-message block or X-Forwarded-Message-Id to say where they came from). This flag does not govern the images the original's body displayed: those are body content and are carried either way (see the tool description for the bound on that). What it does govern, besides the ordinary files, is an image the forwarded block could not display — that one rides as a regular attachment when this is true, and is left behind when it is false. Ignored when asAttachment is set — the .eml already embeds every original attachment."),
            },
            asAttachment: {
              type: ['boolean', 'string'],
              description: lenientBool('Instead of reproducing the original inline, attach the entire original as a raw .eml file (message/rfc822): lossless, including embedded inline images; supersedes includeOriginalAttachments. There is no forwarded-message block on this shape — your note (textBody/htmlBody) is the WHOLE body. IF YOU PASS NO NOTE THE DRAFT IS NOT BODY-LESS: it ships the literal plain-text body "Forwarded message attached." so the message reads as something rather than as an empty page with a file on it. Pass your own note to replace it. NOTE: the raw message carries its full transport headers (Received chain, authentication results) and — when forwarding a message from Sent — any Bcc recipients (see docs/security-model.md), which an inline forward would not expose.'),
            },
            replyTo: {
              type: 'array',
              items: { type: 'string' },
              description: 'Reply-To email addresses (replies go here instead of to the sender). Each entry may be "Name <email>" or a bare address.',
            },
            appendSignature: {
              type: ['boolean', 'string'],
              description: appendSignatureDesc(
                ', above the forwarded-message block (on an asAttachment forward there is no such block — the note is the whole body, and the signature goes BELOW it)',
                'the note you passed was blank — note that an INLINE forward with NO note at all IS signed, the signature becoming the whole of the content above the forwarded-message block, and that an asAttachment forward cannot report this at all: with no note it writes the filler body "Forwarded message attached." and signs that',
              ),
            },
            attachments: attachmentsSchemaProperty(false, "NEW attachments to add (the original's own attachments are carried automatically — see includeOriginalAttachments). "),
          },
          required: ['originalEmailId', 'to'],
        },
      },
      {
        name: 'create_draft',
        description: 'Create an email draft without sending it (transmit it later with send_draft, the only tool that sends mail). Supports threading headers for replies, but for a reply to an existing message prefer reply_email — hand-rolled inReplyTo/references under a mismatched subject permanently detach the draft from its conversation in the Fastmail UI (see the inReplyTo parameter). IMPORTANT: each call creates a new draft — do not call twice for the same message. Set appendSignature=true to have the sending identity\'s configured signature added to the body (off by default). You can embed your own image in the body: give an attachments item a cid and reference it from htmlBody as <img src=\"cid:THE_CID\"> (see the attachments parameter; requires FASTMAIL_ATTACH_DIR for a local file, or FASTMAIL_ALLOW_BLOB_ATTACH for content already in the account).',
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
            appendSignature: {
              type: ['boolean', 'string'],
              description: appendSignatureDesc(''),
            },
            attachments: attachmentsSchemaProperty(false),
          },
        },
      },
      {
        name: 'edit_draft',
        description: 'Edit an existing draft email. Only fields you provide are changed; omit a field to leave it unchanged. Setting a field to an empty value is rejected: to deliberately clear a field, name it in `clearFields`. A cleared draft is still valid (it just may not be sendable, e.g. with no recipients). The plain-text body is an auto-managed fallback of the HTML: editing htmlBody alone regenerates textBody from the new HTML (an html-alone edit discards any custom textBody the draft had); editing textBody alone while htmlBody is present is rejected (it would not change what recipients render); clearFields:[\'textBody\'] while htmlBody is present is rejected (the fallback is auto-managed); clearFields:[\'htmlBody\'] converts the draft to plain text. An edit that would leave the draft with no body is rejected. Editing the body of a reply draft that still carries the quoted original — or of a forward draft that carries the forwarded-message block — requires you to say what happens to it: pass originalEmailId (the id of the message this draft replies to or forwards) to rebuild the body and keep it, or noQuote:true to drop it. Metadata-only edits (subject/recipients/attachments) and plain-text conversion (clearFields:[\'htmlBody\']) keep the quote automatically; each successive body edit that should keep the quote must pass originalEmailId again. Supplying htmlBody to a text-only reply draft converts it to HTML. A draft carrying the sending identity\'s signature keeps it: an edit that writes a body without one re-appends the identity\'s current signature and says so, unless you pass appendSignature:false to drop it (see that parameter). The display name written alongside the From address prefers the name the draft already carries against that address; the identity\'s configured name is only used as a fallback when the draft carries none, so a display name you set yourself is never silently reverted to your account\'s name by an edit that never touched `from` — including a metadata-only edit. A draft whose From matches no identity you can send as keeps its own address and its own display name unchanged. Since JMAP emails are immutable, this creates a replacement draft and moves the old one to Trash (so the returned email ID is new); the edit preserves the draft\'s threading headers (In-Reply-To/References), attachments, and other keywords. The replaced draft is never destroyed: it stays recoverable in Trash until Trash is emptied or auto-purged (Trash retention is a per-account setting), so an edit made from an out-of-date copy of the draft can be undone. The result also reports what the replaced draft contained (subject, recipients, body sizes) — compare it against what you expected to replace, since a draft changed elsewhere (the web UI, another client) in the meantime will have been overwritten. On the rare failure where the replacement is created but the old copy can\'t be moved to Trash, you are left with a duplicate draft rather than none, and the result says so. Drafts with embedded (cid:) images can be edited: an image the edited body still displays keeps its identifier, an image the body no longer displays is taken off the draft if this server put it there and becomes a regular attachment if it came from elsewhere, and the result says what the draft ended up embedding. Two body shapes still can\'t be rebuilt faithfully and are rejected — a body part that is neither text nor a carriable image, audio, video or attached message, and a body that interleaves two parts of the same text type — recreate those drafts instead. If the draft\'s stored body already references an image that isn\'t attached, editing its body is rejected until the edit resolves that (replace the body, or add an attachments item supplying the missing cid); metadata and attachment edits still work on such a draft.',
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
              description: 'Updated sender email address (optional). Switching to a different address you can send as writes that identity\'s display name, since the draft carries no name for the new address yet.',
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
              description: "When editing the body of a REPLY or FORWARD draft, the id of the message this draft replies to OR forwards (NOT this draft's own id, which is emailId). Pass it to rebuild the body and keep the quoted original / forwarded-message block. Without it, a body edit that would drop them is rejected. Rebuilding APPENDS the quote to the body you supply: if the body you are sending already contains the quoted original — which it does whenever you read the draft, edit the words, and send the whole text back — pass noQuote:true instead; such a call is rejected rather than storing the quote twice.",
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
            appendSignature: {
              type: ['boolean', 'string'],
              description: lenientBool("What happens to the sending identity's configured signature when this edit writes a body. OMIT IT and the signature is preserved: if the draft already carries a signature block this server wrote and the body you supply does not, the identity's CURRENT signature is re-appended (above any quoted or forwarded original) and the result says so — rewriting a body is not read as a request to drop a sign-off that was deliberately added. Pass true to add the signature to a draft that never had one. Pass false to remove it: your new body is stored exactly as written, with nothing appended. An edit that writes no body leaves both bodies untouched, so nothing is appended there either; pass true on such an edit and the result says that plainly rather than ignoring the flag in silence. LIMIT: preservation is detected by an HTML class, so it works on drafts with an HTML body ONLY. A PLAIN-TEXT draft carries no marker, so an edit that rewrites its text loses the sign-off even if this server put it there — pass appendSignature:true on that edit to keep it, which is safe to repeat because a text body already carrying EITHER form of this identity's block (the HTML-derived one or the configured plain-text one) is left alone rather than signed twice — whether the block ends the body or sits above a quoted or forwarded original, which is the shape a reply or forward draft you read back will actually have. That covers converting an HTML draft to plain text with clearFields:['htmlBody'] and handing back the text the draft gave you. A signature typed by hand is invisible to this everywhere. WHEN THE SIGN-OFF DOES NOT LAND, THE RESULT SAYS SO AND WHY, and which cases are reported depends on who asked. WITH true, all four: no signature is available for the address this edit sends as (none configured, or not one of your verified identities); the body you supplied already carries one; this edit wrote no body; or the identity's signature is images-only HTML and this edit leaves the draft as plain text (or supplies a plain-text alternative) that has nothing to carry it. WITH THE FLAG OMITTED only the two that LOSE something are reported — no signature available, and the images-only case — because the other two mean the sign-off is still there: a body you supplied that already carries it is what preservation wanted, and an edit that writes no body carries both bodies through untouched. So on the omitted path, silence does not mean all four checks passed."),
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
        // One search tool, not two. The structured filters below (from/to/cc/bcc/subject,
        // hasAttachment/isUnread/isPinned, the date range, and the mailbox scoping params)
        // are the whole of what upstream exposes as a separate `advanced_search`; this
        // server folds them into `search_emails` and ships no second tool. Two tools whose
        // only difference is which filters they accept forces the caller to pick before
        // knowing what it needs, and a filter added to one silently makes the other the
        // weaker choice — the same split-vocabulary failure that made a free-text-only
        // search and an id-only mailbox parameter worth consolidating (#12). Anything new
        // here is a parameter on this tool.
        description: 'Search emails. Provide a free-text query matched across subject, body, and participants (plain words — NOT operator syntax: "from:alice" is matched literally; for structured matching use this tool\'s own from/to/cc/bcc/subject params). All filters combine with AND. Trash and Spam are excluded by default (deleted mail lives in Trash; set includeTrash/includeSpam to include them); drafts are included. Set mailbox (incl. Trash/Spam) to search exactly that mailbox, which ignores the default exclusion; requiredMailboxes/excludeMailboxes scope across several mailboxes at once. ' + SEARCH_SCOPE_RELIABILITY_CONTRACT + ` Recovery example: if a search returns a "2 ${excludedCountPhrase('Trash/Spam')}" note, re-run with includeTrash:true (or mailbox:"trash") to find the deleted message.` + ' Returns simplified format (metadata + preview, no bodies); use raw=true for original JMAP, get_email for bodies. The date field is local time with a UTC offset (raw=true returns canonical JMAP UTC). ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' query is optional: search_emails with no query returns recent mail matching only the structural filters (for a plain folder listing use list_emails). limit default 20, max 100. ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
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
              description: SEARCH_MAILBOX_PARAM_DESC,
            },
            requiredMailboxes: {
              type: 'array',
              items: { type: 'string' },
              description: REQUIRED_MAILBOXES_PARAM_DESC,
            },
            excludeMailboxes: {
              type: 'array',
              items: { type: 'string' },
              description: EXCLUDE_MAILBOXES_PARAM_DESC,
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
          ' This is IRREVERSIBLE: unlike an email, a deleted contact does not go to Trash, so the echoed card is the only copy left — keep it if there is any chance the delete was wrong. create_contact can rebuild the name, emails, phones, addresses and note from it, but NOT the rest of the card (photos, titles, organizations, nicknames, URLs, anniversaries, group membership, the uid, or the per-entry contexts/pref), so recreating gives you a similar contact rather than the one that was deleted. A contact GROUP is REFUSED: create_contact has no kind or members parameter, so this server cannot make a group and will not destroy one it could never put back — delete a group in the Fastmail web interface. (That refusal is about the kind of record, not about fields create_contact cannot set: an ordinary card carrying titles or organizations still deletes.) There is deliberately no confirmation parameter.',
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
        description:
          'List events from a calendar, one entry per occurrence in the requested window. ' +
          'RECURRING EVENTS ARE EXPANDED: give startDate/endDate and a repeating event returns one entry for each occurrence that falls inside the window, with start/end set to that occurrence\'s real dates — so a fortnightly event across three months is several entries, not one. Reading four recurrence fields tells you exactly what a date is: `recurrenceId` present means start/end ARE the in-window occurrence; `recurrenceRule` present (the RRULE) means you are looking at the series master shown at its ORIGINAL start date, which may be years before the window; `recurrenceDates` present (the raw RDATE values, comma-separated) means the same thing for a series that LISTS its occurrences instead of stating a rule, so it too is a master at its original date; `isRecurring` says the entry belongs to a repeating series. `recurrenceDates` CARRIES THE VALUES AND NOT THE PARAMETERS: a TZID, VALUE=DATE or VALUE=PERIOD on the RDATE line is dropped, so a designator-less value in that list does NOT follow the `timeZone` rule that governs `start`, and RDATEs written in different zones are indistinguishable once joined. Read it as proof that other dates exist, never as a date you can place on a clock — pass a window and let the server expand the series if you need real occurrence times. A SERIES THAT ONLY LISTS ITS DATES IS MATCHED BY THE WINDOW OVER ITS OWN START ALONE: where a series states no RRULE and lists its occurrences as RDATEs, Fastmail indexes it across `DTSTART` only, so a window covering one of those listed dates but NOT the series start returns nothing whatsoever for it — the row is ABSENT rather than mis-dated, nothing in the response marks the omission, and an empty result is not proof of a free day. Confirm such a day in the Fastmail web interface, or widen the window until it reaches the series start (fork issue #167). Master and occurrence are the usual split, but the two fields CAN appear together — RFC 5545 allows an override block to carry its own rule — and where they do, `recurrenceId` names the instance this entry is and `recurrenceRule` is that block\'s own rule, not the series\'. In an expanded window the FIRST occurrence of a series arrives with no recurrenceId of its own (the server marks only instances after the first), so it is recognisable as recurring only when a sibling occurrence shares the window; where the window holds that first occurrence and nothing else, the entry is indistinguishable from a one-off — call get_calendar_event on its id to see the series master and its RRULE. ' +
          'EVERY OCCURRENCE OF A SERIES CARRIES THE SAME `id`, and that id names the SERIES, not the occurrence: seven rows of a fortnightly event are seven identical ids. update_calendar_event and delete_calendar_event act on ALL occurrences of the id you give them — deleting the row for one Thursday deletes every Thursday, past and future, and mails every attendee a cancellation. There is no per-occurrence id here and no way to change or remove a single instance through this server. A row\'s `url` is the SAME record under another name and is equally series-wide: delete_calendar_event accepts it in place of `id` and destroys just as much, so it is not a safer handle. ' +
          'ROWS CARRY CORE FIELDS ONLY — never `participants` and never `organizer`, even on an event that has them, because this tool does not fetch them. An absent participant list here does NOT mean the event has no attendees, and empty fields are omitted throughout, so there is nothing to distinguish "none" from "not asked for". Call get_calendar_event on the id before any destructive call: it is the only way to see who is about to be mailed a cancellation. ' +
          'WITHOUT startDate/endDate THE NEXT MONTH IS LISTED: the window runs from the start of today in the configured timezone for 31 days, expanded like any other window, and the response says so in a trailing "Note:" line naming the range actually searched. There is no unwindowed listing — an absent window is an OPEN-ENDED one, and expanding recurrences across it would materialise every occurrence of every repeating event. Pass a window to ask about other days; call get_calendar_event on an id to see the unexpanded series master. ' +
          'A DATE IS A LOCAL DAY: startDate/endDate written as plain dates (2026-08-12) cover that calendar day in the account\'s configured timezone (FASTMAIL_TIMEZONE, falling back to this server\'s own zone — the same zone every email `date` is shown in), NOT the UTC day. A datetime with no Z and no offset is read as local time too. Only a value carrying Z or a numeric offset means the exact instant it names, so that is how to ask for something the local-day rule cannot express. ' +
          'ONE-SIDED WINDOWS ARE BOUNDED: pass only startDate (or only endDate) and the missing half is filled in 31 days away — a month — because expanding recurrences over an open-ended range would materialise every occurrence of every repeating event. The response says so in a trailing "Note:" line naming the range actually searched; pass both bounds to choose the span yourself. ' +
          'A SINGLE REPEATING EVENT TOO DENSE TO MATERIALISE IS LEFT OUT, whatever window you named: a series that expands to more than 5000 blocks in the RANGE THIS SERVER SEARCHED for your window is omitted from the rows AND from the total, and named in a trailing "Note:" line with its title, id, occurrence count and calendar. That range runs up to 14 hours past each edge of the window you named (see the widening described below, under which rows may sit outside the window), so the count reported is over the returned blocks and can exceed the number of occurrences falling strictly inside your window. The span is yours to choose, but the density is not — one FREQ=MINUTELY invitation would otherwise fill the whole page and push every real event off it. The call still answers everything else normally; narrow the window to see the omitted series. ' +
          'Every calendar the account listed is queried before the results are sorted and trimmed, so `limit` is a genuine "earliest N" across all of them (a collection the server failed to list is not in that set — see the discovery clause below). The response opens with a summary line stating how many events matched in total; when that total exceeds the returned count, `limit` cut the rest off and there is no paging, so raise `limit` (up to 500) to see more, or narrow the window if the total is larger than that. ' +
          `CALENDAR TIMES CARRY A ZONE NAME, NEVER AN OFFSET. \`start\`/\`end\` is a bare local wall clock (2026-04-20T10:00:00), a Z-designated UTC instant, or a date-only (all-day) value — this server never puts an offset in either and never asks you to compute one. READ THE VALUE'S OWN DESIGNATOR FIRST: \`timeZone\` only QUALIFIES a value that carries neither Z nor a date-only marker, so "absent means the configured zone" applies to a bare wall-clock \`start\` and nothing else. \`timeZone\` names the IANA zone a wall-clock \`start\` is in, but ONLY when it differs from this server's configured zone (${CONFIGURED_TIMEZONE}): an ABSENT \`timeZone\` means ${CONFIGURED_TIMEZONE}, and \`timeZone: null\` means \`start\` is genuinely FLOATING (RFC 5545 §3.3.5 — no TZID, no Z, a different instant for every reader), which is a different fact from "in the configured zone". A Z-designated value or an all-day value never carries \`timeZone\` at all, because both already name themselves; \`null\` there would wrongly assert "floating". \`endTimeZone\` describes \`end\` the same way but ONLY relative to \`start\` — it appears only when \`end\`'s zone differs from \`start\`'s, which is legal (a flight departing one zone and landing in another), and is omitted whenever \`end\` is absent or shares \`start\`'s zone. ` +
          'WHICH ROWS MAY SIT OUTSIDE THE WINDOW. Rows are filtered EXACTLY against the window you asked for, and all-day events are your account\'s LOCAL days: a date-only value covers that whole day in the configured zone, and an all-day event on a neighbouring day is not returned. Behind that, the range this server REQUESTS of Fastmail is deliberately up to 14 hours wider at each edge than the window you gave, because the server matches an all-day value on its UTC day and reads a floating time as UTC — without the widening it would withhold both from a window narrower than a day, and no filter can keep what was never sent. The extra rows that widening pulls in are then trimmed. TWO KINDS OF ROW CAN STILL SIT OUTSIDE IT. A block that still carries its own recurrence (`recurrenceRule` or `recurrenceDates`) is never dropped whatever its dates say — its start is the series\' ORIGINAL date, which may be years away, and judging it on that would delete a real event rather than misdate it. And a FLOATING timed event comes back from expansion stamped as UTC with the floating marker destroyed, so nothing downstream can move it to your clock; it is judged on UTC and can therefore land in the wrong day for an account far from UTC. THE TWO FAIL IN OPPOSITE DIRECTIONS. A recurrence carrier only ever ADDS a row, so check each `start` against the window you asked for rather than assuming every row is inside it. A floating timed event can be ABSENT from the window it really belongs to, judged into a neighbouring day instead, which is how an account far from UTC loses a row it asked for. A THIRD CASE IS NOT A ROW SITTING OUTSIDE THE WINDOW BUT A ROW THAT NEVER ARRIVES: a series that lists its occurrences as RDATEs and states no RRULE is matched by Fastmail\'s own filter over the series start alone, so a window covering one of its listed dates and not that start returns nothing for it, and no filter on this side can keep what the server never sent. So an empty result is NOT proof of a free day on ANY account, and a "nothing on then" answer built from this call alone can be wrong in the direction that matters; confirm in the Fastmail web interface, or widen the window until it reaches the series start (fork issue #167). ' +
          'A TOTAL calendar-discovery failure is reported as an error, never as an empty list, and a calendarId matching no calendar is an error too — never an empty result. An empty or whitespace-only calendarId is that same error, not "every calendar". A PARTIAL failure is not covered: if the server answers for the account but fails on one collection, that calendar is silently missing from the results and from the total, so a cross-calendar "am I free?" can still be answered from an incomplete set (fork issue #136).',
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: `Which calendar to read (optional; OMIT it to read every calendar — an empty string is NOT a way to ask for all of them). ${CALENDAR_ID_MATCHING_DESC}`,
            },
            startDate: {
              type: 'string',
              description:
                'Start of the window, inclusive. Either a date (2027-03-01, meaning midnight LOCAL time at the start of that day — see the local-day rule in this tool\'s description) or a full datetime (2027-03-01T09:00:00Z or 2027-03-01T09:00:00+10:00, each taken exactly as written; 2027-03-01T09:00:00 with no designator is read as local time). Other spellings such as 2027/03/01 or "March 1 2027" are rejected rather than guessed at, and so is a day its month does not have (2027-02-30) — which would otherwise roll silently into the next month and move the window. THE PAIR IS VALIDATED: if both bounds are given and startDate does not resolve strictly before endDate the call is rejected, and two bounds resolving to the same instant are rejected as a zero-length window rather than answered with nothing.',
            },
            endDate: {
              type: 'string',
              description:
                'End of the window, exclusive. Same accepted spellings and the same local-day rule as startDate, with one difference: a DATE-ONLY endDate covers the WHOLE of that day (2027-03-10 runs through to local midnight at the start of the 11th), so a single-day query is startDate and endDate on the same date. A datetime — with a designator or without — is taken literally as the exclusive end, with no day added. This deliberately differs from create_calendar_event\'s `end`, where a date-only value is exclusive at the START of that day (RFC 5545 DTEND): there the value names one edge of a single event, here it names the last day you are asking about, and reading it as the start of that day would make a one-day query a zero-length window that returns nothing. The pair is validated the same way startDate describes: a backwards or zero-length window is rejected, naming both values and the range they resolved to.',
            },
            limit: {
              type: ['number', 'string'],
              description: 'Maximum number of events to return (default: 50, max: 500). Hard cap, no paging: when the summary line reports a total larger than the number returned, raise `limit` to reach the rest, and narrow the window only if the total is above 500. Recurrence expansion means this is reached sooner than it used to be. The summary line states the total of rows that MATCHED — a series left out for being too dense (see the density clause in the tool description) is not counted there, because the total answers "how many rows did `limit` cut off" and counting occurrences no row represents would send you raising `limit` after rows that are not there. The trailing "Note:" line is what discloses an omitted series.',
              default: 50,
            },
          },
        },
      },
      {
        name: 'get_calendar_event',
        description:
          'Get a specific calendar event by ID. Unlike list_calendar_events this returns `organizer` and `participants` when the event has them, so it is the call to make before deleting or rescheduling anything: it is the only way to see who the server will mail. ' +
          'For a repeating event this returns the SERIES MASTER: `isRecurring` is set, and start/end are the series\' original dates, NOT the next or nearest occurrence. What states the repeat is `recurrenceRule` (the RRULE) or `recurrenceDates` (the raw RDATE values, comma-separated, for a series that LISTS its occurrences rather than stating a rule) — either or both may be present, and this is the path where they show up, because the master is returned unexpanded. `recurrenceDates` CARRIES THE VALUES AND NOT THE PARAMETERS: a TZID, VALUE=DATE or VALUE=PERIOD on the RDATE line is dropped, so a designator-less value in that list does NOT follow the `timeZone` rule that governs `start`, and RDATEs written in different zones are indistinguishable once joined. It is proof that other dates exist, not a date you can place on a clock. To find out when the event actually falls on given days, call list_calendar_events with a startDate/endDate window, which expands the recurrence. THAT ROUTE DOES NOT WORK FOR AN RDATE-ONLY SERIES: Fastmail\'s window filter matches such a record over its `DTSTART` alone, so a window covering one of the `recurrenceDates` values but not the series start returns NOTHING for it — the row never arrives and nothing marks its absence. Widen the window until it reaches the series start (the `start` in this response), or confirm the date in the Fastmail web interface (fork issue #167). The id is the series id, so update_calendar_event and delete_calendar_event on it act on EVERY occurrence. One exception: where a record holds only overridden instances and no master at all, what comes back is one of those overrides and carries a `recurrenceId` — so a `recurrenceId` here means you are NOT looking at the series master. ' +
          `CALENDAR TIMES CARRY A ZONE NAME, NEVER AN OFFSET. \`start\`/\`end\` is a bare local wall clock (2026-04-20T10:00:00), a Z-designated UTC instant, or a date-only (all-day) value. READ THE VALUE'S OWN DESIGNATOR FIRST: \`timeZone\` only QUALIFIES a value that carries neither Z nor a date-only marker, so "absent means the configured zone" applies to a bare wall-clock \`start\` and nothing else. \`timeZone\` names the IANA zone a wall-clock \`start\` is in, but ONLY when it differs from this server's configured zone (${CONFIGURED_TIMEZONE}): an ABSENT \`timeZone\` means ${CONFIGURED_TIMEZONE}, and \`timeZone: null\` means \`start\` is genuinely FLOATING (RFC 5545 §3.3.5 — a different instant for every reader), which is a different fact from "in the configured zone". A Z-designated value or an all-day value never carries \`timeZone\`, because both already name themselves. \`endTimeZone\` describes \`end\` relative to \`start\` — it appears only when \`end\`'s zone differs from \`start\`'s, which is legal (a flight departing one zone and landing in another), and is otherwise omitted. This path returns the stored property exactly as written, with no normalisation, so it is where a floating or differently-zoned value is most likely to show up. Same rule as list_calendar_events.`,
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
        description: `Create a new calendar event. Supports date-only (e.g. 2026-04-01) for all-day events. DTEND is exclusive per RFC 5545 — a one-day event on April 1 needs end: 2026-04-02. start and end must use the SAME form — both date-only, both with a zone designator (Z or +HH:MM), or both without one — and end must be later than start; a mismatched or backwards pair is rejected. ONE EXEMPTION: two values carrying DIFFERENT named time zones (a flight departing one zone and landing in another) are a legal shape whose wall clocks can read backwards, so the ordering check stands down there and a backwards cross-zone pair IS written (fork issue #140). Write both as strict ISO-8601 (2026-04-07, 2026-04-07T14:00:00, 2026-04-07T14:00:00Z, or 2026-04-07T14:00:00+10:00): other spellings such as 2026/04/07 or "April 7 2026" are rejected rather than guessed at, since guessing would place the event on a day that depends on the server's own time zone, and so is a day that does not exist in its month (2026-02-31). ` +
          `CALENDAR TIMES CARRY A ZONE NAME, NEVER AN OFFSET, and this tool never asks you to compute one. \`timeZone\` only QUALIFIES a designator-less start/end — a bare wall clock with no Z and no date-only marker — because that is the one shape with no zone of its own to contradict. Pass \`timeZone\` as an IANA name (e.g. "Australia/Sydney") to say which zone that wall clock is in; OMITTING it does NOT mean floating here — it means writing the event in this server's configured zone (${CONFIGURED_TIMEZONE}), a deliberate difference from update_calendar_event, which never defaults it (see that tool). Combining \`timeZone\` with a start/end that already carries Z/an offset, or with a date-only value, is rejected — both already name their own instant or have no time component, so \`timeZone\` would contradict rather than qualify. \`timeZone: null\` and an empty/whitespace string are rejected too: there is no way to force a genuinely floating write through this parameter. The response states what actually got written — a zone name, UTC, or all-day (no time component) — as one combined statement when start and end land the same way, or a separate statement for each side when they differ (e.g. a cross-zone flight pair), since a caller-named zone, an inherited one, and the configured default all read as the same "no designator" input. ` +
          `participants entries may be { email, name? } objects or bare email-address strings. Text size limits: title, description, location and each participant name/email are capped at ${MAX_ICAL_FIELD_KB}KB each, at most ${MAX_ICAL_PARTICIPANTS} participants, and all of that text together must stay under ${MAX_ICAL_TOTAL_KB}KB. Oversized input is rejected naming the field and the limit — nothing is silently truncated.`,
        inputSchema: {
          type: 'object',
          properties: {
            calendarId: {
              type: 'string',
              description: `Which calendar to create the event in (required). ${CALENDAR_ID_MATCHING_DESC}`,
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
            timeZone: {
              // 'null' is a real, meaningful input here (rejected, with a tailored message
              // explaining there is no way to force a floating write on create) — not absence.
              // A validating client checks a value's type against the schema before this
              // server's own handler ever runs, so restricting this to 'string' would make such
              // a client reject a null timeZone with a generic schema-mismatch error instead of
              // ever reaching that tailored rejection.
              type: ['string', 'null'],
              description: `IANA zone name (e.g. "Australia/Sydney") for a designator-less start/end — the ONE shape it can qualify. Omit to write the account's configured zone (${CONFIGURED_TIMEZONE}); this is create's default and is never floating. MUST contain a region-qualifying slash, or be exactly "UTC" — a bare abbreviation or alias such as "EST", "NZ", "GMT" or "Zulu" is REJECTED even though it resolves to a real zone, because it is ambiguous: "EST" resolves to a fixed-offset zone with no daylight saving, NOT US Eastern. Write "Pacific/Auckland" rather than "NZ". Written as its CANONICAL IANA spelling, which may differ in case or alias from what you passed (e.g. "us/pacific" is written as "America/Los_Angeles"). Rejected: combined with a start/end that already carries Z/an offset or is date-only (both already name themselves), and \`null\`/empty/whitespace (there is no way to force a floating write here — this server never creates a floating calendar time).`,
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
        description: `Update an existing calendar event. SINGLE (NON-REPEATING) EVENTS ONLY: a repeating event is REFUSED, and no parameter overrides that — eventId names the whole series (every occurrence row list_calendar_events returns for a repeating event carries the same id, and so does its url), a patch would move EVERY occurrence past and future, and this server cannot create a repeating event, so it will not rewrite one it has no way to put back. REPEATING covers a series that LISTS its occurrences (an RDATE series, the one that shows \`recurrenceDates\` and no \`recurrenceRule\`) exactly as much as one stating a rule, and a record made only of edited occurrences. Change a repeating event, or one occurrence of it, in the Fastmail web interface instead; get_calendar_event still reads it here. Preserves all existing data (attendees, reminders, recurrence rules, etc.) not being changed. Omit a field to leave it unchanged; passing an empty/whitespace string for title, description, or location is rejected (use clearFields to delete description/location). Floating times (no Z/offset) preserve the original timezone; explicit UTC/offset times convert to UTC. A new start/end is checked against the value it will sit beside — the other one you passed, or the stored one you left alone: they must end up in the same form (both date-only, both UTC, both floating, or both in the same TZID) and end must be later than start, otherwise the update is rejected. ONE EXEMPTION: when both values end up carrying DIFFERENT named time zones the ordering check stands down, because a flight departing one zone and landing in another is a legal event whose wall clocks read backwards — so a backwards cross-zone pair IS written (fork issue #140). So moving an event to a different day or converting only one side to UTC means passing BOTH start and end. Both are read as strict ISO-8601, matching create_calendar_event: a non-ISO spelling like 2026/04/07, or a day its month does not have (2026-02-31), is rejected rather than guessed at. ` +
          `CALENDAR TIMES CARRY A ZONE NAME, NEVER AN OFFSET. \`timeZone\` only QUALIFIES a designator-less start/end you are ALSO passing in this same call — the one shape with no zone of its own to contradict. Unlike create_calendar_event, OMITTING \`timeZone\` never defaults to the configured zone here: it leaves start/end exactly as today — an inherited stored TZID stays, or a value with no stored zone stays floating. Rejected: \`timeZone\` combined with a start/end that already carries Z/an offset or is date-only (both already name themselves); \`timeZone\` with NEITHER start nor end (still reachable — re-send start and/or end unchanged alongside it to re-zone them); \`timeZone\` with only ONE of start/end when the untouched side is stored in a DIFFERENT named zone, because re-zoning just one side would silently strand the other into a two-zone event — pass BOTH start and end (re-sending the one you are not otherwise moving, unchanged) to change the zone; and \`null\`/empty/whitespace (there is no way to force a floating write through this parameter). For each side actually written this call, the response states what ended up there — a zone name, UTC, all-day (no time component), or floating (no zone). ` +
          `WARNING: providing participants replaces ALL existing attendee data (acceptance status, roles, etc.). participants: [] removes all attendees, and its entries may be { email, name? } objects or bare email-address strings. Text size limits match create_calendar_event: title, description, location and each participant name/email are capped at ${MAX_ICAL_FIELD_KB}KB each, at most ${MAX_ICAL_PARTICIPANTS} participants, and all of that text together must stay under ${MAX_ICAL_TOTAL_KB}KB. Oversized input is rejected naming the field and the limit — nothing is silently truncated.`,
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to update, or equivalently the `url` from a list_calendar_events row — both resolve to the same record and both are refused when it repeats. An id taken from an occurrence row of a repeating event is the SERIES id, which is why this tool refuses it.',
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
            timeZone: {
              // Same reasoning as create_calendar_event's timeZone: 'null' is a real, rejected
              // input (there is no way to force a floating write through this parameter), not
              // absence, so it must stay a valid schema type or a validating client would reject
              // it with a generic error before this server's own tailored message is reached.
              type: ['string', 'null'],
              description: 'IANA zone name (e.g. "Australia/Sydney") for a designator-less start/end you are ALSO passing this call — the ONE shape it can qualify. Omitting it never defaults to a configured zone here (unlike create_calendar_event): a stored TZID is inherited unchanged, or the value stays floating. MUST contain a region-qualifying slash, or be exactly "UTC" — a bare abbreviation or alias such as "EST", "NZ", "GMT" or "Zulu" is REJECTED even though it resolves to a real zone, because it is ambiguous: "EST" resolves to a fixed-offset zone with no daylight saving, NOT US Eastern. Write "Pacific/Auckland" rather than "NZ". Written as its CANONICAL IANA spelling, which may differ in case or alias from what you passed (e.g. "us/pacific" is written as "America/Los_Angeles"). Rejected: with neither start nor end (re-send one unchanged alongside it to re-zone); with only one of start/end when the untouched side is stored in a different named zone (pass both, or omit timeZone); combined with a start/end already carrying Z/an offset or date-only; and `null`/empty/whitespace.',
            },
            participants: participantsSchemaProperty(
              `Replaces ALL existing attendees (at most ${MAX_ICAL_PARTICIPANTS}). Empty array removes all attendees. Omit to preserve existing attendees.`,
            ),
            clearFields: {
              type: 'array',
              items: { type: 'string', enum: ['description', 'location'] },
              description: 'Property names to delete from the event. Allowed: description, location. Cannot also pass the same field as a value.',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'delete_calendar_event',
        description:
          'Delete a calendar event by ID.' +
          ' SINGLE (NON-REPEATING) EVENTS ONLY: a repeating event is REFUSED, and no parameter overrides that. eventId names the calendar record, and every occurrence row list_calendar_events returns for a repeating event carries that same id — so deleting the row for one Thursday would destroy the entire series, every past and future occurrence with it, and mail every attendee a cancellation. This server cannot create a repeating event, so it will not destroy one it has no way to put back. REPEATING covers a series that LISTS its occurrences (an RDATE series, the one that shows `recurrenceDates` and no `recurrenceRule`) exactly as much as one stating a rule, and a record made only of edited occurrences. Delete a repeating event, or one occurrence of it, in the Fastmail web interface instead.' +
          ' SENDS MAIL: if the event has attendees, the server emails every one of them a cancellation from this account as soon as it is deleted. That is the server scheduling layer, not a compose tool, so it happens with no send_draft call — deleting an event with a real person on it is contacting them. For a series that is one cancellation for the whole thing, to every attendee.' +
          ' The delete is irreversible and nothing is echoed back: a calendar event does not go to Trash the way an email does, and this tool reads nothing before destroying it. Call get_calendar_event first if there is any chance the id is wrong — that response is the only copy you will have, and it is also how you find out whether anyone is about to be mailed.',
        inputSchema: {
          type: 'object',
          properties: {
            eventId: {
              type: 'string',
              description: 'ID of the event to delete, or equivalently the `url` from a list_calendar_events row — the two name the same record, destroy the same amount, and are both refused when it repeats. The id is checked only for whether it repeats: a wrong id that names a real single event still deletes it and mails its attendees.',
            },
          },
          required: ['eventId'],
        },
      },
      {
        name: 'list_identities',
        description: "List sending identities (email addresses that can be used for sending). Returns simplified format by default (name, email, replyTo, and the identity's configured signature as textSignature/htmlSignature when it has one; a blank signature is omitted). JMAP does not append the signature server-side, so signing is a choice each compose call makes: pass appendSignature:true to create_draft/reply_email/forward_email and the identity's own signature is placed in the body for you (above the quoted history on a reply), or read the field from here and write it into the body yourself. Use verbose=true only if you need extra fields like SMTP config or verification state. Use raw=true for original JMAP response.",
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
        description: 'Get the newest emails across all mailboxes (or one, via mailbox) — the small, cheap read. NOT an arrivals feed: "all mailboxes" means every folder except Trash and Spam, so your own Sent copies, drafts and any custom folder come back alongside newly received mail. Narrow it with mailbox:"inbox" if you want delivered mail only, and excludeDrafts to drop drafts. Trash and Spam are excluded by default (set includeTrash/includeSpam to include them). Set mailbox to scope to a single mailbox (incl. Trash/Spam), which ignores the default exclusion. ' + SCOPE_RELIABILITY_CONTRACT + ' It filters exactly as list_emails does; the difference is size, not scope — this returns 10 by default and caps at 50, list_emails returns 20 and caps at 100, so reach for list_emails when you are working through a folder rather than taking a quick look. One mailbox is all this tool scopes to: to intersect several mailboxes or exclude some, use search_emails (requiredMailboxes / excludeMailboxes) with no query. Returns simplified format (metadata + preview, no bodies). Use raw=true for original JMAP response. For email bodies, use get_email. The date field is rendered in local time with a UTC offset (e.g. 2026-03-02T08:00:00+10:00), not UTC; raw=true returns the canonical JMAP UTC time. ' + LOCATION_FIELDS_DESC + ' ' + PREVIEW_SIZE_DESC + ' ' + COMPACT_ATTACHMENT_DESC + ' ' + FIELDS_TOOL_DESC + ' ' + POSITION_TOOL_DESC,
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
              description: MAILBOX_PARAM_DESC,
            },
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
              description: targetMailboxParamDesc('delete_email'),
            },
          },
          required: ['emailId', 'targetMailbox'],
        },
      },
      {
        name: 'archive_email',
        description:
          'Archive one or more emails the way the Fastmail client does. Takes an ARRAY (`emailIds`), so it handles one message or a batch — there is no separate bulk_archive. ' +
          'It archives exactly the messages you name and nothing else. Fastmail\'s own Archive button acts on a single message or on the whole conversation depending on a per-user display setting this server cannot see, so to archive a CONVERSATION pass every message id in it (get_thread lists them) rather than just one; passing one leaves the rest of the thread in the Inbox. ' +
          'What it actually does: it REMOVES the message from the Inbox and leaves every other folder and label in place, adding the Archive folder only when removing the Inbox would otherwise leave the message filed nowhere. ' +
          'So a message filed in the Inbox plus a label keeps the label and does NOT go to Archive; a message filed only in the Inbox moves to Archive; a message that already left the Inbox is left completely untouched. ' +
          'A no-op is a legitimate, successful outcome here, not a failure. ' +
          'Once a message has LEFT the Inbox, it REFUSES one in Trash, Spam, Drafts, Scheduled, Sent or Snoozed, because the Fastmail client offers no Archive action there either; the refusal says why, and names an alternative in this server where one exists (for a scheduled send or a snooze there is none, and it says so). The Inbox is tested first, so a message somehow in the Inbox AND one of those is archived rather than refused, keeping that membership and gaining nothing — except that a message also in Scheduled may come back as `failed`, because the server appears to reject re-asserting a scheduled membership outside a send request; that combination has not been measured yet. ' +
          'Never throws on a partial failure: the result reports each id separately as movedToArchive, removedFromInbox, notInInbox, refused, notFound or failed, with counts that sum to the number of distinct ids you passed (duplicates are collapsed). The JSON beside the summary is `{ counts, results }`, so both the per-id detail and the counts are parseable. ' +
          'Each entry\'s `mailboxes`/`roles` are the PROJECTED filing for the two branches that wrote, the OBSERVED unchanged filing for notInInbox and refused (which write nothing), and for failed either the filing as OBSERVED BEFORE the write was attempted or, when no write was attempted for it, nothing at all; read roles for "archive" to tell whether a message is in Archive, since the branch name alone will not say. Both fields are ABSENT on a notFound entry (there is no filing to report) and on the failed sub-case where the current filing could not be READ (the server returned no mailboxIds object, or an empty one, which is not a filing a message can have; its filing was never observed, which is why it failed), and `roles` is absent whenever nothing the message is filed in has a role. A `unresolvedMailboxIds` on an entry means a mailbox id could not be resolved to a name, so `mailboxes`/`roles` are incomplete for that message and those raw ids are the remainder — which is also what tells you how to read an absent `roles`: absent with no `unresolvedMailboxIds` means no role mailboxes, absent WITH them means the roles are unknown for the ids listed there. ' +
          'The Archive destination is found by JMAP role, never by folder name, so a folder merely NAMED "archive" is not it, and there is no destination parameter — use move_email to file into anything else. ' +
          'The whole call throws, rather than reporting per message, in four account-wide cases. Three happen BEFORE anything is written, so nothing was archived and the message says so: a read that failed or came back incomplete, an account with no inbox-role mailbox, and an account with no archive-role mailbox when a message actually needed Archive (use move_email instead). The fourth is the write itself failing, which happens AFTER it was dispatched — there the outcome of the batch is unknown and you should re-read the messages rather than assume nothing changed. A per-message problem never throws. ' +
          'No keyword is written, but that is not the same as the read state being untouched: $seen is reported only when every one of a message\'s per-mailbox copies carries it, so dropping an unread Inbox copy can flip a message to read. ' +
          'This describes an account in LABELS mode. In folders mode a message has a single membership, so every archive is the move-to-Archive case.',
        inputSchema: {
          type: 'object',
          properties: {
            emailIds: {
              // Declared array-OR-string so a validating client cannot reject the
              // stringified form before the lenient coercion runs (#98).
              type: ['array', 'string'],
              items: { type: 'string' },
              description: 'IDs of the emails to archive — message ids, NOT thread ids. Pass an array even for a single message; duplicates are collapsed. To archive a whole conversation, pass every message id in the thread.',
            },
          },
          required: ['emailIds'],
        },
      },
      {
        name: 'add_labels',
        description: 'Add labels (mailboxes) to an email without removing existing ones. Each label mailbox may be given by id, role (e.g. inbox), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.' + LABEL_NAMESPACE_DESC,
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
        description: 'Remove specific labels (mailboxes) from an email. Each label mailbox may be given by id, role (e.g. inbox), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.' + LABEL_NAMESPACE_DESC + LABEL_REMOVAL_RESCUE_DESC,
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
        description: 'List an email\'s parts, as raw JMAP part objects (partId, blobId, type, size, name, disposition, cid) rather than the simplified shape the read tools return. ' + UNION_SCOPE_DESC + ' A body-embedded part usually reports disposition:null rather than "inline", and nothing in this raw listing tells it apart from a genuinely attached file — the derived isInline flag lives only in get_email (and get_thread with includeBodies), so cross-check there before acting on an entry, e.g. before handing its blobId to edit_draft removeAttachments, which would strip an image the body still displays. ' + MINTED_CID_NONDURABILITY + ' This listing is where an attachmentId comes from: pass back an entry\'s partId or blobId — those are the durable handles, to download_attachment and to an attachments item that attaches this part to a new message (emailId + attachmentId, which needs FASTMAIL_ALLOW_BLOB_ATTACH). This listing is also what download_attachment\'s entry numbers count from: its first entry is attachmentId "0". Entry numbers are read-only and positional — they shift whenever the listing does, and attaching a part to outgoing mail rejects one.',
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
              description: ATTACHMENT_REF_DESC,
            },
            path: {
              type: 'string',
              description: `DESTINATION file path to WRITE the attachment to. This is the opposite direction from the compose tools' attachments[].path, which is a SOURCE file read under FASTMAIL_ATTACH_DIR; this one is written under FASTMAIL_DOWNLOAD_DIR. May be absolute or relative; relative paths resolve against ${getDownloadDir() || '~/Downloads/fastmail-mcp/'} (configurable via FASTMAIL_DOWNLOAD_DIR), so a bare filename lands there in one step. Absolute paths must fall within that directory; traversal or symlink escape outside it is rejected for security. To save directly into your own location, set FASTMAIL_DOWNLOAD_DIR to that root. Parent directories will be created automatically.`,
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
              description: targetMailboxParamDesc('bulk_delete'),
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
        description: 'Add labels to multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. inbox), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.' + LABEL_NAMESPACE_DESC,
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
        description: 'Remove labels from multiple emails simultaneously. Each label mailbox may be given by id, role (e.g. inbox), name, or path; an unknown or ambiguous mailbox rejects the whole call with the valid list.' + LABEL_NAMESPACE_DESC + LABEL_REMOVAL_RESCUE_DESC + ' The rescue is decided per message, so a batch can archive some and merely unlabel others. A rejection is NOT decided per message: one unservable id rejects the whole batch before anything is written. When some messages succeed and others fail at the server, the error still names any message the call filed in Archive.',
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
              text: raw ? toolJson(email) : toolJson(projectEmail(simplified, fields)),
            },
          ],
        };
      }

      case 'reply_email': {
        // The orchestration (fetch original, assemble reply, upload + thread attachments,
        // save the draft) lives in composeReply so it is unit-testable with a mock client;
        // this handler just maps the result to the response text.
        const result = await composeReply(args, client, getAttachDir(), getAllowBlobAttach());
        const text = `Reply draft saved successfully (Email ID: ${result.emailId}). Use send_draft to transmit it. Subject: ${result.subject}${formatInlineNotes(result.notes)}`;
        return { content: [{ type: 'text', text }] };
      }

      case 'forward_email': {
        // The orchestration (fetch original, assemble the forwarded-message block,
        // carry/.eml attachments, upload new ones, save the draft) lives in
        // composeForward so it is unit-testable with a mock client; this handler just
        // maps the result to the response text.
        const result = await composeForward(args, client, getAttachDir(), getAllowBlobAttach());
        const text = `Forward draft saved successfully (Email ID: ${result.emailId}). Use send_draft to transmit it. Subject: ${result.subject}${formatInlineNotes(result.notes)}`;
        return { content: [{ type: 'text', text }] };
      }

      case 'create_draft': {
        // The orchestration (body validation, recipient coercion, the contentless-draft
        // guard, attachment upload, create) lives in composeDraft so it is unit-testable
        // with a mock client; this handler just maps the result to the response text.
        const { emailId, subject, to, cc, notes } = await composeDraft(args, client, getAttachDir(), getAllowBlobAttach());

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
        // The orchestration (field coercion, the body guard, attachment upload/resolution,
        // the update) lives in editDraft so it is unit-testable with a mock client — the
        // attachment seam decides whether a capability gate refuses the call, and a gate
        // proved only by a live run is not regression protection. This handler just maps
        // the result to the response text.
        const updateResult = await editDraft(args, client, getAttachDir(), getAllowBlobAttach());

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
              text: toolJson(output),
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
        return { content: [{ type: 'text', text: toolJson(calendars) }] };
      }

      case 'list_calendar_events': {
        const { calendarId, limit, startDate, endDate } = args as any;
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const { events, total, windowClamp, denseSeries } = await davClient.getCalendarEvents(calendarId, clampLimit(limit, 50, 500), startDate, endDate);
        // Rendered through the shared seam every other list/search read tool uses, so the
        // summary wording cannot drift here on its own. Unpaged: this tool takes no
        // `position`, so formatQueryResult offers no nextPosition — passing one back would be
        // rejected by the unknown-parameter guard. Recurrence expansion makes the count
        // load-bearing, since one fortnightly event across a quarter is now several rows.
        //
        // The window note rides AFTER the JSON, like the Trash/Spam exclusion note on the
        // email listings, so the JSON block stays parseable. It appears only when the window
        // queried was not the window asked for. Both the wording and the blank-line separator
        // come from the note builder, exactly as buildExclusionNote owns them for the email
        // listings — the handler concatenates and decides nothing.
        //
        // The density note rides after the window note, in that order: the window note says
        // which days were searched, the density note says what inside them was too dense to
        // materialise, and the second only makes sense once the first has named the range.
        return { content: [{ type: 'text', text: `${formatQueryResult({ items: events, total })}${buildCalendarWindowNote(windowClamp)}${buildCalendarDensityNote(denseSeries)}` }] };
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
        return { content: [{ type: 'text', text: toolJson(event) }] };
      }

      case 'create_calendar_event': {
        const { calendarId, title, description, start, end, location, timeZone } = args as any;
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
        const result = await davClient.createCalendarEvent({
          calendarId, title, description, start, end, location, participants, timeZone,
        });
        return { content: [{ type: 'text', text: `Calendar event created. Event ID: ${result.eventId}.${describeCreateCalendarEventResult(result)}` }] };
      }

      case 'update_calendar_event': {
        const { eventId, title, description, start, end, location, timeZone } = args as any;
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
        if (!eventId) {
          throw new McpError(ErrorCode.InvalidParams, 'eventId is required');
        }
        const hasClearFields = Array.isArray(clearFields) && clearFields.length > 0;
        // timeZone counts as a field here so a timeZone-only call reaches
        // updateCalendarEvent's own "timeZone with neither start nor end" rejection
        // instead of this generic message masking it.
        if (title === undefined && description === undefined && start === undefined && end === undefined && location === undefined && participants === undefined && timeZone === undefined && !hasClearFields) {
          throw new McpError(ErrorCode.InvalidParams, 'At least one field to update must be provided (title, description, start, end, location, participants, timeZone, or clearFields)');
        }
        const davClient = initializeCalDAVClient();
        if (!davClient) {
          throw new McpError(ErrorCode.InvalidRequest, 'CalDAV not configured. Set FASTMAIL_CALDAV_USERNAME and FASTMAIL_CALDAV_PASSWORD.');
        }
        const fields = { title, description, start, end, location, participants, clearFields, timeZone };
        const result = await davClient.updateCalendarEvent(eventId, fields);
        return { content: [{ type: 'text', text: `Calendar event updated. Event ID: ${eventId}${describeUpdateCalendarEventResult(result)}` }] };
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
              text: toolJson(output),
            },
          ],
        };
      }

      case 'get_recent_emails': {
        // No `= 'inbox'` default: this tool reads all mail and scopes only when the caller
        // says so (#29). A blank or whitespace value is an omitted value all the way down.
        const { mailbox } = args as any;
        // clampLimit, and it is the ONLY clamp on this path — the client deliberately
        // passes `limit` straight to JMAP. The schema advertises a string limit, and
        // Math.min("abc", 50) is NaN, which JMAP serializes as `"limit": null` — a query
        // with no bound at all.
        const limit = clampLimit((args as any).limit, 10, 50);
        // coerceBool, not !!: a lenient client's stringified "false" is truthy and would
        // silently reverse the sort order. Defaults to false (newest first). The scope
        // flags below coerce the same way.
        //
        // That each flag is read from the argument of its OWN name is asserted by a source
        // scan in tool-schema.test.ts — a swap here type-checks and silently hides the
        // wrong folder. What that scan does not reach is the `?? false` defaults and the
        // note append below: they are only exercised by running the server, and this
        // handler is a destructure-and-delegate, so it stays inline rather than being
        // pulled behind an injected client (the extraction CLAUDE.md prescribes is for
        // handlers that orchestrate). Accepted, not overlooked.
        const ascending = coerceBool((args as any).ascending) ?? false;
        // Same coercion for raw — see list_emails for why `!!` was wrong here.
        const raw = coerceBool((args as any).raw) ?? false;
        // Validated before the query so a typo'd field name costs no round trip.
        const fields = parseEmailFields((args as any).fields, { raw });
        // Same reason: an unusable paging offset is rejected before the query runs.
        const position = coercePosition((args as any).position);
        const client = initializeClient();
        const result = await client.getRecentEmails({
          mailbox,
          limit,
          position,
          ascending,
          includeTrash: coerceBool((args as any).includeTrash) ?? false,
          includeSpam: coerceBool((args as any).includeSpam) ?? false,
          excludeDrafts: coerceBool((args as any).excludeDrafts) ?? false,
        });
        // Append the exclusion note (if any) out of band, the same way list_emails does, so
        // the JSON block stays parseable on both the raw and the simplified path.
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
        // Strict, not the lenient coerceStringArray: that one maps every element through
        // String(), so `emailIds: [null]` would reach Email/get as the literal id "null"
        // and come back reported as notFound — a type error wearing a not-found error's
        // clothes, which would falsify this tool's promise that notFound means the server
        // did not know the id.
        const emailIds = coerceStringArrayStrict((args as any).emailIds, 'emailIds');
        if (!emailIds || emailIds.length === 0) {
          throw new McpError(ErrorCode.InvalidParams, 'emailIds array is required and must not be empty');
        }
        const client = initializeClient();
        const result = await client.archiveEmails(emailIds);
        return {
          content: [
            { type: 'text', text: formatArchiveResult(result) },
            // The whole result as its own item, so the summary prose never has to be parsed
            // and the JSON stays parseable. `counts` rides along with `results` because both
            // descriptions promise counts that sum to the distinct id count, and a caller
            // cannot check that invariant against prose — a bucket with no entries produces
            // no line at all.
            //
            // Redacted like the prose beside it. This path RETURNS rather than throws, so the
            // CallTool catch never sees it, and the renderer's own redaction of a set-error
            // description would be decorative if the same server string then went out verbatim
            // one content item later. Applied to every string VALUE, so it covers descriptions,
            // mailbox names and ids alike rather than a field list that has to be kept in step
            // with the type.
            //
            // Redaction is ALL this item gets, and that is a deliberate difference from the
            // prose: the renderer passes mailbox names through describePart, which also strips
            // format characters such as a bidi override. Here the defence is JSON quoting,
            // which escapes quotes and control characters but leaves a bidi override intact.
            // This item is DATA — a caller parses it rather than reading it as the server
            // speaking — and truncating or rewriting names inside it would corrupt the values
            // it exists to convey. The prose is the neutralised surface; this one is verbatim
            // by design.
            // redactedJson, NOT redactBearerTokens(JSON.stringify(...)): redacting a finished
            // JSON document lets the bearer pattern eat the delimiters that terminate a value,
            // and this item's whole promise is that it parses. See its definition in coerce.ts.
            { type: 'text', text: redactedJson({ counts: result.counts, results: result.results }) },
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
        const { rescued, unchangedCount } = await client.removeLabels(emailId, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: formatLabelRemoval(rescued, 1, unchangedCount),
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
          // Everything else goes to the top-level catch, which maps a bad emailId /
          // attachmentId to InvalidParams naming what to pass instead, and a transport or
          // JMAP failure to InternalError carrying the server's own reason. Both are
          // redacted there, so this tool needs no error handling of its own beyond the
          // path branch above.
          throw error;
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
        // The scope arrays coerce STRICTLY: a present-but-uncoercible value is rejected
        // rather than read as absent. Dropping one would silently widen the query the
        // caller passed it to narrow, and the results would look like a complete answer.
        const requiredMailboxes = coerceStringArrayStrict((args as any).requiredMailboxes, 'requiredMailboxes');
        const excludeMailboxes = coerceStringArrayStrict((args as any).excludeMailboxes, 'excludeMailboxes');
        const client = initializeClient();
        const validLimit = Math.min(Math.max(Number(limit) || 20, 1), 100);
        const result = await client.searchEmails({
          query, from, to, cc, bcc, subject,
          hasAttachment: coerceBool(hasAttachment),
          isUnread: coerceBool(isUnread),
          isPinned: coerceBool(isPinned),
          mailbox, requiredMailboxes, excludeMailboxes,
          after, before, limit: validLimit, position,
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
              text: toolJson(stats),
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
              text: toolJson(summary),
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
        const { rescued, unchangedCount } = await client.bulkRemoveLabels(emailIds, mailboxIds);
        return {
          content: [
            {
              type: 'text',
              text: formatLabelRemoval(rescued, new Set(emailIds).size, unchangedCount),
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
              text: toolJson(availability),
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
        // serializes as `"limit": null` — an unbounded metadata dump. This call
        // bypasses the get_recent_emails handler, so this IS the bound on this path
        // (the client no longer clamps).
        //
        // `mailbox: 'inbox'` stays explicit. get_recent_emails now defaults to all mail,
        // but this diagnostic wants ordinary delivered messages to mark read and pin, and
        // an all-mail read would hand it drafts and sent copies to write to.
        const testLimit = clampLimit(limit, 3, 10);
        const { items: emails } = await client.getRecentEmails({ limit: testLimit, mailbox: 'inbox' });

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
                text: `BULK OPERATIONS TEST (DRY RUN)\n\n${toolJson(results)}\n\nTo actually execute the test, set dryRun: false`,
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
                text: `BULK OPERATIONS TEST (EXECUTED)\n\n${toolJson(results)}`,
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
  // Resolve the display and window-interpretation zone once, before any tool handler can fire.
  // One stored value does both jobs: every email `date` renders in it, AND list_calendar_events
  // interprets a date-only window bound as a whole day in it, so the day a calendar query covers
  // and the day an email is dated cannot drift apart.
  //
  // FASTMAIL_TIMEZONE (or the host zone it falls back to when unset) is held to the same slash
  // rule as the caller-supplied `timeZone` parameter (#157 amendment): a shorthand or
  // unresolvable OPERATOR-SET value refuses to start this server, rather than silently rendering
  // every date in the wrong zone with nothing said. A rejected HOST zone is different — nobody
  // configured it, so it falls back to UTC with a loud one-time warning instead of making an
  // unconfigured machine unusable. This has to run here, not at module load: a module-level throw
  // would make src/index.ts unimportable, breaking every test that imports it, and would produce
  // a stack trace instead of an operator-readable refusal. See resolveConfiguredTimezone's own
  // comment in coerce.ts for the full reasoning.
  let resolvedTimezone: { zone: string; warning?: string };
  try {
    resolvedTimezone = resolveConfiguredTimezone(getTimezone());
  } catch (err) {
    console.error(err instanceof Error ? err.message : String(err));
    process.exit(1);
  }
  if (resolvedTimezone.warning) {
    console.error(resolvedTimezone.warning);
  }
  setDefaultTimezone(resolvedTimezone.zone);
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Fastmail MCP server running on stdio');
}

runServer().catch(() => {
  // Avoid logging raw error objects to prevent accidental PII leakage
  console.error('Fastmail MCP server failed to start');
  process.exit(1);
});