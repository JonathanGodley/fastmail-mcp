import { simplifyEmail } from './email-formatter.js';
import { projectEmail } from './field-projection.js';
import type { QueryResult, ReplacedDraftInfo, UpdateDraftResult } from './jmap-client.js';
import type { SendDraftResult } from './send-draft-handler.js';

// The query-level summary that heads every list/search response, so the count wording
// is written once and can't drift between the raw and simplified paths (#51).
//
// `total` is ALWAYS stated, on every listing tool. A bare "20 results." reads
// identically whether 20 is the whole match set or just the page, and a caller that
// reads a capped page as the whole answer concludes "nothing else matched" when plenty
// did.
//
// `paged` says whether the CALLING TOOL accepts a `position` parameter, and only a
// paged tool is told how to continue:
//
//   - `nextPosition` appears only when more results exist, and only for a paged tool.
//     Its absence on a paged tool is the published "this listing is complete" signal,
//     so there is no `hasMore: false` to interpret — the same discipline as the
//     Trash/Spam note, where silence means nothing was withheld.
//   - Emitting it for an unpaged tool (the contacts listings, which render through
//     formatQueryResult on their raw path) would be an instruction the caller cannot
//     follow: `position` is not in those tools' schemas, so passing it back is
//     rejected outright by the unknown-parameter guard. They still get the total.
//   - The arithmetic uses the position the page actually started at and the number of
//     items actually returned, never the requested `limit`. A short final page (the
//     server had fewer left than asked for) must not advertise another page.
//
// A position past the end is not an error: JMAP clamps it and returns an empty page
// with the real total, which is self-describing ("0 of 137") and idempotent, so the
// caller can see it overshot and re-ask from a valid offset.
export function formatQuerySummary(result: QueryResult, options?: { paged?: boolean }): string {
  const { items, total, position } = result;
  const start = typeof position === 'number' && position > 0 ? position : 0;
  const from = start > 0 ? ` from position ${start}` : '';

  // `calculateTotal` is server-discretionary (RFC 8620 section 5.5), so a total can be
  // absent. Say so rather than printing the page size as if it were the total — that
  // substitution is the exact miscount this summary exists to prevent.
  if (typeof total !== 'number') {
    const consequence = options?.paged ? ', so whether more results exist is unknown' : '';
    return `Showing ${items.length} results${from}; the total match count was not returned${consequence}.`;
  }

  const next = start + items.length;
  const more = options?.paged && next < total
    ? ` nextPosition: ${next} (pass position:${next} for the next page).`
    : '';
  return `Showing ${items.length} of ${total} results${from}.${more}`;
}

// Raw (untransformed JMAP) rendering for a listing tool that does NOT take a
// `position` — the contacts listings. Paging is a property of the tool, so it is
// carried by which renderer the handler picks rather than by a flag each call site has
// to remember to pass: forgetting a flag would silently drop a promised signal, while
// reaching for the wrong function is visible in the handler.
export function formatQueryResult(result: QueryResult): string {
  return `${formatQuerySummary(result)}\n${JSON.stringify(result.items, null, 2)}`;
}

// Raw rendering for the three paging email tools (list_emails, search_emails,
// get_recent_emails), the counterpart to formatEmailQueryResult below: same summary,
// untransformed items.
export function formatRawEmailQueryResult(result: QueryResult): string {
  return `${formatQuerySummary(result, { paged: true })}\n${JSON.stringify(result.items, null, 2)}`;
}

// The one seam every list/search read tool renders through (list_emails,
// search_emails, get_recent_emails), so `fields` projection lands on all three at
// once and cannot drift between them. The summary line (its result count and
// `nextPosition`) and the trailing exclusion note are NOT fields and are never
// projected away — they are out-of-band signals about the query, and a caller
// silently losing "N results were withheld" or "there is another page" while asking
// for a narrower shape would be a scope lie, not a smaller response.
export function formatEmailQueryResult(result: QueryResult, options?: { fields?: ReadonlySet<string> }): string {
  const simplified = result.items.map(e => projectEmail(simplifyEmail(e), options?.fields));
  return `${formatQuerySummary(result, { paged: true })}\n${JSON.stringify(simplified, null, 2)}`;
}

// Recipient lists are capped so a big draft can't turn the result into a wall of
// addresses; the trashed copy holds the full picture.
const MAX_ECHOED_RECIPIENTS = 5;
function formatReplacedRecipients(label: string, addresses?: string[]): string | null {
  if (!addresses?.length) return null;
  const shown = addresses.slice(0, MAX_ECHOED_RECIPIENTS).join(', ');
  const extra = addresses.length - MAX_ECHOED_RECIPIENTS;
  return `${label} ${shown}${extra > 0 ? ` (+${extra} more)` : ''}`;
}

// Render the fingerprint of the draft an edit replaced (#65). Returns '' when that draft
// carried nothing worth echoing.
function formatReplacedDraft(replaced: ReplacedDraftInfo): string {
  const parts = [
    replaced.subject ? `subject "${replaced.subject}"` : null,
    formatReplacedRecipients('to', replaced.to),
    formatReplacedRecipients('cc', replaced.cc),
    replaced.htmlBodySize != null ? `htmlBody ${replaced.htmlBodySize} chars` : null,
    replaced.textBodySize != null ? `textBody ${replaced.textBodySize} chars` : null,
  ].filter(Boolean);
  return parts.join(', ');
}

// The edit_draft result text. An edit recreates the message (JMAP content is immutable),
// so this has to say three things: the new id, where the old copy went, and what that old
// copy contained — the last one so a caller that edited from a stale copy can see at once
// that it overwrote something it didn't know about, and restore it from Trash (#65).
export function formatEditDraftResult(result: UpdateDraftResult): string {
  const disposal = result.trashedOldDraftId
    ? `The previous draft (id ${result.trashedOldDraftId}) was moved to Trash, where it stays recoverable until Trash is emptied or auto-purged.`
    : `WARNING: the previous draft (id ${result.orphanedOldDraftId}) could NOT be moved to Trash (${result.orphanedOldDraftReason ?? 'reason unknown'}), so it remains in place as a duplicate holding the pre-edit content; delete it if you don't want it.`;
  const fingerprint = formatReplacedDraft(result.replacedDraft);
  const replaced = fingerprint
    ? ` It contained: ${fingerprint}. If that isn't what you expected to replace, the draft changed since you last read it and this edit overwrote those changes.`
    : '';
  return `Draft updated successfully. New Email ID: ${result.id}. ${disposal}${replaced}`;
}

// The send_draft result text. Reports the submission, then what happened to the message
// the draft was composed from (#60): marked when the original was identified and updated,
// and — because the caller never named that original — an explicit note when the draft
// pointed at a message this server could not pin down, so the skip is actionable rather
// than invisible. A keyword-write failure after a successful lookup is deliberately not
// reported, matching reply_email/forward_email on the same failure.
export function formatSendDraftResult(result: SendDraftResult): string {
  const base = `Draft sent successfully. Submission ID: ${result.submissionId}`;
  const km = result.keywordMaintenance;
  if (!km) return base;

  const marking = km.kind === 'reply' ? 'answered and read' : 'forwarded and read';
  if (km.marked) return `${base} Original marked ${marking}.`;
  if (!km.skipReason) return base; // keyword write failed; the draft still sent

  const relation = km.kind === 'reply' ? 'replies to' : 'forwards';
  const why = km.skipReason === 'ambiguous'
    ? 'more than one stored message carries that Message-ID'
    : km.skipReason === 'lookup-failed'
      ? 'the lookup failed'
      : 'no stored message carries that Message-ID';
  return `${base} The message this draft ${relation} (Message-ID ${km.messageId}) was not marked ${marking}: ${why}.`;
}

// Build the trailing Trash/Spam exclusion note from QueryResult.exclusion (the
// out-of-band metadata that searchEmails/getEmails populate; the formatters above
// deliberately ignore it). Returns '' when there is nothing to disclose. The note is
// appended by the handler to the formatter's string (raw + simplified), so the JSON
// block stays parseable. Three independent signals, fail-loud ones FRONT-LOADED with
// the imperative so a model that learned "no note = safe" can't skim past them:
//   - unresolved role  -> the folder couldn't be found, so it was NOT excluded
//   - hidden === null   -> excluded, but the count couldn't be confirmed (degraded)
//   - hidden > 0        -> N matches were withheld to Trash/Spam
//   - hidden === 0      -> NO note (silence is the published "nothing matched" signal)
export function buildExclusionNote(exclusion?: QueryResult['exclusion']): string {
  if (!exclusion) return '';
  const { hidden, excludedRoles, unresolvedRoles } = exclusion;
  const flagFor = (role: string) => (role === 'Trash' ? 'includeTrash:true' : 'includeSpam:true');
  const notes: string[] = [];

  if (unresolvedRoles && unresolvedRoles.length > 0) {
    notes.push(
      `Re-run to be sure: the ${unresolvedRoles.join('/')} folder couldn't be found, so it was NOT excluded — these results may include ${unresolvedRoles.join('/')} mail.`,
    );
  }

  if (excludedRoles && excludedRoles.length > 0) {
    const flags = excludedRoles.map(flagFor).join(' / ');
    if (hidden === null) {
      notes.push(
        `Re-run with ${flags}: ${excludedRoles.join('/')} were excluded but the hidden count couldn't be confirmed.`,
      );
    } else if (hidden > 0) {
      notes.push(
        `Note: ${hidden} message(s) in ${excludedRoles.join('/')} were excluded; set ${flags} (or mailbox:"trash"/"junk") to include them.`,
      );
    }
    // hidden === 0 -> no note: silence is the trustworthy "nothing matched in Trash/Spam" signal.
  }

  return notes.length ? `\n\n${notes.join('\n')}` : '';
}

// Disclose what get_email_attachments' `raw` mode withheld. That mode returns the JMAP
// attachments array alone, which omits the parts the server routed into the body lists
// instead — and a bare array is indistinguishable from a complete listing, so the
// withheld count is stated rather than left to be noticed (#13).
//
// Returns null when nothing was withheld: silence is the "this is the whole listing"
// signal, the same discipline as the Trash/Spam note. The handler emits this as its own
// content item, never appended to the JSON, which must stay parseable.
export function buildOmittedPartsNote(omittedCount: number): string | null {
  if (!(omittedCount > 0)) return null;
  return `${omittedCount} body-embedded part(s) omitted (raw lists the JMAP attachments array only; omit raw to include them).`;
}

// The whole response body of get_email_attachments, both modes, so the branch is
// exercised by the test suite rather than only by a live call (#13).
//
// The first content item is ALWAYS the JSON array and nothing else, in either mode: a
// caller parses it directly, so the withheld-count note rides as a separate item and is
// never concatenated onto the JSON string.
export function buildAttachmentListContent(
  result: { attachments: any[]; rawAttachments: any[]; omittedFromRaw: number },
  raw: boolean,
): Array<{ type: 'text'; text: string }> {
  if (!raw) {
    return [{ type: 'text', text: JSON.stringify(result.attachments, null, 2) }];
  }
  const content: Array<{ type: 'text'; text: string }> = [
    { type: 'text', text: JSON.stringify(result.rawAttachments, null, 2) },
  ];
  const note = buildOmittedPartsNote(result.omittedFromRaw);
  if (note) content.push({ type: 'text', text: note });
  return content;
}

export function simplifyMailbox(raw: any, options?: { verbose?: boolean }): any {
  const result: any = {
    id: raw.id,
    name: raw.name,
    role: raw.role || undefined,
    parentId: raw.parentId || undefined,
    totalEmails: raw.totalEmails,
    unreadEmails: raw.unreadEmails,
    totalThreads: raw.totalThreads,
    unreadThreads: raw.unreadThreads,
  };
  if (options?.verbose) {
    // Include all remaining mailbox properties
    const coreKeys = new Set(['id', 'name', 'role', 'parentId', 'totalEmails', 'unreadEmails', 'totalThreads', 'unreadThreads']);
    for (const key of Object.keys(raw)) {
      if (!coreKeys.has(key) && raw[key] !== undefined) {
        result[key] = raw[key];
      }
    }
  }
  return result;
}

export function simplifyIdentity(raw: any, options?: { verbose?: boolean }): any {
  const result: any = {
    id: raw.id,
    name: raw.name,
    email: raw.email,
  };
  if (raw.replyTo) result.replyTo = raw.replyTo;
  if (raw.mayDelete != null) result.mayDelete = raw.mayDelete;
  // The identity's configured signature (RFC 8621 section 6). This is where the Fastmail
  // web UI stores the signature it appends for you; JMAP does not append it server-side,
  // so a caller composing through this server has to read it from here and include it in
  // the body deliberately. Surfaced by default rather than behind verbose, because it is
  // the authoritative sign-off and free-handing one from memory drifts from what the user
  // actually configured (#33). An unset or blank signature is omitted, per the
  // omit-empty-fields convention.
  if (typeof raw.textSignature === 'string' && raw.textSignature.trim() !== '') {
    result.textSignature = raw.textSignature;
  }
  if (typeof raw.htmlSignature === 'string' && raw.htmlSignature.trim() !== '') {
    result.htmlSignature = raw.htmlSignature;
  }
  if (options?.verbose) {
    // Include all remaining identity properties. The signature keys are deliberately NOT
    // in coreKeys: verbose still means "everything the server sent", so a blank signature
    // (omitted above) is restored here rather than narrowed away by the default view.
    const coreKeys = new Set(['id', 'name', 'email', 'replyTo', 'mayDelete']);
    for (const key of Object.keys(raw)) {
      if (!coreKeys.has(key) && raw[key] !== undefined) {
        result[key] = raw[key];
      }
    }
  }
  return result;
}

export function simplifyContact(raw: any, options?: { verbose?: boolean }): any {
  const result: any = { id: raw.id };

  // Name - could be in name.full, name.given+surname, or other forms
  if (raw.name) {
    result.name = raw.name.full || [raw.name.given, raw.name.surname].filter(Boolean).join(' ') || undefined;
  }

  // Emails - map from { "label": { address: "..." } } to array of strings
  if (raw.emails && typeof raw.emails === 'object') {
    const emailList = Object.values(raw.emails).map((e: any) => e.address).filter(Boolean);
    if (emailList.length) result.emails = emailList;
  }

  // Phones - map from { "label": { number: "..." } } to array of strings
  if (raw.phones && typeof raw.phones === 'object') {
    const phoneList = Object.values(raw.phones).map((p: any) => p.number).filter(Boolean);
    if (phoneList.length) result.phones = phoneList;
  }

  // Organization
  if (raw.organizations && typeof raw.organizations === 'object') {
    const org = Object.values(raw.organizations)[0] as any;
    if (org?.name) result.organization = org.name;
  }

  // Notes — JMAP ContactCard returns notes as {hash: {note: "text"}} object
  if (raw.notes) {
    if (typeof raw.notes === 'string') {
      result.notes = raw.notes;
    } else if (typeof raw.notes === 'object') {
      const noteTexts = Object.values(raw.notes).map((n: any) => n.note).filter(Boolean);
      if (noteTexts.length) result.notes = noteTexts.join('\n');
    }
  }

  // Verbose: include fields normally dropped, simplified where possible
  if (options?.verbose) {
    // Addresses — flatten to array of address objects (drop hash keys)
    if (raw.addresses && typeof raw.addresses === 'object') {
      const list = Object.values(raw.addresses).filter(Boolean);
      if (list.length) result.addresses = list;
    }
    // Titles — flatten to array of name strings
    if (raw.titles && typeof raw.titles === 'object') {
      const list = Object.values(raw.titles).map((t: any) => t.name).filter(Boolean);
      if (list.length) result.titles = list;
    }
    // Online/URLs — flatten to array of URI strings
    if (raw.online && typeof raw.online === 'object') {
      const list = Object.values(raw.online).map((o: any) => o.uri).filter(Boolean);
      if (list.length) result.online = list;
    }
    if (raw.photos && typeof raw.photos === 'object') {
      result.photos = raw.photos;
    }
    if (raw.anniversaries && typeof raw.anniversaries === 'object') {
      result.anniversaries = raw.anniversaries;
    }
    // Pass through any remaining fields not already handled
    const handledKeys = new Set([
      'id', 'name', 'emails', 'phones', 'organizations', 'notes',
      'addresses', 'titles', 'online', 'photos', 'anniversaries',
    ]);
    for (const key of Object.keys(raw)) {
      if (!handledKeys.has(key) && result[key] === undefined && raw[key] !== undefined) {
        result[key] = raw[key];
      }
    }
  }

  return result;
}

// The contacts listings are not paged (they take no `position`), so they share the
// summary for its always-stated total and never carry a nextPosition instruction their
// callers could not act on.
export function formatContactQueryResult(result: QueryResult, options?: { verbose?: boolean }): string {
  const simplified = result.items.map(c => simplifyContact(c, options));
  return `${formatQuerySummary(result)}\n${JSON.stringify(simplified, null, 2)}`;
}
