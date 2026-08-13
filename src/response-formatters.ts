import { simplifyEmail } from './email-formatter.js';
import type { QueryResult, ReplacedDraftInfo, UpdateDraftResult } from './jmap-client.js';

export function formatQueryResult(result: QueryResult): string {
  const { items, total } = result;
  const summary = total != null && total > items.length
    ? `Showing ${items.length} of ${total} results.`
    : total != null
      ? `${total} results.`
      : `${items.length} results.`;
  return `${summary}\n${JSON.stringify(items, null, 2)}`;
}

export function formatEmailQueryResult(result: QueryResult): string {
  const { items, total } = result;
  const simplified = items.map(e => simplifyEmail(e));
  const summary = total != null && total > items.length
    ? `Showing ${items.length} of ${total} results.`
    : total != null
      ? `${total} results.`
      : `${items.length} results.`;
  return `${summary}\n${JSON.stringify(simplified, null, 2)}`;
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

export function formatContactQueryResult(result: QueryResult, options?: { verbose?: boolean }): string {
  const { items, total } = result;
  const simplified = items.map(c => simplifyContact(c, options));
  const summary = total != null && total > items.length
    ? `Showing ${items.length} of ${total} results.`
    : total != null
      ? `${total} results.`
      : `${items.length} results.`;
  return `${summary}\n${JSON.stringify(simplified, null, 2)}`;
}
