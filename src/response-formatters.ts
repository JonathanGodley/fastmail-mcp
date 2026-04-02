import { simplifyEmail } from './email-formatter.js';
import type { QueryResult } from './jmap-client.js';

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
  if (options?.verbose) {
    // Include all remaining identity properties
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
