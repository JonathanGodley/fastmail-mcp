// Experimental snooze tools (#35). Registered only when FASTMAIL_EXPERIMENTAL_SNOOZE=1.
//
// Fastmail refuses the snooze property to an API token at every capability name, and serves
// it under https://www.fastmail.com/dev/mail to its own web client's session. So these tools
// talk JMAP on that session, not on JmapClient's API token: the caller supplies the web
// login's session cookie, and the bearer token is fetched from it the way the web client does.
// Measured by scripts/probes/snooze-read.probe.mjs and snooze-write.probe.mjs.

import { coerceUtcDate, describeUntrusted, InvalidInputError, redactedJson, registerSecret } from './coerce.js';
import { getDefaultTimezone } from './email-formatter.js';
import type { ToolContent } from './mailbox-handler.js';

export const SNOOZE_CAPABILITY = 'https://www.fastmail.com/dev/mail';
const USING = ['urn:ietf:params:jmap:core', 'urn:ietf:params:jmap:mail', SNOOZE_CAPABILITY];

// What Fastmail's own client writes when it snoozes: wake into Inbox, unread and new.
export const WAKE_KEYWORDS = { $new: true, $seen: false };

export const WEB_SESSION_EXPIRED =
  'Fastmail did not accept FASTMAIL_WEB_COOKIE: the browser login it was copied from has ended or been logged out. ' +
  'Log in to Fastmail in a browser and copy the __Http-s_… cookie for api.fastmail.com again.';

export class WebSessionError extends Error {}

/** One JMAP request on the web session: method calls in, methodResponses out. */
export interface SnoozeTransport {
  call(methodCalls: any[]): Promise<any[]>;
  accountId(): Promise<string>;
}

interface WebSession { token: string; apiUrl: string; accountId: string }

// The token and the API URL both come back from Fastmail; refuse an API URL off Fastmail's
// own HTTPS hosts, because the cookie and token ride on every request to it.
function isFastmailHttps(url: string): boolean {
  try {
    const u = new URL(url);
    return u.protocol === 'https:' && (u.hostname === 'fastmail.com' || u.hostname.endsWith('.fastmail.com'));
  } catch {
    return false;
  }
}

/**
 * JMAP over Fastmail's web session. The cookie alone fetches the bearer token from
 * /auth/sessions (what the web client does on load and hourly), but the token is honoured
 * only alongside the cookie, so both ride on every call. A 401 refetches the token once.
 */
export class WebSessionTransport implements SnoozeTransport {
  private session: WebSession | null = null;

  constructor(
    private readonly cookie: string,
    private readonly fetchImpl: typeof fetch = fetch,
    private readonly origin = 'https://api.fastmail.com',
  ) {}

  private async open(): Promise<WebSession> {
    const res = await this.fetchImpl(`${this.origin}/auth/sessions`, {
      headers: { Accept: 'application/json', Cookie: this.cookie },
      redirect: 'error',
    });
    if (res.status === 401 || res.status === 403) throw new WebSessionError(WEB_SESSION_EXPIRED);
    if (!res.ok) throw new Error(`Fastmail web session lookup failed: HTTP ${res.status}`);
    const body: any = await res.json().catch(() => null);
    const record = Array.isArray(body) ? body.find((s: any) => typeof s?.accessToken === 'string' && s.accessToken) : null;
    // A cookie that is not logged in still gets 200, with no session carrying a token.
    if (!record) throw new WebSessionError(WEB_SESSION_EXPIRED);
    const accountId = record.primaryAccounts?.['urn:ietf:params:jmap:mail'];
    if (typeof record.apiUrl !== 'string' || !isFastmailHttps(record.apiUrl) || typeof accountId !== 'string') {
      throw new Error('Fastmail web session record has no usable apiUrl or mail account');
    }
    registerSecret(record.accessToken);
    return { token: record.accessToken, apiUrl: record.apiUrl, accountId };
  }

  async accountId(): Promise<string> {
    this.session ??= await this.open();
    return this.session.accountId;
  }

  async call(methodCalls: any[]): Promise<any[]> {
    for (let attempt = 0; attempt < 2; attempt++) {
      if (!this.session || attempt === 1) this.session = await this.open();
      const res = await this.fetchImpl(this.session.apiUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${this.session.token}`,
          Cookie: this.cookie,
          'Content-Type': 'application/json',
          Accept: 'application/json',
        },
        body: JSON.stringify({ using: USING, methodCalls }),
        redirect: 'error',
      });
      if (res.status === 401) continue;
      if (!res.ok) throw new Error(`JMAP request on the web session failed: HTTP ${res.status}`);
      const data: any = await res.json().catch(() => null);
      if (!data || !Array.isArray(data.methodResponses)) throw new Error('Invalid JMAP response: missing methodResponses');
      return data.methodResponses;
    }
    throw new WebSessionError(WEB_SESSION_EXPIRED);
  }
}

function result(responses: any[], index: number, method: string): any {
  const entry = responses[index];
  if (!Array.isArray(entry)) throw new Error(`JMAP response missing ${method}`);
  if (entry[0] === 'error') {
    throw new Error(`JMAP error in ${method}: ${describeUntrusted(entry[1]?.type)}${entry[1]?.description ? ' - ' + describeUntrusted(entry[1].description) : ''}`);
  }
  return entry[1];
}

interface Mailboxes { inbox: any; snoozed: any; byId: Map<string, any> }

async function mailboxes(t: SnoozeTransport, accountId: string): Promise<Mailboxes> {
  const list = result(await t.call([['Mailbox/get', { accountId, properties: ['id', 'name', 'role'] }, 'm']]), 0, 'Mailbox/get').list ?? [];
  const inbox = list.find((m: any) => m?.role === 'inbox');
  const snoozed = list.find((m: any) => m?.role === 'snoozed');
  if (!inbox || !snoozed) throw new Error('This account has no Inbox or no Snoozed mailbox to snooze between');
  return { inbox, snoozed, byId: new Map(list.map((m: any) => [m.id, m])) };
}

async function getOne(t: SnoozeTransport, accountId: string, emailId: string): Promise<any> {
  const got = result(await t.call([['Email/get', { accountId, ids: [emailId], properties: ['id', 'mailboxIds', 'snoozed'] }, 'g']]), 0, 'Email/get');
  const email = got.list?.[0];
  if (!email) throw new InvalidInputError(`Email not found: ${describeUntrusted(emailId)}`);
  return email;
}

async function update(t: SnoozeTransport, accountId: string, emailId: string, patch: Record<string, unknown>): Promise<void> {
  const set = result(await t.call([['Email/set', { accountId, update: { [emailId]: patch } }, 'u']]), 0, 'Email/set');
  const err = set.notUpdated?.[emailId];
  if (err || !set.updated || !(emailId in set.updated)) {
    throw new Error(`Fastmail refused the update: ${describeUntrusted(err?.type ?? 'not updated')}${err?.description ? ' - ' + describeUntrusted(err.description) : ''}`);
  }
}

/**
 * The wake time as JMAP's UTCDate, read as search_emails reads `after`/`before`: a value with
 * no offset is local time in the configured zone. It must be in the future.
 */
export function parseWakeTime(value: unknown, now: Date = new Date(), zone: string | undefined = getDefaultTimezone()): string {
  if (value === undefined || value === null || value === '') {
    throw new InvalidInputError('until is required, e.g. 2026-11-03T09:00:00+11:00');
  }
  const until = coerceUtcDate(value, 'until', zone)!;
  if (Date.parse(until) <= now.getTime()) {
    throw new InvalidInputError(`until is not in the future: ${describeUntrusted(String(value))} is ${until}`);
  }
  return until;
}

/** list_snoozed_emails: the Snoozed mailbox, soonest wake first. */
export async function listSnoozedEmails(args: any, t: SnoozeTransport, limit: number): Promise<ToolContent> {
  const accountId = await t.accountId();
  const mb = await mailboxes(t, accountId);
  const responses = await t.call([
    ['Email/query', { accountId, filter: { inMailbox: mb.snoozed.id }, sort: [{ property: 'snoozedUntil', mailboxId: mb.snoozed.id, isAscending: true }], limit, calculateTotal: true }, 'q'],
    ['Email/get', { accountId, '#ids': { resultOf: 'q', name: 'Email/query', path: '/ids' }, properties: ['id', 'threadId', 'subject', 'from', 'receivedAt', 'snoozed'] }, 'g'],
  ]);
  const query = result(responses, 0, 'Email/query');
  const list = result(responses, 1, 'Email/get').list ?? [];
  const items = list.map((e: any) => {
    const dest = e.snoozed?.moveToMailboxId;
    return {
      id: e.id,
      threadId: e.threadId,
      subject: e.subject,
      from: e.from?.[0]?.email,
      receivedAt: e.receivedAt,
      wakeAt: e.snoozed?.until ?? null,
      wakeInto: dest ? (mb.byId.get(dest)?.name ?? dest) : mb.inbox.name,
    };
  });
  return [{ type: 'text', text: redactedJson({ total: query.total ?? items.length, emails: items }) }];
}

/**
 * snooze_email. From the Inbox it moves the message to Snoozed with the client's own shape;
 * on a message already snoozed it changes the wake time only. Other labels are kept.
 */
export async function snoozeEmail(args: any, t: SnoozeTransport, now: Date = new Date()): Promise<ToolContent> {
  const emailId = typeof args?.emailId === 'string' ? args.emailId.trim() : '';
  if (!emailId) throw new InvalidInputError('emailId is required');
  const until = parseWakeTime(args?.until, now);
  const accountId = await t.accountId();
  const mb = await mailboxes(t, accountId);
  const email = await getOne(t, accountId, emailId);
  const inSnoozed = email.mailboxIds?.[mb.snoozed.id] === true;
  if (inSnoozed) {
    await update(t, accountId, emailId, email.snoozed ? { 'snoozed/until': until } : { snoozed: { until, setKeywords: WAKE_KEYWORDS } });
    return [{ type: 'text', text: redactedJson({ emailId, wakeAt: until, wakeInto: mb.inbox.name, previousWakeAt: email.snoozed?.until ?? null }) }];
  }
  if (email.mailboxIds?.[mb.inbox.id] !== true) {
    throw new InvalidInputError('snooze_email snoozes a message from the Inbox only; this one is not in the Inbox. Move it to the Inbox first.');
  }
  await update(t, accountId, emailId, {
    [`mailboxIds/${mb.inbox.id}`]: null,
    [`mailboxIds/${mb.snoozed.id}`]: true,
    snoozed: { until, setKeywords: WAKE_KEYWORDS },
  });
  return [{ type: 'text', text: redactedJson({ emailId, wakeAt: until, wakeInto: mb.inbox.name }) }];
}

/** unsnooze_email: back to the Inbox now, snooze cleared. */
export async function unsnoozeEmail(args: any, t: SnoozeTransport): Promise<ToolContent> {
  const emailId = typeof args?.emailId === 'string' ? args.emailId.trim() : '';
  if (!emailId) throw new InvalidInputError('emailId is required');
  const accountId = await t.accountId();
  const mb = await mailboxes(t, accountId);
  const email = await getOne(t, accountId, emailId);
  if (email.mailboxIds?.[mb.snoozed.id] !== true && !email.snoozed) {
    throw new InvalidInputError('This message is not snoozed.');
  }
  await update(t, accountId, emailId, {
    [`mailboxIds/${mb.snoozed.id}`]: null,
    [`mailboxIds/${mb.inbox.id}`]: true,
    snoozed: null,
  });
  return [{ type: 'text', text: redactedJson({ emailId, movedTo: mb.inbox.name, wasWakeAt: email.snoozed?.until ?? null }) }];
}

const EXPERIMENTAL_NOTE =
  ' EXPERIMENTAL: runs on the Fastmail web login\'s session cookie (FASTMAIL_WEB_COOKIE), not the API token, because Fastmail refuses snooze to API tokens. It stops working when that browser login ends.';

export const SNOOZE_TOOLS = [
  {
    name: 'list_snoozed_emails',
    description: 'List the messages in the Snoozed mailbox, soonest wake time first. Each item carries wakeAt (UTC, when Fastmail returns it to wakeInto, normally the Inbox, marked unread). Returns { total, emails }.' + EXPERIMENTAL_NOTE,
    inputSchema: {
      type: 'object',
      properties: {
        limit: { type: ['number', 'string'], description: 'Maximum messages to return (default 50, max 100)', default: 50 },
      },
    },
  },
  {
    name: 'snooze_email',
    description: 'Snooze a message until a time: it leaves the Inbox for Snoozed and Fastmail returns it to the Inbox, unread, at that time, as the Fastmail app does. Other labels are kept. Only an Inbox message can be snoozed; on a message already snoozed this changes the wake time. Returns { emailId, wakeAt, wakeInto }, wakeAt in UTC.' + EXPERIMENTAL_NOTE,
    inputSchema: {
      type: 'object',
      properties: {
        emailId: { type: 'string', description: 'ID of the email to snooze' },
        until: { type: 'string', description: 'When it comes back, in the future: a datetime such as 2026-11-03T09:00:00+11:00 or 2026-11-02T22:00:00Z (taken exactly as written), or 2026-11-03T09:00:00 with no offset (read as local time in the configured timezone, FASTMAIL_TIMEZONE or the host\'s). A date alone means midnight at its start.' },
      },
      required: ['emailId', 'until'],
    },
  },
  {
    name: 'unsnooze_email',
    description: 'Return a snoozed message to the Inbox now and clear its snooze. Its read state is left as it is. Returns { emailId, movedTo, wasWakeAt }.' + EXPERIMENTAL_NOTE,
    inputSchema: {
      type: 'object',
      properties: {
        emailId: { type: 'string', description: 'ID of the snoozed email' },
      },
      required: ['emailId'],
    },
  },
];
