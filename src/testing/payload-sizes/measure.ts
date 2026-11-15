// The payload-size baseline (#156): five default reads, produced by the real handlers and
// clients from the synthetic server data in fixtures.ts, measured against baseline.json.
// Run through scripts/payload-sizes.mts. A size change is reported, never failed on.

import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { FastmailAuth } from '../../auth.js';
import { CalDAVCalendarClient } from '../../caldav-client.js';
import { listCalendarEventsTool } from '../../calendar-list-handler.js';
import { ContactsCalendarClient } from '../../contacts-calendar.js';
import { listContactsTool } from '../../contacts-handler.js';
import { getEmailTool } from '../../email-get-handler.js';
import { getDefaultTimezone, setDefaultTimezone } from '../../email-formatter.js';
import { listEmailsTool } from '../../email-list-handler.js';
import { JmapClient } from '../../jmap-client.js';
import type { JmapRequest, JmapResponse } from '../../jmap-client.js';
import { readThread } from '../../thread-handler.js';
import {
  CALENDARS, EXPANDED_OBJECTS, FIXTURE_NOW, FIXTURE_ZONE, GET_EMAIL_ID, GET_EMAIL_RESPONSES,
  GET_THREAD_RESPONSES, LIST_CONTACTS_RESPONSES, LIST_EMAILS_RESPONSES, MAILBOXES, SESSION,
  STORED_OBJECTS, THREAD_ID,
} from './fixtures.js';

export const BASELINE_PATH = fileURLToPath(new URL('./baseline.json', import.meta.url));

export interface Figures {
  bytes: number;
  chars: number;
  sha256: string;
}

export type Baseline = Record<string, Figures>;

export type Produced = { name: string; text: string } | { name: string; error: string };

// Answers each call in a batch from the fixture entry with the same call id. A call the
// fixture has no answer for gets a method-level error in its place (RFC 8620 section 3.6.2),
// so a required response that goes missing fails the row instead of shifting the others.
function answer(request: JmapRequest, responses: Record<string, [string, any]>): JmapResponse {
  return {
    sessionState: 'fixture',
    methodResponses: request.methodCalls.map(([method, , callId]): [string, any, string] => {
      const entry = responses[callId];
      if (!entry || entry[0] !== method) {
        return ['error', { type: 'serverFail', description: `the fixture has no ${method} response for call "${callId}"` }, callId];
      }
      return [method, structuredClone(entry[1]), callId];
    }),
  };
}

function stubJmap<C extends JmapClient>(client: C, responses: Record<string, [string, any]>): C {
  client.getSession = async () => structuredClone(SESSION);
  client.getMailboxes = async () => structuredClone(MAILBOXES);
  client.makeRequest = async (request) => answer(request, responses);
  return client;
}

const auth = () => new FastmailAuth({ apiToken: 'fixture-token' });
const textOf = (content: Array<{ text: string }>) => content.map((c) => c.text).join('');

function calendarClient(): CalDAVCalendarClient {
  const client = new CalDAVCalendarClient({ username: 'fixture', password: 'fixture', now: () => FIXTURE_NOW });
  const toData = (lines: string[]) => lines.join('\r\n');
  const dav = {
    login: async () => {},
    fetchCalendars: async () => structuredClone(CALENDARS),
    fetchCalendarObjects: async ({ calendar }: { calendar: { url: string } }) =>
      EXPANDED_OBJECTS.filter((o) => o.calendarUrl === calendar.url)
        .map((o) => ({ url: o.url, etag: `"${o.url.length}"`, data: toData(o.lines) })),
    calendarMultiGet: async ({ objectUrls }: { objectUrls?: string[] }) =>
      (objectUrls ?? []).flatMap((href) => {
        const url = Object.keys(STORED_OBJECTS).find((u) => new URL(u).pathname === href);
        return url ? [{ href, props: { getetag: '"stored"', calendarData: { _cdata: toData(STORED_OBJECTS[url]) } } }] : [];
      }),
  };
  (client as any).client = dav;
  return client;
}

// Fixed order: the table, the baseline file and the smoke test all follow it.
const PAYLOADS: Array<{ name: string; produce: () => Promise<string> }> = [
  {
    name: 'list_emails',
    produce: async () => textOf(await listEmailsTool({}, 20, stubJmap(new JmapClient(auth()), LIST_EMAILS_RESPONSES))),
  },
  {
    name: 'get_email',
    produce: async () => textOf(await getEmailTool({ emailId: GET_EMAIL_ID }, stubJmap(new JmapClient(auth()), GET_EMAIL_RESPONSES))),
  },
  {
    name: 'get_thread',
    produce: () => readThread({ threadId: THREAD_ID }, stubJmap(new JmapClient(auth()), GET_THREAD_RESPONSES)),
  },
  {
    name: 'list_calendar_events',
    produce: async () => textOf(await listCalendarEventsTool({}, 50, calendarClient())),
  },
  {
    name: 'list_contacts',
    produce: async () => textOf(await listContactsTool({}, 20, stubJmap(new ContactsCalendarClient(auth()), LIST_CONTACTS_RESPONSES))),
  },
];

export const PAYLOAD_NAMES = PAYLOADS.map((p) => p.name);

export async function producePayloads(): Promise<Produced[]> {
  const savedZone = getDefaultTimezone();
  setDefaultTimezone(FIXTURE_ZONE);
  try {
    const produced: Produced[] = [];
    for (const { name, produce } of PAYLOADS) {
      try {
        produced.push({ name, text: await produce() });
      } catch (err) {
        produced.push({ name, error: err instanceof Error ? err.message : String(err) });
      }
    }
    return produced;
  } finally {
    setDefaultTimezone(savedZone);
  }
}

export function measure(text: string): Figures {
  return {
    bytes: Buffer.byteLength(text, 'utf8'),
    chars: text.length,
    sha256: createHash('sha256').update(text, 'utf8').digest('hex'),
  };
}

export function readBaseline(path: string = BASELINE_PATH): Baseline {
  return JSON.parse(readFileSync(path, 'utf8')) as Baseline;
}

export function formatBaseline(produced: Produced[]): string {
  const out: Baseline = {};
  for (const p of produced) {
    if ('text' in p) {
      const { bytes, chars, sha256 } = measure(p.text);
      out[p.name] = { bytes, chars, sha256 };
    }
  }
  return `${JSON.stringify(out, null, 2)}\n`;
}

const signed = (n: number) => (n > 0 ? `+${n}` : String(n));
const percent = (delta: number, base: number) => (base === 0 ? 'n/a' : `${delta > 0 ? '+' : ''}${((delta / base) * 100).toFixed(1)}%`);

export function formatTable(baseline: Baseline, produced: Produced[]): string {
  const header = ['payload', 'bytes base', 'bytes now', 'delta', 'delta %', 'chars base', 'chars now', 'delta', 'digest'];
  const rows: string[][] = [];
  let baseBytes = 0;
  let nowBytes = 0;
  let baseChars = 0;
  let nowChars = 0;
  for (const p of produced) {
    const base = baseline[p.name];
    if (!('text' in p)) {
      rows.push([p.name, base ? String(base.bytes) : '-', 'FAILED', '', '', base ? String(base.chars) : '-', '', '', p.error]);
      continue;
    }
    const now = measure(p.text);
    nowBytes += now.bytes;
    nowChars += now.chars;
    if (!base) {
      rows.push([p.name, '-', String(now.bytes), '', '', '-', String(now.chars), '', 'new']);
      continue;
    }
    baseBytes += base.bytes;
    baseChars += base.chars;
    const digest = base.sha256 === now.sha256
      ? 'match'
      : base.bytes === now.bytes ? 'DIFFERS at equal size (content changed, size did not)' : 'differs';
    rows.push([
      p.name, String(base.bytes), String(now.bytes), signed(now.bytes - base.bytes), percent(now.bytes - base.bytes, base.bytes),
      String(base.chars), String(now.chars), signed(now.chars - base.chars), digest,
    ]);
  }
  for (const name of Object.keys(baseline)) {
    if (!produced.some((p) => p.name === name)) {
      const base = baseline[name];
      baseBytes += base.bytes;
      baseChars += base.chars;
      rows.push([name, String(base.bytes), '-', '', '', String(base.chars), '-', '', 'removed']);
    }
  }
  rows.push([
    'total', String(baseBytes), String(nowBytes), signed(nowBytes - baseBytes), percent(nowBytes - baseBytes, baseBytes),
    String(baseChars), String(nowChars), signed(nowChars - baseChars), '',
  ]);
  const widths = header.map((h, i) => Math.max(h.length, ...rows.map((r) => r[i].length)));
  const line = (cells: string[]) => cells.map((c, i) => (i === 0 || i === cells.length - 1 ? c.padEnd(widths[i]) : c.padStart(widths[i]))).join('  ').trimEnd();
  return [line(header), ...rows.map(line)].join('\n');
}

/** The command: prints the table, or with `--update` rewrites the baseline. Returns the exit code. */
export async function runPayloadSizes(argv: string[], log: (line: string) => void = console.log): Promise<number> {
  const update = argv.includes('--update');
  const produced = await producePayloads();
  const failed = produced.filter((p) => !('text' in p));

  let baseline: Baseline = {};
  try {
    baseline = readBaseline();
  } catch (err) {
    if (!update) {
      log(`Cannot read the baseline at ${BASELINE_PATH}: ${err instanceof Error ? err.message : String(err)}`);
      return 1;
    }
  }

  log(formatTable(baseline, produced));
  if (failed.length > 0) {
    log(`\nCould not produce: ${failed.map((p) => p.name).join(', ')}. A fixture no longer satisfies its client or handler.`);
    return 1;
  }
  if (update) {
    writeFileSync(BASELINE_PATH, formatBaseline(produced));
    log(`\nBaseline written to ${BASELINE_PATH}.`);
  }
  return 0;
}
