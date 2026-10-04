// Every tool that declares `position` must render through the paged summary, which offers
// `nextPosition` while more results remain (#212). The declaring set is read from the BUILT
// server's tools/list, so a tool that gains `position` with no runner below fails here by name.
// The converse (each paged tool declares `position`) is pinned in tool-schema.test.ts.

import { describe, it, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readdirSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createClient } from '../scripts/mcp-harness.mjs';
import { listContactsTool, searchContactsTool } from './contacts-handler.js';
import type { ContactsReadClient } from './contacts-handler.js';
import { listCalendarEventsTool } from './calendar-list-handler.js';
import type { CalendarListClient } from './calendar-list-handler.js';
import { formatEmailQueryResult, formatRawQueryResult } from './response-formatters.js';
import type { QueryResult } from './jmap-client.js';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const SERVER_ENTRY = join(SRC_DIR, '..', 'dist', 'index.js');

// tools/list never reaches Fastmail, so any non-empty value boots the server.
const FAKE_API_VALUE = 'probe-value-not-a-real-credential';

// The number of tools that declare `position`, pinned exact so a dropped or added declaration
// trips it. Change it in the same change that pages or unpages a tool.
const DECLARING_COUNT = 5;

// Each page holds 2 of 10 results, so every page below leaves more to fetch.
const POSITIONS = [0, 4];
const TOTAL = 10;

function emailPage(position: number): QueryResult {
  return { items: [{ id: 'M1' }, { id: 'M2' }], total: TOTAL, position };
}

function contactsClient(position: number): ContactsReadClient {
  const page = (): QueryResult => ({ items: [{ id: 'C1' }, { id: 'C2' }], total: TOTAL, position });
  return { getContacts: async () => page(), searchContacts: async () => page() };
}

function calendarClient(position: number): CalendarListClient {
  return {
    getCalendarEvents: async () => ({
      events: [
        { id: 'a@cal.example', url: '/cal/a.ics', title: 'A', start: '2027-03-02T09:00:00Z' },
        { id: 'b@cal.example', url: '/cal/b.ics', title: 'B', start: '2027-03-03T09:00:00Z' },
      ],
      total: TOTAL,
      position,
    }),
  };
}

type Runner = (position: number) => Promise<string>;

const emailRunners: Record<string, Runner> = {
  simplified: async (position) => formatEmailQueryResult(emailPage(position)),
  raw: async (position) => formatRawQueryResult(emailPage(position)),
};

const contactRunners = (run: typeof listContactsTool): Record<string, Runner> => ({
  simplified: async (position) => (await run({ query: 'q', position }, 2, contactsClient(position)))[0].text,
  raw: async (position) => (await run({ query: 'q', position, raw: true }, 2, contactsClient(position)))[0].text,
});

const RUNNERS: Record<string, Record<string, Runner>> = {
  // The email tools page inline in the CallTool switch, so these run the two renderers that
  // switch calls; that the switch calls them is not checked here.
  list_emails: emailRunners,
  search_emails: emailRunners,
  list_contacts: contactRunners(listContactsTool),
  search_contacts: contactRunners(searchContactsTool),
  list_calendar_events: {
    default: async (position) => (await listCalendarEventsTool({ position }, 2, calendarClient(position)))[0].text,
  },
};

describe('every tool that declares position offers nextPosition (#212)', () => {
  const home = mkdtempSync(join(tmpdir(), 'fastmail-mcp-paged-tools-'));
  after(() => rmSync(home, { recursive: true, force: true, maxRetries: 3 }));

  let declaring: string[] = [];
  let bootError: unknown;

  // A throwing before() reports the suite CANCELLED, not failed, so the boot error is stored
  // and asserted in the first test.
  before(async () => {
    try {
      const built = statSync(SERVER_ENTRY).mtimeMs;
      const newest = readdirSync(SRC_DIR)
        .filter((f) => f.endsWith('.ts') && !f.endsWith('.test.ts'))
        .map((f) => ({ f, m: statSync(join(SRC_DIR, f)).mtimeMs }))
        .reduce((a, b) => (b.m > a.m ? b : a));
      assert.ok(built >= newest.m, `dist/index.js is older than src/${newest.f}. Run "npm run build".`);

      const env: Record<string, string> = {};
      for (const [k, v] of Object.entries(process.env)) {
        if (v !== undefined && !/fastmail/i.test(k)) env[k] = v;
      }
      env.HOME = home;
      env.USERPROFILE = home;
      env.FASTMAIL_API_TOKEN = FAKE_API_VALUE;

      const client = createClient({ env });
      try {
        await client.init();
        const result: any = await client.list();
        declaring = result.tools
          .filter((t: any) => Object.hasOwn(t.inputSchema?.properties ?? {}, 'position'))
          .map((t: any) => t.name);
      } finally {
        client.close();
      }
    } catch (err) {
      bootError = err;
    }
  });

  it(`reads ${DECLARING_COUNT} position-declaring tools from the built server`, () => {
    assert.equal(bootError, undefined, `the built server failed to boot or answer tools/list: ${bootError}`);
    assert.equal(
      declaring.length,
      DECLARING_COUNT,
      `found ${declaring.length} tools declaring position (expected ${DECLARING_COUNT}): ${declaring.join(', ')}`,
    );
  });

  it('has a runner for every declaring tool, and each offers nextPosition', async () => {
    const failures: string[] = [];
    for (const tool of declaring) {
      const runners = RUNNERS[tool];
      if (!runners) {
        failures.push(`${tool}: declares position but has no runner in paged-tools.test.ts`);
        continue;
      }
      for (const [label, run] of Object.entries(runners)) {
        for (const position of POSITIONS) {
          const text = await run(position);
          if (!text.includes(`nextPosition: ${position + 2}`)) {
            failures.push(`${tool} (${label}, position ${position}): no nextPosition in "${text.split('\n')[0]}"`);
          }
        }
      }
    }
    assert.deepEqual(failures, []);
  });

  it('has no runner for a tool that no longer declares position', () => {
    assert.deepEqual(Object.keys(RUNNERS).filter((tool) => !declaring.includes(tool)), []);
  });
});
