// A fake CalDAV server that answers tsdav's real requests with raw multistatus XML, so a test
// drives tsdav's own parse instead of stubbing what it returns. Test-only (see mock-calls.ts).
//
// Every value here is synthetic: hosts under `.invalid`, invented UIDs and paths.

import { mock } from 'node:test';
import { calendarMultiGet, fetchCalendarObjects } from 'tsdav';
import { makeMockDAVClient } from './caldav-mock.js';

/**
 * How one resource answers a read that carries its data:
 * - `ok`: etag and calendar-data in a 200 propstat;
 * - `nodata`: calendar-data in a 403 propstat, which tsdav drops;
 * - `emptydata`: an empty calendar-data element in the 200 propstat;
 * - `noetag`: getetag in a 404 propstat;
 * - `raw`: `rawData` as bare text in calendar-data, so tsdav applies its own type coercion;
 * - `omit`: no response for the resource at all;
 * - `404`: a response holding only a 404 status (the HTTP status of the answer stays 207).
 */
export type FakeAnswerMode = 'ok' | 'nodata' | 'emptydata' | 'noetag' | 'raw' | 'omit' | '404';

export interface FakeResource {
  /** Absolute url. */
  href: string;
  /** The stored iCalendar, which a UID query matches against. */
  data: string;
  etag?: string;
  /** What a time-range query with expand returns as this resource's data; `data` when absent. */
  expandedData?: string;
  /** How it answers a read that a query started (the multiget after a UID query, or an expand query). */
  query?: FakeAnswerMode;
  /** How it answers a multiget for its own href that no query started. */
  fetch?: FakeAnswerMode;
  rawData?: string;
}

export interface FakeRequest {
  /** The request url, which for a REPORT is the calendar's url. */
  url: string;
  kind: 'query' | 'multiget';
  /** The UID text-match of a UID query. */
  uid?: string;
  /** Whether a query asked for calendar-data expanded over a time range. */
  expand: boolean;
  /** The hrefs a multiget asked for, as sent. */
  hrefs: string[];
}

/**
 * Replaces the answer to one request. `body` replaces the multistatus; `status` the HTTP status
 * (207 otherwise); `contentType` the header (`null` sends none).
 */
export interface FakeOverride {
  body?: string;
  status?: number;
  contentType?: string | null;
}

export interface FakeCalDAVServer {
  fetch: typeof globalThis.fetch;
  /** Every request the server received, in order. */
  requests: FakeRequest[];
}

const NULL_BODY_STATUSES = new Set([101, 103, 204, 205, 304]);

function storedUid(data: string): string {
  const unfolded = data.replace(/\r?\n[ \t]/g, '');
  return /^UID:(.*)$/mi.exec(unfolded)?.[1]?.trim() ?? '';
}

function xmlEscape(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function xmlUnescape(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

export function multistatus(inner: string): string {
  return '<?xml version="1.0" encoding="utf-8"?>'
    + `<d:multistatus xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">${inner}</d:multistatus>`;
}

/** One `<d:response>` carrying `props` in a propstat of `status`, plus any further propstat. */
export function resp(href: string, props: string, status = '200 OK', extraPropstat = ''): string {
  return `<d:response><d:href>${href}</d:href><d:propstat><d:prop>${props}</d:prop>`
    + `<d:status>HTTP/1.1 ${status}</d:status></d:propstat>${extraPropstat}</d:response>`;
}

export const etagProp = (etag: string) => `<d:getetag>${xmlEscape(etag)}</d:getetag>`;
export const dataProp = (data: string) => `<c:calendar-data><![CDATA[${data}]]></c:calendar-data>`;

function answerWithData(path: string, o: FakeResource, mode: FakeAnswerMode, data: string): string {
  if (mode === 'omit') return '';
  if (mode === '404') return `<d:response><d:href>${path}</d:href><d:status>HTTP/1.1 404 Not Found</d:status></d:response>`;
  const etagOk = mode !== 'noetag' && o.etag !== undefined;
  const ok: string[] = [];
  const bad: string[] = [];
  if (etagOk) ok.push(etagProp(o.etag!));
  else bad.push('<d:getetag/>');
  if (mode === 'nodata') bad.push('<c:calendar-data/>');
  else if (mode === 'emptydata') ok.push('<c:calendar-data/>');
  else if (mode === 'raw') ok.push(`<c:calendar-data>${xmlEscape(o.rawData ?? '')}</c:calendar-data>`);
  else ok.push(dataProp(data));
  let r = `<d:response><d:href>${path}</d:href>`;
  if (ok.length) r += `<d:propstat><d:prop>${ok.join('')}</d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat>`;
  if (bad.length) {
    r += `<d:propstat><d:prop>${bad.join('')}</d:prop>`
      + `<d:status>HTTP/1.1 ${mode === 'nodata' ? '403 Forbidden' : '404 Not Found'}</d:status></d:propstat>`;
  }
  return r + '</d:response>';
}

/**
 * A fake server over `store`, keyed by calendar url. A UID query answers each resource whose
 * stored UID matches case-insensitively, as Fastmail does; a query without a UID answers every
 * resource in the calendar whatever its time range, with data when it asked for expand. A
 * multiget answers each href asked for, or a 404 response for an unknown one.
 *
 * `override(request, n)`, where `n` counts requests from 0, replaces the answer to that request.
 */
export function makeFakeCalDAVServer(
  store: Record<string, FakeResource[]>,
  override?: (request: FakeRequest, n: number) => FakeOverride | undefined,
): FakeCalDAVServer {
  const requests: FakeRequest[] = [];
  // Resources the last UID query matched, so the multiget it starts answers in `query` mode.
  const viaQuery = new Set<FakeResource>();
  const path = (href: string) => new URL(href).pathname;

  const fetch = async (input: string | URL | Request, init?: RequestInit): Promise<Response> => {
    const url = String(input instanceof Request ? input.url : input);
    const body = String(init?.body ?? '');
    const objects = store[url] ?? [];
    const out: string[] = [];
    let request: FakeRequest;
    if (body.includes('calendar-query')) {
      const match = /<c:text-match[^>]*>([\s\S]*?)<\/c:text-match>/.exec(body);
      const uid = match ? xmlUnescape(match[1]) : undefined;
      const expand = body.includes('expand');
      request = { url, kind: 'query', uid, expand, hrefs: [] };
      for (const o of objects) {
        if (uid !== undefined) {
          if (storedUid(o.data).toLowerCase() !== uid.toLowerCase()) continue;
          out.push(resp(xmlEscape(path(o.href)), etagProp(o.etag ?? '"e"')));
          viaQuery.add(o);
        } else if (expand) {
          out.push(answerWithData(xmlEscape(path(o.href)), o, o.query ?? 'ok', o.expandedData ?? o.data));
        } else {
          out.push(resp(xmlEscape(path(o.href)), etagProp(o.etag ?? '"e"')));
        }
      }
    } else if (body.includes('calendar-multiget')) {
      const hrefs = [...body.matchAll(/<d:href>([\s\S]*?)<\/d:href>/g)].map(m => m[1]);
      request = { url, kind: 'multiget', expand: false, hrefs };
      for (const h of hrefs) {
        const o = objects.find(x => xmlEscape(path(x.href)) === h);
        if (!o) {
          out.push(`<d:response><d:href>${h}</d:href><d:status>HTTP/1.1 404 Not Found</d:status></d:response>`);
          continue;
        }
        const mode = (viaQuery.has(o) ? o.query : o.fetch) ?? 'ok';
        viaQuery.delete(o);
        out.push(answerWithData(h, o, mode, o.data));
      }
    } else {
      throw new Error(`fake CalDAV server: unexpected request to ${url}`);
    }
    const n = requests.length;
    requests.push(request);
    const replaced = override?.(request, n);
    const status = replaced?.status ?? 207;
    const contentType = replaced && 'contentType' in replaced ? replaced.contentType : 'application/xml; charset=utf-8';
    const text = replaced?.body ?? multistatus(out.join(''));
    const response = new Response(NULL_BODY_STATUSES.has(status) ? null : text, {
      status,
      headers: contentType === null || contentType === undefined ? {} : { 'content-type': contentType },
    });
    // A constructed Response has an empty url; a fetched one carries the request's.
    Object.defineProperty(response, 'url', { value: url });
    return response;
  };
  return { fetch: fetch as typeof globalThis.fetch, requests };
}

/**
 * A DAVClient stand-in whose reads are tsdav's real `fetchCalendarObjects` and
 * `calendarMultiGet` against `server`. The server is the client's `fetchOverride`, and each read
 * merges the caller's params over it exactly as DAVClient's own methods do, so a `fetch` the
 * production code passes per call still wins.
 */
export function makeFakeServerDAVClient<Calendar, Rest extends object & { login?: never; fetchCalendars?: never }>(
  calendars: Calendar[],
  server: FakeCalDAVServer,
  rest: Rest,
) {
  const fetchOverride = server.fetch;
  return makeMockDAVClient(calendars, {
    fetchOverride,
    fetchCalendarObjects: mock.fn(async (p: Parameters<typeof fetchCalendarObjects>[0]) =>
      fetchCalendarObjects({ fetch: fetchOverride, ...p })),
    calendarMultiGet: mock.fn(async (p: Parameters<typeof calendarMultiGet>[0]) =>
      calendarMultiGet({ fetch: fetchOverride, ...p })),
    ...rest,
  });
}
