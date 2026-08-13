import { coerceBool, InvalidInputError } from './coerce.js';
import { simplifyEmail } from './email-formatter.js';
import type { SimplifiedEmail } from './email-formatter.js';
import { parseEmailFields, projectEmail } from './field-projection.js';
import { assertStripQuotedNotRaw } from './quote-strip.js';

// get_thread's orchestration, extracted from the CallTool switch so it is unit-testable
// with an injected client: the flag guards, the body-size cap, the per-message body
// signals, and the hidden-draft note. The handler is a thin text wrapper around this.

// The client surface get_thread needs. JmapClient satisfies it structurally.
export interface ThreadClient {
  getThread(
    threadId: string,
    includeDrafts?: boolean,
    includeBodies?: boolean,
  ): Promise<{ emails: any[]; hiddenDraftCount: number }>;
}

// Total plain-text body bytes one get_thread response may carry (#74). Roughly 25k
// tokens: large enough to return a whole ordinary conversation in one call, small enough
// that it cannot quietly consume a context window. Chosen as a cap that ERRORS rather
// than truncates — a silently shortened body is indistinguishable from a short message,
// which is precisely the trap the thread-body feature exists to remove. A thread that
// trips it is almost always one with deep quoted history, where stripQuoted:true brings
// it back under by removing duplication rather than content.
export const THREAD_BODY_BYTE_CAP = 100_000;

const LARGEST_LISTED = 3;

function byteLength(s: string): number {
  return Buffer.byteLength(s, 'utf8');
}

// Per-message body byte counts, largest first, for the cap check and its error message.
function bodyByteEntries(sizes: Array<{ id: string; bytes: number }>) {
  return [...sizes].sort((a, b) => b.bytes - a.bytes);
}

// The remedy has to be one the caller can actually run from where they are: on the raw
// path "retry with stripQuoted:true" would be rejected by assertStripQuotedNotRaw, so the
// raw case gets its own wording.
function overCapRemedy(mode: { stripQuoted: boolean; raw: boolean }): string {
  if (mode.raw) {
    return 'raw returns JMAP unmodified and cannot be stripped, so drop raw and retry with stripQuoted:true, or fetch the messages individually with get_email.';
  }
  if (mode.stripQuoted) {
    return 'Quoted history was already stripped, so fetch these messages individually with get_email (run get_thread without includeBodies for the id list).';
  }
  return 'Retry with stripQuoted:true to drop the repeated quoted history, or fetch the messages individually with get_email.';
}

export function assertThreadBodiesWithinCap(
  sizes: Array<{ id: string; bytes: number }>,
  mode: { stripQuoted: boolean; raw: boolean },
): void {
  const total = sizes.reduce((sum, s) => sum + s.bytes, 0);
  if (total <= THREAD_BODY_BYTE_CAP) return;

  const largest = bodyByteEntries(sizes)
    .slice(0, LARGEST_LISTED)
    .map((s) => `${s.id} (${s.bytes} bytes)`)
    .join(', ');
  const remedy = overCapRemedy(mode);

  throw new InvalidInputError(
    `Thread bodies total ${total} bytes across ${sizes.length} message(s), over the ${THREAD_BODY_BYTE_CAP}-byte limit for a single get_thread response. ` +
    `${remedy} Largest: ${largest}.`,
  );
}

// Raw-path body size: everything Email/get returned in bodyValues. includeBodies fetches
// only the text values, so this is the same quantity the simplified path measures.
function rawBodyBytes(emails: any[]): Array<{ id: string; bytes: number }> {
  return emails.map((e: any) => ({
    id: e?.id ?? '(unknown id)',
    bytes: Object.values(e?.bodyValues ?? {}).reduce(
      (sum: number, bv: any) => sum + byteLength(typeof bv?.value === 'string' ? bv.value : ''),
      0,
    ),
  }));
}

export async function readThread(args: any, client: ThreadClient): Promise<string> {
  const { threadId } = args ?? {};
  // `raw` keeps the plain truthiness test the other read tools use; the new flags follow
  // the lenient-value convention (a stringified "true"/"false" is accepted).
  const raw = !!args?.raw;
  const includeDrafts = coerceBool(args?.includeDrafts) ?? false;
  const includeBodies = coerceBool(args?.includeBodies) ?? false;
  const stripQuoted = coerceBool(args?.stripQuoted) ?? false;

  assertStripQuotedNotRaw(stripQuoted, raw);
  // Validated before the fetch, and rejected with raw for the same reason as the other
  // read tools: raw is untransformed JMAP, whose field names differ (#69).
  const fields = parseEmailFields(args?.fields, { raw });
  // stripQuoted rewrites bodies, and without includeBodies there are none. Rejecting is
  // the same reasoning as the unknown-parameter guard: a flag that silently does nothing
  // lets a caller believe it read stripped bodies when it read previews.
  if (stripQuoted && !includeBodies) {
    throw new InvalidInputError(
      'stripQuoted has nothing to strip without includeBodies: get_thread returns no bodies by default. ' +
      'Set includeBodies:true, or drop stripQuoted.',
    );
  }

  const { emails, hiddenDraftCount } = await client.getThread(threadId, includeDrafts, includeBodies);

  if (raw) {
    // raw is a pure-JSON escape valve that external clients may JSON.parse wholesale; the
    // draft note (below) is appended only on the simplified path so raw output stays
    // faithfully parseable. hiddenDraftCount never leaks into the raw JSON — a raw
    // consumer can pass includeDrafts itself. The size cap still applies: it guards the
    // response, not the formatting.
    if (includeBodies) assertThreadBodiesWithinCap(rawBodyBytes(emails), { stripQuoted: false, raw: true });
    return JSON.stringify(emails, null, 2);
  }

  const simplified: SimplifiedEmail[] = emails.map((e: any) => simplifyEmail(e, { stripQuoted }));

  if (includeBodies) {
    // Never-silent: a message with no plain-text body (an HTML-only one) yields no
    // bodyText, and thread reads deliberately never carry HTML. Say so per message
    // instead of letting the field quietly go missing. Set before projection so the
    // flag rides along with a projected bodyText.
    for (const msg of simplified) {
      if (typeof msg.bodyText !== 'string') msg.bodyTextUnavailable = true;
    }
  }

  // Projection runs BEFORE the cap so the cap measures what is actually returned:
  // a caller who projected bodyText away gets the small response they asked for, not
  // an error about bytes that never reach the output.
  const projected = simplified.map((m) => projectEmail(m, fields));

  if (includeBodies) {
    assertThreadBodiesWithinCap(
      // Ids come from the pre-projection objects so the over-cap error can still name
      // the largest messages when the caller didn't project id.
      projected.map((m, i) => ({ id: simplified[i].id, bytes: byteLength((m as Partial<SimplifiedEmail>).bodyText ?? '') })),
      { stripQuoted, raw: false },
    );
  }

  let text = JSON.stringify(projected, null, 2);
  if (hiddenDraftCount > 0) {
    text += `\n\nNote: ${hiddenDraftCount} draft(s) in this thread are hidden; set includeDrafts:true to include them.`;
  }
  return text;
}
