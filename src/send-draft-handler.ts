import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import type { SendDraftOutcome, SourceReferences } from './jmap-client.js';

// The minimal client surface sendDraftAndMaintainKeywords needs; JmapClient satisfies it
// structurally. Declared here (like ReplyClient/ForwardClient) so the orchestration stays
// unit-testable with a mock and free of a hard dependency on the concrete client.
export interface SendDraftClient {
  sendDraft(emailId: string): Promise<SendDraftOutcome>;
  findEmailIdsByMessageId(messageId: string): Promise<string[]>;
  addKeywords(emailId: string, keywords: string[]): Promise<void>;
}

// Why the original wasn't marked, when the draft DID name one. A keyword-write failure is
// not in this list: it stays silent (see below), so it never reaches the result text.
export type KeywordSkipReason = 'not-found' | 'ambiguous' | 'lookup-failed';

// Present only when the sent draft recorded a source message.
export interface KeywordMaintenance {
  kind: 'reply' | 'forward';
  // The source Message-ID read off the draft (bare, no angle brackets).
  messageId: string;
  // The resolved JMAP id of the original, when the lookup found exactly one.
  originalEmailId?: string;
  // True when the original was actually marked.
  marked: boolean;
  // Set when the original could not be identified; absent when the lookup succeeded and
  // only the keyword write failed.
  skipReason?: KeywordSkipReason;
}

export interface SendDraftResult {
  submissionId: string;
  keywordMaintenance?: KeywordMaintenance;
}

const KEYWORDS: Record<'reply' | 'forward', string[]> = {
  reply: ['$answered', '$seen'],
  forward: ['$forwarded', '$seen'],
};

function firstId(ids: string[] | undefined): string | undefined {
  return (ids ?? []).map(id => String(id ?? '').trim()).find(id => id.length > 0);
}

// Pick the source message the draft was composed from. In-Reply-To wins over
// X-Forwarded-Message-Id on a draft carrying both — mutually exclusive dispatch,
// reply-first, matching how the edit guard chooses its variant on the same pair of
// signals. A multi-id In-Reply-To (never written by this server; some clients list the
// whole ancestry) takes its first entry, the conventional immediate parent.
export function selectSource(refs: SourceReferences): { kind: 'reply' | 'forward'; messageId: string } | undefined {
  const replyTo = firstId(refs?.inReplyTo);
  if (replyTo) return { kind: 'reply', messageId: replyTo };
  const forwarded = firstId(refs?.forwardedMessageId);
  if (forwarded) return { kind: 'forward', messageId: forwarded };
  return undefined;
}

// Send a saved draft, then maintain the thread state of whatever message it was composed
// from: a reply marks the original answered + read, a forward marks it forwarded + read
// (#60, #54). The compose surface is draft-first (#32/#66) — every reply and forward is
// transmitted here — so this is the only place that maintenance can happen. Both forward
// shapes record the header, including asAttachment (whose .eml carries the content but
// is not machine-resolvable as provenance), so both mark their original on send.
//
// Everything after the submission is best-effort: the mail is already gone, so no failure
// here may fail the call or roll anything back. The draft names its source by Message-ID,
// so this path has to work out WHICH message it came from, and that lookup can fail — the
// source Message-ID resolving to no message or to several. That one is reported in the
// result text (the caller never named an original, so silence would leave them no way to
// know maintenance was skipped, nor which message to mark by hand). A keyword-write
// failure stays silent: the send succeeded, and the flags are cosmetic thread state.
export async function sendDraftAndMaintainKeywords(
  args: any,
  client: SendDraftClient,
): Promise<SendDraftResult> {
  const emailId = args?.emailId;
  if (!emailId) {
    throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
  }

  // The provenance rides back on the send: sendDraft reads it off the same pre-send
  // Email/get that validates the draft, so there is no second fetch and no "the send
  // worked but we couldn't tell what it was" state — a draft we can't read never sends.
  const { submissionId, sourceReferences } = await client.sendDraft(emailId);

  const source = selectSource(sourceReferences);
  if (!source) return { submissionId }; // a fresh compose has no original to mark

  let candidates: string[];
  try {
    candidates = await client.findEmailIdsByMessageId(source.messageId);
  } catch {
    return { submissionId, keywordMaintenance: { ...source, marked: false, skipReason: 'lookup-failed' } };
  }

  if (candidates.length === 0) {
    return { submissionId, keywordMaintenance: { ...source, marked: false, skipReason: 'not-found' } };
  }
  if (candidates.length > 1) {
    // Two stored messages carrying the same Message-ID (a duplicate delivery, or a copy
    // kept in another folder). Marking one would be a guess about which the caller means,
    // and marking both would spread a keyword across a message we were never pointed at.
    return { submissionId, keywordMaintenance: { ...source, marked: false, skipReason: 'ambiguous' } };
  }

  const originalEmailId = candidates[0];
  try {
    await client.addKeywords(originalEmailId, KEYWORDS[source.kind]);
    return { submissionId, keywordMaintenance: { ...source, originalEmailId, marked: true } };
  } catch {
    /* best-effort: the draft already sent */
    return { submissionId, keywordMaintenance: { ...source, originalEmailId, marked: false } };
  }
}
