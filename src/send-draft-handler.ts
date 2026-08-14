import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import type { SourceReferences } from './jmap-client.js';

// The minimal client surface sendDraftAndMaintainKeywords needs; JmapClient satisfies it
// structurally. Declared here (like ReplyClient/ForwardClient) so the orchestration stays
// unit-testable with a mock and free of a hard dependency on the concrete client.
export interface SendDraftClient {
  sendDraft(emailId: string): Promise<string>;
  getSourceReferences(emailId: string): Promise<SourceReferences>;
  findEmailIdsByMessageId(messageId: string): Promise<string[]>;
  addKeywords(emailId: string, keywords: string[]): Promise<void>;
}

// Why the original wasn't marked, when the draft DID name one. A keyword-write failure is
// not in this list: it is reported the same way the direct send=true paths report theirs
// (silently, see below), so it never reaches the result text.
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
  // The provenance read itself failed after the send, so we can't say whether this draft
  // replied to or forwarded anything. Distinct from "no provenance": the mail is sent
  // either way, but here the caller has no way to know maintenance was even applicable.
  sourceReadFailed?: boolean;
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
// (#60, #54). Without this, the recommended review-the-draft-then-send workflow would be
// the one flow that leaves the original untouched, while the direct reply/forward
// send=true paths mark it.
//
// Everything after the submission is best-effort: the mail is already gone, so no failure
// here may fail the call or roll anything back. Where the direct paths just mark an id the
// caller handed them, this path has to work out WHICH message the draft came from, so it
// has one failure mode they don't — the source Message-ID resolving to no message or to
// several. That one is reported in the result text (the caller never named an original, so
// silence would leave them no way to know maintenance was skipped, nor which message to
// mark by hand). A keyword-write failure stays silent, matching reply_email/forward_email
// on the identical failure.
export async function sendDraftAndMaintainKeywords(
  args: any,
  client: SendDraftClient,
): Promise<SendDraftResult> {
  const emailId = args?.emailId;
  if (!emailId) {
    throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
  }

  const submissionId = await client.sendDraft(emailId);

  // Read the provenance AFTER the send: submission keeps the same email id (it moves to
  // Sent and drops $draft, it is not recreated), and doing it here costs nothing on the
  // path where the send itself fails.
  let refs: SourceReferences;
  try {
    refs = await client.getSourceReferences(emailId);
  } catch {
    return { submissionId, sourceReadFailed: true };
  }

  const source = selectSource(refs);
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
