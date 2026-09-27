import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import type { SendDraftOutcome, SourceReferences } from './jmap-client.js';

// The client surface sendDraftAndMaintainKeywords needs; JmapClient satisfies it
// structurally.
export interface SendDraftClient {
  sendDraft(emailId: string): Promise<SendDraftOutcome>;
  findEmailIdsByMessageId(messageId: string): Promise<string[]>;
  getEmailMessageId(emailId: string): Promise<string[] | null>;
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
  // What the transmitted message carried as embedded images, reported straight through
  // from the send (#13).
  notes?: string[];
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
// (#60, #54). Every reply and forward is transmitted here (#32/#66), so this is the only
// place that maintenance can happen. Both forward shapes, asAttachment included, record the
// provenance header, so both mark their original.
//
// Everything after the submission is best-effort: the mail is already gone, so no failure
// here may fail the call. A failed source lookup (none or several matches) is reported,
// because the caller never named an original and could not otherwise know to mark one by
// hand. A keyword-write failure stays silent: the flags are cosmetic thread state.
export async function sendDraftAndMaintainKeywords(
  args: any,
  client: SendDraftClient,
): Promise<SendDraftResult> {
  const emailId = args?.emailId;
  if (!emailId) {
    throw new McpError(ErrorCode.InvalidParams, 'emailId is required');
  }

  // sendDraft reads the provenance off the same pre-send Email/get that validates the
  // draft: no second fetch, and a draft that cannot be read never sends.
  const { submissionId, sourceReferences, notes } = await client.sendDraft(emailId);

  // One builder so every return, early ones included, carries the send's receipt.
  const sent = (keywordMaintenance?: KeywordMaintenance): SendDraftResult => ({
    submissionId,
    ...(notes?.length && { notes }),
    ...(keywordMaintenance && { keywordMaintenance }),
  });

  const source = selectSource(sourceReferences);
  if (!source) return sent(); // a fresh compose has no original to mark

  // Exact-instance path. draft_email's reply and forward modes record the JMAP id of the
  // stored instance composed from (X-Fastmail-MCP-Source-Id). An account can hold several
  // copies of one Message-ID, and Fastmail's own client marks only the instance replied to
  // (observed live 2026-08-14). The pointer must still exist and carry the draft's
  // Message-ID; a destroyed instance, a mismatch or a failed read falls through to the
  // Message-ID lookup below.
  const recorded = sourceReferences.sourceEmailId;
  if (recorded) {
    let instanceMessageIds: string[] | null = null;
    try {
      instanceMessageIds = await client.getEmailMessageId(recorded);
    } catch {
      /* transient read failure — the lookup below is the fallback */
    }
    if (instanceMessageIds?.includes(source.messageId)) {
      try {
        await client.addKeywords(recorded, KEYWORDS[source.kind]);
        return sent({ ...source, originalEmailId: recorded, marked: true });
      } catch {
        /* best-effort: the draft already sent; write failure stays silent */
        return sent({ ...source, originalEmailId: recorded, marked: false });
      }
    }
  }

  let candidates: string[];
  try {
    candidates = await client.findEmailIdsByMessageId(source.messageId);
  } catch {
    return sent({ ...source, marked: false, skipReason: 'lookup-failed' });
  }

  if (candidates.length === 0) {
    return sent({ ...source, marked: false, skipReason: 'not-found' });
  }
  if (candidates.length > 1) {
    // Marking one would be a guess; marking all would touch messages we were never
    // pointed at.
    return sent({ ...source, marked: false, skipReason: 'ambiguous' });
  }

  const originalEmailId = candidates[0];
  try {
    await client.addKeywords(originalEmailId, KEYWORDS[source.kind]);
    return sent({ ...source, originalEmailId, marked: true });
  } catch {
    /* best-effort: the draft already sent */
    return sent({ ...source, originalEmailId, marked: false });
  }
}
