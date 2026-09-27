import { redactBearerTokens, toolJson } from './coerce.js';

// The client surface test_bulk_operations writes through. JmapClient satisfies it structurally.
export interface BulkTestClient {
  bulkMarkRead(emailIds: string[], read: boolean): Promise<void>;
}

const STEP_PAUSE_MS = 500;

/**
 * test_bulk_operations, after the handler has read the messages: step 1 marks them all read,
 * step 2 marks unread again only the ones that were unread before, so each message ends in
 * its own prior state. Step 2 runs even when step 1 failed part-way: it writes only the prior
 * state of messages that were unread, so a message step 1 never reached is left as it was.
 * A message whose keywords were not reported has no known prior state, so neither step
 * touches it.
 */
export async function runBulkReadTest(
  emails: any[],
  dryRun: boolean,
  client: BulkTestClient,
  pause: (ms: number) => Promise<void> = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
): Promise<string> {
  const hasReadState = (email: any) =>
    !!email.keywords && typeof email.keywords === 'object' && !Array.isArray(email.keywords);
  const known = emails.filter(hasReadState);
  const emailIds = known.map((email) => email.id);
  const unreadIds = known.filter((email) => !email.keywords.$seen).map((email) => email.id);

  const operations: Array<{ name: string; description: string; parameters: { emailIds: string[]; read: boolean } }> = [
    {
      name: 'bulk_mark_read',
      description: `Mark ${emailIds.length} emails as read`,
      parameters: { emailIds, read: true },
    },
  ];
  if (unreadIds.length > 0) {
    operations.push({
      name: 'bulk_mark_read (restore)',
      description: `Mark the ${unreadIds.length} of them that were unread before as unread again`,
      parameters: { emailIds: unreadIds, read: false },
    });
  }

  const results = {
    testEmails: emails.map((email) => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.[0]?.email || 'unknown',
      receivedAt: email.receivedAt,
      ...(hasReadState(email) ? { wasRead: !!email.keywords.$seen } : { excluded: true }),
    })),
    operations: [] as any[],
  };

  if (dryRun) {
    results.operations = operations.map((op) => ({
      ...op,
      status: 'DRY RUN - Would execute but not actually performed',
      executed: false,
    }));
    return `BULK OPERATIONS TEST (DRY RUN)\n\n${toolJson(results)}\n\nTo actually execute the test, set dryRun: false`;
  }

  for (const [i, operation] of operations.entries()) {
    try {
      await client.bulkMarkRead(operation.parameters.emailIds, operation.parameters.read);
      results.operations.push({ ...operation, status: 'SUCCESS', executed: true, timestamp: new Date().toISOString() });
    } catch (error) {
      results.operations.push({
        ...operation,
        status: 'FAILED',
        executed: false,
        // Folded into result JSON rather than raised, so the top-level catch's redaction
        // never sees it.
        error: redactBearerTokens(error instanceof Error ? error.message : String(error)),
        timestamp: new Date().toISOString(),
      });
    }
    if (i < operations.length - 1) await pause(STEP_PAUSE_MS);
  }

  return `BULK OPERATIONS TEST (EXECUTED)\n\n${toolJson(results)}`;
}
