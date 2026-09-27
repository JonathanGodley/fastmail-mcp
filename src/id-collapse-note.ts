/**
 * The one sentence that discloses a duplicate-id collapse (#185), shared by the bulk
 * success texts (response-formatters.ts) and the bulk failure text (jmap-client.ts's
 * throwBulkSetError trailingNote) so the two cannot disagree. Imports neither module, so
 * either can import it without a cycle.
 *
 * The bulk tools build their write as an id-keyed map, so a duplicated id is written once
 * and the raw submitted count would overstate what changed. Every distinct id was acted on,
 * so the sentence names duplication as the cause and says nothing was skipped.
 */
export function buildIdCollapseNote(rawIds: string[]): string {
  const submitted = rawIds.length;
  const distinct = new Set(rawIds).size;
  if (submitted === distinct) return '';
  // submitted is at least 2 here (one id cannot collapse), so "ids" is always plural.
  const emailsPhrase = distinct === 1 ? '1 distinct email' : `${distinct} distinct emails`;
  return `${submitted} ids were given, but duplicates collapsed them to ${emailsPhrase}; nothing was skipped.`;
}
