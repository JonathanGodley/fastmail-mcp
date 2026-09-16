/**
 * The one sentence that discloses a duplicate-id collapse (#185), shared by the bulk
 * success texts (response-formatters.ts) and the bulk failure text (jmap-client.ts's
 * throwBulkSetError trailingNote) so the two can never say something different about the
 * same call. Imports neither of those modules, so importing this one into either never
 * creates a cycle.
 *
 * `bulk_mark_read`/`bulk_pin`/`bulk_move`/`bulk_delete`/`bulk_add_labels` build their write
 * by assigning into an id-keyed map, so a duplicated id collapses to one entry and is
 * written once - reporting the raw submitted count as the subject would then claim two
 * emails changed when only one was. Silence is correct once submitted and distinct counts
 * agree; otherwise the sentence has to name duplication as the cause and say nothing was
 * skipped, since every distinct id was in fact acted on.
 */
export function buildIdCollapseNote(rawIds: string[]): string {
  const submitted = rawIds.length;
  const distinct = new Set(rawIds).size;
  if (submitted === distinct) return '';
  // submitted is never 1 here: a single-element array has exactly 1 distinct element, so
  // submitted === 1 would force distinct === submitted, already returned '' above.
  const emailsPhrase = distinct === 1 ? '1 distinct email' : `${distinct} distinct emails`;
  return `${submitted} ids were given, but duplicates collapsed them to ${emailsPhrase}; nothing was skipped.`;
}
