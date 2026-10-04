// Content-line folding (RFC 5545 §3.1), shared by `caldav-client.ts` and `vtimezone.ts` (#166).
// Not in `caldav-client.ts`, which imports `vtimezone.ts`: that would be an import cycle.

/**
 * Fold an iCalendar content line at 75 octets per RFC 5545 §3.1.
 * @param lineEnding Line ending to use for fold breaks (default '\r\n')
 */
export function foldICalLine(line: string, lineEnding: string = '\r\n'): string {
  const parts: string[] = [];
  let start = 0;
  let octets = 0;
  for (let i = 0; i < line.length; i++) {
    const code = line.charCodeAt(i);
    const isLow = (code & 0xFC00) === 0xDC00;
    let size: number;
    if (code < 0x80) size = 1;
    else if (code < 0x800) size = 2;
    // A high surrogate counts as 3 octets, so the low that completes a pair in this segment adds 1.
    // Any other surrogate is lone and encodes as U+FFFD, 3 octets, as Buffer.byteLength counts it.
    else if (isLow && i > start && (line.charCodeAt(i - 1) & 0xFC00) === 0xD800) size = 1;
    else size = 3;
    if (octets + size > 75) {
      // Don't split a surrogate pair: a cut before any low surrogate, even a lone one, moves back a unit.
      const cut = isLow ? i - 1 : i;
      parts.push(line.slice(start, cut));
      start = cut;
      i = cut - 1;
      octets = 1; // the continuation line's leading space
      continue;
    }
    octets += size;
  }
  parts.push(line.slice(start));
  return parts.join(lineEnding + ' ');
}
