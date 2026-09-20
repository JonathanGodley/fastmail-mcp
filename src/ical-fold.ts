// Content-line folding (RFC 5545 §3.1), shared by `caldav-client.ts` (every property line it
// writes) and `vtimezone.ts` (every line of a generated VTIMEZONE, #166). Lives here, not in
// `caldav-client.ts`, to avoid an import cycle: `caldav-client.ts` imports `generateVTimezone`
// from `vtimezone.ts`, so `vtimezone.ts` cannot import back from it.

/**
 * Fold an iCalendar content line at 75 octets per RFC 5545 §3.1.
 * @param lineEnding Line ending to use for fold breaks (default '\r\n')
 */
export function foldICalLine(line: string, lineEnding: string = '\r\n'): string {
  const parts: string[] = [];
  while (Buffer.byteLength(line, 'utf8') > 75) {
    // Find the largest character count that fits in 75 bytes
    let cut = 75;
    while (cut > 0 && Buffer.byteLength(line.slice(0, cut), 'utf8') > 75) {
      cut--;
    }
    // Don't split a surrogate pair (characters outside BMP like emoji)
    if (cut > 0 && cut < line.length) {
      const code = line.charCodeAt(cut);
      if (code >= 0xDC00 && code <= 0xDFFF) cut--;
    }
    parts.push(line.slice(0, cut));
    line = ' ' + line.slice(cut);
  }
  parts.push(line);
  return parts.join(lineEnding);
}
