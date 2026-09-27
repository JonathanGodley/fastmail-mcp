// Content-line folding (RFC 5545 §3.1), shared by `caldav-client.ts` and `vtimezone.ts` (#166).
// Not in `caldav-client.ts`, which imports `vtimezone.ts`: that would be an import cycle.

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
