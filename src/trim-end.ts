/**
 * `text` with its trailing run of characters that `drop` accepts removed.
 *
 * A walk back from the end, never a `/[…]+$/` regex: the regex retries the run from every
 * start position, so it is quadratic on a long run followed by anything else, and every
 * caller trims text this server did not write.
 */
export function trimEnd(text: string, drop: (ch: string) => boolean): string {
  let end = text.length;
  while (end > 0 && drop(text[end - 1])) end--;
  return text.slice(0, end);
}

/** What `\s` matches; each such character is a single UTF-16 code unit. */
export const isWhitespace = (ch: string): boolean => /\s/.test(ch);
