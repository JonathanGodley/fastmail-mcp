import { describeUntrusted } from './coerce.js';

// Fastmail's session discovery hands back region-pinned endpoints, so the
// allowlist matches host *shapes*, not a fixed list of names. Observed live:
//   apiUrl / uploadUrl -> phl.api.fastmail.com   (region as a leading label)
//   downloadUrl        -> phl-www.fastmailusercontent.com  (region hyphenated
//                                                           onto the `www` label)
// The two hosts spell the region differently, hence two patterns. Both are
// anchored at each end and the region part excludes `.`, so a single label is
// all that can precede the fixed suffix — `evil.phl-www.fastmailusercontent.com`
// and `evilapi.fastmail.com.attacker.com` both still fail.
const REGION = '[a-z0-9]+(?:-[a-z0-9]+)*';
const FASTMAIL_ALLOWED_HOST_PATTERNS: readonly RegExp[] = [
  new RegExp(`^(?:${REGION}\\.)?api\\.fastmail\\.com$`),
  new RegExp(`^(?:${REGION}-)?www\\.fastmailusercontent\\.com$`),
];

function isAllowedFastmailHost(hostname: string): boolean {
  // URL parsing already lowercases and punycodes the hostname.
  return FASTMAIL_ALLOWED_HOST_PATTERNS.some((re) => re.test(hostname));
}

/**
 * Validate that a URL is acceptable for sending the bearer token to. HTTPS always;
 * `allowUnsafe` (FASTMAIL_ALLOW_UNSAFE_BASE_URL, for self-hosted JMAP) lifts only the host
 * allowlist. Throws on rejection; returns the parsed URL on success.
 */
export function validateFastmailUrl(input: string, fieldName: string, allowUnsafe = false): URL {
  let parsed: URL;
  try {
    parsed = new URL(input);
  } catch {
    throw new Error(`${fieldName} is not a valid URL`);
  }
  if (parsed.protocol !== 'https:') {
    throw new Error(
      `${fieldName} must use HTTPS (got: ${parsed.protocol}). ` +
      `Plain HTTP is rejected because the bearer token would be sent in cleartext.`,
    );
  }
  if (!allowUnsafe && !isAllowedFastmailHost(parsed.hostname)) {
    // The hostname comes from FASTMAIL_BASE_URL or the session response, and the URL parser
    // admits an apostrophe in a host, so it is echoed as untrusted (docs/conventions.md).
    throw new Error(
      `${fieldName} host "${describeUntrusted(parsed.hostname)}" is not in the Fastmail allowlist ` +
      `(api.fastmail.com and www.fastmailusercontent.com, each with an optional ` +
      `regional prefix such as phl.api.fastmail.com). ` +
      `Set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true to opt in for self-hosted JMAP servers.`,
    );
  }
  return parsed;
}
