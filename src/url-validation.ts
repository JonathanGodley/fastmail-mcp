// Validates URLs that will receive the bearer token. Restricts to approved
// Fastmail origins by default, with an explicit opt-out for self-hosted JMAP.

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
 * Validate that a URL is acceptable for sending the bearer token to.
 *
 * Default policy:
 *   - Must be HTTPS.
 *   - Hostname must match FASTMAIL_ALLOWED_HOST_PATTERNS: Fastmail's API and
 *     user-content hosts, with or without a regional prefix.
 *
 * When `allowUnsafe=true` (e.g. user opted in via FASTMAIL_ALLOW_UNSAFE_BASE_URL
 * for a self-hosted JMAP server):
 *   - Must still be HTTPS (plain HTTP is never allowed; the token would be sent
 *     in cleartext).
 *   - Any hostname is accepted.
 *
 * Throws on rejection; returns the parsed URL on success.
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
    throw new Error(
      `${fieldName} host '${parsed.hostname}' is not in the Fastmail allowlist ` +
      `(api.fastmail.com and www.fastmailusercontent.com, each with an optional ` +
      `regional prefix such as phl.api.fastmail.com). ` +
      `Set FASTMAIL_ALLOW_UNSAFE_BASE_URL=true to opt in for self-hosted JMAP servers.`,
    );
  }
  return parsed;
}
