import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { validateFastmailUrl } from './url-validation.js';

describe('validateFastmailUrl (default policy)', () => {
  it('accepts api.fastmail.com over HTTPS', () => {
    const url = validateFastmailUrl('https://api.fastmail.com/jmap/api/', 'apiUrl');
    assert.equal(url.hostname, 'api.fastmail.com');
  });

  it('accepts www.fastmailusercontent.com over HTTPS', () => {
    const url = validateFastmailUrl('https://www.fastmailusercontent.com/jmap/download/x/y/z', 'downloadUrl');
    assert.equal(url.hostname, 'www.fastmailusercontent.com');
  });

  it('accepts a regional API host', () => {
    // Fastmail session discovery returns region-pinned endpoints for some
    // accounts, e.g. apiUrl/uploadUrl on phl.api.fastmail.com.
    const url = validateFastmailUrl('https://phl.api.fastmail.com/jmap/api/', 'session.apiUrl');
    assert.equal(url.hostname, 'phl.api.fastmail.com');
  });

  it('accepts a regional user-content host', () => {
    // The user-content host spells the region with a hyphen, not a dot.
    const url = validateFastmailUrl(
      'https://phl-www.fastmailusercontent.com/jmap/download/x/y/z',
      'session.downloadUrl',
    );
    assert.equal(url.hostname, 'phl-www.fastmailusercontent.com');
  });

  it('accepts a multi-part regional prefix', () => {
    assert.equal(
      validateFastmailUrl('https://us-west1.api.fastmail.com/jmap/api/', 'apiUrl').hostname,
      'us-west1.api.fastmail.com',
    );
    assert.equal(
      validateFastmailUrl('https://us-west1-www.fastmailusercontent.com/x', 'downloadUrl').hostname,
      'us-west1-www.fastmailusercontent.com',
    );
  });

  it('rejects a regional prefix carrying extra labels', () => {
    // Only one label may precede the fixed suffix, so an attacker-controlled
    // host cannot hide in front of a legitimate-looking regional prefix.
    assert.throws(
      () => validateFastmailUrl('https://evil.phl.api.fastmail.com/jmap/api/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
    assert.throws(
      () => validateFastmailUrl('https://evil.phl-www.fastmailusercontent.com/x', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects a regional-looking prefix on the wrong domain', () => {
    assert.throws(
      () => validateFastmailUrl('https://phl.api.fastmail.com.attacker.com/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
    assert.throws(
      () => validateFastmailUrl('https://phl-www.fastmailusercontent.com.attacker.com/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects a regional prefix on a non-API fastmail host', () => {
    // The prefix is only allowed in front of the two endpoint hosts.
    assert.throws(
      () => validateFastmailUrl('https://phl.www.fastmail.com/jmap/api/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects HTTP even on allowed host', () => {
    assert.throws(
      () => validateFastmailUrl('http://api.fastmail.com/jmap/api/', 'baseUrl'),
      /must use HTTPS/,
    );
  });

  it('rejects non-allowlisted host', () => {
    assert.throws(
      () => validateFastmailUrl('https://attacker.example.com/jmap/api/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects subdomain of fastmail.com that is not on the explicit allowlist', () => {
    // www.fastmail.com is NOT in the allowlist — only api and the user-content host.
    assert.throws(
      () => validateFastmailUrl('https://www.fastmail.com/jmap/api/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects host that ends with allowlisted domain (suffix-attack)', () => {
    // Confirms exact-match check, not endsWith.
    assert.throws(
      () => validateFastmailUrl('https://evilapi.fastmail.com.attacker.com/', 'baseUrl'),
      /not in the Fastmail allowlist/,
    );
  });

  it('rejects malformed URL', () => {
    assert.throws(
      () => validateFastmailUrl('not a url', 'baseUrl'),
      /not a valid URL/,
    );
  });

  it('rejects javascript: scheme', () => {
    assert.throws(
      () => validateFastmailUrl('javascript:fetch("https://attacker")', 'baseUrl'),
      /must use HTTPS/,
    );
  });

  it('rejects ftp: scheme', () => {
    assert.throws(
      () => validateFastmailUrl('ftp://api.fastmail.com/jmap/', 'baseUrl'),
      /must use HTTPS/,
    );
  });

  it('error message names the field for diagnostics', () => {
    try {
      validateFastmailUrl('https://attacker.example.com/', 'session.apiUrl');
      assert.fail('should have thrown');
    } catch (e) {
      assert.match((e as Error).message, /session\.apiUrl/);
    }
  });
});

describe('validateFastmailUrl (allowUnsafe=true)', () => {
  it('accepts arbitrary HTTPS host when opted in', () => {
    const url = validateFastmailUrl('https://jmap.self-hosted.example/jmap/api/', 'baseUrl', true);
    assert.equal(url.hostname, 'jmap.self-hosted.example');
  });

  it('still rejects HTTP even with opt-in', () => {
    assert.throws(
      () => validateFastmailUrl('http://jmap.self-hosted.example/jmap/api/', 'baseUrl', true),
      /must use HTTPS/,
    );
  });

  it('still rejects non-URL input', () => {
    assert.throws(
      () => validateFastmailUrl('garbage', 'baseUrl', true),
      /not a valid URL/,
    );
  });
});
