// Drift guard for the DXT configuration surface (manifest.json).
//
// A DXT setting only reaches the server if TWO things agree: `user_config` declares the
// key so the installer prompts for it, and `server.mcp_config.env` maps some environment
// variable to `${user_config.<key>}` so the answer is handed to the process. Neither half
// fails loudly on its own — an unmirrored key produces an input box whose value goes
// nowhere, and a stale `${user_config.x}` reference produces an environment variable set
// to an uninterpolated placeholder string. Both look configured and do nothing.
//
// The third check goes further and asks whether the server reads the variable at all, by
// scanning src/index.ts and its siblings as TEXT. That mirrors src/tool-schema.test.ts and
// is text-based for the same reason: `npm test` runs tsx over src/ and never builds, so a
// guard reading dist/ would validate the last build rather than the current source — which
// is precisely the drift it exists to catch. tsc does not rewrite string literals, so the
// source and the shipped code cannot disagree on these names.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';

const SRC_URL = new URL('./', import.meta.url);
const MANIFEST_URL = new URL('../manifest.json', import.meta.url);
const INDEX_URL = new URL('./index.ts', import.meta.url);

interface Manifest {
  server: { mcp_config: { env: Record<string, string> } };
  user_config: Record<string, unknown>;
}

function readManifest(): Manifest {
  return JSON.parse(readFileSync(MANIFEST_URL, 'utf8')) as Manifest;
}

// The string literals inside every findEnvValue([...]) call in index.ts. The array
// contents are matched as one blob so a multi-line call is picked up whole.
function collectFindEnvValueNames(): string[] {
  const source = readFileSync(INDEX_URL, 'utf8');
  const names: string[] = [];
  for (const call of source.matchAll(/findEnvValue\(\s*\[([^\]]*)\]/g)) {
    for (const literal of call[1].matchAll(/['"]([^'"]+)['"]/g)) {
      names.push(literal[1]);
    }
  }
  return names;
}

// `process.env.NAME` reads outside the findEnvValue helper. Configurable settings are
// meant to go through findEnvValue (that lookup is what lets a DXT user_config key reach
// the server at all), but a module reading its own variable at the point of use is still
// reading it, and the manifest may legitimately map it. Every non-test module is scanned
// rather than a fixed list, so moving such a read between files cannot silently turn a
// live setting into a "dead configuration" failure here. Comment lines are skipped so a
// name that is only discussed in prose does not count as a read.
function collectDirectEnvReads(): string[] {
  const names: string[] = [];
  const files = readdirSync(SRC_URL)
    .filter((f) => f.endsWith('.ts') && !f.endsWith('.test.ts'));
  for (const file of files) {
    const source = readFileSync(new URL(`./${file}`, import.meta.url), 'utf8');
    for (const line of source.split('\n')) {
      const trimmed = line.trim();
      if (trimmed.startsWith('//') || trimmed.startsWith('*')) continue;
      for (const read of line.matchAll(/process\.env\.([A-Za-z_][A-Za-z0-9_]*)/g)) {
        names.push(read[1]);
      }
    }
  }
  return names;
}

function envNamesTheServerReads(): Set<string> {
  return new Set([...collectFindEnvValueNames(), ...collectDirectEnvReads()]);
}

// The env value that carries a user_config answer, e.g. "${user_config.fastmail_timezone}".
function userConfigReference(key: string): string {
  return '${user_config.' + key + '}';
}

describe('DXT configuration surface', () => {
  it('maps every declared user_config key into the server environment', () => {
    const manifest = readManifest();
    const declared = Object.keys(manifest.user_config);
    const mirrored = new Set(Object.values(manifest.server.mcp_config.env));

    const unmirrored = declared.filter((key) => !mirrored.has(userConfigReference(key)));
    assert.deepEqual(
      unmirrored,
      [],
      `these user_config keys are never mapped into server.mcp_config.env, so the installer ` +
        `collects a value that reaches nothing. Add an env entry whose value is exactly ` +
        `\${user_config.<key>}: ${unmirrored.join(', ')}`,
    );
  });

  it('resolves every ${user_config.X} reference to a declared user_config key', () => {
    const manifest = readManifest();
    const declared = new Set(Object.keys(manifest.user_config));

    const dangling: string[] = [];
    for (const [envName, value] of Object.entries(manifest.server.mcp_config.env)) {
      for (const ref of value.matchAll(/\$\{user_config\.([^}]+)\}/g)) {
        if (!declared.has(ref[1])) dangling.push(`${envName} -> ${ref[1]}`);
      }
    }
    assert.deepEqual(
      dangling,
      [],
      `these env entries reference a user_config key that is not declared, so the variable is ` +
        `set to an uninterpolated placeholder rather than the user's answer: ${dangling.join(', ')}`,
    );
  });

  it('maps only environment variables the server actually reads', () => {
    const manifest = readManifest();
    const read = envNamesTheServerReads();

    // A sanity floor: if the source scan ever stops matching (the helper is renamed, the
    // env reads move to another file) the check below would pass vacuously.
    assert.ok(
      read.size >= 10,
      `found only ${read.size} environment names in the source; the scan has probably stopped matching`,
    );

    const dead = Object.keys(manifest.server.mcp_config.env).filter((name) => !read.has(name));
    assert.deepEqual(
      dead,
      [],
      `the manifest maps these environment variables but nothing in the server reads them, so ` +
        `configuring them does nothing: ${dead.join(', ')}`,
    );
    // Deliberately one-directional. The server also reads names the manifest does not
    // declare — the USER_CONFIG_* fallback spellings some hosts use, and the base-URL
    // kill switch below — and that is correct, so the reverse is not asserted.
  });

  it('keeps FASTMAIL_ALLOW_UNSAFE_BASE_URL out of the manifest', () => {
    const manifest = readManifest();
    const envNames = Object.keys(manifest.server.mcp_config.env);
    const configKeys = Object.keys(manifest.user_config);
    const mentions = [
      ...envNames.filter((n) => n === 'FASTMAIL_ALLOW_UNSAFE_BASE_URL'),
      ...configKeys.filter((k) => k.toLowerCase() === 'fastmail_allow_unsafe_base_url'),
    ];

    assert.deepEqual(
      mentions,
      [],
      `FASTMAIL_ALLOW_UNSAFE_BASE_URL disables the endpoint allowlist that decides where the ` +
        `API token may be sent. It is a kill switch for self-hosted JMAP servers and is ` +
        `deliberately environment-only: surfacing it as an installer setting would put a ` +
        `security control one checkbox away for users who have no reason to touch it. Adding ` +
        `it here to "complete" the mapping is the mistake this test exists to stop.`,
    );
  });
});
