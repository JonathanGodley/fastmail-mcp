#!/usr/bin/env node
// Secret & PII scanner - blocks credentials and personal information from being
// committed or published. Runs in CI (all tracked files), as a pre-commit hook
// (staged content), and as a release gate before the .dxt is packed.
// Self-contained: no third-party dependencies, and it embeds NO personal data -
// personal domains live only in a gitignored local denylist (see
// .secret-scan-local.txt.example).
//
// Usage:
//   node scripts/scan-secrets.mjs --all         scan all git-tracked files
//   node scripts/scan-secrets.mjs --staged      scan staged (pre-commit) content
//   node scripts/scan-secrets.mjs file1 file2   scan specific files
//
// Suppress a known-safe match by putting the marker  allowlist-secret  in a
// comment on the line (e.g. a synthetic test fixture). Exit code 1 on any
// finding, and also on any file the scanner could not read: a gate that cannot
// see a file reports that rather than reporting clean.
//
// The limits of what this covers (git history, untracked files, build output,
// the exempt-domain list, the marker's line-total reach) are documented in
// CONTRIBUTING.md. Keep that section in step with this file.

import { execFileSync } from 'node:child_process';
import { readFileSync, existsSync } from 'node:fs';

const PRAGMA = 'allowlist-secret';

// Generous ceiling for git output and blob reads; the default 1MB is smaller
// than files this repo tracks.
const MAX_BUFFER = 64 * 1024 * 1024;

// Paths never scanned: vendored code, build output, the lockfile, the packed
// bundle, and this scanner's own pattern definitions. These are policy
// exclusions - deliberate, listed here, and counted in the run summary.
const IGNORE = [
  /^node_modules\//,
  /^dist\//,
  /(^|\/)package-lock\.json$/,
  /\.lock$/,
  /\.dxt$/,
  /(^|\/)scan-secrets\.mjs$/,
];

// Tracked files that are legitimately binary and therefore cannot be scanned.
// Each entry is an exact repo-relative path and is a named hole in the gate's
// coverage, which is why they are declared here one at a time instead of being
// waved through by a filename or content heuristic. Nothing tracked in this
// repo is binary today, so the list is empty; adding a binary file means adding
// its path here, and the scan fails until you do.
const BINARY_ALLOWANCES = new Set([]);

// Domains considered non-personal placeholders/services. An email on any other
// domain is flagged as possible real PII. This list is intentionally generic -
// it names no personal domains.
// RFC 2606 / RFC 6761 reserve these top-level domains so they can never be
// delegated. An address under one of them cannot belong to a real person, so
// matching on the TLD is not a blind spot the way listing a registered domain
// would be - it closes the whole space rather than one name at a time.
const RESERVED_TLDS = ['.example', '.invalid', '.test', '.localhost'];

// Entries here are exempt by NAME, which means the scanner is permanently blind
// to a genuine address at that domain. Keep the list to names that cannot
// plausibly carry one, and prefer moving a fixture under a RESERVED_TLDS suffix
// over adding to it.
const SAFE_EMAIL_DOMAINS = new Set([
  'example.com', 'example.org', 'example.net', 'localhost',
  'fastmail.com', 'api.fastmail.com', 'caldav.fastmail.com',
  'www.fastmailusercontent.com', 'fastmailusercontent.com',
  'anthropic.com',
  // Adversary placeholders that have to read as a real-looking hostile domain
  // for the fixture to mean anything. Accepted blind spot, recorded in
  // CONTRIBUTING.md.
  'evil.com', 'other.com',
]);

function isSafeEmailDomain(domain) {
  return SAFE_EMAIL_DOMAINS.has(domain) || RESERVED_TLDS.some((tld) => domain.endsWith(tld));
}

const RULES = [
  { name: 'Fastmail API token', re: /fmu\d+-[A-Za-z0-9_-]{20,}/g },
  { name: 'Bearer credential', re: /Bearer\s+[A-Za-z0-9._~+/=-]{16,}/g },
  { name: 'Basic auth credential', re: /Basic\s+[A-Za-z0-9+/=]{16,}/g },
  {
    name: 'hardcoded secret assignment',
    re: /\b(?:api[_-]?key|secret|password|passwd|token)\b\s*[:=]\s*["'`][A-Za-z0-9_\-]{16,}["'`]/gi,
    // skip env refs / obvious placeholders
    skip: (m) => /\$\{|process\.env|your-|REDACTED|example|placeholder|x{6,}|0{6,}|changeme|<[^>]+>/i.test(m),
  },
];

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;

// The marker only suppresses when it sits in a comment, so that a credential
// whose own value happens to contain the marker string cannot suppress itself.
// Comment syntax is not parsed per language: the marker counts if any of these
// introducers appears earlier on the line, which covers //, #, /* */ and HTML
// comments, plus the leading * of a JSDoc continuation line.
const COMMENT_INTRODUCERS = ['//', '#', '/*', '<!--'];
const JSDOC_CONTINUATION_RE = /^\s*\*/;

// The marker must stand as its own token: preceded by start-of-line, whitespace
// or comment punctuation, and not run on into a longer word.
const PRAGMA_RE = new RegExp(`(?:^|[\\s*/#!])${PRAGMA}(?![\\w-])`, 'g');

function commentStartIndex(line) {
  let earliest = -1;
  for (const introducer of COMMENT_INTRODUCERS) {
    const at = line.indexOf(introducer);
    if (at !== -1 && (earliest === -1 || at < earliest)) earliest = at;
  }
  const jsdoc = JSDOC_CONTINUATION_RE.exec(line);
  if (jsdoc) {
    const at = jsdoc[0].length - 1;
    if (earliest === -1 || at < earliest) earliest = at;
  }
  return earliest;
}

// A marker occurrence suppresses the line when it is inside a comment and is not
// itself part of one of the candidate matches on that line. Suppression is
// line-total: it silences every rule on the line, not just the one it sits next
// to. See CONTRIBUTING.md.
function isSuppressed(line, candidates) {
  const commentAt = commentStartIndex(line);
  if (commentAt === -1) return false;
  PRAGMA_RE.lastIndex = 0;
  let m;
  while ((m = PRAGMA_RE.exec(line)) !== null) {
    const at = m.index + (m[0].length - PRAGMA.length);
    if (at <= commentAt) continue;
    const insideMatch = candidates.some((c) => at < c.index + c.length && at + PRAGMA.length > c.index);
    if (!insideMatch) return true;
  }
  return false;
}

// Optional local denylist (gitignored): one literal string or domain per line.
// Lets a developer catch their own personal domains without publishing them.
function loadLocalDenylist() {
  const path = '.secret-scan-local.txt';
  if (!existsSync(path)) return [];
  return readFileSync(path, 'utf8')
    .split('\n')
    .map((l) => l.trim())
    .filter((l) => l && !l.startsWith('#'));
}

function git(args) {
  return execFileSync('git', args, { encoding: 'buffer', maxBuffer: MAX_BUFFER });
}

// git quotes paths containing non-ASCII or unusual characters in its normal
// output, which turns them into names that do not exist on disk. The -z form
// emits raw NUL-separated paths instead, so no path can be dropped or mangled.
function gitPaths(args) {
  const out = git([...args, '-z']).toString('utf8');
  const seen = new Set();
  // git ls-files repeats a path once per stage while a merge is unresolved.
  for (const path of out.split('\0')) if (path) seen.add(path);
  return [...seen];
}

function targetFiles() {
  const mode = process.argv[2];
  if (mode === '--all') {
    return { staged: false, files: gitPaths(['ls-files']) };
  }
  if (mode === '--staged') {
    // A staged deletion (D) leaves no content in the commit, so there is
    // nothing to scan; every other status does put content at a path. R and C
    // name the destination path only, which is the content being committed.
    // Unmerged paths (U) are excluded because git refuses to commit them at
    // all - once resolved and staged they reappear as A or M.
    return {
      staged: true,
      files: gitPaths(['diff', '--cached', '--name-only', '--diff-filter=ACMRTC']),
    };
  }
  return { staged: false, files: process.argv.slice(2) };
}

// Returns the bytes to scan, or a reason string explaining why they could not
// be read. Staged mode reads the index blob rather than the working tree, so
// that editing a secret out of the working tree after `git add` cannot slip it
// past the hook.
function readContent(file, staged) {
  if (staged) {
    try {
      return { bytes: git(['show', `:${file}`]) };
    } catch (err) {
      const detail = String(err?.stderr ?? '').trim().split('\n')[0] || err?.message || 'unknown error';
      return { reason: `staged content unreadable (${detail})` };
    }
  }
  try {
    return { bytes: readFileSync(file) };
  } catch (err) {
    return { reason: `unreadable (${err?.code || err?.message || 'unknown error'})` };
  }
}

const denylist = loadLocalDenylist();
const findings = [];
const skipped = [];
const excluded = [];
let scannedCount = 0;

const { staged, files } = targetFiles();

for (const file of files) {
  if (IGNORE.some((re) => re.test(file))) {
    excluded.push(file);
    continue;
  }
  if (BINARY_ALLOWANCES.has(file)) {
    excluded.push(file);
    continue;
  }

  const { bytes, reason } = readContent(file, staged);
  if (!bytes) {
    skipped.push({ file, reason });
    continue;
  }
  const nulAt = bytes.indexOf(0);
  if (nulAt !== -1) {
    skipped.push({
      file,
      reason: `binary content (NUL byte at offset ${nulAt}) and not declared in BINARY_ALLOWANCES`,
    });
    continue;
  }

  scannedCount += 1;
  const text = bytes.toString('utf8');

  text.split('\n').forEach((line, i) => {
    const lineNo = i + 1;
    const candidates = [];

    for (const rule of RULES) {
      rule.re.lastIndex = 0;
      let m;
      while ((m = rule.re.exec(line)) !== null) {
        if (m[0].length === 0) { rule.re.lastIndex += 1; continue; }
        if (rule.skip && rule.skip(m[0])) continue;
        candidates.push({ index: m.index, length: m[0].length, rule: rule.name, snippet: m[0].slice(0, 40) });
      }
    }

    EMAIL_RE.lastIndex = 0;
    let e;
    while ((e = EMAIL_RE.exec(line)) !== null) {
      const domain = e[0].split('@')[1].toLowerCase();
      if (!isSafeEmailDomain(domain)) {
        candidates.push({ index: e.index, length: e[0].length, rule: 'possible personal email', snippet: e[0] });
      }
    }

    const lowerLine = line.toLowerCase();
    for (const bad of denylist) {
      const at = lowerLine.indexOf(bad.toLowerCase());
      if (at !== -1) {
        candidates.push({ index: at, length: bad.length, rule: 'local denylist match', snippet: bad });
      }
    }

    if (candidates.length === 0) return;
    if (isSuppressed(line, candidates)) return;
    for (const c of candidates) findings.push({ file, lineNo, rule: c.rule, snippet: c.snippet });
  });
}

const coverage = `${scannedCount} file(s) scanned, ${excluded.length} excluded by policy, ${skipped.length} skipped`;

if (skipped.length > 0) {
  console.error(`\n✗ secret/PII scan could not read ${skipped.length} file(s):\n`);
  for (const s of skipped) console.error(`  ${s.file}  ${s.reason}`);
  console.error(`
A file the scanner cannot read is a hole in the gate, so it fails the run
rather than passing quietly. For a missing or unreadable path, restore it or
drop it from the scan. For a file that is legitimately binary and has to stay,
add its exact path to BINARY_ALLOWANCES in scripts/scan-secrets.mjs so the gap
is recorded there.
`);
}

if (findings.length > 0) {
  console.error(`\n✗ secret/PII scan found ${findings.length} issue(s):\n`);
  for (const f of findings) {
    console.error(`  ${f.file}:${f.lineNo}  [${f.rule}]  ${f.snippet}`);
  }
  console.error(`
If a match is a deliberate synthetic fixture (not a real secret), append the
marker "${PRAGMA}" in a comment on that line, with a note saying why it is
safe. If it is a real credential or personal datum, remove it - and if it was
ever committed, rotate/revoke it.
`);
}

if (findings.length > 0 || skipped.length > 0) {
  console.error(`Coverage: ${coverage}.`);
  process.exit(1);
}

console.log(`✓ secret/PII scan clean: ${coverage}`);
process.exit(0);
