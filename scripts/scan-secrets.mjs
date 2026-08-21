#!/usr/bin/env node
// Secret & PII scanner - blocks credentials and personal information from being
// committed or published. Runs in CI (all tracked files), from the repo's git
// hooks (staged content before a commit, the message of a commit, and every
// commit and tag annotation about to be pushed), and as a release gate before
// the .dxt is packed.
// Self-contained: no third-party dependencies, and it embeds NO personal data -
// personal strings live only in local denylists that are never committed (see
// loadLocalDenylist below and .secret-scan-local.txt.example).
//
// Usage:
//   node scripts/scan-secrets.mjs --all              scan all git-tracked files
//   node scripts/scan-secrets.mjs --staged           scan staged (pre-commit) content
//   node scripts/scan-secrets.mjs --message <file>   scan a commit/tag message (commit-msg hook)
//   node scripts/scan-secrets.mjs --pre-push         scan what a push would publish
//                                                    (reads the pre-push ref lines on stdin)
//   node scripts/scan-secrets.mjs file1 file2        scan specific files
//
// Suppress a known-safe match in FILE content by putting the marker
// allowlist-secret  in a comment on the line (e.g. a synthetic test fixture).
// There is no suppression for messages: a flagged message is rewritten.
// Exit code 1 on any finding, and also on any file the scanner could not read:
// a gate that cannot see a file reports that rather than reporting clean.
//
// Findings name the location and the rule, never the matched value. Printing
// the value would put another copy of the thing being protected into terminal
// scrollback and session transcripts; whoever wrote the text can see it at the
// location given.
//
// The limits of what this covers (git history, untracked files, build output,
// the exempt-domain list, the marker's line-total reach, what a push scan can
// and cannot see) are documented in CONTRIBUTING.md. Keep that section in step
// with this file.

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
  'github.com', 'noreply.github.com', 'users.noreply.github.com',
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

// Australian mobile numbers, the shape that turns up in a signature block.
// Deliberately narrow: a leading 04 or +61 4 and exactly ten digits, so an
// ordinary long number cannot trip it.
const AU_MOBILE_RE = /(?:\+?61[\s-]?4|\b04)\d{2}[\s-]?\d{3}[\s-]?\d{3}\b/g;

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
  { name: 'possible personal phone number', re: AU_MOBILE_RE },
];

const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;

const DENYLIST_RULE = 'local denylist match';

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

function git(args, input) {
  return execFileSync('git', args, { encoding: 'buffer', maxBuffer: MAX_BUFFER, input });
}

function gitText(args) {
  return git(args).toString('utf8');
}

// `git config --get` exits 1 when the key is unset, which is the normal case.
function gitConfig(key) {
  try {
    return gitText(['config', '--get', key]).trim();
  } catch {
    return '';
  }
}

// Local denylists: one literal string per line, case-insensitive substring
// match, never committed. Two sources, both optional:
//   - .secret-scan-local.txt in the repo root (gitignored), the per-clone file
//     described in CONTRIBUTING.md;
//   - the file named by `git config secretscan.denylist`, so one list kept
//     outside every repo can be shared with other tooling on the machine and
//     with every worktree of this clone (local git config is shared by them).
// A source that is configured but cannot be read fails the run: a gate that
// cannot see its own list must say so rather than scan without it.
function loadLocalDenylist() {
  const sources = [];
  if (existsSync('.secret-scan-local.txt')) sources.push('.secret-scan-local.txt');
  const configured = gitConfig('secretscan.denylist');
  if (configured) sources.push(configured);

  const entries = [];
  for (const source of sources) {
    let text;
    try {
      text = readFileSync(source, 'utf8');
    } catch (err) {
      console.error(`\n✗ secret/PII scan: denylist ${source} is configured but could not be read (${err?.code || err?.message || 'unknown error'}).`);
      console.error('A gate that cannot see its own list must not report clean. Fix the path or unset `git config secretscan.denylist`.\n');
      process.exit(1);
    }
    for (const raw of text.split('\n')) {
      const l = raw.trim();
      if (l && !l.startsWith('#')) entries.push(l);
    }
  }
  return entries;
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

// ---------------------------------------------------------------------------
// Scanning
// ---------------------------------------------------------------------------

// Scan one body of text line by line. `where(lineNo)` labels a finding;
// `allowPragma` enables the allowlist-secret marker, which only makes sense
// for file content (a commit message has no comment syntax to carry it).
function scanText(text, where, { allowPragma }, denylist, findings) {
  text.split('\n').forEach((line, i) => {
    const candidates = [];

    for (const rule of RULES) {
      rule.re.lastIndex = 0;
      let m;
      while ((m = rule.re.exec(line)) !== null) {
        if (m[0].length === 0) { rule.re.lastIndex += 1; continue; }
        if (rule.skip && rule.skip(m[0])) continue;
        candidates.push({ index: m.index, length: m[0].length, rule: rule.name });
      }
    }

    EMAIL_RE.lastIndex = 0;
    let e;
    while ((e = EMAIL_RE.exec(line)) !== null) {
      const domain = e[0].split('@')[1].toLowerCase();
      if (!isSafeEmailDomain(domain)) {
        candidates.push({ index: e.index, length: e[0].length, rule: 'possible personal email' });
      }
    }

    const lowerLine = line.toLowerCase();
    for (const bad of denylist) {
      const at = lowerLine.indexOf(bad.toLowerCase());
      if (at !== -1) {
        candidates.push({ index: at, length: bad.length, rule: DENYLIST_RULE, denylist: true });
      }
    }

    if (candidates.length === 0) return;
    if (allowPragma && isSuppressed(line, candidates)) return;
    for (const c of candidates) findings.push({ where: where(i + 1), rule: c.rule, denylist: !!c.denylist });
  });
}

// ---------------------------------------------------------------------------
// What to scan
// ---------------------------------------------------------------------------

// Returns the bytes to scan, or a reason string explaining why they could not
// be read. Staged mode reads the index blob rather than the working tree, so
// that editing a secret out of the working tree after `git add` cannot slip it
// past the hook. Commit mode reads the blob as committed.
function readContent(file, { staged, commit }) {
  if (staged || commit) {
    try {
      return { bytes: git(['show', `${commit || ''}:${file}`]) };
    } catch (err) {
      const detail = String(err?.stderr ?? '').trim().split('\n')[0] || err?.message || 'unknown error';
      return { reason: `${commit ? 'committed' : 'staged'} content unreadable (${detail})` };
    }
  }
  try {
    return { bytes: readFileSync(file) };
  } catch (err) {
    return { reason: `unreadable (${err?.code || err?.message || 'unknown error'})` };
  }
}

// Files changed by one commit. A merge is listed with -c, which names only the
// paths whose merged result differs from every parent (the resolutions), so a
// merge of an already-published branch does not re-scan that branch's history.
function filesInCommit(sha) {
  const parents = gitText(['rev-list', '--parents', '-n', '1', sha]).trim().split(/\s+/).length - 1;
  const args = ['diff-tree', '-r', '--no-commit-id', '--name-only', '--diff-filter=ACMRTC', '--root'];
  if (parents > 1) args.push('-c');
  args.push(sha);
  return gitPaths(args);
}

// What a push would publish, from the `<local ref> <local sha> <remote ref>
// <remote sha>` lines git hands a pre-push hook on stdin. Every commit that is
// not already reachable from some remote-tracking ref is new to the world, so
// its message and content are scanned; an annotated tag is scanned as well,
// because git has no hook that sees a tag message at creation.
function prePushTargets(stdinText) {
  const ZERO = /^0+$/;
  const commits = new Set();
  const tags = [];
  for (const raw of stdinText.split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    const [localRef, localSha] = line.split(/\s+/);
    if (!localSha || ZERO.test(localSha)) continue; // a deletion publishes nothing

    if (localRef.startsWith('refs/tags/') && gitText(['cat-file', '-t', localSha]).trim() === 'tag') {
      const object = gitText(['cat-file', '-p', localSha]);
      const blank = object.indexOf('\n\n');
      tags.push({ name: localRef.slice('refs/tags/'.length), message: blank === -1 ? '' : object.slice(blank + 2) });
    }

    const listed = gitText(['rev-list', localSha, '--not', '--remotes']).trim();
    for (const sha of listed ? listed.split('\n') : []) commits.add(sha);
  }
  return { commits: [...commits], tags };
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

const denylist = loadLocalDenylist();
const findings = [];
const skipped = [];
const excluded = [];
let scannedCount = 0;
let messageCount = 0;

function scanFile(file, opts) {
  if (IGNORE.some((re) => re.test(file))) {
    excluded.push(file);
    return;
  }
  if (BINARY_ALLOWANCES.has(file)) {
    excluded.push(file);
    return;
  }
  const label = opts.commit ? `${opts.commit.slice(0, 7)}:${file}` : file;
  const { bytes, reason } = readContent(file, opts);
  if (!bytes) {
    skipped.push({ file: label, reason });
    return;
  }
  const nulAt = bytes.indexOf(0);
  if (nulAt !== -1) {
    skipped.push({
      file: label,
      reason: `binary content (NUL byte at offset ${nulAt}) and not declared in BINARY_ALLOWANCES`,
    });
    return;
  }
  scannedCount += 1;
  scanText(bytes.toString('utf8'), (n) => `${label}:${n}`, { allowPragma: true }, denylist, findings);
}

function scanMessage(text, label) {
  messageCount += 1;
  scanText(text, (n) => `${label}:${n}`, { allowPragma: false }, denylist, findings);
}

const mode = process.argv[2];
if (mode === '--all') {
  for (const file of gitPaths(['ls-files'])) scanFile(file, { staged: false });
} else if (mode === '--staged') {
  // A staged deletion (D) leaves no content in the commit, so there is
  // nothing to scan; every other status does put content at a path. R and C
  // name the destination path only, which is the content being committed.
  // Unmerged paths (U) are excluded because git refuses to commit them at
  // all - once resolved and staged they reappear as A or M.
  for (const file of gitPaths(['diff', '--cached', '--name-only', '--diff-filter=ACMRTC'])) {
    scanFile(file, { staged: true });
  }
} else if (mode === '--message') {
  const file = process.argv[3];
  let text;
  try {
    text = readFileSync(file, 'utf8');
  } catch (err) {
    console.error(`\n✗ secret/PII scan: message file ${file ?? '(none given)'} could not be read (${err?.code || err?.message || 'unknown error'}).\n`);
    process.exit(1);
  }
  scanMessage(text, 'message');
} else if (mode === '--pre-push') {
  const { commits, tags } = prePushTargets(readFileSync(0, 'utf8'));
  for (const tag of tags) scanMessage(tag.message, `tag ${tag.name}`);
  for (const sha of commits) {
    scanMessage(gitText(['log', '-1', '--format=%B', sha]), `commit ${sha.slice(0, 7)} message`);
    for (const file of filesInCommit(sha)) scanFile(file, { commit: sha });
  }
} else {
  for (const file of process.argv.slice(2)) scanFile(file, { staged: false });
}

const parts = [`${scannedCount} file(s) scanned`];
if (messageCount > 0) parts.push(`${messageCount} message(s) scanned`);
parts.push(`${excluded.length} excluded by policy`, `${skipped.length} skipped`);
const coverage = parts.join(', ');

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
  for (const f of findings) console.error(`  ${f.where}  [${f.rule}]`);
  console.error(`
The matched values are deliberately not shown: printing them would make another
copy of the thing being protected. Open the location named and you will see it.
`);
  if (findings.some((f) => f.denylist)) {
    console.error(`A DENYLIST MATCH MEANS REAL PERSONAL DATA FROM THIS MACHINE'S RECORDS WAS ABOUT
TO BE COMMITTED OR PUBLISHED. ESCALATE TO THE USER FOR TRIAGE: stop, and tell
them what was about to go out and where the text came from. Do not rewrite it
and retry, and do not touch the denylist.
`);
  }
  if (findings.some((f) => !f.denylist)) {
    console.error(`For the other rules: if the match is a deliberate synthetic fixture in a FILE,
append the marker "${PRAGMA}" in a comment on that line with a note saying
why it is safe. In a commit or tag message there is no suppression - rewrite
the message. If it is a real credential or personal datum, remove it, and if it
was ever committed, rotate/revoke it.
`);
  }
}

if (findings.length > 0 || skipped.length > 0) {
  console.error(`Coverage: ${coverage}.`);
  process.exit(1);
}

console.log(`✓ secret/PII scan clean: ${coverage}`);
process.exit(0);
