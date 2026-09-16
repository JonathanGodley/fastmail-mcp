#!/usr/bin/env node
// Mutation-test only the source lines a single commit changed.
//
// Asks, for every line a commit touched, "would any test notice if this line did
// something else?". A survived mutant is a line no test covers, which is the same
// question as hand-reverting a change to watch its new test fail - asked of every
// changed line at once instead of one test at a time.
//
// Usage: node scripts/mutate-commit.mjs <commit> [--tests src/a.test.ts,src/b.test.ts]
//        node scripts/mutate-commit.mjs <commit> [--tests=src/a.test.ts,src/b.test.ts]
//
// Test files are chosen BY NAME by default: for each changed src/X.ts, every src/X*.test.ts
// that exists, plus any test file the commit itself changed. `--tests` overrides that, which
// is needed when the list it would otherwise pick includes a test Stryker cannot run.
//
// Two different things push a test out of the by-name list here, and they do not fail the same
// way: one aborts the initial run outright; the other, for one of its two members, does not fail
// at all - it just runs against the wrong code and can never register a kill.
//
// The first is a src/*.ts file getting Stryker's instrumentation AT ALL. Whichever file actually
// receives a mutant is reprinted WHOLE by Stryker's TypeScript printer (@babel/generator over the
// entire file, not just the mutated node), and every instrumented file also gets a header
// prepended that itself reads process.env (to learn which mutant is active). Five tests are
// affected, in three different ways - each measured directly:
//   - index-env.test.ts scans every non-test src/*.ts file for a bare process.env read, and
//     fails the moment ANY of them carries Stryker's own env-reading header - not because its
//     own `function findEnvValue(` text match ever breaks. Measured against a commit touching
//     only src/index.ts and, separately, one touching only src/auth.ts: both fail from the same
//     cause, process.env.__STRYKER_ACTIVE_MUTANT__ at the mutated file's own top lines (the
//     src/index.ts case also fails a second assertion, because the reprint moves the closing
//     brace that ends findEnvValue's located body range). It therefore fails on every run that
//     selects it, whichever src/*.ts file that commit touches - there is no line to avoid
//     mutating that keeps it green.
//   - readme-inventory.test.ts and tool-schema.test.ts each locate a piece of src/index.ts's
//     structure by matching exact text (the `const TOOLS = [` array; tool-schema.test.ts also
//     scans recipient-parameter descriptions for a named constant) rather than by parsing it, so
//     the whole-file reprint above can disturb their match wherever in the file a mutant lands,
//     not only beside it. Measured: a one-line mutation on the Server constructor's `name`
//     field, nowhere near either scan's target, still broke both - readme-inventory.test.ts's
//     search for `const TOOLS = [`, and, separately, tool-schema.test.ts's count of
//     recipient-parameter descriptions.
//   - config-surface.test.ts and jmap-client.test.ts's "version sync" check each match one
//     narrow literal (a findEnvValue([...]) argument list; the `version: '...'` line) and fail
//     only when that text is itself what a mutant lands on, or what the reprint happens to
//     reformat - not on every src/index.ts change. The same name-field mutation above left both
//     passing (their own mutated line was just an ordinary, unrelated survivor); mutating a
//     manifest-mapped findEnvValue() call, by contrast, does break config-surface.test.ts's
//     initial run.
// Of these five, index-env.test.ts and jmap-client.test.ts are the two the by-name rule above
// ever auto-selects on its own (for a change to src/index.ts and to src/jmap-client.ts
// respectively). Only jmap-client.test.ts is also that module's real unit-test suite; the other
// four - index-env.test.ts included - exist for no reason but the guard. A commit touching both
// src/index.ts and src/jmap-client.ts auto-selects both: index-env.test.ts must still be dropped
// (it fails unconditionally, above), but dropping jmap-client.test.ts too loses real coverage of
// src/jmap-client.ts's own changed lines for a version-sync failure that may not even occur in
// that run - check whether it actually fails before excluding the whole file.
//
// The second is READING dist/index.js. dist/ being gitignored does not keep it out of a
// sandbox: Stryker's own file-copy step has no .gitignore handling at all, only a short
// hardcoded ignore list (node_modules, .git and a few framework build directories), so a dist/
// that exists on disk when a run starts is copied in whole - measured by the sandboxed file
// count rising from 131 to 218 once `npm run build` had been run first. What decides the
// outcome is whether the tree was built before the run, and the two dist-reading tests react to
// that differently:
//   - built-server.test.ts's `before` hooks fail the initial run outright either way, for two
//     different reasons depending on whether dist/ existed: with no dist/ present, "dist/index.js
//     does not exist"; with a freshly built dist/ present, the mtime comparison itself fails
//     instead, because Stryker's sandbox copy does not preserve source timestamps, so a dist/
//     built moments earlier can still compare as older than a src file the copy touched
//     afterward. Either way it fails, every time, for every commit, whatever is mutated.
//   - server-lifecycle.test.ts only skips its whole suite when dist/ is absent; with dist/
//     present it actually runs (measured: its dry run succeeds) and spawns the server that was
//     last built, so it can never contribute a kill against the mutated source either way - the
//     same reason a change to src/index.ts itself cannot be measured, below.
//
// A change to src/index.ts specifically cannot be meaningfully mutation-tested by this script:
// it holds the tool schema literal and the CallTool switch, which has no test harness of its
// own (see CLAUDE.md § Testing), so every test that pins its content either IS one of the five
// text-parsing guards above, or is one of the two dist-reading tests above (which either fails
// outright or exercises the last `npm run build`, not the mutated source). `--tests` is no cure
// for either. A single past observation, not reproduced in this pass and not to be re-measured:
// `abc1d9f` (a real index.ts change), run with `--tests` naming only files outside both classes
// above, scored 5 mutants, 5 survived - every one vacuously.
//
// That accounts for two of the three ways the initial run can fail. The third announces itself
// as a MISSING PACKAGE ("Cannot find package '@modelcontextprotocol/sdk'"), which is the
// sandbox having no node_modules to resolve against - no choice of test files changes it.
// DEPS_ROOT below is what stops that happening; read its comment if it recurs.
//
// Runs from a `git worktree` as well as from the primary checkout. A worktree contributes the
// FILES to mutate; the primary checkout contributes the installed DEPENDENCIES.
//
// Refuses to run unless the commit named is REPO's current HEAD and REPO's tree is clean
// (`git status --porcelain` reports nothing) - see the refusal itself for why and for the
// cure. A non-tip commit of a multi-commit branch needs a detached checkout (or a worktree)
// first. A merge commit is diffed against its first parent only (`${rev}~1`) - not refused.
//
// Writes the Stryker config, sandbox and report OUTSIDE the repo tree (os.tmpdir by
// default; override with MUTATE_OUT_DIR). Re-running overwrites; nothing needs deleting.
//
// The node_modules link this makes OUTLIVES the run: it sits beside the sandbox rather than
// inside it, so nothing wipes it between runs. A copy of this script WITHOUT that link step
// will therefore appear to work in a scratch directory some other copy has already used -
// point MUTATE_OUT_DIR at a fresh one before trusting a comparison of two versions of it.

import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, rmdirSync, symlinkSync, unlinkSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT_DIR = path.resolve(process.env.MUTATE_OUT_DIR || path.join(tmpdir(), 'fastmail-mcp-mutate'));
const CONFIG_FILE = path.join(OUT_DIR, 'stryker.config.json');
const REPORT_FILE = path.join(OUT_DIR, 'mutation-report.json');

const argv = process.argv.slice(2);
const USAGE = 'Usage: node scripts/mutate-commit.mjs <commit> '
  + '[--tests src/a.test.ts,src/b.test.ts | --tests=src/a.test.ts,src/b.test.ts]';

// `--tests` takes its value as a following argument OR as `--tests=...`; both forms are
// consumed here so neither leaves its raw token in `rest` to be mistaken for the commit.
let testsFlagPresent = false;
let testsRaw = null;
const rest = [];
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--tests') {
    testsFlagPresent = true;
    testsRaw = argv[i + 1] ?? '';
    i++;
  } else if (a.startsWith('--tests=')) {
    testsFlagPresent = true;
    testsRaw = a.slice('--tests='.length);
  } else {
    rest.push(a);
  }
}
const testsOverride = testsFlagPresent ? testsRaw.split(',').filter(Boolean) : null;

if (testsFlagPresent && testsOverride.length === 0) {
  console.error(USAGE);
  console.error('--tests was given no value. Pass a comma-separated list, e.g. --tests=src/a.test.ts,src/b.test.ts.');
  process.exit(2);
}
if (rest.length !== 1) {
  console.error(USAGE);
  if (testsFlagPresent) {
    console.error(
      `Expected one commit after --tests, found ${rest.length}: ${rest.join(' ') || '(none)'}. `
      + 'Pass every test file as one comma-separated --tests value, not as separate arguments.',
    );
  } else if (rest.length === 0) {
    console.error('A commit is required.');
  } else {
    console.error(`Expected exactly one commit argument, got ${rest.length}: ${rest.join(' ')}.`);
  }
  process.exit(2);
}
const commit = rest[0];

const git = (...args) => execFileSync('git', args, { cwd: REPO, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });

// Peels an annotated tag to the commit it names, so a branch, a short SHA and an annotated
// tag all compare equal to what `git status`/HEAD report. Returns null rather than throwing,
// so a rev that does not resolve is refused below with its own name rather than a git
// stack trace. `--verify --quiet` is load-bearing, not decoration: without it, a rev
// starting with `-` (e.g. `--help`) is read by `git rev-parse` as an option rather than an
// argument and echoed back instead of failing, and a plain failure prints git's own
// "fatal: ambiguous argument" noise to stderr before this function ever gets to react to it.
function resolveCommit(rev) {
  try {
    return git('rev-parse', '--verify', '--quiet', '--end-of-options', `${rev}^{commit}`).trim();
  } catch {
    return null;
  }
}

const headSha = resolveCommit('HEAD');
const resolvedCommit = resolveCommit(commit);
if (!resolvedCommit) {
  console.error(`${commit} does not resolve to a commit in ${REPO}.`);
  process.exit(2);
}

// Refuse unless the tree Stryker is about to copy is exactly the commit being measured.
// Stryker sandboxes whatever it finds on disk, not the named commit's own snapshot: a dirty
// mutated source shifts which lines are "changed", a dirty test earns kills the commit does
// not contain, and a dirty file either of those imports does the same at one remove. The
// check is WHOLE-TREE, not just the changed files, for the same reason.
const headMismatch = resolvedCommit !== headSha;
const status = git('status', '--porcelain').trim();
const dirty = status !== '';
if (headMismatch || dirty) {
  console.error(`Refusing to mutate ${REPO}: it would not describe ${commit}.`);
  const cures = [];
  if (headMismatch) {
    console.error(`${commit} resolves to ${resolvedCommit}, but HEAD is ${headSha}.`);
    cures.push(`check ${commit} out detached (or in a worktree) and re-run from there`);
  }
  if (dirty) {
    console.error(`The tree is not clean (git status --porcelain):\n${status}`);
    cures.push('commit the changes and re-run naming the new HEAD');
    cures.push('stash the changes INCLUDING UNTRACKED FILES (`git stash -u` - a bare `git stash` leaves untracked files behind and this refusal keeps firing)');
  }
  console.error(`Cure: ${cures.join('; or ')}.`);
  process.exit(2);
}

// A root commit resolves fine above but has no parent, so `${rev}~1` below has nothing to
// diff against. Checked only once HEAD/clean has already passed, since a root commit that is
// not also HEAD is refused by that check first.
if (!resolveCommit(`${resolvedCommit}~1`)) {
  console.error(`${commit} (${resolvedCommit}) is a root commit - it has no parent to diff against, so its changed lines cannot be computed.`);
  process.exit(2);
}

// Observable without a dependency tree: this is everything --tests resolved to, printed
// before DEPS_ROOT below can refuse the run for unrelated reasons.
if (testsOverride) console.log(`tests (--tests) ${testsOverride.join(', ')}`);

// The checkout whose node_modules the sandboxed tests resolve against. REPO holds the files to
// mutate; it does NOT necessarily hold the dependencies. `npm install` is run once, in the
// primary checkout, and a `git worktree` checkout has no node_modules of its own.
//
// Stryker links a node_modules into its sandbox only if it finds one by scanning DOWNWARD from
// its own cwd (findNodeModulesList), and cwd must stay REPO or Stryker would mutate the wrong
// checkout's files. So from a worktree it links nothing, the sandbox sits outside every tree,
// and Node's walk-up from a sandboxed src/*.ts reaches no node_modules at all: every test dies
// on "Cannot find package". Naming the primary checkout here is what closes that gap.
//
// `--path-format=absolute` is load-bearing, not decoration: plain `--git-common-dir` answers a
// RELATIVE ".git" from a primary checkout (it answers an absolute path from a worktree, so the
// bug hides in exactly the case that looks fine). path.dirname of that is ".", which then reads
// node_modules relative to whatever cwd the caller happened to have.
function primaryCheckout() {
  try {
    return path.dirname(git('rev-parse', '--path-format=absolute', '--git-common-dir').trim());
  } catch {
    return null; // not a repo, or a git too old for --path-format: fall back to REPO alone
  }
}

// REPO FIRST, so a primary checkout resolves to itself and runs exactly as it always has - the
// primary checkout is a fallback, never an override. It also means a worktree someone HAS run
// `npm install` in uses its own dependencies rather than reaching for another tree's.
const DEPS_ROOT = [REPO, primaryCheckout()].find((dir) => dir && existsSync(path.join(dir, 'node_modules')));
if (!DEPS_ROOT) {
  console.error(`No node_modules in ${REPO}, nor in the primary checkout of its repository.`);
  console.error('Run `npm install` in the primary checkout; a worktree does not need its own.');
  process.exit(1);
}

/**
 * Point `link` at `target`, replacing it if it already points elsewhere.
 *
 * 'junction' is the only directory link Windows creates without elevation; on POSIX Node
 * ignores the type and writes an ordinary symlink. Removal is unlink-then-rmdir because
 * neither call follows a link: unlink takes a POSIX symlink, rmdir takes a Windows junction,
 * and rmdir on a REAL populated directory fails ENOTEMPTY rather than deleting someone's
 * node_modules. Nothing here is ever recursive.
 */
function linkDir(link, target) {
  try {
    if (realpathSync(link) === realpathSync(target)) return; // already correct: the re-run case
  } catch { /* absent or dangling - fall through and (re)create it */ }
  try { unlinkSync(link); } catch { try { rmdirSync(link); } catch { /* nothing to replace */ } }
  symlinkSync(target, link, 'junction');
}

/** Changed line ranges per file, from a zero-context diff of the commit against its parent. */
function changedRanges(rev) {
  const diff = git('diff', `${rev}~1`, rev, '-U0', '--', 'src');
  const ranges = new Map();
  let file = null;
  for (const line of diff.split('\n')) {
    // A REAL header only: a "b/" path, or /dev/null for a file the commit DELETED. Testing
    // for the "+++ " prefix alone also matched an ADDED CONTENT line whose own text begins
    // "++ ", which cleared the file being parsed and dropped its remaining hunks. An added
    // line reading exactly "++ b/<path>" would still fool this; a source line of that shape
    // is not worth tracking `diff --git` state to exclude.
    const header = /^\+\+\+ (b\/(.+)|\/dev\/null)$/.exec(line);
    if (header) {
      // Reset on the deleted file rather than falling through, so its hunks cannot be
      // attributed to whichever file was named before it.
      file = header[2] === undefined ? null : header[2].trim();
      continue;
    }
    const hunk = /^@@ -\S+ \+(\d+)(?:,(\d+))? @@/.exec(line);
    if (!hunk || !file) continue;
    const start = Number(hunk[1]);
    const count = hunk[2] === undefined ? 1 : Number(hunk[2]);
    if (count === 0) continue; // pure deletion: nothing left in the new file to mutate
    if (!ranges.has(file)) ranges.set(file, []);
    ranges.get(file).push([start, start + count - 1]);
  }
  return ranges;
}

const isTest = (f) => f.endsWith('.test.ts');
const isSource = (f) => f.startsWith('src/') && f.endsWith('.ts') && !isTest(f);

const ranges = changedRanges(resolvedCommit);
const sources = [...ranges.keys()].filter(isSource).sort();
if (sources.length === 0) {
  console.error(`${commit} changed no non-test src/*.ts file - nothing to mutate.`);
  process.exit(1);
}

// Test selection: for each changed src/X.ts, every src/X*.test.ts that exists, plus any
// test file the commit itself changed.
const srcFiles = readdirSync(path.join(REPO, 'src'));
const tests = new Set([...ranges.keys()].filter(isTest));
for (const source of sources) {
  const stem = path.basename(source, '.ts');
  for (const candidate of srcFiles) {
    if (candidate.startsWith(stem) && candidate.endsWith('.test.ts')) tests.add(`src/${candidate}`);
  }
}
const testFiles = testsOverride ?? [...tests].sort();
if (testFiles.length === 0) {
  console.error(`No test file matches ${sources.join(', ')} - every mutant would survive by default.`);
  process.exit(1);
}

const mutate = sources.flatMap((f) => ranges.get(f).map(([a, b]) => `${f}:${a}-${b}`));

mkdirSync(OUT_DIR, { recursive: true });

// Give every sandbox a node_modules to walk up into. This sits one level ABOVE the sandboxes
// Stryker creates (tempDirName is OUT_DIR/sandbox, a sandbox is OUT_DIR/sandbox/sandbox-XXXX),
// so Node resolves through it from any sandboxed file without anyone predicting that random
// name. It is a SIBLING of tempDirName rather than a child, which is deliberate: `cleanTempDir`
// wipes tempDirName between runs, and a link inside it would put the real node_modules on the
// far end of a directory tree Stryker deletes.
//
// Stryker's own symlinking is left ON. From the primary checkout it still fires and still wins
// inside the sandbox, so that path is unchanged; this link is what the worktree case falls back
// to, and is simply unused when Stryker has already done the job. It is not pure redundancy
// even there: Stryker's scan keeps only entries whose Dirent isDirectory(), so a node_modules
// that is ITSELF a link - a checkout set up to share one install - is invisible to it, and this
// link is then the only one.
linkDir(path.join(OUT_DIR, 'node_modules'), path.join(DEPS_ROOT, 'node_modules'));

writeFileSync(CONFIG_FILE, JSON.stringify({
  $schema: './node_modules/@stryker-mutator/core/schema/stryker-schema.json',
  packageManager: 'npm',
  testRunner: 'command',
  commandRunner: { command: `npx tsx --test ${testFiles.join(' ')}` },
  coverageAnalysis: 'off', // the command runner reports one exit code, not per-test coverage
  mutate,
  // Everything Stryker writes goes to OUT_DIR: the sandbox copies (tempDirName) and the
  // machine-readable report. 'html' is left out because it writes into reports/ in the repo.
  tempDirName: path.join(OUT_DIR, 'sandbox'),
  cleanTempDir: 'always',
  reporters: ['progress', 'json'],
  jsonReporter: { fileName: REPORT_FILE },
  timeoutMS: 60000,
  // Point Stryker's tsconfig rewriter at a file that does not exist, so it skips. It rewrites
  // `extends`/`references` paths that would break in the sandbox; this repo's tsconfig has
  // neither, and the rewriter calls a TypeScript API that typescript@7 no longer exposes
  // (`ts.parseConfigFileTextToJson`), which aborts the whole run.
  tsconfigFile: 'tsconfig.stryker-skip.json',
}, null, 2) + '\n');

console.log(`commit      ${commit}`);
console.log(`mutating    ${mutate.join('\n            ')}`);
console.log(`tests       ${testFiles.join(' ')}`);
console.log(`scratch     ${OUT_DIR}`);
console.log(`mutating in ${REPO}`);
console.log(`deps from   ${DEPS_ROOT}${DEPS_ROOT === REPO ? '' : ' (primary checkout of this worktree)'}\n`);

// Blank the previous report BEFORE the run, so a Stryker that dies without writing one is
// reported as a failure instead of silently re-reporting the last run's numbers.
writeFileSync(REPORT_FILE, '');

// NO SHELL, and not through `npx`. `npx` on Windows is a .cmd, so running it needed
// `shell: true`, and under a shell the config path went through UNQUOTED: a MUTATE_OUT_DIR
// containing a space arrived as two arguments and Stryker refused the run with "too many
// arguments for 'run'". Naming npx.cmd directly instead is not an option either - Node
// refuses to spawn a .cmd without a shell (EINVAL). So the CLI is resolved here and handed to
// this same Node binary, which takes its arguments as an array and never re-splits them.
// Resolved from DEPS_ROOT, not REPO: from a worktree that is not nested inside the primary
// checkout, a REPO-based resolve has no node_modules anywhere up its chain and throws here,
// before Stryker is ever started.
const strykerPkg = createRequire(path.join(DEPS_ROOT, 'package.json')).resolve('@stryker-mutator/core/package.json');
const strykerBin = path.join(path.dirname(strykerPkg), JSON.parse(readFileSync(strykerPkg, 'utf8')).bin.stryker);

const run = spawnSync(process.execPath, [strykerBin, 'run', CONFIG_FILE], {
  cwd: REPO,
  stdio: ['ignore', 'inherit', 'inherit'],
});

let report;
try {
  report = JSON.parse(readFileSync(REPORT_FILE, 'utf8'));
} catch {
  // Deliberately NOT a diagnosis. Stryker's own output has streamed to this terminal already,
  // so the reason is on screen; naming one cause as though it were the cause sent a
  // space-in-the-path failure off after a test-selection problem it did not have.
  const how = run.error ? `could not start it: ${run.error.message}`
    : run.signal ? `killed by ${run.signal}` : `exit status ${run.status}`;
  console.error(`\nStryker produced no report at ${REPORT_FILE} (${how}).`);
  console.error('Its output above says why. Most often one of the selected tests reads src/*.ts');
  console.error('as text and is seeing Stryker\'s instrumentation; --tests drops that test - except');
  console.error('for a src/index.ts change, where no --tests selection avoids this (see the header).');
  process.exit(run.status || 1);
}

const trim = (s, n) => {
  const one = String(s ?? '').replace(/\s+/g, ' ').trim();
  return one.length > n ? `${one.slice(0, n)}...` : one;
};

const counts = {};
const survivors = [];
for (const [file, entry] of Object.entries(report.files || {})) {
  const lines = (entry.source || '').split('\n');
  for (const m of entry.mutants || []) {
    counts[m.status] = (counts[m.status] || 0) + 1;
    if (m.status !== 'Survived' && m.status !== 'NoCoverage') continue;
    const { start, end } = m.location;
    // Two literals on one line mutate to the same text, so print what was replaced as well.
    const original = (lines[start.line - 1] || '').slice(start.column - 1, end.line === start.line ? end.column - 1 : undefined);
    survivors.push({ file, line: start.line, column: start.column, mutator: m.mutatorName, status: m.status, original, replacement: m.replacement });
  }
}

const total = Object.values(counts).reduce((a, b) => a + b, 0);
const killed = (counts.Killed || 0) + (counts.Timeout || 0);
const scored = killed + (counts.Survived || 0) + (counts.NoCoverage || 0);
const score = scored === 0 ? 'n/a' : `${((killed / scored) * 100).toFixed(2)}%`;

console.log(`\n=== ${total} mutants: ${Object.entries(counts).map(([k, v]) => `${k} ${v}`).join(', ')} ===`);
console.log(`Mutation score ${score}\n`);

if (survivors.length === 0) {
  console.log('No survivors: every mutable line this commit changed is noticed by a test.');
} else {
  console.log(`${survivors.length} survived - each is a line no selected test notices:\n`);
  for (const s of survivors.sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line || a.column - b.column)) {
    console.log(`  ${s.file}:${s.line}:${s.column}  ${s.mutator} [${s.status}]`);
    console.log(`      ${trim(s.original, 120)}`);
    console.log(`      -> ${trim(s.replacement, 120)}`);
  }
}

process.exit(survivors.length > 0 ? 1 : run.status || 0);
