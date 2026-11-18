#!/usr/bin/env node
// Runs levers: deliberate source edits that must each turn one pinned test red at its own
// assertion (CLAUDE.md, "Building and testing, in practice").
//
//   node scripts/run-levers.mjs levers.json
//
// levers.json is an array of { label, target, find, replace, test, pattern }: `target` and
// `test` are paths relative to the current directory, `find` must occur exactly once in
// `target`, and `pattern` is passed to the test runner's --test-name-pattern. For each
// lever, in order, the edit is written, the one test file runs under tsx --test (the
// runner `npm test` uses), and the original bytes are put back and compared. One line per
// lever, for a commit message: RED with the first failing test and the file:line of its
// failing assertion, GREEN (the pin did not notice the lever), or INVALID with the reason.
// Exits 1 if any lever is GREEN or INVALID, 2 on a usage error or a refusal to start.
//
// It refuses to start if any target has uncommitted changes, so a restore never destroys
// work. A test that already fails without the lever makes the lever INVALID, since its red
// would prove nothing.
//
// This file is also the test reporter the child run loads; it prints one marked JSON line
// per finished test, which is all the parent reads.

import { spawn, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { relative, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const MARK = 'run-levers-event ';
const SELF = fileURLToPath(import.meta.url);

export default async function* reporter(source) {
  for await (const { type, data } of source) {
    if (type !== 'test:pass' && type !== 'test:fail') continue;
    if (data.details?.type === 'suite' || data.skip || data.todo) continue;
    let err = data.details?.error;
    if (err?.code === 'ERR_TEST_FAILURE' && err.cause) err = err.cause;
    yield `${MARK}${JSON.stringify({
      failed: type === 'test:fail',
      name: data.name,
      file: data.file,
      stack: err ? String(err.stack ?? err.message ?? err) : '',
    })}\n`;
  }
}

function runTest(testFile, pattern) {
  const tsxCli = createRequire(import.meta.url).resolve('tsx/cli');
  const args = [tsxCli, '--test', `--test-name-pattern=${pattern}`,
    `--test-reporter=${pathToFileURL(SELF).href}`, '--test-reporter-destination=stdout', testFile];
  return new Promise((done, fail) => {
    // NODE_TEST_CONTEXT, inherited when this runs under a test, would make the child report
    // to a parent runner in its own protocol instead of through this reporter.
    const { NODE_TEST_CONTEXT, ...env } = process.env;
    const child = spawn(process.execPath, args, { stdio: ['ignore', 'pipe', 'pipe'], env });
    current.child = child;
    let out = '';
    child.stdout.on('data', (b) => { out += b; });
    child.stderr.on('data', (b) => { out += b; });
    child.on('error', fail);
    child.on('close', (code) => {
      current.child = null;
      // A file that fails to load is reported as a test named by the file's path.
      const events = out.split('\n').filter((l) => l.startsWith(MARK)).map((l) => JSON.parse(l.slice(MARK.length)))
        .filter((e) => !e.file || resolve(e.name) !== resolve(e.file));
      done({ code, events, out });
    });
  });
}

// The first stack frame in the test file itself, else the first in the repo outside
// node_modules: the line of the assertion, not of the test's declaration.
function assertionSite(stack, testFile) {
  const frames = [...stack.matchAll(/(?:file:\/\/\/?)?((?:[A-Za-z]:)?[\\/][^\s():]*?):(\d+):\d+/g)]
    .map((m) => ({ file: resolve(decodeURIComponent(m[1]).replace(/^\/([A-Za-z]:)/, '$1')), line: m[2] }));
  const inRepo = (f) => !relative(process.cwd(), f.file).startsWith('..') && !f.file.includes(`${sep}node_modules${sep}`);
  const hit = frames.find((f) => f.file === testFile) ?? frames.find(inRepo);
  return hit ? `${relative(process.cwd(), hit.file).split(sep).join('/')}:${hit.line}` : null;
}

const current = { target: null, original: null, child: null };

function restore() {
  if (current.target === null) return;
  writeFileSync(current.target, current.original);
  const back = readFileSync(current.target);
  if (!back.equals(current.original)) throw new Error(`${current.target} is not byte-identical after restore`);
  current.target = null;
}

function onSignal(signal) {
  current.child?.kill();
  restore();
  console.error(`run-levers: ${signal}; target restored`);
  process.exit(130);
}

function count(haystack, needle) {
  let n = 0;
  for (let i = haystack.indexOf(needle); i !== -1; i = haystack.indexOf(needle, i + needle.length)) n++;
  return n;
}

async function runLever(lever, baselines) {
  for (const key of ['label', 'target', 'find', 'replace', 'test', 'pattern']) {
    if (typeof lever[key] !== 'string') return `INVALID - "${key}" is not a string`;
  }
  const target = resolve(lever.target);
  const testFile = resolve(lever.test);
  let original;
  try {
    original = readFileSync(target);
  } catch (e) {
    return `INVALID - cannot read ${lever.target}: ${e.code ?? e.message}`;
  }
  const find = Buffer.from(lever.find, 'utf8');
  const hits = count(original, find);
  if (find.length === 0 || hits !== 1) return `INVALID - find occurs ${hits} times in ${lever.target}, not once`;

  const key = `${testFile}\0${lever.pattern}`;
  if (!baselines.has(key)) baselines.set(key, await runTest(testFile, lever.pattern));
  const base = baselines.get(key);
  if (base.events.length === 0) return `INVALID - no test in ${lever.test} matches "${lever.pattern}"`;
  if (base.code !== 0) return 'INVALID - the test fails without the lever';

  const at = original.indexOf(find);
  const changed = Buffer.concat([original.subarray(0, at), Buffer.from(lever.replace, 'utf8'), original.subarray(at + find.length)]);
  let run;
  current.target = target;
  current.original = original;
  try {
    writeFileSync(target, changed);
    run = await runTest(testFile, lever.pattern);
  } finally {
    restore();
  }
  if (run.code === 0) return 'GREEN - the test still passes with the lever in place';
  // RED only for a test failing at a located line: a lever that breaks loading proves no pin.
  const first = run.events.find((e) => e.failed);
  if (!first) return `INVALID - the run failed (exit ${run.code}) but no matching test did; the lever may break loading`;
  const site = assertionSite(first.stack, testFile);
  if (!site) return `INVALID - "${first.name}" failed with no stack frame in the repo`;
  return `RED - "${first.name}" at ${site}`;
}

async function main(argv) {
  if (argv.length !== 1) {
    console.error('usage: node scripts/run-levers.mjs <levers.json>');
    return 2;
  }
  let levers;
  try {
    levers = JSON.parse(readFileSync(argv[0], 'utf8'));
  } catch (e) {
    console.error(`run-levers: cannot read ${argv[0]}: ${e.message}`);
    return 2;
  }
  if (!Array.isArray(levers) || levers.length === 0) {
    console.error('run-levers: the file must hold a non-empty array of levers');
    return 2;
  }
  const targets = [...new Set(levers.map((l) => l?.target).filter((t) => typeof t === 'string'))];
  if (targets.length > 0) {
    const st = spawnSync('git', ['status', '--porcelain', '--', ...targets], { encoding: 'utf8' });
    if (st.status !== 0) {
      console.error(`run-levers: git status failed: ${st.stderr.trim()}`);
      return 2;
    }
    if (st.stdout.trim() !== '') {
      console.error(`run-levers: refusing to start, uncommitted changes in a target:\n${st.stdout.trimEnd()}`);
      return 2;
    }
  }
  process.on('SIGINT', onSignal);
  process.on('SIGTERM', onSignal);

  const baselines = new Map();
  let bad = false;
  for (const lever of levers) {
    const verdict = await runLever(lever ?? {}, baselines);
    if (!verdict.startsWith('RED')) bad = true;
    console.log(`${lever?.label ?? '(no label)'}: ${verdict}`);
  }
  return bad ? 1 : 0;
}

if (process.argv[1] && resolve(process.argv[1]) === SELF) {
  process.exitCode = await main(process.argv.slice(2));
}
