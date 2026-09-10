#!/usr/bin/env node
/**
 * Find interpolations rendered inside a single-quoted span with no echo helper.
 *
 * WHAT UNKNOWN THIS SETTLES
 *
 * `docs/conventions.md` states one rule for untrusted values in prose: redact, then
 * neutralise, sanitising the VALUE and never the finished sentence. The drift guard that
 * enforces part of it lives in `src/coerce.test.ts` under `echo-quoting convention`, and it
 * catches a helper rendered inside `'…'` — a value that IS routed through an echo helper but
 * quoted the wrong way round.
 *
 * It cannot be extended to the class this script finds: a value quoted with NO helper at all.
 * The reason is recorded in `docs/conventions.md` and is not a gap in the guard — whether a
 * value is untrusted is a property of where it CAME FROM, so nothing lexical separates
 * `'${name}'` (caller-authored, needs the helper) from `'${mode}'` (a server-side enum, does
 * not). A test that failed on every match would need a per-site suppression list, which is the
 * shape a rule takes when the rule is wrong.
 *
 * So this stays a hand-run inventory rather than a test: it prints the candidates, and a human
 * decides each one by tracing the value to its origin. A non-empty list is NOT a defect list.
 *
 * USAGE
 *
 *   node scripts/find-raw-interpolations.mjs [dir]     # defaults to ./src
 *
 * Prints `file:line  '${expr}'` per candidate, then a count. Always exits 0: there is no
 * threshold that means "wrong", so an exit code would be inventing a verdict it cannot reach.
 *
 * SCOPE AND LIMITS
 *
 * Walks `.ts` files under the directory, skipping `.test.ts` and `.d.ts`. Comments are blanked
 * before matching (block comments are replaced space-for-space so line numbers survive), so a
 * documented example inside a comment does not register. It matches a single-quoted span
 * holding one `${...}` on ONE line; a folded template or a value quoted with a different
 * delimiter is invisible to it. Treat the output as a floor, not a census.
 */
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';

const HELPERS = /describeUntrusted|echoCallerText|describePart|describeDate|echoDate/;

const stripComments = (s) =>
  s.replace(/\/\*[\s\S]*?\*\//g, (b) => b.replace(/[^\n]/g, ' '))
    .split('\n')
    .map((l) => l.replace(/(^|[^:])\/\/.*$/, '$1'))
    .join('\n');

const walk = (dir) => {
  const out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else if (e.endsWith('.ts') && !e.endsWith('.test.ts') && !e.endsWith('.d.ts')) out.push(p);
  }
  return out;
};

// A raw interpolation sitting inside a single-quoted span: '${ ... }' with no echo helper.
const RAW = /'\$\{([^}]*)\}'/g;

let found = 0;
for (const file of walk(process.argv[2] || './src')) {
  const src = stripComments(readFileSync(file, 'utf8'));
  src.split('\n').forEach((line, i) => {
    for (const m of line.matchAll(RAW)) {
      if (!HELPERS.test(m[1])) {
        found++;
        console.log(`${file.replace(/\\/g, '/')}:${i + 1}  '\${${m[1].trim()}}'`);
      }
    }
  });
}
console.log(`\n${found} candidate(s). Each one is decided by tracing the value to its origin, not by this count.`);
