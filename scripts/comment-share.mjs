#!/usr/bin/env node
// Comment-share report: says when a change carries far more comment than the
// file it lands in. A report, never a gate - it always exits 0, and prints
// nothing when nothing is over the bar, so the pre-commit hook stays quiet on
// an ordinary commit.
//
// Usage:
//   node scripts/comment-share.mjs --staged      the index against HEAD (what pre-commit sees)
//   node scripts/comment-share.mjs --working     the working tree against HEAD, untracked files included
//   node scripts/comment-share.mjs <commit>      that commit against its first parent
//   node scripts/comment-share.mjs <a>..<b>      a range
//   add --all to print every measured file, not only the ones over the bar
//
// What it measures: among the lines a change ADDED to each code file, how many
// are comment and how many are code, against the same ratio for the file as it
// stood before the change. C-like files count `//`, `/*` and `*` lines; hash
// families (.py, .sh, .ps1, ...) count `#` lines. Markdown and JSON are not
// measured. The bar: at least 20 added comment lines (fewer is one doc comment,
// never the problem), and either no added code, or a ratio at least twice the
// file's own, that baseline floored at 0.25 so a sparse file still has a bar
// it can meet; a file with no prior version fires at 1.0. Twice the file's own
// density is where the #194 free/busy change sat (3.6 against 1.4) and where
// its trimmed version did not (2.1).
//
// The same measurement runs live in every session from the global PostToolUse
// hook; this is the checked-in copy for the pre-commit hook and for reviewing a
// commit after the fact.

import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { extname, join } from 'node:path';

const MIN_COMMENT = 20;
const OVER = 2;
const FLOOR = 0.25;
const NEW_FILE_BAR = 1;

const C_LIKE = new Set([
  '.ts', '.tsx', '.mts', '.cts', '.js', '.jsx', '.mjs', '.cjs', '.go', '.rs',
  '.java', '.kt', '.cs', '.c', '.h', '.cpp', '.hpp', '.swift', '.scala', '.php',
  '.css', '.scss',
]);
const HASH = new Set([
  '.py', '.sh', '.bash', '.ps1', '.psm1', '.rb', '.pl', '.yml', '.yaml', '.toml',
  '.r', '.jl',
]);

function family(file) {
  const ext = extname(file).toLowerCase();
  if (C_LIKE.has(ext)) return 'c';
  if (HASH.has(ext)) return 'hash';
  return null;
}

function tally(lines, fam) {
  const t = { comment: 0, code: 0 };
  for (const raw of lines) {
    const line = raw.trim();
    if (line === '') continue;
    const isComment = fam === 'c'
      ? (line.startsWith('//') || line.startsWith('/*') || line.startsWith('*'))
      : line.startsWith('#');
    if (isComment) t.comment += 1; else t.code += 1;
  }
  return t;
}

function git(args) {
  return execFileSync('git', args, {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'ignore'],
    windowsHide: true,
    maxBuffer: 64 * 1024 * 1024,
  });
}

function gitOrNull(args) {
  try {
    return git(args);
  } catch {
    return null;
  }
}

const addedLines = (diff) => diff.split(/\r?\n/)
  .filter((l) => l.startsWith('+') && !l.startsWith('+++'))
  .map((l) => l.slice(1));

const names = (out) => (out ?? '').split('\0').filter((f) => f.length);

/**
 * Resolve the mode to: the changed files, a function giving each file's added
 * lines, and a function giving its prior content (null when it had none).
 */
function scope(args) {
  const flags = new Set(args.filter((a) => a.startsWith('--')));
  const rest = args.filter((a) => !a.startsWith('--'));
  const all = flags.has('--all');

  if (flags.has('--staged')) {
    return {
      all,
      files: names(gitOrNull(['diff', '--cached', '--name-only', '-z', 'HEAD', '--'])),
      added: (f) => addedLines(gitOrNull(['diff', '--cached', '--no-color', '--no-ext-diff', '-U0', 'HEAD', '--', f]) ?? ''),
      before: (f) => gitOrNull(['show', 'HEAD:' + f]),
    };
  }
  if (flags.has('--working')) {
    const tracked = names(gitOrNull(['diff', '--name-only', '-z', 'HEAD', '--']));
    const untracked = names(gitOrNull(['ls-files', '--others', '--exclude-standard', '-z']));
    const root = (gitOrNull(['rev-parse', '--show-toplevel']) ?? '.').trim();
    const untrackedSet = new Set(untracked);
    return {
      all,
      files: Array.from(new Set(tracked.concat(untracked))),
      added: (f) => {
        if (untrackedSet.has(f)) {
          try { return readFileSync(join(root, f), 'utf8').split(/\r?\n/); } catch { return []; }
        }
        return addedLines(gitOrNull(['diff', '--no-color', '--no-ext-diff', '-U0', 'HEAD', '--', f]) ?? '');
      },
      before: (f) => (untrackedSet.has(f) ? null : gitOrNull(['show', 'HEAD:' + f])),
    };
  }
  if (rest.length === 1) {
    const range = rest[0].includes('..') ? rest[0].split('..') : [rest[0] + '~1', rest[0]];
    const [from, to] = range;
    return {
      all,
      files: names(gitOrNull(['diff', '--name-only', '-z', from, to, '--'])),
      added: (f) => addedLines(gitOrNull(['diff', '--no-color', '--no-ext-diff', '-U0', from, to, '--', f]) ?? ''),
      before: (f) => gitOrNull(['show', from + ':' + f]),
    };
  }
  return null;
}

function verdict(added, base) {
  let baseRatio = null;
  let baseLabel = 'new file, no baseline';
  if (base !== null) {
    if (base.code > 0) {
      baseRatio = base.comment / base.code;
      baseLabel = 'the file runs ' + baseRatio.toFixed(1);
    } else {
      baseLabel = 'the file has no code before this change';
    }
  }
  if (added.comment < MIN_COMMENT) return { over: false, baseLabel };
  if (added.code === 0) return { over: true, baseLabel };
  const bar = baseRatio === null ? NEW_FILE_BAR : OVER * Math.max(baseRatio, FLOOR);
  return { over: added.comment / added.code >= bar, baseLabel };
}

function describe(file, added, v) {
  const figures = added.code === 0
    ? '+' + added.comment + ' comment lines and no code'
    : '+' + added.comment + ' comment lines against +' + added.code + ' code ('
      + (added.comment / added.code).toFixed(1) + ' per code line';
  const close = added.code === 0 ? ' (' + v.baseLabel + ')' : '; ' + v.baseLabel + ')';
  return (v.over ? 'comment-share: ' : 'comment-share (under the bar): ') + file + ': ' + figures + close + '.';
}

function main(argv) {
  const s = scope(argv);
  if (s === null) {
    process.stderr.write('usage: comment-share.mjs --staged | --working | <commit> | <a>..<b>  [--all]\n');
    return 0;
  }
  const lines = [];
  for (const file of s.files) {
    const fam = family(file);
    if (fam === null) continue;
    const added = tally(s.added(file), fam);
    if (added.comment === 0 && added.code === 0) continue;
    const prior = s.before(file);
    const v = verdict(added, prior === null ? null : tally(prior.split(/\r?\n/), fam));
    if (v.over || s.all) lines.push(describe(file, added, v));
  }
  if (lines.length) {
    process.stdout.write(lines.join('\n') + '\n');
    if (lines.some((l) => l.startsWith('comment-share: '))) {
      process.stdout.write('Worth a /tidy-comments pass: a figure far above the file\'s own is nearly always restatement or history, not domain complexity.\n');
    }
  }
  return 0;
}

process.exit(main(process.argv.slice(2)));
