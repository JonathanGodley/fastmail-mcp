// Checks the relative links in README.md, CONTRIBUTING.md, docs/*.md and docs/reference/*.md:
// each must name a file that exists, and a #anchor on a .md target must match one of its
// headings (GitHub's slug rule, with -1, -2 for repeats) or an <a name> / <a id>. Links with
// a URL scheme are skipped, as is anything inside a ``` fence or an inline code span.
// An anchor on a target that is not .md is reported but does not fail the run.
//
// Not handled: reference-style links, ~~~ fences and setext (underlined) headings.
//
// Usage: node scripts/check-doc-links.mjs [root]. Exits 1 when any link is bad.

import { existsSync, readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const FENCE = /^\s*```/;
const HEADING = /^(#{1,6})\s+(.*?)\s*#*\s*$/;
const HTML_ANCHOR = /<a\s+(?:name|id)="([^"]+)"/g;
const LINK = /\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g;
const SCHEME = /^[a-z]+:/;

function slug(heading) {
  return heading
    .trim()
    .toLowerCase()
    .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/[^\p{L}\p{N} _-]/gu, '')
    .replaceAll(' ', '-');
}

/** Lines outside ``` fences, each with its 1-based line number. */
function* unfencedLines(file) {
  let inFence = false;
  const lines = readFileSync(file, 'utf8').split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    if (FENCE.test(lines[i])) {
      inFence = !inFence;
      continue;
    }
    if (!inFence) yield [i + 1, lines[i]];
  }
}

function readAnchors(file) {
  const anchors = new Set();
  const seen = new Map();
  for (const [, line] of unfencedLines(file)) {
    const m = HEADING.exec(line);
    if (m) {
      const base = slug(m[2]);
      const n = seen.get(base) ?? 0;
      seen.set(base, n + 1);
      anchors.add(n === 0 ? base : `${base}-${n}`);
    }
    for (const a of line.matchAll(HTML_ANCHOR)) anchors.add(a[1]);
  }
  return anchors;
}

function markdownIn(dir) {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((name) => name.endsWith('.md'))
    .sort()
    .map((name) => path.join(dir, name));
}

/**
 * Checks every relative link under `root`. `checked` counts the links examined; `bad` holds
 * each missing file or unmatched anchor, `notes` each anchor on a non-.md target. Paths in
 * the results are relative to `root` with `/` separators.
 */
export function checkDocLinks(root) {
  const files = [
    path.join(root, 'README.md'),
    ...markdownIn(path.join(root, 'docs')),
    ...markdownIn(path.join(root, 'docs', 'reference')),
    path.join(root, 'CONTRIBUTING.md'),
  ].filter((f) => existsSync(f));
  const anchorCache = new Map();
  const anchorsOf = (file) => {
    if (!anchorCache.has(file)) anchorCache.set(file, readAnchors(file));
    return anchorCache.get(file);
  };
  const rel = (file) => path.relative(root, file).split(path.sep).join('/');

  let checked = 0;
  const bad = [];
  const notes = [];
  for (const file of files) {
    for (const [lineNo, line] of unfencedLines(file)) {
      for (const [, target] of line.replace(/`[^`]*`/g, '').matchAll(LINK)) {
        if (SCHEME.test(target)) continue;
        checked++;
        const hash = target.indexOf('#');
        const filePart = hash === -1 ? target : target.slice(0, hash);
        const anchor = hash === -1 ? '' : target.slice(hash + 1);
        const resolved = filePart ? path.resolve(path.dirname(file), filePart) : file;
        const where = { file: rel(file), line: lineNo, target };
        if (!existsSync(resolved)) {
          bad.push({ ...where, problem: 'missing file' });
        } else if (anchor && !resolved.endsWith('.md')) {
          notes.push({ ...where, problem: 'anchor on a non-.md target' });
        } else if (anchor && !anchorsOf(resolved).has(anchor)) {
          bad.push({ ...where, problem: 'no such anchor' });
        }
      }
    }
  }
  return { checked, bad, notes };
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const { checked, bad, notes } = checkDocLinks(path.resolve(process.argv[2] ?? '.'));
  for (const b of [...bad, ...notes]) console.log(`${b.problem}: ${b.file}:${b.line} -> ${b.target}`);
  console.log(`${checked} links checked, ${bad.length} bad`);
  if (bad.length > 0) process.exitCode = 1;
}
