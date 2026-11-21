import { describe, it, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { checkDocLinks } from '../scripts/check-doc-links.mjs';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..');
const roots: string[] = [];
after(() => {
  for (const root of roots) rmSync(root, { recursive: true, force: true });
});

/** A fixture tree under the real temp dir; keys are `/`-separated paths below the root. */
function tree(files: Record<string, string>): string {
  const root = mkdtempSync(join(tmpdir(), 'fastmail-mcp-doc-links-'));
  roots.push(root);
  for (const [rel, body] of Object.entries(files)) {
    const file = join(root, ...rel.split('/'));
    mkdirSync(dirname(file), { recursive: true });
    writeFileSync(file, body);
  }
  return root;
}

describe('checkDocLinks', () => {
  it('reports a link to a file that does not exist', () => {
    const root = tree({ 'README.md': 'See [the guide](docs/guide.md).\n' });
    const { checked, bad } = checkDocLinks(root);
    assert.equal(checked, 1);
    assert.deepEqual(bad, [{ file: 'README.md', line: 1, target: 'docs/guide.md', problem: 'missing file' }]);
  });

  it('reports an anchor that matches no heading, and passes one that does', () => {
    const root = tree({
      'README.md': '[ok](docs/guide.md#setting-up)\n[bad](docs/guide.md#tearing-down)\n',
      'docs/guide.md': '# Guide\n\n## Setting up\n',
    });
    const { checked, bad } = checkDocLinks(root);
    assert.equal(checked, 2);
    assert.deepEqual(bad, [{ file: 'README.md', line: 2, target: 'docs/guide.md#tearing-down', problem: 'no such anchor' }]);
  });

  it('gives a repeated heading the -1 slug, and an <a id> counts as an anchor', () => {
    const root = tree({
      'docs/a.md': '## Notes\n\n## Notes\n\n<a id="pinned"></a>\n\n[second](#notes-1) [third](#notes-2) [pin](#pinned)\n',
    });
    assert.deepEqual(checkDocLinks(root).bad.map((b) => b.target), ['#notes-2']);
  });

  it('keeps both hyphens either side of a spaced em dash, and keeps code text in the slug', () => {
    const root = tree({
      'docs/a.md': '## Contacts — `kind` rules\n\n[here](#contacts--kind-rules)\n',
    });
    assert.deepEqual(checkDocLinks(root).bad, []);
  });

  it('skips a link inside a code fence and inside an inline code span', () => {
    const root = tree({
      'README.md': '```\n[fenced](missing-one.md)\n```\nInline `[spanned](missing-two.md)` here.\n',
    });
    assert.deepEqual(checkDocLinks(root), { checked: 0, bad: [], notes: [] });
  });

  it('skips a link with a URL scheme', () => {
    const root = tree({ 'README.md': '[site](https://example.com/missing.md) [mail](mailto:someone@example.com)\n' });
    assert.deepEqual(checkDocLinks(root), { checked: 0, bad: [], notes: [] });
  });

  it('checks a link that carries a single-quoted title', () => {
    const root = tree({ 'README.md': "[x](missing.md 'A title')\n" });
    assert.deepEqual(checkDocLinks(root).bad, [{ file: 'README.md', line: 1, target: 'missing.md', problem: 'missing file' }]);
  });

  it('recognises an <a id> or <a name> written with single quotes', () => {
    const root = tree({ 'docs/a.md': "<a id='first'></a> <a name='second'></a>\n\n[1](#first) [2](#second)\n" });
    assert.deepEqual(checkDocLinks(root), { checked: 2, bad: [], notes: [] });
  });

  it('skips a URL scheme written in capitals', () => {
    const root = tree({ 'README.md': '[site](HTTPS://example.com/missing.md)\n' });
    assert.deepEqual(checkDocLinks(root), { checked: 0, bad: [], notes: [] });
  });

  it('notes an anchor on a non-.md target without failing it', () => {
    const root = tree({ 'README.md': '[licence](LICENSE#terms)\n', LICENSE: 'terms\n' });
    const { checked, bad, notes } = checkDocLinks(root);
    assert.equal(checked, 1);
    assert.deepEqual(bad, []);
    assert.deepEqual(notes, [{ file: 'README.md', line: 1, target: 'LICENSE#terms', problem: 'anchor on a non-.md target' }]);
  });

  it('finds no bad link in this repository', () => {
    const { checked, bad } = checkDocLinks(REPO);
    assert.ok(checked > 0, 'no links were checked');
    assert.deepEqual(bad, []);
  });
});
