// Drift guards keeping the two catalogues DERIVED from the tool list in step with it:
// README's tool reference, and the report `check_function_availability` returns. Both are
// hand-maintained lists of tool names sitting next to the `TOOLS` array they describe, and
// both go stale the same way — a tool is added and its entry is simply never written.
//
// The failure this exists to catch is not a stale count. A tool gets appended to TOOLS,
// the "Available Tools (N Total)" heading is bumped along with it, and the tool's own
// README entry is simply never written — a count-only check passes and the tool ships
// undocumented. So the assertion is SET EQUALITY between the names in TOOLS and the names
// README documents, in both directions (the reverse direction catches docs left behind by
// a removed tool, which this fork has done more than once), and the heading number is
// derived from the set rather than checked on its own.
//
// Both sides are read as source TEXT out of src/, never from the built server. `npm test`
// runs tsx over src/ and never builds first, so a check that imported dist/ would read the
// previous build and miss the tool just added — precisely the drift it is here to catch.
// (Same reasoning as the schema scan in tool-schema.test.ts.)
//
// README STRUCTURAL CONTRACT — what counts as a tool's reference entry:
//
//   An UNINDENTED list item whose first content is the tool name in bold, immediately
//   followed by a colon:  `- **send_draft**: Send an existing draft email …`
//   living under a `###` subsection of the `## Available Tools (N Total)` section.
//
// Everything else that names a tool is a mention, not documentation, and is deliberately
// not counted: prose and cross-references name tools in backticks (`send_draft`); the
// "Most Popular Tools" teaser at the top of the section uses the same bullet shape but
// sits above the first `###`, and is a pointer to the entries below rather than an entry
// itself (so a tool listed only there is still reported as undocumented, which is the
// point); the Troubleshooting list near the end of the file is outside the section
// altogether. Nested detail bullets under an entry are indented, so they never match.
//
// If you restructure that part of README, keep the shape above or real entries will stop
// counting here and this guard will report them as missing.
//
// AVAILABILITY REPORT CONTRACT — what `check_function_availability` has to list:
//
//   Every tool the server ships appears in exactly ONE of the report's `functions: [...]`
//   arrays (email / identity / contacts / calendar), so a caller asking what it can do
//   gets the whole surface and never sees a tool claimed by two capability groups. The two
//   meta-tools are exempt: they report on the server rather than acting on an account, so
//   they belong to no capability group and are listed in README's "Most Popular Tools"
//   teaser instead. A tool missing from the report is invisible to any caller that plans
//   its work off it, which is the whole reason the report exists.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const SCHEMA_FILE = join(SRC_DIR, 'index.ts');
const README_FILE = join(SRC_DIR, '..', 'README.md');

const TOOLS_HEADING = /^## Available Tools \((\d+) Total\)\s*$/;

function readLines(file: string): string[] {
  return readFileSync(file, 'utf8').split('\n').map((l) => l.replace(/\r$/, ''));
}

// The tool names the server ships, read off the TOOLS array literal in src/index.ts.
// Scoped to that literal so the server's own `name: 'fastmail-mcp'` and the sample tool
// names in the dry-run handler further down the file cannot leak in. A tool parameter
// called `name` would be written `name: {`, so the string-valued form is unambiguous.
function collectShippedToolNames(): string[] {
  const lines = readLines(SCHEMA_FILE);
  const start = lines.findIndex((l) => l.trim() === 'const TOOLS = [');
  assert.notEqual(start, -1, 'could not find the `const TOOLS = [` literal in src/index.ts');
  const end = lines.findIndex((l, i) => i > start && l.trim() === '];');
  assert.notEqual(end, -1, 'could not find the end of the TOOLS literal in src/index.ts');

  const names: string[] = [];
  for (let i = start + 1; i < end; i++) {
    const match = /^\s*name: '([a-z][a-z0-9_]*)',$/.exec(lines[i]);
    if (match) names.push(match[1]);
  }
  return names;
}

// The tools that deliberately appear in no capability group of the availability report:
// check_function_availability reports on the server itself, and test_bulk_operations is a
// self-test harness. Adding a name here excuses a tool from the coverage assertion below,
// so only a tool that genuinely acts on no account capability belongs in it.
const AVAILABILITY_EXEMPT_TOOLS = new Set(['check_function_availability', 'test_bulk_operations']);

// The `functions: [...]` arrays of the availability report, one entry per capability
// group, read out of the `const availability = {` literal in src/index.ts. Kept as groups
// rather than flattened so a name listed under two groups is detectable.
function collectAvailabilityGroups(): string[][] {
  const lines = readLines(SCHEMA_FILE);
  const start = lines.findIndex((l) => l.trim() === 'const availability = {');
  assert.notEqual(start, -1, 'could not find the `const availability = {` literal in src/index.ts');
  const end = lines.findIndex((l, i) => i > start && l.trim() === '};');
  assert.notEqual(end, -1, 'could not find the end of the availability literal in src/index.ts');

  const groups: string[][] = [];
  for (const array of lines.slice(start, end).join('\n').matchAll(/functions:\s*\[([^\]]*)\]/g)) {
    groups.push([...array[1].matchAll(/'([a-z][a-z0-9_]*)'/g)].map((m) => m[1]));
  }
  return groups;
}

// The tool names README documents, plus the number written into the section heading.
function collectDocumentedToolNames(): { names: string[]; headingCount: number } {
  const lines = readLines(README_FILE);
  const headingIndex = lines.findIndex((l) => TOOLS_HEADING.test(l));
  assert.notEqual(
    headingIndex,
    -1,
    'could not find a `## Available Tools (N Total)` heading in README.md',
  );
  const headingCount = Number(TOOLS_HEADING.exec(lines[headingIndex])![1]);

  const names: string[] = [];
  let inSubsection = false;
  for (let i = headingIndex + 1; i < lines.length; i++) {
    const line = lines[i];
    if (/^## /.test(line)) break; // the next top-level section ends the tool reference
    if (/^### /.test(line)) {
      inSubsection = true;
      continue;
    }
    if (!inSubsection) continue; // the "Most Popular Tools" teaser above the first `###`
    const match = /^- \*\*([a-z][a-z0-9_]*)\*\*:/.exec(line);
    if (match) names.push(match[1]);
  }
  return { names, headingCount };
}

describe('README tool inventory', () => {
  it('documents every tool the server ships', () => {
    const shipped = collectShippedToolNames();
    // A floor, so a scan that has stopped matching fails here rather than passing
    // vacuously against an equally empty README set.
    assert.ok(
      shipped.length >= 30,
      `found only ${shipped.length} tools in the TOOLS literal; the source scan has ` +
        `probably stopped matching`,
    );

    const documented = new Set(collectDocumentedToolNames().names);
    const undocumented = shipped.filter((name) => !documented.has(name));
    assert.deepEqual(
      undocumented,
      [],
      `this tool ships with no README entry: ${undocumented.join(', ')}. Add a ` +
        `\`- **<name>**: …\` entry under the matching \`###\` subsection of ` +
        `"## Available Tools", alongside the other tools of its kind`,
    );
  });

  it('documents no tool the server does not ship', () => {
    const shipped = new Set(collectShippedToolNames());
    const { names: documented } = collectDocumentedToolNames();
    const phantom = documented.filter((name) => !shipped.has(name));
    assert.deepEqual(
      phantom,
      [],
      `the README documents a tool that does not exist: ${phantom.join(', ')}. It was ` +
        `removed from TOOLS (or renamed) and its entry was left behind — delete the entry`,
    );
  });

  it('reports every tool in check_function_availability, in exactly one group', () => {
    const shipped = collectShippedToolNames();
    const groups = collectAvailabilityGroups();
    // A floor on the scan itself, so a reformatted availability literal fails here rather
    // than passing against an empty set of groups.
    assert.ok(
      groups.length >= 4,
      `found only ${groups.length} capability groups in the availability report; the ` +
        `source scan has probably stopped matching`,
    );

    const listed = groups.flat();
    const unreported = shipped.filter(
      (name) => !AVAILABILITY_EXEMPT_TOOLS.has(name) && !listed.includes(name),
    );
    assert.deepEqual(
      unreported,
      [],
      `this tool ships but check_function_availability never reports it: ` +
        `${unreported.join(', ')}. Add it to the \`functions\` array of the capability ` +
        `group it belongs to, so a caller planning off that report can see it`,
    );

    const repeated = listed.filter((name, i) => listed.indexOf(name) !== i);
    assert.deepEqual(
      repeated,
      [],
      `these tools are reported under more than one capability group: ` +
        `${repeated.join(', ')}. A tool's availability then depends on which group the ` +
        `caller read`,
    );

    const shippedNames = new Set(shipped);
    const phantom = listed.filter((name) => !shippedNames.has(name));
    assert.deepEqual(
      phantom,
      [],
      `check_function_availability reports a tool that does not exist: ` +
        `${phantom.join(', ')}. It was removed from TOOLS (or renamed) and the report was ` +
        `left behind`,
    );
  });

  it('counts the documented tools in the Available Tools heading', () => {
    const { names, headingCount } = collectDocumentedToolNames();
    const unique = new Set(names);
    assert.equal(
      unique.size,
      names.length,
      `two README entries document the same tool: ${
        names.filter((n, i) => names.indexOf(n) !== i).join(', ')
      }`,
    );
    assert.equal(
      headingCount,
      unique.size,
      `the "## Available Tools (${headingCount} Total)" heading disagrees with the ` +
        `${unique.size} tools documented below it`,
    );
  });
});
