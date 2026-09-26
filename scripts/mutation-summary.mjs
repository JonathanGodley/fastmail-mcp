#!/usr/bin/env node
// Summarise Stryker JSON reports (one per shard) as Markdown, for a GitHub Actions run summary.
//
//   node scripts/mutation-summary.mjs reports/mutation-*.json >> "$GITHUB_STEP_SUMMARY"

import { readFileSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const UNTESTED = ['Survived', 'NoCoverage'];

/** Markdown for the given parsed reports, listing at most `cap` untested mutants. */
export function summarise(reports, cap = 200) {
  const counts = { Killed: 0, Survived: 0, Timeout: 0, NoCoverage: 0 };
  let total = 0;
  const untested = [];
  for (const report of reports) {
    for (const [file, entry] of Object.entries(report.files ?? {})) {
      for (const m of entry.mutants ?? []) {
        total += 1;
        if (m.status in counts) counts[m.status] += 1;
        if (UNTESTED.includes(m.status)) untested.push({ file, line: m.location.start.line, mutator: m.mutatorName, status: m.status });
      }
    }
  }
  const detected = counts.Killed + counts.Timeout;
  const scored = detected + counts.Survived + counts.NoCoverage;
  const score = scored === 0 ? 'n/a' : `${((detected / scored) * 100).toFixed(2)}%`;
  untested.sort((a, b) => (a.file < b.file ? -1 : a.file > b.file ? 1 : a.line - b.line));

  const out = [
    '## Mutation testing',
    '',
    `Reports read: ${reports.length}`,
    '',
    '| Mutants | Killed | Survived | Timeout | No coverage | Score |',
    '| ---: | ---: | ---: | ---: | ---: | ---: |',
    `| ${total} | ${counts.Killed} | ${counts.Survived} | ${counts.Timeout} | ${counts.NoCoverage} | ${score} |`,
    '',
  ];
  if (untested.length === 0) {
    out.push('No surviving mutants.');
  } else {
    out.push(`### Surviving mutants (${untested.length})`, '');
    for (const u of untested.slice(0, cap)) {
      out.push(`- \`${u.file}:${u.line}\` ${u.mutator}${u.status === 'NoCoverage' ? ' (no coverage)' : ''}`);
    }
    if (untested.length > cap) out.push('', `...and ${untested.length - cap} more; see the report artifacts.`);
  }
  return `${out.join('\n')}\n`;
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const files = process.argv.slice(2);
  if (files.length === 0) {
    process.stdout.write('## Mutation testing\n\nNo shard reports were found; see the shard jobs for the error.\n');
  } else {
    process.stdout.write(summarise(files.map((f) => JSON.parse(readFileSync(f, 'utf8')))));
  }
}
