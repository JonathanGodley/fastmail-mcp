// Three drift guards over what this server's tools present to a caller: the lenient-boolean
// convention (#54), the rule that a tool emitting a recovery note has to expose every
// parameter that note prescribes, and the rule that every JSON result payload goes out
// compact (#40). The first two read the schemas and handlers in src/index.ts; the third
// scans every non-test source file, since any of them can serialise a payload.
//
// Lenient booleans (#54).
//
// Every boolean tool parameter has to satisfy BOTH halves of that convention, and each
// half fails silently on its own:
//
//   1. The schema declares `type: ['boolean', 'string']`. A narrow `type: 'boolean'`
//      makes a validating client reject "true"/"false" before the request is dispatched,
//      so whatever coercion the handler runs is unreachable.
//   2. The handler reads the value through coerceBool, not `!!`. Under `!!`, the string
//      "false" is truthy — the parameter silently inverts.
//
// This asserts against src/index.ts and the handler modules as TEXT rather than by
// spawning the built server and reading tools/list. tools/list would prove what is
// actually shipped, which is the stronger claim, but `npm test` runs tsx over src/ and
// never builds first: a guard read out of a stale dist/ would not see a tool added since
// the last build, which is precisely the drift it exists to catch. tsc does not rewrite
// string literals, so the source and the shipped schema cannot disagree on these, and the
// source check can never go stale. (scripts/mcp-harness.mjs grew a list() for the
// on-demand check against a freshly built server; `node scripts/mcp-harness.mjs --list`.)
//
// Recovery notes name only parameters the emitting tool has.
//
// The Trash/Spam exclusion note ends with a runnable recovery — "set includeTrash:true
// (or mailbox:"trash") to include them" — and a note is worth having only if the caller
// reading it can act on it. Naming a parameter the tool does not declare is worse than
// saying nothing: the retry is rejected outright by the unknown-parameter guard, and the
// caller is left believing the withheld messages are unreachable.
//
// Nothing else enforces the pair. The note text is built from the roles that were actually
// excluded, so it is the same string on every tool that emits it, while the parameters are
// declared per tool — a tool can start emitting notes (get_recent_emails did, #29) without
// declaring the flags they prescribe, and both halves still look correct in isolation. So
// the assertion is derived from both sides at once: which tools emit a note is read out of
// the handlers, and which parameters a note names is read out of the real emitter's output
// rather than a copy of its wording.
//
// Compact result payloads (#40).
//
// Every JSON result item goes out through one of two seams in coerce.ts — toolJson, or
// redactedJson where the values may carry credentials. Neither takes an indent argument, so
// the seams themselves cannot pretty-print, and TypeScript rejects a second argument to
// either. What nothing else catches is a new handler reaching past them for a bare
// JSON.stringify with an indent, which is how every site came to be indented in the first
// place: the cost is invisible at the call site and only shows up in the response. The
// number of sites is deliberately not written down here - it moves whenever a handler is
// refactored or a renderer is shared, and the scan below is what pins it.
//
// So the scan finds JSON.stringify CALL SITES — every shipped .ts under src/, recursively,
// with comments and string/template/regex literals blanked first so only real syntax is left
// — and asserts two things about each: that it passes no third argument (an indent, whatever
// the replacer looks like), and that it is one of the listed non-payload uses.
//
// What that second half actually buys, stated exactly, because the loose version of the
// sentence ("the seam is what redacts") is FALSE and would invite waving a real gap through
// later. toolJson is a bare JSON.stringify with no replacer, and it is nearly every seam call
// site, so routing a payload through it adds no redaction whatsoever. What the second half
// buys is that serialisation stays in ONE place: neither seam accepts an indent, so keeping
// every payload on them makes "no payload can be pretty-printed" a property of two function
// signatures rather than of a text scan that has to keep finding every call site forever. It
// also keeps the choice between the two seams visible when a new handler is reviewed.
//
// Where redaction does and does not run, since this scan cannot tell you and the file should
// not imply otherwise: redactedJson is the only serialiser that redacts, and it has one call
// site (the bulk-operations result in index.ts). Success payloads on every other path are not
// redacted, and were not before the compaction either. The ERROR path is covered independently
// of both seams - every error reply is redacted centrally in index.ts's CallTool catch - so a
// bearer token in a server error description is caught there, not here.
//
// It is not a small cost. Indentation is bytes the reader pays for that nothing parses, and
// it scales with the number of JSON tokens rather than with the content, so it varies with a
// payload's shape rather than its size — measured live against one real account when the
// change landed (2026-08): 17.3% of a 25-message list page, 24.9% of that same page under
// raw:true, 28.5% of a mailbox listing (many small flat objects, the most delimiters per byte
// of content) and 6.5% of a single get_email (one long body string, which has no delimiters
// to indent). A point-in-time measurement, not a rate: the full table is on issue #106.

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildExclusionNote } from './response-formatters.js';

const SRC_DIR = dirname(fileURLToPath(import.meta.url));
const SCHEMA_FILE = join(SRC_DIR, 'index.ts');

// The files that read tool arguments. The formatters and the JMAP client take already-
// coerced values, and email-formatter.ts uses `raw` as the name of a JMAP object, so
// scanning them would only generate false positives.
const HANDLER_FILES = [
  'index.ts',
  'thread-handler.ts',
  'draft-email-handler.ts',
  'edit-draft-handler.ts',
  'send-draft-handler.ts',
  'mailbox-handler.ts',
  'contacts-handler.ts',
];

function readLines(file: string): string[] {
  return readFileSync(join(SRC_DIR, file), 'utf8').split('\n').map((l) => l.replace(/\r$/, ''));
}

// Collect the boolean parameters out of the TOOLS schemas. Matches the declaration line
// exactly (leading whitespace, trailing comma) so the prose in nearby comments — which
// quotes both forms — is not picked up.
function collectBooleanParams(): { unionNames: string[]; narrow: string[] } {
  const lines = readLines('index.ts');
  const unionNames = new Set<string>();
  const narrow: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const isUnion = line === "type: ['boolean', 'string'],";
    const isNarrow = line === "type: 'boolean',";
    if (!isUnion && !isNarrow) continue;
    // The property name is the line that opened this block.
    const name = lines[i - 1].trim().replace(/:\s*\{$/, '');
    if (isUnion) unionNames.add(name);
    else narrow.push(`${name} (src/index.ts:${i + 1})`);
  }
  return { unionNames: [...unionNames], narrow };
}

describe('lenient-boolean convention', () => {
  it('declares every boolean tool parameter as type: [boolean, string]', () => {
    const { unionNames, narrow } = collectBooleanParams();
    // A sanity floor: if the scan ever matches nothing (the schema literal is reformatted,
    // TOOLS moves out of this file) the assertion below would pass vacuously.
    assert.ok(
      unionNames.length >= 15,
      `found only ${unionNames.length} boolean parameters; the schema scan has probably stopped matching`,
    );
    assert.deepEqual(
      narrow,
      [],
      `these boolean parameters declare a narrow type: 'boolean', which makes their coerceBool ` +
        `unreachable from a validating client. Widen to type: ['boolean', 'string'] and wrap the ` +
        `description in lenientBool(): ${narrow.join(', ')}`,
    );
  });

  it('reads every boolean tool parameter through coerceBool, never bare !!', () => {
    const { unionNames } = collectBooleanParams();
    const offenders: string[] = [];
    for (const file of HANDLER_FILES) {
      const lines = readLines(file);
      for (const name of unionNames) {
        // `!!name`, `!!(args as any).name` and `!!args?.name`. The negative lookahead
        // keeps `!!raw.hasAttachment` (a property read off a JMAP object that happens to
        // be called `raw`) from matching.
        const patterns = [
          new RegExp(`!!\\s*${name}\\b(?![.\\[])`),
          new RegExp(`!!\\s*\\(\\s*args\\s+as\\s+any\\s*\\)\\.${name}\\b`),
          new RegExp(`!!\\s*args\\?\\.${name}\\b`),
        ];
        lines.forEach((line, i) => {
          if (line.trimStart().startsWith('//')) return;
          if (patterns.some((p) => p.test(line))) offenders.push(`${name} (src/${file}:${i + 1})`);
        });
      }
    }
    assert.deepEqual(
      offenders,
      [],
      `these boolean parameters are read with bare !!, so a lenient client's "false" reads as ` +
        `true. Use coerceBool(...) ?? <default>: ${offenders.join(', ')}`,
    );
  });
});

// timeZone accepts null as a real, deliberately-rejected input (create_calendar_event and
// update_calendar_event, #157) — not absence. Omitting timeZone is what absence means, and
// that is already handled by the parameter being optional; `null` is a caller explicitly
// asking to force a floating write, which validateCallerTimezone rejects with a tailored
// message ("there is no way to force a floating write through this parameter"). The schema
// has to say `type: ['string', 'null']` for that message to ever be reached: a narrowing edit
// back to `type: 'string'` makes a validating client reject `timeZone: null` with a generic
// type-mismatch error before this server's own handler ever runs — the same failure mode the
// lenient-boolean guard above exists for, one type union over.
function collectTimeZoneNullableParams(): { nullable: string[]; narrow: string[] } {
  const lines = readLines('index.ts');
  const nullable: string[] = [];
  const narrow: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const isNullable = line === "type: ['string', 'null'],";
    const isNarrow = line === "type: 'string',";
    if (!isNullable && !isNarrow) continue;
    // The property name is the nearest non-comment, non-blank line above this one — timeZone's
    // own declaration carries an explanatory comment between the property name and its type.
    let j = i - 1;
    while (j >= 0 && (lines[j].trim() === '' || lines[j].trim().startsWith('//'))) j--;
    const name = lines[j].trim().replace(/:\s*\{$/, '');
    if (name !== 'timeZone') continue;
    if (isNullable) nullable.push(`${name} (src/index.ts:${i + 1})`);
    else narrow.push(`${name} (src/index.ts:${i + 1})`);
  }
  return { nullable, narrow };
}

describe('timeZone parameter accepts null (#157)', () => {
  it('declares every timeZone parameter as type: [string, null], not narrowed to string', () => {
    const { nullable, narrow } = collectTimeZoneNullableParams();
    // A sanity floor matching the two tools that declare timeZone today (create_calendar_event,
    // update_calendar_event): if the scan ever matches nothing, this would pass vacuously.
    assert.ok(
      nullable.length >= 2,
      `found only ${nullable.length} timeZone parameters declared type: ['string', 'null']; ` +
        'the schema scan has probably stopped matching.',
    );
    assert.deepEqual(
      narrow,
      [],
      `these timeZone parameters declare a narrow type: 'string', dropping 'null' from the union — ` +
        `a validating client would then reject timeZone: null with a generic type-mismatch error ` +
        `before this server's own tailored rejection is ever reached: ${narrow.join(', ')}`,
    );
  });
});

// The parameter names each tool declares, keyed by tool name, read off the TOOLS literal.
// Property keys sit one indent level inside `properties: {`, and every nested shape (an
// array's `items`, an object property's own keys) is indented deeper, so the fixed-depth
// match cannot pick one up. Both declaration forms count: `mailbox: {` and the shared
// helpers, `position: positionSchemaProperty(),`.
function collectToolParams(): Map<string, Set<string>> {
  const lines = readLines('index.ts');
  const start = lines.findIndex((l) => l.trim() === 'const TOOLS = [');
  assert.notEqual(start, -1, 'could not find the `const TOOLS = [` literal in src/index.ts');
  const end = lines.findIndex((l, i) => i > start && l.trim() === '];');
  assert.notEqual(end, -1, 'could not find the end of the TOOLS literal in src/index.ts');

  const params = new Map<string, Set<string>>();
  let current: string | undefined;
  let inProperties = false;
  for (let i = start + 1; i < end; i++) {
    const line = lines[i];
    const name = /^ {8}name: '([a-z][a-z0-9_]*)',$/.exec(line);
    if (name) {
      current = name[1];
      params.set(current, new Set());
      inProperties = false;
      continue;
    }
    if (!current) continue;
    if (line === '          properties: {') {
      inProperties = true;
      continue;
    }
    if (inProperties) {
      if (line === '          },') {
        inProperties = false;
        continue;
      }
      const key = /^ {12}([A-Za-z_][A-Za-z0-9_]*):/.exec(line);
      if (key) params.get(current)!.add(key[1]);
    }
  }
  return params;
}

// The CallTool switch in index.ts, split into one body per `case '<tool>':`. A case runs
// until the next case label, which is exact here because no case in that switch falls
// through into another (the label-only ones are single-line delegations to a handler
// module). Comment lines are dropped, so prose quoting a call is never mistaken for one.
//
// This attribution is by POSITION, which is its known limit: logic extracted out of the
// switch into an injected-client handler module — the refactor CLAUDE.md's Testing section
// pushes toward — leaves no line under the case label to find. Every assertion built on
// this therefore carries a floor on how many sites it matched, so the extraction shows up
// as a failure here instead of as silent coverage loss.
function collectCaseBodies(): Map<string, string[]> {
  const bodies = new Map<string, string[]>();
  let current: string | undefined;
  for (const line of readLines('index.ts')) {
    const trimmed = line.trim();
    if (trimmed.startsWith('//')) continue;
    const caseLabel = /^case '([a-z][a-z0-9_]*)':/.exec(trimmed);
    if (caseLabel) {
      current = caseLabel[1];
      if (!bodies.has(current)) bodies.set(current, []);
      continue;
    }
    if (current) bodies.get(current)!.push(trimmed);
  }
  return bodies;
}

// The tools whose handler appends buildExclusionNote to its response. Read from the source
// rather than listed here so a tool that starts emitting notes is covered the moment it
// does.
function collectNoteEmittingTools(): string[] {
  return [...collectCaseBodies()]
    .filter(([, body]) => body.some((line) => line.includes('buildExclusionNote(')))
    .map(([tool]) => tool);
}

// The parameter names the note actually prescribes, taken from the emitter's own output
// over every state it can report — a hidden count, the degraded count, an unresolved role,
// and the mixed case — rather than from a copy of its wording here. `includeTrash:true`,
// `includeSpam:true` and `mailbox:"trash"` all match; ordinary prose colons ("Note: 3 …")
// do not, because the lookahead requires a value straight after the colon.
function collectNoteParams(): string[] {
  const states: NonNullable<Parameters<typeof buildExclusionNote>[0]>[] = [
    { hidden: 3, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] },
    { hidden: 2, excludedRoles: ['Trash'], unresolvedRoles: [] },
    { hidden: 1, excludedRoles: ['Spam'], unresolvedRoles: [] },
    { hidden: null, excludedRoles: ['Trash', 'Spam'], unresolvedRoles: [] },
    { hidden: 0, excludedRoles: [], unresolvedRoles: ['Trash', 'Spam'] },
    { hidden: 4, excludedRoles: ['Spam'], unresolvedRoles: ['Trash'] },
  ];
  const names = new Set<string>();
  for (const state of states) {
    for (const match of buildExclusionNote(state).matchAll(/\b([A-Za-z_][A-Za-z0-9_]*):(?=true|false|")/g)) {
      names.add(match[1]);
    }
  }
  return [...names];
}

describe('recovery notes name only parameters the tool has', () => {
  it('declares every parameter the Trash/Spam exclusion note prescribes', () => {
    const emitters = collectNoteEmittingTools();
    // The floor has TWO jobs, and the second is the less obvious one.
    //
    // It stops a scan that has silently stopped matching from passing against an empty
    // set. And it is the only thing that notices a note emitter this scan can no longer
    // see: attribution is positional (see collectCaseBodies), so moving a handler into an
    // injected-client module removes it from `emitters` and the subset assertion below
    // would then hold vacuously for it. Raise this number with each emitter added; do not
    // lower it to make a failure go away without checking which of the two it is.
    //
    // The subset assertion is not carried by the floor, and neither carries the other. The
    // failure it detects is a tool that emits the note while missing a flag the note
    // prescribes: hold `emitters` fixed, delete includeTrash from one of those tools'
    // schemas, and `missing` is non-empty. (Against the state before get_recent_emails
    // started emitting notes, `missing` is legitimately empty — that tool emitted nothing
    // to be held to — and it is the floor, not the subset check, that fails there.)
    assert.ok(
      emitters.length >= 3,
      `found only ${emitters.length} tools appending buildExclusionNote; either the handler ` +
        `scan has stopped matching, or an emitter moved out of the CallTool switch and is ` +
        `no longer being checked`,
    );
    const noteParams = collectNoteParams();
    assert.ok(
      noteParams.length >= 3,
      `extracted only ${noteParams.length} parameter names from the exclusion note; the note ` +
        `wording or the extraction has drifted`,
    );

    const params = collectToolParams();
    const missing: string[] = [];
    for (const tool of emitters) {
      const declared = params.get(tool);
      assert.ok(declared, `${tool} appends an exclusion note but is not in the TOOLS literal`);
      for (const name of noteParams) {
        if (!declared.has(name)) missing.push(`${tool}.${name}`);
      }
    }
    assert.deepEqual(
      missing,
      [],
      `this tool emits a note prescribing a parameter it does not declare, so following the ` +
        `note is rejected as an unknown parameter: ${missing.join(', ')}. Declare the ` +
        `parameter, or stop the tool emitting the note`,
    );
  });
});

// A tool's README reference entry: the indented lines under its `- **<name>**: …` bullet, up
// to the next unindented list item. Same structural contract readme-inventory.test.ts pins for
// the entry LINE itself (see the comment at the top of that file); this reads the detail
// bullets beneath it, which is where the `- Parameters: …` line lives.
function collectReadmeToolEntries(): Map<string, string> {
  const lines = readFileSync(join(SRC_DIR, '..', 'README.md'), 'utf8')
    .split('\n').map((l) => l.replace(/\r$/, ''));
  const heading = lines.findIndex((l) => /^## Available Tools \(\d+ Total\)\s*$/.test(l));
  assert.notEqual(heading, -1, 'could not find a `## Available Tools (N Total)` heading in README.md');

  const entries = new Map<string, string[]>();
  let current: string | undefined;
  for (let i = heading + 1; i < lines.length; i++) {
    const line = lines[i];
    if (/^## /.test(line)) break;
    const entry = /^- \*\*([a-z][a-z0-9_]*)\*\*:/.exec(line);
    if (entry) {
      current = entry[1];
      entries.set(current, [line]);
      continue;
    }
    if (!current) continue;
    if (/^- /.test(line)) { current = undefined; continue; } // a bullet that is not an entry
    entries.get(current)!.push(line);
  }
  return new Map([...entries].map(([tool, body]) => [tool, body.join('\n')]));
}

// README documents every parameter, not just every tool.
//
// readme-inventory.test.ts asserts set equality on tool NAMES, which is what catches a tool
// shipping with no entry at all. It says nothing about the entry's contents, so a parameter
// added to an existing tool — the far commoner change — was held by diligence alone: the
// entry still exists, the guard still passes, and the parameter is invisible to anyone
// reading the docs instead of the schema.
//
// Matched on the backticked name anywhere in the entry, deliberately loosely. The point is
// that the parameter is MENTIONED; whether its description is any good is not something a
// text scan can judge, and a stricter shape (a position in the `- Parameters:` line, a
// required/optional tag) would make ordinary prose edits fail for no gain.
describe('README documents every tool parameter', () => {
  it('mentions each declared parameter in the tool entry that has it', () => {
    const params = collectToolParams();
    const entries = collectReadmeToolEntries();
    // A floor on both scans, so a restructured README or TOOLS literal fails here rather
    // than passing vacuously against an empty set on either side.
    assert.ok(params.size >= 30, `found only ${params.size} tools in the TOOLS literal; the schema scan has probably stopped matching`);
    assert.ok(entries.size >= 30, `found only ${entries.size} README tool entries; the README scan has probably stopped matching`);

    const missing: string[] = [];
    for (const [tool, declared] of params) {
      const entry = entries.get(tool);
      // A tool with no entry at all is readme-inventory.test.ts's failure to report, not
      // this one's — reporting it twice would just make one fix look like two problems.
      if (entry === undefined) continue;
      for (const name of [...declared].sort()) {
        if (!entry.includes(`\`${name}\``)) missing.push(`${tool}.${name}`);
      }
    }
    assert.deepEqual(
      missing,
      [],
      `this parameter is declared in the tool's schema but named nowhere in its README ` +
        `entry: ${missing.join(', ')}. Add it to that entry's \`- Parameters:\` line (and ` +
        `anywhere else the entry describes behaviour it changes), so a reader of the docs ` +
        `sees the same surface a reader of the schema does`,
    );
  });
});

// What a tool's `limit` schema advertises: the `default:` keyword and the "max: N" its
// description states. Both are read per tool out of the TOOLS literal.
function collectLimitSchemas(): Map<string, { fallback?: number; max?: number }> {
  const lines = readLines('index.ts');
  const start = lines.findIndex((l) => l.trim() === 'const TOOLS = [');
  const end = lines.findIndex((l, i) => i > start && l.trim() === '];');
  const schemas = new Map<string, { fallback?: number; max?: number }>();
  let current: string | undefined;
  let inLimit = false;
  for (let i = start + 1; i < end; i++) {
    const line = lines[i];
    const name = /^ {8}name: '([a-z][a-z0-9_]*)',$/.exec(line);
    if (name) {
      current = name[1];
      inLimit = false;
      continue;
    }
    if (!current) continue;
    if (/^ {12}limit: \{$/.test(line)) {
      inLimit = true;
      schemas.set(current, {});
      continue;
    }
    if (!inLimit) continue;
    if (/^ {12}\},$/.test(line)) {
      inLimit = false;
      continue;
    }
    const fallback = /^ {14}default: (\d+),$/.exec(line);
    if (fallback) schemas.get(current)!.fallback = Number(fallback[1]);
    const max = /max: (\d+)\)/.exec(line);
    if (max) schemas.get(current)!.max = Number(max[1]);
  }
  return schemas;
}

// `clampLimit(<value>, <fallback>, <max>)` inside a handler case body, one entry per tool.
function collectClamps(): Map<string, { fallback: number; max: number }> {
  const clamps = new Map<string, { fallback: number; max: number }>();
  for (const [tool, body] of collectCaseBodies()) {
    for (const line of body) {
      const match = /clampLimit\([^,]+,\s*(\d+),\s*(\d+)\)/.exec(line);
      if (match) clamps.set(tool, { fallback: Number(match[1]), max: Number(match[2]) });
    }
  }
  return clamps;
}

// The limit bound belongs to the handler, not to the JMAP client (#29). getRecentEmails
// used to carry a `Math.min(limit, 50)` of its own; that came off, because a second clamp
// can only ever disagree with the first, and the client now passes `limit` to JMAP exactly
// as given. What made the deletion safe is that every caller clamps BEFORE calling in —
// which is a property of the call sites, so it is asserted at the call sites.
//
// The failure being guarded is quiet and expensive: an unclamped non-numeric limit reaches
// JMAP as `"limit": null`, which is not a small page but no bound at all — a whole-mailbox
// metadata dump. Nothing else notices it. The client-level tests deliberately assert the
// unclamped passthrough, so they would stay green through exactly this regression.
describe('the limit bound is owned by the handlers', () => {
  it('clamps in every handler that calls getRecentEmails', () => {
    // Every call site, across the files that read tool arguments — not index.ts alone. A
    // call from a handler module would have no case body to clamp in, so it must show up
    // as a failure rather than as a site this scan never looked at.
    const callSites: string[] = [];
    for (const file of HANDLER_FILES) {
      readLines(file).forEach((line, i) => {
        if (line.trimStart().startsWith('//')) return;
        if (/\.getRecentEmails\(/.test(line)) callSites.push(`src/${file}:${i + 1}`);
      });
    }
    assert.ok(
      callSites.length >= 2,
      `found only ${callSites.length} getRecentEmails call sites; the scan has probably ` +
        `stopped matching`,
    );

    const clamping = new Set(
      [...collectCaseBodies()]
        .filter(([, body]) => body.some((l) => l.includes('.getRecentEmails(')))
        .filter(([tool]) => collectClamps().has(tool))
        .map(([tool]) => tool),
    );
    const clampedSites = [...collectCaseBodies()]
      .filter(([tool]) => clamping.has(tool))
      .reduce((n, [, body]) => n + body.filter((l) => l.includes('.getRecentEmails(')).length, 0);

    assert.equal(
      clampedSites,
      callSites.length,
      `${callSites.length} call site(s) of getRecentEmails exist (${callSites.join(', ')}) but ` +
        `only ${clampedSites} sit in a handler that calls clampLimit. The client no longer ` +
        `bounds its own limit, so an unclamped call sends "limit": null — no bound at all. ` +
        `Clamp with clampLimit(value, default, max) in the handler`,
    );
  });

  // The clamp is where the tool's real limit default and cap live, so it is where they are
  // asserted. The schema is the other half of the same contract — it is what a client reads
  // and materialises — and the two are written in different files with nothing tying them
  // together. A clamp that disagrees with the number beside it in the schema advertises one
  // default and applies another.
  it('clamps to the default and cap each tool advertises', () => {
    const clamps = collectClamps();
    const schemas = collectLimitSchemas();
    assert.ok(clamps.size >= 4, `found only ${clamps.size} clampLimit call sites; the scan has probably stopped matching`);

    const mismatches: string[] = [];
    for (const [tool, clamp] of clamps) {
      const schema = schemas.get(tool);
      if (!schema) continue; // a clamp on a tool with no `limit` parameter has nothing to agree with
      if (schema.fallback !== undefined && schema.fallback !== clamp.fallback) {
        mismatches.push(`${tool}: schema default ${schema.fallback} vs clamp fallback ${clamp.fallback}`);
      }
      if (schema.max !== undefined && schema.max !== clamp.max) {
        mismatches.push(`${tool}: schema says max ${schema.max} vs clamp max ${clamp.max}`);
      }
    }
    assert.deepEqual(
      mismatches,
      [],
      `a tool's limit clamp disagrees with what its schema advertises, so it applies a bound ` +
        `it never told the caller about: ${mismatches.join('; ')}`,
    );
  });
});

// Each scope/status flag has to be read from the argument of the SAME name. The handlers
// read them positionally into an options object
// (`includeTrash: coerceBool((args as any).includeTrash)`), where swapping two names is a
// one-character edit that type-checks, runs, and silently inverts which folder is hidden.
//
// This covers the wiring only. The DEFAULT each flag falls back to, and the append of the
// exclusion note itself, are still uncovered: they sit in the CallTool switch, which
// CLAUDE.md records as having no test harness. That residual is accepted here rather than
// tracked — the extractable part of these handlers is a destructure-and-delegate, and the
// injected-client extraction the Testing section prescribes is for handlers that
// orchestrate. What made the residual worth narrowing at all is that the mis-wire above is
// both the likeliest edit and the one whose symptom (mail quietly missing from a result)
// looks like an empty mailbox rather than like a bug.
// The same class of untestable handler wiring, for a different coercer. archive_email's
// contract says `notFound` means the server did not know the id; the LENIENT
// coerceStringArray maps every element through String(), so `emailIds: [null]` would reach
// Email/get as the literal id "null" and come back in that bucket. A caller reading the
// report would be told the server does not have a message it was never asked about, and the
// type error would be invisible.
//
// Nothing else catches the swap. coerceStringArrayStrict's own rejection is unit-tested in
// coerce.test.ts, but that pins the COERCER, not which coercer the handler calls — swap the
// call here for the lenient one and every existing test still passes.
describe('archive_email is wired to the strict string-array coercer', () => {
  it('reads emailIds through coerceStringArrayStrict', () => {
    const lines = readLines('index.ts');
    const start = lines.findIndex((l) => l.includes("case 'archive_email':"));
    assert.ok(start >= 0, "could not find the archive_email case in src/index.ts");
    // The handler is short; scanning to the next `case '` keeps this from reading a
    // neighbouring tool's coercion and calling it a pass.
    const end = lines.findIndex((l, i) => i > start && /^\s*case '/.test(l));
    const body = lines.slice(start, end > start ? end : start + 40).join('\n');
    assert.match(
      body,
      /const emailIds = coerceStringArrayStrict\(/,
      'archive_email must coerce emailIds with coerceStringArrayStrict; the lenient ' +
        'coerceStringArray would turn a non-string element into a literal id string and ' +
        'report it as notFound',
    );
    assert.doesNotMatch(
      body,
      // `coerceStringArrayStrict(` does not match this: the pattern requires the open
      // paren immediately after the lenient name.
      /coerceStringArray\(/,
      'archive_email must not use the lenient coerceStringArray for any argument',
    );
  });

  it('serialises counts alongside results, which both descriptions promise', () => {
    // The tool description and the README both tell a caller the counts sum to the number
    // of distinct ids they passed. That invariant is uncheckable from the prose alone,
    // because a bucket with no entries produces no line — so `counts` has to be in the JSON
    // item. Serialising `result.results` on its own satisfies every other test in the repo
    // while quietly making both descriptions wrong.
    const lines = readLines('index.ts');
    const start = lines.findIndex((l) => l.includes("case 'archive_email':"));
    const end = lines.findIndex((l, i) => i > start && /^\s*case '/.test(l));
    const body = lines.slice(start, end > start ? end : start + 60).join('\n');
    assert.match(
      body,
      /redactedJson\(\{\s*counts: result\.counts,\s*results: result\.results\s*\}/,
      'the archive_email JSON content item must carry counts as well as results',
    );
    // And it must go through redactedJson rather than redacting the finished document:
    // BEARER_PATTERN runs to the next whitespace, so over serialised JSON it eats the quote
    // and comma that terminate a value and the item stops parsing. See redactedJson in
    // coerce.ts; the failing input is a mailbox named "Bearer Bonds".
    // Comment lines are stripped first: the code carries a comment naming the wrong form in
    // order to warn against it, and a negative match over the raw text would fire on the
    // warning rather than on a real regression.
    const code = body.split('\n').filter((l) => !l.trimStart().startsWith('//')).join('\n');
    assert.doesNotMatch(
      code,
      /redactBearerTokens\(\s*JSON\.stringify/,
      'the archive_email JSON item must not be redacted after serialisation',
    );
  });
});

describe('scope flags are wired to their own argument', () => {
  it('reads every coerceBool flag from the argument of the same name', () => {
    // `const raw = coerceBool((args as any).raw)`, `raw: coerceBool(args?.raw)` and the
    // `a.` receiver the compose handlers use. A coerceBool over a bare identifier (an
    // already-destructured or renamed local) has no argument name to compare against and
    // is deliberately not matched.
    const pattern =
      /(?:const\s+([A-Za-z_][A-Za-z0-9_]*)\s*=|([A-Za-z_][A-Za-z0-9_]*)\s*:)\s*coerceBool\(\s*(?:\(args as any\)|args\?|args|a)\.([A-Za-z_][A-Za-z0-9_]*)\s*\)/;
    const pairs: string[] = [];
    const mismatches: string[] = [];
    for (const file of HANDLER_FILES) {
      readLines(file).forEach((line, i) => {
        if (line.trimStart().startsWith('//')) return;
        const match = pattern.exec(line);
        if (!match) return;
        const target = match[1] ?? match[2];
        const argument = match[3];
        pairs.push(`${target}<-${argument}`);
        if (target !== argument) mismatches.push(`src/${file}:${i + 1} assigns ${target} from ${argument}`);
      });
    }
    assert.ok(
      pairs.length >= 20,
      `matched only ${pairs.length} coerceBool argument reads; the scan has probably stopped matching`,
    );
    assert.deepEqual(
      mismatches,
      [],
      `a boolean is read from an argument of a different name, which silently applies the ` +
        `caller's answer to the wrong parameter: ${mismatches.join(', ')}`,
    );
  });
});

// Directories under src/ that hold no shipped code. Named, not inferred: the previous
// version of this scan was flat, so EVERY subdirectory was skipped by accident and a
// src/<subdir>/handler.ts would have gone unscanned forever. Listing the exclusion by name
// means a new subdirectory is scanned by default and dropping one is a deliberate edit.
const NON_SHIPPED_DIRS = new Set(['testing']);

// Every shipped source file, recursively: any handler or formatter can serialise a payload,
// wherever it lives. Read as text out of src/ rather than imported from dist/, for the same
// reason as the scans above: `npm test` never builds first, so a dist/ read would miss the
// site just added. Paths come back relative to src/ ('coerce.ts', 'sub/thing.ts').
function collectSourceFiles(dir: string = SRC_DIR, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true })
    .flatMap((entry) => {
      if (entry.isDirectory()) {
        return NON_SHIPPED_DIRS.has(entry.name)
          ? []
          : collectSourceFiles(join(dir, entry.name), `${prefix}${entry.name}/`);
      }
      const shipped = entry.name.endsWith('.ts') && !entry.name.endsWith('.test.ts');
      return shipped ? [`${prefix}${entry.name}`] : [];
    })
    .sort();
}

// The start of a JSON.stringify CALL. `JSON?.stringify(...)` reaches exactly the same
// function, and optional chaining is used throughout this codebase, so a pattern matching
// only the plain dot would report an indented `JSON?.stringify(v, null, 2)` neither as a
// three-argument call nor as an unlisted bare one - invisible to both assertions at once.
// Shared with blankLiterals, which uses it to refuse to blank a span that contains a call.
const CALL_START = String.raw`\bJSON\s*\??\.\s*stringify\s*\(`;

// A reference to JSON.stringify that is NOT a call: `const j = JSON.stringify` and then
// `j(v, null, 2)` later. Aliasing defeats both assertions, and the fix is cheap because a
// bare reference has no legitimate use here - nothing in src/ passes JSON.stringify as a
// callback - so the reference itself is the thing to report.
const BARE_REFERENCE = new RegExp(String.raw`\bJSON\s*\??\.\s*stringify\b(?!\s*\()`, 'g');

// Blank every comment and every string / template / regex literal in `source`, preserving
// both length and newlines so offsets and line numbers still match the original text.
//
// The line-by-line regex this replaced could only see a call written on one line in one
// shape, and dropped comments by testing for a leading `//`. Blanking first is what lets the
// scan below count a call's arguments by walking brackets instead: after this pass, every
// comma, quote and paren left in the text is syntax, so a comma inside a string, a comment
// or a regex cannot be mistaken for an argument separator - and a JSON.stringify quoted in
// prose cannot be mistaken for a call. Text inside a `${}` interpolation is deliberately
// left intact, because real call sites live there (contact-card.ts builds its dropped-value
// sentence that way).
function blankLiterals(source: string): string {
  const out = [...source];
  const blankAt = (i: number) => {
    if (i < out.length && out[i] !== '\n') out[i] = ' ';
  };
  // Consume template-literal TEXT from `i` (just past a backtick or the `}` closing an
  // interpolation), blanking it. Reports where to resume and whether a `${` was entered.
  const scanTemplateText = (start: number): { next: number; entersExpression: boolean } => {
    let i = start;
    while (i < source.length) {
      if (source[i] === '\\') {
        blankAt(i);
        blankAt(i + 1);
        i += 2;
        continue;
      }
      if (source[i] === '`') return { next: i + 1, entersExpression: false };
      if (source[i] === '$' && source[i + 1] === '{') return { next: i + 2, entersExpression: true };
      blankAt(i);
      i++;
    }
    return { next: i, entersExpression: false };
  };
  // A `/` opens a regex only where a value cannot already have ended; after an identifier,
  // a number, `)`, `]` or a closing backtick it is division. `}` counts as a statement end,
  // so a regex may open after it.
  //
  // The single character in front of the `/` does not settle it, and reading only that
  // character got both directions wrong: `a++ / b` ends in `+` yet is division, and
  // `return /re/` ends in an identifier character yet is a regex. Misreading a division as a
  // regex is the dangerous one, because blanking it swallows every character up to the next
  // `/` on the line - a real call site can vanish, and the scan then passes by finding
  // nothing, which is the exact rot this whole guard exists to avoid. So look at the whole
  // preceding TOKEN, and see the second defence at the blanking site below.
  const REGEX_MAY_FOLLOW_KEYWORD = new Set([
    'await', 'case', 'delete', 'do', 'else', 'in', 'instanceof', 'new', 'of', 'return',
    'throw', 'typeof', 'void', 'yield',
  ]);
  // Reads backwards through `out`, not `source`: everything behind the cursor has already
  // been blanked, so a comment or a string in the way is whitespace and the token in front of
  // it is what we land on. Both delimiters of a blanked string / template / regex survive, so
  // a value that ended in one is still recognisable as a value.
  const regexCanOpen = (at: number): boolean => {
    let j = at - 1;
    while (j >= 0 && /\s/.test(out[j])) j--;
    if (j < 0) return true;
    // `++` and `--` end a value, so what follows one is division.
    if ((out[j] === '+' || out[j] === '-') && out[j - 1] === out[j]) return false;
    if (/[A-Za-z0-9_$]/.test(out[j])) {
      let k = j;
      while (k >= 0 && /[A-Za-z0-9_$]/.test(out[k])) k--;
      return REGEX_MAY_FOLLOW_KEYWORD.has(out.slice(k + 1, j + 1).join(''));
    }
    return !/[)\]`]/.test(out[j]);
  };

  // Brace depth of each open `${`, so the `}` that closes one is not read as an object's.
  const templateExpressions: number[] = [];
  let braceDepth = 0;
  let i = 0;
  while (i < source.length) {
    const ch = source[i];
    const next = source[i + 1];
    if (ch === '/' && next === '/') {
      while (i < source.length && source[i] !== '\n') blankAt(i++);
      continue;
    }
    if (ch === '/' && next === '*') {
      const end = source.indexOf('*/', i + 2);
      const stop = end < 0 ? source.length : end + 2;
      while (i < stop) blankAt(i++);
      continue;
    }
    if (ch === '"' || ch === "'") {
      i++;
      // Bounded at a newline, because a `'`/`"` literal cannot contain a raw one: an odd
      // quote that is NOT opening a string - the apostrophe in a comment the comment scan
      // never saw, or one inside a regex this walk left unblanked - would otherwise open a
      // scan that runs to the next quote ANYWHERE in the file and blanks every line between,
      // deleting real call sites and leaving the scan to pass by finding nothing. That is the
      // same silence the regex/division test below defends against, arriving from the other
      // side, and stopping at the newline costs nothing: it leaves every real literal in this
      // codebase blanked exactly as before.
      // The escape branch is bounded at the newline TOO. It blanks the escaped character as
      // well as the backslash, and `blankAt` refuses to blank a newline but still advances
      // past it - so a `\`-newline line continuation carried the scan onto the next line in
      // string mode and blanked the real call sites there, which is the exact silence the
      // bound above exists to close, arriving through the one branch that could skip it.
      while (i < source.length && source[i] !== ch && source[i] !== '\n') {
        if (source[i] === '\\' && source[i + 1] !== '\n') blankAt(i++);
        blankAt(i++);
      }
      i++;
      continue;
    }
    if (ch === '`' || (ch === '}' && templateExpressions.length > 0 && braceDepth - 1 === templateExpressions[templateExpressions.length - 1])) {
      if (ch === '}') {
        braceDepth--;
        templateExpressions.pop();
      }
      const scan = scanTemplateText(i + 1);
      if (scan.entersExpression) {
        templateExpressions.push(braceDepth);
        braceDepth++;
      }
      i = scan.next;
      continue;
    }
    if (ch === '{') {
      braceDepth++;
      i++;
      continue;
    }
    if (ch === '}') {
      braceDepth--;
      i++;
      continue;
    }
    if (ch === '/' && regexCanOpen(i)) {
      let j = i + 1;
      let inClass = false;
      while (j < source.length) {
        const c = source[j];
        if (c === '\\') {
          j += 2;
          continue;
        }
        if (c === '\n') break;
        if (c === '[') inClass = true;
        else if (c === ']') inClass = false;
        else if (c === '/' && !inClass) break;
        j++;
      }
      // Second defence, and the one that does not depend on getting the token test right.
      // No regex literal in this codebase contains a JSON.stringify call, so a candidate
      // span holding one is a division the token test could not settle - `arr[0]! / b` and
      // anything else where the character in front is genuinely ambiguous. Blanking it would
      // delete the very call site this scan exists to find, and delete it silently. Treat it
      // as division. Being wrong that way leaves a regex unblanked, which the argument walk
      // survives - but only because the string scan above stops at a newline: an unblanked
      // regex containing an odd `'` or `"` would otherwise open a string scan that ran past
      // the end of the line and blanked whole call sites, so this direction was silent too
      // until that bound was added. Being wrong the other way is silent unconditionally.
      if (j < source.length && source[j] === '/' && !new RegExp(CALL_START).test(source.slice(i + 1, j))) {
        for (let k = i + 1; k < j; k++) blankAt(k);
        i = j + 1;
        continue;
      }
      // No closing delimiter on this line: it was division after all.
    }
    i++;
  }
  return out.join('');
}

type StringifyCall = { line: number; args: number; snippet: string };

// Every JSON.stringify call site in `source`, with the number of arguments it passes.
// Arguments are counted by walking the bracket depth from the opening paren, so a call
// spread over several lines is one site, and a comma inside a replacer's parameter list,
// an array or an object literal is not counted.
function findStringifyCalls(source: string): StringifyCall[] {
  const code = blankLiterals(source);
  const calls: StringifyCall[] = [];
  const callStart = new RegExp(CALL_START, 'g');
  let match: RegExpExecArray | null;
  while ((match = callStart.exec(code)) !== null) {
    const open = match.index + match[0].length - 1;
    let depth = 0;
    let commas = 0;
    let end = code.length - 1;
    for (let i = open; i < code.length; i++) {
      const c = code[i];
      if (c === '(' || c === '[' || c === '{') depth++;
      else if (c === ')' || c === ']' || c === '}') {
        depth--;
        if (depth === 0) {
          end = i;
          break;
        }
      } else if (depth === 1 && c === ',') commas++;
    }
    // Literals are blanked, but their delimiters are not, so an argument list holding
    // nothing but a string still reads as non-empty here - only a genuinely empty call
    // has no arguments.
    const empty = code.slice(open + 1, end).trim() === '';
    calls.push({
      line: code.slice(0, match.index).split('\n').length,
      args: empty ? 0 : commas + 1,
      snippet: source.slice(match.index, end + 1).replace(/\s+/g, ' ').slice(0, 120),
    });
  }
  return calls;
}

// Every line holding a JSON.stringify that is referenced rather than called - the aliasing
// escape, `const j = JSON.stringify`. Reported as a problem in its own right, because the
// alias's later invocation is invisible to the argument walk above.
function findStringifyAliases(source: string): number[] {
  const code = blankLiterals(source);
  const lines: number[] = [];
  BARE_REFERENCE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = BARE_REFERENCE.exec(code)) !== null) {
    lines.push(code.slice(0, match.index).split('\n').length);
  }
  return lines;
}

// The JSON.stringify calls that are NOT payload serialisation, keyed by file with an exact
// count, so ADDING one to an already-listed file still has to be justified here. The count
// catches an addition and not a substitution: swapping contact-card.ts's prose-quoting call
// for a payload serialisation leaves the count at 2 and passes. Everything else has to go
// through toolJson / redactedJson, which is what keeps serialisation on two signatures that
// cannot take an indent - see the header for what routing through a seam does and does not
// buy, since toolJson itself redacts nothing.
const NON_PAYLOAD_STRINGIFY: Record<string, { count: number; why: string }> = {
  'coerce.ts': { count: 2, why: 'the seams themselves - toolJson and redactedJson' },
  'contact-card.ts': { count: 2, why: 'quotes a dropped/added value into a prose sentence' },
  'jmap-client.ts': { count: 1, why: 'the HTTP request body POSTed to the JMAP endpoint' },
  'response-formatters.ts': { count: 1, why: 'builds a Map key from a value tuple, never emitted' },
};

describe('result payloads are serialised compact', () => {
  const sourceFiles = collectSourceFiles();

  it('no JSON.stringify passes an indent argument', () => {
    const offenders: string[] = [];
    for (const file of sourceFiles) {
      const source = readFileSync(join(SRC_DIR, file), 'utf8');
      for (const call of findStringifyCalls(source)) {
        // Three arguments means an indent was passed, whatever the replacer looks like:
        // `null`, `undefined`, a named function or an inline arrow all reach the same
        // pretty-printed output, and only the third argument decides.
        if (call.args >= 3) offenders.push(`src/${file}:${call.line} (${call.snippet})`);
      }
    }
    assert.deepEqual(
      offenders,
      [],
      `a result payload is pretty-printed. Indentation is bytes the caller pays for that ` +
        `nothing parses (6.5-28.5% of a real payload depending on its shape, measured live ` +
        `on one account in 2026-08 - see the header comment). Serialise through ` +
        `toolJson - or redactedJson where a value could carry a credential - both in ` +
        `coerce.ts: ${offenders.join(', ')}`,
    );
  });

  it('every JSON.stringify outside the seam is a listed non-payload use', () => {
    const counts = new Map<string, StringifyCall[]>();
    for (const file of sourceFiles) {
      const calls = findStringifyCalls(readFileSync(join(SRC_DIR, file), 'utf8'));
      if (calls.length > 0) counts.set(file, calls);
    }
    const problems: string[] = [];
    // Aliasing first: `const j = JSON.stringify` then `j(v, null, 2)` passes both the
    // argument walk and the per-file count, because neither ever sees the invocation.
    for (const file of sourceFiles) {
      const aliases = findStringifyAliases(readFileSync(join(SRC_DIR, file), 'utf8'));
      if (aliases.length > 0) {
        problems.push(
          `src/${file} references JSON.stringify without calling it (${aliases.map((l) => `:${l}`).join(', ')}); ` +
            `an alias is invoked somewhere this scan cannot see it`,
        );
      }
    }
    for (const [file, calls] of counts) {
      const allowed = NON_PAYLOAD_STRINGIFY[file];
      if (!allowed) {
        problems.push(
          `src/${file} has ${calls.length} bare JSON.stringify call(s) ` +
            `(${calls.map((c) => `:${c.line}`).join(', ')})`,
        );
      } else if (calls.length !== allowed.count) {
        problems.push(
          `src/${file} has ${calls.length} bare JSON.stringify call(s), expected ${allowed.count} ` +
            `(${allowed.why}); the extra one is at ${calls.map((c) => `:${c.line}`).join(', ')}`,
        );
      }
    }
    // An entry that matches nothing means the scan has stopped seeing a site it used to see,
    // which is the failure mode a text scan rots into. Caught here so a broken scanner cannot
    // pass by finding nothing.
    for (const file of Object.keys(NON_PAYLOAD_STRINGIFY)) {
      if (!counts.has(file)) problems.push(`src/${file} is listed as a non-payload use but no JSON.stringify was found in it`);
    }
    assert.deepEqual(
      problems,
      [],
      `What is checked: every JSON.stringify call written literally in a shipped src/ file ` +
        `(recursively, excluding *.test.ts and src/testing/), spelled with either a plain dot ` +
        `or an optional one, must either serialise something that is not a tool result ` +
        `payload - and be listed in NON_PAYLOAD_STRINGIFY with the reason and an exact ` +
        `per-file count - or go through toolJson / redactedJson in coerce.ts. Referencing ` +
        `JSON.stringify without calling it is reported too, since an alias would be invoked ` +
        `out of sight. What is NOT checked, and this list is meant to be exhaustive: that ` +
        `the value handed to a seam is the whole payload; a per-file count catching a ` +
        `SUBSTITUTION rather than an addition; serialisation reached indirectly through some ` +
        `other helper that itself calls a seam; a destructured \`const { stringify } = JSON\` ` +
        `or a computed \`JSON['stringify']\`; arguments supplied by a spread ` +
        `(\`JSON.stringify(...[v, null, 2])\` walks as ONE argument, so the indent is invisible ` +
        `to the count); and anything in dist/. ` +
        `Findings: ${problems.join('; ')}`,
    );
  });
});

describe('the compact-serialisation scan sees the forms that used to escape it', () => {
  // The scan this pins replaced a line-by-line regex that only matched a `null`/`undefined`
  // replacer on a single line. Each offender case below is a real way that scan could be
  // defeated while shipping a pretty-printed payload; each allowed case is a form that must
  // NOT be reported, so the guard cannot be "fixed" into firing on prose.
  const offenders: [string, string][] = [
    ['a null replacer with an indent', 'const s = JSON.stringify(v, null, 2);'],
    ['no spaces at all', 'const s=JSON.stringify(v,null,2);'],
    ['a tab indent', "const s = JSON.stringify(v, null, '\\t');"],
    ['an indent held in a constant', 'const s = JSON.stringify(v, null, INDENT);'],
    ['a named replacer function', 'const s = JSON.stringify(v, redact, 2);'],
    ['an inline arrow replacer, whose parameter list contains a comma', 'const s = JSON.stringify(v, (k, val) => val, 2);'],
    ['the exact shape redactedJson had before the indent was removed', 'return JSON.stringify(value, (_key, v) => (typeof v === "string" ? redactBearerTokens(v) : v), 2);'],
    ['a call wrapped across several lines', 'const s = JSON.stringify(\n  payload,\n  null,\n  2\n);'],
    ['an object argument whose own commas are nested', 'const s = JSON.stringify({ a: 1, b: [2, 3] }, null, 2);'],
    ['a call inside a template interpolation', 'const s = `x ${JSON.stringify(v, null, 2)} y`;'],
    ['a call after a regex literal containing a quote and a paren', 'const re = /["(]/g;\nconst s = JSON.stringify(v, null, 2);'],
    // The three below were found by running the scan against them rather than by reading it,
    // and each was invisible to BOTH assertions - not merely miscounted. That is the shape of
    // hole worth pinning: the guard reported nothing and looked healthy.
    ['an optional call, which reaches the same function', 'const s = JSON?.stringify(v, null, 2);'],
    [
      // A `/` after `a++` is division, but the character in front of it is `+`, which reads
      // like an operator. Scanning to the next `/` on the line blanked the call in between.
      'a call on a line whose earlier division could be misread as a regex',
      'const r = a++ / b; const s = JSON.stringify(v, null, 2) / c;',
    ],
    [
      // A `/` after `return` opens a regex, but the character in front of it is a letter, so
      // the regex was read as division and left unblanked - and its `"` then opened a string
      // scan that ate the real call on the next line.
      'a call after a regex that follows a keyword rather than an operator',
      'function f() { return /["(]/g; }\nconst s = JSON.stringify(v, null, 2);',
    ],
    [
      // The same silence from the other direction, and the one the token test CANNOT settle:
      // after `)` a `/` really is division far more often than not, so a regex written there
      // is deliberately left unblanked - and the `"` inside it then opened a string scan that
      // ran past the end of the line and blanked every call below it. Bounding that scan at a
      // newline is what makes "left unblanked" survivable.
      'a call below an unblanked regex whose quote would otherwise swallow the rest of the file',
      'for (const a of b) /["(]/.test(a);\nconst s = JSON.stringify(v, null, 2);\nconst t = toolJson(w);',
    ],
    [
      // …and the one branch that could step over that newline anyway. The escape handling
      // inside the string scan blanks the backslash and then blanks-and-advances again,
      // landing on the next line still in string mode, so a single trailing backslash was
      // enough to resume the scan below and blank the real call sites there.
      'a call below a line whose trailing backslash would otherwise carry the string scan onto it',
      "const doc = 'x \\\nconst s = JSON.stringify(v, null, 2);\nconst t = toolJson(w);",
    ],
  ];
  for (const [label, source] of offenders) {
    it(`reports ${label}`, () => {
      const calls = findStringifyCalls(source);
      assert.equal(calls.length, 1, `expected one call site in: ${source}`);
      assert.ok(calls[0].args >= 3, `counted ${calls[0].args} argument(s) in: ${source}`);
    });
  }

  const allowed: [string, string, number][] = [
    ['one argument', 'const s = JSON.stringify(payload);', 1],
    ['a replacer and no indent, which is what redactedJson does', 'return JSON.stringify(value, (_key, v) => v);', 2],
    ['nested commas but a single argument', 'const s = JSON.stringify({ a: [1, 2], b: f(3, 4) });', 1],
  ];
  for (const [label, source, args] of allowed) {
    it(`counts ${label} correctly`, () => {
      assert.deepEqual(findStringifyCalls(source).map((c) => c.args), [args]);
    });
  }

  const invisible: [string, string][] = [
    ['a line comment naming the forbidden form', '// never write JSON.stringify(v, null, 2) here\nconst s = toolJson(v);'],
    ['a block comment naming it', '/* JSON.stringify(v, null, 2) */\nconst s = toolJson(v);'],
    ['a string literal containing it', "const doc = 'JSON.stringify(v, null, 2)';"],
    ['a template literal containing it', 'const doc = `JSON.stringify(v, null, 2)`;'],
  ];
  for (const [label, source] of invisible) {
    it(`does not report ${label}`, () => {
      assert.deepEqual(findStringifyCalls(source), []);
    });
  }

  it('reports the line a multi-line call starts on', () => {
    const source = 'const a = 1;\nconst b = 2;\nconst s = JSON.stringify(\n  v,\n  null,\n  2\n);';
    assert.equal(findStringifyCalls(source)[0].line, 3);
  });

  // The third hole. An alias is a reference, so it is not a call site and the argument walk
  // will never see the indent - the reference itself has to be the thing reported.
  it('reports an aliased JSON.stringify, which no argument count can see', () => {
    const source = 'const j = JSON.stringify;\nconst s = j(v, null, 2);';
    assert.deepEqual(findStringifyCalls(source), []);
    assert.deepEqual(findStringifyAliases(source), [1]);
  });

  it('reports an optionally-chained alias too', () => {
    assert.deepEqual(findStringifyAliases('const j = JSON?.stringify;'), [1]);
  });

  const notAliases: [string, string][] = [
    ['an ordinary call', 'const s = JSON.stringify(v);'],
    ['an optional call', 'const s = JSON?.stringify(v);'],
    ['a call with the paren on the next line', 'const s = JSON.stringify\n  (v);'],
    ['the name quoted in prose', "const doc = 'JSON.stringify';"],
    ['the name in a comment', '// JSON.stringify\nconst s = toolJson(v);'],
  ];
  for (const [label, source] of notAliases) {
    it(`does not read ${label} as an alias`, () => {
      assert.deepEqual(findStringifyAliases(source), []);
    });
  }
});

describe('the compact-serialisation scan reaches every shipped file', () => {
  it('descends into subdirectories and excludes src/testing by name', () => {
    // Built in a temp tree rather than asserted against src/ itself, because src/ happens to
    // have no shipped subdirectory today - which is exactly how the flat readdirSync that
    // preceded this looked correct while skipping every subdirectory there could ever be.
    const root = mkdtempSync(join(tmpdir(), 'srcscan-'));
    try {
      mkdirSync(join(root, 'sub'));
      mkdirSync(join(root, 'testing'));
      writeFileSync(join(root, 'shipped.ts'), '');
      writeFileSync(join(root, 'shipped.test.ts'), '');
      writeFileSync(join(root, 'notes.md'), '');
      writeFileSync(join(root, 'sub', 'nested.ts'), '');
      writeFileSync(join(root, 'sub', 'nested.test.ts'), '');
      writeFileSync(join(root, 'testing', 'mock-calls.ts'), '');
      assert.deepEqual(collectSourceFiles(root), ['shipped.ts', 'sub/nested.ts']);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it('scans the shipped files that exist today', () => {
    const files = collectSourceFiles();
    assert.ok(files.includes('index.ts') && files.includes('coerce.ts'), files.join(', '));
    assert.ok(!files.some((f) => f.endsWith('.test.ts') || f.startsWith('testing/')), files.join(', '));
  });
});
