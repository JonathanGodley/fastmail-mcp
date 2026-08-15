// Two drift guards over the tool schemas in src/index.ts: the lenient-boolean convention
// (#54), and the rule that a tool emitting a recovery note has to expose every parameter
// that note prescribes.
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

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
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
  'reply-handler.ts',
  'forward-handler.ts',
  'compose-handler.ts',
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
