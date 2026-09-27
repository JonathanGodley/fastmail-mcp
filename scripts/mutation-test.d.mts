// Types for the pure helpers mutation-test.mjs exports to its unit test. Hand-written: keep in
// step with the script's exports.

export const USAGE: string;
export const EXCLUDED_TESTS: string[];
export const MUTATE_ALL: string[];
export const TEST_HEAP_MB: number;
export function parseArgs(argv: string[]):
  | { all: true; shard?: { i: number; n: number } }
  | { commit: string }
  | { error: string };
export function partition(files: [path: string, bytes: number][], n: number): string[][];
export function isMutable(file: string): boolean;
export function testFiles(srcNames: string[]): string[];
export function diffToRanges(diff: string): string[];
export function reportOutputs(args: { all?: true; commit?: string; shard?: { i: number; n: number } }):
  { reporters: string[]; tag: string };
