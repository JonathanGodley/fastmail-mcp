// Types for the pure helpers mutation-test.mjs exports to its unit test. Hand-written: keep in
// step with the script's exports.

export const USAGE: string;
export const EXCLUDED_TESTS: string[];
export const MUTATE_ALL: string[];
export function parseArgs(argv: string[]): { all: true } | { commit: string } | { error: string };
export function isMutable(file: string): boolean;
export function testFiles(srcNames: string[]): string[];
export function diffToRanges(diff: string): string[];
