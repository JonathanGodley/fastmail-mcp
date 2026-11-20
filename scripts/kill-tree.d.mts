// Types for kill-tree.mjs. Hand-written: keep in step with its exports.

import type { ChildProcess } from 'node:child_process';

/** Whether a run is spawned leading its own process group, which is how it is killed off Windows. */
export declare const DETACHED: boolean;

/** Kills `child` and every process under it. Returns why the kill failed, or null. */
export declare function killTree(child: ChildProcess): string | null;
