// Killing a test run means killing its whole process tree. tsx runs the tests in a child node
// process of its own, and `node --test` runs each file in a further child, so killing the
// process this module spawned leaves the rest running and holding its pipes. Off Windows the
// child is spawned leading its own process group (`spawn(..., { detached: DETACHED })`) and
// the group is killed; on Windows taskkill /T kills the tree.

import { spawnSync } from 'node:child_process';

export const DETACHED = process.platform !== 'win32';

// Returns why the kill failed, or null. A tree that is already gone is not a failure.
export function killTree(child) {
  if (child.exitCode !== null || child.signalCode !== null) return null;
  if (DETACHED) {
    try {
      process.kill(-child.pid, 'SIGKILL');
    } catch (e) {
      if (e.code !== 'ESRCH') return `kill of process group ${child.pid} failed: ${e.code ?? e.message}`;
    }
    return null;
  }
  const r = spawnSync('taskkill', ['/pid', String(child.pid), '/T', '/F'], { encoding: 'utf8' });
  // taskkill exits 128 when the process is not found.
  if (r.error) return `taskkill failed: ${r.error.message}`;
  if (r.status !== 0 && r.status !== 128) return `taskkill /pid ${child.pid} exited ${r.status}: ${(r.stderr || r.stdout).trim()}`;
  return null;
}
