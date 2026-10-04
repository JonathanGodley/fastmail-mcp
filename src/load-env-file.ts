// Loads <home>/.fastmail-mcp/.env into process.env. index.ts imports this module before any
// other, so the file is in place before any module, dependencies included, reads a setting
// at import. It therefore imports only Node built-ins.

import { closeSync, openSync, statSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

function isDirectory(path: string): boolean {
  try {
    return statSync(path).isDirectory();
  } catch {
    return false;
  }
}

export function envFilePath(home: string): string {
  return join(home, '.fastmail-mcp', '.env');
}

/**
 * Fills unset variables from the file. A variable already in the environment wins even when
 * blank: process.loadEnvFile never overwrites one, on every Node version CI runs
 * (load-env-file.test.ts). A missing file is silent; any other failure throws naming the path.
 */
export function loadHomeEnvFile(home: string): void {
  const path = envFilePath(home);
  // Opened here first because process.loadEnvFile reports a file it cannot open as ENOENT, the
  // code for a missing file (measured on Windows; on Linux CI an unreadable file was likewise
  // indistinguishable from a missing one).
  try {
    closeSync(openSync(path, 'r'));
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return;
    throw loadError(path, err);
  }
  try {
    process.loadEnvFile(path);
  } catch (err) {
    throw loadError(path, err);
  }
}

function loadError(path: string, err: unknown): Error {
  // Node's own error for a directory differs by platform, and on Windows misstates the cause.
  if (isDirectory(path)) return new Error(`${path} is a directory, not a file`, { cause: err });
  return new Error(`${path} exists but could not be loaded: ${(err as Error).message}`, { cause: err });
}

try {
  loadHomeEnvFile(homedir());
} catch (err) {
  // A message, not a stack trace: this runs at import, before the server can report anything.
  console.error(`Fastmail MCP server failed to start: ${(err as Error).message}`);
  process.exit(1);
}
