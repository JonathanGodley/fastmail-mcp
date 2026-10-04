// Loads <home>/.fastmail-mcp/.env into process.env. index.ts imports this module before any
// other, so the file is in place before any module, dependencies included, reads a setting
// at import. It therefore imports only Node built-ins.

import { homedir } from 'node:os';
import { join } from 'node:path';

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
  try {
    process.loadEnvFile(path);
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === 'ENOENT') return;
    throw new Error(`${path} exists but could not be loaded: ${(err as Error).message}`, { cause: err });
  }
}

try {
  loadHomeEnvFile(homedir());
} catch (err) {
  // A message, not a stack trace: this runs at import, before the server can report anything.
  console.error(`Fastmail MCP server failed to start: ${(err as Error).message}`);
  process.exit(1);
}
