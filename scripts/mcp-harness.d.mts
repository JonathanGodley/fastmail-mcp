// Types for the JSON-RPC harness in mcp-harness.mjs.
//
// The harness is plain JavaScript on purpose: it spawns the BUILT server
// (dist/index.js) and is run directly by node, both from the on-demand live probes
// under scripts/probes/ and from the CLI. A declaration file next to it, rather
// than a declaration scoped inside the one test that imports it, is what makes the
// contract visible to every consumer — the probes are not typechecked, so this is
// also the only written statement of the harness's surface they can be read
// against.
//
// It is hand-written, so it can drift: keep it in step with the exports of
// mcp-harness.mjs. The typecheck over the test corpus is what catches drift on the
// side that matters, since built-server.test.ts is compiled against it.

/** Options for spawning the built server. */
export interface HarnessOptions {
  /**
   * Environment for the child process. Defaults to `process.env`. The harness
   * passes it through untouched and never reads or prints an individual value.
   */
  env?: NodeJS.ProcessEnv;
}

/** A JSON-RPC client bound to one spawned server process. */
export interface HarnessClient {
  /** Send `initialize`, then the `initialized` notification. Resolves with the result. */
  init(): Promise<unknown>;
  /**
   * Send `tools/list`. Needs no credentials and reaches no account, so it is the
   * safe first call when checking what the built server declares.
   */
  list(): Promise<unknown>;
  /**
   * Send `tools/call`. Generic, and therefore able to MUTATE the account
   * (send_draft, delete_email, bulk_* …) — pick the tool name deliberately.
   * Rejects when the server returns a JSON-RPC error.
   */
  call(name: string, args?: object): Promise<unknown>;
  /** End stdin and kill the child. Safe to call more than once. */
  close(): void;
}

/** Spawn the built server (dist/index.js) and return a client for it. */
export function createClient(opts?: HarnessOptions): HarnessClient;
