# Embedding reference

## Embedding this server as a spawned child

The README describes registering the server with a terminal MCP client, which
keeps one long-lived connection. If instead your own program spawns this server per
task - spawn, initialize, run a few tool calls, close stdin - the following applies.

### Spawn node directly, never through a shell wrapper

Spawn the entry point with node itself:

```
node /path/to/fastmail-mcp/dist/index.js
```

Do **not** wrap it in a shell, and on Windows specifically **never** launch it via
`cmd /c node ...`. `cmd.exe` becomes an intermediary process: killing what you
spawned kills `cmd.exe` and leaves the node grandchild running, unparented and
invisible, holding its stdio pipes. The wrapper styles that are perfectly fine for
a terminal MCP client (`cmd /c`, `npx`, a `.cmd` shim, a shell one-liner) are all
wrong for an embedding caller, because an embedding caller has to be able to
terminate what it started.

For the same reason, prefer the absolute path to `dist/index.js` over `npx`: `npx`
adds its own process layer and a package-resolution step to every spawn.

### Closing stdin ends the process

Closing the child's stdin is the normal shutdown signal. On stdin EOF the server
finishes what it is doing and exits with code 0, typically within tens of
milliseconds; no kill is needed. This holds whether or not the handshake ever
completed, so a caller that gives up mid-spawn can still close stdin and walk away.

A defensive caller should keep the usual belt: after closing stdin, wait a
short grace period (a second is ample) and force-kill the child if it is still
alive.

### Startup-to-ready latency

Measured on Windows 11 with Node 25, spawning `dist/index.js` directly: the
`initialize` response comes back in roughly **0.3 seconds** (15 warm runs: median
0.33s, slowest 0.43s). Nothing in startup touches the network or validates
credentials - authentication is resolved lazily on the first tool call - so the
figure is dominated by node's own startup and module loading, and does not vary
with account size or connectivity.

Set the handshake timeout to **10 seconds**. That is far above the observed range,
which leaves room for a cold module cache, an on-access virus scanner, or a loaded
machine, while still being short enough that a genuinely stuck child cannot hang
your own tool call. If `initialize` has not answered inside that window, treat the
spawn as failed and kill the child rather than waiting longer.
