// Generate a durable, script-produced reference of the official Fastmail MCP's
// surface — every tool and its full input schema — straight from the live
// `tools/list`. Uses ONLY protocol methods (initialize, notifications/initialized,
// tools/list): it makes NO tools/call, so no account data is ever fetched and the
// output is purely the server's public tool surface (safe to check in).
//
// Outputs (overwrites in place, so re-running refreshes them):
//   docs/official-mcp-tools.json   raw tools/list result — machine-readable, full fidelity
//   docs/official-mcp-surface.md   human-readable index of methods + fields
//
// Usage:  FASTMAIL_MCP_TOKEN=… node scripts/dump-official-surface.mjs
// (references the token var NAME only; never prints its value)

import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO = join(HERE, '..');
const ENDPOINT = 'https://api.fastmail.com/mcp';
const PROTOCOL_VERSION = '2025-06-18';

function token() {
  const t = process.env.FASTMAIL_MCP_TOKEN;
  if (!t) throw new Error('FASTMAIL_MCP_TOKEN not set in env');
  return t;
}
function headers(sid) {
  const h = {
    Authorization: `Bearer ${token()}`,
    'Content-Type': 'application/json',
    Accept: 'application/json, text/event-stream',
  };
  if (sid) h['Mcp-Session-Id'] = sid;
  return h;
}
// A response may be plain JSON or SSE (data: frames). Return the JSON-RPC message
// matching wantId (or the first message if wantId is null).
async function parse(res, wantId) {
  const ctype = (res.headers.get('content-type') || '').toLowerCase();
  const text = await res.text();
  if (ctype.includes('text/event-stream')) {
    const messages = [];
    for (const frame of text.split(/\r?\n\r?\n/)) {
      const data = [];
      for (const line of frame.split(/\r?\n/)) {
        if (line.startsWith('data:')) data.push(line.slice(5).replace(/^ /, ''));
      }
      if (!data.length) continue;
      try { messages.push(JSON.parse(data.join('\n'))); } catch { /* skip */ }
    }
    return wantId == null ? (messages[0] ?? null) : (messages.find((m) => m && m.id === wantId) ?? null);
  }
  if (!text.trim()) return null;
  return JSON.parse(text);
}

let sid = null;
let nextId = 1;
async function rpc(method, params, { notification = false } = {}) {
  const id = notification ? undefined : nextId++;
  const body = notification ? { jsonrpc: '2.0', method, params } : { jsonrpc: '2.0', id, method, params };
  const res = await fetch(ENDPOINT, { method: 'POST', headers: headers(sid), body: JSON.stringify(body) });
  const newSid = res.headers.get('mcp-session-id');
  if (newSid) sid = newSid;
  if (notification) return null;
  const msg = await parse(res, id);
  if (!res.ok) throw new Error(`${method} HTTP ${res.status}`);
  if (msg && msg.error) throw new Error(`${method} error: ${msg.error.code} ${msg.error.message}`);
  return msg ? msg.result : null;
}

// --- bootstrap (protocol only, no tool calls) ---
const init = await rpc('initialize', {
  protocolVersion: PROTOCOL_VERSION,
  capabilities: {},
  clientInfo: { name: 'surface-dump', version: '1.0' },
});
await rpc('notifications/initialized', {}, { notification: true });
const listResult = await rpc('tools/list', {});
const tools = (listResult.tools || []).slice().sort((a, b) => a.name.localeCompare(b.name));

// --- raw machine-readable artifact (full fidelity) ---
const generatedAt = new Date().toISOString().slice(0, 10);
const rawArtifact = {
  generatedAt,
  source: ENDPOINT,
  serverInfo: init.serverInfo,
  protocolVersion: init.protocolVersion,
  capabilities: init.capabilities,
  toolCount: tools.length,
  tools,
};
fs.writeFileSync(join(REPO, 'docs', 'official-mcp-tools.json'), JSON.stringify(rawArtifact, null, 2) + '\n');

// --- human-readable index ---
function fieldRows(schema) {
  const props = schema?.properties || {};
  const required = new Set(schema?.required || []);
  const names = Object.keys(props).sort((a, b) => {
    // required first, then alpha
    const ra = required.has(a), rb = required.has(b);
    if (ra !== rb) return ra ? -1 : 1;
    return a.localeCompare(b);
  });
  if (!names.length) return '_(no input fields)_\n';
  let out = '| field | type | required | allowed values | description |\n|---|---|---|---|---|\n';
  for (const n of names) {
    const p = props[n] || {};
    let type = p.type || (p.anyOf ? p.anyOf.map((x) => x.type).filter(Boolean).join('\\|') : '?');
    if (type === 'array' && p.items?.type) type = `array<${p.items.type}>`;
    const enumv = p.enum ? p.enum.map((v) => `\`${v}\``).join(', ') : (p.items?.enum ? p.items.enum.map((v) => `\`${v}\``).join(', ') : '');
    const desc = (p.description || '').replace(/\s+/g, ' ').replace(/\|/g, '\\|').trim();
    out += `| \`${n}\` | ${type} | ${required.has(n) ? '**yes**' : 'no'} | ${enumv} | ${desc} |\n`;
  }
  return out;
}

let md = '';
md += '# Official Fastmail MCP — tool surface (methods + fields)\n\n';
md += '> **Generated by `scripts/dump-official-surface.mjs` from the live `tools/list`** — do not hand-edit; re-run the script to refresh.\n';
md += `> Source: \`${ENDPOINT}\` · server **${init.serverInfo?.name} v${init.serverInfo?.version}** · protocol \`${init.protocolVersion}\` · **${tools.length} tools** · generated **${generatedAt}**.\n`;
md += '>\n';
md += `> Capabilities: \`${JSON.stringify(init.capabilities)}\`. This file is the official server's public tool surface only — no account data is read to produce it. The machine-readable twin is \`docs/official-mcp-tools.json\`.\n\n`;

md += '## Tool index\n\n';
md += '| tool | required fields | all fields |\n|---|---|---|\n';
for (const t of tools) {
  const props = t.inputSchema?.properties || {};
  const req = (t.inputSchema?.required || []).map((f) => `\`${f}\``).join(', ') || '—';
  const all = Object.keys(props).map((f) => `\`${f}\``).join(', ') || '—';
  md += `| **${t.name}** | ${req} | ${all} |\n`;
}
md += '\n---\n\n## Tools in detail\n\n';
for (const t of tools) {
  md += `### \`${t.name}\`\n\n`;
  if (t.description) md += `${t.description.replace(/\s+/g, ' ').trim()}\n\n`;
  const ann = t.annotations || {};
  const annBits = [];
  if (ann.readOnlyHint !== undefined) annBits.push(`readOnlyHint=${ann.readOnlyHint}`);
  if (ann.destructiveHint !== undefined) annBits.push(`destructiveHint=${ann.destructiveHint}`);
  if (ann.idempotentHint !== undefined) annBits.push(`idempotentHint=${ann.idempotentHint}`);
  if (t._meta) annBits.push('has `_meta` (widget/UI)');
  if (annBits.length) md += `*annotations: ${annBits.join(', ')}*\n\n`;
  md += fieldRows(t.inputSchema);
  if (t.outputSchema) {
    md += '\n**Output schema (server-declared):**\n';
    md += fieldRows(t.outputSchema);
  }
  md += '\n';
}
fs.writeFileSync(join(REPO, 'docs', 'official-mcp-surface.md'), md);

// --- console summary (no account data) ---
console.log(`server: ${init.serverInfo?.name} v${init.serverInfo?.version} | protocol ${init.protocolVersion} | tools ${tools.length}`);
console.log(`tools-with-outputSchema: ${tools.filter((t) => t.outputSchema).length}`);
console.log('wrote docs/official-mcp-tools.json and docs/official-mcp-surface.md');
console.log('tool names:', tools.map((t) => t.name).join(', '));
