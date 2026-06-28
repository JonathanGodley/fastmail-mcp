import { McpError, ErrorCode } from '@modelcontextprotocol/sdk/types.js';

// Tagged error for filesystem-path access decisions (path confinement and the
// attachment opt-in gate). Thrown by the path guards and attachment upload in
// jmap-client.ts, which deliberately stays free of MCP SDK types — the index
// boundary maps every PathAccessError to McpError(InvalidParams). instanceof is
// the discriminator, so the message text carries no routing burden.
export class PathAccessError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PathAccessError';
  }
}

// Tagged error for caller-supplied input that is well-formed JSON but semantically
// invalid (e.g. a `mailbox` that resolves to nothing, or a label `mailboxIds`
// element that isn't a real id). Thrown from jmap-client.ts (which stays free of
// MCP SDK types); the index boundary maps every InvalidInputError to
// McpError(InvalidParams), mirroring PathAccessError. instanceof is the
// discriminator. Unlike the PathAccessError branch, the index mapping runs this
// message through redactBearerTokens — these messages can reflect caller input
// and mailbox names, so redaction is cheap defense-in-depth against a
// token-shaped echo (it is NOT what makes the reflected-input oracle acceptable;
// see docs/security-model.md).
export class InvalidInputError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'InvalidInputError';
  }
}

// Some MCP clients (e.g. Claude Cowork as of 2026-04-08, issue #54) stringify
// structured params before dispatch. These helpers coerce such values back to
// their expected shapes so the handlers work against both strict and lenient clients.

// Defense-in-depth: scrub bearer-token-shaped substrings from any string that
// might be reflected back to the MCP caller (e.g. a JMAP error message). This
// is intentionally narrow — provider error messages are useful for the LLM to
// recover from, so we don't want to over-sanitize.
const BEARER_PATTERN = /Bearer\s+\S+/gi;
const FASTMAIL_TOKEN_PATTERN = /fmu\d+-[A-Za-z0-9-]{20,}/g;

export function redactBearerTokens(input: string): string {
  return input
    .replace(BEARER_PATTERN, 'Bearer [REDACTED]')
    .replace(FASTMAIL_TOKEN_PATTERN, 'fmu[REDACTED]');
}

export function coerceStringArray(value: unknown): string[] | undefined {
  if (value === undefined || value === null) return undefined;
  if (Array.isArray(value)) return value.map(String);
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  if (!trimmed) return [];
  if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return parsed.map(String);
    } catch { /* fall through to comma-split */ }
  }
  return trimmed.split(',').map(s => s.trim()).filter(Boolean);
}

// Coerce the four recipient list fields from whatever shape a (possibly lenient)
// client sent into string[] | undefined, so the JMAP client's .map(parseAddress)
// calls never receive a bare string (issue #54). Pass the raw tool args; reads
// only to/cc/bcc/replyTo and returns the coerced quartet.
export function coerceRecipients(args: { to?: unknown; cc?: unknown; bcc?: unknown; replyTo?: unknown }): {
  to?: string[]; cc?: string[]; bcc?: string[]; replyTo?: string[];
} {
  return {
    to: coerceStringArray(args.to),
    cc: coerceStringArray(args.cc),
    bcc: coerceStringArray(args.bcc),
    replyTo: coerceStringArray(args.replyTo),
  };
}

// Hard-reject any argument key the tool didn't declare in its inputSchema, so a
// misspelled/hallucinated param (e.g. `mailbox` vs `mailboxId`) fails loudly
// instead of being silently dropped and the tool running with defaults (#11).
// KEY-strictness only — value coercion is handled separately and is untouched.
// `additionalProperties: true` on a tool's schema opts that tool out (none today).
export function assertKnownParams(
  toolName: string,
  args: Record<string, unknown> | null | undefined,
  allowedKeys: Set<string>,
  additionalProperties: boolean,
): void {
  if (additionalProperties) return;
  if (args === null || args === undefined) return;
  const unknown = Object.keys(args).filter(k => !allowedKeys.has(k));
  if (unknown.length === 0) return;
  throw new McpError(
    ErrorCode.InvalidParams,
    `Unknown parameter(s): ${unknown.join(', ')}. Valid: ${[...allowedKeys].join(', ')}`,
  );
}

export function coerceBool(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') return value;
  if (value === 'true') return true;
  if (value === 'false') return false;
  return undefined;
}

// Loud-reject a settable string field that was provided but is empty,
// whitespace-only, or null. Callers invoke this only for fields that were
// actually present (i.e. !== undefined at the call site), so silently omitting
// a field stays distinct from explicitly blanking it. Returns the trimmed value.
export function requireNonEmpty(value: unknown, fieldName: string, hint = 'omit the field to leave it unchanged'): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${fieldName} cannot be empty; ${hint}`);
  }
  return value.trim();
}

// Validate a clearFields list: every entry must be in the allowed set, and no
// entry may also appear as a settable param (can't both set and clear a field).
// No-op when clearFields is empty/undefined.
export function validateClearFields(clearFields: string[] | undefined, allowed: Set<string>, provided: Set<string>): void {
  if (!clearFields || clearFields.length === 0) return;
  for (const field of clearFields) {
    if (!allowed.has(field)) {
      throw new Error(`Cannot clear "${field}"; clearable fields are: ${[...allowed].join(', ')}`);
    }
    if (provided.has(field)) {
      throw new Error(`cannot both set and clear ${field}; pass it as a value or in clearFields, not both`);
    }
  }
}

// Parse an RFC 5322 "Display Name <email>" recipient string into a JMAP
// EmailAddress object. Bare addresses pass through as { email }, and a blank
// display name is omitted. This is a pragmatic parse, not the full RFC grammar.
// Callers map it over already-trimmed, non-empty arrays (coerceStringArray
// filters blanks), so input is assumed non-empty.
export function parseAddress(input: string): { name?: string; email: string } {
  const trimmed = String(input).trim();
  const open = trimmed.lastIndexOf('<');
  const close = trimmed.lastIndexOf('>');
  if (open !== -1 && close > open) {
    const email = trimmed.slice(open + 1, close).trim();
    let name = trimmed.slice(0, open).trim();
    // Strip one pair of surrounding double-quotes from a quoted display name.
    if (name.length >= 2 && name.startsWith('"') && name.endsWith('"')) {
      name = name.slice(1, -1).trim();
    }
    return name ? { name, email } : { email };
  }
  return { email: trimmed };
}

// One file-to-attach spec as it arrives from a tool call (before path confinement
// and upload, which happen in jmap-client.ts).
export interface AttachmentSpec {
  path: string;
  name?: string;
  contentType?: string;
}

const ATTACHMENT_KEYS = new Set(['path', 'name', 'contentType']);

// Coerce the `attachments` tool param into AttachmentSpec[] | undefined. Accepts a
// real array, or a JSON-string array from lenient clients (mirroring
// coerceStringArray). Per element it REJECTS — never silently drops — a non-object,
// a spec missing `path`, or an unexpected per-item key, naming the index so the
// caller can fix it (assertKnownParams is top-level only and won't catch nested
// keys, so this is the sole guard for the item shape). A bare string element is
// rejected rather than guessed as a path (too magic); a JSON-object string is parsed.
export function coerceAttachments(value: unknown): AttachmentSpec[] | undefined {
  if (value === undefined || value === null) return undefined;

  let arr: unknown = value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return undefined;
    try {
      arr = JSON.parse(trimmed);
    } catch {
      throw new McpError(ErrorCode.InvalidParams, 'attachments must be an array of { path, name?, contentType? } objects.');
    }
  }

  if (!Array.isArray(arr)) {
    throw new McpError(ErrorCode.InvalidParams, 'attachments must be an array of { path, name?, contentType? } objects.');
  }

  const specs: AttachmentSpec[] = [];
  for (let i = 0; i < arr.length; i++) {
    let item: unknown = arr[i];
    if (typeof item === 'string') {
      const t = item.trim();
      if (t.startsWith('{') && t.endsWith('}')) {
        try {
          item = JSON.parse(t);
        } catch {
          throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] is a string that isn't valid JSON; pass an object with a path.`);
        }
      } else {
        throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object with a path, not a bare string.`);
      }
    }
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] must be an object with a path.`);
    }
    const obj = item as Record<string, unknown>;
    const unknownKeys = Object.keys(obj).filter(k => !ATTACHMENT_KEYS.has(k));
    if (unknownKeys.length > 0) {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] has unknown key(s): ${unknownKeys.join(', ')}. Valid: ${[...ATTACHMENT_KEYS].join(', ')}`);
    }
    if (typeof obj.path !== 'string' || obj.path.trim() === '') {
      throw new McpError(ErrorCode.InvalidParams, `attachments[${i}] is missing a non-empty 'path'.`);
    }
    // Trim the path (consistent with coerceStringArray's lenient coercion): an accidental
    // leading/trailing space would otherwise reach the filesystem and read as "file not found".
    const spec: AttachmentSpec = { path: obj.path.trim() };
    if (obj.name !== undefined) {
      if (typeof obj.name !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].name must be a string.`);
      spec.name = obj.name;
    }
    if (obj.contentType !== undefined) {
      if (typeof obj.contentType !== 'string') throw new McpError(ErrorCode.InvalidParams, `attachments[${i}].contentType must be a string.`);
      spec.contentType = obj.contentType;
    }
    specs.push(spec);
  }
  return specs;
}
