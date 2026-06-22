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
export function requireNonEmpty(value: unknown, fieldName: string): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${fieldName} cannot be empty; omit the field to leave it unchanged`);
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
