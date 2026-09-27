import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { isBlank } from './body-format.js';

// The `subject` override shared by the two draft_email modes that derive a subject from an
// original message: mode:'reply' ("Re: <original>") and mode:'forward' ("Fwd: <original>").
// One implementation so the two can't drift (#68).
//
// Not in coerce.ts because the blank test must be the zero-width-aware isBlank, and
// body-format.ts already imports coerce.ts, so the two would import each other.

// Validate the optional `subject` override, returning it, or undefined meaning "fall back
// to the tool's derived default". omittedHint names that default in the reject message.
//
// A non-string value is REJECTED rather than ignored: silently replacing it with the
// derived default would ship mail under a line the caller never asked for. `null` counts as
// omitted, which is how several lenient clients spell an unset optional field.
//
// A blank string falls through to the derived default instead of erroring: it cannot say
// whether the caller wants an empty subject or the default, the default is the
// non-destructive read, and the success message echoes the subject either way.
export function coerceSubjectOverride(value: unknown, omittedHint: string): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    const got = Array.isArray(value) ? 'array' : typeof value;
    throw new McpError(ErrorCode.InvalidParams, `subject must be a string; received ${got}. ${omittedHint}`);
  }
  return isBlank(value) ? undefined : value;
}
