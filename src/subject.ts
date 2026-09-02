import { ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { isBlank } from './body-format.js';

// The `subject` override shared by the two draft_email modes that derive a subject from an
// original message: mode:'reply' ("Re: <original>") and mode:'forward' ("Fwd: <original>").
// One implementation so the two can't drift — as separate tools they took the same
// parameter with different postures (#68).
//
// Lives in its own module rather than in coerce.ts because the blank test must be the
// shared, zero-width-aware isBlank predicate: coerce.ts imports no local module today, and
// reaching for isBlank from there would make body-format.ts and coerce.ts import each
// other. The compose handler already imports this kind of small focused helper.

// Validate the optional `subject` override, returning it, or undefined meaning "fall back
// to the tool's derived default". omittedHint names that default in the reject message, so
// the error tells the caller what omitting the parameter would have given them.
//
// A present-but-non-string value is REJECTED rather than ignored, matching the body
// params: a subject that arrives as a number or an array is a caller bug, and silently
// replacing it with the derived default would ship mail under a line the caller never
// asked for. `null` is accepted as "omitted" — that is how several lenient clients spell
// an unset optional field.
//
// A present-but-BLANK string falls through to the derived default instead of erroring. A
// blank string cannot express which of two things the caller wants (a genuinely empty
// subject line, or the default), and the default is the non-destructive read of the two;
// the resulting subject is echoed back in the tool's success message, so the outcome is
// visible either way. Blankness is zero-width-aware, so a subject that renders as empty
// is treated as empty.
export function coerceSubjectOverride(value: unknown, omittedHint: string): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string') {
    const got = Array.isArray(value) ? 'array' : typeof value;
    throw new McpError(ErrorCode.InvalidParams, `subject must be a string; received ${got}. ${omittedHint}`);
  }
  return isBlank(value) ? undefined : value;
}
