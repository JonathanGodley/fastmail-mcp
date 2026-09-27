// Shared helpers for the live probes: PASS/FAIL harness and MCP response parsing.

// One PASS/FAIL line per check; failures() feeds the exit code.
//
// ARGUMENT ORDER: check(label, ok, extra) - the LABEL comes first. probe-exact-instance.mjs
// carries its own `check(ok, label)` with the first two inverted, so moving a line between it
// and this harness silently turns every assertion into a check of a truthy string. Convert
// call sites to this order; do not adapt this signature, which every other probe imports.
export function makeChecker() {
  let fails = 0;
  const check = (label, ok, extra = '') => {
    console.log(`${ok ? 'PASS' : 'FAIL'} ${label}${extra ? ' — ' + extra : ''}`);
    if (!ok) fails++;
  };
  return { check, failures: () => fails };
}

// First content item's text from an MCP tool result.
export const text = r => r?.content?.[0]?.text ?? '';

// Extract the first complete JSON value from a response that may wrap it in prose (list
// tools emit a summary line before the payload and notes after it).
//
// Walks bracket depth with string state, so it ends the payload at ITS OWN closing bracket:
// slicing to `lastIndexOf(']')` would take a bracket in trailing prose, and a bracket inside
// a string VALUE would end the slice early. Candidate openers are tried in order because
// prose before the payload can contain a bracket too; one that does not parse means the
// payload starts later.
export const jsonOf = t => {
  for (let from = 0; from < t.length; from++) {
    const s = Math.min(...['[', '{'].map(ch => { const i = t.indexOf(ch, from); return i < 0 ? Infinity : i; }));
    if (!Number.isFinite(s)) break;
    from = s;
    let d = 0, inStr = false, esc = false;
    for (let i = s; i < t.length; i++) {
      const ch = t[i];
      if (esc) { esc = false; continue; }
      if (inStr && ch === '\\') { esc = true; continue; }
      if (ch === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (ch === '[' || ch === '{') d++;
      else if (ch === ']' || ch === '}') {
        if (--d === 0) {
          try { return JSON.parse(t.slice(s, i + 1)); } catch { /* prose, not a payload: try the next opener */ }
          break;
        }
      }
    }
  }
  throw new Error('no JSON found in: ' + t.slice(0, 120));
};

// Draft id from create/edit responses ("Email ID: X" / "New Email ID: X").
export const idOf = t => t.match(/[Ee]mail ID: ([A-Za-z0-9_-]+)/)?.[1];

// Raw html + non-html-derived text of an email, joined over bodyValues.
export const rawBodies = async (c, id) => {
  const r = await c.call('get_email', { emailId: id, raw: true });
  const raw = jsonOf(text(r));
  // Classify by the part's own media type, NOT by which list it appears in: with only one
  // alternative, JMAP lists that single part in BOTH htmlBody and textBody, so reading
  // htmlBody as "the html" reports a plain-text draft as html-with-no-text.
  const partsOfType = (parts, wanted) => (parts ?? [])
    .filter(p => (p.type || '').toLowerCase() === wanted)
    .map(p => raw.bodyValues?.[p.partId]?.value ?? '')
    .join('');
  return {
    html: partsOfType(raw.htmlBody, 'text/html'),
    txt: partsOfType(raw.textBody, 'text/plain'),
  };
};
