// Shared helpers for the live probes: PASS/FAIL harness and MCP response parsing.

// One PASS/FAIL line per check; failures() feeds the exit code.
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

// Extract the first complete JSON value from a response that may wrap it in
// prose (list tools emit a summary line before and note lines after the JSON).
export const jsonOf = t => {
  const s = Math.min(...['[', '{'].map(ch => { const i = t.indexOf(ch); return i < 0 ? Infinity : i; }));
  let d = 0, inStr = false, esc = false;
  for (let i = s; i < t.length; i++) {
    const ch = t[i];
    if (esc) { esc = false; continue; }
    if (ch === '\\') { esc = true; continue; }
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === '[' || ch === '{') d++;
    else if (ch === ']' || ch === '}') { d--; if (d === 0) return JSON.parse(t.slice(s, i + 1)); }
  }
  throw new Error('no JSON found in: ' + t.slice(0, 120));
};

// Draft id from create/edit responses ("Email ID: X" / "New Email ID: X").
export const idOf = t => t.match(/[Ee]mail ID: ([A-Za-z0-9_-]+)/)?.[1];

// Raw html + non-html-derived text of an email, joined over bodyValues.
export const rawBodies = async (c, id) => {
  const r = await c.call('get_email', { emailId: id, raw: true });
  const raw = jsonOf(text(r));
  const hIds = (raw.htmlBody ?? []).map(p => p.partId);
  return {
    html: hIds.map(pid => raw.bodyValues?.[pid]?.value ?? '').join(''),
    txt: (raw.textBody ?? []).map(p => p.partId).filter(pid => !hIds.includes(pid)).map(pid => raw.bodyValues?.[pid]?.value ?? '').join(''),
  };
};
