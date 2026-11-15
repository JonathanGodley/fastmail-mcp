// Does Fastmail's web session serve the full list property set alongside `snoozed`, so that
// list_snoozed_emails can return the same simplified shape as list_emails (#230)? The unit
// suite answers through a fake transport; only a real session shows whether Email/get under
// https://www.fastmail.com/dev/mail accepts those properties.
//
// The built server, READ-ONLY: list_snoozed_emails twice (simplified, then raw). Needs
// FASTMAIL_WEB_COOKIE. Output: counts and field NAMES only, never a value from the mailbox.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf } from './probelib.mjs';

const { check, failures } = makeChecker();
if (!process.env.FASTMAIL_WEB_COOKIE) { console.error('FASTMAIL_WEB_COOKIE is not set. Stopping.'); process.exit(2); }

const c = createClient({ env: { ...process.env, FASTMAIL_EXPERIMENTAL_SNOOZE: '1' } });
await c.init();
// A failed call check ends the run: the checks after it would read a result that is not there.
const listCall = async (label, args) => {
  const r = await c.call('list_snoozed_emails', args);
  check(`${label}: the call succeeds`, !r.isError, r.isError ? text(r).slice(0, 200) : '');
  return r.isError ? null : jsonOf(text(r));
};
try {
  const simplified = await listCall('simplified', { limit: 5 });
  if (simplified) console.log(`snoozed messages: ${simplified.total ?? 'not returned'}, returned ${simplified.emails.length}`);
  const raw = simplified && await listCall('raw', { limit: 5, raw: true });
  const sameCount = raw && raw.emails.length === simplified.emails.length;
  if (raw) check('raw: same messages as the simplified call', sameCount);
  if (!sameCount) {
    // nothing further can be compared; the summary below reports the failure
  } else if (simplified.emails.length === 0) {
    console.log('NOTE: nothing is snoozed, so the item shape was not exercised. Snooze a message and re-run.');
  } else {
    const item = simplified.emails[0];
    const keys = Object.keys(item);
    console.log('simplified keys:', keys.join(', '));
    check('simplified: carries wakeAt and wakeInto', 'wakeAt' in item && 'wakeInto' in item);
    check('simplified: date is local time with an offset', /[+-]\d\d:\d\d$/.test(item.date ?? ''));
    check('simplified: roles include snoozed', (item.roles ?? []).includes('snoozed'));
    check('simplified: from is a "Name <email>" or bare address string', typeof item.from === 'string');
    check('simplified: no raw snoozed object leaks through', !('snoozed' in item));
    const rawKeys = Object.keys(raw.emails[0]);
    console.log('raw keys:', rawKeys.join(', '));
    check('raw: carries the snoozed object and mailboxIds', rawKeys.includes('snoozed') && rawKeys.includes('mailboxIds'));
    check('raw: carries the list property set (keywords, preview, receivedAt)',
      ['keywords', 'preview', 'receivedAt'].every(k => rawKeys.includes(k)));
  }
} finally {
  c.close();
}
console.log(failures() ? `SNOOZE-LIST: ${failures()} FAILED` : 'SNOOZE-LIST: ALL PASS');
process.exit(failures() ? 1 : 0);
