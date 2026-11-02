// Can a first-party web session snooze, re-snooze and unsnooze a message over JMAP,
// and what does Fastmail refuse? Companion to snooze-read.probe.mjs (#35): that one
// found `snoozed` served under `https://www.fastmail.com/dev/mail`; this one writes it.
//
// Raw JMAP on a synthetic fixture created in Inbox with Email/set create (never sent)
// and destroyed in the finally block, including on failure. Nothing of the account's
// own mail is touched. Needs FASTMAIL_WEB_TOKEN and FASTMAIL_WEB_COOKIE (see README).
//
// The write shape under test is Cyrus's: snoozing is one Email/set update that drops
// the Inbox membership, adds the Snoozed membership and sets `snoozed` to
// `{ until, moveToMailboxId?, setKeywords? }`; waking is the reverse with `snoozed: null`.

const WEB = process.env.FASTMAIL_WEB_TOKEN;
if (!WEB) { console.error('FASTMAIL_WEB_TOKEN is not set; an API token cannot write snoozed. Stopping.'); process.exit(2); }
process.env.FASTMAIL_API_TOKEN = WEB;
if (process.env.FASTMAIL_WEB_COOKIE) process.env.JMAPLIB_COOKIE = process.env.FASTMAIL_WEB_COOKIE;
const { getSession, jmap } = await import('./jmaplib.mjs');
import { makeChecker } from './probelib.mjs';

const FM_DEV_MAIL = 'https://www.fastmail.com/dev/mail';
const TAG = 'zz-snooze-write-probe';
const { check, failures } = makeChecker();

function res(body, i) {
  const entry = body.methodResponses[i];
  if (entry[0] === 'error') throw new Error(`method ${i} errored: ${JSON.stringify(entry[1])}`);
  return entry[1];
}
const iso = d => d.toISOString().replace(/\.\d{3}Z$/, 'Z');

const s = await getSession();
const mailboxes = res(await jmap(s, [['Mailbox/get', { accountId: s.accountId, properties: ['id', 'name', 'role'] }, 'm']]), 0).list;
const inbox = mailboxes.find(m => m.role === 'inbox');
const snoozedMb = mailboxes.find(m => m.role === 'snoozed');
if (!inbox || !snoozedMb) { console.error('need both an Inbox and a role=snoozed mailbox'); process.exit(1); }

const set = (update) => jmap(s, [['Email/set', { accountId: s.accountId, update }, 'u']], [FM_DEV_MAIL]).then(b => res(b, 0));
const get = (id) => jmap(s, [['Email/get', { accountId: s.accountId, ids: [id], properties: ['id', 'mailboxIds', 'keywords', 'snoozed'] }, 'g']], [FM_DEV_MAIL]).then(b => res(b, 0).list[0]);
const inOnly = (e, mb) => Object.keys(e.mailboxIds).length === 1 && e.mailboxIds[mb.id] === true;

let id = null;
try {
  const c = res(await jmap(s, [['Email/set', {
    accountId: s.accountId,
    create: { f: {
      mailboxIds: { [inbox.id]: true },
      from: [{ email: 'probe@example.invalid', name: 'Snooze Probe' }],
      to: [{ email: 'probe@example.invalid' }],
      subject: `${TAG} fixture`,
      keywords: { $seen: true },
      bodyStructure: { type: 'text/plain', partId: 'p' },
      bodyValues: { p: { value: 'Local fixture for the snooze-write probe. Never sent.' } },
    } },
  }, 'c']]), 0);
  if (!c.created?.f) throw new Error(`fixture create failed: ${JSON.stringify(c.notCreated)}`);
  id = c.created.f.id;
  console.log('fixture created in Inbox');

  // 1. Refusal: a snoozed object while the message stays in Inbox.
  const bad = await set({ [id]: { snoozed: { until: iso(new Date(Date.now() + 86400e3)) } } });
  check('snoozed set without the Snoozed membership is refused', !!bad.notUpdated?.[id], JSON.stringify(bad.notUpdated?.[id] ?? bad.updated));

  // 2. Snooze: the client's own shape (wake into Inbox, unread and $new).
  const until1 = iso(new Date(Date.now() + 2 * 86400e3));
  const sn = await set({ [id]: {
    [`mailboxIds/${inbox.id}`]: null,
    [`mailboxIds/${snoozedMb.id}`]: true,
    snoozed: { until: until1, setKeywords: { $new: true, $seen: false } },
  } });
  check('snooze (patch memberships + snoozed object) is accepted', !!sn.updated && id in sn.updated, JSON.stringify(sn.notUpdated?.[id] ?? ''));
  id = sn.updated?.[id]?.id ?? id;
  let e = await get(id);
  check('snoozed message is in Snoozed only', inOnly(e, snoozedMb), JSON.stringify(e.mailboxIds));
  check('snoozed.until reads back as written', e.snoozed?.until === until1, JSON.stringify(e.snoozed));
  const kw = e.snoozed?.setKeywords ?? {};
  check('snoozed.setKeywords reads back as written', kw.$new === true && kw.$seen === false && Object.keys(kw).length === 2, JSON.stringify(kw));
  check('snoozed.moveToMailboxId is absent when not given', !('moveToMailboxId' in (e.snoozed ?? {})) || e.snoozed.moveToMailboxId == null, JSON.stringify(e.snoozed?.moveToMailboxId));

  // 3. Re-snooze: change the wake time of an already-snoozed message. First as a patch
  // of the one field, then, if refused, as the whole object.
  const until2 = iso(new Date(Date.now() + 3 * 86400e3));
  const p = await set({ [id]: { 'snoozed/until': until2 } });
  const patched = !!p.updated && id in p.updated;
  check('re-snooze by patching snoozed/until is accepted', patched, JSON.stringify(p.notUpdated?.[id] ?? ''));
  if (!patched) {
    const w = await set({ [id]: { snoozed: { until: until2, setKeywords: { $new: true, $seen: false } } } });
    check('re-snooze by writing the whole snoozed object is accepted', !!w.updated && id in w.updated, JSON.stringify(w.notUpdated?.[id] ?? ''));
    id = w.updated?.[id]?.id ?? id;
  }
  e = await get(id);
  check('new wake time reads back', e.snoozed?.until === until2, JSON.stringify(e.snoozed));
  check('still in Snoozed only after re-snooze', inOnly(e, snoozedMb), JSON.stringify(e.mailboxIds));

  // 4. Leaving Snoozed without clearing snoozed. Upstream Cyrus refuses this with
  // invalidProperties; Fastmail accepts it, so the read-back records what a plain move
  // out of Snoozed (this server's move_email) leaves behind.
  const mv = await set({ [id]: { [`mailboxIds/${snoozedMb.id}`]: null, [`mailboxIds/${inbox.id}`]: true } });
  const moved = !!mv.updated && id in mv.updated;
  check('leaving Snoozed with snoozed still set (informational: accepted or refused)', true, moved ? 'ACCEPTED' : `refused ${JSON.stringify(mv.notUpdated?.[id])}`);
  if (moved) {
    id = mv.updated[id]?.id ?? id;
    e = await get(id);
    check('after the plain move out of Snoozed (informational)', true, `mailboxes ${JSON.stringify(Object.keys(e.mailboxIds).map(m => mailboxes.find(x => x.id === m)?.role ?? m))} snoozed ${JSON.stringify(e.snoozed)}`);
  }

  // 5. Unsnooze: back to Inbox with snoozed null.
  e = await get(id);
  const un = await set({ [id]: {
    ...(e.mailboxIds[snoozedMb.id] ? { [`mailboxIds/${snoozedMb.id}`]: null } : {}),
    [`mailboxIds/${inbox.id}`]: true,
    snoozed: null,
  } });
  check('unsnooze (back to Inbox, snoozed: null) is accepted', !!un.updated && id in un.updated, JSON.stringify(un.notUpdated?.[id] ?? ''));
  id = un.updated?.[id]?.id ?? id;
  e = await get(id);
  check('woken message is in Inbox only', inOnly(e, inbox), JSON.stringify(e.mailboxIds));
  check('woken message reads snoozed: null', e.snoozed === null, JSON.stringify(e.snoozed));
  console.log('keywords after manual unsnooze:', JSON.stringify(e.keywords), '(setKeywords applies only when the server wakes it)');
} finally {
  if (id) {
    const d = res(await jmap(s, [['Email/set', { accountId: s.accountId, destroy: [id] }, 'd']]), 0);
    console.log(`cleanup: destroyed ${d.destroyed?.length ?? 0} fixture` + (d.notDestroyed && Object.keys(d.notDestroyed).length ? ` NOT destroyed: ${JSON.stringify(d.notDestroyed)}` : ''));
  }
}
process.exit(failures() ? 1 : 0);
