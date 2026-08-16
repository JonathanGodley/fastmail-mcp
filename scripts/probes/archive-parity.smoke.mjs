// Archive semantics against a real account: that the patch form Cyrus actually
// receives is accepted, that a message filed in the Inbox plus a label comes out
// still holding the label, that an Inbox-only message reaches Archive, that a
// message already out of the Inbox is untouched, and that a refusing role writes
// nothing. Also sends one mixed multi-id batch, since "both patch shapes in one
// Email/set" is a server accept/reject claim no mock can settle.
//
// This proves the CODE matches the measured table. It does NOT prove the table
// matches Fastmail — that rests on the UI measurements in
// docs/fastmail-action-availability.md.
//
// Last run 2026-08-16: all checks passed, which is what settled the two claims no mock can
// reach — that Cyrus accepts one Email/set carrying both patch shapes, and that an
// Inbox+Trash message survives the re-assert with Trash kept and Archive not added. That was
// one run against one account, so it is evidence the code was right THEN, not standing
// coverage; re-run it after any change to the branch rule or the patch shape.
//
// Self-fixturing: creates its own messages and a label folder and clears both on exit,
// including on a failed check and on the setup path that bails before the try block the
// finally hangs off. The label folder is DESTROYED (it is a real mailbox, so it needs
// Mailbox/set destroy rather than a Trash move); the fixture messages are TRASHED, because
// teardown goes through delete_email, which moves to Trash by design. Expect seven probe
// messages sitting in Trash after a run. Token via run-probe.py child env; never printed.
import { createClient } from '../mcp-harness.mjs';
import { makeChecker, text, jsonOf } from './probelib.mjs';
import { getSession, jmap } from './jmaplib.mjs';

const LABEL_NAME = `archive-probe-${Date.now()}`;
const session = await getSession();

const boxes = await jmap(session, [
  ['Mailbox/get', { accountId: session.accountId, ids: null, properties: ['id', 'name', 'role'] }, 'm'],
]);
const all = boxes.methodResponses.find(r => r[2] === 'm')[1].list;
const byRole = role => all.find(x => x.role === role);
const inbox = byRole('inbox');
const archive = byRole('archive');
// Sent, not Trash: Trash alone is the least informative of the six refusing roles,
// because a reader can talk themselves into "of course you cannot archive deleted mail".
const refusing = byRole('sent') ?? byRole('drafts');
const trash = byRole('trash');
if (!inbox || !archive || !refusing || !trash) {
  console.error('FIXTURE SETUP FAILED: this account is missing inbox, archive, trash or a refusing role');
  process.exit(1);
}

const madeLabel = await jmap(session, [
  ['Mailbox/set', { accountId: session.accountId, create: { L: { name: LABEL_NAME } } }, 'l'],
]);
const LABEL = madeLabel.methodResponses.find(r => r[2] === 'l')[1].created?.L?.id;
if (!LABEL) { console.error('FIXTURE SETUP FAILED: could not create the label folder'); process.exit(1); }

const fixture = (key, subject, mailboxIds) => [key, {
  mailboxIds: Object.fromEntries(mailboxIds.map(id => [id, true])),
  keywords: { $seen: true },
  from: [{ name: 'Archive Probe', email: 'probe@invalid.example' }],
  subject,
  bodyStructure: { type: 'text/plain', partId: 'b' },
  bodyValues: { b: { value: `${subject}\n` } },
}];

const create = Object.fromEntries([
  fixture('both', 'archive probe: inbox + label', [inbox.id, LABEL]),
  fixture('only', 'archive probe: inbox only', [inbox.id]),
  fixture('gone', 'archive probe: label only', [LABEL]),
  fixture('deny', 'archive probe: refusing role', [refusing.id]),
  // Inbox+Trash: the row the availability doc records as OUR decision rather than parity,
  // because the client cannot produce the state. It is also the one a later tidy-up would
  // most plausibly "fix" into a refusal, and the only place a live run can settle whether
  // the server accepts patching Inbox away while re-asserting Trash.
  fixture('both2', 'archive probe: inbox + trash', [inbox.id, trash.id]),
  fixture('mixA', 'archive probe: batch inbox only', [inbox.id]),
  fixture('mixB', 'archive probe: batch inbox + label', [inbox.id, LABEL]),
]);
const made = await jmap(session, [['Email/set', { accountId: session.accountId, create }, 'c']]);
const created = made.methodResponses.find(r => r[2] === 'c')[1].created ?? {};
const ID = Object.fromEntries(Object.keys(create).map(k => [k, created[k]?.id]));
const missing = Object.entries(ID).filter(([, v]) => !v).map(([k]) => k);
if (missing.length) {
  // Tear down before bailing. By this point the label folder exists and a PARTIAL create may
  // have left real messages in the account, so exiting straight out would leak both into the
  // live mailbox — and the finally block below never runs, because it guards a try this exit
  // is upstream of. Raw JMAP rather than the MCP client, which is not started yet.
  const strays = Object.values(ID).filter(Boolean);
  try {
    await jmap(session, [
      ...(strays.length ? [['Email/set', { accountId: session.accountId, destroy: strays }, 'x']] : []),
      ['Mailbox/set', { accountId: session.accountId, destroy: [LABEL], onDestroyRemoveEmails: true }, 'd'],
    ]);
    console.log('partial fixtures cleaned up');
  } catch (e) { console.log('CLEANUP FAILED after a partial create:', String(e).slice(0, 120)); }
  console.error('FIXTURE CREATE FAILED for', missing.join(', '));
  process.exit(1);
}

const { check, failures } = makeChecker();
// Declared before the try so the finally can see it, but STARTED inside it: createClient and
// init can fail on their own (a stale dist/, a token missing from the child env), and with
// them outside the try that failure exits with all seven fixture messages and the label
// folder left in the live account.
let c;

// The message's real membership, read back over raw JMAP rather than trusted from
// the tool's own projected report — the projection is exactly what is under test.
const filedIn = async id => {
  const got = await jmap(session, [
    ['Email/get', { accountId: session.accountId, ids: [id], properties: ['mailboxIds'] }, 'g'],
  ]);
  return Object.keys(got.methodResponses.find(r => r[2] === 'g')[1].list?.[0]?.mailboxIds ?? {});
};
// The second content item is `{ counts, results }`, not a bare results array. Reading it as
// an array does not degrade into an empty list — `?? []` never fires on a truthy object, so
// `.find` throws and the very first check takes the whole probe down.
const entryFor = (r, id) => (jsonOf(r.content[1].text)?.results ?? []).find(e => e.id === id);

try {
  c = createClient({ env: { ...process.env } });
  await c.init();

  // 1. Inbox + label: the label survives and Archive is NOT added.
  let r = await c.call('archive_email', { emailIds: [ID.both] });
  check('inbox+label reports removedFromInbox', entryFor(r, ID.both)?.action === 'removedFromInbox', text(r).slice(0, 160));
  let now = await filedIn(ID.both);
  check('inbox+label keeps the label', now.includes(LABEL), now.join(','));
  check('inbox+label left the Inbox', !now.includes(inbox.id), now.join(','));
  check('inbox+label did NOT gain Archive', !now.includes(archive.id), now.join(','));

  // 2. Inbox only: reaches Archive, and is filed nowhere else.
  r = await c.call('archive_email', { emailIds: [ID.only] });
  check('inbox-only reports movedToArchive', entryFor(r, ID.only)?.action === 'movedToArchive', text(r).slice(0, 160));
  now = await filedIn(ID.only);
  check('inbox-only is in Archive alone', now.length === 1 && now[0] === archive.id, now.join(','));

  // 3. Already out of the Inbox: untouched, and reported as a success.
  r = await c.call('archive_email', { emailIds: [ID.gone] });
  check('label-only reports notInInbox', entryFor(r, ID.gone)?.action === 'notInInbox', text(r).slice(0, 160));
  check('label-only call is not an error', !r.isError);
  now = await filedIn(ID.gone);
  check('label-only membership unchanged', now.length === 1 && now[0] === LABEL, now.join(','));

  // 4. A refusing role writes nothing at all.
  r = await c.call('archive_email', { emailIds: [ID.deny] });
  const denied = entryFor(r, ID.deny);
  check(`${refusing.role} is refused`, denied?.action === 'refused', JSON.stringify(denied ?? null).slice(0, 160));
  check(`${refusing.role} refusal names the role`, denied?.reason?.role === refusing.role, JSON.stringify(denied?.reason ?? null));
  now = await filedIn(ID.deny);
  check(`${refusing.role} membership unchanged`, now.length === 1 && now[0] === refusing.id, now.join(','));

  // 4b. Inbox+Trash takes the WRITING branch: Inbox goes, Trash stays, Archive is not added.
  r = await c.call('archive_email', { emailIds: [ID.both2] });
  check('inbox+trash is archived, not refused', entryFor(r, ID.both2)?.action === 'removedFromInbox', text(r).slice(0, 160));
  now = await filedIn(ID.both2);
  check('inbox+trash keeps Trash and nothing else', now.length === 1 && now[0] === trash.id, now.join(','));

  // 5. A mixed batch: both patch shapes in one Email/set, which the server must accept.
  r = await c.call('archive_email', { emailIds: [ID.mixA, ID.mixB] });
  check('mixed batch reports both branches', entryFor(r, ID.mixA)?.action === 'movedToArchive' && entryFor(r, ID.mixB)?.action === 'removedFromInbox', text(r).slice(0, 200));
  const a = await filedIn(ID.mixA);
  const b = await filedIn(ID.mixB);
  check('mixed batch: inbox-only half reached Archive', a.length === 1 && a[0] === archive.id, a.join(','));
  check('mixed batch: labelled half kept its label and nothing else', b.length === 1 && b[0] === LABEL, b.join(','));

  // 6. Idempotent: archiving again changes nothing and still succeeds.
  r = await c.call('archive_email', { emailIds: [ID.only] });
  check('re-archiving is a reported no-op', entryFor(r, ID.only)?.action === 'notInInbox', text(r).slice(0, 160));
} finally {
  // Guarded on `c`: if init threw, the harness never came up and the message teardown has no
  // client to go through. The label destroy below runs over raw JMAP either way, and
  // onDestroyRemoveEmails takes the fixtures filed in it with it.
  for (const id of c ? Object.values(ID) : []) {
    try { await c.call('delete_email', { emailId: id }); } catch (e) { console.log('CLEANUP FAILED for', id, String(e).slice(0, 100)); }
  }
  if (c) console.log('message fixtures trashed');
  try {
    await jmap(session, [['Mailbox/set', { accountId: session.accountId, destroy: [LABEL], onDestroyRemoveEmails: true }, 'd']]);
    console.log('label folder destroyed');
  } catch (e) { console.log('CLEANUP FAILED for label', LABEL, String(e).slice(0, 100)); }
  if (c) await c.close();
}

console.log(failures() === 0 ? '\nARCHIVE-PARITY: ALL PASS' : `\nARCHIVE-PARITY: ${failures()} FAILURE(S)`);
process.exit(failures() === 0 ? 0 : 1);
