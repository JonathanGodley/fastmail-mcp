// Does a membership patch that would leave a message filed nowhere get rejected,
// or does it expunge the message?
//
// This is the platform fact #132 turns on. `removeLabels`/`bulkRemoveLabels`
// used to emit a bare `{"mailboxIds/<id>": null}` patch with no read of current
// membership and no emptiness guard of their own, so whatever the server does
// here was exactly what remove_labels did. They now read the current filing and
// re-assert what survives, adding Archive when nothing would; this probe pins the
// platform behaviour that made that necessary and would start failing if the
// server's own guard were ever tightened to cover both branches.
//
// The answer is not one behaviour but two, selected by history the caller cannot
// see. The server's patch-path emptiness guard counts conversations-record
// tombstones rather than live memberships, so a message that has ever been moved
// out of any mailbox carries a `{removed}` tombstone that inflates the count past
// the guard. Two fixtures pin both branches:
//
//   A. clean      — created in one mailbox, never moved. Guard sees no other
//                   record, rejects the patch, message survives.
//   B. tombstoned — created in mailbox 2, moved to mailbox 1. The mailbox-2
//                   tombstone survives the move, the guard is satisfied, the
//                   executor then filters tombstones out and expunges the only
//                   live record. The message is DESTROYED, not moved to Trash.
//
// Fixture B is the common case, not the exotic one: archive_email patches the
// Inbox membership away, so every message it has touched is in that class.
//
// This probe deliberately talks raw JMAP rather than going through the built
// server: the unknown is server behaviour, and routing it through the tool would
// only measure the tool's own guard instead.
//
// Fixtures are created directly with Email/set create. Nothing is ever sent.
// Everything is destroyed in the finally block, including on failure.

import { getSession, jmap } from './jmaplib.mjs';
import { makeChecker } from './probelib.mjs';

const TAG = 'zz-label-emptiness-probe';
const { check, failures } = makeChecker();

function res(body, i) {
  const entry = body.methodResponses[i];
  if (entry[0] === 'error') throw new Error(`method ${i} errored: ${JSON.stringify(entry[1])}`);
  return entry[1];
}

function fixture(mboxId, subject) {
  return {
    mailboxIds: { [mboxId]: true },
    from: [{ email: 'probe@example.invalid', name: 'Emptiness Probe' }],
    to: [{ email: 'probe@example.invalid' }],
    subject,
    keywords: { $seen: true },
    bodyStructure: { type: 'text/plain', partId: 'p' },
    bodyValues: { p: { value: 'Local fixture for the label-emptiness probe. Never sent.' } },
  };
}

const s = await getSession();
const created = { mailboxes: [], emails: [] };

try {
  // --- Two throwaway mailboxes --------------------------------------------
  const mb = res(await jmap(s, [['Mailbox/set', {
    accountId: s.accountId,
    create: {
      one: { name: `${TAG}-1`, parentId: null },
      two: { name: `${TAG}-2`, parentId: null },
    },
  }, 'c']]), 0);
  if (mb.notCreated && Object.keys(mb.notCreated).length) {
    throw new Error(`mailbox create failed: ${JSON.stringify(mb.notCreated)}`);
  }
  const box1 = mb.created.one.id;
  const box2 = mb.created.two.id;
  created.mailboxes.push(box1, box2);

  // --- Fixture A: clean, never moved --------------------------------------
  const ca = res(await jmap(s, [['Email/set', {
    accountId: s.accountId,
    create: { a: fixture(box1, `${TAG} fixture A (clean)`) },
  }, 'c']]), 0);
  if (!ca.created?.a) throw new Error(`fixture A create failed: ${JSON.stringify(ca.notCreated)}`);
  const idA = ca.created.a.id;
  created.emails.push(idA);

  // --- Fixture B: created in box2, whole-value moved to box1 --------------
  const cb = res(await jmap(s, [['Email/set', {
    accountId: s.accountId,
    create: { b: fixture(box2, `${TAG} fixture B (tombstoned)`) },
  }, 'c']]), 0);
  if (!cb.created?.b) throw new Error(`fixture B create failed: ${JSON.stringify(cb.notCreated)}`);
  let idB = cb.created.b.id;
  const mv = res(await jmap(s, [['Email/set', {
    accountId: s.accountId,
    update: { [idB]: { mailboxIds: { [box1]: true } } },
  }, 'm']]), 0);
  if (mv.notUpdated?.[idB]) throw new Error(`fixture B move failed: ${JSON.stringify(mv.notUpdated)}`);
  // A whole-value move can re-issue the id; follow it if so.
  idB = mv.updated?.[idB]?.id ?? idB;
  created.emails.push(idB);

  // --- Empty each one's last remaining membership -------------------------
  const observe = async (id) => {
    const out = res(await jmap(s, [['Email/set', {
      accountId: s.accountId,
      update: { [id]: { [`mailboxIds/${box1}`]: null } },
    }, 'rm']]), 0);
    const got = res(await jmap(s, [['Email/get', {
      accountId: s.accountId, ids: [id], properties: ['id', 'mailboxIds'],
    }, 'g']]), 0);
    return {
      accepted: Object.prototype.hasOwnProperty.call(out.updated ?? {}, id),
      setError: out.notUpdated?.[id],
      alive: (got.list?.length ?? 0) > 0,
    };
  };

  const a = await observe(idA);
  check('clean message: emptying patch is rejected', a.accepted === false,
    `accepted=${a.accepted}`);
  check('clean message: rejection is invalidProperties on mailboxIds',
    a.setError?.type === 'invalidProperties' && a.setError?.properties?.includes('mailboxIds'),
    JSON.stringify(a.setError));
  check('clean message: survives', a.alive === true);

  const b = await observe(idB);
  if (!b.alive) created.emails = created.emails.filter(e => e !== idB);
  check('tombstoned message: emptying patch is ACCEPTED', b.accepted === true,
    `accepted=${b.accepted} setError=${JSON.stringify(b.setError)}`);
  check('tombstoned message: message is DESTROYED, not left filed nowhere',
    b.alive === false, `alive=${b.alive}`);

  // The whole point: identical call, opposite outcome, and the caller cannot
  // tell the two messages apart from anything the API exposes.
  check('the two branches genuinely disagree', a.alive === true && b.alive === false);
} finally {
  // Each step is isolated. res() throws on a method-level error entry, so an unguarded
  // failure in the first step would propagate out of the finally block, skip the mailbox
  // destroy entirely — leaving the probe's folders and their contents in the live account —
  // and replace the real failure with the cleanup's. Both steps always run; both report.
  const sweep = async (label, call) => {
    try {
      const d = res(await jmap(s, [call]), 0);
      console.log(`cleanup: destroyed ${d.destroyed?.length ?? 0} ${label}` +
        (d.notDestroyed && Object.keys(d.notDestroyed).length
          ? ` NOT destroyed: ${JSON.stringify(d.notDestroyed)}` : ''));
    } catch (err) {
      console.log(`cleanup FAILED for ${label}: ${err?.message ?? err}. ` +
        'Remove them by hand: they are named with the zz-label-emptiness-probe prefix.');
    }
  };

  if (created.emails.length) {
    await sweep('fixture email(s)', ['Email/set', {
      accountId: s.accountId, destroy: created.emails,
    }, 'd']);
  }
  if (created.mailboxes.length) {
    await sweep('probe mailbox(es)', ['Mailbox/set', {
      accountId: s.accountId, destroy: created.mailboxes, onDestroyRemoveEmails: true,
    }, 'd']);
  }
}

process.exit(failures() ? 1 : 0);
