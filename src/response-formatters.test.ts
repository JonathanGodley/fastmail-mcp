import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { simplifyMailbox, simplifyIdentity, simplifyContact, formatEmailQueryResult, formatContactQueryResult, formatEditDraftResult } from './response-formatters.js';

// ---------- formatEditDraftResult ----------

describe('formatEditDraftResult', () => {
  const REPLACED = {
    id: 'draft-1',
    subject: 'Lunch plans',
    to: ['bob@example.com'],
    cc: ['carol@example.com'],
    textBodySize: 120,
    htmlBodySize: 240,
  };

  it('reports the new id, the Trash move, and what was replaced', () => {
    const text = formatEditDraftResult({ id: 'draft-2', replacedDraft: REPLACED, trashedOldDraftId: 'draft-1' });
    assert.match(text, /New Email ID: draft-2/);
    assert.match(text, /draft-1\) was moved to Trash/);
    assert.match(text, /recoverable until Trash is emptied or auto-purged/);
    assert.match(text, /subject "Lunch plans"/);
    assert.match(text, /to bob@example\.com/);
    assert.match(text, /cc carol@example\.com/);
    assert.match(text, /htmlBody 240 chars/);
    assert.match(text, /textBody 120 chars/);
    assert.doesNotMatch(text, /WARNING/);
  });

  it('warns plainly when the old draft could not be trashed, and gives the reason', () => {
    const text = formatEditDraftResult({
      id: 'draft-2',
      replacedDraft: REPLACED,
      orphanedOldDraftId: 'draft-1',
      orphanedOldDraftReason: 'this account has no mailbox with the trash role',
    });
    assert.match(text, /New Email ID: draft-2/);
    assert.match(text, /WARNING/);
    assert.match(text, /could NOT be moved to Trash \(this account has no mailbox with the trash role\)/);
    assert.match(text, /remains in place as a duplicate/);
    // The edit still succeeded, and the echo-back is still reported.
    assert.match(text, /Draft updated successfully/);
    assert.match(text, /subject "Lunch plans"/);
  });

  it('caps a long recipient list rather than dumping every address', () => {
    const many = Array.from({ length: 9 }, (_, i) => `p${i}@example.com`);
    const text = formatEditDraftResult({
      id: 'draft-2',
      replacedDraft: { id: 'draft-1', to: many },
      trashedOldDraftId: 'draft-1',
    });
    assert.match(text, /to p0@example\.com, p1@example\.com, p2@example\.com, p3@example\.com, p4@example\.com \(\+4 more\)/);
    assert.doesNotMatch(text, /p5@example\.com/);
  });

  it('omits the echo-back sentence when the replaced draft had nothing to report', () => {
    const text = formatEditDraftResult({ id: 'draft-2', replacedDraft: { id: 'draft-1' }, trashedOldDraftId: 'draft-1' });
    assert.match(text, /moved to Trash/);
    assert.doesNotMatch(text, /It contained/);
  });
});

// ---------- simplifyMailbox ----------

describe('simplifyMailbox', () => {
  const raw = {
    id: 'mb-1',
    name: 'Inbox',
    role: 'inbox',
    parentId: null,
    totalEmails: 100,
    unreadEmails: 5,
    totalThreads: 80,
    unreadThreads: 3,
    sortOrder: 1,
    isSubscribed: true,
    myRights: { mayReadItems: true, mayDelete: false },
  };

  it('returns core fields by default', () => {
    const result = simplifyMailbox(raw);
    assert.equal(result.id, 'mb-1');
    assert.equal(result.name, 'Inbox');
    assert.equal(result.role, 'inbox');
    assert.equal(result.totalEmails, 100);
    assert.equal(result.unreadEmails, 5);
    assert.equal(result.totalThreads, 80);
    assert.equal(result.unreadThreads, 3);
  });

  it('omits verbose fields by default', () => {
    const result = simplifyMailbox(raw);
    assert.equal(result.sortOrder, undefined);
    assert.equal(result.isSubscribed, undefined);
    assert.equal(result.myRights, undefined);
  });

  it('includes verbose fields when verbose=true', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.sortOrder, 1);
    assert.equal(result.isSubscribed, true);
    assert.deepEqual(result.myRights, { mayReadItems: true, mayDelete: false });
  });

  it('sets falsy role and parentId to undefined', () => {
    const result = simplifyMailbox({ ...raw, role: null, parentId: '' });
    assert.equal(result.role, undefined);
    assert.equal(result.parentId, undefined);
  });
});

// ---------- simplifyIdentity ----------

describe('simplifyIdentity', () => {
  const raw = {
    id: 'id-1',
    name: 'Jonathan',
    email: 'jon@example.com',
    replyTo: [{ email: 'reply@example.com' }],
    mayDelete: true,
    bcc: [{ email: 'bcc@example.com' }],
    textSignature: 'Regards, Jon',
    htmlSignature: '<p>Regards, Jon</p>',
  };

  it('returns core fields by default', () => {
    const result = simplifyIdentity(raw);
    assert.equal(result.id, 'id-1');
    assert.equal(result.name, 'Jonathan');
    assert.equal(result.email, 'jon@example.com');
    assert.deepEqual(result.replyTo, [{ email: 'reply@example.com' }]);
    assert.equal(result.mayDelete, true);
  });

  it('returns the configured signatures by default (#33)', () => {
    const result = simplifyIdentity(raw);
    assert.equal(result.textSignature, 'Regards, Jon');
    assert.equal(result.htmlSignature, '<p>Regards, Jon</p>');
  });

  it('omits verbose fields by default', () => {
    const result = simplifyIdentity(raw);
    assert.equal(result.bcc, undefined);
  });

  it('omits a signature that is absent, blank, or not a string (#33)', () => {
    const absent = simplifyIdentity({ id: 'id-3', name: 'Test', email: 'test@example.com' });
    assert.equal(absent.textSignature, undefined);
    assert.equal(absent.htmlSignature, undefined);
    assert.equal('textSignature' in absent, false);

    const blank = simplifyIdentity({ ...raw, textSignature: '', htmlSignature: '   ' });
    assert.equal(blank.textSignature, undefined);
    assert.equal(blank.htmlSignature, undefined);

    const wrongType = simplifyIdentity({ ...raw, textSignature: null, htmlSignature: 42 });
    assert.equal(wrongType.textSignature, undefined);
    assert.equal(wrongType.htmlSignature, undefined);
  });

  it('includes verbose fields when verbose=true', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.deepEqual(result.bcc, [{ email: 'bcc@example.com' }]);
    assert.equal(result.textSignature, 'Regards, Jon');
    assert.equal(result.htmlSignature, '<p>Regards, Jon</p>');
  });

  it('still reports a blank signature under verbose=true (verbose means everything sent)', () => {
    const result = simplifyIdentity({ ...raw, textSignature: '' }, { verbose: true });
    assert.equal(result.textSignature, '');
  });

  it('omits replyTo when not present', () => {
    const result = simplifyIdentity({ id: 'id-2', name: 'Test', email: 'test@example.com' });
    assert.equal(result.replyTo, undefined);
  });
});

// ---------- simplifyContact ----------

describe('simplifyContact', () => {
  const raw = {
    id: 'ct-1',
    name: { full: 'Alice Smith' },
    emails: {
      work: { address: 'alice@work.com' },
      home: { address: 'alice@home.com' },
    },
    phones: {
      mobile: { number: '+1234567890' },
    },
    organizations: {
      org1: { name: 'Acme Corp' },
    },
    notes: 'VIP client',
    addresses: {
      home: { street: '123 Main St', locality: 'Springfield' },
    },
    titles: {
      t1: { name: 'CEO' },
    },
    online: {
      web: { uri: 'https://example.com' },
    },
    photos: {
      photo1: { uri: 'https://example.com/photo.jpg' },
    },
    anniversaries: {
      birthday: { date: '1990-01-15' },
    },
  };

  it('returns core fields by default', () => {
    const result = simplifyContact(raw);
    assert.equal(result.id, 'ct-1');
    assert.equal(result.name, 'Alice Smith');
    assert.deepEqual(result.emails, ['alice@work.com', 'alice@home.com']);
    assert.deepEqual(result.phones, ['+1234567890']);
    assert.equal(result.organization, 'Acme Corp');
    assert.equal(result.notes, 'VIP client');
  });

  it('omits verbose fields by default', () => {
    const result = simplifyContact(raw);
    assert.equal(result.addresses, undefined);
    assert.equal(result.titles, undefined);
    assert.equal(result.online, undefined);
    assert.equal(result.photos, undefined);
    assert.equal(result.anniversaries, undefined);
  });

  it('includes verbose fields when verbose=true', () => {
    const result = simplifyContact(raw, { verbose: true });
    // addresses flattened to array of objects (hash keys stripped)
    assert.deepEqual(result.addresses, [{ street: '123 Main St', locality: 'Springfield' }]);
    // titles flattened to array of name strings
    assert.deepEqual(result.titles, ['CEO']);
    // online flattened to array of URI strings
    assert.deepEqual(result.online, ['https://example.com']);
    assert.deepEqual(result.photos, raw.photos);
    assert.deepEqual(result.anniversaries, raw.anniversaries);
  });

  it('resolves name from given+surname when full is absent', () => {
    const result = simplifyContact({ id: 'ct-2', name: { given: 'Bob', surname: 'Jones' } });
    assert.equal(result.name, 'Bob Jones');
  });

  it('handles missing name gracefully', () => {
    const result = simplifyContact({ id: 'ct-3' });
    assert.equal(result.name, undefined);
  });

  it('handles missing emails/phones gracefully', () => {
    const result = simplifyContact({ id: 'ct-4' });
    assert.equal(result.emails, undefined);
    assert.equal(result.phones, undefined);
  });
});

// ---------- formatEmailQueryResult ----------

describe('formatEmailQueryResult', () => {
  const makeEmail = (id: string) => ({
    id,
    subject: 'Test',
    from: [{ name: 'Alice', email: 'alice@example.com' }],
    receivedAt: '2026-01-01T00:00:00Z',
    preview: 'Hello',
    keywords: {},
    textBody: [{ partId: 'text1', type: 'text/plain' }],
    htmlBody: [{ partId: 'html1', type: 'text/html' }],
    bodyValues: { text1: { value: 'Hello plain' }, html1: { value: '<p>Hello</p>' } },
  });

  it('omits HTML body and includes size hint', () => {
    const result = formatEmailQueryResult({ items: [makeEmail('e1')], total: 1 });
    assert.ok(!result.includes('<p>Hello</p>'), 'should not include HTML content');
    assert.ok(result.includes('bodyHtmlSize'), 'should include bodyHtmlSize hint');
  });

  it('formats summary line', () => {
    const result = formatEmailQueryResult({ items: [makeEmail('e1'), makeEmail('e2')], total: 50 });
    assert.ok(result.startsWith('Showing 2 of 50 results.'));
  });
});

// ---------- formatContactQueryResult ----------

describe('formatContactQueryResult', () => {
  const rawContact = {
    id: 'ct-1',
    name: { full: 'Alice' },
    addresses: { home: { street: '123 Main St' } },
  };

  it('omits verbose contact fields by default', () => {
    const result = formatContactQueryResult({ items: [rawContact], total: 1 });
    assert.ok(!result.includes('123 Main St'));
  });

  it('includes verbose contact fields when verbose=true', () => {
    const result = formatContactQueryResult({ items: [rawContact], total: 1 }, { verbose: true });
    assert.ok(result.includes('123 Main St'));
  });
});

// ==========================================================================
// Functional test issues — these tests document gaps found during live testing.
// Written as TDD: tests first, then fix the code.
// ==========================================================================

// ---------- simplifyContact: notes bug ----------

describe('simplifyContact notes extraction', () => {
  it('extracts notes from JMAP object format { hash: { note: "text" } }', () => {
    const raw = {
      id: 'ct-notes-1',
      notes: {
        'abc123': { note: 'VIP client' },
      },
    };
    const result = simplifyContact(raw);
    assert.equal(result.notes, 'VIP client');
  });

  it('concatenates multiple notes', () => {
    const raw = {
      id: 'ct-notes-2',
      notes: {
        'n1': { note: 'First note' },
        'n2': { note: 'Second note' },
      },
    };
    const result = simplifyContact(raw);
    assert.ok(result.notes.includes('First note'));
    assert.ok(result.notes.includes('Second note'));
  });

  it('omits notes when empty', () => {
    const raw = {
      id: 'ct-notes-3',
      notes: {
        'n1': { note: '' },
      },
    };
    const result = simplifyContact(raw);
    assert.equal(result.notes, undefined);
  });
});

// ---------- simplifyContact: verbose field simplification ----------

describe('simplifyContact verbose field formatting', () => {
  it('simplifies addresses to flat array in verbose mode', () => {
    const raw = {
      id: 'ct-addr-1',
      addresses: {
        'a1': { street: '123 Main St', locality: 'Springfield', country: 'US', contexts: { work: true } },
        'a2': { street: '456 Oak Ave', locality: 'Portland' },
      },
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.ok(Array.isArray(result.addresses), 'addresses should be an array');
    assert.equal(result.addresses.length, 2);
    assert.equal(result.addresses[0].street, '123 Main St');
  });

  it('simplifies titles to flat array in verbose mode', () => {
    const raw = {
      id: 'ct-title-1',
      titles: {
        't1': { name: 'CEO' },
        't2': { name: 'Founder' },
      },
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.ok(Array.isArray(result.titles), 'titles should be an array');
    assert.ok(result.titles.includes('CEO'));
    assert.ok(result.titles.includes('Founder'));
  });

  it('simplifies online/URLs to flat array in verbose mode', () => {
    const raw = {
      id: 'ct-url-1',
      online: {
        'o1': { uri: 'https://example.com', contexts: { work: true } },
        'o2': { uri: 'https://github.com/example' },
      },
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.ok(Array.isArray(result.online), 'online should be an array');
    assert.ok(result.online.includes('https://example.com'));
    assert.ok(result.online.includes('https://github.com/example'));
  });
});

// ---------- simplifyContact: missing verbose fields ----------

describe('simplifyContact missing verbose fields', () => {
  it('includes addressBookIds in verbose mode', () => {
    const raw = {
      id: 'ct-ab-1',
      addressBookIds: { 'R-k': true },
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.ok(result.addressBookIds !== undefined, 'should include addressBookIds');
  });

  it('includes updated timestamp in verbose mode', () => {
    const raw = {
      id: 'ct-upd-1',
      updated: '2026-03-15T10:00:00Z',
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.equal(result.updated, '2026-03-15T10:00:00Z');
  });

  it('includes kind in verbose mode', () => {
    const raw = {
      id: 'ct-kind-1',
      kind: 'individual',
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.equal(result.kind, 'individual');
  });

  it('includes uid in verbose mode', () => {
    const raw = {
      id: 'ct-uid-1',
      uid: 'urn:uuid:abc-123',
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.equal(result.uid, 'urn:uuid:abc-123');
  });

  it('includes version in verbose mode', () => {
    const raw = {
      id: 'ct-ver-1',
      version: '1.0',
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.equal(result.version, '1.0');
  });

  it('includes prodId in verbose mode', () => {
    const raw = {
      id: 'ct-prod-1',
      prodId: 'Fastmail',
    };
    const result = simplifyContact(raw, { verbose: true });
    assert.equal(result.prodId, 'Fastmail');
  });
});

// ---------- simplifyMailbox: missing verbose fields ----------

describe('simplifyMailbox missing verbose fields', () => {
  const raw = {
    id: 'mb-1',
    name: 'Trash',
    role: 'trash',
    totalEmails: 10,
    unreadEmails: 0,
    totalThreads: 8,
    unreadThreads: 0,
    autoPurge: true,
    hidden: 0,
    purgeOlderThanDays: 31,
    isCollapsed: false,
    autoLearn: true,
    sort: [{ property: 'receivedAt', isAscending: false }],
    identityRef: null,
    learnAsSpam: false,
    suppressDuplicates: false,
  };

  it('includes autoPurge in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.autoPurge, true);
  });

  it('includes hidden in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.hidden, 0);
  });

  it('includes purgeOlderThanDays in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.purgeOlderThanDays, 31);
  });

  it('includes isCollapsed in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.isCollapsed, false);
  });

  it('includes autoLearn in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.autoLearn, true);
  });

  it('includes sort in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.deepEqual(result.sort, [{ property: 'receivedAt', isAscending: false }]);
  });

  it('includes identityRef in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.identityRef, null);
  });

  it('includes learnAsSpam in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.learnAsSpam, false);
  });

  it('includes suppressDuplicates in verbose mode', () => {
    const result = simplifyMailbox(raw, { verbose: true });
    assert.equal(result.suppressDuplicates, false);
  });
});

// ---------- simplifyIdentity: missing verbose fields ----------

describe('simplifyIdentity missing verbose fields', () => {
  const raw = {
    id: 'id-1',
    name: 'Jonathan',
    email: 'jon@example.com',
    verificationState: 'autoverified',
    showInCompose: true,
    saveSentToMailboxId: 'mb-sent',
    displayName: 'Jon G',
    isAutoConfigured: true,
    enableExternalSMTP: false,
    server: 'smtp.fastmail.com',
    port: 587,
    ssl: 'starttls',
    addBccOnSMTP: false,
    saveOnSMTP: false,
    externalCredentialId: null,
    warnings: [],
    useForAutoReply: false,
    verificationCheckTime: '2026-03-01T00:00:00Z',
  };

  it('includes verificationState in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.verificationState, 'autoverified');
  });

  it('includes showInCompose in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.showInCompose, true);
  });

  it('includes saveSentToMailboxId in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.saveSentToMailboxId, 'mb-sent');
  });

  it('includes displayName in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.displayName, 'Jon G');
  });

  it('includes isAutoConfigured in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.isAutoConfigured, true);
  });

  it('includes enableExternalSMTP in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.enableExternalSMTP, false);
  });

  it('includes server in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.server, 'smtp.fastmail.com');
  });

  it('includes port in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.port, 587);
  });

  it('includes ssl in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.ssl, 'starttls');
  });

  it('includes addBccOnSMTP in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.addBccOnSMTP, false);
  });

  it('includes saveOnSMTP in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.saveOnSMTP, false);
  });

  it('includes externalCredentialId in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.externalCredentialId, null);
  });

  it('includes warnings in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.deepEqual(result.warnings, []);
  });

  it('includes useForAutoReply in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.useForAutoReply, false);
  });

  it('includes verificationCheckTime in verbose mode', () => {
    const result = simplifyIdentity(raw, { verbose: true });
    assert.equal(result.verificationCheckTime, '2026-03-01T00:00:00Z');
  });
});
