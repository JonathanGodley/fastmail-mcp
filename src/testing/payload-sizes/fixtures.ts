// Synthetic server data for the payload-size baseline (scripts/payload-sizes.mts).
//
// What it represents: one personal account, as the five default reads see it.
// - list_emails: a 20-row Inbox/Archive page out of 1342, mostly plain person-to-person mail.
//   Most rows carry only the always-present fields. A minority carry cc (3), replyTo (2),
//   List-Unsubscribe (4 newsletters), attachments (3), a reply chain (4), a flag (2) or a
//   second, custom folder (2). 23 messages sit in Trash and Spam, so the exclusion note fires.
// - get_email: one received multipart/alternative message with a PDF attached.
// - get_thread: a four-message reply chain across Inbox and Sent.
// - list_calendar_events: a default call (no dates) on 4 May 2026 in Sydney, over two
//   calendars: 12 rows, among them a weekly series, one instance of a monthly series that began
//   in February, two all-day events and a meeting with attendees.
// - list_contacts: a 20-card page out of 143: mostly a name and one address, a few phones and
//   organisations, one note, one contact group.
// Some names and subjects are non-ASCII, so bytes and characters differ as they do in real mail.
//
// Every value is invented. Addresses, Message-IDs and UIDs use example.com or a reserved
// suffix only, because the secret scanner flags any other domain.
//
// The JMAP responses are keyed by the call id each client puts in its batch (see `answer` in
// measure.ts). The DAV stub ignores a request's filters, timeRange, expand and urlFilter: the
// expanded objects are the server's answer as data, so the window and expansion go unexercised.

import type { JmapSession } from '../../jmap-client.js';

export const FIXTURE_ZONE = 'Australia/Sydney';
// 10:00 on Monday 4 May 2026 in Sydney: the default calendar window is 4 May to 4 June.
export const FIXTURE_NOW = Date.parse('2026-05-04T00:00:00Z');

const ACCOUNT = 'u7f3e2a10';

export const SESSION: JmapSession = {
  apiUrl: 'https://api.example.com/jmap/api/',
  accountId: ACCOUNT,
  capabilities: {
    'urn:ietf:params:jmap:core': {},
    'urn:ietf:params:jmap:mail': {},
    'urn:ietf:params:jmap:contacts': {},
  },
  primaryAccounts: {
    'urn:ietf:params:jmap:mail': ACCOUNT,
    'urn:ietf:params:jmap:contacts': ACCOUNT,
  },
};

export const MAILBOXES = [
  { id: 'P-inbox', name: 'Inbox', role: 'inbox', parentId: null, totalEmails: 1204, unreadEmails: 7 },
  { id: 'P-archive', name: 'Archive', role: 'archive', parentId: null, totalEmails: 9381, unreadEmails: 0 },
  { id: 'P-drafts', name: 'Drafts', role: 'drafts', parentId: null, totalEmails: 3, unreadEmails: 0 },
  { id: 'P-sent', name: 'Sent', role: 'sent', parentId: null, totalEmails: 2210, unreadEmails: 0 },
  { id: 'P-trash', name: 'Trash', role: 'trash', parentId: null, totalEmails: 19, unreadEmails: 2 },
  { id: 'P-spam', name: 'Spam', role: 'junk', parentId: null, totalEmails: 4, unreadEmails: 4 },
  { id: 'P-receipts', name: 'Receipts', role: null, parentId: null, totalEmails: 412, unreadEmails: 0 },
  { id: 'P-house', name: 'House', role: null, parentId: null, totalEmails: 88, unreadEmails: 0 },
];

type Address = { name: string | null; email: string };
const person = (name: string | null, email: string): Address => ({ name, email });

const ME = person('Robin Castellano', 'robin@example.com');

// The part structure a compact read returns: no body values, so the simplifier reports
// bodyTextSize from these sizes.
function textPart(partId: string, size: number) {
  return {
    partId, blobId: `G${partId}${size.toString(16)}`, size, name: null, type: 'text/plain',
    charset: 'utf-8', disposition: null, cid: null, language: null, location: null,
  };
}

interface RowSpec {
  id: string;
  from: Address;
  subject: string;
  receivedAt: string;
  preview: string;
  textSize: number;
  size: number;
  seen?: boolean;
  flagged?: boolean;
  answered?: boolean;
  to?: Address[];
  cc?: Address[];
  replyTo?: Address[];
  listUnsubscribe?: string[];
  hasAttachment?: boolean;
  inReplyTo?: string[];
  references?: string[];
  mailboxes?: string[];
}

// One Email/get row with the compact property set, nulls where the message has no value, as
// JMAP returns them.
function compactRow(spec: RowSpec) {
  const keywords: Record<string, true> = {};
  if (spec.seen !== false) keywords.$seen = true;
  if (spec.flagged) keywords.$flagged = true;
  if (spec.answered) keywords.$answered = true;
  return {
    id: spec.id,
    blobId: `B${spec.id}`,
    threadId: `T${spec.id}`,
    mailboxIds: Object.fromEntries((spec.mailboxes ?? ['P-inbox']).map((m) => [m, true])),
    keywords,
    size: spec.size,
    receivedAt: spec.receivedAt,
    sentAt: spec.receivedAt,
    messageId: [`${spec.id.toLowerCase()}.${spec.receivedAt.slice(0, 10)}@example.com`],
    inReplyTo: spec.inReplyTo ?? null,
    references: spec.references ?? null,
    from: [spec.from],
    to: spec.to ?? [ME],
    cc: spec.cc ?? null,
    bcc: null,
    replyTo: spec.replyTo ?? null,
    subject: spec.subject,
    preview: spec.preview,
    hasAttachment: spec.hasAttachment ?? false,
    'header:List-Unsubscribe:asURLs': spec.listUnsubscribe ?? null,
    textBody: [textPart('1', spec.textSize)],
  };
}

const LIST_ROWS: RowSpec[] = [
  {
    id: 'Mb81f0a', from: person('Priya Raman', 'priya.raman@example.com'), subject: 'Re: Saturday plans',
    receivedAt: '2026-04-30T09:14:22Z', preview: 'Sounds good, we can bring the salad. Is 1pm still the plan or did that move?',
    textSize: 640, size: 9120, seen: false,
    inReplyTo: ['ma11c2e.2026-04-29@example.com'], references: ['ma11c2e.2026-04-29@example.com'],
  },
  {
    id: 'Mb81e77', from: person('The Weekly Allotment', 'newsletter@allotment.example'), subject: 'May sowing guide: beans, beets and what to do about slugs',
    receivedAt: '2026-04-30T06:00:03Z', preview: 'This month in the garden: direct-sow beans once the soil warms, thin your beetroot, and three slug deterrents that actually work. Plus readers’ photos from the autumn harvest and a recipe for green tomato chutney.',
    textSize: 11840, size: 68302,
    listUnsubscribe: ['https://allotment.example/u/7c2d9e', 'mailto:leave-7c2d9e@allotment.example'],
  },
  {
    id: 'Mb81d40', from: person('Sam Okonkwo', 'sam.okonkwo@example.com'), subject: 'Photos from the weekend',
    receivedAt: '2026-04-29T22:41:10Z', preview: 'A few from the beach, the rest are in the shared album.',
    textSize: 212, size: 4402911, hasAttachment: true,
  },
  {
    id: 'Mb81c05', from: person('Northside Dental', 'appointments@dental.example'), subject: 'Appointment reminder: Thursday 7 May, 9:00am',
    receivedAt: '2026-04-29T21:00:00Z', preview: 'This is a reminder of your appointment with Dr Lindqvist on Thursday 7 May at 9:00am. Reply C to confirm or call us to reschedule.',
    textSize: 1033, size: 14871,
  },
  {
    id: 'Mb81ab2', from: person('Zoë Laurent', 'zoe.laurent@example.com'), subject: 'Café on Thursday?',
    receivedAt: '2026-04-29T04:12:56Z', preview: 'Are you free for a coffee Thursday arvo? The new place on King St does a great flat white.',
    textSize: 388, size: 6022, answered: true,
  },
  {
    id: 'Mb81a19', from: person('Harbour Energy', 'billing@energy.example'), subject: 'Your April bill is ready',
    receivedAt: '2026-04-28T23:30:00Z', preview: 'Your bill for 28 March to 27 April is $142.65, due 19 May. View it online or download the PDF attached.',
    textSize: 2410, size: 98213, hasAttachment: true, mailboxes: ['P-inbox', 'P-house'],
  },
  {
    id: 'Mb81998', from: person('Priya Raman', 'priya.raman@example.com'), subject: 'Saturday plans',
    receivedAt: '2026-04-28T08:02:41Z', preview: 'We were thinking a picnic at the park if the weather holds. Can you make it?',
    textSize: 702, size: 8904, mailboxes: ['P-archive'],
    cc: [person('Dev Raman', 'dev.raman@example.com')],
  },
  {
    id: 'Mb81870', from: person('Market Lane Books', 'orders@books.example'), subject: 'Your order has shipped',
    receivedAt: '2026-04-27T03:15:09Z', preview: 'Good news: order #48213 is on its way. Tracking number and estimated delivery inside.',
    textSize: 3301, size: 41270, mailboxes: ['P-receipts'],
    replyTo: [person('Market Lane Books Support', 'help@books.example')],
  },
  {
    id: 'Mb81702', from: person('Ellis Moreau', 'ellis.moreau@example.com'), subject: 'Re: Re: Fence quote',
    receivedAt: '2026-04-26T10:47:33Z', preview: 'That works for me. I’ll let the contractor know we’re going with the timber option and split it 50/50.',
    textSize: 2871, size: 17412, flagged: true,
    inReplyTo: ['mb81531.2026-04-25@example.com'],
    references: ['mb81402.2026-04-24@example.com', 'mb81531.2026-04-25@example.com'],
  },
  {
    id: 'Mb81688', from: person('Coastal Running Club', 'club@running.example'), subject: 'Sunday long run: route change',
    receivedAt: '2026-04-26T00:05:17Z', preview: 'Hi all, the coastal path is closed past the point so we’ll loop back through the reserve. Same start time, 6:30am at the kiosk.',
    textSize: 1240, size: 12188,
    listUnsubscribe: ['https://running.example/lists/unsubscribe?m=91f3'],
  },
  {
    id: 'Mb81531', from: ME, subject: 'Re: Fence quote',
    receivedAt: '2026-04-25T07:20:00Z', preview: 'Thanks Ellis. Timber looks better to me too. Shall I send the deposit or will you?',
    textSize: 1960, size: 11031, mailboxes: ['P-archive'],
    to: [person('Ellis Moreau', 'ellis.moreau@example.com')],
    inReplyTo: ['mb81402.2026-04-24@example.com'], references: ['mb81402.2026-04-24@example.com'],
  },
  {
    id: 'Mb81477', from: person('Jordan Whitfield', 'jordan.whitfield@example.com'), subject: 'Book club: May pick',
    receivedAt: '2026-04-24T11:31:48Z', preview: 'The vote is in. May’s book is the one about the lighthouse keeper. See you on the 14th.',
    textSize: 540, size: 7310,
    cc: [person('Ana Sousa', 'ana.sousa@example.com'), person('Kai Brennan', 'kai.brennan@example.com'), person(null, 'mel.h@example.com')],
  },
  {
    id: 'Mb81402', from: person('Ellis Moreau', 'ellis.moreau@example.com'), subject: 'Fence quote',
    receivedAt: '2026-04-24T05:09:12Z', preview: 'Got two quotes back: one for colorbond at $3,850 and one for timber at $4,200. Pros and cons below, let me know what you think.',
    textSize: 3482, size: 186204, hasAttachment: true, mailboxes: ['P-archive', 'P-house'],
  },
  {
    id: 'Mb81366', from: person('City Library', 'noreply@library.example'), subject: 'Items due soon',
    receivedAt: '2026-04-23T20:00:00Z', preview: '2 items are due on 30 April. Renew online to avoid fees.',
    textSize: 812, size: 10240,
  },
  {
    id: 'Mb81290', from: person('Mia Castellano', 'mia.c@example.com'), subject: 'Mum’s birthday',
    receivedAt: '2026-04-23T09:58:01Z', preview: 'Shall we go halves on something? I was thinking the pottery class she mentioned. Let me know before the weekend.',
    textSize: 431, size: 5980, flagged: true, seen: false,
  },
  {
    id: 'Mb81203', from: person('Streamline', 'hello@streamline.example'), subject: 'New this week: three series we think you’ll like',
    receivedAt: '2026-04-23T02:00:00Z', preview: 'Handpicked for you: a slow-burn mystery set in the fjords, a cooking competition with a twist, and the documentary everyone is talking about.',
    textSize: 6120, size: 52807, seen: false,
    listUnsubscribe: ['https://streamline.example/prefs?u=4b1f'],
  },
  {
    id: 'Mb81188', from: person('Tomás Herrera', 'tomas.herrera@example.com'), subject: 'Bike for sale',
    receivedAt: '2026-04-22T13:44:27Z', preview: 'Still interested in the road bike? I can do $450 and throw in the lights.',
    textSize: 298, size: 4710, mailboxes: ['P-archive'],
    replyTo: [person('Tomás Herrera', 'tomas.h@example.com')],
  },
  {
    id: 'Mb81075', from: person('Council Waste Services', 'waste@council.example'), subject: 'Bulky waste collection booked',
    receivedAt: '2026-04-21T22:16:40Z', preview: 'Your collection is booked for Tuesday 12 May. Place items on the kerb the night before, no more than 2 cubic metres.',
    textSize: 1655, size: 15102, mailboxes: ['P-archive'],
  },
  {
    id: 'Mb80f61', from: person('Ana Sousa', 'ana.sousa@example.com'), subject: 'Recipe you asked for',
    receivedAt: '2026-04-21T08:27:05Z', preview: 'Here it is! The trick is to toast the spices first. Let me know how it turns out.',
    textSize: 2204, size: 9433, mailboxes: ['P-archive'],
    cc: [person('Priya Raman', 'priya.raman@example.com')],
  },
  {
    id: 'Mb80e02', from: person('Fieldnotes Digest', 'digest@fieldnotes.example'), subject: 'Weekend reading: 6 long reads',
    receivedAt: '2026-04-20T21:00:00Z', preview: 'On rewilding urban rivers, the economics of the corner shop, a profile of a lighthouse restorer, and three more pieces worth your time this weekend.',
    textSize: 9311, size: 61205, mailboxes: ['P-archive'],
    listUnsubscribe: ['https://fieldnotes.example/unsubscribe/2e9a', 'mailto:unsubscribe-2e9a@fieldnotes.example'],
  },
];

export const LIST_EMAILS_RESPONSES: Record<string, [string, any]> = {
  query: ['Email/query', {
    accountId: ACCOUNT, queryState: 'q1', canCalculateChanges: true, position: 0,
    ids: LIST_ROWS.map((r) => r.id), total: 1342,
  }],
  emails: ['Email/get', { accountId: ACCOUNT, state: 's1', list: LIST_ROWS.map(compactRow), notFound: [] }],
  count: ['Email/query', {
    accountId: ACCOUNT, queryState: 'q1', canCalculateChanges: true, position: 0, ids: [], total: 1365,
  }],
};

const SINGLE_TEXT = [
  'Hi Robin,',
  '',
  'Thanks for coming along to the open day on Saturday. As promised, I’ve attached the',
  'information pack for the community garden plots, including the fee schedule and the',
  'waiting-list form.',
  '',
  'A few things that came up on the day:',
  '',
  '- Plots are 3m x 6m. Half plots are available but the wait is longer.',
  '- Water is included. Tools are shared from the shed (the code is on the noticeboard).',
  '- Working bees are the first Sunday of each month, 9am to 12pm. We ask everyone to come',
  '  to at least four a year.',
  '- Composting is communal; please don’t put weeds that have gone to seed in the bays.',
  '',
  'If you’d like to go ahead, fill in the form and drop it in the letterbox at the gate, or',
  'reply to this email with a scan. We’ll be in touch when a plot comes up, usually within',
  'two to three months.',
  '',
  'Our next open day is in June if you’d like to bring anyone else along.',
  '',
  'Cheers,',
  'Bérénice Aalto',
  'Plot coordinator, Riverside Community Garden',
  '',
].join('\n');

const SINGLE_HTML = '<div dir="ltr"><p>Hi Robin,</p><p>Thanks for coming along to the open day on Saturday. '
  + 'As promised, I’ve attached the information pack for the community garden plots, including the '
  + 'fee schedule and the waiting-list form.</p><p>A few things that came up on the day:</p><ul>'
  + '<li>Plots are 3m x 6m. Half plots are available but the wait is longer.</li>'
  + '<li>Water is included. Tools are shared from the shed (the code is on the noticeboard).</li>'
  + '<li>Working bees are the first Sunday of each month, 9am to 12pm. We ask everyone to come to at '
  + 'least four a year.</li><li>Composting is communal; please don’t put weeds that have gone to '
  + 'seed in the bays.</li></ul><p>If you’d like to go ahead, fill in the form and drop it in the '
  + 'letterbox at the gate, or reply to this email with a scan. We’ll be in touch when a plot comes '
  + 'up, usually within two to three months.</p><p>Our next open day is in June if you’d like to '
  + 'bring anyone else along.</p><p>Cheers,<br>Bérénice Aalto<br>Plot coordinator, Riverside '
  + 'Community Garden</p></div>';

export const GET_EMAIL_ID = 'Mc20a7f';

const SINGLE_EMAIL = {
  id: GET_EMAIL_ID,
  blobId: 'Bc20a7f',
  threadId: 'Tc20a7f',
  mailboxIds: { 'P-inbox': true },
  keywords: { $seen: true },
  size: 312744,
  receivedAt: '2026-04-27T01:03:51Z',
  sentAt: '2026-04-27T11:03:44+10:00',
  messageId: ['c20a7f.2026-04-27@garden.example'],
  inReplyTo: null,
  references: null,
  from: [person('Bérénice Aalto', 'plots@garden.example')],
  to: [ME],
  cc: null,
  bcc: null,
  replyTo: null,
  subject: 'Community garden: information pack',
  preview: 'Hi Robin, Thanks for coming along to the open day on Saturday. As promised, I’ve attached the information pack for the community garden plots, including the fee schedule and the waiting-list form. A few things that came up on the day:',
  hasAttachment: true,
  'header:List-Unsubscribe:asURLs': null,
  textBody: [{ partId: '1', blobId: 'Gc20a7f1', type: 'text/plain', size: 1312, name: null, disposition: null, cid: null }],
  htmlBody: [{ partId: '2', blobId: 'Gc20a7f2', type: 'text/html', size: 1478, name: null, disposition: null, cid: null }],
  attachments: [{
    partId: '3', blobId: 'Gc20a7f3', type: 'application/pdf', size: 304118,
    name: 'Riverside garden plots - information pack.pdf', disposition: 'attachment', cid: null,
  }],
  bodyValues: {
    1: { value: SINGLE_TEXT, isEncodingProblem: false, isTruncated: false },
    2: { value: SINGLE_HTML, isEncodingProblem: false, isTruncated: false },
  },
  'header:X-Forwarded-Message-Id:asMessageIds': null,
  'header:X-Fastmail-MCP-Source-Id:asText': null,
};

export const GET_EMAIL_RESPONSES: Record<string, [string, any]> = {
  email: ['Email/get', { accountId: ACCOUNT, state: 's1', list: [SINGLE_EMAIL], notFound: [] }],
  mailboxes: ['Mailbox/get', { accountId: ACCOUNT, state: 'm1', list: MAILBOXES.map(({ id, name, role }) => ({ id, name, role })), notFound: [] }],
};

export const THREAD_ID = 'Td41c09';
const ELLIS = person('Ellis Moreau', 'ellis.moreau@example.com');
const msgId = (n: number) => `d41c0${n}.2026-04@example.com`;

const THREAD_ROWS: RowSpec[] = [
  {
    id: 'Md41c01', from: ELLIS, subject: 'Shared driveway resurfacing',
    receivedAt: '2026-04-14T00:12:09Z', preview: 'Hi Robin, the council letter about the driveway came through. They want both owners to sign off before the resurfacing starts. Can we talk this week?',
    textSize: 1840, size: 10612,
  },
  {
    id: 'Md41c02', from: ME, subject: 'Re: Shared driveway resurfacing',
    receivedAt: '2026-04-14T06:40:51Z', preview: 'Sure, Wednesday evening works. Do they say who pays for the kerb section?',
    textSize: 2410, size: 12008, mailboxes: ['P-sent'], to: [ELLIS],
    inReplyTo: [msgId(1)], references: [msgId(1)],
  },
  {
    id: 'Md41c03', from: ELLIS, subject: 'Re: Shared driveway resurfacing',
    receivedAt: '2026-04-15T21:03:27Z', preview: 'Council covers the kerb, we split the rest. Quote attached. Wednesday at 7?',
    textSize: 3120, size: 214330, hasAttachment: true, answered: true,
    inReplyTo: [msgId(2)], references: [msgId(1), msgId(2)],
  },
  {
    id: 'Md41c04', from: ME, subject: 'Re: Shared driveway resurfacing',
    receivedAt: '2026-04-16T02:18:44Z', preview: 'Wednesday at 7 is good. I’ve read the quote, looks reasonable. See you then.',
    textSize: 3655, size: 14211, mailboxes: ['P-sent'], to: [ELLIS],
    inReplyTo: [msgId(3)], references: [msgId(1), msgId(2), msgId(3)],
  },
];

export const GET_THREAD_RESPONSES: Record<string, [string, any]> = {
  // get_thread first asks whether its id is an email id; a thread id comes back not found.
  checkEmail: ['Email/get', { accountId: ACCOUNT, state: 's1', list: [], notFound: [THREAD_ID] }],
  getThread: ['Thread/get', { accountId: ACCOUNT, state: 't1', list: [{ id: THREAD_ID, emailIds: THREAD_ROWS.map((r) => r.id) }], notFound: [] }],
  emails: ['Email/get', {
    accountId: ACCOUNT, state: 's1', notFound: [],
    list: THREAD_ROWS.map((r, i) => ({ ...compactRow(r), threadId: THREAD_ID, messageId: [msgId(i + 1)] })),
  }],
  mailboxes: ['Mailbox/get', { accountId: ACCOUNT, state: 'm1', list: MAILBOXES.map(({ id, name, role }) => ({ id, name, role })), notFound: [] }],
};

interface CardSpec {
  id: string;
  given: string;
  surname?: string;
  emails?: string[];
  phones?: string[];
  org?: string;
  note?: string;
}

function card(spec: CardSpec) {
  const c: any = {
    '@type': 'Card',
    version: '1.0',
    id: spec.id,
    uid: `urn:uuid:${spec.id.slice(1).padEnd(8, '0')}-4c1e-4a7b-9d2f-6b8e1c3a5f70`,
    addressBookIds: { personal: true },
    kind: 'individual',
    created: '2025-09-14T03:22:10Z',
    updated: '2026-02-03T08:41:55Z',
    name: {
      components: [
        { kind: 'given', value: spec.given },
        ...(spec.surname ? [{ kind: 'surname', value: spec.surname }] : []),
      ],
      isOrdered: true,
    },
  };
  if (spec.emails) {
    c.emails = Object.fromEntries(spec.emails.map((address, i) => [`e${i + 1}`, { address, contexts: { private: true } }]));
  }
  if (spec.phones) {
    c.phones = Object.fromEntries(spec.phones.map((number, i) => [`p${i + 1}`, { number, features: { mobile: true } }]));
  }
  if (spec.org) c.organizations = { o1: { name: spec.org } };
  if (spec.note) c.notes = { n1: { note: spec.note } };
  return c;
}

const CARDS = [
  card({ id: 'C0a11', given: 'Ana', surname: 'Sousa', emails: ['ana.sousa@example.com'] }),
  card({ id: 'C0a12', given: 'Bérénice', surname: 'Aalto', emails: ['plots@garden.example'], org: 'Riverside Community Garden' }),
  card({ id: 'C0a13', given: 'Dev', surname: 'Raman', emails: ['dev.raman@example.com'] }),
  card({ id: 'C0a14', given: 'Astrid', surname: 'Lindqvist', phones: ['+1 202 555 0143'], org: 'Northside Dental' }),
  card({ id: 'C0a15', given: 'Ellis', surname: 'Moreau', emails: ['ellis.moreau@example.com'], phones: ['+44 20 7946 0811'], note: 'Neighbour at no. 14. Shares the driveway.' }),
  card({ id: 'C0a16', given: 'Hana', surname: 'Kobayashi', emails: ['hana.k@example.com'] }),
  card({ id: 'C0a17', given: 'Jordan', surname: 'Whitfield', emails: ['jordan.whitfield@example.com'] }),
  card({ id: 'C0a18', given: 'Kai', surname: 'Brennan', emails: ['kai.brennan@example.com', 'kai@work.example'] }),
  card({ id: 'C0a19', given: 'Lena', surname: 'Fischer', emails: ['lena.fischer@example.com'] }),
  card({ id: 'C0a1a', given: 'Mel', surname: 'Hartley', emails: ['mel.h@example.com'] }),
  card({ id: 'C0a1b', given: 'Mia', surname: 'Castellano', emails: ['mia.c@example.com'], phones: ['+44 20 7946 0377'] }),
  card({ id: 'C0a1c', given: 'Noah', surname: 'Adeyemi', emails: ['noah.adeyemi@example.com'] }),
  card({ id: 'C0a1d', given: 'Olivia', surname: 'Grant', emails: ['olivia.grant@example.com'], org: 'Grant & Daughters Plumbing' }),
  card({ id: 'C0a1e', given: 'Priya', surname: 'Raman', emails: ['priya.raman@example.com'], phones: ['+1 202 555 0178'] }),
  card({ id: 'C0a1f', given: 'Quinn', surname: 'Delacroix', emails: ['quinn.d@example.com'] }),
  card({ id: 'C0a20', given: 'Sam', surname: 'Okonkwo', emails: ['sam.okonkwo@example.com'] }),
  card({ id: 'C0a21', given: 'Søren', surname: 'Madsen', emails: ['soren.madsen@example.com'] }),
  card({ id: 'C0a22', given: 'Tomás', surname: 'Herrera', emails: ['tomas.herrera@example.com', 'tomas.h@example.com'] }),
  card({ id: 'C0a23', given: 'Zoë', surname: 'Laurent', emails: ['zoe.laurent@example.com'] }),
  {
    '@type': 'Card', version: '1.0', id: 'C0a24', uid: 'urn:uuid:0a240000-4c1e-4a7b-9d2f-6b8e1c3a5f70',
    addressBookIds: { personal: true }, kind: 'group', created: '2025-11-02T10:00:00Z', updated: '2026-03-20T09:12:31Z',
    name: { full: 'Book club' },
    members: { 'urn:uuid:a1700000-4c1e-4a7b-9d2f-6b8e1c3a5f70': true, 'urn:uuid:a1100000-4c1e-4a7b-9d2f-6b8e1c3a5f70': true, 'urn:uuid:a1800000-4c1e-4a7b-9d2f-6b8e1c3a5f70': true },
  },
];

export const LIST_CONTACTS_RESPONSES: Record<string, [string, any]> = {
  query: ['ContactCard/query', {
    accountId: ACCOUNT, queryState: 'cq1', canCalculateChanges: false, position: 0,
    ids: CARDS.map((c) => c.id), total: 143,
  }],
  contacts: ['ContactCard/get', { accountId: ACCOUNT, state: 'c1', list: CARDS, notFound: [] }],
};

const CALDAV_HOME = 'https://caldav.example.invalid/dav/calendars/user/robin@example.com/';
export const CALENDARS = [
  { displayName: 'Personal', url: `${CALDAV_HOME}personal/`, ctag: 'ct-1', components: ['VEVENT'], resourcetype: ['calendar'] },
  { displayName: 'Work', url: `${CALDAV_HOME}work/`, ctag: 'ct-2', components: ['VEVENT'], resourcetype: ['calendar'] },
];
const [PERSONAL, WORK] = CALENDARS.map((c) => c.url);

function vcalendar(...vevents: string[][]): string[] {
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//CyrusIMAP.org/Cyrus 3.9//EN', ...vevents.flat(), 'END:VCALENDAR'];
}

function vevent(lines: string[]): string[] {
  return ['BEGIN:VEVENT', 'DTSTAMP:20260401T000000Z', ...lines, 'END:VEVENT'];
}

export interface CalendarObjectFixture {
  calendarUrl: string;
  url: string;
  // Joined with CRLF at load: a .ics file would be normalised to LF by .gitattributes.
  lines: string[];
}

// What the server answers to an expanded time-range query: each in-window instance of a
// series as its own VEVENT, in UTC, with RRULE stripped.
export const EXPANDED_OBJECTS: CalendarObjectFixture[] = [
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}4e1d2c7a-dent.ics`,
    lines: vcalendar(vevent([
      'UID:4e1d2c7a-0b6f-4d0e-9a51-8c3f2e7d1a90', 'SEQUENCE:0',
      'DTSTART:20260506T230000Z', 'DTEND:20260506T234500Z',
      'SUMMARY:Dentist', 'LOCATION:Northside Dental\\, 12 Grove St',
    ])),
  },
  {
    calendarUrl: WORK, url: `${WORK}9b3a51f0-standup.ics`,
    lines: vcalendar(
      ...['20260506', '20260513', '20260520', '20260527'].map((day, i) => vevent([
        'UID:9b3a51f0-77c2-4f1e-8d6a-2e5b9c0f4d13',
        ...(i === 0 ? [] : [`RECURRENCE-ID:${day}T223000Z`]),
        `DTSTART:${day}T223000Z`, `DTEND:${day}T224500Z`,
        'SUMMARY:Team stand-up', 'LOCATION:https://meet.example.com/standup-7f2',
      ])),
    ),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}c7f02e19-bookclub.ics`,
    lines: vcalendar(vevent([
      'UID:c7f02e19-5a3d-4b8c-a1e6-0d9f7b2c4e85',
      // The series began in February, so this is not its first instance.
      'RECURRENCE-ID:20260514T090000Z',
      'DTSTART:20260514T090000Z', 'DTEND:20260514T110000Z',
      'SUMMARY:Book club', 'LOCATION:Jordan’s place',
      'DESCRIPTION:May: the lighthouse keeper one. Bring a plate.',
    ])),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}2a6e9d40-away.ics`,
    lines: vcalendar(vevent([
      'UID:2a6e9d40-1c8b-4f7a-b3e2-5d0c9a8f6e17', 'TRANSP:TRANSPARENT',
      'DTSTART;VALUE=DATE:20260522', 'DTEND;VALUE=DATE:20260526',
      'SUMMARY:Long weekend away', 'LOCATION:Blue Mountains',
    ])),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}5f8c3b21-mum.ics`,
    lines: vcalendar(vevent([
      'UID:5f8c3b21-9e4d-4a6f-8b0c-7d2e1f3a9c64', 'TRANSP:TRANSPARENT',
      'DTSTART;VALUE=DATE:20260519', 'DTEND;VALUE=DATE:20260520',
      'SUMMARY:Mum’s birthday 🎂',
    ])),
  },
  {
    calendarUrl: WORK, url: `${WORK}e3d7a9c2-planning.ics`,
    lines: vcalendar(vevent([
      'UID:e3d7a9c2-planning-q3@work.example', 'SEQUENCE:2',
      'DTSTART:20260512T000000Z', 'DTEND:20260512T013000Z',
      'SUMMARY:Q3 planning', 'LOCATION:Level 4\\, Room Kestrel',
      'DESCRIPTION:Agenda:\\n1. Q2 review\\n2. Roadmap for Q3\\n3. Hiring\\n\\nPre-read is in the shared folder.',
      'ORGANIZER;CN=Hana Kobayashi:mailto:hana.k@work.example',
      'ATTENDEE;CN=Hana Kobayashi;PARTSTAT=ACCEPTED;ROLE=CHAIR:mailto:hana.k@work.example',
      'ATTENDEE;CN=Robin Castellano;PARTSTAT=ACCEPTED;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:robin@example.com',
      'ATTENDEE;CN=Noah Adeyemi;PARTSTAT=TENTATIVE;ROLE=REQ-PARTICIPANT;RSVP=TRUE:mailto:noah@work.example',
      'ATTENDEE;CN=Lena Fischer;PARTSTAT=NEEDS-ACTION;ROLE=OPT-PARTICIPANT;RSVP=TRUE:mailto:lena@work.example',
      'ATTENDEE;CN=Kestrel;CUTYPE=ROOM;PARTSTAT=ACCEPTED:mailto:room-kestrel@work.example',
    ])),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}81b4f6e3-flight.ics`,
    lines: vcalendar(vevent([
      'UID:81b4f6e3-2d9a-4c5b-a7e0-3f6c8d1b9a42',
      'DTSTART:20260521T210000Z', 'DTEND:20260521T223000Z',
      'SUMMARY:Flight SYD → MEL', 'LOCATION:Sydney Airport T2',
      'DESCRIPTION:Booking ref QX7R2P. Seat 14C.',
    ])),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}d2e5c8a1-cafe.ics`,
    lines: vcalendar(vevent([
      'UID:d2e5c8a1-6f3b-4e9d-8c2a-1b7e0f5d3c96',
      'DTSTART:20260507T050000Z', 'DTEND:20260507T060000Z',
      'SUMMARY:Café with Zoë', 'LOCATION:King St',
    ])),
  },
  {
    calendarUrl: PERSONAL, url: `${PERSONAL}a9c1e7f4-school.ics`,
    lines: vcalendar(vevent([
      'UID:a9c1e7f4-3b8d-4f2e-9a6c-5e0d2b7f1c38',
      'DTSTART:20260528T070000Z', 'DTEND:20260528T074500Z',
      'SUMMARY:Parent-teacher interview',
    ])),
  },
];

// The stored resources the follow-up read returns for each row the expanded listing cannot
// classify on its own. Only the one-off events are asked about, and each is stored as listed.
export const STORED_OBJECTS: Record<string, string[]> = Object.fromEntries(
  EXPANDED_OBJECTS.map((o) => [o.url, o.lines]),
);
