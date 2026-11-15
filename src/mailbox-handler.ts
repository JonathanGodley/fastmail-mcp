import { buildMailboxPathMap, filterMailboxesByParent, filterTopLevelMailboxes } from './jmap-client.js';
import { simplifyMailbox, buildUnpathableMailboxNote } from './response-formatters.js';
import { coerceBool, describeUntrusted, InvalidInputError, toolJson } from './coerce.js';

/**
 * The slice of the JMAP client the mailbox tools need, so the handlers can be exercised
 * with a stub instead of a live account. `JmapClient` satisfies it structurally.
 */
export interface MailboxClient {
  getMailboxes(): Promise<any[]>;
  createMailbox(input: { name: string; parent?: string }): Promise<{ mailbox: any; created: any; path?: string }>;
  deleteMailbox(mailbox: string): Promise<{ deleted: string; deletedMailbox?: any }>;
}

export type ToolContent = Array<{ type: 'text'; text: string }>;

/**
 * list_mailboxes. The first content item is ALWAYS the JSON array and nothing else: a
 * caller parses it directly, so the no-path note rides as a separate item.
 */
export async function listMailboxes(args: any, client: MailboxClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const verbose = coerceBool(args?.verbose, 'verbose') ?? false;
  const topLevelOnly = coerceBool(args?.topLevelOnly, 'topLevelOnly') ?? false;
  // Before the fetch, so it costs no round trip. "Present" is filterMailboxesByParent's test,
  // so a blank parent gets this refusal rather than the blank-parent one.
  if (topLevelOnly && args?.parent !== undefined && args?.parent !== null) {
    throw new InvalidInputError(
      'topLevelOnly cannot be combined with parent: topLevelOnly lists the mailboxes that have no parent, ' +
        "parent lists one mailbox's direct children. Pass one of them.",
    );
  }

  // `path` is root-anchored, so it needs the WHOLE tree even when the listing is narrowed
  // to one parent's children: fetch unnarrowed, then filter the list already in hand.
  const mailboxes = await client.getMailboxes();
  const { paths } = buildMailboxPathMap(mailboxes);
  const shown = topLevelOnly ? filterTopLevelMailboxes(mailboxes) : filterMailboxesByParent(mailboxes, args?.parent);

  // raw promises no path, so it gets no note about one.
  if (raw) return [{ type: 'text', text: toolJson(shown) }];

  const content: ToolContent = [
    {
      type: 'text',
      text: toolJson(shown.map(mb => simplifyMailbox(mb, { verbose, path: paths.get(mb.id) }))),
    },
  ];
  const note = buildUnpathableMailboxNote(shown.filter(mb => !paths.has(mb.id)).map(mb => mb.id));
  if (note) content.push({ type: 'text', text: note });
  return content;
}

/**
 * create_mailbox. The empty-name and slash-in-name rejections live in the client method,
 * which runs before any round trip.
 */
export async function createMailbox(args: any, client: MailboxClient): Promise<ToolContent> {
  const raw = coerceBool(args?.raw, 'raw') ?? false;
  const verbose = coerceBool(args?.verbose, 'verbose') ?? false;

  const { mailbox, created, path } = await client.createMailbox({
    name: args?.name,
    parent: args?.parent,
  });

  // raw is the server's own Mailbox/set created object, untouched.
  if (raw) return [{ type: 'text', text: toolJson(created) }];

  const content: ToolContent = [
    { type: 'text', text: toolJson(simplifyMailbox(mailbox, { verbose, path })) },
  ];
  const note = path === undefined ? buildUnpathableMailboxNote([mailbox.id]) : null;
  if (note) content.push({ type: 'text', text: note });
  return content;
}

/**
 * delete_mailbox. Every refusal lives in the client method. No raw or verbose: the only
 * mailbox returned is the echo, which is always the untransformed JMAP object.
 */
export async function deleteMailbox(args: any, client: MailboxClient): Promise<ToolContent> {
  const { deleted, deletedMailbox } = await client.deleteMailbox(args?.mailbox);

  const content: ToolContent = [{ type: 'text', text: toolJson({ deleted, deletedMailbox }) }];
  if (deletedMailbox === undefined) {
    // Not a throw: that would report a failure for a completed delete.
    content.push({
      type: 'text',
      text:
        `WARNING: mailbox "${describeUntrusted(deleted)}" was deleted, but the server did not return it alongside the ` +
        'destroy, so `deletedMailbox` is absent and its name and parent cannot be read back here.',
    });
  }
  return content;
}
