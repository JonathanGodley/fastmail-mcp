import { buildMailboxPathMap, filterMailboxesByParent } from './jmap-client.js';
import { simplifyMailbox, buildUnpathableMailboxNote } from './response-formatters.js';
import { coerceBool, toolJson } from './coerce.js';

/**
 * The slice of the JMAP client the mailbox tools need, so both handlers can be exercised
 * with a stub instead of a live account. `JmapClient` satisfies it structurally.
 */
export interface MailboxClient {
  getMailboxes(): Promise<any[]>;
  createMailbox(input: { name: string; parent?: string }): Promise<{ mailbox: any; created: any; path?: string }>;
}

export type ToolContent = Array<{ type: 'text'; text: string }>;

/**
 * list_mailboxes. Lives here rather than inline in the CallTool switch because it does
 * real work — builds the path map, applies the parent narrowing, threads a per-mailbox
 * `path` into the formatter, branches on raw, and conditionally emits a second content
 * item — and logic left in the switch can only ever be checked by running the server.
 *
 * The first content item is ALWAYS the JSON array and nothing else, in either mode: a
 * caller parses it directly, so the no-path note rides as a separate item and is never
 * concatenated onto the JSON string.
 */
export async function listMailboxes(args: any, client: MailboxClient): Promise<ToolContent> {
  // coerceBool, not `!!`: a lenient client's stringified "false" is truthy, and under `!!`
  // raw:"false" would return untransformed JMAP to a caller that explicitly asked for the
  // simplified shape. Both default to false. (#54)
  const raw = coerceBool(args?.raw) ?? false;
  const verbose = coerceBool(args?.verbose) ?? false;

  // `path` is root-anchored, so it needs the WHOLE tree even when the listing is narrowed
  // to one parent's children. Fetch unnarrowed, build the path map, then apply the pure
  // filter to the list already in hand — one Mailbox/get either way, rather than a second
  // round trip just for the ancestors.
  const mailboxes = await client.getMailboxes();
  const { paths } = buildMailboxPathMap(mailboxes);
  const shown = filterMailboxesByParent(mailboxes, args?.parent);

  // raw stays untransformed JMAP, which carries no path — and therefore no note about a
  // path being absent, since none was promised.
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
 * which runs before any round trip and inside this same harness.
 *
 * A created mailbox whose parent has no computable path gets no `path` field, and says so
 * in a trailing note rather than letting the promised field vanish — the same discipline
 * the listing follows.
 */
export async function createMailbox(args: any, client: MailboxClient): Promise<ToolContent> {
  // Same coercion as listMailboxes, for the same reason.
  const raw = coerceBool(args?.raw) ?? false;
  const verbose = coerceBool(args?.verbose) ?? false;

  const { mailbox, created, path } = await client.createMailbox({
    name: args?.name,
    parent: args?.parent,
  });

  // raw returns the server's own Mailbox/set created object, untouched — the same meaning
  // raw carries on every other tool here.
  if (raw) return [{ type: 'text', text: toolJson(created) }];

  const content: ToolContent = [
    { type: 'text', text: toolJson(simplifyMailbox(mailbox, { verbose, path })) },
  ];
  const note = path === undefined ? buildUnpathableMailboxNote([mailbox.id]) : null;
  if (note) content.push({ type: 'text', text: note });
  return content;
}
