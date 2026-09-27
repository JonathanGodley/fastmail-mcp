import { InvalidInputError, describeUntrusted } from './coerce.js';

// The per-entry algebra shared by the contact READ shape (src/response-formatters.ts) and
// the update_contact MERGE (src/contacts-calendar.ts), in its own module so the formatter
// need not import the JMAP client to agree with it on what a label is.

export type EntryKeyField = 'address' | 'number';

/**
 * The label of one emails/phones entry, or undefined when it has none. Measured on real cards:
 *
 *  - The entry map's KEY is an opaque server-assigned id, never a label.
 *  - `contexts` is a SET (`{"private": true}`); recent Fastmail UI cards carry it and no `label`.
 *  - `label` is a scalar on older imported cards, observed always `""`, so an EMPTY label
 *    means "no label".
 *
 * The scalar wins when it says something; a `contexts` set with several keys names no label.
 */
export function resolveEntryLabel(entry: any): string | undefined {
  if (!entry || typeof entry !== 'object') return undefined;
  if (typeof entry.label === 'string' && entry.label.trim() !== '') return entry.label;
  const contexts = entry.contexts;
  if (contexts && typeof contexts === 'object' && !Array.isArray(contexts)) {
    const keys = Object.keys(contexts);
    if (keys.length === 1) return keys[0];
  }
  return undefined;
}

/**
 * The default read shape of an emails/phones map: a HYBRID list. An unlabelled entry emits
 * as a BARE STRING, the common case, to save tokens; a labelled one as `{address, label}`.
 * An entry with no value is skipped, still reachable through `verbose` and `raw`; the merge
 * keeps such an entry for the same reason (see `hasEntryValue`).
 */
export function simplifyEntryMap(
  map: any,
  keyField: EntryKeyField,
): Array<string | Record<string, string>> | undefined {
  if (!map || typeof map !== 'object') return undefined;
  const out: Array<string | Record<string, string>> = [];
  for (const entry of Object.values(map)) {
    if (!hasEntryValue(entry, keyField)) continue;
    const value = (entry as any)[keyField];
    const label = resolveEntryLabel(entry);
    out.push(label ? { [keyField]: value, label } : value);
  }
  return out.length ? out : undefined;
}

/**
 * Whether the default read shape shows this entry. One the view hides cannot be named in a
 * resent list, so `mergeEntryMap` carries it over rather than reading its absence as a drop.
 */
function hasEntryValue(entry: any, keyField: EntryKeyField): boolean {
  const value = entry?.[keyField];
  return typeof value === 'string' && value !== '';
}

export interface ContactEntryInput {
  address?: string;
  number?: string;
  label?: string;
}

export interface ContactNameInput {
  given?: string;
  surname?: string;
  full?: string;
}

/**
 * Merge a supplied name into the stored one rather than replacing it wholesale: several real
 * cards carry `components` with NO `full`, so a whole-value replace would delete the only
 * structured given/surname they had. Components of other kinds and every other property of
 * the name object are carried through.
 */
export function mergeContactName(existing: any, incoming: ContactNameInput): Record<string, any> {
  const base: Record<string, any> =
    existing && typeof existing === 'object' && !Array.isArray(existing) ? { ...existing } : {};

  // Copy each component object as well as the array: the merge must not mutate the card the
  // caller is about to be handed back as `previousCard`.
  const components: Array<Record<string, any>> = Array.isArray(base.components)
    ? base.components.map((c: any) => (c && typeof c === 'object' ? { ...c } : c))
    : [];

  const setComponent = (kind: string, value: string | undefined) => {
    if (value === undefined) return;
    const found = components.find((c) => c && c.kind === kind);
    if (found) found.value = value;
    else components.push({ kind, value });
  };
  setComponent('given', incoming.given);
  setComponent('surname', incoming.surname);

  if (incoming.full !== undefined) base.full = incoming.full;
  if (components.length) base.components = components;
  return base;
}

/**
 * Merge a supplied `notes` string into the stored Id-map of notes. One stored note keeps its
 * key and other properties. SEVERAL are REJECTED: a scalar cannot say "replace the second of
 * three", so writing one would silently delete the others.
 */
export function mergeContactNotes(existing: any, note: string): Record<string, any> {
  if (existing && typeof existing === 'object' && !Array.isArray(existing)) {
    const keys = Object.keys(existing);
    if (keys.length > 1) {
      throw new InvalidInputError(
        `This contact stores ${keys.length} separate notes, and \`notes\` sets a single one — writing it ` +
          `would delete the other ${keys.length - 1}. Read them with get_contact (verbose or raw). If ` +
          `replacing all of them with one note is what you want, do it in two deliberate steps: ` +
          `clearFields:['notes'] first, then set notes.`,
      );
    }
    if (keys.length === 1) {
      const key = keys[0];
      const current = existing[key];
      const carried = current && typeof current === 'object' && !Array.isArray(current) ? current : {};
      return { [key]: { ...carried, note } };
    }
  }
  return { n0: { note } };
}

export function buildEntryMap(items: Array<Record<string, any>>, prefix: string): Record<string, any> {
  const map: Record<string, any> = {};
  items.forEach((item, i) => {
    map[`${prefix}${i}`] = item;
  });
  return map;
}

export interface EntryMergeOutcome {
  map: Record<string, any>;
  dropped: Array<{ key: string; entry: any }>;
  /** The key values of supplied entries that matched nothing on the card. */
  added: string[];
}

/**
 * Merge a supplied emails/phones array into the stored map. The array defines WHICH entries
 * exist; a matching entry keeps its map key and every field the read shape never surfaced
 * (`contexts`, `pref`, `@type`, anything future).
 *
 * Matching is EXACT, so a case-differing address is a drop-plus-add, which the ambiguity
 * guard rejects rather than resolving on a guess.
 */
export function mergeEntryMap(
  existingMap: any,
  incoming: ContactEntryInput[],
  keyField: EntryKeyField,
): EntryMergeOutcome {
  const existing: Record<string, any> =
    existingMap && typeof existingMap === 'object' && !Array.isArray(existingMap) ? existingMap : {};
  const existingKeys = Object.keys(existing);
  const prefix = keyField === 'address' ? 'e' : 'p';

  const matched = new Set<string>();
  const map: Record<string, any> = {};
  const added: string[] = [];

  let counter = 0;
  const nextKey = () => {
    let key: string;
    do {
      key = `${prefix}${counter++}`;
    } while (key in existing || key in map);
    return key;
  };

  for (const item of incoming) {
    const value = item[keyField] as string;
    const matchKey = existingKeys.find((k) => !matched.has(k) && existing[k]?.[keyField] === value);
    if (matchKey !== undefined) {
      matched.add(matchKey);
      const stored = existing[matchKey];
      const merged: Record<string, any> =
        stored && typeof stored === 'object' && !Array.isArray(stored) ? { ...stored } : {};
      merged[keyField] = value;
      // Written only when it CHANGES what the read shape showed: a round-tripped label usually
      // came from `contexts`, and writing a scalar `label` would mutate an unchanged card.
      // A real relabel writes the scalar, which then disagrees with `contexts` for other
      // clients; a deliberate residual, stated in the tool description.
      if (item.label !== undefined && item.label !== resolveEntryLabel(stored)) {
        merged.label = item.label;
      }
      map[matchKey] = merged;
    } else {
      added.push(value);
      // A fresh literal, never a passthrough of the caller's object.
      const fresh: Record<string, any> = { [keyField]: value };
      if (item.label !== undefined) fresh.label = item.label;
      map[nextKey()] = fresh;
    }
  }

  const dropped: Array<{ key: string; entry: any }> = [];
  for (const k of existingKeys) {
    if (matched.has(k)) continue;
    if (hasEntryValue(existing[k], keyField)) dropped.push({ key: k, entry: existing[k] });
    else map[k] = existing[k];
  }
  return { map, dropped, added };
}

// Past any real card's entry count, so the echo is whole in practice; the bound only keeps a
// pathological card from producing an unbounded error message.
const MAX_ECHOED_DROPPED_ENTRIES = 50;

/**
 * The card-level `kind` (RFC 9553 section 2.1.4), read ONLY here so the read surface and the
 * write refusals agree on what a group is (#113). What Cyrus does with it:
 *
 *  - ALWAYS present: `jscard_from_vcard` seeds `kind: "individual"` (`imap/jscontact.c:1982`),
 *    so the default value, not absence, marks an ordinary card.
 *  - LOWERCASED on the way out (`imap/jscontact.c:1169`), which is why the group test compares
 *    exactly, as Cyrus does on write (`imap/jmap_contact.c:4282`).
 *  - Any other value passes through verbatim (`org`, `location`, extensions), so this is not a
 *    two-way group-or-not flag.
 */
export function contactCardKind(card: any): string | undefined {
  const kind = card?.kind;
  return typeof kind === 'string' && kind !== '' ? kind : undefined;
}

const DEFAULT_CONTACT_KIND = 'individual';

/**
 * The declared kind, unless it is the default: so a card the write tools may refuse is
 * visible BEFORE the caller spends a write finding out (#113).
 */
export function nonDefaultContactKind(card: any): string | undefined {
  const kind = contactCardKind(card);
  return kind === DEFAULT_CONTACT_KIND ? undefined : kind;
}

/**
 * The kind the contact write tools refuse, or undefined for a card they may write.
 * `create_contact` has no `kind` parameter, so it makes individuals only; any other declared
 * kind is a record this server could not put back, and both write tools refuse it through
 * this single rule.
 */
export function refusedContactKind(card: any): string | undefined {
  return nonDefaultContactKind(card);
}

/** The shared refusal both write tools raise, so they read as one rule. */
export function contactKindRefusal(opts: { id: string; kind: string; tool: string; because: string; recovery: string }): string {
  const group = opts.kind === 'group' ? ' (a contact GROUP)' : '';
  return `Contact ${opts.id} is a card of kind "${describeUntrusted(opts.kind)}"${group}, not a person card, ` +
    `so ${opts.tool} refuses it: ${opts.because} ${opts.recovery}`;
}

/**
 * Entries dropped AND added in one call. Split out from the rejection so `update_contact` can
 * scope its override to the ambiguous field without catching an exception.
 */
export function isAmbiguousEntryEdit(outcome: EntryMergeOutcome): boolean {
  return outcome.dropped.length > 0 && outcome.added.length > 0;
}

/**
 * Reject an edit that both drops a known entry and adds an unknown one: it reads as a
 * correction or as a removal plus an unrelated addition, and silent replace is the lossy one.
 *
 * The dropped entries are echoed in FULL, hidden fields included, up to
 * MAX_ECHOED_DROPPED_ENTRIES, so the lossless retry is cheaper than reaching for
 * `allowEntryReplace`.
 */
export function assertUnambiguousEntryEdit(field: string, outcome: EntryMergeOutcome): void {
  if (!isAmbiguousEntryEdit(outcome)) return;

  const shown = outcome.dropped.slice(0, MAX_ECHOED_DROPPED_ENTRIES);
  const more = outcome.dropped.length - shown.length;
  const droppedText = `${shown.map((d) => JSON.stringify(d.entry)).join(', ')}${
    more > 0 ? `, …and ${more} more (read them in full with get_contact with verbose:true)` : ''
  }`;

  throw new InvalidInputError(
    `This ${field} edit both drops ${outcome.dropped.length} existing entry(ies) and adds ` +
      `${outcome.added.length} the card does not have, which is ambiguous: it reads either as a ` +
      `correction to the existing entry or as its removal plus an unrelated addition. ` +
      `Dropped: ${droppedText}. Added: ${outcome.added.map((v) => JSON.stringify(v)).join(', ')}. ` +
      `To EDIT an entry losslessly, resend it under its existing value (shown above) so it matches, ` +
      `changing only what you meant to change. To REPLACE the ${field} outright, pass ` +
      `allowEntryReplace:true — every ${field} entry is then written fresh from what you ` +
      `supplied, so contexts, pref and any other field the simplified output does not show will ` +
      `NOT carry over. The override applies to ${field} alone; any other array in the same call ` +
      `still merges.`,
  );
}
