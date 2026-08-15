import { InvalidInputError } from './coerce.js';

// The per-entry algebra shared by the contact READ shape (src/response-formatters.ts) and
// the update_contact MERGE (src/contacts-calendar.ts). It lives in its own module so the
// formatter never has to import the JMAP client to agree with it on what a label is: both
// sides read one definition, so the shape a caller sees and the shape a write preserves
// cannot drift apart.
//
// Everything here is pure — no account, no network — so the rules below are exercised
// directly by unit tests rather than only through a live card.

/** The property that identifies an entry within its map: emails by address, phones by number. */
export type EntryKeyField = 'address' | 'number';

/**
 * The label of one emails/phones entry, or undefined when it has none.
 *
 * Read off real cards in a live address book, because the shape is not what the map
 * suggests:
 *
 *  - The entry map's KEY is an opaque server-assigned id, never a label. Older cards carry
 *    40-character sha1-shaped keys; cards written by the current Fastmail UI carry short
 *    6-character ones. Neither is readable, so the key is never used as a label.
 *  - `contexts` is a SET of context names — `{"private": true}` — and every card written by
 *    a recent Fastmail UI carries it and carries no `label` key at all.
 *  - `label` is a scalar string, present on older imported cards, and on every one of those
 *    observed its value was `""`. So an EMPTY label means "no label", not "a label that
 *    happens to be blank"; treating it as a label would emit `{address, label: ""}` objects
 *    for a whole imported address book.
 *
 * Both properties are live in the same account, so both are read here: the scalar wins when
 * it says something, and `contexts` is the fallback. A `contexts` set with more than one key
 * names no single label, so it resolves to none rather than picking one arbitrarily.
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
 * The default read shape of an emails/phones map: a HYBRID list.
 *
 * An entry with no label emits as a BARE STRING (the address / the number) — which is what
 * the overwhelming majority of real entries are, and wrapping every one of them in a
 * single-key object costs tokens to say nothing. An entry that does have a label emits as
 * `{address, label}` / `{number, label}`, because dropping the label would make "home" and
 * "work" indistinguishable in the output.
 *
 * Entries with no address/number at all are skipped: there is no value to emit, and the
 * full entry is still reachable through `verbose` and `raw`.
 */
export function simplifyEntryMap(
  map: any,
  keyField: EntryKeyField,
): Array<string | Record<string, string>> | undefined {
  if (!map || typeof map !== 'object') return undefined;
  const out: Array<string | Record<string, string>> = [];
  for (const entry of Object.values(map)) {
    const value = (entry as any)?.[keyField];
    if (typeof value !== 'string' || value === '') continue;
    const label = resolveEntryLabel(entry);
    out.push(label ? { [keyField]: value, label } : value);
  }
  return out.length ? out : undefined;
}

/** A coerced emails/phones entry as it arrives from a tool call. */
export interface ContactEntryInput {
  address?: string;
  number?: string;
  label?: string;
}

/** A coerced structured name as it arrives from a tool call. */
export interface ContactNameInput {
  given?: string;
  surname?: string;
  full?: string;
}

/**
 * Merge a supplied name into the stored one rather than replacing it wholesale.
 *
 * Real cards carry `{ full?, components?: [{kind, value}] }`, and several carry `components`
 * with NO `full` — so a whole-value replace driven by a bare name string would delete the
 * only structured given/surname the card had. Instead:
 *
 *  - a supplied `full` sets `full` and leaves the components alone;
 *  - a supplied `given`/`surname` updates the component of that kind in place, or appends
 *    one when the card had none, leaving components of every other kind (middle names,
 *    titles, suffixes) untouched;
 *  - every other property of the stored name object (`@type`, `sortAs`, `isOrdered`, …) is
 *    carried through unchanged.
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
 * Merge a supplied `notes` string into the stored notes map.
 *
 * A card's notes are an Id-map of note objects, but the tool surface is a single string. When
 * the card holds exactly one note the merge keeps its map key and any other properties it
 * carries and overwrites only the text, for the same reason the entry merge preserves
 * `contexts`/`pref`. When the card holds none, a single fresh note is written.
 *
 * When the card holds SEVERAL, the write is REJECTED. There is no way to express "replace the
 * second of three" through a scalar parameter, so writing one note would delete the others —
 * silent loss of a field the caller was never shown, which is the exact failure the rest of
 * this module exists to prevent. Same posture as the entry-edit ambiguity guard: refuse and
 * name the deliberate route, rather than resolve it on a heuristic.
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

/** Build an entry map from scratch, with no reference to what the card already held. */
export function buildEntryMap(items: Array<Record<string, any>>, prefix: string): Record<string, any> {
  const map: Record<string, any> = {};
  items.forEach((item, i) => {
    map[`${prefix}${i}`] = item;
  });
  return map;
}

export interface EntryMergeOutcome {
  /** The Id-map to write, existing map keys preserved wherever an entry matched. */
  map: Record<string, any>;
  /** Existing entries no supplied entry matched, with the map keys they were stored under. */
  dropped: Array<{ key: string; entry: any }>;
  /** The key values of supplied entries that matched nothing on the card. */
  added: string[];
}

/**
 * Merge a supplied emails/phones array into the stored map.
 *
 * The supplied array defines WHICH entries exist; the stored map defines what each surviving
 * entry still carries. An entry whose address/number matches one already on the card keeps
 * every field the read shape never surfaced — `contexts`, the near-ubiquitous `pref`,
 * `@type`, and anything a future Fastmail release adds — and only the supplied properties
 * are written over it. Its map key is kept too, so a client holding an entry id still
 * resolves it.
 *
 * Matching is EXACT on the key value. A case-differing address does not match, and therefore
 * surfaces as a drop-plus-add, which the ambiguity guard below rejects rather than resolving
 * on a guess.
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
      // A label is only written when it CHANGES what the entry says. A caller that read the
      // card and sent it straight back is resending the label this server showed it, and on
      // the common shape that label came from `contexts` — the entry has no scalar `label`
      // key at all. Writing one anyway would mutate the card on a round-trip that asked for
      // no change, adding a property the entry never had. So an incoming label equal to the
      // one the read shape resolved is treated as "unchanged" and nothing is written.
      //
      // A label that genuinely differs still writes the scalar `label`, which then wins over
      // `contexts` on the next read — leaving the two properties disagreeing on a card whose
      // other clients read `contexts`. That residual is deliberate and NOT resolved here;
      // whether a relabel should write `contexts` instead is an open behaviour question, and
      // the tool description states what actually happens rather than promising more.
      if (item.label !== undefined && item.label !== resolveEntryLabel(stored)) {
        merged.label = item.label;
      }
      map[matchKey] = merged;
    } else {
      added.push(value);
      // A fresh literal built from the validated keys, never a passthrough of the caller's
      // object — the same discipline the coercion layer follows, for the same reason.
      const fresh: Record<string, any> = { [keyField]: value };
      if (item.label !== undefined) fresh.label = item.label;
      map[nextKey()] = fresh;
    }
  }

  const dropped = existingKeys.filter((k) => !matched.has(k)).map((k) => ({ key: k, entry: existing[k] }));
  return { map, dropped, added };
}

// How many dropped entries are spelled out in the ambiguity rejection. Enough to retry
// losslessly on any realistic card, short enough that a pathological one does not become the
// error message.
const MAX_ECHOED_DROPPED_ENTRIES = 5;

/**
 * Whether a card is a contact GROUP rather than a person card.
 *
 * A group card carries a `members` map of the uids it contains and none of the person fields.
 * This server's contact surface has no `kind` parameter and no `members` parameter, so it can
 * neither create a group nor describe one — which is why both write tools that can meet one
 * refuse it, from a single rule stated in one place rather than two lookalike checks.
 */
export function isContactGroupCard(card: any): boolean {
  return card?.kind === 'group';
}

/**
 * The shared refusal both group-aware write tools raise, so they read as one rule.
 *
 * `because` carries the part that genuinely differs — an update has no parameters that
 * describe a group, a delete cannot put back what it destroys — and `recovery` names where
 * the caller can do it instead. Everything a caller needs to recognise the rule (this is a
 * group; this tool will not touch it) is fixed here.
 */
export function contactGroupRefusal(opts: { id: string; tool: string; because: string; recovery: string }): string {
  return `Contact ${opts.id} is a contact GROUP, not a person card, so ${opts.tool} refuses it: ` +
    `${opts.because} ${opts.recovery}`;
}

/**
 * Whether this field's merge is the one shape that cannot be resolved: entries dropped AND
 * entries added in the same call.
 *
 * Split out from the rejection below so a caller can ask the question without catching an
 * exception — `update_contact` needs it to scope its override to the field that was
 * actually ambiguous instead of applying it to every array in the call.
 */
export function isAmbiguousEntryEdit(outcome: EntryMergeOutcome): boolean {
  return outcome.dropped.length > 0 && outcome.added.length > 0;
}

/**
 * Reject an edit that both drops a known entry and adds an unknown one in the same call.
 *
 * That combination has two readings that produce different cards — "correct the address on
 * this entry, keeping its contexts and pref" and "delete this entry and add an unrelated
 * one" — and the tool cannot tell which was meant. Resolving it silently as a replace is the
 * lossy reading, so it is refused instead.
 *
 * The refusal echoes the dropped entries in FULL, hidden fields included, because the point
 * is to make the lossless retry cheaper than reaching for the override: with `pref` and
 * `contexts` in hand, the caller can resend the entry it meant to keep. A rejection that only
 * named the entries would leave `allowEntryReplace` as the path of least resistance.
 */
export function assertUnambiguousEntryEdit(field: string, outcome: EntryMergeOutcome): void {
  if (!isAmbiguousEntryEdit(outcome)) return;

  const shown = outcome.dropped.slice(0, MAX_ECHOED_DROPPED_ENTRIES);
  const more = outcome.dropped.length - shown.length;
  const droppedText = `${shown.map((d) => JSON.stringify(d.entry)).join(', ')}${
    more > 0 ? `, …and ${more} more` : ''
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
