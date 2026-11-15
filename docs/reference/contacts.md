# Contacts reference

## Contact fields

**Default**: `id`, `name`, `emails`, `phones`, `organization`, `notes`, and `kind` when it is not `individual`

**Verbose adds**: `addresses`, `anniversaries`, every entry of `organizations`, `titles` and `notes` whole and keyed, plus every remaining JMAP field under its own JSContact name and in its stored shape: on Fastmail that includes `links` (URLs), `onlineServices`, `media` (photos), `nicknames` and `kind` on an ordinary card — *and* it widens `emails`/`phones` from the hybrid shape below to the whole stored entry objects, so it adds detail to fields the default already returns, not just extra fields.

**Simplification applied:**
- Name resolved from `name.full`, or when that is absent from the name's `components`: in stored order when the name is marked ordered, otherwise title, given, middle, surname, second surname, generation, credential
- Emails and phones flattened from JMAP's Id-map to an array — see the hybrid shape below
- Default view: `organization` is the first organization's name, and `notes` is the text of every note joined with a newline
- `kind` omitted when it is the default `individual` - see below
- Verbose: addresses as objects, anniversaries as stored; `organizations`, `titles` and `notes` are the card's own Id-maps, keys kept, so a title's `organizationId` refers to a key of `organizations`; `organization` is still emitted; an `online` key, which RFC 9553 does not define, is copied whole when present; a value of any of these four keys that is not a map is passed through as stored, except null or an empty string, which are omitted; the other passed-through fields keep JMAP's Id-map shape
- Verbose `name` is the same display string as the default; the structured name is in `raw` only

### `kind` tells you which cards the write tools will refuse

A JSContact card carries a `kind` ([RFC 9553](https://www.rfc-editor.org/rfc/rfc9553.html) §2.1.4), and the server always emits one - a card whose stored vCard has no `KIND` line still comes back as `"individual"` (`docs/conventions.md` cites the server source for this). **The default view shows `kind` only when it is something else**:

```json
[{"id": "C1", "name": "Ada Lovelace"},
 {"id": "G1", "kind": "group", "name": "Book club"}]
```

No `kind` field means an individual. Any `kind` the default view shows is one `update_contact` and `delete_contact` both refuse (see [Writing contacts](#writing-contacts)) - so a listing shows you that before you spend a call finding out. `kind: "group"` is a **contact group**; RFC 9553 also defines `org`, `location`, `device` and `application`, and the server passes an unregistered value through as-is. None of those kinds is a person, and `create_contact` cannot produce any of them.

`verbose: true` and `raw: true` pass `kind` through untouched, `individual` included - so under either one the presence of the field says nothing about what the card is, and you read its value. Should a card ever arrive without a `kind` at all, the default view reads that as an individual too ([#113](https://github.com/JonathanGodley/fastmail-mcp/issues/113)).

### `emails` and `phones` are a hybrid array — handle both shapes

Each entry is **either a bare string** (the address / the number) when it carries no label, **or an object** when it does:

```json
"emails": ["plain@example.com", {"address": "work@example.com", "label": "work"}],
"phones": [{"number": "+1 555 0100", "label": "mobile"}, "+1 555 0199"]
```

A label is resolved from the entry itself, never from its map key (the keys are opaque server ids — 40-character hashes on older cards, short 6-character ids on newer ones). Two properties can carry one, and both are in live use in the same address book: a scalar `label` string wins when it is non-empty, otherwise a single-key `contexts` set (`{"private": true}`) supplies the name. An **empty** `label` — which every entry on some older imported cards carries — counts as no label at all. A `contexts` set naming more than one context resolves to no label rather than picking one.

Everything the hybrid shape folds away (`contexts`, `pref`, `@type`, …) is still reachable with `verbose: true` or `raw: true`.

## Writing contacts

`create_contact` and `update_contact` take the same entry arrays (`delete_contact` takes only `contactId`), and accept both shapes on input too: `emails` may be `["a@b.example"]` or `[{address, label}]`, `phones` may be `["+1 555 0100"]` or `[{number, label}]`. `addresses` take `{full, label?}` objects only. Any array may also arrive as a JSON string. An unknown per-item key, a key of the wrong type, or a repeated value is rejected naming its position (e.g. `emails[2]`).

**`update_contact` merges per entry rather than overwriting the card.** A JMAP patch replaces a top-level property outright, so writing `emails` from a flat array alone would silently discard the `contexts` and `pref` that sit on nearly every real entry. Instead the tool reads the card first, and:

- **The array you pass decides which entries exist.** An entry whose address/number matches one already stored keeps everything the simplified output does not show, under its existing entry id; only the properties you supply are written over it. Resending an entry exactly as you read it writes nothing at all. Two kinds of stored entry you could not have named are kept as they are: one with an empty address/number, which the default view does not show, and a second entry repeating an address/number you sent, since the input may not repeat a value. `allowEntryReplace` does not remove such an entry unless the edit of that array is ambiguous; to remove one, clear the array with `clearFields` and then write it again.
- **Labels are add-and-override, not a clean rewrite.** A label that differs from the one you read is written as the entry's scalar `label` property, which is what this server resolves first — but the label you read may have come from a `contexts` set, and that set is left exactly as it was. The two can then disagree outside this server, and a Fastmail client may keep showing the old context. There is no way to *remove* a label here: a labelled entry can be relabelled, not unlabelled. Writing `contexts` the way the Fastmail apps do is not implemented.
- **A call that both drops a stored entry and adds an unknown one is rejected as ambiguous** — it reads equally as "correct this entry" and as "delete it and add an unrelated one", which produce different cards. The rejection prints the dropped entries **in full**, hidden fields included, so resending them losslessly is the cheap path. Past 50 dropped entries it prints the first 50 and says how many more there are; read those with `get_contact` and `verbose: true`. `allowEntryReplace: true` proceeds anyway, and is **scoped to the array that was actually ambiguous**: every entry of that one array is rewritten from what you supplied and **not** carrying those hidden fields over, while another array in the same call that merged cleanly still merges.
- **`name` merges into the stored structured name.** A bare string sets the full name and keeps the given/surname components; `{given}`/`{surname}` update just that part, leaving other components (titles, middle names) alone.
- **`addresses` do NOT merge.** A postal entry has no matchable key, so supplying them replaces the whole set.
- **An empty array is rejected**, not read as "clear", and so is a blank string (`emails: ""`). Use `clearFields: ['emails']`, which can only be written on purpose. Clearable: `emails`, `phones`, `addresses`, `notes`. **`name` is not clearable** — it is how the contact is identified in every listing, so delete the card and create a new one instead. Passing a field as a value *and* in `clearFields` in the same call is rejected. An empty or whitespace-only `notes` string is rejected too. `create_contact` rejects both, except that there the fix is to omit the field — a contact being created has nothing to clear. `update_contact` also rejects a call that names no field and no `clearFields` at all.
- **A card storing more than one note is rejected** rather than collapsed. `notes` writes a single note, so a multi-note card would lose the others silently; the rejection names the count and leaves the card alone.
- **A card of any kind but individual cannot be updated** by this tool: its parameters describe a person card, and a group's `members` are not editable here.

**Nor can such a card be deleted here.** `create_contact` has no `kind` and no `members` parameter, so this server can make only individual cards and will not destroy one it could not put back. `delete_contact` rejects a group, org, location, device or application card (or any other non-individual kind) and points you at the Fastmail web interface. The rule is about the *kind* of record, not about fields the create tool cannot set: cards carrying titles, organizations or photos this server cannot write still delete normally.

**You can see which cards those are before you write to them.** `list_contacts`, `get_contact` and `search_contacts` show `kind` on every card that is not an individual - see [`kind` tells you which cards the write tools will refuse](#kind-tells-you-which-cards-the-write-tools-will-refuse).

**Both writes echo the card as it stood before them.** `update_contact` returns `previousCard` and `delete_contact` returns `deletedCard`, the full untransformed JMAP card. That is what makes a wrong write legible: a contact delete is irreversible (a card does not go to Trash the way an email does), so the echo is the only copy left, and an update made from a stale copy is visible in the response. Neither tool has a confirmation parameter.

**The echo is not an undo.** This server writes a name, emails, phones, addresses and a note, and nothing else — so `create_contact` can rebuild that much of a deleted card, but **not** its photos, titles, organizations, nicknames, URLs, anniversaries, group membership, `uid`, or the per-entry `contexts`/`pref` that the merge exists to protect in the first place. The card you get back is a similar contact, not the one that was there. Keep the echo if there is any chance the write was wrong, and restore the rest in a Fastmail client.

**Both writes refuse a card that changed after they read it.** Each tool reads the card first (to merge, or to check its kind) and sends the write only if the address book's state is still the one it read; otherwise the call fails with "changed since it was read" and nothing is written or deleted. Fastmail keeps that state per account rather than per card, so a change to any other contact in the same moment also causes the refusal. Retrying the same call re-reads the card and succeeds. If the server reports no state with the read, the write is refused rather than sent unguarded.

`update_contact` also returns the merged card as `contact`, read back in the same request as the write. On the rare occasion the server does not return it, `contact` is absent and the response says so rather than letting the field vanish.

`deletedCard` degrades the same way, and that is the worst case of the two: `previousCard` is read before the update and the update does not proceed without it, but the delete's copy is read alongside the destroy, so on the rare occasion the server does not return it the card is gone and no copy came back. The response carries an explicit warning saying exactly that instead of returning `{deleted}` on its own. Nothing can reconstruct the card at that point.
