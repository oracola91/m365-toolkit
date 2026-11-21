# What you can do

[Back to the README](../README.md)

Catalogued actions cover the common jobs, grouped by category:

- **Exchange Online.** Check access (who has access to a mailbox or calendar, who can Send As,
  which mail groups or shared mailboxes a user has, who is in a distribution group, where a
  mailbox forwards mail, and a leaver check of everything a user can still access); Give access
  (mailbox, Send As, calendar, distribution group); Remove and change access (mailbox and
  calendar); Mailboxes (look up, change a setting, create or convert shared mailboxes).
- **Entra / Identity (Graph).** Look up a user; see which groups a user is directly in, or all
  groups including nested ones.
- **SharePoint Admin (SPO).** List all sites. **Intune / Devices (Graph).** List enrolled
  devices. **Power Platform (admin).** List environments.

**Beyond the catalogue — full capability:** every category menu has a
**"Run any command by name (advanced)"** option. Type any cmdlet in the connected module
and the engine guides its parameters exactly the same way. So the tool is not limited to
the catalogued actions — it can drive *any* cmdlet the connected module exposes.

**Run once or in bulk:** after choosing an action that changes something (or a run-any
cmdlet) you pick **Run once** (type the values now) or **Bulk from a CSV file** (run it for every row;
UTF-8 or Excel's ANSI encoding, `,` or `;` separated).

**Results / export:** when a command returns objects (e.g. a lookup), the tool offers to
show them as a table, show only columns you choose, or **export to CSV** (all columns or a
chosen subset). CSVs default to the tool's `exports` folder (next to the script). They are
owner-only, formula cells are written as text, and the tool deletes its own exports after 30
days; a path you type elsewhere is not cleaned up.

**Convenience:** the tool remembers your email (UPN) and last-used service between runs
(stored in `config.json`, never committed); after connecting it shows **who you're signed
in as** (and your Graph scopes / Exchange org); every category menu has **Search actions by
keyword**; and for identity fields you can type **`find`** to search the directory and pick a
person or mailbox from a list instead of typing the address.
