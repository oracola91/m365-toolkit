# M365 Guided Command Builder

A plain-language, menu-driven PowerShell tool so the IT team can run Microsoft 365
admin commands **without memorizing cmdlet names or syntax**. You pick an action in
human language and the tool walks you through the parameters one at a time, shows you
the exact command it will run (dry-run), and only runs it when you confirm.

## How it works (independent layers, each extended by adding one row)

1. **Connection registry** (`$Services`) — one row per M365 service. On launch you
   enter your email and choose which service to connect to (Exchange Online, Entra/Graph,
   SharePoint/PnP, Teams, Security & Compliance, SharePoint Admin/SPO, Intune/Devices,
   Power Platform). Add a service = add one row.
2. **Introspection engine** (`Invoke-GuidedCmdlet` / `Get-CmdletParameters`) — the clever
   bit. It reads the *real* parameters straight off the cmdlet with
   `(Get-Command X).Parameters`, so it's always accurate and never goes stale. It sorts
   them into **Required → Optional → Switches**, and for each one it works out the allowed
   values from a `[ValidateSet]`, an **enum type**, or the manual `$ParamHints` table, then
   offers them as a numbered pick or lists them inline (type `?` at any prompt to see them).
   Works for *any* cmdlet in *any* connected module — no per-command code.
3. **Plain-language catalogue** (`$Actions`) — the only per-command work: a friendly
   `Label → Cmdlet` mapping, grouped by category, with optional `Help`, per-parameter
   `ParamHelp` (who → whom), and `Preset` values the action fixes for you. Add an
   action = add one row.
4. **Bulk runner** (`Invoke-BulkCmdlet`) — run the *same* action once per row of a CSV.
   You map each parameter to a CSV column or a fixed value. Every row is checked before
   anything runs (same value rules as a single run; identity columns are looked up for
   catalogued actions), blank cells are skipped unless you choose to clear, and the
   targets are listed. Destructive commands need the command name and the row count
   typed. Each row's outcome (and, for a `Set-*`, the values it replaced) is saved to
   `exports\bulk-<cmdlet>-<timestamp>-results.csv`.

## Requirements / platform support

The tool itself is plain PowerShell and runs on **PowerShell 7+ (Windows, macOS, Linux)**
and **Windows PowerShell 5.1**. It only loads a service's module when you pick that
service, so unsupported services fail with a clear message rather than crashing the tool.

Coverage depends on the vendor module behind each service — and two of Microsoft's
modules are **Windows-only** (built on .NET Framework, which macOS/Linux don't have):

| Service | macOS / Linux (pwsh 7) | Windows (pwsh 7) | Windows PowerShell 5.1 |
|---|:--:|:--:|:--:|
| Exchange Online | ✅ | ✅ | ✅ |
| Security & Compliance (Purview) | ✅ | ✅ | ✅ |
| Entra / Identity (Graph) | ✅ | ✅ | ✅ |
| Intune / Devices (Graph) | ✅ | ✅ | ✅ |
| Microsoft Teams | ✅ | ✅ | ✅ |
| SharePoint Online (PnP) | ✅ | ✅ | ❌ (PnP needs pwsh 7.2+) |
| **SharePoint Admin (SPO)** | ❌ Windows only | ✅ | ✅ |
| **Power Platform (admin)** | ❌ Windows only | ✅ | ✅ |

> The two Windows-only services (**SharePoint Admin**, **Power Platform**) are a Microsoft
> platform limitation, not a bug in this tool — Microsoft doesn't ship those modules for
> macOS/Linux. On a Mac/Linux box the tool detects this and tells you to run them from Windows.

## Run it

In PowerShell (pwsh 7 on Windows/macOS/Linux, or Windows PowerShell 5.1) with the relevant modules:

```powershell
cd M365Toolkit
# Safe first run: preview commands, never execute anything
.\Invoke-M365Toolkit.ps1 -DryRunOnly
# Normal run (asks y/n before running each command)
.\Invoke-M365Toolkit.ps1
```

### Run it remotely (once the repo is public)

> **⚠️ Security note.** `irm … | iex` downloads and runs a script directly in your
> admin session — the same trust model as `curl … | bash`. **Read the script first**,
> and pin the URL to a **specific release tag** (not the moving `main` branch) so you
> always run a known, reviewed version. Cloning the repo and running the local file is
> the safer default for an admin tool; use the one-liner only when you understand it.

Pinned to a release tag (recommended — replace `v0.24` with the tag you reviewed):

```powershell
irm https://raw.githubusercontent.com/oracola91/m365-toolkit/v0.24/Invoke-M365Toolkit.ps1 | iex
```

For the safe dry-run preview mode over the network:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/oracola91/m365-toolkit/v0.24/Invoke-M365Toolkit.ps1))) -DryRunOnly
```

Run this way there is no local file, so logs and CSV exports go to `%USERPROFILE%\M365Toolkit\`.
Using a tag (not `main`) also avoids the CDN-cache lag on fresh pushes. (The repo is
private for now — this works once it's switched to public.)

If a required module isn't installed, the tool offers to install it for the current user.

- Uses **your signed-in admin's RBAC role** — a helpdesk member can only do what their
  role permits.
- Every session is transcript-logged to `.\logs\session-<timestamp>.log` for audit.
- At any parameter prompt, type **`?`** to list the allowed values. In the category and
  action menus, type **`?`** for a short explanation of each choice.

## What you can do

Catalogued actions cover the common jobs, grouped by category:

- **Mailbox access** — grant/remove access to a mailbox, Send As.
- **Shared mailboxes** — create a shared mailbox, convert an existing mailbox to shared,
  give Full Access with control over auto-mapping.
- **Mailboxes** — look up a mailbox, change a mailbox setting.
- **Groups** — add a member to a distribution group.
- **Users** (Graph) — look up a user.
- **Sites** (SharePoint Admin/SPO) — list all SharePoint sites.
- **Devices** (Intune/Graph) — list enrolled devices.
- **Environments** (Power Platform) — list environments.

**Beyond the catalogue — full capability:** every category menu has a
**"Run any command by name (advanced)"** option. Type any cmdlet in the connected module
and the engine guides its parameters exactly the same way. So the tool is not limited to
the catalogued actions — it can drive *any* cmdlet the connected module exposes.

**Run once or in bulk:** after choosing an action (or a run-any cmdlet) you pick
**Run once** (type the values now) or **Bulk from a CSV file** (run it for every row;
UTF-8 or Excel's ANSI encoding, `,` or `;` separated).

**Results / export:** when a command returns objects (e.g. a lookup), the tool offers to
show them as a table, show only columns you choose, or **export to CSV** (all columns or a
chosen subset). CSVs default to `.\exports\<cmdlet>-<timestamp>.csv`.

**Convenience:** the tool remembers your email (UPN) and last-used service between runs
(stored in `config.json`, never committed); after connecting it shows **who you're signed
in as** (and your Graph scopes / Exchange org); every category menu has **Search actions by
keyword**; and for identity fields you can type **`find`** to search the directory and pick a
person or mailbox from a list instead of typing the address.

## Add a new action (the common case)

Append one entry to `$Actions` in `Invoke-M365Toolkit.ps1`:

```powershell
[pscustomobject]@{
    Id        = 'mbx-hide-from-gal'  # required: unique and stable; actions are dispatched by Id, so never change it once shipped
    Service   = 'Exchange Online'
    Category  = 'Mailboxes'
    Label     = 'Hide a mailbox from the address list'
    Cmdlet    = 'Set-Mailbox'
    Prompt    = @('Identity','HiddenFromAddressListsEnabled')  # params to always walk (required OR optional)
    Preset    = @{ }                 # optional: values the action fixes, never prompted (e.g. @{ Type='Shared' })
    Help      = 'One-line description shown when building and via ? in the menu.'
    ParamHelp = @{ Identity = 'The mailbox to change' }        # optional: who -> whom, per parameter
}
```

That's it — the engine discovers and validates the parameters for you.

If a parameter has fixed values but the cmdlet hides them (common with Exchange Online
proxy cmdlets that expose things as `Object[]`), teach the tool by adding a row to
`$ParamHints` (keyed by parameter name), e.g. `AccessRights = @('FullAccess','SendAs',...)`.

## Status / roadmap

- **v0.8 (2026-07-07):** **Calendar access** category — see who has access to a mailbox's
  calendar, and grant / change / remove a role (Reviewer, Editor, AvailabilityOnly, …); the
  mailbox's `\Calendar` folder is targeted automatically, so you only type the person. Merged
  the shared "Give Full Access" action into the general grant (auto-mapping is now an optional
  parameter there). Two reusable engine hooks added: per-action `Transform` (post-process a
  collected value before it runs) and `Hints` (per-action allowed-value override, since
  `AccessRights` means calendar *roles* here but mailbox *rights* elsewhere). Fixed
  `Connect-ExchangeOnline` on macOS / PowerShell 7 with `-SkipLoadingFormatData` (avoids the
  `exchange.format.ps1xml` temp-file error). **UI:** all menus now list **A-Z** (categories,
  actions, services — last-used service still floated to the top); app controls (Search / Run
  any command / Switch service / Quit) render below a `-- app --` divider, visually separated
  from the menu's own items.
- **v0.7 (2026-07-06):** prompt navigation — type `back` to re-enter the previous field or
  `cancel` to abort to the menu; a shared numbered **multi-select** picker (`Read-OptionMenu`)
  for fixed-set fields; and two **mailbox permission view** actions — "See who can Send As a
  mailbox" and "See who has access to a mailbox".
- **v0.6 (2026-07-06):** parameter-set resolution (no more spurious required params like
  `-Owner`); `?` explains each optional/switch parameter (action help → glossary → Get-Help,
  with `-LoadCmdletHelp` on EXO connect); validates values against allowed sets and identity
  fields against the directory (typos/nonsense are flagged, not accepted); post-action
  "What next?" menu (do another / switch service / quit); reports `Failed: <msg>` on error;
  standard white/green/yellow/red colors; SendAs handled via the correct action.
- **v0.5 (2026-07-06):** remembers UPN + last-used service; "who am I" preflight after
  connect; **Search actions by keyword**; `find` directory lookup for identity fields.
- **v0.4 (2026-07-06):** three more services wired — **SharePoint Admin (SPO)**,
  **Intune / Devices (Graph)**, **Power Platform (admin)** — each with a starter list action.
  (Runs from this Mac copy; no VM involved.)
- **v0.3 (2026-07-06):** result handling — after a command returns objects you can show a
  table, choose columns, or export to CSV (all or chosen columns) to `.\exports\`.
- **v0.2 (2026-07-03):** enum + `$ParamHints` value discovery with `?` to list options;
  per-parameter plain-language help (who → whom); `?` help in the menus; `Preset` values;
  **Shared mailboxes** category; **run-any-cmdlet** escape hatch; **CSV bulk mode**.
- **v0.1 (2026-07-03):** connection registry, introspection wizard, dry-run preview,
  transcript logging, a small starter catalogue.
- **Next / still missing:**
  - Legacy AzureAD/MSOnline are retired (not wired, and not worth adding — migrate to Graph).
  - No piping one action's output into the next.
  - Interactive auth only — no app-only/certificate auth for unattended/scheduled runs.
  - No undo, and destructive cmdlets get the same y/n prompt as a lookup.
  - Optional plain-English → cmdlet layer (via the Claude API) so you can type a request
    instead of navigating menus.

Design notes live in the second brain: `second-brain/Knowledge/Office 365 PowerShell reference.md`.
