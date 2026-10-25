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
# Preview mode: never runs a change. It still signs in, reads the tenant for lookups,
# and can install a missing module, so it is not offline.
.\Invoke-M365Toolkit.ps1 -DryRunOnly
# Normal run (asks y/n before running each command)
.\Invoke-M365Toolkit.ps1
# Only if the browser sign-in for Entra / Intune hangs (seen on macOS): device code instead
.\Invoke-M365Toolkit.ps1 -UseDeviceCode
# Also keep a full screen transcript of the session (off by default: it holds every result you view)
.\Invoke-M365Toolkit.ps1 -Transcript
```

After each sign-in the tool shows the account the session is really signed in as, and
stops if it differs from the email you typed. On Quit it signs out of every service it
connected to.

### Run it remotely

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

Run this way there is no local file, so logs and CSV exports go to `%USERPROFILE%\M365Toolkit\`
on Windows or `~/M365Toolkit` on macOS and Linux. Using a tag (not `main`) also avoids the
CDN-cache lag on fresh pushes.

If a required module isn't installed, the tool offers to install it for the current user.

- Uses **your signed-in admin's RBAC role** — a helpdesk member can only do what their
  role permits.
- Every action (sign-in and sign-out, each command, each bulk row, each undo) is recorded in
  `logs\actions-<timestamp>.jsonl`: who, which service and tenant, the command, its target and
  the outcome. Other values are not kept: the Microsoft 365 unified audit log is the full record
  of changes, and this local file is not an audit trail. A full screen transcript
  (`logs\session-<timestamp>.log`) is written only with `-Transcript`. Logs and the tool's own
  exports are owner-only and deleted after 30 days.
- At any parameter prompt, type **`?`** to list the allowed values. In the category and
  action menus, type **`?`** for a short explanation of each choice.

## What you can do

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

## Tests

The engine's logic is covered by Pester tests in `tests/` that run without a tenant (stub
cmdlets, scripted answers). Run them with:

```powershell
Invoke-Pester ./tests
```

GitHub Actions runs PSScriptAnalyzer and the tests on PowerShell 7 (Linux and Windows) and
Windows PowerShell 5.1 for every push and pull request. `TEST-PLAN.md` lists the manual
checks that need a real tenant.

## Status

See [CHANGELOG.md](CHANGELOG.md) for what changed in each version, and
[SECURITY.md](SECURITY.md) to report a security problem.

Known limits, on purpose:

- Interactive, delegated sign-in only: no app-only or certificate auth for unattended runs.
  Adding it would need a service-account review first.
- Legacy AzureAD and MSOnline are retired and not supported; use the Graph services.
- An optional plain-English layer through an external AI service was considered. It would
  send directory data to a new party and turn model output into admin commands, so it needs
  a security review and IT approval before anyone builds it.
