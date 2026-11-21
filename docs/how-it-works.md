# How it works

[Back to the README](../README.md)

The tool is built from independent layers. Each one is extended by adding one row.

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
