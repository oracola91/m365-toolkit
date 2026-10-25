# Changelog

All notable changes to the M365 Guided Command Builder. Newest first.

## v0.32 (2026-10-25)
- "Run any command" search matches the typed text literally; a `[` no longer ends the tool.
- A single run skips the cmdlet's own "Are you sure?" after the tool's confirmation (as bulk
  mode already did), so a declined vendor prompt can't be reported as "Done".

## v0.31 (2026-10-25)
- New read-only check "See where a mailbox forwards mail": mailbox forwarding plus inbox rules
  that forward, redirect, delete or hide mail (the usual signs of a taken-over mailbox).
- New "Leaver check": Send As, Send on behalf, Full Access on shared mailboxes, mail groups,
  forwarding and rules on the user's own mailbox, and who can open it, in one list for
  offboarding evidence. Areas that can't be read are listed, never skipped.
- New "See ALL groups a user is in, including nested groups" (Graph, transitive membership).
- Catalogued read actions run once; bulk mode is offered only for actions that change things.

## v0.30 (2026-10-25)
- Automated tests (`tests/`, Pester 5) and CI on GitHub Actions: PSScriptAnalyzer, and the
  tests on PowerShell 7 (Linux and Windows) and Windows PowerShell 5.1.
- `SECURITY.md` with private vulnerability reporting; this changelog.
- README and TEST-PLAN brought up to date (status, logging, paths, preview mode wording).
- The group lookup help says it lists direct memberships (nested groups are not included).
- Robustness: number input can't overflow, an empty email can't reach sign-in, dead code
  removed.

## v0.29 (2026-10-25)
- Action log by default (`logs/actions-<stamp>.jsonl`: who, service, tenant, command, target,
  outcome); full screen transcript only with `-Transcript`.
- 30-day retention for logs, the tool's exports and undo lines; warning when the tool's folder
  looks synced or shared.
- Owner-only protection, transcript pause and resume, module import, purge, undo-log writes
  and the shared-mailbox access scan now report failures.
- CSV export neutralises formula cells; tenant text is stripped of control characters.
- Modules install from the PowerShell Gallery only, without `-Force`, after showing name,
  version and author; Graph installs only the sub-modules it uses; Exchange needs 3.7+.
- PnP asks for the tenant's app client ID; calendar actions use the localized folder; `last`
  fills a field from the column of the same name.

## v0.28 (2026-10-25)
- Typed confirmation for Invoke- actions, raw non-GET requests, admin-role and mail-flow
  changes, purge/force switches and mail forwarding.
- "Run any command" lists only the connected service's own commands.
- The real signed-in account and tenant are shown after sign-in; a mismatch with the typed
  email stops. Graph warns about write scopes consented earlier.
- Graph signs in with the browser by default (`-UseDeviceCode` to opt in); token kept in the
  process. Every service is signed out on Quit.

## v0.27 (2026-10-25)
- Preview, confirmation and undo lines render values as PowerShell literals, so tenant data
  can't turn a pasted undo line into a command.
- Undo plans are declined when the record is ambiguous and say what they don't restore; the
  undo offer is cleared by later changes and never offered under `-DryRunOnly`.
- Secrets judged by type (fixes `-ResetPasswordOnNextLogon` being read hidden); PSCredential
  built from user name plus hidden password; secrets masked in results.

## v0.26 (2026-10-25)
- Bulk mode gets the single-run safety rules: every row checked first, typed gate with row
  count for destructive commands, blank cells skipped unless you choose to clear, targets
  listed, secrets masked, and a results CSV per run.
- Lists split on commas only; Excel ANSI and `;` CSV files are read correctly.

## v0.25 (2026-10-25)
- Bulk runner passes `-ErrorAction Stop` per row, like the single run and the undo.
- README "Add a new action" template includes the required `Id`.

## v0.18 to v0.24 (2026-07-10 to 2026-07-14)
- v0.24: pre-public security and code review fixes ("Run any command" limited to the
  service, remote launch pinned to a tag, masking of secret-named arrays, quote-aware
  colouring, apostrophe search fallback).
- v0.23: platform support documented; Windows-only services fail cleanly on macOS and Linux.
- v0.22: fixes from live QA. v0.21: actions identified by a stable `Id`.
- v0.20: Intune scopes read-only, owner-only config and undo log, secret masking in
  "Run any command".
- v0.19: calendar change/remove picks from people who already have access.
- v0.18: change safety net and undo for Set-* commands; audit fixes (booleans bind, undo of
  calendar permissions, OPATH quote escaping, typed confirm for destructive verbs).

## v0.9 to v0.17 (2026-07-08 to 2026-07-10)
- v0.17: "Run any command" above the app divider; long menus in columns.
- v0.16: browse "Run any command" by verb. v0.15: keyword picker, single-field jump,
  wide results. v0.14: command name validated in "Run any command".
- v0.13: "See which shared mailboxes a user can access". v0.12: "See who is in a
  distribution group". v0.11: "See which mail groups a user is in".
- v0.10: group-membership check and verb-based menus; read-only Graph scopes.
- v0.9: reuse a value from the last result with `last`.

## v0.1 to v0.8 (2026-07-03 to 2026-07-07)
- v0.8: calendar access category; per-action `Transform` and `Hints`; macOS Exchange connect
  fix; menus sorted A-Z.
- v0.7: `back` / `cancel` at prompts; multi-select picker; mailbox permission views.
- v0.6: parameter-set resolution; `?` explanations; value and identity validation; "What
  next?" menu; failures reported.
- v0.5: remembered email and service; who-am-I; action search; `find` directory lookup.
- v0.4: SharePoint Admin, Intune and Power Platform. v0.3: show and export results.
- v0.2: allowed-value discovery, plain-language help, presets, shared mailboxes,
  "Run any command", CSV bulk mode.
- v0.1: connection registry, introspection wizard, dry-run preview, starter catalogue.
