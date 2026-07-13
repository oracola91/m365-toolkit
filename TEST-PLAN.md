# M365 Toolkit — Manual QA Test Plan

Test plan for `Invoke-M365Toolkit.ps1` (v0.18, single-file app, ~1,552 lines).
Every test below is tied to a real function or menu label in the code so the tester
knows exactly what to look for. Line numbers refer to `Invoke-M365Toolkit.ps1`.

The app is an interactive, guided PowerShell "command builder" for Microsoft 365
(Exchange Online + Entra/Graph and others), run on macOS/pwsh 7 or Windows PS 5.1.

---

## 1. Preconditions / test environment setup

**Host / runtime**
- macOS with PowerShell 7+ (`pwsh`) is the primary target (the code carries macOS-specific
  fixes). If available, also spot-check on Windows PowerShell 5.1.
- Terminal window wide enough to test column wrapping, and a way to resize it narrow
  (see the multi-column menu tests).

**Modules** (install per-user as needed; the tool will offer to install missing ones):
- `ExchangeOnlineManagement` (v3.7+; used by Exchange Online, and by Security & Compliance
  via `Connect-IPPSSession`). Required for most core tests.
- `Microsoft.Graph` (Entra/Identity + Intune tests).
- Optional for full coverage: `PnP.PowerShell`, `MicrosoftTeams`,
  `Microsoft.Online.SharePoint.PowerShell`, `Microsoft.PowerApps.Administration.PowerShell`.

**Credentials / tenant**
- An M365 admin (or helpdesk) account. The tool uses **your** RBAC — note which role you
  test as, since "permission denied" cases depend on it.
- A test tenant is strongly preferred because several actions **write** (Add/Set/Remove/New).
  Use `-DryRunOnly` (below) for any run where you must not mutate the tenant.
- Know at least: one real user UPN, one shared mailbox, one distribution group with members,
  and one mailbox whose calendar has a sharing entry — for the "check" actions.

**Test data to prepare**
- A CSV with a header row and 2–3 rows for bulk tests (e.g. columns `Identity,User,AccessRights`).
- A malformed CSV (no data rows) and a non-existent path for negative bulk tests.

**Two run modes (README lines 33–39)**
- Safe: `pwsh ./Invoke-M365Toolkit.ps1 -DryRunOnly` — builds and previews commands but
  **never executes** (`$DryRunOnly`, checked at lines 871 and 1020). Use this for the bulk
  of menu/navigation testing.
- Live: `pwsh ./Invoke-M365Toolkit.ps1` — asks `y/n` before each execution.

**Where output lands**
- `$script:ToolkitHome` resolves to the script folder when run locally, else
  `~/M365Toolkit` (lines 34–37). Logs → `logs/session-<timestamp>.log`;
  CSV exports → `exports/<cmdlet>-<timestamp>.csv`; undo hints → `logs/undo-hints.log`.
- `config.json` stores `Upn` + `LastService` (currently
  `taner.ibryamov@shelly.com` / `Exchange Online`).

**Before each full pass:** back up / note the current `config.json`, and `ls logs/ exports/`
so you can confirm new files are created.

---

## 2. Test cases by feature area

### A. Startup, config & logging

**A1 — Clean startup banner + transcript**
- Objective: The tool launches, logs, and greets.
- Steps: Run `pwsh ./Invoke-M365Toolkit.ps1 -DryRunOnly`.
- Expected: Banner `Microsoft 365 Guided Command Builder (v0.18)` (lines 1411–1413).
  A new `logs/session-<timestamp>.log` file exists and, after quitting, contains the session
  transcript (`Start-Transcript`/`Stop-Transcript`, lines 1409/1548).
- Edge: Delete the `logs/` folder first — it must be recreated (lines 1405–1406).

**A2 — Remembered UPN default**
- Objective: Config pre-fills the email prompt.
- Steps: At `Your email (UPN) [Enter = taner.ibryamov@shelly.com]`, press Enter.
- Expected: The stored UPN is used (lines 1416–1418).
- Edge: Type a different UPN — it should be used and later saved to `config.json`
  (`Save-ToolkitConfig`, line 1437, only after a successful service selection).
- Negative: Temporarily rename `config.json` so it's missing → prompt shows plain
  `Your email (UPN)` with no default; `Get-ToolkitConfig` returns nulls (lines 1323–1327).
- Negative: Corrupt `config.json` (invalid JSON) → tool must not crash; the `try/catch`
  falls back to null config (line 1325).

**A3 — Last-service floated to top**
- Objective: The most-recently-used service appears first, rest A–Z.
- Steps: With `LastService = Exchange Online`, reach the `Connect to which service?` menu.
- Expected: `Exchange Online` is option 1; remaining services listed A–Z below it
  (lines 1424–1427). `Quit` appears under the `-- app --` divider.

**A4 — ToolkitHome fallback**
- Objective: Home-folder fallback path when there is no script file.
- Steps (optional/advanced): Simulate the `irm | iex` path (no `$PSScriptRoot`).
- Expected: Logs/exports go under `~/M365Toolkit` (lines 34–37). Hard to force locally —
  at minimum verify the local path is used when running the file directly.

### B. Service selection, module import & connect

**B1 — Connect to Exchange Online (happy path)**
- Objective: EXO connects with macOS fixes applied.
- Steps: Choose `Exchange Online`.
- Expected: `Connecting to Exchange Online as <upn> ...` then `Connected to Exchange Online.`
  (lines 553–555). No `exchange.format.ps1xml` temp-file error (the connect splat sets
  `SkipLoadingFormatData = $true` and `LoadCmdletHelp = $true`, line 53). Modern-auth browser
  sign-in completes.
- Then `Show-WhoAmI` prints `Connected as: <upn>` plus `Org:` / `UPN:` from
  `Get-ConnectionInformation` (lines 1341–1345).

**B2 — Connect to Entra / Graph via device code**
- Objective: Graph uses device-code flow (no silent browser hang on macOS).
- Steps: Choose `Entra / Identity (Graph)`.
- Expected: A URL + code are printed to sign in (`UseDeviceCode = $true`, line 62). After
  auth, `Show-WhoAmI` prints `Account:` and `Granted scopes:` including
  `User.Read.All, GroupMember.Read.All` (lines 1346–1351).

**B3 — Missing module offer-to-install**
- Objective: Graceful handling when a module isn't installed.
- Steps: Choose a service whose module you have NOT installed (e.g.
  `Microsoft Teams` or `Power Platform (admin)`).
- Expected: `Module '<name>' is not installed.` then `Install it now for the current user? (y/n)`
  (lines 537–539).
  - Answer `n` → throws `Cannot connect without module '<name>'.` and returns you toward the
    menu without crashing the process (line 541).
  - Answer `y` → `Install-Module ... -Scope CurrentUser -Force` runs (line 540). (Only do this
    if you actually want the module.)

**B4 — Connection caching (don't reconnect)**
- Objective: A service already connected in this session is not reconnected.
- Steps: Connect to Exchange Online, do a task, `Switch service`, pick Exchange Online again.
- Expected: No second connect banner; goes straight to the category menu
  (`$connected[$svcName]` guard, lines 1432–1436).

**B5 — Connect failure surfaces cleanly**
- Objective: A failed sign-in doesn't leave a corrupt state.
- Steps: Start a connect and cancel the auth / use a bad account.
- Expected: The error is shown; because `$ErrorActionPreference='Stop'`, the connect throws.
  Verify the tool does not silently claim "Connected". (Note this is a robustness area — see
  Regression hotspots.)

### C. Menu system (`Get-MenuChoice`, lines 467–527)

**C1 — Basic numbered selection**
- Objective: Pick by number works; numbering is continuous across options + extras.
- Steps: On any menu, choose a valid number.
- Expected: Correct item returned. `ExtraOptions` are numbered continuing from the main list
  and rendered under `-------- app --------` (lines 502–508).

**C2 — "Enter = Back" and 0 = Back**
- Objective: On menus with `-AllowBack`, both `0` and empty Enter go back.
- Steps: At the category menu (has `-AllowBack`, line 1450), press Enter, then try `0`.
- Expected: Both return `$null` → you go back / switch service (line 521).

**C3 — `?` descriptions**
- Objective: Typing `?` lists per-choice explanations where descriptions exist.
- Steps: At the category→action menu (built with `-Descriptions` from each action's `Help`,
  lines 1533–1535), type `?`.
- Expected: Each action printed with its Help text indented beneath (lines 513–519). Then
  it re-prompts for a number.
- Negative: On a menu without descriptions (e.g. the service menu), `?` should just be
  "Not a valid choice" — the `?` branch requires `$Descriptions` (line 513).

**C4 — Invalid input re-prompts**
- Objective: Bad input is rejected, not accepted.
- Steps: Type a letter, a number out of range (e.g. `999`), a negative, whitespace.
- Expected: `Not a valid choice, try again.` and re-prompt (lines 522–525). Range check is
  `1..$all.Count`.

**C5 — Long-menu multi-column layout**
- Objective: Menus with >12 options flow into multiple columns; ≤12 stay single-column.
- Steps: Enter `Run any command by name (advanced)` on Exchange Online → the **verb picker**
  and **command lists** are long (>12). Observe columns.
- Expected: Options flow left-to-right, top-to-bottom across as many columns as fit the window
  (lines 484–498). Numbering still maps correctly (pick a high number and confirm the right
  command is chosen).
- Edge: **Resize the terminal narrow** and repeat → fewer columns, no wrapping past the
  window edge (`$Host.UI.RawUI.WindowSize.Width`, lines 486–488). Very narrow → falls back to
  1 column (`Math::Max(1, …)`, line 488).

### D. Category browsing & "Switch service"

**D1 — Category list A–Z**
- Objective: Categories are unique + sorted, with fixed items appended.
- Steps: On Exchange Online, open the category menu.
- Expected: Categories A–Z (`Check access`, `Give access`, `Mailboxes`,
  `Remove & change access`), then `Run any command by name (advanced)`, then under the divider
  `Search actions by keyword` and `Switch service` (lines 1447–1451).

**D2 — Action list A–Z within a category**
- Objective: Actions sorted by Label.
- Steps: Open `Check access`.
- Expected: The 6 check actions sorted A–Z (`Sort-Object Label`, line 1531), each with `?`
  help available.

**D3 — Switch service**
- Objective: `Switch service` returns to the service picker.
- Steps: Choose `Switch service` (extra option, line 1450), or Back out of the category menu.
- Expected: Back at `Connect to which service?` (lines 1451, and `Get-NextStep` 'switch').

**D4 — Service with no catalogued actions**
- Objective: Services without `$Actions` rows still work via run-any.
- Steps: Connect to a service that has no/one catalogued action (e.g. `Microsoft Teams`,
  `Security & Compliance`).
- Expected: `No plain-language actions catalogued for this service yet.` message
  (lines 1442–1445); category menu still offers `Run any command by name (advanced)`.

### E. Keyword search (`Search actions by keyword`, lines 1454–1472)

**E1 — Match by label / cmdlet / category**
- Objective: Search matches across Label, Cmdlet, and Category.
- Steps: Choose `Search actions by keyword`, type `calendar`.
- Expected: Results include the calendar actions, sorted A–Z, with `?` descriptions
  (lines 1458–1463). Selecting one runs it (`Invoke-Action`).
- Also try `Set-Mailbox` (cmdlet match) and `Give access` (category match).

**E2 — No matches / empty keyword**
- Steps: Search `zzzzz`; then search with empty input.
- Expected: `No actions match 'zzzzz'.` (line 1459). Empty keyword → nothing happens, returns
  to category menu (guard `if ($kw)`, line 1456).
- Edge: Keyword with regex metacharacters (e.g. `access)`) must not error — input is escaped
  via `[regex]::Escape` (line 1457).

### F. Guided run — core wizard (`Invoke-GuidedCmdlet`, lines 706–916)

Use a **read-only** action for most of these (e.g. `See who has access to a mailbox` →
`Get-MailboxPermission`) so you can run live safely.

**F1 — Header + param counts + parameter set**
- Objective: The build header reports the cmdlet, help, and param breakdown.
- Steps: Choose `Give a user access to another mailbox` (`Add-MailboxPermission`).
- Expected: `Building: Add-MailboxPermission`, the Help line, `N required, M optional, K switches`
  (lines 724–727), and `(using parameter set: …)` when applicable (lines 728–730). This action
  exercises **parameter-set resolution** — confirm `-Owner` does NOT appear as required
  (the whole point of `Get-CmdletParameters`, lines 648–672).

**F2 — Presets shown, never prompted**
- Objective: Preset values are auto-bound and displayed.
- Steps: Choose `Create a new shared mailbox` (`New-Mailbox`, `Preset=@{Shared=$true}`), or
  `Let a user send AS another mailbox` (`Preset=@{AccessRights='SendAs'}`).
- Expected: `Fixed by this action: -Shared True` (lines 737–740); `Shared`/`AccessRights` never
  prompted, and it appears in the preview.

**F3 — Required prompts in curated order + back navigation**
- Objective: Required params prompted in `Prompt` order; `back` steps to previous field.
- Steps: In `Add-MailboxPermission`, at the 2nd field (`User`) type `back`.
- Expected: `Going back to -Identity.` and the Identity value is discarded/re-asked
  (lines 759–765). At the **first** field, `back` → `Already at the first field - type 'cancel'…`
  (line 760).

**F4 — Cancel to menu**
- Objective: `cancel` aborts with nothing run.
- Steps: At any prompt type `cancel`.
- Expected: `Cancelled - back to the menu (nothing was run).` and return to menu
  (lines 758, 774, etc.). `$script:ParamCancel` drives this.

**F5 — Optional param picker**
- Objective: Remaining optional params are offered as a numbered multi-pick.
- Steps: In `Add-MailboxPermission`, when asked which optional params to set, type `?`.
- Expected: Each optional listed with a description (ParamHelp → `$ParamDocs` glossary →
  `Get-Help`), then re-prompt (lines 791–807). Enter a couple of numbers → each is then prompted.
  `AutoMapping` should show the glossary text from `$ParamDocs` (line 452).
- Edge: Enter → skips all optional (blank pick). Enter `abc` / out-of-range numbers → ignored
  (regex `^\d+$` + bounds, line 808).

**F6 — Single identity-like optional auto-ask**
- Objective: A lone identity optional is asked directly (no "which numbers?" step).
- Steps: Use a `Get-*` with a single optional Identity in `ParamLookup` (e.g. run-any
  `Get-OrganizationConfig` style, or any cmdlet whose only optional is `Identity`).
- Expected: `Optional: Identity - enter it now, or press Enter to skip.` (lines 780–789).

**F7 — Switches picker**
- Objective: Switch params are offered; chosen ones become `$true`.
- Steps: In an action with switches, select one by number.
- Expected: It appears in the preview as `-SwitchName` (no value). `?` explains each
  (lines 819–842). Preview rendering: switches only show if present (lines 569–570).

**F8 — Dry-run preview + final confirm**
- Objective: The exact command is previewed and requires explicit `y`.
- Steps: Complete a read-only action.
- Expected: `Preview (dry-run):` with the assembled command, then
  `FINAL CONFIRMATION - about to execute this command:` and `Execute it now? (y/n)`
  (lines 866–883). Answer `n` → `Cancelled - nothing was run.` Answer `y` → runs.
- DryRunOnly mode: `-DryRunOnly` → `DryRunOnly is set - not executing.` and no confirm
  (lines 871–873).

**F9 — Preview value formatting (`Format-CommandPreview`, lines 561–582)**
- Objective: Values render correctly by type.
- Verify across actions: booleans render as `$true`/`$false` (line 571–572); arrays as
  comma-joined quoted values (line 573); strings with spaces are single-quoted (line 575);
  switches show as bare `-Name` (line 570).

**F10 — Execution error is caught**
- Objective: A cmdlet failure shows a red message, not a crash.
- Steps: Run an action against a non-existent target so the cmdlet throws (e.g.
  `Get-MailboxPermission -Identity nosuchbox@x` — bypass the resolver via run-any if needed).
- Expected: `Failed: <message>` in red (lines 913–915); tool continues to `What next?`.

### G. Parameter value entry (`Read-ParameterValue`, lines 1093–1241)

**G1 — ValidateSet / enum / hints → numbered picker**
- Objective: Fixed-value params show a picker, not free text.
- Steps: `Give a user access to a calendar` (`Add-MailboxFolderPermission`) → `AccessRights`
  uses the per-action `Hints` list of calendar roles (lines 263, 1114); pick a role.
- Expected: `Read-OptionMenu` shows the numbered roles (lines 1051–1087). Compare with
  `Add-MailboxPermission` where `AccessRights` comes from global `$ParamHints`
  (`FullAccess`, `ExternalAccount`, … line 408) — the **per-action Hints must win** over the
  global hint (Hints checked before ParamHints, lines 1114–1115).

**G2 — `?` at a value prompt**
- Objective: `?` lists allowed values / help without submitting.
- Steps: At a picker or free-text prompt, type `?`.
- Expected: Description and/or `Options: …` printed, then re-prompt (lines 1157–1162,
  1066–1069).

**G3 — Directory lookup via `find`**
- Objective: `find` searches the directory and lets you pick.
- Steps: At an Identity/User field (EXO connected), type `find`, then a search term.
- Expected: Prompt `Search text (name or email)`, results from `Get-EXORecipient`
  shown as `Display <smtp> [Type]`, pick one → `-> using <smtp>` (lines 1168–1179,
  418–427). `find` is only offered when the probe cmdlet is loaded (label built at 1149).
- Edge: Search term with no matches → `No matches - try again, or just type the value.`
  (line 1173). Cancel the pick (Back) → returns to the field prompt.
- Graph variant: On Entra, `See which groups a user is in` → `UserId` field offers `find`
  backed by `Get-MgUser` (lines 435–443).

**G4 — Identity validation (`Resolve`) on typed input**
- Objective: A typed identity is verified to exist.
- Steps: At an EXO Identity field, type a nonsense value like `zzznobody@x.y`.
- Expected: `No mailbox/user found matching 'zzznobody@x.y'. Check the spelling, or type 'find'…`
  and re-prompt (lines 1217–1223). A real recipient is accepted.
- Note: `Filter` field (mail-groups action) also maps to the recipient lookup — confirm typing
  a real person there is accepted (it's resolved then rewritten by Transform).

**G5 — Required empty / optional empty**
- Objective: Empty handling differs by required-ness.
- Steps: At a required free-text field press Enter; at an optional one press Enter.
- Expected: Required → `This one is required.` re-prompt (lines 1210–1212). Optional →
  returns `$null`, field skipped (line 1214).

**G6 — Array field entry**
- Objective: Comma/space-separated arrays parse; allowed-set arrays validate + normalize.
- Steps: For an array `AccessRights` (multi), enter `FullAccess, badvalue`.
- Expected: `Not valid: badvalue. Choose from: …` re-prompt (lines 1230–1233). Enter valid
  values → normalized to canonical casing (line 1235). Free array (no allowed set) → split
  on `[,\s]+` (line 1226).

**G7 — Type hints**
- Objective: The prompt shows a human hint per type.
- Verify: Boolean → `true / false`; numeric → `a number`; DateTime → `a date, e.g. 2026-07-31`;
  string → `text (e.g. an email address or name)` (lines 1130–1137).

**G8 — `Read-OptionMenu` back / cancel**
- Objective: The fixed-set picker supports back (0/`back`) and `cancel`.
- Steps: At a role picker, type `0`, then re-enter and type `cancel`.
- Expected: `0`/`back` sets `ParamBack` (returns to prior field); `cancel` sets `ParamCancel`
  (aborts to menu) (lines 1071–1072). Invalid number → `X is not in the list (1-N).` (line 1080).
- Multi vs single: for a single-value field, entering `1,3` → `Just one number for this field.`
  (line 1075).

### H. Group membership & shared-mailbox checks

**H1 — Which mail groups a user is in (EXO)**
- Objective: Resolves a person to their DN and lists their mail groups.
- Steps: `Check access` → `See which mail groups a user is in` → type a real user at `Filter`.
- Expected: The Transform rewrites `Filter` to `Members -eq '<DN>'` (lines 194–198); result is
  projected to `Name / Email / Type` with friendly types (Distribution group, etc.,
  lines 201–213). Preview shows the rewritten filter.
- Edge: User with no groups → `No objects returned.` (via Show-Result, line 1282).

**H2 — Who is in a distribution group (EXO)**
- Steps: `See who is in a distribution group` → type a group.
- Expected: Members projected to Name/Email/Type (`Get-DistributionGroupMember`, lines 215–237).
- Edge: Empty group → no objects returned. Non-group identity → cmdlet error caught as `Failed:`.

**H3 — Which shared mailboxes a user can access (EXO)**
- Objective: The custom helper scans shared mailboxes for the user's FullAccess.
- Steps: `See which shared mailboxes a user can access` → type a user.
- Expected: Runs `Get-UserSharedMailboxAccess` (lines 117–128) — may take a minute; returns
  shared mailboxes where the user has direct, non-inherited FullAccess, projected to
  Name/Email/Type=`Shared mailbox` (lines 244–248).
- Edge: User with none → no objects returned.

**H4 — Which groups a user is in (Graph)**
- Objective: Graph member-of with type flattening.
- Steps: On Entra, `See which groups a user is in` → `UserId`.
- Expected: `Get-MgUserMemberOf -All` (Preset, line 350); the Project flattens
  `AdditionalProperties` into Name/Email/Type where Type distinguishes Microsoft 365 group /
  Security / Mail-enabled security / Distribution group / Directory role (lines 356–381). This
  is the richest projection — verify types look correct against a known user.

**H5 — Calendar access checks & changes**
- `See who has access to a calendar` (`Get-MailboxFolderPermission`): type just the mailbox;
  Transform appends `:\Calendar` (line 254) — verify preview shows `<mbx>:\Calendar`.
- `Give a user access to a calendar` (`Add-MailboxFolderPermission`): role picker from Hints
  (line 263). Run in DryRunOnly and confirm the built command targets `<mbx>:\Calendar`.
- `Change a user's calendar access` (`Set-MailboxFolderPermission`) — see the safety-net tests (J).
- `Remove a user's calendar access` (`Remove-MailboxFolderPermission`).

### I. Run any command by name (advanced) — lines 1475–1529

**I1 — Verb → command drill-down**
- Objective: Pick a verb, then a command, with no typing.
- Steps: `Run any command by name (advanced)` → pick verb `Get` → pick a command.
- Expected: Verb list built from connected commands (baseline-diffed, lines 1480–1485); after a
  verb, its commands are listed (line 1514). Selecting one enters the run-once/bulk chooser
  (lines 1517–1522).

**I2 — Baseline scoping**
- Objective: Only the connected service's commands are offered (not core PowerShell).
- Steps: On Exchange Online, browse verbs — you should see EXO verbs/commands, not generic
  ones like `Write` from core. (`$script:BaselineCommands` diff, lines 1402–1403, 1480–1481.)
- Edge: If the diff yields zero (can happen with Graph's on-demand submodules) it falls back to
  ALL commands (lines 1482–1484) — verify Graph run-any still lists commands.

**I3 — "Type a name or keyword instead"**
- Objective: Free-form name entry with validation and match-picking.
- Steps: In run-any, choose `Type a name or keyword instead`.
  - Full valid name (e.g. `Get-Mailbox`) → accepted directly (line 1496).
  - Partial (e.g. `mailboxperm`) → shows matching commands to pick (lines 1497–1501).
  - `?` → help text `Type a full cmdlet name to use it…` (line 1495).
- Negative: Bad name (`Get-Nothing`) → `Nothing matches 'Get-Nothing'.` and **re-prompts in
  place** (line 1498). Over-broad term matching >40 commands → `'X' matches N commands - narrow
  it down.` (line 1499). Empty entry → back to the verb menu (line 1494).

**I4 — Long verb command list filter**
- Objective: A verb with >40 commands prompts to filter first.
- Steps: Pick a common verb (`Get` or `Set`) on EXO.
- Expected: `Get- has N commands - type part of the name to filter, or Enter to list all`
  (lines 1507–1512). Filter that matches nothing → `Nothing matched - pick another verb.`

**I5 — Run-any drives the same engine**
- Objective: A picked cmdlet goes through `Invoke-GuidedCmdlet` with full introspection.
- Steps: Pick `Get-Mailbox` via run-any, Run once, supply Identity.
- Expected: Same required/optional/switch walk, preview, confirm, result handling. Errors are
  caught with `Could not run '<cmd>': <msg>` at the outer layer too (line 1523).

### J. Change safety-net & undo (v0.18) — `Get-RevertPlan` lines 590–616; use lines 856–904

**J1 — Before→after display for a Set-***
- Objective: A `Set-*` shows current values that will change.
- Steps: `Change a user's calendar access` (`Set-MailboxFolderPermission`) against an existing
  entry, OR run-any `Set-Mailbox` changing a simple property (e.g. `Office`).
- Expected: `Current value(s) that will change:` with `Prop: old -> new` per changed param
  (lines 857–863). `Format-SafetyValue` renders empty as `(empty)` and arrays comma-joined
  (lines 619–626).

**J2 — Undo command emitted + logged**
- Objective: After a successful Set-*, an exact revert command is printed and logged.
- Steps: Execute the Set-* from J1 (live, non-DryRun) against a disposable target.
- Expected: `To undo this change, run:` followed by the reversed command (old values),
  `(also saved to …/logs/undo-hints.log)` (lines 890–903). Open `logs/undo-hints.log` and
  confirm a timestamped line was appended. **Run the printed undo command** and verify it
  restores the original value.

**J3 — Unsupported cases fall back silently**
- Objective: The safety net degrades gracefully.
- Cases to verify (all should just skip the before/after + undo block, no error):
  - A non-`Set-*` cmdlet (e.g. `Add-*`, `Remove-*`) → `Supported=$false` (line 596).
  - A `Set-*` with no matching `Get-*` (line 598).
  - A `Set-*` with no `Identity/Id/User/Mailbox` bound (lines 599–601).
  - A target whose `Get-*` returns nothing / errors (lines 602–605).
  - A change to only switch params (switches excluded, line 610) → no undo block.

**J4 — Undo dir auto-created**
- Steps: Delete `logs/` then run a Set-* (live) that produces an undo hint.
- Expected: `logs/` recreated and `undo-hints.log` written (lines 898–899).

### K. Output → next-action piping (`last`) — Show-Result store lines 1297–1299; use lines 1183–1209

**K1 — `last` offered after an object-returning action**
- Objective: The next action can reuse a value from the previous result.
- Steps: Run a check that returns objects (e.g. `See who is in a distribution group`). Note the
  tip `Tip: in the next action, type 'last'…` (line 1299). Then start another action and at an
  Identity field, type `last`.
- Expected: The field label includes `- or 'last' to reuse a value from the last result`
  (only when `$script:LastResult` is non-empty and the field is scalar, lines 1145, 1150).
  It auto-picks an identity-like column (`PrimarySmtpAddress`, `UserPrincipalName`, `Email`,
  `Identity`, …, line 1186) and lists rows (with a DisplayName/Name context column, line 1193)
  to choose from; picking → `-> using <value>` (lines 1204–1208).

**K2 — Column fallback when no identity column**
- Objective: If no known identity column exists, the user is asked which column.
- Steps: Produce a result whose columns don't include a standard id column, then `last`.
- Expected: `Which value from the last result?` column picker (lines 1188–1190).

**K3 — `last` not offered / empty**
- Cases:
  - Fresh session, no prior result → no `last` in the label (LastResult empty).
  - Array field → `last` not offered (scalar-only, `-not $isArray`, line 1145).
  - Last result column has only blank values → `The last result has no '<col>' values to use.`
    (line 1202).
- Note: Only **object** results are stored (scalars/strings are not, lines 1286–1298) — after a
  scalar-returning command, `last` should not appear.

### L. Result rendering & CSV export (`Show-Result` lines 1278–1317)

**L1 — Table vs list rendering**
- Objective: One result → Format-List; many → Format-Table.
- Steps: Run a lookup returning exactly 1 object (`Look up a mailbox` for one mailbox), then
  one returning many (`Get-Mailbox` via run-any).
- Expected: Single item shows as a list (`Format-List`); multiple as an auto-sized wrapped
  table (line 1310). This is the "single-identity auto-jump" behavior.

**L2 — Choose columns**
- Steps: `Show a table with only columns I choose` → enter a couple numbers.
- Expected: Only those columns shown (`Select-ResultColumns`, lines 1246–1257). Enter (no
  numbers) → all columns (line 1254).

**L3 — CSV export (all / chosen)**
- Steps: `Save all columns to a CSV file`, accept default path.
- Expected: File written to `exports/<cmdlet>-<timestamp>.csv`, `Saved N row(s) to …`
  (lines 1262–1271). The `exports/` dir is auto-created (line 1268). Repeat with chosen columns.
- Edge: Provide a custom path with quotes/spaces → trimmed (line 1266).

**L4 — No objects / scalar output**
- Steps: Run something returning nothing (e.g. a Remove that returns no object) and something
  returning a string.
- Expected: `No objects returned.` (line 1282); scalar/string just printed line-by-line, no
  table menu (lines 1289–1291).

**L5 — Result menu loops**
- Objective: You can view AND export without re-running.
- Steps: Show table, then export, then `Done`.
- Expected: The menu re-appears until `Done` (while loop, lines 1301–1316).

### M. Bulk / CSV runner (`Invoke-BulkCmdlet` lines 923–1043)

**M1 — Map columns and run**
- Objective: Map each param to a CSV column or fixed value; run per row.
- Steps: Choose any write action → `Bulk - run for every row in a CSV file` → give the CSV
  path. Map `-Identity` to a column, `-AccessRights` to `=FullAccess`.
- Expected: `Loaded N rows. Columns: …` (line 941); guided mapping (`column / =fixed / Enter`,
  lines 942–962); first-row preview; `FINAL CONFIRMATION - about to run this for ALL N rows`;
  per-row `[i/N] OK`/`FAILED`; tally `Bulk done: X succeeded, Y failed, Z total.`
  (lines 1030–1042). Use `-DryRunOnly` to stop before execution (line 1020).

**M2 — Bulk negative cases**
- Missing file → `File not found: <path>` (line 935).
- CSV with header only (no rows) → `That CSV has no data rows.` (line 937).
- Required param left blank in mapping → `Required - give a column name or =value.` (line 956).
- Unknown column name typed → `'X' is not a column. Available: …` (line 961).
- Presets apply automatically to every row (line 946); Transforms applied per row (lines
  1007–1011) — verify e.g. calendar `:\Calendar` appending works in a bulk calendar action.

### N. "What next?" flow (`Get-NextStep` lines 1360–1367)

**N1 — Do another / switch / quit**
- Steps: After any action, at `What next?` choose each option in turn.
- Expected: `Do another task (same service)` → category menu; `Switch service` → service
  picker; `Quit` → `Goodbye.` and transcript stops (lines 1361–1366, 1545–1548).
- Edge: Backing out (Invoke-Action returns 'again' when the run/bulk chooser is cancelled,
  line 1378) → returns to category menu without running.

---

## 3. Regression hotspots (recently added — test hardest)

**RH1 — Change safety-net & undo (v0.18, newest).**
- The `Get-*` pairing (`Set-Foo` → `Get-Foo`) and identity discovery are heuristic
  (lines 597–604). Test Set-* cmdlets where: the Get- takes a positional identity vs `-Identity`
  (both code paths at lines 603–604); the property name in the result differs from the parameter
  name (should just be skipped, line 611); the value being set is an array or empty.
- Verify the **printed undo command actually reverses the change** (round-trip J2) — the highest-
  value check. A wrong undo hint is worse than none.
- Confirm the before/after and undo blocks are **absent** for Add-/Remove-/New- and for
  lookups (must not slow or clutter read-only actions).

**RH2 — Run-any-command (verb browser + name picker).**
- Baseline diff (lines 1402, 1480): on Graph the diff can be empty → fallback to all commands
  (lines 1482–1484). Confirm both EXO (non-empty diff) and Graph (possible fallback) list
  sensible commands.
- Bad-name **re-prompt in place** (line 1498) and the >40-match guard (line 1499) — type junk,
  type a single letter, type an empty string; ensure you can always get back to the verb menu
  (lines 1490, 1494) and never get stuck.
- A run-any cmdlet with unusual parameter sets goes through the same `Get-CmdletParameters`
  resolution — pick something with multiple sets and confirm no spurious "required" params.

**RH3 — Output→next-action piping (`last`).**
- The store happens only for object results in `Show-Result` (lines 1297–1298) — confirm a
  scalar-returning command does NOT enable `last` afterward.
- Column auto-detection order (line 1186) and the blank-value guard (lines 1196, 1202).
- `last` must be suppressed on array fields (line 1145). Cross-action flow: result from EXO
  action → `last` into a following action's Identity, then verify the value is actually used in
  the preview.
- State is per-session in-memory (lines 1394–1395) — after `Switch service` the last result
  should still be usable (it is not cleared), confirm behavior matches expectation.

**RH4 — Parameter-set resolution (`Get-CmdletParameters`, lines 648–679).**
- Underpins F1/F5 and everything in run-any. `Add-MailboxPermission` (AccessRights set vs Owner
  set) and `Set-Mailbox` (many sets) are the canonical stress tests — spurious required params
  here break many actions at once.

**RH5 — macOS EXO connect fixes (lines 53, 551).**
- The `SkipLoadingFormatData` / import-noise-swallow combo is macOS/pwsh-specific. Re-test a
  fresh `pwsh` session AND a session where EXO was already imported once (the stale
  `exchange.format.ps1xml` scenario the comment describes). No format-data error should reach
  the user.

**RH6 — Hints vs ParamHints precedence (lines 1114–1115).**
- Calendar `AccessRights` (per-action Hints = roles) vs mailbox `AccessRights` (global
  ParamHints = rights). Easy to regress into showing the wrong value list.

**RH7 — Multi-column menu math (lines 484–498).**
- Column count depends on live window width; number→item mapping must stay correct at various
  widths and with an odd count that doesn't fill the last row (line 493 bounds check).

---

## 4. Prioritized smoke-test checklist (run after ANY change)

Run in `-DryRunOnly` unless a step needs a live write. ~10–15 minutes.

1. **Launch + config**: `-DryRunOnly` starts, banner shows v0.18, UPN default pre-filled from
   `config.json`, a new `logs/session-*.log` is created. (A1/A2)
2. **Connect EXO**: choose Exchange Online, connects with no `exchange.format.ps1xml` error,
   `Connected as` + Org shown. (B1/RH5)
3. **Menus**: category menu lists A–Z with `-- app --` divider; `?` shows action descriptions;
   invalid input re-prompts; Enter/`0` = Back. (C2/C3/C4/D1)
4. **Guided read-only run**: `See who has access to a mailbox` → find/type an Identity → preview
   → confirm → results render (table/list). (F8/G3/L1)
5. **Parameter-set sanity**: build `Add-MailboxPermission` — confirm `-Owner` is NOT required and
   the param counts look right. (F1/RH4)
6. **Value pickers + precedence**: `Add-MailboxPermission` AccessRights uses mailbox rights,
   `Add-MailboxFolderPermission` uses calendar roles. (G1/RH6)
7. **back / cancel**: mid-wizard `back` re-asks the prior field; `cancel` returns to menu with
   nothing run. (F3/F4)
8. **Identity validation**: type a nonsense address → rejected with "No mailbox/user found". (G4)
9. **Run-any**: verb → command drill-down works; `Type a name or keyword` accepts a valid name
   and re-prompts in place on a bad one. (I1/I3/RH2)
10. **Piping**: after an object result, `last` is offered at the next Identity field and the
    chosen value lands in the preview. (K1/RH3)
11. **Safety-net/undo** (live, disposable target): a `Set-*` shows before→after and prints +
    logs a working undo command. (J1/J2/RH1)
12. **Result export**: export a result to CSV → file appears under `exports/`. (L3)
13. **Exit clean**: `What next?` → Quit → `Goodbye.`, transcript closes (log file complete). (N1)
