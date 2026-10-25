# Manual test plan (v0.31)

The engine's logic is covered by the automated tests in `tests/` (run `Invoke-Pester ./tests`,
and CI runs them on every push). This plan covers what those tests can't: real sign-ins, real
Exchange and Graph behaviour, and the three platforms. Run it before a release.

## Before you start

- Use test mailboxes and test users (or a test tenant), never real people's mailboxes.
- Sign in with a separate admin account, not your everyday account.
- Note the version in the start banner and the action log path it prints.
- Each check: do the steps, compare with "Expect", tick it.

## A. Start, sign-in and sign-out

- [ ] **A1 Start.** Run `.\Invoke-M365Toolkit.ps1`. Expect: banner with the version, the line
  `Action log: ...logs/actions-<stamp>.jsonl`, no transcript file in `logs/`.
- [ ] **A2 Real account.** Connect Exchange Online. Expect: `Signed in as: <account>` and the
  tenant. In the action log: a `sign-in` line with that account.
- [ ] **A3 Mismatch.** Type one email at start, sign in with another account in the browser.
  Expect: a red "You typed X, but this session is signed in as Y" and a `Continue as Y?`
  question. Answer `n`: the tool signs out and returns to the service menu.
- [ ] **A4 Graph browser sign-in.** Connect Entra. Expect: a browser sign-in (no device code),
  then the account, tenant and granted scopes.
- [ ] **A5 Device code fallback.** Start with `-UseDeviceCode`, connect Entra. Expect: a device
  code prompt instead of the browser.
- [ ] **A6 Extra scopes.** If the Graph PowerShell app has write scopes consented in the tenant,
  expect a yellow "This token can also WRITE ..." line after connecting Entra.
- [ ] **A7 Shared Graph sign-in.** Connect Entra, then Intune, then Entra again. Expect: Entra
  signs in again the third time (Intune replaced its session).
- [ ] **A8 Sign-out on Quit.** Connect Exchange and Entra, then Quit. Expect: "Signed out of ..."
  for each. In the same shell, `Get-ConnectionInformation` returns nothing and `Get-MgContext`
  returns nothing.

## B. Single actions

- [ ] **B1 Grant.** Give a test user access to a test mailbox. Expect: preview, final
  confirmation, a plain `y/n`, then "Done." and a `run` line with outcome OK in the action log.
- [ ] **B2 Remove.** Remove that access. Expect: a red line saying it is destructive and a
  request to type the command name; a plain `y` cancels.
- [ ] **B3 Forwarding.** Change a mailbox setting and set `ForwardingSmtpAddress`. Expect: the
  typed confirmation ("it forwards mail elsewhere").
- [ ] **B4 Undo.** Change a mailbox setting such as `Office`. Expect: "Current value(s) that will
  change" with old and new value, an undo line, and in "What next?" an option "Undo the last
  change (Set-Mailbox on Identity=..., made HH:mm)". Run it and check the old value is back.
- [ ] **B5 Undo line with odd text.** Set a test mailbox's auto-reply text to
  `It's done; $(Get-Date)`. Then, via "Run any command", run `Set-MailboxAutoReplyConfiguration`
  on that mailbox with a new `InternalMessage`. Expect: the printed undo line shows the old text
  inside single quotes with the apostrophe doubled; pasting it into a new shell restores the
  text and runs nothing else.
- [ ] **B6 Stale undo.** After B4, run another change (any action). Expect: the undo offer is gone.
- [ ] **B7 Warnings.** Remove a permission the user doesn't have. Expect: "Done, with warning(s)
  above - check whether the change was actually made."
- [ ] **B8 Calendar, localized.** Use a mailbox whose folders are in another language (e.g.
  German "Kalender"). Check, give, change and remove calendar access. Expect: all work.
- [ ] **B9 Group lookup.** "See which groups a user is in" for a user in a nested group. Expect:
  only direct groups (the help says so).
- [ ] **B10 Shared mailbox access.** "See which shared mailboxes a user can access". Expect: the
  list; if some mailboxes can't be read, a warning with the count.
- [ ] **B11 Forwarding check.** On a test mailbox, set forwarding and add an inbox rule that
  forwards, and one that moves mail to "RSS Feeds" and marks it read. Run "See where a mailbox
  forwards mail". Expect: all three listed (with "keeps NO copy" if no copy is kept); an
  ordinary rule (e.g. move newsletters to a folder) is not listed.
- [ ] **B12 Leaver check.** For a test user with Send As, Send on behalf, Full Access on a shared
  mailbox and a distribution group, run "Leaver check". Expect: each appears under its area.
  Save it to CSV. If an area can't be read (e.g. missing role), expect a "(could not be read)"
  row and a warning.
- [ ] **B13 Nested groups.** For a user in a nested group, run "See ALL groups a user is in".
  Expect: the nested group listed too (B9 lists only direct ones).
- [ ] **B14 Reads run once.** Pick any "Check access" action. Expect: no "Run once / Bulk"
  question.

## C. Bulk mode

- [ ] **C1 Problem rows.** Use a CSV with one good row, one unknown user, one blank required
  cell and one invalid `AccessRights` value. Expect: the three problem rows listed before
  anything runs, and the choice to cancel or run only the good row.
- [ ] **C2 Destructive gate.** Bulk remove for 3 rows. Expect: "type exactly:
  Remove-MailboxPermission 3"; a plain `y` cancels.
- [ ] **C3 Blank cells.** Map an optional column with some blank cells. Expect: the blank-cell
  question; with "Skip", those rows keep their value; with "Send it as empty", the value clears.
- [ ] **C4 Results file.** After a run, open `exports/bulk-<cmdlet>-<stamp>-results.csv`. Expect:
  every row with `BulkStatus` (OK, WARNING, FAILED, SKIPPED), and for a `Set-*` the replaced
  values in `BulkBefore`.
- [ ] **C5 Excel CSV.** Save a CSV from Excel (default "CSV", and in an EU locale with `;`) with
  names like "Mueller" written with an umlaut. Expect: the tool reads it and the names are intact.
- [ ] **C6 Secret column.** Bulk run a cmdlet with a password column (e.g. in a test tenant).
  Expect: the value never shows on screen, in errors or in the results file.

## D. Run any command

- [ ] **D1 Exchange.** Expect: only Exchange cmdlets in the verb lists; no `Connect-*`.
- [ ] **D2 Other service after Exchange.** Connect Exchange, then Entra, open "Run any command".
  Expect: Graph user and group cmdlets, no Exchange cmdlets, no `Invoke-MgGraphRequest`.
- [ ] **D3 Non-recipient value.** Run a cmdlet whose `-Identity` is a policy name. Expect: the
  value is accepted (no "No mailbox/user found").

## E. Local files

- [ ] **E1 Action log.** Open `logs/actions-<stamp>.jsonl`. Expect: start, sign-in, run and stop
  lines with account, service, cmdlet, target and outcome; no other parameter values.
- [ ] **E2 Transcript.** Start with `-Transcript` and connect a service. Expect:
  `logs/session-<stamp>.log` exists and holds no device code or sign-in text.
- [ ] **E3 Exports.** Save a result to CSV. Expect: the file is owner-only (macOS/Linux:
  `ls -l` shows `-rw-------`), and a value starting with `=` opens in Excel as text.
- [ ] **E4 Retention.** Set an old date on a file in `logs/` or `exports/` (older than 30 days)
  and start the tool. Expect: "Retention: removed ..." and the file is gone.
- [ ] **E5 Synced folder.** Run the tool from a OneDrive or iCloud folder. Expect: the warning
  about a synced or shared folder.

## F. Platforms and modules

- [ ] **F1 Windows PowerShell 5.1.** Run A1, A2, B1, C1 on Windows PowerShell 5.1.
- [ ] **F2 PowerShell 7 on macOS and Windows.** Run A1 to A4 and B1.
- [ ] **F3 Missing module.** On a machine without a module, connect its service. Expect: name,
  version and author from the PowerShell Gallery, a question, then PowerShellGet's own
  untrusted-repository prompt.
- [ ] **F4 SharePoint (PnP).** Expect: a clear message on PowerShell older than 7.4, and a
  client ID question.
- [ ] **F5 Windows-only services.** On macOS, connect SharePoint Admin or Power Platform. Expect:
  a clear "Windows-only" message.

## G. Preview mode

- [ ] **G1 `-DryRunOnly`.** Expect: previews only, nothing runs, no undo offered. It still signs
  in and reads the tenant.
