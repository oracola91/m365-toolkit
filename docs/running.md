# Running the tool

[Back to the README](../README.md)

## Run it locally

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

## Run it remotely

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

## Modules, permissions and logs

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
