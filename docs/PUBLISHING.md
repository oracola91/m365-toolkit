# Publishing to the PowerShell Gallery

[Back to the README](../README.md)

The tool is one file, so it ships to the Gallery as a script, not a module. Users install it
with `Install-Script -Name Invoke-M365Toolkit`. The Gallery reads the `<#PSScriptInfo ... #>`
block at the top of `Invoke-M365Toolkit.ps1`. CI checks that block on every pull request.

## One time: account and API key

1. Sign in at <https://www.powershellgallery.com> with a Microsoft account.
2. Open your account menu, then **API Keys**, then **Create**.
3. Key name: `m365-toolkit`. Scope: **Push new packages and package versions**.
   Glob pattern: `Invoke-M365Toolkit`. Expiry: 365 days.
4. Copy the key once and keep it in your password manager.

> **Never paste the key in a chat, an issue or a commit.** Anyone with it can publish a
> new version of the tool under your name to every admin who installs it.

## Each release

1. Run [TEST-PLAN.md](../TEST-PLAN.md) on the version you want to release.
2. Set the same version in two places in `Invoke-M365Toolkit.ps1`:
   `.VERSION` in the PSScriptInfo block (for example `0.33.0`) and
   `$script:ToolkitVersion` (for example `'v0.33'`). CI fails if they differ.
3. Add the version to `CHANGELOG.md`. Merge the pull request.
4. Tag the merged `main` (for example `v0.33`) and publish the GitHub release.
   Update the pinned tag in [running.md](running.md).
5. Publish from a clean checkout of that tag, in pwsh. Load the key into the session
   without echoing it, then publish:

```powershell
git checkout v0.33
$env:PSGALLERY_KEY = Read-Host 'PSGallery API key' -MaskInput
Publish-Script -Path ./Invoke-M365Toolkit.ps1 -NuGetApiKey $env:PSGALLERY_KEY
Remove-Item Env:PSGALLERY_KEY
```

6. Check the page: <https://www.powershellgallery.com/packages/Invoke-M365Toolkit>.
   The Gallery scans every upload first, so it can take a few minutes to show.

A published version can be unlisted on the Gallery but not deleted or replaced. A fix
always needs a new version number.
