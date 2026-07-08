<#
.SYNOPSIS
    Guided, plain-language command builder for Microsoft 365 PowerShell.

.DESCRIPTION
    You never memorize cmdlet names or syntax. Launch the tool, enter your email,
    choose which M365 service to connect to, then browse by category and pick an
    action in plain language. The tool introspects the real cmdlet, walks you
    through its parameters (Required -> Optional -> Switches) in human terms,
    shows a dry-run preview of the assembled command, and only runs it when you
    confirm. Everything is transcript-logged for an audit trail.

    Architecture (three independent layers, each extended by adding one row):
      1. $Services  - connection registry: one row per M365 service.
      2. Invoke-GuidedCmdlet - generic introspection engine, works for ANY cmdlet.
      3. $Actions   - plain-language catalogue: friendly label -> cmdlet.

.NOTES
    v0.10 - 2026-07-08. Runs on Windows PowerShell 5.1 or PowerShell 7+.
    Requires the relevant module installed for whichever service you connect to.
#>

[CmdletBinding()]
param(
    # Skip the run confirmation and only ever show the preview (never execute).
    [switch]$DryRunOnly
)

$ErrorActionPreference = 'Stop'

# Where logs and CSV exports go. Prefer the script's own folder, but when the
# tool is run remotely (irm <URL> | iex) there is no script file on disk, so
# $PSScriptRoot is empty - fall back to the user's profile / home folder.
$script:ToolkitHome =
    if     ($PSScriptRoot)     { $PSScriptRoot }
    elseif ($env:USERPROFILE)  { Join-Path $env:USERPROFILE 'M365Toolkit' }
    else                       { Join-Path $HOME 'M365Toolkit' }

# ---------------------------------------------------------------------------
# Layer 1: Connection registry. Add a service = add one row here.
# ---------------------------------------------------------------------------
$Services = @(
    [pscustomobject]@{
        Name        = 'Exchange Online'
        Module      = 'ExchangeOnlineManagement'
        ConnectCmd  = 'Connect-ExchangeOnline'
        # Builds the parameter splat for the connect cmdlet from the entered email.
        # LoadCmdletHelp so Get-Help works (EXO v3.7+ skips it by default) - lets
        # the tool show real Microsoft parameter descriptions via '?'.
        # SkipLoadingFormatData avoids the 'exchange.format.ps1xml not found' temp-file
        # error hit by EXO 3.10 on pwsh 7 / macOS; the toolkit formats output itself
        # (Show-Result / Save-ResultCsv), so the display format data is not needed.
        ConnectArgs = { param($Upn) @{ UserPrincipalName = $Upn; ShowBanner = $false; LoadCmdletHelp = $true; SkipLoadingFormatData = $true } }
        Disconnect  = 'Disconnect-ExchangeOnline'
    }
    [pscustomobject]@{
        Name        = 'Entra / Identity (Graph)'
        Module      = 'Microsoft.Graph'
        ConnectCmd  = 'Connect-MgGraph'
        ConnectArgs = { param($Upn) @{ Scopes = 'User.ReadWrite.All','Group.ReadWrite.All' } }
        Disconnect  = 'Disconnect-MgGraph'
    }
    [pscustomobject]@{
        Name        = 'SharePoint Online (PnP)'
        Module      = 'PnP.PowerShell'
        ConnectCmd  = 'Connect-PnPOnline'
        ConnectArgs = { param($Upn) @{ Url = (Read-Host 'SharePoint site/admin URL'); Interactive = $true } }
        Disconnect  = 'Disconnect-PnPOnline'
    }
    [pscustomobject]@{
        Name        = 'Microsoft Teams'
        Module      = 'MicrosoftTeams'
        ConnectCmd  = 'Connect-MicrosoftTeams'
        ConnectArgs = { param($Upn) @{ AccountId = $Upn } }
        Disconnect  = 'Disconnect-MicrosoftTeams'
    }
    [pscustomobject]@{
        Name        = 'Security & Compliance (Purview)'
        Module      = 'ExchangeOnlineManagement'   # Connect-IPPSSession ships inside EXO module
        ConnectCmd  = 'Connect-IPPSSession'
        ConnectArgs = { param($Upn) @{ UserPrincipalName = $Upn } }
        Disconnect  = $null
    }
    [pscustomobject]@{
        Name        = 'SharePoint Admin (SPO)'
        Module      = 'Microsoft.Online.SharePoint.PowerShell'
        ConnectCmd  = 'Connect-SPOService'
        # Tenant admin cmdlets (Get-SPOSite, Set-SPOTenant...). Needs the -admin URL.
        ConnectArgs = { param($Upn) @{ Url = (Read-Host 'SharePoint ADMIN URL (e.g. https://shelly-admin.sharepoint.com)') } }
        Disconnect  = 'Disconnect-SPOService'
    }
    [pscustomobject]@{
        Name        = 'Intune / Devices (Graph)'
        Module      = 'Microsoft.Graph'
        ConnectCmd  = 'Connect-MgGraph'
        # Same Graph SDK as Entra, but consented for device-management scopes.
        ConnectArgs = { param($Upn) @{ Scopes = 'DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementConfiguration.ReadWrite.All' } }
        Disconnect  = 'Disconnect-MgGraph'
    }
    [pscustomobject]@{
        Name        = 'Power Platform (admin)'
        Module      = 'Microsoft.PowerApps.Administration.PowerShell'
        ConnectCmd  = 'Add-PowerAppsAccount'   # opens an interactive sign-in
        ConnectArgs = { param($Upn) @{ } }
        Disconnect  = 'Remove-PowerAppsAccount'
    }
)

# ---------------------------------------------------------------------------
# Layer 3: Plain-language action catalogue. Add an action = add one row here.
# Cmdlet is the ONLY thing you must name; parameters are discovered by layer 2.
# 'Prompt' (optional) curates which params to walk first; if omitted, the engine
# walks all Required params and then offers the full Optional/Switch lists.
# ---------------------------------------------------------------------------
$Actions = @(
    [pscustomobject]@{
        Service='Exchange Online'; Category='Give access'
        Label='Give a user access to another mailbox (incl. shared mailboxes)'
        Cmdlet='Add-MailboxPermission'; Prompt=@('Identity','User','AccessRights')
        Help='Grant ONE person the right to open/manage ANOTHER mailbox (a colleague''s or a shared mailbox). For a shared mailbox pick FullAccess; to control whether it auto-appears in their Outlook, set the optional AutoMapping parameter.'
        ParamHelp=@{
            Identity     = 'The mailbox that will BE accessed - whose inbox is opened (e.g. a shared mailbox, or a colleague)'
            User         = 'The person who will GET the access'
            AccessRights = 'What they may do. FullAccess = open & manage everything (Send As is a separate action)'
            AutoMapping  = 'Only meaningful with FullAccess. true = the mailbox appears automatically in the person''s Outlook; false = they add it manually (often preferred for large shared mailboxes)'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Remove & change access'
        Label='Remove a user''s access to a mailbox'
        Cmdlet='Remove-MailboxPermission'; Prompt=@('Identity','User','AccessRights')
        Help='Take away access that was previously granted.'
        ParamHelp=@{
            Identity     = 'The mailbox that WAS being accessed'
            User         = 'The person who will LOSE the access'
            AccessRights = 'Which right to remove (usually FullAccess)'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Give access'
        Label='Let a user send AS another mailbox (Send As)'
        Cmdlet='Add-RecipientPermission'; Prompt=@('Identity','Trustee')
        Preset=@{ AccessRights='SendAs' }
        Help='Let a person send email that appears to come FROM another mailbox.'
        ParamHelp=@{
            Identity     = 'The mailbox that mail will appear to come FROM'
            Trustee      = 'The person who will be allowed to send as it'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Check access'
        Label='See who can Send As a mailbox'
        Cmdlet='Get-RecipientPermission'; Prompt=@('Identity')
        Help='List everyone who has Send As on a mailbox - use it to verify a Send As grant.'
        ParamHelp=@{ Identity='The mailbox to check - the one others may send email AS' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Check access'
        Label='See who has access to a mailbox (FullAccess etc.)'
        Cmdlet='Get-MailboxPermission'; Prompt=@('Identity')
        Help='List who can open/manage a mailbox. NOTE: "NT AUTHORITY\SELF" and inherited rows are normal system entries - the real grants are the named users.'
        ParamHelp=@{ Identity='The mailbox to check - whose access list you want to see' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Check access'
        Label='See who has access to a calendar'
        Cmdlet='Get-MailboxFolderPermission'; Prompt=@('Identity')
        Transform=@{ Identity = { param($v) '{0}:\Calendar' -f $v } }
        Help='List who can see or edit a mailbox''s calendar, and the role each has.'
        ParamHelp=@{ Identity='The mailbox whose calendar you want to check (email or name) - the tool targets its \Calendar folder automatically' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Give access'
        Label='Give a user access to a calendar'
        Cmdlet='Add-MailboxFolderPermission'; Prompt=@('Identity','User','AccessRights')
        Transform=@{ Identity = { param($v) '{0}:\Calendar' -f $v } }
        Hints=@{ AccessRights=@('Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor','AvailabilityOnly','LimitedDetails','None') }
        Help='Grant a person a role on a mailbox''s calendar. Use this only if they have NO existing entry - otherwise use "Change a user''s calendar access".'
        ParamHelp=@{
            Identity     = 'The mailbox whose calendar is being shared (its \Calendar folder is targeted automatically)'
            User         = 'The person who will GET access to the calendar'
            AccessRights = 'Their role. Reviewer = read events; Editor = create/edit/delete; AvailabilityOnly = free/busy only; LimitedDetails = free/busy + subject & location'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Remove & change access'
        Label='Change a user''s calendar access'
        Cmdlet='Set-MailboxFolderPermission'; Prompt=@('Identity','User','AccessRights')
        Transform=@{ Identity = { param($v) '{0}:\Calendar' -f $v } }
        Hints=@{ AccessRights=@('Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor','AvailabilityOnly','LimitedDetails','None') }
        Help='Change the role of someone who ALREADY has access to a calendar (fails if they have no existing entry - use "Give a user access" first).'
        ParamHelp=@{
            Identity     = 'The mailbox whose calendar is shared (its \Calendar folder is targeted automatically)'
            User         = 'The person whose access you are changing'
            AccessRights = 'The new role (see the grant action for what each role means)'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Remove & change access'
        Label='Remove a user''s calendar access'
        Cmdlet='Remove-MailboxFolderPermission'; Prompt=@('Identity','User')
        Transform=@{ Identity = { param($v) '{0}:\Calendar' -f $v } }
        Help='Take away a person''s access to a mailbox''s calendar.'
        ParamHelp=@{
            Identity = 'The mailbox whose calendar was shared (its \Calendar folder is targeted automatically)'
            User     = 'The person who will LOSE calendar access'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Mailboxes'; Label='Look up a mailbox'
        Cmdlet='Get-EXOMailbox'; Prompt=@('Identity')
        Help='Show details about a mailbox.'
        ParamHelp=@{ Identity='The mailbox to look up (email address or name)' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Mailboxes'; Label='Change a mailbox setting'
        Cmdlet='Set-Mailbox'; Prompt=$null
        Help='Change a setting on a mailbox (advanced - walks all available options).'
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Mailboxes'
        Label='Create a new shared mailbox'
        Cmdlet='New-Mailbox'; Prompt=@('Name','DisplayName','PrimarySmtpAddress')
        Preset=@{ Shared=$true }
        Help='Create a brand-new shared mailbox (no license needed, no password).'
        ParamHelp=@{
            Name               = 'Internal name / alias, no spaces (e.g. sales)'
            DisplayName        = 'Friendly name shown to people (e.g. Sales Team)'
            PrimarySmtpAddress = 'The email address for it (e.g. sales@shelly.com)'
        }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Mailboxes'
        Label='Convert an existing mailbox to shared'
        Cmdlet='Set-Mailbox'; Prompt=@('Identity')
        Preset=@{ Type='Shared' }
        Help='Turn a normal user mailbox into a shared mailbox (e.g. after someone leaves). Free up the license afterwards.'
        ParamHelp=@{ Identity='The mailbox to convert (email address or name)' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Mailboxes'
        Label='Convert a shared mailbox back to a regular (user) mailbox'
        Cmdlet='Set-Mailbox'; Prompt=@('Identity')
        Preset=@{ Type='Regular' }
        Help='Turn a shared mailbox back into a normal user mailbox. NOTE: a regular mailbox needs an Exchange Online license assigned, or it gets disabled after ~30 days.'
        ParamHelp=@{ Identity='The shared mailbox to convert back (email address or name)' }
    }
    [pscustomobject]@{
        Service='Exchange Online'; Category='Give access'; Label='Add a member to a distribution group'
        Cmdlet='Add-DistributionGroupMember'; Prompt=@('Identity','Member')
        Help='Add someone to a distribution / mailing group.'
        ParamHelp=@{ Identity='The group to add TO'; Member='The person being ADDED' }
    }
    [pscustomobject]@{
        Service='Entra / Identity (Graph)'; Category='Users'; Label='Look up a user'
        Cmdlet='Get-MgUser'; Prompt=@('UserId')
        Help='Show details about a user account.'
        ParamHelp=@{ UserId='The user to look up (email or object id)' }
    }
    [pscustomobject]@{
        Service='Entra / Identity (Graph)'; Category='Check access'
        Label='See which groups a user is in'
        Cmdlet='Get-MgUserMemberOf'; Prompt=@('UserId')
        Preset=@{ All=$true }
        Help='List EVERY group (and directory role) a user belongs to, each labelled with what it is - Distribution group, Microsoft 365 group, Security group, Mail-enabled security, or Directory role.'
        ParamHelp=@{ UserId='The user whose memberships you want (email or object id). Type find to search the directory, or last to reuse someone from a previous result.' }
        # Get-MgUserMemberOf returns raw directory objects whose useful fields sit
        # in an AdditionalProperties bag. Flatten to clean Name / Email / Type so
        # the table is readable (and so the Type says WHAT each group is).
        Project={
            param($r)
            @($r) | ForEach-Object {
                $ap = $_.AdditionalProperties; if (-not $ap) { $ap = @{} }
                $odata = [string]$ap['@odata.type']
                $type =
                    if     ($odata -match 'directoryRole') { 'Directory role' }
                    elseif ($odata -match 'group' -or $ap.ContainsKey('groupTypes')) {
                        $gt   = @($ap['groupTypes'])
                        $mail = [bool]$ap['mailEnabled']
                        $sec  = [bool]$ap['securityEnabled']
                        if     ($gt -contains 'Unified') { 'Microsoft 365 group' }
                        elseif ($mail -and $sec)         { 'Mail-enabled security' }
                        elseif ($mail)                   { 'Distribution group' }
                        elseif ($sec)                    { 'Security group' }
                        else                             { 'Group' }
                    }
                    elseif ($odata) { $odata -replace '#microsoft\.graph\.','' }
                    else            { 'Object' }
                [pscustomobject]@{
                    Name  = [string]$ap['displayName']
                    Email = [string]$ap['mail']
                    Type  = $type
                }
            }
        }
    }
    [pscustomobject]@{
        Service='SharePoint Admin (SPO)'; Category='Sites'; Label='List all SharePoint sites'
        Cmdlet='Get-SPOSite'; Prompt=$null
        Help='List the site collections in the tenant (add -Limit All to see every one).'
    }
    [pscustomobject]@{
        Service='Intune / Devices (Graph)'; Category='Devices'; Label='List enrolled devices'
        Cmdlet='Get-MgDeviceManagementManagedDevice'; Prompt=$null
        Help='List devices enrolled in Intune.'
    }
    [pscustomobject]@{
        Service='Power Platform (admin)'; Category='Environments'; Label='List Power Platform environments'
        Cmdlet='Get-AdminPowerAppEnvironment'; Prompt=$null
        Help='List the Power Platform environments in the tenant.'
    }
)

# ---------------------------------------------------------------------------
# Optional value hints. Some cmdlets have a fixed set of allowed values but do
# not advertise it in a way the engine can read - e.g. Exchange Online proxy
# cmdlets expose -AccessRights as a plain Object[] with no ValidateSet or enum.
# List those values here and the tool will offer them. Keyed by parameter name
# (case-insensitive). Teaching the tool a new field = adding one row.
# ---------------------------------------------------------------------------
$ParamHints = @{
    AccessRights = @('FullAccess','ExternalAccount','DeleteItem','ReadPermission','ChangePermission','ChangeOwner')
}

# ---------------------------------------------------------------------------
# Optional target lookups. For identity-style parameters the user can type
# 'find' to search the directory and pick from a list instead of typing an
# address. Each entry: Probe = a cmdlet whose presence means the lookup can run
# (i.e. you're connected to the right service); Run = { param($term) ... }
# returning objects with .Display (what to show) and .Value (what to submit).
# ---------------------------------------------------------------------------
$recipientLookup = @{
    Probe = 'Get-EXORecipient'
    Run   = {
        param($term)
        Get-EXORecipient -ResultSize 25 -Filter "DisplayName -like '*$term*' -or PrimarySmtpAddress -like '*$term*'" -ErrorAction Stop |
            ForEach-Object { [pscustomobject]@{ Display = ('{0}  <{1}>  [{2}]' -f $_.DisplayName, $_.PrimarySmtpAddress, $_.RecipientTypeDetails); Value = [string]$_.PrimarySmtpAddress } }
    }
    # Resolve an EXACT typed value to a real recipient (or $null if it doesn't exist).
    Resolve = { param($v) Get-EXORecipient -Identity $v -ErrorAction SilentlyContinue | Select-Object -First 1 }
}
$ParamLookup = @{
    Identity = $recipientLookup
    User     = $recipientLookup
    Trustee  = $recipientLookup
    Member   = $recipientLookup
    Mailbox  = $recipientLookup
    UserId   = @{
        Probe = 'Get-MgUser'
        Run   = {
            param($term)
            Get-MgUser -Filter "startswith(displayName,'$term') or startswith(mail,'$term') or startswith(userPrincipalName,'$term')" -Top 25 -ErrorAction Stop |
                ForEach-Object { [pscustomobject]@{ Display = ('{0}  <{1}>' -f $_.DisplayName, $_.UserPrincipalName); Value = [string]$_.UserPrincipalName } }
        }
        Resolve = { param($v) Get-MgUser -UserId $v -ErrorAction SilentlyContinue }
    }
}

# ---------------------------------------------------------------------------
# Plain-language glossary for common parameters that actions don't describe
# themselves - used by the '?' explain option in the optional/switch lists.
# Keyed by parameter name. Add a row = teach the tool about a parameter.
# ---------------------------------------------------------------------------
$ParamDocs = @{
    AutoMapping     = "Whether the mailbox auto-appears in the person's Outlook. Default `$true; set `$false for large/shared mailboxes so they add it manually."
    InheritanceType = 'How the permission propagates on the AD object. Leave the default (All) - almost never changed.'
    Identity        = 'The object the command acts on - a mailbox, user, or group (email address or name).'
    Confirm         = 'Ask for confirmation before running.'
    Force           = 'Skip confirmation prompts and proceed.'
    ResultSize      = 'How many results to return (Unlimited for all).'
    WhatIf          = 'Show what WOULD happen without actually doing it.'
    Deny            = 'Creates a DENY entry - blocks the right instead of granting it. Leave off for a normal grant.'
    GroupMailbox    = 'Allow this to target a Microsoft 365 Group mailbox (normally blocked). Rarely needed.'
    IgnoreDefaultScope = 'Search the whole tenant instead of the session''s default scope when resolving the mailbox. Advanced - only if a lookup fails due to scope.'
}

# ---------------------------------------------------------------------------
# Small UI helper: print a numbered menu and return the chosen item (or $null).
# ---------------------------------------------------------------------------
function Get-MenuChoice {
    param(
        [Parameter(Mandatory)] [string]   $Title,
        [Parameter(Mandatory)] [string[]] $Options,
        [string[]] $Descriptions,   # optional, parallel to $Options: shown when the user types ?
        [string[]] $ExtraOptions,   # app-level controls (Search / Switch service / Quit ...): shown below a divider, still numbered & selectable
        [switch] $AllowBack
    )
    $all = @($Options) + @($ExtraOptions)   # combined for selection; numbering runs continuously
    Write-Host ''
    Write-Host $Title -ForegroundColor White
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host ('  {0,2}. {1}' -f ($i + 1), $Options[$i])
    }
    if ($ExtraOptions) {
        # App-level controls, kept separate from the menu's own items.
        Write-Host '   -------- app --------' -ForegroundColor White
        for ($j = 0; $j -lt $ExtraOptions.Count; $j++) {
            Write-Host ('  {0,2}. {1}' -f ($Options.Count + $j + 1), $ExtraOptions[$j]) -ForegroundColor White
        }
    }
    if ($AllowBack) { Write-Host '   0. Back' -ForegroundColor White }
    if ($Descriptions) { Write-Host '   (type ? for a short explanation of each choice)' -ForegroundColor White }
    while ($true) {
        $raw = Read-Host 'Choose a number'
        if ($raw -eq '?' -and $Descriptions) {
            for ($i = 0; $i -lt $Options.Count; $i++) {
                Write-Host ('  {0,2}. {1}' -f ($i + 1), $Options[$i]) -ForegroundColor White
                $d = if ($i -lt $Descriptions.Count) { $Descriptions[$i] } else { '' }
                if ($d) { Write-Host ('        {0}' -f $d) -ForegroundColor White }
            }
            continue
        }
        if ($AllowBack -and $raw -eq '0') { return $null }
        if ($raw -match '^\d+$' -and [int]$raw -ge 1 -and [int]$raw -le $all.Count) {
            return $all[[int]$raw - 1]
        }
        Write-Host 'Not a valid choice, try again.' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# Layer 1 driver: connect to a chosen service using the entered email.
# ---------------------------------------------------------------------------
function Connect-M365Service {
    param(
        [Parameter(Mandatory)] [pscustomobject] $Service,
        [Parameter(Mandatory)] [string]         $Upn
    )
    if (-not (Get-Module -ListAvailable -Name $Service.Module)) {
        Write-Host ("Module '{0}' is not installed." -f $Service.Module) -ForegroundColor Yellow
        if ((Read-Host 'Install it now for the current user? (y/n)') -match '^(y|yes)$') {
            Install-Module $Service.Module -Scope CurrentUser -Force
        } else {
            throw "Cannot connect without module '$($Service.Module)'."
        }
    }
    Import-Module $Service.Module -ErrorAction SilentlyContinue
    $connectArgs = & $Service.ConnectArgs $Upn
    Write-Host ("Connecting to {0} as {1} ..." -f $Service.Name, $Upn) -ForegroundColor White
    & $Service.ConnectCmd @connectArgs
    Write-Host ("Connected to {0}." -f $Service.Name) -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Render an assembled splat as a copy-pasteable command line (for the preview).
# ---------------------------------------------------------------------------
function Format-CommandPreview {
    param(
        [Parameter(Mandatory)] [string]    $Cmdlet,
        [Parameter(Mandatory)] [hashtable] $BoundParams
    )
    $sb = [System.Text.StringBuilder]::new($Cmdlet)
    foreach ($key in $BoundParams.Keys) {
        $val = $BoundParams[$key]
        if ($val -is [switch]) {
            if ($val.IsPresent) { [void]$sb.Append(" -$key") }
        } elseif ($val -is [bool]) {
            [void]$sb.Append(" -$key `$$($val.ToString().ToLower())")
        } elseif ($val -is [array]) {
            [void]$sb.Append(" -$key " + (($val | ForEach-Object { "'$_'" }) -join ','))
        } elseif ("$val" -match '\s') {
            [void]$sb.Append(" -$key '$val'")
        } else {
            [void]$sb.Append(" -$key $val")
        }
    }
    $sb.ToString()
}

# ---------------------------------------------------------------------------
# Shared introspection: read a cmdlet's real parameters and sort them into
# Required / Optional / Switches (dropping PowerShell's common params). Used by
# both the guided (one-at-a-time) and bulk (CSV) runners so they stay in sync.
# ---------------------------------------------------------------------------
function Get-CmdletParameters {
    param(
        [Parameter(Mandatory)] $Cmd,
        [string[]] $PreferParams   # hint: params the action intends to use, to pick the right set
    )
    $common = [System.Management.Automation.PSCmdlet]::CommonParameters +
              [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
    $all = $Cmd.Parameters.Values | Where-Object { $common -notcontains $_.Name }

    # Many cmdlets have several mutually-exclusive PARAMETER SETS (e.g.
    # Add-MailboxPermission has an AccessRights set AND an Owner set). A param
    # that is mandatory in one set is optional/absent in another, so we must
    # resolve ONE set and classify against it - otherwise unrelated params
    # (like -Owner) wrongly show as required. Pick the set that best matches the
    # action's intended params, else the cmdlet's default set.
    $setNames = @($Cmd.ParameterSets | ForEach-Object { $_.Name })
    $chosen = if ($setNames.Count -le 1) { $setNames | Select-Object -First 1 } else {
        $best = $null
        if ($PreferParams) {
            $best = ($Cmd.ParameterSets | Sort-Object `
                @{ Expression = { @($_.Parameters | Where-Object { $PreferParams -contains $_.Name }).Count }; Descending = $true }, `
                @{ Expression = { [int]$_.IsDefault }; Descending = $true } | Select-Object -First 1).Name
        }
        if (-not $best) { $best = ($Cmd.ParameterSets | Where-Object { $_.IsDefault } | Select-Object -First 1).Name }
        if (-not $best) { $best = $setNames[0] }
        $best
    }

    # Keep only params that belong to the chosen set (or to all sets).
    if ($chosen) {
        $all = $all | Where-Object { $_.ParameterSets.ContainsKey($chosen) -or $_.ParameterSets.ContainsKey('__AllParameterSets') }
    }

    # Mandatory-ness is per-set: read the chosen set's entry (or the all-sets entry).
    $isMandatory = {
        param($p)
        if     ($chosen -and $p.ParameterSets.ContainsKey($chosen))     { return $p.ParameterSets[$chosen].IsMandatory }
        elseif ($p.ParameterSets.ContainsKey('__AllParameterSets'))     { return $p.ParameterSets['__AllParameterSets'].IsMandatory }
        return (@($p.ParameterSets.Values | Where-Object { $_.IsMandatory }).Count -gt 0)
    }

    [pscustomobject]@{
        SetName  = $chosen
        Required = @($all | Where-Object { & $isMandatory $_ })
        Switches = @($all | Where-Object { -not (& $isMandatory $_) -and $_.SwitchParameter })
        Optional = @($all | Where-Object { -not (& $isMandatory $_) -and -not $_.SwitchParameter })
    }
}

# ---------------------------------------------------------------------------
# Best plain-language description for a parameter, in order: the action's own
# ParamHelp, the $ParamDocs glossary, then the cmdlet's built-in help. $null if
# nothing is known. Used by the '?' explain option in the parameter lists.
# ---------------------------------------------------------------------------
function Get-ParamDescription {
    param([string] $Name, [hashtable] $ParamHelp, [string] $Cmdlet)
    if ($ParamHelp -and $ParamHelp[$Name]) { return $ParamHelp[$Name] }
    if ($ParamDocs.ContainsKey($Name))     { return $ParamDocs[$Name] }
    try {
        $h = Get-Help $Cmdlet -Parameter $Name -ErrorAction SilentlyContinue
        if ($h -and $h.description) {
            $txt = ((($h.description | ForEach-Object { $_.Text }) -join ' ') -replace '\s+', ' ').Trim()
            if ($txt) { return $txt }
        }
    } catch { }
    return $null
}

# ---------------------------------------------------------------------------
# Layer 2: the generic introspection engine. Given a cmdlet name it discovers
# the real parameters, categorises them, prompts in plain language, previews,
# and runs. Works for ANY cmdlet in ANY connected module - no per-cmdlet code.
# ---------------------------------------------------------------------------
function Invoke-GuidedCmdlet {
    param(
        [Parameter(Mandatory)] [string]   $Cmdlet,
        [string[]]  $Prompt,     # curated params to always walk (required OR optional), in order
        [string]    $Help,       # optional one-line plain description of the whole action
        [hashtable] $ParamHelp,  # optional per-parameter plain descriptions (name -> text)
        [hashtable] $Preset,     # optional param values fixed by the action (never prompted)
        [hashtable] $Hints,      # optional per-action allowed-value overrides (name -> values)
        [hashtable] $Transform,  # optional post-processing of collected values (name -> scriptblock)
        [scriptblock] $Project   # optional reshaping of the RESULT before display (raw result -> clean objects)
    )

    $cmd      = Get-Command $Cmdlet -ErrorAction Stop
    $parts    = Get-CmdletParameters -Cmd $cmd -PreferParams $Prompt
    $required = $parts.Required
    $switches = $parts.Switches
    $optional = $parts.Optional

    Write-Host ''
    Write-Host ("Building: {0}" -f $Cmdlet) -ForegroundColor White
    if ($Help) { Write-Host ("  {0}" -f $Help) -ForegroundColor White }
    Write-Host ("  {0} required, {1} optional, {2} switches" -f $required.Count, $optional.Count, $switches.Count) -ForegroundColor White
    if ($parts.SetName -and $parts.SetName -ne '__AllParameterSets') {
        Write-Host ("  (using parameter set: {0})" -f $parts.SetName) -ForegroundColor White
    }

    $bound = [ordered]@{}

    # ---- Presets: values fixed by the action itself (e.g. -Shared for a shared
    #      mailbox, -Type Shared to convert). Bound automatically, shown in the
    #      preview, never prompted for.
    if ($Preset) {
        foreach ($k in $Preset.Keys) { $bound[$k] = $Preset[$k] }
        Write-Host ('  Fixed by this action: ' + (($Preset.Keys | ForEach-Object { "-$_ $($Preset[$_])" }) -join ', ')) -ForegroundColor White
    }
    $promptList = @($Prompt)   # params to always walk, even if optional

    # ---- Required params: always prompt, in curated order, skipping any preset.
    $reqOrdered = if ($Prompt) {
        @($Prompt | ForEach-Object { $n = $_; $required | Where-Object { $_.Name -eq $n } }) +
        @($required | Where-Object { $Prompt -notcontains $_.Name })
    } else { $required }

    # Required params: prompt in curated order. Typing 'back' steps to the
    # PREVIOUS field (discarding its value so it can be re-entered); 'back' on
    # the first field cancels the action and returns to the menu.
    $reqList = @($reqOrdered | Where-Object { $_ -and -not $bound.Contains($_.Name) })
    $ri = 0
    while ($ri -lt $reqList.Count) {
        $p = $reqList[$ri]
        $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
        $val = Read-ParameterValue -Param $p -Required -Description $desc -Hints $Hints
        if ($script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
        if ($script:ParamBack) {
            if ($ri -eq 0) { Write-Host "  Already at the first field - type 'cancel' to return to the menu." -ForegroundColor White; continue }
            $ri--
            $bound.Remove($reqList[$ri].Name)
            Write-Host ("  Going back to -{0}." -f $reqList[$ri].Name) -ForegroundColor White
            continue
        }
        $bound[$p.Name] = $val
        $ri++
    }

    # ---- Optional params named in the curated list: ask them outright.
    foreach ($p in ($optional | Where-Object { $promptList -contains $_.Name -and -not $bound.Contains($_.Name) })) {
        $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
        $v = Read-ParameterValue -Param $p -Description $desc -Hints $Hints
        if ($script:ParamBack -or $script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
        if ($null -ne $v) { $bound[$p.Name] = $v }
    }

    # ---- Remaining optional params: offer the list, let the user pick which to set.
    $optRest = @($optional | Where-Object { -not $bound.Contains($_.Name) })
    if ($optRest.Count -gt 0) {
        Write-Host ''
        Write-Host 'Optional parameters you can set:' -ForegroundColor White
        for ($i = 0; $i -lt $optRest.Count; $i++) {
            Write-Host ('  {0,2}. {1}  ({2})' -f ($i+1), $optRest[$i].Name, $optRest[$i].ParameterType.Name)
        }
        Write-Host '   (type ? to explain each one)' -ForegroundColor White
        $pick = $null
        while ($true) {
            $pick = Read-Host 'Enter numbers to set (comma-separated), ? to explain, or Enter to skip'
            if ($pick -ne '?') { break }
            for ($i = 0; $i -lt $optRest.Count; $i++) {
                Write-Host ('  {0,2}. {1}' -f ($i+1), $optRest[$i].Name) -ForegroundColor White
                $d = Get-ParamDescription -Name $optRest[$i].Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
                if ($d) { Write-Host ('        {0}' -f $d) -ForegroundColor White } else { Write-Host '        (no description available)' -ForegroundColor White }
            }
        }
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })) {
            $idx = [int]$n - 1
            if ($idx -ge 0 -and $idx -lt $optRest.Count) {
                $p = $optRest[$idx]
                $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
                $bound[$p.Name] = Read-ParameterValue -Param $p -Description $desc -Hints $Hints
                if ($script:ParamBack -or $script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
            }
        }
    }

    # ---- Switches: offer the list, chosen ones become $true (skip any preset).
    $swRest = @($switches | Where-Object { -not $bound.Contains($_.Name) })
    if ($swRest.Count -gt 0) {
        Write-Host ''
        Write-Host 'Switches you can enable:' -ForegroundColor White
        for ($i = 0; $i -lt $swRest.Count; $i++) {
            Write-Host ('  {0,2}. {1}' -f ($i+1), $swRest[$i].Name)
        }
        Write-Host '   (type ? to explain each one)' -ForegroundColor White
        $pick = $null
        while ($true) {
            $pick = Read-Host 'Enter numbers to enable (comma-separated), ? to explain, or Enter to skip'
            if ($pick -ne '?') { break }
            for ($i = 0; $i -lt $swRest.Count; $i++) {
                Write-Host ('  {0,2}. {1}' -f ($i+1), $swRest[$i].Name) -ForegroundColor White
                $d = Get-ParamDescription -Name $swRest[$i].Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
                if ($d) { Write-Host ('        {0}' -f $d) -ForegroundColor White } else { Write-Host '        (no description available)' -ForegroundColor White }
            }
        }
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })) {
            $idx = [int]$n - 1
            if ($idx -ge 0 -and $idx -lt $swRest.Count) { $bound[$swRest[$idx].Name] = $true }
        }
    }

    # ---- Action transforms: post-process a collected value before it is used
    #      (e.g. append ':\Calendar' to a mailbox identity so a folder-permission
    #      cmdlet targets that mailbox's calendar). Applied just before preview so
    #      the dry-run shows the exact value that will be submitted.
    if ($Transform) {
        foreach ($tk in $Transform.Keys) {
            if ($bound.Contains($tk)) { $bound[$tk] = & $Transform[$tk] $bound[$tk] }
        }
    }

    # ---- Dry-run preview + final confirm. -----------------------------------
    $preview = Format-CommandPreview -Cmdlet $Cmdlet -BoundParams $bound
    Write-Host ''
    Write-Host 'Preview (dry-run):' -ForegroundColor Yellow
    Write-Host ('  ' + $preview) -ForegroundColor White

    if ($DryRunOnly) {
        Write-Host 'DryRunOnly is set - not executing.' -ForegroundColor White
        return
    }

    # Last confirmation: repeat the EXACT command, then require an explicit y.
    Write-Host ''
    Write-Host 'FINAL CONFIRMATION - about to execute this command:' -ForegroundColor Yellow
    Write-Host ('  ' + $preview) -ForegroundColor White
    if ((Read-Host 'Execute it now? (y/n)') -notmatch '^(y|yes)$') {
        Write-Host 'Cancelled - nothing was run.' -ForegroundColor White
        return
    }

    Write-Host 'Running...' -ForegroundColor White
    try {
        $result = & $Cmdlet @bound -ErrorAction Stop
        Write-Host 'Done.' -ForegroundColor Green
        # Optional per-action projection: reshape a raw/cluttered result into clean,
        # readable columns (e.g. flatten Get-MgUserMemberOf into Name/Email/Type)
        # before it is shown AND before it becomes the pipe source for 'last'.
        if ($Project) {
            try   { $result = & $Project $result }
            catch { Write-Host ('  (could not reshape result: {0})' -f $_.Exception.Message) -ForegroundColor Yellow }
        }
        Show-Result -Result $result -Cmdlet $Cmdlet
    } catch {
        Write-Host ('Failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------------
# Bulk runner: run the SAME cmdlet once per row of a CSV file. You map each
# parameter to a CSV column (per-row value) or a fixed value (same every row),
# preview the first row, then it runs them all with a per-row success/fail tally.
# ---------------------------------------------------------------------------
function Invoke-BulkCmdlet {
    param(
        [Parameter(Mandatory)] [string] $Cmdlet,
        [string[]]  $Prompt,
        [hashtable] $ParamHelp,
        [hashtable] $Preset,
        [hashtable] $Transform   # optional post-processing of each row's values (name -> scriptblock)
    )
    $cmd   = Get-Command $Cmdlet -ErrorAction Stop
    $parts = Get-CmdletParameters -Cmd $cmd -PreferParams $Prompt

    $csvPath = (Read-Host 'Path to CSV file (must have a header row)').Trim('"',' ')
    if (-not (Test-Path $csvPath)) { Write-Host "File not found: $csvPath" -ForegroundColor Red; return }
    $rows = @(Import-Csv -Path $csvPath)
    if ($rows.Count -eq 0) { Write-Host 'That CSV has no data rows.' -ForegroundColor Yellow; return }
    $columns = @($rows[0].PSObject.Properties.Name)

    Write-Host ''
    Write-Host ("Loaded {0} rows. Columns: {1}" -f $rows.Count, ($columns -join ', ')) -ForegroundColor White
    Write-Host 'For each parameter: type a COLUMN name (uses each row''s value), or =value for a fixed value on every row, or press Enter to skip.' -ForegroundColor White

    # map: paramName -> @{ Kind = 'column' | 'fixed' | 'switch'; Value = ... }
    $map = [ordered]@{}
    if ($Preset) { foreach ($k in $Preset.Keys) { $map[$k] = @{ Kind='fixed'; Value=$Preset[$k] } } }

    $askParam = {
        param($p, $required)
        if ($map.Contains($p.Name)) { return }
        $desc = if ($ParamHelp) { $ParamHelp[$p.Name] } else { $null }
        if ($desc) { Write-Host ("  -> {0}" -f $desc) -ForegroundColor White }
        while ($true) {
            $ans = Read-Host ("Map -{0} (column / =fixed / Enter){1}" -f $p.Name, $(if ($required) { ' [required]' } else { '' }))
            if ([string]::IsNullOrWhiteSpace($ans)) {
                if ($required) { Write-Host '  Required - give a column name or =value.' -ForegroundColor Yellow; continue }
                return
            }
            if ($ans.StartsWith('=')) { $map[$p.Name] = @{ Kind='fixed'; Value=$ans.Substring(1) }; return }
            if ($columns -contains $ans) { $map[$p.Name] = @{ Kind='column'; Value=$ans }; return }
            Write-Host ("  '{0}' is not a column. Available: {1}" -f $ans, ($columns -join ', ')) -ForegroundColor Yellow
        }
    }

    # Required first, in curated order.
    $reqOrdered = if ($Prompt) {
        @($Prompt | ForEach-Object { $n=$_; $parts.Required | Where-Object { $_.Name -eq $n } }) +
        @($parts.Required | Where-Object { $Prompt -notcontains $_.Name })
    } else { $parts.Required }
    foreach ($p in ($reqOrdered | Where-Object { $_ })) { & $askParam $p $true }

    # Optional: offer the list.
    $optRest = @($parts.Optional | Where-Object { -not $map.Contains($_.Name) })
    if ($optRest.Count -gt 0) {
        Write-Host ''
        Write-Host 'Optional parameters you can map:' -ForegroundColor White
        for ($i=0; $i -lt $optRest.Count; $i++) { Write-Host ('  {0,2}. {1}' -f ($i+1), $optRest[$i].Name) }
        $pick = Read-Host 'Enter numbers to map (comma-separated), or Enter to skip'
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })) {
            $idx=[int]$n-1; if ($idx -ge 0 -and $idx -lt $optRest.Count) { & $askParam $optRest[$idx] $false }
        }
    }

    # Switches: enabled for every row.
    $swRest = @($parts.Switches | Where-Object { -not $map.Contains($_.Name) })
    if ($swRest.Count -gt 0) {
        Write-Host ''
        Write-Host 'Switches to enable for every row:' -ForegroundColor White
        for ($i=0; $i -lt $swRest.Count; $i++) { Write-Host ('  {0,2}. {1}' -f ($i+1), $swRest[$i].Name) }
        $pick = Read-Host 'Enter numbers to enable, or Enter to skip'
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })) {
            $idx=[int]$n-1; if ($idx -ge 0 -and $idx -lt $swRest.Count) { $map[$swRest[$idx].Name] = @{ Kind='switch'; Value=$true } }
        }
    }

    $buildRow = {
        param($row)
        $b = [ordered]@{}
        foreach ($k in $map.Keys) {
            $m = $map[$k]
            switch ($m.Kind) {
                'column' { $b[$k] = $row.$($m.Value) }
                'fixed'  { $b[$k] = $m.Value }
                'switch' { $b[$k] = $true }
            }
        }
        if ($Transform) {
            foreach ($tk in $Transform.Keys) {
                if ($b.Contains($tk)) { $b[$tk] = & $Transform[$tk] $b[$tk] }
            }
        }
        $b
    }

    $preview = Format-CommandPreview -Cmdlet $Cmdlet -BoundParams (& $buildRow $rows[0])
    Write-Host ''
    Write-Host ('Preview - first of {0} rows (dry-run):' -f $rows.Count) -ForegroundColor Yellow
    Write-Host ('  ' + $preview) -ForegroundColor White

    if ($DryRunOnly) { Write-Host 'DryRunOnly is set - not executing.' -ForegroundColor White; return }

    # Last confirmation before running every row.
    Write-Host ''
    Write-Host ('FINAL CONFIRMATION - about to run this for ALL {0} rows, e.g.:' -f $rows.Count) -ForegroundColor Yellow
    Write-Host ('  ' + $preview) -ForegroundColor White
    if ((Read-Host ("Execute for all {0} rows now? (y/n)" -f $rows.Count)) -notmatch '^(y|yes)$') {
        Write-Host 'Cancelled - nothing was run.' -ForegroundColor White; return
    }

    $ok=0; $fail=0; $i=0
    foreach ($row in $rows) {
        $i++
        $splat = & $buildRow $row
        try {
            & $Cmdlet @splat | Out-Null
            $ok++;  Write-Host ("  [{0}/{1}] OK" -f $i, $rows.Count) -ForegroundColor Green
        } catch {
            $fail++; Write-Host ("  [{0}/{1}] FAILED: {2}" -f $i, $rows.Count, $_.Exception.Message) -ForegroundColor Red
        }
    }
    Write-Host ''
    Write-Host ("Bulk done: {0} succeeded, {1} failed, {2} total." -f $ok, $fail, $rows.Count) -ForegroundColor White
}

# ---------------------------------------------------------------------------
# Numbered picker for a fixed set of allowed values. -Multi lets the user pick
# several (comma-separated); otherwise one. Typing 'back' / 0 sets ParamBack,
# 'cancel' sets ParamCancel (the caller decides what those mean). Returns the
# chosen value (string) or values (string[]), or $null when the user backs out.
# ---------------------------------------------------------------------------
function Read-OptionMenu {
    param(
        [Parameter(Mandatory)] [string]   $Name,
        [Parameter(Mandatory)] [string[]] $Options,
        [switch] $Multi,
        [string] $Description
    )
    Write-Host ''
    Write-Host (('Choose {0} for -{1}:' -f $(if ($Multi) { 'one or more' } else { 'a value' }), $Name)) -ForegroundColor White
    for ($i = 0; $i -lt $Options.Count; $i++) { Write-Host ('  {0,2}. {1}' -f ($i + 1), $Options[$i]) }
    Write-Host "   0. Back to previous field   (or type 'cancel' to quit to the menu)" -ForegroundColor White
    $hintNums = if ($Multi) { 'numbers (e.g. 1,3)' } else { 'a number' }
    while ($true) {
        $pick = Read-Host ('Enter {0}, 0 to go back, or ? to explain' -f $hintNums)
        $k = ("$pick").Trim().ToLower()
        if ($k -eq '?') {
            if ($Description) { Write-Host ('  {0}' -f $Description) -ForegroundColor White }
            Write-Host ('  Options: ' + ($Options -join ', ')) -ForegroundColor White
            continue
        }
        if ($k -eq 'back' -or $k -eq '0') { $script:ParamBack   = $true; return $null }
        if ($k -eq 'cancel')              { $script:ParamCancel = $true; return $null }
        $nums = @($k -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })
        if ($nums.Count -eq 0) { Write-Host ('  Type {0} from the list.' -f $hintNums) -ForegroundColor Yellow; continue }
        if (-not $Multi -and $nums.Count -gt 1) { Write-Host '  Just one number for this field.' -ForegroundColor Yellow; continue }
        $sel = @(); $bad = $false
        foreach ($n in $nums) {
            $idx = [int]$n - 1
            if ($idx -ge 0 -and $idx -lt $Options.Count) { $sel += $Options[$idx] }
            else { Write-Host ('  {0} is not in the list (1-{1}).' -f $n, $Options.Count) -ForegroundColor Yellow; $bad = $true; break }
        }
        if ($bad) { continue }
        $sel = @($sel | Select-Object -Unique)
        Write-Host ('  -> ' + ($sel -join ', ')) -ForegroundColor White
        if ($Multi) { return $sel } else { return $sel[0] }
    }
}

# ---------------------------------------------------------------------------
# Prompt for a single parameter value, in plain language. ValidateSet params
# are offered as a numbered pick; everything else gets a friendly type hint.
# ---------------------------------------------------------------------------
function Read-ParameterValue {
    param(
        [Parameter(Mandatory)] $Param,
        [switch] $Required,
        [string] $Description,  # plain-language "what is this field?" shown above the prompt
        [hashtable] $Hints      # per-action allowed-value override (name -> values); beats global $ParamHints
    )
    $script:ParamBack   = $false  # cleared each prompt; set if the user types 'back'   (go to previous field)
    $script:ParamCancel = $false  # cleared each prompt; set if the user types 'cancel' (abort action to menu)
    $name    = $Param.Name
    $type    = $Param.ParameterType
    $isArray = $type.IsArray
    $elem    = if ($isArray) { $type.GetElementType() } else { $type }

    # Work out whether this field has a fixed set of allowed values, from (in order):
    # a [ValidateSet], the values of an enum type, or our manual $ParamHints table
    # (needed for Exchange proxy cmdlets that hide their real type as Object[]).
    $allowed = $null
    $validate = $Param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | Select-Object -First 1
    if     ($validate)                              { $allowed = [string[]]$validate.ValidValues }
    elseif ($elem.IsEnum)                           { $allowed = [string[]][enum]::GetNames($elem) }
    elseif ($Hints -and $Hints.ContainsKey($name))  { $allowed = [string[]]$Hints[$name] }
    elseif ($ParamHints.ContainsKey($name))         { $allowed = [string[]]$ParamHints[$name] }

    # Say what the field means before asking for it.
    if ($Description) { Write-Host ("  -> {0}" -f $Description) -ForegroundColor White }

    # Fixed set of allowed values -> numbered picker (multi-select for arrays).
    # Both support 'back'/'cancel' via the shared helper.
    if ($allowed) {
        return Read-OptionMenu -Name $name -Options $allowed -Multi:$isArray -Description $Description
    }

    # Otherwise a prompt line with a human-friendly hint.
    $hint = if ($allowed) {
        'one or more of: ' + ($allowed -join ', ')
    } else {
        switch -Regex ($type.Name) {
            'Boolean'               { 'true / false' }
            'Int|Double|Long'       { 'a number' }
            'DateTime'              { 'a date, e.g. 2026-07-31' }
            'String\[\]|Object\[\]' { 'one or more values, comma-separated' }
            'String'                { 'text (e.g. an email address or name)' }
            default                 { $type.Name }
        }
    }
    # Directory lookup: offer 'find' when a lookup is registered for this param
    # AND its probe cmdlet is loaded (i.e. we're connected to the right service).
    $lookup  = $ParamLookup[$name]
    $canFind = $lookup -and (Get-Command $lookup.Probe -ErrorAction SilentlyContinue)
    # Piping: offer 'last' when the previous action returned objects. Scalar
    # fields only in v1 (array fields keep their comma-separated free entry).
    $hasLast = (@($script:LastResult).Count -gt 0) -and (-not $isArray)

    $label = ('Enter -{0} [{1}]{2}' -f $name, $hint, $(if ($Required) { ' (required)' } else { '' }))
    if ($allowed) { $label += ' - type ? to list options' }
    if ($canFind) { $label += " - or type 'find' to search" }
    if ($hasLast) { $label += " - or 'last' to reuse a value from the last result" }
    $label += " - or 'back' (previous field) / 'cancel' (menu)"

    while ($true) {
        $val = Read-Host $label
        # Normalise for keyword detection (trim + lower-case) so 'Find', ' find ' etc. still work.
        $key = ("$val").Trim().ToLower()
        if ($key -eq '?') {
            if ($Description) { Write-Host ('  {0}' -f $Description) -ForegroundColor White }
            if     ($allowed) { Write-Host ('  Options: ' + ($allowed -join ', ')) -ForegroundColor White }
            elseif ($canFind) { Write-Host "  Enter an email address or name - or type 'find' to search the directory." -ForegroundColor White }
            else              { Write-Host '  Free text - type a value.' -ForegroundColor White }
            continue
        }
        # 'back' = re-do the previous field; 'cancel' = abort the whole action.
        # The caller decides what each means in context (see Build-CommandFromCmdlet).
        if ($key -eq 'back')   { $script:ParamBack   = $true; return $null }
        if ($key -eq 'cancel') { $script:ParamCancel = $true; return $null }
        if ($key -eq 'find' -and $canFind) {
            $term = Read-Host '  Search text (name or email)'
            if ([string]::IsNullOrWhiteSpace($term)) { continue }
            try { $hits = @(& $lookup.Run $term) }
            catch { Write-Host ('  Lookup failed: {0}' -f $_.Exception.Message) -ForegroundColor Yellow; continue }
            if ($hits.Count -eq 0) { Write-Host '  No matches - try again, or just type the value.' -ForegroundColor Yellow; continue }
            $pick = Get-MenuChoice -Title '  Pick one:' -Options @($hits.Display) -AllowBack
            if (-not $pick) { continue }
            $sel = $hits | Where-Object { $_.Display -eq $pick } | Select-Object -First 1
            Write-Host ('  -> using {0}' -f $sel.Value) -ForegroundColor White
            return $sel.Value
        }
        # 'last' = pipe a value out of the previous action's result. Pick a
        # column (defaulting to the first identity-like one), then pick a row;
        # returns that cell as a scalar, mirroring the 'find' branch above.
        if ($key -eq 'last' -and $hasLast) {
            $rows = @($script:LastResult)
            $cols = @($rows[0].PSObject.Properties.Name)
            $idCols = @('PrimarySmtpAddress','UserPrincipalName','Mail','Email','Identity','Id','WindowsEmailAddress')
            $col = $idCols | Where-Object { $cols -contains $_ } | Select-Object -First 1
            if (-not $col) {
                $col = Get-MenuChoice -Title '  Which value from the last result?' -Options $cols -AllowBack
                if (-not $col) { continue }
            }
            # Show a second column (DisplayName/Name) for context when present.
            $ctxCol = @('DisplayName','Name') | Where-Object { $cols -contains $_ -and $_ -ne $col } | Select-Object -First 1
            $choices = foreach ($r in $rows) {
                $v = [string]$r.$col
                if ([string]::IsNullOrWhiteSpace($v)) { continue }
                $disp = if ($ctxCol) { '{0}  ({1})' -f $v, [string]$r.$ctxCol } else { $v }
                [pscustomobject]@{ Display = $disp; Value = $v }
            }
            $choices = @($choices)
            if ($choices.Count -eq 0) {
                Write-Host ("  The last result has no '{0}' values to use." -f $col) -ForegroundColor Yellow; continue
            }
            $pick = Get-MenuChoice -Title ("  Pick a {0} from the last result:" -f $col) -Options @($choices.Display) -AllowBack
            if (-not $pick) { continue }
            $sel = $choices | Where-Object { $_.Display -eq $pick } | Select-Object -First 1
            Write-Host ('  -> using {0}' -f $sel.Value) -ForegroundColor White
            return $sel.Value
        }
        if ([string]::IsNullOrWhiteSpace($val) -and $Required) {
            Write-Host '  This one is required.' -ForegroundColor Yellow
            continue
        }
        if ([string]::IsNullOrWhiteSpace($val)) { return $null }
        # Identity-style fields: verify the typed value is a real recipient/user.
        # Catches nonsense/typos instead of silently accepting them as an address.
        if ($canFind -and $lookup.Resolve -and -not $isArray) {
            $match = $null
            try { $match = & $lookup.Resolve $val } catch { }
            if (-not $match) {
                Write-Host ("  No mailbox/user found matching '{0}'. Check the spelling, or type 'find' to search." -f $val) -ForegroundColor Yellow
                continue
            }
        }
        if ($isArray) {
            $vals = @($val -split '[,\s]+' | Where-Object { $_ })
            if ($allowed) {
                # -contains / -eq are case-insensitive for strings, so this catches
                # typos like 'Send' (vs 'SendAs') and normalises casing to canonical.
                $bad = @($vals | Where-Object { $allowed -notcontains $_ })
                if ($bad.Count -gt 0) {
                    Write-Host ("  Not valid: {0}. Choose from: {1}" -f ($bad -join ', '), ($allowed -join ', ')) -ForegroundColor Yellow
                    continue
                }
                $vals = @($vals | ForEach-Object { $v = $_; $allowed | Where-Object { $_ -eq $v } | Select-Object -First 1 })
            }
            return $vals
        }
        return $val
    }
}

# ---------------------------------------------------------------------------
# Let the user pick a subset of columns from the returned objects. Enter = all.
# ---------------------------------------------------------------------------
function Select-ResultColumns {
    param([Parameter(Mandatory)] $Items)
    $cols = @($Items[0].PSObject.Properties.Name)
    Write-Host ''
    Write-Host 'Available columns:' -ForegroundColor White
    for ($i=0; $i -lt $cols.Count; $i++) { Write-Host ('  {0,2}. {1}' -f ($i+1), $cols[$i]) }
    $pick = Read-Host 'Enter column numbers (comma-separated), or Enter for all'
    $nums = @($pick -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })
    if ($nums.Count -eq 0) { return $cols }
    $sel = foreach ($n in $nums) { $idx=[int]$n-1; if ($idx -ge 0 -and $idx -lt $cols.Count) { $cols[$idx] } }
    @($sel)
}

# ---------------------------------------------------------------------------
# Save objects to a CSV file, defaulting to .\exports\<cmdlet>-<timestamp>.csv.
# ---------------------------------------------------------------------------
function Save-ResultCsv {
    param([Parameter(Mandatory)] $Items, [string] $Cmdlet)
    $default = Join-Path (Join-Path $script:ToolkitHome 'exports') ("{0}-{1}.csv" -f $Cmdlet, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $path = Read-Host ("CSV path (Enter for {0})" -f $default)
    if ([string]::IsNullOrWhiteSpace($path)) { $path = $default } else { $path = $path.Trim('"',' ') }
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $Items | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    Write-Host ("Saved {0} row(s) to {1}" -f @($Items).Count, $path) -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Post-run handling of whatever a cmdlet returned: show as a table (all or
# chosen columns) or export to CSV. Scalar/string output just prints; nothing
# returned says so. Loops so you can view AND save without re-running.
# ---------------------------------------------------------------------------
function Show-Result {
    param($Result, [string] $Cmdlet)

    $items = @($Result)
    if ($items.Count -eq 0) { Write-Host 'No objects returned.' -ForegroundColor White; return }
    Write-Host ("Returned {0} item(s)." -f $items.Count) -ForegroundColor White

    # Only offer table/CSV handling for real objects; scalars/strings just print.
    $first = $items[0]
    $isObject = $first -is [psobject] -and -not ($first -is [string]) -and -not ($first.GetType().IsPrimitive) `
                -and @($first.PSObject.Properties).Count -gt 0
    if (-not $isObject) {
        $items | ForEach-Object { Write-Host ("  {0}" -f $_) }
        return
    }

    # Remember this result so the NEXT action can pipe a value out of it (the
    # 'last' keyword in Read-ParameterValue). Scalars/strings aren't useful as a
    # pipe source, so only real objects reach here and get stored.
    $script:LastResult = $items
    $script:LastResultCmdlet = $Cmdlet
    Write-Host "Tip: in the next action, type 'last' at a field to reuse a value from these results." -ForegroundColor White

    while ($true) {
        $choice = Get-MenuChoice -Title 'What do you want to do with the results?' -Options @(
            'Show as a table',
            'Show a table with only columns I choose',
            'Save all columns to a CSV file',
            'Save only columns I choose to a CSV file',
            'Done'
        )
        switch ($choice) {
            'Show as a table'                        { $items | Format-Table -AutoSize | Out-Host }
            'Show a table with only columns I choose' { $c = Select-ResultColumns -Items $items; if ($c) { $items | Select-Object $c | Format-Table -AutoSize | Out-Host } }
            'Save all columns to a CSV file'          { Save-ResultCsv -Items $items -Cmdlet $Cmdlet }
            'Save only columns I choose to a CSV file' { $c = Select-ResultColumns -Items $items; if ($c) { Save-ResultCsv -Items ($items | Select-Object $c) -Cmdlet $Cmdlet } }
            default { return }
        }
    }
}

# ---------------------------------------------------------------------------
# Tiny persisted config (UPN + last service) at <ToolkitHome>\config.json so
# you don't retype your email every launch.
# ---------------------------------------------------------------------------
function Get-ToolkitConfig {
    $path = Join-Path $script:ToolkitHome 'config.json'
    if (Test-Path $path) { try { return (Get-Content $path -Raw | ConvertFrom-Json) } catch { } }
    [pscustomobject]@{ Upn = $null; LastService = $null }
}
function Save-ToolkitConfig {
    param([string] $Upn, [string] $LastService)
    if (-not (Test-Path $script:ToolkitHome)) { New-Item -ItemType Directory -Path $script:ToolkitHome -Force | Out-Null }
    [pscustomobject]@{ Upn = $Upn; LastService = $LastService } |
        ConvertTo-Json | Set-Content -Path (Join-Path $script:ToolkitHome 'config.json') -Encoding UTF8
}

# ---------------------------------------------------------------------------
# "Who am I" preflight: after connecting, show the identity in use (and the
# granted Graph scopes / EXO org) so you know your blast radius before acting.
# ---------------------------------------------------------------------------
function Show-WhoAmI {
    param([Parameter(Mandatory)] $Service, [string] $Upn)
    Write-Host ("Connected as: {0}" -f $Upn) -ForegroundColor Green
    try {
        if ($Service.Module -eq 'ExchangeOnlineManagement') {
            $ci = Get-ConnectionInformation -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($ci) { Write-Host ("  Org: {0}  |  UPN: {1}" -f $ci.Organization, $ci.UserPrincipalName) -ForegroundColor White }
        } elseif ($Service.Module -like 'Microsoft.Graph*') {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) {
                Write-Host ("  Account: {0}" -f $ctx.Account) -ForegroundColor White
                Write-Host ("  Granted scopes: {0}" -f ($ctx.Scopes -join ', ')) -ForegroundColor White
            }
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# After an action finishes, ask what to do next. Returns 'again' (back to the
# category menu for this service), 'switch' (pick another service), or 'quit'.
# ---------------------------------------------------------------------------
function Get-NextStep {
    switch (Get-MenuChoice -Title 'What next?' -Options @(
                'Do another task (same service)','Switch service','Quit')) {
        'Switch service' { 'switch' }
        'Quit'           { 'quit' }
        default          { 'again' }
    }
}

# ---------------------------------------------------------------------------
# Given a catalogued action, ask run-once vs bulk and dispatch. Shared by the
# category browser and the keyword search. Returns the Get-NextStep directive
# ('again' if the user backed out before running anything).
# ---------------------------------------------------------------------------
function Invoke-Action {
    param([Parameter(Mandatory)] $Action)
    $mode = Get-MenuChoice -Title "How do you want to run '$($Action.Label)'?" `
                -Options @('Run once (type the values now)','Bulk - run for every row in a CSV file') -AllowBack
    if (-not $mode) { return 'again' }
    if ($mode -like 'Bulk*') {
        Invoke-BulkCmdlet  -Cmdlet $Action.Cmdlet -Prompt $Action.Prompt -ParamHelp $Action.ParamHelp -Preset $Action.Preset -Transform $Action.Transform
    } else {
        Invoke-GuidedCmdlet -Cmdlet $Action.Cmdlet -Prompt $Action.Prompt -Help $Action.Help -ParamHelp $Action.ParamHelp -Preset $Action.Preset -Hints $Action.Hints -Transform $Action.Transform -Project $Action.Project
    }
    Get-NextStep
}

# ===========================================================================
# Main loop
# ===========================================================================
function Start-M365Toolkit {
    # Session state for output piping (the 'last' keyword in Read-ParameterValue):
    # the most recent result set and the cmdlet that produced it. Populated by
    # Show-Result; in-memory only (not persisted across sessions).
    $script:LastResult = @()
    $script:LastResultCmdlet = $null

    $logDir = Join-Path $script:ToolkitHome 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $logFile = Join-Path $logDir "session-$stamp.log"
    Start-Transcript -Path $logFile -Append | Out-Null
    try {
        Write-Host '=====================================================' -ForegroundColor White
        Write-Host '  Microsoft 365 Guided Command Builder  (v0.10)'       -ForegroundColor White
        Write-Host '=====================================================' -ForegroundColor White

        $cfg = Get-ToolkitConfig
        $upnPrompt = if ($cfg.Upn) { "Your email (UPN) [Enter = $($cfg.Upn)]" } else { 'Your email (UPN)' }
        $upn = Read-Host $upnPrompt
        if ([string]::IsNullOrWhiteSpace($upn) -and $cfg.Upn) { $upn = $cfg.Upn }

        $connected = @{}   # ServiceName -> $true
        $quit = $false
        while ($true) {
            # Choose / switch service - offer the last-used one first.
            $svcNames = @($Services.Name | Sort-Object)   # A-Z; last-used floated to the top below
            if ($cfg.LastService -and ($svcNames -contains $cfg.LastService)) {
                $svcNames = @($cfg.LastService) + @($svcNames | Where-Object { $_ -ne $cfg.LastService })
            }
            $svcName = Get-MenuChoice -Title 'Connect to which service?' -Options $svcNames -ExtraOptions @('Quit')
            if ($svcName -eq 'Quit') { break }

            $service = $Services | Where-Object { $_.Name -eq $svcName }
            if (-not $connected[$svcName]) {
                Connect-M365Service -Service $service -Upn $upn
                $connected[$svcName] = $true
                Show-WhoAmI -Service $service -Upn $upn
            }
            Save-ToolkitConfig -Upn $upn -LastService $svcName

            # Inner loop: pick and run actions for this service.
            while ($true) {
                $svcActions = $Actions | Where-Object { $_.Service -eq $svcName }
                if (-not $svcActions) {
                    Write-Host 'No plain-language actions catalogued for this service yet.' -ForegroundColor Yellow
                    Write-Host '(You can still extend $Actions - or point the engine at any cmdlet.)' -ForegroundColor White
                }

                $categories = @($svcActions.Category | Select-Object -Unique | Sort-Object)   # A-Z; fixed items appended after
                $cat = Get-MenuChoice -Title "[$svcName] Choose a category" `
                            -Options @($categories) `
                            -ExtraOptions @('Search actions by keyword','Run any command by name (advanced)','Switch service') -AllowBack
                if (-not $cat -or $cat -eq 'Switch service') { break }

                # Keyword search across this service's actions (name / cmdlet / category).
                if ($cat -eq 'Search actions by keyword') {
                    $kw = (Read-Host 'Keyword (matches action name, cmdlet, or category)').Trim()
                    if ($kw) {
                        $rx = [regex]::Escape($kw)
                        $hits = @($svcActions | Where-Object { $_.Label -match $rx -or $_.Cmdlet -match $rx -or $_.Category -match $rx } | Sort-Object Label)   # results A-Z
                        if ($hits.Count -eq 0) { Write-Host "No actions match '$kw'." -ForegroundColor Yellow }
                        else {
                            $lbl = Get-MenuChoice -Title "Actions matching '$kw'" `
                                        -Options $hits.Label `
                                        -Descriptions @($hits | ForEach-Object { [string]$_.Help }) -AllowBack
                            if ($lbl) {
                                $post = Invoke-Action -Action ($hits | Where-Object { $_.Label -eq $lbl } | Select-Object -First 1)
                                if ($post -eq 'quit')   { $quit = $true; break }
                                if ($post -eq 'switch') { break }
                            }
                        }
                    }
                    continue
                }

                # Escape hatch: drive ANY cmdlet in this connected module by name.
                if ($cat -eq 'Run any command by name (advanced)') {
                    $cmdName = (Read-Host 'Cmdlet name (e.g. Set-Mailbox), or Enter to cancel').Trim()
                    if ($cmdName) {
                        $mode = Get-MenuChoice -Title "How do you want to run '$cmdName'?" `
                                    -Options @('Run once (type the values now)','Bulk - run for every row in a CSV file') -AllowBack
                        try {
                            if     ($mode -like 'Bulk*') { Invoke-BulkCmdlet  -Cmdlet $cmdName }
                            elseif ($mode)               { Invoke-GuidedCmdlet -Cmdlet $cmdName }
                        } catch { Write-Host ("Could not run '{0}': {1}" -f $cmdName, $_.Exception.Message) -ForegroundColor Red }
                        $post = Get-NextStep
                        if ($post -eq 'quit')   { $quit = $true; break }
                        if ($post -eq 'switch') { break }
                    }
                    continue
                }

                $catActions = @($svcActions | Where-Object { $_.Category -eq $cat } | Sort-Object Label)   # actions A-Z
                $label = Get-MenuChoice -Title "[$svcName / $cat] Choose an action" `
                            -Options $catActions.Label `
                            -Descriptions @($catActions | ForEach-Object { [string]$_.Help }) `
                            -AllowBack
                if (-not $label) { continue }

                $action = $catActions | Where-Object { $_.Label -eq $label } | Select-Object -First 1
                $post = Invoke-Action -Action $action
                if ($post -eq 'quit')   { $quit = $true; break }
                if ($post -eq 'switch') { break }
            }
            if ($quit) { break }
        }
        Write-Host 'Goodbye.' -ForegroundColor White
    }
    finally {
        Stop-Transcript | Out-Null
    }
}

Start-M365Toolkit
