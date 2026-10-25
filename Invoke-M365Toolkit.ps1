<#
.SYNOPSIS
    Guided, plain-language command builder for Microsoft 365 PowerShell.

.DESCRIPTION
    You never memorize cmdlet names or syntax. Launch the tool, enter your email,
    choose which M365 service to connect to, then browse by category and pick an
    action in plain language. The tool introspects the real cmdlet, walks you
    through its parameters (Required -> Optional -> Switches) in human terms,
    shows a dry-run preview of the assembled command, and only runs it when you
    confirm. Every action is recorded in a local action log (the Microsoft 365
    audit log is the full record of changes); a full transcript only on request.

    Architecture (three independent layers, each extended by adding one row):
      1. $Services  - connection registry: one row per M365 service.
      2. Invoke-GuidedCmdlet - generic introspection engine, works for ANY cmdlet.
      3. $Actions   - plain-language catalogue: friendly label -> cmdlet.

.NOTES
    v0.30 - 2026-10-25. Engine runs on PowerShell 7+ (Windows/macOS/Linux) and
    Windows PowerShell 5.1. Per service: Exchange, Purview, Entra, Intune and Teams
    work everywhere; SharePoint (PnP) needs pwsh 7.2+; SharePoint Admin (SPO) and
    Power Platform use Windows-only modules and run on Windows only.
    Requires the relevant module installed for whichever service you connect to.
#>

[CmdletBinding()]
param(
    # Skip the run confirmation and only ever show the preview (never execute).
    [switch]$DryRunOnly,
    # Sign in to Graph (Entra, Intune) with a device code instead of the browser.
    # Use it only if the browser sign-in hangs (seen on macOS): many tenants block
    # device code flow because it is a common phishing route.
    [switch]$UseDeviceCode,
    # Also write a full screen transcript to logs\. Off by default: it holds every
    # directory search and result you view. The action log is always written.
    [switch]$Transcript
)

# Run with 'Stop', but remember the caller's setting: under 'irm | iex' this scope
# is the caller's own shell, which gets its setting back when the tool exits.
$script:CallerErrorActionPreference = $ErrorActionPreference
$ErrorActionPreference = 'Stop'
$script:ToolkitVersion = 'v0.30'

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
        MinModuleVersion = '3.7.0'   # LoadCmdletHelp exists from 3.7
    }
    [pscustomobject]@{
        Name        = 'Entra / Identity (Graph)'
        Module      = 'Microsoft.Graph'
        ConnectCmd  = 'Connect-MgGraph'
        # Browser sign-in by default; -UseDeviceCode switches to a device code (only if
        # the browser hangs). ContextScope Process keeps the token in this process only,
        # not in the user's token cache, so it ends with the session.
        Scopes         = @('User.Read.All','GroupMember.Read.All')
        CommandModules = @('Microsoft.Graph.Users','Microsoft.Graph.Groups','Microsoft.Graph.Identity.DirectoryManagement')
        InstallModules = @('Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Groups','Microsoft.Graph.Identity.DirectoryManagement')
        ConnectArgs = { param($Upn, $Svc) $a = @{ Scopes = $Svc.Scopes; ContextScope = 'Process' }; if ($UseDeviceCode) { $a.UseDeviceCode = $true }; $a }
        Disconnect  = 'Disconnect-MgGraph'
    }
    [pscustomobject]@{
        Name        = 'SharePoint Online (PnP)'
        Module      = 'PnP.PowerShell'
        ConnectCmd  = 'Connect-PnPOnline'
        # Current PnP.PowerShell has no shared app any more: interactive sign-in needs
        # the client ID of an app registered in your tenant.
        ConnectArgs = { param($Upn)
            $a = @{ Url = (Read-Host 'SharePoint site/admin URL'); Interactive = $true }
            $cid = "$(Read-Host "Client ID of your tenant's PnP app registration (Enter to try without)")".Trim()
            if ($cid) { $a.ClientId = $cid }
            $a }
        Disconnect  = 'Disconnect-PnPOnline'
        MinPSVersion = '7.4'   # PnP.PowerShell 3.x needs PowerShell 7.4 or later
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
        ConnectArgs = { param($Upn) @{ Url = (Read-Host 'SharePoint ADMIN URL (e.g. https://contoso-admin.sharepoint.com)') } }
        Disconnect  = 'Disconnect-SPOService'
        WinOnly     = $true   # Microsoft.Online.SharePoint.PowerShell is a .NET Framework (Windows-only) module
    }
    [pscustomobject]@{
        Name        = 'Intune / Devices (Graph)'
        Module      = 'Microsoft.Graph'
        ConnectCmd  = 'Connect-MgGraph'
        # Same Graph SDK as Entra, but consented for device-management scopes.
        # Read-only (least privilege): the only catalogued Intune action lists
        # enrolled devices (Get-MgDeviceManagementManagedDevice). If a write
        # action is ever added, widen the matching scope to .ReadWrite.All then.
        Scopes         = @('DeviceManagementManagedDevices.Read.All')
        CommandModules = @('Microsoft.Graph.DeviceManagement')
        InstallModules = @('Microsoft.Graph.Authentication','Microsoft.Graph.DeviceManagement')
        ConnectArgs = { param($Upn, $Svc) $a = @{ Scopes = $Svc.Scopes; ContextScope = 'Process' }; if ($UseDeviceCode) { $a.UseDeviceCode = $true }; $a }
        Disconnect  = 'Disconnect-MgGraph'
    }
    [pscustomobject]@{
        Name        = 'Power Platform (admin)'
        Module      = 'Microsoft.PowerApps.Administration.PowerShell'
        ConnectCmd  = 'Add-PowerAppsAccount'   # opens an interactive sign-in
        ConnectArgs = { param($Upn) @{ } }
        Disconnect  = 'Remove-PowerAppsAccount'
        WinOnly     = $true   # Microsoft.PowerApps.Administration.PowerShell is a .NET Framework (Windows-only) module
    }
)

# ---------------------------------------------------------------------------
# Helper "cmdlets" that back catalogue actions needing more than one call. The
# engine drives these via Get-Command exactly like a real cmdlet, so they just
# expose a normal param() block and return objects (reshaped by the action's
# Project). Add one here when no single built-in cmdlet can answer a question.
# ---------------------------------------------------------------------------
function Get-UserSharedMailboxAccess {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string] $User)
    # No reverse index exists ("which mailboxes can X open?"), so scan the shared
    # mailboxes and keep those where the user has a direct (non-inherited)
    # FullAccess grant. Can take a minute on tenants with many shared mailboxes.
    # A mailbox that can't be read (throttling, a transient error) is counted and
    # reported, never silently dropped - a quiet gap would read as "no access".
    $shared = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox -ErrorAction Stop)
    $failed = 0
    foreach ($m in $shared) {
        try {
            $hit = @(Get-MailboxPermission -Identity $m.Identity -User $User -ErrorAction Stop |
                     Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited })
            if ($hit.Count -gt 0) { $m }
        } catch { $failed++ }
    }
    if ($failed -gt 0) { Write-Warning ('{0} of {1} shared mailboxes could not be read - this list may be incomplete.' -f $failed, $shared.Count) }
}

# The calendar folder of a mailbox as "<mailbox>:\<folder>". The folder name is
# localized (e.g. "Kalender"), so read it from the mailbox; fall back to
# "Calendar" when it can't be read.
function Get-CalendarFolderId {
    param([Parameter(Mandatory)] [string] $Mailbox)
    $mbx = $Mailbox.Trim()
    try {
        $f = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar -ErrorAction Stop |
             Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
        if ($f -and $f.FolderPath) { return ('{0}:{1}' -f $mbx, ($f.FolderPath -replace '/', '\')) }
    } catch { }
    return ('{0}:\Calendar' -f $mbx)
}

# Strip control characters (ESC sequences can rewrite the terminal) from text
# that comes from the tenant before it is shown.
function ConvertTo-SafeText {
    param([AllowNull()] $Text)
    if ($null -eq $Text) { return '' }
    return ([string]$Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '')
}

# Turn a Get-MailboxFolderPermission entry's .User object back into a value that
# Set/Remove-MailboxFolderPermission -User will actually bind to. That object is
# a MailboxFolderUserId whose usable address hides in different places depending
# on whether the grantee is an internal user, a mail-enabled recipient, or an
# unresolved name - so try the reliable SMTP paths first and fall back to the
# display name. Best-effort; returns $null only if nothing at all is present.
function Get-FolderPermissionUserAddress {
    param([Parameter(Mandatory)] $User)
    foreach ($sb in @(
        { $User.ADRecipient.PrimarySmtpAddress },
        { $User.RecipientPrincipal.PrimarySmtpAddress },
        { $User.ADRecipient.WindowsLiveID },
        { $User.DisplayName }
    )) {
        $candidate = $null
        try { $candidate = & $sb } catch { }
        if ($candidate) { return [string]$candidate }
    }
    return $null
}

# A dependent candidate source (see the 'Candidates' action hook): given the
# values collected so far, returns the people who ALREADY have a permission entry
# on the chosen mailbox's calendar - i.e. the only valid targets for CHANGING or
# REMOVING calendar access (Set/Remove-MailboxFolderPermission fail with "no
# existing permission entry" for anyone else). Reads $bound['Identity'] as the
# raw mailbox (the ':\Calendar' Transform hasn't run yet), filters out the
# built-in Default/Anonymous rows, and drops anyone we can't turn into a bindable
# address. Returns objects shaped like the $ParamLookup hits: .Display / .Value.
$calendarUserCandidates = { param($bound)
    $mbx = "$($bound['Identity'])".Trim()
    if ([string]::IsNullOrWhiteSpace($mbx)) { return @() }
    Get-MailboxFolderPermission -Identity (Get-CalendarFolderId -Mailbox $mbx) -ErrorAction Stop |
        Where-Object { $_.User.DisplayName -notin @('Default','Anonymous') } |
        ForEach-Object {
            $addr = Get-FolderPermissionUserAddress -User $_.User
            [pscustomobject]@{
                Display = ('{0}  <{1}>  [{2}]' -f (ConvertTo-SafeText $_.User.DisplayName), $(if ($addr) { $addr } else { '?' }), (@($_.AccessRights) -join ', '))
                Value   = $addr
            }
        } | Where-Object { $_.Value }
}

# ---------------------------------------------------------------------------
# Layer 3: Plain-language action catalogue. Add an action = add one row here.
# Cmdlet is the ONLY thing you must name; parameters are discovered by layer 2.
# 'Prompt' (optional) curates which params to walk first; if omitted, the engine
# walks all Required params and then offers the full Optional/Switch lists.
# ---------------------------------------------------------------------------
$Actions = @(
    [pscustomobject]@{
        Id='mbx-give-access'; Service='Exchange Online'; Category='Give access'
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
        Id='mbx-remove-access'; Service='Exchange Online'; Category='Remove & change access'
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
        Id='mbx-give-sendas'; Service='Exchange Online'; Category='Give access'
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
        Id='mbx-check-sendas'; Service='Exchange Online'; Category='Check access'
        Label='See who can Send As a mailbox'
        Cmdlet='Get-RecipientPermission'; Prompt=@('Identity')
        Help='List everyone who has Send As on a mailbox - use it to verify a Send As grant.'
        ParamHelp=@{ Identity='The mailbox to check - the one others may send email AS' }
    }
    [pscustomobject]@{
        Id='mbx-check-access'; Service='Exchange Online'; Category='Check access'
        Label='See who has access to a mailbox (FullAccess etc.)'
        Cmdlet='Get-MailboxPermission'; Prompt=@('Identity')
        Help='List who can open/manage a mailbox. NOTE: "NT AUTHORITY\SELF" and inherited rows are normal system entries - the real grants are the named users.'
        ParamHelp=@{ Identity='The mailbox to check - whose access list you want to see' }
    }
    [pscustomobject]@{
        Id='mail-groups-of-user'; Service='Exchange Online'; Category='Check access'
        Label='See which mail groups a user is in'
        Cmdlet='Get-Recipient'; Prompt=@('Filter')
        Preset=@{ ResultSize='Unlimited' }
        # The user types a person (email/name); the Transform resolves them to a
        # DistinguishedName and rewrites -Filter into a membership query, so this
        # lists the distribution / mail-enabled groups that user belongs to.
        # Uses only Exchange access - no Graph, no admin consent needed.
        Transform=@{ Filter = { param($v)
            $rcpt = Get-Recipient -Identity $v -ErrorAction SilentlyContinue | Select-Object -First 1
            $dn = if ($rcpt) { $rcpt.DistinguishedName } else { $v }
            # Escape single quotes ('->'') so names like O'Brien don't break the
            # OPATH filter or allow filter injection.
            $dnEsc = "$dn".Replace("'","''")
            "Members -eq '$dnEsc'"
        } }
        Help='List the MAIL groups (distribution lists and mail-enabled security groups) a user is a direct member of. Type the person and the tool finds their groups. Note: Microsoft 365 groups and nested membership are not covered by this Exchange filter.'
        ParamHelp=@{ Filter='The user whose mail groups you want - type their email or name (or find / last). The tool resolves them and looks up their memberships.' }
        Project={ param($r)
            @($r) | ForEach-Object {
                $rt = [string]$_.RecipientTypeDetails
                $type = switch ($rt) {
                    'MailUniversalDistributionGroup' { 'Distribution group' }
                    'MailUniversalSecurityGroup'     { 'Mail-enabled security' }
                    'DynamicDistributionGroup'       { 'Dynamic distribution group' }
                    'GroupMailbox'                   { 'Microsoft 365 group' }
                    default                          { if ($rt) { $rt } else { 'Group' } }
                }
                [pscustomobject]@{ Name = [string]$_.DisplayName; Email = [string]$_.PrimarySmtpAddress; Type = $type }
            }
        }
    }
    [pscustomobject]@{
        Id='dg-list-members'; Service='Exchange Online'; Category='Check access'
        Label='See who is in a distribution group'
        Cmdlet='Get-DistributionGroupMember'; Prompt=@('Identity')
        Preset=@{ ResultSize='Unlimited' }
        Help='List the members of a distribution group / mailing list (who is IN the group). Type the group; find / last work too.'
        ParamHelp=@{ Identity='The distribution group whose members you want (email or name)' }
        Project={ param($r)
            @($r) | ForEach-Object {
                $rt = [string]$_.RecipientTypeDetails; if (-not $rt) { $rt = [string]$_.RecipientType }
                $type = switch ($rt) {
                    'UserMailbox'                    { 'User' }
                    'SharedMailbox'                  { 'Shared mailbox' }
                    'MailUniversalDistributionGroup' { 'Distribution group (nested)' }
                    'MailUniversalSecurityGroup'     { 'Mail-enabled security' }
                    'MailContact'                    { 'Contact' }
                    'MailUser'                       { 'Mail user' }
                    default                          { if ($rt) { $rt } else { 'Member' } }
                }
                [pscustomobject]@{ Name = [string]$_.DisplayName; Email = [string]$_.PrimarySmtpAddress; Type = $type }
            }
        }
    }
    [pscustomobject]@{
        Id='mbx-check-shared-access'; Service='Exchange Online'; Category='Check access'
        Label='See which shared mailboxes a user can access'
        Cmdlet='Get-UserSharedMailboxAccess'; Prompt=@('User')
        Help='List the shared mailboxes a user has FullAccess to. Scans all shared mailboxes, so it can take a minute on large tenants. Type the person; find / last work.'
        ParamHelp=@{ User='The person whose shared-mailbox access you want (email or name)' }
        Project={ param($r)
            @($r) | ForEach-Object {
                [pscustomobject]@{ Name = [string]$_.DisplayName; Email = [string]$_.PrimarySmtpAddress; Type = 'Shared mailbox' }
            }
        }
    }
    [pscustomobject]@{
        Id='cal-check-access'; Service='Exchange Online'; Category='Check access'
        Label='See who has access to a calendar'
        Cmdlet='Get-MailboxFolderPermission'; Prompt=@('Identity')
        Transform=@{ Identity = { param($v) Get-CalendarFolderId -Mailbox $v } }
        Help='List who can see or edit a mailbox''s calendar, and the role each has.'
        ParamHelp=@{ Identity='The mailbox whose calendar you want to check (email or name) - the tool targets its \Calendar folder automatically' }
    }
    [pscustomobject]@{
        Id='cal-give-access'; Service='Exchange Online'; Category='Give access'
        Label='Give a user access to a calendar'
        Cmdlet='Add-MailboxFolderPermission'; Prompt=@('Identity','User','AccessRights')
        Transform=@{ Identity = { param($v) Get-CalendarFolderId -Mailbox $v } }
        Hints=@{ AccessRights=@('Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor','AvailabilityOnly','LimitedDetails','None') }
        Help='Grant a person a role on a mailbox''s calendar. Use this only if they have NO existing entry - otherwise use "Change a user''s calendar access".'
        ParamHelp=@{
            Identity     = 'The mailbox whose calendar is being shared (its \Calendar folder is targeted automatically)'
            User         = 'The person who will GET access to the calendar'
            AccessRights = 'Their role. Reviewer = read events; Editor = create/edit/delete; AvailabilityOnly = free/busy only; LimitedDetails = free/busy + subject & location'
        }
    }
    [pscustomobject]@{
        Id='cal-change-access'; Service='Exchange Online'; Category='Remove & change access'
        Label='Change a user''s calendar access'
        Cmdlet='Set-MailboxFolderPermission'; Prompt=@('Identity','User','AccessRights')
        Transform=@{ Identity = { param($v) Get-CalendarFolderId -Mailbox $v } }
        Candidates=@{ User = $calendarUserCandidates }
        Hints=@{ AccessRights=@('Owner','PublishingEditor','Editor','PublishingAuthor','Author','NonEditingAuthor','Reviewer','Contributor','AvailabilityOnly','LimitedDetails','None') }
        Help='Change the role of someone who ALREADY has access to a calendar (fails if they have no existing entry - use "Give a user access" first). After you pick the calendar, the tool lists only the people who already have an entry, so you can''t hit the "no existing entry" error.'
        ParamHelp=@{
            Identity     = 'The mailbox whose calendar is shared (its \Calendar folder is targeted automatically)'
            User         = 'The person whose access you are changing'
            AccessRights = 'The new role (see the grant action for what each role means)'
        }
    }
    [pscustomobject]@{
        Id='cal-remove-access'; Service='Exchange Online'; Category='Remove & change access'
        Label='Remove a user''s calendar access'
        Cmdlet='Remove-MailboxFolderPermission'; Prompt=@('Identity','User')
        Transform=@{ Identity = { param($v) Get-CalendarFolderId -Mailbox $v } }
        Candidates=@{ User = $calendarUserCandidates }
        Help='Take away a person''s access to a mailbox''s calendar. After you pick the calendar, the tool lists only the people who currently have an entry, so you pick from real grants instead of guessing.'
        ParamHelp=@{
            Identity = 'The mailbox whose calendar was shared (its \Calendar folder is targeted automatically)'
            User     = 'The person who will LOSE calendar access'
        }
    }
    [pscustomobject]@{
        Id='mbx-lookup'; Service='Exchange Online'; Category='Mailboxes'; Label='Look up a mailbox'
        Cmdlet='Get-EXOMailbox'; Prompt=@('Identity')
        Help='Show details about a mailbox.'
        ParamHelp=@{ Identity='The mailbox to look up (email address or name)' }
    }
    [pscustomobject]@{
        Id='mbx-set-setting'; Service='Exchange Online'; Category='Mailboxes'; Label='Change a mailbox setting'
        Cmdlet='Set-Mailbox'; Prompt=$null
        Help='Change a setting on a mailbox (advanced - walks all available options).'
    }
    [pscustomobject]@{
        Id='mbx-create-shared'; Service='Exchange Online'; Category='Mailboxes'
        Label='Create a new shared mailbox'
        Cmdlet='New-Mailbox'; Prompt=@('Name','DisplayName','PrimarySmtpAddress')
        Preset=@{ Shared=$true }
        Help='Create a brand-new shared mailbox (no license needed, no password).'
        ParamHelp=@{
            Name               = 'Internal name / alias, no spaces (e.g. sales)'
            DisplayName        = 'Friendly name shown to people (e.g. Sales Team)'
            PrimarySmtpAddress = 'The email address for it (e.g. sales@contoso.com)'
        }
    }
    [pscustomobject]@{
        Id='mbx-convert-to-shared'; Service='Exchange Online'; Category='Mailboxes'
        Label='Convert an existing mailbox to shared'
        Cmdlet='Set-Mailbox'; Prompt=@('Identity')
        Preset=@{ Type='Shared' }
        Help='Turn a normal user mailbox into a shared mailbox (e.g. after someone leaves). Free up the license afterwards.'
        ParamHelp=@{ Identity='The mailbox to convert (email address or name)' }
    }
    [pscustomobject]@{
        Id='mbx-convert-to-regular'; Service='Exchange Online'; Category='Mailboxes'
        Label='Convert a shared mailbox back to a regular (user) mailbox'
        Cmdlet='Set-Mailbox'; Prompt=@('Identity')
        Preset=@{ Type='Regular' }
        Help='Turn a shared mailbox back into a normal user mailbox. NOTE: a regular mailbox needs an Exchange Online license assigned, or it gets disabled after ~30 days.'
        ParamHelp=@{ Identity='The shared mailbox to convert back (email address or name)' }
    }
    [pscustomobject]@{
        Id='dg-add-member'; Service='Exchange Online'; Category='Give access'; Label='Add a member to a distribution group'
        Cmdlet='Add-DistributionGroupMember'; Prompt=@('Identity','Member')
        Help='Add someone to a distribution / mailing group.'
        ParamHelp=@{ Identity='The group to add TO'; Member='The person being ADDED' }
    }
    [pscustomobject]@{
        Id='user-lookup'; Service='Entra / Identity (Graph)'; Category='Users'; Label='Look up a user'
        Cmdlet='Get-MgUser'; Prompt=@('UserId')
        Help='Show details about a user account.'
        ParamHelp=@{ UserId='The user to look up (email or object id)' }
    }
    [pscustomobject]@{
        Id='user-groups-membership'; Service='Entra / Identity (Graph)'; Category='Check access'
        Label='See which groups a user is in'
        Cmdlet='Get-MgUserMemberOf'; Prompt=@('UserId')
        Preset=@{ All=$true }
        Help='List the groups (and directory roles) a user is a DIRECT member of, each labelled with what it is - Distribution group, Microsoft 365 group, Security group, Mail-enabled security, or Directory role. Groups reached through nesting are not included.'
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
        Id='spo-list-sites'; Service='SharePoint Admin (SPO)'; Category='Sites'; Label='List all SharePoint sites'
        Cmdlet='Get-SPOSite'; Prompt=$null
        Help='List the site collections in the tenant (add -Limit All to see every one).'
    }
    [pscustomobject]@{
        Id='intune-list-devices'; Service='Intune / Devices (Graph)'; Category='Devices'; Label='List enrolled devices'
        Cmdlet='Get-MgDeviceManagementManagedDevice'; Prompt=$null
        Help='List devices enrolled in Intune.'
    }
    [pscustomobject]@{
        Id='pp-list-environments'; Service='Power Platform (admin)'; Category='Environments'; Label='List Power Platform environments'
        Cmdlet='Get-AdminPowerAppEnvironment'; Prompt=$null
        Help='List the Power Platform environments in the tenant.'
    }
)

# Identity guard (parse time): actions are dispatched by their stable Id, so a
# duplicated or missing Id would silently route to the wrong action. Fail loud
# here - before the tool ever runs - so a contributor sees it immediately.
$dupeIds = $Actions | Group-Object Id | Where-Object Count -gt 1
if ($dupeIds) { throw ("Action catalogue has duplicate Id(s): {0}. Ids must be unique." -f (($dupeIds.Name) -join '; ')) }
$missingId = $Actions | Where-Object { [string]::IsNullOrWhiteSpace($_.Id) }
if ($missingId) { throw ("Action(s) missing an Id: {0}" -f (($missingId.Label) -join '; ')) }

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

# Parameter names that look like credentials. One pattern for every place that
# decides what to read hidden and what to mask (prompts, previews, undo, bulk).
$script:SecretNamePattern = '(?i)(password|passwd|pwd|secret|credential|token|apikey|api[_-]?key|clientsecret|access[_-]?key|privatekey|accountkey|encryptionkey|sharedkey|signingkey|storagekey|subscriptionkey|passphrase|connectionstring|certificatebase64)'
# Bearer tokens (JWTs) in command output are masked before display or export.
$script:JwtPattern = 'eyJ[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]{8,}\.[A-Za-z0-9_-]*'

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
        $t = "$term".Replace("'","''")   # escape quotes so O'Brien-type names can't break/inject the OPATH filter
        $rows = $null
        try {
            $rows = Get-EXORecipient -ResultSize 25 -Filter "DisplayName -like '*$t*' -or PrimarySmtpAddress -like '*$t*'" -ErrorAction Stop
        } catch {
            # The escape stops injection, but EXO's OPATH -like still rejects an
            # apostrophe sitting between wildcards ("Invalid wildcard in between").
            # Only that case gets the fallback - auth/throttling/network errors are
            # re-thrown, not masked behind a second broad query.
            if ("$($_.Exception.Message)" -notmatch 'wildcard') { throw }
            # Query the server on the MOST SELECTIVE alphanumeric chunk of the term
            # (e.g. "Brien" from "O'Brien", not "O" - so we don't pull a huge, capped
            # result set that might not even contain the target) then refine
            # client-side, where the apostrophe is just an ordinary character.
            $token = ($term -split "[^A-Za-z0-9]+" | Where-Object { $_ } | Sort-Object { $_.Length } -Descending | Select-Object -First 1)
            if (-not $token) { return }   # nothing searchable (term was only punctuation)
            $tkEsc = "$token".Replace("'","''")
            $rows = @(Get-EXORecipient -ResultSize 200 -Filter "DisplayName -like '*$tkEsc*' -or PrimarySmtpAddress -like '*$tkEsc*'" -ErrorAction Stop |
                        Where-Object { $_.DisplayName -like "*$term*" -or "$($_.PrimarySmtpAddress)" -like "*$term*" } |
                        Select-Object -First 25)
        }
        $rows | ForEach-Object { [pscustomobject]@{ Display = ('{0}  <{1}>  [{2}]' -f (ConvertTo-SafeText $_.DisplayName), $_.PrimarySmtpAddress, $_.RecipientTypeDetails); Value = [string]$_.PrimarySmtpAddress } }
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
    Filter   = $recipientLookup   # only "See which mail groups a user is in" prompts -Filter: type a person; its Transform rewrites it into a membership query
    UserId   = @{
        Probe = 'Get-MgUser'
        Run   = {
            param($term)
            $t = "$term".Replace("'","''")   # escape quotes so O'Brien-type names can't break/inject the OData filter
            Get-MgUser -Filter "startswith(displayName,'$t') or startswith(mail,'$t') or startswith(userPrincipalName,'$t')" -Top 25 -ErrorAction Stop |
                ForEach-Object { [pscustomobject]@{ Display = ('{0}  <{1}>' -f (ConvertTo-SafeText $_.DisplayName), $_.UserPrincipalName); Value = [string]$_.UserPrincipalName } }
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

# True only when input is redirected (piped) AND exhausted. Used to break out of
# prompt loops that would otherwise spin forever on a closed stdin (e.g. the tool
# was run non-interactively). Always false for a normal interactive terminal.
function Test-InputEnded {
    try { return ([Console]::IsInputRedirected -and [Console]::In.Peek() -lt 0) } catch { return $false }
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
    # Render the options. Short lists stay single-column; long ones (e.g. the
    # verb / cmdlet pickers) flow into as many columns as the window is wide,
    # so they don't stretch far down the screen. Numbering is left-to-right,
    # top-to-bottom, and the selection mapping is unchanged (pick by number).
    $labels = @(for ($i = 0; $i -lt $Options.Count; $i++) { '{0,3}. {1}' -f ($i + 1), $Options[$i] })
    $cols = 1
    if ($Options.Count -gt 12) {
        $winW = 80
        try { $w = $Host.UI.RawUI.WindowSize.Width; if ($w -gt 20) { $winW = $w } } catch {}
        $colW = (($labels | Measure-Object -Property Length -Maximum).Maximum) + 3
        $cols = [Math]::Max(1, [Math]::Floor(($winW - 2) / $colW))
        # Only lay out a grid when 2+ columns actually fit. When the labels are so
        # wide that just one column fits, $cols falls back to 1 and we let the
        # single-column block below render it - otherwise the list prints twice
        # (once here as a 1-wide grid, once below).
        if ($cols -gt 1) {
            $rowsN = [Math]::Ceiling($Options.Count / $cols)
            for ($r = 0; $r -lt $rowsN; $r++) {
                $line = ''
                for ($c = 0; $c -lt $cols; $c++) {
                    $idx = $r * $cols + $c
                    if ($idx -lt $labels.Count) { $line += $labels[$idx].PadRight($colW) }
                }
                Write-Host ('  ' + $line.TrimEnd())
            }
        }
    }
    if ($cols -eq 1) {
        foreach ($lbl in $labels) { Write-Host ('  ' + $lbl) }
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
        if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
        $raw = Read-Host 'Choose a number'
        if ($raw -eq '?' -and $Descriptions) {
            for ($i = 0; $i -lt $Options.Count; $i++) {
                Write-Host ('  {0,2}. {1}' -f ($i + 1), $Options[$i]) -ForegroundColor White
                $d = if ($i -lt $Descriptions.Count) { $Descriptions[$i] } else { '' }
                if ($d) { Write-Host ('        {0}' -f $d) -ForegroundColor White }
            }
            continue
        }
        if ($AllowBack -and ($raw -eq '0' -or $raw -eq '')) { return $null }   # Enter = Back, matches Read-Host prompts
        if ($raw -match '^\d{1,5}$' -and [int]$raw -ge 1 -and [int]$raw -le $all.Count) {
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
    # Two vendor modules (SharePoint Admin, Power Platform) are built on .NET
    # Framework and only exist on Windows - they cannot load on macOS/Linux no
    # matter what we do. Fail with a clear message instead of a confusing install
    # or import error. ($IsWindows is undefined on Windows PowerShell 5.1, so test
    # for an explicit $false - that only matches pwsh 6+ on macOS/Linux.)
    if ($Service.WinOnly -and ($IsWindows -eq $false)) {
        throw "'$($Service.Name)' uses a Windows-only module ($($Service.Module)) that Microsoft does not provide for macOS/Linux. Run this service from Windows (PowerShell 5.1 or 7)."
    }
    if ($Service.MinPSVersion -and $PSVersionTable.PSVersion -lt [version]$Service.MinPSVersion) {
        throw "'$($Service.Name)' needs PowerShell $($Service.MinPSVersion) or later (this is $($PSVersionTable.PSVersion))."
    }
    # Modules this service needs. Graph services list the few sub-modules they use
    # instead of the whole Microsoft.Graph meta-module (dozens of packages).
    $mods = if ($Service.InstallModules) { @($Service.InstallModules) } else { @($Service.Module) }
    foreach ($m in $mods) {
        if (Get-Module -ListAvailable -Name $m) { continue }
        Write-Host ("Module '{0}' is not installed." -f $m) -ForegroundColor Yellow
        # Say exactly what would be installed, from where and by whom, before asking.
        $found = $null
        try { $found = Find-Module -Name $m -Repository PSGallery -ErrorAction Stop } catch { }
        if ($found) { Write-Host ("  The PowerShell Gallery has {0} {1} by {2}." -f $found.Name, $found.Version, $found.Author) -ForegroundColor White }
        if ((Read-Host 'Install it now for the current user? (y/n)') -notmatch '^(y|yes)$') { throw "Cannot connect without module '$m'." }
        # From the PowerShell Gallery only, and without -Force, so PowerShellGet's
        # own untrusted-repository prompt (which names the source) still appears.
        Install-Module -Name $m -Repository PSGallery -Scope CurrentUser -ErrorAction Stop
    }
    # Import. Only one known harmless error is ignored: a stale Exchange Online temp
    # format file ('exchange.format.ps1xml') left by an earlier session in the same
    # shell (the toolkit formats its own output). Any other failure is shown.
    foreach ($m in $mods) {
        try { Import-Module $m -ErrorAction Stop -WarningAction SilentlyContinue }
        catch {
            if ("$($_.Exception.Message)" -notmatch 'format\.ps1xml') { throw ("Could not load module '{0}': {1}" -f $m, $_.Exception.Message) }
        }
    }
    $loaded = Get-Module -Name $mods[0] | Sort-Object Version -Descending | Select-Object -First 1
    if ($loaded) {
        Write-Host ("Using {0} {1}" -f $loaded.Name, $loaded.Version) -ForegroundColor White
        if ($Service.MinModuleVersion -and $loaded.Version -lt [version]$Service.MinModuleVersion) {
            throw ("'{0}' needs {1} {2} or later (loaded: {3}). Update it with: Update-Module {1}" -f $Service.Name, $loaded.Name, $Service.MinModuleVersion, $loaded.Version)
        }
    }
    $connectArgs = & $Service.ConnectArgs $Upn $Service
    Write-Host ("Connecting to {0} as {1} ..." -f $Service.Name, $Upn) -ForegroundColor White
    # With -Transcript, pause it across the sign-in so a device code or interactive
    # secret is never written to the log, then resume the same file. Say so if
    # either step fails instead of carrying on silently.
    $transcriptPaused = $false
    if ($script:TranscriptOn) {
        try { Stop-Transcript | Out-Null; $transcriptPaused = $true }
        catch { Write-Host ('  [!] Could not pause the transcript ({0}) - the sign-in may be written to it.' -f $_.Exception.Message) -ForegroundColor Yellow }
    }
    try {
        $script:LastConnectResult = & $Service.ConnectCmd @connectArgs
    } finally {
        if ($transcriptPaused) {
            try { Start-Transcript -Path $script:LogFile -Append | Out-Null }
            catch {
                $script:TranscriptOn = $false
                Write-Host ('[!] Could not resume the transcript ({0}) - the rest of this session is NOT in it.' -f $_.Exception.Message) -ForegroundColor Red
            }
        }
    }
    Write-Host ("Connected to {0}." -f $Service.Name) -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Action log: one JSON line per thing the tool DOES (start, sign-in, sign-out,
# each command run, each bulk row, each undo): who, which service and tenant,
# the command, its target and the outcome. Other parameter values are not kept -
# the Microsoft 365 audit log is the full record of changes. The full screen
# transcript is only written with -Transcript.
# ---------------------------------------------------------------------------
function Write-ActionLog {
    param([Parameter(Mandatory)] [string] $Kind, [string] $Cmdlet, [string] $Target, [string[]] $Params, [string] $Outcome, [string] $Message)
    if (-not $script:ActionLog) { return }
    $entry = [ordered]@{
        time = (Get-Date).ToString('o'); event = $Kind; service = $script:CurrentService
        account = $script:CurrentAccount; tenant = $script:CurrentTenant
        cmdlet = $Cmdlet; target = $Target; params = @($Params | Where-Object { $_ }); outcome = $Outcome; message = $Message
    }
    try {
        [System.IO.File]::AppendAllText($script:ActionLog, ($entry | ConvertTo-Json -Compress -Depth 4) + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
    } catch {
        if (-not $script:ActionLogWarned) {
            $script:ActionLogWarned = $true
            Write-Host ('[!] Could not write the action log ({0}).' -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

# The target of a command, from its key parameters, for screens and the action log.
function Get-TargetText {
    param([System.Collections.IDictionary] $Bound)
    if (-not $Bound) { return '' }
    (@('Identity','User','Mailbox','Member','Trustee','UserId','Id') | Where-Object { $Bound.Contains($_) } |
        ForEach-Object { '{0}={1}' -f $_, (Format-SafetyValue $Bound[$_]) }) -join ', '
}

# Retention: delete logs and the tool's own exports older than 30 days, and drop
# undo lines older than that. Runs at start and on exit; reports what it could
# not delete rather than failing quietly.
function Invoke-RetentionPurge {
    param([int] $Days = 30, [switch] $Quiet)
    $cutoff = (Get-Date).AddDays(-$Days)
    $removed = 0; $failed = 0
    foreach ($sub in 'logs','exports') {
        $dir = Join-Path $script:ToolkitHome $sub
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff -and $_.Name -ne 'undo-hints.log' })) {
            try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $removed++ } catch { $failed++ }
        }
    }
    $undo = Join-Path (Join-Path $script:ToolkitHome 'logs') 'undo-hints.log'
    if (Test-Path -LiteralPath $undo) {
        try {
            $enc   = [System.Text.UTF8Encoding]::new($false)
            $lines = @([System.IO.File]::ReadAllLines($undo, $enc))
            $keep  = @($lines | Where-Object {
                $d = [datetime]::MinValue
                -not ($_ -match '^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d)') -or
                -not [datetime]::TryParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss', [cultureinfo]::InvariantCulture, 'None', [ref]$d) -or
                $d -ge $cutoff })
            if ($keep.Count -ne $lines.Count) { [System.IO.File]::WriteAllLines($undo, [string[]]$keep, $enc); $removed += ($lines.Count - $keep.Count) }
        } catch { $failed++ }
    }
    if (-not $Quiet -and $removed -gt 0) { Write-Host ('Retention: removed {0} log/export file(s) or undo line(s) older than {1} days.' -f $removed, $Days) -ForegroundColor White }
    if ($failed -gt 0) { Write-Host ('[!] Retention: {0} old file(s) could not be deleted.' -f $failed) -ForegroundColor Yellow }
}

# Logs and exports live in the tool's folder. Warn when that folder looks synced
# (OneDrive, iCloud, Dropbox...) or is a network share: the files hold directory data.
function Test-DataFolder {
    $p = [string]$script:ToolkitHome
    if ($p -match '(?i)(OneDrive|iCloud|Mobile Documents|Dropbox|Google Drive|Box Sync)' -or $p.StartsWith('\\')) {
        Write-Host ("[!] The tool's folder ({0}) looks synced or shared. Its logs and exports hold directory data - run the tool from a local folder." -f $p) -ForegroundColor Yellow
    }
}

# Make exported cells safe to open in Excel: a value starting with = + - @ (or a
# tab / CR) is a formula there, so it gets a leading ' and is shown as text.
# Lists are written as "a; b" instead of "System.Object[]".
function ConvertTo-SafeCsvRow {
    param($Items)
    foreach ($it in @($Items)) {
        $o = [ordered]@{}
        foreach ($p in $it.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = (@($v) | ForEach-Object { "$_" }) -join '; ' }
            if ($v -is [string]) { $v = ConvertTo-SafeText $v; if ($v -match '^[=+\-@\t\r]') { $v = "'" + $v } }
            $o[$p.Name] = $v
        }
        [pscustomobject]$o
    }
}

# ---------------------------------------------------------------------------
# The commands "Run any command" offers for a service: that service's own
# commands only. Exchange Online and Purview load their data cmdlets into a temp
# module that Get-ConnectionInformation names; the others list their modules.
# Sign-in commands and raw request cmdlets stay out: they could switch account,
# tenant or scopes, or call any endpoint, outside the tool's gates.
# ---------------------------------------------------------------------------
function Get-ServiceCommands {
    param([Parameter(Mandatory)] $Service)
    $cmds = @()
    if ($Service.Module -eq 'ExchangeOnlineManagement') {
        $isIpps = $Service.ConnectCmd -eq 'Connect-IPPSSession'
        $ci = @(Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' -and ([bool]("$($_.ConnectionUri)" -match 'compliance')) -eq $isIpps })
        foreach ($m in @($ci | ForEach-Object { $_.ModuleName } | Where-Object { $_ } | Select-Object -Unique)) {
            $cmds += @(Get-Command -Module $m -CommandType Cmdlet,Function -ErrorAction SilentlyContinue)
        }
    } else {
        $mods = if ($Service.CommandModules) { $Service.CommandModules } else { @($Service.Module) }
        foreach ($m in $mods) { $cmds += @(Get-Command -Module $m -CommandType Cmdlet,Function -ErrorAction SilentlyContinue) }
    }
    @($cmds | Where-Object { $_.Name -notmatch '^(Connect|Disconnect)-' -and $_.Name -notmatch '^Invoke-.*(Request|RestMethod|GraphMethod)$' } |
        Sort-Object Name -Unique)
}

# The account a service session is really signed in as (the typed email is only a
# hint for some services, and the browser can pick another account or tenant).
# $null when the service gives no way to read it back.
function Get-SignedInAccount {
    param([Parameter(Mandatory)] $Service)
    try {
        if ($Service.Module -eq 'ExchangeOnlineManagement') {
            $isIpps = $Service.ConnectCmd -eq 'Connect-IPPSSession'
            $ci = Get-ConnectionInformation -ErrorAction SilentlyContinue |
                  Where-Object { $_.State -eq 'Connected' -and ([bool]("$($_.ConnectionUri)" -match 'compliance')) -eq $isIpps } | Select-Object -Last 1
            if ($ci) { return [pscustomobject]@{ Account = [string]$ci.UserPrincipalName; Tenant = [string]$ci.Organization; Scopes = $null } }
        } elseif ($Service.ConnectCmd -eq 'Connect-MgGraph') {
            $ctx = Get-MgContext -ErrorAction SilentlyContinue
            if ($ctx) { return [pscustomobject]@{ Account = [string]$ctx.Account; Tenant = [string]$ctx.TenantId; Scopes = @($ctx.Scopes) } }
        } elseif ($Service.ConnectCmd -eq 'Connect-MicrosoftTeams' -and $script:LastConnectResult) {
            $r = $script:LastConnectResult | Select-Object -First 1
            $acct = if ($r.Account.Id) { [string]$r.Account.Id } else { [string]$r.Account }
            if ($acct) { return [pscustomobject]@{ Account = $acct; Tenant = [string]$r.Tenant.Id; Scopes = $null } }
        }
    } catch { }
    return $null
}

# Sign out of one service. Purview shares Exchange Online's sign-out.
function Disconnect-M365Service {
    param([Parameter(Mandatory)] $Service)
    $dc = if ($Service.Disconnect) { $Service.Disconnect } elseif ($Service.Module -eq 'ExchangeOnlineManagement') { 'Disconnect-ExchangeOnline' } else { $null }
    if (-not $dc) { return }
    $cmd = Get-Command $dc -ErrorAction SilentlyContinue
    if (-not $cmd) { return }
    $a = @{}
    if ($cmd.Parameters.ContainsKey('Confirm')) { $a.Confirm = $false }
    try {
        & $dc @a -ErrorAction Stop *> $null
        Write-Host ('Signed out of {0}.' -f $Service.Name) -ForegroundColor White
        Write-ActionLog -Kind 'sign-out' -Message $Service.Name -Outcome 'OK'
    } catch {
        Write-Host ('Could not sign out of {0}: {1}' -f $Service.Name, $_.Exception.Message) -ForegroundColor Yellow
        Write-ActionLog -Kind 'sign-out' -Message ('{0}: {1}' -f $Service.Name, $_.Exception.Message) -Outcome 'FAILED'
    }
}

# Why a command needs a typed confirmation, or $null for a plain y. Read verbs
# never do. Destructive verbs, Invoke- actions (wipe, retire...) and raw non-GET
# requests, changes to admin roles or mail flow, purge/force switches and mail
# forwarding (the classic account-takeover step) all do.
function Get-DestructiveReason {
    param([Parameter(Mandatory)] [string] $Cmdlet, [System.Collections.IDictionary] $Bound)
    $verb, $noun = $Cmdlet -split '-', 2
    if (@('Get','Test','Search','Find','Show','Measure','Compare','Resolve') -contains $verb) { return $null }
    if (@('Remove','Disable','Delete','Reset','Clear','Revoke','Uninstall','Deprovision','Stop','Block','Suspend','Unregister','Purge') -contains $verb) {
        return ('{0} is destructive and has no automatic undo' -f $verb)
    }
    if ($verb -eq 'Invoke') {
        if ($noun -match '(Request|RestMethod|GraphMethod)$') {
            $m = if ($Bound -and $Bound.Contains('Method') -and $Bound['Method']) { "$($Bound['Method'])".ToUpper() } else { 'GET' }
            if ($m -notin @('GET','HEAD')) { return ('it sends a raw {0} request' -f $m) }
        } else { return 'it runs an Invoke- action (such as a wipe or retire)' }
    }
    if ($noun -match '(RoleGroupMember|RoleAssignment|DirectoryRoleMember|RoleMember|TransportRule|JournalRule|InboxRule)') {
        return 'it changes admin roles or mail flow'
    }
    if ($Bound) {
        foreach ($k in @('Purge','PurgeType','PermanentlyDelete','HardDelete','DeleteContent','Force')) {
            if ($Bound.Contains($k) -and $Bound[$k] -and -not ($Bound[$k] -is [switch] -and -not $Bound[$k].IsPresent)) { return ('it uses -{0}' -f $k) }
        }
        foreach ($k in @('ForwardingSmtpAddress','ForwardingAddress','ForwardTo','RedirectTo','ForwardAsAttachmentTo')) {
            if ($Bound.Contains($k) -and $Bound[$k]) { return 'it forwards mail elsewhere' }
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Render one value as a PowerShell literal, so the preview, the final
# confirmation and the logged undo line are exactly what runs if pasted. A
# value can come from the tenant (an undo line carries the CURRENT values of a
# mailbox, which its owner may control), so strings are single-quoted with every
# quote character doubled - PowerShell also treats the typographic quotes
# U+2018 to U+201B as single quotes - and line breaks become spaces.
# ---------------------------------------------------------------------------
function ConvertTo-PSLiteral {
    param($Value)
    if ($null -eq $Value)          { return '$null' }
    if ($Value -is [securestring]) { return "'********'" }
    if ($Value -is [bool])         { return '$' + $Value.ToString().ToLower() }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return $Value.ToString([cultureinfo]::InvariantCulture)
    }
    if ($Value -is [datetime])     { $Value = $Value.ToString('o') }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string] -and $Value -isnot [System.Collections.IDictionary]) {
        $parts = @($Value | ForEach-Object { ConvertTo-PSLiteral $_ })
        if ($parts.Count -eq 0) { return '@()' }
        return ($parts -join ',')
    }
    $s = ([string]$Value) -replace '[\r\n]+', ' '
    $s = $s -replace "['\u2018\u2019\u201A\u201B]", '$0$0'
    return "'" + $s + "'"
}

# ---------------------------------------------------------------------------
# Render an assembled splat as a copy-pasteable command line (for the preview).
# ---------------------------------------------------------------------------
function Format-CommandPreview {
    param(
        [Parameter(Mandatory)] [string]                       $Cmdlet,
        [Parameter(Mandatory)] [System.Collections.IDictionary] $BoundParams,   # IDictionary (not [hashtable]) so an [ordered] dict keeps its insertion order
        [string[]] $Secret = @()   # param names whose values are secrets - render masked, never in plaintext
    )
    $sb = [System.Text.StringBuilder]::new($Cmdlet)
    foreach ($key in $BoundParams.Keys) {
        $val = $BoundParams[$key]
        if ($Secret -contains $key) {
            [void]$sb.Append(" -$key '********'")   # secret: masked so it never reaches the screen/transcript
        } elseif ($val -is [switch]) {
            if ($val.IsPresent) { [void]$sb.Append(" -$key") }
        } else {
            [void]$sb.Append(" -$key " + (ConvertTo-PSLiteral $val))
        }
    }
    $sb.ToString()
}

# ---------------------------------------------------------------------------
# Print a built command line in colour so the cmdlet, its -parameters, and their
# values are easy to tell apart. Cmdlet = Cyan, -Params = Yellow, values = White.
# Splits only on ' -Param' boundaries, so quoted values with spaces stay intact.
# Cosmetic only - logging/undo files still use the plain Format-CommandPreview string.
# ---------------------------------------------------------------------------
function Write-CommandColored {
    param([Parameter(Mandatory)] [string] $Command, [string] $Indent = '  ')
    Write-Host $Indent -NoNewline
    $rest = $Command
    if ($Command -match '^(\S+)(.*)$') {
        Write-Host $Matches[1] -ForegroundColor Cyan -NoNewline
        $rest = $Matches[2]
    }
    # Walk char-by-char tracking quote state, so a "-word" sitting INSIDE a quoted
    # value (e.g. -Filter 'Members -eq ...') is not mistaken for a real -Parameter.
    $i = 0; $n = $rest.Length; $quote = $null
    while ($i -lt $n) {
        $ch = $rest[$i]
        if (-not $quote -and [char]::IsWhiteSpace($ch) -and ($i + 2) -lt $n -and $rest[$i + 1] -eq '-' -and [char]::IsLetter([char]$rest[$i + 2])) {
            $j = $i + 2
            while ($j -lt $n -and ("$($rest[$j])" -match '\w')) { $j++ }
            Write-Host $rest.Substring($i, $j - $i) -ForegroundColor Yellow -NoNewline
            $i = $j; continue
        }
        if ($ch -eq "'" -or $ch -eq '"') {
            if (-not $quote) { $quote = $ch } elseif ($quote -eq $ch) { $quote = $null }
        }
        Write-Host $ch -ForegroundColor White -NoNewline
        $i++
    }
    Write-Host ''
}

# ---------------------------------------------------------------------------
# Change safety net: for a Set-* cmdlet, read the CURRENT values of the params
# about to change (via the matching Get-* cmdlet) so the tool can show what is
# being overwritten and, after the change, print/log an exact command to revert
# it. Best-effort: returns Supported=$false if it can't map the pair cleanly.
# ---------------------------------------------------------------------------
function Get-RevertPlan {
    param(
        [Parameter(Mandatory)] [string]    $Cmdlet,
        [Parameter(Mandatory)] [hashtable] $Bound
    )
    $plan = [pscustomobject]@{ Supported = $false; Keys = @{}; Before = @{}; Missing = @() }
    if ($Cmdlet -notlike 'Set-*') { return $plan }
    $getCmd = 'Get-' + $Cmdlet.Substring(4)
    $getCommand = Get-Command -Name $getCmd -ErrorAction SilentlyContinue
    if (-not $getCommand) { return $plan }

    # A record may be located by a COMPOUND key. Set-MailboxFolderPermission, for
    # example, is keyed by Identity (the calendar) AND User (whose entry) - reading
    # only Identity returns EVERY entry on the calendar and Select -First 1 grabs
    # the Default row, so the "before" and the undo would silently target Default
    # instead of the user we changed. Capture ALL bound key params and query with
    # every one the Get- cmdlet accepts, so we read the single matching record.
    $keyNames = @(@('Identity','Id','User','Mailbox') | Where-Object { $Bound.Contains($_) })
    if ($keyNames.Count -eq 0) { return $plan }
    $getParamNames = @($getCommand.Parameters.Keys)   # null-safe: empty array if Parameters is null
    # A key the Get- cmdlet can't take means it can't be told WHICH record to read,
    # so a "before" it returns could belong to another record. Decline the undo.
    if (@($keyNames | Where-Object { $getParamNames -notcontains $_ }).Count -gt 0) { return $plan }
    $getParams = @{}
    foreach ($k in $keyNames) { $getParams[$k] = $Bound[$k] }
    $found = @()
    try {
        $found = @(& $getCmd @getParams -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object -First 2)
    } catch {
        # Positional fallback only for a SINGLE key - for a compound key a lone
        # positional value would match the wrong record (e.g. Default), so we
        # decline the undo entirely rather than generate a corrupting one.
        if ($keyNames.Count -eq 1) {
            try { $found = @(& $getCmd $Bound[$keyNames[0]] -WarningAction SilentlyContinue -ErrorAction Stop | Select-Object -First 2) } catch { return $plan }
        } else { return $plan }
    }
    # Exactly one record, or no undo: with two matches we can't tell which one changes.
    if ($found.Count -ne 1) { return $plan }
    $cur    = $found[0]
    $props  = @($cur.PSObject.Properties.Name)
    $before = @{}
    $missing = @()
    foreach ($k in @($Bound.Keys)) {
        if ($keyNames -contains $k)   { continue }   # keys locate the record; they are replayed verbatim, never reverted
        if ($Bound[$k] -is [switch])  { continue }   # switches aren't a simple before/after value
        if ($props -contains $k)      { $before[$k] = $cur.$k } else { $missing += $k }
    }
    if ($before.Count -eq 0) { return $plan }
    $plan.Missing = $missing   # changed values the undo will NOT restore (shown to the user)
    # Replay the exact key values the user typed (clean, bindable) - not $cur's
    # key properties, which are display objects that may not re-bind.
    $keys = @{}
    foreach ($k in $keyNames) { $keys[$k] = $Bound[$k] }
    $plan.Supported = $true; $plan.Keys = $keys; $plan.Before = $before
    return $plan
}

# Human-readable rendering of a before/after value for the safety display.
function Format-SafetyValue {
    param($Value)
    if ($null -eq $Value) { return '(empty)' }
    if ($Value -is [array]) { if ($Value.Count -eq 0) { return '(empty)' } else { return (($Value | ForEach-Object { "$_" }) -join ', ') } }
    $s = "$Value"
    if ([string]::IsNullOrWhiteSpace($s)) { return '(empty)' }
    return $s
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
# Shared value rules for the guided and bulk runners, so both treat allowed
# values, booleans, lists, secrets and the run gate the same way.
# ---------------------------------------------------------------------------

# The fixed set of allowed values for a parameter, from (in order) a
# [ValidateSet], the values of an enum type, the action's Hints, or the global
# $ParamHints table. $null when the field is free text.
function Get-AllowedValues {
    param([Parameter(Mandatory)] $Param, [hashtable] $Hints)
    $type = $Param.ParameterType
    $elem = if ($type.IsArray) { $type.GetElementType() } else { $type }
    $vset = $Param.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | Select-Object -First 1
    if     ($vset)                                       { return ,([string[]]$vset.ValidValues) }
    elseif ($elem.IsEnum)                                { return ,([string[]][enum]::GetNames($elem)) }
    elseif ($Hints -and $Hints.ContainsKey($Param.Name)) { return ,([string[]]$Hints[$Param.Name]) }
    elseif ($ParamHints.ContainsKey($Param.Name))        { return ,([string[]]$ParamHints[$Param.Name]) }
    return $null
}

# Secret fields are read hidden and masked in previews and logs: a
# [SecureString] parameter, or one whose name looks like a credential.
function Test-SecretParam {
    param([Parameter(Mandatory)] $Param)
    $type = $Param.ParameterType
    $elem = if ($type.IsArray) { $type.GetElementType() } else { $type }
    if ($elem -eq [securestring] -or $elem -eq [pscredential]) { return $true }
    # A credential-like NAME only counts for text-like types, so a Boolean such
    # as -ResetPasswordOnNextLogon is still asked as true / false.
    return (($Param.Name -match $script:SecretNamePattern) -and ($elem -eq [string] -or $elem -eq [object]))
}

# A credential-like parameter of a structured type (e.g. a password profile
# object): the builder can't enter one safely, so it is skipped, never echoed.
function Test-StructuredSecretParam {
    param([Parameter(Mandatory)] $Param)
    $type = $Param.ParameterType
    $elem = if ($type.IsArray) { $type.GetElementType() } else { $type }
    if (-not ($Param.Name -match $script:SecretNamePattern) -or (Test-SecretParam -Param $Param)) { return $false }
    $simple = $elem.IsPrimitive -or $elem.IsEnum -or $elem -eq [decimal] -or $elem -eq [datetime] -or $elem -eq [switch]
    return (-not $simple)
}

# Turn text typed at a prompt, or read from a CSV cell, into a value the
# parameter binder accepts. Returns @{ Ok; Value; Error }. Booleans map yes/no
# words explicitly ('false' as a string is truthy, so never cast). List fields
# split on commas only, so "Sales Team" stays one value. Values with an allowed
# set are checked and normalised to the canonical casing.
function ConvertTo-ParamValue {
    param([Parameter(Mandatory)] $Param, [AllowEmptyString()] [string] $Text, [string[]] $Allowed)
    $type = $Param.ParameterType
    $elem = if ($type.IsArray) { $type.GetElementType() } else { $type }
    $t    = $Text.Trim()
    if ($elem -eq [bool]) {
        $k = $t.ToLower()
        if ($k -in @('true','yes','y','1'))  { return @{ Ok = $true; Value = $true } }
        if ($k -in @('false','no','n','0')) { return @{ Ok = $true; Value = $false } }
        return @{ Ok = $false; Error = ("'{0}' is not true or false" -f $t) }
    }
    $vals = @(if ($type.IsArray) { $t -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { $t })
    if ($Allowed) {
        # -contains is case-insensitive, so 'sendas' passes and is normalised to 'SendAs'.
        $bad = @($vals | Where-Object { $Allowed -notcontains $_ })
        if ($bad.Count -gt 0) { return @{ Ok = $false; Error = ("not valid: {0}. Choose from: {1}" -f ($bad -join ', '), ($Allowed -join ', ')) } }
        $vals = @($vals | ForEach-Object { $v = $_; $Allowed | Where-Object { $_ -eq $v } | Select-Object -First 1 })
    }
    if ($type.IsArray) { return @{ Ok = $true; Value = [object[]]$vals } }
    return @{ Ok = $true; Value = $vals[0] }
}

# The run gate shared by the guided and bulk runners. Destructive verbs have no
# automatic undo (the safety net only covers Set-*) and any cmdlet can reach
# this point via "Run any command", so they need the command name typed - plus
# the row count in bulk, where one keypress would otherwise hit every row.
# Everything else needs an explicit y. Returns $true to go ahead.
function Confirm-Execution {
    param([Parameter(Mandatory)] [string] $Cmdlet, [int] $RowCount = 0, [System.Collections.IDictionary] $Bound)
    $reason = Get-DestructiveReason -Cmdlet $Cmdlet -Bound $Bound
    if ($reason) {
        Write-Host ''
        if ($RowCount -gt 0) {
            Write-Host ('[!] {0}: {1}, and it will run for {2} row(s).' -f $Cmdlet, $reason, $RowCount) -ForegroundColor Red
            $expect = '{0} {1}' -f $Cmdlet, $RowCount
        } else {
            Write-Host ('[!] {0}: {1}.' -f $Cmdlet, $reason) -ForegroundColor Red
            $expect = $Cmdlet
        }
        $typed = ("$(Read-Host ('To proceed, type exactly: {0}  (anything else cancels)' -f $expect))".Trim()) -replace '\s+', ' '
        return ($typed -eq $expect)
    }
    $q = if ($RowCount -gt 0) { 'Execute for all {0} row(s) now? (y/n)' -f $RowCount } else { 'Execute it now? (y/n)' }
    return ((Read-Host $q) -match '^(y|yes)$')
}

# Read a CSV for the bulk runner. Excel often saves ANSI rather than UTF-8 and,
# in EU locales, uses ';' as the separator. Decode as strict UTF-8 first; if the
# bytes are not UTF-8, decode as Windows-1252; if the header has ';' but no ','
# read with ';'. Says what it did.
function Import-BulkCsv {
    param([Parameter(Mandatory)] [string] $Path)
    $bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Path).ProviderPath)
    try {
        $text = [System.Text.UTF8Encoding]::new($false, $true).GetString($bytes)   # strict: throws on bytes that are not UTF-8
    } catch {
        $enc = $null
        try { $enc = [System.Text.Encoding]::GetEncoding(1252) } catch { $enc = [System.Text.Encoding]::GetEncoding('iso-8859-1') }
        $text = $enc.GetString($bytes)
        Write-Host ('  The file is not UTF-8 - reading it as {0} (the usual Excel "CSV" encoding).' -f $enc.WebName) -ForegroundColor Yellow
    }
    $text   = $text.TrimStart([char]0xFEFF)   # drop a byte order mark
    $header = ($text -split "`r?`n", 2)[0]
    $delim  = ','
    if ($header -notmatch ',' -and $header -match ';') {
        $delim = ';'
        Write-Host "  The file uses ';' as the separator - reading it that way." -ForegroundColor Yellow
    }
    @($text | ConvertFrom-Csv -Delimiter $delim)
}

# Replace each given secret value inside a message with ********, so a binder or
# service error that echoes its input cannot leak a secret to screen or log.
# Mask secrets in command OUTPUT before it is shown, piped with 'last' or saved:
# bearer tokens in text, and properties whose name looks like a credential.
function Hide-SecretResult {
    param($Items)
    foreach ($it in @($Items)) {
        if ($null -eq $it) { continue }
        if ($it -is [string]) { (ConvertTo-SafeText $it) -replace $script:JwtPattern, '********'; continue }
        $names = @($it.PSObject.Properties.Name)
        $hit = @($names | Where-Object { $_ -match $script:SecretNamePattern })
        $jwt = @($names | Where-Object { $hit -notcontains $_ -and $it.$_ -is [string] -and $it.$_ -match $script:JwtPattern })
        $ctl = @($names | Where-Object { $it.$_ -is [string] -and $it.$_ -match '[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]' })
        if ($hit.Count -eq 0 -and $jwt.Count -eq 0 -and $ctl.Count -eq 0) { $it; continue }
        $o = [ordered]@{}
        foreach ($n in $names) {
            $o[$n] = if ($hit -contains $n) { '********' }
                     elseif ($jwt -contains $n) { (ConvertTo-SafeText $it.$n) -replace $script:JwtPattern, '********' }
                     elseif ($ctl -contains $n) { ConvertTo-SafeText $it.$n }
                     else { $it.$n }
        }
        [pscustomobject]$o
    }
}

function Hide-SecretText {
    param([string] $Text, [string[]] $Values)
    $out = "$Text"
    foreach ($v in @($Values | Where-Object { $_ })) { $out = $out.Replace($v, '********') }
    return $out
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
        [scriptblock] $Project,  # optional reshaping of the RESULT before display (raw result -> clean objects)
        [hashtable] $Candidates, # optional dependent pick-lists (name -> { param($bound) ... }): valid values derived from fields already chosen
        [switch]    $Resolve     # check typed identities against the directory (catalogued actions; run-any skips it)
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
    $secretKeys = [System.Collections.Generic.List[string]]::new()   # params read as masked secrets; masked in preview + undo log

    # ---- Presets: values fixed by the action itself (e.g. -Shared for a shared
    #      mailbox, -Type Shared to convert). Bound automatically, shown in the
    #      preview, never prompted for.
    if ($Preset) {
        foreach ($k in $Preset.Keys) {
            $pv = $Preset[$k]
            # A switch preset written as $true should bind and preview as a switch
            # (-Shared), not as -Shared $true. Coerce bool -> switch for switch params.
            if ($pv -is [bool] -and ($switches.Name -contains $k)) { $pv = [switch]$pv }
            $bound[$k] = $pv
        }
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
        $val = Read-ParameterValue -Param $p -Required -Description $desc -Hints $Hints -Candidates $Candidates -Bound $bound -Resolve:$Resolve
        if ($script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
        if ($script:ParamBack) {
            if ($ri -eq 0) { Write-Host "  Already at the first field - type 'cancel' to return to the menu." -ForegroundColor White; continue }
            $ri--
            $bound.Remove($reqList[$ri].Name)
            Write-Host ("  Going back to -{0}." -f $reqList[$ri].Name) -ForegroundColor White
            continue
        }
        $bound[$p.Name] = $val
        if ($script:ParamSecret) { $secretKeys.Add($p.Name) }
        $ri++
    }

    # ---- Optional params named in the curated list: ask them outright.
    foreach ($p in ($optional | Where-Object { $promptList -contains $_.Name -and -not $bound.Contains($_.Name) })) {
        $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
        $v = Read-ParameterValue -Param $p -Description $desc -Hints $Hints -Candidates $Candidates -Bound $bound -Resolve:$Resolve
        # 'cancel' aborts the whole action; 'back' on an OPTIONAL field just skips it.
        if ($script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
        if (-not $script:ParamBack -and $null -ne $v) { $bound[$p.Name] = $v; if ($script:ParamSecret) { $secretKeys.Add($p.Name) } }
    }

    # ---- Remaining optional params: offer the list, let the user pick which to set.
    $optRest = @($optional | Where-Object { -not $bound.Contains($_.Name) })
    if ($optRest.Count -eq 1 -and $ParamLookup.ContainsKey($optRest[0].Name)) {
        # Single identity-like optional field (e.g. Get-*Configuration -Identity):
        # skip the "which numbers do you want to set?" step and ask for its value directly.
        $p = $optRest[0]
        Write-Host ''
        Write-Host ("Optional: {0} - enter it now, or press Enter to skip." -f $p.Name) -ForegroundColor White
        $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
        $v = Read-ParameterValue -Param $p -Description $desc -Hints $Hints -Candidates $Candidates -Bound $bound -Resolve:$Resolve
        # 'cancel' aborts the whole action; 'back' on an OPTIONAL field just skips it.
        if ($script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
        if (-not $script:ParamBack -and $null -ne $v) { $bound[$p.Name] = $v; if ($script:ParamSecret) { $secretKeys.Add($p.Name) } }
    }
    elseif ($optRest.Count -gt 0) {
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
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })) {
            $idx = [int]$n - 1
            if ($idx -ge 0 -and $idx -lt $optRest.Count) {
                $p = $optRest[$idx]
                $desc = Get-ParamDescription -Name $p.Name -ParamHelp $ParamHelp -Cmdlet $Cmdlet
                $v = Read-ParameterValue -Param $p -Description $desc -Hints $Hints -Candidates $Candidates -Bound $bound -Resolve:$Resolve
                # 'cancel' aborts the whole action; 'back' just skips this optional pick.
                if ($script:ParamCancel) { Write-Host 'Cancelled - back to the menu (nothing was run).' -ForegroundColor White; return }
                if (-not $script:ParamBack -and $null -ne $v) { $bound[$p.Name] = $v; if ($script:ParamSecret) { $secretKeys.Add($p.Name) } }
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
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })) {
            $idx = [int]$n - 1
            if ($idx -ge 0 -and $idx -lt $swRest.Count) { $bound[$swRest[$idx].Name] = [switch]$true }   # [switch] so preview shows -Force, not -Force $true
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

    # ---- Change safety net: for a Set-*, capture what the change overwrites and
    #      show current -> new so you see exactly what you are about to replace.
    $revertPlan = Get-RevertPlan -Cmdlet $Cmdlet -Bound $bound
    if ($revertPlan.Supported) {
        Write-Host ''
        Write-Host 'Current value(s) that will change:' -ForegroundColor Yellow
        foreach ($k in $revertPlan.Before.Keys) {
            $mask = ($secretKeys -contains $k) -or ($k -match $script:SecretNamePattern)
            $old  = if ($mask) { '********' } else { Format-SafetyValue $revertPlan.Before[$k] }
            $new  = if ($mask) { '********' } else { Format-SafetyValue $bound[$k] }
            Write-Host ('  {0}:  {1}  ->  {2}' -f $k, $old, $new) -ForegroundColor White
        }
        if ($revertPlan.Missing.Count -gt 0) {
            Write-Host ('  Not covered by the undo (no current value to read back): {0}' -f ($revertPlan.Missing -join ', ')) -ForegroundColor Yellow
        }
    }

    # ---- Dry-run preview + final confirm. -----------------------------------
    $preview = Format-CommandPreview -Cmdlet $Cmdlet -BoundParams $bound -Secret @($secretKeys)
    Write-Host ''
    Write-Host 'Preview (dry-run):' -ForegroundColor Yellow
    Write-CommandColored -Command $preview

    if ($DryRunOnly) {
        Write-Host 'DryRunOnly is set - not executing.' -ForegroundColor White
        return
    }

    # Last confirmation: repeat the EXACT command, then require an explicit y.
    Write-Host ''
    Write-Host 'FINAL CONFIRMATION - about to execute this command:' -ForegroundColor Yellow
    Write-CommandColored -Command $preview
    # Destructive verbs need the command name typed; everything else an explicit y.
    # Same gate as the bulk runner (Confirm-Execution).
    if (-not (Confirm-Execution -Cmdlet $Cmdlet -Bound $bound)) {
        Write-Host 'Cancelled - nothing was run.' -ForegroundColor White
        return
    }

    Write-Host 'Running...' -ForegroundColor White
    $script:LastRevert = $null   # a new change makes any earlier undo stale
    try {
        $wv = $null
        $result = & $Cmdlet @bound -ErrorAction Stop -WarningVariable wv
        # A warning often means nothing changed (e.g. removing an entry that does not exist).
        if (@($wv).Count -gt 0) { Write-Host 'Done, with warning(s) above - check whether the change was actually made.' -ForegroundColor Yellow }
        else                    { Write-Host 'Done.' -ForegroundColor Green }
        Write-ActionLog -Kind 'run' -Cmdlet $Cmdlet -Target (Get-TargetText -Bound $bound) -Params @($bound.Keys) `
            -Outcome $(if (@($wv).Count -gt 0) { 'WARNING' } else { 'OK' }) -Message ((@($wv) | ForEach-Object { "$_" }) -join ' | ')
        # Change safety net: print + log the exact command to undo this change.
        if ($revertPlan.Supported) {
            $rb = [ordered]@{}
            foreach ($k in $revertPlan.Keys.Keys)   { $rb[$k] = $revertPlan.Keys[$k] }
            foreach ($k in $revertPlan.Before.Keys) { $rb[$k] = $revertPlan.Before[$k] }
            # Mask any revert param whose NAME looks like a secret too - the "before"
            # values come straight from Get-RevertPlan's Get-* snapshot (not typed by
            # the user), so they aren't in $secretKeys; without this a credential-ish
            # property could land unmasked in undo-hints.log / on screen.
            $revSecret = @($secretKeys) + @($rb.Keys | Where-Object { $_ -match $script:SecretNamePattern })
            $revCmd = Format-CommandPreview -Cmdlet $Cmdlet -BoundParams $rb -Secret @($revSecret)
            Write-Host ''
            Write-Host 'To undo this change, run:' -ForegroundColor Yellow
            Write-CommandColored -Command $revCmd
            if ($revertPlan.Missing.Count -gt 0) { Write-Host ('  (partial: does not restore {0})' -f ($revertPlan.Missing -join ', ')) -ForegroundColor Yellow }
            # Keep the line in logs\undo-hints.log: owner-only, newest 500, and lines older
            # than 30 days are dropped by the retention purge. One encoding (UTF-8, no
            # BOM) for reading and writing, so non-ASCII names survive on 5.1 too.
            $undoFile = Join-Path (Join-Path $script:ToolkitHome 'logs') 'undo-hints.log'
            try {
                $undoDir = Split-Path -Parent $undoFile
                if (-not (Test-Path -LiteralPath $undoDir)) { New-Item -ItemType Directory -Path $undoDir | Out-Null }
                $enc = [System.Text.UTF8Encoding]::new($false)
                if (-not (Test-Path -LiteralPath $undoFile)) { New-Item -ItemType File -Path $undoFile | Out-Null }
                [void](Protect-LocalPath -Path $undoFile)   # locked down before anything is written
                $lines = @([System.IO.File]::ReadAllLines($undoFile, $enc)) + @(('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $revCmd))
                if ($lines.Count -gt 500) { $lines = $lines[-500..-1] }
                [System.IO.File]::WriteAllLines($undoFile, [string[]]$lines, $enc)
                Write-Host ('  (also saved to {0})' -f $undoFile) -ForegroundColor White
            } catch {
                Write-Host ('  [!] Could not save the undo line to {0}: {1}' -f $undoFile, $_.Exception.Message) -ForegroundColor Yellow
            }
            # Stash the built revert (already a bindable splat) so the user can run
            # it ON DEMAND from the "What next?" menu - no prompt after every change,
            # and no retyping/pasting the command (the tool has no paste-a-command
            # entry by design). Each successful Set replaces the previous stash.
            $script:LastRevert = @{
                Cmdlet  = $Cmdlet; Splat = $rb; Preview = $revCmd
                Target  = (@($revertPlan.Keys.Keys | ForEach-Object { '{0}={1}' -f $_, (Format-SafetyValue $revertPlan.Keys[$_]) }) -join ', ')
                Time    = Get-Date
                Service = $script:CurrentService
                Partial = ($revertPlan.Missing.Count -gt 0)
            }
        } elseif ($Cmdlet -like 'Set-*') {
            Write-Host '  No automatic undo for this change (its current values could not be read back safely).' -ForegroundColor Yellow
        }
        # Optional per-action projection: reshape a raw/cluttered result into clean,
        # readable columns (e.g. flatten Get-MgUserMemberOf into Name/Email/Type)
        # before it is shown AND before it becomes the pipe source for 'last'.
        if ($Project) {
            try   { $result = & $Project $result }
            catch { Write-Host ('  (could not reshape result: {0})' -f $_.Exception.Message) -ForegroundColor Yellow }
        }
        # The command already ran: a problem showing or saving its result must not
        # read as "Failed" (and invite a retry of a change that was made).
        try { Show-Result -Result $result -Cmdlet $Cmdlet }
        catch [System.IO.EndOfStreamException] { throw }
        catch { Write-Host ('The command ran, but showing or saving its result failed: {0}' -f $_.Exception.Message) -ForegroundColor Yellow }
    } catch [System.IO.EndOfStreamException] { throw }
    catch {
        $msg = Hide-SecretText -Text $_.Exception.Message -Values @($secretKeys | ForEach-Object { $bound[$_] } | Where-Object { $_ -is [string] })
        Write-Host ('Failed: {0}' -f $msg) -ForegroundColor Red
        Write-ActionLog -Kind 'run' -Cmdlet $Cmdlet -Target (Get-TargetText -Bound $bound) -Params @($bound.Keys) -Outcome 'FAILED' -Message $msg
    }
}

# ---------------------------------------------------------------------------
# Bulk runner: run the SAME cmdlet once per row of a CSV file. You map each
# parameter to a CSV column (per-row value) or a fixed value (same every row).
# Every row is converted and checked BEFORE anything runs, with the same value
# rules as the guided path; a blank cell is skipped unless you choose to clear;
# the targets are listed; and the same run gate applies (a destructive verb needs
# the command name AND the row count typed). Each row's outcome, and for a Set-*
# the values it replaced, goes to a results CSV.
# ---------------------------------------------------------------------------
function Invoke-BulkCmdlet {
    param(
        [Parameter(Mandatory)] [string] $Cmdlet,
        [string[]]  $Prompt,
        [hashtable] $ParamHelp,
        [hashtable] $Preset,
        [hashtable] $Hints,      # optional per-action allowed-value overrides (name -> values)
        [hashtable] $Transform,  # optional post-processing of each row's values (name -> scriptblock)
        [switch]    $Resolve     # check identity values against the directory (catalogued actions; run-any skips it)
    )
    $cmd   = Get-Command $Cmdlet -ErrorAction Stop
    $parts = Get-CmdletParameters -Cmd $cmd -PreferParams $Prompt
    $byName = @{}
    foreach ($p in (@($parts.Required) + @($parts.Optional) + @($parts.Switches))) { if ($p) { $byName[$p.Name] = $p } }
    $requiredNames = @($parts.Required | ForEach-Object { $_.Name })

    $csvPath = ("$(Read-Host 'Path to CSV file (must have a header row)')").Trim('"',' ')
    if (-not $csvPath -or -not (Test-Path -LiteralPath $csvPath -PathType Leaf)) { Write-Host "File not found: $csvPath" -ForegroundColor Red; return }
    try   { $rows = @(Import-BulkCsv -Path $csvPath) }
    catch { Write-Host ('Could not read that CSV: {0}' -f $_.Exception.Message) -ForegroundColor Red; return }
    if ($rows.Count -eq 0) { Write-Host 'That CSV has no data rows.' -ForegroundColor Yellow; return }
    $columns = @($rows[0].PSObject.Properties.Name)

    Write-Host ''
    Write-Host ("Loaded {0} rows. Columns: {1}" -f $rows.Count, ($columns -join ', ')) -ForegroundColor White
    Write-Host 'For each parameter: type a COLUMN name (uses each row''s value), or =value for a fixed value on every row, or press Enter to skip.' -ForegroundColor White

    # map: paramName -> @{ Kind = 'column' | 'fixed' | 'switch'; Value; Allowed; Secret }
    $map = [ordered]@{}
    if ($Preset) {
        foreach ($k in $Preset.Keys) {
            $pv = $Preset[$k]
            # A switch preset written as $true binds and previews as a switch (-Shared), not -Shared $true.
            if ($pv -is [bool] -and (@($parts.Switches | ForEach-Object { $_.Name }) -contains $k)) {
                if ($pv) { $map[$k] = @{ Kind = 'switch' } }
                continue
            }
            $map[$k] = @{ Kind = 'fixed'; Value = $pv }
        }
    }

    $askParam = {
        param($p, $required)
        if ($map.Contains($p.Name)) { return }
        $desc = if ($ParamHelp) { $ParamHelp[$p.Name] } else { $null }
        if ($desc) { Write-Host ("  -> {0}" -f $desc) -ForegroundColor White }
        # A credential needs a user name AND a password; one CSV column can't carry it.
        $pType = if ($p.ParameterType.IsArray) { $p.ParameterType.GetElementType() } else { $p.ParameterType }
        if ($pType -eq [pscredential] -or (Test-StructuredSecretParam -Param $p)) {
            if ($required) { throw ("Bulk mode can't fill -{0} (a credential or structured secret). Run this command once instead." -f $p.Name) }
            Write-Host ("  -{0} can't be filled in bulk mode - skipped." -f $p.Name) -ForegroundColor Yellow; return
        }
        # Same allowed-value and secret rules as the guided prompt.
        $allowed  = Get-AllowedValues -Param $p -Hints $Hints
        $isSecret = Test-SecretParam -Param $p
        if ($allowed)  { Write-Host ('     allowed values: ' + ($allowed -join ', ')) -ForegroundColor White }
        if ($isSecret) { Write-Host '     secret: map a column, or type = on its own to enter one fixed value hidden' -ForegroundColor White }
        while ($true) {
            $ans = "$(Read-Host ("Map -{0} (column / =fixed / Enter){1}" -f $p.Name, $(if ($required) { ' [required]' } else { '' })))"
            if ([string]::IsNullOrWhiteSpace($ans)) {
                if ($required) {
                    if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
                    Write-Host '  Required - give a column name or =value.' -ForegroundColor Yellow; continue
                }
                return
            }
            if ($ans.StartsWith('=')) {
                if ($isSecret) {
                    # A fixed secret typed after '=' would already be on screen and in the transcript.
                    if ($ans -ne '=') { Write-Host '  For a secret, type = on its own and enter the value hidden. What you typed was not used.' -ForegroundColor Yellow; continue }
                    $sec = Read-Host ("Fixed -{0} for every row (hidden)" -f $p.Name) -AsSecureString
                    $len = 0; try { $len = $sec.Length } catch { }
                    if ($len -eq 0) { continue }
                    $map[$p.Name] = @{ Kind = 'fixed'; Value = $sec; Secret = $true }; return
                }
                $conv = ConvertTo-ParamValue -Param $p -Text $ans.Substring(1) -Allowed $allowed
                if (-not $conv.Ok) { Write-Host ("  {0}." -f $conv.Error) -ForegroundColor Yellow; continue }
                $map[$p.Name] = @{ Kind = 'fixed'; Value = $conv.Value }; return
            }
            if ($columns -contains $ans) { $map[$p.Name] = @{ Kind = 'column'; Value = $ans; Allowed = $allowed; Secret = $isSecret }; return }
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
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })) {
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
        foreach ($n in ($pick -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })) {
            $idx=[int]$n-1; if ($idx -ge 0 -and $idx -lt $swRest.Count) { $map[$swRest[$idx].Name] = @{ Kind = 'switch' } }
        }
    }
    if ($map.Count -eq 0) { Write-Host 'Nothing mapped - nothing to run.' -ForegroundColor Yellow; return }

    # Blank cells: in the guided path an empty answer means "skip", so by default a
    # blank cell skips that parameter on that row. Clearing needs an explicit choice.
    $blankClears = $false
    $blankCols = @(foreach ($k in $map.Keys) {
        $m = $map[$k]
        if ($m.Kind -ne 'column' -or $requiredNames -contains $k) { continue }
        $n = @($rows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.($m.Value)) }).Count
        if ($n -gt 0) { [pscustomobject]@{ Param = $k; Column = $m.Value; Blank = $n } }
    })
    if ($blankCols.Count -gt 0) {
        Write-Host ''
        Write-Host 'Some mapped columns have blank cells:' -ForegroundColor Yellow
        foreach ($bc in $blankCols) { Write-Host ('  -{0} (column {1}): {2} of {3} rows blank' -f $bc.Param, $bc.Column, $bc.Blank, $rows.Count) -ForegroundColor White }
        $bq = Get-MenuChoice -Title 'What should a blank cell do?' -Options @(
            'Skip that parameter on that row (nothing changes for it) - recommended',
            'Send it as empty, to CLEAR the current value')
        $blankClears = ($bq -like 'Send it as empty*')
    }

    # Build and check EVERY row before anything runs, with the guided path's rules:
    # value conversion, allowed values, required fields, and (for catalogued
    # actions) the directory check on identity fields.
    $lookupOk = @{}
    if ($Resolve) {
        foreach ($k in $map.Keys) {
            $lk = $ParamLookup[$k]; $p = $byName[$k]
            $lookupOk[$k] = [bool]($map[$k].Kind -eq 'column' -and $lk -and $lk.Resolve -and $p -and -not $p.ParameterType.IsArray -and
                                   (Get-Command $lk.Probe -ErrorAction SilentlyContinue))
        }
    }
    $keyNames = @('Identity','User','Mailbox','Member','Trustee','UserId','Id')
    $plans = [System.Collections.Generic.List[object]]::new()
    $rowNo = 0
    foreach ($row in $rows) {
        $rowNo++
        if ($Resolve -and $rows.Count -gt 25 -and ($rowNo % 25) -eq 0) { Write-Host ('  checking row {0} of {1}...' -f $rowNo, $rows.Count) -ForegroundColor White }
        $b = [ordered]@{}
        $problems   = [System.Collections.Generic.List[string]]::new()
        $secretText = [System.Collections.Generic.List[string]]::new()
        foreach ($k in $map.Keys) {
            $m = $map[$k]; $p = $byName[$k]
            $elem = if ($p) { if ($p.ParameterType.IsArray) { $p.ParameterType.GetElementType() } else { $p.ParameterType } } else { $null }
            if ($m.Kind -eq 'switch') { $b[$k] = [switch]$true; continue }
            if ($m.Kind -eq 'fixed') {
                $v = $m.Value
                if ($v -is [securestring] -and $elem -ne [securestring]) {
                    $v = [System.Net.NetworkCredential]::new('', $v).Password   # string-typed secret: bind plaintext, masked everywhere
                    $secretText.Add($v)
                }
                $b[$k] = $v; continue
            }
            $raw = [string]$row.($m.Value)
            if ([string]::IsNullOrWhiteSpace($raw)) {
                if ($requiredNames -contains $k)                  { $problems.Add(("-{0} is blank" -f $k)) }
                elseif ($blankClears -and $elem -ne [bool])       { $b[$k] = $null }   # $null clears the value
                continue
            }
            $raw = $raw.Trim()
            if ($m.Secret) {
                $secretText.Add($raw)
                # The plaintext is already in the CSV file; a SecureString parameter needs it as one.
                if ($elem -eq [securestring]) {
                    $ss = [System.Security.SecureString]::new()
                    foreach ($ch in $raw.ToCharArray()) { $ss.AppendChar($ch) }
                    $ss.MakeReadOnly()
                    $b[$k] = $ss
                } else { $b[$k] = $raw }
                continue
            }
            $val = $raw
            if ($p) {
                $conv = ConvertTo-ParamValue -Param $p -Text $raw -Allowed $m.Allowed
                if (-not $conv.Ok) { $problems.Add(("-{0}: {1}" -f $k, $conv.Error)); continue }
                $val = $conv.Value
            }
            if ($lookupOk[$k]) {
                $match = $null
                try { $match = & $ParamLookup[$k].Resolve $val } catch { }
                if (-not $match) { $problems.Add(("-{0}: no mailbox/user found matching '{1}'" -f $k, $val)); continue }
            }
            $b[$k] = $val
        }
        if ($Transform -and $problems.Count -eq 0) {
            foreach ($tk in $Transform.Keys) {
                if (-not $b.Contains($tk)) { continue }
                try   { $b[$tk] = & $Transform[$tk] $b[$tk] }
                catch { $problems.Add(("-{0}: {1}" -f $tk, $_.Exception.Message)) }
            }
        }
        $target = (@($keyNames | Where-Object { $b.Contains($_) } | ForEach-Object { '{0}={1}' -f $_, (Format-SafetyValue $b[$_]) }) -join ', ')
        if (-not $target) { $target = 'row {0}' -f $rowNo }
        $plans.Add([pscustomobject]@{ Row = $rowNo; Source = $row; Splat = $b; Problems = @($problems); Secrets = @($secretText); Target = $target })
    }

    $bad  = @($plans | Where-Object { $_.Problems.Count -gt 0 })
    $todo = @($plans | Where-Object { $_.Problems.Count -eq 0 })
    if ($bad.Count -gt 0) {
        Write-Host ''
        Write-Host ('{0} of {1} rows have problems and will not run:' -f $bad.Count, $plans.Count) -ForegroundColor Yellow
        foreach ($pl in ($bad | Select-Object -First 20)) { Write-Host ('  row {0}: {1}' -f $pl.Row, ($pl.Problems -join '; ')) -ForegroundColor Yellow }
        if ($bad.Count -gt 20) { Write-Host ('  ... and {0} more.' -f ($bad.Count - 20)) -ForegroundColor Yellow }
        if ($todo.Count -eq 0) { Write-Host 'No row can run - fix the CSV and try again.' -ForegroundColor White; return }
        $c = Get-MenuChoice -Title 'What now?' -Options @('Cancel - fix the CSV first (nothing runs)', ('Run only the {0} good row(s)' -f $todo.Count))
        if ($c -notlike 'Run only*') { Write-Host 'Cancelled - nothing was run.' -ForegroundColor White; return }
    }

    # Preview: first and last command, then every target (the first 15, or all on request).
    $secretNames   = @($map.Keys | Where-Object { $map[$_].Secret })
    $secretColumns = @($map.Keys | Where-Object { $map[$_].Kind -eq 'column' -and $map[$_].Secret } | ForEach-Object { $map[$_].Value })
    Write-Host ''
    Write-Host ('Preview - {0} row(s) will run (dry-run):' -f $todo.Count) -ForegroundColor Yellow
    Write-CommandColored -Command (Format-CommandPreview -Cmdlet $Cmdlet -BoundParams $todo[0].Splat -Secret $secretNames)
    if ($todo.Count -gt 1) {
        Write-Host '  ...' -ForegroundColor White
        Write-CommandColored -Command (Format-CommandPreview -Cmdlet $Cmdlet -BoundParams $todo[-1].Splat -Secret $secretNames)
    }
    Write-Host ''
    Write-Host 'Targets:' -ForegroundColor White
    $showN = [Math]::Min($todo.Count, 15)
    foreach ($pl in $todo[0..($showN - 1)]) { Write-Host ('  row {0}: {1}' -f $pl.Row, $pl.Target) -ForegroundColor White }
    if ($todo.Count -gt $showN) {
        $more = Read-Host ("  ... and {0} more. Type 'all' to list every row, or Enter to continue" -f ($todo.Count - $showN))
        if ("$more".Trim().ToLower() -eq 'all') {
            foreach ($pl in $todo[$showN..($todo.Count - 1)]) { Write-Host ('  row {0}: {1}' -f $pl.Row, $pl.Target) -ForegroundColor White }
        }
    }

    if ($DryRunOnly) { Write-Host 'DryRunOnly is set - not executing.' -ForegroundColor White; return }

    # Last confirmation before running every row. The cmdlet's own per-row
    # "Are you sure?" prompt is skipped after it, so a declined prompt can't be
    # counted as OK and the typed gate below covers every row.
    $skipVendorConfirm = $cmd.Parameters.ContainsKey('Confirm')
    Write-Host ''
    Write-Host ('FINAL CONFIRMATION - about to run {0} for {1} row(s).' -f $Cmdlet, $todo.Count) -ForegroundColor Yellow
    if ($skipVendorConfirm) { Write-Host '  (the cmdlet''s own per-row confirmation is skipped - this one covers every row)' -ForegroundColor White }
    $boundUnion = [ordered]@{}   # every parameter any row binds, so forwarding or -Force in one row counts
    foreach ($pl in $todo) { foreach ($k in $pl.Splat.Keys) { if (-not $boundUnion.Contains($k) -or $null -eq $boundUnion[$k]) { $boundUnion[$k] = $pl.Splat[$k] } } }
    if (-not (Confirm-Execution -Cmdlet $Cmdlet -RowCount $todo.Count -Bound $boundUnion)) {
        Write-Host 'Cancelled - nothing was run.' -ForegroundColor White; return
    }

    $script:LastRevert = $null   # a new change makes any earlier undo stale
    $isSet = $Cmdlet -like 'Set-*'
    $newRec = {
        param($pl, $status, $message, $before)
        $rec = [ordered]@{ BulkRow = $pl.Row; BulkStatus = $status; BulkMessage = $message; BulkTarget = $pl.Target }
        if ($isSet) { $rec['BulkBefore'] = $before }
        foreach ($c in $columns) { if (-not $rec.Contains($c)) { $rec[$c] = if ($secretColumns -contains $c) { '********' } else { $pl.Source.$c } } }
        [pscustomobject]$rec
    }
    $results = [System.Collections.Generic.List[object]]::new()
    $ok = 0; $warn = 0; $fail = 0; $i = 0
    foreach ($pl in $todo) {
        $i++
        $splat = $pl.Splat
        # For a Set-*, record the values this row replaces (masked if credential-like).
        $beforeJson = ''
        if ($isSet) {
            try {
                $rp = Get-RevertPlan -Cmdlet $Cmdlet -Bound $splat
                if ($rp.Supported) {
                    $bf = [ordered]@{}
                    foreach ($bk in $rp.Before.Keys) { $bf[$bk] = if ($bk -match $script:SecretNamePattern) { '********' } else { Format-SafetyValue $rp.Before[$bk] } }
                    $beforeJson = $bf | ConvertTo-Json -Compress
                }
            } catch { }
        }
        $run = [ordered]@{}
        foreach ($k in $splat.Keys) { $run[$k] = $splat[$k] }
        if ($skipVendorConfirm) { $run['Confirm'] = $false }
        $wv = $null
        try {
            & $Cmdlet @run -ErrorAction Stop -WarningVariable wv -WarningAction SilentlyContinue | Out-Null
            if (@($wv).Count -gt 0) {
                # A warning usually means nothing changed (e.g. removing an entry that does not exist).
                $warn++; $status = 'WARNING'
                $msg = Hide-SecretText -Text ((@($wv) | ForEach-Object { "$_" }) -join ' | ') -Values $pl.Secrets
                Write-Host ('  [{0}/{1}] WARNING {2}: {3}' -f $i, $todo.Count, $pl.Target, $msg) -ForegroundColor Yellow
            } else {
                $ok++; $status = 'OK'; $msg = ''
                Write-Host ('  [{0}/{1}] OK      {2}' -f $i, $todo.Count, $pl.Target) -ForegroundColor Green
            }
        } catch {
            $fail++; $status = 'FAILED'
            $msg = Hide-SecretText -Text $_.Exception.Message -Values $pl.Secrets
            Write-Host ('  [{0}/{1}] FAILED  {2}: {3}' -f $i, $todo.Count, $pl.Target, $msg) -ForegroundColor Red
        }
        Write-ActionLog -Kind 'bulk-row' -Cmdlet $Cmdlet -Target $pl.Target -Params @($splat.Keys) -Outcome $status -Message $msg
        $results.Add((& $newRec $pl $status $msg $beforeJson))
    }
    foreach ($pl in $bad) { $results.Add((& $newRec $pl 'SKIPPED' ($pl.Problems -join '; ') '')) }

    Write-Host ''
    Write-Host ("Bulk done: {0} OK, {1} with warnings, {2} failed, {3} skipped, {4} row(s) in the file." -f $ok, $warn, $fail, $bad.Count, $rows.Count) -ForegroundColor White
    # Results file: every row and its outcome (and, for a Set-*, what it replaced),
    # so a partial or wrong run can be checked and retried. Owner-only, like logs.
    try {
        $expDir = Join-Path $script:ToolkitHome 'exports'
        if (-not (Test-Path -LiteralPath $expDir)) { New-Item -ItemType Directory -Path $expDir -Force | Out-Null }
        [void](Protect-LocalPath -Path $expDir -Directory)
        $resPath = Join-Path $expDir ('bulk-{0}-{1}-results.csv' -f $Cmdlet, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType File -Path $resPath -Force | Out-Null
        [void](Protect-LocalPath -Path $resPath)
        ConvertTo-SafeCsvRow -Items @($results | Sort-Object BulkRow) | Export-Csv -LiteralPath $resPath -NoTypeInformation -Encoding UTF8
        Write-Host ('Results saved to {0}' -f $resPath) -ForegroundColor White
        Write-ActionLog -Kind 'bulk' -Cmdlet $Cmdlet -Outcome 'DONE' -Message ('{0} OK, {1} warnings, {2} failed, {3} skipped; results {4}' -f $ok, $warn, $fail, $bad.Count, $resPath)
        if ($fail -gt 0 -or $bad.Count -gt 0) { Write-Host '  To retry, keep the FAILED / SKIPPED rows of that file and run it again as a CSV.' -ForegroundColor White }
    } catch {
        Write-Host ('Could not save the results file: {0}' -f $_.Exception.Message) -ForegroundColor Yellow
    }
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
        if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
        $pick = Read-Host ('Enter {0}, 0 to go back, or ? to explain' -f $hintNums)
        $k = ("$pick").Trim().ToLower()
        if ($k -eq '?') {
            if ($Description) { Write-Host ('  {0}' -f $Description) -ForegroundColor White }
            Write-Host ('  Options: ' + ($Options -join ', ')) -ForegroundColor White
            continue
        }
        if ($k -eq 'back' -or $k -eq '0') { $script:ParamBack   = $true; return $null }
        if ($k -eq 'cancel')              { $script:ParamCancel = $true; return $null }
        $nums = @($k -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })
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
        [hashtable] $Hints,     # per-action allowed-value override (name -> values); beats global $ParamHints
        [hashtable] $Candidates, # per-action dependent pick-lists (name -> { param($bound) ... }) derived from fields already chosen
        $Bound,                 # values collected so far (an [ordered] dict), passed to a Candidates scriptblock
        [switch] $Resolve       # verify a typed identity is a real recipient/user (catalogued actions only)
    )
    $script:ParamBack   = $false  # cleared each prompt; set if the user types 'back'   (go to previous field)
    $script:ParamCancel = $false  # cleared each prompt; set if the user types 'cancel' (abort action to menu)
    $script:ParamSecret = $false  # cleared each prompt; set true when the value read was a masked secret (mask it in preview/logs)
    $name    = $Param.Name
    $type    = $Param.ParameterType
    $isArray = $type.IsArray
    $elem    = if ($isArray) { $type.GetElementType() } else { $type }

    # Work out whether this field has a fixed set of allowed values, from (in order):
    # a [ValidateSet], the values of an enum type, or our manual $ParamHints table
    # (needed for Exchange proxy cmdlets that hide their real type as Object[]).
    $allowed = Get-AllowedValues -Param $Param -Hints $Hints

    # Say what the field means before asking for it.
    if ($Description) { Write-Host ("  -> {0}" -f $Description) -ForegroundColor White }

    # Secret fields: never echo to screen or transcript. Triggered by a
    # [SecureString] param, or a plain param whose name looks like a credential
    # (Password/Secret/Token/Key/...). Read masked via Read-Host -AsSecureString
    # (its input is NOT captured by Start-Transcript). Bind a SecureString param
    # as-is; a string-typed secret is converted to plaintext only at return, and
    # $script:ParamSecret tells the caller to mask it in the preview and logs.
    # Mainly a guard for the "Run any command" hatch (the curated catalogue has
    # no secret params). No back/cancel here - you can't type a keyword blind.
    # NB: no $isArray exclusion - a secret-NAMED array param (e.g. -Passwords) must
    # still be read hidden rather than echoed through the plain array prompt. Reading
    # a single masked value for such a param is acceptable (they are vanishingly rare)
    # and far safer than leaking it to screen/transcript/undo log.
    if (Test-StructuredSecretParam -Param $Param) {
        Write-Host ("  -{0} takes a structured {1} object that this builder can't enter safely - skipped." -f $name, $elem.Name) -ForegroundColor Yellow
        if ($Required) { Write-Host '  It is required, so this command can''t be built here.' -ForegroundColor Yellow; $script:ParamCancel = $true }
        return $null
    }
    $isSecret = Test-SecretParam -Param $Param
    if ($isSecret) {
        $credUser = $null
        if ($elem -eq [pscredential]) {
            # A credential is a user name plus a password: ask the name in clear, the password hidden.
            $credUser = "$(Read-Host ("User name for -{0}{1}" -f $name, $(if ($Required) { ' (required)' } else { ' - or press Enter to skip' })))".Trim()
            if (-not $credUser) {
                if ($Required) { $script:ParamCancel = $true; Write-Host '  A user name is required here - cancelled.' -ForegroundColor Yellow }
                return $null
            }
        }
        while ($true) {
            $sec = Read-Host ("Enter -{0} (hidden){1}" -f $name, $(if ($Required) { ' (required)' } else { ' - or press Enter to skip' })) -AsSecureString
            $len = 0; try { $len = $sec.Length } catch { }
            if ($len -eq 0) {
                if ($Required) {
                    if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
                    Write-Host '  This one is required.' -ForegroundColor Yellow; continue
                }
                return $null
            }
            $script:ParamSecret = $true
            if ($elem -eq [pscredential]) { return [pscredential]::new($credUser, $sec) }
            if ($elem -eq [securestring]) { return $sec }
            return [System.Net.NetworkCredential]::new('', $sec).Password   # string-typed secret: bind plaintext, masked in preview/logs
        }
    }

    # Fixed set of allowed values -> numbered picker (multi-select for arrays).
    # Both support 'back'/'cancel' via the shared helper.
    if ($allowed) {
        return Read-OptionMenu -Name $name -Options $allowed -Multi:$isArray -Description $Description
    }

    # Dependent pick-list: some fields are constrained by a field already chosen
    # (e.g. calendar -User must be someone who ALREADY has a permission entry on
    # the chosen calendar). The action supplies a scriptblock that, given the
    # values collected so far, returns the valid candidates (.Display / .Value).
    # We present them as a numbered list so the user can't fill in someone with no
    # entry and then hit the cmdlet's "no existing permission entry" error. An
    # 'enter a different value' escape falls through to the normal free-text prompt
    # (in case the live list missed a valid target). If the source errors, we warn
    # and fall through rather than block. Single-value fields only.
    if ($Candidates -and $Candidates.ContainsKey($name) -and -not $isArray) {
        $cands = $null
        try { $cands = @(& $Candidates[$name] $Bound) }
        catch {
            Write-Host ('  Could not build the list of eligible values ({0}).' -f $_.Exception.Message) -ForegroundColor Yellow
            Write-Host '  Falling back to typing the value.' -ForegroundColor Yellow
        }
        if ($null -ne $cands) {
            if ($cands.Count -eq 0) {
                Write-Host ''
                Write-Host '  Nobody (other than the built-in Default/Anonymous entries) currently has a' -ForegroundColor Yellow
                Write-Host '  permission entry on this calendar - so there is nothing to change or remove here.' -ForegroundColor Yellow
                $alt = Get-MenuChoice -Title '  What do you want to do?' `
                            -Options @('Cancel this action','Type an address manually anyway') -AllowBack
                if (-not $alt)                     { $script:ParamBack   = $true; return $null }  # 0/Back = re-pick the calendar (previous field)
                if ($alt -eq 'Cancel this action') { $script:ParamCancel = $true; return $null }
                # 'Type an address manually anyway' -> fall through to the normal prompt below.
            }
            else {
                $pick = Get-MenuChoice -Title ("  Who already has access? Pick the person for -{0}:" -f $name) `
                            -Options @($cands.Display) `
                            -ExtraOptions @('Type a different address','Cancel this action') -AllowBack
                if (-not $pick)                        { $script:ParamBack   = $true; return $null }  # 0/Back = previous field
                if ($pick -eq 'Cancel this action')    { $script:ParamCancel = $true; return $null }
                if ($pick -ne 'Type a different address') {
                    $sel = $cands | Where-Object { $_.Display -eq $pick } | Select-Object -First 1
                    Write-Host ('  -> using {0}' -f $sel.Value) -ForegroundColor White
                    return $sel.Value
                }
                # 'Type a different address' -> fall through to the normal prompt below.
            }
        }
    }

    # Otherwise a prompt line with a human-friendly hint.
    $hint = switch -Regex ($type.Name) {
        'Boolean'               { 'true / false'; break }
        'Int|Double|Long'       { 'a number'; break }
        'DateTime'              { 'a date, e.g. 2026-07-31'; break }
        'String\[\]|Object\[\]' { 'one or more values, comma-separated'; break }
        'String'                { 'text (e.g. an email address or name)'; break }
        default                 { $type.Name }
    }
    # Directory lookup: offer 'find' when a lookup is registered for this param
    # AND its probe cmdlet is loaded (i.e. we're connected to the right service).
    $lookup  = $ParamLookup[$name]
    $canFind = $lookup -and (Get-Command $lookup.Probe -ErrorAction SilentlyContinue)
    # Piping: offer 'last' when the previous action returned objects. Scalar
    # fields only in v1 (array fields keep their comma-separated free entry).
    $hasLast = (@($script:LastResult).Count -gt 0) -and (-not $isArray)

    $label = ('Enter -{0} [{1}]{2}' -f $name, $hint, $(if ($Required) { ' (required)' } else { '' }))
    if ($canFind) { $label += " - or type 'find' to search" }
    if ($hasLast) { $label += " - or 'last' to reuse a value from the last result" }
    $label += " - or 'back' (previous field) / 'cancel' (menu)"

    while ($true) {
        $val = Read-Host $label
        # Normalise for keyword detection (trim + lower-case) so 'Find', ' find ' etc. still work.
        $key = ("$val").Trim().ToLower()
        if ($key -eq '?') {
            if ($Description) { Write-Host ('  {0}' -f $Description) -ForegroundColor White }
            if     ($canFind) { Write-Host "  Enter an email address or name - or type 'find' to search the directory." -ForegroundColor White }
            else              { Write-Host '  Free text - type a value.' -ForegroundColor White }
            continue
        }
        # 'back' = re-do the previous field; 'cancel' = abort the whole action.
        # The caller decides what each means in context (see Invoke-GuidedCmdlet).
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
            # Prefer the column named like the field being filled: a -User field takes
            # the User column of a permission result, not the mailbox's Identity.
            $idCols = @($name) + @('PrimarySmtpAddress','UserPrincipalName','Mail','Email','Identity','Id','WindowsEmailAddress')
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
            if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
            Write-Host '  This one is required.' -ForegroundColor Yellow
            continue
        }
        if ([string]::IsNullOrWhiteSpace($val)) { return $null }
        # Identity-style fields: verify the typed value is a real recipient/user.
        # Catches nonsense/typos instead of silently accepting them as an address.
        if ($Resolve -and $canFind -and $lookup.Resolve -and -not $isArray) {
            $match = $null
            try { $match = & $lookup.Resolve $val } catch { }
            if (-not $match) {
                Write-Host ("  No mailbox/user found matching '{0}'. Check the spelling, or type 'find' to search." -f $val) -ForegroundColor Yellow
                continue
            }
        }
        # Boolean and list fields: the binder rejects a String bound to a [bool]
        # (e.g. -AutoMapping), and a list must not be split inside a name ("Sales
        # Team"), so both go through the converter shared with the bulk runner.
        if ($elem -eq [bool] -or $isArray) {
            $conv = ConvertTo-ParamValue -Param $Param -Text $val -Allowed $allowed
            if (-not $conv.Ok) { Write-Host ("  {0}." -f $conv.Error) -ForegroundColor Yellow; continue }
            return $conv.Value
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
    $nums = @($pick -split '[,\s]+' | Where-Object { $_ -match '^\d{1,5}$' })
    if ($nums.Count -eq 0) { return $cols }
    $sel = foreach ($n in $nums) { $idx=[int]$n-1; if ($idx -ge 0 -and $idx -lt $cols.Count) { $cols[$idx] } }
    @($sel)
}

# ---------------------------------------------------------------------------
# Save objects to a CSV file, defaulting to .\exports\<cmdlet>-<timestamp>.csv.
# ---------------------------------------------------------------------------
function Save-ResultCsv {
    param([Parameter(Mandatory)] $Items, [string] $Cmdlet)
    $expDir  = Join-Path $script:ToolkitHome 'exports'
    $default = Join-Path $expDir ("{0}-{1}.csv" -f $Cmdlet, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $path = Read-Host ("CSV path (Enter for {0})" -f $default)
    if ([string]::IsNullOrWhiteSpace($path)) { $path = $default } else { $path = $path.Trim('"',' ') }
    try {
        $full   = [System.IO.Path]::GetFullPath($path)
        $dir    = Split-Path -Parent $full
        $inside = $dir -eq [System.IO.Path]::GetFullPath($expDir)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ($inside) { [void](Protect-LocalPath -Path $dir -Directory) }
        New-Item -ItemType File -Path $full -Force | Out-Null   # create and lock down before writing directory data
        [void](Protect-LocalPath -Path $full)
        ConvertTo-SafeCsvRow -Items $Items | Export-Csv -LiteralPath $full -NoTypeInformation -Encoding UTF8
        Write-Host ("Saved {0} row(s) to {1}" -f @($Items).Count, $full) -ForegroundColor Green
        if (-not $inside) { Write-Host '  This file is outside the exports folder, so the 30-day clean-up will not delete it. It holds directory data - delete it when you are done.' -ForegroundColor Yellow }
    } catch {
        Write-Host ('Could not save the CSV: {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------------
# Post-run handling of whatever a cmdlet returned: show as a table (all or
# chosen columns) or export to CSV. Scalar/string output just prints; nothing
# returned says so. Loops so you can view AND save without re-running.
# ---------------------------------------------------------------------------
function Show-Result {
    param($Result, [string] $Cmdlet)

    $items = @(Hide-SecretResult -Items $Result)
    $script:LastResult = @()   # a new result replaces the old one for 'last', even when it is empty
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
            'Show as a table'                        { if (@($items).Count -eq 1) { $items | Format-List | Out-Host } else { $items | Format-Table -AutoSize -Wrap | Out-Host } }
            'Show a table with only columns I choose' { $c = Select-ResultColumns -Items $items; if ($c) { $items | Select-Object $c | Format-Table -AutoSize -Wrap | Out-Host } }
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
# Restrict a file or directory to the current user only, cross-platform: an NTFS
# ACL (no inheritance, owner FullControl) on Windows, chmod 600 (file) / 700
# (dir) elsewhere. Never throws - a perms hiccup must not stop the tool - but
# says so and returns $false, because these files hold directory data.
function Protect-LocalPath {
    param([Parameter(Mandatory)] [string] $Path, [switch] $Directory)
    $onWindows = $IsWindows -or ($env:OS -eq 'Windows_NT')   # $IsWindows is $null on Windows PowerShell 5.1
    try {
        if ($onWindows) {
            $acl = Get-Acl $Path
            $acl.SetAccessRuleProtection($true, $false)   # stop inheriting; drop inherited entries
            $me  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $inherit = if ($Directory) { 'ContainerInherit,ObjectInherit' } else { 'None' }
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($me, 'FullControl', $inherit, 'None', 'Allow')
            $acl.SetAccessRule($rule)
            Set-Acl -Path $Path -AclObject $acl
        } else {
            & chmod ($(if ($Directory) { '700' } else { '600' })) $Path 2>$null
            if ($LASTEXITCODE -ne 0) { throw "chmod exited with code $LASTEXITCODE" }
        }
        return $true
    } catch {
        Write-Host ("[!] Could not restrict {0} to your account ({1}). Other users of this computer may be able to read it." -f $Path, $_.Exception.Message) -ForegroundColor Yellow
        return $false
    }
}
function Get-ToolkitConfig {
    $path = Join-Path $script:ToolkitHome 'config.json'
    if (Test-Path $path) { try { return (Get-Content $path -Raw | ConvertFrom-Json) } catch { } }
    [pscustomobject]@{ Upn = $null; LastService = $null }
}
function Save-ToolkitConfig {
    param([string] $Upn, [string] $LastService)
    if (-not (Test-Path $script:ToolkitHome)) { New-Item -ItemType Directory -Path $script:ToolkitHome -Force | Out-Null }
    $cfgPath = Join-Path $script:ToolkitHome 'config.json'
    # Create the file empty and lock it down BEFORE writing content, so there is no
    # window where config.json exists with default (possibly group/world-readable)
    # perms. It only holds a UPN + last-used service, but this closes the TOCTOU gap.
    if (-not (Test-Path $cfgPath)) { New-Item -ItemType File -Path $cfgPath -Force | Out-Null; [void](Protect-LocalPath -Path $cfgPath) }
    [pscustomobject]@{ Upn = $Upn; LastService = $LastService } |
        ConvertTo-Json | Set-Content -Path $cfgPath -Encoding UTF8
    [void](Protect-LocalPath -Path $cfgPath)   # re-assert (also covers the already-existed case)
}

# ---------------------------------------------------------------------------
# "Who am I" preflight: after connecting, show the account the session is REALLY
# signed in as (not just the email typed at start), its tenant, and for Graph the
# granted scopes. A mismatch with the typed email stops until you confirm.
# Returns $false when you decline, so the caller can sign out again.
# ---------------------------------------------------------------------------
function Show-WhoAmI {
    param([Parameter(Mandatory)] $Service, [string] $Upn)
    $who = Get-SignedInAccount -Service $Service
    if (-not $who -or -not $who.Account) {
        Write-Host ("You typed {0}. {1} gives no way to read the signed-in account back - check the sign-in window used the right one." -f $Upn, $Service.Name) -ForegroundColor Yellow
        $script:CurrentAccount = '{0} (typed; not read back)' -f $Upn; $script:CurrentTenant = $null
        return $true
    }
    $script:CurrentAccount = $who.Account; $script:CurrentTenant = $who.Tenant
    Write-Host ("Signed in as: {0}" -f $who.Account) -ForegroundColor Green
    if ($who.Tenant) { Write-Host ("  Tenant: {0}" -f $who.Tenant) -ForegroundColor White }
    if ($who.Scopes) {
        Write-Host ("  Granted scopes: {0}" -f ($who.Scopes -join ', ')) -ForegroundColor White
        # A delegated token carries every scope ever consented to this app, not only
        # the ones requested, so "read-only" is a minimum. Say so when it matters.
        $extra = @($who.Scopes | Where-Object { $Service.Scopes -notcontains $_ -and $_ -notin @('openid','profile','offline_access','email') })
        $write = @($extra | Where-Object { $_ -match 'Write|Manage|FullControl|AccessAsUser|PrivilegedAccess' })
        if ($write.Count -gt 0) {
            Write-Host ("  [!] This token can also WRITE (consented earlier): {0}. 'Run any command' can use these." -f ($write -join ', ')) -ForegroundColor Yellow
        }
    }
    if ($Upn -and $who.Account -ne $Upn) {
        Write-Host ("[!] You typed {0}, but this session is signed in as {1}." -f $Upn, $who.Account) -ForegroundColor Red
        if ((Read-Host ('Continue as {0}? (y/n)' -f $who.Account)) -notmatch '^(y|yes)$') { return $false }
    }
    return $true
}

# ---------------------------------------------------------------------------
# After an action finishes, ask what to do next. Returns 'again' (back to the
# category menu for this service), 'switch' (pick another service), or 'quit'.
# ---------------------------------------------------------------------------
function Get-NextStep {
    # Loop (not recurse) so repeatedly choosing "Undo" - e.g. a scripted/piped
    # stdin - can't build unbounded call depth.
    while ($true) {
        $opts = @('Do another task (same service)','Switch service','Quit')
        # Only surface an undo when the last Set produced a runnable revert. It is
        # opt-in here rather than a prompt after every command.
        if ($script:LastRevert -and -not $DryRunOnly) {
            $lr = $script:LastRevert
            $opts = @(('Undo the last change ({0} on {1}, made {2}{3})' -f $lr.Cmdlet, $lr.Target, $lr.Time.ToString('HH:mm'), $(if ($lr.Partial) { ', partial' } else { '' }))) + $opts
        }
        $choice = Get-MenuChoice -Title 'What next?' -Options $opts
        if ($choice -like 'Undo the last change*') {
            $rv = $script:LastRevert
            Write-Host ''
            Write-Host 'Reverting with:' -ForegroundColor Yellow
            Write-CommandColored -Command $rv.Preview
            if ((Read-Host 'Run this undo now? (y/n)') -match '^(y|yes)$') {
                Write-Host 'Reverting...' -ForegroundColor White
                try {
                    $rvSplat = $rv.Splat
                    & $rv.Cmdlet @rvSplat -ErrorAction Stop | Out-Null
                    Write-Host 'Reverted.' -ForegroundColor Green
                    Write-ActionLog -Kind 'undo' -Cmdlet $rv.Cmdlet -Target $rv.Target -Params @($rvSplat.Keys) -Outcome 'OK'
                    $script:LastRevert = $null   # consumed - don't offer the same undo again
                } catch {
                    Write-Host ('Revert failed: {0}' -f $_.Exception.Message) -ForegroundColor Red
                    Write-ActionLog -Kind 'undo' -Cmdlet $rv.Cmdlet -Target $rv.Target -Outcome 'FAILED' -Message $_.Exception.Message
                }
            }
            continue   # re-show the menu after handling the undo
        }
        switch ($choice) {
            'Switch service' { return 'switch' }
            'Quit'           { return 'quit' }
            default          { return 'again' }
        }
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
    # Same guard as the run-any path: an unexpected error reports and returns to
    # the menu instead of ending the session (and losing the in-memory undo).
    try {
        if ($mode -like 'Bulk*') {
            Invoke-BulkCmdlet  -Cmdlet $Action.Cmdlet -Prompt $Action.Prompt -ParamHelp $Action.ParamHelp -Preset $Action.Preset -Hints $Action.Hints -Transform $Action.Transform -Resolve
        } else {
            Invoke-GuidedCmdlet -Cmdlet $Action.Cmdlet -Prompt $Action.Prompt -Help $Action.Help -ParamHelp $Action.ParamHelp -Preset $Action.Preset -Hints $Action.Hints -Transform $Action.Transform -Project $Action.Project -Candidates $Action.Candidates -Resolve
        }
    } catch [System.IO.EndOfStreamException] { throw }
    catch { Write-Host ("Could not run '{0}': {1}" -f $Action.Label, $_.Exception.Message) -ForegroundColor Red }
    Get-NextStep
}

# ===========================================================================
# Main loop
# ===========================================================================
function Start-M365Toolkit {
    # Session state for output piping (the 'last' keyword in Read-ParameterValue):
    # the most recent result set. Populated by
    # Show-Result; in-memory only (not persisted across sessions).
    $script:LastResult = @()
    # The undo offer lives for this run only (under 'iex' script scope is the
    # shell's, so a stale undo from an earlier run must not survive).
    $script:LastRevert = $null
    $script:CurrentService = $null


    $script:CurrentAccount = $null
    $script:CurrentTenant  = $null
    $logDir = Join-Path $script:ToolkitHome 'logs'
    if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
    # Logs and exports hold names, addresses and directory data: owner-only, and
    # nothing older than 30 days is kept (checked now and again on exit).
    [void](Protect-LocalPath -Path $logDir -Directory)
    Invoke-RetentionPurge
    Test-DataFolder
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:ActionLog = Join-Path $logDir "actions-$stamp.jsonl"
    New-Item -ItemType File -Path $script:ActionLog -Force | Out-Null
    [void](Protect-LocalPath -Path $script:ActionLog)
    # The full screen transcript only on request (-Transcript): it holds every
    # search hit and result viewed. Script-scoped so the connect step can pause it.
    $script:TranscriptOn = $false
    $script:LogFile = $null
    if ($Transcript) {
        $script:LogFile = Join-Path $logDir "session-$stamp.log"
        $ta = @{ Path = $script:LogFile; Append = $true }
        if ((Get-Command Start-Transcript).Parameters.ContainsKey('UseMinimalHeader')) { $ta.UseMinimalHeader = $true }
        try { Start-Transcript @ta | Out-Null; $script:TranscriptOn = $true }
        catch { Write-Host ('[!] Could not start the transcript: {0}' -f $_.Exception.Message) -ForegroundColor Yellow }
    }
    try {
        Write-Host '=====================================================' -ForegroundColor White
        Write-Host ('  Microsoft 365 Guided Command Builder  ({0})' -f $script:ToolkitVersion) -ForegroundColor White
        Write-Host '=====================================================' -ForegroundColor White
        Write-Host ('Action log: {0}{1}' -f $script:ActionLog, $(if ($script:TranscriptOn) { "  |  transcript: $($script:LogFile)" } else { '' })) -ForegroundColor White
        Write-ActionLog -Kind 'start' -Message ('{0}{1}{2}' -f $script:ToolkitVersion, $(if ($DryRunOnly) { ' DryRunOnly' } else { '' }), $(if ($script:TranscriptOn) { ' Transcript' } else { '' }))

        $cfg = Get-ToolkitConfig
        $upnPrompt = if ($cfg.Upn) { "Your email (UPN) [Enter = $($cfg.Upn)]" } else { 'Your email (UPN)' }
        $upn = $null
        while ([string]::IsNullOrWhiteSpace($upn)) {
            $upn = "$(Read-Host $upnPrompt)".Trim()
            if (-not $upn -and $cfg.Upn) { $upn = $cfg.Upn }
            if (-not $upn) {
                if (Test-InputEnded) { throw [System.IO.EndOfStreamException]::new('No more input (stdin closed).') }
                Write-Host '  Your email is needed to sign in.' -ForegroundColor Yellow
            }
        }

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
            $script:CurrentService = $svcName
            if (-not $connected[$svcName]) {
                # A failed sign-in (bad creds, cancelled login, no network) throws
                # under $ErrorActionPreference='Stop'. Without this catch it unwinds
                # past the whole main loop and kills the tool. Instead, report it and
                # drop back to the service menu so the user can retry or pick another.
                try {
                    Connect-M365Service -Service $service -Upn $upn
                    # Services that share one sign-in (Entra and Intune both use
                    # Connect-MgGraph) replace each other's session: only the latest counts.
                    foreach ($other in @($Services | Where-Object { $_.ConnectCmd -eq $service.ConnectCmd -and $_.Name -ne $svcName })) { $connected.Remove($other.Name) }
                    $connected[$svcName] = $true
                    if (-not (Show-WhoAmI -Service $service -Upn $upn)) {
                        Write-ActionLog -Kind 'sign-in' -Outcome 'DECLINED' -Message 'signed-in account differs from the typed email'
                        Disconnect-M365Service -Service $service
                        $connected.Remove($svcName)
                        Write-Host 'Back to the service menu.' -ForegroundColor White
                        continue
                    }
                    Write-ActionLog -Kind 'sign-in' -Outcome 'OK'
                } catch [System.IO.EndOfStreamException] { throw }
                catch {
                    Write-ActionLog -Kind 'sign-in' -Outcome 'FAILED' -Message $_.Exception.Message
                    Write-Host ("Could not connect to {0}: {1}" -f $svcName, $_.Exception.Message) -ForegroundColor Red
                    Write-Host 'Back to the service menu - try again, or pick another service / Quit.' -ForegroundColor White
                    continue
                }
            }
            Save-ToolkitConfig -Upn $upn -LastService $svcName
            if ($script:LastRevert -and $script:LastRevert.Service -ne $svcName) { $script:LastRevert = $null }   # an undo belongs to the service it was made in
            $script:CurrentService = $svcName

            # Inner loop: pick and run actions for this service.
            while ($true) {
                $svcActions = $Actions | Where-Object { $_.Service -eq $svcName }
                if (-not $svcActions) {
                    Write-Host 'No plain-language actions catalogued for this service yet.' -ForegroundColor Yellow
                    Write-Host '(You can still extend $Actions - or point the engine at any cmdlet.)' -ForegroundColor White
                }

                $categories = @($svcActions.Category | Select-Object -Unique | Sort-Object)   # A-Z; fixed items appended after
                $cat = Get-MenuChoice -Title "[$svcName] Choose a category" `
                            -Options (@($categories) + @('Run any command by name (advanced)')) `
                            -ExtraOptions @('Search actions by keyword','Switch service') -AllowBack
                if (-not $cat -or $cat -eq 'Switch service') { break }

                # Keyword search across this service's actions (name / cmdlet / category).
                if ($cat -eq 'Search actions by keyword') {
                    $kw = (Read-Host 'Keyword (matches action name, cmdlet, or category)').Trim()
                    if ($kw) {
                        $rx = [regex]::Escape($kw)
                        $hits = @($svcActions | Where-Object { $_.Label -match $rx -or $_.Cmdlet -match $rx -or $_.Category -match $rx } | Sort-Object Label)   # results A-Z
                        if ($hits.Count -eq 0) { Write-Host "No actions match '$kw'." -ForegroundColor Yellow }
                        else {
                            $shown = @($hits)   # the exact objects whose labels are displayed, same order
                            $lbl = Get-MenuChoice -Title "Actions matching '$kw'" `
                                        -Options @($shown.Label) `
                                        -Descriptions @($shown | ForEach-Object { [string]$_.Help }) -AllowBack
                            if ($lbl) {
                                # Map the chosen label back to its object by position, then
                                # dispatch by the object's stable Id - never re-match on Label.
                                $idx = [array]::IndexOf(@($shown.Label), $lbl)
                                $action = if ($idx -ge 0) { $shown[$idx] } else { $null }
                                if ($action) {
                                    $post = Invoke-Action -Action $action
                                    if ($post -eq 'quit')   { $quit = $true; break }
                                    if ($post -eq 'switch') { break }
                                }
                            }
                        }
                    }
                    continue
                }

                # Escape hatch: drive ANY cmdlet in this connected module by name.
                if ($cat -eq 'Run any command by name (advanced)') {
                    # Browse the connected service's commands with NO typing: pick a verb
                    # (Get / Set / New ...), then pick the command. "Type a name or keyword"
                    # stays available for people who know the name.
                    # Only this service's own commands (Get-ServiceCommands) - never local
                    # system cmdlets (Remove-Item, Invoke-Expression, ...), another
                    # service's commands, sign-in commands or raw request cmdlets.
                    $svcCmds = @(Get-ServiceCommands -Service $service)
                    if ($svcCmds.Count -eq 0) {
                        Write-Host ("No commands found for {0} - is the session still connected?" -f $svcName) -ForegroundColor Yellow
                        continue
                    }
                    $verbs = @($svcCmds | ForEach-Object { $_.Verb } | Where-Object { $_ } | Sort-Object -Unique)
                    $cmdName = $null
                    while (-not $cmdName) {
                        $verb = Get-MenuChoice -Title 'Pick a verb (what do you want to do?)' `
                                    -Options $verbs -ExtraOptions @('Type a name or keyword instead') -AllowBack
                        if (-not $verb) { break }   # 0 / Enter = Back -> leave the advanced path
                        if ($verb -eq 'Type a name or keyword instead') {
                            while ($true) {
                                $entry = (Read-Host 'Cmdlet name or keyword (Enter to go back)').Trim()
                                if (-not $entry) { break }
                                if ($entry -eq '?') { Write-Host 'Type a full cmdlet name to use it, or part of a name to pick from matches.' -ForegroundColor White; continue }
                                # Only accept an exact name that belongs to THIS service's
                                # command set - not any cmdlet loaded in the session - so the
                                # advanced hatch stays scoped to M365 and can't run arbitrary
                                # local cmdlets (Remove-Item, Invoke-Expression, ...).
                                if ($svcCmds.Name -contains $entry) { $cmdName = $entry; break }
                                $hits = @($svcCmds | Where-Object { $_.Name -like "*$entry*" } | Select-Object -ExpandProperty Name | Sort-Object)
                                if ($hits.Count -eq 0)  { Write-Host ("Nothing matches '{0}'." -f $entry) -ForegroundColor Yellow; continue }
                                if ($hits.Count -gt 40) { Write-Host ("'{0}' matches {1} commands - narrow it down." -f $entry, $hits.Count) -ForegroundColor Yellow; continue }
                                $p = Get-MenuChoice -Title ("Commands matching '{0}':" -f $entry) -Options $hits -AllowBack
                                if ($p) { $cmdName = $p; break }
                            }
                            continue   # got a name -> loop exits; otherwise back to the verb menu
                        }
                        # A verb was chosen: list its commands (nouns), filtering first if long.
                        $inVerb = @($svcCmds | Where-Object { $_.Verb -eq $verb } | Select-Object -ExpandProperty Name | Sort-Object)
                        if ($inVerb.Count -gt 40) {
                            $kw = (Read-Host ("  {0}- has {1} commands - type part of the name to filter, or Enter to list all" -f $verb, $inVerb.Count)).Trim()
                            if ($kw) {
                                $inVerb = @($inVerb | Where-Object { $_ -like "*$kw*" })
                                if ($inVerb.Count -eq 0) { Write-Host '  Nothing matched - pick another verb.' -ForegroundColor Yellow; continue }
                            }
                        }
                        $picked = Get-MenuChoice -Title ("{0} commands:" -f $verb) -Options $inVerb -AllowBack
                        if ($picked) { $cmdName = $picked }   # else 0/Enter = Back -> re-pick a verb
                    }
                    if ($cmdName) {
                        $mode = Get-MenuChoice -Title "How do you want to run '$cmdName'?" `
                                    -Options @('Run once (type the values now)','Bulk - run for every row in a CSV file') -AllowBack
                        try {
                            if     ($mode -like 'Bulk*') { Invoke-BulkCmdlet  -Cmdlet $cmdName }
                            elseif ($mode)               { Invoke-GuidedCmdlet -Cmdlet $cmdName }
                        } catch [System.IO.EndOfStreamException] { throw }
                        catch { Write-Host ("Could not run '{0}': {1}" -f $cmdName, $_.Exception.Message) -ForegroundColor Red }
                        $post = Get-NextStep
                        if ($post -eq 'quit')   { $quit = $true; break }
                        if ($post -eq 'switch') { break }
                    }
                    continue
                }

                $shown = @($svcActions | Where-Object { $_.Category -eq $cat } | Sort-Object Label)   # exact displayed objects, A-Z
                $label = Get-MenuChoice -Title "[$svcName / $cat] Choose an action" `
                            -Options @($shown.Label) `
                            -Descriptions @($shown | ForEach-Object { [string]$_.Help }) `
                            -AllowBack
                if (-not $label) { continue }

                # Map the chosen label back to its object by position, then dispatch
                # by the object's stable Id - never re-match on the Label string.
                $idx = [array]::IndexOf(@($shown.Label), $label)
                $action = if ($idx -ge 0) { $shown[$idx] } else { $null }
                if (-not $action) { continue }
                $post = Invoke-Action -Action $action
                if ($post -eq 'quit')   { $quit = $true; break }
                if ($post -eq 'switch') { break }
            }
            if ($quit) { break }
        }
        Write-Host 'Goodbye.' -ForegroundColor White
    }
    catch [System.IO.EndOfStreamException] {
        # Non-interactive run with no more piped input - exit cleanly instead of
        # spinning forever on an unanswerable prompt.
        Write-Host 'Input ended - exiting.' -ForegroundColor White
    }
    finally {
        # Sign out of every service this run connected to, so no admin session stays
        # live in the shell after Quit (under 'irm | iex' that shell is the caller's).
        if ($connected) {
            $signedOut = @()
            foreach ($n in @($connected.Keys)) {
                $svc = $Services | Where-Object { $_.Name -eq $n }
                $dc  = if ($svc.Disconnect) { $svc.Disconnect } else { 'Disconnect-ExchangeOnline' }
                if ($signedOut -contains $dc) { continue }
                $signedOut += $dc
                Disconnect-M365Service -Service $svc
            }
        }
        Write-ActionLog -Kind 'stop'
        if ($script:TranscriptOn) { try { Stop-Transcript | Out-Null } catch { } }
        Invoke-RetentionPurge -Quiet
        Set-Variable -Name ErrorActionPreference -Scope Script -Value $script:CallerErrorActionPreference
    }
}

# Dot-sourcing (". .\Invoke-M365Toolkit.ps1") loads the functions without starting
# the menu, so tests can call them. Every normal launch starts the tool.
if ($MyInvocation.InvocationName -ne '.') { Start-M365Toolkit }
