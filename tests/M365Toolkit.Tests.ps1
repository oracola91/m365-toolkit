<#
    Logic tests for Invoke-M365Toolkit.ps1 (Pester 5).

    The script is dot-sourced, which loads its functions without starting the
    menu. Read-Host is replaced with scripted answers, and the engine runs
    against stub cmdlets in an in-memory module, so no tenant, network or
    Microsoft module is needed.

    Run:  Invoke-Pester ./tests
#>

BeforeAll {
    . (Join-Path (Split-Path $PSScriptRoot -Parent) 'Invoke-M365Toolkit.ps1')
    Set-Variable -Name ToolkitHome -Scope Script -Value $TestDrive

    $global:Answers = [System.Collections.Generic.Queue[string]]::new()
    $global:Prompts = [System.Collections.Generic.List[string]]::new()
    $global:Calls   = [System.Collections.Generic.List[object]]::new()
    function Read-Host {
        param([Parameter(Position = 0)] $Prompt, [switch] $AsSecureString)
        $global:Prompts.Add("$Prompt")
        if ($global:Answers.Count -eq 0) { throw "No scripted answer for: $Prompt" }
        $a = $global:Answers.Dequeue()
        if ($AsSecureString) { $s = [System.Security.SecureString]::new(); foreach ($c in $a.ToCharArray()) { $s.AppendChar($c) }; return $s }
        $a
    }
    function Test-InputEnded { $false }
    # Run a block with scripted answers; returns everything it printed as one string.
    function Invoke-Scripted {
        param([string[]] $Answers, [scriptblock] $Block)
        $global:Answers.Clear(); foreach ($ans in $Answers) { $global:Answers.Enqueue($ans) }
        $global:Calls.Clear(); $global:Prompts.Clear()
        $script:Out = @(& $Block 6>&1 3>&1 | ForEach-Object { "$_" })
        ($script:Out -join '')
    }

    $global:Store = @{ m1 = [pscustomobject]@{ Identity = 'm1'; InternalMessage = "x' ; Write-Host PWNED ; '"; Office = 'OldOffice'; HiddenFromAddressListsEnabled = $true; GrantSendOnBehalfTo = @('old1') } }
    New-Module -Name ToolkitTestStubs -ScriptBlock {
        function Get-EXORecipient { param([string] $Identity, $ResultSize, $Filter)
            if ($Identity -in 'mbx1','mbx2','mbx3','mbx5','alice','bob','ghost','warnme','m1','m2') { [pscustomobject]@{ Identity = $Identity } } }
        function Remove-FakePermission { [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
            param([Parameter(Mandatory)][string] $Identity, [Parameter(Mandatory)][string] $User, [object[]] $AccessRights)
            if ($User -eq 'ghost')  { Write-Error "Recipient $User not found"; return }
            if ($User -eq 'warnme') { Write-Warning "ACE for $User does not exist"; return }
            if ($PSCmdlet.ShouldProcess($Identity)) { $global:Calls.Add([pscustomobject]@{ Identity = $Identity; User = $User }) } }
        function Get-FakeMailbox { [CmdletBinding()] param([string] $Identity) $global:Store['m1'] | Select-Object Identity, Office, HiddenFromAddressListsEnabled, GrantSendOnBehalfTo }
        function Set-FakeMailbox { [CmdletBinding()] param([Parameter(Mandatory)][string] $Identity, [string] $Office, [bool] $HiddenFromAddressListsEnabled, [object[]] $GrantSendOnBehalfTo)
            $global:Calls.Add([pscustomobject]@{ Identity = $Identity; Bound = ($PSBoundParameters.Keys -join ','); Hidden = $HiddenFromAddressListsEnabled
                HiddenType = $(if ($PSBoundParameters.ContainsKey('HiddenFromAddressListsEnabled')) { $HiddenFromAddressListsEnabled.GetType().Name }); Delegates = @($GrantSendOnBehalfTo) }) }
        function Get-FakeAutoReply { [CmdletBinding()] param([string] $Identity) $global:Store[$Identity] }
        function Set-FakeAutoReply { [CmdletBinding()] param([Parameter(Mandatory)][string] $Identity, [string] $InternalMessage, [string] $Office, [string] $NotAProp) $global:Calls.Add('set') }
        function Get-FakeTwo { [CmdletBinding()] param([string] $Identity) [pscustomobject]@{ Office = 'A' }; [pscustomobject]@{ Office = 'B' } }
        function Set-FakeTwo { [CmdletBinding()] param([Parameter(Mandatory)][string] $Identity, [string] $Office) $global:Calls.Add('two') }
        function New-FakeApp { [CmdletBinding()] param([Parameter(Mandatory)][string] $Name, [string] $ApiToken) Write-Error "token $ApiToken rejected for $Name" }
        function New-FakeCred { [CmdletBinding()] param([Parameter(Mandatory)][pscredential] $Credentials) $global:Calls.Add($Credentials.UserName + '|' + $Credentials.GetNetworkCredential().Password) }
        function Set-FakeThing { [CmdletBinding()] param([Parameter(Mandatory)][string] $Identity) $global:Calls.Add($Identity) }
        function Get-ConnectionInformation { [CmdletBinding()] param()
            [pscustomobject]@{ State = 'Connected'; ConnectionUri = 'https://outlook.office365.com'; ModuleName = 'ToolkitTestStubs'; UserPrincipalName = 'admin@contoso.com'; Organization = 'contoso.onmicrosoft.com' } }
        function Get-MgContext { [CmdletBinding()] param() [pscustomobject]@{ Account = 'other@contoso.com'; TenantId = 't-1'; Scopes = @('User.Read.All','User.ReadWrite.All','openid') } }
        function Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')] param() $global:Calls.Add('disconnect confirm=' + $PSBoundParameters['Confirm']) }
        function Get-MailboxFolderStatistics { [CmdletBinding()] param([string] $Identity, [string] $FolderScope)
            if ($Identity -eq 'de-user') { [pscustomobject]@{ FolderType = 'Calendar'; FolderPath = '/Kalender' } } else { throw 'not found' } }
        function Get-Mailbox { [CmdletBinding()] param($ResultSize, $RecipientTypeDetails) 's1','s2','s3' | ForEach-Object { [pscustomobject]@{ Identity = $_ } } }
        function Get-MailboxPermission { [CmdletBinding()] param($Identity, $User)
            if ($Identity -eq 's2') { throw 'throttled' }
            if ($Identity -eq 's3') { [pscustomobject]@{ AccessRights = @('FullAccess'); IsInherited = $false } } }
        Export-ModuleMember -Function *
    } | Import-Module -DisableNameChecking -Force
}

AfterAll {
    Remove-Module ToolkitTestStubs -ErrorAction SilentlyContinue
    Remove-Variable Answers, Prompts, Calls, Store -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Catalogue' {
    It 'every action has a unique Id' {
        @($Actions | Where-Object { [string]::IsNullOrWhiteSpace($_.Id) }).Count | Should -Be 0
        @($Actions | Group-Object Id | Where-Object Count -gt 1).Count | Should -Be 0
    }
    It 'the README "Add a new action" example includes an Id' {
        (Get-Content (Join-Path (Split-Path $PSScriptRoot -Parent) 'README.md') -Raw) | Should -Match "Id\s+=\s+'"
    }
}

Describe 'Shared value rules' {
    BeforeAll { $p = (Get-Command Set-FakeMailbox).Parameters }
    It 'list fields split on commas only ("Sales Team" stays one value)' {
        @((ConvertTo-ParamValue -Param $p.GrantSendOnBehalfTo -Text 'Sales Team').Value).Count | Should -Be 1
        @((ConvertTo-ParamValue -Param $p.GrantSendOnBehalfTo -Text 'a, b').Value).Count | Should -Be 2
    }
    It 'booleans map yes/no words and reject anything else' {
        (ConvertTo-ParamValue -Param $p.HiddenFromAddressListsEnabled -Text 'no').Value | Should -BeFalse
        (ConvertTo-ParamValue -Param $p.HiddenFromAddressListsEnabled -Text 'maybe').Ok | Should -BeFalse
    }
    It 'a CSV in Excel ANSI with ; separators is read correctly' {
        $f = Join-Path $TestDrive 'ansi.csv'
        [System.IO.File]::WriteAllBytes($f, [byte[]](0x4E,0x61,0x6D,0x65,0x3B,0x43,0x69,0x74,0x79,0x0A,0x4D,0xFC,0x6C,0x6C,0x65,0x72,0x3B,0x57,0x69,0x65,0x6E))
        $r = @(Import-BulkCsv -Path $f 6>$null)
        $r[0].Name | Should -Be ('M' + [char]0xFC + 'ller')
        $r[0].City | Should -Be 'Wien'
    }
}

Describe 'Run gate' {
    It 'read verbs and plain sets need no typed confirmation' {
        Get-DestructiveReason -Cmdlet 'Get-Mailbox' | Should -BeNullOrEmpty
        Get-DestructiveReason -Cmdlet 'Set-Mailbox' -Bound ([ordered]@{ Identity = 'x'; Office = 'A' }) | Should -BeNullOrEmpty
    }
    It 'destructive verbs, forwarding, raw writes, role changes and purge need one' {
        Get-DestructiveReason -Cmdlet 'Remove-MailboxPermission' | Should -Not -BeNullOrEmpty
        Get-DestructiveReason -Cmdlet 'Set-Mailbox' -Bound ([ordered]@{ ForwardingSmtpAddress = 'a@evil.test' }) | Should -Match 'forwards'
        Get-DestructiveReason -Cmdlet 'Invoke-MgGraphRequest' -Bound ([ordered]@{ Method = 'delete' }) | Should -Match 'DELETE'
        Get-DestructiveReason -Cmdlet 'Invoke-MgGraphRequest' -Bound ([ordered]@{ Method = 'GET' }) | Should -BeNullOrEmpty
        Get-DestructiveReason -Cmdlet 'Add-RoleGroupMember' | Should -Match 'admin roles'
        Get-DestructiveReason -Cmdlet 'New-ComplianceSearchAction' -Bound ([ordered]@{ Purge = [switch]$true }) | Should -Match 'Purge'
    }
    It 'a plain y does not pass a typed gate' {
        Invoke-Scripted @('y') { Confirm-Execution -Cmdlet 'Set-Mailbox' -Bound ([ordered]@{ ForwardingSmtpAddress = 'a@evil.test' }) } | Out-Null
        $script:Out[-1] | Should -Be 'False'
    }
}

Describe 'Bulk mode' {
    BeforeAll {
        "Mailbox,User,Rights`nmbx1,alice,FullAccess`nmbx2,ghost,FullAccess`nmbx3,warnme,fullaccess`nmbx4,,FullAccess`nmbx5,bob,Bogus" | Set-Content (Join-Path $TestDrive 'perm.csv')
        "Identity,Office,Hidden,Delegates`nm1,Sofia,true,Sales Team`nm2,,no,`"a, b`"" | Set-Content (Join-Path $TestDrive 'set.csv')
        "Name,Tok`napp1,SuperSecret123" | Set-Content (Join-Path $TestDrive 'sec.csv')
    }
    It 'checks every row first, gates on name and row count, tallies warnings and failures' {
        $o = Invoke-Scripted @((Join-Path $TestDrive 'perm.csv'),'Mailbox','User','1','Rights','2','Remove-FakePermission 3') { Invoke-BulkCmdlet -Cmdlet Remove-FakePermission -Prompt @('Identity','User','AccessRights') -Resolve }
        $o | Should -Match 'row 4:[^\n]*-User is blank'
        $o | Should -Match 'row 5: -AccessRights: not valid: Bogus'
        ($global:Prompts -join '|') | Should -Match 'type exactly: Remove-FakePermission 3'
        $global:Calls.Count | Should -Be 1
        $o | Should -Match 'Bulk done: 1 OK, 1 with warnings, 1 failed, 2 skipped'
        $res = Get-ChildItem (Join-Path $TestDrive 'exports') -Filter 'bulk-Remove-FakePermission-*-results.csv' | Select-Object -First 1
        (@(Import-Csv $res.FullName).BulkStatus -join ',') | Should -Be 'OK,FAILED,WARNING,SKIPPED,SKIPPED'
    }
    It 'cancels a destructive run when the typed text does not match' {
        $o = Invoke-Scripted @((Join-Path $TestDrive 'perm.csv'),'Mailbox','User','','2','yes') { Invoke-BulkCmdlet -Cmdlet Remove-FakePermission -Resolve }
        $o | Should -Match 'Cancelled - nothing was run'
        $global:Calls.Count | Should -Be 0
    }
    It 'binds real booleans and comma-only lists, skips blank cells, records replaced values' {
        Invoke-Scripted @((Join-Path $TestDrive 'set.csv'),'Identity','1,2,3','Office','Hidden','Delegates','1','y') { Invoke-BulkCmdlet -Cmdlet Set-FakeMailbox -Resolve } | Out-Null
        $c1 = $global:Calls | Where-Object Identity -eq 'm1'; $c2 = $global:Calls | Where-Object Identity -eq 'm2'
        $c1.HiddenType | Should -Be 'Boolean'
        @($c1.Delegates).Count | Should -Be 1
        @($c2.Delegates).Count | Should -Be 2
        $c2.Bound | Should -Not -Match 'Office'
        $res = Get-ChildItem (Join-Path $TestDrive 'exports') -Filter 'bulk-Set-FakeMailbox-*-results.csv' | Select-Object -First 1
        (Import-Csv $res.FullName)[0].BulkBefore | Should -Match 'OldOffice'
    }
    It 'masks secret columns in preview, errors and the results file' {
        $o = Invoke-Scripted @((Join-Path $TestDrive 'sec.csv'),'Name','1','Tok','y') { Invoke-BulkCmdlet -Cmdlet New-FakeApp }
        $o | Should -Not -Match 'SuperSecret123'
        $res = Get-ChildItem (Join-Path $TestDrive 'exports') -Filter 'bulk-New-FakeApp-*-results.csv' | Select-Object -First 1
        (Get-Content $res.FullName -Raw) | Should -Not -Match 'SuperSecret123'
    }
}

Describe 'Undo and command text' {
    It 'renders hostile text as one inert single-quoted literal' {
        $evil = "x' ; Write-Host PWNED ; '" + [char]0x2019 + "; `$(whoami)"
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('Set-X -A ' + (ConvertTo-PSLiteral $evil), [ref]$null, [ref]$null)
        @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)).Count | Should -Be 1
        ($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $args[0].StringConstantType -eq 'SingleQuoted' }, $true) | Select-Object -First 1).Value | Should -Be $evil
    }
    It 'writes an undo line that parses back to a single command' {
        Invoke-Scripted @('m1','1','new text','y') { Invoke-GuidedCmdlet -Cmdlet Set-FakeAutoReply -Prompt @('Identity','InternalMessage') } | Out-Null
        $line = (Get-Content (Join-Path $TestDrive 'logs/undo-hints.log') | Select-Object -Last 1) -split '  ', 2
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($line[1], [ref]$null, [ref]$null)
        @($ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)).Count | Should -Be 1
        $script:LastRevert.Target | Should -Be 'Identity=m1'
    }
    It 'flags a partial undo and offers none when the record is ambiguous' {
        $o = Invoke-Scripted @('m1','v1','v2','','y') { Invoke-GuidedCmdlet -Cmdlet Set-FakeAutoReply -Prompt @('Identity','NotAProp','Office') }
        $o | Should -Match 'Not covered by the undo'
        $o = Invoke-Scripted @('m1','1','Z','y') { Invoke-GuidedCmdlet -Cmdlet Set-FakeTwo }
        $o | Should -Match 'No automatic undo'
        $script:LastRevert | Should -BeNullOrEmpty
    }
    It 'never offers an undo under DryRunOnly' {
        $script:LastRevert = @{ Cmdlet = 'Set-Old'; Target = 'Identity=old'; Time = (Get-Date); Partial = $false; Preview = 'x'; Splat = @{} }
        $DryRunOnly = $true
        $o = Invoke-Scripted @('1') { Get-NextStep }
        $o | Should -Not -Match 'Undo the last change'
    }
}

Describe 'Secrets' {
    BeforeAll { $fake = { param($n, $t) [pscustomobject]@{ Name = $n; ParameterType = $t; Attributes = @() } } }
    It 'judges secrets by type, not just by name' {
        Test-SecretParam (& $fake 'ResetPasswordOnNextLogon' ([bool])) | Should -BeFalse
        Test-SecretParam (& $fake 'Password' ([string])) | Should -BeTrue
        Test-SecretParam (& $fake 'Credentials' ([pscredential])) | Should -BeTrue
        Test-SecretParam (& $fake 'Key' ([string])) | Should -BeFalse
        Test-StructuredSecretParam (& $fake 'PasswordProfile' ([hashtable])) | Should -BeTrue
    }
    It 'builds a PSCredential from a user name and a hidden password' {
        $o = Invoke-Scripted @('svc-user','S3cret!','y') { Invoke-GuidedCmdlet -Cmdlet New-FakeCred }
        $global:Calls[0] | Should -Be 'svc-user|S3cret!'
        $o | Should -Not -Match 'S3cret!'
    }
    It 'masks credential-like properties and bearer tokens in results' {
        $r = @(Hide-SecretResult -Items @([pscustomobject]@{ Name = 'app'; ClientSecret = 'abc' }, 'Bearer eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0In0.sig'))
        $r[0].ClientSecret | Should -Be '********'
        $r[1] | Should -Not -Match 'eyJhbGci'
    }
}

Describe 'Sessions and scope' {
    It 'shows the real account, flags extra write scopes, stops on a mismatch' {
        $svc = [pscustomobject]@{ Name = 'Entra'; Module = 'Microsoft.Graph'; ConnectCmd = 'Connect-MgGraph'; Scopes = @('User.Read.All') }
        $o = Invoke-Scripted @('n') { Show-WhoAmI -Service $svc -Upn 'admin@contoso.com' }
        $o | Should -Match 'Signed in as: other@contoso.com'
        $o | Should -Match 'can also WRITE'
        $script:Out[-1] | Should -Be 'False'
    }
    It 'offers only the service''s own commands in Run any command' {
        $exo = [pscustomobject]@{ Name = 'Exchange Online'; Module = 'ExchangeOnlineManagement'; ConnectCmd = 'Connect-ExchangeOnline' }
        $names = @(Get-ServiceCommands -Service $exo | ForEach-Object Name)
        $names | Should -Contain 'Get-FakeMailbox'
        $names | Should -Not -Contain 'Disconnect-ExchangeOnline'
    }
    It 'checks identities against the directory only for catalogued actions' {
        Invoke-Scripted @('not-a-recipient','y') { Invoke-GuidedCmdlet -Cmdlet Set-FakeThing } | Out-Null
        $global:Calls | Should -Contain 'not-a-recipient'
        $o = Invoke-Scripted @('not-a-recipient','cancel') { Invoke-GuidedCmdlet -Cmdlet Set-FakeThing -Resolve }
        $o | Should -Match 'No mailbox/user found'
    }
    It 'signs out without a prompt' {
        Invoke-Scripted @() { Disconnect-M365Service -Service ([pscustomobject]@{ Name = 'Purview'; Module = 'ExchangeOnlineManagement'; Disconnect = $null }) } | Out-Null
        $global:Calls | Should -Contain 'disconnect confirm=False'
    }
}

Describe 'Local data' {
    It 'writes action log lines with names of parameters only' {
        $script:ActionLog = Join-Path $TestDrive 'actions-test.jsonl'
        Write-ActionLog -Kind 'run' -Cmdlet 'Set-Mailbox' -Target 'Identity=m1' -Params @('Identity','Office') -Outcome 'OK'
        $j = Get-Content $script:ActionLog | Select-Object -Last 1 | ConvertFrom-Json
        $j.target | Should -Be 'Identity=m1'
        (@($j.params) -join ',') | Should -Be 'Identity,Office'
    }
    It 'deletes old logs and exports and old undo lines' {
        New-Item -ItemType Directory -Force (Join-Path $TestDrive 'logs'), (Join-Path $TestDrive 'exports') | Out-Null
        $old = Join-Path $TestDrive 'exports/old.csv'; 'x' | Set-Content $old; (Get-Item $old).LastWriteTime = (Get-Date).AddDays(-40)
        $undo = Join-Path $TestDrive 'logs/undo-hints.log'
        @(('{0}  Set-X -Identity ''a''' -f (Get-Date).AddDays(-45).ToString('yyyy-MM-dd HH:mm:ss')), ('{0}  Set-X -Identity ''b''' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))) | Set-Content $undo
        Invoke-RetentionPurge 6>$null
        Test-Path $old | Should -BeFalse
        @(Get-Content $undo).Count | Should -Be 1
    }
    It 'neutralises formula cells in exports' {
        $r = @(ConvertTo-SafeCsvRow -Items @([pscustomobject]@{ Name = '=HYPERLINK("http://x")'; Tags = @('a','b') }))
        $r[0].Name | Should -Be '''=HYPERLINK("http://x")'
        $r[0].Tags | Should -Be 'a; b'
    }
    It 'strips terminal control characters from tenant text' {
        ConvertTo-SafeText ('Bob' + [char]27 + '[2J') | Should -Be 'Bob[2J'
    }
    It 'uses the real (localized) calendar folder' {
        Get-CalendarFolderId -Mailbox 'de-user' | Should -Be 'de-user:\Kalender'
        Get-CalendarFolderId -Mailbox 'x' | Should -Be 'x:\Calendar'
    }
    It 'reports shared mailboxes the access scan could not read' {
        $o = Invoke-Scripted @() { Get-UserSharedMailboxAccess -User 'alice' }
        $o | Should -Match '1 of 3 shared mailboxes could not be read'
    }
}

Describe 'Added checks' {
    BeforeAll {
        New-Module -Name ToolkitCheckStubs -ScriptBlock {
            function Get-Mailbox { [CmdletBinding()] param($Identity, $ResultSize, $RecipientTypeDetails, $Filter)
                if ($Identity) { return [pscustomobject]@{ Identity = $Identity; ForwardingSmtpAddress = 'smtp:x@evil.test'; ForwardingAddress = $null; DeliverToMailboxAndForward = $false } }
                if ($Filter) { $global:LastFilter = $Filter; return [pscustomobject]@{ PrimarySmtpAddress = 'boss@contoso.com' } }
                [pscustomobject]@{ Identity = 'shared1'; PrimarySmtpAddress = 'shared1@contoso.com' } }
            function Get-InboxRule { [CmdletBinding()] param($Mailbox, [switch] $IncludeHidden)
                [pscustomobject]@{ Name = 'Forward all'; Enabled = $true; ForwardTo = @('x@evil.test') }
                [pscustomobject]@{ Name = 'Newsletters'; Enabled = $true; MoveToFolder = 'Newsletters' }
                [pscustomobject]@{ Name = 'Hide replies'; Enabled = $true; MoveToFolder = 'RSS Feeds'; MarkAsRead = $true } }
            function Get-Recipient { [CmdletBinding()] param($Identity, $ResultSize, $Filter)
                if ($Identity) { return [pscustomobject]@{ DistinguishedName = "CN=O'Brien,OU=contoso"; PrimarySmtpAddress = 'obrien@contoso.com' } }
                [pscustomobject]@{ PrimarySmtpAddress = 'sales@contoso.com'; RecipientTypeDetails = 'MailUniversalDistributionGroup' } }
            function Get-RecipientPermission { [CmdletBinding()] param($Trustee, $ResultSize) throw 'access denied' }
            function Get-MailboxPermission { [CmdletBinding()] param($Identity, $User)
                if ($User) { return [pscustomobject]@{ AccessRights = @('FullAccess'); IsInherited = $false } }
                [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess'); IsInherited = $false }
                [pscustomobject]@{ User = 'helpdesk@contoso.com'; AccessRights = @('FullAccess'); IsInherited = $false } }
            Export-ModuleMember -Function *
        } | Import-Module -DisableNameChecking -Force
    }
    AfterAll { Remove-Module ToolkitCheckStubs -ErrorAction SilentlyContinue; Remove-Variable LastFilter -Scope Global -ErrorAction SilentlyContinue }

    It 'forwarding check lists mailbox forwarding and suspicious rules, not ordinary ones' {
        $r = @(Get-MailboxForwardingReport -Identity 'm1')
        $r.Count | Should -Be 3
        $r[0].WhatItDoes | Should -Match 'keeps NO copy'
        ($r | Where-Object Name -eq 'Hide replies').WhatItDoes | Should -Match 'moves to RSS Feeds, marks as read'
        $r.Name | Should -Not -Contain 'Newsletters'
    }
    It 'leaver check covers every area, lists the unreadable one, and escapes the DN' {
        $o = @(Get-UserAccessReport -User 'obrien' -WarningVariable w -WarningAction SilentlyContinue)
        ($o | Where-Object Area -eq 'Send As').Object | Should -Be '(could not be read)'
        ($o | Where-Object Area -eq 'Send on behalf').Object | Should -Be 'boss@contoso.com'
        ($o | Where-Object Area -eq 'Mail groups').Object | Should -Be 'sales@contoso.com'
        ($o | Where-Object Area -eq 'Own mailbox: who can open it').Object | Should -Be 'helpdesk@contoso.com'
        @($o | Where-Object Area -eq 'Own mailbox: forwarding and rules').Count | Should -Be 3
        $global:LastFilter | Should -Match "O''Brien"
        "$w" | Should -Match 'Could not fully read: Send As'
    }
    It 'transitive group action shares the membership projection' {
        $t = $Actions | Where-Object Id -eq 'user-groups-transitive'
        $t.Cmdlet | Should -Be 'Get-MgUserTransitiveMemberOf'
        $row = & $t.Project @([pscustomobject]@{ AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @('Unified'); displayName = 'Team A'; mail = 'a@contoso.com' } })
        $row.Type | Should -Be 'Microsoft 365 group'
    }
    It 'read actions run once without the bulk question' {
        $act = $Actions | Where-Object Id -eq 'mbx-check-forwarding'
        $o = Invoke-Scripted @('mbx1','y','5','3') { Invoke-Action -Action $act }
        $o | Should -Not -Match 'How do you want to run'
        $o | Should -Match 'Returned 3 item'
    }
}
