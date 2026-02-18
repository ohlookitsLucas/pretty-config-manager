<#
.SYNOPSIS
    PCM Diagnostic Test Script - Signatures + Permissions
.DESCRIPTION
    Run this script in PowerShell 5+ (with Outlook open) to exercise all major
    code paths and produce a detailed diagnostic report.

    Usage:
        powershell -ExecutionPolicy Bypass -File "tests\Test-PcmFunctionality.ps1"

    The report is written to $env:TEMP\pcm-test-report-<timestamp>.txt and
    opened in Notepad automatically when done.

    IMPORTANT: This script will ADD and then immediately REMOVE a "Can view"
    permission on the Inbox folder of your first account, using your second
    account as the test user.  It cleans up after itself (removes what it added).
    Run only when Outlook is open.
#>

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ---- Locate modules ---------------------------------------------------------

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir     = Join-Path $scriptDir '..\src'
$permModule = Join-Path $srcDir 'PermissionsManager.psm1'
$sigModule  = Join-Path $srcDir 'SignatureManager.psm1'

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$report  = [System.Text.StringBuilder]::new()
$ts      = Get-Date -Format 'yyyyMMdd-HHmmss'
$outFile = Join-Path $env:TEMP "pcm-test-report-$ts.txt"

function Add-Result {
    param([string]$section, [string]$test, [string]$status, [string]$detail)
    $results.Add([PSCustomObject]@{ Section=$section; Test=$test; Status=$status; Detail=$detail })
    $colour = if ($status -eq 'PASS') { 'Green' } elseif ($status -eq 'INFO') { 'Cyan' } else { 'Red' }
    $line = "[$status] $section > $test"
    Write-Host $line -ForegroundColor $colour
    if ($detail) { Write-Host "        $detail" -ForegroundColor Gray }
    [void]$report.AppendLine($line)
    if ($detail) { [void]$report.AppendLine("        $detail") }
}

function Add-Section {
    param([string]$title)
    $bar = '=' * 60
    Write-Host ''
    Write-Host $bar -ForegroundColor Yellow
    Write-Host "  $title" -ForegroundColor Yellow
    Write-Host $bar -ForegroundColor Yellow
    [void]$report.AppendLine('')
    [void]$report.AppendLine($bar)
    [void]$report.AppendLine("  $title")
    [void]$report.AppendLine($bar)
}

function Join-WithSemicolon {
    param([string[]]$Items)
    return $Items -join '; '
}

# ---- SECTION 0: SETUP -------------------------------------------------------

Add-Section 'SETUP'

foreach ($mod in @($permModule, $sigModule)) {
    $leaf = Split-Path $mod -Leaf
    if (Test-Path $mod) {
        Add-Result 'Setup' "Module exists: $leaf" 'PASS' $mod
    } else {
        Add-Result 'Setup' "Module exists: $leaf" 'FAIL' "NOT FOUND at $mod"
    }
}

try {
    Import-Module $permModule -Force -ErrorAction Stop
    Add-Result 'Setup' 'Import PermissionsManager' 'PASS' ''
} catch {
    Add-Result 'Setup' 'Import PermissionsManager' 'FAIL' "$_"
}

try {
    Import-Module $sigModule -Force -ErrorAction Stop
    Add-Result 'Setup' 'Import SignatureManager' 'PASS' ''
} catch {
    Add-Result 'Setup' 'Import SignatureManager' 'FAIL' "$_"
}

$outlookProc = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
if ($outlookProc) {
    Add-Result 'Setup' 'Outlook process running' 'PASS' "PID: $($outlookProc.Id)"
} else {
    Add-Result 'Setup' 'Outlook process running' 'FAIL' 'Outlook.exe not found -- COM calls WILL fail. Please open Outlook and re-run.'
}

# ---- SECTION 1: SIGNATURE TAB -----------------------------------------------

Add-Section 'SECTION 1: SIGNATURE TAB'

$sigPath = Join-Path $env:APPDATA 'Microsoft\Signatures'
if (Test-Path $sigPath) {
    Add-Result 'Signatures' 'Signature path exists' 'PASS' $sigPath
} else {
    Add-Result 'Signatures' 'Signature path exists' 'FAIL' "Not found: $sigPath"
}

$sigs = $null
try {
    $sigs = Get-Signatures
    if ($sigs.Count -gt 0) {
        $sigList = $sigs -join ', '
        Add-Result 'Signatures' 'Get-Signatures' 'PASS' "$($sigs.Count) found: $sigList"
    } else {
        Add-Result 'Signatures' 'Get-Signatures' 'INFO' 'No .htm files found in Signatures folder'
    }
} catch {
    Add-Result 'Signatures' 'Get-Signatures' 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
}

if ($null -ne $sigs -and $sigs.Count -gt 0) {
    foreach ($sigName in $sigs) {
        try {
            $st = Get-SignatureStatus -Name $sigName
            $detail = "HasHtm=$($st.HasHtm) HasRtf=$($st.HasRtf) HasTxt=$($st.HasTxt) IsValid=$($st.IsValid)"
            if ($st.Warning) { $detail = "$detail WARNING=$($st.Warning)" }
            Add-Result 'Signatures' "Status of '$sigName'" 'PASS' $detail
        } catch {
            Add-Result 'Signatures' "Status of '$sigName'" 'FAIL' "$_"
        }
    }
}

$assignments = $null
try {
    $assignments = Get-SignatureAssignments
    if ($assignments.Count -gt 0) {
        Add-Result 'Signatures' 'Get-SignatureAssignments' 'PASS' "$($assignments.Count) account(s) returned"
        foreach ($a in $assignments) {
            $detail = "SMTP='$($a.SmtpAddress)' New='$($a.NewSignature)' Reply='$($a.ReplySignature)' RegPath='$($a.RegistryPath)'"
            Add-Result 'Signatures' "  Account: $($a.AccountName)" 'INFO' $detail
        }
    } else {
        Add-Result 'Signatures' 'Get-SignatureAssignments' 'FAIL' 'Returned 0 accounts -- check sigmanager.log for details'
    }
} catch {
    Add-Result 'Signatures' 'Get-SignatureAssignments' 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
}

$sigLog = Join-Path $env:APPDATA 'outlookmAnAger\sigmanager.log'
if (Test-Path $sigLog) {
    try {
        $logLines = Get-Content $sigLog -Tail 50 -ErrorAction Stop
        $detail = $logLines -join "`n        "
        Add-Result 'Signatures' 'sigmanager.log (last 50 lines)' 'INFO' $detail
    } catch {
        Add-Result 'Signatures' 'sigmanager.log read' 'FAIL' "$_"
    }
} else {
    Add-Result 'Signatures' 'sigmanager.log' 'INFO' "Not found at $sigLog (normal on first run)"
}

# ---- SECTION 2: PERMISSIONS - ACCOUNT ENUMERATION --------------------------

Add-Section 'SECTION 2: PERMISSIONS -- ACCOUNT ENUMERATION'

$accounts = $null
try {
    $accounts = Get-SignedInAccounts
    if ($accounts.Count -gt 0) {
        Add-Result 'Permissions' 'Get-SignedInAccounts' 'PASS' "$($accounts.Count) account(s)"
        foreach ($a in $accounts) {
            Add-Result 'Permissions' "  Account: $($a.Name)" 'INFO' "SMTP='$($a.SmtpAddress)'"
        }
    } else {
        Add-Result 'Permissions' 'Get-SignedInAccounts' 'FAIL' 'Returned 0 accounts -- Outlook COM may not be ready'
    }
} catch {
    Add-Result 'Permissions' 'Get-SignedInAccounts' 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
}

# ---- SECTION 3: PERMISSIONS - FOLDER ENUMERATION ---------------------------

Add-Section 'SECTION 3: PERMISSIONS -- FOLDER ENUMERATION'

$foldersByAccount = @{}
if ($null -ne $accounts -and $accounts.Count -gt 0) {
    foreach ($acct in $accounts) {
        try {
            $folders = Get-MailboxFolders -SmtpAddress $acct.SmtpAddress
            $foldersByAccount[$acct.SmtpAddress] = $folders
            if ($folders.Count -gt 0) {
                $topNames = @($folders | Where-Object { $_.Depth -le 1 } | ForEach-Object { $_.Name })
                $topList  = $topNames -join ', '
                Add-Result 'Permissions' "Get-MailboxFolders '$($acct.Name)'" 'PASS' "$($folders.Count) folders. Top: $topList"
            } else {
                Add-Result 'Permissions' "Get-MailboxFolders '$($acct.Name)'" 'FAIL' '0 folders returned'
            }
        } catch {
            Add-Result 'Permissions' "Get-MailboxFolders '$($acct.Name)'" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
        }
    }
} else {
    Add-Result 'Permissions' 'Get-MailboxFolders (skipped)' 'INFO' 'No accounts available'
}

# ---- SECTION 4: PERMISSIONS - READ PERMISSIONS (Inbox) ---------------------

Add-Section 'SECTION 4: PERMISSIONS -- READ PERMISSIONS (Inbox)'

$inboxByAccount = @{}
if ($null -ne $accounts -and $accounts.Count -gt 0) {
    foreach ($acct in $accounts) {
        $smtp    = $acct.SmtpAddress
        $folders = if ($foldersByAccount.ContainsKey($smtp)) { $foldersByAccount[$smtp] } else { @() }

        $inbox = @($folders | Where-Object { $_.Name -like 'Inbox' -or $_.Name -like 'Posteingang' } | Select-Object -First 1)
        if ($inbox.Count -eq 0) {
            $inbox = @($folders | Where-Object { $_.Depth -eq 1 } | Select-Object -First 1)
        }

        if ($inbox.Count -eq 0) {
            Add-Result 'Permissions' "Read Inbox perms '$($acct.Name)'" 'INFO' 'Inbox not found in folder list'
            continue
        }

        $inboxByAccount[$smtp] = $inbox[0]

        try {
            $perms = Get-FolderPermissions -EntryID $inbox[0].EntryID -StoreID $inbox[0].StoreID
            $permParts = @($perms | ForEach-Object { "$($_.User)=$($_.PermissionLevelName)" })
            $permSummary = $permParts -join '; '
            if ($perms.Count -gt 0) {
                Add-Result 'Permissions' "Read Inbox perms '$($acct.Name)'" 'PASS' $permSummary
            } else {
                Add-Result 'Permissions' "Read Inbox perms '$($acct.Name)'" 'INFO' '0 permissions (unusual - Default should exist)'
            }
        } catch {
            Add-Result 'Permissions' "Read Inbox perms '$($acct.Name)'" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
        }
    }
}

# ---- SECTION 5: WIZARD SIMULATION ------------------------------------------

Add-Section 'SECTION 5: WIZARD SIMULATION (add / verify / remove)'

$wizSource = if ($null -ne $accounts -and $accounts.Count -ge 1) { $accounts[0] } else { $null }
$wizTarget = if ($null -ne $accounts -and $accounts.Count -ge 2) { $accounts[1] } else { $null }

if ($null -eq $wizSource) {
    Add-Result 'Wizard' 'Setup' 'FAIL' 'No source account available -- skipping wizard simulation'
} elseif ($null -eq $wizTarget) {
    Add-Result 'Wizard' 'Setup' 'INFO' 'Only 1 account -- using own SMTP as target (self-permission round-trip)'
    $wizTarget = [PSCustomObject]@{ Name = $wizSource.Name; SmtpAddress = $wizSource.SmtpAddress }
} else {
    $info = "Source: '$($wizSource.Name)' ($($wizSource.SmtpAddress))  |  Target: '$($wizTarget.Name)' ($($wizTarget.SmtpAddress))"
    Add-Result 'Wizard' 'Setup' 'INFO' $info
}

$wizInbox = $null
if ($null -ne $wizSource) {
    if ($inboxByAccount.ContainsKey($wizSource.SmtpAddress)) {
        $wizInbox = $inboxByAccount[$wizSource.SmtpAddress]
    } else {
        try {
            $wf = @(Get-MailboxFolders -SmtpAddress $wizSource.SmtpAddress)
            $wi = @($wf | Where-Object { $_.Name -like 'Inbox' -or $_.Name -like 'Posteingang' } | Select-Object -First 1)
            if ($wi.Count -eq 0) { $wi = @($wf | Where-Object { $_.Depth -eq 1 } | Select-Object -First 1) }
            if ($wi.Count -gt 0) { $wizInbox = $wi[0] }
        } catch {}
    }
}

if ($null -ne $wizSource -and $null -ne $wizTarget -and $null -ne $wizInbox) {
    $targetUser = $wizTarget.SmtpAddress

    # 5a -- Add permission
    try {
        $addResult = Set-FolderPermissionWithAncestors -EntryID $wizInbox.EntryID -StoreID $wizInbox.StoreID -User $targetUser -Level 1
        $autoNames = @($addResult.AutoGranted | ForEach-Object { $_.FolderName })
        $autoMsg   = if ($autoNames.Count -gt 0) { " | AutoGranted: $($autoNames -join ', ')" } else { '' }
        Add-Result 'Wizard' "Add perm on '$($wizInbox.Name)' for '$targetUser'" 'PASS' "Success$autoMsg"
    } catch {
        Add-Result 'Wizard' "Add perm on '$($wizInbox.Name)'" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
    }

    # 5b -- Verify permission was written
    try {
        Start-Sleep -Milliseconds 500
        $verifyPerms = Get-FolderPermissions -EntryID $wizInbox.EntryID -StoreID $wizInbox.StoreID
        $smtpLocal   = ($targetUser -split '@')[0]
        $match = @($verifyPerms | Where-Object { $_.User -eq $targetUser -or $_.User -like "*$smtpLocal*" })
        if ($match.Count -gt 0) {
            $m = $match[0]
            Add-Result 'Wizard' "Verify perm present for '$targetUser'" 'PASS' "User='$($m.User)' Level='$($m.PermissionLevelName)' ($($m.PermissionLevel))"
        } else {
            $allPerms = @($verifyPerms | ForEach-Object { "$($_.User)=$($_.PermissionLevelName)" })
            Add-Result 'Wizard' "Verify perm present for '$targetUser'" 'FAIL' "User NOT found. Current: $($allPerms -join '; ')"
        }
    } catch {
        Add-Result 'Wizard' "Verify perm present" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
    }

    # 5c -- Remove permission
    try {
        Remove-FolderPermission -EntryID $wizInbox.EntryID -StoreID $wizInbox.StoreID -User $targetUser
        Add-Result 'Wizard' "Remove perm for '$targetUser'" 'PASS' ''
    } catch {
        Add-Result 'Wizard' "Remove perm for '$targetUser'" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
    }

    # 5d -- Verify removal
    try {
        Start-Sleep -Milliseconds 500
        $afterPerms  = Get-FolderPermissions -EntryID $wizInbox.EntryID -StoreID $wizInbox.StoreID
        $smtpLocal   = ($targetUser -split '@')[0]
        $stillThere  = @($afterPerms | Where-Object { $_.User -eq $targetUser -or $_.User -like "*$smtpLocal*" })
        if ($stillThere.Count -eq 0) {
            Add-Result 'Wizard' "Verify perm removed for '$targetUser'" 'PASS' 'User no longer in permissions'
        } else {
            $s = $stillThere[0]
            Add-Result 'Wizard' "Verify perm removed for '$targetUser'" 'FAIL' "STILL present: $($s.User)=$($s.PermissionLevelName)"
        }
    } catch {
        Add-Result 'Wizard' "Verify perm removed" 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
    }

} else {
    $reason = 'unknown'
    if ($null -eq $wizSource) { $reason = 'no source account' }
    elseif ($null -eq $wizTarget) { $reason = 'no target account' }
    elseif ($null -eq $wizInbox) { $reason = 'Inbox folder not found' }
    Add-Result 'Wizard' 'Wizard simulation (skipped)' 'INFO' "Skipped: $reason"
}

# ---- SECTION 6: PERMISSIONS OVERVIEW ----------------------------------------

Add-Section 'SECTION 6: PERMISSIONS OVERVIEW'

try {
    $overview = Get-PermissionsOverview
    Add-Result 'Overview' 'Get-PermissionsOverview' 'PASS' "$($overview.Count) mailbox(es)"
    foreach ($o in $overview) {
        if ($o.Entries.Count -eq 0) {
            Add-Result 'Overview' "  $($o.Mailbox) ($($o.SmtpAddress))" 'INFO' '(no custom permissions)'
        } else {
            foreach ($e in $o.Entries) {
                $folderParts = @($e.Folders | ForEach-Object { "$($_.FolderName)=$($_.Level)" })
                Add-Result 'Overview' "  $($o.Mailbox) > '$($e.User)'" 'INFO' ($folderParts -join ', ')
            }
        }
    }
} catch {
    Add-Result 'Overview' 'Get-PermissionsOverview' 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
}

# ---- SECTION 7: STORE / OST / PST DIAGNOSTICS ------------------------------

Add-Section 'SECTION 7: STORE / OST / PST DIAGNOSTICS'

try {
    $olApp = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $ns    = $olApp.GetNameSpace('MAPI')
    Add-Result 'Stores' 'Outlook COM connect' 'PASS' ''
    Add-Result 'Stores' 'Stores count' 'INFO' "$($ns.Stores.Count) store(s)"

    for ($i = 1; $i -le $ns.Stores.Count; $i++) {
        $st = $ns.Stores.Item($i)
        $dn = ''; try { $dn = $st.DisplayName } catch {}
        $fp = ''; try { $fp = $st.FilePath    } catch {}
        Add-Result 'Stores' "  Store[$i] '$dn'" 'INFO' "FilePath='$fp'"
    }

    Add-Result 'Stores' 'Accounts count' 'INFO' "$($ns.Accounts.Count) account(s)"
    for ($i = 1; $i -le $ns.Accounts.Count; $i++) {
        $acc  = $ns.Accounts.Item($i)
        $smtp = ''; try { $smtp = $acc.SmtpAddress } catch {}
        $ds   = ''; try { $ds   = $acc.DeliveryStore.DisplayName } catch { $ds = '(none)' }
        Add-Result 'Stores' "  Account[$i] '$($acc.DisplayName)'" 'INFO' "SMTP='$smtp' DeliveryStore='$ds'"
    }
} catch {
    Add-Result 'Stores' 'Outlook COM connect' 'FAIL' "$_ | Stack: $($_.ScriptStackTrace)"
}

# ---- SUMMARY ----------------------------------------------------------------

Add-Section 'SUMMARY'

$pass = @($results | Where-Object { $_.Status -eq 'PASS' }).Count
$fail = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
$info = @($results | Where-Object { $_.Status -eq 'INFO' }).Count

$summary     = "PASS: $pass   FAIL: $fail   INFO: $info"
$finalStatus = if ($fail -eq 0) { 'PASS' } else { 'FAIL' }
Add-Result 'Summary' $summary $finalStatus ''

if ($fail -gt 0) {
    [void]$report.AppendLine('')
    [void]$report.AppendLine('FAILED TESTS:')
    foreach ($r in ($results | Where-Object { $_.Status -eq 'FAIL' })) {
        [void]$report.AppendLine("  [$($r.Section)] $($r.Test)")
        if ($r.Detail) { [void]$report.AppendLine("    $($r.Detail)") }
    }
}

# ---- Write report -----------------------------------------------------------

$headerLines = @(
    'PCM Diagnostic Report',
    "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
    "User:      $env:USERNAME",
    "Machine:   $env:COMPUTERNAME",
    "ScriptDir: $scriptDir",
    ''
)
$header = $headerLines -join "`n"

Set-Content -Path $outFile -Value ($header + $report.ToString()) -Encoding UTF8

Write-Host ''
Write-Host "Report written to: $outFile" -ForegroundColor Green

$summaryColour = if ($fail -eq 0) { 'Green' } else { 'Red' }
Write-Host $summary -ForegroundColor $summaryColour

try { Start-Process notepad.exe -ArgumentList $outFile } catch {}
