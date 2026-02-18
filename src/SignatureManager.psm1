<#
  SignatureManager.psm1
  Full CRUD for Outlook signatures: list, validate, new, rename, delete, duplicate,
  export/import (zip), registry read/write for mailbox assignments, file-lock detection,
  structured logging, atomic file operations.

  Outlook stores signatures in: %APPDATA%\Microsoft\Signatures\
  Each signature is a set of files sharing the same base name:
    <Name>.htm   - HTML version (required)
    <Name>.rtf   - Rich Text version (optional)
    <Name>.txt   - Plain text version (optional)
    <Name>_files - Resource folder for embedded images (optional)

  Registry assignments (per Outlook profile/account):
    HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings
      NewSignature    = <name>
      ReplySignature  = <name>
  When multiple accounts are present, per-account keys may live under:
    HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\<profile>\9375CFF0413111d3B88A00104B2A6676\<acct-index>
      New Signature   = <name>   (REG_SZ)
      Reply Signature = <name>
#>

# Detect highest installed Office version dynamically; fall back to 16.0
$script:OutlookVer = '16.0'
$_officeBase = 'HKCU:\Software\Microsoft\Office'
if (Test-Path $_officeBase) {
    $_found = Get-ChildItem $_officeBase -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object { [double]$_.PSChildName } -Descending |
        Select-Object -First 1
    if ($_found) { $script:OutlookVer = $_found.PSChildName }
}
$script:SigPath      = Join-Path $env:APPDATA 'Microsoft\Signatures'
$script:RegBase      = "HKCU:\Software\Microsoft\Office\$($script:OutlookVer)\Common\MailSettings"
$script:ProfilesBase = "HKCU:\Software\Microsoft\Office\$($script:OutlookVer)\Outlook\Profiles"
$script:LogFile      = Join-Path $env:APPDATA 'outlookmAnAger\sigmanager.log'
$script:BackupDir    = Join-Path $env:APPDATA 'outlookmAnAger\registry-backups'

# ─── Logging ──────────────────────────────────────────────────────────────────

function Write-SigLog {
    param([string]$Level, [string]$Message)
    $logDir = Split-Path $script:LogFile -Parent
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

# ─── Path override (for testing) ──────────────────────────────────────────────

function Set-SignatureManagerPaths {
    <# Redirect module-scope paths so unit tests / sandbox can run against a temp directory. #>
    param(
        [string]$SignaturePath,
        [string]$RegistryBase,
        [string]$ProfilesBase,
        [string]$LogFile,
        [string]$BackupDir
    )
    if ($SignaturePath) { $script:SigPath      = $SignaturePath }
    if ($RegistryBase)  { $script:RegBase       = $RegistryBase }
    if ($ProfilesBase)  { $script:ProfilesBase  = $ProfilesBase }
    if ($LogFile)       { $script:LogFile        = $LogFile }
    if ($BackupDir)     { $script:BackupDir      = $BackupDir }
}

# ─── Path helpers ──────────────────────────────────────────────────────────────

function Get-SigFolder { return $script:SigPath }

function Get-SignatureFiles {
    param([Parameter(Mandatory)][string]$Name)
    return @{
        Htm    = Join-Path $script:SigPath "$Name.htm"
        Rtf    = Join-Path $script:SigPath "$Name.rtf"
        Txt    = Join-Path $script:SigPath "$Name.txt"
        Folder = Join-Path $script:SigPath "${Name}_files"
    }
}

function Get-SignatureHtmlPath {
    param([Parameter(Mandatory)][string]$Name)
    return Join-Path $script:SigPath "$Name.htm"
}

# ─── File-lock detection ───────────────────────────────────────────────────────

function Test-FileLocked {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        $stream.Dispose()
        return $false
    } catch {
        return $true
    }
}

function Assert-NotLocked {
    param([Parameter(Mandatory)][string]$Name)
    $files = Get-SignatureFiles -Name $Name
    foreach ($f in @($files.Htm, $files.Rtf, $files.Txt)) {
        if (Test-FileLocked -Path $f) {
            throw "File is locked (Outlook may have it open): $f"
        }
    }
}

# ─── Outlook running detection ─────────────────────────────────────────────────

function Test-OutlookRunning {
    return [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
}

# ─── Validation ────────────────────────────────────────────────────────────────

function Get-SignatureStatus {
    <#
    Returns a PSCustomObject per signature with validation state:
      Name, HasHtm, HasRtf, HasTxt, HasFolder, IsValid, Warning
    IsValid = HasHtm (minimum requirement)
    Warning = missing optional files
    #>
    param([string]$Name)
    $files = Get-SignatureFiles -Name $Name
    $hasHtm    = Test-Path $files.Htm
    $hasRtf    = Test-Path $files.Rtf
    $hasTxt    = Test-Path $files.Txt
    $hasFolder = Test-Path $files.Folder

    $warnings = @()
    if (-not $hasRtf) { $warnings += 'no .rtf' }
    if (-not $hasTxt) { $warnings += 'no .txt' }

    [PSCustomObject]@{
        Name      = $Name
        HasHtm    = $hasHtm
        HasRtf    = $hasRtf
        HasTxt    = $hasTxt
        HasFolder = $hasFolder
        IsValid   = $hasHtm
        Warning   = if ($warnings) { $warnings -join ', ' } else { $null }
    }
}

# ─── List ──────────────────────────────────────────────────────────────────────

function Get-Signatures {
    if (-not (Test-Path $script:SigPath)) { return @() }
    Get-ChildItem -Path $script:SigPath -Filter '*.htm' -File -ErrorAction SilentlyContinue |
        ForEach-Object { $_.BaseName } | Sort-Object -Unique
}

function Get-SignatureStatusList {
    Get-Signatures | ForEach-Object { Get-SignatureStatus -Name $_ }
}

# ─── Create (new blank signature) ─────────────────────────────────────────────

function New-Signature {
    param([Parameter(Mandatory)][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { throw "Name cannot be blank." }
    if ($Name -match '[\\/:*?"<>|]') { throw "Name contains invalid characters." }

    if (-not (Test-Path $script:SigPath)) {
        New-Item -ItemType Directory -Path $script:SigPath -Force | Out-Null
    }

    $files = Get-SignatureFiles -Name $Name
    if (Test-Path $files.Htm) { throw "Signature '$Name' already exists." }

    # Write minimal stub files
    $htmlStub = @"
<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="font-family:Calibri,sans-serif;font-size:11pt;">
<p>$Name</p>
</body>
</html>
"@
    Set-Content -Path $files.Htm -Value $htmlStub -Encoding UTF8
    Set-Content -Path $files.Txt -Value $Name -Encoding UTF8
    Write-SigLog 'INFO' "Created new signature: $Name"
}

# ─── Delete ────────────────────────────────────────────────────────────────────

function Remove-Signature {
    param([Parameter(Mandatory)][string]$Name)
    Assert-NotLocked -Name $Name
    $files = Get-SignatureFiles -Name $Name

    foreach ($f in @($files.Htm, $files.Rtf, $files.Txt)) {
        if (Test-Path $f) { Remove-Item -Path $f -Force }
    }
    if (Test-Path $files.Folder) {
        Remove-Item -Path $files.Folder -Recurse -Force
    }
    Write-SigLog 'INFO' "Deleted signature: $Name"
}

# ─── Rename (atomic) ──────────────────────────────────────────────────────────

function Rename-Signature {
    param(
        [Parameter(Mandatory)][string]$OldName,
        [Parameter(Mandatory)][string]$NewName
    )
    if ($OldName -eq $NewName) { return }
    if ([string]::IsNullOrWhiteSpace($NewName)) { throw "New name cannot be blank." }
    if ($NewName -match '[\\/:*?"<>|]') { throw "New name contains invalid characters." }

    $srcFiles = Get-SignatureFiles -Name $OldName
    $dstFiles = Get-SignatureFiles -Name $NewName

    if (Test-Path $dstFiles.Htm) { throw "A signature named '$NewName' already exists." }
    Assert-NotLocked -Name $OldName

    # Rename each file that exists
    $pairs = @(
        @{ Src = $srcFiles.Htm;    Dst = $dstFiles.Htm }
        @{ Src = $srcFiles.Rtf;    Dst = $dstFiles.Rtf }
        @{ Src = $srcFiles.Txt;    Dst = $dstFiles.Txt }
        @{ Src = $srcFiles.Folder; Dst = $dstFiles.Folder }
    )
    foreach ($p in $pairs) {
        if (Test-Path $p.Src) { Rename-Item -Path $p.Src -NewName (Split-Path $p.Dst -Leaf) -Force }
    }

    # Update registry references
    Update-RegistrySignatureName -OldName $OldName -NewName $NewName

    Write-SigLog 'INFO' "Renamed signature: '$OldName' -> '$NewName'"
}

# ─── Duplicate ─────────────────────────────────────────────────────────────────

function Copy-Signature {
    param(
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string]$TargetName
    )
    if ([string]::IsNullOrWhiteSpace($TargetName)) { throw "Target name cannot be blank." }
    if ($TargetName -match '[\\/:*?"<>|]') { throw "Target name contains invalid characters." }

    $src = Get-SignatureFiles -Name $SourceName
    $dst = Get-SignatureFiles -Name $TargetName
    if (-not (Test-Path $src.Htm)) { throw "Source signature '$SourceName' not found." }
    if (Test-Path $dst.Htm) { throw "Target signature '$TargetName' already exists." }

    Assert-NotLocked -Name $SourceName

    foreach ($pair in @(
        @{ S = $src.Htm; D = $dst.Htm }
        @{ S = $src.Rtf; D = $dst.Rtf }
        @{ S = $src.Txt; D = $dst.Txt }
    )) {
        if (Test-Path $pair.S) { Copy-Item -Path $pair.S -Destination $pair.D -Force }
    }
    if (Test-Path $src.Folder) {
        if (Test-Path $dst.Folder) { Remove-Item -Path $dst.Folder -Recurse -Force }
        Copy-Item -Path $src.Folder -Destination $dst.Folder -Recurse -Force
    }
    Write-SigLog 'INFO' "Duplicated signature: '$SourceName' -> '$TargetName'"
}

# ─── Save HTML (WYSIWYG edit result) ──────────────────────────────────────────

function Save-SignatureHtml {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$HtmlContent
    )
    Assert-NotLocked -Name $Name
    $files = Get-SignatureFiles -Name $Name

    # Atomic write via temp file
    $tmp = $files.Htm + '.tmp'
    try {
        Set-Content -Path $tmp -Value $HtmlContent -Encoding UTF8
        if (Test-Path $files.Htm) { Remove-Item $files.Htm -Force }
        Rename-Item -Path $tmp -NewName (Split-Path $files.Htm -Leaf) -Force

        # Regenerate .txt from HTML (strip tags)
        $plain = $HtmlContent -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
        $plain = [System.Text.RegularExpressions.Regex]::Replace($plain, '\s+', ' ').Trim()
        Set-Content -Path $files.Txt -Value $plain -Encoding UTF8

        # Leave .rtf untouched (per design decision)
        Write-SigLog 'INFO' "Saved HTML for signature: $Name"
    } catch {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
}

# ─── Registry: read assignments ────────────────────────────────────────────────

function Get-SignatureAssignments {
    <#
    Returns a list of PSCustomObjects:
      @{ AccountName; SmtpAddress; NewSignature; ReplySignature; RegistryPath }
    Per-account entries are built first. The global (Default) entry is only
    prepended when at least one per-account entry exists, so machines with no
    Outlook accounts configured return an empty list.
    #>
    $perAccount = @()

    # Per-account profile keys — include ALL accounts, not just those with sigs assigned
    $profilesBase = $script:ProfilesBase
    if (Test-Path $profilesBase) {
        Get-ChildItem -Path $profilesBase -ErrorAction SilentlyContinue | ForEach-Object {
            $profilePath = $_.PSPath
            $acctGuid = '9375CFF0413111d3B88A00104B2A6676'
            $acctBase = Join-Path $profilePath $acctGuid
            if (Test-Path $acctBase) {
                Get-ChildItem -Path $acctBase -ErrorAction SilentlyContinue | ForEach-Object {
                    $acctPath  = $_.PSPath
                    $acctProps = Get-ItemProperty -Path $acctPath -ErrorAction SilentlyContinue
                    if ($null -eq $acctProps) { return }
                    # Identify real account entries by the presence of an Account Name value
                    $rawName = $acctProps.'Account Name'
                    if ([string]::IsNullOrEmpty($rawName)) { return }
                    # Extract SMTP from "Display Name <smtp@addr>" or direct value
                    $smtp = $rawName -replace '^.*<(.+)>.*$','$1'
                    if ($smtp -eq $rawName) { $smtp = '' }   # no angle-bracket pattern found
                    if ([string]::IsNullOrEmpty($smtp) -and $acctProps.PSObject.Properties['Email']) {
                        $smtp = $acctProps.Email
                    }
                    # Skip non-email registry entries (transport/service subkeys with no valid SMTP)
                    if ([string]::IsNullOrEmpty($smtp)) { return }
                    $displayName = if ($acctProps.'Display Name') { $acctProps.'Display Name' } else { $smtp }
                    $perAccount += [PSCustomObject]@{
                        AccountName    = $displayName
                        SmtpAddress    = $smtp
                        NewSignature   = if ($acctProps.'New Signature')   { $acctProps.'New Signature' }   else { '' }
                        ReplySignature = if ($acctProps.'Reply Signature') { $acctProps.'Reply Signature' } else { '' }
                        RegistryPath   = $acctPath
                    }
                }
            }
        }
    }

    Write-SigLog 'INFO' "Get-SignatureAssignments: registry scan found $($perAccount.Count) account(s)"

    # Always merge COM accounts — catches delegated/shared mailboxes (OG- accounts added
    # via Exchange delegation) that have no independent Outlook profile registry entries.
    # Deduplicates by SMTP so registry-sourced entries are not duplicated.
    try {
        $ol = New-Object -ComObject Outlook.Application -ErrorAction Stop
        $ns = $ol.GetNameSpace('MAPI')
        $existingSmtps = @($perAccount | ForEach-Object { $_.SmtpAddress })
        Write-SigLog 'INFO' "Get-SignatureAssignments: COM merge — Outlook.Accounts.Count=$($ns.Accounts.Count)"
        for ($i = 1; $i -le $ns.Accounts.Count; $i++) {
            $acc  = $ns.Accounts.Item($i)
            $smtp = ''; try { $smtp = $acc.SmtpAddress } catch {}
            if ([string]::IsNullOrEmpty($smtp)) {
                Write-SigLog 'INFO' "  COM[$i] '$($acc.DisplayName)' — SMTP empty, skipped"
                continue
            }
            if ($existingSmtps -contains $smtp) {
                Write-SigLog 'INFO' "  COM[$i] '$($acc.DisplayName)' <$smtp> — already in registry list, skipped"
                continue
            }
            Write-SigLog 'INFO' "  COM[$i] '$($acc.DisplayName)' <$smtp> — added (not in registry)"
            $perAccount += [PSCustomObject]@{
                AccountName    = $acc.DisplayName
                SmtpAddress    = $smtp
                NewSignature   = ''
                ReplySignature = ''
                RegistryPath   = $script:RegBase
            }
        }
    } catch {
        Write-SigLog 'WARN' "Get-SignatureAssignments: COM merge failed — $_"
    }

    # Only expose the global (Default) entry when real accounts exist
    if ($perAccount.Count -eq 0) { return @() }

    $results = @()
    $globalKey = $script:RegBase
    if (Test-Path $globalKey) {
        $vals = Get-ItemProperty -Path $globalKey -ErrorAction SilentlyContinue
        $results += [PSCustomObject]@{
            AccountName    = 'All accounts (global default)'
            SmtpAddress    = ''
            NewSignature   = if ($vals.NewSignature)   { $vals.NewSignature }   else { '' }
            ReplySignature = if ($vals.ReplySignature) { $vals.ReplySignature } else { '' }
            RegistryPath   = $globalKey
        }
    }
    $results += $perAccount
    return $results
}

# ─── Registry: write assignments ──────────────────────────────────────────────

function Set-SignatureAssignment {
    param(
        [Parameter(Mandatory)][string]$RegistryPath,
        [string]$NewSignature   = $null,
        [string]$ReplySignature = $null
    )

    # Backup the key before modifying
    $backupDir = $script:BackupDir
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
    $backupFile = Join-Path $backupDir "backup-$(Get-Date -Format 'yyyyMMdd-HHmmss').reg"
    try {
        # Export via reg.exe for a clean backup
        $regPath = $RegistryPath -replace 'HKCU:\\','HKCU\'
        reg export $regPath $backupFile /y 2>$null | Out-Null
        Write-SigLog 'INFO' "Registry backup: $backupFile"
    } catch {
        Write-SigLog 'WARN' "Could not backup registry key: $_"
    }

    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    if ($null -ne $NewSignature) {
        # Detect key name style (global uses 'NewSignature', per-account uses 'New Signature')
        $existingProps = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue
        $keyNew = if ($existingProps.PSObject.Properties['New Signature']) { 'New Signature' } else { 'NewSignature' }
        Set-ItemProperty -Path $RegistryPath -Name $keyNew -Value $NewSignature -Type String
        Write-SigLog 'INFO' "Set $keyNew='$NewSignature' at $RegistryPath"
    }
    if ($null -ne $ReplySignature) {
        $existingProps = Get-ItemProperty -Path $RegistryPath -ErrorAction SilentlyContinue
        $keyReply = if ($existingProps.PSObject.Properties['Reply Signature']) { 'Reply Signature' } else { 'ReplySignature' }
        Set-ItemProperty -Path $RegistryPath -Name $keyReply -Value $ReplySignature -Type String
        Write-SigLog 'INFO' "Set $keyReply='$ReplySignature' at $RegistryPath"
    }
}

# ─── Registry: rename references ──────────────────────────────────────────────

function Update-RegistrySignatureName {
    param([string]$OldName, [string]$NewName)
    $assignments = Get-SignatureAssignments
    foreach ($a in $assignments) {
        $changed = $false
        $newNew   = $a.NewSignature
        $newReply = $a.ReplySignature
        if ($a.NewSignature   -eq $OldName) { $newNew   = $NewName; $changed = $true }
        if ($a.ReplySignature -eq $OldName) { $newReply = $NewName; $changed = $true }
        if ($changed) {
            Set-SignatureAssignment -RegistryPath $a.RegistryPath -NewSignature $newNew -ReplySignature $newReply
            Write-SigLog 'INFO' "Updated registry reference '$OldName'->'$NewName' at $($a.RegistryPath)"
        }
    }
}

# ─── Export / Import ──────────────────────────────────────────────────────────

function Export-Signature {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Destination
    )
    if (-not (Test-Path $script:SigPath)) { throw "Signatures folder not found." }
    Assert-NotLocked -Name $Name

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $temp | Out-Null
    try {
        $files = Get-SignatureFiles -Name $Name
        foreach ($f in @($files.Htm, $files.Rtf, $files.Txt)) {
            if (Test-Path $f) { Copy-Item -Path $f -Destination $temp -Force }
        }
        if (Test-Path $files.Folder) {
            Copy-Item -Path $files.Folder -Destination (Join-Path $temp "${Name}_files") -Recurse -Force
        }
        if (Test-Path $Destination) { Remove-Item -Path $Destination -Force }
        Compress-Archive -Path (Join-Path $temp '*') -DestinationPath $Destination -Force
        Write-SigLog 'INFO' "Exported '$Name' to $Destination"
    } finally {
        Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Import-Signature {
    param([Parameter(Mandatory)][string]$ZipPath)
    if (-not (Test-Path $ZipPath)) { throw "Zip not found: $ZipPath" }
    if (-not (Test-Path $script:SigPath)) { New-Item -ItemType Directory -Path $script:SigPath -Force | Out-Null }

    $temp = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $temp | Out-Null
    try {
        Expand-Archive -Path $ZipPath -DestinationPath $temp -Force
        Get-ChildItem -Path $temp -Recurse | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
            $rel  = $_.FullName.Substring($temp.Length).TrimStart('\')
            $dest = Join-Path $script:SigPath $rel
            $dir  = Split-Path $dest -Parent
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            Copy-Item -Path $_.FullName -Destination $dest -Force
        }
        Write-SigLog 'INFO' "Imported signature from $ZipPath"
    } finally {
        Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Export-ModuleMember -Function `
    Get-SigFolder, Get-SignatureFiles, Get-SignatureHtmlPath, `
    Get-Signatures, Get-SignatureStatusList, Get-SignatureStatus, `
    New-Signature, Remove-Signature, Rename-Signature, Copy-Signature, `
    Save-SignatureHtml, `
    Get-SignatureAssignments, Set-SignatureAssignment, `
    Export-Signature, Import-Signature, `
    Test-OutlookRunning, Test-FileLocked, `
    Set-SignatureManagerPaths
