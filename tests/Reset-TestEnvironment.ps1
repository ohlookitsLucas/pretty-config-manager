<#
  Reset-TestEnvironment.ps1
  Creates (or resets) a sandbox environment for offline testing.

  Usage:
    # Create/reset the sandbox and get all redirected paths
    $sandbox = .\tests\Reset-TestEnvironment.ps1

    # Tear down everything
    .\tests\Reset-TestEnvironment.ps1 -Teardown

  Returns a PSCustomObject with:
    .SignaturePath  - path to sandbox signatures folder
    .RegBase        - registry path for global MailSettings
    .ProfilesBase   - registry path for Outlook profiles
    .LogFile        - path to sandbox log file
    .BackupDir      - path to sandbox registry-backups folder
    .SettingsFile   - path to sandbox settings.json
    .SandboxRoot    - root of the entire sandbox
    .RegistryRoot   - root of all test registry keys
#>
param(
    [switch]$Teardown,
    [string]$SandboxRoot = (Join-Path $env:TEMP 'outlookmanager-sandbox')
)

$testRegRoot = 'HKCU:\Software\outlookmAnAger-TEST'

if ($Teardown) {
    # Clean up filesystem
    if (Test-Path $SandboxRoot) {
        Remove-Item -Path $SandboxRoot -Recurse -Force
        Write-Host "Removed sandbox folder: $SandboxRoot"
    }
    # Clean up registry
    if (Test-Path $testRegRoot) {
        Remove-Item -Path $testRegRoot -Recurse -Force
        Write-Host "Removed test registry keys: $testRegRoot"
    }
    return
}

# ── Resolve paths ──
$fixturesDir  = Join-Path $PSScriptRoot 'fixtures'
$sigFixtures  = Join-Path $fixturesDir 'signatures'
$settingsFixture = Join-Path $fixturesDir 'settings.json'

$sandboxSigs    = Join-Path $SandboxRoot 'Signatures'
$sandboxAppData = Join-Path $SandboxRoot 'AppData'
$sandboxLog     = Join-Path $sandboxAppData 'sigmanager.log'
$sandboxBackups = Join-Path $sandboxAppData 'registry-backups'
$sandboxSettings = Join-Path $sandboxAppData 'settings.json'

# ── Wipe and recreate filesystem ──
if (Test-Path $SandboxRoot) {
    Remove-Item -Path $SandboxRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $SandboxRoot -Force | Out-Null
New-Item -ItemType Directory -Path $sandboxSigs -Force | Out-Null
New-Item -ItemType Directory -Path $sandboxAppData -Force | Out-Null
New-Item -ItemType Directory -Path $sandboxBackups -Force | Out-Null

# Copy fixture signatures
if (Test-Path $sigFixtures) {
    Get-ChildItem -Path $sigFixtures -Recurse | ForEach-Object {
        $rel  = $_.FullName.Substring($sigFixtures.Length).TrimStart('\')
        $dest = Join-Path $sandboxSigs $rel
        if ($_.PSIsContainer) {
            New-Item -ItemType Directory -Path $dest -Force | Out-Null
        } else {
            $dir = Split-Path $dest -Parent
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Copy-Item -Path $_.FullName -Destination $dest -Force
        }
    }
}

# Copy settings
if (Test-Path $settingsFixture) {
    Copy-Item -Path $settingsFixture -Destination $sandboxSettings -Force
}

# Create empty log
Set-Content -Path $sandboxLog -Value '' -Encoding UTF8

# ── Wipe and recreate registry ──
if (Test-Path $testRegRoot) {
    Remove-Item -Path $testRegRoot -Recurse -Force
}

# Global MailSettings
$globalKey = "$testRegRoot\MailSettings"
New-Item -Path $globalKey -Force | Out-Null
Set-ItemProperty -Path $globalKey -Name 'NewSignature'   -Value 'Corporate Standard' -Type String
Set-ItemProperty -Path $globalKey -Name 'ReplySignature' -Value 'Personal'            -Type String

# Outlook Profiles — mirror the real structure:
#   Profiles\<ProfileName>\9375CFF0413111d3B88A00104B2A6676\<index>
$acctGuid    = '9375CFF0413111d3B88A00104B2A6676'
$profileName = 'Outlook'
$profileBase = "$testRegRoot\Profiles\$profileName\$acctGuid"

# Account 1: Alice (primary)
$acct1 = "$profileBase\00000001"
New-Item -Path $acct1 -Force | Out-Null
Set-ItemProperty -Path $acct1 -Name 'Account Name'    -Value 'Alice Johnson <alice@contoso.com>' -Type String
Set-ItemProperty -Path $acct1 -Name 'Display Name'    -Value 'Alice Johnson'                      -Type String
Set-ItemProperty -Path $acct1 -Name 'Email'            -Value 'alice@contoso.com'                  -Type String
Set-ItemProperty -Path $acct1 -Name 'New Signature'    -Value 'Corporate Standard'                 -Type String
Set-ItemProperty -Path $acct1 -Name 'Reply Signature'  -Value 'Personal'                           -Type String

# Account 2: Shared mailbox
$acct2 = "$profileBase\00000002"
New-Item -Path $acct2 -Force | Out-Null
Set-ItemProperty -Path $acct2 -Name 'Account Name'    -Value 'Shared Services <shared@contoso.com>' -Type String
Set-ItemProperty -Path $acct2 -Name 'Display Name'    -Value 'Shared Services'                        -Type String
Set-ItemProperty -Path $acct2 -Name 'Email'            -Value 'shared@contoso.com'                     -Type String
Set-ItemProperty -Path $acct2 -Name 'New Signature'    -Value 'Shared Mailbox Sig'                     -Type String
Set-ItemProperty -Path $acct2 -Name 'Reply Signature'  -Value 'Shared Mailbox Sig'                     -Type String

# Account 3: Bob (delegate)
$acct3 = "$profileBase\00000003"
New-Item -Path $acct3 -Force | Out-Null
Set-ItemProperty -Path $acct3 -Name 'Account Name'    -Value 'Bob (Delegate) <bob.delegate@contoso.com>' -Type String
Set-ItemProperty -Path $acct3 -Name 'Display Name'    -Value 'Bob (Delegate)'                              -Type String
Set-ItemProperty -Path $acct3 -Name 'Email'            -Value 'bob.delegate@contoso.com'                    -Type String
Set-ItemProperty -Path $acct3 -Name 'New Signature'    -Value ''                                            -Type String
Set-ItemProperty -Path $acct3 -Name 'Reply Signature'  -Value ''                                            -Type String

# ── Return sandbox config ──
$result = [PSCustomObject]@{
    SignaturePath = $sandboxSigs
    RegBase       = $globalKey
    ProfilesBase  = "$testRegRoot\Profiles"
    LogFile       = $sandboxLog
    BackupDir     = $sandboxBackups
    SettingsFile  = $sandboxSettings
    SandboxRoot   = $SandboxRoot
    RegistryRoot  = $testRegRoot
}

Write-Host "Sandbox ready at: $SandboxRoot"
Write-Host "  Signatures:  $sandboxSigs"
Write-Host "  Registry:    $testRegRoot"
Write-Host "  3 fixture signatures, 3 mailbox accounts"

return $result
