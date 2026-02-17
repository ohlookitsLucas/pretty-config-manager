<#
  Start-Sandbox.ps1
  Launches the real outlookmAnAger WPF app against sandbox data.
  All signature files and registry keys come from the test fixtures.
  COM (Outlook) and AD are replaced with in-memory mocks.

  Usage:
    .\tests\Start-Sandbox.ps1               # launch with sandbox data
    .\tests\Start-Sandbox.ps1 -NoTeardown   # keep sandbox after closing
#>
param(
    [switch]$NoTeardown
)

$ErrorActionPreference = 'Stop'

trap {
    Write-Host "`n=== CRASH ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Read-Host 'Press Enter to close'
    exit 1
}

# ── Setup sandbox ──
Write-Host 'Setting up sandbox environment...' -ForegroundColor Cyan
$sandbox = & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1')

# ── Load WPF ──
Add-Type -AssemblyName PresentationFramework

$srcDir   = Join-Path $PSScriptRoot '..\src'
$xamlPath = Join-Path $srcDir 'MainWindow.xaml'

if (-not (Test-Path $xamlPath)) {
    Write-Error "Cannot find XAML at $xamlPath"
    exit 1
}

[xml]$xaml = Get-Content -Path $xamlPath -Raw
$reader   = New-Object System.Xml.XmlNodeReader($xaml.DocumentElement)
$window   = [Windows.Markup.XamlReader]::Load($reader)

# ── Import modules ──
Import-Module (Join-Path $srcDir 'SignatureManager.psm1')  -Force
Import-Module (Join-Path $srcDir 'PermissionsManager.psm1') -Force

# ── Redirect SignatureManager to sandbox paths ──
Set-SignatureManagerPaths `
    -SignaturePath $sandbox.SignaturePath `
    -RegistryBase  $sandbox.RegBase `
    -ProfilesBase  $sandbox.ProfilesBase `
    -LogFile       $sandbox.LogFile `
    -BackupDir     $sandbox.BackupDir

# ── Inject mock COM + AD ──
. (Join-Path $PSScriptRoot 'Mocks\New-MockOutlookNS.ps1')

$mockNS = New-MockOutlookNS
Set-OutlookNSFactory { $mockNS }.GetNewClosure()
Set-ADSearchFactory  (New-MockADSearchFactory)

# ── Initialize UI ──
# Load Theme + Language first (extracted from Ui.ps1), then Ui.ps1 itself.
# Uses explicit UTF-8 read so Windows PowerShell 5.1 doesn't mangle non-ASCII chars.
$themeContent = Get-Content -Path (Join-Path $srcDir 'Theme.ps1')    -Raw -Encoding UTF8
$langContent  = Get-Content -Path (Join-Path $srcDir 'Language.ps1') -Raw -Encoding UTF8
$uiContent    = Get-Content -Path (Join-Path $srcDir 'Ui.ps1')      -Raw -Encoding UTF8
Invoke-Expression $themeContent
Invoke-Expression $langContent
Invoke-Expression $uiContent

# Redirect Ui.ps1's settings path to sandbox (set after dot-source)
$script:SettingsPath = $sandbox.SettingsFile

Initialize-Ui -Window $window -ScriptRoot (Resolve-Path $srcDir).Path

Write-Host ''
Write-Host '  Sandbox app is running.' -ForegroundColor Green
Write-Host '  Signatures folder:  ' $sandbox.SignaturePath
Write-Host '  Registry root:      ' $sandbox.RegistryRoot
Write-Host '  3 fixture sigs, 3 mock accounts, mock AD with 8 users'
Write-Host ''

# ── Catch unhandled WPF dispatcher exceptions (same as PrettyConfMan.ps1) ──
$app = [System.Windows.Application]::Current
if ($null -eq $app) { $app = New-Object System.Windows.Application; $app.ShutdownMode = 'OnExplicitShutdown' }
$app.add_DispatcherUnhandledException({
    param($s, $e)
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "`n[WPF Error $ts] $($e.Exception.Message)" -ForegroundColor Red
    $e.Handled = $true
})

# ── Show window ──
$null = $window.ShowDialog()

# ── Teardown ──
if (-not $NoTeardown) {
    Write-Host 'Tearing down sandbox...' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'Reset-TestEnvironment.ps1') -Teardown
    Write-Host 'Done.' -ForegroundColor Green
} else {
    Write-Host "Sandbox preserved at: $($sandbox.SandboxRoot)" -ForegroundColor Yellow
    Write-Host "Registry preserved at: $($sandbox.RegistryRoot)" -ForegroundColor Yellow
    Write-Host "Run: .\tests\Reset-TestEnvironment.ps1 -Teardown  to clean up" -ForegroundColor Yellow
}
