<#
  PrettyConfMan.ps1 - launcher for outlookmAnAger
  Runs the WPF XAML and initializes the UI code-behind.
#>

Add-Type -AssemblyName PresentationFramework

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$xamlPath = Join-Path $scriptDir 'MainWindow.xaml'

if (-not (Test-Path $xamlPath)) {
    Write-Error "Cannot find XAML at $xamlPath"
    exit 1
}

$crashLogPath = Join-Path $env:APPDATA 'outlookmAnAger\crash.log'

try {
    [xml]$xaml = Get-Content -Path $xamlPath -Raw
    # Create an XmlNodeReader from the document element for XamlReader
    $reader = New-Object System.Xml.XmlNodeReader($xaml.DocumentElement)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Catch unhandled WPF dispatcher exceptions so they don't silently kill the app
    $app = [System.Windows.Application]::Current
    if ($null -eq $app) { $app = New-Object System.Windows.Application; $app.ShutdownMode = 'OnExplicitShutdown' }
    $app.add_DispatcherUnhandledException({
        param($s, $e)
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $msg = "$ts`n$($e.Exception.GetType().FullName): $($e.Exception.Message)`n$($e.Exception.StackTrace)"
        if ($e.Exception.InnerException) {
            $msg += "`n--- Inner: $($e.Exception.InnerException.Message)`n$($e.Exception.InnerException.StackTrace)"
        }
        $logDir = Split-Path $crashLogPath -Parent
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $msg | Out-File -FilePath $crashLogPath -Append -Encoding UTF8
        $e.Handled = $true
    })

    # Import modules and UI scripts
    Import-Module (Join-Path $scriptDir 'SignatureManager.psm1') -Force
    Import-Module (Join-Path $scriptDir 'PermissionsManager.psm1') -Force
    . (Join-Path $scriptDir 'Theme.ps1')
    . (Join-Path $scriptDir 'Language.ps1')
    . (Join-Path $scriptDir 'Ui.ps1')
    . (Join-Path $scriptDir 'Ui.Signatures.ps1')
    . (Join-Path $scriptDir 'Ui.Permissions.ps1')
    . (Join-Path $scriptDir 'Ui.Wizard.ps1')
    . (Join-Path $scriptDir 'Ui.Extras.ps1')

    Initialize-Ui -Window $window -ScriptRoot $scriptDir

    # Show window and keep app running
    $null = $window.ShowDialog()
} catch {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $msg = "$ts`n$($_.Exception.GetType().FullName): $($_.Exception.Message)`n$($_.ScriptStackTrace)"
    $logDir = Split-Path $crashLogPath -Parent
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $msg | Out-File -FilePath $crashLogPath -Append -Encoding UTF8
    Write-Host "`n=== CRASH ===" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Host "`nDetails written to: $crashLogPath" -ForegroundColor Cyan
    Write-Host "Press Enter to close..." -ForegroundColor Cyan
    Read-Host
}
