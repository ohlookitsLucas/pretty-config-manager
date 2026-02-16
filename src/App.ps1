<#
  App.ps1 - launcher for outlookmAnAger
  Runs the WPF XAML and initializes the UI code-behind.
#>

Add-Type -AssemblyName PresentationFramework

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$xamlPath = Join-Path $scriptDir 'MainWindow.xaml'

if (-not (Test-Path $xamlPath)) {
    Write-Error "Cannot find XAML at $xamlPath"
    exit 1
}

[xml]$xaml = Get-Content -Path $xamlPath -Raw
# Create an XmlNodeReader from the document element for XamlReader
$reader = New-Object System.Xml.XmlNodeReader($xaml.DocumentElement)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Import modules and UI script
Import-Module (Join-Path $scriptDir 'SignatureManager.psm1') -Force
Import-Module (Join-Path $scriptDir 'PermissionsManager.psm1') -Force
. (Join-Path $scriptDir 'Ui.ps1')

Initialize-Ui -Window $window -ScriptRoot $scriptDir

# Show window and keep app running
$null = $window.ShowDialog()
