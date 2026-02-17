<#
  Launch.ps1 - launches outlookmAnAger without a visible console window.
  Run this instead of App.ps1 when you don't want the PowerShell console visible.
  For troubleshooting, run App.ps1 directly instead.
#>
$appPath = Join-Path $PSScriptRoot 'src\App.ps1'
Start-Process powershell.exe `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$appPath`"" `
    -WindowStyle Hidden
