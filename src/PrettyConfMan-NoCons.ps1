<#
  PrettyConfMan-NoCons.ps1 - launches outlookmAnAger without a visible console window.
  Run this instead of PrettyConfMan.ps1 when you don't want the PowerShell console visible.
  For troubleshooting, run PrettyConfMan.ps1 directly instead.
#>
$appPath = Join-Path $PSScriptRoot 'PrettyConfMan.ps1'
Start-Process powershell.exe `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$appPath`"" `
    -WindowStyle Hidden
