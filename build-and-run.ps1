<#
  build-and-run.ps1
  Convenience script to launch the outlookmAnAger UI from this folder.
  Run in PowerShell 5.1. If script execution policy prevents running, start PS as Administrator
  and run: Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass
#>

$root = Split-Path -Parent $MyInvocation.MyCommand.Definition
Push-Location $root

# WPF requires the STA apartment state. If the current PowerShell session is not STA,
# relaunch a new powershell.exe process with -STA to run the UI.
try {
  $apartment = [System.Threading.Thread]::CurrentThread.ApartmentState
} catch {
  $apartment = $null
}

if ($apartment -ne 'STA') {
  Write-Host "Current PowerShell is not STA. Relaunching a new process with -STA to run the UI..."
  $psExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
  if (-not $psExe) { $psExe = 'powershell.exe' }
  $appPath = Join-Path $root 'src\App.ps1'
  $launchArgs = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$appPath`""
  Start-Process -FilePath $psExe -ArgumentList $launchArgs -WorkingDirectory $root
  Pop-Location
  return
}

# If already STA, just run the app script directly
& "$root\src\App.ps1"

Pop-Location
