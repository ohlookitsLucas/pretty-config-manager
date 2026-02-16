<#
  close-app.ps1 - Closes the running outlookmAnAger window.
#>
Get-Process | Where-Object { $_.MainWindowTitle -eq 'outlookmAnAger' } | Stop-Process -Force -ErrorAction SilentlyContinue
Write-Host "Closed outlookmAnAger"
