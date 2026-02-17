# Infinite Popup
# Shows a friendly message box every 10 seconds.
# Stop via Task Manager -> End Task on the PowerShell process, or Ctrl+Alt+Del.
Add-Type -AssemblyName System.Windows.Forms

while ($true) {
    [System.Windows.Forms.MessageBox]::Show("Hello there!", "Friendly Message") | Out-Null
    Start-Sleep -Seconds 10
}
