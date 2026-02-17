# Fake Progress Bar
# Displays an official-looking "Installing Super Secret Update..." progress bar
# in the PowerShell console. Does absolutely nothing.
for ($i = 0; $i -le 100; $i++) {
    Write-Progress -Activity "Installing Super Secret Update..." `
                   -Status   "$i% complete" `
                   -PercentComplete $i
    Start-Sleep -Milliseconds 400
}
Write-Progress -Activity "Installing Super Secret Update..." -Completed
Write-Host "Update complete. Nothing happened." -ForegroundColor Green
Write-Host "Press any key to close..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
