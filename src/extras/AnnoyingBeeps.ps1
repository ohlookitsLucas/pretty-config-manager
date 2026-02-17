# Annoying Beeps
# Produces a beep every 500 ms. Stop with Ctrl+C or by ending the process.
while ($true) {
    [console]::beep(800, 300)
    Start-Sleep -Milliseconds 500
}
