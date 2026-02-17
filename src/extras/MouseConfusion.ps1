# Mouse Confusion
# Moves the mouse cursor slightly every few seconds — subtle and annoying.
Add-Type -AssemblyName System.Windows.Forms

while ($true) {
    Start-Sleep -Seconds 5
    $pos = [System.Windows.Forms.Cursor]::Position
    $pos.X += 10
    [System.Windows.Forms.Cursor]::Position = $pos
}
